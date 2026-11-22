import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

/**
 * AR4: a WASM-init failure must not leave `load()` hanging forever. The worker
 * used to swallow the init error (log-only) and keep a `ready` flag false, so the
 * main thread — which blocked on a `ready` handshake — never resolved. The fix
 * moves pptx to the docx/xlsx `initPromise` pattern: every request `await`s the
 * init promise, so a rejected init rejects the request (surfacing an `error`
 * response the bridge turns into a rejected `load()`), never a silent hang.
 *
 * These drive the worker's `onmessage` directly against a mocked WASM module and
 * a stubbed `self`, so no real Worker / WASM is needed.
 */

const initMock = vi.fn();
const openSourceMock = vi.fn();
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => { resolve = resolvePromise; });
  return { promise, resolve };
}
const resourcePolicy = {
  maxArchiveEntryBytes: null,
  maxTotalInflatedBytes: null,
  maxArchiveEntries: null,
} as const;
class FakePptxArchive {
  constructor(_bytes: Uint8Array, _max?: bigint) {}
  presentation_bootstrap(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      slideCount: 0,
      slideWidth: 914400,
      slideHeight: 914400,
      defaultTextColor: null,
      majorFont: null,
      minorFont: null,
      hlinkColor: null,
      folHlinkColor: null,
      embeddedFonts: [],
      slides: [],
    }));
  }
  close_presentation_session(): void {}
  assert_healthy(): void {}
  extract_media(_p: string): Uint8Array {
    return new Uint8Array([1, 2, 3]);
  }
  extract_image(_p: string): Uint8Array {
    return new Uint8Array([4, 5, 6]);
  }
  extract_font(_p: string): Uint8Array {
    return new Uint8Array([7, 8, 9]);
  }
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: (arg: unknown) => initMock(arg),
  // RB6: the worker wires `reinit` (the forced-fresh-instance recovery hook) into
  // WasmParserHost. These init-hang tests never trap, so route it through the same
  // init mock as `default`; the recovery semantics are proven in the core / node
  // suites against the real glue.
  reinit: (arg: unknown) => initMock(arg),
  PptxArchive: FakePptxArchive,
}));
vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
}));
vi.mock('@silurus/ooxml-core/internal/model-source', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core/internal/model-source')>(),
  openModelSourceModule: (...args: unknown[]) => openSourceMock(...args),
}));

const modelSource = {
  protocol: 'ooxml-model-source-module/v2',
  target: 'pptx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
} as const;

/** A model-source archive: the parser archive's reads minus every optional capability. */
function sourceArchive() {
  const parser = new FakePptxArchive(new Uint8Array());
  return {
    presentation_bootstrap: () => parser.presentation_bootstrap(),
    close_presentation_session: () => parser.close_presentation_session(),
    assert_healthy: () => parser.assert_healthy(),
    extract_image: (path: string) => parser.extract_image(path),
  };
}

interface FakeSelf {
  onmessage: ((e: MessageEvent) => void) | null;
  posted: unknown[];
  postMessage: (msg: unknown, transfer?: Transferable[]) => void;
}

function installSelf(): FakeSelf {
  const posted: unknown[] = [];
  const fake: FakeSelf = {
    onmessage: null,
    posted,
    postMessage: (msg: unknown) => {
      posted.push(msg);
    },
  };
  vi.stubGlobal('self', fake);
  return fake;
}

/** Import the worker module fresh (its top-level `self.onmessage = …` runs on
 *  import), after `self` and the WASM mock are installed. */
async function loadWorker(variant: 'default' | 'source' = 'source'): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  if (variant === 'source') await import('./worker-source.js');
  else await import('./worker.js');
  return fake;
}

beforeEach(() => {
  initMock.mockReset();
  openSourceMock.mockReset();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('pptx worker.ts — init failure never hangs a request (AR4)', () => {
  it('transfers only source-returned image, media and font subarray bytes', async () => {
    const bytes = (middle: number) => new Uint8Array([91, middle, 92]).subarray(1, 2);
    const archive = {
      ...sourceArchive(),
      extract_image: vi.fn(() => bytes(1)),
      extract_media: vi.fn(() => bytes(2)),
      extract_font: vi.fn(() => bytes(3)),
    };
    openSourceMock.mockResolvedValue({ archive, viewDefaults: {}, close: vi.fn() });
    const fake = await loadWorker();
    fake.onmessage?.({ data: {
      kind: 'parse', id: 40, buffer: new ArrayBuffer(4), resourcePolicy, source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js',
    } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationOpened', id: 40 })));
    for (const [kind, response, id, expected] of [
      ['extractImage', 'imageExtracted', 41, 1],
      ['extractMedia', 'mediaExtracted', 42, 2],
      ['extractFont', 'fontExtracted', 43, 3],
    ] as const) {
      fake.onmessage?.({ data: { kind, id, path: 'part' } } as MessageEvent);
      await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: response, id })));
      const posted = fake.posted.find((item) => (item as { id?: number }).id === id) as { bytes: ArrayBuffer };
      expect(Array.from(new Uint8Array(posted.bytes))).toEqual([expected]);
    }
  });
  it('opens a model-source cursor without initializing OOXML WASM and closes it on a trap', async () => {
    const archive = sourceArchive();
    const close = vi.fn();
    openSourceMock.mockResolvedValue({ archive, viewDefaults: {}, close });
    const fake = await loadWorker();
    fake.onmessage?.({ data: {
      kind: 'parse', id: 30, buffer: new ArrayBuffer(4), resourcePolicy, source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js',
    } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationOpened', id: 30,
    })));
    expect(initMock).not.toHaveBeenCalled();
    expect(openSourceMock).toHaveBeenCalledTimes(1);

    // Missing optional capabilities degrade instead of reaching for OOXML.
    fake.onmessage?.({ data: { kind: 'resourceUsage', id: 31 } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual({ kind: 'resourceUsage', id: 31, usage: undefined }));
    fake.onmessage?.({ data: { kind: 'extractMedia', id: 34, path: 'ppt/media/1' } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 34, message: 'media extraction is unsupported for this source',
    })));
    expect(close).not.toHaveBeenCalled();

    vi.spyOn(archive, 'extract_image').mockImplementation(() => {
      throw new WebAssembly.RuntimeError('source trap');
    });
    fake.onmessage?.({ data: { kind: 'extractImage', id: 32, path: 'ppt/media/1' } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 32, message: expect.stringContaining('source trap'),
    })));
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('terminally closes a model source on bootstrap trap without losing the trap', async () => {
    const archive = sourceArchive();
    vi.spyOn(archive, 'presentation_bootstrap').mockImplementation(() => {
      throw new WebAssembly.RuntimeError('bootstrap trap');
    });
    const close = vi.fn(() => { throw new Error('cleanup failed'); });
    openSourceMock.mockResolvedValue({ archive, viewDefaults: {}, close });
    const fake = await loadWorker();
    fake.onmessage?.({ data: {
      kind: 'parse', id: 33, buffer: new ArrayBuffer(4), resourcePolicy, source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js',
    } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 33, message: expect.stringContaining('bootstrap trap'),
    })));
    expect(close).toHaveBeenCalledTimes(1);
    expect(initMock).not.toHaveBeenCalled();
  });

  it('a parse after a REJECTED init responds with an error (not a hang)', async () => {
    initMock.mockRejectedValue(new Error('wasm boom'));
    const fake = await loadWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 7, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    // Let the awaited (rejected) initPromise settle and the handler run its catch.
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'error')).toBe(true);
    });

    const err = fake.posted.find((m) => (m as { kind?: string }).kind === 'error') as {
      kind: string;
      id: number;
      message: string;
    };
    expect(err.id).toBe(7);
    expect(err.message).toContain('boom');
    // Crucially: the request settled — no pending promise is left hanging.
  });

  it('a parse after a SUCCESSFUL init responds with a compact bootstrap', async () => {
    initMock.mockResolvedValue(undefined);
    const fake = await loadWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 3, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'presentationOpened')).toBe(true);
    });

    const parsed = fake.posted.find((m) => (m as { kind?: string }).kind === 'presentationOpened') as {
      kind: string;
      id: number;
    };
    expect(parsed.id).toBe(3);
    expect(initMock).toHaveBeenCalledTimes(1);
    // No `ready` handshake is emitted anymore (initPromise pattern replaces it).
    expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'ready')).toBe(false);

    fake.onmessage?.({
      data: { kind: 'extractFont', id: 4, path: 'ppt/fonts/font1.fntdata' },
    } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'fontExtracted',
      id: 4,
    })));
    const extracted = fake.posted.find((message) =>
      (message as { kind?: string }).kind === 'fontExtracted') as { bytes: ArrayBuffer };
    expect(Array.from(new Uint8Array(extracted.bytes))).toEqual([7, 8, 9]);
  });

  it('rejects a second parse reserved while the first parse is still opening', async () => {
    const init = deferred<void>();
    initMock.mockReturnValue(init.promise);
    const fake = await loadWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 10, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 11, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);

    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error',
      id: 11,
      code: 'ooxml-pptx-parse-already-started',
    })));
    init.resolve();
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationOpened',
      id: 10,
    })));
  });

  it('correlates a synchronous slide-open reservation failure instead of hanging', async () => {
    initMock.mockResolvedValue(undefined);
    const fake = await loadWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: {
        kind: 'openSlideSession',
        id: 20,
        slideIndex: 0,
        sessionId: 0,
        operationId: 1,
        generation: 1,
      },
    } as MessageEvent);

    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error',
      id: 20,
      message: expect.stringContaining('session id'),
    })));
  });
});


it.each(['default', 'source'] as const)('carries completed compact demand from the parser owner (%s)', async (variant) => {
  initMock.mockResolvedValue(undefined);
  if (variant === 'source') openSourceMock.mockResolvedValue({ archive: sourceArchive(), viewDefaults: {}, close: vi.fn() });
  const fake = await loadWorker(variant);
  fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
  fake.onmessage?.({ data: { kind: 'parse', id: 80, buffer: new ArrayBuffer(4), resourcePolicy, cjkFallback: 'jp', collectFontDemand: new URL('./font-preload-demand.ts', import.meta.url).href,
    ...(variant === 'source' ? { source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js' } : {}),
  } } as MessageEvent);
  await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationOpened', id: 80 })));
  fake.onmessage?.({ data: { kind: 'finishPresentationPreflight', id: 81 } } as MessageEvent);
  await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
    kind: 'presentationPreflightReady', id: 81, preflight: expect.objectContaining({ fontPreloadDemand: [] }),
  })));
});

it.each(['default', 'source'] as const)('still loads when the collector sidecar is not served (%s)', async (variant) => {
  initMock.mockResolvedValue(undefined);
  if (variant === 'source') openSourceMock.mockResolvedValue({ archive: sourceArchive(), viewDefaults: {}, close: vi.fn() });
  const fake = await loadWorker(variant);
  fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
  // A bundler that does not emit module-relative assets (esbuild) leaves the URL unserved.
  fake.onmessage?.({ data: { kind: 'parse', id: 90, buffer: new ArrayBuffer(4), resourcePolicy, cjkFallback: 'jp', collectFontDemand: './missing-font-demand-sidecar.mjs',
    ...(variant === 'source' ? { source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js' } : {}),
  } } as MessageEvent);
  await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationOpened', id: 90 })));
  fake.onmessage?.({ data: { kind: 'finishPresentationPreflight', id: 91 } } as MessageEvent);
  await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
    kind: 'presentationPreflightReady', id: 91, preflight: expect.objectContaining({ fontPreloadDemand: expect.anything() }),
  })));
  // Every slide then demands every subset (font-demand-collector.ts).
  const { importFontDemandCollector } = await import('./font-demand-collector.js');
  for (const url of ['./missing-font-demand-sidecar.mjs', '', new URL('./missing-font-demand-sidecar.mjs', import.meta.url).href,
    // A served module without the export.
    new URL('./office-face-name.ts', import.meta.url).href]) {
    const collector = await importFontDemandCollector(url);
    expect(collector({ index: 0, elements: [] } as never, new Set(), 1000), url).toBe('all');
  }
});
