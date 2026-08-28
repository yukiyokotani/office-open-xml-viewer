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
async function loadWorker(): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  await import('./worker.js');
  return fake;
}

beforeEach(() => {
  initMock.mockReset();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('pptx worker.ts — init failure never hangs a request (AR4)', () => {
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
