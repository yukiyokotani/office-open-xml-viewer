import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

/**
 * AR4 (worker-render mode twin of worker-init-hang): the render-capable worker
 * carried the same `ready`-flag hazard. After the initPromise conversion a
 * REJECTED WASM init must reject the pending request (`error` response) rather
 * than leaving `load()` hanging. Driven against a mocked WASM module + stubbed
 * `self`; only the parse arm is exercised (no OffscreenCanvas render needed).
 */

const initMock = vi.fn();
const openSourceMock = vi.fn();
const fontMocks = vi.hoisted(() => ({
  load: vi.fn(),
  unload: vi.fn(),
  requests: vi.fn(),
  render: vi.fn(),
}));
let bootstrapEmbeddedFonts: unknown[] = [];
let extractedFontCount = 0;
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
      slideCount: 1,
      slideWidth: 914400,
      slideHeight: 914400,
      defaultTextColor: null,
      majorFont: null,
      minorFont: null,
      hlinkColor: null,
      folHlinkColor: null,
      embeddedFonts: bootstrapEmbeddedFonts,
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    }));
  }
  pull_slide(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      index: 0,
      slideNumber: 1,
      partName: 'ppt/slides/slide1.xml',
      background: null,
      elements: [],
      notes: 'worker note',
      hidden: true,
    }));
  }
  slide_cursor_resource_usage(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      archiveEntryCount: 1,
      declaredInflatedBytes: 1,
      distinctInflatedBytes: 1,
      operationInflatedBytes: 1,
    }));
  }
  acknowledge_slide(): void {}
  cancel_slide(): void {}
  close_presentation_session(): void {}
  assert_healthy(): void {}
  extract_media(): Uint8Array {
    return new Uint8Array([1]);
  }
  extract_image(): Uint8Array {
    return new Uint8Array([2]);
  }
  extract_font(): Uint8Array {
    extractedFontCount += 1;
    return new Uint8Array([3]);
  }
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: (arg: unknown) => initMock(arg),
  // RB6: mirror the worker's `reinit` recovery hook (see worker-init-hang.test).
  reinit: (arg: unknown) => initMock(arg),
  PptxArchive: FakePptxArchive,
}));
vi.mock('@silurus/ooxml-core', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  openModelSourceModule: (...args: unknown[]) => openSourceMock(...args),
  loadOfficeFontFallbacks: fontMocks.load,
  unloadOfficeFontFallbacks: fontMocks.unload,
}));
vi.mock('./google-fonts', async (importOriginal) => ({
  ...await importOriginal<typeof import('./google-fonts')>(),
  pptxSlideOfficeFontRequests: fontMocks.requests,
}));
vi.mock('./renderer', () => ({ renderSlideWithEmbeddedFonts: fontMocks.render }));

const modelSource = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'pptx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
} as const;

interface FakeSelf {
  onmessage: ((e: MessageEvent) => void) | null;
  posted: unknown[];
  postMessage: (msg: unknown, transfer?: Transferable[]) => void;
  fonts?: FontFaceSet;
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

async function loadRenderWorker(): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  await import('./render-worker.js');
  return fake;
}

beforeEach(() => {
  initMock.mockReset();
  openSourceMock.mockReset();
  fontMocks.load.mockReset();
  fontMocks.unload.mockReset();
  fontMocks.requests.mockReset();
  fontMocks.render.mockReset();
  bootstrapEmbeddedFonts = [];
  extractedFontCount = 0;
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('pptx render-worker.ts — init failure never hangs a request (AR4)', () => {
  it('preflights a model-source cursor without initializing OOXML WASM', async () => {
    const archive = new FakePptxArchive(new Uint8Array());
    openSourceMock.mockResolvedValue({ archive, viewDefaults: {}, close: vi.fn() });
    const fake = await loadRenderWorker();
    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({ data: {
      kind: 'parse', id: 40, buffer: new ArrayBuffer(4), resourcePolicy,
      source: modelSource,
    } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationReady', id: 40,
    })));
    expect(initMock).not.toHaveBeenCalled();
    expect(openSourceMock).toHaveBeenCalledTimes(1);

    fake.onmessage?.({ data: { kind: 'toMarkdown', id: 41 } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 41, message: 'Markdown conversion is unsupported for this source',
    })));
  });

  it('closes a model source when bootstrap traps', async () => {
    const archive = new FakePptxArchive(new Uint8Array());
    vi.spyOn(archive, 'presentation_bootstrap').mockImplementation(() => {
      throw new WebAssembly.RuntimeError('render bootstrap trap');
    });
    const closeArchive = vi.fn();
    openSourceMock.mockResolvedValue({ archive, viewDefaults: {}, close: closeArchive });
    const fake = await loadRenderWorker();
    fake.onmessage?.({ data: {
      kind: 'parse', id: 42, buffer: new ArrayBuffer(4), resourcePolicy,
      source: modelSource,
    } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 42, message: expect.stringContaining('render bootstrap trap'),
    })));
    expect(closeArchive).toHaveBeenCalledTimes(1);
    expect(initMock).not.toHaveBeenCalled();
  });

  it('a parse after a REJECTED init responds with an error (not a hang)', async () => {
    initMock.mockRejectedValue(new Error('render wasm boom'));
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 9, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'error')).toBe(true);
    });

    const err = fake.posted.find((m) => (m as { kind?: string }).kind === 'error') as {
      id: number;
      message: string;
    };
    expect(err.id).toBe(9);
    expect(err.message).toContain('boom');
  });

  it('a parse after a SUCCESSFUL init responds with compact preflight and no ready handshake', async () => {
    initMock.mockResolvedValue(undefined);
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 2, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'presentationReady')).toBe(true);
    });

    const ready = fake.posted.find(
      (message) => (message as { kind?: string }).kind === 'presentationReady',
    ) as { preflight: { slides: Array<{ notes: string | null; hidden: boolean }> } };
    expect(ready.preflight.slides).toEqual([
      expect.objectContaining({ notes: 'worker note', hidden: true }),
    ]);
    expect(initMock).toHaveBeenCalledTimes(1);

    expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'ready')).toBe(false);
  });

  it('loads embedded font parts into the worker FontFaceSet', async () => {
    initMock.mockResolvedValue(undefined);
    bootstrapEmbeddedFonts = [{
      fontName: 'Worker Deck Font',
      style: 'boldItalic',
      partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }];
    const added: Array<{ family: string; descriptors: FontFaceDescriptors; loadCalls: number }> = [];
    class FakeFontFace {
      loadCalls = 0;
      constructor(public family: string, _source: ArrayBuffer, public descriptors: FontFaceDescriptors) {}
      load() { this.loadCalls += 1; return Promise.resolve(this); }
    }
    vi.stubGlobal('FontFace', FakeFontFace);
    const fake = await loadRenderWorker();
    fake.fonts = {
      add: (face: FontFace) => { added.push(face as unknown as typeof added[number]); },
      ready: Promise.resolve(),
    } as unknown as FontFaceSet;

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 22, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);

    await vi.waitFor(() => expect(added).toHaveLength(1));
    expect(extractedFontCount).toBe(1);
    expect(added[0]).toMatchObject({
      family: expect.stringMatching(/^__ooxml_pptx_/),
      descriptors: { weight: 'bold', style: 'italic' },
      loadCalls: 1,
    });
  });

  it('rejects a second parse reserved while the first render-worker parse is opening', async () => {
    const init = deferred<void>();
    initMock.mockReturnValue(init.promise);
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 12, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);
    fake.onmessage?.({
      data: { kind: 'parse', id: 13, buffer: new ArrayBuffer(4), resourcePolicy },
    } as MessageEvent);

    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error',
      id: 13,
      code: 'ooxml-pptx-parse-already-started',
    })));
    init.resolve();
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationReady',
      id: 12,
    })));
  });

  it('keeps the current slide and its Office font after rejecting a second parse', async () => {
    initMock.mockResolvedValue(undefined);
    const face = { family: 'active-deck-face' };
    const route = { family: 'active-deck-face' };
    fontMocks.requests.mockReturnValue([{ family: 'Calibri', weight: 400, style: 'normal' }]);
    fontMocks.load.mockResolvedValue({ faces: [face], routes: { 'calibri:400:normal': route } });
    fontMocks.render.mockResolvedValue(undefined);
    vi.stubGlobal('OffscreenCanvas', class { constructor(_width: number, _height: number) {} });
    const fake = await loadRenderWorker();
    const send = (data: unknown) => fake.onmessage?.({ data } as MessageEvent);

    send({ kind: 'init', wasmUrl: 'x' });
    send({ kind: 'parse', id: 30, buffer: new ArrayBuffer(4), resourcePolicy });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'presentationReady', id: 30,
    })));
    send({ kind: 'collectRuns', id: 31, slideIndex: 0, width: 100 });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'runsCollected', id: 31,
    })));
    expect(fontMocks.load).toHaveBeenCalledTimes(1);

    send({ kind: 'parse', id: 32, buffer: new ArrayBuffer(4), resourcePolicy });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'error', id: 32, code: 'ooxml-pptx-parse-already-started',
    })));
    send({ kind: 'collectRuns', id: 33, slideIndex: 0, width: 100 });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({
      kind: 'runsCollected', id: 33,
    })));

    expect(fontMocks.unload).not.toHaveBeenCalled();
    expect(fontMocks.load).toHaveBeenCalledTimes(1);
    expect(fontMocks.render).toHaveBeenLastCalledWith(
      expect.anything(), expect.anything(), expect.any(Number), expect.any(Number),
      expect.objectContaining({ officeFontRoutes: { 'calibri:400:normal': route } }),
      expect.any(Function),
    );
  });
});
