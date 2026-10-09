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
let subsetSlides: string[] | undefined;
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
      slideCount: subsetSlides?.length ?? 1,
      slideWidth: 914400,
      slideHeight: 914400,
      defaultTextColor: null,
      majorFont: subsetSlides ? 'Calibri' : null,
      minorFont: null,
      hlinkColor: null,
      folHlinkColor: null,
      embeddedFonts: bootstrapEmbeddedFonts,
      slides: Array.from({ length: subsetSlides?.length ?? 1 }, (_, index) => ({ index, partName: `ppt/slides/slide${index + 1}.xml` })),
    }));
  }
  pull_slide(index = 0, _operation?: number, _generation?: number, _credit?: number): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      index,
      slideNumber: index + 1,
      partName: `ppt/slides/slide${index + 1}.xml`,
      background: null,
      elements: subsetSlides ? [{ type: 'shape', textBody: { vert: 'horz', paragraphs: [{ bullet: { type: 'none' }, runs: [{ type: 'text', text: subsetSlides[index], fontFamily: 'Calibri' }] }] } }] : [],
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
  loadOfficeFontFallbacks: fontMocks.load,
  unloadOfficeFontFallbacks: fontMocks.unload,
}));
vi.mock('@silurus/ooxml-core/internal/model-source', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core/internal/model-source')>(),
  openModelSourceModule: (...args: unknown[]) => openSourceMock(...args),
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

async function loadRenderWorker(variant: 'source' | 'default' = 'source'): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  if (variant === 'source') await import('./render-worker-source.js');
  else await import('./render-worker.js');
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
  subsetSlides = undefined;
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
    fake.onmessage?.({ data: {
      kind: 'parse', id: 40, buffer: new ArrayBuffer(4), resourcePolicy,
      source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js',
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
      source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js',
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


describe('Google subset readiness in progressive render owners', () => {
  it.each(['default', 'source'] as const)('loads a later subset of an already requested family before publication (%s)', async (variant) => {
    subsetSlides = ['A', 'B'];
    initMock.mockResolvedValue(undefined);
    if (variant === 'source') openSourceMock.mockResolvedValue({ archive: new FakePptxArchive(new Uint8Array()), viewDefaults: {}, close: vi.fn() });
    const a = deferred<void>(), b = deferred<void>();
    const faces: Array<{ source: string; loadCalls: number }> = [];
    class SubsetFace {
      loadCalls = 0;
      constructor(public family: string, public source: string) {}
      load() { this.loadCalls++; return (this.source.includes('a.woff2') ? a.promise : b.promise).then(() => this); }
    }
    vi.stubGlobal('FontFace', SubsetFace);
    vi.stubGlobal('fetch', vi.fn(async () => ({ ok: true, text: async () => `
      @font-face { font-family: Carlito; src: url(b.woff2); unicode-range: U+0042; }
      @font-face { font-family: Carlito; src: url(a.woff2); unicode-range: U+0041; }
    ` })));
    const fake = await loadRenderWorker(variant);
    fake.fonts = { add: (f: FontFace) => faces.push(f as unknown as typeof faces[number]), delete: () => true } as unknown as FontFaceSet;
    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({ data: { kind: 'parse', id: 71, buffer: new ArrayBuffer(4), resourcePolicy, useGoogleFonts: true, progressiveLayout: true,
      ...(variant === 'source' ? { source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js' } : {}),
    } } as MessageEvent);
    await vi.waitFor(() => expect(faces.map(f => f.loadCalls)).toEqual([0, 1]));
    expect(fake.posted.some((m) => (m as { kind: string }).kind === 'presentationLayoutPartial')).toBe(false);
    a.resolve();
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationLayoutPartial', availableSlides: 1 })));
    fake.onmessage?.({ data: { kind: 'continuePresentationPreflight', forId: 71, availableSlides: 1 } } as MessageEvent);
    await vi.waitFor(() => expect(faces.map(f => f.loadCalls)).toEqual([1, 1]));
    expect(fake.posted).not.toContainEqual(expect.objectContaining({ kind: 'presentationLayoutPartial', availableSlides: 2 }));
    b.resolve();
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationLayoutPartial', availableSlides: 2 })));
    expect(faces).toHaveLength(2);
    fake.onmessage?.({ data: { kind: 'continuePresentationPreflight', forId: 71, availableSlides: 2 } } as MessageEvent);
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationReady', id: 71 })));
  });

  // Main mode never forwards renderer descriptors; the worker must load the
  // same subset for the same text-only slide when one is registered.
  it.each(['default', 'source'] as const)('a registered but unpainted renderer does not widen text-only subset loading (%s)', async (variant) => {
    subsetSlides = ['A'];
    initMock.mockResolvedValue(undefined);
    fontMocks.requests.mockReturnValue([]);
    fontMocks.render.mockResolvedValue(undefined);
    if (variant === 'source') openSourceMock.mockResolvedValue({ archive: new FakePptxArchive(new Uint8Array()), viewDefaults: {}, close: vi.fn() });
    const faces: Array<{ source: string; status: string }> = [];
    class SubsetFace {
      status = 'unloaded';
      constructor(public family: string, public source: string) {}
      load() { this.status = 'loaded'; return Promise.resolve(this); }
    }
    vi.stubGlobal('FontFace', SubsetFace);
    vi.stubGlobal('OffscreenCanvas', class { constructor(_width: number, _height: number) {} });
    // Same stylesheet as above: core caches it per URL across module resets.
    // Demand for 'A' plus the ii/M/space sentinels excludes only U+0042.
    vi.stubGlobal('fetch', vi.fn(async () => ({ ok: true, text: async () => `
      @font-face { font-family: Carlito; src: url(b.woff2); unicode-range: U+0042; }
      @font-face { font-family: Carlito; src: url(a.woff2); unicode-range: U+0041; }
    ` })));
    const fake = await loadRenderWorker(variant);
    fake.fonts = { add: (f: FontFace) => faces.push(f as unknown as typeof faces[number]), delete: () => true } as unknown as FontFaceSet;
    const send = (data: unknown) => fake.onmessage?.({ data } as MessageEvent);
    send({ kind: 'init', wasmUrl: 'x' });
    send({ kind: 'parse', id: 90, buffer: new ArrayBuffer(4), resourcePolicy, useGoogleFonts: true,
      renderers: { chartEx: { protocol: 'ooxml-worker-renderer-module/v1', builtin: 'chartEx' } },
      ...(variant === 'source' ? { source: modelSource, sourceOwnerUrl: './internal/worker-presentation-source.js' } : {}),
    });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'presentationReady', id: 90 })));
    // collectRuns awaits the worker's settled Google preload.
    send({ kind: 'collectRuns', id: 91, slideIndex: 0, width: 100 });
    await vi.waitFor(() => expect(fake.posted).toContainEqual(expect.objectContaining({ kind: 'runsCollected', id: 91 })));
    expect(faces.filter((f) => f.status === 'loaded').map((f) => f.source)).toEqual([expect.stringContaining('/a.woff2')]);
  });
});
