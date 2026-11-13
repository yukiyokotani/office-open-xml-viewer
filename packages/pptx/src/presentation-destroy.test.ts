import { describe, it, expect, afterEach, beforeEach, vi } from 'vitest';
import {
  WorkerBridge,
  type ModelSource,
  type WorkerLike,
} from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
let { PptxPresentation } = await import('./presentation');
let { registerEmbeddedFonts } = await import('@silurus/ooxml-core');
import type { PptxPresentation as Presentation } from './presentation';
let { loadEmbeddedFonts } = await import('./embedded-fonts');
// Font sources must come from the module instance the presentation uses
// (`fonts` accepts only that instance's factory results).
let { fontFiles, googleFonts } = await import('@silurus/ooxml-core/internal/font-sources');
import type { PptxEmbeddedFontRef } from './worker-protocol';

/**
 * `PptxPresentation.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight so a
 * `load()` / image extraction awaiting the worker cannot hang after the deck is
 * disposed. Pinned with a real {@link WorkerBridge} over an in-memory worker
 * (the constructor opens a real Worker, so we build off-prototype and inject
 * the collaborators destroy() touches — the pattern from
 * `presentation.image.test.ts`).
 */

class SilentWorker implements WorkerLike {
  static instances: SilentWorker[] = [];

  constructor() {
    SilentWorker.instances.push(this);
  }

  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminated = false;
  terminate(): void {
    this.terminated = true;
  }
}

interface DestroyProbe {
  destroy(): void;
}

// ── Fake FontFaceSet so destroy()'s Google-Fonts release is observable ───────
const G = globalThis as Record<string, unknown>;
const ORIG_FONTS = {
  document: G.document,
  self: G.self,
  fetch: G.fetch,
  FontFace: G.FontFace,
  Worker: G.Worker,
  location: G.location,
};
beforeEach(async () => {
  vi.resetModules();
  ({ PptxPresentation } = await import('./presentation'));
  ({ registerEmbeddedFonts } = await import('@silurus/ooxml-core'));
  ({ loadEmbeddedFonts } = await import('./embedded-fonts'));
  ({ fontFiles, googleFonts } = await import('@silurus/ooxml-core/internal/font-sources'));
});

afterEach(() => {
  G.document = ORIG_FONTS.document;
  G.self = ORIG_FONTS.self;
  G.fetch = ORIG_FONTS.fetch;
  G.FontFace = ORIG_FONTS.FontFace;
  G.Worker = ORIG_FONTS.Worker;
  G.location = ORIG_FONTS.location;
  SilentWorker.instances = [];
  vi.restoreAllMocks();
});

const CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
interface FakeFace { family: string }
function installFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: string | ArrayBuffer, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  // The stylesheet, and the face files the library fetches itself.
  G.fetch = async (url: string) => ({ ok: true, status: 200, url, text: async () => CSS, arrayBuffer: async () => new Uint8Array(16).buffer });
  delete G.self;
  return { added };
}

describe('PptxPresentation.destroy() — rejects in-flight worker requests', () => {
  function makePresentation() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance._rawParts = new BoundedRawPartCache({ maxEntries: 2, maxBytes: 1024 });
    instance._googleFontFaces = [];
    instance._officeFontLoads = new Map();
    instance._symbolSentinels = new Map();
    instance._embeddedFontFaces = [];
    instance._layoutWaiters = new Set();
    instance._layoutLifecycle = new ProgressiveLayoutLifecycle();
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { pres: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { pres, bridge, worker } = makePresentation();
    const inFlight = bridge.request((id) => ({ id }));
    pres.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice', () => {
    const { pres } = makePresentation();
    pres.destroy();
    expect(() => pres.destroy()).not.toThrow();
  });

  it('does not start a font load when a pending slide read completes after destroy', async () => {
    const { pres } = makePresentation();
    const instance = pres as unknown as Record<string, unknown>;
    instance._mode = 'main';
    instance._preflight = {
      slideCount: 1, slideWidth: 914400, slideHeight: 914400,
      majorFont: null, minorFont: 'Calibri', defaultTextColor: null,
    };
    instance._availableSlideCount = 1;
    let releaseSlide: (() => void) | undefined;
    const slideReady = new Promise<void>((resolve) => { releaseSlide = resolve; });
    const slide = { elements: [{ type: 'shape', textBody: { paragraphs: [
      { bullet: { type: 'none' }, runs: [{ type: 'text', text: 'A', fontFamily: 'Calibri' }] },
    ] } }] };
    const withSlide = vi.fn(async (_index: number, read: (slide: unknown) => unknown) => {
      await slideReady;
      return read(slide);
    });
    instance._slides = { withSlide, clear: vi.fn() };
    const loadRoutes = vi.spyOn(pres as unknown as {
      _officeRoutesForRequests(requests: unknown[], set: FontFaceSet | null): Promise<object>;
    }, '_officeRoutesForRequests').mockResolvedValue({});
    const rendering = (pres as unknown as Presentation).renderSlide({} as HTMLCanvasElement, 0);
    await Promise.resolve();
    pres.destroy();
    releaseSlide?.();
    await expect(rendering).rejects.toThrow(/destroyed/i);
    expect(loadRoutes).not.toHaveBeenCalled();
    expect(withSlide).toHaveBeenCalledTimes(1);
  });

  it('terminates the owned worker when a partially initialized load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(
          buffer: ArrayBuffer,
          resourcePolicy: object,
          fonts?: unknown,
          timeoutMs?: number,
        ): Promise<void>;
      },
      '_parse',
    ).mockRejectedValueOnce(failure);

    await expect(PptxPresentation.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('terminates an opened model-source worker when release throws after parse', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('release failed');
    const release = vi.fn(() => { throw failure; });
    const source: ModelSource<'pptx'> = {
      target: 'pptx',
      claim: () => true,
      beginLoad: () => ({
        module: {
          protocol: 'ooxml-model-source-module/v2',
          target: 'pptx', moduleUrl: 'https://example.test/source.mjs', config: {},
        },
        release,
      }),
    };
    vi.spyOn(PptxPresentation.prototype as unknown as {
      _parse(buffer: ArrayBuffer, resourcePolicy: object): Promise<void>;
    }, '_parse').mockResolvedValueOnce(undefined);

    await expect(PptxPresentation.load(new ArrayBuffer(0), { legacyPpt: source }))
      .rejects.toBe(failure);
    expect(release).toHaveBeenCalledOnce();
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('preserves the load error and terminates directly when destroy throws', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(
          buffer: ArrayBuffer,
          resourcePolicy: object,
          fonts?: unknown,
          timeoutMs?: number,
        ): Promise<void>;
      },
      '_parse',
    ).mockRejectedValueOnce(failure);
    vi.spyOn(PptxPresentation.prototype, 'destroy').mockImplementationOnce(() => {
      throw new Error('cleanup failure');
    });

    await expect(PptxPresentation.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('main-mode load registers embedded fonts before returning the presentation', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const { added } = installFontFaceSet();
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(buffer: ArrayBuffer, resourcePolicy: object): Promise<void>;
      },
      '_parse',
    ).mockImplementationOnce(async function (this: Presentation) {
      const embeddedFonts: PptxEmbeddedFontRef[] = [{
        fontName: 'Main Deck Font',
        style: 'regular' as const,
        partPath: 'ppt/fonts/font1.fntdata',
        contentType: 'application/x-font-ttf',
      }];
      (this as unknown as { _preflight: object })._preflight = {
        slideCount: 0,
        slideWidth: 914400,
        slideHeight: 914400,
        defaultTextColor: null,
        majorFont: null,
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts,
        slides: [],
        fontPreloadNames: [],
      };
      const loaded = await loadEmbeddedFonts(
        embeddedFonts,
        async () => new Uint8Array([0, 1, 0, 0]),
      );
      (this as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces = loaded.faces;
      (this as unknown as { _embeddedFontAliases: ReadonlyMap<string, string> })
        ._embeddedFontAliases = loaded.aliases;
      (this as unknown as { _embeddedFontAuthoredFamilies: ReadonlyMap<string, string> })
        ._embeddedFontAuthoredFamilies = loaded.authoredFamilies;
    });

    const presentation = await PptxPresentation.load(new ArrayBuffer(0));
    expect(added).toEqual([expect.objectContaining({ family: expect.stringMatching(/^__ooxml_pptx_/) })]);
    presentation.destroy();
    expect(added).toHaveLength(0);
  });

  it('does not request remote fonts without font sources', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    installFontFaceSet();
    const fetch = vi.fn(async () => ({ ok: true, text: async () => CSS, arrayBuffer: async () => new Uint8Array(16).buffer }));
    G.fetch = fetch;
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(buffer: ArrayBuffer, resourcePolicy: object): Promise<void>;
      },
      '_parse',
    ).mockImplementationOnce(async function (this: Presentation) {
      (this as unknown as { _preflight: object })._preflight = {
        slideCount: 0,
        slideWidth: 914400,
        slideHeight: 914400,
        defaultTextColor: null,
        majorFont: 'Noto Sans CJK SC',
        minorFont: null,
        hlinkColor: null,
        folHlinkColor: null,
        embeddedFonts: [],
        slides: [],
        fontPreloadNames: ['Noto Sans CJK SC'],
      };
    });

    const presentation = await PptxPresentation.load(new ArrayBuffer(0), {
      mode: 'main',
    });

    expect(fetch).not.toHaveBeenCalled();
    presentation.destroy();
  });

  it('terminates directly when construction fails before the factory owns an instance', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'not a valid base URL' };

    await expect(
      PptxPresentation.load(new ArrayBuffer(0), { wasmUrl: 'relative.wasm' }),
    ).rejects.toThrow();
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('rejects invalid resource options before fetch or worker creation', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const fetch = vi.fn();
    G.fetch = fetch;

    await expect(
      PptxPresentation.load('/presentation.pptx', { debug: 'yes' as never }),
    ).rejects.toThrow(/debug must be a boolean/);
    expect(fetch).not.toHaveBeenCalled();
    expect(SilentWorker.instances).toHaveLength(0);
  });

  // Wiring guard: destroy() must actually release the Google-Fonts substitutes
  // the deck preloaded into the shared FontFaceSet. The other tests set
  // `_googleFontFaces = []`, so they never exercise the unload branch — a dropped
  // call (or a wrong field name) would go unnoticed. Preload a real face through
  // core, hand it to the deck, then assert destroy() removes it and clears the array.
  it('destroy() releases the deck’s font-source faces from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const { pres } = makePresentation();
    Object.assign(pres, { _fonts: [googleFonts()], _embeddedFontAliases: new Map() });
    await (pres as unknown as {
      _ensureFontSources(names: string[], facts: undefined, scalars: undefined): Promise<void>;
    })._ensureFontSources(['Calibri'], undefined, undefined);
    expect(added).toHaveLength(1);
    pres.destroy();
    expect(added).toHaveLength(0);

  });

  it('never registers a source face for an installed authored family', async () => {
    const { added } = installFontFaceSet();
    for (const installed of [['arial'], []]) {
      const { pres } = makePresentation();
      const { PptxInstalledFamilies } = await import('./font-source-load.js');
      Object.assign(pres, {
        _fonts: [fontFiles({ faces: [{ family: 'Arial', url: 'https://app.example/arial.woff2' }] })],
        _embeddedFontAliases: new Map(),
        _installedFamilies: new PptxInstalledFamilies(async () => ({ installed: new Set(installed), checked: new Set(['arial']) })),
      });
      await (pres as unknown as {
        _ensureFontSources(names: string[], facts: undefined, scalars: undefined): Promise<void>;
      })._ensureFontSources(['Arial'], undefined, undefined);
      expect(added).toHaveLength(installed.length ? 0 : 1);
      pres.destroy();
    }
  });

  it('destroy() drops the font-source descriptors and their bytes', async () => {
    installFontFaceSet();
    const { pres } = makePresentation();
    const bytes = new ArrayBuffer(1 << 20);
    Object.assign(pres, { _fonts: [fontFiles({ faces: [{ family: 'Deck Face', data: bytes }] })], _embeddedFontAliases: new Map() });
    await (pres as unknown as {
      _ensureFontSources(names: string[], facts: undefined, scalars: undefined): Promise<void>;
    })._ensureFontSources(['Deck Face'], undefined, undefined);
    pres.destroy();
    // Nothing reachable from the destroyed presentation still holds the bytes.
    const reachable = (root: unknown, seen = new Set<unknown>()): boolean => {
      if (root === bytes) return true;
      if (root === null || typeof root !== 'object' || seen.has(root)) return false;
      seen.add(root);
      if (root instanceof Map) return [...root.entries()].some(([key, value]) => reachable(key, seen) || reachable(value, seen));
      if (root instanceof Set) return [...root].some((value) => reachable(value, seen));
      return Object.values(root).some((value) => reachable(value, seen));
    };
    expect(reachable(pres)).toBe(false);
  });

  it('destroy() releases the deck’s embedded fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await registerEmbeddedFonts([{
      family: 'Deck Sans',
      bytes: new Uint8Array([0, 1, 0, 0]),
      odttf: false,
      weight: 'normal',
      style: 'normal',
    }]);
    expect(added).toHaveLength(1);

    const { pres } = makePresentation();
    (pres as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces = held;
    pres.destroy();

    expect(added).toHaveLength(0);
    expect((pres as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces).toHaveLength(0);
  });
});


it.each(['ordinary', 'selected-source'] as const)('filters ordinary Window preload and releases its retained rules (%s)', async (variant) => {
  G.Worker = SilentWorker;
  G.location = { href: 'http://localhost/' };
  const { added } = installFontFaceSet();
  G.fetch = vi.fn(async (url: string) => ({ ok: true, status: 200, url, text: async () => `
    @font-face { font-family: Carlito; src: url(b.woff2); unicode-range: U+0042; }
    @font-face { font-family: Carlito; src: url(a.woff2); unicode-range: U+0041; }
  `, arrayBuffer: async () => new TextEncoder().encode(url).buffer }));
  const loads: string[] = [];
  G.FontFace = class {
    readonly source: string;
    // A fetched subset is registered from its bytes (the fake body names it).
    constructor(public family: string, source: string | ArrayBuffer) {
      this.source = typeof source === 'string' ? source : new TextDecoder().decode(source);
    }
    load() { loads.push(this.source); return Promise.resolve(this); }
  };
  vi.spyOn(PptxPresentation.prototype as unknown as {
    _parse(buffer: ArrayBuffer, resourcePolicy: object): Promise<void>;
  }, '_parse').mockImplementationOnce(async function (this: Presentation) {
    (this as unknown as { _preflight: object })._preflight = {
      slideCount: 0, slideWidth: 914400, slideHeight: 914400,
      defaultTextColor: null, majorFont: 'Calibri', minorFont: null,
      hlinkColor: null, folHlinkColor: null, embeddedFonts: [], slides: [],
      fontPreloadNames: ['Calibri'], fontPreloadDemand: [65],
    };
  });
  const source: ModelSource<'pptx'> = { target: 'pptx', claim: () => true, beginLoad: () => ({
    module: { protocol: 'ooxml-model-source-module/v2', target: 'pptx', moduleUrl: 'https://example.test/source.mjs', config: {} },
    release: () => undefined,
  }) };
  const presentation = await PptxPresentation.load(new ArrayBuffer(0), { fonts: [googleFonts()],
    ...(variant === 'selected-source' ? { legacyPpt: source } : {}),
  });
  // Only the subset the demanded scalars cover is fetched and registered.
  expect(added).toHaveLength(1);
  // The family presence probe loads local() sources first; of the source
  // faces only the subset the demanded scalars cover loads.
  expect(loads.filter((source) => !source.startsWith('local('))).toEqual([expect.stringContaining('a.woff2')]);
  presentation.destroy();
  expect(added).toHaveLength(0);
});
