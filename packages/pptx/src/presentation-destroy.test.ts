import { describe, it, expect, afterEach, vi } from 'vitest';
import {
  WorkerBridge,
  preloadGoogleFonts,
  registerEmbeddedFonts,
  type WorkerLike,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { PptxPresentation } from './presentation';
import { loadEmbeddedFonts } from './embedded-fonts';
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
  G.fetch = async () => ({ ok: true, text: async () => CSS });
  delete G.self;
  return { added };
}
const MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

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

  it('terminates the owned worker when a partially initialized load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(
          buffer: ArrayBuffer,
          resourcePolicy: object,
          useGoogleFonts?: boolean,
          timeoutMs?: number,
        ): Promise<void>;
      },
      '_parse',
    ).mockRejectedValueOnce(failure);

    await expect(PptxPresentation.load(new ArrayBuffer(0))).rejects.toBe(failure);
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
          useGoogleFonts?: boolean,
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
    ).mockImplementationOnce(async function (this: PptxPresentation) {
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

  it('does not request remote fonts when useGoogleFonts is false', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    installFontFaceSet();
    const fetch = vi.fn(async () => ({ ok: true, text: async () => CSS }));
    G.fetch = fetch;
    vi.spyOn(
      PptxPresentation.prototype as unknown as {
        _parse(buffer: ArrayBuffer, resourcePolicy: object): Promise<void>;
      },
      '_parse',
    ).mockImplementationOnce(async function (this: PptxPresentation) {
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
      useGoogleFonts: false,
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
  it('destroy() releases the deck’s Google fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(1); // the web font is in the shared set

    const { pres } = makePresentation();
    (pres as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces = held;
    pres.destroy();

    expect(added).toHaveLength(0); // face left the set
    expect((pres as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces).toHaveLength(0);
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
