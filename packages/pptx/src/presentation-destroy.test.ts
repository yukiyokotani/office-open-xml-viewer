import { describe, it, expect, vi, afterEach } from 'vitest';
import { WorkerBridge, preloadGoogleFonts, type WorkerLike, type FontPreloadEntry } from '@silurus/ooxml-core';
import { PptxPresentation } from './presentation';

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
});

const CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
interface FakeFace { family: string }
function installFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: string, public descriptors?: object) {}
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
    instance._mediaCache = new Map();
    instance._imageCache = new Map();
    instance._googleFontFaces = [];
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

  it('destroys a partially initialized presentation when load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    const parse = vi
      .spyOn(
        PptxPresentation.prototype as unknown as {
          _parse(
            buffer: ArrayBuffer,
            maxZipEntryBytes?: number,
            useGoogleFonts?: boolean,
            timeoutMs?: number,
          ): Promise<void>;
        },
        '_parse',
      )
      .mockRejectedValueOnce(failure);
    const destroy = vi.spyOn(PptxPresentation.prototype, 'destroy');

    await expect(PptxPresentation.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(destroy).toHaveBeenCalledTimes(1);

    parse.mockRestore();
    destroy.mockRestore();
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
});
