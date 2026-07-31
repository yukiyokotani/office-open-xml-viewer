import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  WorkerBridge,
  preloadGoogleFonts,
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  type WorkerLike,
  type FontPreloadEntry,
  type OffscreenFactory,
} from '@silurus/ooxml-core';
import { XlsxWorkbook } from './workbook.js';

/**
 * `XlsxWorkbook.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight so a
 * `load()` / image extraction awaiting the worker cannot hang after the
 * workbook is disposed. Pinned with a real {@link WorkerBridge} over an
 * in-memory worker (the constructor opens a real Worker, so we build
 * off-prototype and inject the collaborators destroy() touches — the pattern
 * from `workbook.image.test.ts`).
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

/** Flush pending microtasks so a drop's close-through-promise has run. */
const flush = () => new Promise((r) => setTimeout(r, 0));

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

describe('XlsxWorkbook.destroy() — rejects in-flight worker requests', () => {
  function makeWorkbook() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance.sheetCache = new Map();
    instance.imageCache = new Map();
    instance.imageBlobCache = new Map();
    instance.googleFontFaces = [];
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { wb: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { wb, bridge, worker } = makeWorkbook();
    const inFlight = bridge.request((id) => ({ id }));
    wb.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice', () => {
    const { wb } = makeWorkbook();
    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
  });

  it('destroys a partially initialized workbook when load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    const load = vi
      .spyOn(
        XlsxWorkbook.prototype as unknown as {
          _load(buffer: ArrayBuffer, opts: object): Promise<void>;
        },
        '_load',
      )
      .mockRejectedValueOnce(failure);
    const destroy = vi.spyOn(XlsxWorkbook.prototype, 'destroy');

    await expect(XlsxWorkbook.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(destroy).toHaveBeenCalledTimes(1);

    load.mockRestore();
    destroy.mockRestore();
  });

  // Wiring guard: destroy() must actually release the Google-Fonts substitutes
  // the workbook preloaded. The other tests set `googleFontFaces = []`, so they
  // never exercise the unload branch — a dropped call would go unnoticed.
  it('destroy() releases the workbook’s Google fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(1);

    const { wb } = makeWorkbook();
    (wb as unknown as { googleFontFaces: FontFace[] }).googleFontFaces = held;
    wb.destroy();

    expect(added).toHaveLength(0);
    expect((wb as unknown as { googleFontFaces: FontFace[] }).googleFontFaces).toHaveLength(0);
  });
});

/**
 * After #781 the decoded bitmaps are owned by the shared, per-`_fetchImage` core
 * caches (base raster via getCachedBitmapByPath, `<a:duotone>` recolour via
 * getCachedDuotoneBitmapByPath), NOT by the per-instance `imageCache` lookup map.
 * `destroy()` must therefore drop those shared caches — a bare `imageCache.clear()`
 * would drop only lookup references and leak the GPU backing until GC (which is
 * not guaranteed to run promptly for GPU-backed objects). This is the same
 * teardown discipline #779 fixed, expressed through the shared cache the way
 * docx/pptx do. Pins the wiring: after a decode has landed in the shared caches
 * keyed by the instance's `_fetchImage`, destroy() closes those bitmaps.
 */
describe('XlsxWorkbook.destroy() — drops the shared image caches (GPU-leak guard)', () => {
  afterEach(() => vi.unstubAllGlobals());

  function makeWorkbook(fetchImage: (path: string, mime: string) => Promise<Blob>) {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, { correlate: (r) => r.id });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    instance.sheetCache = new Map();
    instance.imageCache = new Map();
    instance.imageBlobCache = new Map();
    instance.googleFontFaces = [];
    instance._fetchImage = fetchImage;
    return instance as unknown as DestroyProbe & { imageCache: Map<string, unknown> };
  }

  /** An offscreen surface for the duotone pixel pass in node (no OffscreenCanvas). */
  function recordingFactory(): OffscreenFactory {
    return ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext: () => ({
        drawImage() {},
        getImageData(_x: number, _y: number, sw: number, sh: number) {
          const data = new Uint8ClampedArray(sw * sh * 4).fill(246);
          for (let i = 3; i < data.length; i += 4) data[i] = 255;
          return { data, width: sw, height: sh } as unknown as ImageData;
        },
        putImageData() {},
      }),
    })) as unknown as OffscreenFactory;
  }

  it('closes a base ImageBitmap decoded into the shared cache and empties the lookup map', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const wb = makeWorkbook(fetchImage);
    // Warm the shared cache the same way prefetchImages does — keyed by _fetchImage.
    await getCachedBitmapByPath('xl/media/image1.png', 'image/png', fetchImage);
    wb.imageCache.set('xl/media/image1.png', {});

    wb.destroy();
    await flush(); // the drop closes through the settled promise (a microtask)

    expect(close).toHaveBeenCalledTimes(1); // dropBitmapCacheByPath closed it
    expect(wb.imageCache.size).toBe(0);
  });

  it('closes a duotone recolour decoded into the shared duotone cache', async () => {
    const baseClose = vi.fn();
    const duoClose = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (src: unknown) =>
        (src instanceof Blob
          ? { width: 4, height: 4, close: baseClose }
          : { width: 4, height: 4, close: duoClose }) as unknown as ImageBitmap,
      ),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const wb = makeWorkbook(fetchImage);
    await getCachedDuotoneBitmapByPath(
      'xl/media/image1.png',
      'image/png',
      { clr1: '000000', clr2: 'FFF3F4' },
      fetchImage,
      { offscreenFactory: recordingFactory() },
    );

    wb.destroy();
    await flush();

    // dropDuotoneBitmapCache closed the recolour; dropBitmapCacheByPath the base.
    expect(duoClose).toHaveBeenCalledTimes(1);
    expect(baseClose).toHaveBeenCalledTimes(1);
  });

  it('is safe to destroy() twice (dropping an already-dropped shared cache is a no-op)', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const wb = makeWorkbook(fetchImage);
    await getCachedBitmapByPath('xl/media/image1.png', 'image/png', fetchImage);

    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
    await flush();
    // The shared cache was forgotten on the first destroy(), so the second pass
    // has nothing to close — close() runs exactly once total.
    expect(close).toHaveBeenCalledTimes(1);
  });
});
