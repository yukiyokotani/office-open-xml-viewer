import { describe, it, expect, vi, afterEach } from 'vitest';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import {
  WorkerBridge,
  FontProviderSession,
  GoogleFontsProvider,
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  type WorkerLike,
  type OffscreenFactory,
} from '@silurus/ooxml-core';
import { XlsxWorkbook, retainXlsxViewerFonts } from './workbook.js';

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
  SilentWorker.instances = [];
  vi.restoreAllMocks();
});

const CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
const CALADEA_CSS = `@font-face { font-family: 'Caladea'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/caladea/y.woff2) format('woff2'); }`;
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
  G.fetch = async (input: RequestInfo | URL) => String(input).includes('fonts.googleapis')
    ? new Response(CSS)
    : new Response(new Uint8Array([1, 2, 3]));
  delete G.self;
  return { added };
}
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
    instance.sheetLoads = new Map();
    instance.rawParts = new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 });
    instance.fontsDestroyed = false;
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

  it('terminates the owned worker when a partially initialized load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      XlsxWorkbook.prototype as unknown as {
        _load(buffer: ArrayBuffer, opts: object): Promise<void>;
      },
      '_load',
    ).mockRejectedValueOnce(failure);

    await expect(XlsxWorkbook.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('preserves the load error and terminates directly when destroy throws', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      XlsxWorkbook.prototype as unknown as {
        _load(buffer: ArrayBuffer, opts: object): Promise<void>;
      },
      '_load',
    ).mockRejectedValueOnce(failure);
    vi.spyOn(XlsxWorkbook.prototype, 'destroy').mockImplementationOnce(() => {
      throw new Error('cleanup failure');
    });

    await expect(XlsxWorkbook.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('terminates directly when construction fails before the factory owns an instance', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'not a valid base URL' };

    await expect(
      XlsxWorkbook.load(new ArrayBuffer(0), { wasmUrl: 'relative.wasm' }),
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
      XlsxWorkbook.load('/workbook.xlsx', {
        resourceLimits: { maxTotalInflatedBytes: Number.POSITIVE_INFINITY },
      }),
    ).rejects.toThrow(/resourceLimits\.maxTotalInflatedBytes/);
    expect(fetch).not.toHaveBeenCalled();
    expect(SilentWorker.instances).toHaveLength(0);
  });

  // Google fonts use the same session-owned cleanup as private providers.
  it('destroy() releases the workbook’s Google fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const session = new FontProviderSession(new GoogleFontsProvider());
    await session.ensure(['Calibri']);
    expect(added).toHaveLength(1);

    const { wb } = makeWorkbook();
    (wb as unknown as { fontSession: FontProviderSession }).fontSession = session;
    wb.destroy();

    expect(added).toHaveLength(0);
  });

  it('releases a viewer-owned FontFaceSet before workbook destruction', async () => {
    const { added } = installFontFaceSet();
    const { wb } = makeWorkbook();
    (wb as unknown as { fontSession: FontProviderSession }).fontSession =
      new FontProviderSession(new GoogleFontsProvider());
    (wb as unknown as { providerFontNames: string[] }).providerFontNames = ['Calibri'];

    const release = await (wb as unknown as XlsxWorkbook)[retainXlsxViewerFonts](
      G.document as Document,
    );
    expect(added).toHaveLength(1);

    release();
    expect(added).toHaveLength(0);
    wb.destroy();
  });

  it('an in-flight destroy releases once and preserves another workbook holder', async () => {
    const { added } = installFontFaceSet();
    let resolveFetch!: (response: Response) => void;
    G.fetch = vi.fn((input: RequestInfo | URL) => String(input).includes('fonts.googleapis')
      ? new Promise<Response>((resolve) => { resolveFetch = resolve; })
      : Promise.resolve(new Response(new Uint8Array([1, 2, 3]))));
    const first = makeWorkbook().wb as unknown as XlsxWorkbook;
    const second = makeWorkbook().wb as unknown as XlsxWorkbook;
    (first as unknown as { fontSession: FontProviderSession }).fontSession =
      new FontProviderSession(new GoogleFontsProvider());
    (second as unknown as { fontSession: FontProviderSession }).fontSession =
      new FontProviderSession(new GoogleFontsProvider());
    (first as unknown as { providerFontNames: string[] }).providerFontNames = ['Cambria'];
    (second as unknown as { providerFontNames: string[] }).providerFontNames = ['Cambria'];
    const targetDocument = G.document as Document;

    const firstRetain = first[retainXlsxViewerFonts](targetDocument);
    const secondRetain = second[retainXlsxViewerFonts](targetDocument);
    await Promise.resolve();
    first.destroy();
    resolveFetch(new Response(CALADEA_CSS));
    const [releaseFirst] = await Promise.all([firstRetain, secondRetain]);
    releaseFirst(); // stale viewer completion after workbook teardown: no-op

    expect(added).toHaveLength(1);
    second.destroy();
    expect(added).toHaveLength(0);
  });
});

/**
 * After #781 the decoded bitmaps are owned by the shared, per-`_fetchImage` core
 * decoded owner (base raster plus `<a:duotone>` derivatives), not by a
 * workbook-lifetime lookup map. Viewport lookup is frame-local. `destroy()` must
 * drop that shared owner so GPU-backed objects are released promptly. This is
 * the same teardown discipline #779 fixed, expressed through the shared cache
 * the way docx/pptx do.
 */
describe('XlsxWorkbook.destroy() — drops the shared image caches (GPU-leak guard)', () => {
  afterEach(() => vi.unstubAllGlobals());

  function makeWorkbook(fetchImage: (path: string, mime: string) => Promise<Blob>) {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, { correlate: (r) => r.id });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    instance.sheetCache = new Map();
    instance.sheetLoads = new Map();
    instance.rawParts = new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 });
    instance.fontsDestroyed = false;
    instance._fetchImage = fetchImage;
    return instance as unknown as DestroyProbe;
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

  it('closes a base ImageBitmap decoded into the shared owner', async () => {
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
    wb.destroy();
    await flush(); // the drop closes through the settled promise (a microtask)

    expect(close).toHaveBeenCalledTimes(1); // dropBitmapCacheByPath closed it
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
