import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  peekCachedBitmapByPath,
  resolvedCachedBitmapVariantKey,
  dropBitmapCacheByPath,
  dropCachedDerivedBitmapNamespace,
  acquireBitmapCacheLease,
  inspectCachedRasterSource,
  releaseOwnedBitmap,
  withBitmapCacheLease,
} from './bitmap-image-by-path';
import {
  HARD_MAX_DECODED_IMAGE_BYTES,
  MAX_CONCURRENT_IMAGE_DECODES,
  MAX_DECODED_IMAGE_BYTES,
} from './pixel-budget';

/** Build a minimal standard (non-placeable) WMF that draws one polyline, so the
 *  shared player produces non-empty geometry (→ a non-null bitmap). Mirrors the
 *  byte layout exercised in wmf.test.ts. */
function buildMinimalWmf(): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff);
  const i16 = (v: number) => u16(v & 0xffff);
  const u32 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  // 18-byte standard header (type=1, headerSize=9 words).
  u16(1); u16(9); u16(0x0300); u32(0); u16(8); u32(0); u16(0);
  const rec = (fn: number, params: number[]) => { u32(3 + params.length); u16(fn); for (const p of params) i16(p); };
  rec(0x020b, [0, 0]);                       // SETWINDOWORG (y,x)
  rec(0x020c, [100, 100]);                   // SETWINDOWEXT (y,x)
  rec(0x02fa, [0, 1, 0, 0, 0]);              // CREATEPENINDIRECT
  rec(0x012d, [0]);                          // SELECTOBJECT idx 0
  rec(0x0325, [2, 0, 0, 50, 50]);            // POLYLINE 2 pts (0,0)-(50,50)
  u32(3); u16(0x0000);                       // EOF
  return new Uint8Array(b);
}

function stubMetafileSurfaceCreation(): {
  create: ReturnType<typeof vi.fn>;
  closes: Array<ReturnType<typeof vi.fn>>;
} {
  vi.stubGlobal(
    'OffscreenCanvas',
    class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() {
        return {
          fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
          lineJoin: 'miter', lineCap: 'butt',
          save() {}, restore() {}, beginPath() {}, closePath() {},
          moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
        };
      }
    },
  );
  const closes: Array<ReturnType<typeof vi.fn>> = [];
  const create = vi.fn(async (source: { width: number; height: number }) => {
    const close = vi.fn();
    closes.push(close);
    return {
      width: source.width,
      height: source.height,
      close,
    } as unknown as ImageBitmap;
  });
  vi.stubGlobal('createImageBitmap', create);
  return { create, closes };
}

/** Coded 400×100 JPEG whose EXIF orientation 6 gives a 100×400 natural grid. */
function buildOrientedJpeg(): Uint8Array {
  const exif = new Uint8Array(32);
  exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00, 0x49, 0x49], 0);
  const exifView = new DataView(exif.buffer);
  exifView.setUint16(8, 42, true);
  exifView.setUint32(10, 8, true);
  exifView.setUint16(14, 1, true);
  exifView.setUint16(16, 0x0112, true);
  exifView.setUint16(18, 3, true);
  exifView.setUint32(20, 1, true);
  exifView.setUint16(24, 6, true);
  const bytes = new Uint8Array(2 + 4 + exif.length + 9 + 2);
  let offset = 0;
  bytes.set([0xff, 0xd8, 0xff, 0xe1, 0x00, exif.length + 2], offset); offset += 6;
  bytes.set(exif, offset); offset += exif.length;
  bytes.set([0xff, 0xc0, 0x00, 0x07, 0x08, 0x00, 0x64, 0x01, 0x90], offset); offset += 9;
  bytes.set([0xff, 0xda], offset);
  return bytes;
}

/**
 * The decoded-bitmap cache (sibling of getCachedSvgImageByPath) pulls bytes via
 * the injected `fetchImage(path, mime)` and caches by zip path plus any raster
 * resolution band, namespaced per document by the closure identity.
 */
describe('getCachedBitmapByPath', () => {
  beforeEach(() => {
    // `createImageBitmap` doesn't exist in the node test env; stub it to a
    // sentinel with the .close() the LRU eviction / drop calls.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('decodes by path via fetchImage and caches across draws (single fetch, single decode)', async () => {
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }));
    // Unique path so the module-level LRU isn't pre-warmed by another test.
    const path = 'word/media/cachehit-a.png';

    const first = await getCachedBitmapByPath(path, 'image/png', fetchImage);
    const second = await getCachedBitmapByPath(path, 'image/png', fetchImage);

    expect(first).toBe(second);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(path, 'image/png');
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).toHaveBeenCalledTimes(1);
  });

  it('reuses the smallest cached resolution that fully covers a later request', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 12_000);
    view.setUint32(20, 9_000);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 12_000,
      height: options?.resizeHeight ?? 9_000,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);

    const path = 'ppt/media/poster-resolution-bands.png';
    const a = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1010,
      targetHeightPx: 758,
    });
    const covered = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1000,
      targetHeightPx: 750,
    });
    const larger = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1400,
      targetHeightPx: 1050,
    });

    expect(a).toBe(covered);
    expect(larger).not.toBe(a);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(cib).toHaveBeenCalledTimes(2);
  });

  it('ranks reusable variants by retained pixels rather than target-key area', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 300);
    view.setUint32(20, 300);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 300,
      height: options?.resizeHeight ?? 300,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const path = 'word/media/cross-aspect-resolution-bands.png';

    const square = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 100,
      targetHeightPx: 100,
    });
    const wideRequest = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 200,
      targetHeightPx: 40,
    });
    const smallestCover = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 80,
      targetHeightPx: 30,
    });

    expect(square).toMatchObject({ width: 100, height: 100 });
    expect(wideRequest).toMatchObject({ width: 200, height: 40 });
    expect(smallestCover).toBe(wideRequest);
    expect(cib).toHaveBeenCalledTimes(2);
  });

  it('keeps sufficient display variants and one native entry for an in-budget raster', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 1920);
    view.setUint32(20, 1080);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 1920,
      height: options?.resizeHeight ?? 1080,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);

    const path = 'ppt/media/native-bands.png';
    const first = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 640,
      targetHeightPx: 360,
    });
    const covered = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 320,
      targetHeightPx: 180,
    });
    const zoomed = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1600,
      targetHeightPx: 900,
    });
    const atNative = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1920,
      targetHeightPx: 1080,
    });
    const native = await getCachedBitmapByPath(path, 'image/png', fetchImage);

    expect(covered).toBe(first);
    expect(zoomed).not.toBe(first);
    expect(native).not.toBe(zoomed);
    expect(atNative).toBe(native);
    expect(fetchImage).toHaveBeenCalledTimes(3);
    expect(cib).toHaveBeenCalledTimes(3);
    expect(cib).toHaveBeenNthCalledWith(1, expect.any(Blob), {
      resizeWidth: 640,
      resizeHeight: 360,
      resizeQuality: 'high',
    });
    expect(cib).toHaveBeenNthCalledWith(2, expect.any(Blob), {
      resizeWidth: 1600,
      resizeHeight: 900,
      resizeQuality: 'high',
    });
    expect(cib).toHaveBeenNthCalledWith(3, expect.any(Blob));
  });

  it('keys JPEG variants from the browser-oriented EXIF dimensions', async () => {
    const jpeg = buildOrientedJpeg();
    const fetchImage = vi.fn(async () => new Blob([jpeg as BlobPart], { type: 'image/jpeg' }));
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      const width = options?.resizeWidth ?? 100;
      const height = options?.resizeHeight ?? 400;
      return { width, height, close() {} } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', cib);
    const path = 'word/media/exif-oriented-bands.jpg';

    const reduced = await getCachedBitmapByPath(path, 'image/jpeg', fetchImage, {
      targetWidthPx: 50,
      targetHeightPx: 200,
    });
    const crossAspect = await getCachedBitmapByPath(path, 'image/jpeg', fetchImage, {
      targetWidthPx: 200,
      targetHeightPx: 50,
    });

    expect(reduced).toMatchObject({ width: 50, height: 200 });
    expect(crossAspect).toMatchObject({ width: 100, height: 50 });
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(cib).toHaveBeenNthCalledWith(1, expect.any(Blob), {
      resizeWidth: 50,
      resizeHeight: 200,
      resizeQuality: 'high',
    });
    expect(cib).toHaveBeenNthCalledWith(2, expect.any(Blob), {
      resizeWidth: 100,
      resizeHeight: 50,
      resizeQuality: 'high',
    });
  });

  it('keeps TIFF display-resolution variants across target changes', async () => {
    const bytes = new Uint8Array(38);
    const view = new DataView(bytes.buffer);
    bytes.set([0x49, 0x49], 0);
    view.setUint16(2, 42, true);
    view.setUint32(4, 8, true);
    view.setUint16(8, 2, true);
    view.setUint16(10, 256, true);
    view.setUint16(12, 4, true);
    view.setUint32(14, 1, true);
    view.setUint32(18, 4_249, true);
    view.setUint16(22, 257, true);
    view.setUint16(24, 4, true);
    view.setUint32(26, 1, true);
    view.setUint32(30, 6_137, true);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const render = vi.fn(async (
      _input: Uint8Array,
      options?: Readonly<{ targetWidthPx?: number; targetHeightPx?: number }>,
    ) => ({
      width: options?.targetWidthPx ?? 4_249,
      height: options?.targetHeightPx ?? 6_137,
      close() {},
    }) as unknown as ImageBitmap);
    const tiff = { render };
    const path = 'word/media/scan-target-bands.tiff';

    const first = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff,
      targetWidthPx: 320,
      targetHeightPx: 463,
    });
    const covered = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff,
      targetWidthPx: 160,
      targetHeightPx: 232,
    });
    const larger = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff,
      targetWidthPx: 640,
      targetHeightPx: 925,
    });

    expect(covered).toBe(first);
    expect(larger).not.toBe(first);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(render).toHaveBeenCalledTimes(2);
    expect(render).toHaveBeenNthCalledWith(1, bytes, expect.objectContaining({
      targetWidthPx: 320,
      targetHeightPx: 463,
    }));
    expect(render).toHaveBeenNthCalledWith(2, bytes, expect.objectContaining({
      targetWidthPx: 640,
      targetHeightPx: 925,
    }));
  });

  it('does not cache a one-axis TIFF downsample as the native bitmap', async () => {
    const bytes = new Uint8Array(38);
    const view = new DataView(bytes.buffer);
    bytes.set([0x49, 0x49], 0);
    view.setUint16(2, 42, true);
    view.setUint32(4, 8, true);
    view.setUint16(8, 2, true);
    view.setUint16(10, 256, true);
    view.setUint16(12, 4, true);
    view.setUint32(14, 1, true);
    view.setUint32(18, 4_000, true);
    view.setUint16(22, 257, true);
    view.setUint16(24, 4, true);
    view.setUint32(26, 1, true);
    view.setUint32(30, 3_000, true);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const render = vi.fn(async (
      _input: Uint8Array,
      options?: Readonly<{ targetWidthPx?: number; targetHeightPx?: number }>,
    ) => {
      const scale = options?.targetWidthPx ? Math.min(1, options.targetWidthPx / 4_000) : 1;
      return {
        width: Math.ceil(4_000 * scale),
        height: Math.ceil(3_000 * scale),
        close() {},
      } as unknown as ImageBitmap;
    });
    const path = 'word/media/one-axis-target.tiff';

    const reduced = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff: { render },
      targetWidthPx: 400,
    });
    const native = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff: { render },
    });

    expect(reduced).toMatchObject({ width: 400, height: 300 });
    expect(native).toMatchObject({ width: 4_000, height: 3_000 });
    expect(native).not.toBe(reduced);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(render).toHaveBeenCalledTimes(2);
  });

  it('prefers the smallest retained TIFF surface across one- and two-axis keys', async () => {
    const bytes = new Uint8Array(38);
    const view = new DataView(bytes.buffer);
    bytes.set([0x49, 0x49], 0);
    view.setUint16(2, 42, true);
    view.setUint32(4, 8, true);
    view.setUint16(8, 2, true);
    view.setUint16(10, 256, true);
    view.setUint16(12, 4, true);
    view.setUint32(14, 1, true);
    view.setUint32(18, 10_000, true);
    view.setUint16(22, 257, true);
    view.setUint16(24, 4, true);
    view.setUint32(26, 1, true);
    view.setUint32(30, 10_000, true);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const render = vi.fn(async (
      _input: Uint8Array,
      options?: Readonly<{ targetWidthPx?: number; targetHeightPx?: number }>,
    ) => {
      const scale = Math.min(1, Math.max(
        (options?.targetWidthPx ?? 0) / 10_000,
        (options?.targetHeightPx ?? 0) / 10_000,
      ));
      const size = Math.max(1, Math.ceil(10_000 * scale));
      return { width: size, height: size, close() {} } as unknown as ImageBitmap;
    });
    const path = 'word/media/one-axis-variant-ranking.tiff';

    const wide = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff: { render },
      targetWidthPx: 2_828,
    });
    const square = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff: { render },
      targetWidthPx: 1_000,
      targetHeightPx: 1_000,
    });
    const smallestCover = await getCachedBitmapByPath(path, 'image/tiff', fetchImage, {
      tiff: { render },
      targetWidthPx: 500,
    });

    expect(wide).toMatchObject({ width: 2_828, height: 2_828 });
    expect(square).toMatchObject({ width: 1_000, height: 1_000 });
    expect(smallestCover).toBe(square);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(render).toHaveBeenCalledTimes(2);
  });

  it('does not let a permissive native entry bypass a later retained-pixel limit', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 4000);
    view.setUint32(20, 3000);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const close = vi.fn();
    const native = { width: 4000, height: 3000, close } as unknown as ImageBitmap;
    const cib = vi.fn(async () => native);
    vi.stubGlobal('createImageBitmap', cib);
    const path = 'word/media/native-restricted.png';

    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage)).resolves.toBe(native);
    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage, {
      maxRetainedPixels: 1 << 23,
    })).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 1 << 23,
      observed: 4000 * 3000,
    });
    expect(fetchImage).toHaveBeenCalledOnce();
    expect(cib).toHaveBeenCalledOnce();
    expect(close).not.toHaveBeenCalled();

    dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    expect(close).toHaveBeenCalledOnce();
  });

  it('namespaces the cache by fetchImage — two documents sharing a zip path decode independently', async () => {
    // Different files reuse the SAME internal paths (…/media/image1.png). Opening
    // document B after document A must NOT paint A's bytes for B: the cache is
    // scoped per byte source, not by path alone.
    const path = 'word/media/image1.png';
    const fetchA = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const fetchB = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([2])], { type: mime }));
    await getCachedBitmapByPath(path, 'image/png', fetchA);
    await getCachedBitmapByPath(path, 'image/png', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1);
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).toHaveBeenCalledTimes(2);
    // Within one document the path still dedupes across draws.
    await getCachedBitmapByPath(path, 'image/png', fetchA);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('peek returns undefined until the decode resolves, then the warmed bitmap (sync bullet contract)', async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => { release = r; });
    const fetchImage = vi.fn(async (_p: string, mime: string) => {
      await gate; // hold the fetch open so the decode hasn't settled yet
      return new Blob([new Uint8Array([7])], { type: mime });
    });
    const path = 'word/media/peek.png';
    const p = getCachedBitmapByPath(path, 'image/png', fetchImage);
    // Not warmed yet → the synchronous peek must see nothing (bullet skips).
    expect(peekCachedBitmapByPath(path, fetchImage)).toBeUndefined();
    release();
    await p;
    // After the warm pass awaited the decode, the peek sees the settled bitmap.
    const bmp = peekCachedBitmapByPath(path, fetchImage);
    expect(bmp).not.toBeUndefined();
    expect(bmp).not.toBeNull();
  });

  it('a WMF blip rasterizes through the shared player (opts.widthPt/heightPt size the raster)', async () => {
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) { this.width = w; this.height = h; }
        getContext() {
          return {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
          };
        }
      },
    );
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap));

    const wmf = buildMinimalWmf();
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([wmf as BlobPart], { type: 'image/wmf' }));

    const bmp = await getCachedBitmapByPath('word/media/wmf.wmf', 'image/wmf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(200); // wmfRasterTarget(100,100) → 200×200
  });

  it('content-sniffs a generically labeled WMF and keeps small and large frame variants', async () => {
    const { create, closes } = stubMetafileSurfaceCreation();
    const wmf = buildMinimalWmf();
    const fetchImage = vi.fn(async () =>
      new Blob([wmf as BlobPart], { type: 'application/octet-stream' }));
    const path = 'word/media/wmf-small-then-large.wmf';

    const small = await getCachedBitmapByPath(path, 'application/octet-stream', fetchImage, {
      widthPt: 50,
      heightPt: 25,
    });
    const large = await getCachedBitmapByPath(path, 'application/octet-stream', fetchImage, {
      widthPt: 100,
      heightPt: 50,
    });
    const smallAgain = await getCachedBitmapByPath(path, 'application/octet-stream', fetchImage, {
      widthPt: 50,
      heightPt: 25,
    });

    expect(small).toMatchObject({ width: 100, height: 50 });
    expect(large).toMatchObject({ width: 200, height: 100 });
    expect(large).not.toBe(small);
    expect(smallAgain).toBe(small);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(create).toHaveBeenCalledTimes(2);

    // Different point values that normalize to the same player grid share the
    // collision-safe raster variant instead of retaining a duplicate surface.
    await expect(getCachedBitmapByPath(path, 'application/octet-stream', fetchImage, {
      widthPt: 50.24,
      heightPt: 25.24,
    })).resolves.toBe(small);
    expect(fetchImage).toHaveBeenCalledTimes(2);

    dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    await Promise.resolve();
    expect(closes).toHaveLength(2);
    expect(closes.every((close) => close.mock.calls.length === 1)).toBe(true);
  });

  it('uses ordinary raster variants when PNG bytes are mislabeled as WMF', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 400);
    view.setUint32(20, 200);
    const fetchImage = vi.fn(async () =>
      new Blob([png as BlobPart], { type: 'image/wmf' }));
    const create = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 400,
      height: options?.resizeHeight ?? 200,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', create);
    const path = 'word/media/png-labeled-wmf.wmf';

    const reduced = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 50,
      heightPt: 25,
      targetWidthPx: 100,
      targetHeightPx: 50,
    });
    const sameRasterTarget = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 500,
      heightPt: 500,
      targetWidthPx: 100,
      targetHeightPx: 50,
      suppressBoundaryFrame: true,
    });
    const native = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 10,
      heightPt: 10,
    });

    expect(reduced).toMatchObject({ width: 100, height: 50 });
    expect(sameRasterTarget).toBe(reduced);
    expect(native).toMatchObject({ width: 400, height: 200 });
    expect(native).not.toBe(reduced);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(create).toHaveBeenCalledTimes(2);
    expect(create).toHaveBeenNthCalledWith(1, expect.any(Blob), {
      resizeWidth: 100,
      resizeHeight: 50,
      resizeQuality: 'high',
    });
    expect(create).toHaveBeenNthCalledWith(2, expect.any(Blob));
  });

  it('reuses a larger WMF raster for a smaller frame but isolates suppression variants', async () => {
    const { create, closes } = stubMetafileSurfaceCreation();
    const wmf = buildMinimalWmf();
    const fetchImage = vi.fn(async () =>
      new Blob([wmf as BlobPart], { type: 'image/wmf' }));
    const path = 'word/media/wmf-large-then-small.wmf';

    const large = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 150,
      heightPt: 100,
    });
    const coveredSmall = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 50,
      heightPt: 40,
    });
    const suppressedSmall = await getCachedBitmapByPath(path, 'image/wmf', fetchImage, {
      widthPt: 50,
      heightPt: 40,
      suppressBoundaryFrame: true,
    });

    expect(large).toMatchObject({ width: 300, height: 200 });
    expect(coveredSmall).toBe(large);
    expect(suppressedSmall).toMatchObject({ width: 100, height: 80 });
    expect(suppressedSmall).not.toBe(large);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(create).toHaveBeenCalledTimes(2);
    const [regularKey, suppressedKey] = await Promise.all([
      resolvedCachedBitmapVariantKey(
        path,
        'image/wmf',
        fetchImage,
        { widthPt: 150, heightPt: 100 },
        undefined,
        large as ImageBitmap,
      ),
      resolvedCachedBitmapVariantKey(
        path,
        'image/wmf',
        fetchImage,
        { widthPt: 50, heightPt: 40, suppressBoundaryFrame: true },
        undefined,
        suppressedSmall as ImageBitmap,
      ),
    ]);
    expect(suppressedKey).not.toBe(regularKey);

    dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    await Promise.resolve();
    expect(closes).toHaveLength(2);
    expect(closes.every((close) => close.mock.calls.length === 1)).toBe(true);
  });

  it('a true EMF blip is cached as null (skipped, not crashed)', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () =>
      ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap));
    // ENHMETAHEADER: u32@0=1 (EMR_HEADER), u32@40=0x464D4520 (" EMF").
    const emf = new Uint8Array(48);
    const dv = new DataView(emf.buffer);
    dv.setUint32(0, 1, true);
    dv.setUint32(40, 0x464d4520, true);
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([emf as BlobPart], { type: 'image/emf' }));

    const bmp = await getCachedBitmapByPath('word/media/diagram.emf', 'image/emf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).not.toHaveBeenCalled();
    // The null is cached — a second draw does not re-fetch/re-sniff.
    await getCachedBitmapByPath('word/media/diagram.emf', 'image/emf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(fetchImage).toHaveBeenCalledTimes(1);
  });

  it('self-evicts on a failed decode (no poisoned cache) and retries on the next call', async () => {
    let calls = 0;
    const fetchImage = vi.fn(async (_p: string, _m: string) => {
      calls++;
      throw new Error('byte source unavailable');
    });
    const path = 'word/media/fail.png';
    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage)).rejects.toThrow();
    // Second call must RETRY (cache self-evicted), not return a cached rejection.
    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage)).rejects.toThrow();
    expect(calls).toBe(2);
    // The failed entry left nothing warm for the sync peek.
    expect(peekCachedBitmapByPath(path, fetchImage)).toBeUndefined();
  });

  it('evicts the LRU-oldest past the cap and closes its GPU backing', async () => {
    const closed: string[] = [];
    // Each decode returns a bitmap tagged with the blob's first byte so we can
    // see which one gets closed.
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => {
      const tag = new Uint8Array(await blob.arrayBuffer())[0];
      return { width: 1, height: 1, close: () => closed.push(`b${tag}`) } as unknown as ImageBitmap;
    }));
    // A dedicated fetchImage → dedicated (empty) cache, so the 256 cap is reached
    // deterministically by this test alone.
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    // Fill to the cap (256 distinct paths), then one more to force one eviction.
    for (let i = 0; i < 256; i++) {
      await getCachedBitmapByPath(`word/media/lru-${i}.png`, 'image/png', fetchImage);
    }
    expect(closed.length).toBe(0); // nothing evicted yet at the cap
    await getCachedBitmapByPath('word/media/lru-256.png', 'image/png', fetchImage);
    // Let the eviction close-through-promise microtask run.
    await Promise.resolve();
    expect(closed.length).toBe(1); // the oldest (lru-0) was closed
  });

  it('dropBitmapCacheByPath closes a document\'s bitmaps and lets it re-decode', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([3])], { type: mime }));
    await getCachedBitmapByPath('word/media/drop-a.png', 'image/png', fetchImage);
    await getCachedBitmapByPath('word/media/drop-b.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage); // e.g. on Document.destroy()
    await Promise.resolve();
    expect(closes.length).toBe(2);
    await getCachedBitmapByPath('word/media/drop-a.png', 'image/png', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });

  it('keeps compatibility namespace teardown a no-op without an initialized owner', () => {
    expect(() => dropCachedDerivedBitmapNamespace(
      undefined as unknown as object,
      'legacy-effect',
    )).not.toThrow();
    expect(() => dropCachedDerivedBitmapNamespace(
      null as unknown as object,
      'legacy-effect',
    )).not.toThrow();
  });

  it('closes an in-flight decode exactly once when it completes after teardown', async () => {
    let finishDecode!: (bitmap: ImageBitmap) => void;
    const close = vi.fn();
    const bitmap = { width: 1, height: 1, close } as unknown as ImageBitmap;
    const cib = vi.fn(() => new Promise<ImageBitmap>((resolve) => { finishDecode = resolve; }));
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([3])], { type: mime }));

    const pending = getCachedBitmapByPath('word/media/inflight-drop.png', 'image/png', fetchImage);
    await vi.waitFor(() => expect(cib).toHaveBeenCalledOnce());
    dropBitmapCacheByPath(fetchImage);
    finishDecode(bitmap);

    await expect(pending).resolves.toBe(bitmap);
    await Promise.resolve();
    await Promise.resolve();
    expect(close).toHaveBeenCalledOnce();
  });

  it('does not recreate a cache when teardown wins a pending profile inspection', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 1920);
    view.setUint32(20, 1080);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    let finishFirstFetch!: (blob: Blob) => void;
    let fetchCount = 0;
    const fetchImage = vi.fn((_path: string, _mime: string) => {
      fetchCount++;
      return fetchCount === 1
        ? new Promise<Blob>((resolve) => { finishFirstFetch = resolve; })
        : Promise.resolve(blob);
    });
    const cib = vi.fn(async () => ({
      width: 960,
      height: 540,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const path = 'word/media/profile-drop-race.png';
    const options = { targetWidthPx: 960, targetHeightPx: 540 };

    const pending = getCachedBitmapByPath(path, 'image/png', fetchImage, options);
    await vi.waitFor(() => expect(fetchImage).toHaveBeenCalledOnce());
    const rejected = expect(pending).rejects.toThrow(/cache.*dropped/i);
    dropBitmapCacheByPath(fetchImage);
    finishFirstFetch(blob);
    await rejected;
    expect(cib).not.toHaveBeenCalled();

    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage, options))
      .resolves.toMatchObject({ width: 960, height: 540 });
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(cib).toHaveBeenCalledOnce();
  });

  it('rejects and closes an oversized decode when its header was not recognized', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 8192, height: 8192, close }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1, 2, 3])], { type: mime }));

    await expect(getCachedBitmapByPath('word/media/unknown.bin', 'application/octet-stream', fetchImage))
      .rejects.toMatchObject({
        code: 'ooxml-decoded-image-limit',
        metric: 'image-pixels',
        observed: 8192 * 8192,
      });
    expect(close).toHaveBeenCalledOnce();
    expect(peekCachedBitmapByPath('word/media/unknown.bin', fetchImage)).toBeUndefined();
  });

  it('admits at most the shared number of concurrent decodes per document', async () => {
    let active = 0;
    let maximumActive = 0;
    const releases: Array<() => void> = [];
    vi.stubGlobal('createImageBitmap', vi.fn(() => new Promise<ImageBitmap>((resolve) => {
      active++;
      maximumActive = Math.max(maximumActive, active);
      releases.push(() => {
        active--;
        resolve({ width: 1, height: 1, close() {} } as unknown as ImageBitmap);
      });
    })));
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));

    const decodes = Array.from({ length: MAX_CONCURRENT_IMAGE_DECODES + 2 }, (_, index) =>
      getCachedBitmapByPath(`word/media/concurrent-${index}.png`, 'image/png', fetchImage));
    await vi.waitFor(() => expect(active).toBe(MAX_CONCURRENT_IMAGE_DECODES));

    for (const release of releases.splice(0)) release();
    await vi.waitFor(() => expect(releases).toHaveLength(2));
    for (const release of releases.splice(0)) release();
    await Promise.all(decodes);
    expect(maximumActive).toBe(MAX_CONCURRENT_IMAGE_DECODES);
    dropBitmapCacheByPath(fetchImage);
  });

  it('admits source inspection through the shared per-document gate', async () => {
    let active = 0;
    let maximumActive = 0;
    const releases: Array<() => void> = [];
    const fetchImage = vi.fn((_path: string, mime: string) => new Promise<Blob>((resolve) => {
      active++;
      maximumActive = Math.max(maximumActive, active);
      releases.push(() => {
        active--;
        resolve(new Blob([new Uint8Array([1])], { type: mime }));
      });
    }));

    const inspections = Array.from({ length: MAX_CONCURRENT_IMAGE_DECODES + 2 }, (_, index) =>
      inspectCachedRasterSource(`word/media/inspect-${index}.png`, 'image/png', fetchImage));
    await vi.waitFor(() => expect(active).toBe(MAX_CONCURRENT_IMAGE_DECODES));

    for (const release of releases.splice(0)) release();
    await vi.waitFor(() => expect(releases).toHaveLength(2));
    for (const release of releases.splice(0)) release();
    await Promise.all(inspections);
    expect(maximumActive).toBe(MAX_CONCURRENT_IMAGE_DECODES);
  });
});

describe('releaseOwnedBitmap', () => {
  it('treats ImageBitmap.close failures as best-effort cleanup', () => {
    const bitmap = {
      close: vi.fn(() => {
        throw new Error('cleanup failed');
      }),
    } as unknown as ImageBitmap;

    expect(() => releaseOwnedBitmap(bitmap)).not.toThrow();
    expect(bitmap.close).toHaveBeenCalledTimes(1);
  });
});

/**
 * Render-pass lease (`acquireBitmapCacheLease`): a renderer resolves EVERY image
 * a page/sheet/slide references and then draws from those references. Without a
 * lease, resolving more images than the LRU cap in one pass evicts — and
 * GPU-closes — bitmaps the pass still holds, so the draw would paint a closed
 * bitmap. While a lease is active, evictions/drops still remove the cache entry
 * (size stays bounded) but the close is deferred to the LAST release.
 */
describe('acquireBitmapCacheLease (render-pass liveness)', () => {
  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  /** Flush the close-through-promise microtasks. */
  const flush = async () => {
    await Promise.resolve();
    await Promise.resolve();
  };

  it('handles a failed cleanup promise while its render lease remains open', async () => {
    const owner = {};
    const failure = new Error('expected image decode failure');
    const unhandled: unknown[] = [];
    const recordUnhandled = (reason: unknown) => unhandled.push(reason);
    process.on('unhandledRejection', recordUnhandled);

    try {
      await withBitmapCacheLease(owner, undefined, async () => {
        await expect(getCachedDerivedBitmap('test', 'failed', owner, async () => {
          throw failure;
        })).rejects.toBe(failure);

        // A sibling image can keep the render pass alive after this decode has
        // been contained. Node reports an unhandled rejection before the next
        // timer when an internal cleanup branch has no rejection handler yet.
        await new Promise((resolve) => setTimeout(resolve, 0));
        expect(unhandled).toEqual([]);
      });
    } finally {
      process.off('unhandledRejection', recordUnhandled);
    }
  });

  it('defers an LRU-eviction close past the cap until the lease is released', async () => {
    const closed: string[] = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => {
      const tag = new Uint8Array(await blob.arrayBuffer())[0];
      return { width: 1, height: 1, close: () => closed.push(`b${tag}`) } as unknown as ImageBitmap;
    }));
    const fetchImage = vi.fn(async (path: string, mime: string) => {
      // Tag the blob with the path's index so each bitmap is identifiable.
      const i = Number(/lease-(\d+)/.exec(path)?.[1] ?? 0);
      return new Blob([new Uint8Array([i % 256])], { type: mime });
    });

    const release = acquireBitmapCacheLease(fetchImage);
    // Resolve one more than the cap (257) in a single leased pass, the way a
    // render pass over a 257-image document would.
    for (let i = 0; i <= 256; i++) {
      await getCachedBitmapByPath(`word/media/lease-${i}.png`, 'image/png', fetchImage);
    }
    await flush();
    // The eviction happened (entry removed → a re-resolve would re-fetch), but
    // the GPU close is deferred: the pass's reference is still drawable.
    expect(closed).toEqual([]);

    release();
    await flush();
    expect(closed).toEqual(['b0']); // the evicted oldest closes at release
  });

  it('defers a dropBitmapCacheByPath close while leased (drop racing an in-flight render)', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));

    const release = acquireBitmapCacheLease(fetchImage);
    await getCachedBitmapByPath('word/media/leased-drop-a.png', 'image/png', fetchImage);
    await getCachedBitmapByPath('word/media/leased-drop-b.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage); // e.g. destroy()/re-parse mid-render
    await flush();
    expect(closes.length).toBe(0); // still drawable for the in-flight pass
    // The cache itself was forgotten immediately: a re-resolve re-decodes.
    await getCachedBitmapByPath('word/media/leased-drop-a.png', 'image/png', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3);

    release();
    await flush();
    expect(closes.length).toBe(2); // the dropped bitmaps close at release
  });

  it('nested leases (concurrent passes): deferred closes run at the LAST release only', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));

    const releaseA = acquireBitmapCacheLease(fetchImage);
    const releaseB = acquireBitmapCacheLease(fetchImage);
    await getCachedBitmapByPath('word/media/nested.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage);
    releaseA();
    await flush();
    expect(closes.length).toBe(0); // pass B still holds the document
    releaseB();
    await flush();
    expect(closes.length).toBe(1);
  });

  it('release is idempotent — a double release neither double-closes nor steals a sibling lease', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));

    const releaseA = acquireBitmapCacheLease(fetchImage);
    const releaseB = acquireBitmapCacheLease(fetchImage);
    await getCachedBitmapByPath('word/media/idem.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage);
    releaseA();
    releaseA(); // double release must NOT decrement B's hold
    await flush();
    expect(closes.length).toBe(0);
    releaseB();
    await flush();
    expect(closes.length).toBe(1);
  });

  it('with no lease active, eviction and drop close immediately (unchanged baseline)', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    await getCachedBitmapByPath('word/media/unleased.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage);
    await flush();
    expect(closes.length).toBe(1);
  });

  it('rejects a render pass before its live decoded images exceed the byte ceiling', async () => {
    const closes: number[] = [];
    const bitmapBytes = MAX_DECODED_IMAGE_BYTES / 2;
    const width = 4096;
    const height = bitmapBytes / width / 4;
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width, height, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));

    const release = acquireBitmapCacheLease(fetchImage);
    await getCachedBitmapByPath('word/media/live-a.png', 'image/png', fetchImage);
    await getCachedBitmapByPath('word/media/live-b.png', 'image/png', fetchImage);
    await expect(getCachedBitmapByPath('word/media/live-c.png', 'image/png', fetchImage))
      .rejects.toMatchObject({
        name: 'OoxmlDecodedImageLimitError',
        code: 'ooxml-decoded-image-limit',
        metric: 'active-decoded-bytes',
        limit: MAX_DECODED_IMAGE_BYTES,
        observed: MAX_DECODED_IMAGE_BYTES + bitmapBytes,
      });
    expect(closes.length).toBe(1);
    release();
    dropBitmapCacheByPath(fetchImage);
  });

  it('accounts base and derived surfaces against one render-pass ceiling', async () => {
    const width = 4096;
    const height = MAX_DECODED_IMAGE_BYTES / 2 / width / 4;
    const closeBase = vi.fn();
    const closeDerived = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width, height, close: closeBase }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async () => new Blob([new Uint8Array([1])]));
    const release = acquireBitmapCacheLease(fetchImage);

    await getCachedBitmapByPath('word/media/base.bin', 'application/octet-stream', fetchImage);
    await getCachedDerivedBitmap('effect', 'first', fetchImage, async () => ({
      bitmap: { width, height, close: closeDerived } as unknown as ImageBitmap,
      owned: true,
    }));
    await expect(getCachedDerivedBitmap('effect', 'second', fetchImage, async () => ({
      bitmap: { width, height, close: closeDerived } as unknown as ImageBitmap,
      owned: true,
    }))).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'active-decoded-bytes',
      observed: MAX_DECODED_IMAGE_BYTES * 1.5,
    });
    expect(closeDerived).toHaveBeenCalledTimes(1);

    release();
    dropBitmapCacheByPath(fetchImage);
    await flush();
    expect(closeBase).toHaveBeenCalledTimes(1);
    expect(closeDerived).toHaveBeenCalledTimes(2);
  });

  it('honours a caller-configured aggregate budget above the adaptive default', async () => {
    const width = 4096;
    const height = 4096; // 64 MiB per RGBA surface; three need 192 MiB.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width, height, close() {} }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const budget = Math.min(HARD_MAX_DECODED_IMAGE_BYTES, MAX_DECODED_IMAGE_BYTES * 2);
    const release = acquireBitmapCacheLease(fetchImage, {
      decodedByteBudget: budget,
      strategy: 'adaptive',
    });

    await expect(Promise.all([
      getCachedBitmapByPath('word/media/configured-a.bin', 'application/octet-stream', fetchImage),
      getCachedBitmapByPath('word/media/configured-b.bin', 'application/octet-stream', fetchImage),
      getCachedBitmapByPath('word/media/configured-c.bin', 'application/octet-stream', fetchImage),
    ])).resolves.toHaveLength(3);

    release();
    dropBitmapCacheByPath(fetchImage);
  });

  it('does not leak one paint\'s configured retained limit into unleased cache work', async () => {
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const release = acquireBitmapCacheLease(fetchImage, {
      decodedByteBudget: 4,
      strategy: 'adaptive',
    });
    await getCachedBitmapByPath('word/media/transient-limit-a.png', 'image/png', fetchImage);
    release();

    await getCachedBitmapByPath('word/media/transient-limit-b.png', 'image/png', fetchImage);
    await getCachedBitmapByPath('word/media/transient-limit-a.png', 'image/png', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    dropBitmapCacheByPath(fetchImage);
  });

  it('serializes overlapping image-bearing paints for the same document owner', async () => {
    const owner = vi.fn(async () => new Blob());
    let releaseFirst!: () => void;
    const firstGate = new Promise<void>((resolve) => { releaseFirst = resolve; });
    const order: string[] = [];

    const first = withBitmapCacheLease(owner, undefined, async () => {
      order.push('first:start');
      await firstGate;
      order.push('first:end');
    });
    const second = withBitmapCacheLease(owner, undefined, async () => {
      order.push('second:start');
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(order).toEqual(['first:start']);

    releaseFirst();
    await Promise.all([first, second]);
    expect(order).toEqual(['first:start', 'first:end', 'second:start']);
  });

  it('releases queue admission after a paint or option-validation failure', async () => {
    const owner = vi.fn(async () => new Blob());
    await expect(withBitmapCacheLease(owner, undefined, async () => {
      throw new Error('paint failed');
    })).rejects.toThrow('paint failed');
    await expect(withBitmapCacheLease(owner, {
      decodedByteBudget: 0,
    }, async () => 'unreachable')).rejects.toThrow(RangeError);

    await expect(withBitmapCacheLease(owner, undefined, async () => 'recovered'))
      .resolves.toBe('recovered');
  });

  it('does not retain an inspected compressed blob after a failed paint', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 100);
    view.setUint32(20, 100);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const path = 'word/media/failed-paint-profile.png';

    await expect(withBitmapCacheLease(fetchImage, undefined, async () => {
      await inspectCachedRasterSource(path, 'image/png', fetchImage);
      throw new Error('paint failed after inspection');
    })).rejects.toThrow('paint failed after inspection');

    await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 10,
      targetHeightPx: 10,
    });
    expect(fetchImage).toHaveBeenCalledTimes(2);
  });
});
