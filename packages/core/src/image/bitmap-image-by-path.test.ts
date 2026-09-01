import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  peekCachedBitmapByPath,
  dropBitmapCacheByPath,
  acquireBitmapCacheLease,
} from './bitmap-image-by-path';
import {
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

  it('keys display-sized raster variants by stable resolution bands', async () => {
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
      targetWidthPx: 1000,
      targetHeightPx: 750,
    });
    const sameBand = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1010,
      targetHeightPx: 758,
    });
    const larger = await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 1400,
      targetHeightPx: 1050,
    });

    expect(a).toBe(sameBand);
    expect(larger).not.toBe(a);
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(cib).toHaveBeenCalledTimes(2);
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
});
