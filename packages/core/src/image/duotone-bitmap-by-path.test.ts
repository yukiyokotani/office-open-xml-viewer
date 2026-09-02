import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './duotone-bitmap-by-path';
import {
  getCachedBitmapByPath,
  dropBitmapCacheByPath,
  acquireBitmapCacheLease,
} from './bitmap-image-by-path';
import type { OffscreenFactory } from './duotone';

/**
 * The core second-layer duotone cache: decode the base blip once (shared
 * path-keyed cache), then run the `<a:duotone>` recolour once per (path +
 * colours). Shared by the docx and pptx renderers. The recolour reads the base
 * bitmap's pixels via getImageData → transform → putImageData → a NEW bitmap, so
 * we inject an offscreen factory + stub createImageBitmap to exercise the path
 * without a real canvas.
 */
describe('getCachedDuotoneBitmapByPath', () => {
  afterEach(() => vi.unstubAllGlobals());

  /** An offscreen surface whose getImageData returns a near-white pixel buffer;
   *  putImageData records the recoloured bytes so a test can confirm the ramp ran. */
  function recordingFactory(record: { out?: Uint8ClampedArray }): OffscreenFactory {
    return ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext() {
        return {
          drawImage() {},
          getImageData(_sx: number, _sy: number, sw: number, sh: number) {
            // One near-white opaque pixel (luminance ≈ 1 → maps to clr2).
            const data = new Uint8ClampedArray(sw * sh * 4).fill(255);
            return { data, width: sw, height: sh } as unknown as ImageData;
          },
          putImageData(img: ImageData) {
            record.out = img.data;
          },
        };
      },
    })) as unknown as OffscreenFactory;
  }

  it('decodes the base once and recolours once per colour pair, caching both', async () => {
    const path = 'ppt/media/duo-cachehit-a.png';
    const baseBitmap = { width: 4, height: 4, close() {} } as unknown as ImageBitmap;
    const recoloured = { width: 4, height: 4, tag: 'duo', close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (src: unknown) => {
      // The base decode passes a Blob; applyDuotone passes the offscreen surface.
      return src instanceof Blob ? baseBitmap : recoloured;
    });
    vi.stubGlobal('createImageBitmap', cib);

    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);

    // Both calls return the SAME recoloured bitmap (memoized), the base blip was
    // fetched + decoded once, and the recolour ran once.
    expect(first).toBe(recoloured);
    expect(second).toBe(recoloured);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    // createImageBitmap: once for the base decode + once for the recolour = 2.
    expect(cib).toHaveBeenCalledTimes(2);
    // The recolour mapped a near-white pixel toward clr2 (DAB6BA): R≈0xDA, not 0.
    expect(record.out?.[0]).toBeGreaterThan(200);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('applies duotone on the authored grid before display-target resampling', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const pngView = new DataView(png.buffer);
    pngView.setUint32(16, 2);
    pngView.setUint32(20, 2);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const surfaces: Array<{ width: number; height: number }> = [];
    const factory = ((width: number, height: number) => {
      surfaces.push({ width, height });
      return {
        width,
        height,
        getContext: () => ({
          drawImage() {},
          getImageData: () => ({
            data: new Uint8ClampedArray(width * height * 4).fill(255),
            width,
            height,
          }),
          putImageData() {},
        }),
      };
    }) as unknown as OffscreenFactory;
    const createBitmap = vi.fn(async (
      source: Blob | { width: number; height: number },
      options?: ImageBitmapOptions,
    ) => ({
      width: options?.resizeWidth ?? (source instanceof Blob ? 2 : source.width),
      height: options?.resizeWidth ?? (source instanceof Blob ? 2 : source.height),
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', createBitmap);
    const path = 'ppt/media/duotone-effect-order.png';

    const result = await getCachedDuotoneBitmapByPath(
      path,
      'image/png',
      { clr1: '000000', clr2: 'FFFFFF' },
      fetchImage,
      {
        targetWidthPx: 1,
        targetHeightPx: 1,
        offscreenFactory: factory,
      },
    );

    expect(surfaces).toEqual([{ width: 2, height: 2 }]);
    expect(result).toMatchObject({ width: 1, height: 1 });
    expect(createBitmap).toHaveBeenNthCalledWith(1, expect.any(Blob));
    expect(createBitmap).toHaveBeenNthCalledWith(2, expect.anything(), {
      resizeWidth: 1,
      resizeQuality: 'high',
    });

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('keys post-effect display variants independently while base variants evolve', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const pngView = new DataView(png.buffer);
    pngView.setUint32(16, 300);
    pngView.setUint32(20, 300);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    let derivedIndex = 0;
    vi.stubGlobal('createImageBitmap', vi.fn(async (
      source: Blob | { width: number; height: number },
      options?: ImageBitmapOptions,
    ) => source instanceof Blob
      ? ({
          width: options?.resizeWidth ?? 300,
          height: options?.resizeWidth ?? 300,
          close() {},
        } as unknown as ImageBitmap)
      : ({
          width: options?.resizeWidth ?? source.width,
          height: options?.resizeWidth ?? source.height,
          derivedIndex: derivedIndex++,
          close() {},
        } as unknown as ImageBitmap)));
    const factory = ((width: number, height: number) => ({
      width,
      height,
      getContext: () => ({
        drawImage() {},
        getImageData: () => ({
          data: new Uint8ClampedArray(width * height * 4),
          width,
          height,
        }),
        putImageData() {},
      }),
    })) as unknown as OffscreenFactory;
    const path = 'ppt/media/duotone-variant-race.png';
    const maxRetainedPixels = 1 << 23;
    const duotone = { clr1: '000000', clr2: 'FFFFFF' };

    await getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 100,
      targetHeightPx: 100,
      maxRetainedPixels,
    });
    const smallDerivedPromise = getCachedDuotoneBitmapByPath(
      path,
      'image/png',
      duotone,
      fetchImage,
      {
        targetWidthPx: 80,
        targetHeightPx: 30,
        maxRetainedPixels,
        offscreenFactory: factory,
      },
    );
    const largeBasePromise = getCachedBitmapByPath(path, 'image/png', fetchImage, {
      targetWidthPx: 200,
      targetHeightPx: 40,
      maxRetainedPixels,
    });
    const [smallDerived, largeBase] = await Promise.all([smallDerivedPromise, largeBasePromise]);
    const largeDerived = await getCachedDuotoneBitmapByPath(
      path,
      'image/png',
      duotone,
      fetchImage,
      {
        targetWidthPx: 200,
        targetHeightPx: 40,
        maxRetainedPixels,
        offscreenFactory: factory,
      },
    );

    expect(smallDerived).toMatchObject({ width: 80, height: 80 });
    expect(largeBase).toMatchObject({ width: 200, height: 200 });
    expect(largeDerived).toMatchObject({ width: 200, height: 200 });
    expect(largeDerived).not.toBe(smallDerived);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('passes through to the base cache (no recolour) when duotone is null', async () => {
    const path = 'ppt/media/duo-passthrough-b.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => baseBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );

    const out = await getCachedDuotoneBitmapByPath(path, 'image/png', null, fetchImage, {});

    // Returns the base bitmap directly; no second recolour decode.
    expect(out).toBe(baseBitmap);
    expect(cib).toHaveBeenCalledTimes(1);

    dropBitmapCacheByPath(fetchImage);
  });

  it('keys separate colour pairs independently (same path, two duotones)', async () => {
    const path = 'ppt/media/duo-two-c.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    let n = 0;
    const cib = vi.fn(async (src: unknown) =>
      src instanceof Blob
        ? baseBitmap
        : ({ width: 2, height: 2, tag: `duo${n++}`, close() {} } as unknown as ImageBitmap),
    );
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const a = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: 'FF0000' }, fetchImage, opts);
    const b = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: '00FF00' }, fetchImage, opts);

    // Two distinct recolour results, but the base was decoded only once.
    expect(a).not.toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('rejects before decode when the required duotone base exceeds its working-set budget', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 12_000);
    new DataView(png.buffer).setUint32(20, 9_000);
    const created: Array<{ width: number; height: number }> = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async (source: Blob | { width: number; height: number }, options?: ImageBitmapOptions) => {
      const width = source instanceof Blob ? options?.resizeWidth ?? 12_000 : source.width;
      const height = source instanceof Blob ? Math.floor(width * 0.75) : source.height;
      const bitmap = { width, height, close() {} } as unknown as ImageBitmap;
      created.push({ width, height });
      return bitmap;
    }));
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const factory = ((width: number, height: number) => ({
      width,
      height,
      getContext: () => ({
        drawImage() {},
        getImageData: () => ({ data: new Uint8ClampedArray([255, 255, 255, 255]), width: 1, height: 1 }),
        putImageData() {},
      }),
    })) as unknown as OffscreenFactory;
    await expect(getCachedDuotoneBitmapByPath(
      'ppt/media/poster.png',
      'image/png',
      { clr1: '000000', clr2: 'FFFFFF' },
      fetchImage,
      { targetWidthPx: 1_200, targetHeightPx: 900, offscreenFactory: factory },
    )).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 8_388_608,
      observed: 12_000 * 9_000,
    });
    expect(created).toEqual([]);
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('duotoneCacheKey suffixes the path with both colours only when a duotone is set', () => {
    expect(duotoneCacheKey('word/media/image1.png')).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', null)).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', { clr1: '000000', clr2: 'DAB6BA' })).toBe(
      'word/media/image1.png|duo:000000:DAB6BA',
    );
  });

  it('keeps strict chart fail-closed results separate from compatibility pass-throughs', async () => {
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async () => bitmap));
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-strict.png';

    // With no OffscreenCanvas the established shape path keeps its source,
    // while chart marker paint must not silently drop the authored effect.
    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage, {},
    )).resolves.toBe(bitmap);
    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage,
      { failClosedOnDuotoneFailure: true },
    )).resolves.toBeNull();

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it.each([
    [
      'an unavailable effect surface',
      (() => null) as unknown as OffscreenFactory,
    ],
    [
      'unavailable pixel readback',
      ((width: number, height: number) => ({
        width,
        height,
        getContext: () => ({
          drawImage() {},
          getImageData() { throw new Error('readback unavailable'); },
          putImageData() {},
        }),
      })) as unknown as OffscreenFactory,
    ],
  ])('resamples the current source after %s while strict callers fail closed', async (_name, factory) => {
    const baseClose = vi.fn();
    const resizedClose = vi.fn();
    const base = { width: 4, height: 2, close: baseClose } as unknown as ImageBitmap;
    const resized = { width: 2, height: 1, close: resizedClose } as unknown as ImageBitmap;
    const createBitmap = vi.fn(async (source: Blob | ImageBitmap) => (
      source instanceof Blob ? base : resized
    ));
    vi.stubGlobal('createImageBitmap', createBitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const path = `ppt/media/duotone-fallback-${_name}.png`;
    const duotone = { clr1: '000000', clr2: 'FFFFFF' };
    const options = {
      targetWidthPx: 2,
      targetHeightPx: 1,
      offscreenFactory: factory,
    };

    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage, options,
    )).resolves.toBe(resized);
    await expect(getCachedDuotoneBitmapByPath(
      path, 'image/png', duotone, fetchImage,
      { ...options, failClosedOnDuotoneFailure: true },
    )).resolves.toBeNull();

    expect(createBitmap).toHaveBeenCalledTimes(2);
    expect(createBitmap).toHaveBeenNthCalledWith(1, expect.any(Blob));
    expect(createBitmap).toHaveBeenNthCalledWith(2, base, {
      resizeWidth: 2,
      resizeQuality: 'high',
    });

    dropDuotoneBitmapCache(fetchImage);
    await Promise.resolve();
    expect(resizedClose).toHaveBeenCalledOnce();
    expect(baseClose).not.toHaveBeenCalled();
    dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    expect(baseClose).toHaveBeenCalledOnce();
  });

  // ── Second-layer × base-eviction interaction ────────────────────────────────
  // A PASS-THROUGH entry (the pixel pipeline was unavailable, so the recolour
  // resolved to the base bitmap itself) must not outlive the base: the base LRU
  // protects itself by removing the entry at eviction so the next resolve
  // re-decodes, but a lingering second-layer entry would bypass that re-decode
  // and keep serving the (now closed) base bitmap.
  it('a pass-through entry never outlives the base: after base LRU eviction, a re-resolve returns a live bitmap', async () => {
    // No offscreenFactory and no OffscreenCanvas in this env → applyDuotone
    // returns the base unchanged (pass-through).
    const made: Array<{ closed: boolean }> = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      const bmp = { width: 2, height: 2, closed: false, close(): void { this.closed = true; } };
      made.push(bmp);
      return bmp as unknown as ImageBitmap;
    }));
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-passthrough-evict.png';

    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {});
    expect(first).toBe(made[0] as unknown as ImageBitmap); // pass-through: the base itself

    // Evict the base entry with LRU pressure (256 more distinct paths, no lease
    // held) — the base bitmap is closed.
    for (let i = 0; i < 256; i++) {
      await getCachedBitmapByPath(`ppt/media/duo-pressure-${i}.png`, 'image/png', fetchImage);
    }
    await new Promise((r) => setTimeout(r, 0));
    expect(made[0].closed).toBe(true);

    // The next render pass re-resolves the same (path, duotone): it must NOT be
    // served the stale pass-through entry (a closed bitmap) — the base
    // re-decodes and the pass-through re-derives from the live base.
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {});
    expect(second).not.toBeNull();
    expect((second as unknown as { closed: boolean }).closed).toBe(false);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('dropping BOTH caches around an in-flight pass-through closes the shared bitmap exactly once', async () => {
    // A pass-through entry still in flight at drop time resolves to the base
    // bitmap, so both the duotone drop and the base drop would target the SAME
    // bitmap. The funneled close dedupe (closeBitmapOnce) must close it exactly
    // once — whichever interleaving occurs — with no reliance on
    // ImageBitmap.close() idempotence.
    const closes: number[] = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({
      width: 2,
      height: 2,
      close: () => closes.push(1),
    }) as unknown as ImageBitmap));
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const path = 'ppt/media/duo-double-close.png';
    let markEffectStarted!: () => void;
    const effectStarted = new Promise<void>((resolve) => { markEffectStarted = resolve; });
    const unavailableFactory = (() => {
      markEffectStarted();
      return null;
    }) as unknown as OffscreenFactory;

    const release = acquireBitmapCacheLease(fetchImage);
    // Settle the base first so the duotone wrapper reaches its second-layer
    // entry creation promptly.
    await getCachedBitmapByPath(path, 'image/png', fetchImage);
    const p = getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, {
      offscreenFactory: unavailableFactory,
    });
    // Wait until the derived producer is running, which proves its entry was
    // inserted with a current epoch before either drop occurs.
    await effectStarted;
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
    await p;
    release();
    await new Promise((r) => setTimeout(r, 0));

    expect(closes.length).toBe(1);
  });

  it('does not recreate a duotone entry when a full drop wins after the base resolves', async () => {
    let finishBase!: (bitmap: ImageBitmap) => void;
    const baseClose = vi.fn();
    const derivedClose = vi.fn();
    const base = { width: 2, height: 2, close: baseClose } as unknown as ImageBitmap;
    const derived = { width: 2, height: 2, close: derivedClose } as unknown as ImageBitmap;
    const cib = vi.fn()
      .mockImplementationOnce(() => new Promise<ImageBitmap>((resolve) => { finishBase = resolve; }))
      .mockResolvedValueOnce(derived);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const release = acquireBitmapCacheLease(fetchImage);
    const pending = getCachedDuotoneBitmapByPath(
      'ppt/media/duo-owner-drop-race.png',
      'image/png',
      { clr1: '000000', clr2: 'FFFFFF' },
      fetchImage,
      { offscreenFactory: recordingFactory({}) },
    );
    await vi.waitFor(() => expect(cib).toHaveBeenCalledOnce());
    const rejected = expect(pending).rejects.toThrow(/cache.*dropped/i);

    // Resolve the base and tear down in the same task, before the wrapper's
    // continuation can insert its derived entry.
    finishBase(base);
    dropBitmapCacheByPath(fetchImage);
    await rejected;

    expect(cib).toHaveBeenCalledOnce();
    expect(derivedClose).not.toHaveBeenCalled();
    release();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(baseClose).toHaveBeenCalledOnce();
  });

  it('does not recreate a duotone entry when its namespace is dropped after the base resolves', async () => {
    let finishBase!: (bitmap: ImageBitmap) => void;
    const baseClose = vi.fn();
    const derivedClose = vi.fn();
    const base = { width: 2, height: 2, close: baseClose } as unknown as ImageBitmap;
    const derived = { width: 2, height: 2, close: derivedClose } as unknown as ImageBitmap;
    const cib = vi.fn()
      .mockImplementationOnce(() => new Promise<ImageBitmap>((resolve) => { finishBase = resolve; }))
      .mockResolvedValueOnce(derived);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const release = acquireBitmapCacheLease(fetchImage);
    const pending = getCachedDuotoneBitmapByPath(
      'ppt/media/duo-namespace-drop-race.png',
      'image/png',
      { clr1: '000000', clr2: 'FFFFFF' },
      fetchImage,
      { offscreenFactory: recordingFactory({}) },
    );
    await vi.waitFor(() => expect(cib).toHaveBeenCalledOnce());
    const rejected = expect(pending).rejects.toThrow(/cache.*dropped/i);

    finishBase(base);
    dropDuotoneBitmapCache(fetchImage);
    await rejected;

    expect(cib).toHaveBeenCalledOnce();
    expect(derivedClose).not.toHaveBeenCalled();
    release();
    dropBitmapCacheByPath(fetchImage);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(baseClose).toHaveBeenCalledOnce();
  });
});
