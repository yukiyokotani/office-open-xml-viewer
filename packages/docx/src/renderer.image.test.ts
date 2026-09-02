import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { decodeRaster, preloadImages } from './test-support/renderer-internals.test-support.js';
import { dropColorReplacedCache } from './renderer';
import * as core from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types';

/**
 * docx raster blips decode through `fetchImage(path, mime)` (twin of pptx's
 * lazy-bytes path) instead of `fetch`-ing an inlined data URL. `preloadImages`
 * keys the decoded-image map by `imageKey(imagePath, colorReplaceFrom)` and must
 * decode each distinct key exactly once. The base (colour-replacement-free)
 * bitmap now comes from the shared, per-document, path-keyed core cache, so a
 * plain + recoloured reference to the same path share ONE fetch/decode; this
 * test pins that raster + keying + shared-base contract.
 */
/** Stub OffscreenCanvas + 2D context so applyColorReplacement's
 *  getImageData/putImageData make-transparent pass runs in the node test env. */
function stubOffscreen(): void {
  class FakeOffscreen {
    width: number;
    height: number;
    constructor(w: number, h: number) { this.width = w; this.height = h; }
    getContext() {
      return {
        drawImage: () => {},
        getImageData: (_x: number, _y: number, w: number, h: number) => ({
          data: new Uint8ClampedArray(Math.max(1, w) * Math.max(1, h) * 4),
          width: w,
          height: h,
        }),
        putImageData: () => {},
      };
    }
  }
  vi.stubGlobal('OffscreenCanvas', FakeOffscreen);
}

describe('docx lazy image bytes', () => {
  beforeEach(() => {
    // `createImageBitmap` doesn't exist in the node test env; stub it to a
    // sentinel image bitmap.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_src: unknown) => ({ width: 2, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('decodeRaster pulls bytes by path via fetchImage and passes the MIME through', async () => {
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    const bmp = await decodeRaster('word/media/image1.png', 'image/png', undefined, fetchImage);
    expect(bmp).toBeTruthy();
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('word/media/image1.png', 'image/png');
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('threads the laid-out device-pixel target into an oversized raster decode', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 12_090);
    new DataView(png.buffer).setUint32(20, 9_063);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{
          type: 'image', imagePath: 'word/media/poster.png', mimeType: 'image/png',
          widthPt: 100, heightPt: 50,
        }],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await preloadImages(doc, fetchImage, undefined, 2);

    expect(globalThis.createImageBitmap).toHaveBeenCalledWith(
      expect.any(Blob),
      expect.objectContaining({ resizeWidth: 200, resizeQuality: 'high' }),
    );
  });

  it('applies exact clrChange matching before display-target resampling', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 2);
    view.setUint32(20, 2);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    let written: Uint8ClampedArray | undefined;
    class ExactColorSurface {
      constructor(readonly width: number, readonly height: number) {}
      getContext() {
        return {
          drawImage() {},
          getImageData: () => ({
            // A decoder-side 2→1 resize would interpolate away the authored
            // exact white pixel. The effect must instead see both source pixels.
            data: this.width === 2 && this.height === 2
              ? new Uint8ClampedArray([
                  255, 255, 255, 255,
                  254, 255, 255, 255,
                  253, 255, 255, 255,
                  252, 255, 255, 255,
                ])
              : new Uint8ClampedArray([254, 255, 255, 255]),
            width: this.width,
            height: this.height,
          }),
          putImageData: (data: ImageData) => { written = new Uint8ClampedArray(data.data); },
        };
      }
    }
    vi.stubGlobal('OffscreenCanvas', ExactColorSurface);
    const createBitmap = vi.fn(async (
      source: Blob | ExactColorSurface,
      options?: ImageBitmapOptions,
    ) => ({
      width: options?.resizeWidth ?? (source instanceof Blob ? 2 : source.width),
      height: options?.resizeWidth ? 1 : source instanceof Blob ? 2 : source.height,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', createBitmap);

    const result = await decodeRaster(
      'word/media/exact-clr-change.png',
      'image/png',
      'FFFFFF',
      fetchImage,
      0,
      0,
      undefined,
      false,
      undefined,
      { targetWidthPx: 1, targetHeightPx: 1 },
    );

    expect(written).toEqual(new Uint8ClampedArray([
      255, 255, 255, 0,
      254, 255, 255, 255,
      253, 255, 255, 255,
      252, 255, 255, 255,
    ]));
    expect(result).toMatchObject({ width: 1, height: 1 });
    expect(createBitmap).toHaveBeenNthCalledWith(1, expect.any(Blob));
    expect(createBitmap).toHaveBeenNthCalledWith(2, expect.any(ExactColorSurface), {
      resizeWidth: 1,
      resizeQuality: 'high',
    });
  });

  it('applies chained clrChange and duotone in one source-grid pass and keys non-square targets', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 8);
    view.setUint32(20, 4);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const writes: Uint8ClampedArray[] = [];
    const surfaces: ChainSurface[] = [];
    class ChainSurface {
      constructor(readonly width: number, readonly height: number) { surfaces.push(this); }
      getContext() {
        return {
          drawImage() {},
          getImageData: () => {
            const data = new Uint8ClampedArray(this.width * this.height * 4);
            data.set([
              255, 255, 255, 255,
              128, 128, 128, 255,
            ]);
            return { data, width: this.width, height: this.height };
          },
          putImageData: (data: ImageData) => writes.push(new Uint8ClampedArray(data.data)),
        };
      }
    }
    vi.stubGlobal('OffscreenCanvas', ChainSurface);
    let finalIndex = 0;
    const createBitmap = vi.fn(async (
      source: Blob | ChainSurface,
      options?: ImageBitmapOptions,
    ) => {
      if (source instanceof Blob) {
        return { width: 8, height: 4, close: vi.fn() } as unknown as ImageBitmap;
      }
      const width = options?.resizeWidth ?? source.width;
      return {
        width,
        height: Math.ceil(source.height * width / source.width),
        finalIndex: finalIndex++,
        close: vi.fn(),
      } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', createBitmap);
    const args = [
      'word/media/chained-effects.png', 'image/png', 'FFFFFF', fetchImage,
      0, 0, { clr1: '000000', clr2: 'FF0000' }, false, undefined,
    ] as const;

    const wide = await decodeRaster(...args, { targetWidthPx: 4, targetHeightPx: 1 });
    const tall = await decodeRaster(...args, { targetWidthPx: 2, targetHeightPx: 3 });
    const wideAgain = await decodeRaster(...args, { targetWidthPx: 4, targetHeightPx: 1 });

    expect(wide).toMatchObject({ width: 4, height: 2 });
    expect(tall).toMatchObject({ width: 6, height: 3 });
    expect(wideAgain).toBe(wide);
    expect(fetchImage).toHaveBeenCalledOnce();
    expect(surfaces).toHaveLength(2);
    expect(surfaces.every(surface => surface.width === 8 && surface.height === 4)).toBe(true);
    expect(createBitmap).toHaveBeenCalledTimes(3);
    expect(createBitmap).toHaveBeenNthCalledWith(1, expect.any(Blob));
    expect(createBitmap).toHaveBeenNthCalledWith(2, surfaces[0], {
      resizeWidth: 4,
      resizeQuality: 'high',
    });
    expect(createBitmap).toHaveBeenNthCalledWith(3, surfaces[1], {
      resizeWidth: 6,
      resizeQuality: 'high',
    });
    // Exact white becomes transparent first; the still-opaque gray neighbour is
    // then mapped once through the black→red duotone ramp in the same buffer.
    expect([...writes[0].slice(0, 8)]).toEqual([
      255, 255, 255, 0,
      128, 0, 0, 255,
    ]);
  });

  it('resamples the clrChange result when chained duotone fails, while strict mode returns null', async () => {
    const written: Uint8ClampedArray[] = [];
    class FallbackSurface {
      constructor(readonly width: number, readonly height: number) {}
      getContext() {
        return {
          drawImage() {},
          getImageData: () => ({
            data: new Uint8ClampedArray([
              255, 255, 255, 255,
              128, 128, 128, 255,
              0, 0, 0, 0,
              0, 0, 0, 0,
            ]),
            width: this.width,
            height: this.height,
          }),
          putImageData: (data: ImageData) => written.push(new Uint8ClampedArray(data.data)),
        };
      }
    }
    vi.stubGlobal('OffscreenCanvas', FallbackSurface);
    const base = { width: 2, height: 2, close: vi.fn() } as unknown as ImageBitmap;
    const fallback = { width: 1, height: 1, close: vi.fn() } as unknown as ImageBitmap;
    const createBitmap = vi.fn(async (source: Blob | FallbackSurface) => (
      source instanceof Blob ? base : fallback
    ));
    vi.stubGlobal('createImageBitmap', createBitmap);
    const duotoneTransform = vi.spyOn(core, 'duotoneImageData')
      .mockImplementation(() => { throw new Error('duotone transform unavailable'); });
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const args = [
      'word/media/chained-effect-fallback.png', 'image/png', 'FFFFFF', fetchImage,
      0, 0, { clr1: '000000', clr2: 'FFFFFF' },
    ] as const;

    try {
      await expect(decodeRaster(
        ...args,
        false,
        undefined,
        { targetWidthPx: 1, targetHeightPx: 1 },
      )).resolves.toBe(fallback);
      await expect(decodeRaster(
        ...args,
        true,
        undefined,
        { targetWidthPx: 1, targetHeightPx: 1 },
      )).resolves.toBeNull();

      expect([...written[0].slice(0, 8)]).toEqual([
        255, 255, 255, 0,
        128, 128, 128, 255,
      ]);
      expect(createBitmap).toHaveBeenCalledTimes(2);
      expect(createBitmap).toHaveBeenNthCalledWith(2, expect.any(FallbackSurface), {
        resizeWidth: 1,
        resizeQuality: 'high',
      });
    } finally {
      duotoneTransform.mockRestore();
      dropColorReplacedCache(fetchImage);
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('admits a chained effect source within the four-surface working-set budget', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 3_000);
    view.setUint32(20, 2_500); // 7.5 MP: above 32 MP / 5, below 32 MP / 4.
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    class BudgetSurface {
      constructor(readonly width: number, readonly height: number) {}
      getContext() {
        return {
          drawImage() {},
          getImageData: () => ({
            data: new Uint8ClampedArray([255, 255, 255, 255]),
            width: this.width,
            height: this.height,
          }),
          putImageData() {},
        };
      }
    }
    vi.stubGlobal('OffscreenCanvas', BudgetSurface);
    const createBitmap = vi.fn(async (source: Blob | BudgetSurface, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? (source instanceof Blob ? 3_000 : source.width),
      height: options?.resizeWidth ? 250 : source instanceof Blob ? 2_500 : source.height,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', createBitmap);

    await expect(decodeRaster(
      'word/media/four-surface-budget.png',
      'image/png',
      'FFFFFF',
      fetchImage,
      0,
      0,
      { clr1: '000000', clr2: 'FFFFFF' },
      false,
      undefined,
      { targetWidthPx: 300, targetHeightPx: 250 },
    )).resolves.toMatchObject({ width: 300, height: 250 });
    expect(createBitmap).toHaveBeenCalledTimes(2);
  });

  it('preserves native decoding for an ordinary raster that fits the document budget', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 800);
    new DataView(png.buffer).setUint32(20, 600);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const fetchImage = vi.fn(async () => blob);
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{
          type: 'image', imagePath: 'word/media/photo.png', mimeType: 'image/png',
          widthPt: 100, heightPt: 75,
        }],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await preloadImages(doc, fetchImage, undefined, 1);

    expect(globalThis.createImageBitmap).toHaveBeenCalledWith(blob);
  });

  it('keeps DrawingML pixel effects on the authored source grid during preload', async () => {
    stubOffscreen();
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 800);
    new DataView(png.buffer).setUint32(20, 600);
    const fetchImage = vi.fn(async () =>
      new Blob([png as BlobPart], { type: 'image/png' }));
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{
          type: 'image',
          imagePath: 'word/media/duotone.png',
          mimeType: 'image/png',
          widthPt: 100,
          heightPt: 75,
          duotone: { clr1: '000000', clr2: 'FFFFFF' },
        }],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await preloadImages(doc, fetchImage, undefined, 2);

    const decode = globalThis.createImageBitmap as ReturnType<typeof vi.fn>;
    expect(decode).toHaveBeenCalledTimes(2);
    expect(decode.mock.calls[0]).toHaveLength(1);
    expect(decode.mock.calls[1]).toHaveLength(1);
  });

  it('does not apply a display-sized pixel cap to a TIFF codec result', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const bitmap = { width: 1000, height: 1000, close() {} } as unknown as ImageBitmap;
    const render = vi.fn(async () => bitmap);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{
          type: 'image', imagePath: 'word/media/photo.tiff', mimeType: 'image/tiff',
          widthPt: 10, heightPt: 10,
        }],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await expect(preloadImages(
      doc,
      fetchImage,
      undefined,
      1,
      undefined,
      { render },
    )).resolves.toBeDefined();
    expect(render).toHaveBeenCalledOnce();
  });

  it('display-decodes the redacted issue #1426 image class even though each source is below 32 MP', async () => {
    const png = (width: number, height: number) => {
      const bytes = new Uint8Array(24);
      bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      bytes.set([0x49, 0x48, 0x44, 0x52], 12);
      const view = new DataView(bytes.buffer);
      view.setUint32(16, width);
      view.setUint32(20, height);
      return bytes;
    };
    const blobs = new Map([
      ['word/media/photo.png', new Blob([png(4518, 6777) as BlobPart], { type: 'image/png' })],
      ['word/media/header.png', new Blob([png(8000, 2311) as BlobPart], { type: 'image/png' })],
    ]);
    const fetchImage = vi.fn(async (path: string) => blobs.get(path) as Blob);
    const image = (imagePath: string, widthPt: number, heightPt: number) => ({
      type: 'image', imagePath, mimeType: 'image/png', widthPt, heightPt,
    });
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [
          image('word/media/photo.png', 300, 450),
          image('word/media/header.png', 100, 28.8875),
        ],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await preloadImages(doc, fetchImage, undefined, 2);

    expect(globalThis.createImageBitmap).toHaveBeenCalledWith(
      blobs.get('word/media/photo.png'),
      expect.objectContaining({ resizeWidth: 600, resizeQuality: 'high' }),
    );
    expect(globalThis.createImageBitmap).toHaveBeenCalledWith(
      blobs.get('word/media/header.png'),
      expect.objectContaining({ resizeWidth: 201, resizeQuality: 'high' }),
    );
  });

  it('shares one adaptive quality scale across a page whose display targets exceed its budget', async () => {
    const bytes = new Uint8Array(24);
    bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    bytes.set([0x49, 0x48, 0x44, 0x52], 12);
    const view = new DataView(bytes.buffer);
    view.setUint32(16, 2000);
    view.setUint32(20, 2000);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/png' }));
    const image = (imagePath: string) => ({
      type: 'image', imagePath, mimeType: 'image/png', widthPt: 100, heightPt: 100,
    });
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [image('word/media/a.png'), image('word/media/b.png')],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await preloadImages(doc, fetchImage, undefined, 10, {
      decodedByteBudget: 2_000_000,
      strategy: 'adaptive',
    });

    expect(globalThis.createImageBitmap).toHaveBeenNthCalledWith(
      1,
      expect.any(Blob),
      expect.objectContaining({ resizeWidth: 500, resizeQuality: 'high' }),
    );
    expect(globalThis.createImageBitmap).toHaveBeenNthCalledWith(
      2,
      expect.any(Blob),
      expect.objectContaining({ resizeWidth: 500, resizeQuality: 'high' }),
    );
  });

  it('adapts a default working set above 128 MiB while strict mode rejects it', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    const view = new DataView(png.buffer);
    view.setUint32(16, 4096);
    view.setUint32(20, 4096);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const image = (imagePath: string) => ({
      type: 'image', imagePath, mimeType: 'image/png', widthPt: 4096, heightPt: 4096,
    });
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [
          image('word/media/a.png'),
          image('word/media/b.png'),
          image('word/media/c.png'),
        ],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    await expect(preloadImages(doc, fetchImage, undefined, 1, {
      strategy: 'strict',
    })).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'active-decoded-bytes',
      limit: 128 * 1024 * 1024,
      observed: 4096 * 4096 * 4 * 3,
    });
    expect(globalThis.createImageBitmap).not.toHaveBeenCalled();

    await expect(preloadImages(doc, fetchImage, undefined, 1)).resolves.toHaveLength(3);
    const resizeWidths = (globalThis.createImageBitmap as ReturnType<typeof vi.fn>).mock.calls
      .map(([, options]) => (options as ImageBitmapOptions | undefined)?.resizeWidth)
      .filter((width): width is number => typeof width === 'number');
    expect(resizeWidths).toHaveLength(3);
    expect(resizeWidths.every((width) => width < 4096)).toBe(true);
    expect(resizeWidths.reduce(
      (bytes, width) => bytes + width ** 2 * 4,
      0,
    )).toBeLessThanOrEqual(128 * 1024 * 1024);
  });

  it('threads the opt-in TIFF codec through the DOCX image path', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const bitmap = { width: 8, height: 4, close() {} } as unknown as ImageBitmap;
    const render = vi.fn(async () => bitmap);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));

    await expect(decodeRaster(
      'word/media/image1.tiff',
      'image/tiff',
      undefined,
      fetchImage,
      0,
      0,
      undefined,
      false,
      { render },
    )).resolves.toBe(bitmap);

    expect(render).toHaveBeenCalledTimes(1);
    expect(globalThis.createImageBitmap).not.toHaveBeenCalled();
  });

  it('preloadImages keys by imagePath and decodes each distinct key exactly once', async () => {
    const fetchImage = vi.fn(
      async (path: string, mime: string) => new Blob([new Uint8Array([path.length])], { type: mime }),
    );
    const imgRun = (imagePath: string) => ({
      type: 'image',
      imagePath,
      mimeType: 'image/png',
      widthPt: 10,
      heightPt: 10,
    });
    // image1 is referenced twice (must collapse to ONE decode); image2 once.
    const doc = {
      body: [
        { type: 'paragraph', runs: [imgRun('word/media/image1.png')] },
        { type: 'paragraph', runs: [imgRun('word/media/image1.png')] }, // dup → same key
        { type: 'paragraph', runs: [imgRun('word/media/image2.png')] }, // distinct key
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    const map = await preloadImages(doc, fetchImage);

    // Two distinct keys (the raster path itself, no colorReplaceFrom suffix).
    expect(map.has('word/media/image1.png')).toBe(true);
    expect(map.has('word/media/image2.png')).toBe(true);
    expect(map.size).toBe(2);
    // The duplicate reference must NOT trigger a second fetch/decode for its key:
    // one fetch per distinct path = 2 total (not 3).
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
  });

  it('keeps authored occurrence order when registry key sorting would choose another SVG fallback', async () => {
    const rasterPath = 'word/media/shared.png';
    const earlySvgPath = 'word/media/early.svg';
    const lateSvgPath = 'word/media/late.svg';
    const paragraphs = Array.from({ length: 11 }, (_, index) => ({
      type: 'paragraph',
      runs: index === 2 || index === 10 ? [{
        type: 'image',
        imagePath: rasterPath,
        mimeType: 'image/png',
        svgImagePath: index === 2 ? earlySvgPath : lateSvgPath,
        widthPt: 10,
        heightPt: 10,
      }] : [],
    }));
    const doc = {
      section: {}, body: paragraphs, headers: {}, footers: {},
    } as unknown as DocxDocumentModel;
    const vector = { width: 2, height: 2 } as unknown as HTMLImageElement;
    const svg = vi.spyOn(core, 'getCachedSvgImageByPath').mockResolvedValue(vector);
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([], { type: mime }));

    try {
      const map = await preloadImages(doc, fetchImage);
      expect(map.get(rasterPath)).toBe(vector);
      expect(svg).toHaveBeenCalledTimes(1);
      expect(svg).toHaveBeenCalledWith(earlySvgPath, fetchImage);
    } finally {
      svg.mockRestore();
    }
  });

  it('uses resolved numbering marker metadata for no-extent picture-bullet decode size', async () => {
    const path = 'word/media/bullet.wmf';
    const paragraph = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
      runs: [{ type: 'text', text: 'body', fontSize: 10 }],
      numbering: {
        numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 0, tab: 18,
        suff: 'tab', picBulletImagePath: path, picBulletMimeType: 'image/x-wmf',
        fontFacts: { fontSize: 18 },
      },
    };
    const doc = {
      section: {}, body: [paragraph], headers: {}, footers: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc);
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const raster = vi.spyOn(core, 'getCachedBitmapByPath').mockResolvedValue(bitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([], { type: mime }));

    try {
      await preloadImages(doc, fetchImage, services);
      expect(raster).toHaveBeenCalledWith(
        path,
        'image/x-wmf',
        fetchImage,
        expect.objectContaining({ widthPt: 18, heightPt: 18 }),
      );
    } finally {
      raster.mockRestore();
    }
  });

  it('uses the first text run size when picture-bullet metadata has no explicit font size', async () => {
    const path = 'word/media/run-sized-bullet.wmf';
    const paragraph = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [], defaultFontSize: 10,
      runs: [{ type: 'text', text: 'body', fontSize: 16 }],
      numbering: {
        numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 0, tab: 18,
        suff: 'tab', picBulletImagePath: path, picBulletMimeType: 'image/x-wmf',
        fontFacts: {},
      },
    };
    const doc = {
      section: {}, body: [paragraph], headers: {}, footers: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc);
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const raster = vi.spyOn(core, 'getCachedBitmapByPath').mockResolvedValue(bitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([], { type: mime }));

    try {
      await preloadImages(doc, fetchImage, services);
      expect(raster).toHaveBeenCalledWith(
        path,
        'image/x-wmf',
        fetchImage,
        expect.objectContaining({ widthPt: 16, heightPt: 16 }),
      );
    } finally {
      raster.mockRestore();
    }
  });

  it('propagates a corrupt image decode and allows a later retry', async () => {
    const path = 'word/media/corrupt.png';
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{ type: 'image', imagePath: path, mimeType: 'image/png', widthPt: 10, heightPt: 10 }],
      }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    const decode = globalThis.createImageBitmap as ReturnType<typeof vi.fn>;
    decode.mockRejectedValueOnce(new Error('corrupt PNG payload'));

    await expect(preloadImages(doc, fetchImage)).rejects.toThrow('corrupt PNG payload');
    await expect(preloadImages(doc, fetchImage)).resolves.toEqual(new Map([
      [path, expect.any(Object)],
    ]));
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(decode).toHaveBeenCalledTimes(2);
  });

  it('a recoloured ref reuses the shared base bitmap: distinct map key, ONE fetch/decode', async () => {
    // The colorReplaceFrom variant is a distinct cache key (its make-transparent
    // result differs), but its BASE bitmap now comes from the shared path-keyed
    // core cache — so a plain + recoloured reference to the same path share one
    // fetch and one decode; only the recolour pass runs per (path, colour).
    // Stub OffscreenCanvas so applyColorReplacement's getImageData/putImageData
    // pass actually runs (otherwise it throws and the entry is dropped).
    stubOffscreen();
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const plain = { type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png', widthPt: 1, heightPt: 1 };
    const recoloured = { ...plain, colorReplaceFrom: 'FFFFFF' };
    const doc = {
      body: [
        { type: 'paragraph', runs: [plain] },
        { type: 'paragraph', runs: [recoloured] },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    try {
      const map = await preloadImages(doc, fetchImage);
      // Two distinct map keys: the plain path and its recolour suffix.
      expect(map.has('word/media/image1.png')).toBe(true);
      expect(map.has('word/media/image1.png|clr:FFFFFF')).toBe(true);
      // Shared base → ONE fetch and ONE createImageBitmap decode for both refs
      // (down from two before the shared core cache).
      expect(fetchImage).toHaveBeenCalledTimes(1);
      const decodes = (globalThis.createImageBitmap as ReturnType<typeof vi.fn>).mock.calls.length;
      // One decode for the base blob + one for the recoloured OffscreenCanvas.
      expect(decodes).toBe(2);
      // The recolour produced a distinct bitmap, not the base itself.
      expect(map.get('word/media/image1.png')).not.toBe(map.get('word/media/image1.png|clr:FFFFFF'));
    } finally {
      dropColorReplacedCache(fetchImage);
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('decodeRaster memoizes the recolour per (path, colour): a repeat call re-runs neither decode nor recolour', async () => {
    stubOffscreen();
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    try {
      const a = await decodeRaster('word/media/image1.png', 'image/png', 'FFFFFF', fetchImage);
      const decodesAfterFirst = (globalThis.createImageBitmap as ReturnType<typeof vi.fn>).mock.calls.length;
      const b = await decodeRaster('word/media/image1.png', 'image/png', 'FFFFFF', fetchImage);
      expect(b).toBe(a); // memoized recolour result reused
      expect(fetchImage).toHaveBeenCalledTimes(1); // base fetched once
      // No further createImageBitmap on the repeat: neither base decode nor recolour re-ran.
      expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>).mock.calls.length).toBe(decodesAfterFirst);
    } finally {
      dropColorReplacedCache(fetchImage);
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('retries a failed chained final bake without leaking or closing the cached base', async () => {
    stubOffscreen();
    const baseClose = vi.fn();
    const finalClose = vi.fn();
    const base = { width: 2, height: 2, close: baseClose } as unknown as ImageBitmap;
    const final = { width: 2, height: 2, close: finalClose } as unknown as ImageBitmap;
    const createBitmap = vi.fn()
      .mockResolvedValueOnce(base)
      .mockRejectedValueOnce(new Error('final effect bake failed'))
      .mockResolvedValueOnce(final);
    vi.stubGlobal('createImageBitmap', createBitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const decode = () => decodeRaster(
      'word/media/chained-effect-retry.png',
      'image/png',
      'FFFFFF',
      fetchImage,
      0,
      0,
      { clr1: '000000', clr2: 'FFFFFF' },
    );

    await expect(decode()).rejects.toThrow('final effect bake failed');
    await expect(decode()).resolves.toBe(final);

    expect(fetchImage).toHaveBeenCalledOnce();
    expect(createBitmap).toHaveBeenCalledTimes(3);
    expect(baseClose).not.toHaveBeenCalled();
    dropColorReplacedCache(fetchImage);
    await Promise.resolve();
    expect(finalClose).toHaveBeenCalledOnce();
    expect(baseClose).not.toHaveBeenCalled();
    core.dropBitmapCacheByPath(fetchImage);
    await Promise.resolve();
    expect(baseClose).toHaveBeenCalledOnce();
  });

  it('defers closing a cached recolour until the active render lease is released', async () => {
    stubOffscreen();
    const base = { width: 2, height: 2, close: vi.fn() } as unknown as ImageBitmap;
    const recoloured = { width: 2, height: 2, close: vi.fn() } as unknown as ImageBitmap;
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn()
        .mockResolvedValueOnce(base)
        .mockResolvedValueOnce(recoloured),
    );
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const release = core.acquireBitmapCacheLease(fetchImage);

    try {
      expect(await decodeRaster(
        'word/media/image1.png',
        'image/png',
        'FFFFFF',
        fetchImage,
      )).toBe(recoloured);
      dropColorReplacedCache(fetchImage);
      await Promise.resolve();
      expect(recoloured.close).not.toHaveBeenCalled();

      release();
      await Promise.resolve();
      expect(recoloured.close).toHaveBeenCalledTimes(1);
    } finally {
      release();
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('does not recreate a colour-effect entry when a full drop wins after the base resolves', async () => {
    stubOffscreen();
    let finishBase!: (bitmap: ImageBitmap) => void;
    const baseClose = vi.fn();
    const recolouredClose = vi.fn();
    const base = { width: 2, height: 2, close: baseClose } as unknown as ImageBitmap;
    const recoloured = { width: 2, height: 2, close: recolouredClose } as unknown as ImageBitmap;
    const decode = vi.fn()
      .mockImplementationOnce(() => new Promise<ImageBitmap>((resolve) => { finishBase = resolve; }))
      .mockResolvedValueOnce(recoloured);
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const release = core.acquireBitmapCacheLease(fetchImage);
    const pending = decodeRaster(
      'word/media/colour-owner-drop-race.png',
      'image/png',
      'FFFFFF',
      fetchImage,
    );
    await vi.waitFor(() => expect(decode).toHaveBeenCalledOnce());
    const rejected = expect(pending).rejects.toThrow(/cache.*dropped/i);

    finishBase(base);
    core.dropBitmapCacheByPath(fetchImage);
    await rejected;

    expect(decode).toHaveBeenCalledOnce();
    expect(recolouredClose).not.toHaveBeenCalled();
    release();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(baseClose).toHaveBeenCalledOnce();
  });

  it('does not recreate a colour-effect entry when its namespace is dropped after the base resolves', async () => {
    stubOffscreen();
    let finishBase!: (bitmap: ImageBitmap) => void;
    const baseClose = vi.fn();
    const recolouredClose = vi.fn();
    const base = { width: 2, height: 2, close: baseClose } as unknown as ImageBitmap;
    const recoloured = { width: 2, height: 2, close: recolouredClose } as unknown as ImageBitmap;
    const decode = vi.fn()
      .mockImplementationOnce(() => new Promise<ImageBitmap>((resolve) => { finishBase = resolve; }))
      .mockResolvedValueOnce(recoloured);
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([new Uint8Array([1])], { type: mime }));
    const release = core.acquireBitmapCacheLease(fetchImage);
    const pending = decodeRaster(
      'word/media/colour-namespace-drop-race.png',
      'image/png',
      'FFFFFF',
      fetchImage,
    );
    await vi.waitFor(() => expect(decode).toHaveBeenCalledOnce());
    const rejected = expect(pending).rejects.toThrow(/cache.*dropped/i);

    finishBase(base);
    dropColorReplacedCache(fetchImage);
    await rejected;

    expect(decode).toHaveBeenCalledOnce();
    expect(recolouredClose).not.toHaveBeenCalled();
    release();
    core.dropBitmapCacheByPath(fetchImage);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(baseClose).toHaveBeenCalledOnce();
  });

  it('decodeRaster applies a <a:duotone> recolour on the raster: base decode + one recolour, memoized', async () => {
    stubOffscreen();
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    try {
      const a = await decodeRaster('word/media/duo.png', 'image/png', undefined, fetchImage, 0, 0, duotone);
      // Base blob decode + the duotone OffscreenCanvas → 2 createImageBitmap calls.
      expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
      const b = await decodeRaster('word/media/duo.png', 'image/png', undefined, fetchImage, 0, 0, duotone);
      // Repeat: memoized recolour reused, base fetched once, no further decode.
      expect(b).toBe(a);
      expect(fetchImage).toHaveBeenCalledTimes(1);
      expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
    } finally {
      dropColorReplacedCache(fetchImage);
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('preloadImages keys a duotone picture separately from the raw blip (distinct map keys, shared base)', async () => {
    stubOffscreen();
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const path = 'word/media/duo2.png';
    const doc = {
      body: [
        {
          type: 'paragraph',
          runs: [
            { type: 'image', imagePath: path, mimeType: 'image/png', widthPt: 10, heightPt: 10 },
            {
              type: 'image',
              imagePath: path,
              mimeType: 'image/png',
              widthPt: 10,
              heightPt: 10,
              duotone: { clr1: '000000', clr2: 'DAB6BA' },
            },
          ],
        },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    try {
      const map = await preloadImages(doc, fetchImage);
      // Two distinct keys: the raw path + the duotone-suffixed variant.
      expect(map.has(path)).toBe(true);
      expect(map.has(`${path}|duo:000000:DAB6BA`)).toBe(true);
      // Shared base → ONE fetch for both refs.
      expect(fetchImage).toHaveBeenCalledTimes(1);
    } finally {
      dropColorReplacedCache(fetchImage);
      core.dropBitmapCacheByPath(fetchImage);
    }
  });

  it('preloadImages with no fetchImage yields an empty map (no byte source)', async () => {
    const doc = {
      body: [
        {
          type: 'paragraph',
          runs: [{ type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png', widthPt: 10, heightPt: 10 }],
        },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    const map = await preloadImages(doc, undefined);
    expect(map.size).toBe(0);
  });
});
