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
