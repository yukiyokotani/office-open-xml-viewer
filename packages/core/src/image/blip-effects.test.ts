import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  applyBlipPixelEffects,
  assertBlipPixelEffectsBudget,
  blipGrayLevel,
  blipLuminance,
  type BlipEffect,
  type BlipPixelEffects,
} from './blip-effects';
import {
  MAX_IMAGE_EFFECT_BASE_PIXELS,
  MAX_IMAGE_EFFECT_PASSES,
  MAX_IMAGE_EFFECT_PIXEL_WORK,
  isOoxmlDecodedImageLimitError,
} from './pixel-budget';
import {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './duotone-bitmap-by-path';
import { dropBitmapCacheByPath } from './bitmap-image-by-path';
import type { OffscreenFactory } from './duotone';

const buffer = (...pixels: number[][]) => ({
  data: new Uint8ClampedArray(pixels.flat()),
  width: pixels.length,
  height: 1,
});

describe('applyBlipPixelEffects (ECMA-376 §20.1.8.11/16/34)', () => {
  it('grayscl + biLevel(50%) thresholds the Rec. 709 luma PowerPoint uses', () => {
    // PowerPoint PDF evidence: (2,167,223) → white; (0,147,190), (0,126,229)
    // and (9,74,178) → black under its "Black and White" picture setting.
    const buf = buffer([2, 167, 223, 255], [0, 147, 190, 255], [0, 126, 229, 128], [9, 74, 178, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }] });
    expect([...buf.data]).toEqual([255, 255, 255, 255, 0, 0, 0, 255, 0, 0, 0, 128, 0, 0, 0, 255]);
  });

  it('grayscl keeps alpha and writes the luma to every channel', () => {
    const buf = buffer([255, 0, 0, 7]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }] });
    const gray = Math.floor(blipLuminance(255, 0, 0) * 255);
    expect([...buf.data]).toEqual([gray, gray, gray, 7]);
  });

  it('grayscl truncates the luma to a whole level, as the PowerPoint boundary evidence shows', () => {
    // Luma 127.59 must become 127 (black under biLevel 50%), 128.02 stays 128.
    expect(blipGrayLevel(99, 131, 178)).toBe(127);
    expect(blipGrayLevel(98, 132, 177)).toBe(128);
    expect(blipGrayLevel(255, 255, 255)).toBe(255);
    const buf = buffer([99, 131, 178, 255], [98, 132, 177, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }] });
    expect([...buf.data]).toEqual([0, 0, 0, 255, 255, 255, 255, 255]);
  });

  it('clrChange replaces exact RGB matches with the target colour and alpha', () => {
    const buf = buffer([255, 255, 255, 255], [254, 255, 255, 255], [255, 255, 255, 10]);
    applyBlipPixelEffects(buf, {
      effects: [{ type: 'colorChange', from: 'FFFFFF', fromAlpha: 1, to: 'FFFFFF', toAlpha: 0, useAlpha: false }],
    });
    expect([...buf.data]).toEqual([255, 255, 255, 0, 254, 255, 255, 255, 255, 255, 255, 0]);
    const strict = buffer([255, 255, 255, 255], [255, 255, 255, 10]);
    applyBlipPixelEffects(strict, {
      effects: [{ type: 'colorChange', from: 'FFFFFF', fromAlpha: 1, to: '000000', toAlpha: 0.5, useAlpha: true }],
    });
    expect([...strict.data]).toEqual([0, 0, 0, 128, 255, 255, 255, 10]);
  });

  it('applies effects in document order, with the duotone at its marker', () => {
    const duotone = { clr1: '000000', clr2: 'FF0000' };
    // Duotone first maps white to red; grayscale after it grays the red.
    const first = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(first, { effects: [{ type: 'duotone' }, { type: 'grayscale' }], duotone });
    expect(first.data[0]).toBe(first.data[1]);
    // Grayscale first leaves white, which the later duotone maps to red.
    const second = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(second, { effects: [{ type: 'grayscale' }, { type: 'duotone' }], duotone });
    expect([...second.data]).toEqual([255, 0, 0, 255]);
    // Without a marker the duotone applies after the listed effects.
    const unmarked = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(unmarked, { effects: [{ type: 'grayscale' }], duotone });
    expect([...unmarked.data]).toEqual([255, 0, 0, 255]);
  });

  // Levels read from PowerPoint's PDF export of a 256-step gray ramp under
  // <a:lum> (inputs 0, 32, 64, 96, 128, 160, 192, 224, 255).
  it('lum follows the brightness/contrast levels PowerPoint renders', () => {
    const inputs = [0, 32, 64, 96, 128, 160, 192, 224, 255];
    const levels = (bright: number, contrast: number) => {
      const buf = buffer(inputs.flatMap((v) => [v, v, v, 200]));
      applyBlipPixelEffects(buf, { effects: [{ type: 'luminance', bright, contrast }] });
      expect([...buf.data].filter((_, i) => i % 4 === 3).every((a) => a === 200)).toBe(true);
      return [...buf.data].filter((_, i) => i % 4 === 0);
    };
    const near = (actual: number[], office: number[]) => actual.forEach((value, i) => {
      expect(Math.abs(value - office[i])).toBeLessThanOrEqual(1);
    });
    near(levels(0, -0.7), [89, 99, 108, 118, 128, 137, 147, 156, 166]);
    near(levels(0, -0.35), [45, 65, 86, 107, 128, 149, 169, 190, 210]);
    near(levels(0.35, 0), [89, 121, 153, 185, 217, 249, 255, 255, 255]);
    near(levels(0.7, -0.7), [205, 215, 224, 234, 244, 253, 255, 255, 255]);
    expect(levels(1, 0)).toEqual(Array(9).fill(255));
    expect(levels(-1, 0)).toEqual(Array(9).fill(0));
    expect(levels(0, -1)).toEqual(Array(9).fill(128));
    expect(levels(0, 1)).toEqual([0, 0, 0, 0, 255, 255, 255, 255, 255]);
  });
});

describe('getCachedDuotoneBitmapByPath with CT_Blip effects', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('runs the ordered effects once through the shared decode cache, keyed apart from the duotone', async () => {
    const base = { width: 1, height: 1, close() {} } as unknown as ImageBitmap;
    const transformed = { width: 1, height: 1, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (src: unknown) => (src instanceof Blob ? base : transformed));
    vi.stubGlobal('createImageBitmap', cib);
    const written: Uint8ClampedArray[] = [];
    const factory = ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext() {
        return {
          drawImage() {},
          getImageData() {
            return { data: new Uint8ClampedArray([2, 167, 223, 255]), width: 1, height: 1 } as unknown as ImageData;
          },
          putImageData(img: ImageData) {
            written.push(img.data);
          },
        };
      },
    })) as unknown as OffscreenFactory;
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const effects: BlipPixelEffects = {
      effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }],
      duotone: null,
    };
    const path = 'ppt/media/blip-effects.png';
    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', effects, fetchImage, { offscreenFactory: factory });
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', effects, fetchImage, { offscreenFactory: factory });
    expect(first).toBe(transformed);
    expect(second).toBe(transformed);
    expect(written).toHaveLength(1);
    expect([...written[0]]).toEqual([255, 255, 255, 255]);
    expect(duotoneCacheKey(path, effects)).toBe(`${path}|fx:g,b0.5`);
    expect(duotoneCacheKey(path, { clr1: '000000', clr2: 'FFFFFF' })).toBe(`${path}|duo:000000:FFFFFF`);
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });
});

describe('CT_Blip effect workload budget (shared resource policy)', () => {
  const grays = (count: number): BlipEffect[] => Array.from({ length: count }, () => ({ type: 'grayscale' }));
  const limitOf = (run: () => unknown) => {
    try {
      run();
    } catch (error) {
      if (!isOoxmlDecodedImageLimitError(error)) throw error;
      return { metric: error.metric, limit: error.limit, observed: error.observed };
    }
    return undefined;
  };

  it('derives the pass limit from the work ceiling at the largest effect base', () => {
    expect(MAX_IMAGE_EFFECT_PASSES).toBe(16);
    expect(MAX_IMAGE_EFFECT_PASSES * MAX_IMAGE_EFFECT_BASE_PIXELS).toBe(MAX_IMAGE_EFFECT_PIXEL_WORK);
  });

  it('admits the pass limit and rejects one more pass before touching a pixel', () => {
    const atLimit = buffer([255, 0, 0, 9]);
    applyBlipPixelEffects(atLimit, { effects: grays(MAX_IMAGE_EFFECT_PASSES) });
    expect(atLimit.data[0]).toBe(atLimit.data[1]);

    const over = buffer([255, 0, 0, 9]);
    expect(limitOf(() => applyBlipPixelEffects(over, { effects: grays(MAX_IMAGE_EFFECT_PASSES + 1) })))
      .toEqual({ metric: 'image-effect-count', limit: 16, observed: 17 });
    expect([...over.data]).toEqual([255, 0, 0, 9]);
    // A duotone without a position marker is one more pass.
    expect(limitOf(() => applyBlipPixelEffects(buffer([255, 0, 0, 9]), {
      effects: grays(MAX_IMAGE_EFFECT_PASSES),
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    }))).toEqual({ metric: 'image-effect-count', limit: 16, observed: 17 });
  });

  it('admits cumulative work at the ceiling and rejects one pixel visit more', () => {
    const two = { effects: grays(2) };
    const half = MAX_IMAGE_EFFECT_PIXEL_WORK / 2;
    expect(limitOf(() => assertBlipPixelEffectsBudget(two, half))).toBeUndefined();
    expect(limitOf(() => assertBlipPixelEffectsBudget(two, half + 1)))
      .toEqual({ metric: 'image-effect-work', limit: MAX_IMAGE_EFFECT_PIXEL_WORK, observed: 2 * (half + 1) });
    // The applied transform checks the grid it would scan before reading it:
    // this buffer has no pixel storage at all, only its length.
    const unread = { data: { length: (half + 1) * 4 } as unknown as Uint8ClampedArray, width: half + 1, height: 1 };
    expect(limitOf(() => applyBlipPixelEffects(unread, two))?.metric).toBe('image-effect-work');
  });
});

describe('getCachedDuotoneBitmapByPath effect budget', () => {
  afterEach(() => vi.unstubAllGlobals());

  const factory = ((w: number, h: number) => ({
    width: w,
    height: h,
    getContext() {
      return {
        drawImage() {},
        getImageData() {
          return { data: new Uint8ClampedArray([255, 0, 0, 255]), width: 1, height: 1 } as unknown as ImageData;
        },
        putImageData() {},
      };
    },
  })) as unknown as OffscreenFactory;
  const grays = (count: number): BlipEffect[] => Array.from({ length: count }, () => ({ type: 'grayscale' }));

  it('rejects an over-long effect list before fetching or decoding the blip', async () => {
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const rejected = getCachedDuotoneBitmapByPath(
      'ppt/media/too-many-effects.png',
      'image/png',
      { effects: grays(MAX_IMAGE_EFFECT_PASSES + 1), duotone: null },
      fetchImage,
      { offscreenFactory: factory },
    );
    await expect(rejected).rejects.toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-count' });
    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('runs the pass limit on the largest effect base (work exactly at the ceiling)', async () => {
    // 4096 × 2048 = MAX_IMAGE_EFFECT_BASE_PIXELS; 16 passes = 2^27 visits.
    const base = { width: 4096, height: 2048, close() {} } as unknown as ImageBitmap;
    const transformed = { width: 4096, height: 2048, close() {} } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: unknown) => (src instanceof Blob ? base : transformed)));
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const result = await getCachedDuotoneBitmapByPath(
      'ppt/media/largest-effect-base.png',
      'image/png',
      { effects: grays(MAX_IMAGE_EFFECT_PASSES), duotone: null },
      fetchImage,
      { offscreenFactory: factory },
    );
    expect(result).toBe(transformed);
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });
});
