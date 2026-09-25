import { describe, it, expect, vi, afterEach } from 'vitest';
import { applyBlipPixelEffects, blipGrayLevel, blipLuminance, type BlipPixelEffects } from './blip-effects';
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
