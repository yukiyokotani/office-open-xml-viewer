import { describe, it, expect, vi } from 'vitest';
import {
  cropSourceRect,
  drawImageCropped,
  imageNaturalSize,
  metafileRasterSize,
  sourceRasterTargetSize,
  srcRectHasVisibleArea,
} from './crop';

/**
 * The shared `<a:srcRect>` crop (ECMA-376 §20.1.8.55) used by the docx, pptx and
 * xlsx renderers. `drawImageCropped` draws only the visible source sub-rectangle
 * into the (unchanged) destination box via the 9-arg `ctx.drawImage`. The crop is
 * uniform across raster blips AND metafiles — a cropped metafile is first
 * rasterized at its FULL picture frame (`metafileRasterSize` scales the box up by
 * 1/(1−crop)), so its bitmap pixels map to the same source fractions a raster
 * blip's native grid does.
 */

/** A decoded bitmap stand-in exposing native pixel `width`/`height`. */
const fakeImg = (w: number, h: number): CanvasImageSource =>
  ({ width: w, height: h }) as unknown as CanvasImageSource;

function spyCtx(): { ctx: CanvasRenderingContext2D; drawImage: ReturnType<typeof vi.fn> } {
  const drawImage = vi.fn();
  return { ctx: { drawImage } as unknown as CanvasRenderingContext2D, drawImage };
}

describe('imageNaturalSize', () => {
  it('reads ImageBitmap-style width/height', () => {
    expect(imageNaturalSize(fakeImg(640, 480))).toEqual({ w: 640, h: 480 });
  });
  it('prefers HTMLImageElement naturalWidth/naturalHeight over CSS width/height', () => {
    const el = { naturalWidth: 800, naturalHeight: 600, width: 100, height: 50 } as unknown as CanvasImageSource;
    expect(imageNaturalSize(el)).toEqual({ w: 800, h: 600 });
  });
  it('returns 0×0 when no size is reported', () => {
    expect(imageNaturalSize({} as unknown as CanvasImageSource)).toEqual({ w: 0, h: 0 });
  });
});

describe('cropSourceRect', () => {
  it('returns null for no crop, an all-zero crop, or no native size', () => {
    expect(cropSourceRect(fakeImg(100, 100), undefined)).toBeNull();
    expect(cropSourceRect(fakeImg(100, 100), { l: 0, t: 0, r: 0, b: 0 })).toBeNull();
    expect(cropSourceRect(fakeImg(0, 0), { l: 0.1, t: 0, r: 0.1, b: 0 })).toBeNull();
  });
  it('computes the source sub-rectangle from fractional insets', () => {
    const c = cropSourceRect(fakeImg(2860, 1368), { l: 0.3256, t: 0, r: 0.03829, b: 0 });
    expect(c).not.toBeNull();
    expect(c!.sx).toBeCloseTo(0.3256 * 2860, 3);
    expect(c!.sy).toBe(0);
    expect(c!.sw).toBeCloseTo((1 - 0.3256 - 0.03829) * 2860, 3);
    expect(c!.sh).toBe(1368);
  });
  it('intersects a normative negative outset with the actual source bitmap', () => {
    const c = cropSourceRect(fakeImg(100, 100), { l: -0.5, t: 0, r: 0, b: 0 });
    expect(c).toEqual({ sx: 0, sy: 0, sw: 100, sh: 100 });
  });
});

describe('drawImageCropped', () => {
  it('draws the visible sub-rectangle (9-arg) for a horizontal raster crop', () => {
    const { ctx, drawImage } = spyCtx();
    const img = fakeImg(2860, 1368); // sample-27 PNG native pixel size
    drawImageCropped(ctx, img, { l: 0.3256, t: 0, r: 0.03829, b: 0 }, 10, 20, 305, 229);
    const call = drawImage.mock.calls[0];
    expect(call).toHaveLength(9);
    const [, sx, sy, sw, sh, dx, dy, dw, dh] = call;
    expect(sx).toBeCloseTo(0.3256 * 2860, 3);
    expect(sy).toBe(0);
    expect(sw).toBeCloseTo((1 - 0.3256 - 0.03829) * 2860, 3);
    expect(sh).toBe(1368);
    expect([dx, dy, dw, dh]).toEqual([10, 20, 305, 229]);
  });
  it('crops the vertical axis (sy/sh) for a top+bottom crop', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(400, 200), { l: 0, t: 0.1, r: 0, b: 0.25 }, 0, 0, 80, 40);
    const [, sx, , sw, sh] = drawImage.mock.calls[0];
    expect(sx).toBe(0);
    expect(sw).toBe(400);
    expect(drawImage.mock.calls[0][2]).toBeCloseTo(0.1 * 200, 6); // sy = 20
    expect(sh).toBeCloseTo((1 - 0.1 - 0.25) * 200, 6); // 130
  });
  it('draws the whole image (4-arg) with no crop or an all-zero crop', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(100, 100), undefined, 0, 0, 50, 50);
    drawImageCropped(ctx, fakeImg(100, 100), { l: 0, t: 0, r: 0, b: 0 }, 0, 0, 50, 50);
    expect(drawImage.mock.calls[0]).toHaveLength(5);
    expect(drawImage.mock.calls[1]).toHaveLength(5);
  });
  it('maps a negative left outset to transparent destination space', () => {
    const { ctx, drawImage } = spyCtx();
    drawImageCropped(ctx, fakeImg(100, 100), { l: -0.5, t: 0, r: 0, b: 0 }, 10, 20, 150, 90);
    expect(drawImage.mock.calls[0]).toEqual([
      expect.anything(), 0, 0, 100, 100,
      60, 20, 100, 90,
    ]);
  });
  it('crops a metafile too — its full-frame raster maps to the same fractions', () => {
    const { ctx, drawImage } = spyCtx();
    // A full-frame EMF raster; crop = sample-13 Fig.2 subfigure (a) insets.
    const img = fakeImg(900, 556);
    drawImageCropped(ctx, img, { l: 0.0883, t: 0.0595, r: 0.6421, b: 0.6592 }, 10, 20, 122, 78);
    const call = drawImage.mock.calls[0];
    expect(call).toHaveLength(9);
    const [, sx, sy, sw, sh] = call;
    expect(sx).toBeCloseTo(0.0883 * 900, 2);
    expect(sw).toBeCloseTo((1 - 0.0883 - 0.6421) * 900, 2);
    expect(sy).toBeCloseTo(0.0595 * 556, 2);
    expect(sh).toBeCloseTo((1 - 0.0595 - 0.6592) * 556, 2);
  });
});

describe('metafileRasterSize', () => {
  it('scales a cropped metafile up to its full picture frame', () => {
    const s = metafileRasterSize('image/emf', { l: 0.1, t: 0.2, r: 0.1, b: 0.2 }, 80, 50);
    expect(s?.widthPt).toBeCloseTo(100, 6); // 80 / (1 − 0.1 − 0.1)
    expect(s?.heightPt).toBeCloseTo(83.333, 3); // 50 / (1 − 0.2 − 0.2)
  });
  it('passes an uncropped metafile through unchanged', () => {
    expect(metafileRasterSize('image/emf', null, 80, 50)).toEqual({ widthPt: 80, heightPt: 50 });
  });
  it('passes a raster blip through unchanged even with a crop (raster decodes native)', () => {
    expect(metafileRasterSize('image/png', { l: 0.1, t: 0, r: 0.1, b: 0 }, 80, 50)).toEqual({
      widthPt: 80,
      heightPt: 50,
    });
  });
  it('fails closed instead of rasterizing a fully cropped metafile', () => {
    expect(srcRectHasVisibleArea({ l: 0.6, t: 0, r: 0.6, b: 0 })).toBe(false);
    expect(metafileRasterSize('image/wmf', { l: 0.6, t: 0, r: 0.6, b: 0 }, 80, 50)).toBeNull();
    const narrow = metafileRasterSize('image/wmf', { l: 0.499, t: 0, r: 0.5, b: 0 }, 80, 50);
    expect(narrow?.widthPt).toBeCloseTo(80 / 0.001, 6);
    expect(narrow?.heightPt).toBe(50);
  });
});

describe('sourceRasterTargetSize', () => {
  it('uses the destination device-pixel box for an uncropped raster', () => {
    expect(sourceRasterTargetSize(1200, 900)).toEqual({ width: 1200, height: 900 });
  });

  it('requests enough full-source pixels for a cropped slice to fill the destination', () => {
    expect(sourceRasterTargetSize(1200, 900, { l: 0.25, t: 0.1, r: 0.25, b: 0.1 }))
      .toEqual({ width: 2400, height: 1125 });
  });

  it('accounts for negative crop outsets without over-decoding transparent margins', () => {
    expect(sourceRasterTargetSize(1200, 900, { l: -0.1, t: 0, r: -0.1, b: 0 }))
      .toEqual({ width: 1000, height: 900 });
  });

  it('rejects non-visible and non-finite crop targets', () => {
    expect(sourceRasterTargetSize(1200, 900, { l: 0.6, t: 0, r: 0.6, b: 0 })).toBeNull();
    expect(sourceRasterTargetSize(Number.NaN, 900)).toBeNull();
  });
});
