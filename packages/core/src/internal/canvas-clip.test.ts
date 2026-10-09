import { describe, expect, it } from 'vitest';
import { clipCanvasHorizontally } from './canvas-clip.js';

// Record the clipping region sent to Canvas. The contract is geometry, not
// glyph pixels; extremely scaled matrices exceed native Skia's scalar domain.
function clipRegion(a: number, d: number, f = 0) {
  const regions: number[][] = [];
  const ctx = {
    canvas: { width: 240, height: 240 },
    getTransform: () => ({ a, b: 0, c: 0, d, e: 0, f }),
    beginPath() {}, rect: (...region: number[]) => regions.push(region), clip() {},
  } as unknown as CanvasRenderingContext2D;
  clipCanvasHorizontally(ctx, 40, 60);
  return regions;
}

describe('horizontal Canvas clip numeric domain', () => {
  it.each([1e160, 1e-200])('preserves a representable viewport at scale %s', (scale) => {
    // A determinant computed directly overflows/underflows although the
    // inverse viewport height 240 / scale is representable and nonzero.
    const regions = clipRegion(scale, scale);
    expect(regions).toHaveLength(1);
    expect(regions[0].slice(0, 3)).toEqual([40, 0, 60]);
    expect(regions[0][3]).toBe(240 / scale);
  });

  it('leaves a genuinely singular transform unchanged', () => {
    expect(clipRegion(0, 1)).toEqual([]);
  });

  it('declines unrepresentable projected corners or interval height', () => {
    expect(clipRegion(1, 1e-310)).toEqual([]);
    // Both corner y values are finite, but their difference exceeds Number.
    expect(clipRegion(1, 1e-306, 120)).toEqual([]);
  });
});
