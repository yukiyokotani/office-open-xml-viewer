import { describe, it, expect } from 'vitest';
import { tileAnchorOffset, tileSourceExtent } from './renderer';

/**
 * ECMA-376 §20.1.8.58 (CT_TileInfoProperties) `algn` + §20.1.8.41
 * (ST_RectAlignment). The tile grid registers against a corner / edge / centre
 * of the fill box; `tileAnchorOffset` returns the top-left position (px) of the
 * tile that sits at that anchor, before the tx/ty offset is added.
 *
 * Box 100×100, tile 40×30 — distinct W/H so axis mix-ups surface.
 */
describe('tileAnchorOffset — §20.1.8.41 ST_RectAlignment', () => {
  const W = 100;
  const H = 100;
  const TW = 40;
  const TH = 30;

  // Horizontal: left = 0, centre = (W-TW)/2 = 30, right = W-TW = 60.
  // Vertical:   top  = 0, middle = (H-TH)/2 = 35, bottom = H-TH = 70.
  const cases: Array<[string, number, number]> = [
    ['tl', 0, 0],
    ['t', 30, 0],
    ['tr', 60, 0],
    ['l', 0, 35],
    ['ctr', 30, 35],
    ['r', 60, 35],
    ['bl', 0, 70],
    ['b', 30, 70],
    ['br', 60, 70],
  ];

  for (const [algn, ax, ay] of cases) {
    it(`algn="${algn}" anchors at (${ax}, ${ay})`, () => {
      const got = tileAnchorOffset(algn, W, H, TW, TH);
      expect(got.ax).toBeCloseTo(ax, 6);
      expect(got.ay).toBeCloseTo(ay, 6);
    });
  }

  it('defaults an unknown algn to top-left (tl)', () => {
    // §20.1.8.58 gives no schema default; PowerPoint treats absent/unknown as
    // tl, which the parser already normalises to "tl".
    expect(tileAnchorOffset('???', W, H, TW, TH)).toEqual({ ax: 0, ay: 0 });
  });
});

describe('tile source rectangle — §20.1.8.55', () => {
  it('duplicates the cropped logical window at its cropped native extent', () => {
    expect(tileSourceExtent(800, 600, { l: 0.25, t: 0.1, r: 0.15, b: 0.2 }))
      .toEqual({ width: 480, height: 420 });
  });

  it('retains negative source outsets as transparent tile extent', () => {
    expect(tileSourceExtent(800, 600, { l: -0.1, t: 0, r: -0.15, b: 0 }))
      .toEqual({ width: 1_000, height: 600 });
  });
});
