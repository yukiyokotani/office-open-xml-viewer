import { describe, expect, it } from 'vitest';
import { resolveLineFloatWindow, createFloatLineBandPredicate, type FloatRect } from './float-wrap.js';

// An anchor line exempt from its own topAndBottom float (exemptLineTopPt, set
// by WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL when the float only reaches the
// line's trailing grid leading) keeps its top; other line tops still clear it.
const float = (exemptLineTopPt?: number): FloatRect => ({
  kind: 'shape', mode: 'topAndBottom', authoredWrap: 'topAndBottom', imageKey: 'test',
  imageX: 0, imageY: 75.7, imageW: 106.3, imageH: 21,
  xLeft: 0, xRight: 106.3, yTop: 75.7, yBottom: 96.7, side: 'bothSides',
  distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
  ...(exemptLineTopPt === undefined ? {} : { exemptLineTopPt }),
});

describe('topAndBottom anchor-line exemption', () => {
  it('pushes an overlapping line below the float without an exemption', () => {
    expect(resolveLineFloatWindow(56.7, 10, 20, 0, 400, [float()]).topY).toBeCloseTo(96.7, 10);
    expect(createFloatLineBandPredicate(float())(56.7, 20)).toBe(true);
  });
  it('keeps the exempt anchor line at its top', () => {
    expect(resolveLineFloatWindow(56.7, 10, 20, 0, 400, [float(56.7)]).topY).toBeCloseTo(56.7, 10);
    expect(createFloatLineBandPredicate(float(56.7))(56.7, 20)).toBe(false);
    // The next line still clears the float.
    expect(resolveLineFloatWindow(76.7, 10, 20, 0, 400, [float(56.7)]).topY).toBeCloseTo(96.7, 10);
  });
});
