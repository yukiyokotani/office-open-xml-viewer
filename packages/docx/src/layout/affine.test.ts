import { describe, expect, it } from 'vitest';
import {
  canonicalLogicalToPhysical,
  mapAffinePoint,
  mapAffineRect,
} from './affine.js';

describe('canonical logical/physical affine contract', () => {
  it.each([
    ['horizontal-tb', { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }, [
      { xPt: 17, yPt: 31 }, { xPt: 60, yPt: 31 },
      { xPt: 17, yPt: 50 }, { xPt: 60, yPt: 50 },
    ], { xPt: 17, yPt: 31, widthPt: 43, heightPt: 19 }],
    ['vertical-rl', { a: 0, b: 1, c: -1, d: 0, e: 641, f: 0 }, [
      { xPt: 610, yPt: 17 }, { xPt: 610, yPt: 60 },
      { xPt: 591, yPt: 17 }, { xPt: 591, yPt: 60 },
    ], { xPt: 591, yPt: 17, widthPt: 19, heightPt: 43 }],
    ['vertical-lr', { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 }, [
      { xPt: 31, yPt: 17 }, { xPt: 31, yPt: 60 },
      { xPt: 50, yPt: 17 }, { xPt: 50, yPt: 60 },
    ], { xPt: 31, yPt: 17, widthPt: 19, heightPt: 43 }],
  ] as const)(
    'maps every corner and AABB for %s',
    (writingMode, matrix, corners, bounds) => {
      const actual = canonicalLogicalToPhysical(writingMode, 641);
      expect(actual).toEqual(matrix);
      expect([
        mapAffinePoint(actual, { xPt: 17, yPt: 31 }),
        mapAffinePoint(actual, { xPt: 60, yPt: 31 }),
        mapAffinePoint(actual, { xPt: 17, yPt: 50 }),
        mapAffinePoint(actual, { xPt: 60, yPt: 50 }),
      ]).toEqual(corners);
      expect(mapAffineRect(actual, {
        xPt: 17, yPt: 31, widthPt: 43, heightPt: 19,
      })).toEqual(bounds);
    },
  );
});
