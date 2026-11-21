import { describe, expect, it } from 'vitest';
import type { PhysicalAnchorFrame } from './acquisition-context.js';
import { logicalToPhysicalMatrix, physicalToLogicalMatrix, transformRect } from './coordinate-space.js';
import type { WritingMode } from './types.js';
import {
  physicalAnchorBoxToLogical,
  physicalDistToLogical,
  physicalInlineBandTopPt,
  uprightBlockPhysicalOrigin,
  uprightPhysicalRectToLogical,
} from './upright-physical-frame.js';

const frame = (writingMode: WritingMode): PhysicalAnchorFrame => ({
  pageWidth: 842,
  pageHeight: 595,
  marginLeft: 56.7,
  marginRight: 56.7,
  marginTop: 70.9,
  marginBottom: 70.9,
  writingMode,
});
const page = { widthPt: 842, heightPt: 595 };
const anchor = { x: 444.3, y: 397.85, w: 96.2, h: 123 };

describe('upright physical frame projections', () => {
  // Extents are carried, never recovered by differencing transformed corners,
  // and the clockwise and counter-clockwise frames keep their established
  // arithmetic order: these values are exact.
  it('projects an anchor box exactly in the clockwise and counter-clockwise frames', () => {
    expect(physicalAnchorBoxToLogical(frame('vertical-rl'), anchor)).toEqual({
      x: 397.85, y: 842 - (444.3 + 96.2), w: 123, h: 96.2,
    });
    expect(physicalAnchorBoxToLogical(frame('vertical-rl'), anchor).h).toBe(96.2);
    expect(physicalAnchorBoxToLogical(frame('sideways-lr'), anchor)).toEqual({
      x: 595 - (397.85 + 123), y: 444.3, w: 123, h: 96.2,
    });
  });

  it('places an upright block exactly in the clockwise and counter-clockwise frames', () => {
    const cursor = { xPt: 70.9, yPt: 56.7 };
    expect(uprightBlockPhysicalOrigin(frame('vertical-rl'), cursor, 96.2, 123))
      .toEqual({ xPt: 689.0999999999999, yPt: 70.9 });
    expect(uprightBlockPhysicalOrigin(frame('sideways-lr'), cursor, 96.2, 123))
      .toEqual({ xPt: 56.7, yPt: 595 - 70.9 - 123 });
  });

  it('projects an upright rectangle and a column top exactly as before', () => {
    const rect = { xPt: 444.3, yPt: 397.85, widthPt: 96.2, heightPt: 123 };
    expect(uprightPhysicalRectToLogical(frame('vertical-rl'), rect)).toEqual({
      xPt: 397.85, yPt: 842 - 444.3 - 96.2, widthPt: 123, heightPt: 96.2,
    });
    expect(uprightPhysicalRectToLogical(frame('sideways-lr'), rect)).toEqual(
      transformRect(physicalToLogicalMatrix('sideways-lr', page), rect),
    );
    expect(physicalInlineBandTopPt(frame('vertical-rl'), 70.9, 453.2)).toBe(70.9);
    expect(physicalInlineBandTopPt(frame('sideways-lr'), 70.9, 453.2)).toBe(595 - (70.9 + 453.2));
  });

  it.each(['vertical-rl', 'vertical-lr', 'sideways-lr'] as const)(
    'agrees with the %s paint matrix',
    (writingMode) => {
      const toLogical = physicalToLogicalMatrix(writingMode, page);
      const toPhysical = logicalToPhysicalMatrix(writingMode, page);
      const box = physicalAnchorBoxToLogical(frame(writingMode), anchor);
      const expected = transformRect(toLogical, {
        xPt: anchor.x, yPt: anchor.y, widthPt: anchor.w, heightPt: anchor.h,
      });
      expect(box.x).toBeCloseTo(expected.xPt, 9);
      expect(box.y).toBeCloseTo(expected.yPt, 9);
      expect(box.w).toBeCloseTo(expected.widthPt, 9);
      expect(box.h).toBeCloseTo(expected.heightPt, 9);

      const cursor = { xPt: 70.9, yPt: 56.7 };
      const origin = uprightBlockPhysicalOrigin(frame(writingMode), cursor, 96.2, 123);
      const block = transformRect(toPhysical, { ...cursor, widthPt: 123, heightPt: 96.2 });
      expect(origin.xPt).toBeCloseTo(block.xPt, 9);
      expect(origin.yPt).toBeCloseTo(block.yPt, 9);

      const band = transformRect(toPhysical, { xPt: 70.9, yPt: 0, widthPt: 453.2, heightPt: 0 });
      expect(physicalInlineBandTopPt(frame(writingMode), 70.9, 453.2)).toBeCloseTo(band.yPt, 9);
    },
  );

  it('relabels dist* padding with the edges of each frame', () => {
    const dist = { dl: 1, dr: 2, dt: 3, db: 4 };
    expect(physicalDistToLogical(frame('vertical-rl'), dist)).toEqual({ dl: 3, dr: 4, dt: 2, db: 1 });
    expect(physicalDistToLogical(frame('vertical-lr'), dist)).toEqual({ dl: 3, dr: 4, dt: 1, db: 2 });
    expect(physicalDistToLogical(frame('sideways-lr'), dist)).toEqual({ dl: 4, dr: 3, dt: 1, db: 2 });
  });
});
