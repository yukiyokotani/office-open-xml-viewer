import { describe, expect, it } from 'vitest';
import {
  createSectionRegionCoordinateSpace,
  logicalPageExtent,
  logicalToPhysicalMatrix,
  physicalToLogicalMatrix,
  sectionWritingMode,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
  transformPoint,
  transformRect,
  transformRectEdges,
} from './coordinate-space.js';
import type { LayoutRect, PointPt, WritingMode } from './types.js';

const page = { widthPt: 612, heightPt: 792 };
const point: PointPt = { xPt: 80, yPt: 120 };
const rect: LayoutRect = { xPt: 80, yPt: 120, widthPt: 140, heightPt: 60 };

describe('coordinate-space', () => {
  it.each<readonly [WritingMode, PointPt, LayoutRect]>([
    ['horizontal-tb', { xPt: 80, yPt: 120 }, rect],
    ['vertical-rl', { xPt: 492, yPt: 80 }, { xPt: 432, yPt: 80, widthPt: 60, heightPt: 140 }],
    ['vertical-lr', { xPt: 120, yPt: 80 }, { xPt: 120, yPt: 80, widthPt: 60, heightPt: 140 }],
    // Counter-clockwise: inline +x runs up from the bottom edge, block +y right.
    ['sideways-lr', { xPt: 120, yPt: 712 }, { xPt: 120, yPt: 572, widthPt: 60, heightPt: 140 }],
  ])('maps %s logical points and rectangles to upright physical space', (mode, expectedPoint, expectedRect) => {
    const matrix = logicalToPhysicalMatrix(mode, page);

    expect(transformPoint(matrix, point)).toEqual(expectedPoint);
    expect(transformRect(matrix, rect)).toEqual(expectedRect);
  });

  it.each<WritingMode>(['horizontal-tb', 'vertical-rl', 'vertical-lr', 'sideways-lr'])
    ('uses an exact closed-form inverse for %s points and rectangles', (mode) => {
      const forward = logicalToPhysicalMatrix(mode, page);
      const inverse = physicalToLogicalMatrix(mode, page);

      expect(transformPoint(inverse, transformPoint(forward, point))).toEqual(point);
      expect(transformRect(inverse, transformRect(forward, rect))).toEqual(rect);
      expect(createSectionRegionCoordinateSpace(mode, page)).toEqual({
        writingMode: mode,
        logicalToPhysical: forward,
        physicalToLogical: inverse,
      });
    });

  it.each([
    ['horizontal-tb', { top: 'top', right: 'right', bottom: 'bottom', left: 'left' }],
    ['vertical-rl', { top: 'right', right: 'bottom', bottom: 'left', left: 'top' }],
    ['vertical-lr', { top: 'left', right: 'bottom', bottom: 'right', left: 'top' }],
    ['sideways-lr', { top: 'left', right: 'top', bottom: 'right', left: 'bottom' }],
  ] as const)('projects physical edge ownership through %s', (mode, expected) => {
    expect(transformRectEdges(
      physicalToLogicalMatrix(mode, page),
      { top: 'top', right: 'right', bottom: 'bottom', left: 'left' },
    )).toEqual(expected);
  });

  it('derives vertical-rl translation from each physical page width', () => {
    expect(logicalToPhysicalMatrix('vertical-rl', { widthPt: 500, heightPt: 700 }).e).toBe(500);
    expect(logicalToPhysicalMatrix('vertical-rl', { widthPt: 840, heightPt: 600 }).e).toBe(840);
  });

  it.each([
    ['tb', 'horizontal-tb'],
    ['tbV', 'horizontal-tb'],
    ['lrTb', 'horizontal-tb'],
    ['lrTbV', 'horizontal-tb'],
    ['rl', 'vertical-rl'],
    ['rlV', 'vertical-rl'],
    ['tbRl', 'vertical-rl'],
    ['tbRlV', 'vertical-rl'],
    ['lr', 'vertical-lr'],
    ['lrV', 'vertical-lr'],
    ['tbLrV', 'vertical-lr'],
    ['btLr', 'vertical-rl'],
  ] as const)('maps Transitional text direction %s to %s', (token, expected) => {
    expect(writingModeFromTextDirection(token)).toBe(expected);
  });

  it.each(['', 'unknown', 'sideways-lr', 'bottomToTop'])('rejects unsupported text direction %j', (token) => {
    expect(() => writingModeFromTextDirection(token)).toThrow(RangeError);
    // An authored token never selects the native counter-clockwise frame.
    expect(() => sectionWritingMode({ textDirection: token })).toThrow(RangeError);
  });

  it('selects the counter-clockwise frame only from a section flow fact on its nominal token', () => {
    expect(sectionWritingMode({ textDirection: 'btLr' })).toBe('vertical-rl');
    expect(sectionWritingMode({ textDirection: 'tbRl' })).toBe('vertical-rl');
    expect(sectionWritingMode({ textDirection: 'btLr', sectionFlow: 'bottomToTop' }))
      .toBe('sideways-lr');
    for (const textDirection of ['lrTb', 'tbRl', 'tbLrV']) {
      expect(() => sectionWritingMode({ textDirection, sectionFlow: 'bottomToTop' }))
        .toThrow(RangeError);
    }
  });

  it('places the counter-clockwise inline start at the physical bottom and block start at the left', () => {
    const forward = logicalToPhysicalMatrix('sideways-lr', page);
    expect(forward).toEqual({ a: 0, b: -1, c: 1, d: 0, e: 0, f: 792 });
    expect(physicalToLogicalMatrix('sideways-lr', page)).toEqual({ a: 0, b: 1, c: -1, d: 0, e: 792, f: 0 });
    // Logical origin: physical bottom-left corner. Inline advance: upward.
    // Block advance: rightward.
    expect(transformPoint(forward, { xPt: 0, yPt: 0 })).toEqual({ xPt: 0, yPt: 792 });
    expect(transformPoint(forward, { xPt: 10, yPt: 0 })).toEqual({ xPt: 0, yPt: 782 });
    expect(transformPoint(forward, { xPt: 0, yPt: 10 })).toEqual({ xPt: 10, yPt: 792 });
    expect(logicalPageExtent(page, 'sideways-lr')).toEqual({ widthPt: 792, heightPt: 612 });
    expect(uprightPhysicalExtent({ widthPt: 792, heightPt: 612 }, 'sideways-lr')).toEqual(page);
  });

  it('derives logical and upright extents without conflating vertical axes', () => {
    expect(logicalPageExtent(page, 'horizontal-tb')).toEqual(page);
    expect(logicalPageExtent(page, 'vertical-rl')).toEqual({ widthPt: 792, heightPt: 612 });
    expect(logicalPageExtent(page, 'vertical-lr')).toEqual({ widthPt: 792, heightPt: 612 });
    expect(uprightPhysicalExtent({ widthPt: 792, heightPt: 612 }, 'vertical-rl'))
      .toEqual(page);
    expect(uprightPhysicalExtent(page, 'horizontal-tb')).toEqual(page);
  });

  it.each([
    { widthPt: 0, heightPt: 792 },
    { widthPt: -1, heightPt: 792 },
    { widthPt: Number.NaN, heightPt: 792 },
    { widthPt: 612, heightPt: Number.POSITIVE_INFINITY },
  ])('rejects invalid physical page extent $widthPt x $heightPt', (invalidPage) => {
    expect(() => logicalToPhysicalMatrix('horizontal-tb', invalidPage)).toThrow(RangeError);
    expect(() => physicalToLogicalMatrix('vertical-rl', invalidPage)).toThrow(RangeError);
  });

  it('rejects non-finite points and invalid rectangles', () => {
    const identity = logicalToPhysicalMatrix('horizontal-tb', page);

    expect(() => transformPoint(identity, { xPt: Number.NaN, yPt: 0 })).toThrow(RangeError);
    expect(() => transformRect(identity, { ...rect, widthPt: -1 })).toThrow(RangeError);
    expect(() => transformRect(identity, { ...rect, heightPt: Number.POSITIVE_INFINITY }))
      .toThrow(RangeError);
  });

  it.each(['a', 'b', 'c', 'd', 'e', 'f'] as const)
    ('rejects a non-finite %s matrix coefficient', (coefficient) => {
      const matrix = {
        a: 1, b: 0, c: 0, d: 1, e: 0, f: 0,
        [coefficient]: Number.NaN,
      };

      expect(() => transformPoint(matrix, point)).toThrow(RangeError);
      expect(() => transformRect(matrix, rect)).toThrow(RangeError);
    });
});
