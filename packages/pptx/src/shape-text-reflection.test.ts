import { describe, expect, it } from 'vitest';
import { reflectedShapeTextRotation } from './renderer.js';

describe('reflected shape text rotation', () => {
  it.each([
    { rotation: 180, flipH: false, flipV: true, expected: 0 },
    { rotation: 180, flipH: true, flipV: false, expected: 0 },
    { rotation: 0, flipH: false, flipV: true, expected: 0 },
    { rotation: 90, flipH: true, flipV: false, expected: 90 },
    { rotation: 270, flipH: false, flipV: true, expected: -90 },
  ])(
    'uses the readable equivalent for rotation=$rotation flipH=$flipH flipV=$flipV',
    ({ rotation, flipH, flipV, expected }) => {
      expect(reflectedShapeTextRotation(rotation, flipH, flipV)).toBe(expected);
    },
  );

  it.each([
    { rotation: 180, flipH: false, flipV: false },
    { rotation: 180, flipH: true, flipV: true },
  ])(
    'leaves non-reflected rotation unchanged for flipH=$flipH flipV=$flipV',
    ({ rotation, flipH, flipV }) => {
      expect(reflectedShapeTextRotation(rotation, flipH, flipV)).toBe(rotation);
    },
  );
});
