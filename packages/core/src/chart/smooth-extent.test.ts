import { expect, test } from 'vitest';
import { smoothCurveExtent } from './smooth-extent.js';

test('a smoothed run leaving a flat stretch of zeros dips below zero', () => {
  const [lo, hi] = smoothCurveExtent([0, 0, 0, 0.1, 1, 3, 4])!;
  expect(lo).toBeLessThan(0);
  expect(hi).toBeGreaterThanOrEqual(4);
});

test('two points draw a straight segment and blanks break the run', () => {
  expect(smoothCurveExtent([0, 5])).toEqual([0, 5]);
  expect(smoothCurveExtent([null, 2, null, 3, null])).toEqual([2, 3]);
  expect(smoothCurveExtent([null, null])).toBeNull();
});
