import { describe, expect, it } from 'vitest';
import { correctLineMetrics, fontWinLineHeightRatio, intendedSingleLinePx } from './line-metrics.js';

describe('legacy family-keyed metric API', () => {
  it('does not give an authored family metric authority over the measured face', () => {
    for (const family of ['Meiryo', 'Times New Roman', 'Sakkal Majalla', 'unknown']) {
      expect(fontWinLineHeightRatio(family, true)).toBeNull();
      expect(intendedSingleLinePx(family, 16, true)).toBe(0);
      expect(correctLineMetrics(family, 16, 18, 8, true)).toEqual({ ascent: 18, descent: 8 });
    }
  });
});
