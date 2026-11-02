import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { textAdvanceWidth } from './line-breaker/advance.js';
import { recordSpacingFixture, spacingFixtures } from './test-support/spacing-arithmetic.test-support.js';
const mainStreams = JSON.parse(readFileSync(
  new URL('./spacing-arithmetic-main.json', import.meta.url), 'utf8',
)) as Record<string, unknown>;

// Follow word-space-fit-parity.test.ts's captured main-oracle data pattern.
// Captured with this recording canvas at 4bcb91c487e38393aadb3f3b70a732adf2c7b708;
// these are previous-renderer results, never Office fidelity references.
// The fixture now declares Arial present (absent faces size lines from their
// painting fallback), so only the recorded font strings gained the local
// route's resolved family; every geometry value is unchanged.
// 'review-justified' was re-recorded when WORD_LINE_END_SPACE_HANGS took the
// line-end space out of justification slack: only its gap offsets changed.
// 'kashida' (highKashida) was re-recorded for WORD_KASHIDA_LINE_FIT: its
// over-wide Latin run now force-breaks at 0.8 of the width; every advance,
// transform and fill keeps its arithmetic.
// All numeric fields are unrounded doubles, including retained cluster advances
// and offsets, every transform argument, and every fill's complete paint state.
describe('unexcluded text keeps main spacing arithmetic bit for bit', () => {
  it.each(spacingFixtures)('$id paint stream and advances', async fixture => {
    expect(await recordSpacingFixture(fixture)).toEqual(mainStreams[fixture.id]);
  });

  it('adds grid and tracking to the scaled natural width in main order', () => {
    for (const text of ['AV To WAVE AVATAR', '甲乙😀', 'سلام']) {
      const n = [...text].length;
      for (const [natural, grid, scale, tracking] of [
        [143.125, 1200 / 4096, 0.66, 0.35],
        [77.33333333333333, -0.1, 1.33, -0.15],
        [175.7744903564453, 0.1, 0.66, 0.35],
      ]) {
        expect(textAdvanceWidth(natural, text, grid, scale, tracking)).toBe(
          natural * scale + n * grid + n * tracking,
        );
      }
    }
  });
});
