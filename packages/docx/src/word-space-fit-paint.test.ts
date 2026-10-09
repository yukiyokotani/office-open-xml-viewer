import { describe, expect, it } from 'vitest';
import { sourceOwnedTextPlacements } from './layout/text-source-ownership.js';
import type { TextPlacement } from './layout/types.js';
import { FONTS, advancePt, retainStubParagraph } from './test-support/word-space-fit.test-support.js';

// WORD_COMPRESSED_SPACE_LINE_FIT shrinks every U+0020 on a mixed line by the
// same amount, including consecutive spaces held by one segment. The retained
// clusters and source-owned fragments must follow that shortened advance.
describe('WORD_COMPRESSED_SPACE_LINE_FIT retained geometry', () => {
  it('shares a consecutive-space reduction across its clusters and source owners', () => {
    const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
    const face = FONTS['BIZ UDGothic|700']!;
    const space = advancePt(face, ' ', 8.5);
    const natural = advancePt(face, '甲   AB', 8.5);
    // Three spaces can give up 3 × (space − 8.5/4); ask for 5.5pt of that.
    const reductionPt = 5.5;
    expect(reductionPt).toBeLessThan(3 * (space - 8.5 / 4));
    const shrunk = space - reductionPt / 3;
    for (const chunks of [['甲   AB'], ['甲  ', ' AB']]) {
      const node = retainStubParagraph({
        runs: chunks.map((text) => ({ ...run, text })),
        environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
        bandPt: natural - reductionPt,
        justification: 'left',
      });
      expect(node.lines).toHaveLength(1);
      const placements = node.lines[0]!.placements
        .filter((placement): placement is TextPlacement => placement.kind === 'text');
      const gap = placements.find((placement) => placement.text.endsWith(' '))!;
      const next = placements[placements.indexOf(gap) + 1]!;
      expect(gap.advancePt).toBeCloseTo(3 * shrunk, 9);
      expect(next.bounds.xPt).toBeCloseTo(gap.bounds.xPt + gap.advancePt, 9);
      const spaces = gap.clusters.filter((cluster) =>
        gap.text.slice(cluster.range.start - gap.range.start, cluster.range.end - gap.range.start) === ' ');
      expect(spaces).toHaveLength(3);
      for (const cluster of spaces) expect(cluster.advancePt).toBeCloseTo(shrunk, 9);
      const last = gap.clusters.at(-1)!;
      expect(last.offset.xPt + last.advancePt).toBeCloseTo(gap.advancePt, 9);
      // Every source-owned fragment stays inside the placed segment box and
      // reports only the reduction already included in its own advance.
      const owners = sourceOwnedTextPlacements(gap);
      expect(owners).toHaveLength(chunks.length);
      expect(owners.reduce((sum, owner) => sum + owner.advancePt, 0)).toBeCloseTo(gap.advancePt, 9);
      for (const owner of owners) {
        expect(owner.bounds.xPt + owner.bounds.widthPt).toBeLessThanOrEqual(next.bounds.xPt + 1e-9);
        const ownedSpaces = owner.text.length - owner.text.trimEnd().length;
        expect(owner.trailingSpaceCompressionPt).toBeCloseTo(ownedSpaces * (space - shrunk), 9);
      }
    }
  });
});
