import { describe, expect, it } from 'vitest';
import { resolveDocxCommentThreads } from './comments.js';
import { intersect, pageBoxOf, paintedDocument } from './test-support/painted-glyphs.test-support.js';

describe('a comment on text that crosses its cell clip', () => {
  it('tints the visible part of a condensed glyph whose advance lies outside the cell', async () => {
    // 16pt `W` condensed by 6pt, 3.6pt before the cell edge: its 2pt advance
    // ends left of the cell, its 8pt glyph reaches 4.4pt into it.
    const { runs, glyphs } = await paintedDocument(
      '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
      + '<w:p><w:pPr><w:ind w:left="-72"/></w:pPr>'
      + '<w:r><w:rPr><w:spacing w:val="-120"/><w:sz w:val="32"/></w:rPr><w:t>W</w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>normal</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>');
    const run = runs.find((candidate) => candidate.text === 'W')!;
    const glyph = glyphs.find((candidate) => candidate.character === 'W')!;
    const visible = intersect(glyph.box, glyph.clip)!;
    // The glyph is cut at the cell edge, and its advance box lies outside.
    expect(glyph.clip).not.toBeNull();
    expect(visible.left).toBeGreaterThan(glyph.box.left);
    expect(run.x + run.clusterBoxes![0]!.x + run.clusterBoxes![0]!.width).toBeLessThan(visible.left);
    const source = run.source!;
    const [thread] = resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [{
      commentId: '1', source, startRunIndex: run.sourceRunIndex!, endRunIndex: run.sourceRunIndex! + 1,
      reference: { source, runIndex: run.sourceRunIndex! + 1, affinity: 'preceding' },
    }], runs);
    const rects = thread!.anchors[0]!.rects.map(pageBoxOf);
    expect(rects).toHaveLength(1);
    for (const side of ['left', 'top', 'right', 'bottom'] as const) expect(rects[0]![side]).toBeCloseTo(visible[side], 6);
  });
});
