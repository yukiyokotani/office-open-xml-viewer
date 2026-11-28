import { describe, expect, it } from 'vitest';
import { resolveDocxCommentThreads } from './comments.js';
import { DocxFindController } from './find.js';
import {
  inlinePictureRun, intersect, pageBoxOf, paintedDocument, PICTURE_PART,
} from './test-support/painted-glyphs.test-support.js';

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

describe('find in text a cell clips away', () => {
  it('returns no match for a run painted wholly beyond the cell edge', async () => {
    // A -60pt indent puts the short first run left of the cell.
    const { runs, glyphs } = await paintedDocument(
      '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
      + '<w:p><w:pPr><w:ind w:left="-1200"/></w:pPr>'
      + '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Hidden </w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>and the shown text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>');
    expect(glyphs.filter((glyph) => glyph.character === 'H').every((glyph) => intersect(glyph.box, glyph.clip) === null)).toBe(true);
    const find = new DocxFindController(() => 1, async () => runs);
    expect(await find.find('Hidden')).toEqual([]);
    expect((await find.find('shown')).map((match) => match.text)).toEqual(['shown']);
  });
});

describe('a comment on a turned picture its cell clips away', () => {
  it('marks nothing where the picture is turned out of the cell', async () => {
    // A 30pt picture turned 45 degrees, 65pt into a 60pt cell with an exact
    // 12pt row: its bounding box reaches back into the cell, its outline
    // does not.
    const { runs } = await paintedDocument(
      '<w:tbl><w:tblPr><w:tblW w:w="1200" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
      + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="240" w:hRule="exact"/></w:trPr>'
      + '<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:ind w:left="1300"/></w:pPr>'
      + `${inlinePictureRun(30, 30, '', 45)}</w:p></w:tc></w:tr></w:tbl><w:p/>`, PICTURE_PART);
    const picture = runs.find((run) => run.inlineContent === 'object')!;
    // The picture keeps its source run, with nothing to show.
    expect(picture.sourceRunIndex).toBeDefined();
    const source = picture.source!;
    expect(resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [{
      commentId: '1', source, startRunIndex: picture.sourceRunIndex!, endRunIndex: picture.sourceRunIndex! + 1,
      reference: { source, runIndex: picture.sourceRunIndex! + 1, affinity: 'preceding' },
    }], runs)).toEqual([]);
  });
});
