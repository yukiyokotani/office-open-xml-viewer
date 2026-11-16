import { expect, it } from 'vitest';
import { DocxFindController } from './find.js';
import { docxRunsBreak } from './find-run-break.js';
import type { DocxTextRunInfo } from './renderer';

it('maps shared search results and highlight geometry to DOCX pages', async () => {
  const run = (text: string): DocxTextRunInfo => ({
    text, x: 0, y: 0, w: text.length, h: 10, fontSize: 10, font: '10px monospace',
  });
  const pages = [[run('Hel'), run('lo')], [run('hello')]];
  const find = new DocxFindController(() => pages.length, async (page) => pages[page]);
  expect((await find.find('hello')).map((match) => match.location)).toEqual([{ page: 0 }, { page: 1 }]);
  find.next();
  expect(find.activePage()).toBe(0);
  expect(find.pageHighlights(0)[0].slices).toHaveLength(2);
  expect(find.pageRuns(1)).toEqual(pages[1]);
});

const paragraph = (index: number): DocxTextRunInfo['source'] =>
  ({ story: 'body', storyInstance: 'body', path: [index] });
const at = (text: string, x: number, y: number, extra: Partial<DocxTextRunInfo> = {}): DocxTextRunInfo => ({
  text, x, y, w: text.length * 5, h: 12, fontSize: 10, font: '10px monospace', ...extra,
});

it('separates runs of different paragraphs, lines, or apart on a line, and joins the rest', () => {
  const head = at('Intro', 0, 0, { source: paragraph(0) });
  // The same paragraph, continuing the line where the previous run ends.
  expect(docxRunsBreak(head, at('duction', 25, 0, { source: paragraph(0) }))).toBe(false);
  // Another paragraph (a table cell on the same row keeps the same y).
  expect(docxRunsBreak(head, at('Total', 25, 0, { source: paragraph(1) }))).toBe(true);
  expect(docxRunsBreak(head, at('Total', 25, 0, { source: { ...paragraph(0)!, story: 'header' } }))).toBe(true);
  expect(docxRunsBreak(head, at('Total', 25, 0, { source: { ...paragraph(0)!, path: [0, 1] } }))).toBe(true);
  // The next line of the same paragraph (an explicit line break leaves no space).
  expect(docxRunsBreak(head, at('next', 0, 12, { source: paragraph(0) }))).toBe(true);
  // A tab stop: same line, a gap wider than a quarter of the font size.
  expect(docxRunsBreak(head, at('3', 80, 0, { source: paragraph(0) }))).toBe(true);
  expect(docxRunsBreak(head, at('duction', 27, 0, { source: paragraph(0) }))).toBe(false);
  // Right-to-left order: the next run ends where the previous one starts.
  expect(docxRunsBreak(at('ab', 50, 0), at('cd', 40, 0))).toBe(false);
  // A vertical page: a line is a column, so its runs share x.
  const rotate = 'rotate(90deg)';
  expect(docxRunsBreak(at('縦', 100, 0, { transform: rotate }), at('書', 100, 10, { transform: rotate }))).toBe(false);
  expect(docxRunsBreak(at('縦', 100, 0, { transform: rotate }), at('書', 80, 0, { transform: rotate }))).toBe(true);
});

it('highlights whole words that end at a paragraph, and reports group and term', async () => {
  const pages = [[
    at('Introduction', 0, 0, { source: paragraph(0) }),
    at('The intro ', 0, 20, { source: paragraph(1) }),
    at('duction', 50, 20, { source: paragraph(1) }),
  ]];
  const find = new DocxFindController(() => pages.length, async (page) => pages[page]);
  const matches = await find.highlight([
    { id: 'system', terms: ['introduction', 'the'], wholeWord: true, color: 'rgba(0,0,255,0.3)' },
    { id: 'user', terms: ['duction'] },
  ]);
  expect(matches.map((match) => [match.group, match.term, match.text, match.location.page])).toEqual([
    ['system', 'introduction', 'Introduction', 0],
    ['user', 'duction', 'duction', 0],
    ['system', 'the', 'The', 0],
    ['user', 'duction', 'duction', 0],
  ]);
  expect(find.pageHighlights(0).map((box) => box.color)).toEqual([
    'rgba(0,0,255,0.3)', 'rgba(0, 184, 148, 0.35)', 'rgba(0,0,255,0.3)', 'rgba(0, 184, 148, 0.35)',
  ]);
  expect((await find.highlightNext())?.matchIndex).toBe(0);
});
