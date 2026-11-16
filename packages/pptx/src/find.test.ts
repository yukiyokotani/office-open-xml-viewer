import { expect, it } from 'vitest';
import { PptxFindController } from './find.js';
import { pptxRunsBreak } from './find-run-break.js';
import type { PptxTextRunInfo } from './renderer';

function run(text: string, overrides: Partial<PptxTextRunInfo> = {}): PptxTextRunInfo {
  return {
    text, inShapeX: 0, inShapeY: 0, w: text.length, h: 10,
    fontSize: 10, font: '10px monospace', shapeX: 0, shapeY: 0,
    shapeW: 100, shapeH: 20, rotation: 0, ...overrides,
  };
}

it('maps shared search results and highlights to PPTX slides', async () => {
  const slides = [[run('Hel'), run('lo')], [run('hello')]];
  const find = new PptxFindController(() => slides.length, async (slide) => slides[slide]);
  expect((await find.find('hello')).map((match) => match.location)).toEqual([{ slide: 0 }, { slide: 1 }]);
  find.next();
  expect(find.activeSlide()).toBe(0);
  expect(find.slideHighlights(0)[0].slices).toHaveLength(2);
  expect(find.slideRuns(1)).toEqual(slides[1]);
});

it('matches within a table cell but rejects text crossing cell boundaries', async () => {
  const sameCell = [
    run('Hel', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('lo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
  ];
  const adjacentCells = [
    run('foo', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('bar', { elementIndex: 0, tableCell: { row: 0, column: 1 } }),
  ];
  const find = new PptxFindController(() => 1, async () => sameCell);
  await expect(find.find('Hello')).resolves.toHaveLength(1);
  find.setSlideRuns(0, adjacentCells);
  await expect(find.find('foobar')).resolves.toEqual([]);
});

it('keeps preview tables separate when they have no source-model locator', async () => {
  const previewRuns = [
    run('Hel', { renderElementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('lo', { renderElementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('world', { renderElementIndex: 1, tableCell: { row: 0, column: 0 } }),
  ];
  const find = new PptxFindController(() => 1, async () => previewRuns);
  await expect(find.find('Hello')).resolves.toHaveLength(1);
  await expect(find.find('Helloworld')).resolves.toEqual([]);
});

it('separates runs of different shapes, cells or lines, or apart on a line, and joins the rest', () => {
  const head = run('Intro', { elementIndex: 0, shapeId: '4' });
  const next = (overrides: Partial<PptxTextRunInfo>) =>
    run('duction', { elementIndex: 0, shapeId: '4', inShapeX: 5, ...overrides });
  expect(pptxRunsBreak(head, next({}))).toBe(false);
  expect(pptxRunsBreak(head, next({ elementIndex: 1 }))).toBe(true);
  expect(pptxRunsBreak(head, next({ shapeId: '5' }))).toBe(true);
  expect(pptxRunsBreak(head, next({ shapeX: 40 }))).toBe(true);
  expect(pptxRunsBreak(head, next({ origin: 'layout' as PptxTextRunInfo['origin'] }))).toBe(true);
  expect(pptxRunsBreak(head, next({ inShapeY: 12 }))).toBe(true);
  expect(pptxRunsBreak(head, next({ inShapeX: 40 }))).toBe(true);
  expect(pptxRunsBreak(head, next({ inShapeX: 7 }))).toBe(false);
  const cell = (column: number, text: string) =>
    run(text, { elementIndex: 0, tableCell: { row: 0, column }, inShapeX: column === 0 ? 0 : 5 });
  expect(pptxRunsBreak(cell(0, 'Total'), cell(1, 'Amount'))).toBe(true);
  expect(pptxRunsBreak(cell(0, 'To'), { ...cell(0, 'tal'), inShapeX: 2 })).toBe(false);
});

it('highlights whole words that end at a line of a title, and reports group and term', async () => {
  const slides = [[
    run('Quarterly', { shapeId: '2' }),
    run('Revenue up', { shapeId: '2', inShapeY: 20 }),
    run('revenue', { shapeId: '3', shapeY: 60 }),
  ]];
  const find = new PptxFindController(() => slides.length, async (slide) => slides[slide]);
  const matches = await find.highlight([
    { id: 'system', terms: ['quarterly', 'revenue'], wholeWord: true, color: 'rgba(0,0,255,0.3)' },
    { id: 'user', terms: ['up'] },
  ]);
  expect(matches.map((match) => [match.group, match.term, match.text, match.location.slide])).toEqual([
    ['system', 'quarterly', 'Quarterly', 0],
    ['system', 'revenue', 'Revenue', 0],
    ['user', 'up', 'up', 0],
    ['system', 'revenue', 'revenue', 0],
  ]);
  expect(find.slideHighlights(0).map((box) => box.color)).toEqual([
    'rgba(0,0,255,0.3)', 'rgba(0,0,255,0.3)', 'rgba(0, 184, 148, 0.35)', 'rgba(0,0,255,0.3)',
  ]);
  expect((await find.highlightPrev())?.matchIndex).toBe(3);
});
