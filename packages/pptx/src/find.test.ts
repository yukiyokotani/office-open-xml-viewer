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
  // Painted group by group in the order of the set.
  expect(find.slideHighlights(0).map((box) => box.color)).toEqual([
    'rgba(0,0,255,0.3)', 'rgba(0,0,255,0.3)', 'rgba(0,0,255,0.3)', 'rgba(0, 184, 148, 0.35)',
  ]);
  expect((await find.highlightPrev())?.matchIndex).toBe(3);
});

it('keeps a term inside one table cell when a longer term of its group would cross cells', async () => {
  const cells = [
    run('cat', { elementIndex: 0, tableCell: { row: 0, column: 0 } }),
    run('dog', { elementIndex: 0, tableCell: { row: 0, column: 1 } }),
  ];
  const find = new PptxFindController(() => 1, async () => cells);
  // "catdog" exists only across the cell boundary, where nothing may match.
  const matches = await find.highlight([{ id: 'a', terms: ['catdog', 'cat', 'dog'] }]);
  expect(matches.map((match) => [match.term, match.text])).toEqual([['cat', 'cat'], ['dog', 'dog']]);
  expect((await find.find('cat', { wholeWord: true })).map((match) => match.text)).toEqual(['cat']);
  await expect(find.find('catdog', { wholeWord: true })).resolves.toEqual([]);
});

it('joins the runs of one authored line across wraps and stacked cells, and separates authored lines', () => {
  const line = (text: string, textSpan: number, overrides: Partial<PptxTextRunInfo> = {}) =>
    run(text, { shapeId: '7', textSpan, ...overrides });
  // An automatic wrap inside a word: the next drawn line continues the line.
  expect(pptxRunsBreak(line('cat', 0), line('alog', 0, { inShapeY: 12 }))).toBe(false);
  // A paragraph end or a manual break starts another authored line.
  expect(pptxRunsBreak(line('cat', 0), line('alog', 1, { inShapeY: 12 }))).toBe(true);
  expect(pptxRunsBreak(line('cat', 0), line('alog', 1, { inShapeX: 3 }))).toBe(true);
  // The span alone decides: justification may move a piece of a word away,
  // and a tab may be drawn with no gap at all.
  expect(pptxRunsBreak(line('cat', 0), line('alog', 0, { inShapeX: 3 }))).toBe(false);
  expect(pptxRunsBreak(line('cat', 0), line('alog', 0, { inShapeX: 40 }))).toBe(false);
  expect(pptxRunsBreak(line('cat', 0), line('dog', 1, { inShapeX: 3 }))).toBe(true);
  // Stacked vertical text: one run per glyph, down a column and into the next.
  expect(pptxRunsBreak(line('c', 0), line('a', 0, { inShapeY: 28 }))).toBe(false);
  expect(pptxRunsBreak(line('t', 0, { inShapeY: 56 }), line('a', 0, { inShapeX: 30 }))).toBe(false);
  expect(pptxRunsBreak(line('t', 0, { inShapeY: 56 }), line('d', 1, { inShapeX: 30 }))).toBe(true);
  // Another shape separates whatever the line index.
  expect(pptxRunsBreak(line('t', 0), run('d', { shapeId: '8', textSpan: 0, inShapeX: 1 }))).toBe(true);
});

it('does not take a wrapped or stacked word for separate words', async () => {
  // "catalog" wrapped after "cat", then a second paragraph "cat".
  const wrapped = [
    run('cat', { shapeId: '7', textSpan: 0 }),
    run('alog', { shapeId: '7', textSpan: 0, inShapeY: 12 }),
    run('cat', { shapeId: '7', textSpan: 1, inShapeY: 24 }),
  ];
  // The same text stacked: one glyph per run.
  const stacked = [...'catalog'].map((text, index) => run(text, { shapeId: '7', textSpan: 0, inShapeY: index * 28 }))
    .concat([...'cat'].map((text, index) => run(text, { shapeId: '7', textSpan: 1, inShapeY: index * 28, inShapeX: 30 })));
  for (const [runs, expected] of [[wrapped, [2]], [stacked, [7, 8, 9]]] as const) {
    const find = new PptxFindController(() => 1, async () => [...runs]);
    expect(await find.find('cat', { wholeWord: true })).toHaveLength(1);
    expect(find.slideHighlights(0)[0].slices.map((slice) => slice.runIndex)).toEqual(expected);
    expect(await find.find('cat')).toHaveLength(2);
    expect((await find.highlight([{ id: 'a', terms: ['cat'], wholeWord: true }])).map((match) => match.text)).toEqual(['cat']);
    // The whole word is still found across the wrap.
    expect(await find.find('catalog', { wholeWord: true })).toHaveLength(1);
  }
});
