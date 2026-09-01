import { describe, expect, it } from 'vitest';
import type { RetainedTableAcquisition } from './layout/table-acquisition.js';
import {
  startTableFragmentCursor,
  takeTableFragment,
  type TableFragmentCursor,
  type TableFragmentLayout,
} from './layout/table-pagination.js';
import { layoutParagraph } from './layout/paragraph.js';
import { layoutTable } from './layout/table.js';
import type {
  AcquiredParagraphLayoutInput,
  LayoutServices,
  ParagraphLayout,
  TableEdgeInputs,
  TableLayoutInput,
  TableRowLayoutInput,
} from './layout/types.js';

const noBorders: TableEdgeInputs = {
  top: null,
  right: null,
  bottom: null,
  left: null,
  insideH: null,
  insideV: null,
};

function paragraph(id: string, lineHeightsPt: readonly number[]): ParagraphLayout {
  let yPt = 0;
  const lines = lineHeightsPt.map((heightPt, lineIndex) => {
    const line = {
      range: { start: lineIndex, end: lineIndex + 1 },
      bounds: { xPt: 0, yPt, widthPt: 80, heightPt },
      baselinePt: yPt + heightPt * 0.8,
      advancePt: heightPt,
      placements: [],
    };
    yPt += heightPt;
    return line;
  });
  return layoutParagraph({
    kind: 'paragraph',
    id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'cell',
    ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt: 80, heightPt: yPt },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 80, heightPt: yPt },
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines,
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  } satisfies AcquiredParagraphLayoutInput);
}

function row(
  logicalRowIndex: number,
  lineHeightsPt: readonly number[],
  options: {
    cantSplit?: boolean;
    repeatedHeader?: boolean;
    verticalMerge?: 'none' | 'restart' | 'continue';
    exactHeightPt?: number;
  } = {},
): TableRowLayoutInput {
  const verticalMerge = options.verticalMerge ?? 'none';
  return {
    id: `row-${logicalRowIndex}`,
    source: { story: 'body', storyInstance: 'body', path: [0, logicalRowIndex] },
    logicalRowIndex,
    cantSplit: options.cantSplit ?? false,
    heightPt: options.exactHeightPt ?? null,
    heightRule: options.exactHeightPt === undefined ? 'auto' : 'exact',
    cellSpacingPt: 0,
    exceptionBorders: null,
    alignment: 'left',
    indentPt: 0,
    repeatedHeader: options.repeatedHeader ?? false,
    cells: [{
      id: `cell-${logicalRowIndex}`,
      source: { story: 'body', storyInstance: 'body', path: [0, logicalRowIndex, 0] },
      columnStart: 0,
      columnSpan: 1,
      verticalMerge,
      margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      vAlign: 'top',
      borders: noBorders,
      blocks: verticalMerge === 'continue' ? [] : [{
        layout: paragraph(`paragraph-${logicalRowIndex}`, lineHeightsPt),
        sourceBlockIndex: 0,
      }],
    }],
  };
}

function parallelRow(
  logicalRowIndex: number,
  cellLineHeightsPt: readonly (readonly number[])[],
): TableRowLayoutInput {
  return {
    ...row(logicalRowIndex, []),
    cells: cellLineHeightsPt.map((lineHeightsPt, cellIndex) => ({
      id: `cell-${logicalRowIndex}-${cellIndex}`,
      source: { story: 'body', storyInstance: 'body', path: [0, logicalRowIndex, cellIndex] },
      columnStart: cellIndex,
      columnSpan: 1,
      verticalMerge: 'none',
      margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      vAlign: 'top',
      borders: noBorders,
      blocks: [{
        layout: paragraph(`paragraph-${logicalRowIndex}-${cellIndex}`, lineHeightsPt),
        sourceBlockIndex: 0,
      }],
    })),
  };
}

function acquisition(rows: readonly TableRowLayoutInput[]): RetainedTableAcquisition {
  const columnCount = Math.max(1, ...rows.map((item) => item.cells.reduce(
    (max, cell) => Math.max(max, cell.columnStart + cell.columnSpan), 0,
  )));
  const input: TableLayoutInput = {
    kind: 'table',
    id: 'table',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body',
    ordinaryFlow: true,
    alignment: 'left',
    indentPt: 0,
    bidiVisual: false,
    columnWidthsPt: Array.from({ length: columnCount }, () => 100 / columnCount),
    borders: noBorders,
    rows,
  };
  const placement = {
    container: {
      id: 'body',
      kind: 'body' as const,
      bounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 10_000 },
    },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 10_000 },
  };
  return Object.freeze({
    input,
    layout: layoutTable(input, placement, {} as LayoutServices).layout,
    nestedById: Object.freeze({}),
    floatingTables: Object.freeze([]),
  });
}

function paginate(
  source: RetainedTableAcquisition,
  firstPageHeightPt: number,
  freshPageHeightPt: number,
  compatibility: 'word' | 'standard' = 'standard',
): readonly (TableFragmentLayout | null)[] {
  const pages: Array<TableFragmentLayout | null> = [];
  let cursor: TableFragmentCursor | null = startTableFragmentCursor();
  let availableHeightPt = firstPageHeightPt;
  let guard = 0;
  while (cursor) {
    guard += 1;
    if (guard > 100) throw new Error('table pagination did not make progress');
    const pageIndex = pages.length;
    const result = takeTableFragment(source, cursor, {
      availableHeightPt,
      freshPageHeightPt,
      placement: {
        container: {
          id: `page-${pageIndex}`,
          kind: 'body',
          bounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: availableHeightPt },
        },
        cursor: { xPt: 0, yPt: 0 },
        availableBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: availableHeightPt },
      },
      services: {} as LayoutServices,
      compatibility,
      page: {
        physicalPageIndex: pageIndex,
        displayPageNumber: pageIndex + 1,
        occurrenceId: `page-${pageIndex}`,
      },
    });
    if (!result.fragment && !result.requiresFreshPage) {
      throw new Error('table pagination returned no fragment or fresh-page request');
    }
    pages.push(result.fragment);
    cursor = result.nextCursor;
    availableHeightPt = freshPageHeightPt;
  }
  return pages;
}

function sourceRows(pages: readonly (TableFragmentLayout | null)[]) {
  return pages.flatMap((page) => page?.rows
    .filter((item) => item.ownership === 'source')
    .map((item) => [item.logicalRowIndex, item.fragmentIndex] as const) ?? []);
}

describe('retained table pagination across pages', () => {
  it('emits the largest fitting whole-row prefix on each page without loss or duplication', () => {
    const source = acquisition(Array.from({ length: 10 }, (_, index) => row(index, [100])));

    const pages = paginate(source, 350, 350);

    expect(pages.map((page) => page?.rows.length)).toEqual([3, 3, 3, 1]);
    expect(pages.every((page) => page !== null && page.advancePt <= 350)).toBe(true);
    expect(sourceRows(pages)).toEqual(Array.from(
      { length: 10 },
      (_, index) => [index, 0],
    ));
  });

  it('relocates a cantSplit row that fits a fresh page', () => {
    // ECMA-376 §17.4.6 requires relocation when the row fits a fresh page but
    // not the remaining page band.
    const source = acquisition([
      row(0, [100], { cantSplit: true }),
      row(1, [100]),
      row(2, [100]),
      row(3, [100]),
    ]);

    const pages = paginate(source, 50, 350);

    expect(pages[0]).toBeNull();
    expect(pages.slice(1).map((page) => page?.rows.length)).toEqual([3, 1]);
    expect(sourceRows(pages)).toEqual([[0, 0], [1, 0], [2, 0], [3, 0]]);
  });

  it('repeats only the consecutive leading tblHeader rows', () => {
    // ECMA-376 §17.4.49 limits repetition to the leading header prefix; a later
    // authored tblHeader after a body row does not restart that prefix.
    const source = acquisition([
      row(0, [100], { repeatedHeader: true }),
      row(1, [100]),
      row(2, [100]),
      row(3, [100], { repeatedHeader: true }),
      row(4, [100]),
      row(5, [100]),
    ]);

    const pages = paginate(source, 250, 250);

    expect(pages.map((page) => page?.rows.map((item) => (
      [item.logicalRowIndex, item.ownership]
    )))).toEqual([
      [[0, 'source'], [1, 'source']],
      [[0, 'repeated-header'], [2, 'source']],
      [[0, 'repeated-header'], [3, 'source']],
      [[0, 'repeated-header'], [4, 'source']],
      [[0, 'repeated-header'], [5, 'source']],
    ]);
    expect(sourceRows(pages)).toEqual([[0, 0], [1, 0], [2, 0], [3, 0], [4, 0], [5, 0]]);
  });

  it('continues an over-tall splittable row at retained line boundaries', () => {
    const source = acquisition([row(0, [60, 60, 60])]);

    const pages = paginate(source, 100, 100);

    expect(sourceRows(pages)).toEqual([[0, 0], [0, 1], [0, 2]]);
    expect(pages.map((page) => page?.rows[0]?.cells[0]?.contentRanges)).toEqual([
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 }],
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 1, lineEnd: 2 }],
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 2, lineEnd: 3 }],
    ]);
  });

  it('relocates a parallel one-line row when no common legal cut fits the remaining band', () => {
    // Word-verified: when only the shorter one-line cell fits, Word relocates
    // both cells instead of emitting the short cell alone on the first page.
    const source = acquisition([parallelRow(0, [[20], [30]])]);

    const pages = paginate(source, 25, 100, 'word');

    expect(pages[0]).toBeNull();
    expect(sourceRows(pages)).toEqual([[0, 0]]);
    expect(pages[1]?.rows[0]?.cells.map((cell) => cell.contentRanges)).toEqual([
      [{ kind: 'whole', blockIndex: 0 }],
      [{ kind: 'whole', blockIndex: 0 }],
    ]);
  });

  it('still splits a parallel row when every unfinished cell reaches a legal cut', () => {
    // Word-verified: both cells own a first-line boundary in the retained band,
    // so the longer cell may continue while the shorter cell is complete.
    const source = acquisition([parallelRow(0, [[20, 20, 20], [20]])]);

    const pages = paginate(source, 25, 100, 'word');

    expect(sourceRows(pages)).toEqual([[0, 0], [0, 1]]);
    expect(pages[0]?.rows[0]?.cells.map((cell) => cell.contentRanges)).toEqual([
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 }],
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 }],
    ]);
    expect(pages[1]?.rows[0]?.cells.map((cell) => cell.contentRanges)).toEqual([
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 1, lineEnd: 3 }],
      [],
    ]);
  });

  it('relocates when one cell completes but another cannot reach its first boundary', () => {
    // Word-verified: completing one multi-line cell is insufficient when the
    // parallel cell's first indivisible line crosses the page cut.
    const source = acquisition([parallelRow(0, [[10, 10], [30]])]);

    const pages = paginate(source, 25, 100, 'word');

    expect(pages[0]).toBeNull();
    expect(sourceRows(pages)).toEqual([[0, 0]]);
    expect(pages[1]?.rows[0]?.cells.map((cell) => cell.contentRanges)).toEqual([
      [{ kind: 'whole', blockIndex: 0 }],
      [{ kind: 'whole', blockIndex: 0 }],
    ]);
  });

  it('does not apply the Word-observed common row cut in standard mode', () => {
    const source = acquisition([parallelRow(0, [[20], [30]])]);

    const pages = paginate(source, 25, 100, 'standard');

    expect(sourceRows(pages)).toEqual([[0, 0], [0, 1]]);
    expect(pages[0]?.rows[0]?.cells.map((cell) => cell.contentRanges)).toEqual([
      [{ kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 }],
      [],
    ]);
  });

  it('keeps vMerge semantics immutable when a page boundary crosses the span', () => {
    // ECMA-376 §17.4.84 defines restart/continue semantics but does not make a
    // vertical merge page-atomic; only page-local visual ownership is derived.
    const source = acquisition([
      row(0, [100]),
      row(1, [100], { verticalMerge: 'restart' }),
      row(2, [], { verticalMerge: 'continue', exactHeightPt: 100 }),
      row(3, [100]),
    ]);

    const pages = paginate(source, 200, 200);

    expect(sourceRows(pages)).toEqual([[0, 0], [1, 0], [2, 0], [3, 0]]);
    expect(pages[0]?.rows[1]?.cells[0]?.verticalMerge).toBe('restart');
    expect(pages[1]?.rows[0]?.cells[0]?.verticalMerge).toBe('continue');
    expect(pages[1]?.rows[0]?.cells[0]?.visualMergeOwnership).toBe('continuation');
    expect(source.input.rows[1]?.cells[0]?.verticalMerge).toBe('restart');
    expect(source.input.rows[2]?.cells[0]?.verticalMerge).toBe('continue');
  });
});
