import { describe, expect, it } from 'vitest';
import type { DocParagraph, DocTable } from '../types.js';
import { tableColumnLayoutInput } from '../parser-model.js';
import {
  acquireRetainedTable,
} from './table-acquisition.js';
import type {
  LayoutServices,
  ParagraphLayout,
  TableFormatInput,
} from './types.js';

const noBorders = Object.freeze({
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
});

function retainedParagraph(widthPt: number): ParagraphLayout {
  const bounds = { xPt: 0, yPt: 0, widthPt, heightPt: 8 };
  return {
    kind: 'paragraph', id: 'paragraph',
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: 'table-cell', ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: 8,
    alignment: 'left', bidi: false,
    spacing: { beforePt: 0, afterPt: 0 },
    borders: [], lines: [],
  } as unknown as ParagraphLayout;
}

describe('retained table acquisition', () => {
  it('retains the immutable layout input and recursive acquisitions beside final geometry', () => {
    const nested = {
      type: 'table', rows: [], colWidths: [], borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const table = {
      rows: [{
        cells: [{
          content: [nested], colSpan: 1, vMerge: null, borders: noBorders,
          background: null, vAlign: 'top',
        }],
        gridBefore: 0, gridAfter: 0, isHeader: false, cantSplit: false,
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const formats = new WeakMap<object, TableFormatInput>([
      [table, {
        effectiveStyleId: null, ordinaryFlow: true, positioning: null,
        firstRowException: null,
        rows: [{
          height: null, cantSplit: false, repeatedHeader: false,
          cellSpacingPt: 0, justification: null, exception: null,
          cells: [{ marginsPt: { top: 0, right: 0, bottom: 0, left: 0 } }],
        }],
      }],
      [nested, {
        effectiveStyleId: null, ordinaryFlow: true, positioning: null,
        firstRowException: null, rows: [],
      }],
    ]);

    const acquisition = acquireRetainedTable(
      table,
      [100],
      100,
      { yPt: 0 },
      [0],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: (source) => formats.get(source)!,
        resolveColumns: () => [],
        createCellState: (state) => state,
        acquireParagraph: () => retainedParagraph(100),
        registerFloatingTable: () => null,
        advanceState: () => {},
      },
    );

    expect(acquisition.input.columnWidthsPt).toEqual([100]);
    expect(acquisition.layout.columnWidthsPt).toEqual([100]);
    expect(acquisition.input.rows[0]).toMatchObject({
      logicalRowIndex: 0,
      cantSplit: false,
      repeatedHeader: false,
    });
    expect(acquisition.input.rows[0]?.cells[0]?.blocks[0]).toMatchObject({
      sourceBlockIndex: 0,
    });
    expect(Object.keys(acquisition.nestedById)).toHaveLength(1);
    const nestedAcquisition = Object.values(acquisition.nestedById)[0];
    expect(nestedAcquisition?.input.rows).toEqual([]);
    expect(nestedAcquisition?.layout.rows).toEqual([]);
    expect(Object.isFrozen(acquisition.input)).toBe(true);
    expect(Object.isFrozen(acquisition.nestedById)).toBe(true);
  });

  it('keeps an authored but Word-ignored tblpPr in ordinary table flow', () => {
    const table = {
      type: 'table', rows: [], colWidths: [], borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', bidiVisual: false,
      tblpPr: {
        leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
        horzAnchor: 'text', horzSpecified: true, vertAnchor: 'margin',
        tblpX: 0, tblpY: 0,
      },
    } as unknown as DocTable;

    const acquisition = acquireRetainedTable(
      table,
      [],
      100,
      { yPt: 0 },
      [],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: () => ({
          effectiveStyleId: 'TableNormal', ordinaryFlow: true, positioning: null,
          firstRowException: null, rows: [],
        }),
        resolveColumns: () => [],
        createCellState: (state) => state,
        acquireParagraph: () => retainedParagraph(100),
        registerFloatingTable: () => null,
        advanceState: () => {},
      },
    );

    expect(acquisition.input.ordinaryFlow).toBe(true);
    expect(acquisition.floatingTables).toEqual([]);
  });

  it('owns a nested floating table outside cell flow and anchors it to the next regular paragraph', () => {
    // ECMA-376 §17.4.57 makes tblpPr tables out-of-flow while retaining their
    // logical position for anchoring to the next regular paragraph. Keeping the
    // same child in `blocks` and in a floating-placement registry would charge
    // its height and paint it twice.
    const anchor = {
      type: 'paragraph', runs: [{ type: 'text' }],
    } as unknown as DocParagraph;
    const framedParagraph = {
      type: 'paragraph', runs: [{ type: 'text' }], framePr: {},
    } as unknown as DocParagraph;
    const floating = {
      type: 'table',
      rows: [{
        cells: [{
          content: [{ type: 'paragraph' } as unknown as DocParagraph],
          colSpan: 1, vMerge: null, borders: noBorders,
          background: null, vAlign: 'top',
        }],
        gridBefore: 0, gridAfter: 0, isHeader: false, cantSplit: false,
      }],
      colWidths: [40], borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', bidiVisual: false, overlap: 'never',
      tblpPr: {
        leftFromText: 1, rightFromText: 2, topFromText: 3, bottomFromText: 4,
        horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
        tblpX: 5, tblpY: 6,
      },
    } as unknown as DocTable;
    const interveningTable = {
      ...floating,
      rows: [],
      tblpPr: undefined,
      overlap: undefined,
    } as unknown as DocTable;
    const table = {
      rows: [{
        cells: [{
          content: [floating, interveningTable, framedParagraph, anchor],
          colSpan: 1, vMerge: null,
          borders: noBorders, background: null, vAlign: 'top',
        }],
        gridBefore: 0, gridAfter: 0, isHeader: false, cantSplit: false,
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const oneRowFormat: TableFormatInput = {
      effectiveStyleId: null, ordinaryFlow: true, positioning: null,
      firstRowException: null,
      rows: [{
        height: null, cantSplit: false, repeatedHeader: false,
        cellSpacingPt: 0, justification: null, exception: null,
        cells: [{ marginsPt: { top: 0, right: 0, bottom: 0, left: 0 } }],
      }],
    };

    const cellState = { wrapWidthPt: null as number | null };
    const acquisition = acquireRetainedTable(
      table,
      [100],
      100,
      cellState,
      [0],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: (source) => source === floating
          ? {
              effectiveStyleId: null,
              ordinaryFlow: false,
              positioning: {
                leftFromTextPt: 1, rightFromTextPt: 2,
                topFromTextPt: 3, bottomFromTextPt: 4,
                horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
                xPt: 5, yPt: 6,
              },
              firstRowException: null,
              rows: oneRowFormat.rows,
            }
          : source === interveningTable
            ? {
                effectiveStyleId: null, ordinaryFlow: true, positioning: null,
                firstRowException: null, rows: [],
              }
            : oneRowFormat,
        resolveColumns: (source) => source === floating ? [40] : [100],
        createCellState: (state) => state,
        acquireParagraph: (state, paragraph, widthPt) => {
          const retained = retainedParagraph(widthPt);
          if (paragraph !== anchor || state.wrapWidthPt == null) return retained;
          return {
            ...retained,
            lines: [{
              range: { start: 0, end: 4 },
              bounds: { xPt: state.wrapWidthPt, yPt: 0, widthPt: 40, heightPt: 8 },
              baselinePt: 7,
              advancePt: 8,
              placements: [],
            }],
          } as ParagraphLayout;
        },
        registerFloatingTable: (state) => {
          // The production adapter writes a float oracle exclusion here. This
          // synthetic width makes the resulting anchor-line geometry observable.
          state.wrapWidthPt = 60;
          return { xPt: 5, yPt: 6 };
        },
        advanceState: () => {},
      },
    );
    const cellInput = acquisition.input.rows[0]!.cells[0]!;
    const nested = Object.values(acquisition.nestedById)
      .find((candidate) => !candidate.input.ordinaryFlow)!;
    const floatingPlacements = acquisition.floatingTables;

    expect(nested.input.ordinaryFlow).toBe(false);
    expect(cellInput.blocks.map((block) => block.sourceBlockIndex)).toEqual([1, 2, 3]);
    expect(acquisition.layout.rows[0]?.cells[0]?.blocks).toHaveLength(3);
    expect(acquisition.layout.advancePt).toBe(16);
    expect(floatingPlacements).toEqual([{
      hostCellId: cellInput.id,
      sourceBlockIndex: 0,
      anchorBlockIndex: 3,
      tableId: nested.layout.id,
      overlap: 'never',
      positioning: {
        leftFromTextPt: 1,
        rightFromTextPt: 2,
        topFromTextPt: 3,
        bottomFromTextPt: 4,
        horzAnchor: 'text',
        horzSpecified: true,
        vertAnchor: 'text',
        xPt: 5,
        yPt: 6,
      },
      acquiredTextOffsetPt: { xPt: 5, yPt: 6 },
    }]);
    expect(Object.isFrozen(floatingPlacements)).toBe(true);
    expect(Object.isFrozen(floatingPlacements[0]?.positioning)).toBe(true);
    const anchorLayout = cellInput.blocks.find((block) => block.sourceBlockIndex === 3)?.layout;
    expect(anchorLayout?.kind).toBe('paragraph');
    expect(anchorLayout?.kind === 'paragraph' ? anchorLayout.lines[0]?.bounds : null).toEqual({
      xPt: 60,
      yPt: 0,
      widthPt: 40,
      heightPt: 8,
    });

    floating.tblpPr!.horzAnchor = 'page';
    floating.overlap = 'overlap';
    expect(floatingPlacements[0]?.positioning.horzAnchor).toBe('text');
    expect(floatingPlacements[0]?.overlap).toBe('never');
    expect(JSON.parse(JSON.stringify(floatingPlacements[0]))).toEqual(floatingPlacements[0]);
  });

  it('includes cell-spacing in autofit intrinsic grid constraints', () => {
    const table = {
      colWidths: [50, 50],
      rows: [{
        gridBefore: 0, gridAfter: 0,
        cells: [0, 1].map(() => ({
          content: [], colSpan: 1, vMerge: null, borders: noBorders,
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        })),
        __tableRowLayout: {
          height: null, justification: null, beforeWidth: null, afterWidth: null,
          cellSpacing: { kind: 'dxa', value: '80' }, exception: null,
        },
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [{ width: '1000' }, { width: '1000' }],
          requiredColumnCount: 2,
        },
        preferredWidth: null, layout: { kind: 'autofit' },
        cellSpacing: null, cellMargins: null,
      },
    } as unknown as DocTable;

    const input = tableColumnLayoutInput(
      table,
      100,
      () => ({ minWidthPt: 10, maxWidthPt: 20 }),
    );

    expect(input.rows[0]?.cells.map((cell) => ({
      min: cell.minContentWidthPt,
      max: cell.maxContentWidthPt,
    }))).toEqual([{ min: 16, max: 26 }, { min: 16, max: 26 }]);
  });

  it('acquires cell children at the final width after cell spacing and margins', () => {
    const paragraph = { type: 'paragraph' } as unknown as DocParagraph;
    const table = {
      rows: [{
        cells: [
          { content: [paragraph], colSpan: 1, vMerge: null, borders: noBorders,
            background: null, vAlign: 'top' },
          { content: [paragraph], colSpan: 1, vMerge: null, borders: noBorders,
            background: null, vAlign: 'top' },
        ],
        gridBefore: 0, gridAfter: 0, isHeader: false,
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 1, cellMarginBottom: 0, cellMarginLeft: 1,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const format: TableFormatInput = {
      effectiveStyleId: null, ordinaryFlow: true, positioning: null,
      firstRowException: null,
      rows: [{
        height: null,
        cantSplit: false,
        repeatedHeader: false,
        cellSpacingPt: 4,
        justification: null,
        exception: null,
        cells: [
          { marginsPt: { top: 0, right: 1, bottom: 0, left: 1 } },
          { marginsPt: { top: 0, right: 1, bottom: 0, left: 1 } },
        ],
      }],
    };
    const acquiredWidths: number[] = [];

    const acquisition = acquireRetainedTable(
      table,
      [50, 50],
      100,
      { yPt: 0 },
      [0],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: () => format,
        resolveColumns: () => [],
        createCellState: (state) => state,
        acquireParagraph: (_state, _paragraph, widthPt) => {
          acquiredWidths.push(widthPt);
          return retainedParagraph(widthPt);
        },
        registerFloatingTable: () => null,
        advanceState: () => {},
      },
    );

    // Each outer cell owns one full outer spacing inset and half of the shared
    // inner gap: 50 - (4 + 2) - (1 + 1) = 42pt.
    expect(acquiredWidths).toEqual([42, 42]);
    expect(acquisition.layout.rows[0]?.cells.map((cell) => cell.contentBounds.widthPt))
      .toEqual([42, 42]);
  });

  it('acquires a spanning cell from its final gridBefore-adjusted column band', () => {
    const paragraph = { type: 'paragraph' } as unknown as DocParagraph;
    const table = {
      rows: [{
        cells: [{
          content: [paragraph], colSpan: 2, vMerge: null, borders: noBorders,
          background: null, vAlign: 'top',
        }],
        gridBefore: 1, gridAfter: 0, isHeader: false,
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 3, cellMarginBottom: 0, cellMarginLeft: 2,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const format: TableFormatInput = {
      effectiveStyleId: null, ordinaryFlow: true, positioning: null,
      firstRowException: null,
      rows: [{
        height: null,
        cantSplit: false,
        repeatedHeader: false,
        cellSpacingPt: 0,
        justification: null,
        exception: null,
        cells: [{ marginsPt: { top: 0, right: 3, bottom: 0, left: 2 } }],
      }],
    };
    const acquiredWidths: number[] = [];

    const acquisition = acquireRetainedTable(
      table,
      [20, 30, 50],
      100,
      { yPt: 0 },
      [0],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: () => format,
        resolveColumns: () => [],
        createCellState: (state) => state,
        acquireParagraph: (_state, _paragraph, widthPt) => {
          acquiredWidths.push(widthPt);
          return retainedParagraph(widthPt);
        },
        registerFloatingTable: () => null,
        advanceState: () => {},
      },
    );

    expect(acquiredWidths).toEqual([75]);
    expect(acquisition.input.rows[0]?.cells[0]).toMatchObject({
      columnStart: 1,
      columnSpan: 2,
    });
    expect(acquisition.layout.rows[0]?.cells[0]).toMatchObject({
      flowBounds: { xPt: 20, widthPt: 80 },
      contentBounds: { xPt: 22, widthPt: 75 },
    });
  });
});
