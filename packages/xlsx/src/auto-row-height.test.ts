import { describe, expect, it } from 'vitest';
import {
  applyAutoRowHeights,
  getGridGeometryForWorksheet,
  invalidateAutoRowHeights,
  markAutoRowHeightsPrepared,
  rowHeightToPx,
} from './renderer.js';
import { worksheetWithAutoRowHeights } from './render-orchestrator.js';
import { GridGeometry } from './internal/grid-geometry.js';
import type { CellFont, CellXf, Styles, Worksheet } from './types.js';

function measurementContext(): CanvasRenderingContext2D {
  let font = '15px sans-serif';
  return {
    canvas: { width: 1, height: 1 },
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 7 }) as TextMetrics,
    save() {},
    restore() {},
  } as unknown as CanvasRenderingContext2D;
}

function countingMeasurementContext(): {
  ctx: CanvasRenderingContext2D;
  count: () => number;
} {
  let calls = 0;
  const ctx = measurementContext();
  ctx.measureText = (text: string) => {
    calls++;
    return { width: [...text].length * 7 } as TextMetrics;
  };
  return { ctx, count: () => calls };
}

const fonts: CellFont[] = [
  { bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: 'Calibri' },
  { bold: false, italic: false, underline: false, strike: false, size: 20, color: null, name: 'Calibri' },
  { bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: 'Meiryo' },
];

const xf = (fontId: number, over: Partial<CellXf> = {}): CellXf => ({
  fontId,
  fillId: 0,
  borderId: 0,
  numFmtId: 0,
  alignH: 'left',
  alignV: 'top',
  wrapText: false,
  ...over,
});

const styles: Styles = {
  fonts,
  fills: [],
  borders: [],
  cellXfs: [
    xf(0),
    xf(0, { wrapText: true }),
    xf(1),
    xf(0, { textRotation: 90 }),
    xf(2, { textRotation: 255 }),
    { ...xf(0, { wrapText: true }), numFmtId: 164 },
    { ...xf(0, { wrapText: true }), numFmtId: 165 },
    xf(0, { textRotation: 1 }),
  ],
  numFmts: [
    { numFmtId: 164, formatCode: '"one two three"' },
    { numFmtId: 165, formatCode: '"漢字"' },
  ],
  dxfs: [],
};

function worksheet(): Worksheet {
  return {
    name: 'auto height',
    rows: [
      {
        index: 1,
        height: null,
        cells: [{ row: 1, col: 1, styleIndex: 1, value: { type: 'text', text: 'abcdefghi klmnopqrst' } }],
      },
      {
        index: 2,
        height: 15,
        cells: [{ row: 2, col: 1, styleIndex: 1, value: { type: 'text', text: 'abcdefghi klmnopqrst' } }],
      },
      {
        index: 3,
        height: null,
        cells: [{ row: 3, col: 1, styleIndex: 1, value: { type: 'text', text: 'abcdefghi klmnopqrst' } }],
      },
      {
        index: 4,
        height: null,
        cells: [{ row: 4, col: 1, styleIndex: 2, value: { type: 'text', text: 'large' } }],
      },
      {
        index: 5,
        height: null,
        cells: [{ row: 5, col: 1, styleIndex: 3, value: { type: 'text', text: 'rotated text' } }],
      },
      {
        index: 6,
        height: null,
        cells: [{
          row: 6,
          col: 1,
          styleIndex: 1,
          value: {
            type: 'text',
            text: 'large rich tail',
            runs: [
              { text: 'large ', font: { ...fonts[1] } },
              { text: 'rich tail' },
            ],
          },
        }],
      },
    ],
    colWidths: { 1: 10 },
    rowHeights: { 2: 15 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [{ top: 3, left: 1, bottom: 3, right: 2 }],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  };
}

describe('XLSX automatic row height (ECMA-376 §18.3.1.73 / Office auto-fit)', () => {
  it('fits wrapped, rich, large, and rotated text while preserving explicit and merged rows', () => {
    const ws = worksheet();
    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);

    // 19 glyphs × 7px wrap into two lines in the 74px content band:
    // 2 × 18px Calibri line boxes + 4px top/bottom inset = 40px = 30pt.
    expect(ws.rowHeights[1]).toBe(30);
    expect(ws.rowHeights[2]).toBe(15); // authored row@ht always wins
    expect(ws.rowHeights[3]).toBeUndefined(); // merged cells do not drive auto-fit
    expect(ws.rowHeights[4]).toBeGreaterThan(15); // single large-font line
    expect(ws.rowHeights[5]).toBeGreaterThan(15); // rotated width becomes vertical extent
    expect(ws.rowHeights[6]).toBeGreaterThan(ws.rowHeights[1]); // rich max run metrics

    const rowAxis = getGridGeometryForWorksheet(ws).row;
    expect(rowAxis.sizeOf(1)).toBe(rowHeightToPx(30));
    expect(rowAxis.sizeOf(3)).toBe(rowHeightToPx(15));
    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(false);
  });

  it('keeps a wrap-enabled value that still fits one line at the default height', () => {
    const ws = worksheet();
    ws.rows = [{
      index: 1,
      height: null,
      cells: [{ row: 1, col: 1, styleIndex: 1, value: { type: 'text', text: 'short' } }],
    }];
    ws.rowHeights = {};

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(false);
    expect(ws.rowHeights).toEqual({});
  });

  it('refits after a column-width change without replacing manual row overrides', () => {
    const ws = worksheet();
    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);
    expect(ws.rowHeights[1]).toBe(30);

    ws.colWidths[1] = 30;
    invalidateAutoRowHeights(ws, [2]);
    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);
    expect(ws.rowHeights[1]).toBeUndefined();
    expect(ws.rowHeights[2]).toBe(15);

    const manual = worksheet();
    applyAutoRowHeights(measurementContext(), manual, styles);
    const sameNumericHeight = manual.rowHeights[1];
    invalidateAutoRowHeights(manual, [1]);
    applyAutoRowHeights(measurementContext(), manual, styles);
    expect(manual.rowHeights[1]).toBe(sameNumericHeight);
  });

  it('does not measure ordinary default-height cells in a dense sheet', () => {
    const ws = worksheet();
    ws.rows = Array.from({ length: 1_000 }, (_, index) => ({
      index: index + 1,
      height: null,
      cells: [{
        row: index + 1,
        col: 1,
        styleIndex: 0,
        value: { type: 'number' as const, number: index },
      }],
    }));
    ws.rowHeights = {};
    ws.mergeCells = [];
    const measured = countingMeasurementContext();

    expect(applyAutoRowHeights(measured.ctx, ws, styles)).toBe(false);
    expect(measured.count()).toBe(0);
  });

  it('does not grow a Normal-font single line because the default row rounds down to 19px', () => {
    const ws = worksheet();
    ws.defaultRowHeight = 14.25;
    ws.defaultFontFamily = 'Calibri';
    ws.defaultFontSize = 11;
    ws.rows = [{
      index: 1,
      height: null,
      cells: [{ row: 1, col: 1, styleIndex: 0, value: { type: 'text', text: 'ordinary' } }],
    }];
    ws.rowHeights = {};
    ws.mergeCells = [];
    const measured = countingMeasurementContext();

    expect(rowHeightToPx(ws.defaultRowHeight)).toBe(19);
    expect(applyAutoRowHeights(measured.ctx, ws, styles)).toBe(false);
    expect(ws.rowHeights).toEqual({});
    expect(measured.count()).toBe(0);
  });

  it('uses the tallest rich run for an unwrapped automatic row', () => {
    const ws = worksheet();
    ws.rows = [{
      index: 1,
      height: null,
      cells: [{
        row: 1,
        col: 1,
        styleIndex: 0,
        value: {
          type: 'text',
          text: 'ordinary large',
          runs: [
            { text: 'ordinary ' },
            { text: 'large', font: { ...fonts[1] } },
          ],
        },
      }],
    }];
    ws.rowHeights = {};
    ws.mergeCells = [];

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);
    expect(ws.rowHeights[1]).toBeGreaterThan(15);
  });

  it('preserves a manually authored sheet-wide default row height', () => {
    const ws = worksheet();
    ws.defaultRowHeight = 30;
    ws.defaultRowHeightCustom = true;
    ws.rowHeights = {};
    ws.mergeCells = [];

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(false);
    expect(ws.rowHeights).toEqual({});
    expect(getGridGeometryForWorksheet(ws).row.sizeOf(1)).toBe(rowHeightToPx(30));
  });

  it('preserves a row marked customHeight even when ht is absent', () => {
    const ws = worksheet();
    ws.rows = [{ ...ws.rows[0], customHeight: true }];
    ws.rowHeights = {};
    ws.mergeCells = [];

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(false);
    expect(ws.rowHeights).toEqual({});
  });

  it('uses the same compact glyph slots as paint for stacked text', () => {
    const ws = worksheet();
    ws.rows = [{
      index: 1,
      height: null,
      cells: [{ row: 1, col: 1, styleIndex: 4, value: { type: 'text', text: '縦書' } }],
    }];
    ws.rowHeights = {};
    ws.mergeCells = [];

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);
    expect(ws.rowHeights[1]).toBe(27);
  });

  it('reserves conditional-format icon width when wrapping', () => {
    const base = worksheet();
    base.colWidths = { 1: 13 };
    base.rows = [{
      index: 1,
      height: null,
      cells: [{ row: 1, col: 1, styleIndex: 5, value: { type: 'number', number: 123456 } }],
    }];
    base.rowHeights = {};
    base.mergeCells = [];
    const withIcon: Worksheet = {
      ...base,
      rows: base.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell })) })),
      rowHeights: {},
      conditionalFormats: [{
        sqref: [{ top: 1, left: 1, bottom: 1, right: 1 }],
        rules: [{
          type: 'iconSet',
          iconSet: '3TrafficLights1',
          cfvos: [
            { kind: 'percent', value: '0' },
            { kind: 'percent', value: '33' },
            { kind: 'percent', value: '67' },
          ],
          reverse: false,
          priority: 1,
        }],
      }],
    };

    applyAutoRowHeights(measurementContext(), base, styles);
    applyAutoRowHeights(measurementContext(), withIcon, styles);
    expect(withIcon.rowHeights[1]).toBeGreaterThan(base.rowHeights[1] ?? base.defaultRowHeight);
  });

  it('measures a wrapped General boolean with the inset paint uses', () => {
    // General centres booleans (§18.18.40); like paint, measurement keeps the
    // General indent inset, so an indented TRUE in a narrow wrapping column
    // fits exactly as the same General text does.
    const general = (value: Worksheet['rows'][number]['cells'][number]['value']): Worksheet => ({
      ...worksheet(),
      rows: [{ index: 1, height: null, cells: [{ row: 1, col: 1, styleIndex: 0, value }] }],
      colWidths: { 1: 1.5 },
      rowHeights: {},
      mergeCells: [],
    });
    const generalStyles: Styles = {
      ...styles,
      cellXfs: [{ ...xf(0, { wrapText: true, indent: 1 }), alignH: null }],
    };
    const bool = general({ type: 'bool', bool: true });
    const text = general({ type: 'text', text: 'TRUE' });
    applyAutoRowHeights(measurementContext(), bool, generalStyles);
    applyAutoRowHeights(measurementContext(), text, generalStyles);
    expect(bool.rowHeights[1]).toBe(text.rowHeights[1]);
  });

  it('preserves the caller-authoritative MDW on a render-local auto-height projection', () => {
    const source = worksheet();
    source.rowHeights = {};
    source.mergeCells = [];
    GridGeometry.forWorksheet(source, 13);

    const projection = worksheetWithAutoRowHeights(measurementContext(), source, styles);

    expect(projection).not.toBe(source);
    expect(source.rowHeights).toEqual({});
    expect(projection.rowHeights[1]).toBeDefined();
    expect(getGridGeometryForWorksheet(projection).maximumDigitWidth).toBe(13);
  });

  it('solves icon wrapping against the rounded row height used by paint', () => {
    const ws = worksheet();
    ws.rows = [{
      index: 1,
      height: null,
      cells: [
        { row: 1, col: 1, styleIndex: 6, value: { type: 'number', number: 1 } },
        { row: 1, col: 2, styleIndex: 7, value: { type: 'text', text: 'driver' } },
      ],
    }];
    ws.colWidths = { 1: 4.375, 2: 20 };
    ws.rowHeights = {};
    ws.mergeCells = [];
    ws.conditionalFormats = [{
      sqref: [{ top: 1, left: 1, bottom: 1, right: 1 }],
      rules: [{
        type: 'iconSet',
        iconSet: '3TrafficLights1',
        cfvos: [
          { kind: 'percent', value: '0' },
          { kind: 'percent', value: '33' },
          { kind: 'percent', value: '67' },
        ],
        reverse: false,
        priority: 1,
      }],
    }];
    GridGeometry.forWorksheet(ws, 8); // 4.375 × 8 MDW = 35px

    expect(applyAutoRowHeights(measurementContext(), ws, styles)).toBe(true);
    expect(rowHeightToPx(ws.rowHeights[1])).toBeGreaterThanOrEqual(40);
  });

  it('does not rescan a worker projection whose viewer supplied derived row heights', () => {
    const workerProjection = worksheet();
    workerProjection.rowHeights = { 1: 30 };
    const measured = countingMeasurementContext();
    markAutoRowHeightsPrepared(workerProjection);

    const rendered = worksheetWithAutoRowHeights(measured.ctx, workerProjection, styles);

    expect(rendered).toBe(workerProjection);
    expect(rendered.rowHeights[1]).toBe(30);
    expect(measured.count()).toBe(0);
  });
});
