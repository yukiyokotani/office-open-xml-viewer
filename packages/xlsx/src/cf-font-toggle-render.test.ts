import { describe, expect, it } from 'vitest';
import { applyAutoRowHeights, renderViewport } from './renderer.js';
import type { CellFont, CfRule, Dxf, PivotTableMetadata, Styles, Worksheet } from './types.js';

// A conditional-formatting dxf is a differential format over the cell's own
// formatting (ECMA-376 §18.3.1.10 cfRule, §18.8.14-15 dxf): `<b val="0"/>`
// (§18.8.2 CT_BooleanProperty) turns off bold the cell style turns on, while a
// dxf `<font>` that omits `<b>` leaves the cell's bold alone. Excel's PDF
// export of a control workbook draws the cell-style bold+italic value regular
// under an explicit-off rule and keeps it bold+italic under a colour-only rule.

const STYLE_FONT: CellFont = {
  bold: true,
  italic: true,
  underline: false,
  strike: false,
  size: 10,
  color: null,
  name: 'Arial',
};

const PLAIN_FONT: CellFont = { ...STYLE_FONT, bold: false, italic: false };

const STYLES: Styles = {
  fonts: [STYLE_FONT, PLAIN_FONT],
  fills: [],
  borders: [],
  cellXfs: [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
    { fontId: 1, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
  ],
  numFmts: [],
  dxfs: [
    // <font><b val="0"/><i val="0"/></font>
    { font: { ...PLAIN_FONT, color: null }, fill: null, border: null, fontToggles: { bold: false, italic: false } },
    // <font><color rgb="FFFF0000"/></font>
    { font: { ...PLAIN_FONT, color: '#FF0000' }, fill: null, border: null, fontToggles: {} },
    // <font><b/><i/></font>
    { font: { ...STYLE_FONT, color: null }, fill: null, border: null, fontToggles: { bold: true, italic: true } },
  ] satisfies Dxf[],
};

function rule(dxfId: number, priority: number): CfRule {
  return { type: 'expression', formula: 'TRUE', dxfId, priority } as CfRule;
}

function worksheet(): Worksheet {
  const cell = (row: number, styleIndex: number, text: string) => ({
    index: row,
    height: null,
    cells: [{ row, col: 1, styleIndex, value: { type: 'text' as const, text } }],
  });
  const at = (row: number) => [{ top: row, bottom: row, left: 1, right: 1 }];
  return {
    name: 'CF toggles',
    rows: [
      cell(1, 0, 'explicit off'), cell(2, 0, 'omitted'), cell(3, 1, 'explicit on'),
      { index: 4, height: null, cells: [{ row: 4, col: 1, styleIndex: 0, value: { type: 'number' as const, number: 5 } }] },
    ],
    colWidths: { 1: 20 },
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [
      { sqref: at(1), rules: [rule(0, 1)] },
      { sqref: at(2), rules: [rule(1, 2)] },
      { sqref: at(3), rules: [rule(2, 3)] },
      // ECMA-376 §18.3.1.10 stopIfTrue: a matching colour-only `cellIs`
      // rule stops the lower-priority explicit-off rule, so the cell-style
      // bold+italic stays.
      {
        sqref: at(4),
        rules: [
          { type: 'cellIs', operator: 'greaterThan', formulas: ['0'], dxfId: 1, priority: 4, stopIfTrue: true },
          rule(0, 5),
        ],
      },
    ],
    images: [],
    charts: [],
  } as Worksheet;
}

/** A recording canvas: the font in effect for each drawn and measured text. */
function recorder() {
  let font = '13px sans-serif';
  const drawn = new Map<string, string>();
  const measured = new Map<string, string>();
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 400, height: 200 },
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText(text: string) { measured.set(text, font); return { width: [...text].length * 7 } as TextMetrics; },
    fillText(text: string) { drawn.set(text, font); },
    createLinearGradient: () => ({ addColorStop: noop }),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => (key in target ? target[key] : noop),
    set: (target, key, value) => { target[key] = value; return true; },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, drawn, measured };
}

const isBold = (font: string | undefined) => /(^| )(bold|[6-9]00) /.test(font ?? '');
const isItalic = (font: string | undefined) => /^italic /.test(font ?? '');

function fontsByText(): Map<string, string> {
  const { ctx, drawn } = recorder();
  renderViewport(ctx, worksheet(), STYLES, { row: 1, col: 1, rows: 4, cols: 1 });
  return drawn;
}

describe('conditional-format font toggles over the cell style', () => {
  it('turns cell-style bold/italic off for an explicit off, keeps them when omitted', () => {
    const drawn = fontsByText();
    const style = (text: string) => {
      const font = drawn.get(text);
      expect(font, text).toBeDefined();
      return { bold: /(^| )(bold|[6-9]00) /.test(font as string), italic: /^italic /.test(font as string) };
    };
    expect(style('explicit off')).toEqual({ bold: false, italic: false });
    expect(style('omitted')).toEqual({ bold: true, italic: true });
    expect(style('explicit on')).toEqual({ bold: true, italic: true });
  });

  it('keeps cell-style bold/italic when a matching stopIfTrue rule precedes an explicit off', () => {
    const font = fontsByText().get('5');
    expect({ bold: isBold(font), italic: isItalic(font) }).toEqual({ bold: true, italic: true });
  });
});

// Paths other than the main paint loop compose the same layers: a
// PivotTable style's wholeTable bold+italic (§18.8.41) reaches the text of
// a merged anchor painted by the off-screen pre-pass, and the font measured
// for automatic row height.
describe('PivotTable style font on the secondary paint and measurement paths', () => {
  const text = 'pivot styled label';
  function pivotSheet(options: { merged?: boolean; wrap?: boolean }): { ws: Worksheet; styles: Styles } {
    const pivotDxf: Dxf = {
      font: { ...PLAIN_FONT, bold: true, italic: true }, fill: null, border: null,
      fontToggles: { bold: true, italic: true },
    };
    const pivot = {
      name: 'P', cacheId: 1,
      location: { top: 1, left: 1, bottom: 2, right: 3, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
      rowFields: [0], columnFields: [], pageFields: [], dataFields: [],
      status: { state: 'complete' },
      rowItems: [{ kind: 'data', depth: 0 }], columnItems: [],
      style: {
        name: 'S', showRowHeaders: true, showColumnHeaders: true,
        showRowStripes: false, showColumnStripes: false, showLastColumn: false,
        elements: [{ kind: 'wholeTable', size: 1, dxf: pivotDxf }],
      },
    } as unknown as PivotTableMetadata;
    const styles: Styles = {
      ...STYLES,
      cellXfs: [{ ...STYLES.cellXfs[1], wrapText: !!options.wrap }],
    };
    const ws = {
      ...worksheet(),
      rows: [{ index: 1, height: null, cells: [{ row: 1, col: 1, styleIndex: 0, value: { type: 'text' as const, text } }] }],
      conditionalFormats: [],
      mergeCells: options.merged ? [{ top: 1, bottom: 1, left: 1, right: 3 }] : [],
      pivotTables: [pivot],
    } as Worksheet;
    return { ws, styles };
  }

  it('reaches an off-screen merged anchor', () => {
    const { ws, styles } = pivotSheet({ merged: true });
    const { ctx, drawn } = recorder();
    // Viewport starts at column 2: the anchor at column 1 is off-screen.
    renderViewport(ctx, ws, styles, { row: 1, col: 2, rows: 1, cols: 2 });
    const font = drawn.get(text);
    expect(font).toBeDefined();
    expect({ bold: isBold(font), italic: isItalic(font) }).toEqual({ bold: true, italic: true });
  });

  it('reaches the automatic row-height measurement', () => {
    const { ws, styles } = pivotSheet({ wrap: true });
    const { ctx, measured } = recorder();
    applyAutoRowHeights(ctx, ws, styles);
    const fonts = [...measured].filter(([t]) => text.includes(t) && t.trim()).map(([, f]) => f);
    expect(fonts.length).toBeGreaterThan(0);
    for (const font of fonts) expect(isBold(font) && isItalic(font), font).toBe(true);
  });
});
