import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { CellFont, CfRule, Dxf, Styles, Worksheet } from './types.js';

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
    rows: [cell(1, 0, 'explicit off'), cell(2, 0, 'omitted'), cell(3, 1, 'explicit on')],
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
    ],
    images: [],
    charts: [],
  } as Worksheet;
}

function fontsByText(): Map<string, string> {
  let font = '13px sans-serif';
  const drawn = new Map<string, string>();
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 400, height: 200 },
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 7 }) as TextMetrics,
    fillText(text: string) { drawn.set(text, font); },
    createLinearGradient: () => ({ addColorStop: noop }),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => (key in target ? target[key] : noop),
    set: (target, key, value) => { target[key] = value; return true; },
  });
  renderViewport(ctx as unknown as CanvasRenderingContext2D, worksheet(), STYLES,
    { row: 1, col: 1, rows: 3, cols: 1 });
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
});
