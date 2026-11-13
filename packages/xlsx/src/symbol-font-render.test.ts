import { describe, expect, it } from 'vitest';
import { createSymbolFaceOracleCache } from '@silurus/ooxml-core';
import { applyAutoRowHeights, derivedAutoRowHeights, drawShapeText, drawSlicerFrame, renderViewport } from './renderer.js';
import { SYMBOL_ENCODED, symbolRealm, type SymbolRealm } from './test-support/symbol-realm.js';
import type { CellFont, CellValue, CellXf, ShapeText, Styles, Worksheet } from './types.js';

// Symbol-encoded cell, run, shape and slicer text (symbol-paint.ts) in a
// realm whose route sentinels work: a face the realm draws keeps its codes,
// a missing one converts them to Unicode, and every measurement uses the
// painted form and font the paint uses.

const font = (name: string, size = 11): CellFont => ({ name, size, bold: false, italic: false,
  underline: false, strike: false, color: null });
const WINGDINGS = { wingdings: SYMBOL_ENCODED };
const VS15 = '︎';

interface CellSpec { font: CellFont; value: CellValue; xf?: Partial<CellXf>; numFmt?: string }

function sheet(cells: CellSpec[], colWidth = 40): { worksheet: Worksheet; styles: Styles } {
  const styles: Styles = {
    fonts: cells.map((cell) => cell.font), fills: [], borders: [], dxfs: [],
    numFmts: cells.flatMap((cell, index) => cell.numFmt ? [{ numFmtId: 164 + index, formatCode: cell.numFmt }] : []),
    cellXfs: cells.map((cell, index) => ({ fontId: index, fillId: 0, borderId: 0,
      numFmtId: cell.numFmt ? 164 + index : 0, alignH: null, alignV: null, wrapText: false, ...cell.xf })),
  } as Styles;
  const worksheet = { name: 'Symbols', defaultColWidth: colWidth, defaultRowHeight: 15,
    colWidths: {}, rowHeights: {}, mergeCells: [], freezeRows: 0, freezeCols: 0,
    conditionalFormats: [], images: [], charts: [],
    rows: cells.map((cell, index) => ({ index: index + 1, height: null,
      cells: [{ row: index + 1, col: 1, styleIndex: index, value: cell.value }] })) } as unknown as Worksheet;
  return { worksheet, styles };
}

function render(realm: SymbolRealm, cells: CellSpec[], colWidth?: number) {
  const { worksheet, styles } = sheet(cells, colWidth);
  renderViewport(realm.ctx, worksheet, styles, { row: 1, col: 1, rows: cells.length, cols: 2 }, { dpr: 1 });
  return worksheet;
}

const text = (value: string): CellValue => ({ type: 'text', text: value });
const painted = (realm: SymbolRealm, needle: string) => realm.fills.filter((fill) => fill.text.includes(needle));

describe('XLSX symbol-font cell text', () => {
  it('converts the codes of a missing face and keeps those of a face the realm draws', () => {
    const missing = symbolRealm();
    render(missing, [
      { font: font('Symbol'), value: text('abg') },
      { font: font('Wingdings'), value: text('üJ') },
    ]);
    expect(painted(missing, 'αβγ')[0]?.font).not.toMatch(/Symbol/u);
    expect(painted(missing, `✓☺${VS15}`)[0]?.font).not.toMatch(/Wingdings/u);
    expect(painted(missing, 'abg')).toEqual([]);

    const installed = symbolRealm(WINGDINGS);
    render(installed, [{ font: font('Wingdings'), value: text('üJ') }]);
    expect(painted(installed, 'üJ')[0]?.font).toMatch(/^15px "Wingdings"/u);
    expect(painted(installed, '✓')).toEqual([]);
  });

  it('keeps a partial face for a text whose drawn code cannot convert', () => {
    // The face draws J (the smiley) but not ü (the check): one face per text,
    // so J converts too, unless its smiley only paints in colour.
    const partial = { wingdings: (ch: string) => ch === 'J' || ch === '\uF04A' };
    const converted = symbolRealm(partial);
    render(converted, [{ font: font('Wingdings'), value: text('Jü') }]);
    expect(painted(converted, `☺${VS15}✓`)[0]?.font).not.toMatch(/Wingdings/u);
    const colour = symbolRealm(partial, undefined, (value) => value.includes('☺'));
    render(colour, [{ font: font('Wingdings'), value: text('Jü') }]);
    expect(painted(colour, 'Jü')[0]?.font).toMatch(/^15px "Wingdings"/u);
  });

  it('paints rotated and stacked rich runs in their own faces', () => {
    const run = (value: string, name: string) => ({ text: value,
      font: { bold: false, italic: false, underline: false, strike: false, name } });
    for (const textRotation of [45, 255]) {
      const realm = symbolRealm();
      render(realm, [
        { font: font('Calibri'), xf: { textRotation },
          value: { type: 'text', text: 'xl', runs: [run('x', 'Calibri'), run('l', 'Wingdings')] } },
        { font: font('Symbol'), xf: { textRotation },
          value: { type: 'text', text: 'ab', runs: [run('a', 'Calibri'), run('b', 'Symbol')] } },
      ]);
      expect(painted(realm, '⚫')[0]?.font, `rotation ${textRotation}`).not.toMatch(/Wingdings/u);
      expect(painted(realm, 'x')[0]?.font).toMatch(/"Calibri"/u);
      expect(painted(realm, 'a')[0]?.font).toMatch(/"Calibri"/u);
      expect(painted(realm, 'β')).toHaveLength(1);
      expect(painted(realm, 'α')).toEqual([]);
    }
  });

  it('converts number-format output and the invalid-value hash fill in the cell face', () => {
    const realm = symbolRealm();
    render(realm, [
      { font: font('Symbol'), value: { type: 'number', number: 7 }, numFmt: '"a"0' },
      { font: font('Wingdings'), value: { type: 'number', number: -1 }, numFmt: 'yyyy' },
    ]);
    expect(painted(realm, 'α7')).toHaveLength(1);
    // Wingdings '#' is the scissors, which takes the text colour here.
    expect(painted(realm, '✁✁').length).toBeGreaterThan(0);
    expect(realm.fills.some((fill) => fill.text.includes('#'))).toBe(false);
  });

  it('paints each rich run in its own painted form and face', () => {
    const realm = symbolRealm();
    render(realm, [{ font: font('Calibri'), value: { type: 'text', text: 'Done ü', runs: [
      { text: 'Done ', font: { bold: false, italic: false, underline: false, strike: false, name: 'Calibri' } },
      { text: 'ü', font: { bold: false, italic: false, underline: false, strike: false, name: 'Wingdings' } },
    ] } }]);
    expect(painted(realm, 'Done ')[0]?.font).toMatch(/"Calibri"/u);
    const check = painted(realm, '✓')[0];
    expect(check?.font).not.toMatch(/Wingdings/u);
    // The check mark follows the run before it on the same line.
    expect(check!.x).toBeGreaterThan(painted(realm, 'Done ')[0]!.x);
  });

  it('wraps and auto-fits a wrapped cell by the widths of its painted form', () => {
    // Each converted smiley carries VS15: twice the authored code points, so
    // a width measured on the authored letters would need fewer lines.
    const realm = symbolRealm();
    const cells = [{ font: font('Wingdings'), value: text('JJJJ JJJJ JJJJ JJJJ'), xf: { wrapText: true } }];
    const { worksheet, styles } = sheet(cells, 8);
    applyAutoRowHeights(realm.ctx, worksheet, styles);
    const height = derivedAutoRowHeights(worksheet).get(1);
    renderViewport(realm.ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 }, { dpr: 1 });
    // The cell's lines (headers paint at 11 px).
    const lines = realm.fills.filter((fill) => fill.font.startsWith('15px'));
    expect(painted(realm, '☺').length).toBeGreaterThan(1);
    expect(lines.every((line) => !/J/u.test(line.text))).toBe(true);
    // The automatic height (points) holds exactly the painted lines: 18 px
    // line boxes plus the top and bottom insets.
    expect(height).toBeCloseTo((lines.length * 18 + 4) * 0.75, 5);
  });

  it('overflows an unwrapped value by its painted width', () => {
    const realm = symbolRealm();
    render(realm, [{ font: font('Wingdings'), value: text('JJJJJJJJJJJJ') }], 4);
    const fill = painted(realm, '☺')[0]!;
    const measured = realm.measures.filter((entry) => entry.text === fill.text && entry.font === fill.font);
    expect(measured.length).toBeGreaterThan(0);
  });

  it('keeps one oracle across renders while the realm is unchanged', () => {
    const realm = symbolRealm(WINGDINGS);
    const cells = [{ font: font('Wingdings'), value: text('ü') }];
    const { worksheet, styles } = sheet(cells);
    const symbolFaces = createSymbolFaceOracleCache();
    const probes = () => realm.measures.length;
    renderViewport(realm.ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 }, { dpr: 1, symbolFaces });
    const first = probes();
    renderViewport(realm.ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 }, { dpr: 1, symbolFaces });
    const second = probes() - first;
    const fresh = symbolRealm(WINGDINGS);
    renderViewport(fresh.ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 }, { dpr: 1 });
    // The painting context measures only the cell text on the second render;
    // a render without the cache judges on the painting context itself.
    expect(second).toBeLessThan(fresh.measures.length);
    expect(painted(realm, 'ü')).toHaveLength(2);
  });
});

describe('XLSX symbol-font shape and slicer text', () => {
  it('paints a:sym codes in the symbol face and converts them where it is missing', () => {
    const body: ShapeText = {
      anchor: 't', wrap: 'square', autoFit: 'none',
      paragraphs: [{ align: 'l', runs: [
        { type: 'text', text: 'Item ', bold: false, italic: false, size: 11, fontFace: 'Calibri' },
        { type: 'text', text: 'x', bold: false, italic: false, size: 11, fontFace: 'Calibri', fontFaceSym: 'Wingdings' },
        { type: 'text', text: 'ab', bold: false, italic: false, size: 11, fontFace: 'Symbol' },
      ] }],
    } as ShapeText;
    const missing = symbolRealm();
    drawShapeText(missing.ctx, body, 400, 100, 1);
    expect(painted(missing, '✓')[0]?.font).not.toMatch(/Wingdings/u);
    expect(painted(missing, 'x')[0]?.font).toMatch(/"Calibri"/u);
    expect(painted(missing, 'αβ')).toHaveLength(1);
    expect(missing.fills.some((fill) => /[-]/u.test(fill.text))).toBe(false);

    const installed = symbolRealm(WINGDINGS);
    drawShapeText(installed.ctx, body, 400, 100, 1);
    expect(painted(installed, '')[0]?.font).toMatch(/"Wingdings"/u);
  });

  it('clips a slicer label by its painted width', () => {
    const realm = symbolRealm();
    const style = { fontFamily: 'Wingdings', fontSize: 11 };
    drawSlicerFrame(realm.ctx, 'JJJJJJJJJJJJJJJJJJJJ', [{ name: 'ü', selected: true }], 0, 0, 120, 120, 1,
      { header: style, selectedItemWithData: style } as never);
    const caption = painted(realm, '☺')[0]!;
    expect(caption.text.endsWith('…')).toBe(true);
    expect(caption.font).not.toMatch(/Wingdings/u);
    expect(painted(realm, '✓')).toHaveLength(1);
  });
});
