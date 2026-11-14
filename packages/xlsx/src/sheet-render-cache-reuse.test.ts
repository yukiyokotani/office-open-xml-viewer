import { describe, expect, it } from 'vitest';
import { getSheetRenderCache, inheritSheetRenderCache, renderViewport } from './renderer.js';
import type { CellFont, Styles, Worksheet } from './types.js';

// The per-worksheet render index (cell map, merge topology, compiled CF,
// structured-style map, …) is viewport- and measurement-independent. It must
// survive the per-frame re-inheritance a viewer-owned projection receives, and
// must never carry scaled or measured values that a zoom or a late font load
// would invalidate.

const FONT: CellFont = {
  bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: 'Calibri',
};

function styles(): Styles {
  return {
    fonts: [FONT],
    fills: [],
    borders: [],
    cellXfs: [
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
      // General (§18.8.30): a number too wide for its column is shortened to
      // the measured width (fitGeneralNumber).
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
    ],
    numFmts: [],
    dxfs: [],
  };
}

function sheet(name: string, text: string): Worksheet {
  return {
    name,
    rows: [
      { index: 1, height: null, cells: [{ row: 1, col: 1, styleIndex: 0, value: { type: 'text', text } }] },
      { index: 2, height: null, cells: [{ row: 2, col: 1, styleIndex: 1, value: { type: 'number', number: 12345.678 } }] },
    ],
    colWidths: { 1: 3 },
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    images: [],
    charts: [],
  } as unknown as Worksheet;
}

/** A recording canvas whose text advance per code point is `advance()`. */
function recorder(advance: () => number) {
  let font = '13px sans-serif';
  const drawn: { text: string; font: string }[] = [];
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 400, height: 200 },
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText(text: string) { return { width: [...text].length * advance() } as TextMetrics; },
    fillText(text: string) { drawn.push({ text, font }); },
    createLinearGradient: () => ({ addColorStop: noop }),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => (key in target ? target[key] : noop),
    set: (target, key, value) => { target[key] = value; return true; },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, drawn };
}

const VIEWPORT = { row: 1, col: 1, rows: 3, cols: 2 };

describe('sheet render index reuse across frames', () => {
  it('keeps one index for a projection re-inherited on every frame', () => {
    const source = sheet('A', 'label');
    const st = styles();
    const projection: Worksheet = { ...source, colWidths: { ...source.colWidths } };
    const { ctx } = recorder(() => 7);
    const entries = [];
    for (let frame = 0; frame < 3; frame++) {
      // XlsxWorkbook.renderViewport re-inherits the source index before every
      // paint of a viewer-owned worksheet.
      inheritSheetRenderCache(source, projection, st);
      renderViewport(ctx, projection, st, VIEWPORT);
      entries.push(getSheetRenderCache(projection, st));
    }
    expect(entries[1]).toBe(entries[0]);
    expect(entries[2]).toBe(entries[0]);
    expect(entries[0]).toBe(getSheetRenderCache(source, st));
  });

  it('keeps the index across zoom and device-pixel-ratio changes; text follows the scale', () => {
    const ws = sheet('A', 'label');
    const st = styles();
    const { ctx, drawn } = recorder(() => 7);
    renderViewport(ctx, ws, st, VIEWPORT, { cellScale: 1, dpr: 1 });
    const entry = getSheetRenderCache(ws, st);
    const at1 = drawn.find((d) => d.text === 'label')?.font;
    drawn.length = 0;
    renderViewport(ctx, ws, st, VIEWPORT, { cellScale: 2, dpr: 2 });
    const at2 = drawn.find((d) => d.text === 'label')?.font;
    expect(getSheetRenderCache(ws, st)).toBe(entry);
    expect(at1).toMatch(/^15px /);
    expect(at2).toMatch(/^29px /);
  });

  it('re-measures after a font load: no measured width is reused across frames', () => {
    const ws = sheet('A', 'label');
    const st = styles();
    // A face that loads late changes advances between two frames.
    let advance = 1;
    const { ctx, drawn } = recorder(() => advance);
    renderViewport(ctx, ws, st, VIEWPORT);
    const entry = getSheetRenderCache(ws, st);
    expect(drawn.some((d) => d.text === '12345.678')).toBe(true);
    drawn.length = 0;
    advance = 9;
    renderViewport(ctx, ws, st, VIEWPORT);
    expect(getSheetRenderCache(ws, st)).toBe(entry);
    // Row headers paint single digits; the value is the only longer number.
    const value = drawn.filter((d) => d.text.length > 1 && /^[\d.E+#]+$/.test(d.text)).map((d) => d.text);
    expect(value).toHaveLength(1);
    expect(value[0]).not.toBe('12345.678');
  });

  it('keeps one index per worksheet instance across sheet switches', () => {
    const a = sheet('A', 'alpha');
    const b = sheet('B', 'beta');
    const st = styles();
    const { ctx, drawn } = recorder(() => 7);
    renderViewport(ctx, a, st, VIEWPORT);
    const entryA = getSheetRenderCache(a, st);
    renderViewport(ctx, b, st, VIEWPORT);
    const entryB = getSheetRenderCache(b, st);
    expect(entryB).not.toBe(entryA);
    drawn.length = 0;
    renderViewport(ctx, a, st, VIEWPORT);
    expect(getSheetRenderCache(a, st)).toBe(entryA);
    expect(drawn.some((d) => d.text === 'alpha')).toBe(true);
    expect(drawn.some((d) => d.text === 'beta')).toBe(false);
  });

  it('rebuilds when provisional rows grow or the workbook dxfs change', () => {
    const ws = sheet('A', 'label');
    const st = styles();
    const first = getSheetRenderCache(ws, st);
    ws.rows.push({ index: 3, height: null, cells: [{ row: 3, col: 1, styleIndex: 0, value: { type: 'text', text: 'late' } }] } as Worksheet['rows'][number]);
    const grown = getSheetRenderCache(ws, st);
    expect(grown).not.toBe(first);
    expect(grown.cellMap.get('3:1')).toBeDefined();
    const restyled = getSheetRenderCache(ws, { ...st, dxfs: [] });
    expect(restyled).not.toBe(grown);
  });
});
