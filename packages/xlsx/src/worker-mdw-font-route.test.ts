import { describe, expect, it } from 'vitest';
import type { OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { colWidthToPx } from './internal/grid-metrics.js';
import type { RenderViewportOptions, Styles, Worksheet, XlsxTextRunInfo } from './types.js';

const STYLES = {
  fonts: [{ name: 'Arial', size: 12, bold: false, italic: false,
    underline: false, strike: false, color: null }],
  fills: [], borders: [], numFmts: [], dxfs: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
} as Styles;

function worksheet(): Worksheet {
  return {
    name: 'Geometry',
    rows: [{ index: 1, height: null, cells: [{ row: 1, col: 2, styleIndex: 0,
      value: { type: 'text', text: 'Visible' } }] }],
    colWidths: { 1: 10, 2: 10 }, rowHeights: {},
    defaultColWidth: 10, defaultRowHeight: 15,
    defaultFontFamily: 'Arial', defaultFontSize: 12,
    mergeCells: [], freezeRows: 0, freezeCols: 0,
    conditionalFormats: [], images: [], charts: [],
  } as Worksheet;
}

function target(fallbackDigitWidth = 9): { canvas: OffscreenCanvas; paintedFonts: string[] } {
  const paintedFonts: string[] = [];
  let font = '16px Arial';
  const fontStack: string[] = [];
  const canvas = { width: 400, height: 100, getContext: () => ctx };
  const ctx = new Proxy({
    canvas,
    get font() { return font; },
    set font(value: string) { font = value; },
    save: () => fontStack.push(font),
    restore: () => { font = fontStack.pop() ?? font; },
    measureText: (_text: string) => ({ width: font.includes('__exact_arial') ? 7 : fallbackDigitWidth }),
    fillText: () => paintedFonts.push(font),
  } as Record<string, unknown>, {
    get(value, property) { return property in value ? value[property as string] : () => {}; },
    set(value, property, next) { value[property as string] = next; return true; },
  }) as unknown as CanvasRenderingContext2D;
  return { canvas: canvas as unknown as OffscreenCanvas, paintedFonts };
}

describe('XLSX render font and geometry authority', () => {
  it('keeps a viewer-supplied MDW through the worker render and auto-height projection', async () => {
    const ws = worksheet();
    for (const fallbackDigitWidth of [9, 11]) {
      const { canvas } = target(fallbackDigitWidth);
      const textRuns: XlsxTextRunInfo[] = [];
      await renderWorksheetViewport({ ws, styles: STYLES }, canvas,
        { row: 1, col: 1, rows: 1, cols: 2 }, {
          authoritativeMdw: 8,
          onTextRun: (run) => textRuns.push(run),
        } as RenderViewportOptions & { authoritativeMdw: number });
      expect(textRuns.find((run) => run.cellRef === 'B1')?.width).toBe(colWidthToPx(10, 8));
    }
  });

  it('paints a non-Calibri cell with the same exact local face used for Normal MDW', async () => {
    const { canvas, paintedFonts } = target();
    const route: OfficeFontFallbackRoute = {
      requestedFamily: 'Arial', family: '__exact_arial', source: 'local',
      resourceIdentity: 'office-local:test', weight: 400, style: 'normal',
      metric: { family: '__exact_arial' },
    };
    const textRuns: XlsxTextRunInfo[] = [];
    await renderWorksheetViewport({ ws: worksheet(), styles: STYLES }, canvas,
      { row: 1, col: 1, rows: 1, cols: 2 }, {
        officeFontRoutes: { arial: route },
        onTextRun: (run) => textRuns.push(run),
      });
    expect(textRuns.find((run) => run.cellRef === 'B1')?.width).toBe(colWidthToPx(10, 7));
    expect(paintedFonts.some((font) => font.includes('__exact_arial'))).toBe(true);
  });

  it('remeasures when a non-Calibri exact face becomes available on the next render', async () => {
    const ws = worksheet();
    const first = target();
    const firstRuns: XlsxTextRunInfo[] = [];
    await renderWorksheetViewport({ ws, styles: STYLES }, first.canvas,
      { row: 1, col: 1, rows: 1, cols: 2 }, { onTextRun: (run) => firstRuns.push(run) });
    expect(firstRuns.find((run) => run.cellRef === 'B1')?.width).toBe(colWidthToPx(10, 9));

    const second = target();
    const secondRuns: XlsxTextRunInfo[] = [];
    const route: OfficeFontFallbackRoute = {
      requestedFamily: 'Arial', family: '__exact_arial', source: 'local',
      resourceIdentity: 'office-local:test', weight: 400, style: 'normal',
      metric: { family: '__exact_arial' },
    };
    await renderWorksheetViewport({ ws, styles: STYLES }, second.canvas,
      { row: 1, col: 1, rows: 1, cols: 2 }, {
        officeFontRoutes: { arial: route }, onTextRun: (run) => secondRuns.push(run),
      });
    expect(secondRuns.find((run) => run.cellRef === 'B1')?.width).toBe(colWidthToPx(10, 7));
    expect(second.paintedFonts.some((font) => font.includes('__exact_arial'))).toBe(true);
  });
});
