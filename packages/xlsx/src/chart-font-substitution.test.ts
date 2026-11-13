import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { OFFICE_FONT_SUBSTITUTIONS } from '@silurus/ooxml-core/internal/font-sources';
import { renderViewport } from './renderer.js';
import { XlsxFontDemand } from './font-source-load.js';
import type { ParsedWorkbook, Styles, Worksheet, XlsxLoadedFontSources } from './types.js';

// Charts demand the faces they name and paint through the same loaded
// substitution map as cell text (LoadOptions.fonts).

const CHART = {
  chartType: 'clusteredBar', title: 'Chart Title', categories: ['A', 'B'],
  series: [{ name: 'Series', color: null, values: [1, 2] }],
  showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null, valAxisTitle: null,
  catAxisHidden: false, valAxisHidden: false, catAxisLineHidden: false, valAxisLineHidden: false,
  plotAreaBg: null, chartBg: null, showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
  valAxisMajorTickMark: 'out', catAxisMajorTickMark: 'out', titleFontSizeHpt: null, titleFontColor: null,
  titleFontFace: 'Chart Face', catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
  dataLabelFontSizeHpt: null, subtotalIndices: [],
} as unknown as ChartModel;

const STYLES = { fonts: [{ name: 'Calibri', bold: false, italic: false, size: 11 }], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as unknown as Styles;

function worksheet(): Worksheet {
  return {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], shapeGroups: [],
    charts: [{ fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 8, toColOff: 0, toRow: 16, toRowOff: 0, chart: CHART }],
  } as unknown as Worksheet;
}

function paintedFonts(fontSources?: XlsxLoadedFontSources): Array<{ text: string; font: string }> {
  const calls: Array<{ text: string; font: string }> = [];
  let font = '10px sans-serif';
  const noop = () => {};
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    canvas: { width: 1200, height: 800 },
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 7, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
    fillText: (text: string) => { calls.push({ text, font }); },
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    getLineDash: () => [],
  } as Record<string | symbol, unknown>, {
    get(target, key) { return key in target ? target[key] : noop; },
    set(target, key, value) { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, worksheet(), STYLES, { row: 1, col: 1, rows: 30, cols: 12 },
    { dpr: 1, loadedImages: new Map(), ...(fontSources ? { fontSources } : {}) });
  return calls;
}

describe('XLSX chart text and font sources', () => {
  it('demands the faces a chart names', () => {
    const demand = new XlsxFontDemand({ styles: STYLES, sharedStrings: [] } as unknown as ParsedWorkbook, 'jp', OFFICE_FONT_SUBSTITUTIONS);
    demand.take();
    demand.worksheetParts(worksheet());
    expect(demand.take()?.families).toEqual(['Chart Face']);
  });

  it('paints chart text through the private family of the loaded unscoped substitute', () => {
    const title = (calls: ReturnType<typeof paintedFonts>) => calls.find((call) => call.text === 'Chart Title')?.font ?? '';
    const offline = title(paintedFonts());
    expect(offline).toContain('"Chart Face"');
    expect(offline).not.toContain('ooxml-font-source');
    const loaded = (script?: 'arabic') => ({
      substitutions: new Map([['chart face', { family: 'Served Face', ...(script ? { script } : {}) }]]),
      regularAliases: new Map<string, string>(),
      families: new Map([['served face', 'ooxml-font-source-1-0']]),
    });
    expect(title(paintedFonts(loaded()))).toContain('"Chart Face", "ooxml-font-source-1-0"');
    // A script-scoped substitute never reaches chart text.
    expect(title(paintedFonts(loaded('arabic')))).toBe(offline);
  });
});
