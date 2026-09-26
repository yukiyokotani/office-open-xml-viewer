import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const EMU_PER_PX = 9525;

const STYLES: Styles = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

const CHART = {
  chartType: 'clusteredBar',
  title: null,
  categories: [],
  series: [],
  showDataLabels: false,
  valMin: null,
  valMax: null,
  catAxisTitle: null,
  valAxisTitle: null,
  catAxisHidden: false,
  valAxisHidden: false,
  catAxisLineHidden: false,
  valAxisLineHidden: false,
  plotAreaBg: null,
  chartBg: 'ABCDEF',
  showLegend: false,
  legendPos: null,
  catAxisCrossBetween: 'between',
  valAxisMajorTickMark: 'out',
  catAxisMajorTickMark: 'out',
  titleFontSizeHpt: null,
  titleFontColor: null,
  titleFontFace: null,
  catAxisFontSizeHpt: null,
  valAxisFontSizeHpt: null,
  dataLabelFontSizeHpt: null,
  subtotalIndices: [],
} as ChartModel;

function chartFill(viewport: { row: number; col: number }): { x: number; y: number } | undefined {
  const fills: Array<{ x: number; y: number; color: string }> = [];
  const state: Record<string, unknown> = {
    canvas: { width: 800, height: 400 },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '11px sans-serif',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillRect(x: number, y: number) {
      fills.push({ x, y, color: String(state.fillStyle) });
    },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noop = () => {};
  const ctx = new Proxy(state, {
    get(target, property) { return property in target ? target[property as string] : noop; },
    set(target, property, value) { target[property as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const worksheet = {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    // Anchored in C2 with a 10 px offset.
    charts: [{
      fromCol: 2,
      fromColOff: 10 * EMU_PER_PX,
      fromRow: 1,
      fromRowOff: 0,
      toCol: 5,
      toColOff: 0,
      toRow: 8,
      toRowOff: 0,
      chart: CHART,
    }],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as unknown as Worksheet;
  renderViewport(ctx, worksheet, STYLES, { ...viewport, rows: 20, cols: 10 });
  return fills.find((fill) => fill.color === '#ABCDEF');
}

describe('viewport origin of anchored drawings', () => {
  it('registers drawings with the grid when the viewport starts below row/column 1', () => {
    const first = chartFill({ row: 1, col: 1 });
    expect(first).toBeDefined();
    // Rows and columns are 1-based; a start of 0 is clamped like the grid
    // bands, not treated as one default column/row before the sheet.
    expect(chartFill({ row: 0, col: 0 })).toEqual(first);
    expect(chartFill({ row: -3, col: -3 })).toEqual(first);
  });
});
