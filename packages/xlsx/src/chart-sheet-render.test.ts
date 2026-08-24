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

const CHART: ChartModel = {
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
};

function context(width = 320, height = 180): {
  ctx: CanvasRenderingContext2D;
  fills: Array<{ x: number; y: number; w: number; h: number; color: string }>;
  texts: string[];
} {
  const fills: Array<{ x: number; y: number; w: number; h: number; color: string }> = [];
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    canvas: { width, height },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '11px sans-serif',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillRect(x: number, y: number, w: number, h: number) {
      fills.push({ x, y, w, h, color: String(state.fillStyle) });
    },
    fillText(text: string) { texts.push(text); },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noop = () => {};
  const ctx = new Proxy(state, {
    get(target, property) { return property in target ? target[property as string] : noop; },
    set(target, property, value) { target[property as string] = value; return true; },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills, texts };
}

describe('XLSX chart-sheet rendering', () => {
  it('uses the canvas origin for absolute anchors and omits worksheet chrome', () => {
    const worksheet: Worksheet = {
      name: 'Chart1',
      isChartSheet: true,
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
      charts: [{
        fromCol: 0,
        fromColOff: 0,
        fromRow: 0,
        fromRowOff: 0,
        toCol: 0,
        toColOff: 240 * EMU_PER_PX,
        toRow: 0,
        toRowOff: 140 * EMU_PER_PX,
        chart: CHART,
      }],
      defaultFontFamily: 'Calibri',
      defaultFontSize: 11,
    } as Worksheet;
    const recording = context();

    renderViewport(recording.ctx, worksheet, STYLES, { row: 1, col: 1, rows: 20, cols: 20 });

    expect(recording.fills).toContainEqual({ x: 0, y: 0, w: 240, h: 140, color: '#ABCDEF' });
    expect(recording.texts).not.toContain('A');
    expect(recording.texts).not.toContain('1');
  });

  it('keeps a manual chart legend invariant across worksheet zoom', () => {
    const names = ['First', 'Second', 'Third', 'Fourth'];
    const chart: ChartModel = {
      ...CHART,
      chartType: 'line',
      categories: ['A', 'B'],
      series: names.map((name, index) => ({
        name,
        color: null,
        values: [index + 1, index + 2],
      })),
      showLegend: true,
      legendPos: 'b',
      legendFontSizeHpt: 1200,
      legendManualLayout: {
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.055, y: 0.134, w: 0.9, h: 0.044,
      },
    };
    const worksheet: Worksheet = {
      name: 'Chart1',
      isChartSheet: true,
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
      charts: [{
        fromCol: 0,
        fromColOff: 0,
        fromRow: 0,
        fromRowOff: 0,
        toCol: 0,
        toColOff: 640 * EMU_PER_PX,
        toRow: 0,
        toRowOff: 600 * EMU_PER_PX,
        chart,
      }],
      defaultFontFamily: 'Calibri',
      defaultFontSize: 11,
    } as Worksheet;

    for (const cellScale of [0.1, 0.25, 0.5, 0.75, 1, 2]) {
      const recording = context(800, 700);
      renderViewport(
        recording.ctx,
        worksheet,
        STYLES,
        { row: 1, col: 1, rows: 40, cols: 40 },
        { cellScale },
      );
      expect(recording.texts.filter(text => names.includes(text)), `${cellScale * 100}%`)
        .toEqual(names);
    }
  });
});
