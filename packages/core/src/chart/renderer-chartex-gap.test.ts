import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';
import { renderChartExChart } from './chart-ex-renderer.js';
import { renderChart } from './renderer.js';

const testChartEx = { render: renderChartExChart };

interface RectCall {
  x: number;
  y: number;
  w: number;
  h: number;
}

function recordingContext(): { ctx: CanvasRenderingContext2D; rects: RectCall[] } {
  const rects: RectCall[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state) return state[prop];
      switch (prop) {
        case 'measureText':
          return (text: string) => ({ width: String(text).length * 6 });
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) => {
            rects.push({ x, y, w, h });
          };
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save':
        case 'restore':
        case 'beginPath':
        case 'closePath':
        case 'fill':
        case 'stroke':
        case 'strokeRect':
        case 'fillText':
        case 'moveTo':
        case 'lineTo':
        case 'arc':
        case 'bezierCurveTo':
        case 'quadraticCurveTo':
        case 'rect':
        case 'clip':
        case 'clearRect':
        case 'strokeText':
        case 'setLineDash':
        case 'translate':
        case 'rotate':
        case 'scale':
        case 'setTransform':
        case 'resetTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_target, prop: string, value) {
      state[prop] = value;
      return true;
    },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    rects,
  };
}

function series(overrides: Partial<ChartSeries> = {}): ChartSeries {
  return { name: 'Series 1', color: '4472C4', values: [10, 20, 15], ...overrides };
}

function model(overrides: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'waterfall',
    title: null,
    categories: ['A', 'B', 'C'],
    series: [series()],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: true,
    valAxisHidden: true,
    catAxisLineHidden: true,
    valAxisLineHidden: true,
    plotAreaBg: null,
    chartBg: null,
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
    ...overrides,
  };
}

type ChartExGapFamily =
  | 'waterfall'
  | 'clusteredColumn'
  | 'histogram'
  | 'funnel'
  | 'boxWhisker';

function familyModel(family: ChartExGapFamily, barGapWidth?: number): ChartModel {
  const common = { barGapWidth };
  switch (family) {
    case 'waterfall':
      return model({ chartType: family, ...common });
    case 'clusteredColumn':
      return model({ chartType: family, ...common });
    case 'histogram':
      return model({
        chartType: family,
        ...common,
        categories: [],
        series: [series({ values: [1, 2, 3, 4, 5, 6] })],
        chartexHistogramBinning: { binCount: 3, intervalClosed: 'l' },
      });
    case 'funnel':
      return model({ chartType: family, ...common });
    case 'boxWhisker':
      return model({
        chartType: family,
        ...common,
        categories: ['A'],
        series: [series({ values: [] })],
        chartexBox: {
          categories: ['A'],
          series: [{
            name: 'Series 1',
            color: '4472C4',
            valuesByCategory: [[1, 2, 3, 4, 5]],
            meanMarker: false,
            meanLine: false,
            showOutliers: false,
            showNonoutliers: false,
            quartileMethod: 'inclusive',
          }],
        },
      });
  }
}

function renderRects(chart: ChartModel, bounds: ChartRect): RectCall[] {
  const recording = recordingContext();
  renderChart(
    recording.ctx,
    chart,
    bounds,
    1,
    undefined,
    undefined,
    undefined,
    undefined,
    testChartEx,
  );
  return recording.rects;
}

const BOUNDS: ChartRect[] = [
  { x: 0, y: 0, w: 360, h: 360 },
  { x: 0, y: 0, w: 480, h: 240 },
  { x: 0, y: 0, w: 240, h: 480 },
];

describe('ChartEx omitted category gap policy (#1227)', () => {
  for (const family of [
    'waterfall',
    'clusteredColumn',
    'histogram',
    'funnel',
    'boxWhisker',
  ] as const) {
    it(`${family}: omitted gap has the same geometry as authored 33% for square, wide, and tall bounds`, () => {
      for (const bounds of BOUNDS) {
        expect(renderRects(familyModel(family), bounds)).toEqual(
          renderRects(familyModel(family, 33), bounds),
        );
      }
    });

    it(`${family}: an authored 80% gap remains authoritative`, () => {
      const bounds = BOUNDS[0];
      expect(renderRects(familyModel(family, 80), bounds)).not.toEqual(
        renderRects(familyModel(family), bounds),
      );
    });
  }

  it('legacy bar/column omission retains the 150% geometry', () => {
    const bounds = BOUNDS[0];
    const legacy = (barGapWidth?: number): ChartModel => model({
      chartType: 'clusteredBar',
      barGapWidth,
    });
    expect(renderRects(legacy(), bounds)).toEqual(renderRects(legacy(150), bounds));
    expect(renderRects(legacy(), bounds)).not.toEqual(renderRects(legacy(33), bounds));
  });
});
