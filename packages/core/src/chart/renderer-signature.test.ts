// Characterization test for the chart renderer (Phase 4 A1 — layout frame
// extraction). This test does NOT assert any specific "correct" appearance;
// it PINS the exact sequence of CanvasRenderingContext2D calls the renderer
// emits for a fixed battery of chart models. Any refactor that is truly a
// verbatim extraction of the layout math will emit a byte-identical call log.
//
// A recording mock captures every mutation (fillStyle, font, lineWidth, …)
// and every drawing op (fillRect, moveTo, lineTo, arc, fillText, …) with its
// numeric arguments. Sub-pixel coordinate drift — which pixel-hashing would
// hide behind rounding — shows up here as a changed argument. That makes this
// a stronger "1px unchanged" guarantee than a rendered-image snapshot.
//
// The battery spans all eight families dispatched by `renderChart`
// (bar/column, line, area, pie, doughnut, radar, scatter, waterfall) plus the
// combo (bar+line secondary axis), legend sides, axis titles, hidden axes,
// manual layouts, and data labels — i.e. every branch of the frame/pad code.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChartExChart } from './chart-ex-renderer.js';
import { renderChart } from './renderer.js';

const testChartEx = { render: renderChartExChart };

// ─── Recording mock context ─────────────────────────────────────────────────

/** Numeric args are rounded to 4 decimals so genuinely-identical floating math
 *  that differs only in the last ULP (e.g. `a*b` vs `b*a` reassociation) does
 *  not cause spurious diffs, while any real ≥0.0001px layout change is caught.
 *  4 decimals is far finer than a device pixel. */
function r(v: unknown): unknown {
  return typeof v === 'number' ? Math.round(v * 1e4) / 1e4 : v;
}

interface RecordingCtx {
  log: string[];
  measureText(t: string): { width: number };
}

/** Build a mock 2D context. `measureText` uses a deterministic width model
 *  (0.6em per char, CJK 1em) so text-dependent layout (legend widths, tick
 *  gutters, elision) is reproducible without a real font backend. The exact
 *  numbers don't matter — only that they're identical before and after. */
function makeRecordingCtx(): RecordingCtx & CanvasRenderingContext2D {
  const log: string[] = [];
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

  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };

  const handler: ProxyHandler<Record<string, unknown>> = {
    get(target, prop: string) {
      // Property reads (ctx.lineWidth etc.) return the tracked value.
      if (prop in state && typeof state[prop] !== 'function') {
        return state[prop];
      }
      switch (prop) {
        case 'log':
          return log;
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) {
              w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            }
            return { width: w };
          };
        case 'save':
          return () => log.push('save');
        case 'restore':
          return () => log.push('restore');
        case 'beginPath':
          return () => log.push('beginPath');
        case 'closePath':
          return () => log.push('closePath');
        case 'fill':
          return () => log.push(`fill fs=${state.fillStyle} ga=${r(state.globalAlpha)}`);
        case 'stroke':
          return () =>
            log.push(
              `stroke ss=${state.strokeStyle} lw=${r(state.lineWidth)} dash=${(state.__dash as number[] | undefined)?.join(',') ?? ''}`,
            );
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) =>
            log.push(`fillRect ${r(x)},${r(y)},${r(w)},${r(h)} fs=${state.fillStyle}`);
        case 'strokeRect':
          return (x: number, y: number, w: number, h: number) =>
            log.push(`strokeRect ${r(x)},${r(y)},${r(w)},${r(h)} ss=${state.strokeStyle} lw=${r(state.lineWidth)}`);
        case 'clearRect':
          return (x: number, y: number, w: number, h: number) =>
            log.push(`clearRect ${r(x)},${r(y)},${r(w)},${r(h)}`);
        case 'moveTo':
          return (x: number, y: number) => log.push(`moveTo ${r(x)},${r(y)}`);
        case 'lineTo':
          return (x: number, y: number) => log.push(`lineTo ${r(x)},${r(y)}`);
        case 'arc':
          return (x: number, y: number, rad: number, a0: number, a1: number) =>
            log.push(`arc ${r(x)},${r(y)},${r(rad)},${r(a0)},${r(a1)}`);
        case 'bezierCurveTo':
          return (a: number, b: number, c: number, d: number, e: number, f: number) =>
            log.push(`bezier ${r(a)},${r(b)},${r(c)},${r(d)},${r(e)},${r(f)}`);
        case 'quadraticCurveTo':
          return (a: number, b: number, c: number, d: number) =>
            log.push(`quad ${r(a)},${r(b)},${r(c)},${r(d)}`);
        case 'rect':
          return (x: number, y: number, w: number, h: number) =>
            log.push(`rect ${r(x)},${r(y)},${r(w)},${r(h)}`);
        case 'fillText':
          return (t: string, x: number, y: number) =>
            log.push(
              `fillText "${t}" ${r(x)},${r(y)} fs=${state.fillStyle} font=${state.font} al=${state.textAlign} bl=${state.textBaseline}`,
            );
        case 'strokeText':
          return (t: string, x: number, y: number) =>
            log.push(`strokeText "${t}" ${r(x)},${r(y)} ss=${state.strokeStyle}`);
        case 'setLineDash':
          return (d: number[]) => {
            state.__dash = d;
            log.push(d.length ? `setLineDash ${d.join(',')}` : 'setLineDash');
          };
        case 'translate':
          return (x: number, y: number) => log.push(`translate ${r(x)},${r(y)}`);
        case 'rotate':
          return (a: number) => log.push(`rotate ${r(a)}`);
        case 'scale':
          return (x: number, y: number) => log.push(`scale ${r(x)},${r(y)}`);
        case 'clip':
          return () => log.push('clip');
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'setTransform':
        case 'resetTransform':
        case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(target, prop: string, value) {
      // Record the mutation ONLY when the value actually changes, so the log
      // reflects the same "effective state churn" the renderer performs.
      if (state[prop] !== value) {
        log.push(`set ${prop}=${r(value)}`);
        state[prop] = value;
      }
      return true;
    },
  };

  return new Proxy(state, handler) as unknown as RecordingCtx & CanvasRenderingContext2D;
}

// ─── Model builders ─────────────────────────────────────────────────────────

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
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
    ...over,
  };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

const RECT: ChartRect = { x: 12, y: 20, w: 640, h: 360 };

// A representative model per branch. Each carries a title, legend, axis titles
// and data labels where the family supports them, to exercise the frame math.
function batteryModels(): Array<[string, ChartModel, ChartRect]> {
  const cats3 = ['Q1', 'Q2', 'Q3'];
  const cats12 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const barSeries = [
    series({ name: 'Alpha', values: [10, 24, 17] }),
    series({ name: 'Beta', values: [8, 15, 22] }),
  ];

  const out: Array<[string, ChartModel, ChartRect]> = [];

  out.push([
    'clusteredBar+title+legendR+dataLabels+axisTitles',
    baseModel({
      chartType: 'clusteredBar',
      title: 'Quarterly',
      categories: cats3,
      series: barSeries,
      showLegend: true,
      legendPos: 'r',
      showDataLabels: true,
      catAxisTitle: 'Quarter',
      valAxisTitle: 'Units',
    }),
    RECT,
  ]);

  out.push([
    'stackedBar+legendBottom',
    baseModel({
      chartType: 'stackedBar',
      title: 'Stacked',
      categories: cats3,
      series: barSeries,
      showLegend: true,
      legendPos: 'b',
    }),
    RECT,
  ]);

  out.push([
    'stackedBarPct+legendLeft',
    baseModel({
      chartType: 'stackedBarPct',
      categories: cats3,
      series: barSeries,
      showLegend: true,
      legendPos: 'l',
    }),
    RECT,
  ]);

  out.push([
    'clusteredBarH+dataLabels',
    baseModel({
      chartType: 'clusteredBarH',
      title: 'Horizontal',
      categories: cats3,
      series: barSeries,
      showDataLabels: true,
      showLegend: true,
    }),
    RECT,
  ]);

  out.push([
    'clusteredBarH+valAxisHidden(manual-ish)',
    baseModel({
      chartType: 'stackedBarH',
      categories: cats3,
      series: barSeries,
      valAxisHidden: true,
      catAxisHidden: true,
    }),
    RECT,
  ]);

  // Combo: bar + line on secondary axis.
  out.push([
    'combo bar+line secondary',
    baseModel({
      chartType: 'clusteredBar',
      title: 'Combo',
      categories: cats3,
      series: [
        series({ name: 'Rev', values: [100, 140, 120] }),
        series({ name: 'Margin', values: [12, 18, 15], seriesType: 'line', useSecondaryAxis: true }),
      ],
      showLegend: true,
      secondaryValAxis: {
        min: null,
        max: null,
        title: 'Margin %',
        hidden: false,
        majorTickMark: 'out',
        lineHidden: false,
      },
    }),
    RECT,
  ]);

  out.push([
    'line+markers+legend+axisTitles+12cats',
    baseModel({
      chartType: 'line',
      title: 'Trend',
      categories: cats12,
      series: [
        series({ name: 'A', values: [3, 5, 4, 6, 7, 5, 8, 6, 9, 7, 10, 8], showMarker: true }),
        series({ name: 'B', values: [2, 3, 5, 4, 6, 8, 7, 9, 6, 8, 7, 9], showMarker: true }),
      ],
      showLegend: true,
      catAxisTitle: 'Month',
      valAxisTitle: 'Value',
      showDataLabels: true,
    }),
    RECT,
  ]);

  out.push([
    'stackedLine+midCat',
    baseModel({
      chartType: 'stackedLine',
      categories: cats3,
      series: barSeries,
      catAxisCrossBetween: 'midCat',
    }),
    RECT,
  ]);

  out.push([
    'area+legend',
    baseModel({
      chartType: 'area',
      title: 'Area',
      categories: cats12,
      series: [series({ name: 'A', values: [3, 5, 4, 6, 7, 5, 8, 6, 9, 7, 10, 8] })],
      showLegend: true,
      catAxisTitle: 'Month',
    }),
    RECT,
  ]);

  out.push([
    'stackedArea',
    baseModel({
      chartType: 'stackedArea',
      categories: cats3,
      series: barSeries,
      showLegend: true,
      legendPos: 't',
    }),
    RECT,
  ]);

  out.push([
    'pie+legendR+dataLabels',
    baseModel({
      chartType: 'pie',
      title: 'Share',
      categories: cats3,
      series: [series({ name: 'Region', values: [30, 45, 25] })],
      showLegend: true,
      showDataLabels: true,
    }),
    RECT,
  ]);

  out.push([
    'doughnut+legendBottom',
    baseModel({
      chartType: 'doughnut',
      categories: cats3,
      series: [series({ name: 'Region', values: [30, 45, 25], dataPointColors: ['AA0000', null, 'CC0000'] })],
      showLegend: true,
      legendPos: 'b',
      showDataLabels: true,
    }),
    RECT,
  ]);

  out.push([
    'radar+filled+legend',
    baseModel({
      chartType: 'radar',
      title: 'Profile',
      categories: ['A', 'B', 'C', 'D', 'E'],
      series: [
        series({ name: 'X', values: [4, 3, 5, 2, 4] }),
        series({ name: 'Y', values: [2, 5, 3, 4, 3] }),
      ],
      radarStyle: 'filled',
      showLegend: true,
    }),
    RECT,
  ]);

  out.push([
    'scatter+lineMarker+axisTitles',
    baseModel({
      chartType: 'scatter',
      title: 'XY',
      categories: [],
      series: [
        series({
          name: 'S1',
          values: [2, 4, 3, 5, 6],
          categories: ['1', '2', '3', '4', '5'],
          markerSymbol: 'diamond',
        }),
      ],
      scatterStyle: 'lineMarker',
      catAxisTitle: 'X',
      valAxisTitle: 'Y',
      showLegend: true,
    }),
    RECT,
  ]);

  out.push([
    'bubble',
    baseModel({
      chartType: 'bubble',
      categories: [],
      series: [
        series({
          name: 'B1',
          values: [2, 4, 3],
          categories: ['1', '2', '3'],
          bubbleSizes: [5, 10, 7],
        }),
      ],
    }),
    RECT,
  ]);

  out.push([
    'waterfall',
    baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Q1', 'Q2', 'End'],
      series: [series({ name: 'W', values: [100, 30, -20, 110] })],
      subtotalIndices: [3],
    }),
    RECT,
  ]);

  out.push([
    'waterfall+valAxisHidden',
    baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Q1', 'Q2', 'End'],
      series: [series({ name: 'W', values: [100, 30, -20, 110] })],
      subtotalIndices: [3],
      valAxisHidden: true,
    }),
    RECT,
  ]);

  out.push([
    'no-data',
    baseModel({ chartType: 'clusteredBar', categories: [], series: [] }),
    RECT,
  ]);

  return out;
}

// ─── The test ────────────────────────────────────────────────────────────────

describe('renderChart draw-call signature (characterization)', () => {
  for (const [label, model, rect] of batteryModels()) {
    it(`is stable for: ${label}`, () => {
      const ctx = makeRecordingCtx();
      renderChart(
        ctx,
        model,
        rect,
        1.05,
        undefined,
        undefined,
        undefined,
        undefined,
        testChartEx,
      );
      // Snapshot the full ordered call log. A verbatim layout extraction must
      // not change a single entry.
      expect(ctx.log.join('\n')).toMatchSnapshot();
    });
  }
});
