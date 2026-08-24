// CH — value-axis MAJOR gridlines must be painted UNDER the data series, not
// over them. PowerPoint draws `<c:valAx><c:majorGridlines>` beneath the plotted
// geometry so an opaque area fill occludes the gridlines inside its region
// (verified against private/sample-14.pdf slide-6: every gridline that falls
// inside the teal ARR area reads solid teal, only the gridlines above the fill
// top are visible). The bar/line/stock/scatter/waterfall/box renderers already
// stroke gridlines before their series; this pins that ordering for the AREA
// family (which historically drew fills first, then gridlines on top) and guards
// the others against regressing.
import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

// Ordered event recorder: logs each fill()/stroke() with the style in effect at
// the call, so we can assert the RELATIVE order of gridline strokes vs series
// fills. These fixtures have no other path fill, so `fill()` marks a series
// area fill; a
// thin hairline strokeStyle (the resolved gridline color, default `#e0e0e0`)
// marks a gridline. We tag events by role and check the first gridline precedes
// the first series fill.
type Ev = { op: 'fill' | 'stroke' | 'text'; fillStyle: string; strokeStyle: string; lineWidth: number; text?: string };

function orderedRecordingCtx(): { ctx: CanvasRenderingContext2D; events: Ev[] } {
  const events: Ev[] = [];
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
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'fill':
          return () => events.push({ op: 'fill', fillStyle: String(state.fillStyle), strokeStyle: String(state.strokeStyle), lineWidth: Number(state.lineWidth) });
        case 'stroke':
          return () => events.push({ op: 'stroke', fillStyle: String(state.fillStyle), strokeStyle: String(state.strokeStyle), lineWidth: Number(state.lineWidth) });
        case 'fillText':
          return (text: string) => events.push({ op: 'text', text, fillStyle: String(state.fillStyle), strokeStyle: String(state.strokeStyle), lineWidth: Number(state.lineWidth) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'moveTo': case 'lineTo': case 'arc': case 'fillRect':
        case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect':
        case 'strokeRect': case 'clearRect': case 'strokeText':
        case 'setLineDash': case 'translate': case 'rotate': case 'scale':
        case 'clip': case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, events };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'area',
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

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

// The default value-axis gridline is a thin hairline (`#e0e0e0`, 0.5 px). A
// series fill for the area family is an authored opaque solid fill. These
// predicates classify the recorded events by role.
const isGridlineStroke = (e: Ev): boolean =>
  e.op === 'stroke' && e.lineWidth <= 1 && e.strokeStyle.toLowerCase() === '#e0e0e0';
const isSeriesFill = (e: Ev): boolean => e.op === 'fill';

describe('CH — value-axis gridlines paint under the data series', () => {
  it('an area chart strokes its major gridlines BEFORE filling the series area', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [series({ name: 'ARR', values: [37, 40, 44, 48] })],
    }), RECT, 1);

    const firstGridline = rec.events.findIndex(isGridlineStroke);
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);

    expect(firstSeriesFill).toBeGreaterThanOrEqual(0); // the area fill happened
    expect(firstGridline).toBeGreaterThanOrEqual(0);    // a gridline was stroked
    // The gridline must be laid down first so the opaque/translucent area sits
    // on top of it (PowerPoint occludes gridlines inside the fill region).
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });

  it('a stacked area chart also strokes gridlines before any fill', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [
        series({ name: 'A', values: [10, 12, 14] }),
        series({ name: 'B', values: [5, 6, 7] }),
      ],
    }), RECT, 1);

    const firstGridline = rec.events.findIndex(isGridlineStroke);
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);
    expect(firstSeriesFill).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });

  it('paints a standard area chart in document order so the later series is on top', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['Q1', 'Q2'],
      series: [
        series({ name: 'North', color: '156082', values: [10, 15] }),
        series({ name: 'South', color: 'E97132', values: [18, 13] }),
      ],
    }), RECT, 1);

    expect(rec.events.filter(event => event.op === 'fill').map(event => event.fillStyle))
      .toEqual(['#156082', '#E97132']);
  });

  it('a stacked area/line combo stacks only area series and paints the line as an overlay', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B', 'C'],
      series: [
        series({
          name: 'Invisible baseline', seriesType: 'area', color: '00000000',
          lineHidden: true, values: [80, 90, 85],
        }),
        series({ name: 'Band', seriesType: 'area', color: '1696D2', values: [10, 12, 11] }),
        series({
          name: 'Forecast', seriesType: 'line', color: '000000', lineColor: '000000',
          showMarker: false, values: [85, 95, 91],
        }),
      ],
    }), RECT, 1);

    const fills = rec.events.filter(event => event.op === 'fill');
    expect(fills.map(event => event.fillStyle.toLowerCase())).toEqual(['#00000000', '#1696d2']);
    const lineOverlay = rec.events.find(event =>
      event.op === 'stroke'
      && event.strokeStyle.toLowerCase() === '#000000'
      && event.lineWidth > 1,
    );
    expect(lineOverlay).toBeDefined();
  });

  it('an area chart honors <c:minorGridlines> with the same count as line, all before the fill', () => {
    // Regression for #883: renderAreaChart silently ignored `<c:minorGridlines>`
    // (§21.2.2.129) while bar/line/stock honored it. With an identical value-axis
    // config the area renderer must stroke the same number of gridlines as the
    // line renderer (major + minor), and every gridline must precede the series
    // fill (same z-order as the major pass moved by #881).
    const cfg: Partial<ChartModel> = {
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [series({ name: 'ARR', values: [37, 40, 44, 48] })],
      valMax: 50,
      valAxisMajorUnit: 5,
      valAxisMinorGridlines: true,
      valAxisMinorUnit: 2.5,
    };

    const areaRec = orderedRecordingCtx();
    renderChart(areaRec.ctx, baseModel({ chartType: 'area', ...cfg }), RECT, 1);
    const lineRec = orderedRecordingCtx();
    renderChart(lineRec.ctx, baseModel({ chartType: 'line', ...cfg }), RECT, 1);

    const areaGrid = areaRec.events.filter(isGridlineStroke).length;
    const lineGrid = lineRec.events.filter(isGridlineStroke).length;
    // Line = 10 major + minor; area must match now that #883 is fixed.
    expect(lineGrid).toBeGreaterThan(0);
    expect(areaGrid).toBe(lineGrid);

    // Every gridline (major AND minor) sits before the first series fill.
    const firstFill = areaRec.events.findIndex(isSeriesFill);
    const lastGridline =
      areaRec.events.length - 1 -
      [...areaRec.events].reverse().findIndex(isGridlineStroke);
    expect(firstFill).toBeGreaterThanOrEqual(0);
    expect(lastGridline).toBeLessThan(firstFill);
  });

  it('a plain line/area combo keeps gridlines below the area fill', () => {
    // Guard the line renderer (which draws area-like fills? no — line strokes).
    // Line chart already gridlines-first; assert it does not regress by pinning a
    // filled marker/area does not precede the gridline. Uses the line renderer.
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A', 'B', 'C', 'D', 'E'],
      series: [series({ name: 'S', values: [1, 2, 3, 2, 4] })],
      // Explicit gridline color exercises the `grid.explicit` uniform-stroke path.
      valAxisGridlineColor: '888888',
    }), RECT, 1);
    const firstGridline = rec.events.findIndex(e => e.op === 'stroke' && e.strokeStyle.toLowerCase() === '#888888');
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);
    expect(firstGridline).toBeGreaterThanOrEqual(0);
    expect(firstSeriesFill).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });

  it('a horizontal bar/scatter dot plot paints category gridlines below its markers', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      catAxisMajorGridlines: true,
      catAxisGridlineColor: 'D9D9D9',
      series: [
        series({ name: 'range anchor', values: [0, 0] }),
        series({
          name: 'dot',
          seriesType: 'scatter',
          categories: ['2', '4'],
          values: [1, 2],
          markerSymbol: 'circle',
          markerFill: '1696D2',
        }),
      ],
      secondaryCatAxis: {
        min: 0, max: 5, title: null, hidden: true, lineHidden: true, majorTickMark: 'none',
      },
      secondaryValAxis: {
        min: 0, max: 3, title: null, hidden: true, lineHidden: true, majorTickMark: 'none',
      },
    }), RECT, 1);

    const firstGridline = rec.events.findIndex(event =>
      event.op === 'stroke' && event.strokeStyle.toLowerCase() === '#d9d9d9',
    );
    const firstMarker = rec.events.findIndex(event =>
      event.op === 'fill' && event.fillStyle.toLowerCase() === '#1696d2',
    );
    expect(firstGridline).toBeGreaterThanOrEqual(0);
    expect(firstMarker).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeLessThan(firstMarker);
  });

  it('paints every scatter error-range series before every marker and data label in a dot plot combo', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [
        series({ name: 'category anchor', values: [0, 0] }),
        series({
          name: 'start',
          seriesType: 'scatter',
          categories: ['0.2', '0.4'],
          values: [1, 2],
          markerSymbol: 'circle',
          markerFill: 'EC008C',
          errBars: [{
            dir: 'x', barType: 'plus', plus: [0.3, 0.3], minus: [null, null],
            noEndCap: false, color: 'E7E6E6', lineWidthEmu: 85_725,
          }],
          seriesDataLabels: {
            showVal: false, showCatName: true, showSerName: false, showPercent: false,
            position: 'l', fontColor: 'EC008C',
          },
        }),
        series({
          name: 'end',
          seriesType: 'scatter',
          categories: ['0.5', '0.7'],
          values: [1, 2],
          markerSymbol: 'circle',
          markerFill: '1596D2',
        }),
        // This final scatter series authors the full-width horizontal guides.
        // Per-series painting used to draw these AFTER the preceding series'
        // markers and labels, visibly crossing through both.
        series({
          name: 'guides',
          seriesType: 'scatter',
          categories: ['0', '0'],
          values: [1, 2],
          markerSymbol: 'none',
          showMarker: false,
          errBars: [{
            dir: 'x', barType: 'plus', plus: [1, 1], minus: [null, null],
            noEndCap: true, color: 'D9D9D9', lineWidthEmu: 9_525,
          }],
        }),
      ],
      secondaryCatAxis: {
        min: 0, max: 1, title: null, hidden: true, lineHidden: true, majorTickMark: 'none',
      },
      secondaryValAxis: {
        min: 0, max: 3, title: null, hidden: true, lineHidden: true, majorTickMark: 'none',
      },
    }), RECT, 1);

    let lastGuide = -1;
    for (let index = rec.events.length - 1; index >= 0; index--) {
      const event = rec.events[index];
      if (event.op === 'stroke' && event.strokeStyle.toLowerCase() === '#d9d9d9') {
        lastGuide = index;
        break;
      }
    }
    const firstMarker = rec.events.findIndex(event =>
      event.op === 'fill' && event.fillStyle.toLowerCase() === '#ec008c',
    );
    const firstDataLabel = rec.events.findIndex(event =>
      event.op === 'text' && event.fillStyle.toLowerCase() === '#ec008c',
    );
    expect(lastGuide).toBeGreaterThanOrEqual(0);
    expect(firstMarker).toBeGreaterThanOrEqual(0);
    expect(firstDataLabel).toBeGreaterThanOrEqual(0);
    expect(lastGuide).toBeLessThan(firstMarker);
    expect(lastGuide).toBeLessThan(firstDataLabel);
  });
});
