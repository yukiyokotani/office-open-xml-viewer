// Regression test for issue #803 (follow-up to PR #802):
//
//   A scatter series whose connecting line is suppressed — either the whole
//   chart group is markers-only (`<c:scatterStyle val="marker">`, the default)
//   or a specific series overrides the group line with `<a:noFill/>`
//   (§21.2.2.198, surfaced as `ChartSeries.lineHidden`) — must draw a MARKER
//   glyph as its legend key, matching Excel. Before this fix the legend key was
//   always a horizontal line swatch for any scatter chart (`legendSwatchStyle`
//   returned 'line' unconditionally for scatter), which read as a
//   line-with-no-markers series that the plot never draws.
//
//   A scatter series that DOES draw a connecting line (group scatterStyle =
//   line / lineMarker / smooth… and the series line is not hidden) keeps the
//   line swatch — the legend key must mirror what the plot actually renders.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'scatter',
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
    showLegend: true,
    legendPos: 'b',
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

/** Recording context that classifies the legend KEY drawn next to each label.
 *
 *  A legend key is drawn immediately before the label text at the same y. The
 *  bottom legend sits below the plot, so we scope classification to the legend
 *  band (`y >= LEGEND_Y_MIN`). Within that band the key is either:
 *    • a MARKER glyph — a `fill()`/`fillRect()` following a `beginPath` (arc /
 *      diamond / triangle / star / square), or an `x`/`plus`/`dash` stroke; OR
 *    • a LINE swatch — a 2-vertex horizontal `stroke()` (`drawLegendSwatch`'s
 *      'line' style).
 *  We record, per `fillText` in the band, whichever key kind was seen most
 *  recently. Plot-area strokes (y above the band) are ignored. */
const LEGEND_Y_MIN = 320; // RECT.h = 360; the bottom legend band starts well below the plot.

function recordingCtx(): {
  ctx: CanvasRenderingContext2D;
  keys: Array<'marker' | 'line' | 'line-marker'>;
  plotMarkerYs: number[];
} {
  const keys: Array<'marker' | 'line' | 'line-marker'> = [];
  const plotMarkerYs: number[] = [];
  let pathVerts = 0;
  let lastPathY = 0;   // y of the last moveTo/lineTo/arc vertex (band gating).
  let pendingMarker = false;
  let pendingLine = false;
  const inBand = (): boolean => lastPathY >= LEGEND_Y_MIN;
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    globalAlpha: 1,
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'save':
        case 'restore':
          return () => undefined;
        case 'beginPath':
          return () => { pathVerts = 0; };
        case 'moveTo':
        case 'lineTo':
        case 'bezierCurveTo':
        case 'quadraticCurveTo':
          return (_x: number, y: number) => { pathVerts += 1; lastPathY = y; };
        case 'arc':
          return (_cx: number, cy: number) => {
            pathVerts += 1;
            lastPathY = cy;
            if (cy < LEGEND_Y_MIN) plotMarkerYs.push(cy);
          };
        case 'fill':
          // A marker glyph (arc/diamond/triangle/star) fills a closed path.
          return () => { if (inBand()) pendingMarker = true; };
        case 'fillRect':
          return (_x: number, y: number) => { if (y >= LEGEND_Y_MIN) pendingMarker = true; };
        case 'stroke':
          return () => {
            if (!inBand()) return;
            // A legend line swatch is a 2-vertex horizontal stroke. x/plus/dash
            // markers also stroke but from a fresh beginPath with 2 vertices —
            // those are handled by 'marker' fills; the only 2-vertex stroke a
            // markers-only legend emits comes from the line swatch itself, so a
            // 2-vertex band stroke unambiguously means a line key here.
            if (pathVerts === 2) pendingLine = true;
          };
        case 'fillText':
          return (_text: string, _x: number, y: number) => {
            if (y < LEGEND_Y_MIN) return;
            if (pendingLine && pendingMarker) keys.push('line-marker');
            else if (pendingLine) keys.push('line');
            else if (pendingMarker) keys.push('marker');
            pendingLine = false;
            pendingMarker = false;
          };
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    keys,
    plotMarkerYs,
  };
}

const scatterModel = (over: Partial<ChartModel>, serOver: Partial<ChartSeries> = {}): ChartModel =>
  baseModel({
    categories: ['1', '2', '3'],
    series: [series({
      name: 'S',
      color: '4f81bd',
      values: [10, 20, 15],
      categories: ['1', '2', '3'],
      showMarker: true,
      ...serOver,
    })],
    ...over,
  });

describe('scatter legend key reflects the plotted mark (#803, §21.2.2.42/.198)', () => {
  it('draws a LINE + MARKER key for Excel automatic marker scatter style', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, scatterModel({ scatterStyle: 'marker' }), RECT, 1);
    expect(rec.keys).toContain('line-marker');
  });

  it('draws a MARKER key when a series line is noFill (lineHidden) despite lineMarker style', () => {
    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      scatterModel({ scatterStyle: 'lineMarker' }, { lineHidden: true }),
      RECT, 1,
    );
    expect(rec.keys).toContain('marker');
    expect(rec.keys).not.toContain('line');
  });

  it('honors the series marker SYMBOL for the legend key (square)', () => {
    // A square marker draws a fillRect — still classified as 'marker'. The point
    // is that a non-line scatter never emits a line swatch.
    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      scatterModel({ scatterStyle: 'marker' }, { markerSymbol: 'square' }),
      RECT, 1,
    );
    expect(rec.keys).toEqual(['line-marker']);
  });

  it('draws a LINE + MARKER key for a lineMarker scatter', () => {
    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      scatterModel({ scatterStyle: 'lineMarker' }, { lineHidden: null }),
      RECT, 1,
    );
    expect(rec.keys).toContain('line-marker');
  });

  it('keeps the LINE key for a lineNoMarker scatter', () => {
    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      scatterModel({ scatterStyle: 'lineNoMarker' }),
      RECT, 1,
    );
    expect(rec.keys).toContain('line');
    expect(rec.keys).not.toContain('marker');
  });

  it('draws a LINE + MARKER key for a marked line chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({
        name: 'Debt',
        color: '1696D2',
        values: [10, 20],
        showMarker: true,
        markerSymbol: 'circle',
        markerFill: 'FFFFFF',
        markerLine: '1696D2',
      })],
    }), RECT, 1);
    expect(rec.keys).toContain('line-marker');
  });

  it('does not reuse chart-level X values for an explicitly empty later series', () => {
    const resolved = series({ name: 'resolved', values: [10], categories: null, showMarker: true });
    const baseline = recordingCtx();
    renderChart(
      baseline.ctx,
      scatterModel({ showLegend: false, scatterStyle: 'marker', series: [resolved] }),
      RECT,
      1,
    );

    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      scatterModel({
        showLegend: false,
        scatterStyle: 'marker',
        series: [
          resolved,
          series({
            name: 'unresolved',
            values: [1_000_000],
            categories: [],
            showMarker: true,
          }),
        ],
      }),
      RECT,
      1,
    );

    expect(baseline.plotMarkerYs).toHaveLength(1);
    expect(rec.plotMarkerYs).toHaveLength(1);
    expect(rec.plotMarkerYs[0]).toBeCloseTo(baseline.plotMarkerYs[0]);
  });
});
