import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';
import { renderChart } from './renderer.js';
import { classicLineRunIndices, smoothCurveExtent } from './smooth-extent.js';
import { computeSecondaryAxis } from './shared/secondary-axis.js';

function labelRecorder(): { ctx: CanvasRenderingContext2D; labels: string[] } {
  const labels: string[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const ctx = new Proxy(state, {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      if (prop === 'measureText') return (text: string) => ({ width: String(text).length * 7 });
      if (prop === 'fillText') return (text: string) => { labels.push(text); };
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      return () => undefined;
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, labels };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };
const DIP = [6, 0, 0, 0, 0, 7, 3];
const CATS = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

function chart(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'line', title: null, categories: CATS, series: [], showDataLabels: false,
    valMin: null, valMax: null, catAxisTitle: null, valAxisTitle: null,
    catAxisHidden: false, valAxisHidden: false, plotAreaBg: null, chartBg: null,
    showLegend: false, legendPos: null, titleFontSizeHpt: null, titleFontColor: null,
    titleFontFace: null, catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
    ...over,
  } as ChartModel;
}

function paintedLabels(model: ChartModel): string[] {
  const recorder = labelRecorder();
  renderChart(recorder.ctx, model, RECT, 1);
  return recorder.labels;
}

describe('smoothCurveExtent', () => {
  it('covers the Catmull-Rom overshoot of a run leaving a flat stretch', () => {
    const [lo, hi] = smoothCurveExtent([DIP]) as [number, number];
    // Segment 0 -> 0 before 7: control points 0 and 0 - 7/6; minimum at t = 2/3.
    expect(lo).toBeCloseTo(-14 / 27, 12);
    expect(hi).toBeGreaterThan(7);
  });

  it('keeps extrema of tiny magnitudes', () => {
    const [lo] = smoothCurveExtent([DIP.map(v => v * 1e-15)]) as [number, number];
    expect(lo).toBeCloseTo(-14 / 27 * 1e-15, 25);
  });

  it('ignores straight short runs and evaluates a log axis in log space', () => {
    expect(smoothCurveExtent([[0, 5], [9]])).toBeNull();
    const [lo] = smoothCurveExtent([[1000, 1, 1, 1000]], 10) as [number, number];
    expect(lo).toBeLessThan(1);
    expect(lo).toBeGreaterThan(0);
    // A non-positive value has no position on a log axis: no curve.
    expect(smoothCurveExtent([[1000, 0, 1, 1000]], 10)).toBeNull();
  });
});

describe('classicLineRunIndices', () => {
  const s = series({ values: [1, null, 2, 3, 4], sourceHidden: [false, false, false, true, false] });
  it('follows dispBlanksAs and breaks at hidden cells', () => {
    expect(classicLineRunIndices(s, 5, 'gap')).toEqual([[0], [2], [4]]);
    expect(classicLineRunIndices(s, 5, 'span')).toEqual([[0, 2], [4]]);
    expect(classicLineRunIndices(s, 5, 'zero')).toEqual([[0, 1, 2], [4]]);
    expect(classicLineRunIndices(s, 5, 'gap', true)).toEqual([[0, 1, 2], [4]]);
  });
});

describe('automatic value axes cover a drawn smoothed curve', () => {
  it('line: smoothed widens below zero, straight does not', () => {
    expect(paintedLabels(chart({ series: [series({ values: DIP, smooth: true })] }))).toContain('-1');
    expect(paintedLabels(chart({ series: [series({ values: DIP })] }))).not.toContain('-1');
  });

  it('line: hiding only point segments still widens; removing the series line does not', () => {
    const hiddenSegments = [2, 3, 4, 5].map(idx => ({ idx, lineHidden: true }));
    expect(paintedLabels(chart({
      series: [series({ values: DIP, smooth: true, dataPointOverrides: hiddenSegments })],
    }))).toContain('-1');
    expect(paintedLabels(chart({
      series: [series({ values: DIP, smooth: true, lineHidden: true })],
    }))).not.toContain('-1');
  });

  it('percentStacked line snaps to the percent unit of the widened range', () => {
    const labels = paintedLabels(chart({
      chartType: 'stackedLinePct',
      series: [
        series({ values: [1, 1, 0, 1, 1, 1, 1], smooth: true }),
        series({ values: [0, 0, 1, 0, 0, 0, 0], smooth: true }),
      ],
      plotGroups: [{
        kind: 'line', seriesStart: 0, seriesCount: 2, grouping: 'percentStacked',
        categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
      }],
    }));
    expect(labels).toContain('120%');
    expect(labels).not.toContain('110%');
  });

  it('a secondary percentStacked axis snaps a crossed bound to the percent unit', () => {
    const scale = (extra: number[]) => computeSecondaryAxis(
      { hidden: false } as NonNullable<ChartModel['secondaryValAxis']>,
      [series({ values: [0, 100], useSecondaryAxis: true })],
      300, 'y', true, false, undefined, undefined, extra,
    );
    expect(scale([-2.3, 100])?.min).toBe(-20);
    expect(scale([0, 106])?.max).toBe(120);
  });

  it('stacked line', () => {
    const labels = paintedLabels(chart({
      chartType: 'stackedLine',
      series: [
        series({ values: DIP, smooth: true }),
        series({ values: [1, 0, 0, 0, 0, 1, 1], smooth: true }),
      ],
      plotGroups: [{
        kind: 'line', seriesStart: 0, seriesCount: 2, grouping: 'stacked',
        categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
      }],
    }));
    expect(labels).toContain('-1');
  });

  it('secondary-axis line over columns', () => {
    const labels = paintedLabels(chart({
      chartType: 'clusteredBar',
      series: [
        series({ values: [20, 30, 20, 30, 20, 30, 20] }),
        series({ values: DIP.map(v => v * 100), smooth: true, seriesType: 'line', useSecondaryAxis: true }),
      ],
      secondaryValAxis: { hidden: false } as ChartModel['secondaryValAxis'],
      plotGroups: [
        { kind: 'bar', seriesStart: 0, seriesCount: 1, barDirection: 'col', grouping: 'clustered',
          categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none' },
        { kind: 'line', seriesStart: 1, seriesCount: 1, grouping: 'standard',
          categoryAxis: 'primary', valueAxis: 'secondary', seriesAxis: 'none' },
      ],
    }));
    expect(labels.some(label => /^-\d+$/.test(label) && Number(label) <= -100)).toBe(true);
  });

  it('smooth scatter, but not a smoothed series whose line is removed', () => {
    const scatter = (lineHidden: boolean): ChartModel => chart({
      chartType: 'scatter',
      scatterStyle: 'smoothMarker',
      categories: ['1', '2', '3', '4', '5', '6', '7'],
      series: [series({
        categories: ['1', '2', '3', '4', '5', '6', '7'], values: DIP, smooth: true, lineHidden,
      })],
    });
    expect(paintedLabels(scatter(false))).toContain('-1');
    expect(paintedLabels(scatter(true))).not.toContain('-1');
  });
});
