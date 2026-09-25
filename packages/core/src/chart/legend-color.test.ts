import { describe, it, expect } from 'vitest';
import type { ChartSeries, ChartModel } from '../types/chart';
import { legendEntryColor, chartVariesColorsByPoint, CHART_PALETTE } from './renderer.js';

function series(partial: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...partial };
}

const pal = (i: number) => `#${CHART_PALETTE[i % CHART_PALETTE.length]}`;

describe('legendEntryColor', () => {
  describe('multi-series chart (bar/line) — one legend entry per series', () => {
    const s: ChartSeries[] = [
      series({ name: 'A', color: null }),
      series({ name: 'B', color: null }),
      series({ name: 'C', color: null }),
    ];

    it('falls back to the per-series palette index', () => {
      expect(legendEntryColor('bar', s, 0)).toBe(pal(0));
      expect(legendEntryColor('bar', s, 1)).toBe(pal(1));
      expect(legendEntryColor('bar', s, 2)).toBe(pal(2));
    });

    it('honors an explicit series color', () => {
      const withColor: ChartSeries[] = [
        series({ name: 'A', color: 'FF0000' }),
        series({ name: 'B', color: null }),
      ];
      expect(legendEntryColor('line', withColor, 0)).toBe('#FF0000');
      expect(legendEntryColor('line', withColor, 1)).toBe(pal(1));
    });
  });

  describe('pie / doughnut — one legend entry per category (data point of series[0])', () => {
    it('uses the per-index palette and matches slice colors, ignoring the series-level color', () => {
      // Excel sets a series-level solidFill on pie series; that must NOT
      // collapse every legend swatch to the same color. Each entry uses the
      // same palette index its slice does.
      const s: ChartSeries[] = [
        series({ name: 'Region', color: '4472C4', values: [10, 20, 30] }),
      ];
      expect(legendEntryColor('pie', s, 0)).toBe(pal(0));
      expect(legendEntryColor('pie', s, 1)).toBe(pal(1));
      expect(legendEntryColor('pie', s, 2)).toBe(pal(2));
      // Distinct colors, not all the series color.
      expect(legendEntryColor('pie', s, 0)).not.toBe(legendEntryColor('pie', s, 1));
    });

    it('honors explicit per-point dPt colors', () => {
      const s: ChartSeries[] = [
        series({
          name: 'Region',
          color: '4472C4',
          values: [10, 20, 30],
          dataPointColors: ['AA0000', null, 'CC0000'],
        }),
      ];
      expect(legendEntryColor('doughnut', s, 0)).toBe('#AA0000');
      // null inside the array -> palette fallback for that slice
      expect(legendEntryColor('doughnut', s, 1)).toBe(pal(1));
      expect(legendEntryColor('doughnut', s, 2)).toBe('#CC0000');
    });
  });

  describe('§21.2.2.227 varyColors single-series bar — one legend entry per point', () => {
    // The `varyByPoint` flag makes a bar legend resolve per DATA POINT of the
    // first series (like a pie) instead of per series (issue #931).
    const s: ChartSeries[] = [
      series({ name: 'Region', color: '4472C4', values: [10, 20, 30, 40] }),
    ];

    it('falls back to the per-point palette when varyByPoint is set', () => {
      expect(legendEntryColor('clusteredBar', s, 0, true)).toBe(pal(0));
      expect(legendEntryColor('clusteredBar', s, 1, true)).toBe(pal(1));
      expect(legendEntryColor('clusteredBar', s, 2, true)).toBe(pal(2));
      // Distinct colors, not all the single series color.
      expect(legendEntryColor('clusteredBar', s, 0, true)).not.toBe(
        legendEntryColor('clusteredBar', s, 1, true),
      );
    });

    it('honors explicit dPt colors carried in dataPointColors', () => {
      const withAccents: ChartSeries[] = [
        series({
          name: 'Region',
          color: '4472C4',
          values: [10, 20, 30, 40],
          dataPointColors: ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000'],
        }),
      ];
      expect(legendEntryColor('clusteredBar', withAccents, 0, true)).toBe('#4472C4');
      expect(legendEntryColor('clusteredBar', withAccents, 2, true)).toBe('#A5A5A5');
    });

    it('without varyByPoint, a bar legend stays per-series', () => {
      expect(legendEntryColor('clusteredBar', s, 0)).toBe(pal(0));
      expect(legendEntryColor('clusteredBar', s, 1)).toBe(pal(1));
    });
  });
});

describe('chartVariesColorsByPoint', () => {
  const model = (partial: Partial<ChartModel>): ChartModel =>
    ({ chartType: 'clusteredBar', series: [], varyColors: false, ...partial }) as ChartModel;
  const oneSeries = [{ name: 'A', color: null, values: [1, 2, 3] }] as ChartSeries[];
  const twoSeries = [
    { name: 'A', color: null, values: [1, 2] },
    { name: 'B', color: null, values: [3, 4] },
  ] as ChartSeries[];

  it('is true for a single-series bar/column with varyColors set', () => {
    expect(chartVariesColorsByPoint(model({ chartType: 'clusteredBar', series: oneSeries, varyColors: true }))).toBe(true);
    expect(chartVariesColorsByPoint(model({ chartType: 'clusteredBarH', series: oneSeries, varyColors: true }))).toBe(true);
    expect(chartVariesColorsByPoint(model({ chartType: 'stackedBar', series: oneSeries, varyColors: true }))).toBe(true);
  });

  it('varies lone line, scatter, and non-filled radar groups but not filled radar or area', () => {
    expect(chartVariesColorsByPoint(model({ chartType: 'line', series: oneSeries, varyColors: true }))).toBe(true);
    expect(chartVariesColorsByPoint(model({ chartType: 'scatter', series: oneSeries, varyColors: true }))).toBe(true);
    expect(chartVariesColorsByPoint(model({ chartType: 'radar', series: oneSeries, varyColors: true }))).toBe(true);
    expect(chartVariesColorsByPoint(model({
      chartType: 'radar', radarStyle: 'filled', series: oneSeries, varyColors: true,
    }))).toBe(false);
    expect(chartVariesColorsByPoint(model({ chartType: 'area', series: oneSeries, varyColors: true }))).toBe(false);
  });

  it('is false without the flag or for multiple series', () => {
    expect(chartVariesColorsByPoint(model({ series: oneSeries, varyColors: false }))).toBe(false);
    expect(chartVariesColorsByPoint(model({ series: twoSeries, varyColors: true }))).toBe(false);
    expect(chartVariesColorsByPoint(model({ chartType: 'pie', series: oneSeries, varyColors: true }))).toBe(false);
  });

  it('uses point entries for a lone bubble series whose x source is textual', () => {
    expect(chartVariesColorsByPoint(model({
      chartType: 'bubble',
      varyColors: false,
      series: [{ ...oneSeries[0], bubbleXSourceIsString: true }],
    }))).toBe(true);
    expect(chartVariesColorsByPoint(model({
      chartType: 'bubble',
      varyColors: true,
      series: [{ ...oneSeries[0], bubbleXSourceIsString: false }],
    }))).toBe(false);
  });
});
