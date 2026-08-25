import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartSeries } from '../types/chart';
import {
  applyPlotVisibleOnly,
  hasFilteredScatterAutomaticPointStyle,
} from './source-visibility';

const series = (overrides: Partial<ChartSeries> = {}): ChartSeries => ({
  name: 'Series',
  color: null,
  values: [10, 20, 30, 40],
  ...overrides,
});

const chart = (overrides: Partial<ChartModel> = {}): ChartModel => ({
  chartType: 'line',
  title: null,
  categories: ['A', 'B', 'C', 'D'],
  series: [series()],
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
  valAxisMajorTickMark: 'cross',
  catAxisMajorTickMark: 'cross',
  titleFontSizeHpt: null,
  titleFontColor: null,
  titleFontFace: null,
  catAxisFontSizeHpt: null,
  valAxisFontSizeHpt: null,
  dataLabelFontSizeHpt: null,
  subtotalIndices: [],
  ...overrides,
});

describe('plotVisOnly source filtering', () => {
  it('does not alter cached data when the chart-level element is absent or false', () => {
    const input = chart({
      categorySourceHidden: [false, true, false, false],
      series: [series({ sourceHidden: [false, true, false, false] })],
    });
    expect(applyPlotVisibleOnly(input)).toBe(input);
    expect(applyPlotVisibleOnly({ ...input, plotVisibleOnly: false }).categories).toEqual([
      'A', 'B', 'C', 'D',
    ]);
  });

  it('compacts hidden category slots and reindexes every point-aligned payload once', () => {
    const input = chart({
      plotVisibleOnly: true,
      categorySourceHidden: [false, true, false, false],
      categoryLevels: [
        ['A', 'B', 'C', 'D'],
        ['First', '', 'Second', ''],
      ],
      series: [
        series({
          categories: ['a', 'b', 'c', 'd'],
          sourceHidden: [false, false, true, false],
          dataPointColors: ['111111', '222222', '333333', '444444'],
          dataLabelColors: ['AAAAAA', 'BBBBBB', 'CCCCCC', 'DDDDDD'],
          catFormatCodes: ['0', '0.0', '0%', 'General'],
          bubbleSizes: [1, 2, 3, 4],
          dataPointOverrides: [{ idx: 1, color: 'ABCDEF' }, { idx: 3, color: 'FEDCBA' }],
          dataLabelOverrides: [
            { idx: 1, text: 'hidden category' },
            { idx: 3, text: 'kept' },
          ],
          errBars: [{
            dir: 'y',
            barType: 'both',
            plus: [1, 2, 3, 4],
            minus: [5, 6, 7, 8],
            noEndCap: false,
          }],
        }),
        series({
          name: 'Hidden column',
          sourceHidden: [true, true, true, true],
        }),
      ],
    });

    const filtered = applyPlotVisibleOnly(input);
    expect(filtered.categories).toEqual(['A', 'C', 'D']);
    expect(filtered.categoryLevels).toEqual([
      ['A', 'C', 'D'],
      ['First', 'Second', ''],
    ]);
    expect(filtered.series[0]).toMatchObject({
      categories: ['a', 'c', 'd'],
      values: [10, null, 40],
      dataPointColors: ['111111', '333333', '444444'],
      dataLabelColors: ['AAAAAA', 'CCCCCC', 'DDDDDD'],
      catFormatCodes: ['0', '0%', 'General'],
      bubbleSizes: [1, null, 4],
      dataPointOverrides: [{ idx: 2, color: 'FEDCBA' }],
      dataLabelOverrides: [{ idx: 2, text: 'kept' }],
      errBars: [{ plus: [1, null, 4], minus: [5, null, 8] }],
    });
    expect(filtered.series).toHaveLength(1);
    expect(input.series[0].values).toEqual([10, 20, 30, 40]);
  });

  it('compacts scatter and bubble points when any required source cell is hidden', () => {
    const input = chart({
      chartType: 'bubble',
      plotVisibleOnly: true,
      categorySourceHidden: [true, false, false, false],
      series: [series({
        sourceHidden: [false, true, false, true],
        bubbleSizes: [10, 20, 30, 40],
        dataPointOverrides: [{ idx: 2, markerSize: 9 }, { idx: 3, markerSize: 11 }],
      })],
    });

    const filtered = applyPlotVisibleOnly(input);
    expect(filtered.categories).toEqual(['A', 'B', 'C', 'D']);
    expect(filtered.series[0]).toMatchObject({
      categories: ['A', 'C'],
      values: [10, 30],
      bubbleSizes: [10, 30],
      dataPointOverrides: [{ idx: 1, markerSize: 9 }],
    });
  });

  it('applies Excel automatic point styles after hidden scatter points are removed', () => {
    const filtered = applyPlotVisibleOnly(chart({
      chartType: 'scatter',
      scatterStyle: 'marker',
      plotVisibleOnly: true,
      themeAccentColors: ['156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'],
      series: [series({
        color: '156082',
        categories: ['1', '2', '3', '4', '5'],
        values: [10, 20, 90, 40, 50],
        sourceHidden: [false, false, true, false, false],
      })],
    }));

    expect(filtered.series[0].dataPointColors).toEqual([
      '156082', 'E97132', '196B24', '0F9ED5',
    ]);
    expect(filtered.series[0].dataPointOverrides).toMatchObject([
      { idx: 0, markerSymbol: 'diamond' },
      { idx: 1, markerSymbol: 'square' },
      { idx: 2, markerSymbol: 'triangle' },
      { idx: 3, markerSymbol: 'x' },
    ]);
    expect(hasFilteredScatterAutomaticPointStyle(filtered.series[0])).toBe(true);

    const coincidentalPublicSeries = {
      ...filtered.series[0],
      dataPointColors: [...(filtered.series[0].dataPointColors ?? [])],
      dataPointOverrides: (filtered.series[0].dataPointOverrides ?? []).map(point => ({ ...point })),
    };
    expect(hasFilteredScatterAutomaticPointStyle(coincidentalPublicSeries)).toBe(false);
  });
});
