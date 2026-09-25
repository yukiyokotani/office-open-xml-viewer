import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import {
  deletedLegendEntryIndices,
  legendEntryGlobalIndex,
  legendEntryIsVisible,
  legendEntryRanges,
  legendSeriesHasVisibleEntry,
} from './legend-entry-plan.js';

const combo = (): ChartModel => ({
  chartType: 'clusteredBar',
  categories: ['A', 'B', 'C'],
  series: [
    {
      name: 'ordinary', color: '112233', values: [1, 2, 3],
      trendLines: [{ trendlineType: 'linear' }],
    },
    { name: 'varying', color: '445566', values: [4, 5, 6] },
    { name: 'after', color: '778899', values: [7, 8, 9] },
  ],
  plotGroups: [
    {
      kind: 'line', seriesStart: 0, seriesCount: 1,
      categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
    },
    {
      kind: 'bar', seriesStart: 1, seriesCount: 1, varyColors: true,
      categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
    },
    {
      kind: 'line', seriesStart: 2, seriesCount: 1,
      categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
    },
  ],
} as ChartModel);

describe('global legend entry planning', () => {
  it('flattens series, point-driven groups, and visible trendlines in source order', () => {
    const chart = combo();
    const withoutTrendlines = legendEntryRanges(chart);
    expect(withoutTrendlines).toEqual([
      { firstIndex: 0, count: 1, pointDriven: false },
      { firstIndex: 1, count: 3, pointDriven: true },
      { firstIndex: 4, count: 1, pointDriven: false },
    ]);
    const withTrendlines = legendEntryRanges(chart, true);
    expect(withTrendlines).toEqual([
      { firstIndex: 0, count: 1, pointDriven: false },
      { firstIndex: 2, count: 3, pointDriven: true },
      { firstIndex: 5, count: 1, pointDriven: false },
    ]);
    expect(legendEntryGlobalIndex(withTrendlines, 1, 2)).toBe(4);
    expect(legendEntryGlobalIndex(withTrendlines, 1, 3)).toBeNull();
  });

  it('applies deleted legend indexes in the same flattened domain', () => {
    const chart = { ...combo(), legendEntries: [{ idx: 2, deleted: true }] };
    const ranges = legendEntryRanges(chart, true);
    const deleted = deletedLegendEntryIndices(chart);
    expect(legendEntryIsVisible(ranges, deleted, 1, 0)).toBe(false);
    expect(legendEntryIsVisible(ranges, deleted, 1, 1)).toBe(true);
    expect(legendSeriesHasVisibleEntry(ranges, deleted, 1)).toBe(true);
  });

  it('treats ofPie as a point-driven category legend', () => {
    const chart = {
      chartType: 'ofPie', categories: ['A', 'B', 'C'],
      series: [{ name: 'Pie', color: '112233', values: [1, 2, 3] }],
      legendEntries: [{ idx: 0, deleted: true }],
    } as ChartModel;
    const ranges = legendEntryRanges(chart, true);
    const deleted = deletedLegendEntryIndices(chart);
    expect(ranges).toEqual([{ firstIndex: 0, count: 3, pointDriven: true }]);
    expect(legendEntryIsVisible(ranges, deleted, 0, 0)).toBe(false);
    expect(legendEntryIsVisible(ranges, deleted, 0, 1)).toBe(true);
  });
});
