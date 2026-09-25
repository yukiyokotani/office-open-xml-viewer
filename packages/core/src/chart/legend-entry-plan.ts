import type { ChartModel } from '../types/chart.js';
import { chartPlotGroupForSeries, chartSeriesVariesByPoint } from './effective-style.js';

/** Chart families whose legend identifies points from the first series. */
export function legendIsCategoryDriven(
  chartType: string | undefined,
  seriesCount = 1,
  pieVaryColors = true,
): boolean {
  if (chartType === 'pie' || chartType === 'ofPie') return true;
  return chartType === 'doughnut' && (seriesCount <= 1 || pieVaryColors);
}

/** Whole-chart point legend used by the legacy single-group projection. */
export function chartVariesColorsByPoint(chart: {
  chartType?: string | null;
  radarStyle?: string | null;
  series: Array<{ bubbleXSourceIsString?: boolean | null }>;
  varyColors?: boolean | null;
}): boolean {
  if (
    chart.chartType === 'bubble'
    && chart.series.length === 1
    && chart.series[0]?.bubbleXSourceIsString === true
  ) return true;
  return !!chart.varyColors
    && chart.series.length === 1
    && typeof chart.chartType === 'string'
    && (/Bar/.test(chart.chartType)
      || chart.chartType === 'line' || chart.chartType === 'stackedLine'
      || chart.chartType === 'stackedLinePct' || chart.chartType === 'scatter'
      || (chart.chartType === 'radar' && chart.radarStyle !== 'filled'));
}

export interface LegendEntryRange {
  readonly firstIndex: number;
  readonly count: number;
  readonly pointDriven: boolean;
}

function visibleTrendlineCount(chart: ChartModel, sourceIndex: number): number {
  let count = 0;
  for (const trendline of chart.series[sourceIndex]?.trendLines ?? []) {
    if (trendline.lineHidden === true
      || (trendline.linePaintAuthored === true && trendline.lineColor == null)) continue;
    count++;
  }
  return count;
}

/**
 * Map source series/points to the global flattened `<c:legendEntry idx>`
 * domain. Combo legends insert every point of a varying group and each
 * paintable trendline before the next source series, so raw source or point
 * indexes are not interchangeable with legend indexes.
 */
export function legendEntryRanges(
  chart: ChartModel,
  includeTrendlines = false,
): readonly LegendEntryRange[] {
  const ranges: LegendEntryRange[] = chart.series.map(() => ({
    firstIndex: -1, count: 0, pointDriven: false,
  }));
  if (chart.series.length === 0) return ranges;
  const categoryDriven = legendIsCategoryDriven(
    chart.chartType, chart.series.length, chart.varyColors !== false,
  );
  const wholeChartPointDriven = chartVariesColorsByPoint(chart);
  if (categoryDriven || wholeChartPointDriven) {
    ranges[0] = {
      firstIndex: 0,
      count: chart.series[0]?.values.length ?? 0,
      pointDriven: true,
    };
    return ranges;
  }

  let offset = 0;
  for (let sourceIndex = 0; sourceIndex < chart.series.length; sourceIndex++) {
    const group = chartPlotGroupForSeries(chart, sourceIndex);
    const pointDriven = group?.kind !== 'pie' && group?.kind !== 'pie3D'
      && group?.kind !== 'doughnut' && group?.kind !== 'ofPie'
      && chartSeriesVariesByPoint(chart, sourceIndex);
    const count = pointDriven ? chart.series[sourceIndex]!.values.length : 1;
    ranges[sourceIndex] = { firstIndex: offset, count, pointDriven };
    offset += count + (includeTrendlines ? visibleTrendlineCount(chart, sourceIndex) : 0);
  }
  return ranges;
}

export function legendEntryGlobalIndex(
  ranges: readonly LegendEntryRange[],
  sourceIndex: number,
  pointIndex = 0,
): number | null {
  const range = ranges[sourceIndex];
  if (!range || range.firstIndex < 0 || range.count === 0) return null;
  const localIndex = range.pointDriven ? pointIndex : 0;
  return localIndex >= 0 && localIndex < range.count
    ? range.firstIndex + localIndex : null;
}

export function deletedLegendEntryIndices(chart: ChartModel): ReadonlySet<number> {
  const deleted = new Set<number>();
  for (const entry of chart.legendEntries ?? []) {
    if (entry.deleted === true) deleted.add(entry.idx);
  }
  return deleted;
}

export function legendEntryIsVisible(
  ranges: readonly LegendEntryRange[],
  deleted: ReadonlySet<number>,
  sourceIndex: number,
  pointIndex = 0,
): boolean {
  const globalIndex = legendEntryGlobalIndex(ranges, sourceIndex, pointIndex);
  return globalIndex != null && !deleted.has(globalIndex);
}

export function legendSeriesHasVisibleEntry(
  ranges: readonly LegendEntryRange[],
  deleted: ReadonlySet<number>,
  sourceIndex: number,
): boolean {
  const range = ranges[sourceIndex];
  if (!range || range.firstIndex < 0) return false;
  for (let localIndex = 0; localIndex < range.count; localIndex++) {
    if (!deleted.has(range.firstIndex + localIndex)) return true;
  }
  return false;
}
