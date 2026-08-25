import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartModel,
  ChartSeries,
} from '../types/chart';

interface VisibilityPlan {
  keep: number[];
  remap: Int32Array;
}

function makeVisibilityPlan(hidden: readonly boolean[], pointCount: number): VisibilityPlan | null {
  if (!hidden.some(Boolean)) return null;
  const count = Math.max(pointCount, hidden.length);
  const keep: number[] = [];
  const remap = new Int32Array(count);
  remap.fill(-1);
  for (let index = 0; index < count; index++) {
    if (hidden[index] === true) continue;
    remap[index] = keep.length;
    keep.push(index);
  }
  return { keep, remap };
}

function select<T>(values: readonly T[] | null | undefined, plan: VisibilityPlan): T[] | null | undefined {
  if (values == null) return values;
  const selected: T[] = [];
  for (const index of plan.keep) {
    if (index < values.length) selected.push(values[index]);
  }
  return selected;
}

function reindexOverrides<T extends ChartDataPointOverride | ChartDataLabelOverride>(
  overrides: readonly T[] | null | undefined,
  plan: VisibilityPlan,
): T[] | null | undefined {
  if (overrides == null) return overrides;
  const result: T[] = [];
  for (const override of overrides) {
    if (override.idx >= plan.remap.length) continue;
    const mapped = plan.remap[override.idx];
    if (mapped >= 0) result.push({ ...override, idx: mapped });
  }
  return result;
}

function pointCount(series: ChartSeries): number {
  return Math.max(
    series.values.length,
    series.categories?.length ?? 0,
    series.sourceHidden?.length ?? 0,
    series.dataPointColors?.length ?? 0,
    series.dataLabelColors?.length ?? 0,
    series.catFormatCodes?.length ?? 0,
    series.bubbleSizes?.length ?? 0,
    ...((series.errBars ?? []).flatMap(errorBars => [
      errorBars.plus.length,
      errorBars.minus.length,
    ])),
  );
}

function applyPlanToSeries(series: ChartSeries, plan: VisibilityPlan): ChartSeries {
  return {
    ...series,
    values: select(series.values, plan) ?? [],
    categories: select(series.categories, plan),
    sourceHidden: select(series.sourceHidden, plan),
    dataPointColors: select(series.dataPointColors, plan),
    dataLabelColors: select(series.dataLabelColors, plan),
    catFormatCodes: select(series.catFormatCodes, plan),
    bubbleSizes: select(series.bubbleSizes, plan),
    dataPointOverrides: reindexOverrides(series.dataPointOverrides, plan),
    dataLabelOverrides: reindexOverrides(series.dataLabelOverrides, plan),
    errBars: series.errBars?.map(errorBars => ({
      ...errorBars,
      plus: select(errorBars.plus, plan) ?? [],
      minus: select(errorBars.minus, plan) ?? [],
    })),
  };
}

function suppressHiddenSeriesPoints(series: ChartSeries): ChartSeries {
  const hidden = series.sourceHidden;
  if (!hidden?.some(Boolean)) return series;
  const suppress = <T>(values: readonly T[] | null | undefined, missing: T) => {
    if (values == null) return values;
    return values.map((value, index) => hidden[index] === true ? missing : value);
  };
  return {
    ...series,
    values: suppress(series.values, null) ?? [],
    bubbleSizes: suppress(series.bubbleSizes, null),
    errBars: series.errBars?.map(errorBars => ({
      ...errorBars,
      plus: suppress(errorBars.plus, null) ?? [],
      minus: suppress(errorBars.minus, null) ?? [],
    })),
  };
}

function isScatterLike(chart: ChartModel): boolean {
  return chart.chartType === 'scatter' || chart.chartType === 'bubble';
}

function allSourcePointsHidden(series: ChartSeries): boolean {
  const hidden = series.sourceHidden;
  return hidden != null && hidden.length > 0 && hidden.every(Boolean);
}

export const EXCEL_AUTOMATIC_SCATTER_MARKERS = [
  'diamond', 'square', 'triangle', 'x', 'star', 'circle',
] as const;

// Provenance for the Office compatibility style synthesized below. Keep this
// out of the public ChartSeries model: the fact exists only for the filtered
// object identity produced in this realm, and must never be reconstructed from
// coincidentally equal paint arrays.
const filteredScatterAutomaticPointStyles = new WeakSet<ChartSeries>();

export function hasFilteredScatterAutomaticPointStyle(series: ChartSeries): boolean {
  return filteredScatterAutomaticPointStyles.has(series);
}

function applyExcelFilteredScatterAutomaticStyle(
  chart: ChartModel,
  series: ChartSeries,
): ChartSeries {
  const accents = chart.themeAccentColors;
  const hasDirectPointStyle = series.dataPointColors?.some(color => color != null) === true
    || (series.dataPointOverrides?.length ?? 0) > 0;
  const hasDirectMarkerStyle = series.markerSymbol != null
    || series.markerFill != null
    || series.markerFillPaintAuthored === true
    || series.markerLine != null;
  if (
    chart.scatterStyle !== 'marker'
    || !accents?.length
    || hasDirectPointStyle
    || hasDirectMarkerStyle
    || series.lineHidden === true
  ) return series;

  // Current Excel restyles the compacted visible points of an otherwise
  // unformatted `scatterStyle="marker"` series as consecutive automatic
  // theme entries. Keep this compatibility behavior after the normative
  // plotVisOnly filtering step so hidden source slots consume no palette entry.
  const styled: ChartSeries = {
    ...series,
    dataPointColors: series.values.map((_, index) => accents[index % accents.length]),
    dataPointOverrides: series.values.map((_, index) => ({
      idx: index,
      markerSymbol: EXCEL_AUTOMATIC_SCATTER_MARKERS[
        index % EXCEL_AUTOMATIC_SCATTER_MARKERS.length
      ],
    })),
  };
  filteredScatterAutomaticPointStyles.add(styled);
  return styled;
}

/**
 * Apply ECMA-376 §21.2.2.146 once, before any chart-family extent, stacking,
 * trendline, error-bar, or label work. Hosts only contribute aligned source
 * visibility facts; this shared step owns their rendering semantics.
 */
export function applyPlotVisibleOnly(chart: ChartModel): ChartModel {
  if (chart.plotVisibleOnly !== true) return chart;

  if (isScatterLike(chart)) {
    return {
      ...chart,
      series: chart.series.flatMap(series => {
        const effective = series.categories == null
          ? { ...series, categories: chart.categories }
          : series;
        const hidden = effective.sourceHidden;
        const plan = hidden && makeVisibilityPlan(hidden, pointCount(effective));
        if (!plan) return [effective];
        if (plan.keep.length === 0) return [];
        const filtered = applyPlanToSeries(effective, plan);
        return chart.chartType === 'scatter'
          ? [applyExcelFilteredScatterAutomaticStyle(chart, filtered)]
          : [filtered];
      }),
    };
  }

  const globalPointCount = Math.max(
    chart.categories.length,
    ...(chart.categoryLevels?.map(level => level.length) ?? []),
    ...chart.series.map(pointCount),
  );
  const categoryPlan = chart.categorySourceHidden
    ? makeVisibilityPlan(chart.categorySourceHidden, globalPointCount)
    : null;
  const series = chart.series.flatMap(source => {
    const categoryFiltered = categoryPlan ? applyPlanToSeries(source, categoryPlan) : source;
    if (categoryPlan?.keep.length === 0 || allSourcePointsHidden(categoryFiltered)) return [];
    return [suppressHiddenSeriesPoints(categoryFiltered)];
  });
  if (!categoryPlan) return { ...chart, series };

  const subtotalIndices = chart.subtotalIndices.flatMap(index => {
    if (index >= categoryPlan.remap.length) return [];
    const mapped = categoryPlan.remap[index];
    return mapped >= 0 ? [mapped] : [];
  });
  return {
    ...chart,
    categories: select(chart.categories, categoryPlan) ?? [],
    categoryLevels: chart.categoryLevels?.map(level => select(level, categoryPlan) ?? []),
    categorySourceHidden: select(chart.categorySourceHidden, categoryPlan),
    subtotalIndices,
    series,
  };
}
