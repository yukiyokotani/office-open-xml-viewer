import type { ChartModel, ChartPlotGroup } from '../types/chart.js';

export type ClassicPlotDispatch =
  | 'legacy' | 'bar-combo' | 'line-groups' | 'area-groups' | 'scatter-bubble'
  | 'stock-line' | 'unsupported';

const THREE_D_GROUPS = new Set(['area3D', 'line3D', 'pie3D', 'bar3D', 'surface3D']);

function legacyTypeMatchesGroup(chartType: string, kind: ChartPlotGroup['kind']): boolean {
  if (kind === 'bar' || kind === 'bar3D') return chartType.includes('Bar');
  if (kind === 'line' || kind === 'line3D') return chartType.includes('Line');
  if (kind === 'area' || kind === 'area3D') return chartType.includes('Area');
  if (kind === 'pie' || kind === 'pie3D') return chartType === 'pie';
  if (kind === 'surface' || kind === 'surface3D') return chartType === 'surface'
    || chartType === 'surface3D';
  return chartType === kind;
}

/** Validate the compact contiguous-range contract before group planning.
 * Axis resolution is dispatch-specific: a small number of Office-authored
 * compatibility layouts deliberately retain an ambiguous same-side numeric
 * axis while their legacy renderer still has an exact primary/secondary route. */
function chartPlotGroupsAreValid(chart: ChartModel): boolean {
  if (chart.plotGroups == null) return true;
  let next = 0;
  for (const group of chart.plotGroups) {
    if (!Number.isSafeInteger(group.seriesStart) || group.seriesStart !== next
      || !Number.isSafeInteger(group.seriesCount) || group.seriesCount < 0
      || group.seriesCount > chart.series.length - next) return false;
    next += group.seriesCount;
  }
  return next === chart.series.length;
}

/** Build the one linear series-to-group lookup shared by paint preflights. */
export function indexChartPlotGroups(
  chart: ChartModel,
): Array<ChartPlotGroup | undefined> {
  const result = new Array<ChartPlotGroup | undefined>(chart.series.length);
  for (const group of chart.plotGroups ?? []) {
    const end = Math.min(chart.series.length, group.seriesStart + group.seriesCount);
    for (let index = group.seriesStart; index < end; index++) result[index] = group;
  }
  return result;
}

/** Canonical family token used by existing marker visibility rules while an
 * exact plot-group kind remains available separately for dispatch. */
export function markerChartTypeForPlotGroup(
  chartType: string,
  group: ChartPlotGroup | undefined,
): string {
  if (!group) return chartType;
  if (group.kind === 'bubble') return 'bubble';
  if (group.kind === 'line') {
    return group.grouping === 'percentStacked'
      ? 'stackedLinePct' : group.grouping === 'stacked' ? 'stackedLine' : 'line';
  }
  if (group.kind === 'area') {
    return group.grouping === 'percentStacked'
      ? 'stackedAreaPct' : group.grouping === 'stacked' ? 'stackedArea' : 'area';
  }
  return group.kind;
}

/** Determine only the Office-observed/spec-safe mixed-family routing boundary. */
export function classicPlotDispatch(chart: ChartModel): ClassicPlotDispatch {
  if (chart.plotGroups == null || chart.plotGroups.length <= 1) return 'legacy';
  if (!chartPlotGroupsAreValid(chart)) return 'unsupported';
  const groups = chart.plotGroups.filter(group => group.seriesCount > 0);
  if (groups.length === 0) return 'legacy';
  if (groups.length === 1) {
    return legacyTypeMatchesGroup(chart.chartType, groups[0].kind)
      ? 'legacy' : 'unsupported';
  }
  if (groups.some(group => group.categoryAxis === 'none' || group.valueAxis === 'none')) {
    return 'unsupported';
  }
  if (groups.some(group =>
    (group.categoryAxis === 'secondary' && chart.secondaryCatAxis == null)
    || (group.valueAxis === 'secondary' && chart.secondaryValAxis == null)
  )) return 'unsupported';
  const kinds = new Set(groups.map(group => group.kind));
  if (groups.some(group => THREE_D_GROUPS.has(group.kind))) return 'unsupported';
  if (groups.length === 2 && groups[0].kind === 'bar' && groups[1].kind === 'scatter') {
    const [bar, scatter] = groups;
    const barSeries = chart.series.slice(bar.seriesStart, bar.seriesStart + bar.seriesCount);
    const scatterSeries = chart.series.slice(
      scatter.seriesStart, scatter.seriesStart + scatter.seriesCount,
    );
    // Office-authored range/dumbbell charts can contain two bottom numeric
    // axes: the horizontal bar value axis and the scatter X axis. Their axPos
    // values alone are ambiguous, but the exact group order, cross-axis pairs,
    // secondary series flags, and independent secondary X/Y models retain the
    // established legacy route without guessing either coordinate.
    if (bar.categoryAxis === 'primary'
      && (bar.valueAxis === 'primary' || bar.valueAxis === 'unresolved')
      && (scatter.categoryAxis === 'secondary' || scatter.categoryAxis === 'unresolved')
      && scatter.valueAxis === 'secondary'
      && chart.secondaryCatAxis != null && chart.secondaryValAxis != null
      && barSeries.every(series => series.useSecondaryAxis !== true)
      && scatterSeries.every(series => series.useSecondaryAxis === true)) {
      return 'bar-combo';
    }
    return 'unsupported';
  }
  if (groups.some(group => group.categoryAxis === 'unresolved'
    || group.valueAxis === 'unresolved' || group.seriesAxis === 'unresolved')) {
    return 'unsupported';
  }
  if (kinds.size === 1) {
    const kind = groups[0].kind;
    if (kind === 'line' || kind === 'area') {
      // The shared line/area painters currently expose one category scale.
      // A group may still bind that primary category scale to either primary
      // or secondary value space (the ordinary column + secondary-line form),
      // but a distinct top category axis needs its own measured layout pass.
      // Keep that unobserved boundary explicit instead of silently painting it
      // through the bottom axis.
      if (groups.some(group => group.categoryAxis !== 'primary')) return 'unsupported';
      return kind === 'line' ? 'line-groups' : 'area-groups';
    }
    if (kind === 'scatter' || kind === 'bubble') return 'scatter-bubble';
    if (kind === 'bar') {
      const fallbackDirection = chart.chartType.endsWith('H') ? 'bar' : 'col';
      const directions = new Set(groups.map(group => group.barDirection ?? fallbackDirection));
      if (directions.size > 1 && (groups.length !== 2 || groups.some(group =>
        group.categoryAxis !== 'primary' || group.valueAxis !== 'primary'
      ))) return 'unsupported';
      return 'bar-combo';
    }
    return 'unsupported';
  }
  if (groups.length === 2 && kinds.has('bar') && kinds.has('area')) {
    const overlay = groups.find(group => group.kind === 'line' || group.kind === 'area');
    if (overlay?.categoryAxis !== 'primary') return 'unsupported';
    return 'bar-combo';
  }
  if (groups.length === 2 && groups[0].kind === 'area' && groups[1].kind === 'line'
    && groups.every(group => group.categoryAxis === 'primary'
      && group.valueAxis === 'primary')) {
    return 'area-groups';
  }
  if (kinds.size === 2 && kinds.has('bar') && kinds.has('line')) {
    const barGroups = groups.filter(group => group.kind === 'bar');
    const lineGroups = groups.filter(group => group.kind === 'line');
    const fallbackDirection = chart.chartType.endsWith('H') ? 'bar' : 'col';
    const directions = new Set(
      barGroups.map(group => group.barDirection ?? fallbackDirection),
    );
    // The retained Office boundary covers the ordinary bar+line pair and the
    // common two-column-group + one-line-group form. Keep larger/unrelated
    // scenes explicit until their z-order and axis ownership are observed.
    // A secondary category axis that is deleted has no layout of its own: Office
    // draws its group against the primary categories (the common "columns on the
    // secondary axis, line on the primary" combo).
    const onPrimaryCategories = (group: ChartPlotGroup) => group.categoryAxis === 'primary'
      || (group.categoryAxis === 'secondary' && chart.secondaryCatAxis?.hidden === true);
    if (barGroups.length <= 2 && lineGroups.length === 1 && groups.length <= 3
      && directions.size === 1
      && groups.every(onPrimaryCategories)) {
      return 'bar-combo';
    }
    return 'unsupported';
  }
  if ([...kinds].every(kind => kind === 'scatter' || kind === 'bubble')) {
    if (groups.length > 2) return 'unsupported';
    return 'scatter-bubble';
  }
  if (groups.length === 2 && groups[0].kind === 'stock' && groups[1].kind === 'line'
    && (groups[0].seriesCount === 3 || groups[0].seriesCount === 4)
    && (groups[1].grouping == null || groups[1].grouping === 'standard')
    && groups.every(group => group.categoryAxis === 'primary'
      && group.valueAxis === 'primary')) {
    return 'stock-line';
  }
  return 'unsupported';
}
