import type { Fill } from '../types/common';
import type { ChartDataPointOverride, ChartModel, ChartSeries } from '../types/chart';
import { dataLabelIsDeleted } from './data-label-style.js';
export { deletedLegendEntryIndices } from './legend-entry-plan.js';

/** True when a series has authored marker geometry/paint or point overrides. */
export function seriesHasMarkerDetail(series: ChartSeries): boolean {
  return series.markerSymbol != null
    || series.markerSize != null
    || series.markerFill != null
    || series.markerFillPaint !== undefined
    || series.markerFillPaintAuthored === true
    || series.markerStyle != null
    || series.markerLine != null
    || series.markerLinePaintAuthored === true
    || series.markerLineWidthEmu != null;
}

export function pointHasMarkerDetail(point: ChartDataPointOverride | undefined): boolean {
  return point != null && (
    point.markerSymbol != null
    || point.markerSize != null
    || point.markerFill != null
    || point.color != null
    || point.markerFillPaint !== undefined
    || point.markerFillPaintAuthored === true
    || point.chartexStyle != null
    || point.markerStyle != null
    || point.markerLine != null
    || point.markerLinePaintAuthored === true
    || point.markerLineWidthEmu != null
  );
}

/** Apply CT_BubbleChart.showNegBubbles before bubble paint/prefetch/work. */
export function visibleBubbleSize(
  chart: Pick<ChartModel, 'showNegativeBubbles'>,
  value: number | null | undefined,
): number | null {
  if (value == null || !Number.isFinite(value) || value === 0) return null;
  if (value < 0 && chart.showNegativeBubbles !== true) return null;
  return Math.abs(value);
}

/** Effective Office `bubble3D` paint for one bubble. Current Excel retains
 * point-level CT_DPt provenance but paints every point from the owning series'
 * value. The owning bubble-chart group remains the ECMA fallback when the
 * series omits the property. Complete omission means an ordinary flat bubble. */
export function bubblePointIsThreeD(
  series: ChartSeries,
  _point: ChartDataPointOverride | undefined,
): boolean {
  return series.bubble3D
    ?? series.bubble3DGroupDefault
    ?? false;
}

/** Open stroke-only symbols do not consume a fill recipe. */
export function markerSymbolConsumesFill(symbol: string | null | undefined): boolean {
  return symbol !== 'none' && symbol !== 'x' && symbol !== 'plus';
}

/** Whether one classic-series point can reach the marker painter. Shared with
 * image prefetch and paint-work preflight so hidden/null points do no I/O. */
export function classicMarkerPointIsPainted(
  chart: ChartModel,
  series: ChartSeries,
  family: string,
  index: number,
  scatterHasNumericX: boolean,
  groupSettings?: {
    chartType?: string;
    bubbleScale?: number | null;
    showNegativeBubbles?: boolean | null;
  },
): boolean {
  if (series.sourceHidden?.[index] === true) return false;
  const chartType = groupSettings?.chartType ?? chart.chartType;
  const value = series.values[index];
  let painted = value != null;
  if (!painted && (family === 'line' || family === 'stackedLine'
    || family === 'stackedLinePct')) {
    const renderedByLineFamily = chartType === 'line'
      || chartType === 'stackedLine' || chartType === 'stackedLinePct';
    painted = renderedByLineFamily
      && (chartType !== 'line' || chart.dispBlanksAs === 'zero');
  }
  if (!painted) return false;
  if (family === 'scatter' && scatterHasNumericX) {
    const category = (series.categories ?? chart.categories)[index];
    if (category == null || !Number.isFinite(Number.parseFloat(category))) return false;
  }
  if (family === 'scatter' && chartType === 'bubble') {
    const size = series.bubbleSizes?.[index];
    if (size == null || !Number.isFinite(size) || size === 0) return false;
    const showNegative = groupSettings?.showNegativeBubbles ?? chart.showNegativeBubbles;
    if (size < 0 && showNegative !== true) return false;
    return (groupSettings?.bubbleScale ?? chart.bubbleScale ?? 100) > 0;
  }
  return true;
}

/** Whether one classic-series point reaches the shared data-label painter.
 * This deliberately differs from marker visibility: area and stacked-line
 * families materialize authored null cells at zero for labels. Keep resource
 * preflight, picture warming, and paint on this single family-aware rule. */
export function classicDataLabelPointIsPainted(
  chart: ChartModel,
  series: ChartSeries,
  family: string,
  index: number,
  scatterHasNumericX: boolean,
  sourceSeriesIndex = 0,
): boolean {
  if (series.sourceHidden?.[index] === true) return false;
  if (family === 'surface') return false;
  if (family === 'ofPie' || (family === 'doughnut' && sourceSeriesIndex > 0)) return false;
  if (family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct') {
    return index < Math.max(
      chart.categories.length,
      series.categories?.length ?? 0,
      series.values.length,
    );
  }
  const value = series.values[index];
  if ((family === 'line' || family === 'stackedLine' || family === 'stackedLinePct')
    && value == null) {
    const renderedByLineFamily = chart.chartType === 'line'
      || chart.chartType === 'stackedLine' || chart.chartType === 'stackedLinePct';
    return renderedByLineFamily
      && (chart.chartType !== 'line' || chart.dispBlanksAs === 'zero');
  }
  if (value == null || !Number.isFinite(value)) return false;
  if (family === 'pie' || family === 'doughnut' || family === 'ofPie') {
    // Pie geometry uses absolute source values. Rich labels intentionally keep
    // zero-valued categories when the ring has at least one visible slice, so
    // prefetch/effect/paint-work must retain the same label domain.
    return series.values.some(candidate =>
      candidate != null && Number.isFinite(candidate) && Math.abs(candidate) > 0
    );
  }
  if (family === 'scatter') {
    if (!scatterHasNumericX) return true;
    const category = (series.categories ?? chart.categories)[index];
    return category != null && Number.isFinite(Number.parseFloat(category));
  }
  return true;
}

/** Count visible data-label legend keys that reuse the series marker paint. */
export function dataLabelLegendKeyCount(
  chart: ChartModel,
  series: ChartSeries,
  family: string,
  pointCount: number,
  scatterHasNumericX: boolean,
  _groupSettings?: Parameters<typeof classicMarkerPointIsPainted>[5],
  sourceSeriesIndex = 0,
): number {
  // Radar currently has no data-label consumer; do not prefetch or charge keys
  // that the family renderer cannot paint.
  if (family === 'radar') return 0;
  if (!series.seriesDataLabels && !(series.dataLabelOverrides?.length)) return 0;
  const overrides = new Map((series.dataLabelOverrides ?? []).map(point => [point.idx, point]));
  let count = 0;
  for (let index = 0; index < pointCount; index++) {
    const point = overrides.get(index);
    if (dataLabelIsDeleted(series.seriesDataLabels, point)) continue;
    if ((point?.showLegendKey ?? series.seriesDataLabels?.showLegendKey ?? false) !== true) continue;
    if (!classicDataLabelPointIsPainted(
      chart, series, family, index, scatterHasNumericX, sourceSeriesIndex,
    )) continue;
    count++;
  }
  return count;
}

/** Whether the current classic family routes CT_DTable through its renderer. */
export function chartDataTableFamilyIsPainted(chartType: string): boolean {
  return chartType === 'line' || chartType === 'stackedLine'
    || chartType === 'stackedLinePct' || chartType === 'area'
    || chartType === 'stackedArea' || chartType === 'stackedAreaPct'
    || chartType === 'stock' || chartType === 'clusteredBar'
    || chartType === 'clusteredBarH' || chartType === 'stackedBar'
    || chartType === 'stackedBarH' || chartType === 'stackedBarPct'
    || chartType === 'stackedBarHPct';
}

/** A point marker is more specific than the series marker visibility. */
export function hasVisiblePointMarkerOverride(series: ChartSeries): boolean {
  return series.dataPointOverrides?.some(point =>
    point.markerSymbol != null && point.markerSymbol !== 'none'
  ) === true;
}

/** Chart-group styles that suppress marker geometry before any series/point
 * formatting is considered. Shared by painters and availability preflight so
 * a structured fill is charged exactly when that family can consume it. */
export function markersSuppressedByChartStyle(
  family: string,
  chartType: string,
  scatterStyle: string | null | undefined,
  radarStyle: string | null | undefined,
): boolean {
  if (family === 'radar') return radarStyle === 'filled';
  return family === 'scatter' && chartType !== 'bubble'
    && (scatterStyle === 'lineNoMarker' || scatterStyle === 'smoothNoMarker');
}

/** Whether a series-driven legend entry owns a marker glyph. Point-level
 * formatting never participates because a legend entry represents the series. */
export function seriesLegendMarkerIsVisible(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  series: ChartSeries,
  radarStyle?: string | null,
): boolean {
  const family = series.seriesType ?? chartType;
  const lineFamily = family === 'line' || family === 'stackedLine'
    || family === 'stackedLinePct' || family === 'stock' || family === 'radar';
  if (!lineFamily && family !== 'scatter' && family !== 'bubble') return false;
  if (family === 'radar' && radarStyle === 'filled') return false;
  if (family === 'scatter'
    && (scatterStyle === 'lineNoMarker' || scatterStyle === 'smoothNoMarker')) return false;
  const symbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
  return symbol !== 'none' && series.showMarker !== false;
}

/** Resolve marker symbol visibility without letting a series-level `none`
 * suppress a more-specific `<c:dPt><c:marker><c:symbol>`. */
export function effectiveMarkerSymbol(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  fallback: string,
  seriesVisible: boolean,
): string {
  if (point?.markerSymbol != null) return point.markerSymbol;
  if (!seriesVisible || series.markerSymbol === 'none') return 'none';
  return series.markerSymbol ?? fallback;
}

/** Structured marker paint with direct point/series precedence. Authored but
 * unresolved/unsupported paint suppresses inherited paint rather than being
 * replaced with a less-specific linked or automatic fill. */
export function markerFillPaintFor(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
): Fill | null | undefined {
  if (point?.markerFillPaint !== undefined) return point.markerFillPaint;
  if (point?.markerFill != null || point?.color != null
    || point?.markerFillPaintAuthored === true) return undefined;
  if (series.dataPointColors?.[pointIndex] != null) return undefined;
  if (series.markerFillPaint !== undefined) return series.markerFillPaint;
  return undefined;
}

/** Legacy solid fallback paired with {@link markerFillPaintFor}. Transparent
 * is intentional when a direct paint exists but cannot currently be painted;
 * this keeps direct formatting authoritative without inventing a replacement. */
export function markerFillColorFor(
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  fallback: string,
): string {
  if (point?.markerFill != null) return point.markerFill;
  if (point?.color != null) return point.color;
  const pointColor = series.dataPointColors?.[pointIndex];
  if (pointColor != null) return pointColor;
  if (point?.markerFillPaintAuthored === true) return '00000000';
  if (series.markerFill != null) return series.markerFill;
  if (series.markerFillPaintAuthored === true) return '00000000';
  return fallback;
}

/** Series-level marker paint for a legend key. A legend must not inherit the
 * first data point's `dPt`/varyColors formatting. */
export function seriesMarkerFillPaint(series: ChartSeries): Fill | null | undefined {
  return series.markerFillPaint;
}

/** Series-level solid fallback paired with {@link seriesMarkerFillPaint}. */
export function seriesMarkerFillColor(series: ChartSeries, fallback: string): string {
  if (series.markerFill != null) return series.markerFill;
  if (series.markerFillPaintAuthored === true) return '00000000';
  return fallback;
}

/** Fill-component work for one effective marker paint. */
export function markerPaintComponents(fill: Fill | null | undefined): number {
  if (fill?.fillType === 'gradient') return fill.stops.length;
  return fill == null ? 0 : 1;
}
