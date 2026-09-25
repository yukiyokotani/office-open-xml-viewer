import type {
  ChartDataPointOverride,
  ChartModel,
  ChartSeries,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import {
  chartDataPointStyleRole,
  chartSeriesVariesByPoint,
  rawLinkedChartStyleRole,
} from './effective-style.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectNoFillDecision,
  chartStyleFillCascade,
} from './style-paint.js';

export function threeDDatumStyleIndex(
  chart: ChartModel,
  series: ChartSeries,
  pointIndex: number,
  seriesIndex: number,
): number {
  return chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex
    : series.chartexFormatIdx ?? seriesIndex;
}

/** Resolve the authored fill selected for one classic 3-D datum. This pure
 * decision is shared by the optional mesh painter and the base-package image
 * preflight so a relationship-backed recipe is decoded iff paint can reach it.
 * `undefined` delegates to the semantic automatic color; `null` is an
 * authoritative no-fill or unresolved authored paint. */
export function chartThreeDDatumFillDecision(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  seriesIndex: number,
): Fill | null | undefined {
  const styleIndex = threeDDatumStyleIndex(chart, series, pointIndex, seriesIndex);
  const linked = chartDataPointStyleRole(chart, 'dataPoint3D', seriesIndex);
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint3D');
  const pointDecision = chartStyleDirectFillDecision(
    point?.chartexStyle, rawLinked, pointIndex,
  );
  if (pointDecision !== undefined) return pointDecision;
  if (point?.fillHidden === true) {
    const noFill = chartStyleDirectNoFillDecision(rawLinked);
    if (noFill !== undefined) return noFill;
  }
  if (point?.color === '00000000') return null;
  if (point?.color) return { fillType: 'solid', color: point.color };

  const seriesDecision = chartStyleDirectFillDecision(
    series.chartexStyle, rawLinked, styleIndex,
  );
  if (seriesDecision !== undefined) return seriesDecision;
  return chartStyleFillCascade(
    linked,
    rawLinked,
    styleIndex,
    point?.chartexStyle,
    series.chartexStyle,
  );
}
