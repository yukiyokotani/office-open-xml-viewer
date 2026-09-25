import type { Fill } from '../types/common.js';
import type { ChartDataPointOverride, ChartModel, ChartSeries } from '../types/chart.js';
import {
  chartDataPointStyleRole,
  rawLinkedChartStyleRole,
  chartSeriesSourceIndex,
  chartStyleDashChoice,
} from './effective-style.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectLineDecision,
  chartStyleDirectNoFillDecision,
  chartStyleFillDecision,
  chartStyleLineDecision,
} from './style-paint.js';

export interface ClassicDataPointLineStyle {
  paint: ChartModel['plotAreaLineFill'] | null | undefined;
  widthEmu: number | null | undefined;
  dash: string | null | undefined;
  customDash: ChartModel['plotAreaLineCustomDash'];
  cap: string | null | undefined;
  join: string | null | undefined;
}

/** Resolve the classic mark outline once for plot paint, resource preflight,
 * legend swatches and data-label keys. Direct dPt/series formatting owns paint
 * before the effective linked/numeric role; geometry cascades independently. */
export function classicDataPointLineStyle(
  chart: ChartModel,
  role: 'dataPoint' | 'dataPointLine',
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  styleIndex: number,
): ClassicDataPointLineStyle {
  const pointStyle = point?.chartexStyle;
  const seriesStyle = series.chartexStyle;
  const sourceSeriesIndex = chartSeriesSourceIndex(chart, series);
  const linkedStyle = chartDataPointStyleRole(
    chart, role, sourceSeriesIndex >= 0 ? sourceSeriesIndex : styleIndex,
  ) ?? (role === 'dataPoint' ? chart.chartexDataPointStyle : chart.chartexDataPointLineStyle);
  const rawLinked = rawLinkedChartStyleRole(chart, role);

  // Classic `<c:dPt>` / `<c:ser>` `spPr/a:ln/a:noFill` removes the chart
  // mark's own outline or connecting line. That semantic visibility decision
  // happens before a Chart Style role supplies the formatting of a line that
  // still exists: Excel keeps marker-only scatter series and borderless pie
  // points borderless even when the linked dataPointLine/dataPoint entry has
  // paint and no `allowNoLineOverride` modifier. By contrast, a no-line value
  // coming from a structured style layer is a replacement while resolving
  // CT_StyleEntry and remains modifier-gated below.
  let paint = point?.lineHidden === true
    ? null
    : chartStyleDirectLineDecision(pointStyle, rawLinked, point?.idx ?? styleIndex);
  if (paint === undefined) {
    if (point?.lineColor) paint = { fillType: 'solid', color: point.lineColor };
  }
  if (paint === undefined) {
    paint = series.lineHidden === true
      ? null
      : chartStyleDirectLineDecision(seriesStyle, rawLinked, styleIndex);
  }
  if (paint === undefined) {
    if (series.lineColor) paint = { fillType: 'solid', color: series.lineColor };
  }
  if (paint === undefined) paint = chartStyleLineDecision(linkedStyle, styleIndex);

  const dashChoice = chartStyleDashChoice(
    point?.lineDash != null ? { lineDash: point.lineDash, lineDashAuthored: true } : undefined,
    pointStyle,
    seriesStyle,
    linkedStyle,
  );
  return {
    paint,
    widthEmu: point?.lineWidthEmu
      ?? pointStyle?.lineWidthEmu
      ?? series.lineWidthEmu
      ?? seriesStyle?.lineWidthEmu
      ?? linkedStyle?.lineWidthEmu,
    dash: dashChoice?.lineDash,
    customDash: dashChoice?.lineCustomDash,
    cap: pointStyle?.lineCap ?? seriesStyle?.lineCap ?? linkedStyle?.lineCap,
    join: pointStyle?.lineJoin ?? seriesStyle?.lineJoin ?? linkedStyle?.lineJoin,
  };
}

/** Resolve one classic 2-D mark fill without collapsing semantic automatic
 * colour into direct formatting. This pure decision is shared by render and
 * picture-resource preflight so only an image that can actually reach a mark
 * is decoded. `undefined` delegates to the family automatic colour; `null` is
 * an authoritative no-fill or unresolved authored paint. */
export function classicDataPointFillDecision(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  styleIndex: number,
  pointIndex?: number,
): Fill | null | undefined {
  const sourceSeriesIndex = chartSeriesSourceIndex(chart, series);
  const linkedRole = chartDataPointStyleRole(
    chart, 'dataPoint', sourceSeriesIndex >= 0 ? sourceSeriesIndex : styleIndex,
  ) ?? chart.chartexDataPointStyle;
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint');
  const pointDecision = chartStyleDirectFillDecision(
    point?.chartexStyle, rawLinked, point?.idx ?? styleIndex,
  );
  if (pointDecision !== undefined) return pointDecision;
  if (point?.fillHidden === true) {
    const noFill = chartStyleDirectNoFillDecision(rawLinked);
    if (noFill !== undefined) return noFill;
  }
  if (point?.color === '00000000') return null;
  if (point?.color) return { fillType: 'solid', color: point.color };
  const indexedPointColor = pointIndex != null ? series.dataPointColors?.[pointIndex] : null;
  if (indexedPointColor === '00000000') return null;
  if (indexedPointColor) return { fillType: 'solid', color: indexedPointColor };

  const seriesDecision = chartStyleDirectFillDecision(
    series.chartexStyle, rawLinked, styleIndex,
  );
  if (seriesDecision !== undefined) return seriesDecision;
  if (series.color === '00000000') return null;
  if (series.fillPattern) return series.fillPattern;
  return chartStyleFillDecision(linkedRole, styleIndex);
}
