// Classic chart legend helpers.
import type { Fill } from '../../types/common';
import type { ChartDataPointOverride, ChartExElementStyle, ChartLegendEntryOverride, ChartModel, ChartSeries } from '../../types/chart';
import { chartDataPointStyleRole, chartPlotGroupForSeries, chartSeriesSourceIndex, chartSeriesVariesByPoint, chartStyleDashChoice, rawLinkedChartStyleRole } from '../effective-style.js';
import { markerChartTypeForPlotGroup } from '../plot-groups.js';
import { bubblePointIsThreeD, effectiveMarkerSymbol, markerFillColorFor, markerFillPaintFor, seriesLegendMarkerIsVisible, seriesMarkerFillColor, seriesMarkerFillPaint } from '../marker-style.js';
import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import { chartStyleDirectFillDecision, chartStyleDirectLineDecision, chartStyleFillDecision, chartStyleLineDecision } from '../style-paint.js';
import { axisLineWidthPx } from '../axis-style.js';
import { resolveFill } from '../../shape/paint.js';
import { drawingmlLineDashArray } from '../../draw/dash.js';
import { chartVariesColorsByPoint, legendEntryGlobalIndex, legendEntryRanges, legendIsCategoryDriven } from '../legend-entry-plan.js';
import { classicDataPointFillDecision, classicDataPointLineStyle } from '../classic-data-point-style.js';
import { chartLegendReserve, chartTextFontSizePx, packLegendRows, resolveManualLayoutRect } from '../layout.js';
import type { ChartLegendReserve } from '../layout.js';
import { elideToWidth } from '../text-elide.js';
import { paintLegendFrame } from '../legend-frame.js';
import { CHART_PALETTE, chartColor, chartExSeriesFormatIndex, piePointStyleIndex, pieSliceColor } from './palette.js';
import { bubblePointFill, bubblePointLine } from './bubble-paint.js';
import { drawMarker } from './markers.js';
import { paintClassicDataPointRect } from './chartex-style.js';
import { resolveThemeFontRef } from './fonts.js';
import { dashPatternForPreset } from './geometry.js';
import { wrapMeasuredText } from './axis.js';
import { legendSeriesWithTrendlines } from './trendline.js';


/** Line-shaped legend swatch styles match Excel's actual chart-type
 *  conventions: bar/column/area/pie use a filled rectangle ("swatch");
 *  line/radar/scatter use a horizontal line segment (the same stroke
 *  weight the series uses). Without this, line-chart legends rendered as
 *  filled squares, which read as a different chart-type marker.
 */
export type LegendSwatchStyle = 'fill' | 'line' | 'none';


export function legendSwatchStyle(chartType: string | undefined): LegendSwatchStyle {
  if (!chartType) return 'fill';
  if (
    chartType === 'line' || chartType === 'stackedLine' || chartType === 'stackedLinePct' ||
    chartType === 'radar' || chartType === 'scatter' || chartType === 'stock'
  ) {
    return 'line';
  }
  return 'fill';
}


/** A resolved marker legend key: the glyph a scatter series draws for its
 *  points, used as the legend swatch when the series has no connecting line
 *  (§21.2.2.32). `fill`/`line` are hex without `#` (chartColor / markerFill). */
export interface LegendMarker {
  symbol: string;
  fill: string;
  fillPaint?: Fill | null;
  line: string | null;
  lineWidthEmu: number | null;
  linePaint?: ChartModel['plotAreaLineFill'] | null;
  lineDash?: string | null;
  lineCustomDash?: ChartModel['plotAreaLineCustomDash'];
  lineCap?: string | null;
  lineJoin?: string | null;
  bubble3D?: boolean;
  directEffect?: ChartExElementStyle;
  fallbackEffect?: ChartExElementStyle;
  directEffectIndex?: number;
  fallbackEffectIndex?: number;
  /** True when the plotted series draws both a connecting line and markers. */
  withLine: boolean;
}


/** Whether a scatter/bubble series draws a connecting line in the plot, so its
 *  legend key should be a line swatch rather than a marker glyph. Mirrors the
 *  plot gate in {@link renderScatterChart}: the group `<c:scatterStyle>` decides
 *  whether points are connected, and a series-level `<a:noFill/>` line override
 *  (§21.2.2.198, `lineHidden`) suppresses the connecting line even when the group
 *  style is `line`/`lineMarker`. Bubble charts are always markers-only. */
export function scatterSeriesDrawsLine(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  series: ChartSeries,
): boolean {
  if (chartType !== 'scatter') return false;
  const style = scatterStyle ?? 'marker';
  const styleDrawsLine =
    style === 'marker' || style === 'line' || style === 'lineMarker' || style === 'lineNoMarker' ||
    style === 'smooth' || style === 'smoothMarker' || style === 'smoothNoMarker';
  return styleDrawsLine && series.lineHidden !== true;
}


/** The marker legend key for a scatter series that draws no connecting line
 *  (markers-only, whether by group style or a series `<a:noFill/>` override).
 *  Excel renders such a series' legend key as its point marker, not a line
 *  swatch. Returns null when a marker key does not apply (non-scatter, or a
 *  scatter series that does draw a line). Colors/symbol resolve exactly like the
 *  plotted markers in {@link renderScatterChart}. */
export function legendMarkerFor(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  radarStyle: string | null | undefined,
  series: ChartSeries[],
  entryIndex: number,
  chart?: ChartModel,
): LegendMarker | null {
  const s = series[entryIndex];
  if (!s) return null;
  const sourceSeriesIndex = chart
    ? Math.max(0, chartSeriesSourceIndex(chart, s))
    : entryIndex;
  const group = chart ? chartPlotGroupForSeries(chart, sourceSeriesIndex) : undefined;
  const effectiveChartType = chart && group
    ? markerChartTypeForPlotGroup(chart.chartType, group)
    : s.seriesType ?? chartType;
  const effectiveScatterStyle = group?.scatterStyle ?? scatterStyle;
  const effectiveRadarStyle = group?.radarStyle ?? radarStyle;
  const family = group?.kind === 'bubble' ? 'bubble' : effectiveChartType;
  const isStock = family === 'stock';
  const isLineFamily = family === 'line' || family === 'stackedLine' ||
    family === 'stackedLinePct' || family === 'radar' || isStock;
  const isBubble = family === 'bubble';
  const isScatter = family === 'scatter' || isBubble;
  if (!isLineFamily && !isScatter) return null;
  const visibilitySeries = group
    ? { ...s, seriesType: effectiveChartType }
    : s;
  if (!seriesLegendMarkerIsVisible(
    effectiveChartType,
    effectiveScatterStyle,
    visibilitySeries,
    effectiveRadarStyle,
  )) return null;
  const symbol = s.markerSymbol
    ?? s.automaticMarkerSymbol
    ?? (isStock ? 'none' : 'circle');
  const base = chartColor(sourceSeriesIndex, s); // '#RRGGBB'
  const fill = seriesMarkerFillColor(s, base.replace(/^#/, ''));
  const withLine = isBubble ? false : isScatter
    ? scatterSeriesDrawsLine('scatter', effectiveScatterStyle, s)
    : s.lineHidden !== true;
  if (isBubble && chart) {
    const styleIndex = chartExSeriesFormatIndex(s, sourceSeriesIndex);
    const bubbleFill = bubblePointFill(chart, s, undefined, 0, styleIndex, base);
    const bubbleLine = bubblePointLine(chart, s, undefined, 0, styleIndex);
    return {
      symbol: 'circle',
      fill: bubbleFill.color,
      fillPaint: bubbleFill.paint,
      line: bubbleLine.color,
      lineWidthEmu: bubbleLine.widthEmu ?? null,
      linePaint: bubbleLine.paint,
      lineDash: bubbleLine.dash,
      lineCustomDash: bubbleLine.customDash,
      lineCap: bubbleLine.cap,
      lineJoin: bubbleLine.join,
      bubble3D: bubblePointIsThreeD(s, undefined),
      withLine: false,
      directEffect: chartStyleEffectOwner(s.chartexStyle),
      fallbackEffect: chartDataPointStyleRole(
        chart, bubblePointIsThreeD(s, undefined) ? 'dataPoint3D' : 'dataPoint',
        sourceSeriesIndex,
      ),
      directEffectIndex: styleIndex,
      fallbackEffectIndex: chartSeriesVariesByPoint(chart, sourceSeriesIndex)
        ? 0 : styleIndex,
    };
  }
  if (chart) {
    return classicPointLegendMarker(
      chart, s, undefined, 0, sourceSeriesIndex, family,
      effectiveScatterStyle, effectiveRadarStyle,
    );
  }
  return {
    symbol,
    fill,
    fillPaint: seriesMarkerFillPaint(s),
    line: s.markerLine ?? null,
    lineWidthEmu: s.markerLineWidthEmu ?? null,
    withLine,
  };
}


export function bubblePointLegendMarker(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
): LegendMarker {
  const sourceSeriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const fallback = chartColor(sourceSeriesIndex, series);
  const styleIndex = chartExSeriesFormatIndex(series, sourceSeriesIndex);
  const fill = bubblePointFill(chart, series, point, pointIndex, styleIndex, fallback);
  const line = bubblePointLine(chart, series, point, pointIndex, styleIndex);
  const pointEffect = chartStyleEffectOwner(point?.chartexStyle);
  const seriesEffect = chartStyleEffectOwner(series.chartexStyle);
  const bubble3D = bubblePointIsThreeD(series, point);
  return {
    symbol: 'circle',
    fill: fill.color,
    fillPaint: fill.paint,
    line: line.color,
    lineWidthEmu: line.widthEmu ?? null,
    linePaint: line.paint,
    lineDash: line.dash,
    lineCustomDash: line.customDash,
    lineCap: line.cap,
    lineJoin: line.join,
    bubble3D,
    withLine: false,
    directEffect: pointEffect ?? seriesEffect,
    fallbackEffect: chartDataPointStyleRole(
      chart, bubble3D ? 'dataPoint3D' : 'dataPoint', sourceSeriesIndex,
    ),
    directEffectIndex: pointEffect ? pointIndex : styleIndex,
    fallbackEffectIndex: chartSeriesVariesByPoint(chart, sourceSeriesIndex)
      ? pointIndex : styleIndex,
  };
}


/** Resolve a point-driven line/scatter/radar legend marker through the same
 * direct-marker then numeric/linked `dataPointMarker` cascade as the plotted
 * glyph. The legend key is geometry-independent, but its paint, outline and
 * effect ownership are not. */
export function classicPointLegendMarker(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  sourceSeriesIndex: number,
  family: string | undefined,
  scatterStyle: string | null | undefined,
  radarStyle: string | null | undefined,
): LegendMarker | null {
  const lineFamily = family === 'line' || family === 'stackedLine'
    || family === 'stackedLinePct' || family === 'radar' || family === 'stock';
  const scatterFamily = family === 'scatter';
  if (!lineFamily && !scatterFamily) return null;
  const seriesVisible = seriesLegendMarkerIsVisible(
    family, scatterStyle, series, radarStyle,
  );
  const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
  if (symbol === 'none') return null;

  const fallback = chartColor(sourceSeriesIndex, series).replace(/^#/, '');
  const styleIndex = series.chartexFormatIdx ?? sourceSeriesIndex;
  const linked = chartDataPointStyleRole(chart, 'dataPointMarker', sourceSeriesIndex);
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPointMarker');
  const linkedIndex = chartSeriesVariesByPoint(chart, sourceSeriesIndex)
    ? pointIndex : styleIndex;
  const pointFill = chartStyleDirectFillDecision(
    point?.markerStyle, rawLinked, pointIndex,
  );
  const seriesFill = chartStyleDirectFillDecision(
    series.markerStyle, rawLinked, styleIndex,
  );
  const pointOwnsFill = pointFill !== undefined || point?.markerFill != null
    || point?.markerFillPaint !== undefined
    || point?.markerFillPaintAuthored === true && point.markerStyle?.fillHidden !== true;
  const seriesOwnsFill = seriesFill !== undefined || series.markerFill != null
    || series.markerFillPaint !== undefined
    || series.markerFillPaintAuthored === true && series.markerStyle?.fillHidden !== true;
  let fill = point
    ? markerFillColorFor(series, point, pointIndex, fallback)
    : seriesMarkerFillColor(series, fallback);
  let fillPaint = point
    ? markerFillPaintFor(series, point, pointIndex)
    : seriesMarkerFillPaint(series);
  if (pointFill === null || (pointFill === undefined && seriesFill === null)) {
    fill = '00000000';
    fillPaint = null;
  } else if (!pointOwnsFill && !seriesOwnsFill) {
    const linkedFill = chartStyleFillDecision(linked, linkedIndex);
    if (linkedFill === null) {
      fill = '00000000';
      fillPaint = null;
    } else if (linkedFill?.fillType === 'solid') {
      fill = linkedFill.color;
      fillPaint = undefined;
    } else if (linkedFill !== undefined) fillPaint = linkedFill;
  }

  const pointLine = chartStyleDirectLineDecision(
    point?.markerStyle, rawLinked, pointIndex,
  );
  const seriesLine = chartStyleDirectLineDecision(
    series.markerStyle, rawLinked, styleIndex,
  );
  let line = point?.markerLine ?? series.markerLine ?? null;
  let linePaint: ChartModel['plotAreaLineFill'] | null | undefined;
  if (pointLine !== undefined) {
    if (pointLine?.fillType === 'solid') line = pointLine.color;
    else linePaint = pointLine;
  } else if (point?.markerLinePaintAuthored === true
    && point.markerStyle?.lineHidden !== true && point.markerLine == null) {
    linePaint = null;
  } else if (seriesLine !== undefined) {
    if (seriesLine?.fillType === 'solid') line = seriesLine.color;
    else linePaint = seriesLine;
  } else if (series.markerLinePaintAuthored === true
    && series.markerStyle?.lineHidden !== true && series.markerLine == null) {
    linePaint = null;
  } else if (point?.markerLine == null && series.markerLine == null) {
    const linkedLine = chartStyleLineDecision(linked, linkedIndex);
    if (linkedLine?.fillType === 'solid') line = linkedLine.color;
    else {
      linePaint = linkedLine;
      if (linkedLine === null) line = null;
    }
  }
  const linkedGeometry = linked;
  const dash = chartStyleDashChoice(point?.markerStyle, series.markerStyle, linkedGeometry);
  const pointEffect = chartStyleEffectOwner(point?.markerStyle, point?.chartexStyle);
  const seriesEffect = chartStyleEffectOwner(series.markerStyle);
  return {
    symbol,
    fill,
    fillPaint,
    line,
    linePaint,
    lineWidthEmu: point?.markerLineWidthEmu ?? point?.markerStyle?.lineWidthEmu
      ?? series.markerLineWidthEmu ?? series.markerStyle?.lineWidthEmu
      ?? linkedGeometry?.lineWidthEmu ?? null,
    lineDash: dash?.lineDash,
    lineCustomDash: dash?.lineCustomDash,
    lineCap: point?.markerStyle?.lineCap ?? series.markerStyle?.lineCap
      ?? linkedGeometry?.lineCap,
    lineJoin: point?.markerStyle?.lineJoin ?? series.markerStyle?.lineJoin
      ?? linkedGeometry?.lineJoin,
    withLine: lineFamily || scatterSeriesDrawsLine('scatter', scatterStyle, series),
    directEffect: pointEffect ?? seriesEffect,
    fallbackEffect: linked,
    directEffectIndex: pointEffect ? pointIndex : styleIndex,
    fallbackEffectIndex: linkedIndex,
  };
}


export function drawLegendSwatch(
  ctx: CanvasRenderingContext2D,
  style: LegendSwatchStyle,
  color: string,
  x: number, y: number, w: number, h: number,
  marker: LegendMarker | null = null,
  /** undefined = no structured override (use `color`); null = authored
   * noFill; Fill = authored/resolved swatch paint. */
  fillPaint: Fill | null | undefined = undefined,
  outlinePaint: ChartModel['plotAreaLineFill'] | null | undefined = undefined,
  outlineColor: string | null = null,
  outlineWidthEmu: number | null = null,
  outlineDash: string | null = null,
  outlineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  outlineCap: string | null = null,
  outlineJoin: string | null = null,
  ptToPx = 1,
  shapeRotationDeg = 0,
  directEffect?: ChartExElementStyle | null,
  fallbackEffect?: ChartExElementStyle | null,
  directEffectIndex = 0,
  fallbackEffectIndex = directEffectIndex,
): void {
  if (style === 'none') return;
  // A line/scatter series with markers shows the same compound key as Excel:
  // connecting stroke first, then the marker centered on it. Markers-only
  // scatter skips the stroke.
  if (marker && !marker.withLine) {
    // Excel's legend marker is about 7pt beside a 12pt label; keeping it near
    // 0.58× the row height also leaves the surrounding key visually balanced.
    drawMarker(
      ctx, x + w / 2, y + h / 2, marker.symbol, h * 0.58 / ptToPx,
      marker.fill, marker.line, ptToPx,
      marker.lineWidthEmu != null ? axisLineWidthPx(marker.lineWidthEmu, ptToPx) : undefined,
      marker.fillPaint, shapeRotationDeg,
      marker.linePaint, marker.lineDash, marker.lineCustomDash,
      marker.lineCap, marker.lineJoin, marker.bubble3D,
      marker.directEffect, marker.fallbackEffect,
      marker.directEffectIndex, marker.fallbackEffectIndex,
    );
    return;
  }
  const body = (target: CanvasRenderingContext2D): void => drawLegendSwatchBody(
    target, style, color, x, y, w, h, fillPaint,
    outlinePaint, outlineColor, outlineWidthEmu, outlineDash,
    outlineCustomDash, outlineCap, outlineJoin, ptToPx, shapeRotationDeg,
  );
  paintChartStyleEffects(
    ctx, directEffect, fallbackEffect, directEffectIndex,
    { x, y, w, h }, ptToPx, body, fallbackEffectIndex,
  );
  if (marker) {
    drawMarker(
      ctx, x + w / 2, y + h / 2, marker.symbol, h * 0.58 / ptToPx,
      marker.fill, marker.line, ptToPx,
      marker.lineWidthEmu != null ? axisLineWidthPx(marker.lineWidthEmu, ptToPx) : undefined,
      marker.fillPaint, shapeRotationDeg,
      marker.linePaint, marker.lineDash, marker.lineCustomDash,
      marker.lineCap, marker.lineJoin, marker.bubble3D,
      marker.directEffect, marker.fallbackEffect,
      marker.directEffectIndex, marker.fallbackEffectIndex,
    );
  }
}


export function drawLegendSwatchBody(
  ctx: CanvasRenderingContext2D,
  style: LegendSwatchStyle,
  color: string,
  x: number, y: number, w: number, h: number,
  fillPaint: Fill | null | undefined,
  outlinePaint: ChartModel['plotAreaLineFill'] | null | undefined,
  outlineColor: string | null,
  outlineWidthEmu: number | null,
  outlineDash: string | null,
  outlineCustomDash: ChartModel['plotAreaLineCustomDash'],
  outlineCap: string | null,
  outlineJoin: string | null,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  ctx.fillStyle = color;
  if (style === 'line') {
    // A line key is the same authored stroke as the plotted series. Only the
    // fully automatic key retains the historical font-relative fallback.
    const hasAuthoredStroke = outlinePaint !== undefined || outlineColor != null
      || outlineWidthEmu != null
      || outlineDash != null
      || outlineCap != null
      || outlineJoin != null;
    if (!hasAuthoredStroke) {
      ctx.strokeStyle = color;
      const previousWidth = ctx.lineWidth;
      ctx.lineWidth = Math.max(1.5, h * 0.15);
      ctx.beginPath();
      const ly = y + h / 2;
      ctx.moveTo(x, ly);
      ctx.lineTo(x + w, ly);
      ctx.stroke();
      ctx.lineWidth = previousWidth;
      return;
    }
    if (outlinePaint === null) {
      return;
    }
    ctx.save();
    const resolvedOutline = outlinePaint
      ? resolveFill(outlinePaint, ctx, x, y, w, h, shapeRotationDeg)
      : outlineColor ? `#${outlineColor}` : color;
    if (!resolvedOutline) {
      ctx.restore();
      return;
    }
    ctx.strokeStyle = resolvedOutline;
    ctx.lineWidth = outlineWidthEmu != null
      ? axisLineWidthPx(outlineWidthEmu, ptToPx)
      : Math.max(1.5, h * 0.15);
    ctx.setLineDash(drawingmlLineDashArray(
      outlineCustomDash, outlineDash, ctx.lineWidth,
    ));
    ctx.lineCap = outlineCap === 'rnd' ? 'round' : outlineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = outlineJoin === 'round' || outlineJoin === 'bevel'
      ? outlineJoin
      : 'miter';
    ctx.beginPath();
    const ly = y + h / 2;
    ctx.moveTo(x, ly);
    ctx.lineTo(x + w, ly);
    ctx.stroke();
    ctx.restore();
  } else {
    paintClassicDataPointRect(
      ctx, fillPaint, { x, y, w, h }, color, ptToPx, shapeRotationDeg,
    );
    const resolvedOutline = outlinePaint === null
      ? null
      : outlinePaint
        ? resolveFill(outlinePaint, ctx, x, y, w, h, shapeRotationDeg)
        : outlineColor ? `#${outlineColor}` : null;
    if (resolvedOutline) {
      const outlineWidth = axisLineWidthPx(outlineWidthEmu, ptToPx);
      ctx.save();
      ctx.strokeStyle = resolvedOutline;
      ctx.lineWidth = outlineWidth;
      ctx.setLineDash(drawingmlLineDashArray(
        outlineCustomDash, outlineDash, ctx.lineWidth,
      ));
      ctx.lineCap = outlineCap === 'rnd' ? 'round' : outlineCap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = outlineJoin === 'round' || outlineJoin === 'bevel'
        ? outlineJoin
        : 'miter';
      ctx.strokeRect(
        x + outlineWidth / 2,
        y + outlineWidth / 2,
        Math.max(0, w - outlineWidth),
        Math.max(0, h - outlineWidth),
      );
      ctx.restore();
    }
  }
}


/** A single legend row: a label and the color of its swatch. Built so that the
 *  swatch color is resolved exactly like the mark it represents (slice / bar /
 *  line). See {@link legendEntryColor}. `marker` is set only for markers-only
 *  scatter series, whose key is a point glyph instead of the line swatch (#803). */
export interface LegendEntry {
  label: string;
  color: string;
  marker: LegendMarker | null;
  swatchStyle: LegendSwatchStyle;
  fillPaint: Fill | null | undefined;
  /** Effective mark outline paint; undefined keeps the legacy color fallback,
   * null is an authored noFill/unresolved paint and suppresses the key line. */
  outlinePaint: ChartModel['plotAreaLineFill'] | null | undefined;
  outlineColor: string | null;
  outlineWidthEmu: number | null;
  outlineDash: string | null;
  outlineCustomDash: ChartModel['plotAreaLineCustomDash'];
  outlineCap: string | null;
  outlineJoin: string | null;
  directEffect: ChartExElementStyle | null | undefined;
  fallbackEffect: ChartExElementStyle | null | undefined;
  directEffectIndex: number;
  fallbackEffectIndex: number;
  textOverride: ChartLegendEntryOverride | null;
}


export function applyLegendEntryOverrides(
  entries: readonly LegendEntry[],
  overrides: readonly ChartLegendEntryOverride[],
): LegendEntry[] {
  if (overrides.length === 0) return [...entries];
  const byIndex = new Map<number, ChartLegendEntryOverride>();
  for (const override of overrides) byIndex.set(override.idx, override);
  const effective: LegendEntry[] = [];
  for (let index = 0; index < entries.length; index++) {
    const override = byIndex.get(index);
    if (override?.deleted === true) continue;
    effective.push({ ...entries[index], textOverride: override ?? null });
  }
  return effective;
}


/** The legend key embedded in a classic data label (`<c:showLegendKey>`).
 * Styling is the same resolved {@link LegendEntry} used by the chart legend;
 * only placement belongs to the data-label layout. */
export interface DataLabelLegendKey {
  entry: LegendEntry;
  ptToPx: number;
  shapeRotationDeg: number;
}


/** Build the legend entries for a chart. Pie/doughnut and the explicitly
 *  resolved point-driven compatibility cases use one row per data point;
 *  ordinary charts use one row per series. */
export function buildLegendEntries(
  series: ChartSeries[],
  chartType: string | undefined,
  scatterStyle?: string | null,
  varyByPoint = false,
  chartCategories: string[] = [],
  fillPaints: ReadonlyArray<Fill | null | undefined> = [],
  pieVaryColors = true,
  entryOverrides: readonly ChartLegendEntryOverride[] = [],
  radarStyle?: string | null,
  chart?: ChartModel,
): LegendEntry[] {
  const categoryDriven = legendIsCategoryDriven(chartType, series.length, pieVaryColors);
  if (varyByPoint || categoryDriven) {
    // Point-driven: one entry per data point of the first series, labeled by
    // its category and colored exactly like the mark the plot draws for that
    // point (pie slice, varyColors bar, or string-X bubble).
    const first = series[0];
    const n = first ? first.values.length : 0;
    const cats = first?.categories ?? chartCategories;
    const sourceSeriesIndex = chart && first
      ? Math.max(0, chartSeriesSourceIndex(chart, first))
      : 0;
    const group = chart ? chartPlotGroupForSeries(chart, sourceSeriesIndex) : undefined;
    const family = group?.kind === 'bubble' ? 'bubble' : first?.seriesType ?? chartType;
    const effectiveScatterStyle = group?.scatterStyle ?? scatterStyle;
    const effectiveRadarStyle = group?.radarStyle ?? radarStyle;
    const overrides = new Map(first?.dataPointOverrides?.map(point => [point.idx, point]) ?? []);
    const entries = Array.from({ length: n }, (_, i) => {
      const point = overrides.get(i);
      const styleIndex = chart && first
        ? piePointStyleIndex(chart, first, sourceSeriesIndex, i) : i;
      const swatchStyle = legendSwatchStyle(family);
      const lineStyle = chart && first
        ? classicDataPointLineStyle(
            chart, swatchStyle === 'line' ? 'dataPointLine' : 'dataPoint',
            first, point, styleIndex,
          )
        : undefined;
      const resolvedFill = i < fillPaints.length
        ? fillPaints[i]
        : chart && first
          ? classicDataPointFillDecision(chart, first, point, styleIndex, i)
          : point?.fillHidden === true
            ? null
            : point?.color || first?.dataPointColors?.[i]
              ? undefined
              : first?.fillPattern ?? undefined;
      const bubbleMarker = family === 'bubble' && chart && first
        ? bubblePointLegendMarker(chart, first, point, i)
        : null;
      const marker = bubbleMarker ?? (chart && first
        ? classicPointLegendMarker(
            chart, first, point, i, sourceSeriesIndex, family,
            effectiveScatterStyle, effectiveRadarStyle,
          )
        : null);
      const pointEffect = chartStyleEffectOwner(point?.chartexStyle);
      const seriesEffect = chartStyleEffectOwner(first?.chartexStyle);
      const fallbackEffect = chart && first
        ? chartDataPointStyleRole(
            chart, swatchStyle === 'line' ? 'dataPointLine' : 'dataPoint',
            sourceSeriesIndex,
          )
        : undefined;
      const legacyColor = categoryDriven && first
        ? pieSliceColor(i, first, pieVaryColors, 0)
        : legendEntryColor(chartType, series, i, varyByPoint, pieVaryColors);
      return {
        label: (cats[i] ?? `Item ${i + 1}`).toString(),
        color: resolvedFill?.fillType === 'solid'
          ? (resolvedFill.color.startsWith('#') ? resolvedFill.color : `#${resolvedFill.color}`)
          : legacyColor,
        marker,
        swatchStyle,
        fillPaint: resolvedFill,
        outlinePaint: lineStyle?.paint?.fillType === 'solid' ? undefined : lineStyle?.paint,
        outlineColor: lineStyle?.paint?.fillType === 'solid'
          ? lineStyle.paint.color : point?.lineColor ?? first?.lineColor ?? null,
        outlineWidthEmu: lineStyle?.widthEmu ?? null,
        outlineDash: lineStyle?.dash ?? null,
        outlineCustomDash: lineStyle?.customDash,
        outlineCap: lineStyle?.cap ?? null,
        outlineJoin: lineStyle?.join ?? null,
        directEffect: pointEffect ?? seriesEffect,
        fallbackEffect,
        directEffectIndex: pointEffect
          ? i : first ? chartExSeriesFormatIndex(first, sourceSeriesIndex) : styleIndex,
        fallbackEffectIndex: styleIndex,
        textOverride: null,
      };
    });
    return applyLegendEntryOverrides(entries, entryOverrides);
  }
  if (chart && series.some((candidate) => {
    const sourceIndex = chartSeriesSourceIndex(chart, candidate);
    const group = chartPlotGroupForSeries(chart, sourceIndex);
    return sourceIndex >= 0 && group?.kind !== 'pie' && group?.kind !== 'pie3D'
      && group?.kind !== 'doughnut' && group?.kind !== 'ofPie'
      && chartSeriesVariesByPoint(chart, sourceIndex);
  })) {
    // `varyColors` belongs to a chart-group child, not the plot area. Expand
    // only the series owned by a varying single-series group and retain every
    // unrelated combo series in its original order.
    const entries = series.flatMap((candidate): LegendEntry[] => {
      const sourceIndex = chartSeriesSourceIndex(chart, candidate);
      const group = chartPlotGroupForSeries(chart, sourceIndex);
      const family = group?.kind === 'bubble' ? 'bubble' : candidate.seriesType ?? chartType;
      const pointDriven = sourceIndex >= 0 && group?.kind !== 'pie'
        && group?.kind !== 'pie3D' && group?.kind !== 'doughnut'
        && group?.kind !== 'ofPie' && chartSeriesVariesByPoint(chart, sourceIndex);
      if (pointDriven) {
        return buildLegendEntries(
          [candidate], family, group?.scatterStyle ?? scatterStyle, true,
          candidate.categories ?? chartCategories, [], pieVaryColors, [],
          group?.radarStyle ?? radarStyle, chart,
        );
      }
      return buildLegendEntries(
        [candidate], family, group?.scatterStyle ?? scatterStyle, false,
        candidate.categories ?? chartCategories, [], pieVaryColors, [],
        group?.radarStyle ?? radarStyle, chart,
      );
    });
    return applyLegendEntryOverrides(entries, entryOverrides);
  }
  const entries = series.map((s, i) => {
    // A combo chart has multiple chart groups under one plotArea. The legend
    // key describes the individual series' group, not the first/primary group.
    const sourceSeriesIndex = chart ? chartSeriesSourceIndex(chart, s, i) : i;
    const group = chart ? chartPlotGroupForSeries(chart, sourceSeriesIndex) : undefined;
    const family = group?.kind === 'bubble' ? 'bubble' : s.seriesType ?? chartType;
    const lineVisible = s.lineHidden !== true;
    const lineColor = lineVisible ? (s.lineColor ?? null) : null;
    const marker = legendMarkerFor(chartType, scatterStyle, radarStyle, series, i, chart);
    const swatchStyle: LegendSwatchStyle = family === 'stock' && !lineVisible && !marker
      ? 'none'
      : legendSwatchStyle(family);
    const styleIndex = chartExSeriesFormatIndex(s, sourceSeriesIndex >= 0 ? sourceSeriesIndex : i);
    const resolvedFill = i < fillPaints.length
      ? fillPaints[i]
      : chart
        ? classicDataPointFillDecision(chart, s, undefined, styleIndex)
        : chartStyleFillDecision(s.chartexStyle, s.chartexFormatIdx ?? i)
          ?? s.fillPattern ?? undefined;
    const lineStyle = chart
      ? classicDataPointLineStyle(
          chart, swatchStyle === 'line' ? 'dataPointLine' : 'dataPoint',
          s, undefined, styleIndex,
        )
      : undefined;
    const fallbackEffect = chart
      ? chartDataPointStyleRole(
          chart, swatchStyle === 'line' ? 'dataPointLine' : 'dataPoint',
          sourceSeriesIndex,
        )
      : undefined;
    return {
      label: s.name || `Series ${i + 1}`,
      color: swatchStyle === 'line' && lineStyle?.paint?.fillType === 'solid'
        ? (lineStyle.paint.color.startsWith('#')
            ? lineStyle.paint.color : `#${lineStyle.paint.color}`)
        : swatchStyle === 'line' && lineColor
          ? `#${lineColor}`
          : resolvedFill?.fillType === 'solid'
            ? (resolvedFill.color.startsWith('#')
                ? resolvedFill.color : `#${resolvedFill.color}`)
            : legendEntryColor(chartType, series, i, false, pieVaryColors),
      marker,
      swatchStyle,
      fillPaint: resolvedFill,
      outlinePaint: lineStyle?.paint?.fillType === 'solid' ? undefined : lineStyle?.paint,
      outlineColor: lineStyle?.paint?.fillType === 'solid'
        ? lineStyle.paint.color : lineColor,
      // A DrawingML noFill line may still carry width/dash/cap/join. Those
      // geometry attributes do not make the stroke visible and must not revive
      // it in the legend through the automatic-color fallback.
      outlineWidthEmu: lineStyle?.widthEmu ?? (lineVisible ? (s.lineWidthEmu ?? null) : null),
      outlineDash: lineStyle?.dash ?? (lineVisible ? (s.chartexStyle?.lineDash ?? null) : null),
      outlineCustomDash: lineStyle?.customDash,
      outlineCap: lineStyle?.cap ?? (lineVisible ? (s.chartexStyle?.lineCap ?? null) : null),
      outlineJoin: lineStyle?.join ?? (lineVisible ? (s.chartexStyle?.lineJoin ?? null) : null),
      directEffect: chartStyleEffectOwner(s.chartexStyle),
      fallbackEffect,
      directEffectIndex: styleIndex,
      fallbackEffectIndex: styleIndex,
      textOverride: null,
    };
  });
  return applyLegendEntryOverrides(entries, entryOverrides);
}


/** Resolve data-label keys through the chart's existing legend-style pipeline.
 * ECMA-376 §21.2.2.179 only requires the corresponding legend key to be shown;
 * it does not define a second paint model. Category-driven families therefore
 * use the point entry, while every other family uses its series entry. */
export function createDataLabelLegendKeyResolver(
  chart: ChartModel,
  ptToPx: number,
  shapeRotationDeg = 0,
): (seriesIndex: number, pointIndex: number) => DataLabelLegendKey | undefined {
  const categoryDriven = chartVariesColorsByPoint(chart)
    || legendIsCategoryDriven(
      chart.chartType,
      chart.series.length,
      chart.varyColors !== false,
    );
  const entries = buildLegendEntries(
    chart.series,
    chart.chartType,
    chart.scatterStyle,
    chartVariesColorsByPoint(chart),
    chart.categories,
    [],
    chart.varyColors !== false,
    [],
    chart.radarStyle,
    chart,
  );
  const ranges = legendEntryRanges(chart);
  return (seriesIndex, pointIndex) => {
    const pointDriven = categoryDriven || chartSeriesVariesByPoint(chart, seriesIndex);
    const entryIndex = categoryDriven
      ? pointIndex
      : legendEntryGlobalIndex(ranges, seriesIndex, pointDriven ? pointIndex : 0);
    const entry = entryIndex == null ? undefined : entries[entryIndex];
    return entry ? { entry, ptToPx, shapeRotationDeg } : undefined;
  };
}


/** Resolved legend text styling (CH10). `fontFamily` already carries the
 *  theme-body fallback; `sizePx` overrides the shared automatic 10pt size only
 *  when the file authored one. */
export interface LegendTextStyle {
  fontFamily: string;
  color: string;
  bold: boolean;
  italic: boolean;
  sizePx: number | null;
}


export function legendEntryTextStyle(
  chart: ChartModel,
  base: LegendTextStyle,
  override: ChartLegendEntryOverride | null,
  ptToPx: number,
): LegendTextStyle {
  if (!override) return base;
  const face = resolveThemeFontRef(chart, override.fontFace) ?? override.fontFace;
  return {
    fontFamily: face ? `"${face}", Calibri, Arial, sans-serif` : base.fontFamily,
    color: override.fontColor ? `#${override.fontColor}` : base.color,
    bold: override.fontBold ?? base.bold,
    italic: override.fontItalic ?? base.italic,
    sizePx: chartTextFontSizePx(override.fontSizeHpt, ptToPx) ?? base.sizePx,
  };
}


export function setLegendFont(ctx: CanvasRenderingContext2D, style: LegendTextStyle, ptToPx: number): number {
  const size = legendFontSizePx(style, ptToPx);
  ctx.font = `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${size}px ${style.fontFamily}`;
  return size;
}


export const DEFAULT_LEGEND_STYLE: LegendTextStyle = {
  fontFamily: 'sans-serif',
  color: '#333',
  bold: false,
  italic: false,
  sizePx: null,
};


export const LEGEND_SWATCH_TEXT_GAP = 4;

export const LEGEND_ITEM_GAP = 12;

export const LEGEND_ROW_EXTRA_PX = 4;

export const LEGEND_HORIZONTAL_INSET = 4;

export const LEGEND_HORIZONTAL_PADDING = LEGEND_HORIZONTAL_INSET * 2;

export const LEGEND_VERTICAL_PADDING = 4;

export const LEGEND_SIDE_PADDING = 8;

// The item width is measured as key + gap + text, then split back into those
// components for paint. Fractional Canvas metrics can lose one ULP in that
// subtraction (e.g. 24.453475952148438 becomes 24.453475952148434), falsely
// eliding a label that was already proven to fit. A hundredth of a CSS pixel is
// safely below a visible/device-pixel overflow while absorbing that arithmetic
// roundoff.
export const LEGEND_MEASUREMENT_EPSILON_PX = 0.01;

// Office vector output keeps a filled legend key at roughly 7pt square even
// when the legend text is larger (for example, a 15pt legend still has a 7pt
// box). Line/marker keys remain font-relative because their glyph geometry is
// tied to the legend row rather than a filled-area key.
export const FILLED_LEGEND_KEY_PT = 7;


export function legendFontSizePx(style: LegendTextStyle, ptToPx: number): number {
  return style.sizePx ?? 10 * ptToPx;
}


export function legendSwatchWidths(
  entries: readonly LegendEntry[],
  fontSize: number,
  ptToPx: number,
): number[] {
  return entries.map(entry => {
    if (entry.swatchStyle === 'fill') return FILLED_LEGEND_KEY_PT * ptToPx;
    const fallback = fontSize * 1.6;
    if (entry.swatchStyle !== 'line' || !entry.outlineDash) return fallback;
    const lineWidth = entry.outlineWidthEmu != null
      ? axisLineWidthPx(entry.outlineWidthEmu, ptToPx)
      : Math.max(1.5, fontSize * 0.15);
    const dash = dashPatternForPreset(entry.outlineDash, lineWidth);
    // End a legend key at a complete dash boundary. One whole pattern plus
    // its first dash renders two complete visible strokes for the common
    // DrawingML `dash` preset instead of clipping the second stroke midway.
    const completePatternWidth = dash.length > 0
      ? dash.reduce((sum, length) => sum + length, 0) + dash[0]
      : 0;
    return Math.max(fallback, completePatternWidth);
  });
}


export function legendSwatchHeight(entry: LegendEntry, fontSize: number, ptToPx: number): number {
  return entry.swatchStyle === 'fill' ? FILLED_LEGEND_KEY_PT * ptToPx : fontSize;
}


/** Resolve side-legend text into the same bounded lines that paint consumes.
 * Office wraps a long series name at word boundaries in a left/right legend
 * (rather than replacing the whole second half with an ellipsis). Two lines
 * keep the automatic side band bounded; only text that still exceeds that
 * contract is elided on the final line. */
export function sideLegendLabelLines(
  ctx: CanvasRenderingContext2D,
  label: string,
  maxWidth: number,
): string[] {
  if (!(maxWidth > 0)) return [];
  const wrapped = wrapMeasuredText(ctx, label, maxWidth);
  if (wrapped.length <= 2) return wrapped;
  return [
    wrapped[0],
    elideToWidth(ctx, wrapped.slice(1).join(' '), maxWidth),
  ];
}


export interface MeasuredLegendLayout extends ChartLegendReserve {
  measuredLabels: string[];
  entryStyles: LegendTextStyle[];
  fontSizes: number[];
  swatches: number[];
  itemWidths: number[];
}


/** Resolve the shared automatic legend reserve from real Canvas text metrics. */
export function measuredLegendReserve(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  w: number,
  h: number,
  sideReserveFrac: number,
  ptToPx: number,
): MeasuredLegendLayout | null {
  if (!chart.showLegend) return null;
  const style = legendTextStyle(chart, ptToPx);
  const entries = buildLegendEntries(
    legendSeriesWithTrendlines(chart),
    chart.chartType,
    chart.scatterStyle,
    chartVariesColorsByPoint(chart),
    chart.categories,
    [],
    chart.varyColors !== false,
    chart.legendEntries ?? [],
    chart.radarStyle,
    chart,
  );
  const entryStyles = entries.map(entry =>
    legendEntryTextStyle(chart, style, entry.textOverride, ptToPx)
  );
  const fontSizes = entryStyles.map(entryStyle => legendFontSizePx(entryStyle, ptToPx));
  const swatches = entries.map((entry, index) =>
    legendSwatchWidths([entry], fontSizes[index], ptToPx)[0]
  );
  ctx.save();
  const itemWidths = entries.map((entry, index) => {
    setLegendFont(ctx, entryStyles[index], ptToPx);
    return swatches[index] + LEGEND_SWATCH_TEXT_GAP + ctx.measureText(entry.label).width;
  });
  ctx.restore();
  const pos = chart.legendPos ?? 'r';
  const horizontal = pos === 't' || pos === 'b';
  const reserve = chartLegendReserve(chart, w, h, sideReserveFrac, {
    itemWidths,
    rowHeight: Math.max(0, ...fontSizes) + LEGEND_ROW_EXTRA_PX,
    itemGap: LEGEND_ITEM_GAP,
    horizontalPadding: horizontal ? LEGEND_HORIZONTAL_PADDING : LEGEND_SIDE_PADDING,
    verticalPadding: LEGEND_VERTICAL_PADDING,
  });
  return reserve ? {
    ...reserve,
    measuredLabels: entries.map(entry => entry.label),
    entryStyles,
    fontSizes,
    swatches,
    itemWidths,
  } : null;
}


export function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartSeries[],
  lx: number, ly: number, lw: number, lh: number,
  orient: 'vertical' | 'horizontal' = 'vertical',
  chartType?: string,
  style: LegendTextStyle = DEFAULT_LEGEND_STYLE,
  scatterStyle?: string | null,
  varyByPoint = false,
  chartCategories: string[] = [],
  ptToPx = 1,
  fillPaints: ReadonlyArray<Fill | null | undefined> = [],
  shapeRotationDeg = 0,
  pieVaryColors = true,
  chartForEntryStyles?: ChartModel,
  measured?: MeasuredLegendLayout | null,
): void {
  const gap = LEGEND_SWATCH_TEXT_GAP;
  const entries = buildLegendEntries(
    series,
    chartType,
    scatterStyle,
    varyByPoint,
    chartCategories,
    fillPaints,
    pieVaryColors,
    chartForEntryStyles?.legendEntries ?? [],
    chartForEntryStyles?.radarStyle,
    chartForEntryStyles,
  );
  const canReuseMeasure = measured != null
    && measured.measuredLabels.length === entries.length
    && measured.measuredLabels.every((label, index) => label === entries[index].label);
  const entryStyles = canReuseMeasure
    ? measured.entryStyles
    : entries.map(entry => chartForEntryStyles
        ? legendEntryTextStyle(chartForEntryStyles, style, entry.textOverride, ptToPx)
        : style
      );
  const fontSizes = canReuseMeasure
    ? measured.fontSizes
    : entryStyles.map(entryStyle => legendFontSizePx(entryStyle, ptToPx));
  if (entryStyles[0]) setLegendFont(ctx, entryStyles[0], ptToPx);
  ctx.textBaseline = 'middle';
  const rowH = Math.max(0, ...fontSizes) + LEGEND_ROW_EXTRA_PX;
  const swatches = canReuseMeasure
    ? measured.swatches
    : entries.map((entry, index) =>
        legendSwatchWidths([entry], fontSizes[index], ptToPx)[0]
      );
  const itemWidths = canReuseMeasure
    ? measured.itemWidths
    : entries.map((entry, index) => {
        setLegendFont(ctx, entryStyles[index], ptToPx);
        return swatches[index] + gap + ctx.measureText(entry.label).width;
      });
  if (orient === 'horizontal') {
    const rows = packLegendRows(itemWidths, lw, LEGEND_ITEM_GAP);
    const visibleRows = rows.slice(
      0,
      Math.max(0, Math.floor((lh - LEGEND_VERTICAL_PADDING) / rowH)),
    );
    const top = ly + LEGEND_VERTICAL_PADDING / 2;
    for (let rowIndex = 0; rowIndex < visibleRows.length; rowIndex++) {
      const row = visibleRows[rowIndex];
      const widths = row.map(index => Math.min(lw, itemWidths[index]));
      const total = widths.reduce((sum, width) => sum + width, 0)
        + LEGEND_ITEM_GAP * Math.max(0, row.length - 1);
      let rx = lx + Math.max(0, (lw - total) / 2);
      const ry = top + rowIndex * rowH + rowH / 2;
      for (let item = 0; item < row.length; item++) {
        const index = row[item];
        const sw = swatches[index];
        const effectiveWidth = widths[item];
        if (effectiveWidth < sw) {
          rx += effectiveWidth + LEGEND_ITEM_GAP;
          continue;
        }
        const maxTextPx = Math.max(
          0,
          effectiveWidth - sw - gap + LEGEND_MEASUREMENT_EPSILON_PX,
        );
        setLegendFont(ctx, entryStyles[index], ptToPx);
        const label = elideToWidth(ctx, entries[index].label, maxTextPx);
        const swatchH = legendSwatchHeight(entries[index], fontSizes[index], ptToPx);
        drawLegendSwatch(
          ctx, entries[index].swatchStyle, entries[index].color,
          rx, ry - swatchH / 2, sw, swatchH,
          entries[index].marker, entries[index].fillPaint,
          entries[index].outlinePaint,
          entries[index].outlineColor, entries[index].outlineWidthEmu,
          entries[index].outlineDash, entries[index].outlineCustomDash,
          entries[index].outlineCap, entries[index].outlineJoin,
          ptToPx, shapeRotationDeg,
          entries[index].directEffect, entries[index].fallbackEffect,
          entries[index].directEffectIndex, entries[index].fallbackEffectIndex,
        );
        ctx.fillStyle = entryStyles[index].color;
        ctx.textAlign = 'left';
        ctx.fillText(label, rx + sw + gap, ry);
        rx += effectiveWidth + LEGEND_ITEM_GAP;
      }
    }
    return;
  }
  // Vertical legend: each label runs from just after the swatch to the right
  // edge of the reserved legend column. Long series names wrap to at most two
  // measured lines; use those same lines to plan row heights and to paint, so
  // reserve and draw cannot disagree about which words fit.
  const maxTextPx = lw - Math.max(...swatches, 0) - gap;
  // Point/category legends can contain dozens of independent keys; keep their
  // historical one-line/elided rows so wrapping one category cannot starve
  // later entries. Series-driven legends are the bounded multi-line case.
  const wrapSeriesNames = !varyByPoint
    && !legendIsCategoryDriven(chartType, series.length, pieVaryColors);
  const labelLines = entries.map((entry, index) => {
    setLegendFont(ctx, entryStyles[index], ptToPx);
    return wrapSeriesNames
      ? sideLegendLabelLines(ctx, entry.label, maxTextPx)
      : [elideToWidth(ctx, entry.label, maxTextPx)];
  });
  const entryHeights = labelLines.map((lines, index) =>
    lines.length * fontSizes[index] + LEGEND_ROW_EXTRA_PX
  );
  let visibleCount = 0;
  let visibleHeight = 0;
  while (visibleCount < entries.length
    && visibleHeight + entryHeights[visibleCount] <= lh) {
    visibleHeight += entryHeights[visibleCount];
    visibleCount++;
  }
  let ry = visibleCount === entries.length
    ? ly + (lh - visibleHeight) / 2
    : ly;
  for (let i = 0; i < visibleCount; i++) {
    const sw = swatches[i];
    const entryH = entryHeights[i];
    if (lw < sw) {
      ry += entryH;
      continue;
    }
    const swatchH = legendSwatchHeight(entries[i], fontSizes[i], ptToPx);
    drawLegendSwatch(
      ctx, entries[i].swatchStyle, entries[i].color,
      lx, ry + (entryH - swatchH) / 2, sw, swatchH,
      entries[i].marker, entries[i].fillPaint,
      entries[i].outlinePaint,
      entries[i].outlineColor, entries[i].outlineWidthEmu,
      entries[i].outlineDash, entries[i].outlineCustomDash,
      entries[i].outlineCap, entries[i].outlineJoin,
      ptToPx, shapeRotationDeg,
      entries[i].directEffect, entries[i].fallbackEffect,
      entries[i].directEffectIndex, entries[i].fallbackEffectIndex,
    );
    setLegendFont(ctx, entryStyles[i], ptToPx);
    ctx.fillStyle = entryStyles[i].color; ctx.textAlign = 'left';
    labelLines[i].forEach((line, lineIndex) =>
      // Preserve the established single-line baseline byte-for-byte. Extra
      // wrapped lines continue at one authored font-size interval.
      ctx.fillText(line, lx + sw + gap, ry + fontSizes[i] * (lineIndex + 0.5))
    );
    ry += entryH;
  }
}


/** Build the resolved legend text style for a chart (CH10). Absent legend
 *  `<c:txPr>` fields use the theme minor face when available and the shared
 *  automatic defaults otherwise. */
export function legendTextStyle(chart: ChartModel, ptToPx: number): LegendTextStyle {
  const face = resolveThemeFontRef(chart, chart.legendFontFace) ?? chart.themeMinorFontLatin;
  return {
    fontFamily: face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif',
    color: chart.legendFontColor ? `#${chart.legendFontColor}` : '#333',
    bold: chart.legendFontBold ?? false,
    italic: chart.legendFontItalic ?? false,
    sizePx: chartTextFontSizePx(chart.legendFontSizeHpt, ptToPx),
  };
}


// Legend placement is resolved by `chartLegendReserve` (layout.ts). This alias
// keeps the drawing helper's signature readable while sharing the single source
// of truth for the reserve shape.
export type LegendLayout = MeasuredLegendLayout;


/** Draw a legend in the band reserved by {@link chartLegendReserve}. */
export function drawLegendForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  leg: LegendLayout | null,
  x: number, y: number, w: number, h: number,
  _px0: number, py0: number, _pw: number, ph: number,
  topBand: number,
  ptToPx: number,
  fillPaints: ReadonlyArray<Fill | null | undefined> = [],
  shapeRotationDeg = 0,
): void {
  if (!leg) return;
  const legStyle = legendTextStyle(chart, ptToPx);
  const legendSeries = legendSeriesWithTrendlines(chart);
  // Point-driven legends list one entry per data point, so key paint and plot
  // paint resolve through the same point precedence.
  const varyByPoint = chartVariesColorsByPoint(chart);
  const sideInset = Math.min(
    LEGEND_SIDE_PADDING / 2,
    Math.max(0, leg.reserveW) / 2,
  );
  const sideContentWidth = Math.max(0, leg.reserveW - sideInset * 2);
  const defaultBox = leg.side === 'r'
    ? {
        x: x + w - leg.reserveW + sideInset,
        y: py0,
        w: sideContentWidth,
        h: ph,
      }
    : leg.side === 'l'
      ? { x: x + sideInset, y: py0, w: sideContentWidth, h: ph }
      : leg.side === 't'
        ? {
            x: x + LEGEND_HORIZONTAL_INSET,
            y: y + topBand,
            w: Math.max(0, w - LEGEND_HORIZONTAL_PADDING),
            h: leg.reserveH,
          }
        : {
            x: x + LEGEND_HORIZONTAL_INSET,
            y: y + h - leg.reserveH,
            w: Math.max(0, w - LEGEND_HORIZONTAL_PADDING),
            h: leg.reserveH,
          };
  const defaultOrientation = leg.side === 't' || leg.side === 'b' ? 'horizontal' : 'vertical';
  // `<c:legend><c:manualLayout>` (§21.2.2.31) wins over the side-based
  // rectangle. The shared resolver applies all four factor/edge modes relative
  // to this automatic box, including the schema's omitted-mode=factor default.
  const ml = chart.legendManualLayout;
  const manualBox = ml
    ? resolveManualLayoutRect(ml, { x, y, w, h }, defaultBox)
    : null;
  if (manualBox) {
    const orient = manualBox.w >= manualBox.h ? 'horizontal' : 'vertical';
    paintLegendFrame(ctx, chart, manualBox, ptToPx, shapeRotationDeg);
    drawLegend(ctx, legendSeries, manualBox.x, manualBox.y, manualBox.w, manualBox.h, orient, chart.chartType, legStyle, chart.scatterStyle, varyByPoint, chart.categories, ptToPx, fillPaints, shapeRotationDeg, chart.varyColors !== false, chart, leg);
    return;
  }
  paintLegendFrame(ctx, chart, defaultBox, ptToPx, shapeRotationDeg);
  drawLegend(ctx, legendSeries, defaultBox.x, defaultBox.y, defaultBox.w, defaultBox.h,
    defaultOrientation, chart.chartType, legStyle, chart.scatterStyle, varyByPoint,
    chart.categories, ptToPx, fillPaints, shapeRotationDeg, chart.varyColors !== false, chart, leg);
}

/** Resolve the color for legend entry `entryIndex`, matching the marks the
 *  plot actually draws.
 *
 *  - Category-driven legends (pie and varying doughnut): the entry maps to
 *    data point `entryIndex` of the first series, so it must use the *same*
 *    resolution as {@link pieSliceColor} — explicit per-point `dPt` color,
 *    else the palette indexed by point. The series-level fill is deliberately
 *    ignored: a pie series carries a single `<c:spPr>` solidFill that, if
 *    honored here, would collapse every swatch to one color while the slices
 *    stay multi-colored. An explicitly non-varying multi-series doughnut is
 *    series-driven instead, matching Excel's ring legend.
 *  - Series-driven legends (bar / line / area / …): the entry maps to series
 *    `entryIndex`, so it uses {@link chartColor} — explicit series fill else
 *    the palette indexed by series. */
export function legendEntryColor(
  chartType: string | undefined,
  series: ChartSeries[],
  entryIndex: number,
  varyByPoint = false,
  pieVaryColors = true,
): string {
  if (varyByPoint || legendIsCategoryDriven(chartType, series.length, pieVaryColors)) {
    const first = series[0];
    if (first) return pieSliceColor(entryIndex, first, pieVaryColors, 0);
    return `#${CHART_PALETTE[entryIndex % CHART_PALETTE.length]}`;
  }
  return chartColor(entryIndex, series[entryIndex]);
}
