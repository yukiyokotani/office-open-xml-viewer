// Classic chart chartex style helpers.
import type { ChartModel, ChartRect, ChartSeries, ChartStyleRole } from '../../types/chart';
import { chartStyleColor, chartStyleDirectFillDecision, chartStyleDirectLineDecision, chartStyleDirectNoFillDecision, chartStyleDirectNoLineDecision, chartStyleFillDecision, chartStyleLineDecision } from '../style-paint.js';
import type { Fill } from '../../types/common';
import { rawLinkedChartStyleRole } from '../effective-style.js';
import { resolveFill } from '../../shape/paint.js';
import { paintChartImageFill } from '../image-fill.js';
import { axisLineWidthPx } from '../axis-style.js';
import { CHARTEX_DEFAULT_PALETTE } from './palette.js';
import { dashPatternForLine } from './geometry.js';


// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

export type ChartExStyle = NonNullable<ChartModel['chartexDataPointStyle']>;


/** Resolve CT_DataLabels + indexed CT_DataLabel/dataLabelHidden without
 * renderer-specific precedence. Defaults describe the semantic label layer of
 * the chart type; authored visibility always overrides those defaults. */
export function chartExStyleColor(
  _chart: ChartModel,
  style: ChartExStyle | null | undefined,
  kind: 'fill' | 'line',
  index: number,
  _count: number,
): string | null {
  return chartStyleColor(style, kind, index);
}


export function chartExPaletteColor(
  chart: ChartModel,
  colors: ReadonlyArray<string | null | undefined>,
  colorIndex: number,
  _count: number,
): string | null {
  if (!colors.length) return null;
  const method = chart.chartexColorStyleMethod;
  const knownMethod = method === 'withinLinear'
    || method === 'acrossLinear'
    || method === 'withinLinearReversed'
    || method === 'acrossLinearReversed';
  // MS-ODRAWXML §2.8.4.1: unknown method strings have cycle semantics.
  if (!knownMethod) return colors[colorIndex % colors.length] ?? null;
  const within = method === 'withinLinear' || method === 'withinLinearReversed';
  // The specification defines which base color linear methods use, but does
  // not define the brightness range or color space. Preserve the authored
  // color here instead of inventing an Office compatibility curve. Once an
  // observed/approved rule exists, brightness belongs before styleClr/style
  // matrix transforms in the shared parser model, not as a post-paint tweak.
  return colors[within ? 0 : colorIndex % colors.length] ?? null;
}


export function chartExSemanticFill(chart: ChartModel, index: number, count: number): string {
  return (chart.chartexColorPalette
      ? chartExPaletteColor(chart, chart.chartexColorPalette, index, count)
      : null)
    ?? chart.chartexAccents?.[index % (chart.chartexAccents.length || 1)]
    ?? CHARTEX_DEFAULT_PALETTE[index % CHARTEX_DEFAULT_PALETTE.length];
}


export function chartExDataPointFill(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle?: ChartExStyle | null,
): string {
  return chartExStyleColor(chart, localStyle, 'fill', index, count)
    ?? chartExStyleColor(chart, chart.chartexDataPointStyle, 'fill', index, count)
    ?? chartExSemanticFill(chart, index, count);
}


/** Line-paint counterpart of chartExStylePaintDecision. `undefined` means the
 * layer did not author a line paint, while `null` is an authored noFill or an
 * authored-but-unresolved paint that must suppress lower-precedence color. */
export function chartExStyleLinePaintDecision(
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  index: number,
  count: number,
): ChartModel['plotAreaLineFill'] | null | undefined {
  void chart;
  void count;
  return chartStyleLineDecision(style, index);
}


/** Resolve one ChartEx style paint layer. `undefined` means this layer supplied
 * no paint, while `null` records an explicit no-fill for consumers whose own
 * shape is governed by that layer. */
export function chartExStylePaintDecision(
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  index: number,
  count: number,
): Fill | null | undefined {
  void chart;
  void count;
  return chartStyleFillDecision(style, index);
}


export function chartExMarkerPaint(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle: ChartExStyle | null | undefined,
  legacyColor: string | null | undefined,
  linkedStyle: ChartExStyle | null | undefined,
): Fill | null {
  const role = linkedStyle === chart.chartexDataPointMarkerStyle
    ? 'dataPointMarker'
    : linkedStyle === chart.chartexDataPointStyle
      ? 'dataPoint'
      : (Object.entries(chart.chartStyleRoles ?? {}).find(
          ([, style]) => style === linkedStyle,
        )?.[0] as ChartStyleRole | undefined);
  const rawLinkedStyle = role
    ? rawLinkedChartStyleRole(chart, role)
      ?? (chart.classicChartStyleRoles == null ? linkedStyle : undefined)
    : linkedStyle;
  const local = chartStyleDirectFillDecision(localStyle, rawLinkedStyle, index);
  if (local !== undefined) return local;
  if (legacyColor) return { fillType: 'solid', color: legacyColor };
  const linked = chartExStylePaintDecision(chart, linkedStyle, index, count);
  if (linked !== undefined) return linked;
  return { fillType: 'solid', color: chartExSemanticFill(chart, index, count) };
}


export function chartExDataPointPaint(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle?: ChartExStyle | null,
  legacyColor?: string | null,
  linkedStyle: ChartExStyle | null | undefined = chart.chartexDataPointStyle,
): Fill | null {
  // CT_Series.spPr formats the series shape; ChartEx semantic data points
  // (waterfall roles, box bodies, hierarchy nodes) still obtain their own
  // paint from the dataPoint Chart Style. A conventional series-level
  // `<a:noFill>` therefore does not erase every point. Positive local series
  // fills remain direct formatting and do override the linked recipe.
  const rawLinkedStyle = linkedStyle === chart.chartexDataPointStyle
    ? rawLinkedChartStyleRole(chart, 'dataPoint')
      ?? (chart.classicChartStyleRoles == null ? linkedStyle : undefined)
    : linkedStyle;
  const local = localStyle?.fillHidden
    ? chartStyleDirectNoFillDecision(rawLinkedStyle)
    : chartStyleDirectFillDecision(localStyle, rawLinkedStyle, index);
  if (local !== undefined) return local;
  if (localStyle && legacyColor) return { fillType: 'solid', color: legacyColor };
  if (legacyColor) return { fillType: 'solid', color: legacyColor };
  const linked = chartStyleFillDecision(linkedStyle, index);
  if (linked !== undefined) return linked;
  return { fillType: 'solid', color: chartExSemanticFill(chart, index, count) };
}


export function chartExFillStyle(
  ctx: CanvasRenderingContext2D,
  paint: Fill,
  x: number,
  y: number,
  w: number,
  h: number,
  fallbackColor: string,
  shapeRotationDeg = 0,
): string | CanvasGradient | CanvasPattern {
  // Keep solid ChartEx paints byte-compatible with the renderer's historical
  // `#RRGGBB` path. The shared resolver is needed only for structured fills;
  // routing solids through it would rewrite equivalent colors as rgba().
  if (paint.fillType === 'solid') {
    return paint.color.startsWith('#') ? paint.color : `#${paint.color}`;
  }
  return resolveFill(paint, ctx, x, y, w, h, shapeRotationDeg) ?? fallbackColor;
}


/** Paint an already-constructed classic mark path. Picture fills require the
 * current geometric path as an outer clip; the shared image painter then owns
 * crop/tile/stretch and its rectangular destination clip. Missing decoded
 * sources remain transparent rather than reviving the semantic fallback. */
export function paintClassicDataPointPath(
  ctx: CanvasRenderingContext2D,
  paint: Fill | null | undefined,
  bounds: ChartRect,
  fallbackColor: string,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (paint === null) return false;
  if (paint?.fillType === 'image') {
    ctx.save();
    ctx.clip();
    const painted = paintChartImageFill(
      ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h, ptToPx, shapeRotationDeg,
    );
    ctx.restore();
    return painted;
  }
  ctx.fillStyle = paint
    ? chartExFillStyle(
        ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h,
        fallbackColor, shapeRotationDeg,
      )
    : fallbackColor;
  ctx.fill();
  return true;
}


export function paintClassicDataPointRect(
  ctx: CanvasRenderingContext2D,
  paint: Fill | null | undefined,
  bounds: ChartRect,
  fallbackColor: string,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (paint === null || !(bounds.w > 0) || !(bounds.h > 0)) return false;
  if (paint?.fillType === 'image') {
    ctx.beginPath();
    ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
    return paintClassicDataPointPath(
      ctx, paint, bounds, fallbackColor, ptToPx, shapeRotationDeg,
    );
  }
  ctx.fillStyle = paint
    ? chartExFillStyle(
        ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h,
        fallbackColor, shapeRotationDeg,
      )
    : fallbackColor;
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  return true;
}


export interface ResolvedChartExLineStyle {
  visible: boolean;
  color: string;
  paint: ChartModel['plotAreaLineFill'] | null | undefined;
  widthEmu: number | null;
  dash: string | null;
  customDash: ChartModel['plotAreaLineCustomDash'];
  cap: string | null;
  join: string | null;
}


export type ChartExSeriesStyleCarrier = Pick<
  ChartSeries,
  'chartexStyle' | 'lineHidden' | 'lineColor' | 'lineWidthEmu'
>;


/** Resolve a ChartEx mark's effective outline once for both plot and legend.
 * Direct CT_Series formatting wins over the linked Chart Style role. `NoStyle`
 * is absence of decoration (and may expose a family semantic outline), while
 * an explicit noFill suppresses the outline. */
export function resolveChartExSeriesLineStyle(
  chart: ChartModel,
  linkedStyle: ChartExStyle | null | undefined,
  series: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  index: number,
  count: number,
  fallbackColor: string,
  options: { linkedNoStyleFallback?: boolean } = {},
): ResolvedChartExLineStyle {
  void count;
  const local = series?.chartexStyle;
  const role = linkedStyle === chart.chartexSeriesLineStyle
    ? 'seriesLine'
    : linkedStyle === chart.chartexDataPointLineStyle
      ? 'dataPointLine'
      : linkedStyle === chart.chartexDataPointMarkerStyle
        ? 'dataPointMarker'
        : (Object.entries(chart.chartStyleRoles ?? {}).find(
            ([, style]) => style === linkedStyle,
          )?.[0] as ChartStyleRole | undefined);
  const rawLinkedStyle = role ? rawLinkedChartStyleRole(chart, role) : linkedStyle;
  const localPaint = chartStyleDirectLineDecision(local, rawLinkedStyle, index);
  const legacyNoLine = series?.lineHidden === true
    ? chartStyleDirectNoLineDecision(rawLinkedStyle) : undefined;
  const legacyPaintAuthored = series?.lineColor != null
    || legacyNoLine !== undefined;
  const linkedPaint = chartStyleLineDecision(linkedStyle, index);
  const selectedPaint = localPaint !== undefined
    ? localPaint
    : legacyPaintAuthored
      ? legacyNoLine !== undefined
        ? legacyNoLine
        : { fillType: 'solid' as const, color: series?.lineColor ?? fallbackColor }
      : linkedPaint !== undefined
        ? linkedPaint
        : undefined;
  // NoStyle applies only to paint. It deliberately exposes the semantic
  // family fallback while local geometry beside the NoStyle reference remains
  // part of the effective outline.
  const linkedNoStyleSuppressesSemanticPaint = selectedPaint === undefined
    && linkedStyle?.lineNoStyle === true
    && options.linkedNoStyleFallback !== true;
  return {
    visible: selectedPaint !== null && !linkedNoStyleSuppressesSemanticPaint,
    color: selectedPaint?.fillType === 'solid'
      ? selectedPaint.color
      : fallbackColor,
    paint: selectedPaint?.fillType === 'solid' ? undefined : selectedPaint,
    widthEmu: local?.lineWidthEmu
      ?? series?.lineWidthEmu ?? linkedStyle?.lineWidthEmu ?? null,
    dash: local?.lineCustomDash != null
      ? null : local?.lineDash ?? linkedStyle?.lineDash ?? null,
    customDash: local?.lineCustomDash ?? linkedStyle?.lineCustomDash ?? null,
    cap: local?.lineCap ?? linkedStyle?.lineCap ?? null,
    join: local?.lineJoin ?? linkedStyle?.lineJoin ?? null,
  };
}


export function applyResolvedChartExLineStyle(
  ctx: CanvasRenderingContext2D,
  line: ResolvedChartExLineStyle,
  ptToPx: number,
): boolean {
  if (!line.visible) return false;
  ctx.strokeStyle = line.color.startsWith('#') ? line.color : `#${line.color}`;
  ctx.lineWidth = line.widthEmu != null
    ? axisLineWidthPx(line.widthEmu, ptToPx)
    : 1;
  ctx.setLineDash(dashPatternForLine(line.customDash, line.dash, ctx.lineWidth));
  ctx.lineCap = line.cap === 'rnd' ? 'round' : line.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = line.join === 'round' || line.join === 'bevel' ? line.join : 'miter';
  return true;
}


/** Apply CT_Series local shape properties before the linked Chart Style.
 *  [MS-ODRAWXML] 2.24.3.77 makes `<cx:series><cx:spPr>` the series' own
 *  OfficeArt formatting, so an authored line (including `noFill`) overrides
 *  the default data-point recipe instead of being merged underneath it. */
export function applyChartExSeriesLineStyle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  series: Pick<ChartSeries, 'chartexStyle' | 'lineHidden' | 'lineColor' | 'lineWidthEmu'> | null | undefined,
  index: number,
  count: number,
  fallbackColor: string,
  ptToPx: number,
  options: { linkedNoStyleFallback?: boolean } = {},
): boolean {
  return applyResolvedChartExLineStyle(
    ctx,
    resolveChartExSeriesLineStyle(
      chart, style, series, index, count, fallbackColor, options,
    ),
    ptToPx,
  );
}


/** Build a synthetic legend series from the same resolved line contract used
 * by plot paint. `chartexStyle` carries dash/cap/join through the generic
 * legend pipeline without widening the public ChartSeries surface. */
export function chartExLegendSeries(
  chart: ChartModel,
  name: string,
  series: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  linkedStyle: ChartExStyle | null | undefined,
  index: number,
  count: number,
  fillColor: string,
  semanticNoStyleFallback = false,
  inheritPlotOutline = true,
): ChartSeries {
  const line = resolveChartExSeriesLineStyle(
    chart,
    linkedStyle,
    series,
    index,
    count,
    fillColor,
    { linkedNoStyleFallback: semanticNoStyleFallback },
  );
  return {
    name,
    values: [],
    color: fillColor.replace(/^#/, ''),
    lineHidden: !inheritPlotOutline || !line.visible,
    lineColor: inheritPlotOutline && line.visible ? line.color.replace(/^#/, '') : null,
    lineWidthEmu: inheritPlotOutline ? line.widthEmu : null,
    chartexStyle: {
      linePaints: inheritPlotOutline && line.paint !== undefined ? [line.paint] : null,
      linePaintAuthored: inheritPlotOutline && line.paint !== undefined ? true : null,
      lineDash: inheritPlotOutline ? line.dash : null,
      lineCustomDash: inheritPlotOutline ? line.customDash : null,
      lineCap: inheritPlotOutline ? line.cap : null,
      lineJoin: inheritPlotOutline ? line.join : null,
    },
  };
}
