// Shared chart layout, style, axes, legends, and painting helpers.
import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartDecorationLineStyle,
  ChartDisplayUnits,
  ChartExElementStyle,
  ChartLabelBox,
  ChartLegendEntryOverride,
  ChartManualLayout,
  ChartModel,
  ChartRect,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartStockUpDownBarStyle,
  ChartStyleRole,
  ChartTextBox,
  ChartTextRun,
  ChartTrendline,
  SecondaryValueAxis,
} from '../../types/chart';
import type { Fill } from '../../types/common';
import {
  classicDataPointFillDecision,
  classicDataPointLineStyle,
} from '../classic-data-point-style.js';

import {
  chartImageFillPaintWorkUpperBound,
  paintChartImageFill,
  type ChartImageLookup,
} from '../image-fill.js';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import {
  chartLabelBoxHasVisiblePaint,
  effectiveChartLabelBoxFill,
  mergeChartLabelBoxes,
  paintChartLabelBox,
} from '../label-box.js';

import {
  anchoredDataLabelPoint,
  dataLabelCanvasTextAlign,
  dataLabelIsDeleted,
  dataLabelInsets,
  effectiveDataLabelTextStyle,
  fitStyledDataLabelLines,
  rotatedDataLabelSize,
  transformDataLabelText,
  type DataLabelTextStyle,
} from '../data-label-style.js';
import {
  bubblePointIsThreeD,
  classicDataLabelPointIsPainted,
  classicMarkerPointIsPainted,
  chartDataTableFamilyIsPainted,
  dataLabelLegendKeyCount,
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markersSuppressedByChartStyle,
  markerFillColorFor,
  markerFillPaintFor,
  markerPaintComponents,
  markerSymbolConsumesFill,
  pointHasMarkerDetail,
  seriesMarkerFillColor,
  seriesMarkerFillPaint,
  seriesLegendMarkerIsVisible,
  seriesHasMarkerDetail,
  visibleBubbleSize,
} from '../marker-style.js';
import {
  chartVariesColorsByPoint,
  deletedLegendEntryIndices,
  legendEntryGlobalIndex,
  legendEntryIsVisible,
  legendEntryRanges,
  legendIsCategoryDriven,
} from '../legend-entry-plan.js';
import {
  cartesianTitleBand,
  chartLegendReserve,
  packLegendRows,
  axisTitleFontPx,
  axisTitleRotationRad,
  chartTextFontSizePx,
  categoryTickLabelGapPx,
  axisTitleMargin,
  resolveManualLayoutRect,
  type ChartLegendReserve,
  type ChartAxisTitleSide,
  type ChartTitleBand,
} from '../layout.js';
import {
  automaticPercentMajorUnit,
  planNumericValueAxis,
  fitTrendline,
  linearTrendlineStats,
} from '../axis-scale.js';
import {
  axisLineWidthPx,
  resolveAxisLine,
  resolveGridline,
  isCrossBetween,
} from '../axis-style.js';
import {
  formatChartVal,
  formatChartValWithCode,
  formatCategoryLabel,
  formatLocalizedExcelShortDate,
} from '../chart-number-format.js';
import { elideToWidth } from '../text-elide.js';
import { categoryLabelAnchorFraction, categoryLabelOffsetPx } from '../category-spacing.js';

import { computeBoxWhiskerStats } from '../box-whisker.js';
import { planDateCategoryAxis } from '../date-axis.js';
import {
  classicCanvasPointFamilyIsPainted,
  MAX_CANVAS_CHART_POINTS,
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
} from '../resource-limits.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from '../plot-groups.js';
import {
  THREE_D_MAX_SHAPE_FACES_PER_DATUM,
  type ChartThreeDRenderer,
} from '../three-d-contract.js';

import {
  paintRichDataLabelBlock,
  resolveRichDataLabelBlock,
  type RichDataLabelOptions,
} from '../rich-data-label.js';
import { effectiveDataLabelText } from '../data-label-content.js';

import {
  chartDataPointStyleRole,
  chartPlotGroupForSeries,
  chartSeriesSourceIndex,
  chartSeriesVariesByPoint,
  chartStyleDashChoice,
  effectiveChartStyleRole,
  rawLinkedChartStyleRole,
} from '../effective-style.js';

import { placeTrendlineLabel } from '../trendline-label.js';
import { paintLegendFrame } from '../legend-frame.js';

import {
  chartStyleColor,
  chartStyleDirectFillDecision,
  chartStyleDirectLineDecision,
  chartStyleDirectNoFillDecision,
  chartStyleDirectNoLineDecision,
  chartStyleFillCascade,
  chartStyleFillDecision,
  chartStyleFontColor,
  chartStyleLineCascade,
  chartStyleLineDecision,
} from '../style-paint.js';
import { hasFilteredScatterAutomaticPointStyle } from '../source-visibility.js';
import {
  boundDataLabelText,
  resolveDataLabelPlacement,
  type DataLabelAnchor,
  type DataLabelRect,
} from '../data-label-layout.js';
import { resolveFill } from '../../shape/paint.js';
import { drawingmlLineDashArray, pptxPresetDashArray } from '../../draw/dash.js';

import {
  DEFAULT_TEXT_INSET_LR_EMU,
  DEFAULT_TEXT_INSET_TB_EMU,
  EMU_PER_PT,
  PT_TO_PX,
} from '../../units.js';

// ─── Palette + helpers ──────────────────────────────────────────────────────

export const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

/** Office 2013+ ChartEx fallback accents when no theme/colors sidecar resolves. */
export const CHARTEX_DEFAULT_PALETTE = [
  '5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47',
] as const;

export function chartColor(idx: number, series?: { color?: string | null } | null): string {
  if (series?.color) return `#${series.color}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

/** Index point-scoped OOXML overrides once while preserving first-in-document
 * precedence for duplicate indexes. */
export function indexPointOverrides<T extends { idx: number }>(
  values: readonly T[] | null | undefined,
): ReadonlyMap<number, T> {
  const indexed = new Map<number, T>();
  for (const value of values ?? []) {
    if (!indexed.has(value.idx)) indexed.set(value.idx, value);
  }
  return indexed;
}

export function pieSliceColor(
  idx: number,
  series: ChartSeries,
  varyColors = true,
  seriesIndex = idx,
): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  // When varyColors is off (or the parser deliberately suppresses automatic
  // point colours for a series noFill), every unspecified slice inherits the
  // series fill. Falling straight to the built-in palette would revive a
  // noFill series and recolour a single-colour pie point by point.
  if (series.color === '00000000') return '#00000000';
  return varyColors
    ? `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`
    : chartColor(seriesIndex, series);
}

/** Select the numeric/linked style palette domain used by a pie-family mark.
 * Excel replays a point palette in every ring while `varyColors` is effective,
 * but colors each complete series/ring when it is explicitly disabled. */
export function piePointStyleIndex(
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  pointIndex: number,
): number {
  return chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex
    : chartExSeriesFormatIndex(series, seriesIndex);
}

/** Apply one classic series/point outline with component-wise DrawingML
 * precedence. Series-local paint is direct formatting, while the linked or
 * numeric Chart Style role supplies only missing paint/geometry before the
 * family semantic fallback. */
export function applyClassicStyleLine(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  role: 'dataPoint' | 'dataPointLine',
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  styleIndex: number,
  fallbackColor: string,
  fallbackWidthPx: number,
  ptToPx: number,
  bounds: ChartRect,
  shapeRotationDeg: number,
  semanticVisible = true,
  resetSolidDash = true,
): boolean {
  const line = classicDataPointLineStyle(chart, role, series, point, styleIndex);
  let { paint } = line;
  if (paint === undefined && !semanticVisible) return false;
  if (paint === undefined) {
    paint = { fillType: 'solid', color: fallbackColor.replace(/^#/, '') };
  }
  if (paint === null) return false;

  const stroke = paint.fillType === 'solid'
    ? (paint.color.startsWith('#') ? paint.color : `#${paint.color}`)
    : resolveFill(
        paint, ctx, bounds.x, bounds.y, bounds.w, bounds.h, shapeRotationDeg,
      );
  if (!stroke) return false;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = line.widthEmu != null
    ? axisLineWidthPx(line.widthEmu, ptToPx) : fallbackWidthPx;
  const lineDash = drawingmlLineDashArray(
    line.customDash, line.dash, ctx.lineWidth,
  );
  const currentLineDash = typeof ctx.getLineDash === 'function'
    ? ctx.getLineDash() ?? []
    : [];
  if (resetSolidDash || lineDash.length > 0 || currentLineDash.length > 0) {
    ctx.setLineDash(lineDash);
  }
  ctx.lineCap = line.cap === 'rnd' ? 'round' : line.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = line.join === 'round' || line.join === 'bevel' ? line.join : 'miter';
  return true;
}

export interface IndexedLinePoint {
  x: number;
  y: number;
  index: number;
}

/** Paint a point-varying line as independently styled destination segments.
 * Excel applies §21.2.2.227 to a lone line/scatter/radar series by changing
 * both its marker and the segment that arrives at that point. A multi-series
 * group remains series-coloured. Smooth curves keep the same Catmull-Rom
 * geometry; only the segment paint/effect owner changes at each endpoint. */
export function paintClassicVaryingLineSegments(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  runs: IndexedLinePoint[][],
  smooth: boolean,
  closed: boolean,
  fallbackColor: string,
  fallbackWidthPx: number,
  ptToPx: number,
  bounds: ChartRect,
  shapeRotationDeg: number,
  semanticVisible = true,
): void {
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const pointOverrides = indexPointOverrides(series.dataPointOverrides);
  const fallbackRole = chartDataPointStyleRole(chart, 'dataPointLine', seriesIndex);
  const paintSegment = (
    from: IndexedLinePoint,
    to: IndexedLinePoint,
    run: IndexedLinePoint[],
    segmentIndex: number,
    closesRun = false,
  ): void => {
    const point = pointOverrides.get(to.index);
    const paint = (target: CanvasRenderingContext2D): void => {
      target.save();
      if (applyClassicStyleLine(
        target, chart, 'dataPointLine', series, point, to.index,
        fallbackColor, fallbackWidthPx, ptToPx, bounds, shapeRotationDeg,
        semanticVisible,
      )) {
        target.beginPath();
        target.moveTo(from.x, from.y);
        if (smooth && !closesRun) {
          const p0 = run[segmentIndex - 1] ?? from;
          const p3 = run[segmentIndex + 2] ?? to;
          target.bezierCurveTo(
            from.x + (to.x - p0.x) / 6,
            from.y + (to.y - p0.y) / 6,
            to.x - (p3.x - from.x) / 6,
            to.y - (p3.y - from.y) / 6,
            to.x,
            to.y,
          );
        } else {
          target.lineTo(to.x, to.y);
        }
        target.stroke();
      }
      target.restore();
    };
    paintChartStyleEffects(
      ctx,
      chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle),
      fallbackRole,
      to.index,
      bounds,
      ptToPx,
      paint,
      to.index,
    );
  };

  for (const run of runs) {
    for (let index = 0; index + 1 < run.length; index++) {
      paintSegment(run[index], run[index + 1], run, index);
    }
    if (closed && run.length > 1) {
      paintSegment(run[run.length - 1], run[0], run, run.length - 1, true);
    }
  }
}

export function paintClassicPiePointOutline(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  styleIndex: number,
  fallbackColor: string,
  ptToPx: number,
  bounds: ChartRect,
  shapeRotationDeg: number,
): void {
  const pointStyle = point?.chartexStyle;
  const sourceSeriesIndex = chartSeriesSourceIndex(chart, series);
  const linkedPointStyle = chartDataPointStyleRole(
    chart, 'dataPoint', sourceSeriesIndex >= 0 ? sourceSeriesIndex : styleIndex,
  );
  const seriesStyle = series.chartexStyle;
  const pointOwnsPaint = chartStyleLineDecision(pointStyle, point?.idx ?? styleIndex) !== undefined
    || point?.lineHidden === true || point?.lineColor != null;
  const seriesOwnsPaint = chartStyleLineDecision(seriesStyle, styleIndex) !== undefined
    || series.lineHidden === true || series.lineColor != null;
  const linkedOwnsPaint = chartStyleLineDecision(linkedPointStyle, styleIndex) !== undefined;
  // A programmatically constructed legacy ChartModel may carry neither a
  // numeric/linked style role nor direct line formatting. Preserve that public
  // contract's historical no-outline semantic; parsed OOXML always supplies
  // the numeric dataPoint role, including its explicit Table 5 No Line cases.
  if (!pointOwnsPaint && !seriesOwnsPaint && !linkedOwnsPaint) return;
  ctx.save();
  if (applyClassicStyleLine(
    ctx,
    chart,
    'dataPoint',
    series,
    point,
    styleIndex,
    fallbackColor,
    Math.max(.5, ptToPx * .75),
    ptToPx,
    bounds,
    shapeRotationDeg,
    false,
  )) ctx.stroke();
  ctx.restore();
}

// ─── Font-face resolution (CH10) ─────────────────────────────────────────────
// Chart text elements draw with, in priority order: the element's own
// `<a:latin typeface>` (from its `<c:txPr>`), else the theme font-scheme face
// (heading `majorFont` for titles, body `minorFont` for tick labels / data
// labels / legend, ECMA-376 §20.1.4.2), else the built-in `sans-serif`. When
// neither a per-element face nor a theme face is present the result is exactly
// `sans-serif`, so charts that specify no faces render byte-identically to
// before. A resolved face is quoted and given the same Calibri/Arial fallback
// chain as the chart title, so a font the platform lacks still degrades to a
// sans-serif rather than a serif default.
export type ChartFontRole = 'major' | 'minor';

/** Resolve a DrawingML theme font-scheme reference (`+mj-lt` / `+mn-lt` etc.,
 *  ECMA-376 §20.1.4.1.16) to the concrete theme face. `+mj-*` = heading
 *  (majorFont), `+mn-*` = body (minorFont); the axis suffix (`-lt`/`-ea`/`-cs`)
 *  is ignored here — chart text is Latin. A non-reference face passes through.
 *  Returns null when a reference can't be resolved (theme not threaded). */
export function resolveThemeFontRef(chart: ChartModel, face: string | null | undefined): string | null | undefined {
  if (!face) return face;
  if (face.startsWith('+mj')) return chart.themeMajorFontLatin ?? null;
  if (face.startsWith('+mn')) return chart.themeMinorFontLatin ?? null;
  return face;
}

export function chartFontFamily(
  chart: ChartModel,
  elementFace: string | null | undefined,
  role: ChartFontRole,
): string {
  const themeFace = role === 'major' ? chart.themeMajorFontLatin : chart.themeMinorFontLatin;
  const face = resolveThemeFontRef(chart, elementFace) ?? themeFace;
  return face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif';
}

export function chartFontCss(
  fontSizePx: number,
  fontFamily: string,
  bold = false,
  italic = false,
): string {
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
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

/** Draw an axis title at an explicit anchor in the outer gutter band. The
 *  side-based compatibility rotation is resolved in one place, with authored
 *  DrawingML body orientation remaining authoritative. */
export function drawAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchorX: number, anchorY: number,
  side: ChartAxisTitleSide,
  fontSizePx: number,
  bold: boolean,
  italic: boolean,
  color: string,
  // Available run length along the axis (plot width for the bottom cat title,
  // plot height for the rotated val title). Titles longer than the axis are
  // elided with an ellipsis rather than hard-cut at a fixed char count.
  maxPx: number,
  // Resolved CSS font-family (element face ?? theme heading ?? sans-serif).
  fontFamily = 'sans-serif',
  authoredRotation?: number | null,
  authoredVerticalMode?: ChartModel['catAxisTitleVerticalMode'],
  manualLayout?: ChartManualLayout | null,
  chartRect?: ChartRect,
  box?: ChartLabelBox | null,
  ptToPx = 1,
): void {
  ctx.save();
  ctx.font = chartFontCss(fontSizePx, fontFamily, bold, italic);
  ctx.fillStyle = color;
  // Automatic titles stay bounded to the axis run. An authored title layout
  // is authoritative and keeps its complete text rather than being elided by
  // the automatic plot-width estimate.
  const label = manualLayout ? text : elideToWidth(ctx, text, maxPx);
  const rotation = axisTitleRotationRad(side, authoredRotation, authoredVerticalMode);
  let resolvedAnchorX = anchorX;
  let resolvedAnchorY = anchorY;
  if (manualLayout && chartRect) {
    const textWidth = ctx.measureText(label).width;
    // CT_Title manual-layout x/y position the title's axis-aligned box after
    // DrawingML rotation. A vertical title therefore has a box approximately
    // one font line wide and one text run tall; using the unrotated dimensions
    // shifts it into the tick-label/plot bands by half the text length.
    const cos = Math.abs(Math.cos(rotation));
    const sin = Math.abs(Math.sin(rotation));
    const fittedWidth = textWidth * cos + fontSizePx * sin;
    const fittedHeight = textWidth * sin + fontSizePx * cos;
    const automatic = {
      x: anchorX - fittedWidth / 2,
      y: anchorY - fittedHeight / 2,
      w: fittedWidth,
      h: fittedHeight,
    };
    // CT_Title manual layout positions the title box, while Office keeps the
    // box fitted to its text. Match the existing chart-title rule: x/y win,
    // authored w/h do not stretch or shrink the text box.
    const resolved = resolveManualLayoutRect(
      { ...manualLayout, w: undefined, h: undefined },
      chartRect,
      automatic,
    );
    if (resolved) {
      resolvedAnchorX = resolved.x + resolved.w / 2;
      resolvedAnchorY = resolved.y + resolved.h / 2;
    }
  }
  ctx.translate(resolvedAnchorX, resolvedAnchorY);
  if (rotation !== 0) ctx.rotate(rotation);
  const textWidth = ctx.measureText(label).width;
  paintChartLabelBox(ctx, box, {
    x: -textWidth / 2,
    y: -fontSizePx / 2,
    w: textWidth,
    h: fontSizePx,
  }, ptToPx);
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

/** Resolve the per-axis title color string for `drawAxisTitle`. Returns
 *  '#rrggbb' when the XML supplied a srgb color, else the legacy '#555'. */
export function axisTitleColor(hex: string | null | undefined): string {
  return hex ? `#${hex}` : '#555';
}

/** Draw both axis titles for a cartesian chart (bar/line/area/scatter),
 *  anchored in the reserved outer gutter bands so they sit OUTSIDE the tick
 *  labels. `catTitlePx`/`valTitlePx` are the title font sizes the caller used
 *  to size `catTitleH`/`valTitleW`; the anchor centers each title within its
 *  band. Column/line/area/scatter use cat-bottom + val-left. Horizontal bars
 *  use cat-left + val-bottom because their value axis runs horizontally.
 *  Bold and italic are independent DrawingML character properties. The parser
 *  resolves authored/inherited OOXML values, including the regular-weight
 *  DrawingML base fallback. A hand-built public model that leaves bold unset
 *  retains the renderer's established bold compatibility fallback. The
 *  separate 10pt size fallback is the product policy in #1228. */
export function drawAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  legLeftW: number, legBottomH: number,
  catTitlePx: number, valTitlePx: number,
  horizontalValueAxis = false,
): void {
  const drawPrimaryTitle = (
    text: string,
    side: ChartAxisTitleSide,
    fontSizePx: number,
    bold: boolean,
    italic: boolean,
    color: string,
    fontFamily: string,
    authoredRotation: number | null | undefined,
    authoredVerticalMode: ChartModel['catAxisTitleVerticalMode'],
    manualLayout: ChartManualLayout | null | undefined,
    directStyle: ChartExStyle | null | undefined,
  ): void => {
    const box = effectiveLinkedLabelBox(
      chart,
      directStyle ? { style: directStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
      true,
    );
    if (side === 'left') {
      drawAxisTitle(
        ctx, text,
        x + legLeftW + axisTitleMargin(w) + fontSizePx / 2,
        py0 + ph / 2,
        side, fontSizePx, bold, italic, color, ph, fontFamily, authoredRotation,
        authoredVerticalMode, manualLayout, { x, y, w, h }, box,
        fontSizePx / 10,
      );
      return;
    }
    drawAxisTitle(
      ctx, text,
      px0 + pw / 2,
      y + h - legBottomH - axisTitleMargin(h) - fontSizePx / 2,
      side, fontSizePx, bold, italic, color, pw, fontFamily, authoredRotation,
      authoredVerticalMode, manualLayout, { x, y, w, h }, box,
      fontSizePx / 10,
    );
  };
  if (chart.valAxisTitle) {
    drawPrimaryTitle(
      chart.valAxisTitle, horizontalValueAxis ? 'horizontal' : 'left',
      valTitlePx, chart.valAxisTitleFontBold ?? true, chart.valAxisTitleFontItalic ?? false,
      axisTitleColor(chart.valAxisTitleFontColor),
      chartFontFamily(chart, chart.valAxisTitleFontFace, 'major'), chart.valAxisTitleRotation,
      chart.valAxisTitleVerticalMode,
      chart.valAxisTitleManualLayout,
      chart.valAxisTitleStyle,
    );
  }
  if (chart.catAxisTitle) {
    drawPrimaryTitle(
      chart.catAxisTitle, horizontalValueAxis ? 'left' : 'horizontal',
      catTitlePx, chart.catAxisTitleFontBold ?? true, chart.catAxisTitleFontItalic ?? false,
      axisTitleColor(chart.catAxisTitleFontColor),
      chartFontFamily(chart, chart.catAxisTitleFontFace, 'major'), chart.catAxisTitleRotation,
      chart.catAxisTitleVerticalMode,
      chart.catAxisTitleManualLayout,
      chart.catAxisTitleStyle,
    );
  }
}

export type ChartDataTableLayout = {
  fontPx: number;
  lineHeight: number;
  headerLines: string[][];
  headerHeight: number;
  rowHeight: number;
  totalHeight: number;
};

/** Office only paints a classic chart data table for category-axis families.
 * CT_DTable is syntactically allowed under plotArea, but an authored table on
 * an XY scatter plot is ignored (confirmed with an Office vector boundary).
 * Keeping this gate beside the shared layout prevents family renderers from
 * inventing different applicability rules. */
export function chartHasDataTable(chart: ChartModel): boolean {
  return chart.dataTable != null && chartDataTableFamilyIsPainted(chart.chartType);
}

export function chartDataTableRows(chart: ChartModel): Array<{ series: ChartSeries; sourceIndex: number }> {
  const horizontal = chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
  const rows = chart.series.map((series, sourceIndex) => ({ series, sourceIndex }));
  return horizontal ? rows.reverse() : rows;
}

/** Minimum data-table band reserved before the final plot width is known. The
 * header starts as one line; after `computeChartFrame` the measured category
 * cell width may add wrapped lines and the caller shrinks the plot by exactly
 * that measured delta. */
export function chartDataTableBaseHeight(chart: ChartModel, ptToPx: number): number {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return 0;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const lineHeight = Math.max(1, fontPx * 1.2);
  const rowHeight = lineHeight + 4 * ptToPx;
  return (chart.series.length + 1) * rowHeight;
}

export function chartDataTableHeaderWidth(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  ptToPx: number,
): number {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return 0;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  ctx.save();
  ctx.font = chartFontCss(fontPx, face, table.fontBold ?? false, table.fontItalic ?? false);
  const nameWidth = chart.series.reduce(
    (width, series) => Math.max(width, ctx.measureText(series.name).width),
    0,
  );
  ctx.restore();
  const keyWidth = table.showKeys ? Math.max(12 * ptToPx, fontPx * 1.7) : 0;
  const keyGap = table.showKeys ? 4 * ptToPx : 0;
  return nameWidth + keyWidth + keyGap + 6 * ptToPx;
}

export function measureChartDataTable(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  categoryWidth: number,
  ptToPx: number,
): ChartDataTableLayout | null {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return null;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const lineHeight = Math.max(1, fontPx * 1.2);
  const rowHeight = lineHeight + 4 * ptToPx;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  ctx.save();
  ctx.font = chartFontCss(fontPx, face, table.fontBold ?? false, table.fontItalic ?? false);
  const categoryFormat = chart.series.find(series => series.catFormatCode)?.catFormatCode
    ?? chart.catAxisFormatCode;
  const categoryBuiltinId = chart.series
    .find(series => series.catFormatBuiltinId != null)?.catFormatBuiltinId;
  const headerLines = chartCategories(chart).map(category => {
    const numeric = category.trim() === '' ? Number.NaN : Number(category);
    const label = categoryBuiltinId === 14 && Number.isFinite(numeric)
      ? formatLocalizedExcelShortDate(numeric, chart.date1904)
      : formatCategoryLabel(category, categoryFormat, chart.date1904);
    return wrapMeasuredText(ctx, label, Math.max(1, categoryWidth - 4 * ptToPx));
  });
  ctx.restore();
  const maxHeaderLines = Math.max(1, ...headerLines.map(lines => lines.length));
  const headerHeight = maxHeaderLines * lineHeight + 4 * ptToPx;
  return {
    fontPx,
    lineHeight,
    headerLines,
    headerHeight,
    rowHeight,
    totalHeight: headerHeight + chartDataTableRows(chart).length * rowHeight,
  };
}

/** Draw `CT_DTable` as a measured chart foreground band. Category columns are
 * aligned to the plot's category span; the leading key/name column occupies
 * the already-reserved value-axis gutter. Border switches are honored
 * independently, as authored by the four CT_Boolean children. */
export function drawChartDataTable(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  layout: ChartDataTableLayout | null,
  plotX: number,
  tableY: number,
  plotWidth: number,
  chartLeft: number,
  ptToPx: number,
): void {
  const table = chart.dataTable;
  if (!table || !layout) return;
  const categories = chartCategories(chart);
  if (categories.length === 0) return;
  const categoryWidth = plotWidth / categories.length;
  const face = chartFontFamily(chart, table.fontFace, 'minor');
  const font = chartFontCss(
    layout.fontPx, face, table.fontBold ?? false, table.fontItalic ?? false,
  );
  const keyWidth = table.showKeys ? Math.max(12 * ptToPx, layout.fontPx * 1.7) : 0;
  const keyGap = table.showKeys ? 4 * ptToPx : 0;
  ctx.save();
  ctx.font = font;
  const longestName = chart.series.reduce(
    (width, series) => Math.max(width, ctx.measureText(series.name).width),
    0,
  );
  const desiredHeaderWidth = longestName + keyWidth + keyGap + 6 * ptToPx;
  const headerWidth = Math.min(Math.max(0, plotX - chartLeft), desiredHeaderWidth);
  const tableX = plotX - headerWidth;
  const tableWidth = headerWidth + plotWidth;
  const tableBottom = tableY + layout.totalHeight;
  const tableRows = chartDataTableRows(chart);
  const keyEntries = buildLegendEntries(
    chart.series,
    chart.chartType,
    chart.scatterStyle,
    false,
    chart.categories,
    [],
    true,
    [],
    chart.radarStyle,
    chart,
  );
  const keyEntryRanges = legendEntryRanges(chart);
  // A direct solid dTable fill belongs to each generated body-text box. That
  // semantic is independent of the owning chart family, plot-group count,
  // manual plot layout, line wrapping, and sparse values. Unsupported fill
  // recipes remain present in the model but do not masquerade as a solid.
  const bodyFillColor = table.fillColor ?? null;
  ctx.beginPath();
  ctx.rect(tableX, tableY, tableWidth, layout.totalHeight);
  ctx.clip();
  const fontColor = table.fontHidden === true
    || (table.fontPaintAuthored === true && table.fontColor == null)
    ? 'transparent'
    : table.fontColor ? `#${table.fontColor}` : '#000000';
  const drawBodyText = (text: string, centerX: number, centerY: number): void => {
    // Desktop Excel scopes a direct dTable/spPr fill to the generated body
    // text boxes. It does not fill the table frame or the leading series-name
    // cells. The text layout box is the measured advance by the measured line
    // height, so this remains tied to authored typography rather than a cell-
    // or sample-specific inset.
    if (bodyFillColor && text !== '') {
      const width = ctx.measureText(text).width;
      ctx.fillStyle = `#${bodyFillColor}`;
      ctx.fillRect(
        centerX - width / 2,
        centerY - layout.lineHeight / 2,
        width,
        layout.lineHeight,
      );
    }
    ctx.fillStyle = fontColor;
    ctx.textAlign = 'center';
    ctx.fillText(text, centerX, centerY);
  };
  ctx.fillStyle = fontColor;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
    const centerX = plotX + (categoryIndex + 0.5) * categoryWidth;
    const lines = layout.headerLines[categoryIndex] ?? [''];
    const textBlockHeight = lines.length * layout.lineHeight;
    const firstY = tableY + (layout.headerHeight - textBlockHeight) / 2 + layout.lineHeight / 2;
    lines.forEach((line, lineIndex) => {
      drawBodyText(line, centerX, firstY + lineIndex * layout.lineHeight);
    });
  }

  for (let seriesIndex = 0; seriesIndex < tableRows.length; seriesIndex++) {
    const { series, sourceIndex } = tableRows[seriesIndex];
    const rowTop = tableY + layout.headerHeight + seriesIndex * layout.rowHeight;
    const rowCenter = rowTop + layout.rowHeight / 2;
    if (headerWidth > 0) {
      const textLeft = tableX + 3 * ptToPx + keyWidth + keyGap;
      ctx.textAlign = 'left';
      ctx.fillText(
        elideToWidth(ctx, series.name, Math.max(0, plotX - textLeft - 2 * ptToPx)),
        textLeft,
        rowCenter,
      );
      if (table.showKeys && keyWidth > 0) {
        const keyX = tableX + 3 * ptToPx;
        const keyEntryIndex = legendEntryGlobalIndex(keyEntryRanges, sourceIndex);
        const entry = keyEntryIndex == null ? undefined : keyEntries[keyEntryIndex];
        if (entry) {
          const keyHeight = Math.min(layout.fontPx, layout.rowHeight - 2 * ptToPx);
          drawLegendSwatch(
            ctx,
            entry.swatchStyle,
            entry.color,
            keyX,
            rowCenter - keyHeight / 2,
            keyWidth,
            keyHeight,
            entry.marker,
            entry.fillPaint,
            entry.outlinePaint,
            entry.outlineColor,
            entry.outlineWidthEmu,
            entry.outlineDash,
            entry.outlineCustomDash,
            entry.outlineCap,
            entry.outlineJoin,
            ptToPx,
            0,
            entry.directEffect,
            entry.fallbackEffect,
            entry.directEffectIndex,
            entry.fallbackEffectIndex,
          );
        }
        ctx.fillStyle = fontColor;
      }
    }
    for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
      const value = series.values[categoryIndex];
      const text = value == null ? '' : formatChartValWithCode(value, series.valFormatCode);
      drawBodyText(text, plotX + (categoryIndex + 0.5) * categoryWidth, rowCenter);
    }
  }

  if (table.lineHidden !== true
    && (table.linePaintAuthored !== true || table.lineColor != null)) {
    ctx.strokeStyle = table.lineColor ? `#${table.lineColor}` : '#808080';
    ctx.lineWidth = table.lineWidthEmu != null
      ? axisLineWidthPx(table.lineWidthEmu, ptToPx)
      : Math.max(0.5, ptToPx * 0.75);
    ctx.setLineDash(dashPatternForPreset(table.lineDash ?? undefined, ctx.lineWidth));
    if (table.showHorizontalBorder) {
      let lineY = tableY + layout.headerHeight;
      for (let row = 0; row < tableRows.length; row++) {
        ctx.beginPath(); ctx.moveTo(tableX, lineY); ctx.lineTo(tableX + tableWidth, lineY); ctx.stroke();
        lineY += layout.rowHeight;
      }
    }
    if (table.showVerticalBorder) {
      ctx.beginPath(); ctx.moveTo(plotX, tableY); ctx.lineTo(plotX, tableBottom); ctx.stroke();
      for (let category = 1; category < categories.length; category++) {
        const lineX = plotX + category * categoryWidth;
        ctx.beginPath(); ctx.moveTo(lineX, tableY); ctx.lineTo(lineX, tableBottom); ctx.stroke();
      }
    }
    if (table.showOutline) {
      const half = ctx.lineWidth / 2;
      ctx.strokeRect(
        tableX + half, tableY + half,
        Math.max(0, tableWidth - ctx.lineWidth),
        Math.max(0, layout.totalHeight - ctx.lineWidth),
      );
    }
  }
  ctx.restore();
}

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

/** Expand a cartesian plot's automatic top inset around an authored top legend.
 * A non-overlay manual legend still participates in chart layout: its authored
 * rectangle replaces the automatic legend rectangle, so the plot must start
 * below its actual bottom rather than below the shorter measured reserve. Keep
 * the existing automatic legend-to-plot clearance unchanged. */
export function manualTopLegendPlotInset(
  chart: ChartModel,
  legend: MeasuredLegendLayout | null,
  x: number,
  y: number,
  w: number,
  h: number,
  titleBandH: number,
  automaticInset: number,
): number {
  if (!legend
    || legend.side !== 't'
    || chart.legendOverlay === true
    || chart.legendManualLayout == null) return automaticInset;
  const defaultBox = {
    x: x + LEGEND_HORIZONTAL_INSET,
    y: y + titleBandH + 2,
    w: Math.max(0, w - LEGEND_HORIZONTAL_PADDING),
    h: legend.reserveH,
  };
  const manualBox = resolveManualLayoutRect(
    chart.legendManualLayout,
    { x, y, w, h },
    defaultBox,
  );
  if (!manualBox) return automaticInset;
  const automaticGap = Math.max(0, y + automaticInset - (defaultBox.y + defaultBox.h));
  return Math.max(
    automaticInset,
    manualBox.y + manualBox.h - y + automaticGap,
  );
}

export function drawAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  axis: 'val' | 'cat',
  anchorXOrY: number,
  perpendicular: number,
  color?: string,
  lineWidth?: number,
  // For a vertical value axis "outside" is to the LEFT (the axis sits on the
  // left). A secondary value axis sits on the RIGHT, where "outside" points
  // right — pass `opposite` to flip the out/in direction.
  opposite = false,
  lineHidden = false,
  level: 'major' | 'minor' = 'major',
  ptToPx = 1,
  dash?: string | null,
): void {
  // Axis shape properties style both the rule and its tick marks. An authored
  // `<a:ln><a:noFill/>` therefore suppresses the ticks too, while labels and
  // gridlines remain independently visible.
  if (lineHidden || mode === 'none' || !mode) return;
  // Office's vector output uses 6pt major ticks and 4pt minor ticks. Tick
  // length still scales mildly with an unusually thick authored axis rule.
  const len = axisTickLengthPx(level, lineWidth, ptToPx);
  // Office's 6pt/4pt observation is the complete cross-tick length, not the
  // length on each side of the axis. out/in use the full length on one side;
  // cross splits it evenly around the rule.
  const sideLen = mode === 'cross' ? len / 2 : len;
  const prevS = ctx.strokeStyle;
  const prevW = ctx.lineWidth;
  const prevDash = ctx.getLineDash?.() ?? [];
  ctx.strokeStyle = color ?? '#888';
  ctx.lineWidth = lineWidth ?? 1;
  ctx.setLineDash(dashPatternForPreset(dash ?? undefined, ctx.lineWidth));
  ctx.beginPath();
  if (axis === 'val') {
    // val axis is vertical (x = anchor, y varies). Ticks extend horizontally;
    // `outSign` points away from the plot (left for a left axis, right for a
    // right/secondary axis).
    const x0 = anchorXOrY;
    const y = perpendicular;
    const outSign = opposite ? 1 : -1;
    const outer = mode === 'out' || mode === 'cross' ? outSign * sideLen : 0;
    const inner = mode === 'in' || mode === 'cross' ? -outSign * sideLen : 0;
    ctx.moveTo(x0 + outer, y);
    ctx.lineTo(x0 + inner, y);
  } else {
    // cat axis is horizontal (y = anchor, x varies). Ticks extend vertically.
    const y0 = anchorXOrY;
    const xc = perpendicular;
    const outSign = opposite ? -1 : 1;
    const outer = mode === 'out' || mode === 'cross' ? outSign * sideLen : 0;
    const inner = mode === 'in' || mode === 'cross' ? -outSign * sideLen : 0;
    ctx.moveTo(xc, y0 + outer);
    ctx.lineTo(xc, y0 + inner);
  }
  ctx.stroke();
  ctx.strokeStyle = prevS;
  ctx.lineWidth = prevW;
  ctx.setLineDash(prevDash);
}

export function strokeAxisSegment(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  lineWidth: number,
  dash?: string | null,
): void {
  const previousDash = ctx.getLineDash?.() ?? [];
  const resolvedDash = dashPatternForPreset(dash ?? undefined, lineWidth);
  const dashChanged = resolvedDash.length !== previousDash.length
    || resolvedDash.some((value, index) => value !== previousDash[index]);
  ctx.strokeStyle = color;
  ctx.lineWidth = lineWidth;
  if (dashChanged) ctx.setLineDash(resolvedDash);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  if (dashChanged) ctx.setLineDash(previousDash);
}

export function axisTickLengthPx(
  level: 'major' | 'minor',
  lineWidth: number | undefined,
  ptToPx: number,
): number {
  const baseLen = (level === 'minor' ? 4 : 6) * ptToPx;
  return lineWidth ? Math.max(baseLen, lineWidth + 2 * ptToPx) : baseLen;
}

/** Distance an axis tick occupies outside the plot-side axis rule. */
export function axisTickOutwardExtentPx(
  mode: string | null | undefined,
  level: 'major' | 'minor',
  lineWidth: number | undefined,
  ptToPx: number,
): number {
  if (mode !== 'out' && mode !== 'cross') return 0;
  const length = axisTickLengthPx(level, lineWidth, ptToPx);
  return mode === 'cross' ? length / 2 : length;
}

/** Stroke one horizontal value-axis gridline spanning the plot width at `gy`.
 *  Extracted from the identical stroke the column-bar, line and area renderers
 *  each emitted inline. `isZero` is the caller's "this is the value-0 line"
 *  predicate (`si === 0` / `v === 0`). Callers set their own font/label
 *  BEFORE/AFTER this call, which is why those (drifted) parts stay at the call
 *  sites. Scatter is deliberately NOT a caller — it has no baseline special-case.
 *
 *  `grid` is the resolved `{ color, width }` from `resolveGridline` (the file's
 *  `<c:majorGridlines><c:spPr><a:ln>` or the faint `#e0e0e0`/0.5 px default).
 *  When the file supplies NO explicit gridline color (`grid.explicit === false`)
 *  the historical baseline emphasis applies: the value-0 line is a darker
 *  `#aaa` 1 px rule. When the file DOES pin a gridline color, PowerPoint strokes
 *  every major gridline in that one color/width uniformly, so the zero-line
 *  override is suppressed. Omitting `grid` reproduces the pre-CH-gridline
 *  default exactly (byte-stable for callers that haven't resolved a style). */
export function strokeValueGridlineH(
  ctx: CanvasRenderingContext2D,
  px0: number,
  pw: number,
  gy: number,
  isZero: boolean,
  grid?: { color: string; width: number; explicit: boolean; dash: number[] },
): void {
  if (grid && grid.explicit) {
    ctx.strokeStyle = grid.color;
    ctx.lineWidth = grid.width;
  } else {
    ctx.strokeStyle = isZero ? '#aaa' : grid?.color ?? '#e0e0e0';
    ctx.lineWidth = isZero ? 1 : grid?.width ?? 0.5;
  }
  const authoredDash = grid?.dash ?? [];
  const previousDash = authoredDash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
  if (authoredDash.length > 0) ctx.setLineDash(authoredDash);
  ctx.beginPath();
  ctx.moveTo(px0, gy);
  ctx.lineTo(px0 + pw, gy);
  ctx.stroke();
  if (authoredDash.length > 0) ctx.setLineDash(previousDash);
}

/** Resolve the value-axis MAJOR gridline stroke for `chart` at the current
 *  display scale. `explicit` is true when the file pinned any line property
 *  (color, width, or dash) under `<c:valAx><c:majorGridlines><c:spPr><a:ln>`;
 *  that flag tells
 *  `strokeValueGridlineH` to stroke every gridline in the resolved color
 *  uniformly (no `#aaa` zero-line emphasis), matching PowerPoint. With no
 *  explicit color the resolved `{ color: '#e0e0e0', width: 0.5 }` reproduces the
 *  historical faint hairline (byte-stable). */
export function valGridStroke(
  chart: ChartModel,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, ptToPx);
  return {
    color,
    width,
    explicit: chart.valAxisGridlineColor != null
      || chart.valAxisGridlineWidthEmu != null
      || chart.valAxisGridlineDash != null,
    dash: dashPatternForPreset(chart.valAxisGridlineDash ?? undefined, width),
  };
}

export function valMinorGridStroke(
  chart: ChartModel,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    chart.valAxisMinorGridlineColor,
    chart.valAxisMinorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: chart.valAxisMinorGridlineColor != null,
    dash: dashPatternForPreset(chart.valAxisMinorGridlineDash ?? undefined, width),
  };
}

export function secondaryMinorGridStroke(
  axis: SecondaryValueAxis,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    axis.minorGridlineColor,
    axis.minorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: axis.minorGridlineColor != null
      || axis.minorGridlineWidthEmu != null
      || axis.minorGridlineDash != null,
    dash: dashPatternForPreset(axis.minorGridlineDash ?? undefined, width),
  };
}

export function secondaryMajorGridStroke(
  axis: SecondaryValueAxis,
  ptToPx: number,
): { color: string; width: number; explicit: boolean; dash: number[] } {
  const { color, width } = resolveGridline(
    axis.majorGridlineColor,
    axis.majorGridlineWidthEmu,
    ptToPx,
  );
  return {
    color,
    width,
    explicit: axis.majorGridlineColor != null
      || axis.majorGridlineWidthEmu != null
      || axis.majorGridlineDash != null,
    dash: dashPatternForPreset(axis.majorGridlineDash ?? undefined, width),
  };
}

/** Whether to draw CATEGORY-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`,
 *  ECMA-376 §21.2.2.100). Office omits them by default, so only `true` turns
 *  them on (null/undefined/false ⇒ off, byte-stable). */
export function drawCatMajorGridlines(chart: ChartModel): boolean {
  return chart.catAxisMajorGridlines === true;
}

/** Resolve the CATEGORY-axis major gridline stroke, mirroring
 *  {@link valGridStroke}. `<c:catAx><c:majorGridlines><c:spPr><a:ln>` gives the
 *  color/width (`chart.catAxisGridlineColor`/`catAxisGridlineWidthEmu`); absent
 *  ⇒ the same faint `#e0e0e0`/0.5 px default as the value axis. Category
 *  gridlines have no zero-line emphasis (there is no "zero category"), so a
 *  single resolved stroke suffices. */
export function catGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
  const stroke = resolveGridline(chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, ptToPx);
  return {
    ...stroke,
    dash: dashPatternForPreset(chart.catAxisGridlineDash ?? undefined, stroke.width),
  };
}

export function catMinorGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
  const stroke = resolveGridline(
    chart.catAxisMinorGridlineColor,
    chart.catAxisMinorGridlineWidthEmu,
    ptToPx,
  );
  return {
    ...stroke,
    dash: dashPatternForPreset(chart.catAxisMinorGridlineDash ?? undefined, stroke.width),
  };
}

/** The plot-fraction positions (0..1 across the category extent) of the CATEGORY
 *  major gridlines / ticks for `n` categories. With crossBetween="between" (the
 *  bar/column default) they sit on the `n+1` band BOUNDARIES; under "midCat"
 *  they sit at the `n` category CENTERS. Shared by the category tick loop and
 *  the category-gridline pass so both stay aligned (§21.2.2.100/§21.2.2.32). */
export function catGridlineFractions(chart: ChartModel, n: number): number[] {
  if (n <= 0) return [];
  const onBoundary = isCrossBetween(chart);
  const fracs: number[] = [];
  const last = onBoundary ? n : n - 1;
  for (let ci = 0; ci <= last; ci++) {
    fracs.push(onBoundary ? ci / n : (n === 1 ? 0.5 : ci / (n - 1)));
  }
  return fracs;
}

/** True when the value axis is reversed (`<c:valAx><c:scaling><c:orientation
 *  val="maxMin">`, ECMA-376 §21.2.2.130). Absent/"minMax" ⇒ false (byte-stable). */
export function valAxisReversed(chart: ChartModel): boolean {
  return chart.valAxisOrientation === 'maxMin';
}

/** True when the category axis is reversed (`<c:catAx>…orientation="maxMin">`). */
export function catAxisReversed(chart: ChartModel): boolean {
  return chart.catAxisOrientation === 'maxMin';
}

/** Whether to draw value-axis MAJOR gridlines. Office writes `<c:majorGridlines>`
 *  on the value axis by default, so the historical always-on behavior maps to
 *  "draw unless the model explicitly says the element is absent". `undefined`
 *  (parser didn't model it) ⇒ true (byte-stable); `false` (axis present without
 *  the element) ⇒ off. */
export function drawValMajorGridlines(chart: ChartModel): boolean {
  return chart.valAxisMajorGridlines !== false;
}

/** A resolved value-axis plan: rounded bounds, the major gridline VALUES to
 *  stroke, an optional minor gridline VALUES list, and the value→fraction map
 *  (0 at the axis min end, 1 at the max end — before any pixel flip). Centralizes
 *  the CH6 major unit / logBase / orientation handling so every value-axis
 *  family shares one spec-faithful code path. With no CH6 fields set the plan is
 *  byte-identical to the old inline math: `step`/bounds from `valueAxisScale`,
 *  `majorLines = [min, min+step, … max]`, `frac(v) = (v-min)/(max-min)`. */
export interface ValueAxisPlan {
  min: number;
  max: number;
  step: number;
  majorLines: number[];
  minorLines: number[];
  minorTicks: number[];
  /** 0..1 position of `v` from the axis minimum toward the maximum (log-aware,
   *  orientation-aware). Renderers turn this into a pixel with
   *  `plotBottom - frac(v) * plotHeight` (vertical) — the reversal is already
   *  baked in, so callers keep their existing `- frac*len` form. */
  frac: (v: number) => number;
}

/** Convert an OOXML percent-axis value (stored as a 0..1 ratio) into the
 * renderer's percentStacked geometry space (0..100 percentage points). */
export function valueAxisUnitInRendererSpace(
  value: number | null | undefined,
  percentStacked: boolean,
): number | null | undefined {
  return value == null || !percentStacked ? value : value * 100;
}

/** Format a primary value-axis tick from the renderer's data space. For a
 * percentStacked chart the plotted values are percentage points, while the
 * axis numFmt still expects the OOXML ratio (0.5 → 50%). */
export function formatPrimaryValueAxisTick(
  chart: ChartModel,
  value: number,
  percentStacked: boolean,
): string {
  return formatChartValWithCode(
    (percentStacked ? value / 100 : value) / displayUnitDivisor(chart.valAxisDisplayUnits),
    percentStacked ? (chart.valAxisFormatCode ?? '0%') : chart.valAxisFormatCode,
    chart.date1904,
  );
}

export function displayUnitDivisor(units: ChartDisplayUnits | null | undefined): number {
  const divisor = units?.divisor;
  return divisor != null && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
}

export function formatAxisTickWithUnits(
  value: number,
  formatCode: string | null | undefined,
  date1904: boolean | undefined,
  units: ChartDisplayUnits | null | undefined,
): string {
  return formatChartValWithCode(value / displayUnitDivisor(units), formatCode, date1904);
}

export function automaticDisplayUnitLabel(units: ChartDisplayUnits): string {
  const names: Record<string, string> = {
    hundreds: 'Hundreds',
    thousands: 'Thousands',
    tenThousands: 'Ten Thousands',
    hundredThousands: 'Hundred Thousands',
    millions: 'Millions',
    tenMillions: 'Ten Millions',
    hundredMillions: 'Hundred Millions',
    billions: 'Billions',
    trillions: 'Trillions',
  };
  return units.builtInUnit ? (names[units.builtInUnit] ?? units.builtInUnit) : formatChartVal(units.divisor);
}

/** Paint the optional §21.2.2.46 display-unit labels after the family painter.
 * Their manual layout is chart-space (not plot-space), so this shared overlay
 * avoids a separate approximation in every chart family. */
export function drawChartDisplayUnitLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
): void {
  const entries = [
    { units: chart.valAxisDisplayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.08, fallbackY: rect.y + rect.h * 0.12, axis: { size: chart.valAxisFontSizeHpt, bold: chart.valAxisFontBold, italic: chart.valAxisFontItalic, color: chart.valAxisFontColor, paintAuthored: chart.valAxisFontPaintAuthored, face: chart.valAxisFontFace } },
    { units: chart.catAxisDisplayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.82, axis: { size: chart.catAxisFontSizeHpt, bold: chart.catAxisFontBold, italic: chart.catAxisFontItalic, color: chart.catAxisFontColor, paintAuthored: chart.catAxisFontPaintAuthored, face: chart.catAxisFontFace } },
    { units: chart.secondaryValAxis?.displayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.92, fallbackY: rect.y + rect.h * 0.12, axis: { size: chart.secondaryValAxis?.fontSizeHpt, bold: chart.secondaryValAxis?.fontBold, italic: chart.secondaryValAxis?.fontItalic, color: chart.secondaryValAxis?.fontColor, paintAuthored: chart.secondaryValAxis?.fontPaintAuthored, face: chart.secondaryValAxis?.fontFace } },
    { units: chart.secondaryCatAxis?.displayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.08, axis: { size: chart.secondaryCatAxis?.fontSizeHpt, bold: chart.secondaryCatAxis?.fontBold, italic: chart.secondaryCatAxis?.fontItalic, color: chart.secondaryCatAxis?.fontColor, paintAuthored: chart.secondaryCatAxis?.fontPaintAuthored, face: chart.secondaryCatAxis?.fontFace } },
  ];
  for (const { units, vertical, fallbackX, fallbackY, axis } of entries) {
    const label = units?.label;
    if (!units || !label) continue;
    const text = label.text ?? automaticDisplayUnitLabel(units);
    const chartText = chart.chartTextStyle;
    const fontPx = chartTextFontSizePx(
      label.fontSizeHpt ?? axis.size ?? chartText?.fontSizeHpt,
      ptToPx,
    ) ?? 10 * ptToPx;
    const fontBold = label.fontBold ?? axis.bold ?? chartText?.fontBold ?? false;
    const fontItalic = label.fontItalic ?? axis.italic ?? chartText?.fontItalic ?? false;
    const fontPaint = label.fontPaintAuthored === true
      ? { color: label.fontColor, hidden: label.fontHidden === true || label.fontColor == null }
      : axis.paintAuthored === true
        ? { color: axis.color, hidden: axis.color == null }
        : chartText?.fontPaintAuthored === true
          ? { color: chartText.fontColor, hidden: chartText.fontColor == null }
          : { color: label.fontColor ?? axis.color ?? chartText?.fontColor, hidden: false };
    if (fontPaint.hidden) continue;
    const fontColor = fontPaint.color;
    const fontFace = label.fontFace ?? axis.face ?? chartText?.fontFace;
    ctx.save();
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, fontFace, 'minor'),
      fontBold,
      fontItalic,
    );
    const rotation = label.rotation != null
      ? (label.rotation / 60_000) * Math.PI / 180
      : vertical ? -Math.PI / 2 : 0;
    const textWidth = ctx.measureText(text).width;
    const rotatedW = Math.abs(Math.cos(rotation)) * textWidth + Math.abs(Math.sin(rotation)) * fontPx;
    const rotatedH = Math.abs(Math.sin(rotation)) * textWidth + Math.abs(Math.cos(rotation)) * fontPx;
    const automatic = {
      x: fallbackX - rotatedW / 2,
      y: fallbackY - rotatedH / 2,
      w: rotatedW,
      h: rotatedH,
    };
    const positioned = label.manualLayout
      ? resolveManualLayoutRect(
          { ...label.manualLayout, w: undefined, h: undefined },
          rect,
          automatic,
        )
      : automatic;
    if (!positioned) { ctx.restore(); continue; }
    const cx = positioned.x + positioned.w / 2;
    const cy = positioned.y + positioned.h / 2;
    paintChartLabelBox(
      ctx,
      effectiveLinkedLabelBox(
        chart,
        label.boxStyle,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
        true,
      ),
      positioned,
      ptToPx,
    );
    ctx.translate(cx, cy);
    if (rotation !== 0) ctx.rotate(rotation);
    ctx.fillStyle = fontColor ? `#${fontColor}` : '#595959';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 0, 0);
    ctx.restore();
  }
}

/** Build a {@link ValueAxisPlan} for the primary value axis. `dataMin`/`dataMax`
 *  are the raw data extents already massaged by the caller (0-anchoring, pct
 *  normalization, explicit valMin/valMax). `axisLenPt` drives the auto major
 *  unit. Reversal is read from the chart's value-axis orientation. */
export function planValueAxis(
  chart: ChartModel,
  dataMin: number,
  dataMax: number,
  axisLenPt?: number,
  percentStacked = false,
  axisOrientation: 'vertical' | 'horizontal' = 'vertical',
): ValueAxisPlan {
  const reversed = valAxisReversed(chart);
  const logBase = chart.valAxisLogBase;
  // c:valAx values remain ratios for percentStacked charts, but all plotted
  // geometry in this renderer is expressed as percentage points. Explicit
  // bounds/units therefore cross the same ×100 boundary as the series values.
  // With no explicit bounds, percentStacked uses its exact normalized extent
  // (0..100 or -100..100) instead of adding ordinary numeric-axis headroom.
  const explicitMin = valueAxisUnitInRendererSpace(chart.valMin, percentStacked)
    ?? (percentStacked ? dataMin : chart.valMin);
  const explicitMax = valueAxisUnitInRendererSpace(chart.valMax, percentStacked)
    ?? (percentStacked ? dataMax : chart.valMax);
  const authoredMajorUnit = valueAxisUnitInRendererSpace(
    chart.valAxisMajorUnit,
    percentStacked,
  );
  const majorUnit = percentStacked
    && !(logBase != null && isFinite(logBase) && logBase >= 2)
    && !(authoredMajorUnit != null && isFinite(authoredMajorUnit) && authoredMajorUnit > 0)
      ? automaticPercentMajorUnit(dataMin, dataMax, axisOrientation, axisLenPt)
      : authoredMajorUnit;
  const needsMinorTicks = chart.valAxisMinorTickMark != null
    && chart.valAxisMinorTickMark !== 'none';
  const mu = valueAxisUnitInRendererSpace(chart.valAxisMinorUnit, percentStacked);
  const numeric = planNumericValueAxis({
    dataMin,
    dataMax,
    explicitMin,
    explicitMax,
    axisLenPt,
    axisOrientation,
    majorUnit,
    minorUnit: mu,
    needMinor: chart.valAxisMinorGridlines === true || needsMinorTicks,
    logBase,
    reversed,
  });
  const { min, max, majorUnit: step, majorTicks: majorLines } = numeric;
  const minorLines = chart.valAxisMinorGridlines ? numeric.minorTicks : [];
  return {
    min, max, step, majorLines, minorLines, minorTicks: numeric.minorTicks,
    frac: numeric.fraction,
  };
}

export interface TrendlineLabelContext {
  chart: ChartModel;
  chartRect: ChartRect;
  plotRect: ChartRect;
  clipLineToPlot?: boolean;
  automaticAnchor?: { x: number; y: number };
  shapeRotationDeg?: number;
}

export function compactTrendlineNumber(value: number, formatCode?: string | null): string {
  if (formatCode && formatCode.trim().toLowerCase() !== 'general') {
    return formatChartValWithCode(value, formatCode);
  }
  return formatChartVal(Number(value.toPrecision(6)));
}

export function generatedTrendlineLabel(
  tl: ChartTrendline,
  stats: ReturnType<typeof linearTrendlineStats>,
  sourceFormatCode?: string | null,
): string[] {
  if (tl.labelText) return tl.labelText.split(/\r?\n/);
  if (!stats) return [];
  const formatCode = tl.labelFormatSourceLinked === true
    ? sourceFormatCode
    : tl.labelFormatCode;
  const lines: string[] = [];
  if (tl.dispEq) {
    const sign = stats.intercept < 0 ? '−' : '+';
    lines.push(
      `y = ${compactTrendlineNumber(stats.slope, formatCode)}x ${sign} ${compactTrendlineNumber(Math.abs(stats.intercept), formatCode)}`,
    );
  }
  if (tl.dispRSqr) lines.push(`R² = ${compactTrendlineNumber(stats.rSquared, formatCode)}`);
  return lines;
}

export function drawTrendlineLabel(
  ctx: CanvasRenderingContext2D,
  tl: ChartTrendline,
  stats: ReturnType<typeof linearTrendlineStats>,
  ptToPx: number,
  labelContext?: TrendlineLabelContext,
  sourceFormatCode?: string | null,
): void {
  if (!labelContext) return;
  const lines = generatedTrendlineLabel(tl, stats, sourceFormatCode);
  if (lines.length === 0) return;
  const { chart, chartRect, plotRect } = labelContext;
  const fontPx = chartTextFontSizePx(tl.labelFontSizeHpt, ptToPx)
    ?? chartTextFontSizePx(chart.dataLabelFontSizeHpt, ptToPx)
    ?? 10 * ptToPx;
  const face = chartFontFamily(chart, tl.labelFontFace ?? chart.dataLabelFontFace, 'minor');
  const bold = tl.labelFontBold ?? chart.dataLabelFontBold ?? false;
  const italic = tl.labelFontItalic ?? false;
  ctx.font = chartFontCss(fontPx, face, bold, italic);
  const lineHeight = fontPx * 1.2;
  const color = tl.labelFontColor ?? chart.dataLabelFontColor;
  const rich = tl.labelRichRuns?.length
      ? resolveRichDataLabelBlock(ctx, {
        runs: tl.labelRichRuns,
        ptToPx,
        fontFamily: face,
        fallbackBold: bold,
        fallbackItalic: italic,
        fallbackBaseline: tl.labelFontBaseline ?? undefined,
        fallbackColorHidden: tl.labelFontPaintAuthored === true
          && (tl.labelFontHidden === true || tl.labelFontColor == null),
        fontFamilyForFace: runFace => chartFontFamily(chart, runFace, 'minor'),
      }, fontPx, color ? `#${color}` : '#595959')
    : null;
  const naturalTextWidth = rich?.width
    ?? Math.max(...lines.map(line => ctx.measureText(line).width));
  const textStyle: DataLabelTextStyle = {
    fontColor: tl.labelFontColor ?? undefined,
    fontItalic: italic,
    fontPaintAuthored: tl.labelFontPaintAuthored ?? undefined,
    fontHidden: tl.labelFontHidden ?? undefined,
    fontLanguage: tl.labelFontLanguage ?? undefined,
    fontBaseline: tl.labelFontBaseline ?? undefined,
    textRotation: tl.labelTextRotation ?? undefined,
    textWrap: tl.labelTextWrap ?? undefined,
    textVerticalAnchor: tl.labelTextVerticalAnchor ?? undefined,
    textVerticalMode: tl.labelTextVerticalMode ?? undefined,
    textLInsEmu: tl.labelTextLInsEmu ?? undefined,
    textTInsEmu: tl.labelTextTInsEmu ?? undefined,
    textRInsEmu: tl.labelTextRInsEmu ?? undefined,
    textBInsEmu: tl.labelTextBInsEmu ?? undefined,
    textBodyAuthored: tl.labelTextBodyAuthored ?? undefined,
  };
  const insets = dataLabelInsets(textStyle, ptToPx);
  const naturalWidth = naturalTextWidth + insets.left + insets.right;
  const naturalHeight = (rich?.height ?? lines.length * lineHeight)
    + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    naturalWidth, naturalHeight,
    tl.labelTextRotation ?? undefined,
    tl.labelTextVerticalMode ?? undefined,
  );
  const placement = placeTrendlineLabel(
    chartRect,
    plotRect,
    rotated.w,
    rotated.h,
    fontPx,
    tl.labelManualLayout,
    labelContext.automaticAnchor,
  );
  if (!placement) return;

  ctx.save();
  if (placement.automatic) {
    ctx.beginPath();
    ctx.rect(plotRect.x, plotRect.y, plotRect.w, plotRect.h);
    ctx.clip();
  }
  const centerX = placement.x + placement.w / 2;
  const centerY = placement.y + placement.h / 2;
  const hasTextBody = tl.labelTextBodyAuthored === true
    || tl.labelTextRotation != null
    || tl.labelTextWrap != null
    || tl.labelTextVerticalAnchor != null
    || tl.labelTextVerticalMode != null
    || tl.labelTextLInsEmu != null
    || tl.labelTextTInsEmu != null
    || tl.labelTextRInsEmu != null
    || tl.labelTextBInsEmu != null;
  // `manualLayout` sizes the authored label shape. Automatic labels are sized
  // from measured text. `bodyPr@rot` rotates text inside that shape, not the
  // shape paint itself (ECMA-376 §20.1.10.83/§21.2.2.216).
  const boxRect = placement.automatic
    ? {
        x: centerX - naturalWidth / 2,
        y: centerY - naturalHeight / 2,
        w: naturalWidth,
        h: naturalHeight,
      }
    : { x: placement.x, y: placement.y, w: placement.w, h: placement.h };
  // chartStyleRoleTrendlineLabel materializes the effective box before paint.
  // Reapplying the linked role here would reinterpret an already-resolved
  // fail-closed paint as a fresh direct noFill and could revive the fallback.
  const labelBox = tl.labelBox;
  paintChartLabelBox(
    ctx,
    labelBox,
    boxRect,
    ptToPx,
    labelContext.shapeRotationDeg ?? 0,
  );
  const alignment = tl.labelTextAlign;
  ctx.textAlign = alignment === 'r' ? 'right' : alignment === 'ctr' ? 'center' : 'left';
  ctx.textBaseline = 'top';
  ctx.fillStyle = color ? `#${color}` : '#595959';
  const maxTextWidth = Math.max(0, boxRect.w - insets.left - insets.right);
  const maxTextHeight = Math.max(0, boxRect.h - insets.top - insets.bottom);
  const automaticNaturalBlockFits = placement.automatic
    && rotated.radians === 0
    && placement.w === rotated.w
    && placement.h === rotated.h
    && (tl.labelTextWrap == null || tl.labelTextWrap === 'none');
  // When the automatic box was measured from these exact generated lines,
  // preserve them directly. Re-fitting an exact two-line body through a
  // floating-point height division could floor 1.999… to one line and drop
  // the authored R² output even though the measured box had room for both.
  const displayLines = rich ? [] : hasTextBody && !automaticNaturalBlockFits
    ? fitStyledDataLabelLines(
        lines.join('\n'), maxTextWidth, maxTextHeight, lineHeight,
        value => ctx.measureText(value).width, textStyle,
      )
    : lines;
  if (!rich && displayLines.length === 0) {
    ctx.restore();
    return;
  }
  const textX = !hasTextBody
    ? (ctx.textAlign === 'right'
      ? placement.x + placement.w
      : ctx.textAlign === 'center' ? placement.x + placement.w / 2 : placement.x)
    : ctx.textAlign === 'right'
    ? boxRect.x + boxRect.w - insets.right
    : ctx.textAlign === 'center'
      ? boxRect.x + (boxRect.w + insets.left - insets.right) / 2
      : boxRect.x + insets.left;
  const baselineShift = (tl.labelFontBaseline ?? 0) * fontPx;
  const textTop = !hasTextBody
    ? placement.y
    : tl.labelTextVerticalAnchor === 'b'
    ? boxRect.y + boxRect.h - insets.bottom - (rich?.height ?? displayLines.length * lineHeight)
    : tl.labelTextVerticalAnchor === 'ctr'
      ? boxRect.y
        + (boxRect.h - (rich?.height ?? displayLines.length * lineHeight)
          + insets.top - insets.bottom) / 2
      : boxRect.y + insets.top;
  const completeLines = placement.automatic
    ? displayLines.length
    : Math.min(displayLines.length, Math.floor(maxTextHeight / lineHeight));
  if (rotated.radians !== 0) {
    ctx.translate(centerX, centerY);
    ctx.rotate(rotated.radians);
    ctx.translate(-centerX, -centerY);
  }
  if (rich) {
    paintRichDataLabelBlock(
      ctx, rich, textX, textTop, ctx.textAlign, 'top',
      Math.max(rich.width, maxTextWidth),
    );
  } else if (!(tl.labelFontPaintAuthored === true
    && (tl.labelFontHidden === true || tl.labelFontColor == null))) {
    for (let index = 0; index < completeLines; index++) {
      ctx.fillText(
        hasTextBody && textStyle.textWrap === 'none'
          ? displayLines[index]
          : elideToWidth(ctx, displayLines[index], Math.max(0, maxTextWidth || naturalTextWidth)),
        textX,
        textTop + index * lineHeight - baselineShift,
      );
    }
  }
  ctx.restore();
}

/** Draw a series' `<c:trendline>` regression lines (ECMA-376 §21.2.2.211).
 *  Each trendline is fitted over the series' non-null `(categoryIndex, value)`
 *  points via {@link fitTrendline} and stroked through the chart's
 *  `toX` (category-index → pixel) and `toY` (value → pixel) maps. `forward` /
 *  `backward` extend the linear fit past the data ends by that many category
 *  units. Nonlinear types are sampled by the same bounded fitter. `seriesColor`
 *  is the fallback stroke when the trendline declares no
 *  `<a:ln>` color. Byte-stable no-op for series with no trendline. */
export function drawSeriesTrendlines(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  seriesColor: string,
  toX: (i: number) => number,
  toY: (v: number) => number,
  ptToPx: number,
  xValues?: readonly (number | null)[],
  labelContext?: TrendlineLabelContext,
  mapPoint?: (categoryValue: number, seriesValue: number) => { x: number; y: number },
): void {
  const tls = s.trendLines;
  if (!tls || tls.length === 0) return;
  // Collect the fittable (index, value) points once.
  const xs: number[] = []; const ys: number[] = [];
  for (let i = 0; i < s.values.length; i++) {
    const v = s.values[i];
    const x = xValues ? xValues[i] : i;
    if (v != null && x != null && Number.isFinite(v) && Number.isFinite(x)) {
      xs.push(x);
      ys.push(v);
    }
  }
  if (xs.length < 2) return;
  const prevDash = ctx.getLineDash ? ctx.getLineDash() : [];
  for (const tl of tls) {
    const fit = fitTrendline(xs, ys, tl.trendlineType, {
      period: tl.period,
      order: tl.order,
      intercept: tl.intercept,
      forward: tl.forward,
      backward: tl.backward,
    });
    if (fit.xs.length < 2) continue;
    if (![...fit.xs, ...fit.ys].every(Number.isFinite)) continue;
    const candidateStats = tl.trendlineType === 'linear'
      ? linearTrendlineStats(xs, ys, tl.intercept)
      : null;
    const stats = candidateStats && [
      candidateStats.slope,
      candidateStats.intercept,
      candidateStats.rSquared,
    ].every(Number.isFinite)
      ? candidateStats
      : null;
    // For a linear fit, forward/backward extend the two endpoints along the
    // fitted slope (in category-index units).
    let fxs = fit.xs; let fys = fit.ys;
    if (tl.trendlineType === 'linear') {
      const m = (fit.ys[1] - fit.ys[0]) / ((fit.xs[1] - fit.xs[0]) || 1);
      const bwd = tl.backward ?? 0; const fwd = tl.forward ?? 0;
      const x0 = fit.xs[0] - bwd; const x1 = fit.xs[1] + fwd;
      fxs = [x0, x1];
      fys = [fit.ys[0] - m * bwd, fit.ys[1] + m * fwd];
    }
    if (![...fxs, ...fys].every(Number.isFinite)) continue;
    const mapped = fxs.map((x, index) => mapPoint
      ? mapPoint(x, fys[index])
      : ({ x: toX(x), y: toY(fys[index]) }));
    if (!mapped.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) continue;
  if (!tl.lineHidden && (tl.linePaintAuthored !== true || tl.lineColor != null)) {
      if (labelContext?.clipLineToPlot) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(
          labelContext.plotRect.x,
          labelContext.plotRect.y,
          labelContext.plotRect.w,
          labelContext.plotRect.h,
        );
        ctx.clip();
      }
      ctx.strokeStyle = tl.lineColor ? `#${tl.lineColor}` : seriesColor;
      ctx.lineWidth = tl.lineWidthEmu ? axisLineWidthPx(tl.lineWidthEmu, ptToPx) : 1.5;
      // DrawingML line presets are authored paint; omission means solid.
      ctx.setLineDash(dashPatternForPreset(tl.lineDash ?? undefined, ctx.lineWidth));
      ctx.beginPath();
      for (let i = 0; i < mapped.length; i++) {
        const { x: px, y: py } = mapped[i];
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.stroke();
      if (labelContext?.clipLineToPlot) ctx.restore();
    }
    drawTrendlineLabel(ctx, tl, stats, ptToPx, labelContext ? {
      ...labelContext,
      automaticAnchor: mapped.at(-1),
    } : undefined, s.valFormatCode);
  }
  ctx.setLineDash(prevDash);
}

/** Office adds every visible trendline to a series-driven legend. The legend
 * entry is a line key whose authored paint comes from `<c:trendline><c:spPr>`;
 * when `<c:name>` is absent, the application-generated label combines the
 * localized trendline kind with the source-series name. We use the invariant
 * OOXML kind names here and keep an authored name verbatim. */
export function trendlineLegendSeries(series: readonly ChartSeries[]): ChartSeries[] {
  const kindLabel = (kind: string): string => {
    switch (kind) {
      case 'exp': return 'Exponential';
      case 'log': return 'Logarithmic';
      case 'poly': return 'Polynomial';
      case 'power': return 'Power';
      case 'movingAvg': return 'Moving Average';
      default: return 'Linear';
    }
  };
  const entries: ChartSeries[] = [];
  for (const source of series) {
    for (const trendline of source.trendLines ?? []) {
      if (trendline.lineHidden === true
        || (trendline.linePaintAuthored === true && trendline.lineColor == null)) continue;
      const fallbackColor = source.lineColor ?? source.color;
      entries.push({
        name: trendline.name
          ?? `${kindLabel(trendline.trendlineType)} (${source.name || 'Series'})`,
        color: trendline.lineColor ?? fallbackColor,
        lineColor: trendline.lineColor ?? fallbackColor,
        lineWidthEmu: trendline.lineWidthEmu,
        lineHidden: false,
        chartexStyle: { lineDash: trendline.lineDash },
        values: [],
        seriesType: 'line',
        showMarker: false,
      });
    }
  }
  return entries;
}

export function legendSeriesWithTrendlines(chart: ChartModel): ChartSeries[] {
  if (legendIsCategoryDriven(
    chart.chartType,
    chart.series.length,
    chart.varyColors !== false,
  ) || chartVariesColorsByPoint(chart)) {
    return chart.series;
  }
  return chart.series.flatMap(series => [series, ...trendlineLegendSeries([series])]);
}

/** Resolve an axis label font size (px) from <c:txPr> hpt or a proportional
 *  fallback. ptToPx comes from the host renderer (EMU/px scale at display). */
export function axisLabelPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  return chartTextFontSizePx(sizeHpt, ptToPx) ?? Math.max(8, h * 0.045);
}

/** Wrap text against the active canvas font without discarding characters.
 * Words are kept intact when possible; a single over-wide token is split at
 * measured character boundaries. Used by chart families whose category-label
 * band is an input to plot layout. */
export function wrapMeasuredText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  singleTokenOverhangPx = 0,
): string[] {
  const words = text.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return [''];
  const lines: string[] = [];
  let line = '';
  const pushToken = (token: string): void => {
    const trial = line ? `${line} ${token}` : token;
    if (ctx.measureText(trial).width <= maxWidth) {
      line = trial;
      return;
    }
    if (line) {
      lines.push(line);
      line = '';
    }
    if (ctx.measureText(token).width <= maxWidth + singleTokenOverhangPx) {
      line = token;
      return;
    }
    // Find each largest fitting code-point prefix by binary search. Measuring
    // every growing prefix makes a single long unbroken label quadratic.
    const chars = Array.from(token);
    let start = 0;
    while (start < chars.length) {
      let low = start + 1;
      let high = chars.length;
      let end = start + 1; // Always make progress, even if one glyph is wider.
      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (ctx.measureText(chars.slice(start, mid).join('')).width <= maxWidth) {
          end = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      const chunk = chars.slice(start, end).join('');
      start = end;
      if (start < chars.length) lines.push(chunk);
      else line = chunk;
    }
  };
  for (const word of words) pushToken(word);
  if (line) lines.push(line);
  return lines.length ? lines : [''];
}

/** Office keeps short numeric category labels on one line when its native
 * theme-font metrics fit the slot. A browser without that Office font may use
 * a slightly wider fallback and otherwise split `10` into `1` / `0`. Permit a
 * small metric-only overhang for numeric tokens; ordinary text still obeys the
 * exact measured slot and genuinely over-wide numbers continue to wrap. */
export function numericCategoryMetricTolerance(text: string, fontPx: number): number {
  return /^[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)%?$/.test(text)
    ? fontPx * 0.15
    : 0;
}

/** Whether the CATEGORY tick labels should be drawn. `<c:catAx><c:tickLblPos
 *  val="none">` (ECMA-376 §21.2.2.207) hides them; anything else (incl. absent)
 *  shows them, so the default is byte-stable. */
export function catLabelsVisible(chart: ChartModel): boolean {
  return chart.catAxisTickLabelPos !== 'none';
}

/** 90° in 60000ths of a degree. `ST_FixedAngle` (ECMA-376 §20.1.10.23) bounds
 *  a fixed-range angle to the OPEN interval "greater than -5400000 / less than
 *  5400000", so ±5400000 itself lies outside the schema type — but Office's
 *  Format-Axis "Custom angle" control accepts -90°…+90° INCLUSIVE, so the code
 *  below deliberately uses a closed boundary (`> LIMIT` rejects, `== LIMIT`
 *  honors) to keep genuine ±90° (vertical) axis labels working. */
export const FIXED_ANGLE_LIMIT_60K = 5_400_000;

/** Category-axis label rotation in RADIANS (canvas convention), from
 *  `<c:catAx|dateAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`
 *  §20.1.10.3, 60000ths of a degree). Returns 0 when unset — the un-rotated
 *  fast path callers keep.
 *
 *  `bodyPr@rot` is typed `ST_Angle` (a restriction of XML Schema `int`, so any
 *  integer is schema-valid), but a *text* rotation is only meaningful within
 *  the `ST_FixedAngle` (§20.1.10.23) fixed-angle domain — an open interval
 *  (-90°, 90°) at the schema level, which Office's Format-Axis "Custom angle"
 *  control widens to -90°…+90° inclusive (we follow the UI's closed range; see
 *  {@link FIXED_ANGLE_LIMIT_60K}). Office writes `rot="-60000000"` (-1000°,
 *  ≈2.8 full turns) as a sentinel for "auto / horizontal" axis text and renders
 *  those labels horizontal; the identical value even appears on the numeric
 *  value axes whose Office-rendered labels are horizontal. So a rot whose magnitude exceeds ±90°
 *  is outside the valid text-rotation domain and is treated as no rotation
 *  (0°) rather than reduced mod 360 (which would map -1000° → +80°,
 *  near-vertical). Genuine rotations within the
 *  closed range (-45° = -2700000, -90° = -5400000) are honored unchanged. */
export function catLabelRotationRad(chart: ChartModel): number {
  const rot = chart.catAxisLabelRotation;
  if (rot == null || rot === 0) return 0;
  if (Math.abs(rot) > FIXED_ANGLE_LIMIT_60K) return 0;
  return (rot / 60000) * (Math.PI / 180);
}

/** Draw a category label at `(x, y)` with optional rotation. `rotRad === 0`
 *  keeps the exact non-rotated draw the callers used before (byte-stable):
 *  `ctx.fillText(text, x, y)` with the caller's current align/baseline. When
 *  rotated, the label pivots around `(x, y)` and is right-aligned+middle so the
 *  text trails up-left from the tick, matching PowerPoint's angled axis labels. */
export function drawRotatedCatLabel(
  ctx: CanvasRenderingContext2D, text: string, x: number, y: number, rotRad: number,
): void {
  if (rotRad === 0) {
    ctx.fillText(text, x, y);
    return;
  }
  ctx.save();
  ctx.translate(x, y);
  ctx.rotate(rotRad);
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

/** Resolved secondary value-axis scale (combo charts). `min`/`max`/`step` are
 *  the "nice" bounds + major unit; `makeToY(py0, ph)` builds the value→pixel
 *  mapping once the final plot rect is known (the scale is computed BEFORE the
 *  pad/gutter math from an estimated plot height, so the mapping factory is
 *  split out). See {@link computeSecondaryAxis}. */
export interface SecondaryAxisScale {
  min: number;
  max: number;
  step: number;
  majorLines: number[];
  minorTicks: number[];
  makeToY: (py0: number, ph: number) => (v: number) => number;
}

/** Build the shared calendar/category mapping for every classic chart family
 * that can be bound to `<c:dateAx>`. Keeping this in one place prevents combo,
 * line, area and stock renderers from assigning different x coordinates to
 * the same authored date cache. */
export function chartDateAxisPlan(
  chart: ChartModel,
  categories: readonly string[],
  reversed = catAxisReversed(chart),
): ReturnType<typeof planDateCategoryAxis> {
  if (chart.catAxisIsDate !== true) return null;
  return planDateCategoryAxis({
    categories,
    date1904: chart.date1904,
    baseTimeUnit: chart.catAxisBaseTimeUnit,
    majorTimeUnit: chart.catAxisMajorTimeUnit,
    majorUnit: chart.catAxisMajorUnit,
    minorTimeUnit: chart.catAxisMinorTimeUnit,
    minorUnit: chart.catAxisMinorUnit,
    explicitMin: chart.catAxisMin,
    explicitMax: chart.catAxisMax,
    crossBetween: isCrossBetween(chart),
    reversed,
  });
}

/** Visit the authored error-bar endpoint values on one numeric axis. The
 * parser has already expanded percentage/fixed/custom forms into per-point
 * positive magnitudes, so scale planning only needs the same plus/minus gates
 * used by paint. */
export function forEachErrorBarEndpoint(
  series: ChartSeries,
  direction: 'x' | 'y',
  baseAt: (index: number) => number | null,
  visit: (value: number) => void,
): void {
  for (const errorBars of series.errBars ?? []) {
    if (errorBars.dir !== direction) continue;
    const drawPlus = errorBars.barType === 'plus' || errorBars.barType === 'both';
    const drawMinus = errorBars.barType === 'minus' || errorBars.barType === 'both';
    const count = Math.max(series.values.length, errorBars.plus.length, errorBars.minus.length);
    for (let index = 0; index < count; index++) {
      const base = baseAt(index);
      if (base == null || !Number.isFinite(base)) continue;
      const plus = errorBars.plus[index];
      const minus = errorBars.minus[index];
      if (drawPlus && plus != null && Number.isFinite(plus)) visit(base + plus);
      if (drawMinus && minus != null && Number.isFinite(minus)) visit(base - minus);
    }
  }
}

/** Compute the INDEPENDENT scale of a secondary value axis from the series that
 *  opt into it (`useSecondaryAxis === true`). Shared by every axis family that
 *  supports a secondary axis (bar-combo line series, and plain line / area
 *  series): the axis has its own bounded automatic plan, with an explicit
 *  `<c:scaling><c:min/max>` (`sec.min`/`sec.max`) overriding. Returns
 *  null when no `SecondaryValueAxis` was parsed OR no series opts into it — the
 *  caller then keeps the single-axis path unchanged.
 *
 *  `plotHeightPt` is the estimated plot height in points (the axis is the
 *  vertical right edge, so its length drives the auto major unit). `getValues`
 *  yields each opted-in series' raw values.
 *
 *  Empty secondary data keeps the neutral 0..1 fallback. */
export function computeSecondaryAxis(
  sec: SecondaryValueAxis | null,
  seriesForSecondary: ChartSeries[],
  plotHeightPt: number,
  errorBarDirection: 'x' | 'y' = 'y',
  percentStacked = false,
  includeZero = false,
  isSecondary: (series: ChartSeries, index: number) => boolean = series =>
    series.useSecondaryAxis === true,
  valueAt: (series: ChartSeries, pointIndex: number, seriesIndex: number) => number | null =
    () => null,
): SecondaryAxisScale | null {
  if (!sec) return null;
  let dMin = Infinity;
  let dMax = -Infinity;
  const include = (value: number): void => {
    if (!Number.isFinite(value)) return;
    dMin = Math.min(dMin, value);
    dMax = Math.max(dMax, value);
  };
  if (includeZero) include(0);
  for (let seriesIndex = 0; seriesIndex < seriesForSecondary.length; seriesIndex++) {
    const s = seriesForSecondary[seriesIndex];
    if (!isSecondary(s, seriesIndex)) continue;
    for (let pointIndex = 0; pointIndex < s.values.length; pointIndex++) {
      const value = valueAt(s, pointIndex, seriesIndex) ?? s.values[pointIndex];
      if (value != null) include(value);
    }
    forEachErrorBarEndpoint(
      s,
      errorBarDirection,
      pointIndex => valueAt(s, pointIndex, seriesIndex) ?? s.values[pointIndex] ?? null,
      include,
    );
  }
  if (!Number.isFinite(dMin) || !Number.isFinite(dMax)) {
    dMin = 0;
    dMax = 1;
  }
  // An explicit `<c:valAx><c:majorUnit>` on the secondary axis (§21.2.2.103)
  // overrides the auto step, mirroring the primary axis. null ⇒ auto.
  const numeric = planNumericValueAxis({
    dataMin: dMin,
    dataMax: dMax,
    explicitMin: valueAxisUnitInRendererSpace(sec.min, percentStacked),
    explicitMax: valueAxisUnitInRendererSpace(sec.max, percentStacked),
    axisLenPt: plotHeightPt,
    axisOrientation: 'vertical',
    majorUnit: valueAxisUnitInRendererSpace(sec.majorUnit, percentStacked),
    minorUnit: valueAxisUnitInRendererSpace(sec.minorUnit, percentStacked),
    needMinor: sec.minorGridlines === true
      || (sec.minorTickMark != null && sec.minorTickMark !== 'none'),
    logBase: sec.logBase,
    reversed: sec.orientation === 'maxMin',
  });
  const { min, max, majorUnit: step } = numeric;
  return {
    min,
    max,
    step,
    majorLines: numeric.majorTicks,
    minorTicks: numeric.minorTicks,
    makeToY: (py0: number, ph: number) => (v: number): number =>
      py0 + ph - numeric.fraction(v) * ph,
  };
}

/** Paint secondary-axis gridlines below chart data. Axis rules, ticks and
 * labels remain in {@link drawSecondaryValueAxis}, which is intentionally
 * called after the series. Keeping these layers separate prevents translucent
 * fills and bars from being incorrectly overpainted by right-axis grids. */
export function drawSecondaryValueGridlines(
  ctx: CanvasRenderingContext2D,
  sec: SecondaryValueAxis,
  secScale: SecondaryAxisScale,
  toYSecondary: (v: number) => number,
  px0: number,
  pw: number,
  ptToPx: number,
): void {
  if (sec.hidden) return;
  ctx.save();
  if (sec.minorGridlines) {
    const grid = secondaryMinorGridStroke(sec, ptToPx);
    for (const value of secScale.minorTicks) {
      strokeValueGridlineH(ctx, px0, pw, toYSecondary(value), false, grid);
    }
  }
  if (sec.majorGridlines) {
    const grid = secondaryMajorGridStroke(sec, ptToPx);
    for (const value of secScale.majorLines) {
      strokeValueGridlineH(ctx, px0, pw, toYSecondary(value), false, grid);
    }
  }
  ctx.restore();
}

/** Draw a secondary value axis on the RIGHT edge of the plot: its rule, mirrored
 *  tick marks + labels, and rotated title. Its scale is INDEPENDENT of the
 *  primary axis (its own "nice" major unit; NOT aligned to the primary
 *  gridlines) — PowerPoint places these marks independently. Shared by the
 *  bar, line and area families so no combo path can retain a divergent rule,
 *  tick, label, or title policy.
 *  Callers pass:
 *  - `secScale`   the resolved scale (from {@link computeSecondaryAxis}),
 *  - `toYSecondary` the value→pixel map (`secScale.makeToY(py0, ph)`),
 *  - `secFontPx` / `secLabelBandW` the tick-label font size + reserved gutter
 *    width (measured up front so the title clears the labels),
 *  - `primaryLabelColor` the fallback tick-label color when the axis specifies
 *    none (the primary value-axis label color). */
export function drawSecondaryValueAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  sec: SecondaryValueAxis,
  secScale: SecondaryAxisScale,
  toYSecondary: (v: number) => number,
  chartRect: ChartRect,
  px0: number, py0: number, pw: number, ph: number,
  ptToPx: number,
  secFontPx: number,
  secLabelBandW: number,
  primaryLabelColor: string,
  date1904: boolean | undefined,
  percentStacked = false,
): void {
  const axX = px0 + pw;
  const { color: secLineColor, width: secLineW } = resolveAxisLine(sec.lineColor, sec.lineWidthEmu, ptToPx);
  if (!sec.lineHidden) {
    strokeAxisSegment(ctx, axX, py0, axX, py0 + ph, secLineColor, secLineW, sec.lineDash);
  }
  if (!sec.hidden) {
    ctx.font = `${sec.fontItalic ? 'italic ' : ''}${sec.fontBold ? 'bold ' : ''}${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    ctx.fillStyle = sec.fontColor ? `#${sec.fontColor}` : primaryLabelColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    for (const sval of secScale.majorLines) {
      const gy = toYSecondary(sval);
      // Same tick geometry as the left axis, mirrored to the right edge.
      drawAxisTick(ctx, sec.majorTickMark, 'val', axX, gy, secLineColor, secLineW, true, sec.lineHidden, 'major', ptToPx, sec.lineDash);
      if (sec.tickLabelPos !== 'none') {
        ctx.fillText(
          formatAxisTickWithUnits(
            percentStacked ? sval / 100 : sval,
            sec.formatCode ?? null,
            date1904,
            sec.displayUnits,
          ),
          axX + 14,
          gy,
        );
      }
    }
    if (sec.minorTickMark && sec.minorTickMark !== 'none') {
      for (const value of secScale.minorTicks) {
        drawAxisTick(ctx, sec.minorTickMark, 'val', axX, toYSecondary(value), secLineColor, secLineW, true, sec.lineHidden, 'minor', ptToPx, sec.lineDash);
      }
    }
  }
  if (sec.title) {
    drawSecondaryAxisTitle(
      ctx, chart, sec, chartRect, px0, py0, pw, ph, secLabelBandW, ptToPx,
    );
  }
}

/** Draw a right-side secondary value-axis title. Both the duplicated combo-bar
 *  path and the shared line/area path use this helper, so the fixed 10pt
 *  fallback and +90° reading direction cannot drift. */
export function drawSecondaryAxisTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  sec: SecondaryValueAxis,
  chartRect: ChartRect,
  px0: number, py0: number, pw: number, ph: number,
  secLabelBandW: number,
  ptToPx: number,
): void {
  if (!sec.title) return;
  const fontSizePx = axisTitleFontPx(sec.titleFontSizeHpt, ptToPx);
  const color = sec.titleFontColor
    ? `#${sec.titleFontColor}`
    : (sec.fontColor ? `#${sec.fontColor}` : '#555');
  drawAxisTitle(
    ctx,
    sec.title,
    px0 + pw + secLabelBandW + fontSizePx * 0.6,
    py0 + ph / 2,
    'right',
    fontSizePx,
    sec.titleFontBold ?? true,
    sec.titleFontItalic ?? false,
    color,
    ph,
    chartFontFamily(chart, sec.titleFontFace, 'major'),
    sec.titleRotation,
    sec.titleVerticalMode,
    sec.titleManualLayout,
    chartRect,
    effectiveLinkedLabelBox(
      chart,
      sec.titleStyle ? { style: sec.titleStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
      true,
    ),
    ptToPx,
  );
}

/** Draw the categorical axis paired with a secondary bar/column group. Its
 * top rule, ticks, labels and title are distinct authored objects from the
 * primary bottom category axis; the secondary value axis remains responsible
 * for the paired right-hand numeric scale. */
export function drawSecondaryCategoryAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  axis: SecondaryValueAxis,
  categories: readonly string[],
  chartRect: ChartRect,
  px0: number,
  py0: number,
  pw: number,
  ptToPx: number,
): void {
  if (axis.hidden || categories.length === 0) return;
  const { color, width } = resolveAxisLine(axis.lineColor, axis.lineWidthEmu, ptToPx);
  if (!axis.lineHidden) {
    strokeAxisSegment(ctx, px0, py0, px0 + pw, py0, color, width, axis.lineDash);
  }
  const reversed = axis.orientation === 'maxMin';
  const labelSkip = Math.max(1, Math.floor(axis.tickLabelSkip ?? 1));
  const markSkip = Math.max(1, Math.floor(axis.tickMarkSkip ?? 1));
  const count = categories.length;
  const labelAnchor = (index: number) => categoryLabelAnchorFraction(
    index,
    count,
    isCrossBetween(chart),
    reversed,
    axis.labelAlignment,
  );
  if (!axis.lineHidden && axis.majorTickMark !== 'none') {
    const onBoundary = isCrossBetween(chart);
    const last = onBoundary ? count : count - 1;
    for (let index = 0; index <= last; index += markSkip) {
      const logical = reversed ? last - index : index;
      const fraction = onBoundary
        ? logical / count
        : count === 1 ? 0.5 : logical / (count - 1);
      drawAxisTick(
        ctx, axis.majorTickMark, 'cat', py0, px0 + fraction * pw,
        color, width, true, axis.lineHidden, 'major', ptToPx, axis.lineDash,
      );
    }
  }
  const fontPx = chartTextFontSizePx(axis.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  if (axis.tickLabelPos !== 'none') {
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, axis.fontFace, 'minor'),
      axis.fontBold ?? false,
      axis.fontItalic ?? false,
    );
    ctx.fillStyle = axis.fontColor ? `#${axis.fontColor}` : '#555';
    ctx.textBaseline = 'bottom';
    const budget = Math.max(1, pw / count - 4);
    const labelGap = categoryLabelOffsetPx(
      categoryTickLabelGapPx(fontPx),
      axis.labelOffsetPercent,
    );
    for (let index = 0; index < count; index += labelSkip) {
      const anchor = labelAnchor(index);
      ctx.textAlign = anchor.textAlign;
      ctx.fillText(
        elideToWidth(ctx, formatCategoryLabel(categories[index], axis.formatCode, chart.date1904), budget),
        px0 + anchor.fraction * pw,
        py0 - labelGap,
      );
    }
  }
  if (axis.title) {
    const titleFontPx = axisTitleFontPx(axis.titleFontSizeHpt, ptToPx);
    drawAxisTitle(
      ctx,
      axis.title,
      px0 + pw / 2,
      py0 - (axis.tickLabelPos === 'none'
        ? 0
        : fontPx + categoryLabelOffsetPx(
          categoryTickLabelGapPx(fontPx),
          axis.labelOffsetPercent,
        ))
        - titleFontPx / 2 - 4,
      'horizontal',
      titleFontPx,
      axis.titleFontBold ?? true,
      axis.titleFontItalic ?? false,
      axis.titleFontColor ? `#${axis.titleFontColor}` : '#555',
      pw,
      chartFontFamily(chart, axis.titleFontFace, 'major'),
      axis.titleRotation,
      axis.titleVerticalMode,
      axis.titleManualLayout,
      chartRect,
      effectiveLinkedLabelBox(
        chart,
        axis.titleStyle ? { style: axis.titleStyle } : undefined,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
        true,
      ),
      ptToPx,
    );
  }
}

export interface ResolvedTitlePiece {
  text: string;
  width: number;
  font: string;
  color: string;
}

export interface ResolvedTitleLine {
  pieces: ResolvedTitlePiece[];
  width: number;
  height: number;
}

export function chartTitleRunFont(
  chart: ChartModel,
  run: ChartTextRun,
  fallbackFontSize: number,
): { font: string; fontSize: number; color: string } {
  const titleSizePt = chart.titleFontSizeHpt != null
    && chart.titleFontSizeHpt >= 100
    && chart.titleFontSizeHpt <= 400_000
    ? chart.titleFontSizeHpt / 100
    : 14;
  const effectivePtToPx = fallbackFontSize / titleSizePt;
  const fontSize = chartTextFontSizePx(run.fontSizeHpt, effectivePtToPx) ?? fallbackFontSize;
  const titleFace = resolveThemeFontRef(chart, run.fontFace ?? chart.titleFontFace);
  const face = titleFace ? `"${titleFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
  return {
    font: chartFontCss(
      fontSize,
      face,
      // DrawingML run `b` defaults to false when neither direct text nor an
      // effective numeric/linked title role owns the property. Keep the
      // historical bold fallback only for the public plain-title model, which
      // has no run-level provenance.
      run.bold ?? chart.titleFontBold ?? false,
      run.italic ?? chart.titleFontItalic ?? false,
    ),
    fontSize,
    color: run.colorPaintAuthored === true
      ? run.colorHidden === true || !run.color ? 'transparent' : `#${run.color}`
      : run.color
        ? `#${run.color}`
        : chart.titleFontPaintAuthored === true
          ? chart.titleFontColor ? `#${chart.titleFontColor}` : 'transparent'
          : chart.titleFontColor ? `#${chart.titleFontColor}` : '#333',
  };
}

/** Measure DrawingML title runs against the chart's finite title box. Explicit
 * newlines remain hard breaks; ordinary whitespace is the only automatic wrap
 * opportunity, matching DrawingML's default square text wrapping. */
export function resolveChartTitleLines(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  maxWidth: number,
  fallbackFontSize: number,
): ResolvedTitleLine[] {
  const runs: ChartTextRun[] = chart.titleRichRuns?.length
    ? chart.titleRichRuns
    : chart.title ? [{ text: chart.title }] : [];
  const lines: ResolvedTitleLine[] = [{ pieces: [], width: 0, height: fallbackFontSize }];
  const pushLine = (): ResolvedTitleLine => {
    const line = { pieces: [], width: 0, height: fallbackFontSize } as ResolvedTitleLine;
    lines.push(line);
    return line;
  };
  let line = lines[0];
  for (const run of runs) {
    const style = chartTitleRunFont(chart, run, fallbackFontSize);
    for (const token of run.text.split(/(\n|[\t ]+)/).filter(part => part.length > 0)) {
      if (token === '\n') {
        line = pushLine();
        continue;
      }
      ctx.font = style.font;
      const width = ctx.measureText(token).width;
      const isSpace = /^[\t ]+$/.test(token);
      if (!isSpace && line.pieces.length > 0 && line.width + width > maxWidth) {
        line = pushLine();
      }
      if (isSpace && line.pieces.length === 0) continue;
      line.pieces.push({ text: token, width, font: style.font, color: style.color });
      line.width += width;
      line.height = Math.max(line.height, style.fontSize);
    }
  }
  return lines;
}

export function measuredCartesianTitleBand(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  w: number,
  h: number,
  ptToPx: number,
): ChartTitleBand {
  const base = cartesianTitleBand(chart, h, ptToPx);
  if (base.bandH === 0) return base;
  if (!chart.titleRichRuns?.length) return base;
  const previousFont = ctx.font;
  const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), base.fontPx);
  ctx.font = previousFont;
  const textHeight = lines.reduce((sum, line) => sum + line.height, 0);
  return {
    ...base,
    bandH: base.topPad + textHeight + base.bottomPad,
  };
}

export function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  const titleBox = effectiveLinkedLabelBox(
    chart,
    chart.titleStyle ? { style: chart.titleStyle } : undefined,
    chart.chartStyleRoles?.title,
    rawLinkedChartStyleRole(chart, 'title'),
    true,
  );
  const titlePtToPx = fontSize / Math.max(1, (chart.titleFontSizeHpt ?? 1_400) / 100);
  // Preserve the established single-fillText path for callers/models without
  // formatted DrawingML runs. Besides avoiding unnecessary tokenization, this
  // keeps the public Canvas contract (center-aligned title anchor) unchanged.
  if (!chart.titleRichRuns?.length) {
    const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
    const face = titleFace
      ? `"${titleFace}", Calibri, Arial, sans-serif`
      : 'Calibri, Arial, sans-serif';
    ctx.font = chartFontCss(
      fontSize,
      face,
      chart.titleFontBold ?? true,
      chart.titleFontItalic ?? false,
    );
    ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const measuredWidth = Math.min(w, ctx.measureText(chart.title).width);
    paintChartLabelBox(ctx, titleBox, {
      x: x + (w - measuredWidth) / 2,
      y,
      w: measuredWidth,
      h: fontSize * 1.2,
    }, titlePtToPx);
    ctx.fillText(chart.title, x + w / 2, y);
    return;
  }
  ctx.save();
  const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), fontSize);
  const boxWidth = Math.min(w, Math.max(...lines.map(line => line.width), 0));
  const boxHeight = lines.reduce((sum, line) => sum + line.height, 0);
  paintChartLabelBox(ctx, titleBox, {
    x: x + (w - boxWidth) / 2,
    y,
    w: boxWidth,
    h: boxHeight,
  }, titlePtToPx);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';
  let lineY = y;
  for (const line of lines) {
    let pieceX = x + (w - line.width) / 2;
    for (const piece of line.pieces) {
      ctx.font = piece.font;
      ctx.fillStyle = piece.color;
      ctx.fillText(piece.text, pieceX, lineY);
      pieceX += piece.width;
    }
    lineY += line.height;
  }
  ctx.restore();
}

/** Draw the title at its authored manual-layout position. Office ignores w/h
 * for title descendants and fits the box to text (MS-OI29500 §2.1.1573), while
 * x/y still use the shared factor/edge rules. */
export function drawChartTitleForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  defaultY: number,
  fontSize: number,
): void {
  if (!chart.title) return;
  const ml = chart.titleManualLayout;
  if (ml) {
    const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
    const face = titleFace ? `"${titleFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
    ctx.font = chartFontCss(
      fontSize,
      face,
      chart.titleFontBold ?? true,
      chart.titleFontItalic ?? false,
    );
    const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), fontSize);
    const autoWidth = Math.min(w, Math.max(...lines.map(line => line.width), 0));
    const automatic = {
      x: x + (w - autoWidth) / 2,
      y: defaultY,
      w: autoWidth,
      h: fontSize,
    };
    const resolved = resolveManualLayoutRect(
      { ...ml, w: undefined, h: undefined },
      { x, y, w, h },
      automatic,
    );
    if (resolved) {
      drawChartTitle(ctx, chart, resolved.x, resolved.y, resolved.w, fontSize);
      return;
    }
  }
  drawChartTitle(ctx, chart, x, defaultY, w, fontSize);
}

// ─── Category helper ────────────────────────────────────────────────────────

export function chartCategories(chart: ChartModel): string[] {
  if (chart.categories.length > 0) return chart.categories;
  const first = chart.series[0];
  if (first?.categories && first.categories.length > 0) return first.categories;
  // ECMA-376 §21.2.2.24 — when <c:cat> is absent the category axis uses
  // integer values starting at 1. Fall back to the longest series so the
  // chart still renders instead of bailing out at n === 0.
  let n = 0;
  for (const s of chart.series) if (s.values.length > n) n = s.values.length;
  return n > 0 ? Array.from({ length: n }, (_, i) => String(i + 1)) : [];
}

export function dataLabelRectIntersection(a: DataLabelRect, b: DataLabelRect): DataLabelRect | null {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.w, b.x + b.w);
  const bottom = Math.min(a.y + a.h, b.y + b.h);
  return right > x && bottom > y ? { x, y, w: right - x, h: bottom - y } : null;
}

/** ECMA-376 §21.2.2.180: omission/false suppresses a label whose data-point value
 * is numerically greater than the effective value-axis maximum. This gate runs
 * after the shared axis planner has resolved authored and automatic bounds; it
 * never changes those bounds. For stacked charts the comparison remains the
 * point's authored value; the cumulative stack endpoint is layout geometry,
 * not the value represented by that label. */
export function dataLabelWithinAxisMaximum(
  chart: Pick<ChartModel, 'showDataLabelsOverMax'>,
  plottedValue: number,
  axisMaximum: number,
): boolean {
  return chart.showDataLabelsOverMax === true
    || !Number.isFinite(axisMaximum)
    || plottedValue <= axisMaximum;
}

/**
 * Draw a bar data label with the ECMA-376 §21.2.2.16 `dLblPos` semantics.
 *
 * For a vertical bar the coordinates describe the rectangle top-left + width +
 * height; for a horizontal bar they describe the bar's left-edge `bx`, top `by`,
 * length `barL`, and thickness `barW`. When `position` is "inBase" / "inEnd" /
 * "ctr" the label sits inside the bar; "outEnd" (default for clustered bars)
 * nudges the text just past the far edge. An explicit `color` overrides the
 * default dark label fill — Excel's workbook typically pairs "inBase" with a
 * white text color so labels stay readable against the bar fill.
 */
export function drawBarDataLabel(
  ctx: CanvasRenderingContext2D,
  text: string,
  bx: number, by: number, barL: number, barW: number,
  orient: 'vertical' | 'horizontal',
  position: string | null,
  color: string | null,
  fontSizePx: number,
  bounds: DataLabelRect,
  layoutReferenceRect: DataLabelRect,
  manualLayout?: ChartDataLabelOverride['manualLayout'],
  negative = false,
  rich?: RichDataLabelOptions,
  legendKey?: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  ptToPx = 1,
  labelBox?: ChartLabelBox,
  shapeRotationDeg = 0,
): void {
  const rect = orient === 'vertical'
    ? { x: bx, y: by, w: barW, h: barL }
    : { x: bx, y: by, w: barL, h: barW };
  drawBoundedDataLabelText(
    ctx,
    text,
    { kind: 'bar', rect, orientation: orient, negative, position: position ?? 'outEnd' },
    bounds,
    fontSizePx,
    color ? `#${color}` : '#333',
    manualLayout,
    layoutReferenceRect,
    rich,
    legendKey,
    textStyle,
    ptToPx,
    labelBox,
    shapeRotationDeg,
  );
}

export function applyDecorationLineStyle(
  ctx: CanvasRenderingContext2D,
  style: ChartDecorationLineStyle,
  ptToPx: number,
  bounds: ChartRect = {
    x: 0,
    y: 0,
    w: Math.max(1, ctx.canvas?.width ?? 1),
    h: Math.max(1, ctx.canvas?.height ?? 1),
  },
  shapeRotationDeg = 0,
): boolean {
  if (style.hidden === true
    || (style.paintAuthored === true && style.color == null && style.fill == null)) return false;
  const stroke = style.fill != null
    ? resolveFill(
        style.fill, ctx, bounds.x, bounds.y, bounds.w, bounds.h, shapeRotationDeg,
      )
    : style.color != null ? `#${style.color}` : '#000000';
  if (stroke == null) return false;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = style.widthEmu != null
    ? axisLineWidthPx(style.widthEmu, ptToPx)
    : Math.max(1, 0.75 * ptToPx);
  ctx.setLineDash(dashPatternForPreset(style.dash ?? undefined, ctx.lineWidth));
  ctx.lineCap = style.cap === 'rnd' ? 'round' : style.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = style.join === 'round' || style.join === 'bevel' ? style.join : 'miter';
  return true;
}

export function chartStyleRoleLine(
  chart: ChartModel,
  direct: ChartDecorationLineStyle,
  role: ChartStyleRole,
  compatibility?: ChartExElementStyle,
): ChartDecorationLineStyle {
  // A bounded Office compatibility recipe is a host/family default, not
  // authored OOXML. It therefore sits between the raw linked role and the
  // ECMA numeric role. `chartStyleRoles` is already linked-over-numeric and
  // cannot express that extra layer without recovering the retained sources.
  const linked = compatibility
    ? effectiveChartStyleRole(
        effectiveChartStyleRole(chart.classicChartStyleRoles?.[role], compatibility),
        chart.linkedChartStyleRoles?.[role] ?? (
          chart.classicChartStyleRoles == null ? chart.chartStyleRoles?.[role] : undefined
        ),
      )
    : chart.chartStyleRoles?.[role];
  const rawLinked = rawLinkedChartStyleRole(chart, role);
  const directPaintAuthored = direct.paintAuthored === true
    || direct.fill != null || direct.color != null || direct.hidden === true;
  const directStyleLine = chartStyleDirectLineDecision(direct.style, rawLinked, 0);
  const directNoLine = direct.hidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const lineDecision = directNoLine !== undefined ? directNoLine
    : direct.fill ?? (direct.color ? { fillType: 'solid' as const, color: direct.color }
      : directStyleLine !== undefined
        ? directStyleLine
        : direct.paintAuthored === true && direct.hidden !== true
        ? null
        : chartStyleLineCascade(linked, rawLinked, 0, direct.style));
  const dash = chartStyleDashChoice(
    direct.dash != null ? { lineDash: direct.dash, lineDashAuthored: true } : undefined,
    direct.style,
    linked,
  );
  return {
    style: chartStyleEffectOwner(direct.style, linked),
    fill: lineDecision != null && lineDecision.fillType !== 'solid'
      ? lineDecision : null,
    color: lineDecision?.fillType === 'solid' ? lineDecision.color : null,
    paintAuthored: directPaintAuthored
      ? direct.paintAuthored
      : lineDecision !== undefined ? true : undefined,
    widthEmu: direct.widthEmu ?? direct.style?.lineWidthEmu ?? linked?.lineWidthEmu ?? null,
    dash: dash?.lineDash ?? null,
    cap: direct.cap ?? direct.style?.lineCap ?? linked?.lineCap ?? null,
    join: direct.join ?? direct.style?.lineJoin ?? linked?.lineJoin ?? null,
    hidden: lineDecision === null ? true : null,
  };
}

export function chartStyleRoleBarPaint(
  chart: ChartModel,
  direct: ChartStockUpDownBarStyle['up'],
  role: 'upBar' | 'downBar',
  automaticPaint?: {
    lineColor: string;
    lineWidthEmu: number;
    upFillColor: string;
    downFillColor: string;
  },
): ChartStockUpDownBarStyle['up'] {
  const compatibility: ChartExElementStyle | undefined = automaticPaint ? {
    fillColors: [role === 'upBar' ? automaticPaint.upFillColor : automaticPaint.downFillColor],
    fillPaintAuthored: true,
    lineColors: [automaticPaint.lineColor],
    linePaintAuthored: true,
    lineWidthEmu: automaticPaint.lineWidthEmu,
  } : undefined;
  const linked = compatibility
    ? effectiveChartStyleRole(
        effectiveChartStyleRole(chart.classicChartStyleRoles?.[role], compatibility),
        chart.linkedChartStyleRoles?.[role] ?? (
          chart.classicChartStyleRoles == null ? chart.chartStyleRoles?.[role] : undefined
        ),
      )
    : chart.chartStyleRoles?.[role];
  const rawLinked = rawLinkedChartStyleRole(chart, role);
  const directFillAuthored = direct.fillPaintAuthored === true
    || direct.fillColor != null || direct.fill != null || direct.fillHidden === true;
  const directLineAuthored = direct.linePaintAuthored === true
    || direct.lineColor != null || direct.lineHidden === true;
  const directStyleFill = chartStyleDirectFillDecision(direct.style, rawLinked, 0);
  const directStyleLine = chartStyleDirectLineDecision(direct.style, rawLinked, 0);
  const directNoFill = direct.fillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directNoLine = direct.lineHidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const fillDecision = directNoFill !== undefined ? directNoFill
    : direct.fill != null ? direct.fill
    : direct.fillColor != null ? { fillType: 'solid' as const, color: direct.fillColor }
    : directStyleFill !== undefined
      ? directStyleFill
      : direct.fillPaintAuthored === true && direct.fillHidden !== true ? null
      : chartStyleFillCascade(linked, rawLinked, 0, direct.style);
  const lineDecision = directNoLine !== undefined ? directNoLine
    : direct.lineColor ? { fillType: 'solid' as const, color: direct.lineColor }
      : directStyleLine !== undefined
        ? directStyleLine
        : direct.linePaintAuthored === true && direct.lineHidden !== true
        ? null
        : chartStyleLineCascade(linked, rawLinked, 0, direct.style);
  return {
    style: chartStyleEffectOwner(direct.style, linked),
    fillColor: fillDecision?.fillType === 'solid' ? fillDecision.color : null,
    fill: fillDecision != null && fillDecision.fillType !== 'solid'
      && fillDecision.fillType !== 'none' ? fillDecision : null,
    fillPaintAuthored: directFillAuthored
      ? direct.fillPaintAuthored
      : fillDecision !== undefined ? true : undefined,
    fillHidden: fillDecision === null ? true : null,
    lineColor: lineDecision?.fillType === 'solid' ? lineDecision.color : null,
    linePaintAuthored: directLineAuthored
      ? direct.linePaintAuthored
      : lineDecision !== undefined ? true : undefined,
    lineWidthEmu: direct.lineWidthEmu ?? linked?.lineWidthEmu ?? null,
    lineDash: direct.lineDash ?? linked?.lineDash ?? null,
    lineCap: direct.lineCap ?? linked?.lineCap ?? null,
    lineJoin: direct.lineJoin ?? linked?.lineJoin ?? null,
    lineHidden: lineDecision === null ? true : null,
  };
}

/** Draws the one-per-category drop-line envelope shared by classic line,
 * area, and stock charts. ECMA-376 assigns the geometry to the owning chart
 * group; each envelope joins its effective category-axis crossing to every
 * finite plotted point at that category. */
export function drawDropLineEnvelopes(
  ctx: CanvasRenderingContext2D,
  members: ChartSeries[],
  pointCount: number,
  toX: (index: number) => number,
  yMapFor: (series: ChartSeries) => (value: number) => number,
  categoryAxisYFor: (series: ChartSeries) => number,
  valueFor: (series: ChartSeries, index: number) => number | null,
): void {
  for (let index = 0; index < pointCount; index++) {
    let minY = Infinity;
    let maxY = -Infinity;
    let hasPoint = false;
    for (const series of members) {
      const value = valueFor(series, index);
      if (value == null || !Number.isFinite(value)) continue;
      const pointY = yMapFor(series)(value);
      const axisY = categoryAxisYFor(series);
      if (!Number.isFinite(pointY) || !Number.isFinite(axisY)) continue;
      minY = Math.min(minY, pointY, axisY);
      maxY = Math.max(maxY, pointY, axisY);
      hasPoint = true;
    }
    if (!hasPoint || Math.abs(maxY - minY) < 0.01) continue;
    ctx.beginPath();
    ctx.moveTo(toX(index), minY);
    ctx.lineTo(toX(index), maxY);
    ctx.stroke();
  }
}

export function chartStyleRoleErrorBar(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['errBars']>[number],
): NonNullable<ChartSeries['errBars']>[number] {
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.color,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.dash,
    hidden: direct.hidden,
  }, 'errorBar');
  return {
    ...direct,
    color: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    dash: linked.dash ?? undefined,
    hidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}

export function chartStyleRoleLeaderLine(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    style: direct.leaderLineStyle,
    color: direct.leaderLineColor,
    paintAuthored: direct.leaderLinePaintAuthored,
    widthEmu: direct.leaderLineWidthEmu,
    dash: direct.leaderLineDash,
    hidden: direct.leaderLineHidden,
  }, 'leaderLine');
}

export function chartStyleRoleTrendline(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['trendLines']>[number],
): NonNullable<ChartSeries['trendLines']>[number] {
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.lineColor,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.lineDash,
    hidden: direct.lineHidden,
  }, 'trendline');
  return {
    ...direct,
    lineColor: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    lineDash: linked.dash ?? undefined,
    lineHidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}

export function chartStyleRoleDataTable(
  chart: ChartModel,
  direct: NonNullable<ChartModel['dataTable']>,
): NonNullable<ChartModel['dataTable']> {
  const role = chart.chartStyleRoles?.dataTable;
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataTable');
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.lineColor,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.lineDash,
    hidden: direct.lineHidden,
  }, 'dataTable');
  const directStyleFill = chartStyleDirectFillDecision(direct.style, rawLinked, 0);
  const directNoFill = direct.fillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const fillDecision = directNoFill !== undefined ? directNoFill
    : direct.fill ?? (direct.fillColor ? { fillType: 'solid' as const, color: direct.fillColor }
      : directStyleFill !== undefined
        ? directStyleFill
        : direct.fillPaintAuthored === true && direct.fillHidden !== true
        ? null
        : chartStyleFillCascade(role, rawLinked, 0, direct.style));
  const fill = fillDecision != null && fillDecision.fillType !== 'solid'
    && fillDecision.fillType !== 'image' && fillDecision.fillType !== 'none'
    ? fillDecision : null;
  const fillColor = fillDecision?.fillType === 'solid' ? fillDecision.color : null;
  const fillHidden = fillDecision === null ? true : null;
  const fillPaintAuthored = direct.fillPaintAuthored
    ?? (fillDecision !== undefined ? true : undefined);
  const textPaint = effectiveInheritedChartTextPaint(chart,
    direct.fontColor,
    direct.fontPaintAuthored,
    role,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? role?.fontSizeHpt,
    fontBold: direct.fontBold ?? chart.chartTextStyle?.fontBold ?? role?.fontBold,
    fontItalic: direct.fontItalic ?? chart.chartTextStyle?.fontItalic ?? role?.fontItalic,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: direct.fontPaintAuthored === true
      ? direct.fontHidden
      : role?.fontHidden,
    fontFace: direct.fontFace ?? chart.chartTextStyle?.fontFace ?? role?.fontFace,
    fill,
    fillColor,
    fillHidden,
    fillPaintAuthored,
    lineColor: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    lineDash: linked.dash ?? undefined,
    lineHidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}

export interface LinkedGridlineResult {
  visible: boolean | null | undefined;
  color?: string | null;
  widthEmu?: number | null;
  dash?: string | null;
  paintAuthored?: boolean | null;
}

export function chartStyleRoleGridline(
  chart: ChartModel,
  role: 'gridlineMajor' | 'gridlineMinor',
  visible: boolean | null | undefined,
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  paintAuthored: boolean | null | undefined,
  directStyle?: ChartExStyle | null,
): LinkedGridlineResult {
  if (visible !== true) {
    return { visible, color, widthEmu, dash, paintAuthored };
  }
  const linked = chartStyleRoleLine(
    chart, { style: directStyle, color, widthEmu, dash, paintAuthored }, role,
  );
  return {
    visible: linked.hidden !== true
      && !(linked.paintAuthored === true && linked.color == null),
    color: linked.color,
    widthEmu: linked.widthEmu,
    dash: linked.dash,
    paintAuthored: linked.paintAuthored,
  };
}

export function chartStyleRoleSecondaryGridlines(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
): SecondaryValueAxis | null | undefined {
  if (!axis) return axis;
  if (!chart.chartStyleRoles?.gridlineMajor && !chart.chartStyleRoles?.gridlineMinor) return axis;
  const major = chartStyleRoleGridline(
    chart, 'gridlineMajor', axis.majorGridlines,
    axis.majorGridlineColor, axis.majorGridlineWidthEmu, axis.majorGridlineDash,
    axis.majorGridlinePaintAuthored,
    axis.majorGridlineStyle,
  );
  const minor = chartStyleRoleGridline(
    chart, 'gridlineMinor', axis.minorGridlines,
    axis.minorGridlineColor, axis.minorGridlineWidthEmu, axis.minorGridlineDash,
    axis.minorGridlinePaintAuthored,
    axis.minorGridlineStyle,
  );
  const changed = major.visible !== axis.majorGridlines
    || major.color !== axis.majorGridlineColor
    || major.widthEmu !== axis.majorGridlineWidthEmu
    || major.dash !== axis.majorGridlineDash
    || major.paintAuthored !== axis.majorGridlinePaintAuthored
    || minor.visible !== axis.minorGridlines
    || minor.color !== axis.minorGridlineColor
    || minor.widthEmu !== axis.minorGridlineWidthEmu
    || minor.dash !== axis.minorGridlineDash
    || minor.paintAuthored !== axis.minorGridlinePaintAuthored;
  return changed ? {
    ...axis,
    majorGridlines: major.visible ?? undefined,
    majorGridlineColor: major.color,
    majorGridlineWidthEmu: major.widthEmu,
    majorGridlineDash: major.dash,
    majorGridlinePaintAuthored: major.paintAuthored,
    minorGridlines: minor.visible ?? undefined,
    minorGridlineColor: minor.color,
    minorGridlineWidthEmu: minor.widthEmu,
    minorGridlineDash: minor.dash,
    minorGridlinePaintAuthored: minor.paintAuthored,
  } : axis;
}

export function chartStyleRoleAxisLine(
  chart: ChartModel,
  role: 'categoryAxis' | 'valueAxis',
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  hidden: boolean,
  paintAuthored: boolean | null | undefined,
  directStyle?: ChartExStyle | null,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    style: directStyle,
    color,
    widthEmu,
    dash,
    paintAuthored,
    // The shared axis model stores the effective boolean, so false means no
    // direct noFill rather than an authored visible override.
    hidden: hidden ? true : undefined,
  }, role);
}

/** The flat axis model can carry a resolved solid color but not an arbitrary
 * DrawingML stroke paint. Preserve authored ownership by suppressing an
 * unresolved/structured paint here instead of reviving the semantic black
 * fallback in `resolveAxisLine`. */
export function chartAxisLineIsHidden(line: ChartDecorationLineStyle): boolean {
  return line.hidden === true || (line.paintAuthored === true && line.color == null);
}

export function chartStyleRoleSecondaryAxisLine(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
  role: 'categoryAxis' | 'valueAxis',
): SecondaryValueAxis | null | undefined {
  const style = chart.chartStyleRoles?.[role];
  if (!axis || (!style && !chart.chartTextStyle)) return axis;
  const line = chartStyleRoleAxisLine(
    chart, role, axis.lineColor, axis.lineWidthEmu, axis.lineDash, axis.lineHidden,
    axis.linePaintAuthored, axis.style,
  );
  const lineHidden = chartAxisLineIsHidden(line);
  const fontSizeHpt = axis.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? style?.fontSizeHpt;
  const fontBold = axis.fontBold ?? chart.chartTextStyle?.fontBold ?? style?.fontBold;
  const fontItalic = axis.fontItalic ?? chart.chartTextStyle?.fontItalic ?? style?.fontItalic;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, axis.fontColor, axis.fontPaintAuthored, style,
  );
  const fontColor = textPaint.color;
  const fontFace = axis.fontFace ?? chart.chartTextStyle?.fontFace ?? style?.fontFace;
  const titleStyle = chart.chartStyleRoles?.axisTitle;
  const titleFontSizeHpt = axis.titleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? titleStyle?.fontSizeHpt;
  const titleFontBold = axis.titleFontBold ?? chart.chartTextStyle?.fontBold ?? titleStyle?.fontBold;
  const titleFontItalic = axis.titleFontItalic
    ?? chart.chartTextStyle?.fontItalic ?? titleStyle?.fontItalic;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart, axis.titleFontColor, axis.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color;
  const titleFontFace = axis.titleFontFace
    ?? chart.chartTextStyle?.fontFace ?? titleStyle?.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && textPaint.authored === axis.fontPaintAuthored
    && fontFace === axis.fontFace
    && titleFontSizeHpt === axis.titleFontSizeHpt
    && titleFontBold === axis.titleFontBold
    && titleFontItalic === axis.titleFontItalic
    && titleFontColor === axis.titleFontColor
    && titleTextPaint.authored === axis.titleFontPaintAuthored
    && titleFontFace === axis.titleFontFace) return axis;
  return {
    ...axis,
    lineColor: line.color,
    lineWidthEmu: line.widthEmu,
    lineDash: line.dash,
    linePaintAuthored: line.paintAuthored,
    lineHidden,
    fontSizeHpt,
    fontBold,
    fontItalic,
    fontColor,
    fontPaintAuthored: textPaint.authored,
    fontFace,
    titleFontSizeHpt,
    titleFontBold,
    titleFontItalic,
    titleFontColor,
    titleFontPaintAuthored: titleTextPaint.authored,
    titleFontFace,
  };
}

export function chartStyleRoleSeriesAxis(chart: ChartModel): ChartModel {
  const axis = chart.threeD?.seriesAxis;
  const style = chart.chartStyleRoles?.seriesAxis;
  if (!axis || (!style && !chart.chartTextStyle)) return chart;
  const line = chartStyleRoleLine(chart, {
    style: axis.style,
    color: axis.lineColor,
    widthEmu: axis.lineWidthEmu,
    dash: axis.lineDash,
    hidden: axis.lineHidden ? true : undefined,
    paintAuthored: axis.linePaintAuthored,
  }, 'seriesAxis');
  const lineHidden = chartAxisLineIsHidden(line);
  const fontSizeHpt = axis.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? style?.fontSizeHpt;
  const fontBold = axis.fontBold ?? chart.chartTextStyle?.fontBold ?? style?.fontBold;
  const fontItalic = axis.fontItalic ?? chart.chartTextStyle?.fontItalic ?? style?.fontItalic;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, axis.fontColor, axis.fontPaintAuthored, style,
  );
  const fontColor = textPaint.color;
  const fontFace = axis.fontFace ?? chart.chartTextStyle?.fontFace ?? style?.fontFace;
  const titleStyle = chart.chartStyleRoles?.axisTitle;
  const titleFontSizeHpt = axis.titleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? titleStyle?.fontSizeHpt;
  const titleFontBold = axis.titleFontBold ?? chart.chartTextStyle?.fontBold ?? titleStyle?.fontBold;
  const titleFontItalic = axis.titleFontItalic
    ?? chart.chartTextStyle?.fontItalic ?? titleStyle?.fontItalic;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart, axis.titleFontColor, axis.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color;
  const titleFontFace = axis.titleFontFace
    ?? chart.chartTextStyle?.fontFace ?? titleStyle?.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && textPaint.authored === axis.fontPaintAuthored
    && fontFace === axis.fontFace
    && titleFontSizeHpt === axis.titleFontSizeHpt
    && titleFontBold === axis.titleFontBold
    && titleFontItalic === axis.titleFontItalic
    && titleFontColor === axis.titleFontColor
    && titleTextPaint.authored === axis.titleFontPaintAuthored
    && titleFontFace === axis.titleFontFace) return chart;
  return {
    ...chart,
    threeD: {
      ...chart.threeD,
      seriesAxis: {
        ...axis,
        lineColor: line.color,
        lineWidthEmu: line.widthEmu,
        lineDash: line.dash,
        lineHidden,
        fontSizeHpt,
        fontBold,
        fontItalic,
        fontColor,
        fontPaintAuthored: textPaint.authored,
        fontFace,
        titleFontSizeHpt,
        titleFontBold,
        titleFontItalic,
        titleFontColor,
        titleFontPaintAuthored: titleTextPaint.authored,
        titleFontFace,
      },
    },
  };
}

export function isClassicMarkerSeries(
  chart: ChartModel,
  series: ChartSeries,
  group: NonNullable<ChartModel['plotGroups']>[number] | undefined,
): boolean {
  if (group?.kind === 'bubble' || (group == null && chart.chartType === 'bubble')) return false;
  const family = group?.kind === 'scatter' ? 'scatter' : series.seriesType ?? chart.chartType;
  return family === 'line'
    || family === 'stackedLine'
    || family === 'stackedLinePct'
    || family === 'area'
    || family === 'stackedArea'
    || family === 'stackedAreaPct'
    || family === 'scatter'
    || family === 'radar'
    || family === 'stock';
}

export function chartStyleRoleMarker(
  chart: ChartModel,
  direct: ChartSeries,
  index: number,
  count: number,
  group: NonNullable<ChartModel['plotGroups']>[number] | undefined,
): ChartSeries {
  void count;
  // A point-varying group uses a separate formatting-index domain. Do not
  // flatten that role into series defaults here: drawChartMarker selects the
  // point role at paint time so every marker keeps its own index.
  const linked = chartSeriesVariesByPoint(chart, index)
    ? undefined
    : chart.chartStyleRoles?.dataPointMarker;
  const rawLinked = chartSeriesVariesByPoint(chart, index)
    ? undefined
    : rawLinkedChartStyleRole(chart, 'dataPointMarker');
  if (!isClassicMarkerSeries(chart, direct, group)
    || ((direct.showMarker === false || direct.markerSymbol === 'none')
      && !hasVisiblePointMarkerOverride(direct))) return direct;
  const styleIndex = chartExSeriesFormatIndex(direct, index);
  const directStyleFill = chartStyleDirectFillDecision(
    direct.markerStyle, rawLinked, styleIndex,
  );
  const directFillAuthored = direct.markerFillPaintAuthored === true
      && direct.markerStyle?.fillHidden !== true
    || direct.markerFill != null || direct.markerFillPaint !== undefined
    || directStyleFill !== undefined;
  const effectiveFill = directFillAuthored
    ? directStyleFill
    : chartStyleFillCascade(linked, rawLinked, styleIndex, direct.markerStyle);
  const markerFill = direct.markerFill
    ?? (effectiveFill?.fillType === 'solid' ? effectiveFill.color
      : effectiveFill === null ? '00000000' : null);
  const markerFillPaint = direct.markerFillPaint !== undefined
    ? direct.markerFillPaint
    : effectiveFill?.fillType === 'gradient'
        || effectiveFill?.fillType === 'pattern'
        || effectiveFill?.fillType === 'image'
      ? effectiveFill
      : undefined;
  const markerFillPaintAuthored = directFillAuthored
    ? direct.markerFillPaintAuthored
    : effectiveFill !== undefined
      ? true
      : undefined;
  const directStyleLine = chartStyleDirectLineDecision(
    direct.markerStyle, rawLinked, styleIndex,
  );
  const directLineAuthored = direct.markerLine != null
    || direct.markerLinePaintAuthored === true && direct.markerStyle?.lineHidden !== true
    || directStyleLine !== undefined;
  const effectiveLine = directLineAuthored
    ? directStyleLine
    : chartStyleLineCascade(linked, rawLinked, styleIndex, direct.markerStyle);
  const markerLine = direct.markerLine
    ?? (effectiveLine?.fillType === 'solid' ? effectiveLine.color
      : effectiveLine === null ? '00000000'
        : directLineAuthored ? '00000000'
        : null);
  const markerLineWidthEmu = direct.markerLineWidthEmu ?? direct.markerStyle?.lineWidthEmu
    ?? linked?.lineWidthEmu ?? null;
  const markerSize = direct.markerSize ?? chart.chartStyleMarkerSizePt;
  const markerSymbol = direct.markerSymbol ?? chart.chartStyleMarkerSymbol;
  const dataPointOverrides = direct.dataPointOverrides?.map(point => {
    if (!pointHasMarkerDetail(point)) return point;
    const directPointStyleFill = chartStyleDirectFillDecision(
      point.markerStyle, rawLinked, point.idx,
    );
    const directPointStyleLine = chartStyleDirectLineDecision(
      point.markerStyle, rawLinked, point.idx,
    );
    const pointFillAuthored = point.markerFillPaintAuthored === true
        && point.markerStyle?.fillHidden !== true
      || point.markerFill != null || point.markerFillPaint !== undefined
      || directPointStyleFill !== undefined;
    const pointLineAuthored = point.markerLine != null
      || point.markerLinePaintAuthored === true && point.markerStyle?.lineHidden !== true
      || directPointStyleLine !== undefined;
    const linkedPointFill = !pointFillAuthored
      ? directFillAuthored ? effectiveFill
          : chartStyleFillCascade(linked, rawLinked, styleIndex, direct.markerStyle)
      : undefined;
    const nextFill = point.markerFill
      ?? (directPointStyleFill?.fillType === 'solid' ? directPointStyleFill.color
        : directPointStyleFill === null ? '00000000' : undefined)
      ?? (linkedPointFill?.fillType === 'solid' ? linkedPointFill.color
        : linkedPointFill === null ? '00000000' : undefined);
    const nextFillPaint = point.markerFillPaint !== undefined
      ? point.markerFillPaint
      : directPointStyleFill?.fillType === 'gradient'
          || directPointStyleFill?.fillType === 'pattern'
          || directPointStyleFill?.fillType === 'image'
        ? directPointStyleFill
      : linkedPointFill?.fillType === 'gradient'
          || linkedPointFill?.fillType === 'pattern'
          || linkedPointFill?.fillType === 'image'
        ? linkedPointFill
        : undefined;
    const nextLine = point.markerLine
      ?? (directPointStyleLine?.fillType === 'solid' ? directPointStyleLine.color
        : directPointStyleLine === null ? '00000000' : undefined)
      ?? (pointLineAuthored ? '00000000'
        : effectiveLine?.fillType === 'solid' ? effectiveLine.color
        : effectiveLine === null ? '00000000'
        : undefined);
    const nextLineWidth = point.markerLineWidthEmu ?? point.markerStyle?.lineWidthEmu
      ?? (!pointLineAuthored && !directLineAuthored
        ? linked?.lineWidthEmu ?? undefined : undefined);
    if (nextFill === point.markerFill
      && nextFillPaint === point.markerFillPaint
      && nextLine === point.markerLine
      && nextLineWidth === point.markerLineWidthEmu) return point;
    return {
      ...point,
      markerFill: nextFill,
      markerFillPaint: nextFillPaint,
      markerFillPaintAuthored: point.markerFillPaintAuthored
        ?? (linkedPointFill !== undefined ? true : undefined),
      markerLine: nextLine,
      markerLinePaintAuthored: pointLineAuthored
        ? point.markerLinePaintAuthored
        : undefined,
      markerLineWidthEmu: nextLineWidth,
    };
  });
  if (markerFill === direct.markerFill
    && markerFillPaint === direct.markerFillPaint
    && markerFillPaintAuthored === direct.markerFillPaintAuthored
    && markerLine === direct.markerLine
    && markerLineWidthEmu === direct.markerLineWidthEmu
    && markerSize === direct.markerSize
    && markerSymbol === direct.markerSymbol
    && dataPointOverrides?.every((point, pointIndex) =>
      point === direct.dataPointOverrides?.[pointIndex]
    ) !== false) return direct;
  return {
    ...direct,
    markerFill,
    markerFillPaint,
    markerFillPaintAuthored,
    markerLine,
    markerLinePaintAuthored: directLineAuthored
      ? direct.markerLinePaintAuthored
      : undefined,
    markerLineWidthEmu,
    markerSize,
    markerSymbol,
    dataPointOverrides,
  };
}

export interface EffectiveFrameLineStyle {
  style?: ChartExStyle | null;
  color?: string | null;
  fill?: ChartModel['plotAreaLineFill'];
  widthEmu?: number | null;
  dash?: string | null;
  dashAuthored?: boolean | null;
  customDash?: ChartModel['plotAreaLineCustomDash'];
  cap?: string | null;
  join?: string | null;
  compound?: string | null;
  hidden?: boolean | null;
  paintAuthored?: boolean | null;
}

/** Merge one chart-frame outline property-by-property. Direct DrawingML paint
 * and dash choices remain authoritative; linked Chart Style geometry fills
 * only genuinely omitted properties. */
export function effectiveFrameLineStyle(
  chart: ChartModel,
  direct: EffectiveFrameLineStyle,
  linked: ChartExStyle | null | undefined,
  rawLinked: ChartExStyle | null | undefined,
  directIndex = 0,
  linkedIndex = directIndex,
): EffectiveFrameLineStyle {
  void chart;
  if (!linked) return direct;
  let { color, fill, hidden } = direct;
  const directNoLine = hidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const directStyleLine = chartStyleDirectLineDecision(
    direct.style, rawLinked, directIndex,
  );
  const directPaint = fill != null || color != null
    || directNoLine !== undefined || directStyleLine !== undefined
    || direct.paintAuthored === true && hidden !== true;
  const linkedPaint = linked.lineNoStyle !== true && (linked.linePaintAuthored === true
    || linked.lineHidden === true || linked.linePaints != null || linked.lineColors != null);
  if (!directPaint) {
    const decision = chartStyleLineDecision(linked, linkedIndex);
    if (decision === null) {
      hidden = true;
    } else if (decision?.fillType === 'solid') {
      color = decision.color;
      fill = null;
      hidden = null;
    } else if (decision !== undefined) {
      fill = decision;
      color = null;
      hidden = null;
    }
  } else if (directNoLine !== undefined) {
    hidden = true;
    color = null;
    fill = null;
  } else if (fill != null || color != null) {
    hidden = null;
  } else if (fill == null && color == null) {
    const decision = directStyleLine !== undefined
      ? directStyleLine
      : direct.paintAuthored === true ? null : undefined;
    if (decision === null) hidden = true;
    else if (decision?.fillType === 'solid') {
      color = decision.color;
      fill = null;
      hidden = null;
    } else if (decision !== undefined) {
      fill = decision;
      color = null;
      hidden = null;
    }
  }
  let dash = direct.dash;
  let customDash = direct.customDash;
  let dashAuthored = direct.dashAuthored;
  if (dashAuthored !== true && dash == null && customDash == null) {
    dash = linked.lineDash;
    customDash = linked.lineCustomDash;
    dashAuthored = linked.lineDashAuthored;
  }
  return {
    color,
    fill,
    hidden,
    paintAuthored: directPaint ? true : linkedPaint ? true : direct.paintAuthored,
    widthEmu: direct.widthEmu ?? direct.style?.lineWidthEmu ?? linked.lineWidthEmu,
    dash,
    dashAuthored,
    customDash,
    cap: direct.cap ?? direct.style?.lineCap ?? linked.lineCap,
    join: direct.join ?? direct.style?.lineJoin ?? linked.lineJoin,
    compound: direct.compound ?? direct.style?.lineCompound ?? linked.lineCompound,
  };
}

export function effectiveLinkedLabelBox(
  chart: ChartModel,
  direct: ChartLabelBox | null | undefined,
  linked: ChartExStyle | null | undefined,
  rawLinked: ChartExStyle | null | undefined,
  createFromLinked: boolean,
  linkedIndex = 0,
): ChartLabelBox | undefined {
  if (!linked || (!direct && !createFromLinked)) return direct ?? undefined;
  const source = direct ?? {};
  const effectiveFill = effectiveChartLabelBoxFill(
    source, linked, rawLinked, createFromLinked, 0, linkedIndex,
  );
  const line = effectiveFrameLineStyle(chart, {
    style: source.style,
    color: source.borderColor,
    fill: source.borderFill,
    widthEmu: source.borderWidthEmu,
    dash: source.borderDash,
    dashAuthored: source.borderDashAuthored,
    customDash: source.borderCustomDash,
    cap: source.borderCap,
    join: source.borderJoin,
    compound: source.borderCompound,
    hidden: source.borderHidden,
    paintAuthored: source.borderPaintAuthored,
  }, linked, rawLinked, 0, linkedIndex);
  return {
    ...source,
    style: source.style,
    effectFallbackStyle: linked,
    effectStyleIndex: 0,
    effectFallbackIndex: linkedIndex,
    ...effectiveFill,
    borderColor: line.color ?? undefined,
    borderFill: (line.fill as ChartLabelBox['borderFill']) ?? undefined,
    borderWidthEmu: line.widthEmu ?? undefined,
    borderDash: line.dash ?? undefined,
    borderDashAuthored: line.dashAuthored ?? undefined,
    borderCustomDash: line.customDash ?? undefined,
    borderCap: line.cap ?? undefined,
    borderJoin: line.join ?? undefined,
    borderCompound: line.compound ?? undefined,
    borderHidden: line.hidden ?? undefined,
    borderPaintAuthored: line.paintAuthored ?? undefined,
  };
}

/** Merge two directly-authored label shapes property-by-property. The higher
 * precedence shape owns an authored paint/noFill choice even when that choice
 * cannot be resolved to a Canvas paint; omitted geometry continues to inherit
 * from the lower-precedence series/linked shape. */

export function chartStyleRoleDataLabels(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
  styleIndex: number,
): ChartSeriesDataLabels {
  // A transparent dLbls/spPr is ordinary label formatting, not a request for
  // Office's filled `dataLabelCallout` recipe. Select that role only when the
  // directly authored box itself has visible paint.
  const usesCalloutRole = chartLabelBoxHasVisiblePaint(direct.labelBox);
  const linked = usesCalloutRole
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  const rawLinked = usesCalloutRole
    ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
      ?? rawLinkedChartStyleRole(chart, 'dataLabel')
    : rawLinkedChartStyleRole(chart, 'dataLabel');
  // The role itself is a legitimate label-shape source. A paint-bearing
  // linked/numeric dataLabel role therefore materializes a box even when the
  // chart has no direct dLbls/spPr; an empty/no-style role still paints
  // nothing because the resulting carrier has no visible fill or outline.
  const labelBox = linked
    ? effectiveLinkedLabelBox(chart, direct.labelBox, linked, rawLinked, true, styleIndex)
    : direct.labelBox;
  const directFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, direct.fontColor, direct.fontPaintAuthored, linked, styleIndex,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic ?? undefined,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: directFontPaint ? direct.fontHidden
      : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? chart.chartTextStyle?.textRotation
      ?? linked?.textRotation ?? undefined,
    textWrap: direct.textWrap ?? chart.chartTextStyle?.textWrap ?? linked?.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? chart.chartTextStyle?.textVerticalAnchor
      ?? linked?.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? chart.chartTextStyle?.textVerticalMode
      ?? linked?.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? chart.chartTextStyle?.textLInsEmu
      ?? linked?.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? chart.chartTextStyle?.textTInsEmu
      ?? linked?.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? chart.chartTextStyle?.textRInsEmu
      ?? linked?.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? chart.chartTextStyle?.textBInsEmu
      ?? linked?.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
    labelBox,
  };
}

export function chartStyleRoleTrendlineLabel(
  chart: ChartModel,
  direct: ChartTrendline,
  styleIndex: number,
): ChartTrendline {
  const linked = chart.chartStyleRoles?.trendlineLabel;
  const rawLinked = rawLinkedChartStyleRole(chart, 'trendlineLabel');
  const directFontPaint = direct.labelFontPaintAuthored === true
    || direct.labelFontColor != null || direct.labelFontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, direct.labelFontColor, direct.labelFontPaintAuthored, linked, styleIndex,
  );
  return {
    ...direct,
    // Unlike `dataLabelCallout`, the `trendlineLabel` role styles the generated
    // equation/R² label shape even when the chart does not carry a local spPr.
    // Materialize it before the chart-wide paint preflight so linked gradient
    // work is charged before any family starts painting.
    labelBox: linked
      ? effectiveLinkedLabelBox(chart, direct.labelBox, linked, rawLinked, true, styleIndex)
      : direct.labelBox,
    labelFontSizeHpt: direct.labelFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    labelFontBold: direct.labelFontBold ?? chart.chartTextStyle?.fontBold
      ?? linked?.fontBold ?? undefined,
    labelFontItalic: direct.labelFontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic ?? undefined,
    labelFontColor: textPaint.color ?? undefined,
    labelFontPaintAuthored: textPaint.authored,
    labelFontHidden: directFontPaint ? direct.labelFontHidden
      : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    labelFontFace: direct.labelFontFace ?? chart.chartTextStyle?.fontFace
      ?? linked?.fontFace ?? undefined,
    labelFontLanguage: direct.labelFontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage ?? undefined,
    labelFontBaseline: direct.labelFontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline ?? undefined,
    labelTextRotation: direct.labelTextRotation ?? chart.chartTextStyle?.textRotation
      ?? linked?.textRotation ?? undefined,
    labelTextWrap: direct.labelTextWrap ?? chart.chartTextStyle?.textWrap
      ?? linked?.textWrap ?? undefined,
    labelTextVerticalAnchor: direct.labelTextVerticalAnchor
      ?? chart.chartTextStyle?.textVerticalAnchor
      ?? linked?.textVerticalAnchor ?? undefined,
    labelTextVerticalMode: direct.labelTextVerticalMode ?? chart.chartTextStyle?.textVerticalMode
      ?? linked?.textVerticalMode ?? undefined,
    labelTextLInsEmu: direct.labelTextLInsEmu ?? chart.chartTextStyle?.textLInsEmu
      ?? linked?.textLInsEmu ?? undefined,
    labelTextTInsEmu: direct.labelTextTInsEmu ?? chart.chartTextStyle?.textTInsEmu
      ?? linked?.textTInsEmu ?? undefined,
    labelTextRInsEmu: direct.labelTextRInsEmu ?? chart.chartTextStyle?.textRInsEmu
      ?? linked?.textRInsEmu ?? undefined,
    labelTextBInsEmu: direct.labelTextBInsEmu ?? chart.chartTextStyle?.textBInsEmu
      ?? linked?.textBInsEmu ?? undefined,
    labelTextBodyAuthored: direct.labelTextBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
  };
}

export function chartStyleRoleDataLabelOverride(
  chart: ChartModel,
  direct: ChartDataLabelOverride,
  seriesDirect: ChartSeriesDataLabels | null | undefined,
): ChartDataLabelOverride {
  // A visible directly-authored box opts into `dataLabelCallout`; a bare or
  // transparent spPr remains an ordinary label. Applying the callout recipe
  // merely because an indexed override exists invents a white box around
  // ordinary point labels in Office styles.
  const directAndSeriesBox = mergeChartLabelBoxes(
    direct.labelBox, seriesDirect?.labelBox,
  );
  const hasCalloutShape = chartLabelBoxHasVisiblePaint(directAndSeriesBox);
  const linked = hasCalloutShape
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  const rawLinked = hasCalloutShape
    ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
      ?? rawLinkedChartStyleRole(chart, 'dataLabel')
    : rawLinkedChartStyleRole(chart, 'dataLabel');
  const labelBox = linked
    ? effectiveLinkedLabelBox(
        chart, directAndSeriesBox, linked, rawLinked, true, direct.idx,
      )
    : directAndSeriesBox;
  const pointFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  const seriesFontPaint = seriesDirect?.fontPaintAuthored === true
    || seriesDirect?.fontColor != null || seriesDirect?.fontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart,
    pointFontPaint ? direct.fontColor : seriesFontPaint ? seriesDirect?.fontColor : undefined,
    pointFontPaint ? direct.fontPaintAuthored : seriesDirect?.fontPaintAuthored,
    linked,
    direct.idx,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? seriesDirect?.fontSizeHpt
      ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? seriesDirect?.fontBold
      ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? seriesDirect?.fontItalic
      ?? chart.chartTextStyle?.fontItalic ?? linked?.fontItalic ?? undefined,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: pointFontPaint
      ? direct.fontHidden
      : seriesFontPaint ? seriesDirect?.fontHidden
        : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? seriesDirect?.fontFace
      ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? seriesDirect?.fontLanguage
      ?? chart.chartTextStyle?.fontLanguage ?? linked?.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? seriesDirect?.fontBaseline
      ?? chart.chartTextStyle?.fontBaseline ?? linked?.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? seriesDirect?.textRotation
      ?? chart.chartTextStyle?.textRotation ?? linked?.textRotation ?? undefined,
    textWrap: direct.textWrap ?? seriesDirect?.textWrap
      ?? chart.chartTextStyle?.textWrap ?? linked?.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? seriesDirect?.textVerticalAnchor
      ?? chart.chartTextStyle?.textVerticalAnchor ?? linked?.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? seriesDirect?.textVerticalMode
      ?? chart.chartTextStyle?.textVerticalMode ?? linked?.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? seriesDirect?.textLInsEmu
      ?? chart.chartTextStyle?.textLInsEmu ?? linked?.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? seriesDirect?.textTInsEmu
      ?? chart.chartTextStyle?.textTInsEmu ?? linked?.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? seriesDirect?.textRInsEmu
      ?? chart.chartTextStyle?.textRInsEmu ?? linked?.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? seriesDirect?.textBInsEmu
      ?? chart.chartTextStyle?.textBInsEmu ?? linked?.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || seriesDirect?.textBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
    textAlign: direct.textAlign ?? seriesDirect?.textAlign,
    labelBox,
  };
}

export function chartStyleRoleLegend(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.legend;
  const rawLinked = rawLinkedChartStyleRole(chart, 'legend');
  if (!linked && !chart.chartTextStyle) return chart;
  let legendFill = chart.legendFill;
  let legendFillColor = chart.legendFillColor;
  let legendFillHidden = chart.legendFillHidden;
  let legendFillPaintAuthored = chart.legendFillPaintAuthored;
  const directNoFill = legendFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directFillPaint = legendFill != null || legendFillColor != null
    || directNoFill !== undefined
    || chart.legendFillPaintAuthored === true && legendFillHidden !== true;
  if (!directFillPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.legendStyle);
    if (decision === null) {
      legendFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      legendFillColor = decision.color;
      legendFill = null;
      legendFillHidden = null;
    } else if (decision !== undefined) {
      legendFill = decision;
      legendFillColor = null;
      legendFillHidden = null;
    }
    if (decision !== undefined) {
      legendFillPaintAuthored = true;
    }
  }

  const legendLine = effectiveFrameLineStyle(chart, {
    style: chart.legendStyle,
    color: chart.legendLineColor,
    fill: chart.legendLineFill,
    widthEmu: chart.legendLineWidthEmu,
    dash: chart.legendLineDash,
    dashAuthored: chart.legendLineDashAuthored,
    customDash: chart.legendLineCustomDash,
    cap: chart.legendLineCap,
    join: chart.legendLineJoin,
    compound: chart.legendLineCompound,
    hidden: chart.legendLineHidden,
    paintAuthored: chart.legendLinePaintAuthored,
  }, linked, rawLinkedChartStyleRole(chart, 'legend'));
  const legendTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.legendFontColor,
    chart.legendFontPaintAuthored,
    linked,
  );
  if (legendFill === chart.legendFill
    && legendFillColor === chart.legendFillColor
    && legendFillHidden === chart.legendFillHidden
    && legendFillPaintAuthored === chart.legendFillPaintAuthored
    && legendLine.color === chart.legendLineColor
    && legendLine.fill === chart.legendLineFill
    && legendLine.widthEmu === chart.legendLineWidthEmu
    && legendLine.dash === chart.legendLineDash
    && legendLine.dashAuthored === chart.legendLineDashAuthored
    && legendLine.customDash === chart.legendLineCustomDash
    && legendLine.cap === chart.legendLineCap
    && legendLine.join === chart.legendLineJoin
    && legendLine.compound === chart.legendLineCompound
    && legendLine.hidden === chart.legendLineHidden
    && legendLine.paintAuthored === chart.legendLinePaintAuthored
    && (chart.legendFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt) === chart.legendFontSizeHpt
    && (chart.legendFontBold ?? chart.chartTextStyle?.fontBold
      ?? linked?.fontBold) === chart.legendFontBold
    && (chart.legendFontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic) === chart.legendFontItalic
    && (chart.legendFontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage) === chart.legendFontLanguage
    && (chart.legendFontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline) === chart.legendFontBaseline
    && legendTextPaint.color === chart.legendFontColor
    && legendTextPaint.authored === chart.legendFontPaintAuthored
    && (chart.legendFontFace ?? chart.chartTextStyle?.fontFace
      ?? linked?.fontFace) === chart.legendFontFace) return chart;
  return {
    ...chart,
    legendFontSizeHpt: chart.legendFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt,
    legendFontBold: chart.legendFontBold ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold,
    legendFontItalic: chart.legendFontItalic
      ?? chart.chartTextStyle?.fontItalic ?? linked?.fontItalic,
    legendFontLanguage: chart.legendFontLanguage
      ?? chart.chartTextStyle?.fontLanguage ?? linked?.fontLanguage,
    legendFontBaseline: chart.legendFontBaseline
      ?? chart.chartTextStyle?.fontBaseline ?? linked?.fontBaseline,
    legendFontColor: legendTextPaint.color,
    legendFontPaintAuthored: legendTextPaint.authored,
    legendFontFace: chart.legendFontFace ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace,
    legendFill,
    legendFillColor,
    legendFillHidden,
    legendFillPaintAuthored,
    legendLineColor: legendLine.color,
    legendLineFill: legendLine.fill,
    legendLineWidthEmu: legendLine.widthEmu,
    legendLineDash: legendLine.dash,
    legendLineDashAuthored: legendLine.dashAuthored,
    legendLineCustomDash: legendLine.customDash,
    legendLineCap: legendLine.cap,
    legendLineJoin: legendLine.join,
    legendLineCompound: legendLine.compound,
    legendLineHidden: legendLine.hidden,
    legendLinePaintAuthored: legendLine.paintAuthored,
  };
}

/** Resolve a chart text paint as one atomic DrawingML component. Authored
 * noFill and paints that cannot be represented by Canvas become transparent;
 * neither may reveal a lower linked/numeric/default colour. */
export function effectiveChartTextPaint(
  directColor: string | null | undefined,
  directAuthored: boolean | null | undefined,
  linked: ChartExStyle | null | undefined,
  index = 0,
): { color: string | null | undefined; authored: boolean | undefined } {
  if (directAuthored === true || directColor != null) {
    return {
      color: directColor ?? '00000000',
      authored: true,
    };
  }
  if (linked && (linked.fontPaintAuthored === true
    || linked.fontColor != null || linked.fontColors != null || linked.fontHidden === true)) {
    return {
      color: linked.fontHidden === true
        ? '00000000'
        : chartStyleFontColor(linked, index) ?? '00000000',
      authored: true,
    };
  }
  return { color: directColor, authored: undefined };
}

/** Resolve the chart-wide txPr between element-local text and the style role.
 * Paint remains atomic so an authored noFill/unresolved chart default cannot
 * leak through to linked or numeric colors. */
export function effectiveInheritedChartTextPaint(
  chart: ChartModel,
  directColor: string | null | undefined,
  directAuthored: boolean | null | undefined,
  linked: ChartExStyle | null | undefined,
  index = 0,
): { color: string | null | undefined; authored: boolean | undefined } {
  const global = effectiveChartTextPaint(
    directColor, directAuthored, chart.chartTextStyle, index,
  );
  if (global.authored === true || global.color != null) return global;
  return effectiveChartTextPaint(directColor, directAuthored, linked, index);
}

export function chartStyleRolePlotArea(chart: ChartModel): ChartModel {
  // MS-ODRAWXML defines plotArea and plotArea3D as separate required style
  // entries. Do not infer one from the other in a malformed/partial sidecar;
  // direct chart formatting stays authoritative below.
  const linked = chart.threeD
    ? chart.chartStyleRoles?.plotArea3D
    : chart.chartStyleRoles?.plotArea;
  const rawLinked = rawLinkedChartStyleRole(
    chart, chart.threeD ? 'plotArea3D' : 'plotArea',
  );
  if (!linked) return chart;
  let plotAreaFill = chart.plotAreaFill;
  let plotAreaBg = chart.plotAreaBg;
  let plotAreaFillHidden = chart.plotAreaFillHidden;
  let plotAreaFillPaintAuthored = chart.plotAreaFillPaintAuthored;
  const directNoFill = plotAreaFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directPaint = ((plotAreaFill != null || plotAreaBg != null)
      && chart.plotAreaFillAutomatic !== true)
    || directNoFill !== undefined
    || chart.plotAreaFillPaintAuthored === true && plotAreaFillHidden !== true;
  if (!directPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.plotAreaStyle);
    if (decision === null) {
      plotAreaFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      plotAreaBg = decision.color;
      plotAreaFill = null;
      plotAreaFillHidden = null;
    } else if (decision !== undefined) {
      plotAreaFill = decision;
      plotAreaBg = null;
      plotAreaFillHidden = null;
    }
    if (decision !== undefined) {
      plotAreaFillPaintAuthored = true;
    }
  }

  const plotAreaLine = effectiveFrameLineStyle(chart, {
    style: chart.plotAreaStyle,
    color: chart.plotAreaLineColor,
    fill: chart.plotAreaLineFill,
    widthEmu: chart.plotAreaLineWidthEmu,
    dash: chart.plotAreaLineDash,
    dashAuthored: chart.plotAreaLineDashAuthored,
    customDash: chart.plotAreaLineCustomDash,
    cap: chart.plotAreaLineCap,
    join: chart.plotAreaLineJoin,
    compound: chart.plotAreaLineCompound,
    hidden: chart.plotAreaLineHidden,
    paintAuthored: chart.plotAreaLinePaintAuthored,
  }, linked, rawLinkedChartStyleRole(
    chart, chart.threeD ? 'plotArea3D' : 'plotArea',
  ));
  if (plotAreaFill === chart.plotAreaFill
    && plotAreaBg === chart.plotAreaBg
    && plotAreaFillHidden === chart.plotAreaFillHidden
    && plotAreaFillPaintAuthored === chart.plotAreaFillPaintAuthored
    && plotAreaLine.color === chart.plotAreaLineColor
    && plotAreaLine.fill === chart.plotAreaLineFill
    && plotAreaLine.widthEmu === chart.plotAreaLineWidthEmu
    && plotAreaLine.dash === chart.plotAreaLineDash
    && plotAreaLine.dashAuthored === chart.plotAreaLineDashAuthored
    && plotAreaLine.customDash === chart.plotAreaLineCustomDash
    && plotAreaLine.cap === chart.plotAreaLineCap
    && plotAreaLine.join === chart.plotAreaLineJoin
    && plotAreaLine.compound === chart.plotAreaLineCompound
    && plotAreaLine.hidden === chart.plotAreaLineHidden
    && plotAreaLine.paintAuthored === chart.plotAreaLinePaintAuthored) return chart;
  return {
    ...chart,
    plotAreaFill,
    plotAreaBg,
    plotAreaFillHidden,
    plotAreaFillPaintAuthored,
    plotAreaLineColor: plotAreaLine.color,
    plotAreaLineFill: plotAreaLine.fill,
    plotAreaLineWidthEmu: plotAreaLine.widthEmu,
    plotAreaLineDash: plotAreaLine.dash,
    plotAreaLineDashAuthored: plotAreaLine.dashAuthored,
    plotAreaLineCustomDash: plotAreaLine.customDash,
    plotAreaLineCap: plotAreaLine.cap,
    plotAreaLineJoin: plotAreaLine.join,
    plotAreaLineCompound: plotAreaLine.compound,
    plotAreaLineHidden: plotAreaLine.hidden,
    plotAreaLinePaintAuthored: plotAreaLine.paintAuthored,
  };
}

export function chartStyleRoleChartArea(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.chartArea;
  const rawLinked = rawLinkedChartStyleRole(chart, 'chartArea');
  if (!linked) return chart;
  let chartFill = chart.chartFill;
  let chartBg = chart.chartBg;
  let chartFillHidden = chart.chartFillHidden;
  let chartFillPaintAuthored = chart.chartFillPaintAuthored;
  const directNoFill = chartFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directPaint = chartFill != null || directNoFill !== undefined
    || chart.chartFillPaintAuthored === true && chartFillHidden !== true;
  if (!directPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.chartAreaStyle);
    if (decision === null) {
      chartFill = null;
      chartBg = null;
      chartFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      chartBg = decision.color;
      chartFill = null;
      chartFillHidden = null;
    } else if (decision !== undefined) {
      chartFill = decision;
      chartBg = null;
      chartFillHidden = null;
    }
    if (decision !== undefined) {
      chartFillPaintAuthored = true;
    }
  }

  const chartBorder = effectiveFrameLineStyle(chart, {
    style: chart.chartAreaStyle,
    color: chart.chartBorderColor,
    fill: chart.chartBorderLineFill,
    widthEmu: chart.chartBorderWidthEmu,
    dash: chart.chartBorderDash,
    dashAuthored: chart.chartBorderDashAuthored,
    customDash: chart.chartBorderCustomDash,
    cap: chart.chartBorderCap,
    join: chart.chartBorderJoin,
    compound: chart.chartBorderCompound,
    hidden: chart.chartBorderHidden,
    paintAuthored: chart.chartBorderPaintAuthored,
  }, linked, rawLinkedChartStyleRole(chart, 'chartArea'));
  if (chartFill === chart.chartFill
    && chartBg === chart.chartBg
    && chartFillHidden === chart.chartFillHidden
    && chartFillPaintAuthored === chart.chartFillPaintAuthored
    && chartBorder.color === chart.chartBorderColor
    && chartBorder.fill === chart.chartBorderLineFill
    && chartBorder.widthEmu === chart.chartBorderWidthEmu
    && chartBorder.dash === chart.chartBorderDash
    && chartBorder.dashAuthored === chart.chartBorderDashAuthored
    && chartBorder.customDash === chart.chartBorderCustomDash
    && chartBorder.cap === chart.chartBorderCap
    && chartBorder.join === chart.chartBorderJoin
    && chartBorder.compound === chart.chartBorderCompound
    && chartBorder.hidden === chart.chartBorderHidden
    && chartBorder.paintAuthored === chart.chartBorderPaintAuthored) return chart;
  return {
    ...chart,
    chartFill,
    chartBg,
    chartFillHidden,
    chartFillPaintAuthored,
    chartBorderColor: chartBorder.color,
    chartBorderLineFill: chartBorder.fill,
    chartBorderWidthEmu: chartBorder.widthEmu,
    chartBorderDash: chartBorder.dash,
    chartBorderDashAuthored: chartBorder.dashAuthored,
    chartBorderCustomDash: chartBorder.customDash,
    chartBorderCap: chartBorder.cap,
    chartBorderJoin: chartBorder.join,
    chartBorderCompound: chartBorder.compound,
    chartBorderHidden: chartBorder.hidden,
    chartBorderPaintAuthored: chartBorder.paintAuthored,
  };
}

/** Materialize the linked decoration roles that an optional family renderer
 * consumes directly from `ChartSeries`. Keeping this projection in core means
 * the 2-D, 3-D, DOCX, XLSX, and PPTX paths receive one effective precedence
 * result without teaching an optional renderer about package sidecars. */
export function withOfficeStyleRasterLineFloor(chart: ChartModel, ptToPx: number): ChartModel {
  const roles = chart.chartStyleRoles;
  if (!roles || !(Number.isFinite(ptToPx) && ptToPx > 0)) return chart;
  // Office keeps the authored/theme width in its vector output (for example,
  // classic Style 2 emits a 6,350 EMU / 0.5pt black rule), then stroke-adjusts
  // that vector to one opaque device pixel when an axis or gridline would
  // otherwise land below a pixel. Canvas instead alpha-antialiases the
  // subpixel stroke into a grey rule. Apply the observed raster floor only to
  // those evidenced style roles in this render projection: direct `<a:ln w>`
  // remains exact, other role families stay untouched, the source model keeps
  // its ECMA-376 width, and zoomed widths naturally exceed the floor.
  const minimumWidthEmu = EMU_PER_PT / ptToPx;
  let changed = false;
  const strokeAdjustedRoles = new Set<ChartStyleRole>([
    'categoryAxis', 'seriesAxis', 'valueAxis', 'gridlineMajor', 'gridlineMinor',
  ]);
  const adjusted = Object.fromEntries(Object.entries(roles).map(([role, style]) => {
    if (!strokeAdjustedRoles.has(role as ChartStyleRole)) return [role, style];
    if (!style || style.lineWidthEmu == null
      || !Number.isFinite(style.lineWidthEmu)
      || style.lineWidthEmu <= 0
      || style.lineWidthEmu >= minimumWidthEmu) return [role, style];
    changed = true;
    return [role, { ...style, lineWidthEmu: minimumWidthEmu }];
  })) as typeof roles;
  return changed ? { ...chart, chartStyleRoles: adjusted } : chart;
}

export function applyLinkedChartStyleRoles(chart: ChartModel, ptToPx: number): ChartModel {
  chart = withOfficeStyleRasterLineFloor(chart, ptToPx);
  if (!chart.chartStyleRoles?.errorBar
    && !chart.chartStyleRoles?.leaderLine
    && !chart.chartStyleRoles?.trendline
    && !chart.chartStyleRoles?.trendlineLabel
    && !chart.chartStyleRoles?.dataLabel
    && !chart.chartStyleRoles?.dataLabelCallout
    && !chart.chartStyleRoles?.dataTable
    && !chart.chartStyleRoles?.gridlineMajor
    && !chart.chartStyleRoles?.gridlineMinor
    && !chart.chartStyleRoles?.categoryAxis
    && !chart.chartStyleRoles?.valueAxis
    && !chart.chartStyleRoles?.seriesAxis
    && !chart.chartStyleRoles?.dataPointMarker
    && !chart.chartStyleRoles?.legend
    && !chart.chartStyleRoles?.plotArea
    && !chart.chartStyleRoles?.plotArea3D
    && !chart.chartStyleRoles?.chartArea
    && !chart.chartStyleRoles?.title
    && !chart.chartStyleRoles?.axisTitle
    && chart.chartTextStyle == null
    && chart.chartStyleMarkerSizePt == null
    && chart.chartStyleMarkerSymbol == null) {
    return chart;
  }
  let changed = false;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const series = chart.series.map((sourceItem, seriesIndex) => {
    const item = chartStyleRoleMarker(
      chart, sourceItem, seriesIndex, chart.series.length, plotGroupBySeries[seriesIndex],
    );
    changed ||= item !== sourceItem;
    const errBars = chart.chartStyleRoles?.errorBar ? item.errBars?.map(errorBar => {
      const effective = chartStyleRoleErrorBar(chart, errorBar);
      changed ||= effective.color !== errorBar.color
        || effective.lineWidthEmu !== errorBar.lineWidthEmu
        || effective.dash !== errorBar.dash
        || effective.hidden !== errorBar.hidden;
      return effective;
    }) : item.errBars;
    let seriesDataLabels = item.seriesDataLabels;
    if (seriesDataLabels
      && (chart.chartTextStyle
        || chart.chartStyleRoles?.dataLabel || chart.chartStyleRoles?.dataLabelCallout)) {
      const effective = chartStyleRoleDataLabels(
        chart,
        seriesDataLabels,
        chartExSeriesFormatIndex(item, seriesIndex),
      );
      changed ||= effective !== seriesDataLabels;
      seriesDataLabels = effective;
    }
    const dataLabelOverrides = (chart.chartTextStyle || chart.chartStyleRoles?.dataLabelCallout
      || chart.chartStyleRoles?.dataLabel)
      ? item.dataLabelOverrides?.map(override => {
          const effective = chartStyleRoleDataLabelOverride(
            chart,
            override,
            sourceItem.seriesDataLabels,
          );
          changed ||= effective !== override;
          return effective;
        })
      : item.dataLabelOverrides;
    if (seriesDataLabels && chart.chartStyleRoles?.leaderLine) {
      const effective = chartStyleRoleLeaderLine(chart, seriesDataLabels);
      const merged = {
        ...seriesDataLabels,
        leaderLineColor: effective.color ?? undefined,
        leaderLineWidthEmu: effective.widthEmu ?? undefined,
        leaderLineDash: effective.dash ?? undefined,
        leaderLineHidden: effective.hidden ?? undefined,
        leaderLinePaintAuthored: effective.paintAuthored,
      };
      changed ||= merged.leaderLineColor !== seriesDataLabels.leaderLineColor
        || merged.leaderLineWidthEmu !== seriesDataLabels.leaderLineWidthEmu
        || merged.leaderLineDash !== seriesDataLabels.leaderLineDash
        || merged.leaderLineHidden !== seriesDataLabels.leaderLineHidden
        || merged.leaderLinePaintAuthored !== seriesDataLabels.leaderLinePaintAuthored;
      seriesDataLabels = merged;
    }
    const trendLines = (chart.chartStyleRoles?.trendline || chart.chartTextStyle
      || chart.chartStyleRoles?.trendlineLabel)
      ? item.trendLines?.map(trendline => {
      let effective = chart.chartStyleRoles?.trendline
        ? chartStyleRoleTrendline(chart, trendline)
        : trendline;
      if (chart.chartTextStyle || chart.chartStyleRoles?.trendlineLabel) {
        effective = chartStyleRoleTrendlineLabel(
          chart,
          effective,
          chartExSeriesFormatIndex(item, seriesIndex),
        );
      }
      changed ||= effective.lineColor !== trendline.lineColor
        || effective.lineWidthEmu !== trendline.lineWidthEmu
        || effective.lineDash !== trendline.lineDash
        || effective.lineHidden !== trendline.lineHidden
        || effective !== trendline;
      return effective;
    }) : item.trendLines;
    if (errBars === item.errBars
      && seriesDataLabels === item.seriesDataLabels
      && dataLabelOverrides === item.dataLabelOverrides
      && trendLines === item.trendLines) return item;
    return { ...item, errBars, seriesDataLabels, dataLabelOverrides, trendLines };
  });
  let dataTable = chart.dataTable;
  if (dataTable && (chart.chartStyleRoles?.dataTable || chart.chartTextStyle)) {
    const effective = chartStyleRoleDataTable(chart, dataTable);
    changed ||= effective !== dataTable;
    dataTable = effective;
  }
  const valMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.valAxisMajorGridlines,
    chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, chart.valAxisGridlineDash,
    chart.valAxisGridlinePaintAuthored,
    chart.valAxisMajorGridlineStyle,
  );
  const catMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.catAxisMajorGridlines,
    chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, chart.catAxisGridlineDash,
    chart.catAxisGridlinePaintAuthored,
    chart.catAxisMajorGridlineStyle,
  );
  const valMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.valAxisMinorGridlines,
    chart.valAxisMinorGridlineColor,
    chart.valAxisMinorGridlineWidthEmu,
    chart.valAxisMinorGridlineDash,
    chart.valAxisMinorGridlinePaintAuthored,
    chart.valAxisMinorGridlineStyle,
  );
  const catMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.catAxisMinorGridlines,
    chart.catAxisMinorGridlineColor,
    chart.catAxisMinorGridlineWidthEmu,
    chart.catAxisMinorGridlineDash,
    chart.catAxisMinorGridlinePaintAuthored,
    chart.catAxisMinorGridlineStyle,
  );
  const secondaryValGridlines = chartStyleRoleSecondaryGridlines(chart, chart.secondaryValAxis);
  const secondaryCatGridlines = chartStyleRoleSecondaryGridlines(chart, chart.secondaryCatAxis);
  const secondaryValAxis = chartStyleRoleSecondaryAxisLine(
    chart, secondaryValGridlines, 'valueAxis',
  );
  const secondaryCatAxis = chartStyleRoleSecondaryAxisLine(
    chart, secondaryCatGridlines, 'categoryAxis',
  );
  const catAxisLine = chartStyleRoleAxisLine(
    chart, 'categoryAxis',
    chart.catAxisLineColor, chart.catAxisLineWidthEmu, chart.catAxisLineDash,
    chart.catAxisLineHidden,
    chart.catAxisLinePaintAuthored,
    chart.catAxisStyle,
  );
  const valAxisLine = chartStyleRoleAxisLine(
    chart, 'valueAxis',
    chart.valAxisLineColor, chart.valAxisLineWidthEmu, chart.valAxisLineDash,
    chart.valAxisLineHidden,
    chart.valAxisLinePaintAuthored,
    chart.valAxisStyle,
  );
  const catAxisLineHidden = chartAxisLineIsHidden(catAxisLine);
  const valAxisLineHidden = chartAxisLineIsHidden(valAxisLine);
  const catAxisStyle = chart.chartStyleRoles?.categoryAxis;
  const valAxisStyle = chart.chartStyleRoles?.valueAxis;
  const titleStyle = chart.chartStyleRoles?.title;
  const axisTitleStyle = chart.chartStyleRoles?.axisTitle;
  const dataLabelStyle = chart.chartStyleRoles?.dataLabel;
  const catAxisFontSizeHpt = chart.catAxisFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? catAxisStyle?.fontSizeHpt ?? null;
  const catAxisFontBold = chart.catAxisFontBold ?? chart.chartTextStyle?.fontBold
    ?? catAxisStyle?.fontBold;
  const catAxisFontItalic = chart.catAxisFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? catAxisStyle?.fontItalic;
  const catAxisTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.catAxisFontColor, chart.catAxisFontPaintAuthored, catAxisStyle,
  );
  const catAxisFontColor = catAxisTextPaint.color;
  const catAxisFontFace = chart.catAxisFontFace ?? chart.chartTextStyle?.fontFace
    ?? catAxisStyle?.fontFace;
  const valAxisFontSizeHpt = chart.valAxisFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? valAxisStyle?.fontSizeHpt ?? null;
  const valAxisFontBold = chart.valAxisFontBold ?? chart.chartTextStyle?.fontBold
    ?? valAxisStyle?.fontBold;
  const valAxisFontItalic = chart.valAxisFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? valAxisStyle?.fontItalic;
  const valAxisTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.valAxisFontColor, chart.valAxisFontPaintAuthored, valAxisStyle,
  );
  const valAxisFontColor = valAxisTextPaint.color;
  const valAxisFontFace = chart.valAxisFontFace ?? chart.chartTextStyle?.fontFace
    ?? valAxisStyle?.fontFace;
  const titleFontSizeHpt = chart.titleFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? titleStyle?.fontSizeHpt ?? null;
  const titleFontBold = chart.titleFontBold ?? chart.chartTextStyle?.fontBold
    ?? titleStyle?.fontBold;
  const titleFontItalic = chart.titleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? titleStyle?.fontItalic;
  const titleFontLanguage = chart.titleFontLanguage ?? chart.chartTextStyle?.fontLanguage
    ?? titleStyle?.fontLanguage;
  const titleFontBaseline = chart.titleFontBaseline ?? chart.chartTextStyle?.fontBaseline
    ?? titleStyle?.fontBaseline;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.titleFontColor, chart.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color ?? null;
  const titleFontFace = chart.titleFontFace ?? chart.chartTextStyle?.fontFace
    ?? titleStyle?.fontFace ?? null;
  const catAxisTitleFontSizeHpt = chart.catAxisTitleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? axisTitleStyle?.fontSizeHpt;
  const catAxisTitleFontBold = chart.catAxisTitleFontBold ?? chart.chartTextStyle?.fontBold
    ?? axisTitleStyle?.fontBold;
  const catAxisTitleFontItalic = chart.catAxisTitleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? axisTitleStyle?.fontItalic;
  const catAxisTitleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.catAxisTitleFontColor, chart.catAxisTitleFontPaintAuthored, axisTitleStyle,
  );
  const catAxisTitleFontColor = catAxisTitleTextPaint.color;
  const catAxisTitleFontFace = chart.catAxisTitleFontFace ?? chart.chartTextStyle?.fontFace
    ?? axisTitleStyle?.fontFace;
  const valAxisTitleFontSizeHpt = chart.valAxisTitleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? axisTitleStyle?.fontSizeHpt;
  const valAxisTitleFontBold = chart.valAxisTitleFontBold ?? chart.chartTextStyle?.fontBold
    ?? axisTitleStyle?.fontBold;
  const valAxisTitleFontItalic = chart.valAxisTitleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? axisTitleStyle?.fontItalic;
  const valAxisTitleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.valAxisTitleFontColor, chart.valAxisTitleFontPaintAuthored, axisTitleStyle,
  );
  const valAxisTitleFontColor = valAxisTitleTextPaint.color;
  const valAxisTitleFontFace = chart.valAxisTitleFontFace ?? chart.chartTextStyle?.fontFace
    ?? axisTitleStyle?.fontFace;
  const dataLabelFontSizeHpt = chart.dataLabelFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? dataLabelStyle?.fontSizeHpt ?? null;
  const dataLabelFontBold = chart.dataLabelFontBold ?? chart.chartTextStyle?.fontBold
    ?? dataLabelStyle?.fontBold;
  const dataLabelFontItalic = chart.dataLabelFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? dataLabelStyle?.fontItalic;
  const dataLabelFontLanguage = chart.dataLabelFontLanguage
    ?? chart.chartTextStyle?.fontLanguage ?? dataLabelStyle?.fontLanguage;
  const dataLabelFontBaseline = chart.dataLabelFontBaseline
    ?? chart.chartTextStyle?.fontBaseline ?? dataLabelStyle?.fontBaseline;
  const dataLabelTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.dataLabelFontColor, chart.dataLabelFontPaintAuthored, dataLabelStyle,
  );
  const dataLabelFontColor = dataLabelTextPaint.color;
  const dataLabelFontFace = chart.dataLabelFontFace ?? chart.chartTextStyle?.fontFace
    ?? dataLabelStyle?.fontFace;
  changed ||= valMajor.visible !== chart.valAxisMajorGridlines
    || valMajor.color !== chart.valAxisGridlineColor
    || valMajor.widthEmu !== chart.valAxisGridlineWidthEmu
    || valMajor.dash !== chart.valAxisGridlineDash
    || valMajor.paintAuthored !== chart.valAxisGridlinePaintAuthored
    || catMajor.visible !== chart.catAxisMajorGridlines
    || catMajor.color !== chart.catAxisGridlineColor
    || catMajor.widthEmu !== chart.catAxisGridlineWidthEmu
    || catMajor.dash !== chart.catAxisGridlineDash
    || catMajor.paintAuthored !== chart.catAxisGridlinePaintAuthored
    || valMinor.visible !== chart.valAxisMinorGridlines
    || valMinor.color !== chart.valAxisMinorGridlineColor
    || valMinor.widthEmu !== chart.valAxisMinorGridlineWidthEmu
    || valMinor.dash !== chart.valAxisMinorGridlineDash
    || valMinor.paintAuthored !== chart.valAxisMinorGridlinePaintAuthored
    || catMinor.visible !== chart.catAxisMinorGridlines
    || catMinor.color !== chart.catAxisMinorGridlineColor
    || catMinor.widthEmu !== chart.catAxisMinorGridlineWidthEmu
    || catMinor.dash !== chart.catAxisMinorGridlineDash
    || catMinor.paintAuthored !== chart.catAxisMinorGridlinePaintAuthored
    || secondaryValAxis !== chart.secondaryValAxis
    || secondaryCatAxis !== chart.secondaryCatAxis
    || catAxisLine.color !== chart.catAxisLineColor
    || catAxisLine.widthEmu !== chart.catAxisLineWidthEmu
    || catAxisLine.dash !== chart.catAxisLineDash
    || catAxisLineHidden !== chart.catAxisLineHidden
    || catAxisLine.paintAuthored !== chart.catAxisLinePaintAuthored
    || valAxisLine.color !== chart.valAxisLineColor
    || valAxisLine.widthEmu !== chart.valAxisLineWidthEmu
    || valAxisLine.dash !== chart.valAxisLineDash
    || valAxisLineHidden !== chart.valAxisLineHidden
    || valAxisLine.paintAuthored !== chart.valAxisLinePaintAuthored
    || catAxisFontSizeHpt !== chart.catAxisFontSizeHpt
    || catAxisFontBold !== chart.catAxisFontBold
    || catAxisFontItalic !== chart.catAxisFontItalic
    || catAxisFontColor !== chart.catAxisFontColor
    || catAxisTextPaint.authored !== chart.catAxisFontPaintAuthored
    || catAxisFontFace !== chart.catAxisFontFace
    || valAxisFontSizeHpt !== chart.valAxisFontSizeHpt
    || valAxisFontBold !== chart.valAxisFontBold
    || valAxisFontItalic !== chart.valAxisFontItalic
    || valAxisFontColor !== chart.valAxisFontColor
    || valAxisTextPaint.authored !== chart.valAxisFontPaintAuthored
    || valAxisFontFace !== chart.valAxisFontFace
    || titleFontSizeHpt !== chart.titleFontSizeHpt
    || titleFontBold !== chart.titleFontBold
    || titleFontItalic !== chart.titleFontItalic
    || titleFontLanguage !== chart.titleFontLanguage
    || titleFontBaseline !== chart.titleFontBaseline
    || titleFontColor !== chart.titleFontColor
    || titleTextPaint.authored !== chart.titleFontPaintAuthored
    || titleFontFace !== chart.titleFontFace
    || catAxisTitleFontSizeHpt !== chart.catAxisTitleFontSizeHpt
    || catAxisTitleFontBold !== chart.catAxisTitleFontBold
    || catAxisTitleFontItalic !== chart.catAxisTitleFontItalic
    || catAxisTitleFontColor !== chart.catAxisTitleFontColor
    || catAxisTitleTextPaint.authored !== chart.catAxisTitleFontPaintAuthored
    || catAxisTitleFontFace !== chart.catAxisTitleFontFace
    || valAxisTitleFontSizeHpt !== chart.valAxisTitleFontSizeHpt
    || valAxisTitleFontBold !== chart.valAxisTitleFontBold
    || valAxisTitleFontItalic !== chart.valAxisTitleFontItalic
    || valAxisTitleFontColor !== chart.valAxisTitleFontColor
    || valAxisTitleTextPaint.authored !== chart.valAxisTitleFontPaintAuthored
    || valAxisTitleFontFace !== chart.valAxisTitleFontFace
    || dataLabelFontSizeHpt !== chart.dataLabelFontSizeHpt
    || dataLabelFontBold !== chart.dataLabelFontBold
    || dataLabelFontItalic !== chart.dataLabelFontItalic
    || dataLabelFontLanguage !== chart.dataLabelFontLanguage
    || dataLabelFontBaseline !== chart.dataLabelFontBaseline
    || dataLabelFontColor !== chart.dataLabelFontColor
    || dataLabelTextPaint.authored !== chart.dataLabelFontPaintAuthored
    || dataLabelFontFace !== chart.dataLabelFontFace;
  const effective = changed ? {
    ...chart,
    series,
    dataTable,
    valAxisMajorGridlines: valMajor.visible,
    valAxisGridlineColor: valMajor.color,
    valAxisGridlineWidthEmu: valMajor.widthEmu,
    valAxisGridlineDash: valMajor.dash,
    valAxisGridlinePaintAuthored: valMajor.paintAuthored,
    catAxisMajorGridlines: catMajor.visible,
    catAxisGridlineColor: catMajor.color,
    catAxisGridlineWidthEmu: catMajor.widthEmu,
    catAxisGridlineDash: catMajor.dash,
    catAxisGridlinePaintAuthored: catMajor.paintAuthored,
    valAxisMinorGridlines: valMinor.visible,
    valAxisMinorGridlineColor: valMinor.color,
    valAxisMinorGridlineWidthEmu: valMinor.widthEmu,
    valAxisMinorGridlineDash: valMinor.dash,
    valAxisMinorGridlinePaintAuthored: valMinor.paintAuthored,
    catAxisMinorGridlines: catMinor.visible,
    catAxisMinorGridlineColor: catMinor.color,
    catAxisMinorGridlineWidthEmu: catMinor.widthEmu,
    catAxisMinorGridlineDash: catMinor.dash,
    catAxisMinorGridlinePaintAuthored: catMinor.paintAuthored,
    secondaryValAxis,
    secondaryCatAxis,
    catAxisLineColor: catAxisLine.color,
    catAxisLineWidthEmu: catAxisLine.widthEmu,
    catAxisLineDash: catAxisLine.dash,
    catAxisLineHidden,
    catAxisLinePaintAuthored: catAxisLine.paintAuthored,
    valAxisLineColor: valAxisLine.color,
    valAxisLineWidthEmu: valAxisLine.widthEmu,
    valAxisLineDash: valAxisLine.dash,
    valAxisLineHidden,
    valAxisLinePaintAuthored: valAxisLine.paintAuthored,
    catAxisFontSizeHpt,
    catAxisFontBold,
    catAxisFontItalic,
    catAxisFontColor,
    catAxisFontPaintAuthored: catAxisTextPaint.authored,
    catAxisFontFace,
    valAxisFontSizeHpt,
    valAxisFontBold,
    valAxisFontItalic,
    valAxisFontColor,
    valAxisFontPaintAuthored: valAxisTextPaint.authored,
    valAxisFontFace,
    titleFontSizeHpt,
    titleFontBold,
    titleFontItalic,
    titleFontLanguage,
    titleFontBaseline,
    titleFontColor,
    titleFontPaintAuthored: titleTextPaint.authored,
    titleFontFace,
    catAxisTitleFontSizeHpt,
    catAxisTitleFontBold,
    catAxisTitleFontItalic,
    catAxisTitleFontColor,
    catAxisTitleFontPaintAuthored: catAxisTitleTextPaint.authored,
    catAxisTitleFontFace,
    valAxisTitleFontSizeHpt,
    valAxisTitleFontBold,
    valAxisTitleFontItalic,
    valAxisTitleFontColor,
    valAxisTitleFontPaintAuthored: valAxisTitleTextPaint.authored,
    valAxisTitleFontFace,
    dataLabelFontSizeHpt,
    dataLabelFontBold,
    dataLabelFontItalic,
    dataLabelFontLanguage,
    dataLabelFontBaseline,
    dataLabelFontColor,
    dataLabelFontPaintAuthored: dataLabelTextPaint.authored,
    dataLabelFontFace,
  } : chart;
  return chartStyleRoleLegend(chartStyleRolePlotArea(chartStyleRoleChartArea(
    chartStyleRoleSeriesAxis(effective),
  )));
}

export function drawUpDownBars(
  ctx: CanvasRenderingContext2D,
  startValueAt: (index: number) => number | null,
  endValueAt: (index: number) => number | null,
  pointCount: number,
  toX: (index: number) => number,
  toYStart: (value: number) => number,
  toYEnd: (value: number) => number,
  slotWidth: number,
  style: ChartStockUpDownBarStyle,
  ptToPx: number,
  automaticPaint?: {
    lineColor: string;
    lineWidthEmu: number;
    upFillColor: string;
    downFillColor: string;
  },
  shapeRotationDeg = 0,
): void {
  const gapPercent = Number.isFinite(style.gapWidthPercent) && style.gapWidthPercent >= 0
    ? style.gapWidthPercent
    : 150;
  const barWidth = Math.max(0, slotWidth / (1 + gapPercent / 100));
  for (let index = 0; index < pointCount; index++) {
    const start = startValueAt(index);
    const end = endValueAt(index);
    if (start == null || end == null || !Number.isFinite(start) || !Number.isFinite(end)) continue;
    const startY = toYStart(start);
    const endY = toYEnd(end);
    const barHeight = Math.abs(endY - startY);
    if (!(barWidth > 0) || !(barHeight > 0) || !Number.isFinite(barHeight)) continue;
    const paint = end >= start ? style.up : style.down;
    const fillOwned = paint.fillPaintAuthored === true
      || paint.fill != null || paint.fillColor != null || paint.fillHidden === true;
    const automaticFill = fillOwned
      ? undefined
      : end >= start ? automaticPaint?.upFillColor : automaticPaint?.downFillColor;
    const fillColor = paint.fillColor ?? automaticFill;
    const barX = toX(index) - barWidth / 2;
    const barY = Math.min(startY, endY);
    const lineOwned = paint.linePaintAuthored === true
      || paint.lineColor != null || paint.lineHidden === true;
    const lineColor = paint.lineColor ?? (lineOwned ? undefined : automaticPaint?.lineColor);
    const lineWidthEmu = paint.lineWidthEmu
      ?? (lineOwned ? undefined : automaticPaint?.lineWidthEmu);
    const paintBar = (target: CanvasRenderingContext2D): void => {
      if (!paint.fillHidden && (paint.fill != null || fillColor != null)) {
        paintClassicDataPointRect(
          target,
          paint.fill ?? (fillColor ? { fillType: 'solid', color: fillColor } : null),
          { x: barX, y: barY, w: barWidth, h: barHeight },
          fillColor ? `#${fillColor}` : 'rgba(0,0,0,0)',
          ptToPx,
          shapeRotationDeg,
        );
      }
      if (!paint.lineHidden
        && (paint.linePaintAuthored !== true || lineColor != null) && (
        lineColor != null || lineWidthEmu != null
      )) {
        const previousDash = target.getLineDash();
        const previousCap = target.lineCap;
        const previousJoin = target.lineJoin;
        target.strokeStyle = `#${lineColor ?? '000000'}`;
        target.lineWidth = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx)
          : Math.max(1, 0.75 * ptToPx);
        target.setLineDash(dashPatternForPreset(paint.lineDash ?? undefined, target.lineWidth));
        target.lineCap = paint.lineCap === 'rnd'
          ? 'round' : paint.lineCap === 'sq' ? 'square' : 'butt';
        target.lineJoin = paint.lineJoin === 'round' || paint.lineJoin === 'bevel'
          ? paint.lineJoin : 'miter';
        target.strokeRect(barX, barY, barWidth, barHeight);
        target.setLineDash(previousDash);
        target.lineCap = previousCap;
        target.lineJoin = previousJoin;
      }
    };
    paintChartStyleEffects(
      ctx,
      paint.style,
      undefined,
      index,
      { x: barX, y: barY, w: barWidth, h: barHeight },
      ptToPx,
      paintBar,
    );
  }
}

export function drawLineGroupDecorations(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  pointCount: number,
  toX: (index: number) => number,
  yMapFor: (series: ChartSeries) => (value: number) => number,
  categoryAxisYFor: (series: ChartSeries) => number,
  valueFor: (series: ChartSeries, index: number) => number | null,
  slotWidth: number,
  ptToPx: number,
  shapeRotationDeg: number,
  phase: 'background' | 'foreground',
): void {
  for (const decoration of chart.lineGroupDecorations ?? []) {
    let members = chart.series.filter(series => series.lineGroupIndex === decoration.groupIndex);
    // Hand-authored callers predating line-group provenance still represent a
    // single ordinary line group as the whole series list.
    if (members.length === 0 && decoration.groupIndex === 0
      && ['line', 'stackedLine', 'stackedLinePct'].includes(chart.chartType)) {
      members = chart.series.filter(series => series.seriesType == null || series.seriesType === 'line');
    }
    if (members.length === 0) continue;

    if (phase === 'foreground' && decoration.upDownBars && members.length >= 2) {
      const first = members[0];
      const last = members[members.length - 1];
      // Empty upBars/downBars paint is application-defined. The retained
      // Office observation is limited to classic Style 2; it is a family
      // default above the numeric role, while direct and linked paint remain
      // authoritative.
      const automaticPaint = chart.legacyChartStyle === 2 ? {
        lineColor: '000000', lineWidthEmu: 9525,
        upFillColor: 'FFFFFF', downFillColor: '000000',
      } : undefined;
      const upDownBars = {
        ...decoration.upDownBars,
        up: chartStyleRoleBarPaint(chart, decoration.upDownBars.up, 'upBar', automaticPaint),
        down: chartStyleRoleBarPaint(chart, decoration.upDownBars.down, 'downBar', automaticPaint),
      };
      drawUpDownBars(
        ctx, index => valueFor(first, index), index => valueFor(last, index), pointCount, toX,
        yMapFor(first), yMapFor(last), slotWidth, upDownBars, ptToPx,
        undefined, shapeRotationDeg,
      );
    }

    if (phase === 'foreground') continue;

    const dropLineStyle = decoration.dropLines
      ? chartStyleRoleLine(chart, decoration.dropLines, 'dropLine')
      : null;
    if (dropLineStyle && applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      // Office paints one envelope per category, not one line per series. The
      // envelope includes the effective category-axis crossing and every
      // plotted point in the owning line group. This is observable in vector
      // output for both ordinary and interior crossings; painting per-series
      // segments produces coincident seams and the wrong visible endpoints.
      drawDropLineEnvelopes(
        ctx, members, pointCount, toX, yMapFor, categoryAxisYFor, valueFor,
      );
    }

    const hiLowLineStyle = decoration.hiLowLines
      ? chartStyleRoleLine(chart, decoration.hiLowLines, 'hiLoLine')
      : null;
    if (hiLowLineStyle && members.length >= 2
      && applyDecorationLineStyle(ctx, hiLowLineStyle, ptToPx)) {
      const toY = yMapFor(members[0]);
      for (let index = 0; index < pointCount; index++) {
        let low = Infinity;
        let high = -Infinity;
        for (const series of members) {
          const value = valueFor(series, index);
          if (value == null || !Number.isFinite(value)) continue;
          low = Math.min(low, value);
          high = Math.max(high, value);
        }
        if (!Number.isFinite(low) || !Number.isFinite(high)) continue;
        ctx.beginPath();
        ctx.moveTo(toX(index), toY(low));
        ctx.lineTo(toX(index), toY(high));
        ctx.stroke();
      }
    }
  }
}

export function axisCrossingValue(
  crossesAt: number | null | undefined,
  crosses: string | null | undefined,
  min: number,
  max: number,
): number {
  if (crossesAt != null && Number.isFinite(crossesAt)) {
    return clamp(crossesAt, min, max);
  }
  if (crosses === 'max') return max;
  if (crosses === 'min') return min;
  return clamp(0, min, max);
}

export function categoryAxisCrossingValue(chart: ChartModel, min: number, max: number): number {
  return axisCrossingValue(chart.catAxisCrossesAt, chart.catAxisCrosses, min, max);
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

// NB: scatter deliberately has NO secondary value axis. Unlike bar/line/area,
// an XY scatter's X axis is already a numeric VALUE axis (not a category axis),
// and Excel/PowerPoint do not define a second Y value axis for a scatter combo
// (`useSecondaryAxis` / a right-hand `<c:valAx>` pairs with a category-based
// family). So `computeSecondaryAxis` is never called here — the CH7 helper is
// wired only into the category-axis families (bar already; line + area now).
export function scatterXValue(cats: string[], index: number, useIndexX: boolean): number | null {
  // A string-backed `<c:xVal>` is plotted by Office as the one-based ordinal
  // sequence 1..N. Zero-based array indices remain an implementation detail.
  if (useIndexX) return index + 1;
  const raw = cats[index];
  if (raw == null) return null;
  const value = parseFloat(raw);
  return Number.isNaN(value) ? null : value;
}

/** Return the linear bubble magnitude prescribed by ST_SizeRepresents.
 * `area` is the schema default, hence sqrt(value); `w` makes radius linear. */
export type BubbleGroupSettings = Pick<
  ChartModel, 'bubbleScale' | 'bubbleSizeRepresents' | 'showNegativeBubbles'
>;

export function bubbleSizeMagnitude(chart: BubbleGroupSettings, value: number): number {
  return chart.bubbleSizeRepresents === 'w' ? value : Math.sqrt(value);
}

export type ScatterSeriesLayer = {
  series: ChartSeries;
  seriesIndex: number;
  fallbackColor: string;
  cats: string[];
  pointOverrides: Map<number, NonNullable<ChartSeries['dataPointOverrides']>[number]>;
};

export function scatterPointFill(
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  index: number,
  fallbackColor: string,
): string {
  return markerFillColorFor(series, point, index, fallbackColor);
}

/** Resolve the classic bubble shape fill without collapsing DrawingML
 * provenance into the marker fallback. CT_DPt shape paint wins over CT_Ser
 * shape paint, which wins over the linked dataPoint role. */
export function bubblePointFill(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  seriesStyleIndex: number,
  fallbackColor: string,
  bubble3D = bubblePointIsThreeD(series, point),
): { color: string; paint: Fill | null | undefined } {
  const bubbleSize = series.bubbleSizes?.[pointIndex];
  if (bubbleSize != null && Number.isFinite(bubbleSize) && bubbleSize < 0) {
    // MS-OE376 §2.1.1504(b): Office always inverts a negative bubble,
    // regardless of `<c:invertIfNegative>`. The application-generated default
    // is outline-only for a flat bubble and white material for a 3-D bubble;
    // an authored c14 alternate fill remains authoritative.
    if (series.invertedFillHidden === true) return { color: '00000000', paint: null };
    if (series.invertedFill) {
      return {
        color: series.invertedFill.fillType === 'solid'
          ? series.invertedFill.color : fallbackColor,
        paint: series.invertedFill,
      };
    }
    return bubble3D
      ? { color: 'FFFFFF', paint: undefined }
      : { color: '00000000', paint: null };
  }
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint');
  const directPoint = chartStyleDirectFillDecision(
    point?.chartexStyle, rawLinked, pointIndex,
  );
  if (directPoint !== undefined) {
    return {
      color: directPoint?.fillType === 'solid' ? directPoint.color : fallbackColor,
      paint: directPoint,
    };
  }
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const linkedRole = chartDataPointStyleRole(chart, 'dataPoint', seriesIndex);
  const linkedIndex = chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex : seriesStyleIndex;
  if (point?.fillHidden === true) {
    const noFill = chartStyleDirectNoFillDecision(rawLinked);
    if (noFill !== undefined) return { color: '00000000', paint: noFill };
  }
  if (point?.color != null) return { color: point.color, paint: undefined };
  const pointColor = series.dataPointColors?.[pointIndex];
  if (pointColor != null) return { color: pointColor, paint: undefined };

  const directSeries = chartStyleDirectFillDecision(
    series.chartexStyle, rawLinked, seriesStyleIndex,
  );
  if (directSeries !== undefined) {
    return {
      color: directSeries?.fillType === 'solid' ? directSeries.color : fallbackColor,
      paint: directSeries,
    };
  }
  if (series.color != null) return { color: series.color, paint: undefined };
  const linkedPoint = chartExStylePaintDecision(
    chart,
    linkedRole,
    linkedIndex,
    series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : fallbackColor,
      paint: linkedPoint,
    };
  }
  return {
    color: scatterPointFill(series, point, pointIndex, fallbackColor),
    paint: markerFillPaintFor(series, point, pointIndex),
  };
}

export function bubblePointLine(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  seriesStyleIndex: number,
): {
  color: string | null;
  paint: ChartModel['plotAreaLineFill'] | null | undefined;
  widthEmu: number | null | undefined;
  dash: string | null | undefined;
  customDash: ChartModel['plotAreaLineCustomDash'];
  cap: string | null | undefined;
  join: string | null | undefined;
} {
  const pointStyle = point?.chartexStyle;
  const seriesStyle = series.chartexStyle;
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const linkedStyle = chartDataPointStyleRole(chart, 'dataPoint', seriesIndex);
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint');
  const linkedIndex = chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex : seriesStyleIndex;
  const linkedGeometry = linkedStyle;
  const dashLayers = [pointStyle, seriesStyle, linkedGeometry];
  let dash: string | null | undefined = point?.lineDash;
  let customDash: ChartModel['plotAreaLineCustomDash'];
  if (dash == null) {
    for (const layer of dashLayers) {
      if (layer?.lineDash != null || layer?.lineCustomDash != null
        || layer?.lineDashAuthored === true) {
        dash = layer.lineDash;
        customDash = layer.lineCustomDash ?? undefined;
        break;
      }
    }
  }
  const geometry = {
    widthEmu: point?.lineWidthEmu
      ?? pointStyle?.lineWidthEmu
      ?? series.lineWidthEmu
      ?? seriesStyle?.lineWidthEmu
      ?? linkedGeometry?.lineWidthEmu
      ?? point?.markerLineWidthEmu
      ?? series.markerLineWidthEmu,
    dash,
    customDash,
    cap: pointStyle?.lineCap ?? seriesStyle?.lineCap ?? linkedGeometry?.lineCap,
    join: pointStyle?.lineJoin ?? seriesStyle?.lineJoin ?? linkedGeometry?.lineJoin,
  };
  const pointPaint = chartStyleDirectLineDecision(pointStyle, rawLinked, pointIndex);
  if (pointPaint !== undefined) {
    return {
      color: pointPaint?.fillType === 'solid' ? pointPaint.color : point?.lineColor ?? null,
      paint: pointPaint,
      ...geometry,
    };
  }
  if (point?.lineHidden === true) {
    const noLine = chartStyleDirectNoLineDecision(rawLinked);
    if (noLine !== undefined) return { color: null, paint: noLine, ...geometry };
  }
  if (point?.lineColor != null) {
    return { color: point.lineColor, paint: undefined, ...geometry };
  }

  const seriesPaint = chartStyleDirectLineDecision(
    seriesStyle, rawLinked, seriesStyleIndex,
  );
  if (seriesPaint !== undefined) {
    return {
      color: seriesPaint?.fillType === 'solid' ? seriesPaint.color : series.lineColor ?? null,
      paint: seriesPaint,
      ...geometry,
    };
  }
  if (series.lineHidden === true) {
    const noLine = chartStyleDirectNoLineDecision(rawLinked);
    if (noLine !== undefined) return { color: null, paint: noLine, ...geometry };
  }
  if (series.lineColor != null) {
    return { color: series.lineColor, paint: undefined, ...geometry };
  }
  const linkedPoint = chartExStyleLinePaintDecision(
    chart, linkedStyle, linkedIndex, series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : null,
      paint: linkedPoint,
      ...geometry,
    };
  }
  const bubbleSize = series.bubbleSizes?.[pointIndex];
  const automaticNegativeThreeDLine = bubbleSize != null
    && Number.isFinite(bubbleSize)
    && bubbleSize < 0
    && bubblePointIsThreeD(series, point)
    ? '000000'
    : null;
  return {
    // Current Excel gives its generated white negative 3-D material a black
    // outline. Direct or linked no-line returned above remains authoritative.
    color: point?.markerLine
      ?? series.markerLine
      ?? series.lineColor
      ?? automaticNegativeThreeDLine,
    paint: undefined,
    ...geometry,
  };
}

export function makeScatterSeriesLayer(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
): ScatterSeriesLayer {
  return {
    series,
    seriesIndex: index,
    fallbackColor: chartColor(index, series),
    cats: series.categories ?? chart.categories,
    pointOverrides: new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point])),
  };
}

/** One `<c:bubbleChart>` group has one size scale: every series must therefore
 * be normalized against the same maximum bubble magnitude. */
export function bubbleSizeToDiameterScale(
  chart: BubbleGroupSettings,
  layers: readonly ScatterSeriesLayer[],
  useIndexX: boolean,
  pw: number,
  ph: number,
): number {
  const bubbleScale = clamp(chart.bubbleScale ?? 100, 0, 300);
  if (bubbleScale <= 0) return 0;
  let maxMagnitude = 0;
  for (const { series, cats, pointOverrides } of layers) {
    if (series.showMarker === false || series.markerSymbol === 'none') continue;
    for (let index = 0; index < series.values.length; index++) {
      if (series.values[index] == null || scatterXValue(cats, index, useIndexX) == null) continue;
      if (pointOverrides.get(index)?.markerSymbol === 'none') continue;
      const value = visibleBubbleSize(chart, series.bubbleSizes?.[index]);
      if (value != null) {
        maxMagnitude = Math.max(maxMagnitude, bubbleSizeMagnitude(chart, value));
      }
    }
  }
  if (maxMagnitude <= 0) return 0;
  // ECMA-376 defines bubbleScale as 0..300% of an application-defined default,
  // but intentionally leaves that default to the consumer. Excel's vector
  // output across the complete 0/25/50/75/100/150/200/300 boundary set follows
  // a bounded scale curve: 0 hides bubbles, 100 uses one quarter of the shorter
  // plot dimension, and 300 approaches one half. The equivalent closed form is
  // `shortSide * scale / (300 + scale)`. Keeping it here (rather than a sample-
  // specific diameter constant) makes the Office compatibility rule depend
  // only on the authored scale and the resolved plot geometry.
  const maximumDiameterPx = Math.min(pw, ph) * bubbleScale / (300 + bubbleScale);
  return maximumDiameterPx / maxMagnitude;
}

/** Paint scatter series into an already-computed plot rectangle. Axis/gridline
 * layout stays with the owning chart renderer, which lets a scatter group be
 * overlaid on a bar chart without duplicating either chart's frame. */
export function drawScatterSeriesLayer(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  entries: Array<{ series: ChartSeries; index: number }>,
  useIndexX: boolean,
  toX: (value: number) => number,
  toY: (value: number) => number,
  chartRect: ChartRect,
  px0: number,
  py0: number,
  pw: number,
  ph: number,
  ptToPx: number,
  isBubble: boolean,
  style: string,
  layoutReferenceRect: DataLabelRect,
  valueAxisMaximum: number,
  valueDisplayUnits?: ChartDisplayUnits | null,
  shapeRotationDeg = 0,
  bubbleSettings?: BubbleGroupSettings,
): void {
  const groupDrawsLines = style === 'line' || style === 'lineMarker' || style === 'lineNoMarker';
  const groupDrawsSmooth = style === 'smooth'
    || style === 'smoothMarker'
    || style === 'smoothNoMarker';
  const hideMarkersByStyle = markersSuppressedByChartStyle(
    'scatter', chart.chartType, style, chart.radarStyle,
  );
  const layers = entries.map(({ series, index }) => makeScatterSeriesLayer(chart, series, index));
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  const effectiveBubbleSettings = bubbleSettings ?? chart;
  const bubbleScale = isBubble
    ? bubbleSizeToDiameterScale(effectiveBubbleSettings, layers, useIndexX, pw, ph)
    : 0;

  // Excel paints a scatter group by geometry phase, not one complete series at
  // a time: all series lines/error bars first, then all markers, then all data
  // labels. This is observable in dot/range plots where a final invisible
  // scatter series authors full-width horizontal guides. Painting per series
  // placed those guides on top of earlier series' dots and labels.
  for (const { series: s, fallbackColor, cats } of layers) {
    for (const eb of s.errBars ?? []) {
      drawSeriesErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), cats, useIndexX, toX, toY,
        fallbackColor,
      );
    }
  }

  for (const { series: s, seriesIndex, fallbackColor, cats } of layers) {
    const automaticPointStyle = style === 'marker'
      && hasFilteredScatterAutomaticPointStyle(s);
    const drawLines = automaticPointStyle || groupDrawsLines;
    const drawSmooth = (!automaticPointStyle && style === 'marker' && !isBubble)
      || groupDrawsSmooth;
    if (drawLines || drawSmooth) {
      const pts: IndexedLinePoint[] = [];
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci];
        if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        pts.push({ x: toX(xv), y: toY(yv), index: ci });
      }
      if (pts.length >= 2) {
        const styleIndex = chartExSeriesFormatIndex(s, seriesIndex);
        const scatterBounds = { x: px0, y: py0, w: pw, h: ph };
        if (chartSeriesVariesByPoint(chart, seriesIndex)) {
          paintClassicVaryingLineSegments(
            ctx, chart, s, [pts], drawSmooth, false, fallbackColor, 1.5,
            ptToPx, scatterBounds, shapeRotationDeg, true,
          );
          continue;
        }
        const paintScatterLine = (target: CanvasRenderingContext2D): void => {
          target.save();
          const paintLine = applyClassicStyleLine(
            target, chart, 'dataPointLine', s, undefined, styleIndex, fallbackColor,
            1.5, ptToPx, scatterBounds, shapeRotationDeg,
            true, false,
          );
          if (paintLine && automaticPointStyle && s.dataPointColors?.some(Boolean)) {
            target.lineWidth = 1.5;
            for (let i = 1; i < pts.length; i++) {
              target.strokeStyle = `#${s.dataPointColors[i] ?? s.color ?? fallbackColor.replace(/^#/, '')}`;
              target.beginPath();
              target.moveTo(pts[i - 1].x, pts[i - 1].y);
              target.lineTo(pts[i].x, pts[i].y);
              target.stroke();
            }
          } else if (paintLine) {
            target.beginPath();
            target.moveTo(pts[0].x, pts[0].y);
            if (drawSmooth && pts.length >= 3) {
              for (let i = 0; i < pts.length - 1; i++) {
                const p0 = pts[i - 1] ?? pts[i];
                const p1 = pts[i];
                const p2 = pts[i + 1];
                const p3 = pts[i + 2] ?? p2;
                target.bezierCurveTo(
                  p1.x + (p2.x - p0.x) / 6,
                  p1.y + (p2.y - p0.y) / 6,
                  p2.x - (p3.x - p1.x) / 6,
                  p2.y - (p3.y - p1.y) / 6,
                  p2.x,
                  p2.y,
                );
              }
            } else {
              for (let i = 1; i < pts.length; i++) target.lineTo(pts[i].x, pts[i].y);
            }
            target.stroke();
          }
          target.restore();
        };
        paintChartStyleEffects(
          ctx,
          chartStyleEffectOwner(s.chartexStyle),
          chart.chartStyleRoles?.dataPointLine,
          styleIndex,
          scatterBounds,
          ptToPx,
          paintScatterLine,
        );
      }
    }
  }

  for (const { series: s, seriesIndex, fallbackColor, cats, pointOverrides } of layers) {
    const seriesMarkersVisible = !hideMarkersByStyle
      && s.showMarker !== false
      && s.markerSymbol !== 'none';
    if (seriesMarkersVisible || (!hideMarkersByStyle && hasVisiblePointMarkerOverride(s))) {
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci];
        if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        const dpt = pointOverrides.get(ci);
        const defaultSymbol = isBubble ? 'circle' : (s.automaticMarkerSymbol ?? 'circle');
        const symbol = effectiveMarkerSymbol(s, dpt, defaultSymbol, seriesMarkersVisible);
        if (symbol === 'none') continue;
        let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
        if (isBubble) {
          if (bubbleScale <= 0) continue;
          const bubbleSize = visibleBubbleSize(effectiveBubbleSettings, s.bubbleSizes?.[ci]);
          if (bubbleSize == null) continue;
          sizePt = (bubbleSizeMagnitude(effectiveBubbleSettings, bubbleSize) * bubbleScale) / ptToPx;
        }
        const bubbleFill = isBubble
          ? bubblePointFill(
              chart, s, dpt, ci, chartExSeriesFormatIndex(s, seriesIndex), fallbackColor,
            )
          : null;
        const fill = bubbleFill?.color ?? scatterPointFill(s, dpt, ci, fallbackColor);
        // Bubble geometry is the series shape itself, so its outline comes from
        // `<c:ser><c:spPr><a:ln>` rather than a `<c:marker>` block. Ordinary
        // scatter markers continue to use markerLine only.
        const bubbleLine = isBubble
          ? bubblePointLine(chart, s, dpt, ci, chartExSeriesFormatIndex(s, seriesIndex))
          : null;
        const line = isBubble
          ? bubbleLine!.color
          : dpt?.markerLine ?? s.markerLine ?? null;
        const markerLineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
        const bubbleLineWidthEmu = bubbleLine?.widthEmu;
        const lineWidthEmu = isBubble ? bubbleLineWidthEmu : markerLineWidthEmu;
        const lineWidthPx = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx)
          : undefined;
        const isThreeD = isBubble ? bubblePointIsThreeD(s, dpt) : false;
        drawChartMarker(
          ctx, chart, s, dpt, ci, toX(xv), toY(yv), symbol, sizePt, fill, line, ptToPx, lineWidthPx,
          isBubble ? bubbleFill!.paint : markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
          isBubble ? bubbleLine!.paint : undefined,
          isBubble ? bubbleLine!.dash : undefined,
          isBubble ? bubbleLine!.customDash : undefined,
          isBubble ? bubbleLine!.cap : undefined,
          isBubble ? bubbleLine!.join : undefined,
          isThreeD,
          isBubble,
        );
      }
    }
  }

  for (const { series: s, seriesIndex, cats, pointOverrides } of layers) {
    const markerGapAt = (pointIndex: number): number => {
      if (hideMarkersByStyle) return 0;
      const seriesMarkerVisible = s.showMarker !== false && s.markerSymbol !== 'none';
      const dpt = pointOverrides.get(pointIndex);
      const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkerVisible);
      if (symbol === 'none') return 0;
      let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
      if (isBubble) {
        if (bubbleScale <= 0) return 0;
        const bubbleSize = visibleBubbleSize(
          effectiveBubbleSettings, s.bubbleSizes?.[pointIndex],
        );
        if (bubbleSize == null) return 0;
        sizePt = bubbleSizeMagnitude(effectiveBubbleSettings, bubbleSize) * bubbleScale / ptToPx;
      }
      return Math.max(0, sizePt * ptToPx / 2);
    };
    drawSeriesDataLabels(
      ctx,
      s,
      cats,
      useIndexX,
      toX,
      toY,
      ph,
      ptToPx,
      chart.date1904,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      chart.dataLabelPosition ?? 'r',
      // Office lets automatic left/right endpoint labels occupy the chart-area
      // gutter while keeping their vertical placement constrained to the plot.
      { x: chartRect.x, y: py0, w: chartRect.w, h: ph },
      layoutReferenceRect,
      face => chartFontFamily(chart, face, 'minor'),
      valueDisplayUnits,
      pointIndex => dataLabelLegendKey(seriesIndex, pointIndex),
      value => dataLabelWithinAxisMaximum(chart, value, valueAxisMaximum),
      shapeRotationDeg,
      markerGapAt,
    );
  }

  for (const { series: s, fallbackColor, cats } of layers) {
    const trendlineX = s.values.map((_, index) => scatterXValue(cats, index, useIndexX));
    drawSeriesTrendlines(
      ctx, s, fallbackColor, toX, toY, ptToPx, trendlineX,
      {
        chart,
        chartRect,
        plotRect: { x: px0, y: py0, w: pw, h: ph },
        clipLineToPlot: true,
        shapeRotationDeg,
      },
    );
  }
}

// Three fixed gradients with 4 + 5 + 6 stops. Keep the work count aligned
// with the complete material so a bubble is admitted or rejected atomically.
export const BUBBLE_3D_MATERIAL_COMPONENTS = 15;

/** ECMA-376 §21.2.2.21 only enables `bubble3D`; it does not define a lighting
 * material. Paint the bounded application-defined material observed in desktop
 * Excel vector output. A single radial envelope cannot independently
 * express the diffuse highlight, right/lower falloff, and narrow lower
 * reflected-light band, so those three components are composited in order.
 * The recipe is normalized to bubble-local coordinates and is therefore shared
 * by every colour, size, and host transform rather than fitted per sample.
 * `source-atop` preserves the authored fill alpha on every pass. */
export function paintBubble3DMaterial(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  sizePx: number,
): void {
  const previousComposite = ctx.globalCompositeOperation;
  const previousFill = ctx.fillStyle;
  ctx.save();
  ctx.clip();
  const paintLayer = (material: CanvasGradient) => {
    ctx.globalCompositeOperation = 'source-atop';
    ctx.fillStyle = material;
    ctx.fillRect(cx - sizePx / 2, cy - sizePx / 2, sizePx, sizePx);
  };

  const diffuseX = cx - sizePx * 0.08;
  const diffuseY = cy - sizePx * 0.17;
  const diffuse = ctx.createRadialGradient(
    diffuseX, diffuseY, 0,
    diffuseX, diffuseY, sizePx * 0.55,
  );
  diffuse.addColorStop(0, 'rgba(255,255,255,0.72)');
  diffuse.addColorStop(0.14, 'rgba(255,255,255,0.48)');
  diffuse.addColorStop(0.38, 'rgba(255,255,255,0.1)');
  diffuse.addColorStop(1, 'rgba(255,255,255,0)');
  paintLayer(diffuse);

  const shadeX = cx - sizePx * 0.08;
  const shadeY = cy - sizePx * 0.18;
  const shade = ctx.createRadialGradient(
    shadeX, shadeY, 0,
    shadeX, shadeY, sizePx * 0.78,
  );
  shade.addColorStop(0, 'rgba(0,0,0,0)');
  shade.addColorStop(0.3, 'rgba(0,0,0,0)');
  shade.addColorStop(0.46, 'rgba(0,0,0,0.22)');
  shade.addColorStop(0.66, 'rgba(0,0,0,0.48)');
  shade.addColorStop(1, 'rgba(0,0,0,0.62)');
  paintLayer(shade);

  // The annulus centre is above-left. Its narrow 0.8--0.95 radius band
  // crosses the lower-left/lower-centre rim while staying clear of the dark
  // lower-right shoulder, matching the material boundary observed in Excel.
  const rimX = cx - sizePx * 0.2;
  const rimY = cy - sizePx * 0.45;
  const lowerRim = ctx.createRadialGradient(
    rimX, rimY, 0,
    rimX, rimY, sizePx,
  );
  lowerRim.addColorStop(0, 'rgba(255,255,255,0)');
  lowerRim.addColorStop(0.76, 'rgba(255,255,255,0)');
  lowerRim.addColorStop(0.82, 'rgba(255,255,255,0.05)');
  lowerRim.addColorStop(0.87, 'rgba(255,255,255,0.12)');
  lowerRim.addColorStop(0.95, 'rgba(255,255,255,0.28)');
  lowerRim.addColorStop(1, 'rgba(255,255,255,0)');
  paintLayer(lowerRim);

  // Recording contexts used by hosts/tests do not necessarily model a full
  // Canvas state stack, so restore the property explicitly as well.
  ctx.globalCompositeOperation = previousComposite;
  ctx.fillStyle = previousFill;
  ctx.restore();
}

export function drawChartMarker(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  cx: number,
  cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
  lineWidthPx: number | undefined,
  fillPaint: Fill | null | undefined,
  shapeRotationDeg: number,
  linePaint: ChartModel['plotAreaLineFill'] | null | undefined = undefined,
  lineDash: string | null | undefined = undefined,
  lineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  lineCap: string | null | undefined = undefined,
  lineJoin: string | null | undefined = undefined,
  bubble3D = false,
  bubble = false,
): void {
  const pointEffect = bubble
    ? chartStyleEffectOwner(point?.chartexStyle)
    // A point marker is the painted CT_DPt shape: its nested marker/spPr is
    // most specific, then dPt/spPr. The series marker/spPr supplies the marker
    // default; series/spPr belongs to the line/area body and must not leak onto
    // every marker.
    : chartStyleEffectOwner(
        point?.markerStyle,
        point?.chartexStyle,
      );
  const seriesEffect = bubble
    ? chartStyleEffectOwner(series.chartexStyle)
    : chartStyleEffectOwner(series.markerStyle);
  const directEffect = pointEffect ?? seriesEffect;
  const sourceSeriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const seriesStyleIndex = series.chartexFormatIdx
    ?? sourceSeriesIndex;
  const variesByPoint = chartSeriesVariesByPoint(chart, sourceSeriesIndex);
  const directEffectIndex = pointEffect ? pointIndex : seriesStyleIndex;
  const linkedMarkerStyle = chartDataPointStyleRole(
    chart, 'dataPointMarker', sourceSeriesIndex,
  );
  const rawLinkedMarkerStyle = rawLinkedChartStyleRole(chart, 'dataPointMarker');
  const linkedStyleIndex = variesByPoint ? pointIndex : seriesStyleIndex;

  const pointFillDecision = chartStyleDirectFillDecision(
    point?.markerStyle, rawLinkedMarkerStyle, pointIndex,
  );
  const seriesFillDecision = chartStyleDirectFillDecision(
    series.markerStyle, rawLinkedMarkerStyle, seriesStyleIndex,
  );
  const pointOwnsFill = pointFillDecision !== undefined
    || point?.markerFill != null
    || point?.markerFillPaintAuthored === true && point.markerStyle?.fillHidden !== true;
  const seriesOwnsFill = seriesFillDecision !== undefined
    || series.markerFill != null
    || series.markerFillPaintAuthored === true && series.markerStyle?.fillHidden !== true;
  let effectiveFill = fill;
  let effectiveFillPaint = fillPaint;
  if (!bubble && !pointOwnsFill && !seriesOwnsFill) {
    const directFillOwner = point?.markerStyle?.shapePropertiesPresent === true
      ? point.markerStyle
      : series.markerStyle;
    const linkedFillDecision = chartStyleFillCascade(
      linkedMarkerStyle,
      rawLinkedMarkerStyle,
      linkedStyleIndex,
      directFillOwner,
    );
    if (linkedFillDecision === null) {
      effectiveFill = '00000000';
      effectiveFillPaint = null;
    } else if (linkedFillDecision?.fillType === 'solid') {
      effectiveFill = linkedFillDecision.color;
      effectiveFillPaint = undefined;
    } else if (linkedFillDecision !== undefined) {
      effectiveFillPaint = linkedFillDecision;
    }
  }
  const pointLineDecision = chartStyleDirectLineDecision(
    point?.markerStyle, rawLinkedMarkerStyle, pointIndex,
  );
  const seriesLineDecision = chartStyleDirectLineDecision(
    series.markerStyle, rawLinkedMarkerStyle, seriesStyleIndex,
  );
  let effectiveLine = line;
  let effectiveLinePaint = linePaint;
  if (!bubble && effectiveLinePaint === undefined) {
    if (pointLineDecision !== undefined) {
      effectiveLinePaint = pointLineDecision?.fillType === 'solid' ? undefined : pointLineDecision;
    }
    else if (point?.markerLine != null) effectiveLinePaint = undefined;
    else if (point?.markerLinePaintAuthored === true
      && point.markerStyle?.lineHidden !== true
      && (point.markerLine == null || point.markerLine === '00000000')) effectiveLinePaint = null;
    else if (seriesLineDecision !== undefined) {
      effectiveLinePaint = seriesLineDecision?.fillType === 'solid' ? undefined : seriesLineDecision;
    }
    else if (series.markerLine != null) effectiveLinePaint = undefined;
    else if (series.markerLinePaintAuthored === true
      && series.markerStyle?.lineHidden !== true
      && (series.markerLine == null || series.markerLine === '00000000')) effectiveLinePaint = null;
    else {
      const directLineOwner = point?.markerStyle?.shapePropertiesPresent === true
        ? point.markerStyle
        : series.markerStyle;
      const linkedLineDecision = chartStyleLineCascade(
        linkedMarkerStyle,
        rawLinkedMarkerStyle,
        linkedStyleIndex,
        directLineOwner,
      );
      if (linkedLineDecision?.fillType === 'solid') {
        effectiveLine = linkedLineDecision.color;
        effectiveLinePaint = undefined;
      } else {
        effectiveLinePaint = linkedLineDecision;
        if (linkedLineDecision === null) effectiveLine = null;
      }
    }
  }
  const linkedLineGeometry = bubble ? undefined : linkedMarkerStyle;
  const effectiveLineWidthPx = lineWidthPx ?? (() => {
    const widthEmu = point?.markerStyle?.lineWidthEmu
      ?? series.markerStyle?.lineWidthEmu ?? linkedLineGeometry?.lineWidthEmu;
    return widthEmu != null ? axisLineWidthPx(widthEmu, ptToPx) : undefined;
  })();
  const markerDashChoice = chartStyleDashChoice(
    lineDash != null || lineCustomDash != null
      ? { lineDash, lineCustomDash, lineDashAuthored: true }
      : undefined,
    point?.markerStyle,
    series.markerStyle,
    linkedLineGeometry,
  );
  const effectiveLineDash = markerDashChoice?.lineDash;
  const effectiveLineCustomDash = markerDashChoice?.lineCustomDash;
  const effectiveLineCap = lineCap ?? point?.markerStyle?.lineCap
    ?? series.markerStyle?.lineCap ?? linkedLineGeometry?.lineCap;
  const effectiveLineJoin = lineJoin ?? point?.markerStyle?.lineJoin
    ?? series.markerStyle?.lineJoin ?? linkedLineGeometry?.lineJoin;
  const fallbackEffect = bubble
    ? chartDataPointStyleRole(
        chart,
        bubble3D ? 'dataPoint3D' : 'dataPoint',
        sourceSeriesIndex,
      )
    : linkedMarkerStyle;
  const fallbackEffectIndex = bubble && chartSeriesVariesByPoint(
    chart, sourceSeriesIndex,
  ) ? pointIndex : (variesByPoint ? pointIndex : seriesStyleIndex);
  drawMarker(
    ctx,
    cx, cy,
    symbol,
    sizePt,
    effectiveFill,
    effectiveLine,
    ptToPx,
    effectiveLineWidthPx,
    effectiveFillPaint,
    shapeRotationDeg,
    effectiveLinePaint,
    effectiveLineDash,
    effectiveLineCustomDash,
    effectiveLineCap,
    effectiveLineJoin,
    bubble3D,
    directEffect,
    fallbackEffect,
    directEffectIndex,
    fallbackEffectIndex,
  );
}

/** Linked/numeric marker styles are part of the effective marker even when the
 * series itself has no `<c:marker><c:spPr>`. Route those automatic glyphs
 * through the shared resolver so fill, line, and atomic effect precedence are
 * identical to explicitly formatted markers. This does not make a marker
 * visible for families (notably area) whose own visibility rule disables it. */
export function seriesHasResolvedMarkerDetail(
  chart: ChartModel,
  series: ChartSeries,
  sourceSeriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series)),
): boolean {
  return seriesHasMarkerDetail(series)
    || chartDataPointStyleRole(chart, 'dataPointMarker', sourceSeriesIndex) != null;
}

/** Draw a single ECMA-376 §21.2.2.32 marker shape centered at `(cx, cy)`.
 *  `sizePt` is the spec's marker side length in points (Excel's default
 *  is 5). `fill` and `line` are hex strings; a leading `#` is tolerated so
 *  callers that route through `chartColor` (which returns `#RRGGBB`)
 *  don't end up double-prefixing into an invalid `##RRGGBB`. `line` may
 *  be null in which case no outline is drawn. `picture` uses the host-warmed
 *  image lookup and fails closed when its authored relationship is unresolved. */
export function drawMarker(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
  lineWidthPx: number = 1,
  /** undefined uses `fill`; null is authored noFill. */
  fillPaint: Fill | null | undefined = undefined,
  shapeRotationDeg = 0,
  /** undefined uses `line`; null is authored line noFill. */
  linePaint: ChartModel['plotAreaLineFill'] | null | undefined = undefined,
  lineDash: string | null | undefined = undefined,
  lineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  lineCap: string | null | undefined = undefined,
  lineJoin: string | null | undefined = undefined,
  bubble3D = false,
  /** Direct marker/point effect component. */
  effectDirect: import('../../types/chart.js').ChartExElementStyle | null | undefined = undefined,
  /** Resolved linked/numeric `dataPointMarker` or `dataPoint3D` role. */
  effectFallback: import('../../types/chart.js').ChartExElementStyle | null | undefined = undefined,
  effectIndex = 0,
  effectFallbackIndex = effectIndex,
): void {
  const sizePx = Math.max(2, sizePt * ptToPx);
  const half = sizePx / 2;
  if (effectDirect !== undefined || effectFallback !== undefined) {
    paintChartStyleEffects(
      ctx,
      effectDirect,
      effectFallback,
      effectIndex,
      { x: cx - half, y: cy - half, w: sizePx, h: sizePx },
      ptToPx,
      target => drawMarker(
        target,
        cx, cy,
        symbol,
        sizePt,
        fill,
        line,
        ptToPx,
        lineWidthPx,
        fillPaint,
        shapeRotationDeg,
        linePaint,
        lineDash,
        lineCustomDash,
        lineCap,
        lineJoin,
        bubble3D,
      ),
      effectFallbackIndex,
    );
    return;
  }
  const fillCss = fill.startsWith('#') ? fill : `#${fill}`;
  const lineCss = line ? (line.startsWith('#') ? line : `#${line}`) : null;
  ctx.save();
  ctx.fillStyle = fillPaint === undefined
    ? fillCss
    : (fillPaint == null
        ? 'rgba(0,0,0,0)'
        : resolveFill(
            fillPaint, ctx, cx - half, cy - half, sizePx, sizePx, shapeRotationDeg,
          ) ?? 'rgba(0,0,0,0)');
  const resolvedLineStyle = linePaint === undefined
    ? lineCss
    : linePaint == null
      ? null
      : resolveFill(
          linePaint, ctx, cx - half, cy - half, sizePx, sizePx, shapeRotationDeg,
        );
  const hasLine = resolvedLineStyle != null;
  if (resolvedLineStyle) {
    ctx.strokeStyle = resolvedLineStyle;
    ctx.lineWidth = lineWidthPx;
    ctx.setLineDash(dashPatternForLine(lineCustomDash, lineDash, lineWidthPx));
    ctx.lineCap = lineCap === 'rnd' ? 'round' : lineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = lineJoin === 'round' || lineJoin === 'bevel' ? lineJoin : 'miter';
  }
  const imageFill = fillPaint?.fillType === 'image' ? fillPaint : undefined;
  const fillCurrentPath = () => {
    if (!imageFill) {
      if (fillPaint !== null) ctx.fill();
      return;
    }
    ctx.save();
    ctx.clip();
    paintChartImageFill(
      ctx, imageFill, cx - half, cy - half, sizePx, sizePx, ptToPx, shapeRotationDeg,
    );
    ctx.restore();
  };
  const paintMaterial = () => {
    if (bubble3D && fillPaint !== null) paintBubble3DMaterial(ctx, cx, cy, sizePx);
  };
  switch (symbol) {
    case 'square': {
      if (imageFill || bubble3D) {
        ctx.beginPath();
        ctx.rect(cx - half, cy - half, sizePx, sizePx);
        fillCurrentPath();
        paintMaterial();
      } else if (fillPaint !== null) {
        ctx.fillRect(cx - half, cy - half, sizePx, sizePx);
      }
      if (hasLine) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      break;
    }
    case 'diamond': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy);
      ctx.lineTo(cx, cy + half);
      ctx.lineTo(cx - half, cy);
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'triangle': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy + half);
      ctx.lineTo(cx - half, cy + half);
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'x': {
      ctx.strokeStyle = resolvedLineStyle ?? ctx.fillStyle;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy - half); ctx.lineTo(cx + half, cy + half);
      ctx.moveTo(cx - half, cy + half); ctx.lineTo(cx + half, cy - half);
      ctx.stroke();
      break;
    }
    case 'plus': {
      ctx.strokeStyle = resolvedLineStyle ?? ctx.fillStyle;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy); ctx.lineTo(cx + half, cy);
      ctx.moveTo(cx, cy - half); ctx.lineTo(cx, cy + half);
      ctx.stroke();
      break;
    }
    case 'star': {
      // 5-point star inscribed in a circle of radius `half`.
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? half : half * 0.45;
        const a = -Math.PI / 2 + i * Math.PI / 5;
        const px = cx + Math.cos(a) * r;
        const py = cy + Math.sin(a) * r;
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'dot': {
      // ECMA-376 §21.2.3.27: width=1/2 and height=1/5 of marker size.
      ctx.beginPath();
      ctx.ellipse(cx, cy, sizePx * 0.25, sizePx * 0.1, 0, 0, Math.PI * 2);
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'dash': {
      // ECMA-376 §21.2.3.27: height=1/5 of marker size.
      const dh = sizePx * 0.2;
      if (imageFill || bubble3D) {
        ctx.beginPath(); ctx.rect(cx - half, cy - dh / 2, sizePx, dh); fillCurrentPath();
        paintMaterial();
      } else if (fillPaint !== null) {
        ctx.fillRect(cx - half, cy - dh / 2, sizePx, dh);
      }
      if (hasLine) ctx.strokeRect(cx - half, cy - dh / 2, sizePx, dh);
      break;
    }
    case 'picture': {
      ctx.beginPath();
      ctx.rect(cx - half, cy - half, sizePx, sizePx);
      if (imageFill) {
        paintChartImageFill(
          ctx, imageFill, cx - half, cy - half, sizePx, sizePx, ptToPx, shapeRotationDeg,
        );
      }
      paintMaterial();
      // Fill and line are independent CT_ShapeProperties components. An
      // authored noFill/unresolved blip must not suppress the picture outline.
      if (hasLine) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      ctx.restore();
      return;
    }
    case 'circle':
    default: {
      ctx.beginPath();
      ctx.arc(cx, cy, half, 0, Math.PI * 2);
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
  }
  ctx.restore();
}

/** Draw error bars for one series + one direction. Each segment is a line
 *  from the data point to the offset point, plus an optional perpendicular
 *  end-cap (skipped when `eb.noEndCap`). */
export function drawSeriesErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true || (eb.linePaintAuthored === true && eb.color == null)) return;
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash, ctx.lineWidth));
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  const isX = eb.dir === 'x';
  // Office's error-bar cap spans one stroke width. Keeping the cap square with
  // the authored error-bar stroke also lets a same-size endpoint marker cover
  // it, as Excel does; the former 3× stroke-width cap protruded above/below
  // overlaid markers.
  const capHalf = ctx.lineWidth / 2;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const px = toX(xv); const py = toY(yv);
    const drawSeg = (dataDelta: number) => {
      let x2 = px, y2 = py;
      if (isX) {
        // X delta is in data X units, so map (xv + delta) → px. For the
        // minus side delta is already a positive magnitude, flip the sign.
        x2 = toX(xv + dataDelta);
      } else {
        // Y delta similar; positive moves the bar toward higher data values
        // (visually upward for our orientation).
        y2 = toY(yv + dataDelta);
      }
      ctx.beginPath();
      ctx.moveTo(px, py); ctx.lineTo(x2, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        if (isX) {
          ctx.moveTo(x2, y2 - capHalf); ctx.lineTo(x2, y2 + capHalf);
        } else {
          ctx.moveTo(x2 - capHalf, y2); ctx.lineTo(x2 + capHalf, y2);
        }
        ctx.stroke();
        ctx.restore();
      }
    };
    // ECMA-376 §21.2.2.20: plus side is `point + plus[i]`, minus side is
    // `point - minus[i]`. For `cust` errValType the values may be signed
    // (e.g. negative minus values that effectively flip direction); for
    // `fixedVal`/`stdErr`/`stdDev`/`percentage` the parser stores positive
    // magnitudes, so the same formula gives the expected direction.
    if (drawPlus) {
      const v = eb.plus[i]; if (v != null) drawSeg(v);
    }
    if (drawMinus) {
      const v = eb.minus[i]; if (v != null) drawSeg(-v);
    }
  }
  ctx.restore();
}

/** Draw per-point data labels: position-aware text near each marker. */
export function drawSeriesDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  ph: number,
  ptToPx: number,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). Threaded so date-format
   *  value labels resolve against the correct epoch. Defaults to false, which
   *  also accepts the optional `ChartModel.date1904` when it is undefined. */
  date1904 = false,
  /** Resolved data-label CSS font-family; defaults to sans-serif (byte-stable). */
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default (scatter defaults to `'r'`). */
  defaultPos = 'r',
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  layoutReferenceRect: DataLabelRect = bounds,
  richFontFamilyForFace?: (face: string) => string,
  valueDisplayUnits?: ChartDisplayUnits | null,
  legendKeyAt?: (pointIndex: number) => DataLabelLegendKey | undefined,
  isValueVisible?: (value: number) => boolean,
  shapeRotationDeg = 0,
  markerGapAt?: (pointIndex: number) => number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  if (overrides.length === 0 && !s.seriesDataLabels) return;
  const seriesDef = s.seriesDataLabels;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    if (isValueVisible && !isValueVisible(yv)) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const ovr = overridesByIndex.get(i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the point; a per-point
    // `<c:dLbl>` that only carries style / flag overrides (empty `<c:tx>`) is NOT
    // a delete — key off the explicit `deleted` flag, then honor per-point
    // show-flags (§21.2.2.47) over the series defaults.
    if (dataLabelIsDeleted(seriesDef, ovr)) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    const showBubbleSize = ovr?.showBubbleSize ?? seriesDef?.showBubbleSize;
    const showLegendKey = ovr?.showLegendKey ?? seriesDef?.showLegendKey ?? false;
    const text = effectiveDataLabelText({
      customText: ovr?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showBubbleSize,
      category: useIndexX
        ? formatCategoryLabel(
          (cats[i] ?? String(xv)).toString(),
          s.catFormatCodes?.[i] ?? s.catFormatCode ?? null,
          date1904,
        )
        : formatChartValWithCode(
          xv, s.catFormatCodes?.[i] ?? s.catFormatCode ?? null, date1904,
        ),
      seriesName: s.name,
      sourceValue: yv,
      bubbleSize: s.bubbleSizes?.[i] ?? undefined,
      valueDivisor: displayUnitDivisor(valueDisplayUnits),
      formatCode: ovr?.formatCode ?? seriesDef?.formatCode ?? null,
      date1904,
      separator: ovr?.separator ?? seriesDef?.separator,
    });
    const legendKey = showLegendKey ? legendKeyAt?.(i) : undefined;
    if (!text && !legendKey) continue;
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = chartTextFontSizePx(sizeHpt, ptToPx)
      ?? Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    const labelFace = ovr?.fontFace ?? seriesDef?.fontFace;
    const labelFont = labelFace && richFontFamilyForFace
      ? richFontFamilyForFace(labelFace)
      : fontFamily;
    drawDataLabelText(
      ctx, toX(xv), toY(yv), text, pos, fontSizePx, color, bold, labelFont,
      markerGapAt?.(i) ?? 0,
      bounds, ovr?.manualLayout,
      layoutReferenceRect,
      ovr?.richRuns,
      ptToPx,
      richFontFamilyForFace,
      legendKey,
      effectiveDataLabelTextStyle(ovr, seriesDef),
      mergeChartLabelBoxes(ovr?.labelBox, seriesDef?.labelBox),
      shapeRotationDeg,
    );
  }
}

export function drawDataLabelText(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  text: string,
  position: string,
  fontSizePx: number,
  color: string | undefined,
  bold: boolean,
  fontFamily = 'sans-serif',
  /** Extra gap (px) added to the text offset in the label's direction so the
   *  text clears an anchor glyph (e.g. a line-chart marker). The shared base
   *  inset is one half-em; markerGap is added outside that inset. */
  markerGap = 0,
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  manualLayout?: ChartDataLabelOverride['manualLayout'],
  layoutReferenceRect: DataLabelRect = bounds,
  richRuns?: readonly ChartTextRun[],
  ptToPx = 1,
  richFontFamilyForFace?: (face: string) => string,
  legendKey?: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  labelBox?: ChartLabelBox,
  shapeRotationDeg = 0,
): void {
  ctx.save();
  ctx.font = `${textStyle?.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
  drawBoundedDataLabelText(
    ctx,
    text,
    { kind: 'point', x: cx, y: cy, position, markerGap },
    bounds,
    fontSizePx,
    color ? `#${color}` : '#333',
    manualLayout,
    layoutReferenceRect,
    richRuns && richRuns.length > 0
      ? {
          runs: richRuns,
          ptToPx,
          fontFamily,
          fallbackBold: bold,
          fallbackItalic: textStyle?.fontItalic,
          fallbackBaseline: textStyle?.fontBaseline,
          fallbackColorHidden: textStyle?.fontPaintAuthored === true
            && (textStyle.fontHidden === true || textStyle.fontColor == null),
          fontFamilyForFace: richFontFamilyForFace,
        }
      : undefined,
    legendKey,
    textStyle,
    ptToPx,
    labelBox,
    shapeRotationDeg,
  );
  ctx.restore();
}

/** A `<c:tx><c:rich>` body is authoritative only with non-empty custom text.
 * Empty override text means the visible label is composed from show/format
 * flags, so stale/empty rich payload must not replace that composition. */
export function customRichDataLabelOptions(
  chart: ChartModel,
  override: ChartDataLabelOverride | undefined,
  ptToPx: number,
  fontFamily: string,
  fallbackBold: boolean,
  textStyle?: DataLabelTextStyle,
): RichDataLabelOptions | undefined {
  if (!override?.text || !override.richRuns || override.richRuns.length === 0) return undefined;
  return richDataLabelOptions(
    chart, override.richRuns, ptToPx, fontFamily, fallbackBold, textStyle,
  );
}

export function richDataLabelOptions(
  chart: ChartModel,
  runs: ChartDataLabelOverride['richRuns'],
  ptToPx: number,
  fontFamily: string,
  fallbackBold: boolean,
  textStyle?: DataLabelTextStyle,
): RichDataLabelOptions | undefined {
  if (!runs || runs.length === 0) return undefined;
  return {
    runs,
    ptToPx,
    fontFamily,
    fallbackBold,
    fallbackItalic: textStyle?.fontItalic,
    fallbackBaseline: textStyle?.fontBaseline,
    fallbackColorHidden: textStyle?.fontPaintAuthored === true
      && (textStyle.fontHidden === true || textStyle.fontColor == null),
    fontFamilyForFace: face => chartFontFamily(chart, face, 'minor'),
  };
}

/** Measure, fit, clip, and paint one label through the shared pure resolver. */
export function drawBoundedDataLabelText(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchor: DataLabelAnchor,
  bounds: DataLabelRect,
  fontSizePx: number,
  color: string,
  manualLayout?: ChartDataLabelOverride['manualLayout'],
  layoutReferenceRect: DataLabelRect = bounds,
  rich?: RichDataLabelOptions,
  legendKey?: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  textPtToPx = 1,
  labelBox?: ChartLabelBox,
  shapeRotationDeg = 0,
): void {
  if ((!text && !legendKey) || !Number.isFinite(fontSizePx) || fontSizePx <= 0) return;
  if (legendKey) {
    drawBoundedDataLabelWithLegendKey(
      ctx, text, anchor, bounds, fontSizePx, color, manualLayout,
      layoutReferenceRect, rich, legendKey,
      textStyle,
      labelBox,
    );
    return;
  }
  if (rich) {
    const block = resolveRichDataLabelBlock(ctx, rich, fontSizePx, color);
    if (!block) return;
    const insets = dataLabelInsets(textStyle, textPtToPx);
    const rotated = rotatedDataLabelSize(
      block.width + insets.left + insets.right,
      block.height + insets.top + insets.bottom,
      textStyle?.textRotation,
      textStyle?.textVerticalMode,
    );
    const placement = resolveDataLabelPlacement(
      anchor, bounds, { w: rotated.w, h: rotated.h }, fontSizePx, manualLayout,
      layoutReferenceRect,
    );
    if (!placement) return;

    ctx.save();
    ctx.beginPath();
    ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
    ctx.clip();
    paintChartLabelBox(ctx, labelBox, placement.rect, textPtToPx, shapeRotationDeg);
    const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
    const anchored = anchoredDataLabelPoint(
      placement.x, placement.y, placement.rect,
      block.height + insets.top + insets.bottom, textStyle, manualLayout != null,
      paintAlign, placement.textAlign,
      block.width + insets.left + insets.right, rotated.radians,
    );
    const transformed = transformDataLabelText(
      ctx, anchored.x, anchored.y, rotated.radians, paintAlign,
      placement.textBaseline, insets,
    );
    paintRichDataLabelBlock(
      ctx, block, transformed.x, transformed.y, paintAlign, placement.textBaseline,
      manualLayout ? Math.max(0, placement.rect.w - insets.left - insets.right) : block.width,
    );
    ctx.restore();
    return;
  }
  const lineHeight = fontSizePx * 1.15;
  const sourceLines = boundDataLabelText(text).value.split(/\r?\n/);
  const measuredW = sourceLines.reduce((max, line) => Math.max(max, ctx.measureText(line).width), 0);
  const measuredH = Math.max(lineHeight, sourceLines.length * lineHeight);
  const insets = dataLabelInsets(textStyle, textPtToPx);
  const measuredRotated = rotatedDataLabelSize(
    measuredW + insets.left + insets.right,
    measuredH + insets.top + insets.bottom,
    textStyle?.textRotation,
    textStyle?.textVerticalMode,
  );
  let placement = resolveDataLabelPlacement(
    anchor, bounds, { w: measuredRotated.w, h: measuredRotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;
  const measure = (value: string): number => ctx.measureText(value).width;
  const lines = fitStyledDataLabelLines(
    text, placement.maxWidth, placement.maxHeight, lineHeight, measure, textStyle,
  );
  if (lines.length === 0) return;
  const fittedW = lines.reduce((max, line) => Math.max(max, measure(line)), 0);
  const fittedH = lines.length * lineHeight;
  const fittedRotated = rotatedDataLabelSize(
    fittedW + insets.left + insets.right,
    fittedH + insets.top + insets.bottom,
    textStyle?.textRotation,
    textStyle?.textVerticalMode,
  );
  placement = resolveDataLabelPlacement(
    anchor, bounds, { w: fittedRotated.w, h: fittedRotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  paintChartLabelBox(ctx, labelBox, placement.rect, textPtToPx, shapeRotationDeg);
  const textPaintUnavailable = textStyle?.fontPaintAuthored === true
    && (textStyle.fontHidden === true || textStyle.fontColor == null);
  ctx.fillStyle = color;
  const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
  ctx.textAlign = paintAlign;
  ctx.textBaseline = placement.textBaseline;
  const anchored = anchoredDataLabelPoint(
    placement.x, placement.y, placement.rect,
    fittedH + insets.top + insets.bottom, textStyle, manualLayout != null,
    paintAlign, placement.textAlign,
    fittedW + insets.left + insets.right, fittedRotated.radians,
  );
  const transformed = transformDataLabelText(
    ctx, anchored.x, anchored.y, fittedRotated.radians, paintAlign,
    placement.textBaseline, insets,
  );
  const baselineShift = (textStyle?.fontBaseline ?? 0) * fontSizePx;
  const firstY = placement.textBaseline === 'middle'
    ? transformed.y - ((lines.length - 1) * lineHeight) / 2
    : placement.textBaseline === 'bottom'
      ? transformed.y - ((lines.length - 1) * lineHeight)
      : transformed.y;
  if (!textPaintUnavailable) for (let index = 0; index < lines.length; index++) {
    ctx.fillText(lines[index], transformed.x, firstY + index * lineHeight - baselineShift);
  }
  ctx.restore();
}

/** Measure and paint a data-label legend key and its optional text as one
 * bounded block. Existing legend swatch geometry is reused verbatim, while the
 * shared data-label placement resolver owns clipping and manual layout. */
export function drawBoundedDataLabelWithLegendKey(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchor: DataLabelAnchor,
  bounds: DataLabelRect,
  fontSizePx: number,
  color: string,
  manualLayout: ChartDataLabelOverride['manualLayout'] | undefined,
  layoutReferenceRect: DataLabelRect,
  rich: RichDataLabelOptions | undefined,
  legendKey: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  labelBox?: ChartLabelBox,
): void {
  const { entry, ptToPx, shapeRotationDeg } = legendKey;
  const keyWidth = legendSwatchWidths([entry], fontSizePx, ptToPx)[0] ?? 0;
  const keyHeight = legendSwatchHeight(entry, fontSizePx, ptToPx);
  const gap = text ? LEGEND_SWATCH_TEXT_GAP : 0;
  const richBlock = text && rich
    ? resolveRichDataLabelBlock(ctx, rich, fontSizePx, color)
    : null;
  if (text && rich && !richBlock) return;
  const lineHeight = fontSizePx * 1.15;
  const sourceLines = text && !richBlock
    ? boundDataLabelText(text).value.split(/\r?\n/)
    : [];
  const sourceTextWidth = richBlock?.width ?? sourceLines.reduce(
    (max, line) => Math.max(max, ctx.measureText(line).width), 0,
  );
  const sourceTextHeight = richBlock?.height
    ?? (sourceLines.length > 0 ? Math.max(lineHeight, sourceLines.length * lineHeight) : 0);
  const insets = dataLabelInsets(textStyle, ptToPx);
  const sourceWidth = keyWidth + gap + sourceTextWidth + insets.left + insets.right;
  const sourceHeight = Math.max(keyHeight, sourceTextHeight) + insets.top + insets.bottom;
  const sourceRotated = rotatedDataLabelSize(
    sourceWidth, sourceHeight, textStyle?.textRotation, textStyle?.textVerticalMode,
  );
  let placement = resolveDataLabelPlacement(
    anchor,
    bounds,
    { w: sourceRotated.w, h: sourceRotated.h },
    fontSizePx,
    manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  let lines = sourceLines;
  if (text && !richBlock) {
    lines = fitStyledDataLabelLines(
      text,
      Math.max(0, placement.maxWidth - keyWidth - gap),
      placement.maxHeight,
      lineHeight,
      value => ctx.measureText(value).width,
      textStyle,
    );
    if (lines.length === 0) return;
  }
  const textWidth = richBlock?.width ?? lines.reduce(
    (max, line) => Math.max(max, ctx.measureText(line).width), 0,
  );
  const textHeight = richBlock?.height ?? (lines.length * lineHeight);
  const contentWidth = keyWidth + gap + textWidth;
  const contentHeight = Math.max(keyHeight, textHeight);
  const totalWidth = contentWidth + insets.left + insets.right;
  const totalHeight = contentHeight + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    totalWidth, totalHeight, textStyle?.textRotation, textStyle?.textVerticalMode,
  );
  placement = resolveDataLabelPlacement(
    anchor, bounds, { w: rotated.w, h: rotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  let centerX = placement.textAlign === 'left'
    ? placement.x + rotated.w / 2
    : placement.textAlign === 'right'
      ? placement.x - rotated.w / 2
      : placement.x;
  let centerY = placement.textBaseline === 'top'
    ? placement.y + rotated.h / 2
    : placement.textBaseline === 'bottom'
      ? placement.y - rotated.h / 2
      : placement.y;
  if (manualLayout) {
    const paintAlign = dataLabelCanvasTextAlign(textStyle, 'center');
    const anchored = anchoredDataLabelPoint(
      centerX, centerY, placement.rect, totalHeight, textStyle, true, paintAlign,
    );
    centerX = paintAlign === 'left' ? anchored.x + totalWidth / 2
      : paintAlign === 'right' ? anchored.x - totalWidth / 2 : anchored.x;
    centerY = anchored.y;
  }
  const left = centerX - totalWidth / 2 + insets.left;
  const top = centerY - totalHeight / 2 + insets.top;
  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  paintChartLabelBox(ctx, labelBox, placement.rect, ptToPx, shapeRotationDeg);
  if (rotated.radians !== 0) {
    ctx.translate(centerX, centerY);
    ctx.rotate(rotated.radians);
    ctx.translate(-centerX, -centerY);
  }
  drawLegendSwatch(
    ctx,
    entry.swatchStyle,
    entry.color,
    left,
    top + (contentHeight - keyHeight) / 2,
    keyWidth,
    keyHeight,
    entry.marker,
    entry.fillPaint,
    entry.outlinePaint,
    entry.outlineColor,
    entry.outlineWidthEmu,
    entry.outlineDash,
    entry.outlineCustomDash,
    entry.outlineCap,
    entry.outlineJoin,
    ptToPx,
    shapeRotationDeg,
    entry.directEffect,
    entry.fallbackEffect,
    entry.directEffectIndex,
    entry.fallbackEffectIndex,
  );
  if (text) {
    const textX = left + keyWidth + gap;
    if (richBlock) {
      paintRichDataLabelBlock(
        ctx, richBlock, textX, top + (contentHeight - textHeight) / 2, 'left', 'top',
      );
    } else if (!(textStyle?.fontPaintAuthored === true
      && (textStyle.fontHidden === true || textStyle.fontColor == null))) {
      ctx.fillStyle = color;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'top';
      const baselineShift = (textStyle?.fontBaseline ?? 0) * fontSizePx;
      const firstY = top + (contentHeight - textHeight) / 2 - baselineShift;
      for (let index = 0; index < lines.length; index++) {
        ctx.fillText(lines[index], textX, firstY + index * lineHeight);
      }
    }
  }
  ctx.restore();
}

export function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/** Append `pts` to the CURRENT path starting from `pts[0]` (which the caller has
 *  already `moveTo`'d, or the first point is the current pen position). When
 *  `smooth` and there are ≥3 points, draw a Catmull-Rom → cubic-Bézier curve
 *  through the points (tangents from neighbours, the same formula scatter uses,
 *  ECMA-376 §21.2.2.194); otherwise straight `lineTo` segments. The caller owns
 *  `beginPath`/`moveTo`/`stroke`/`fill` so this composes into both the line
 *  stroke and the area fill's top edge. */
export function appendCurve(
  ctx: CanvasRenderingContext2D,
  pts: Array<{ x: number; y: number }>,
  smooth: boolean,
): void {
  if (pts.length === 0) return;
  if (smooth && pts.length >= 3) {
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] ?? pts[i];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[i + 2] ?? p2;
      const cp1x = p1.x + (p2.x - p0.x) / 6;
      const cp1y = p1.y + (p2.y - p0.y) / 6;
      const cp2x = p2.x - (p3.x - p1.x) / 6;
      const cp2y = p2.y - (p3.y - p1.y) / 6;
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
    }
  } else {
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
  }
}

export function dashPatternForPreset(preset: string | undefined, lineWidth = 1): number[] {
  const scale = Number.isFinite(lineWidth) && lineWidth > 0 ? lineWidth : 1;
  return pptxPresetDashArray(preset ?? 'solid', scale);
}

export function dashPatternForLine(
  customDash: ChartModel['chartBorderCustomDash'],
  preset: string | null | undefined,
  lineWidth = 1,
): number[] {
  const scale = Number.isFinite(lineWidth) && lineWidth > 0 ? lineWidth : 1;
  return drawingmlLineDashArray(customDash, preset, scale);
}

/** Draw error bars for a category-axis series (line / area). Mirrors the scatter
 *  {@link drawSeriesErrorBars} cap/dash geometry, but maps points by CATEGORY
 *  INDEX (`xAt(ci)`) with a per-series value→px mapping (`yAt`) instead of the
 *  numeric X mapping scatter uses. Only the Y direction is drawn: a category
 *  axis has no data-unit X scale, so `<c:errBars dir="x">` cannot be positioned
 *  (Excel likewise only shows Y error bars on category charts). `plotted`
 *  returns the point's plotted (possibly stacked) value so bars ride the drawn
 *  line. Null cells are skipped. */
export function drawCategoryErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true || (eb.linePaintAuthored === true && eb.color == null)
    || eb.dir === 'x') return; // no data-unit X scale on a category axis
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash, ctx.lineWidth));
  const capHalf = ctx.lineWidth / 2;
  for (let ci = 0; ci < n; ci++) {
    if (s.values[ci] == null) continue;
    const pv = plotted(ci);
    const px = xAt(ci); const py = yAt(pv);
    const drawSeg = (dataDelta: number): void => {
      const y2 = yAt(pv + dataDelta);
      ctx.beginPath(); ctx.moveTo(px, py); ctx.lineTo(px, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(px - capHalf, y2); ctx.lineTo(px + capHalf, y2);
        ctx.stroke();
        ctx.restore();
      }
    };
    if (drawPlus) { const v = eb.plus[ci]; if (v != null) drawSeg(v); }
    if (drawMinus) { const v = eb.minus[ci]; if (v != null) drawSeg(-v); }
  }
  ctx.restore();
}

/** Draw value-axis error bars for a bar/column series. CT_ErrBars `errDir`
 * follows the numeric axis: Y for columns and X for horizontal bars. Deltas
 * have already been expanded by the shared parser (percentage/fixed/stdDev/
 * custom), so this layer only maps the authored geometry. */
export function drawBarErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  n: number,
  horizontal: boolean,
  categoryAt: (ci: number) => number,
  valueAt: (value: number) => number,
  plotted: (ci: number) => number,
  fallbackColor: string,
  ptToPx: number,
): void {
  if (eb.hidden === true || (eb.linePaintAuthored === true && eb.color == null)) return;
  if ((!horizontal && eb.dir === 'x') || (horizontal && eb.dir === 'y')) return;
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu
    ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT * ptToPx)
    : Math.max(0.5, ptToPx * 0.75);
  ctx.setLineDash(dashPatternForPreset(eb.dash, ctx.lineWidth));
  const capHalf = Math.max(ctx.lineWidth / 2, 2 * ptToPx);
  for (let ci = 0; ci < n; ci++) {
    if (s.values[ci] == null) continue;
    const pv = plotted(ci);
    const category = categoryAt(ci);
    const origin = valueAt(pv);
    const drawSegment = (delta: number): void => {
      const endpoint = valueAt(pv + delta);
      ctx.beginPath();
      if (horizontal) {
        ctx.moveTo(origin, category); ctx.lineTo(endpoint, category);
      } else {
        ctx.moveTo(category, origin); ctx.lineTo(category, endpoint);
      }
      ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]); ctx.beginPath();
        if (horizontal) {
          ctx.moveTo(endpoint, category - capHalf); ctx.lineTo(endpoint, category + capHalf);
        } else {
          ctx.moveTo(category - capHalf, endpoint); ctx.lineTo(category + capHalf, endpoint);
        }
        ctx.stroke(); ctx.restore();
      }
    };
    if (drawPlus) { const value = eb.plus[ci]; if (value != null) drawSegment(value); }
    if (drawMinus) { const value = eb.minus[ci]; if (value != null) drawSegment(-value); }
  }
  ctx.restore();
}

/** Per-point data labels for a category-axis series (line / area). Consumes the
 *  same `<c:dLbl idx>` overrides and series-level `<c:dLbls>` block scatter does
 *  ({@link drawSeriesDataLabels}), but maps points by CATEGORY INDEX with the
 *  series' plotted value → px mapping. Returns true when it handled the labels
 *  for this series (so the caller skips the family's legacy `showDataLabels`
 *  path), false when the series has no override/series-level label config.
 *
 *  `plotNullAsZero` mirrors the marker loop's dispBlanksAs gate (§21.2.2.42):
 *  a null cell normally has no label (gap/span leave the point unplotted), but
 *  in "zero" mode the blank IS a plotted point (value 0) and gets a label like
 *  any other — the line-chart caller passes `dispBlanks === 'zero'`. The area
 *  caller passes `true` unconditionally: area's fill has always read a blank
 *  cell as 0 (`?? 0`, dispBlanksAs is a no-op for the filled region), so its
 *  per-point labels have likewise always covered every category index. */
export function drawCategoryDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  ph: number,
  ptToPx: number,
  date1904: boolean,
  plotNullAsZero: boolean,
  // Resolved data-label CSS font-family (element face ?? theme body ??
  // sans-serif). Defaults to sans-serif so callers that don't pass it stay
  // byte-stable.
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default. Line defaults to `'r'` (PowerPoint), area to
   *  `'ctr'`. */
  defaultPos = 't',
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  layoutReferenceRect: DataLabelRect = bounds,
  percentRatioAt?: (index: number) => number,
  markerGapAt?: (index: number) => number,
  richFontFamilyForFace?: (face: string) => string,
  valueDisplayUnits?: ChartDisplayUnits | null,
  legendKeyAt?: (pointIndex: number) => DataLabelLegendKey | undefined,
  isValueVisible?: (value: number) => boolean,
  shapeRotationDeg = 0,
): boolean {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  const seriesDef = s.seriesDataLabels;
  if (overrides.length === 0 && !seriesDef) return false;
  for (let ci = 0; ci < n; ci++) {
    if (s.sourceHidden?.[ci] === true) continue;
    if (s.values[ci] == null && !plotNullAsZero) continue;
    const anchorValue = plotted(ci);
    if (isValueVisible && !isValueVisible(anchorValue)) continue;
    const sourceValue = s.values[ci] ?? 0;
    const ovr = overridesByIndex.get(ci);
    // Genuine `<c:delete val="1"/>` (§21.2.2.43) skips; a style/flag-only
    // override is not a delete. Per-point show-flags (§21.2.2.47) win over the
    // series defaults.
    if (dataLabelIsDeleted(seriesDef, ovr)) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    const showPercent = ovr?.showPercent ?? seriesDef?.showPercent;
    const showLegendKey = ovr?.showLegendKey ?? seriesDef?.showLegendKey ?? false;
    const text = effectiveDataLabelText({
      customText: ovr?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: cats[ci] ?? '',
      seriesName: s.name,
      sourceValue,
      valueDivisor: displayUnitDivisor(valueDisplayUnits),
      percentRatio: percentRatioAt?.(ci),
      formatCode: ovr?.formatCode ?? seriesDef?.formatCode ?? null,
      date1904,
      separator: ovr?.separator ?? seriesDef?.separator,
    });
    const legendKey = showLegendKey ? legendKeyAt?.(ci) : undefined;
    if (!text && !legendKey) continue;
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = chartTextFontSizePx(sizeHpt, ptToPx)
      ?? Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    const labelFace = ovr?.fontFace ?? seriesDef?.fontFace;
    const labelFont = labelFace && richFontFamilyForFace
      ? richFontFamilyForFace(labelFace)
      : fontFamily;
    drawDataLabelText(
      ctx, xAt(ci), yAt(anchorValue), text, pos, fontSizePx, color, bold, labelFont,
      markerGapAt?.(ci) ?? 0,
      bounds, ovr?.manualLayout,
      layoutReferenceRect,
      ovr?.richRuns,
      ptToPx,
      richFontFamilyForFace,
      legendKey,
      effectiveDataLabelTextStyle(ovr, seriesDef),
      mergeChartLabelBoxes(ovr?.labelBox, seriesDef?.labelBox),
      shapeRotationDeg,
    );
  }
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

export type ChartExStyle = NonNullable<ChartModel['chartexDataPointStyle']>;

/** Effective CT_Series formatting index. The shared parser preserves authored
 * `formatIdx` and resolves omission to the original document-order index so a
 * hidden series cannot renumber the visible series' linked Chart Style. */
export function chartExSeriesFormatIndex(
  series: Pick<ChartSeries, 'chartexFormatIdx'> | null | undefined,
  fallbackIndex: number,
): number {
  return series?.chartexFormatIdx ?? fallbackIndex;
}

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

// The parser may preserve up to the OOXML cache ceiling, but expanding every
// point into several synchronous Canvas calls can monopolize the UI thread.
// Refuse an oversized paint atomically instead of drawing a misleading prefix.
// This is an availability boundary, not an automatic chart-layout heuristic.

// Marker gradients are resolved for each painted marker. Bound both one
// recipe and the chart-wide stop registrations so a valid public model cannot
// turn a bounded point count into unbounded synchronous Canvas work.
export const MAX_CANVAS_MARKER_GRADIENT_STOPS = MAX_CHART_PAINT_RECIPE_COMPONENTS;
export const MAX_CANVAS_MARKER_PAINT_COMPONENTS = MAX_CHART_PAINT_COMPONENTS;
export const MAX_CANVAS_LABEL_GRADIENT_STOPS = MAX_CANVAS_MARKER_GRADIENT_STOPS;
export const MAX_CANVAS_LABEL_PAINT_COMPONENTS = MAX_CANVAS_MARKER_PAINT_COMPONENTS;

export const CLASSIC_THREE_D_FAMILIES = new Set([
  'pie',
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
]);

export function markerKeyPaintSizesPx(
  chart: ChartModel,
  series: ChartSeries,
  ptToPx: number,
): { legend: number; table: number; labels: number } {
  const legendFontPx = chartTextFontSizePx(chart.legendFontSizeHpt, ptToPx) ?? 10 * ptToPx;
  const tableFontPx = chartTextFontSizePx(chart.dataTable?.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  let labelFontPx = chartTextFontSizePx(
    series.seriesDataLabels?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
    ptToPx,
  ) ?? 10 * ptToPx;
  for (const label of series.dataLabelOverrides ?? []) {
    labelFontPx = Math.max(
      labelFontPx,
      chartTextFontSizePx(label.fontSizeHpt, ptToPx) ?? labelFontPx,
    );
  }
  // Side legends retain at most two lines. Their marker receives 0.58× the
  // final row height; data-table and data-label keys are no larger than their
  // respective font boxes. Use those actual consumer bounds, not markerSize.
  return {
    legend: Math.max(2, (2 * legendFontPx + LEGEND_ROW_EXTRA_PX) * 0.58),
    table: Math.max(2, tableFontPx),
    labels: Math.max(2, labelFontPx),
  };
}

/** @internal Exported for resource-boundary regression tests; package entry
 * points do not expose the renderer module as public API. */
export function classicMarkerPaintWorkCount(
  chart: ChartModel,
  imageLookup?: ChartImageLookup,
  ptToPx = PT_TO_PX,
  chartRect?: ChartRect,
): number | null {
  const hasClassicMarkers = classicCanvasPointFamilyIsPainted(chart.chartType);
  const hasBoxMarkers = chart.chartexBox != null;
  if (!hasClassicMarkers && !hasBoxMarkers) return null;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const scatterHasNumericX = chart.series.some((series, seriesIndex) => {
    const group = plotGroupBySeries[seriesIndex];
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    if (family !== 'scatter') return false;
    return (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  const dataTableMarkerKeysVisible = chartHasDataTable(chart)
    && chartCategories(chart).length > 0
    && chart.dataTable?.showKeys === true;
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  const legendRanges = legendEntryRanges(chart, true);
  const legacyBubbleScale = chart.chartType === 'bubble' && chartRect
    ? bubbleSizeToDiameterScale(
        chart,
        chart.series.map((series, index) => makeScatterSeriesLayer(chart, series, index)),
        !scatterHasNumericX,
        chartRect.w,
        chartRect.h,
      )
    : 0;
  const bubbleScaleByGroup = new Map<NonNullable<ChartModel['plotGroups']>[number], number>();
  if (chartRect) for (const group of chart.plotGroups ?? []) {
    if (group.kind !== 'bubble' || group.seriesCount === 0) continue;
    const layers = chart.series
      .slice(group.seriesStart, group.seriesStart + group.seriesCount)
      .map((series, offset) => makeScatterSeriesLayer(chart, series, group.seriesStart + offset));
    bubbleScaleByGroup.set(group, bubbleSizeToDiameterScale({
      bubbleScale: group.bubbleScale ?? chart.bubbleScale,
      bubbleSizeRepresents: group.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
      showNegativeBubbles: group.showNegativeBubbles ?? chart.showNegativeBubbles,
    }, layers, !scatterHasNumericX, chartRect.w, chartRect.h));
  }
  let total = 0;
  const chargeComponents = (components: number, repetitions = 1): boolean => {
    if (repetitions <= 0 || components <= 0) return true;
    if (!Number.isSafeInteger(repetitions)
      || components > Math.floor((MAX_CANVAS_MARKER_PAINT_COMPONENTS - total) / repetitions)) {
      return false;
    }
    total += components * repetitions;
    return true;
  };
  const chargePaint = (
    paint: Fill | null | undefined,
    repetitions = 1,
    sizePx = Math.max(2, 5 * ptToPx),
  ): boolean => {
    if (repetitions <= 0 || paint == null) return true;
    const components = paint.fillType === 'image'
      ? chartImageFillPaintWorkUpperBound(paint, imageLookup, sizePx, sizePx, ptToPx)
      : markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CANVAS_MARKER_GRADIENT_STOPS) {
      return false;
    }
    return chargeComponents(components, repetitions);
  };
  if (hasClassicMarkers) for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const series = chart.series[seriesIndex];
    const group = plotGroupBySeries[seriesIndex];
    const isBubble = group?.kind === 'bubble'
      || (group == null && chart.chartType === 'bubble');
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    const effectiveChartType = markerChartTypeForPlotGroup(chart.chartType, group);
    const effectiveScatterStyle = group?.scatterStyle ?? chart.scatterStyle;
    const effectiveRadarStyle = group?.radarStyle ?? chart.radarStyle;
    const markerContext = {
      chartType: effectiveChartType,
      bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
      showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
    };
    const bubbleSettings = isBubble ? {
      bubbleScale: markerContext.bubbleScale,
      bubbleSizeRepresents: group?.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
      showNegativeBubbles: markerContext.showNegativeBubbles,
    } : undefined;
    const bubbleScale = group?.kind === 'bubble'
      ? bubbleScaleByGroup.get(group) ?? 0
      : legacyBubbleScale;
    const markerFamily = family === 'line' || family === 'stackedLine'
      || family === 'stackedLinePct' || family === 'area'
      || family === 'stackedArea' || family === 'stackedAreaPct'
      || family === 'scatter' || family === 'radar' || family === 'stock';
    if (!markerFamily) continue;
    // Family-level style choices override every series marker. Keep the
    // availability preflight on the same path as the painters: filled radar
    // never paints markers, and the two no-marker scatter styles suppress even
    // point-local marker overrides. Bubble geometry is unaffected by the
    // scatter style token and therefore remains chargeable.
    if (markersSuppressedByChartStyle(
      family, effectiveChartType, effectiveScatterStyle, effectiveRadarStyle,
    )) continue;
    const areaFamily = family === 'area' || family === 'stackedArea'
      || family === 'stackedAreaPct';
    const seriesVisible = areaFamily
      ? (series.showMarker === true || seriesHasMarkerDetail(series))
        && series.markerSymbol !== 'none'
      : family === 'stock'
        ? series.markerSymbol != null && series.markerSymbol !== 'none'
        : series.showMarker !== false && series.markerSymbol !== 'none';
    if (!seriesVisible && !hasVisiblePointMarkerOverride(series)) continue;
    const pointCount = Math.max(
      series.values.length,
      series.categories?.length ?? 0,
      chart.categories.length,
    );
    const overrides = indexPointOverrides(series.dataPointOverrides);
    const labelOverrides = new Map(
      (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
    );
    const keySizes = markerKeyPaintSizesPx(chart, series, ptToPx);
    const pointDrivenKeys = legendRanges[seriesIndex]?.pointDriven === true;
    for (let index = 0; index < pointCount; index++) {
      const plotPointVisible = classicMarkerPointIsPainted(
        chart, series, family, index, scatterHasNumericX, markerContext,
      ) && (!isBubble
        || (bubbleScale > 0
          && visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]) != null));
      const legendKeyVisible = pointDrivenKeys && chart.showLegend
        && legendEntryIsVisible(legendRanges, deletedLegendEntries, seriesIndex, index);
      const label = labelOverrides.get(index);
      const labelKeyVisible = pointDrivenKeys && plotPointVisible
        && !dataLabelIsDeleted(series.seriesDataLabels, label)
        && (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey ?? false) === true;
      const tableKeyVisible = pointDrivenKeys && dataTableMarkerKeysVisible && index === 0;
      if (!plotPointVisible && !legendKeyVisible && !labelKeyVisible && !tableKeyVisible) continue;
      if (isBubble && plotPointVisible
        && (bubbleScale <= 0
          || visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]) == null)) continue;
      const point = overrides.get(index);
      const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
      if (symbol === 'none') continue;
      const resolvedMarker = isBubble
        ? bubblePointLegendMarker(chart, series, point, index)
        : classicPointLegendMarker(
            chart, series, point, index, seriesIndex, family,
            effectiveScatterStyle, effectiveRadarStyle,
          );
      const fillPaint = resolvedMarker
        ? resolvedMarker.fillPaint
        : (isBubble ? undefined : markerFillPaintFor(series, point, index));
      const linePaint = resolvedMarker?.linePaint;
      const markerHasFill = markerSymbolConsumesFill(resolvedMarker?.symbol ?? symbol);
      const chargeMarker = (sizePx: number): boolean =>
        (!markerHasFill || chargePaint(fillPaint, 1, sizePx))
        && chargePaint(linePaint, 1, sizePx)
        && (!isBubble || resolvedMarker?.bubble3D !== true || fillPaint === null
          || chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS));
      let sizePx = Math.max(2, (point?.markerSize ?? series.markerSize ?? 5) * ptToPx);
      if (family === 'scatter' && isBubble) {
        const size = visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]);
        sizePx = size == null ? 0 : bubbleSizeMagnitude(bubbleSettings!, size) * bubbleScale;
      } else if (family === 'radar' && point?.markerSize == null && series.markerSize == null
        && chartRect) {
        sizePx = Math.max(4 * ptToPx, Math.min(chartRect.w, chartRect.h) * 0.025);
      }
      if (plotPointVisible && !chargeMarker(sizePx)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (legendKeyVisible && !chargeMarker(keySizes.legend)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (tableKeyVisible && !chargeMarker(keySizes.table)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (labelKeyVisible && !chargeMarker(keySizes.labels)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
    }

    const seriesKeySymbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
    if (!pointDrivenKeys && markerSymbolConsumesFill(seriesKeySymbol) && seriesLegendMarkerIsVisible(
      effectiveChartType, effectiveScatterStyle, series, effectiveRadarStyle,
    )) {
      const legendEntryVisible = legendEntryIsVisible(
        legendRanges, deletedLegendEntries, seriesIndex,
      );
      const labelKeys = dataLabelLegendKeyCount(
        chart, series, family, pointCount, scatterHasNumericX, markerContext,
      );
      const bubbleKeyFill = isBubble
        ? bubblePointFill(
            chart, series, undefined, seriesIndex,
            chartExSeriesFormatIndex(series, seriesIndex), chartColor(seriesIndex, series),
          )
        : null;
      const keyPaint = isBubble ? bubbleKeyFill!.paint : seriesMarkerFillPaint(series);
      const bubbleKeyLine = isBubble
        ? bubblePointLine(
            chart, series, undefined, seriesIndex,
            chartExSeriesFormatIndex(series, seriesIndex),
          )
        : null;
      const chargeKey = (repetitions: number, sizePx: number): boolean =>
        chargePaint(keyPaint, repetitions, sizePx)
        && (!isBubble || chargePaint(bubbleKeyLine!.paint, repetitions, sizePx))
        && (!isBubble || !bubblePointIsThreeD(series, undefined)
          || bubbleKeyFill!.paint === null
          || chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS, repetitions));
      if ((chart.showLegend && legendEntryVisible
          && !chargeKey(1, keySizes.legend))
        || (dataTableMarkerKeysVisible && !chargeKey(1, keySizes.table))
        || !chargeKey(labelKeys, keySizes.labels)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
    }
  }
  const box = chart.chartexBox;
  if (box) {
    const seriesCount = box.series.length;
    const markerStyle = chart.chartexDataPointMarkerStyle ?? chart.chartexDataPointStyle;
    const symbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
    if (markerSymbolConsumesFill(symbol)) {
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
        const series = box.series[seriesIndex];
        if (!series.showNonoutliers && !series.showOutliers) continue;
        let repetitions = 0;
        for (const values of series.valuesByCategory) {
          const stats = computeBoxWhiskerStats(values, series.quartileMethod);
          if (!stats) continue;
          if (series.showNonoutliers) repetitions += stats.inner.length;
          if (series.showOutliers) repetitions += stats.outliers.length;
        }
        const styleIndex = chartExSeriesFormatIndex(series, seriesIndex);
        const paint = chartExMarkerPaint(
          chart, styleIndex, seriesCount, series.chartexStyle, series.color, markerStyle,
        );
        if (!chargePaint(paint, repetitions, Math.max(2, 3 * ptToPx))) {
          return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
        }
      }
    }
  }
  return total;
}

export function chartLabelBoxPaintComponents(
  box: ChartLabelBox | null | undefined,
  imageLookup: ChartImageLookup | undefined,
  destinationWidth: number,
  destinationHeight: number,
  ptToPx: number,
): number | null {
  let total = 0;
  for (const paint of [box?.fillPaint, box?.borderFill]) {
    if (!paint) continue;
    const components = paint.fillType === 'image'
      ? chartImageFillPaintWorkUpperBound(
          paint, imageLookup, destinationWidth, destinationHeight, ptToPx,
        )
      : markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CANVAS_LABEL_GRADIENT_STOPS) {
      return null;
    }
    total += components;
  }
  return total;
}

export function dataLabelHasContent(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
  override: ChartDataLabelOverride | undefined,
): boolean {
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return false;
  return Boolean(
    override?.text
    || (override?.showVal ?? defaults?.showVal ?? chart.showDataLabels)
    || (override?.showCatName ?? defaults?.showCatName)
    || (override?.showSerName ?? defaults?.showSerName)
    || (override?.showPercent ?? defaults?.showPercent)
    || (override?.showBubbleSize ?? defaults?.showBubbleSize)
    || (override?.showLegendKey ?? defaults?.showLegendKey)
  ) && index < Math.max(series.values.length, series.categories?.length ?? 0, chart.categories.length);
}

/** Bound structured label-shape work before any family starts painting. The
 * count follows the shared 2-D/ChartEx label placement and the optional 3-D
 * label path, plus one generated box per visible 2-D trendline label. */
/** @internal Exported for resource-boundary regression tests. */
export function chartLabelPaintWorkCount(
  chart: ChartModel,
  threeD: ChartThreeDRenderer | undefined,
  imageLookup?: ChartImageLookup,
  ptToPx = PT_TO_PX,
  chartRect?: ChartRect,
): number | null {
  // ChartEx hierarchy labels are expanded only by the optional ChartEx
  // renderer, which owns their resource preflight with the hierarchy model.
  if (chart.chartexSunburst || chart.chartexTreemap) return null;
  let total = 0;
  // Label fitting is family- and text-dependent. The chart rectangle is a
  // monotonic upper bound on every generated label box and therefore on tiled
  // drawImage repetitions; a missing rectangle keeps standalone tests and
  // non-rendering callers on a small bounded estimate.
  const destinationWidth = Math.max(1, chartRect?.w ?? 32 * ptToPx);
  const destinationHeight = Math.max(1, chartRect?.h ?? 16 * ptToPx);
  const charge = (box: ChartLabelBox | null | undefined): boolean => {
    const components = chartLabelBoxPaintComponents(
      box, imageLookup, destinationWidth, destinationHeight, ptToPx,
    );
    if (components == null || components > MAX_CANVAS_LABEL_PAINT_COMPONENTS - total) {
      return false;
    }
    total += components;
    return true;
  };
  const threeDLabels = chart.threeD != null && threeD != null
    && CLASSIC_THREE_D_FAMILIES.has(chart.chartType);
  const scatterHasNumericX = chart.series.some(series => {
    const family = series.seriesType ?? chart.chartType;
    return family === 'scatter' && (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  for (let sourceSeriesIndex = 0; sourceSeriesIndex < chart.series.length; sourceSeriesIndex++) {
    const series = chart.series[sourceSeriesIndex]!;
    const overrides = indexPointOverrides(series.dataLabelOverrides);
    const family = series.seriesType ?? chart.chartType;
    const pointCount = Math.max(
      series.values.length, series.categories?.length ?? 0, chart.categories.length,
    );
    for (let index = 0; index < pointCount; index++) {
      const value = series.values[index];
      if (threeDLabels) {
        if (value == null || !Number.isFinite(value)) continue;
        if (chart.showDataLabelsOverMax !== true) {
          const maximum = series.useSecondaryAxis
            ? chart.secondaryValAxis?.max : chart.valMax;
          if (maximum != null && Number.isFinite(maximum) && value > maximum) continue;
        }
      } else if (!classicDataLabelPointIsPainted(
        chart, series, family, index, scatterHasNumericX, sourceSeriesIndex,
      )) {
        continue;
      }
      const override = overrides.get(index);
      if (!dataLabelHasContent(chart, series, index, override)) continue;
      const box = mergeChartLabelBoxes(override?.labelBox, series.seriesDataLabels?.labelBox);
      if (box && !charge(box)) return MAX_CANVAS_LABEL_PAINT_COMPONENTS + 1;
    }
    if (!threeDLabels) for (const trendline of series.trendLines ?? []) {
      const hasLabelContent = trendline.dispEq === true || trendline.dispRSqr === true
        || Boolean(trendline.labelText)
        || trendline.labelRichRuns?.some(run => run.text.length > 0) === true;
      if (hasLabelContent && trendline.labelBox
        && !charge(trendline.labelBox)) return MAX_CANVAS_LABEL_PAINT_COMPONENTS + 1;
    }
  }

  return total;
}

/** Estimate the expanded synchronous paint work for classic 3-D families.
 * A source point is not one Canvas primitive: a pie slice emits up to 32 wall
 * quads plus its top face, and a round/tapered bar emits a bounded revolved
 * mesh whose cap+facet count is owned by the 3-D renderer.
 * Apply the same 10k availability budget to that derived work before arrays
 * and sort keys are allocated. */
export function classicThreeDWorkCount(
  chart: ChartModel,
  threeD: ChartThreeDRenderer | undefined,
): number | null {
  // The expanded-face budget belongs to the optional mesh renderer. Without
  // the renderer this chart intentionally follows its canonical 2-D family, so
  // rejecting it by a cost that will never be allocated would make the
  // tree-shaken fallback less capable than an ordinary 2-D chart.
  if (!chart.threeD || !threeD) return null;
  if (!CLASSIC_THREE_D_FAMILIES.has(chart.chartType)) return null;
  let total = 0;
  for (const series of chart.series) {
    const points = Math.max(1, series.values.length, series.categories?.length ?? 0);
    const shape = series.threeDShape ?? chart.threeD.shape ?? 'box';
    const weight = chart.chartType === 'pie'
      ? THREE_D_MAX_SHAPE_FACES_PER_DATUM
      : chart.chartType.toLowerCase().includes('bar')
        ? (shape === 'box' ? 4 : THREE_D_MAX_SHAPE_FACES_PER_DATUM)
        // A normal area interval contributes only a small bounded set of
        // visible slab faces. Axis clipping can split it further, but that
        // data-dependent amplification is enforced by the renderer's exact
        // cumulative scene budget. Charging the worst-case split here would
        // reject ordinary charts long before they approach the real limit.
        : chart.chartType.toLowerCase().includes('area') ? 4
          : series.smooth === true ? 25 : 3;
    if (!Number.isSafeInteger(points) || points > Math.floor(MAX_CANVAS_CHART_POINTS / weight)) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    total += points * weight;
    if (total > MAX_CANVAS_CHART_POINTS) return MAX_CANVAS_CHART_POINTS + 1;
  }
  return total;
}

export function rejectOversizedCanvasChart(
  ctx: CanvasRenderingContext2D,
  rect: ChartRect,
  pointCount: number,
): boolean {
  if (pointCount <= MAX_CANVAS_CHART_POINTS) return false;
  ctx.fillStyle = '#888';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('(too many data points)', rect.x + rect.w / 2, rect.y + rect.h / 2);
  return true;
}

export function drawChartTextBoxes(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
): void {
  const boxes = chart.chartTextBoxes;
  if (!boxes?.length) return;

  for (const box of boxes) {
    const bx = rect.x + box.x * rect.w;
    const by = rect.y + box.y * rect.h;
    const bw = box.w * rect.w;
    const bh = box.h * rect.h;
    if (!(bw > 0 && bh > 0)) continue;
    const contentX = bx + ((box.lIns ?? DEFAULT_TEXT_INSET_LR_EMU) / EMU_PER_PT) * ptToPx;
    const contentY0 = by + ((box.tIns ?? DEFAULT_TEXT_INSET_TB_EMU) / EMU_PER_PT) * ptToPx;
    const contentRight = bx + bw - ((box.rIns ?? DEFAULT_TEXT_INSET_LR_EMU) / EMU_PER_PT) * ptToPx;
    const contentBottom = by + bh - ((box.bIns ?? DEFAULT_TEXT_INSET_TB_EMU) / EMU_PER_PT) * ptToPx;
    const contentW = contentRight - contentX;
    const contentH = contentBottom - contentY0;
    if (!(contentW > 0 && contentH > 0)) continue;

    type MeasuredTextRun = {
      run: ChartTextBox['paragraphs'][number]['runs'][number];
      text: string;
      fontPx: number;
      font: string;
      width: number;
    };
    type MeasuredLine = {
      paragraph: ChartTextBox['paragraphs'][number];
      runs: MeasuredTextRun[];
      width: number;
      height: number;
      baseline: number;
    };

    const makeLine = (
      paragraph: ChartTextBox['paragraphs'][number],
      runs: MeasuredTextRun[],
    ): MeasuredLine => {
      const maxFontPx = Math.max(1, ...runs.map(run => run.fontPx));
      return {
        paragraph,
        runs,
        width: runs.reduce((sum, run) => sum + run.width, 0),
        height: maxFontPx * 1.2,
        baseline: maxFontPx * 0.9,
      };
    };

    const lines = box.paragraphs.flatMap(paragraph => {
      const measuredRuns = paragraph.runs.map(run => {
        const fontPx = Math.max(1, ((run.fontSizeHpt ?? 1000) / 100) * ptToPx);
        const font = `${run.bold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(chart, run.fontFace, 'minor')}`;
        ctx.font = font;
        return { run, text: run.text, fontPx, font, width: ctx.measureText(run.text).width };
      });
      const paragraphWidth = measuredRuns.reduce((sum, run) => sum + run.width, 0);
      if (box.wrap === 'none' || paragraphWidth <= contentW) {
        return [makeLine(paragraph, measuredRuns)];
      }

      const wrapped: MeasuredLine[] = [];
      let current: MeasuredTextRun[] = [];
      let currentWidth = 0;
      const flush = () => {
        if (!current.length) return;
        wrapped.push(makeLine(paragraph, current));
        current = [];
        currentWidth = 0;
      };

      for (const measured of measuredRuns) {
        const tokens = measured.text.match(/\s+|\S+/g) ?? [];
        for (const token of tokens) {
          const whitespace = /^\s+$/.test(token);
          ctx.font = measured.font;
          const tokenWidth = ctx.measureText(token).width;
          if (current.length && currentWidth + tokenWidth > contentW) {
            flush();
          }
          // A wrapped line does not begin with the inter-word whitespace that
          // caused the previous line to overflow.
          if (whitespace && !current.length) continue;
          current.push({ ...measured, text: token, width: tokenWidth });
          currentWidth += tokenWidth;
        }
      }
      flush();
      return wrapped.length ? wrapped : [makeLine(paragraph, measuredRuns)];
    });
    const textHeight = lines.reduce((sum, line) => sum + line.height, 0);
    const contentY = box.verticalAnchor === 'b'
      ? contentBottom - textHeight
      : box.verticalAnchor === 'ctr'
        ? contentY0 + (contentH - textHeight) / 2
        : contentY0;

    ctx.save();
    ctx.beginPath();
    ctx.rect(bx, by, bw, bh);
    ctx.clip();
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    let lineY = contentY;
    for (const metric of lines) {
      const align = metric.paragraph.align;
      let runX = align === 'ctr'
        ? contentX + (contentW - metric.width) / 2
        : align === 'r'
          ? contentRight - metric.width
          : contentX;
      for (const measured of metric.runs) {
        ctx.font = measured.font;
        ctx.fillStyle = measured.run.color ? `#${measured.run.color}` : '#000000';
        ctx.fillText(measured.text, runX, lineY + metric.baseline);
        runX += measured.width;
      }
      lineY += metric.height;
    }
    ctx.restore();
  }
}

// ─── Background frame + dispatcher ──────────────────────────────────────────

/** ECMA-376 §21.2.2.159 defines only whether chart-space corners are rounded,
 * not the application geometry. Desktop Excel vector output uses a fixed 10pt
 * radius across square, wide, and tall chart frames; keep that observed Office
 * policy isolated from fill, border, and clipping semantics. */
export const CHART_SPACE_CORNER_RADIUS_PT = 10;

export function chartSpaceRoundedPath(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  radius: number,
): void {
  const r = Math.max(0, Math.min(radius, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}
