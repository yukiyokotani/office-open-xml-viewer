// Unified chart renderer. Dispatches on canonical `ChartModel.chartType` and
// delegates to per-family implementations (bar, line, area, pie, radar,
// scatter, waterfall). Ported from the xlsx implementation with pptx
// extensions (valMin-aware axis, plotAreaBg, dataPointColors, waterfall).

import type { ChartDataLabelOverride, ChartDecorationLineStyle, ChartDisplayUnits, ChartLabelBox, ChartLegendEntryOverride, ChartManualLayout, ChartModel, ChartRect, ChartSeries, ChartSeriesDataLabels, ChartStockUpDownBarStyle, ChartStyleRole, ChartTextBox, ChartTextRun, ChartTrendline, SecondaryValueAxis } from '../types/chart';
import type { Fill } from '../types/common';
import {
  chartImageFillSource,
  chartImageFillPaintWorkUpperBound,
  paintChartImageFill,
  type ChartImageLookup,
  withChartImageLookup,
} from './image-fill.js';
import { paintChartThreeDSurfacePicture } from './three-d-surface-picture.js';
import {
  mergeChartLabelBoxes,
  paintChartLabelBox,
} from './label-box.js';
import { strokeChartFrameRect } from './compound-frame.js';
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
} from './data-label-style.js';
import {
  bubblePointIsThreeD,
  classicMarkerPointIsPainted,
  chartDataTableFamilyIsPainted,
  dataLabelLegendKeyCount,
  deletedLegendEntryIndices,
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
} from './marker-style.js';
import {
  AXIS_OUTER_TEXT_MARGIN_PT,
  computeChartFrame,
  cartesianTitleBand,
  catAxisLabelBandH,
  chartLegendReserve,
  chartLegendBands,
  packLegendRows,
  chartAxisTitleBands,
  axisTitleFontPx,
  axisTitleRotationRad,
  chartTextFontSizePx,
  chartManualOuterAxisInsets,
  categoryTickLabelGapPx,
  axisTitleMargin,
  resolveManualLayoutRect,
  valueTickLabelGapPx,
  type ChartLegendReserve,
  type ChartAxisTitleSide,
  type ChartTitleBand,
} from './layout.js';
import {
  automaticPercentMajorUnit,
  automaticRadarMajorUnit,
  automaticSurfaceMajorUnit,
  MAX_AXIS_TICKS,
  planNumericValueAxis,
  fitTrendline,
  linearTrendlineStats,
  finiteDataExtent,
} from './axis-scale.js';
import { axisLineWidthPx, resolveAxisLine, resolveGridline, isCrossBetween } from './axis-style.js';
import {
  formatChartVal,
  formatChartValWithCode,
  formatCategoryLabel,
  formatLocalizedExcelShortDate,
} from './chart-number-format.js';
import { elideToWidth } from './text-elide.js';
import {
  categoryMinorGridlineFractions,
  categoryLabelAnchorFraction,
  categoryLabelOffsetPx,
  categoryPositionFraction,
  resolveCategoryGapWidthPercent,
  type CategoryGapPolicy,
} from './category-spacing.js';
import { planOfPieSecondaryIndices } from './of-pie.js';
import { computeBoxWhiskerStats } from './box-whisker.js';
import { planDateCategoryAxis } from './date-axis.js';
import {
  classicCanvasPointCount,
  classicCanvasPointFamilyIsPainted,
  MAX_CANVAS_CHART_POINTS,
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
  sourceChartStructureCount,
} from './resource-limits.js';
import {
  classicPlotDispatch,
  indexChartPlotGroups,
  markerChartTypeForPlotGroup,
} from './plot-groups.js';
import {
  THREE_D_MAX_SHAPE_FACES_PER_DATUM,
  type ChartThreeDRenderer,
} from './three-d-contract.js';
import type { ChartRegionMapRenderer } from './region-map-contract.js';
import type { ChartExRenderer } from './chart-ex-contract.js';
import {
  paintRichDataLabelBlock,
  resolveRichDataLabelBlock,
  type RichDataLabelBlock,
  type RichDataLabelOptions,
} from './rich-data-label.js';
import { effectiveDataLabelText } from './data-label-content.js';
import { placeTrendlineLabel } from './trendline-label.js';
import { paintLegendFrame } from './legend-frame.js';
import { paintPlotAreaFrame } from './plot-area-frame.js';
import {
  chartThreeDSurfacePaint,
  chartStyleColor,
  chartStyleFillDecision,
  chartStyleFillPaint,
  chartStyleLineDecision,
  chartStyleLinePaint,
} from './style-paint.js';
import {
  applyPlotVisibleOnly,
  hasFilteredScatterAutomaticPointStyle,
} from './source-visibility.js';
import {
  boundDataLabelText,
  resolveDataLabelPlacement,
  type DataLabelAnchor,
  type DataLabelRect,
} from './data-label-layout.js';
import { hexToRgba, resolveFill } from '../shape/paint.js';
import { drawingmlLineDashArray, pptxPresetDashArray } from '../draw/dash.js';
import {
  isObservedAutomaticSurfaceCamera,
  legacyPattern2Color,
  scaleHexColor,
  surfaceMaterialFactor,
  surfacePerspectiveTangentGain,
} from './material-color.js';
import {
  fitChartThreeDProjectionToWallThickness,
  planChartThreeDSurfaceGridSegments,
  planChartThreeDSurfaceGeometry,
  planChartThreeDProjection,
  type ThreeDScenePoint,
} from './three-d.js';
import {
  DEFAULT_TEXT_INSET_LR_EMU,
  DEFAULT_TEXT_INSET_TB_EMU,
  EMU_PER_PT,
  PT_TO_PX,
} from '../units.js';

// ─── Palette + helpers ──────────────────────────────────────────────────────

export const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

/** Office 2013+ ChartEx fallback accents when no theme/colors sidecar resolves. */
const CHARTEX_DEFAULT_PALETTE = [
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

function pieSliceColor(idx: number, series: ChartSeries, varyColors = true): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  // When varyColors is off (or the parser deliberately suppresses automatic
  // point colours for a series noFill), every unspecified slice inherits the
  // series fill. Falling straight to the built-in palette would revive a
  // noFill series and recolour a single-colour pie point by point.
  if (series.color === '00000000') return '#00000000';
  return varyColors ? `#${CHART_PALETTE[idx % CHART_PALETTE.length]}` : chartColor(idx, series);
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
type ChartFontRole = 'major' | 'minor';

/** Resolve a DrawingML theme font-scheme reference (`+mj-lt` / `+mn-lt` etc.,
 *  ECMA-376 §20.1.4.1.16) to the concrete theme face. `+mj-*` = heading
 *  (majorFont), `+mn-*` = body (minorFont); the axis suffix (`-lt`/`-ea`/`-cs`)
 *  is ignored here — chart text is Latin. A non-reference face passes through.
 *  Returns null when a reference can't be resolved (theme not threaded). */
function resolveThemeFontRef(chart: ChartModel, face: string | null | undefined): string | null | undefined {
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

/** Chart types whose legend lists one entry per category (data point of the
 *  first series) rather than one entry per series. Excel/PowerPoint draw pie
 *  and doughnut legends this way: each slice gets its own row, colored with
 *  the slice's color. ECMA-376 §21.2.2.114 (`<c:varyColors>` defaults true for
 *  pie/doughnut). */
function legendIsCategoryDriven(chartType: string | undefined): boolean {
  return chartType === 'pie' || chartType === 'doughnut';
}

/** Whether the legend is point-driven outside the pie/doughnut families.
 * §21.2.2.227 `<c:varyColors val="1"/>` drives the single-series bar case.
 * Office also exposes a lone bubble series whose `<c:xVal>` is string-backed
 * as one entry per point; the parser preserves that source provenance rather
 * than inferring it from the cached values. */
export function chartVariesColorsByPoint(chart: {
  chartType?: string | null;
  series: Array<{ bubbleXSourceIsString?: boolean | null }>;
  varyColors?: boolean | null;
}): boolean {
  if (
    chart.chartType === 'bubble' &&
    chart.series.length === 1 &&
    chart.series[0]?.bubbleXSourceIsString === true
  ) return true;
  return (
    !!chart.varyColors &&
    chart.series.length === 1 &&
    typeof chart.chartType === 'string' &&
    /Bar/.test(chart.chartType)
  );
}

/** Resolve the color for legend entry `entryIndex`, matching the marks the
 *  plot actually draws.
 *
 *  - Category-driven legends (pie / doughnut): the entry maps to data point
 *    `entryIndex` of the first series, so it must use the *same* resolution as
 *    {@link pieSliceColor} — explicit per-point `dPt` color, else the palette
 *    indexed by point. The series-level fill is deliberately ignored: a pie
 *    series carries a single `<c:spPr>` solidFill that, if honored here, would
 *    collapse every swatch to one color while the slices stay multi-colored.
 *  - Series-driven legends (bar / line / area / …): the entry maps to series
 *    `entryIndex`, so it uses {@link chartColor} — explicit series fill else
 *    the palette indexed by series. */
export function legendEntryColor(
  chartType: string | undefined,
  series: ChartSeries[],
  entryIndex: number,
  varyByPoint = false,
): string {
  if (varyByPoint || legendIsCategoryDriven(chartType)) {
    const first = series[0];
    if (first) return pieSliceColor(entryIndex, first);
    return `#${CHART_PALETTE[entryIndex % CHART_PALETTE.length]}`;
  }
  return chartColor(entryIndex, series[entryIndex]);
}

/** Draw an axis title at an explicit anchor in the outer gutter band. The
 *  side-based compatibility rotation is resolved in one place, with authored
 *  DrawingML body orientation remaining authoritative. */
function drawAxisTitle(
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
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

/** Resolve the per-axis title color string for `drawAxisTitle`. Returns
 *  '#rrggbb' when the XML supplied a srgb color, else the legacy '#555'. */
function axisTitleColor(hex: string | null | undefined): string {
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
  ): void => {
    if (side === 'left') {
      drawAxisTitle(
        ctx, text,
        x + legLeftW + axisTitleMargin(w) + fontSizePx / 2,
        py0 + ph / 2,
        side, fontSizePx, bold, italic, color, ph, fontFamily, authoredRotation,
        authoredVerticalMode, manualLayout, { x, y, w, h },
      );
      return;
    }
    drawAxisTitle(
      ctx, text,
      px0 + pw / 2,
      y + h - legBottomH - axisTitleMargin(h) - fontSizePx / 2,
      side, fontSizePx, bold, italic, color, pw, fontFamily, authoredRotation,
      authoredVerticalMode, manualLayout, { x, y, w, h },
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
    );
  }
}

type ChartDataTableLayout = {
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
function chartHasDataTable(chart: ChartModel): boolean {
  return chart.dataTable != null && chartDataTableFamilyIsPainted(chart.chartType);
}

function chartDataTableRows(chart: ChartModel): Array<{ series: ChartSeries; sourceIndex: number }> {
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
function chartDataTableBaseHeight(chart: ChartModel, ptToPx: number): number {
  const table = chartHasDataTable(chart) ? chart.dataTable : null;
  if (!table) return 0;
  const fontPx = chartTextFontSizePx(table.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const lineHeight = Math.max(1, fontPx * 1.2);
  const rowHeight = lineHeight + 4 * ptToPx;
  return (chart.series.length + 1) * rowHeight;
}

function chartDataTableHeaderWidth(
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

function measureChartDataTable(
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
function drawChartDataTable(
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
  // A direct solid dTable fill belongs to each generated body-text box. That
  // semantic is independent of the owning chart family, plot-group count,
  // manual plot layout, line wrapping, and sparse values. Unsupported fill
  // recipes remain present in the model but do not masquerade as a solid.
  const bodyFillColor = table.fillColor ?? null;
  ctx.beginPath();
  ctx.rect(tableX, tableY, tableWidth, layout.totalHeight);
  ctx.clip();
  const fontColor = table.fontColor ? `#${table.fontColor}` : '#000000';
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
        const entry = keyEntries[sourceIndex];
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
            entry.outlineColor,
            entry.outlineWidthEmu,
            entry.outlineDash,
            entry.outlineCap,
            entry.outlineJoin,
            ptToPx,
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

  if (table.lineHidden !== true) {
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
type LegendSwatchStyle = 'fill' | 'line' | 'none';

function legendSwatchStyle(chartType: string | undefined): LegendSwatchStyle {
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
interface LegendMarker {
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
  /** True when the plotted series draws both a connecting line and markers. */
  withLine: boolean;
}

/** Whether a scatter/bubble series draws a connecting line in the plot, so its
 *  legend key should be a line swatch rather than a marker glyph. Mirrors the
 *  plot gate in {@link renderScatterChart}: the group `<c:scatterStyle>` decides
 *  whether points are connected, and a series-level `<a:noFill/>` line override
 *  (§21.2.2.198, `lineHidden`) suppresses the connecting line even when the group
 *  style is `line`/`lineMarker`. Bubble charts are always markers-only. */
function scatterSeriesDrawsLine(
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
function legendMarkerFor(
  chartType: string | undefined,
  scatterStyle: string | null | undefined,
  radarStyle: string | null | undefined,
  series: ChartSeries[],
  entryIndex: number,
  chart?: ChartModel,
): LegendMarker | null {
  const s = series[entryIndex];
  if (!s) return null;
  const family = s.seriesType ?? chartType;
  const isStock = family === 'stock';
  const isLineFamily = family === 'line' || family === 'stackedLine' ||
    family === 'stackedLinePct' || family === 'radar' || isStock;
  const isBubble = family === 'bubble';
  const isScatter = family === 'scatter' || isBubble;
  if (!isLineFamily && !isScatter) return null;
  if (!seriesLegendMarkerIsVisible(chartType, scatterStyle, s, radarStyle)) return null;
  const symbol = s.markerSymbol
    ?? s.automaticMarkerSymbol
    ?? (isStock ? 'none' : 'circle');
  const base = chartColor(entryIndex, s); // '#RRGGBB'
  const fill = seriesMarkerFillColor(s, base.replace(/^#/, ''));
  const withLine = isBubble ? false : isScatter
    ? scatterSeriesDrawsLine('scatter', scatterStyle, s)
    : s.lineHidden !== true;
  if (isBubble && chart) {
    const bubbleFill = bubblePointFill(chart, s, undefined, entryIndex, base);
    const bubbleLine = bubblePointLine(chart, s, undefined, entryIndex);
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
    };
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

function bubblePointLegendMarker(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
): LegendMarker {
  const fallback = chartColor(0, series);
  const fill = bubblePointFill(chart, series, point, pointIndex, fallback);
  const line = bubblePointLine(chart, series, point, pointIndex);
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
    bubble3D: bubblePointIsThreeD(series, point),
    withLine: false,
  };
}

function drawLegendSwatch(
  ctx: CanvasRenderingContext2D,
  style: LegendSwatchStyle,
  color: string,
  x: number, y: number, w: number, h: number,
  marker: LegendMarker | null = null,
  /** undefined = no structured override (use `color`); null = authored
   * noFill; Fill = authored/resolved swatch paint. */
  fillPaint: Fill | null | undefined = undefined,
  outlineColor: string | null = null,
  outlineWidthEmu: number | null = null,
  outlineDash: string | null = null,
  outlineCap: string | null = null,
  outlineJoin: string | null = null,
  ptToPx = 1,
  shapeRotationDeg = 0,
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
    );
    return;
  }
  ctx.fillStyle = color;
  if (style === 'line') {
    // A line key is the same authored stroke as the plotted series. Only the
    // fully automatic key retains the historical font-relative fallback.
    const hasAuthoredStroke = outlineColor != null
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
      if (marker) {
        drawMarker(
          ctx, x + w / 2, y + h / 2, marker.symbol, h * 0.58 / ptToPx,
          marker.fill, marker.line, ptToPx,
          marker.lineWidthEmu != null ? axisLineWidthPx(marker.lineWidthEmu, ptToPx) : undefined,
          marker.fillPaint, shapeRotationDeg,
          marker.linePaint, marker.lineDash, marker.lineCustomDash,
          marker.lineCap, marker.lineJoin, marker.bubble3D,
        );
      }
      return;
    }
    ctx.save();
    ctx.strokeStyle = outlineColor ? `#${outlineColor}` : color;
    ctx.lineWidth = outlineWidthEmu != null
      ? axisLineWidthPx(outlineWidthEmu, ptToPx)
      : Math.max(1.5, h * 0.15);
    ctx.setLineDash(dashPatternForPreset(outlineDash ?? undefined, ctx.lineWidth));
    ctx.lineCap = outlineCap === 'rnd' ? 'round' : outlineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = outlineJoin === 'round' || outlineJoin === 'bevel'
      ? outlineJoin
      : 'miter';
    ctx.beginPath();
    const ly = y + h / 2;
    ctx.moveTo(x, ly);
    ctx.lineTo(x + w, ly);
    ctx.stroke();
    if (marker) {
      drawMarker(
        ctx, x + w / 2, y + h / 2, marker.symbol, h * 0.58 / ptToPx,
        marker.fill, marker.line, ptToPx,
        marker.lineWidthEmu != null ? axisLineWidthPx(marker.lineWidthEmu, ptToPx) : undefined,
        marker.fillPaint, shapeRotationDeg,
        marker.linePaint, marker.lineDash, marker.lineCustomDash,
        marker.lineCap, marker.lineJoin, marker.bubble3D,
      );
    }
    ctx.restore();
  } else {
    if (fillPaint !== null) {
      if (fillPaint) {
        ctx.fillStyle = fillPaint.fillType === 'solid'
          ? (fillPaint.color.startsWith('#') ? fillPaint.color : `#${fillPaint.color}`)
          : (resolveFill(fillPaint, ctx, x, y, w, h, shapeRotationDeg) ?? color);
      }
      ctx.fillRect(x, y, w, h);
    }
    if (outlineColor) {
      const outlineWidth = axisLineWidthPx(outlineWidthEmu, ptToPx);
      ctx.save();
      ctx.strokeStyle = `#${outlineColor}`;
      ctx.lineWidth = outlineWidth;
      ctx.setLineDash(dashPatternForPreset(outlineDash ?? undefined, ctx.lineWidth));
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
interface LegendEntry {
  label: string;
  color: string;
  marker: LegendMarker | null;
  swatchStyle: LegendSwatchStyle;
  fillPaint: Fill | null | undefined;
  outlineColor: string | null;
  outlineWidthEmu: number | null;
  outlineDash: string | null;
  outlineCap: string | null;
  outlineJoin: string | null;
  textOverride: ChartLegendEntryOverride | null;
}

function applyLegendEntryOverrides(
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
interface DataLabelLegendKey {
  entry: LegendEntry;
  ptToPx: number;
  shapeRotationDeg: number;
}

/** Build the legend entries for a chart. Pie/doughnut and the explicitly
 *  resolved point-driven compatibility cases use one row per data point;
 *  ordinary charts use one row per series. */
function buildLegendEntries(
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
  if (varyByPoint || legendIsCategoryDriven(chartType)) {
    // Point-driven: one entry per data point of the first series, labeled by
    // its category and colored exactly like the mark the plot draws for that
    // point (pie slice, varyColors bar, or string-X bubble).
    const first = series[0];
    const n = first ? first.values.length : 0;
    const cats = first?.categories ?? chartCategories;
    const overrides = new Map(first?.dataPointOverrides?.map(point => [point.idx, point]) ?? []);
    const entries = Array.from({ length: n }, (_, i) => {
      const point = overrides.get(i);
      const lineVisible = (point?.lineHidden ?? first?.lineHidden) !== true;
      const bubbleMarker = chartType === 'bubble' && chart && first
        ? bubblePointLegendMarker(chart, first, point, i)
        : null;
      return {
        label: (cats[i] ?? `Item ${i + 1}`).toString(),
        color: legendIsCategoryDriven(chartType) && first
          ? pieSliceColor(i, first, pieVaryColors)
          : legendEntryColor(chartType, series, i, varyByPoint),
        marker: bubbleMarker,
        swatchStyle: legendSwatchStyle(chartType),
        fillPaint: point?.fillHidden === true
          ? null
          : point?.color || first?.dataPointColors?.[i]
            ? undefined
            : i < fillPaints.length
              ? fillPaints[i]
              : first?.fillPattern ?? undefined,
        outlineColor: lineVisible ? (point?.lineColor ?? first?.lineColor ?? null) : null,
        outlineWidthEmu: lineVisible
          ? (point?.lineWidthEmu ?? first?.lineWidthEmu ?? null) : null,
        outlineDash: lineVisible
          ? (point?.lineDash ?? first?.chartexStyle?.lineDash ?? null) : null,
        outlineCap: lineVisible ? (first?.chartexStyle?.lineCap ?? null) : null,
        outlineJoin: lineVisible ? (first?.chartexStyle?.lineJoin ?? null) : null,
        textOverride: null,
      };
    });
    return applyLegendEntryOverrides(entries, entryOverrides);
  }
  const entries = series.map((s, i) => {
    // A combo chart has multiple chart groups under one plotArea. The legend
    // key describes the individual series' group, not the first/primary group.
    const family = s.seriesType ?? chartType;
    const lineVisible = s.lineHidden !== true;
    const lineColor = lineVisible ? (s.lineColor ?? null) : null;
    const marker = legendMarkerFor(chartType, scatterStyle, radarStyle, series, i, chart);
    const swatchStyle: LegendSwatchStyle = family === 'stock' && !lineVisible && !marker
      ? 'none'
      : legendSwatchStyle(family);
    return {
      label: s.name || `Series ${i + 1}`,
      color: swatchStyle === 'line' && lineColor
        ? `#${lineColor}`
        : legendEntryColor(chartType, series, i),
      marker,
      swatchStyle,
      fillPaint: i < fillPaints.length ? fillPaints[i] : (s.fillPattern ?? undefined),
      outlineColor: lineColor,
      // A DrawingML noFill line may still carry width/dash/cap/join. Those
      // geometry attributes do not make the stroke visible and must not revive
      // it in the legend through the automatic-color fallback.
      outlineWidthEmu: lineVisible ? (s.lineWidthEmu ?? null) : null,
      outlineDash: lineVisible ? (s.chartexStyle?.lineDash ?? null) : null,
      outlineCap: lineVisible ? (s.chartexStyle?.lineCap ?? null) : null,
      outlineJoin: lineVisible ? (s.chartexStyle?.lineJoin ?? null) : null,
      textOverride: null,
    };
  });
  return applyLegendEntryOverrides(entries, entryOverrides);
}

/** Resolve data-label keys through the chart's existing legend-style pipeline.
 * ECMA-376 §21.2.2.179 only requires the corresponding legend key to be shown;
 * it does not define a second paint model. Category-driven families therefore
 * use the point entry, while every other family uses its series entry. */
function createDataLabelLegendKeyResolver(
  chart: ChartModel,
  ptToPx: number,
  shapeRotationDeg = 0,
): (seriesIndex: number, pointIndex: number) => DataLabelLegendKey | undefined {
  const categoryDriven = chartVariesColorsByPoint(chart)
    || legendIsCategoryDriven(chart.chartType);
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
  return (seriesIndex, pointIndex) => {
    const entry = entries[categoryDriven ? pointIndex : seriesIndex];
    return entry ? { entry, ptToPx, shapeRotationDeg } : undefined;
  };
}

/** Resolved legend text styling (CH10). `fontFamily` already carries the
 *  theme-body fallback; `sizePx` overrides the shared automatic 10pt size only
 *  when the file authored one. */
interface LegendTextStyle {
  fontFamily: string;
  color: string;
  bold: boolean;
  sizePx: number | null;
}

function legendEntryTextStyle(
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
    sizePx: chartTextFontSizePx(override.fontSizeHpt, ptToPx) ?? base.sizePx,
  };
}

function setLegendFont(ctx: CanvasRenderingContext2D, style: LegendTextStyle, ptToPx: number): number {
  const size = legendFontSizePx(style, ptToPx);
  ctx.font = `${style.bold ? 'bold ' : ''}${size}px ${style.fontFamily}`;
  return size;
}

const DEFAULT_LEGEND_STYLE: LegendTextStyle = {
  fontFamily: 'sans-serif',
  color: '#333',
  bold: false,
  sizePx: null,
};

const LEGEND_SWATCH_TEXT_GAP = 4;
const LEGEND_ITEM_GAP = 12;
const LEGEND_ROW_EXTRA_PX = 4;
const LEGEND_HORIZONTAL_INSET = 4;
const LEGEND_HORIZONTAL_PADDING = LEGEND_HORIZONTAL_INSET * 2;
const LEGEND_VERTICAL_PADDING = 4;
const LEGEND_SIDE_PADDING = 8;
// The item width is measured as key + gap + text, then split back into those
// components for paint. Fractional Canvas metrics can lose one ULP in that
// subtraction (e.g. 24.453475952148438 becomes 24.453475952148434), falsely
// eliding a label that was already proven to fit. A hundredth of a CSS pixel is
// safely below a visible/device-pixel overflow while absorbing that arithmetic
// roundoff.
const LEGEND_MEASUREMENT_EPSILON_PX = 0.01;
// Office vector output keeps a filled legend key at roughly 7pt square even
// when the legend text is larger (for example, a 15pt legend still has a 7pt
// box). Line/marker keys remain font-relative because their glyph geometry is
// tied to the legend row rather than a filled-area key.
const FILLED_LEGEND_KEY_PT = 7;

function legendFontSizePx(style: LegendTextStyle, ptToPx: number): number {
  return style.sizePx ?? 10 * ptToPx;
}

function legendSwatchWidths(
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

function legendSwatchHeight(entry: LegendEntry, fontSize: number, ptToPx: number): number {
  return entry.swatchStyle === 'fill' ? FILLED_LEGEND_KEY_PT * ptToPx : fontSize;
}

/** Resolve side-legend text into the same bounded lines that paint consumes.
 * Office wraps a long series name at word boundaries in a left/right legend
 * (rather than replacing the whole second half with an ellipsis). Two lines
 * keep the automatic side band bounded; only text that still exceeds that
 * contract is elided on the final line. */
function sideLegendLabelLines(
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

interface MeasuredLegendLayout extends ChartLegendReserve {
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

function drawLegend(
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
          entries[index].outlineColor, entries[index].outlineWidthEmu,
          entries[index].outlineDash, entries[index].outlineCap, entries[index].outlineJoin,
          ptToPx, shapeRotationDeg,
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
  const wrapSeriesNames = !varyByPoint && !legendIsCategoryDriven(chartType);
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
      entries[i].outlineColor, entries[i].outlineWidthEmu,
      entries[i].outlineDash, entries[i].outlineCap, entries[i].outlineJoin,
      ptToPx, shapeRotationDeg,
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
function legendTextStyle(chart: ChartModel, ptToPx: number): LegendTextStyle {
  const face = resolveThemeFontRef(chart, chart.legendFontFace) ?? chart.themeMinorFontLatin;
  return {
    fontFamily: face ? `"${face}", Calibri, Arial, sans-serif` : 'sans-serif',
    color: chart.legendFontColor ? `#${chart.legendFontColor}` : '#333',
    bold: chart.legendFontBold ?? false,
    sizePx: chartTextFontSizePx(chart.legendFontSizeHpt, ptToPx),
  };
}

// Legend placement is resolved by `chartLegendReserve` (layout.ts). This alias
// keeps the drawing helper's signature readable while sharing the single source
// of truth for the reserve shape.
type LegendLayout = MeasuredLegendLayout;

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
function manualTopLegendPlotInset(
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

function axisTickLengthPx(
  level: 'major' | 'minor',
  lineWidth: number | undefined,
  ptToPx: number,
): number {
  const baseLen = (level === 'minor' ? 4 : 6) * ptToPx;
  return lineWidth ? Math.max(baseLen, lineWidth + 2 * ptToPx) : baseLen;
}

/** Distance an axis tick occupies outside the plot-side axis rule. */
function axisTickOutwardExtentPx(
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

function secondaryMinorGridStroke(
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

function secondaryMajorGridStroke(
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
function drawCatMajorGridlines(chart: ChartModel): boolean {
  return chart.catAxisMajorGridlines === true;
}

/** Resolve the CATEGORY-axis major gridline stroke, mirroring
 *  {@link valGridStroke}. `<c:catAx><c:majorGridlines><c:spPr><a:ln>` gives the
 *  color/width (`chart.catAxisGridlineColor`/`catAxisGridlineWidthEmu`); absent
 *  ⇒ the same faint `#e0e0e0`/0.5 px default as the value axis. Category
 *  gridlines have no zero-line emphasis (there is no "zero category"), so a
 *  single resolved stroke suffices. */
function catGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
  const stroke = resolveGridline(chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, ptToPx);
  return {
    ...stroke,
    dash: dashPatternForPreset(chart.catAxisGridlineDash ?? undefined, stroke.width),
  };
}

function catMinorGridStroke(chart: ChartModel, ptToPx: number): { color: string; width: number; dash: number[] } {
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
function catGridlineFractions(chart: ChartModel, n: number): number[] {
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
function valAxisReversed(chart: ChartModel): boolean {
  return chart.valAxisOrientation === 'maxMin';
}

/** True when the category axis is reversed (`<c:catAx>…orientation="maxMin">`). */
function catAxisReversed(chart: ChartModel): boolean {
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
interface ValueAxisPlan {
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
function valueAxisUnitInRendererSpace(
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

function displayUnitDivisor(units: ChartDisplayUnits | null | undefined): number {
  const divisor = units?.divisor;
  return divisor != null && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
}

function formatAxisTickWithUnits(
  value: number,
  formatCode: string | null | undefined,
  date1904: boolean | undefined,
  units: ChartDisplayUnits | null | undefined,
): string {
  return formatChartValWithCode(value / displayUnitDivisor(units), formatCode, date1904);
}

function automaticDisplayUnitLabel(units: ChartDisplayUnits): string {
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
function drawChartDisplayUnitLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
): void {
  const entries = [
    { units: chart.valAxisDisplayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.08, fallbackY: rect.y + rect.h * 0.12 },
    { units: chart.catAxisDisplayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.82 },
    { units: chart.secondaryValAxis?.displayUnits, vertical: true, fallbackX: rect.x + rect.w * 0.92, fallbackY: rect.y + rect.h * 0.12 },
    { units: chart.secondaryCatAxis?.displayUnits, vertical: false, fallbackX: rect.x + rect.w * 0.82, fallbackY: rect.y + rect.h * 0.08 },
  ];
  for (const { units, vertical, fallbackX, fallbackY } of entries) {
    const label = units?.label;
    if (!units || !label) continue;
    const text = label.text ?? automaticDisplayUnitLabel(units);
    const fontPx = chartTextFontSizePx(label.fontSizeHpt, ptToPx) ?? 10 * ptToPx;
    ctx.save();
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, label.fontFace, 'minor'),
      label.fontBold ?? false,
      label.fontItalic ?? false,
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
    if (label.boxStyle?.fill) {
      ctx.fillStyle = `#${label.boxStyle.fill}`;
      ctx.fillRect(positioned.x, positioned.y, positioned.w, positioned.h);
    }
    if (label.boxStyle?.borderColor) {
      ctx.strokeStyle = `#${label.boxStyle.borderColor}`;
      ctx.lineWidth = label.boxStyle.borderWidthEmu
        ? Math.max(0.5, label.boxStyle.borderWidthEmu / EMU_PER_PT * ptToPx)
        : 1;
      ctx.strokeRect(positioned.x, positioned.y, positioned.w, positioned.h);
    }
    ctx.translate(cx, cy);
    if (rotation !== 0) ctx.rotate(rotation);
    ctx.fillStyle = label.fontColor ? `#${label.fontColor}` : '#595959';
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

interface TrendlineLabelContext {
  chart: ChartModel;
  chartRect: ChartRect;
  plotRect: ChartRect;
  clipLineToPlot?: boolean;
  automaticAnchor?: { x: number; y: number };
  shapeRotationDeg?: number;
}

function compactTrendlineNumber(value: number, formatCode?: string | null): string {
  if (formatCode && formatCode.trim().toLowerCase() !== 'general') {
    return formatChartValWithCode(value, formatCode);
  }
  return formatChartVal(Number(value.toPrecision(6)));
}

function generatedTrendlineLabel(
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

function drawTrendlineLabel(
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
  const labelBox = effectiveLinkedLabelBox(
    chart,
    tl.labelBox,
    chart.chartStyleRoles?.trendlineLabel,
    true,
  );
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
function drawSeriesTrendlines(
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
    if (!tl.lineHidden) {
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
function trendlineLegendSeries(series: readonly ChartSeries[]): ChartSeries[] {
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
      if (trendline.lineHidden === true) continue;
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

function legendSeriesWithTrendlines(chart: ChartModel): ChartSeries[] {
  if (legendIsCategoryDriven(chart.chartType) || chartVariesColorsByPoint(chart)) {
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
function numericCategoryMetricTolerance(text: string, fontPx: number): number {
  return /^[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)%?$/.test(text)
    ? fontPx * 0.15
    : 0;
}

/** Whether the CATEGORY tick labels should be drawn. `<c:catAx><c:tickLblPos
 *  val="none">` (ECMA-376 §21.2.2.207) hides them; anything else (incl. absent)
 *  shows them, so the default is byte-stable. */
function catLabelsVisible(chart: ChartModel): boolean {
  return chart.catAxisTickLabelPos !== 'none';
}

/** 90° in 60000ths of a degree. `ST_FixedAngle` (ECMA-376 §20.1.10.23) bounds
 *  a fixed-range angle to the OPEN interval "greater than -5400000 / less than
 *  5400000", so ±5400000 itself lies outside the schema type — but Office's
 *  Format-Axis "Custom angle" control accepts -90°…+90° INCLUSIVE, so the code
 *  below deliberately uses a closed boundary (`> LIMIT` rejects, `== LIMIT`
 *  honors) to keep genuine ±90° (vertical) axis labels working. */
const FIXED_ANGLE_LIMIT_60K = 5_400_000;

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
function catLabelRotationRad(chart: ChartModel): number {
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
function drawRotatedCatLabel(
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
interface SecondaryAxisScale {
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
function chartDateAxisPlan(
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
function forEachErrorBarEndpoint(
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
function computeSecondaryAxis(
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
function drawSecondaryValueGridlines(
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
function drawSecondaryValueAxis(
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
function drawSecondaryAxisTitle(
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
  );
}

/** Draw the categorical axis paired with a secondary bar/column group. Its
 * top rule, ticks, labels and title are distinct authored objects from the
 * primary bottom category axis; the secondary value axis remains responsible
 * for the paired right-hand numeric scale. */
function drawSecondaryCategoryAxis(
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
    );
  }
}

interface ResolvedTitlePiece {
  text: string;
  width: number;
  font: string;
  color: string;
}

interface ResolvedTitleLine {
  pieces: ResolvedTitlePiece[];
  width: number;
  height: number;
}

function chartTitleRunFont(
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
      run.bold ?? (chart.titleRichRuns?.length ? false : chart.titleFontBold ?? true),
      run.italic ?? false,
    ),
    fontSize,
    color: run.color ? `#${run.color}` : chart.titleFontColor ? `#${chart.titleFontColor}` : '#333',
  };
}

/** Measure DrawingML title runs against the chart's finite title box. Explicit
 * newlines remain hard breaks; ordinary whitespace is the only automatic wrap
 * opportunity, matching DrawingML's default square text wrapping. */
function resolveChartTitleLines(
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

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  // Preserve the established single-fillText path for callers/models without
  // formatted DrawingML runs. Besides avoiding unnecessary tokenization, this
  // keeps the public Canvas contract (center-aligned title anchor) unchanged.
  if (!chart.titleRichRuns?.length) {
    const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
    const face = titleFace
      ? `"${titleFace}", Calibri, Arial, sans-serif`
      : 'Calibri, Arial, sans-serif';
    ctx.font = `${(chart.titleFontBold ?? true) ? 'bold ' : ''}${fontSize}px ${face}`;
    ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(chart.title, x + w / 2, y);
    return;
  }
  ctx.save();
  const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), fontSize);
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
    ctx.font = `${(chart.titleFontBold ?? true) ? 'bold ' : ''}${fontSize}px ${face}`;
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

function chartCategories(chart: ChartModel): string[] {
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

function dataLabelRectIntersection(a: DataLabelRect, b: DataLabelRect): DataLabelRect | null {
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
function dataLabelWithinAxisMaximum(
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
function drawBarDataLabel(
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

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

/**
 * MS ChartEx does not expose a value-axis label-offset property. Office vector
 * output from histogram, box-and-whisker, and Pareto charts (10 pt and 12 pt
 * labels, with and without cross ticks) places the visible label edge about
 * 6.2–7.5 pt from the axis centreline. Keep this compatibility fallback
 * ChartEx-only; classic axes retain their existing font-relative contract.
 */
export function chartExValueTickLabelOffsetPx(ptToPx: number): number {
  return 7 * ptToPx;
}

export function renderBarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  options: {
    gapPolicy?: CategoryGapPolicy;
    semanticLineNoStyleFallback?: boolean;
  } = {},
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const allBarSeries = chart.series.filter(s => s.seriesType !== 'line' && s.seriesType !== 'scatter' && s.seriesType !== 'area');
  const groupIsHorizontal = (series: ChartSeries): boolean => series.barGroupDirection != null
    ? series.barGroupDirection === 'bar'
    : isH;
  // The first bar group owns the visible axis orientation, but Excel retains
  // each later CT_BarChart group's own barDir. Consequently a schema-valid
  // shared-axis plot may contain vertical columns over horizontal bars. Keep
  // every group in this one data pass; category/value geometry is selected per
  // series below while axes remain owned by the first compatibility family.
  const barSeries = allBarSeries;
  const fallbackBarGroupKey = (series: ChartSeries): string =>
    series.useSecondaryAxis === true ? 'secondary-default' : 'primary-default';
  const barGroupKey = (series: ChartSeries): string => series.barGroupIndex != null
    ? `group-${series.barGroupIndex}`
    : fallbackBarGroupKey(series);
  const groupGrouping = (series: ChartSeries): string => series.barGroupGrouping
    ?? (pct ? 'percentStacked' : stacked ? 'stacked' : 'clustered');
  const groupIsStacked = (series: ChartSeries): boolean => {
    const grouping = groupGrouping(series);
    return grouping === 'stacked' || grouping === 'percentStacked';
  };
  const groupIsPercent = (series: ChartSeries): boolean =>
    groupGrouping(series) === 'percentStacked';
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');
  const areaSeries = chart.series.filter(s => s.seriesType === 'area');
  const scatterSeries = chart.series.filter(s => s.seriesType === 'scatter');
  const sourceSeriesIndices = new Map(chart.series.map((series, index) => [series, index]));
  const plotGroupBySeries = new Map<ChartSeries, NonNullable<ChartModel['plotGroups']>[number]>();
  for (const group of chart.plotGroups ?? []) {
    for (let index = group.seriesStart; index < group.seriesStart + group.seriesCount; index++) {
      const series = chart.series[index];
      if (series) plotGroupBySeries.set(series, group);
    }
  }
  const axisPercentState = new Map<string, { count: number; percentCount: number }>();
  for (const group of chart.plotGroups ?? []) {
    if (group.seriesCount === 0) continue;
    const state = axisPercentState.get(group.valueAxis) ?? { count: 0, percentCount: 0 };
    state.count++;
    if (group.grouping === 'percentStacked') state.percentCount++;
    axisPercentState.set(group.valueAxis, state);
  }
  const axisUsesPercentSpace = (group: NonNullable<ChartModel['plotGroups']>[number]): boolean => {
    const state = axisPercentState.get(group.valueAxis);
    return state != null && state.count === state.percentCount;
  };
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);

  // Combo charts (bar + line) may bind the line series to a SECONDARY value
  // axis drawn on the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with
  // axPos="r" / `<c:crosses val="max">`). `sec` is non-null only when both the
  // axis is declared AND at least one line series opts into it; horizontal bar
  // charts never carry one.
  const hasSecondarySeries = chart.series.some(series => series.useSecondaryAxis === true);
  const sec = !isH && chart.secondaryValAxis && hasSecondarySeries
    ? chart.secondaryValAxis
    : null;
  const secondaryBarSeries = sec
    ? barSeries.filter(series => series.useSecondaryAxis === true)
    : [];
  const primaryBarSeries = sec
    ? barSeries.filter(series => series.useSecondaryAxis !== true)
    : barSeries;
  const secondaryCat = secondaryBarSeries.length > 0 ? chart.secondaryCatAxis : null;
  const secondaryCategories = secondaryBarSeries[0]?.categories?.length
    ? secondaryBarSeries[0].categories
    : chart.categories;

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;
  const barGroups = new Map<string, ChartSeries[]>();
  for (const barSeriesEntry of barSeries) {
    const key = barGroupKey(barSeriesEntry);
    const group = barGroups.get(key);
    if (group) group.push(barSeriesEntry);
    else barGroups.set(key, [barSeriesEntry]);
  }
  const barGroupFor = (series: ChartSeries): ChartSeries[] =>
    barGroups.get(barGroupKey(series)) ?? [series];
  const percentDenominators = new Map<string, number[]>();
  for (const [key, members] of barGroups) {
    const totals = new Array<number>(n).fill(0);
    for (const member of members) {
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        totals[categoryIndex] += Math.abs(member.values[categoryIndex] ?? 0);
      }
    }
    percentDenominators.set(key, totals);
  }
  const percentDenominator = (series: ChartSeries, categoryIndex: number): number =>
    percentDenominators.get(barGroupKey(series))?.[categoryIndex] || 1;
  const percentGroupMultiplier = (series: ChartSeries): number => {
    const group = plotGroupBySeries.get(series);
    return group == null || axisUsesPercentSpace(group) ? 100 : 1;
  };
  const percentFactor = (series: ChartSeries, categoryIndex: number): number => {
    if (!groupIsPercent(series)) return 1;
    return percentGroupMultiplier(series) / percentDenominator(series, categoryIndex);
  };
  const overlayPlottedValues = new Map<ChartSeries, number[]>();
  const overlayBaseValues = new Map<ChartSeries, number[]>();
  for (const series of [...lineSeries, ...areaSeries]) {
    overlayPlottedValues.set(series, Array.from(
      { length: n }, (_, index) => series.values[index] ?? 0,
    ));
    overlayBaseValues.set(series, new Array<number>(n).fill(0));
  }
  for (const group of chart.plotGroups ?? []) {
    if (group.kind !== 'line' && group.kind !== 'area') continue;
    const members = chart.series.slice(group.seriesStart, group.seriesStart + group.seriesCount);
    const stackedGroup = group.grouping === 'stacked' || group.grouping === 'percentStacked';
    const percentGroup = group.grouping === 'percentStacked';
    const multiplier = percentGroup && axisUsesPercentSpace(group) ? 100 : 1;
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      const denominator = percentGroup
        ? members.reduce((sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0) || 1
        : 1;
      let running = 0;
      for (const series of members) {
        const raw = series.values[categoryIndex] ?? 0;
        const contribution = percentGroup ? raw / denominator * multiplier : raw;
        const baseValues = overlayBaseValues.get(series);
        const plottedValues = overlayPlottedValues.get(series);
        if (baseValues == null || plottedValues == null) continue;
        baseValues[categoryIndex] = stackedGroup ? running : 0;
        running = stackedGroup ? running + contribution : contribution;
        plottedValues[categoryIndex] = running;
      }
    }
  }
  const overlayValue = (series: ChartSeries, categoryIndex: number): number =>
    overlayPlottedValues.get(series)?.[categoryIndex] ?? series.values[categoryIndex] ?? 0;
  const overlayBase = (series: ChartSeries, categoryIndex: number): number =>
    overlayBaseValues.get(series)?.[categoryIndex] ?? 0;
  const effectiveBarErrorBars = (
    series: ChartSeries,
  ): NonNullable<ChartSeries['errBars']> => {
    if (!groupIsPercent(series)) return series.errBars ?? [];
    return (series.errBars ?? []).map(errorBars => ({
      ...errorBars,
      plus: errorBars.plus.map((value, index) => value == null
        ? value
        : value * percentFactor(series, index)),
      minus: errorBars.minus.map((value, index) => value == null
        ? value
        : value * percentFactor(series, index)),
    }));
  };

  // §21.2.2.227 varyColors on a single-series bar: color each bar per DATA
  // POINT (its category index) from the palette/theme sequence instead of the
  // one series color — `pieSliceColor` honors an explicit `dPt` fill first,
  // then the accent/palette for that point. Only ever true for a single bar
  // series (see {@link chartVariesColorsByPoint}), so combo/multi-series bars
  // are byte-identical.
  const varyByPoint = chartVariesColorsByPoint(chart);
  const pointOverrides = barSeries.map(series =>
    new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point])),
  );
  const labelOverrides = barSeries.map(series =>
    new Map((series.dataLabelOverrides ?? []).map(label => [label.idx, label])),
  );
  const barStyleIndices = barSeries.map((series, index) =>
    chartExSeriesFormatIndex(series, index)
  );
  const isChartExColumn = chart.chartexDataPointStyle != null
    || chart.chartexColorPalette != null;
  const styledBarLegendSeries = new Map<ChartSeries, ChartSeries>();
  if (isChartExColumn) {
    barSeries.forEach((series, index) => {
      const styleIndex = barStyleIndices[index];
      const fill = series.color
        ?? chartExDataPointFill(chart, styleIndex, barSeries.length, series.chartexStyle);
      styledBarLegendSeries.set(series, chartExLegendSeries(
        chart,
        series.name,
        series,
        chart.chartexDataPointStyle,
        styleIndex,
        barSeries.length,
        fill,
      ));
    });
  }
  const legendChart: ChartModel = {
    ...chart,
    series: (isChartExColumn ? barSeries : chart.series).map(series =>
      styledBarLegendSeries.get(series) ?? series
    ),
  };

  // Honor the parser-resolved title font size when present; otherwise use the
  // shared fixed fallback. Reserve the title band from the actual drawn size
  // so the plot shrinks to avoid overlap.
  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  let titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  let titleFontPx = titleBand.fontPx;
  let titleTopPad = titleBand.topPad;
  let titleH = titleBand.bandH;
  // Axis-label font (XML @sz when set) — sizes the bottom tick-label band the
  // same way the line/area families do.
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const categoryLevels = !isH
    && !chartHasDataTable(chart)
    && chart.catAxisNoMultiLevelLabels !== true
    && (chart.categoryLevels?.length ?? 0) > 1
    ? chart.categoryLevels!
    : null;
  const multiLevelCategoryBandH = categoryLevels
    ? (categoryLevels.length - 1) * (catAxFontPx + 4)
    : 0;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  // Axis-title bands sized from the *actual* title font (honoring XML @sz)
  // plus a small gap, so big titles get a wide enough gutter
  // and never collide with the tick labels.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  // Horizontal bars swap semantic axes: category title belongs in the left
  // band, while the horizontal value-axis title belongs in the bottom band.
  const catTitleH = isH
    ? (chart.valAxisTitle ? valTitlePx + axisTitleMargin(h) + 4 : 0)
    : axBands.catBandH;
  const valTitleW = isH
    ? (chart.catAxisTitle ? catTitlePx + axisTitleMargin(w) + 4 : 0)
    : axBands.valBandW;
  // Value-axis scales are computed up-front (before `pad`) so the side gutters
  // can be sized to the actual tick-label widths instead of a fixed fraction of
  // the chart width — short numeric labels otherwise leave a big empty gap
  // between the axis title and the labels (PowerPoint sizes the gutter to fit
  // the labels). The scales depend only on the series data, not on `pad`.
  // Vertical pads first (independent of the side gutters) so the plot height —
  // and the value-axis length — are known before the scale + label measuring.
  // The value-axis LENGTH drives the auto major unit (Excel targets a roughly
  // constant gridline spacing, so a longer axis gets finer ticks).
  // Top: title band + a small breathing gap above the topmost gridline.
  // Bottom: PowerPoint's tick-label band (gap + line-height + margin) sized to
  // the label font — the category labels for columns, the value-axis labels for
  // horizontal bars (both a single line of text). A hidden bottom axis keeps a
  // minimal gap. Matches the line/area reserve so the four families agree.
  const secondaryCatFontPx = chartTextFontSizePx(secondaryCat?.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const secondaryCatLabelBandH = secondaryCat && !secondaryCat.hidden
    && secondaryCat.tickLabelPos !== 'none'
    ? secondaryCatFontPx + categoryLabelOffsetPx(
      categoryTickLabelGapPx(secondaryCatFontPx),
      secondaryCat.labelOffsetPercent,
    ) + 2
    : 0;
  const secondaryCatTitleBandH = secondaryCat?.title
    ? axisTitleFontPx(secondaryCat.titleFontSizeHpt, ptToPx) + 6
    : 0;
  let padT = titleH + legTopH + valAxLabelFontPx / 2 + 2
    + secondaryCatLabelBandH + secondaryCatTitleBandH;
  const padB = isH
    ? (chart.valAxisHidden ? h * 0.02 : catAxisLabelBandH(valAxLabelFontPx))
      + dataTableBaseH + catTitleH + legBottomH
    : (hasDataTable ? 0 : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
      + multiLevelCategoryBandH + dataTableBaseH + catTitleH + legBottomH;
  const phEst = h - padT - padB;
  let horizontalCategoryLabelBandW = 0;
  if (isH && !chart.catAxisHidden && catLabelsVisible(chart)) {
    // The paint path derives an unauthored tick font from the category slot,
    // not the overall chart height. Use the same resolver here so tall charts
    // do not reserve a gutter for a much larger font than they actually draw.
    const measuredHorizontalCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, (phEst / n) * 0.5));
    ctx.save();
    ctx.font = chartFontCss(
      measuredHorizontalCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    for (const category of cats) {
      horizontalCategoryLabelBandW = Math.max(
        horizontalCategoryLabelBandW,
        ctx.measureText(formatCategoryLabel(category, chart.catAxisFormatCode, chart.date1904)).width,
      );
    }
    ctx.restore();
    horizontalCategoryLabelBandW += categoryLabelOffsetPx(
      chart.catAxisFontSizeHpt != null
        ? valueTickLabelGapPx(measuredHorizontalCatTickFontPx)
        : 4,
      chart.catAxisLabelOffsetPercent,
    ) + AXIS_OUTER_TEXT_MARGIN_PT * ptToPx;
  }
  // Auto layout keeps at least half of the frame available to the data plot;
  // labels beyond that measured budget are elided at paint time. Authored outer
  // layouts still use the full measured band in `manualOuterInsets` below.
  const automaticHorizontalCategoryLabelBandW = Math.min(
    horizontalCategoryLabelBandW,
    Math.max(0, w / 2 - valTitleW - legLeftW),
  );
  // Horizontal bars run the value axis along the (wide) bottom, so its length is
  // the plot WIDTH. Estimate it from the same measured category-label band that
  // the final frame uses so automatic tick density agrees with painted geometry.
  const pwEst = isH
    ? w - ((chart.catAxisHidden ? w * 0.03 : automaticHorizontalCategoryLabelBandW) + valTitleW + legLeftW) - (legRightW + w * 0.03)
    : 0;
  // A deleted value axis has no ticks whose density needs adapting to the
  // available screen length. Office falls back to its default automatic scale
  // target in that case. Feeding the plot length into the visible-tick planner
  // over-refines the major unit and stretches bars relative to slide-authored
  // overlay labels.
  const valAxisLenPt = chart.valAxisHidden ? undefined : (isH ? pwEst : phEst) / ptToPx;

  const plottedBarValue = (seriesIndex: number, categoryIndex: number): number => {
    const owner = barSeries[seriesIndex];
    const raw = owner?.values[categoryIndex] ?? 0;
    if (!owner || !groupIsStacked(owner)) return raw;
    const group = barGroupFor(owner);
    const percent = groupIsPercent(owner);
    let denominator = 1;
    if (percent) {
      denominator = group.reduce(
        (sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0,
      ) || 1;
    }
    const percentMultiplier = percent ? percentFactor(owner, categoryIndex) * denominator : 1;
    const value = percent ? raw / denominator * percentMultiplier : raw;
    let cumulative = 0;
    const ownerIndex = group.indexOf(owner);
    for (let index = 0; index <= ownerIndex; index++) {
      const candidateRaw = group[index]?.values[categoryIndex] ?? 0;
      const candidate = percent ? candidateRaw / denominator * percentMultiplier : candidateRaw;
      if ((value < 0) === (candidate < 0)) cumulative += candidate;
    }
    return cumulative;
  };

  // Value-axis extent. Bars extend from the zero line (the category-axis
  // crossing) toward each value, so the axis must span both the positive and
  // negative reach of the data (ECMA-376 §21.2.2.16 barChart). Negative values
  // pull the axis minimum below 0; positive values push the maximum above it.
  // Clustered charts take the raw extremes; stacked charts accumulate positive
  // and negative contributions on separate sides of the zero line (Excel stacks
  // opposite signs opposite ways), so `dataMax`/`dataMin` come from each
  // category's positive-sum and negative-sum.
  const primaryPlotGroups = (chart.plotGroups ?? []).filter(group =>
    group.seriesCount > 0 && group.valueAxis !== 'secondary'
  );
  const primaryPercentAxis = primaryPlotGroups.length > 0
    ? primaryPlotGroups.every(group => group.grouping === 'percentStacked')
    : primaryBarSeries.some(groupIsPercent);
  let dataMax = 0;
  let dataMin = 0;
  for (let ci = 0; ci < n; ci++) {
    const primaryGroups = new Map<string, ChartSeries[]>();
    for (const series of primaryBarSeries) {
      const key = barGroupKey(series);
      const members = primaryGroups.get(key);
      if (members) members.push(series); else primaryGroups.set(key, [series]);
    }
    for (const members of primaryGroups.values()) {
      const owner = members[0];
      const isStackedGroup = groupIsStacked(owner);
      const isPercentGroup = groupIsPercent(owner);
      const denominator = isPercentGroup
        ? members.reduce((sum, series) => sum + Math.abs(series.values[ci] ?? 0), 0) || 1
        : 1;
      let posSum = 0;
      let negSum = 0;
      for (const series of members) {
        const raw = series.values[ci] ?? 0;
        const value = isPercentGroup
          ? raw / denominator * percentGroupMultiplier(series)
          : raw;
        if (isStackedGroup) {
          if (value >= 0) posSum += value; else negSum += value;
        } else {
          dataMax = Math.max(dataMax, value);
          dataMin = Math.min(dataMin, value);
        }
      }
      if (isStackedGroup) {
        dataMax = Math.max(dataMax, posSum);
        dataMin = Math.min(dataMin, negSum);
      }
    }
  }
  // Combo line series plotted on the PRIMARY value axis (a bar+line chart whose
  // line rides the same `<c:valAx>` as the bars — no secondary axis, or one the
  // line doesn't opt into) must expand the primary axis extent just like the
  // bars do. Excel scales a shared value axis to encompass EVERY series on it,
  // regardless of chart type; a tall line point can exceed the bar stack, so
  // sizing to the bars alone would clip the line. The line is an unstacked overlay, so each raw datum widens the range
  // directly. Secondary-axis line series are excluded (they own an independent
  // scale, mirrored by the `yOf` split below). `sec` matches the draw-time gate.
  for (const s of [...lineSeries, ...areaSeries]) {
    if (sec && s.useSecondaryAxis === true) continue;
    for (let ci = 0; ci < n; ci++) {
      if (s.values[ci] == null) continue;
      const value = overlayValue(s, ci);
      dataMax = Math.max(dataMax, value);
      dataMin = Math.min(dataMin, value);
    }
  }

  // Error bars are part of the plotted value geometry. Their endpoints must
  // participate in automatic scaling; otherwise a valid endpoint can extend
  // beyond an axis planned only from the underlying series values.
  for (const series of primaryBarSeries) {
    const seriesIndex = barSeries.indexOf(series);
    for (const errorBars of effectiveBarErrorBars(series)) {
      forEachErrorBarEndpoint(
        { ...series, errBars: [errorBars] },
        isH ? 'x' : 'y',
        categoryIndex => series.values[categoryIndex] == null
          ? null
          : plottedBarValue(seriesIndex, categoryIndex),
        value => {
          dataMax = Math.max(dataMax, value);
          dataMin = Math.min(dataMin, value);
        },
      );
    }
  }
  for (const series of [...lineSeries, ...areaSeries]) {
    if (sec && series.useSecondaryAxis === true) continue;
    forEachErrorBarEndpoint(
      series,
      'y',
      index => series.values[index] ?? null,
      value => {
        const effective = value;
        dataMax = Math.max(dataMax, effective);
        dataMin = Math.min(dataMin, effective);
      },
    );
  }
  if (primaryPercentAxis) {
    if (primaryBarSeries.some(series => series.values.some(value => value != null && value > 0))) {
      dataMax = Math.max(dataMax, 100);
    }
    if (primaryBarSeries.some(series => series.values.some(value => value != null && value < 0))) {
      dataMin = Math.min(dataMin, -100);
    }
  }
  if (chart.valMax != null) {
    dataMax = primaryPercentAxis ? chart.valMax * 100 : chart.valMax;
  }
  if (chart.valMin != null) {
    dataMin = primaryPercentAxis ? chart.valMin * 100 : chart.valMin;
  }
  if (dataMax === 0 && dataMin === 0) dataMax = 1;
  // `planValueAxis` folds in the CH6 major unit / logBase / orientation; with
  // none set it is byte-identical to `valueAxisScale` + a linear map.
  const plan = planValueAxis(
    chart,
    dataMin,
    dataMax,
    valAxisLenPt,
    primaryPercentAxis,
    isH ? 'horizontal' : 'vertical',
  );
  const { step } = plan;

  // Secondary value-axis scale (combo charts). INDEPENDENT of the primary: its
  // own "nice" major unit / gridline count. Its axis is the vertical right edge,
  // so its length is the plot height. Explicit `<c:scaling>` wins. Computed by
  // the shared `computeSecondaryAxis` helper (same math the line/area families
  // reuse); the fallback keeps the no-secondary path unchanged.
  const renderedBarSeries = new Set(barSeries);
  const allBarSeriesSet = new Set(allBarSeries);
  const secondaryPlotGroups = (chart.plotGroups ?? []).filter(group =>
    group.seriesCount > 0 && group.valueAxis === 'secondary'
  );
  const secondaryPercentAxis = secondaryPlotGroups.length > 0
    ? secondaryPlotGroups.every(group => group.grouping === 'percentStacked')
    : secondaryBarSeries.some(groupIsPercent);
  const secondaryScaleSeries = chart.series
    .filter(series => !allBarSeriesSet.has(series) || renderedBarSeries.has(series))
    .map(series => {
      if (series.useSecondaryAxis !== true) return series;
      if (renderedBarSeries.has(series)) {
        const seriesIndex = barSeries.indexOf(series);
        return {
          ...series,
          values: series.values.map((value, index) => value == null
            ? value
            : plottedBarValue(seriesIndex, index)),
          errBars: effectiveBarErrorBars(series),
        };
      }
      if (!secondaryPercentAxis) return series;
      return {
        ...series,
        values: series.values.map(value => value == null ? value : value * 100),
        errBars: (series.errBars ?? []).map(errorBars => ({
          ...errorBars,
          plus: errorBars.plus.map(value => value == null ? value : value * 100),
          minus: errorBars.minus.map(value => value == null ? value : value * 100),
        })),
      };
    });
  if (secondaryPercentAxis && secondaryBarSeries[0]) {
    const percentBounds: number[] = [];
    if (secondaryBarSeries.some(series =>
      series.values.some(value => value != null && value > 0))) percentBounds.push(100);
    if (secondaryBarSeries.some(series =>
      series.values.some(value => value != null && value < 0))) percentBounds.push(-100);
    secondaryScaleSeries.push({
      ...secondaryBarSeries[0],
      values: percentBounds,
      errBars: [],
    });
  }
  const secScale = computeSecondaryAxis(
    sec,
    secondaryScaleSeries,
    phEst / ptToPx,
    isH ? 'x' : 'y',
    secondaryPercentAxis,
    secondaryBarSeries.length > 0,
  );

  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const measuredValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, phEst / 20));
  const prevFont = ctx.font;
  // Primary value-axis label band (column charts only; horizontal bars keep a
  // wider left band for the category labels).
  let valLabelTextW = 0;
  let valLabelBandW = 0;
  if (!isH && !chart.valAxisHidden) {
    // Measure with the same face the value-axis ticks draw with (below), so the
    // reserved gutter width matches the painted labels when a real face is set.
    ctx.font = chartFontCss(
      measuredValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    let wmax = 0;
    for (const val of plan.majorLines) {
      const label = formatPrimaryValueAxisTick(chart, val, primaryPercentAxis);
      wmax = Math.max(wmax, ctx.measureText(label).width);
    }
    valLabelTextW = wmax;
    valLabelBandW = valLabelTextW + 16; // ~12px tick+gap to the axis + ~4px to the title
  }
  // Secondary value-axis label band (right edge). Measure with the SAME font
  // and number format the axis is drawn with (`secFontPx` / `sec.formatCode`),
  // otherwise a `%`/thousands format or an explicit font size makes the
  // reserved gutter disagree with the painted labels.
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && !sec.hidden) {
    ctx.font = `${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    let wmax = 0;
    for (const value of secScale?.majorLines ?? []) {
      wmax = Math.max(wmax, ctx.measureText(formatAxisTickWithUnits(
        secondaryPercentAxis ? value / 100 : value,
        sec.formatCode ?? null,
        chart.date1904,
        sec.displayUnits,
      )).width);
    }
    secLabelBandW = wmax + 18;
  }
  ctx.font = prevFont;
  const secTitleBandW = sec && sec.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.03 + secLabelBandW + secTitleBandW,
    b: padB,
    // Column charts: title band + measured label band, tight to the axis.
    // Horizontal bars: keep the wider left band for the category labels
    // (`c:catAx/c:delete val="1"` → no category labels, so tighten).
    l: isH
      ? legLeftW + Math.max(
        (chart.catAxisHidden ? w * 0.03 : automaticHorizontalCategoryLabelBandW) + valTitleW,
        dataTableHeaderW,
      )
      : legLeftW + Math.max(valTitleW + valLabelBandW, dataTableHeaderW),
  };
  pad.t = manualTopLegendPlotInset(
    chart, leg, x, y, w, h, titleH, pad.t,
  );

  // `layoutTarget="outer"` includes tick labels and axis titles, but not the
  // chart title or legend. Convert only those measured axis bands to the inner
  // bar/column plot rectangle; an explicit `inner` target ignores the insets in
  // `computeChartFrame`.
  const manualOuterInsets = isH
    ? {
        t: 0,
        r: chart.valAxisHidden ? 0 : measuredValTickFontPx / 2,
        b: chart.valAxisHidden ? 0 : measuredValTickFontPx + catTitleH,
        l: chart.catAxisHidden ? 0 : horizontalCategoryLabelBandW + valTitleW,
      }
    : chartManualOuterAxisInsets({
        valAxisHidden: chart.valAxisHidden,
        catAxisHidden: chart.catAxisHidden,
        valLabelWidth: valLabelTextW,
        valLabelFontPx: measuredValTickFontPx,
        catLabelFontPx: catAxFontPx,
        valLabelGapPx: chart.valAxisFontSizeHpt != null
          ? valueTickLabelGapPx(measuredValTickFontPx)
          : 12,
        catLabelGapPx: chart.catAxisFontSizeHpt != null
          ? categoryLabelOffsetPx(
            categoryTickLabelGapPx(catAxFontPx),
            chart.catAxisLabelOffsetPercent,
          )
          : categoryLabelOffsetPx(3, chart.catAxisLabelOffsetPercent),
        outerTextMarginPx: AXIS_OUTER_TEXT_MARGIN_PT * ptToPx,
        valTitleBandW: valTitleW,
        catTitleBandH: catTitleH,
        secondaryBandW: secLabelBandW + secTitleBandW,
      });

  // Plot-area placement: honor `<c:plotArea><c:layout><c:manualLayout>` when
  // present (ECMA-376 §21.2.2.32). Templates use this to keep bars from
  // overflowing into side annotations; an explicit inner rectangle keeps the
  // data region separate from adjacent authored content.
  // `layoutTarget="inner"` (default) means the rectangle covers the inner
  // data region; "outer" includes axes/labels. We treat both identically
  // because the inner padding stays the same either way. computeChartFrame
  // applies the pad → plot rect and the manual-layout override.
  let frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    // The cartesian title band is already folded into `pad.t`; pass it so
    // `frame.title` (if read) matches the reserved band instead of a stale frac.
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  const plotWidthTitleBand = measuredCartesianTitleBand(
    ctx,
    chart,
    frame.plotRect.pw,
    h,
    ptToPx,
  );
  if (Math.abs(plotWidthTitleBand.bandH - titleBand.bandH) > 0.01) {
    titleBand = plotWidthTitleBand;
    titleFontPx = titleBand.fontPx;
    titleTopPad = titleBand.topPad;
    titleH = titleBand.bandH;
    padT = titleH + legTopH + valAxLabelFontPx / 2 + 2
      + secondaryCatLabelBandH + secondaryCatTitleBandH;
    pad.t = manualTopLegendPlotInset(
      chart, leg, x, y, w, h, titleH, padT,
    );
    frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
      titleBand,
      legendSideReserveFrac: 0.22,
      legendReserve: leg,
      pad,
      honorPlotAreaManualLayout: true,
      manualOuterInsets,
    });
  }
  const { px0, py0, pw } = frame.plotRect;
  let { ph } = frame.plotRect;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleTopPad, titleFontPx,
  );
  if (pw <= 0 || ph <= 0) return;

  // Horizontal bar categories run from bottom to top by default in the
  // existing category-slot contract; preserve that orientation while using
  // the same calendar coordinate plan as every other date-axis family.
  const dateAxisPlan = chartDateAxisPlan(
    chart,
    cats,
    isH ? !catAxisReversed(chart) : catAxisReversed(chart),
  );

  // Horizontal DrawingML category text (`wrap="square"`) wraps within its
  // category slot. Measure the complete strings with the actual tick font and
  // preserve every word instead of replacing most labels with an ellipsis.
  // An authored inner plot rectangle already leaves its own label band; for an
  // automatic/outer layout, move the plot bottom up by the additional wrapped
  // lines so they remain inside the chart frame.
  const catLabelRotation = catLabelRotationRad(chart);
  const wrappedColumnCategories: string[][] = [];
  let wrappedCategoryExtraH = 0;
  if (!hasDataTable && !isH && !dateAxisPlan && !chart.catAxisHidden && catLabelsVisible(chart) && catLabelRotation === 0) {
    const slotW = pw / n;
    const wrapFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, slotW * 0.5));
    ctx.save();
    ctx.font = chartFontCss(
      wrapFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    for (const category of cats) {
      const formattedCategory = formatCategoryLabel(
        category,
        chart.catAxisFormatCode,
        chart.date1904,
      );
      wrappedColumnCategories.push(wrapMeasuredText(
        ctx,
        formattedCategory,
        Math.max(1, slotW),
        numericCategoryMetricTolerance(formattedCategory, wrapFontPx),
      ));
    }
    ctx.restore();
    const maxLines = Math.max(1, ...wrappedColumnCategories.map(lines => lines.length));
    const manualInner = chart.plotAreaManualLayout?.layoutTarget === 'inner' &&
      chart.plotAreaManualLayout.w != null && chart.plotAreaManualLayout.h != null;
    if (!manualInner && maxLines > 1) {
      wrappedCategoryExtraH = (maxLines - 1) * (wrapFontPx + 2);
      ph = Math.max(1, ph - wrappedCategoryExtraH);
    }
  }

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // `axMax`/`step` (primary) and `sMin`/`sMax`/`sStep` (secondary) were computed
  // above the `pad` block so the gutters could be sized to the labels. The
  // line-mapping helpers need the now-final plot rect, so they live here. Line
  // series bound to the secondary axis map through `toYSecondary`; everything
  // else uses the primary `axMax`.
  // Primary value → pixel. `axRange`/`axMin` generalize the old `v / axMax`
  // mapping so the zero line sits wherever the axis crosses it (mid-plot when
  // the data straddles zero); positive-only data keeps `axMin === 0`, so the
  // mapping is unchanged. `valX`/`valY` give the on-axis pixel for a value on
  // the value axis (X for horizontal bars, Y for columns).
  const valY = (v: number): number => py0 + ph - plan.frac(v) * ph;
  const valX = (v: number): number => px0 + plan.frac(v) * pw;
  const zeroY = valY(0); // column zero line
  const zeroX = valX(0); // horizontal-bar zero line
  const toYPrimaryLine = (value: number): number => valY(value);
  // Secondary line series map through the shared scale's factory (identical to
  // the old inline `py0 + ph - ((v - sMin) / sRange) * ph`; `makeToY` uses the
  // same `(max - min) || 1` range). Falls back to the primary map when there is
  // no secondary axis so `toYSecondary` stays callable.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : valY;
  const toYSecondarySeries = (value: number): number => toYSecondary(value);

  // Resolved value-axis gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or
  // the faint `#e0e0e0`/0.5 px default). The vertical (horizontal-bar) path
  // strokes gridlines inline, so it reads `grid.color`/`grid.width` directly.
  const grid = valGridStroke(chart, ptToPx);
  ctx.textBaseline = 'middle';
  const drawnValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, ph / 20));
  ctx.font = chartFontCss(
    drawnValTickFontPx,
    chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
    chart.valAxisFontBold ?? false,
    chart.valAxisFontItalic ?? false,
  );
  // Honor `<c:valAx><c:txPr>…<a:solidFill>` when present (ECMA-376 §21.2.2.*);
  // otherwise keep the neutral gray default.
  const valLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
  ctx.fillStyle = valLabelColor;

  if (!chart.valAxisHidden) {
    // Minor gridlines (under the majors) when the file declares them.
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    for (const val of plan.minorLines) {
      if (!isH) {
        strokeValueGridlineH(ctx, px0, pw, valY(val), false, minorGrid);
      } else {
        const gx = valX(val);
        ctx.strokeStyle = minorGrid.color; ctx.lineWidth = minorGrid.width;
        const previousDash = minorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (minorGrid.dash.length > 0) ctx.setLineDash(minorGrid.dash);
        ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
        if (minorGrid.dash.length > 0) ctx.setLineDash(previousDash);
      }
    }
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const val of plan.majorLines) {
      // The zero line is the emphasized gridline (`si === 0` was that line only
      // while the axis was anchored at 0; with a negative minimum it moves up).
      const isZero = Math.abs(val) < step * 1e-9;
      const label = formatPrimaryValueAxisTick(chart, val, primaryPercentAxis);
      if (!isH) {
        const gy = valY(val);
        if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, isZero, grid);
        if (drawLabels) {
          ctx.textAlign = 'right';
          const gap = options.gapPolicy === 'chartex'
            ? chartExValueTickLabelOffsetPx(ptToPx)
            : chart.valAxisFontSizeHpt != null
              ? valueTickLabelGapPx(drawnValTickFontPx)
              : 12;
          ctx.fillText(label, px0 - gap, gy);
        }
      } else {
        const gx = valX(val);
        if (drawMajorGrid) {
          // Explicit gridline color ⇒ uniform stroke (no zero-line emphasis),
          // matching PowerPoint; otherwise keep the `#aaa`/1 px baseline rule.
          ctx.strokeStyle = grid.explicit ? grid.color : isZero ? '#aaa' : grid.color;
          ctx.lineWidth = grid.explicit ? grid.width : isZero ? 1 : grid.width;
          const previousDash = grid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
          if (grid.dash.length > 0) ctx.setLineDash(grid.dash);
          ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
          if (grid.dash.length > 0) ctx.setLineDash(previousDash);
        }
        if (drawLabels) {
          ctx.textAlign = 'center';
          const gap = chart.valAxisFontSizeHpt != null
            ? categoryTickLabelGapPx(drawnValTickFontPx)
            : 10;
          ctx.fillText(label, gx, py0 + ph + gap);
        }
      }
    }
  }

  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100).
  // Perpendicular to the value gridlines: vertical for a column chart (cat axis
  // runs along x), horizontal for a horizontal-bar chart (cat axis runs along
  // y). Positioned at the same fractions as the category ticks — band
  // boundaries under crossBetween="between" (bar default), category centers
  // under "midCat". Drawn under the bars (like value gridlines). Office omits
  // these by default so the common path is byte-stable.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    const previousDash = cg.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (cg.dash.length > 0) ctx.setLineDash(cg.dash);
    const gridlineFractions = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
      : catGridlineFractions(chart, n);
    for (const frac of gridlineFractions) {
      ctx.beginPath();
      if (!isH) {
        const gx = px0 + frac * pw;
        ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph);
      } else {
        const gy = py0 + frac * ph;
        ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy);
      }
      ctx.stroke();
    }
    if (cg.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // Axis rules. The CATEGORY axis runs along the bars' baseline — bottom
  // (horizontal) for a column chart, left (vertical) for a horizontal bar
  // chart — and the VALUE axis is perpendicular to it. The previous code
  // assumed the left rule was always the value axis, so a horizontal bar
  // chart whose value axis is `<c:delete val="1">` drew
  // no axis line at all even though its category axis carries an explicit
  // `<c:spPr><a:ln>`. `<a:noFill>` on a line suppresses just the rule (labels
  // stay) → `*AxisLineHidden`; an `<a:solidFill>` gives `*AxisLineColor`/Width
  // (ECMA-376 §21.2.2.* line props). Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  // Colour defaults to '#aaa' (Office's faint default rule); the EMU `<a:ln@w>`
  // is scaled to canvas px by `ptToPx`. See `resolveAxisLine`.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null;
  const primaryCatCrossValue = categoryAxisCrossingValue(chart, plan.min, plan.max);
  const primaryCatAxisY = !isH ? valY(primaryCatCrossValue) : py0 + ph;
  const primaryCatAxisX = isH ? valX(primaryCatCrossValue) : px0;
  const categoryLabelAxisY = !isH && (chart.catAxisTickLabelPos ?? 'nextTo') === 'nextTo'
    ? primaryCatAxisY
    : py0 + ph;
  const catMajorTickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
  const multiLevelBoundariesOwnMajorTicks = !isH
    && categoryLevels != null
    && dateAxisPlan == null
    && !chart.catAxisLineHidden
    && isCrossBetween(chart)
    && Math.abs(categoryLabelAxisY - primaryCatAxisY) < 0.01;
  // Axis rules + tick marks are drawn AFTER the bars/line (see `drawAxesOnTop`
  // below) so the bars don't paint over the category baseline — PowerPoint
  // keeps the axis line crisp on top of the columns.
  const drawAxesOnTop = (): void => {
    if (!isH) {
      if (drawCatLine) strokeAxisSegment(ctx, px0, primaryCatAxisY, px0 + pw, primaryCatAxisY, catLineColor, catLineW, chart.catAxisLineDash);
      if (drawValLine) strokeAxisSegment(ctx, px0, py0, px0, py0 + ph, valLineColor, valLineW, chart.valAxisLineDash);           // left
    } else {
      if (drawCatLine) strokeAxisSegment(ctx, primaryCatAxisX, py0, primaryCatAxisX, py0 + ph, catLineColor, catLineW, chart.catAxisLineDash);
      if (drawValLine) strokeAxisSegment(ctx, px0, py0 + ph, px0 + pw, py0 + ph, valLineColor, valLineW, chart.valAxisLineDash); // bottom
    }

    // Axis major tick marks (`<c:*Ax><c:majorTickMark>` — ECMA-376 §21.2.2.101).
    // PowerPoint draws short ruler ticks even when the axis rule itself is light,
    // so the bar renderer must emit them too (the line renderer already does).
    // `drawAxisTick`'s `axis` arg selects GEOMETRY: 'val' = vertical rule with
    // horizontal ticks, 'cat' = horizontal rule with vertical ticks. For a
    // column chart the value axis is vertical (left) and the category axis
    // horizontal (bottom); a horizontal bar chart swaps the two.
    if (!chart.valAxisHidden && chart.valAxisMajorTickMark && chart.valAxisMajorTickMark !== 'none') {
      for (const val of plan.majorLines) {
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, valY(val), valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'cat', py0 + ph, valX(val), valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
        }
      }
    }
    if (!chart.valAxisHidden && chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of plan.minorTicks) {
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, valY(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.valAxisMinorTickMark, 'cat', py0 + ph, valX(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
        }
      }
    }
    // Category-axis major ticks share the major-grid positions: with
    // crossBetween="between" they mark the N+1 interval boundaries, while
    // labels and minor ticks sit at the N category centres. This distinction
    // matters when both levels use `cross`: Office's 6pt major boundary marks
    // remain visibly longer than its 4pt centred minor marks.
    if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
      // Multi-level category brackets already occupy every interval boundary.
      // They absorb the coincident major tick into one continuous stroke below,
      // avoiding darker/thicker Canvas seams from painting the same segment
      // twice. Mid-category ticks remain independent of those boundaries.
      const ordinalFractions = multiLevelBoundariesOwnMajorTicks
        ? []
        : catGridlineFractions(chart, n);
      const tickFractions = dateAxisPlan
        ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
        : ordinalFractions.filter((_, index) => index % catMajorTickSkip === 0);
      for (const frac of tickFractions) {
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCatAxisY, px0 + frac * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'val', primaryCatAxisX, py0 + frac * ph, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
        }
      }
    }
    if (!chart.catAxisHidden && chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none') {
      const ordinalFractions = isCrossBetween(chart)
        ? Array.from({ length: n }, (_, ci) => (ci + 0.5) / n)
        : Array.from({ length: Math.max(0, n - 1) }, (_, ci) => (ci + 0.5) / (n - 1));
      const tickFractions = dateAxisPlan
        ? dateAxisPlan.minorTicks.map(tick => tick.fraction)
        : ordinalFractions;
      for (const frac of tickFractions) {
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMinorTickMark, 'cat', primaryCatAxisY, px0 + frac * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.catAxisMinorTickMark, 'val', primaryCatAxisX, py0 + frac * ph, catLineColor, catLineW, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
        }
      }
    }
  };

  // Bar cluster geometry — ECMA-376 §21.2.2.13 (gapWidth = % of bar width
  // between categories, default 150) and §21.2.2.25 (overlap = signed % of
  // bar width within a cluster, default 0). Within a cluster the pitch
  // between consecutive bars is `barW * (1 - overlap/100)`, so with N series:
  //   clusterWidth = barW + (N - 1) * barW * (1 - overlap/100)
  //   catGap       = clusterWidth + barW * gapWidth/100
  //                = barW * (1 + (N-1) * (1 - overlap/100) + gapWidth/100)
  // Solving for barW gives the formula below. Stacked charts render one bar
  // per category so we treat them as N=1 and overlap=0.
  const categoryGap = (horizontal: boolean): number => horizontal ? ph / n : pw / n;
  const catGap = categoryGap(isH);
  const catRev = catAxisReversed(chart);
  const categorySlotIndex = (ci: number, horizontal: boolean): number => horizontal
    ? (catRev ? ci : n - 1 - ci)
    : (catRev ? n - 1 - ci : ci);
  const categoryBandSize = (ci: number, horizontal = isH): number => dateAxisPlan
    ? dateAxisPlan.categoryBandFractions[ci]! * (horizontal ? ph : pw)
    : categoryGap(horizontal);
  const categoryStart = (ci: number, horizontal = isH): number => dateAxisPlan
    ? (horizontal ? py0 : px0)
      + dateAxisPlan.positions[ci]! * (horizontal ? ph : pw)
      - categoryBandSize(ci, horizontal) / 2
    : (horizontal ? py0 : px0)
      + categorySlotIndex(ci, horizontal) * categoryGap(horizontal);
  const categoryCenterX = (ci: number): number => dateAxisPlan
    ? px0 + dateAxisPlan.positions[ci]! * pw
    : px0 + categorySlotIndex(ci, false) * categoryGap(false) + categoryGap(false) / 2;
  const clusterGeometry = (group: readonly ChartSeries[], categorySize: number) => {
    const owner = group[0];
    const isStackedGroup = owner ? groupIsStacked(owner) : stacked;
    const effective = isStackedGroup ? 1 : Math.max(1, group.length);
    const rawOverlap = owner?.barGroupOverlap ?? chart.barOverlap ?? 0;
    const overlapPct = isStackedGroup || !Number.isFinite(rawOverlap)
      ? 0
      : Math.max(-100, Math.min(100, rawOverlap));
    const gapWidthPct = resolveCategoryGapWidthPercent(
      owner?.barGroupGapWidth ?? chart.barGapWidth,
      options.gapPolicy ?? 'legacy',
    );
    const denom = 1 + (effective - 1) * (1 - overlapPct / 100) + gapWidthPct / 100;
    const barW = categorySize / denom;
    const clusterGap = isStackedGroup ? 0 : barW * (1 - overlapPct / 100);
    const clusterWidth = barW + (effective - 1) * clusterGap;
    return { barW, clusterGap, catStart: (categorySize - clusterWidth) / 2 };
  };
  type BarSeriesLinePoint = {
    categoryStart: number;
    categoryEnd: number;
    valueEnd: number;
  };
  const hasVerifiedBarSeriesLines = (chart.barGroupDecorations ?? []).some(
    decoration => decoration.seriesLines?.length === 1,
  );
  const barSeriesLinePoints: Array<Array<BarSeriesLinePoint | null>> | null =
    hasVerifiedBarSeriesLines
      ? barSeries.map(() => new Array<BarSeriesLinePoint | null>(n).fill(null))
      : null;

  // A classic combo chart may place an `<c:areaChart>` group behind a
  // `<c:barChart>` group. Area series are not bars: they share the category
  // coordinate system, map through their bound value axis, and fill from the
  // zero baseline to their authored top edge. Paint them before columns so the
  // later bar group remains visible, matching Excel's foreground columns for
  // this mixed-family layout.
  for (let si = 0; si < areaSeries.length; si++) {
    const series = areaSeries[si];
    const pointOverrides = indexPointOverrides(series.dataPointOverrides);
    const color = chartColor(sourceSeriesIndices.get(series) ?? si, series);
    const yOf = sec && series.useSecondaryAxis === true
      ? toYSecondarySeries
      : toYPrimaryLine;
    const dispBlanks = chart.dispBlanksAs ?? 'zero';
    let run: Array<{ x: number; y: number; baseY: number }> = [];
    const paintRun = (): void => {
      if (run.length === 0) return;
      if (dateAxisPlan) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(px0, py0, pw, ph);
        ctx.clip();
      }
      ctx.beginPath();
      ctx.moveTo(run[0].x, run[0].baseY);
      ctx.lineTo(run[0].x, run[0].y);
      appendCurve(ctx, run, false);
      for (let index = run.length - 1; index >= 0; index--) {
        ctx.lineTo(run[index].x, run[index].baseY);
      }
      ctx.closePath();
      ctx.fillStyle = series.fillPattern
        ? (resolveFill(
          series.fillPattern,
          ctx,
          run[0].x,
          py0,
          Math.max(1, run[run.length - 1].x - run[0].x),
          ph,
        ) ?? color)
        : color;
      ctx.fill();
      if (series.lineHidden !== true) {
        ctx.strokeStyle = series.lineColor ? `#${series.lineColor}` : color;
        ctx.lineWidth = series.lineWidthEmu != null
          ? axisLineWidthPx(series.lineWidthEmu, ptToPx)
          : 1.5;
        ctx.setLineDash([]);
        ctx.stroke();
      }
      if (dateAxisPlan) ctx.restore();
      run = [];
    };
    for (let ci = 0; ci < n; ci++) {
      const value = series.values[ci];
      if (value == null) {
        if (dispBlanks === 'gap') paintRun();
        if (dispBlanks !== 'zero') continue;
      }
      run.push({
        x: categoryCenterX(ci),
        y: yOf(overlayValue(series, ci)),
        baseY: yOf(overlayBase(series, ci)),
      });
    }
    paintRun();

    const seriesMarkersVisible = (series.showMarker === true || seriesHasMarkerDetail(series))
      && series.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(series)) {
      const markerRadius = Math.max(2, 2.5 * ptToPx);
      for (let ci = 0; ci < n; ci++) {
        const value = series.values[ci];
        if (value == null) continue;
        const point = pointOverrides.get(ci);
        const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        const markerX = categoryCenterX(ci);
        const markerY = yOf(overlayValue(series, ci));
        if (seriesHasMarkerDetail(series) || pointHasMarkerDetail(point)) {
          const lineWidthEmu = point?.markerLineWidthEmu ?? series.markerLineWidthEmu;
          drawMarker(
            ctx, markerX, markerY, symbol,
            point?.markerSize ?? series.markerSize ?? 5,
            markerFillColorFor(series, point, ci, color),
            point?.markerLine ?? series.markerLine ?? null,
            ptToPx,
            lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
            markerFillPaintFor(series, point, ci),
            shapeRotationDeg,
          );
        } else {
          ctx.fillStyle = color;
          ctx.beginPath();
          ctx.arc(markerX, markerY, markerRadius, 0, Math.PI * 2);
          ctx.fill();
        }
      }
    }
  }

  for (let ci = 0; ci < n; ci++) {
    // Stacked charts accumulate positive and negative contributions on opposite
    // sides of the zero line, so each category tracks two running offsets.
    const positiveOffsets = new Map<string, number>();
    const negativeOffsets = new Map<string, number>();
    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const seriesIsHorizontal = groupIsHorizontal(s);
      const categorySize = categoryBandSize(ci, seriesIsHorizontal);
      const secondary = sec != null && s.useSecondaryAxis === true;
      const group = barGroupFor(s);
      const groupKey = barGroupKey(s);
      const isStackedGroup = groupIsStacked(s);
      const isPercentGroup = groupIsPercent(s);
      const stackSum = isPercentGroup
        ? group.reduce((sum, member) => sum + Math.abs(member.values[ci] ?? 0), 0) || 1
        : 1;
      const posOffset = positiveOffsets.get(groupKey) ?? 0;
      const negOffset = negativeOffsets.get(groupKey) ?? 0;
      const groupIndex = Math.max(0, group.indexOf(s));
      const { barW, clusterGap, catStart } = clusterGeometry(group, categorySize);
      const valueY = secondary && secScale ? secScale.makeToY(py0, ph) : valY;
      // `<c:catAx><c:crosses>` locates the category axis on its paired value
      // axis (ECMA-376 §21.2.2.33/.34). A secondary top axis commonly uses
      // `crosses="max"`; its columns therefore start at the maximum rule and
      // extend downward. Treating every secondary group as zero-based reverses
      // that authored geometry even though the top rule itself is painted.
      const secondaryBase = secScale
        ? secondaryCat?.crossesAt != null && Number.isFinite(secondaryCat.crossesAt)
          ? Math.max(secScale.min, Math.min(secScale.max, secondaryCat.crossesAt))
          : secondaryCat?.crosses === 'max'
            ? secScale.max
            : secondaryCat?.crosses === 'min'
              ? secScale.min
              : Math.max(secScale.min, Math.min(secScale.max, 0))
        : 0;
      const groupZeroY = secondary ? valueY(secondaryBase) : zeroY;
      const raw = s.values[ci] ?? 0;
      // Signed value in axis units (percent keeps its sign — a negative slice of
      // a percentStacked chart reaches below the zero line).
      const sv = isPercentGroup
        ? (raw / stackSum) * percentGroupMultiplier(s)
        : raw;
      const negative = sv < 0;
      const labelAxisMaximum = secondary && secScale ? secScale.max : plan.max;
      const useNegativeStyle = negative
        && (s.invertIfNegative === true || s.automaticNegativeStyle === true);
      // A `<c:dPt>` fill is an explicit point override regardless of the
      // chart-group `varyColors` flag (§21.2.2.52). varyColors only controls
      // the fallback palette for points without an override.
      const pointOverride = pointOverrides[si].get(ci);
      const pointColor = pointOverride?.color ?? s.dataPointColors?.[ci];
      const color = pointColor
        ? `#${pointColor}`
        : varyByPoint ? pieSliceColor(ci, s) : chartColor(si, s);
      const invertedPaint = useNegativeStyle
        ? s.automaticNegativeStyle === true || s.invertedFillHidden === true
          ? null
          : s.invertedFill
        : undefined;
      const pointPaint = pointOverride?.fillHidden
        ? null
        : pointOverride?.color
          ? { fillType: 'solid' as const, color: pointOverride.color }
          : invertedPaint !== undefined
            ? invertedPaint
          : isChartExColumn
            // MS-ODRAWXML §2.8.4.5: styleClr="auto" uses the relative
            // index of the styled element. A clustered-column dataPoint style
            // colors a series, so every point in that series resolves with
            // the series index; CT_DataPoint direct formatting above remains
            // point-local and wins.
            ? chartExDataPointPaint(
              chart, barStyleIndices[si], barSeries.length, s.chartexStyle, s.color,
            )
            : undefined;
      const applyPointOutline = (): boolean => {
        const hasPointLine = pointOverride?.lineHidden != null
          || pointOverride?.lineColor != null
          || pointOverride?.lineWidthEmu != null
          || pointOverride?.lineDash != null;
        if (hasPointLine) {
          if (pointOverride?.lineHidden) return false;
          const fallback = chartExStyleColor(
            chart, chart.chartexDataPointStyle, 'line', barStyleIndices[si], barSeries.length,
          )
            ?? s.lineColor
            ?? color;
          ctx.strokeStyle = `#${pointOverride?.lineColor ?? fallback.replace(/^#/, '')}`;
          ctx.lineWidth = pointOverride?.lineWidthEmu != null
            ? axisLineWidthPx(pointOverride.lineWidthEmu, ptToPx)
            : s.lineWidthEmu != null
              ? axisLineWidthPx(s.lineWidthEmu, ptToPx)
              : 1;
          ctx.setLineDash(dashPatternForPreset(pointOverride?.lineDash, ctx.lineWidth));
          return true;
        }
        if (useNegativeStyle
          && (s.invertedLineHidden != null
            || s.invertedLineColor != null
            || s.invertedLineWidthEmu != null)) {
          if (s.invertedLineHidden) return false;
          ctx.strokeStyle = `#${s.invertedLineColor ?? '000000'}`;
          ctx.lineWidth = axisLineWidthPx(s.invertedLineWidthEmu, ptToPx);
          ctx.setLineDash([]);
          return true;
        }
        if (s.automaticNegativeStyle === true) {
          ctx.strokeStyle = '#000000';
          ctx.lineWidth = 0.75 * ptToPx;
          ctx.setLineDash([]);
          return true;
        }
        const omittedAlternateLineFallback = useNegativeStyle
          && chart.chartType === 'clusteredBar'
          && chart.legacyChartStyle === 2
          && s.invertedFillAuthored === true
          && s.invertedFill != null
          && s.invertedLineAuthored === false;
        const seriesOwnsOutline = s.lineHidden === true
          || s.lineColor != null
          || s.lineWidthEmu != null;
        if (omittedAlternateLineFallback && seriesOwnsOutline) {
          if (s.lineHidden || !s.lineColor) return false;
          ctx.strokeStyle = `#${s.lineColor}`;
          ctx.lineWidth = axisLineWidthPx(s.lineWidthEmu, ptToPx);
          ctx.setLineDash([]);
          return true;
        }
        // Office 2010's alternate negative-fill extension records no line
        // provenance when `<a:ln>` is omitted. The observed Style 2 clustered
        // column boundary supplies a black 0.75pt outline; keep that
        // application default here, after direct point/series line ownership,
        // instead of forging an authored line in the shared parser model.
        if (omittedAlternateLineFallback) {
          ctx.strokeStyle = '#000000';
          ctx.lineWidth = 0.75 * ptToPx;
          ctx.setLineDash([]);
          return true;
        }
        if (isChartExColumn) {
          return applyChartExSeriesLineStyle(
            ctx, chart, chart.chartexDataPointStyle, s,
            barStyleIndices[si], barSeries.length, color, ptToPx,
          );
        }
        if (!s.lineColor || s.lineHidden) return false;
        ctx.strokeStyle = `#${s.lineColor}`;
        ctx.lineWidth = axisLineWidthPx(s.lineWidthEmu, ptToPx);
        ctx.setLineDash([]);
        return true;
      };

      if (!seriesIsHorizontal) {
        const bx = isStackedGroup
          ? categoryStart(ci, false) + catStart
          : categoryStart(ci, false) + catStart + groupIndex * clusterGap;
        // A date axis can explicitly crop categories through min/max. Marks
        // wholly outside that authored plot interval do not bleed into the
        // value-axis/title gutter.
        if (bx + barW <= px0 || bx >= px0 + pw) continue;
        // Column: the bar spans between the zero line and the value. Stacked
        // bars start at the running offset for their sign; clustered bars start
        // at the zero line.
        const y0 = isStackedGroup ? valueY(negative ? negOffset : posOffset) : groupZeroY;
        const y1 = isStackedGroup
          ? valueY((negative ? negOffset : posOffset) + sv)
          : valueY(sv);
        const by = clamp(Math.min(y0, y1), py0, py0 + ph);
        const barBottom = clamp(Math.max(y0, y1), py0, py0 + ph);
        const barH = Math.max(0, barBottom - by);
        if (barSeriesLinePoints && s.values[ci] != null) {
          barSeriesLinePoints[si][ci] = {
            categoryStart: bx,
            categoryEnd: bx + barW,
            valueEnd: clamp(y1, py0, py0 + ph),
          };
        }
        if (pointPaint !== null) {
          ctx.fillStyle = pointPaint
            ? chartExFillStyle(ctx, pointPaint, bx, by, barW, barH, color)
            : s.fillPattern
              ? (resolveFill(s.fillPattern, ctx, bx, by, barW, barH) ?? color)
              : color;
          ctx.fillRect(bx, by, barW, barH);
        }
        if (barW > 0 && barH > 0 && applyPointOutline()) {
          const outlineW = ctx.lineWidth;
          ctx.strokeRect(
            bx + outlineW / 2,
            by + outlineW / 2,
            Math.max(0, barW - outlineW),
            Math.max(0, barH - outlineW),
          );
        }
        const seriesLabels = s.seriesDataLabels;
        const label = resolveChartExLabel(
          chart, s, ci, s.categories?.[ci] ?? cats[ci] ?? '', raw,
          {
            visible: chart.showDataLabels,
            showVal: chart.showDataLabels && !isPercentGroup,
            showPercent: chart.showDataLabels && isPercentGroup,
            showCatName: false,
          },
          labelOverrides[si],
          isPercentGroup ? sv / 100 : undefined,
          s.useSecondaryAxis && sec ? sec.displayUnits : chart.valAxisDisplayUnits,
        );
        if (label && dataLabelWithinAxisMaximum(chart, raw, labelAxisMaximum)) {
          // ECMA-376 §21.2.2.30 / §21.1.2.3.2 — data label font size comes from
          // `<c:dLbls><c:txPr>...<a:defRPr@sz>` (hundredths of a point). When
          // the file specifies one we honor it; otherwise the proportional
          // heuristic keeps small bars readable.
          const sizeHpt = label.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = chartTextFontSizePx(sizeHpt, ptToPx)
            ?? Math.max(7, Math.min(11, barW * 0.6));
          const authoredLabel = labelOverrides[si].get(ci);
          const bold = label.fontBold
            || (seriesLabels?.fontBold == null && authoredLabel?.fontBold == null);
          const labelFont = chartFontFamily(
            chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
          );
          ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${lsz}px ${labelFont}`;
          // drawBarDataLabel takes (bx, by, barL=length, barW=thickness). For
          // a vertical column bar, "length" is the bar's height and
          // "thickness" is its horizontal width — pass them in that order.
          // Previously the args were (barW, barH) which silently swapped the
          // two and made `cx = bx + barW/2` (the horizontal-center formula
          // inside the helper) use the bar's HEIGHT instead of its width,
          // pushing data labels far to the right of the bar.
          drawBarDataLabel(
            ctx, label.text,
            bx, by, barH, barW,
            'vertical',
            label.position ?? chart.dataLabelPosition ?? (isStackedGroup ? 'ctr' : null),
            s.dataLabelColors?.[ci] ?? label.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            lsz,
            { x: px0, y: py0, w: pw, h: ph },
            { x, y, w, h },
            authoredLabel?.manualLayout,
            negative,
            customRichDataLabelOptions(
              chart,
              authoredLabel,
              ptToPx,
              labelFont,
              bold,
              label.textStyle,
            ),
            label.showLegendKey
              ? dataLabelLegendKey(sourceSeriesIndices.get(s) ?? si, ci)
              : undefined,
            label.textStyle,
            ptToPx,
            mergeChartLabelBoxes(authoredLabel?.labelBox, seriesLabels?.labelBox),
            shapeRotationDeg,
          );
        }
      } else {
        // Cluster positions are local to the owning `<c:barChart>` group.
        // A second overlay group restarts at its own first bar; using the
        // flattened plot-area series index pushes it outside the category
        // slot when preceding groups contain more series.
        const siVisual = groupIndex;
        const by = isStackedGroup
          ? categoryStart(ci, true) + catStart
          : categoryStart(ci, true) + catStart + siVisual * clusterGap;
        const x0 = isStackedGroup ? valX(negative ? negOffset : posOffset) : zeroX;
        const x1 = isStackedGroup ? valX((negative ? negOffset : posOffset) + sv) : valX(sv);
        const bx = clamp(Math.min(x0, x1), px0, px0 + pw);
        const barRight = clamp(Math.max(x0, x1), px0, px0 + pw);
        const barL = Math.max(0, barRight - bx);
        if (barSeriesLinePoints && s.values[ci] != null) {
          barSeriesLinePoints[si][ci] = {
            categoryStart: by,
            categoryEnd: by + barW,
            valueEnd: clamp(x1, px0, px0 + pw),
          };
        }
        if (pointPaint !== null) {
          ctx.fillStyle = pointPaint
            ? chartExFillStyle(ctx, pointPaint, bx, by, barL, barW, color)
            : s.fillPattern
              ? (resolveFill(s.fillPattern, ctx, bx, by, barL, barW) ?? color)
              : color;
          ctx.fillRect(bx, by, barL, barW);
        }
        if (barL > 0 && barW > 0 && applyPointOutline()) {
          const outlineW = ctx.lineWidth;
          ctx.strokeRect(
            bx + outlineW / 2,
            by + outlineW / 2,
            Math.max(0, barL - outlineW),
            Math.max(0, barW - outlineW),
          );
        }
        const seriesLabels = s.seriesDataLabels;
        const label = resolveChartExLabel(
          chart, s, ci, s.categories?.[ci] ?? cats[ci] ?? '', raw,
          {
            visible: chart.showDataLabels,
            showVal: chart.showDataLabels && !isPercentGroup,
            showPercent: chart.showDataLabels && isPercentGroup,
            showCatName: false,
          },
          labelOverrides[si],
          isPercentGroup ? sv / 100 : undefined,
          s.useSecondaryAxis && sec ? sec.displayUnits : chart.valAxisDisplayUnits,
        );
        if (label && dataLabelWithinAxisMaximum(chart, raw, labelAxisMaximum)) {
          const sizeHpt = label.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = chartTextFontSizePx(sizeHpt, ptToPx)
            ?? Math.max(7, Math.min(11, barW * 0.6));
          const authoredLabel = labelOverrides[si].get(ci);
          const bold = label.fontBold
            || (seriesLabels?.fontBold == null && authoredLabel?.fontBold == null);
          const labelFont = chartFontFamily(
            chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
          );
          ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${lsz}px ${labelFont}`;
          drawBarDataLabel(
            ctx, label.text,
            bx, by, barL, barW,
            'horizontal',
            label.position ?? chart.dataLabelPosition ?? (isStackedGroup ? 'ctr' : null),
            s.dataLabelColors?.[ci] ?? label.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            lsz,
            { x: px0, y: py0, w: pw, h: ph },
            { x, y, w, h },
            authoredLabel?.manualLayout,
            negative,
            customRichDataLabelOptions(
              chart,
              authoredLabel,
              ptToPx,
              labelFont,
              bold,
              label.textStyle,
            ),
            label.showLegendKey
              ? dataLabelLegendKey(sourceSeriesIndices.get(s) ?? si, ci)
              : undefined,
            label.textStyle,
            ptToPx,
            mergeChartLabelBoxes(authoredLabel?.labelBox, seriesLabels?.labelBox),
            shapeRotationDeg,
          );
        }
      }
      if (isStackedGroup) {
        if (negative) negativeOffsets.set(groupKey, negOffset + sv);
        else positiveOffsets.set(groupKey, posOffset + sv);
      }
    }
  }

  // `CT_BarChart/serLines` uses one group-owned line style. MS-OE376
  // 2.1.1578 defines each segment as joining adjacent data points in the same
  // series. Office vector output resolves those points to the value-end edge of
  // each stacked bar and clips the segment to the category gap: columns join
  // right/left facing edges, horizontal bars join bottom/top facing edges.
  // Missing points break the sequence instead of inventing a zero-valued end.
  if (barSeriesLinePoints) {
    const barSeriesByGroup = new Map<number, number[]>();
    for (let seriesIndex = 0; seriesIndex < barSeries.length; seriesIndex++) {
      const groupIndex = barSeries[seriesIndex].barGroupIndex ?? 0;
      const members = barSeriesByGroup.get(groupIndex);
      if (members) members.push(seriesIndex);
      else barSeriesByGroup.set(groupIndex, [seriesIndex]);
    }
    for (const decoration of chart.barGroupDecorations ?? []) {
      // CT_BarChart permits multiple serLines children. The single-child Office
      // geometry is verified; precedence/association for multiple children is
      // application-defined and remains fail-closed until its boundary output is
      // adjudicated rather than guessing first/last/cyclic style semantics.
      if (decoration.seriesLines?.length !== 1) {
        continue;
      }
      ctx.save();
      const seriesLineStyle = chartStyleRoleLine(
        chart, decoration.seriesLines[0], 'seriesLine',
      );
      if (!applyDecorationLineStyle(ctx, seriesLineStyle, ptToPx)) {
        ctx.restore();
        continue;
      }
      ctx.beginPath();
      ctx.rect(px0, py0, pw, ph);
      ctx.clip();
      for (const seriesIndex of barSeriesByGroup.get(decoration.groupIndex) ?? []) {
        const points = barSeriesLinePoints[seriesIndex];
        const seriesIsHorizontal = groupIsHorizontal(barSeries[seriesIndex]);
        for (let categoryIndex = 0; categoryIndex + 1 < n; categoryIndex++) {
          const current = points[categoryIndex];
          const next = points[categoryIndex + 1];
          if (!current || !next) continue;
          const currentCenter = (current.categoryStart + current.categoryEnd) / 2;
          const nextCenter = (next.categoryStart + next.categoryEnd) / 2;
          const forward = nextCenter >= currentCenter;
          ctx.beginPath();
          if (!seriesIsHorizontal) {
            ctx.moveTo(
              forward ? current.categoryEnd : current.categoryStart,
              current.valueEnd,
            );
            ctx.lineTo(
              forward ? next.categoryStart : next.categoryEnd,
              next.valueEnd,
            );
          } else {
            ctx.moveTo(
              current.valueEnd,
              forward ? current.categoryEnd : current.categoryStart,
            );
            ctx.lineTo(
              next.valueEnd,
              forward ? next.categoryStart : next.categoryEnd,
            );
          }
          ctx.stroke();
        }
      }
      ctx.restore();
    }
  }

  // CT_BarSer permits the same trendline and error-bar children as line
  // series. Paint both above the filled rectangles and below axes/labels.
  // Geometry is derived from the same gapWidth/overlap cluster calculation as
  // the bars, so a clustered series' adornments remain centered on its bars.
  const barCategoryCenter = (series: ChartSeries, categoryIndex: number): number => {
    const group = barGroupFor(series);
    const groupIndex = Math.max(0, group.indexOf(series));
    const horizontal = groupIsHorizontal(series);
    const categorySize = categoryBandSize(categoryIndex, horizontal);
    const geometry = clusterGeometry(group, categorySize);
    const start = categoryStart(categoryIndex, horizontal) + geometry.catStart;
    if (groupIsStacked(series)) return start + geometry.barW / 2;
    return start + groupIndex * geometry.clusterGap + geometry.barW / 2;
  };
  const continuousBarCategoryCenter = (series: ChartSeries, index: number): number => {
    if (Number.isInteger(index) && index >= 0 && index < n) {
      return barCategoryCenter(series, index);
    }
    const group = barGroupFor(series);
    const groupIndex = Math.max(0, group.indexOf(series));
    const horizontal = groupIsHorizontal(series);
    const gap = categoryGap(horizontal);
    const geometry = clusterGeometry(group, gap);
    const slot = horizontal
      ? (catRev ? index : n - 1 - index)
      : (catRev ? n - 1 - index : index);
    const start = (horizontal ? py0 : px0) + slot * gap + geometry.catStart;
    const visualIndex = groupIsStacked(series) ? 0 : groupIndex;
    return start + visualIndex * geometry.clusterGap + geometry.barW / 2;
  };
  for (let seriesIndex = 0; seriesIndex < barSeries.length; seriesIndex++) {
    const series = barSeries[seriesIndex];
    const seriesIsHorizontal = groupIsHorizontal(series);
    const secondary = sec != null && series.useSecondaryAxis === true;
    const valueAt = seriesIsHorizontal
      ? valX
      : secondary && secScale ? secScale.makeToY(py0, ph) : valY;
    const color = chartColor(seriesIndex, series);
    const plotted = (categoryIndex: number): number =>
      plottedBarValue(seriesIndex, categoryIndex);
    for (const errorBars of effectiveBarErrorBars(series)) {
      drawBarErrorBars(
        ctx, series, chartStyleRoleErrorBar(chart, errorBars), n, seriesIsHorizontal,
        categoryIndex => barCategoryCenter(series, categoryIndex),
        valueAt,
        plotted, color, ptToPx,
      );
    }
    drawSeriesTrendlines(
      ctx, series, color,
      index => continuousBarCategoryCenter(series, index - 1),
      valueAt,
      ptToPx,
      series.values.map((_value, index) => index + 1),
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
      (index, value) => seriesIsHorizontal
        ? ({
          x: valueAt(value),
          y: continuousBarCategoryCenter(series, index - 1),
        })
        : ({
          x: continuousBarCategoryCenter(series, index - 1),
          y: valueAt(value),
        }),
    );
  }

  if ((!hasDataTable || isH) && !chart.catAxisHidden && catLabelsVisible(chart)) {
    // `<c:catAx><c:txPr>…<a:solidFill>` colors the category tick labels.
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    const drawnCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, catGap * 0.5));
    ctx.font = chartFontCss(
      drawnCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Column: each label is centered in a category slot of width `catGap`, so
    // cap it just under that so neighbours don't collide. Horizontal bars: the
    // label sits right-aligned in the left gutter between the val-title/legend
    // band and the plot edge, so cap it at that band width.
    const catSlotMaxPx = catGap - 4;
    const horizLabelMaxPx = (px0 - 4) - (x + legLeftW + valTitleW);
    // `<c:catAx><c:txPr><a:bodyPr rot>` rotates the column labels (0 = flat).
    const rotRad = catLabelRotation;
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        raw: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        fraction: tick.fraction,
        categoryIndex: -1,
      }))
      : cats.map((category, categoryIndex) => ({
        // §21.2.2.71: a category-axis numFmt formats numeric-serial categories
        // (e.g. dateAx serials → real dates). No-op for string categories.
        raw: formatCategoryLabel(category.toString(), chart.catAxisFormatCode, chart.date1904),
        fraction: null,
        categoryIndex,
      }));
    for (const entry of labelEntries) {
      const { raw } = entry;
      if (!isH) {
        const anchor = entry.fraction != null || entry.categoryIndex < 0
          ? { fraction: entry.fraction ?? 0.5, textAlign: 'center' as CanvasTextAlign }
          : categoryLabelAnchorFraction(
            entry.categoryIndex,
            n,
            isCrossBetween(chart),
            catRev,
            chart.catAxisLabelAlignment,
          );
        const lx = px0 + anchor.fraction * pw;
        ctx.textAlign = anchor.textAlign; ctx.textBaseline = 'top';
        // Rotation elides against a longer diagonal budget. Horizontal labels
        // use the measured word-wrap computed from this category slot.
        const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
        const gap = categoryLabelOffsetPx(
          chart.catAxisFontSizeHpt != null
            ? categoryTickLabelGapPx(drawnCatTickFontPx)
            : 3,
          chart.catAxisLabelOffsetPercent,
        );
        if (rotRad === 0) {
          const lines = entry.categoryIndex >= 0
            ? (wrappedColumnCategories[entry.categoryIndex] ?? [raw])
            : [raw];
          lines.forEach((line, lineIndex) => {
            ctx.fillText(line, lx, categoryLabelAxisY + gap + lineIndex * (drawnCatTickFontPx + 2));
          });
        } else {
          drawRotatedCatLabel(ctx, elideToWidth(ctx, raw, budget), lx, categoryLabelAxisY + gap, rotRad);
        }
      } else {
        const ly = entry.fraction != null
          ? py0 + entry.fraction * ph
          : py0 + categorySlotIndex(entry.categoryIndex, true) * catGap + catGap / 2;
        const gap = categoryLabelOffsetPx(
          chart.catAxisFontSizeHpt != null
            ? valueTickLabelGapPx(drawnCatTickFontPx)
            : 4,
          chart.catAxisLabelOffsetPercent,
        );
        const boxStart = x + legLeftW + valTitleW;
        const boxEnd = px0 - gap;
        const alignment = chart.catAxisLabelAlignment;
        const lx = alignment === 'l'
          ? boxStart
          : alignment === 'ctr' ? (boxStart + boxEnd) / 2 : boxEnd;
        // Horizontal bars historically right-align omitted category labels in
        // the left gutter. `lblAlgn` only replaces that default when authored.
        ctx.textAlign = alignment === 'l' ? 'left' : alignment === 'ctr' ? 'center' : 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(elideToWidth(ctx, raw, horizLabelMaxPx), lx, ly);
      }
    }

    if (!isH && categoryLevels) {
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(drawnCatTickFontPx)
          : 3,
        chart.catAxisLabelOffsetPercent,
      );
      const rowHeight = drawnCatTickFontPx + 4;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      ctx.strokeStyle = catLineColor;
      ctx.lineWidth = catLineW;
      ctx.setLineDash([]);
      const boundaryStartY = (boundaryIndex: number): number => {
        if (!multiLevelBoundariesOwnMajorTicks
          || boundaryIndex % catMajorTickSkip !== 0) {
          return categoryLabelAxisY;
        }
        const tickLength = axisTickLengthPx('major', catLineW, ptToPx);
        if (chart.catAxisMajorTickMark === 'cross') {
          return categoryLabelAxisY - tickLength / 2;
        }
        if (chart.catAxisMajorTickMark === 'in') {
          return categoryLabelAxisY - tickLength;
        }
        return categoryLabelAxisY;
      };

      // Each boundary in the innermost category level separates adjacent
      // labels. Boundaries shared by an outer level are extended below by the
      // outer-level brackets drawn next; extend the remaining boundaries
      // through the first label band here.
      const outerBoundaryIndices = new Set<number>();
      for (let levelIndex = 1; levelIndex < categoryLevels.length; levelIndex++) {
        const level = categoryLevels[levelIndex] ?? [];
        for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
          if ((level[categoryIndex] ?? '') !== '') outerBoundaryIndices.add(categoryIndex);
        }
        outerBoundaryIndices.add(n);
      }
      const firstBandBottom = categoryLabelAxisY + gap + drawnCatTickFontPx + 2;
      for (let boundaryIndex = 0; boundaryIndex <= n; boundaryIndex++) {
        if (outerBoundaryIndices.has(boundaryIndex)) continue;
        const boundary = px0 + boundaryIndex / n * pw;
        ctx.beginPath();
        ctx.moveTo(boundary, boundaryStartY(boundaryIndex));
        ctx.lineTo(boundary, firstBandBottom);
        ctx.stroke();
      }

      const outerBoundaryBottoms = new Map<number, number>();
      for (let levelIndex = 1; levelIndex < categoryLevels.length; levelIndex++) {
        const level = categoryLevels[levelIndex] ?? [];
        const starts: number[] = [];
        for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
          if ((level[categoryIndex] ?? '') !== '') starts.push(categoryIndex);
        }
        for (let groupIndex = 0; groupIndex < starts.length; groupIndex++) {
          const start = starts[groupIndex];
          const end = starts[groupIndex + 1] ?? n;
          const label = level[start] ?? '';
          const left = px0 + start / n * pw;
          const right = px0 + end / n * pw;
          const rowTop = categoryLabelAxisY + gap + levelIndex * rowHeight;
          const alignment = chart.catAxisLabelAlignment;
          const labelX = alignment === 'l'
            ? left
            : alignment === 'r' ? right : (left + right) / 2;
          ctx.textAlign = alignment === 'l' ? 'left' : alignment === 'r' ? 'right' : 'center';
          ctx.fillText(
            elideToWidth(ctx, label, Math.max(0, right - left - 4)),
            labelX,
            rowTop,
          );
          const bracketBottom = rowTop + drawnCatTickFontPx + 2;
          for (const boundaryIndex of [start, end]) {
            outerBoundaryBottoms.set(
              boundaryIndex,
              Math.max(outerBoundaryBottoms.get(boundaryIndex) ?? categoryLabelAxisY, bracketBottom),
            );
          }
        }
      }
      // A boundary can belong to both neighbouring groups, and at deeper
      // levels the same category index can recur. Paint each authored
      // boundary once to its deepest required extent so coincident strokes do
      // not become darker or wider.
      for (const [boundaryIndex, bracketBottom] of outerBoundaryBottoms) {
        const boundary = px0 + boundaryIndex / n * pw;
        ctx.beginPath();
        ctx.moveTo(boundary, boundaryStartY(boundaryIndex));
        ctx.lineTo(boundary, bracketBottom);
        ctx.stroke();
      }
    }
  }

  if (lineSeries.length > 0 && !isH) {
    drawLineGroupDecorations(
      ctx, chart, n, categoryCenterX,
      series => sec && series.useSecondaryAxis === true ? toYSecondarySeries : toYPrimaryLine,
      () => primaryCatAxisY,
      (series, index) => series.values[index] == null ? null : overlayValue(series, index),
      catGap, ptToPx, shapeRotationDeg, 'background',
    );
    if (dateAxisPlan) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(px0, py0, pw, ph);
      ctx.clip();
    }
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      const color = chartColor(barSeries.length + si, s);
      // Series bound to the secondary axis map through its scale; others use
      // the primary (bar) value axis.
      const yOf = sec && s.useSecondaryAxis === true
        ? toYSecondarySeries
        : toYPrimaryLine;
      const hasAuthoredLine = s.chartexStyle != null
        || s.lineHidden != null
        || s.lineColor != null
        || s.lineWidthEmu != null
        || chart.chartexDataPointLineStyle != null;
      const paintLine = hasAuthoredLine
        ? applyChartExSeriesLineStyle(
          ctx,
          chart,
          chart.chartexDataPointLineStyle,
          s,
          chartExSeriesFormatIndex(s, sourceSeriesIndices.get(s) ?? si),
          lineSeries.length,
          color,
          ptToPx,
          { linkedNoStyleFallback: options.semanticLineNoStyleFallback },
        )
        : true;
      if (!hasAuthoredLine) {
        ctx.strokeStyle = color;
        ctx.lineWidth = 2;
        ctx.setLineDash([]);
      }
      ctx.beginPath();
      const smooth = s.smooth === true;
      const dispBlanks = chart.dispBlanksAs ?? 'gap';
      let run: Array<{ x: number; y: number }> = [];
      const flushRun = (): void => {
        if (run.length === 0) return;
        ctx.moveTo(run[0].x, run[0].y);
        appendCurve(ctx, run, smooth);
        run = [];
      };
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) {
          if (dispBlanks === 'gap') flushRun();
          if (dispBlanks !== 'zero') continue;
        }
        const lx = categoryCenterX(ci);
        run.push({ x: lx, y: yOf(overlayValue(s, ci)) });
      }
      flushRun();
      if (paintLine) ctx.stroke();
      const seriesMarkersVisible = s.showMarker !== false && s.markerSymbol !== 'none';
      const drawMarkers = seriesMarkersVisible || hasVisiblePointMarkerOverride(s);
      const hasMarkerDetail = seriesHasMarkerDetail(s);
      if (drawMarkers) {
        for (let ci = 0; ci < n; ci++) {
          const v = s.values[ci];
          if (v == null) continue;
          const lx = categoryCenterX(ci);
          const ly = yOf(overlayValue(s, ci));
          const point = pointOverrides.get(ci);
          const symbol = effectiveMarkerSymbol(s, point, 'circle', seriesMarkersVisible);
          if (symbol === 'none') continue;
          if (hasMarkerDetail || pointHasMarkerDetail(point)) {
            const lineWidthEmu = point?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawMarker(
              ctx, lx, ly, symbol,
              point?.markerSize ?? s.markerSize ?? 5,
              markerFillColorFor(s, point, ci, color),
              point?.markerLine ?? s.markerLine ?? null,
              ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, point, ci),
              shapeRotationDeg,
            );
          } else {
            ctx.fillStyle = color;
            ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
          }
        }
      }
      // Trendlines (`<c:trendline>`, §21.2.2.211) for the combo line series.
      drawSeriesTrendlines(
        ctx, s, color,
        (i) => categoryCenterX(i),
        yOf, ptToPx, undefined,
        {
          chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
          shapeRotationDeg,
        },
      );
    }
    drawLineGroupDecorations(
      ctx, chart, n, categoryCenterX,
      series => sec && series.useSecondaryAxis === true ? toYSecondarySeries : toYPrimaryLine,
      () => primaryCatAxisY,
      (series, index) => series.values[index] == null ? null : overlayValue(series, index),
      catGap, ptToPx, shapeRotationDeg, 'foreground',
    );
    if (dateAxisPlan) ctx.restore();
  }

  // A scatter group can be overlaid on a bar chart with its own pair of
  // numeric axes (ECMA-376 CT_ScatterChart `axId`, first X then Y). This is the
  // standard construction for dot/range plots: an invisible horizontal bar
  // series supplies category labels and the visible scatter markers plus
  // custom X error bars supply the dots and connecting ranges.
  if (scatterSeries.length > 0) {
    const allX: number[] = [];
    const allY: number[] = [];
    for (const s of scatterSeries) {
      const sx = s.categories ?? [];
      for (let i = 0; i < s.values.length; i++) {
        const xv = scatterXValue(sx, i, false);
        const yv = s.values[i];
        if (xv == null || yv == null) continue;
        allX.push(xv);
        allY.push(yv);
      }
    }
    if (allX.length && allY.length) {
      const xAxis = chart.secondaryCatAxis;
      const yAxis = chart.secondaryValAxis;
      const xExtent = finiteDataExtent(allX);
      const yExtent = finiteDataExtent(allY);
      const needsMinor = (axis: SecondaryValueAxis | null | undefined): boolean =>
        axis?.minorGridlines === true
        || (axis?.minorTickMark != null && axis.minorTickMark !== 'none');
      const xScale = planNumericValueAxis({
        dataMin: xExtent.min,
        dataMax: xExtent.max,
        explicitMin: xAxis?.min,
        explicitMax: xAxis?.max,
        axisLenPt: pw / ptToPx,
        axisOrientation: 'horizontal',
        majorUnit: xAxis?.majorUnit,
        minorUnit: xAxis?.minorUnit,
        needMinor: needsMinor(xAxis),
        logBase: xAxis?.logBase,
        reversed: xAxis?.orientation === 'maxMin',
      });
      const yScale = planNumericValueAxis({
        dataMin: yExtent.min,
        dataMax: yExtent.max,
        explicitMin: yAxis?.min,
        explicitMax: yAxis?.max,
        axisLenPt: ph / ptToPx,
        axisOrientation: 'vertical',
        majorUnit: yAxis?.majorUnit,
        minorUnit: yAxis?.minorUnit,
        needMinor: needsMinor(yAxis),
        logBase: yAxis?.logBase,
        reversed: yAxis?.orientation === 'maxMin',
      });
      const scatterToX = (value: number): number =>
        px0 + xScale.fraction(value) * pw;
      const scatterToY = (value: number): number =>
        py0 + ph - yScale.fraction(value) * ph;
      drawScatterSeriesLayer(
        ctx,
        chart,
        scatterSeries.map((series, index) => ({
          series,
          index: sourceSeriesIndices.get(series) ?? index,
        })),
        false,
        scatterToX,
        scatterToY,
        r,
        px0,
        py0,
        pw,
        ph,
        ptToPx,
        false,
        chart.scatterStyle ?? 'marker',
        { x, y, w, h },
        yScale.max,
        undefined,
        shapeRotationDeg,
      );
    }
  }

  // Primary axis rules + ticks on top of the bars/line so the category
  // baseline stays visible (the bars would otherwise paint over it).
  drawAxesOnTop();

  if (secondaryCat && !isH) {
    drawSecondaryCategoryAxis(
      ctx, chart, secondaryCat, secondaryCategories, r, px0, py0, pw, ptToPx,
    );
  }

  // Secondary value axis (right edge). Independent scale: its own "nice" major
  // unit drives the tick labels, positioned via `toYSecondary` (NOT aligned to
  // the primary gridlines — PowerPoint places them independently). Draws its
  // rule + ticks on the right; ticks mirror the left axis ("out" points right).
  if (sec && secScale) {
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r, px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, valLabelColor, chart.date1904, secondaryPercentAxis,
    );
  }

  if (dataTableLayout) {
    const tableY = py0 + ph + (isH
      ? (chart.valAxisHidden ? h * 0.02 : catAxisLabelBandH(valAxLabelFontPx))
      : 0);
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, tableY, pw, x + legLeftW, ptToPx,
    );
  }

  const legendPaints = isChartExColumn
    ? barSeries.flatMap((series, index) => [
      chartExDataPointPaint(
        chart, barStyleIndices[index], barSeries.length, series.chartexStyle, series.color,
      ),
      ...trendlineLegendSeries([series]).map(() => undefined),
    ])
    : [];
  drawLegendForLayout(
    ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph,
    titleH + 2, ptToPx, legendPaints,
  );
  drawAxisTitles(
    ctx, chart, x, y, w, h, px0, py0, pw, ph,
    legLeftW, legBottomH, catTitlePx, valTitlePx, isH,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function applyDecorationLineStyle(
  ctx: CanvasRenderingContext2D,
  style: ChartDecorationLineStyle,
  ptToPx: number,
): boolean {
  if (style.hidden === true || (style.paintAuthored === true && style.color == null)) return false;
  ctx.strokeStyle = `#${style.color ?? '000000'}`;
  ctx.lineWidth = style.widthEmu != null
    ? axisLineWidthPx(style.widthEmu, ptToPx)
    : Math.max(1, 0.75 * ptToPx);
  ctx.setLineDash(dashPatternForPreset(style.dash ?? undefined, ctx.lineWidth));
  ctx.lineCap = style.cap === 'rnd' ? 'round' : style.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = style.join === 'round' || style.join === 'bevel' ? style.join : 'miter';
  return true;
}

function chartStyleRoleLine(
  chart: ChartModel,
  direct: ChartDecorationLineStyle,
  role: ChartStyleRole,
): ChartDecorationLineStyle {
  const linked = chart.chartStyleRoles?.[role];
  const linkedApplies = linked != null && linked.lineNoStyle !== true;
  const directPaintAuthored = direct.paintAuthored === true
    || direct.color != null || direct.hidden === true;
  const linkedPaintAuthored = linkedApplies && (linked.linePaintAuthored === true
    || linked.lineHidden === true
    || linked.lineColors?.some(color => color != null) === true
    || linked.linePaints?.some(paint => paint != null) === true);
  return {
    color: direct.color
      ?? (!directPaintAuthored && linkedApplies
        ? chartExStyleColor(chart, linked, 'line', 0, 1) : null),
    paintAuthored: directPaintAuthored
      ? direct.paintAuthored
      : linkedPaintAuthored ? true : undefined,
    widthEmu: direct.widthEmu ?? (linkedApplies ? linked.lineWidthEmu : null),
    dash: direct.dash ?? (linkedApplies ? linked.lineDash : null),
    cap: direct.cap ?? (linkedApplies ? linked.lineCap : null),
    join: direct.join ?? (linkedApplies ? linked.lineJoin : null),
    hidden: direct.hidden
      ?? (!directPaintAuthored && linkedApplies && linked.lineHidden === true ? true : null),
  };
}

function chartStyleRoleBarPaint(
  chart: ChartModel,
  direct: ChartStockUpDownBarStyle['up'],
  role: 'upBar' | 'downBar',
): ChartStockUpDownBarStyle['up'] {
  const linked = chart.chartStyleRoles?.[role];
  const fillApplies = linked != null && linked.fillNoStyle !== true;
  const lineApplies = linked != null && linked.lineNoStyle !== true;
  const linkedFill = linked?.fillPaints?.[0];
  const directFillAuthored = direct.fillPaintAuthored === true
    || direct.fillColor != null || direct.fill != null || direct.fillHidden === true;
  const directLineAuthored = direct.linePaintAuthored === true
    || direct.lineColor != null || direct.lineHidden === true;
  const linkedFillAuthored = fillApplies && (linked.fillPaintAuthored === true
    || linked.fillHidden === true || linkedFill != null
    || linked.fillColors?.some(color => color != null) === true);
  const linkedLineAuthored = lineApplies && (linked.linePaintAuthored === true
    || linked.lineHidden === true
    || linked.lineColors?.some(color => color != null) === true
    || linked.linePaints?.some(paint => paint != null) === true);
  return {
    fillColor: direct.fillColor
      ?? (!directFillAuthored && fillApplies
        ? chartExStyleColor(chart, linked, 'fill', 0, 1) : null),
    fill: direct.fill ?? (
      !directFillAuthored && fillApplies
        && linkedFill != null
        && linkedFill.fillType !== 'image'
        && linkedFill.fillType !== 'none' ? linkedFill : null
    ),
    fillPaintAuthored: directFillAuthored
      ? direct.fillPaintAuthored
      : linkedFillAuthored ? true : undefined,
    fillHidden: direct.fillHidden
      ?? (!directFillAuthored && fillApplies && linked.fillHidden === true ? true : null),
    lineColor: direct.lineColor
      ?? (!directLineAuthored && lineApplies
        ? chartExStyleColor(chart, linked, 'line', 0, 1) : null),
    linePaintAuthored: directLineAuthored
      ? direct.linePaintAuthored
      : linkedLineAuthored ? true : undefined,
    lineWidthEmu: direct.lineWidthEmu ?? (lineApplies ? linked.lineWidthEmu : null),
    lineDash: direct.lineDash ?? (lineApplies ? linked.lineDash : null),
    lineCap: direct.lineCap ?? (lineApplies ? linked.lineCap : null),
    lineJoin: direct.lineJoin ?? (lineApplies ? linked.lineJoin : null),
    lineHidden: direct.lineHidden
      ?? (!directLineAuthored && lineApplies && linked.lineHidden === true ? true : null),
  };
}

/** Draws the one-per-category drop-line envelope shared by classic line,
 * area, and stock charts. ECMA-376 assigns the geometry to the owning chart
 * group; each envelope joins its effective category-axis crossing to every
 * finite plotted point at that category. */
function drawDropLineEnvelopes(
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

function chartStyleRoleErrorBar(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['errBars']>[number],
): NonNullable<ChartSeries['errBars']>[number] {
  const linked = chartStyleRoleLine(chart, {
    color: direct.color,
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
  };
}

function chartStyleRoleLeaderLine(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    color: direct.leaderLineColor,
    widthEmu: direct.leaderLineWidthEmu,
    dash: direct.leaderLineDash,
    hidden: direct.leaderLineHidden,
  }, 'leaderLine');
}

function chartStyleRoleTrendline(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['trendLines']>[number],
): NonNullable<ChartSeries['trendLines']>[number] {
  const linked = chartStyleRoleLine(chart, {
    color: direct.lineColor,
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
  };
}

function chartStyleRoleDataTable(
  chart: ChartModel,
  direct: NonNullable<ChartModel['dataTable']>,
): NonNullable<ChartModel['dataTable']> {
  const linked = chartStyleRoleLine(chart, {
    color: direct.lineColor,
    widthEmu: direct.lineWidthEmu,
    dash: direct.lineDash,
    hidden: direct.lineHidden,
  }, 'dataTable');
  return {
    ...direct,
    lineColor: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    lineDash: linked.dash ?? undefined,
    lineHidden: linked.hidden ?? undefined,
  };
}

interface LinkedGridlineResult {
  visible: boolean | null | undefined;
  color?: string | null;
  widthEmu?: number | null;
  dash?: string | null;
}

function chartStyleRoleGridline(
  chart: ChartModel,
  role: 'gridlineMajor' | 'gridlineMinor',
  visible: boolean | null | undefined,
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
): LinkedGridlineResult {
  if (visible !== true || !chart.chartStyleRoles?.[role]) {
    return { visible, color, widthEmu, dash };
  }
  const linked = chartStyleRoleLine(chart, { color, widthEmu, dash }, role);
  return {
    visible: linked.hidden !== true,
    color: linked.color,
    widthEmu: linked.widthEmu,
    dash: linked.dash,
  };
}

function chartStyleRoleSecondaryGridlines(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
): SecondaryValueAxis | null | undefined {
  if (!axis) return axis;
  if (!chart.chartStyleRoles?.gridlineMajor && !chart.chartStyleRoles?.gridlineMinor) return axis;
  const major = chartStyleRoleGridline(
    chart, 'gridlineMajor', axis.majorGridlines,
    axis.majorGridlineColor, axis.majorGridlineWidthEmu, axis.majorGridlineDash,
  );
  const minor = chartStyleRoleGridline(
    chart, 'gridlineMinor', axis.minorGridlines,
    axis.minorGridlineColor, axis.minorGridlineWidthEmu, axis.minorGridlineDash,
  );
  const changed = major.visible !== axis.majorGridlines
    || major.color !== axis.majorGridlineColor
    || major.widthEmu !== axis.majorGridlineWidthEmu
    || major.dash !== axis.majorGridlineDash
    || minor.visible !== axis.minorGridlines
    || minor.color !== axis.minorGridlineColor
    || minor.widthEmu !== axis.minorGridlineWidthEmu
    || minor.dash !== axis.minorGridlineDash;
  return changed ? {
    ...axis,
    majorGridlines: major.visible ?? undefined,
    majorGridlineColor: major.color,
    majorGridlineWidthEmu: major.widthEmu,
    majorGridlineDash: major.dash,
    minorGridlines: minor.visible ?? undefined,
    minorGridlineColor: minor.color,
    minorGridlineWidthEmu: minor.widthEmu,
    minorGridlineDash: minor.dash,
  } : axis;
}

function chartStyleRoleAxisLine(
  chart: ChartModel,
  role: 'categoryAxis' | 'valueAxis',
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  hidden: boolean,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    color,
    widthEmu,
    dash,
    // The shared axis model stores the effective boolean, so false means no
    // direct noFill rather than an authored visible override.
    hidden: hidden ? true : undefined,
  }, role);
}

function chartStyleRoleSecondaryAxisLine(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
  role: 'categoryAxis' | 'valueAxis',
): SecondaryValueAxis | null | undefined {
  const style = chart.chartStyleRoles?.[role];
  if (!axis || !style) return axis;
  const line = chartStyleRoleAxisLine(
    chart, role, axis.lineColor, axis.lineWidthEmu, axis.lineDash, axis.lineHidden,
  );
  const lineHidden = line.hidden === true;
  const fontSizeHpt = axis.fontSizeHpt ?? style.fontSizeHpt;
  const fontBold = axis.fontBold ?? style.fontBold;
  const fontItalic = axis.fontItalic ?? style.fontItalic;
  const fontColor = axis.fontColor ?? style.fontColor;
  const fontFace = axis.fontFace ?? style.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && fontFace === axis.fontFace) return axis;
  return {
    ...axis,
    lineColor: line.color,
    lineWidthEmu: line.widthEmu,
    lineDash: line.dash,
    lineHidden,
    fontSizeHpt,
    fontBold,
    fontItalic,
    fontColor,
    fontFace,
  };
}

function chartStyleRoleSeriesAxis(chart: ChartModel): ChartModel {
  const axis = chart.threeD?.seriesAxis;
  const style = chart.chartStyleRoles?.seriesAxis;
  if (!axis || !style) return chart;
  const line = chartStyleRoleLine(chart, {
    color: axis.lineColor,
    widthEmu: axis.lineWidthEmu,
    dash: axis.lineDash,
    hidden: axis.lineHidden ? true : undefined,
  }, 'seriesAxis');
  const lineHidden = line.hidden === true;
  const fontSizeHpt = axis.fontSizeHpt ?? style.fontSizeHpt;
  const fontBold = axis.fontBold ?? style.fontBold;
  const fontItalic = axis.fontItalic ?? style.fontItalic;
  const fontColor = axis.fontColor ?? style.fontColor;
  const fontFace = axis.fontFace ?? style.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && fontFace === axis.fontFace) return chart;
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
        fontFace,
      },
    },
  };
}

function isClassicMarkerSeries(chart: ChartModel, series: ChartSeries): boolean {
  if (chart.chartType === 'bubble') return false;
  const family = series.seriesType ?? chart.chartType;
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

function chartStyleRoleMarker(
  chart: ChartModel,
  direct: ChartSeries,
  index: number,
  count: number,
): ChartSeries {
  const linked = chart.chartStyleRoles?.dataPointMarker;
  if (!isClassicMarkerSeries(chart, direct)
    || ((direct.showMarker === false || direct.markerSymbol === 'none')
      && !hasVisiblePointMarkerOverride(direct))) return direct;
  const fillApplies = linked != null && linked.fillNoStyle !== true;
  const lineApplies = linked != null && linked.lineNoStyle !== true;
  const directFillAuthored = direct.markerFillPaintAuthored === true
    || direct.markerFill != null || direct.markerFillPaint !== undefined;
  const markerFill = direct.markerFill
    ?? (!directFillAuthored && fillApplies && linked?.fillHidden === true
      ? '00000000'
      : !directFillAuthored && fillApplies && linked
        ? chartExStyleColor(chart, linked, 'fill', index, count) : null);
  const linkedFillPaint = !directFillAuthored && fillApplies && linked
    ? chartExStylePaintDecision(chart, linked, index, count)
    : undefined;
  const markerFillPaint = direct.markerFillPaint !== undefined
    ? direct.markerFillPaint
    : linkedFillPaint?.fillType === 'gradient'
        || linkedFillPaint?.fillType === 'pattern'
        || linkedFillPaint?.fillType === 'image'
      ? linkedFillPaint
      : undefined;
  const markerFillPaintAuthored = directFillAuthored
    ? direct.markerFillPaintAuthored
    : fillApplies && linked?.fillPaintAuthored === true
      ? true
      : undefined;
  const markerLine = direct.markerLine
    ?? (lineApplies && linked?.lineHidden === true
      ? '00000000'
      : lineApplies && linked ? chartExStyleColor(chart, linked, 'line', index, count) : null);
  const markerLineWidthEmu = direct.markerLineWidthEmu
    ?? (lineApplies ? linked?.lineWidthEmu : null);
  const markerSize = direct.markerSize ?? chart.chartStyleMarkerSizePt;
  const markerSymbol = direct.markerSymbol ?? chart.chartStyleMarkerSymbol;
  if (markerFill === direct.markerFill
    && markerFillPaint === direct.markerFillPaint
    && markerFillPaintAuthored === direct.markerFillPaintAuthored
    && markerLine === direct.markerLine
    && markerLineWidthEmu === direct.markerLineWidthEmu
    && markerSize === direct.markerSize
    && markerSymbol === direct.markerSymbol) return direct;
  return {
    ...direct,
    markerFill,
    markerFillPaint,
    markerFillPaintAuthored,
    markerLine,
    markerLineWidthEmu,
    markerSize,
    markerSymbol,
  };
}

interface EffectiveFrameLineStyle {
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
function effectiveFrameLineStyle(
  chart: ChartModel,
  direct: EffectiveFrameLineStyle,
  linked: ChartExStyle | null | undefined,
): EffectiveFrameLineStyle {
  if (!linked || linked.lineNoStyle === true) return direct;
  let { color, fill, hidden } = direct;
  const directPaint = direct.paintAuthored === true
    || fill != null || color != null || hidden === true;
  if (!directPaint) {
    if (linked.lineHidden === true) {
      hidden = true;
    } else {
      fill = chartExStyleLinePaint(linked, 0);
      color = fill == null ? chartExStyleColor(chart, linked, 'line', 0, 1) : null;
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
    paintAuthored: direct.paintAuthored,
    widthEmu: direct.widthEmu ?? linked.lineWidthEmu,
    dash,
    dashAuthored,
    customDash,
    cap: direct.cap ?? linked.lineCap,
    join: direct.join ?? linked.lineJoin,
    compound: direct.compound ?? linked.lineCompound,
  };
}

function effectiveLinkedLabelBox(
  chart: ChartModel,
  direct: ChartLabelBox | null | undefined,
  linked: ChartExStyle | null | undefined,
  createFromLinked: boolean,
): ChartLabelBox | undefined {
  if (!linked || (!direct && !createFromLinked)) return direct ?? undefined;
  const source = direct ?? {};
  let fill = source.fill;
  let fillPaint = source.fillPaint;
  let fillHidden = source.fillHidden;
  const directFill = source.fillPaintAuthored === true
      || fill != null || fillPaint != null || fillHidden === true;
  if (!directFill && linked.fillNoStyle !== true) {
    fillHidden = linked.fillHidden;
    fillPaint = linked.fillHidden === true
      ? undefined
      : chartExStyleFillPaint(linked, 0) as ChartLabelBox['fillPaint'];
    fill = fillPaint == null && linked.fillHidden !== true
      ? chartExStyleColor(chart, linked, 'fill', 0, 1) ?? undefined
      : undefined;
  }
  const line = effectiveFrameLineStyle(chart, {
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
  }, linked);
  return {
    ...source,
    fill,
    fillPaint,
    fillHidden,
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
  };
}

/** Merge two directly-authored label shapes property-by-property. The higher
 * precedence shape owns an authored paint/noFill choice even when that choice
 * cannot be resolved to a Canvas paint; omitted geometry continues to inherit
 * from the lower-precedence series/linked shape. */

function chartStyleRoleDataLabels(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
): ChartSeriesDataLabels {
  const linked = direct.labelBox
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  if (!linked) return direct;
  const labelBox = effectiveLinkedLabelBox(chart, direct.labelBox, linked, false);
  const directFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? linked.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? linked.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? linked.fontItalic ?? undefined,
    fontColor: directFontPaint ? direct.fontColor : linked.fontColor ?? undefined,
    fontPaintAuthored: directFontPaint || linked.fontPaintAuthored === true || undefined,
    fontHidden: directFontPaint ? direct.fontHidden : linked.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? linked.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? linked.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? linked.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? linked.textRotation ?? undefined,
    textWrap: direct.textWrap ?? linked.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? linked.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? linked.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? linked.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? linked.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? linked.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? linked.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || linked.textBodyAuthored === true || undefined,
    labelBox,
  };
}

function chartStyleRoleTrendlineLabel(
  chart: ChartModel,
  direct: ChartTrendline,
): ChartTrendline {
  const linked = chart.chartStyleRoles?.trendlineLabel;
  if (!linked) return direct;
  const directFontPaint = direct.labelFontPaintAuthored === true
    || direct.labelFontColor != null || direct.labelFontHidden === true;
  return {
    ...direct,
    // Unlike `dataLabelCallout`, the `trendlineLabel` role styles the generated
    // equation/R² label shape even when the chart does not carry a local spPr.
    // Materialize it before the chart-wide paint preflight so linked gradient
    // work is charged before any family starts painting.
    labelBox: effectiveLinkedLabelBox(chart, direct.labelBox, linked, true),
    labelFontSizeHpt: direct.labelFontSizeHpt ?? linked.fontSizeHpt ?? undefined,
    labelFontBold: direct.labelFontBold ?? linked.fontBold ?? undefined,
    labelFontItalic: direct.labelFontItalic ?? linked.fontItalic ?? undefined,
    labelFontColor: directFontPaint ? direct.labelFontColor : linked.fontColor ?? undefined,
    labelFontPaintAuthored: directFontPaint || linked.fontPaintAuthored === true || undefined,
    labelFontHidden: directFontPaint ? direct.labelFontHidden : linked.fontHidden ?? undefined,
    labelFontFace: direct.labelFontFace ?? linked.fontFace ?? undefined,
    labelFontLanguage: direct.labelFontLanguage ?? linked.fontLanguage ?? undefined,
    labelFontBaseline: direct.labelFontBaseline ?? linked.fontBaseline ?? undefined,
    labelTextRotation: direct.labelTextRotation ?? linked.textRotation ?? undefined,
    labelTextWrap: direct.labelTextWrap ?? linked.textWrap ?? undefined,
    labelTextVerticalAnchor: direct.labelTextVerticalAnchor
      ?? linked.textVerticalAnchor ?? undefined,
    labelTextVerticalMode: direct.labelTextVerticalMode ?? linked.textVerticalMode ?? undefined,
    labelTextLInsEmu: direct.labelTextLInsEmu ?? linked.textLInsEmu ?? undefined,
    labelTextTInsEmu: direct.labelTextTInsEmu ?? linked.textTInsEmu ?? undefined,
    labelTextRInsEmu: direct.labelTextRInsEmu ?? linked.textRInsEmu ?? undefined,
    labelTextBInsEmu: direct.labelTextBInsEmu ?? linked.textBInsEmu ?? undefined,
    labelTextBodyAuthored: direct.labelTextBodyAuthored === true
      || linked.textBodyAuthored === true || undefined,
  };
}

function chartStyleRoleDataLabelOverride(
  chart: ChartModel,
  direct: ChartDataLabelOverride,
  seriesDirect: ChartSeriesDataLabels | null | undefined,
): ChartDataLabelOverride {
  // `dataLabelCallout` is the style for a label that authors shape properties;
  // an indexed `<dLbl>`/`<cx:dataLabel>` alone is still an ordinary data label.
  // Applying the callout recipe merely because an indexed override exists
  // invents a white box around every ordinary point label in Office styles.
  const hasCalloutShape = direct.labelBox != null || seriesDirect?.labelBox != null;
  const linked = hasCalloutShape
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  const seriesAndLinkedBox = linked
    ? effectiveLinkedLabelBox(chart, seriesDirect?.labelBox, linked, true)
    : seriesDirect?.labelBox;
  const pointFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  const seriesFontPaint = seriesDirect?.fontPaintAuthored === true
    || seriesDirect?.fontColor != null || seriesDirect?.fontHidden === true;
  const fontSource = pointFontPaint ? direct : seriesFontPaint ? seriesDirect : linked;
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? seriesDirect?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? seriesDirect?.fontBold ?? linked?.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? seriesDirect?.fontItalic ?? linked?.fontItalic ?? undefined,
    fontColor: fontSource?.fontColor ?? undefined,
    fontPaintAuthored: pointFontPaint || seriesFontPaint
      || linked?.fontPaintAuthored === true || undefined,
    fontHidden: fontSource?.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? seriesDirect?.fontFace ?? linked?.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? seriesDirect?.fontLanguage
      ?? linked?.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? seriesDirect?.fontBaseline
      ?? linked?.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? seriesDirect?.textRotation
      ?? linked?.textRotation ?? undefined,
    textWrap: direct.textWrap ?? seriesDirect?.textWrap ?? linked?.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? seriesDirect?.textVerticalAnchor
      ?? linked?.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? seriesDirect?.textVerticalMode
      ?? linked?.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? seriesDirect?.textLInsEmu
      ?? linked?.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? seriesDirect?.textTInsEmu
      ?? linked?.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? seriesDirect?.textRInsEmu
      ?? linked?.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? seriesDirect?.textBInsEmu
      ?? linked?.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || seriesDirect?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
    textAlign: direct.textAlign ?? seriesDirect?.textAlign,
    labelBox: mergeChartLabelBoxes(direct.labelBox, seriesAndLinkedBox),
  };
}

function chartStyleRoleLegend(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.legend;
  if (!linked) return chart;
  let legendFill = chart.legendFill;
  let legendFillColor = chart.legendFillColor;
  let legendFillHidden = chart.legendFillHidden;
  const directFillPaint = chart.legendFillPaintAuthored === true
    || legendFill != null || legendFillColor != null || legendFillHidden === true;
  if (!directFillPaint && linked.fillNoStyle !== true) {
    if (linked.fillHidden === true) {
      legendFillHidden = true;
    } else {
      legendFill = chartExStyleFillPaint(linked, 0);
      legendFillColor = legendFill == null
        ? chartExStyleColor(chart, linked, 'fill', 0, 1)
        : null;
    }
  }

  const legendLine = effectiveFrameLineStyle(chart, {
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
  }, linked);
  if (legendFill === chart.legendFill
    && legendFillColor === chart.legendFillColor
    && legendFillHidden === chart.legendFillHidden
    && legendLine.color === chart.legendLineColor
    && legendLine.fill === chart.legendLineFill
    && legendLine.widthEmu === chart.legendLineWidthEmu
    && legendLine.dash === chart.legendLineDash
    && legendLine.dashAuthored === chart.legendLineDashAuthored
    && legendLine.customDash === chart.legendLineCustomDash
    && legendLine.cap === chart.legendLineCap
    && legendLine.join === chart.legendLineJoin
    && legendLine.compound === chart.legendLineCompound
    && legendLine.hidden === chart.legendLineHidden) return chart;
  return {
    ...chart,
    legendFill,
    legendFillColor,
    legendFillHidden,
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
  };
}

function chartStyleRolePlotArea(chart: ChartModel): ChartModel {
  // MS-ODRAWXML defines plotArea and plotArea3D as separate required style
  // entries. Do not infer one from the other in a malformed/partial sidecar;
  // direct chart formatting stays authoritative below.
  const linked = chart.threeD
    ? chart.chartStyleRoles?.plotArea3D
    : chart.chartStyleRoles?.plotArea;
  if (!linked) return chart;
  let plotAreaFill = chart.plotAreaFill;
  let plotAreaBg = chart.plotAreaBg;
  let plotAreaFillHidden = chart.plotAreaFillHidden;
  const directPaint = chart.plotAreaFillPaintAuthored === true
    || ((plotAreaFill != null || plotAreaBg != null) && chart.plotAreaFillAutomatic !== true)
    || plotAreaFillHidden === true;
  if (!directPaint && linked.fillNoStyle !== true) {
    if (linked.fillHidden === true) {
      plotAreaFillHidden = true;
    } else {
      plotAreaFill = chartExStyleFillPaint(linked, 0);
      plotAreaBg = plotAreaFill == null
        ? chartExStyleColor(chart, linked, 'fill', 0, 1)
        : null;
    }
  }

  const plotAreaLine = effectiveFrameLineStyle(chart, {
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
  }, linked);
  if (plotAreaFill === chart.plotAreaFill
    && plotAreaBg === chart.plotAreaBg
    && plotAreaFillHidden === chart.plotAreaFillHidden
    && plotAreaLine.color === chart.plotAreaLineColor
    && plotAreaLine.fill === chart.plotAreaLineFill
    && plotAreaLine.widthEmu === chart.plotAreaLineWidthEmu
    && plotAreaLine.dash === chart.plotAreaLineDash
    && plotAreaLine.dashAuthored === chart.plotAreaLineDashAuthored
    && plotAreaLine.customDash === chart.plotAreaLineCustomDash
    && plotAreaLine.cap === chart.plotAreaLineCap
    && plotAreaLine.join === chart.plotAreaLineJoin
    && plotAreaLine.compound === chart.plotAreaLineCompound
    && plotAreaLine.hidden === chart.plotAreaLineHidden) return chart;
  return {
    ...chart,
    plotAreaFill,
    plotAreaBg,
    plotAreaFillHidden,
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
  };
}

function chartStyleRoleChartArea(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.chartArea;
  if (!linked) return chart;
  let chartFill = chart.chartFill;
  let chartBg = chart.chartBg;
  let chartFillHidden = chart.chartFillHidden;
  const directPaint = chart.chartFillPaintAuthored === true
    || chartFill != null || chartFillHidden === true;
  if (!directPaint && linked.fillNoStyle !== true) {
    if (linked.fillHidden === true) {
      chartFill = null;
      chartBg = null;
      chartFillHidden = true;
    } else {
      chartFill = chartExStyleFillPaint(linked, 0);
      chartBg = chartFill == null
        ? chartExStyleColor(chart, linked, 'fill', 0, 1)
        : null;
      chartFillHidden = null;
    }
  }

  const chartBorder = effectiveFrameLineStyle(chart, {
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
  }, linked);
  if (chartFill === chart.chartFill
    && chartBg === chart.chartBg
    && chartFillHidden === chart.chartFillHidden
    && chartBorder.color === chart.chartBorderColor
    && chartBorder.fill === chart.chartBorderLineFill
    && chartBorder.widthEmu === chart.chartBorderWidthEmu
    && chartBorder.dash === chart.chartBorderDash
    && chartBorder.dashAuthored === chart.chartBorderDashAuthored
    && chartBorder.customDash === chart.chartBorderCustomDash
    && chartBorder.cap === chart.chartBorderCap
    && chartBorder.join === chart.chartBorderJoin
    && chartBorder.compound === chart.chartBorderCompound
    && chartBorder.hidden === chart.chartBorderHidden) return chart;
  return {
    ...chart,
    chartFill,
    chartBg,
    chartFillHidden,
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
  };
}

/** Materialize the linked decoration roles that an optional family renderer
 * consumes directly from `ChartSeries`. Keeping this projection in core means
 * the 2-D, 3-D, DOCX, XLSX, and PPTX paths receive one effective precedence
 * result without teaching an optional renderer about package sidecars. */
function applyLinkedChartStyleRoles(chart: ChartModel): ChartModel {
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
    && chart.chartStyleMarkerSizePt == null
    && chart.chartStyleMarkerSymbol == null) {
    return chart;
  }
  let changed = false;
  const series = chart.series.map((sourceItem, seriesIndex) => {
    const item = chartStyleRoleMarker(chart, sourceItem, seriesIndex, chart.series.length);
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
      && (chart.chartStyleRoles?.dataLabel || chart.chartStyleRoles?.dataLabelCallout)) {
      const effective = chartStyleRoleDataLabels(chart, seriesDataLabels);
      changed ||= effective !== seriesDataLabels;
      seriesDataLabels = effective;
    }
    const dataLabelOverrides = (chart.chartStyleRoles?.dataLabelCallout
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
      };
      changed ||= merged.leaderLineColor !== seriesDataLabels.leaderLineColor
        || merged.leaderLineWidthEmu !== seriesDataLabels.leaderLineWidthEmu
        || merged.leaderLineDash !== seriesDataLabels.leaderLineDash
        || merged.leaderLineHidden !== seriesDataLabels.leaderLineHidden;
      seriesDataLabels = merged;
    }
    const trendLines = (chart.chartStyleRoles?.trendline || chart.chartStyleRoles?.trendlineLabel)
      ? item.trendLines?.map(trendline => {
      let effective = chart.chartStyleRoles?.trendline
        ? chartStyleRoleTrendline(chart, trendline)
        : trendline;
      if (chart.chartStyleRoles?.trendlineLabel) {
        effective = chartStyleRoleTrendlineLabel(chart, effective);
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
  if (dataTable && chart.chartStyleRoles?.dataTable) {
    const effective = chartStyleRoleDataTable(chart, dataTable);
    changed ||= effective.lineColor !== dataTable.lineColor
      || effective.lineWidthEmu !== dataTable.lineWidthEmu
      || effective.lineDash !== dataTable.lineDash
      || effective.lineHidden !== dataTable.lineHidden;
    dataTable = effective;
  }
  const valMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.valAxisMajorGridlines,
    chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, chart.valAxisGridlineDash,
  );
  const catMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.catAxisMajorGridlines,
    chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, chart.catAxisGridlineDash,
  );
  const valMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.valAxisMinorGridlines,
    chart.valAxisMinorGridlineColor,
    chart.valAxisMinorGridlineWidthEmu,
    chart.valAxisMinorGridlineDash,
  );
  const catMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.catAxisMinorGridlines,
    chart.catAxisMinorGridlineColor,
    chart.catAxisMinorGridlineWidthEmu,
    chart.catAxisMinorGridlineDash,
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
  );
  const valAxisLine = chartStyleRoleAxisLine(
    chart, 'valueAxis',
    chart.valAxisLineColor, chart.valAxisLineWidthEmu, chart.valAxisLineDash,
    chart.valAxisLineHidden,
  );
  const catAxisStyle = chart.chartStyleRoles?.categoryAxis;
  const valAxisStyle = chart.chartStyleRoles?.valueAxis;
  const catAxisFontSizeHpt = chart.catAxisFontSizeHpt ?? catAxisStyle?.fontSizeHpt ?? null;
  const catAxisFontBold = chart.catAxisFontBold ?? catAxisStyle?.fontBold;
  const catAxisFontItalic = chart.catAxisFontItalic ?? catAxisStyle?.fontItalic;
  const catAxisFontColor = chart.catAxisFontColor ?? catAxisStyle?.fontColor;
  const catAxisFontFace = chart.catAxisFontFace ?? catAxisStyle?.fontFace;
  const valAxisFontSizeHpt = chart.valAxisFontSizeHpt ?? valAxisStyle?.fontSizeHpt ?? null;
  const valAxisFontBold = chart.valAxisFontBold ?? valAxisStyle?.fontBold;
  const valAxisFontItalic = chart.valAxisFontItalic ?? valAxisStyle?.fontItalic;
  const valAxisFontColor = chart.valAxisFontColor ?? valAxisStyle?.fontColor;
  const valAxisFontFace = chart.valAxisFontFace ?? valAxisStyle?.fontFace;
  changed ||= valMajor.visible !== chart.valAxisMajorGridlines
    || valMajor.color !== chart.valAxisGridlineColor
    || valMajor.widthEmu !== chart.valAxisGridlineWidthEmu
    || valMajor.dash !== chart.valAxisGridlineDash
    || catMajor.visible !== chart.catAxisMajorGridlines
    || catMajor.color !== chart.catAxisGridlineColor
    || catMajor.widthEmu !== chart.catAxisGridlineWidthEmu
    || catMajor.dash !== chart.catAxisGridlineDash
    || valMinor.visible !== chart.valAxisMinorGridlines
    || valMinor.color !== chart.valAxisMinorGridlineColor
    || valMinor.widthEmu !== chart.valAxisMinorGridlineWidthEmu
    || valMinor.dash !== chart.valAxisMinorGridlineDash
    || catMinor.visible !== chart.catAxisMinorGridlines
    || catMinor.color !== chart.catAxisMinorGridlineColor
    || catMinor.widthEmu !== chart.catAxisMinorGridlineWidthEmu
    || catMinor.dash !== chart.catAxisMinorGridlineDash
    || secondaryValAxis !== chart.secondaryValAxis
    || secondaryCatAxis !== chart.secondaryCatAxis
    || catAxisLine.color !== chart.catAxisLineColor
    || catAxisLine.widthEmu !== chart.catAxisLineWidthEmu
    || catAxisLine.dash !== chart.catAxisLineDash
    || (catAxisLine.hidden === true) !== chart.catAxisLineHidden
    || valAxisLine.color !== chart.valAxisLineColor
    || valAxisLine.widthEmu !== chart.valAxisLineWidthEmu
    || valAxisLine.dash !== chart.valAxisLineDash
    || (valAxisLine.hidden === true) !== chart.valAxisLineHidden
    || catAxisFontSizeHpt !== chart.catAxisFontSizeHpt
    || catAxisFontBold !== chart.catAxisFontBold
    || catAxisFontItalic !== chart.catAxisFontItalic
    || catAxisFontColor !== chart.catAxisFontColor
    || catAxisFontFace !== chart.catAxisFontFace
    || valAxisFontSizeHpt !== chart.valAxisFontSizeHpt
    || valAxisFontBold !== chart.valAxisFontBold
    || valAxisFontItalic !== chart.valAxisFontItalic
    || valAxisFontColor !== chart.valAxisFontColor
    || valAxisFontFace !== chart.valAxisFontFace;
  const effective = changed ? {
    ...chart,
    series,
    dataTable,
    valAxisMajorGridlines: valMajor.visible,
    valAxisGridlineColor: valMajor.color,
    valAxisGridlineWidthEmu: valMajor.widthEmu,
    valAxisGridlineDash: valMajor.dash,
    catAxisMajorGridlines: catMajor.visible,
    catAxisGridlineColor: catMajor.color,
    catAxisGridlineWidthEmu: catMajor.widthEmu,
    catAxisGridlineDash: catMajor.dash,
    valAxisMinorGridlines: valMinor.visible,
    valAxisMinorGridlineColor: valMinor.color,
    valAxisMinorGridlineWidthEmu: valMinor.widthEmu,
    valAxisMinorGridlineDash: valMinor.dash,
    catAxisMinorGridlines: catMinor.visible,
    catAxisMinorGridlineColor: catMinor.color,
    catAxisMinorGridlineWidthEmu: catMinor.widthEmu,
    catAxisMinorGridlineDash: catMinor.dash,
    secondaryValAxis,
    secondaryCatAxis,
    catAxisLineColor: catAxisLine.color,
    catAxisLineWidthEmu: catAxisLine.widthEmu,
    catAxisLineDash: catAxisLine.dash,
    catAxisLineHidden: catAxisLine.hidden === true,
    valAxisLineColor: valAxisLine.color,
    valAxisLineWidthEmu: valAxisLine.widthEmu,
    valAxisLineDash: valAxisLine.dash,
    valAxisLineHidden: valAxisLine.hidden === true,
    catAxisFontSizeHpt,
    catAxisFontBold,
    catAxisFontItalic,
    catAxisFontColor,
    catAxisFontFace,
    valAxisFontSizeHpt,
    valAxisFontBold,
    valAxisFontItalic,
    valAxisFontColor,
    valAxisFontFace,
  } : chart;
  return chartStyleRoleLegend(chartStyleRolePlotArea(chartStyleRoleChartArea(
    chartStyleRoleSeriesAxis(effective),
  )));
}

function drawUpDownBars(
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
    if (!paint.fillHidden && (paint.fill != null || fillColor != null)) {
      const resolvedFill = paint.fill
        ? resolveFill(paint.fill, ctx, barX, barY, barWidth, barHeight, shapeRotationDeg)
        : `#${fillColor}`;
      // An authored/linked structured fill owns this component even when it
      // cannot be resolved. Do not replace it with application-default paint.
      if (resolvedFill != null) {
        ctx.fillStyle = resolvedFill;
        ctx.fillRect(barX, barY, barWidth, barHeight);
      }
    }
    const lineOwned = paint.linePaintAuthored === true
      || paint.lineColor != null || paint.lineHidden === true;
    const lineColor = paint.lineColor ?? (lineOwned ? undefined : automaticPaint?.lineColor);
    const lineWidthEmu = paint.lineWidthEmu
      ?? (lineOwned ? undefined : automaticPaint?.lineWidthEmu);
    if (!paint.lineHidden
      && (paint.linePaintAuthored !== true || lineColor != null) && (
      lineColor != null || lineWidthEmu != null
    )) {
      const previousDash = ctx.getLineDash();
      const previousCap = ctx.lineCap;
      const previousJoin = ctx.lineJoin;
      ctx.strokeStyle = `#${lineColor ?? '000000'}`;
      ctx.lineWidth = lineWidthEmu != null
        ? axisLineWidthPx(lineWidthEmu, ptToPx)
        : Math.max(1, 0.75 * ptToPx);
      ctx.setLineDash(dashPatternForPreset(paint.lineDash ?? undefined, ctx.lineWidth));
      ctx.lineCap = paint.lineCap === 'rnd'
        ? 'round' : paint.lineCap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = paint.lineJoin === 'round' || paint.lineJoin === 'bevel'
        ? paint.lineJoin : 'miter';
      ctx.strokeRect(barX, barY, barWidth, barHeight);
      ctx.setLineDash(previousDash);
      ctx.lineCap = previousCap;
      ctx.lineJoin = previousJoin;
    }
  }
}

function drawLineGroupDecorations(
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
      const upDownBars = {
        ...decoration.upDownBars,
        up: chartStyleRoleBarPaint(chart, decoration.upDownBars.up, 'upBar'),
        down: chartStyleRoleBarPaint(chart, decoration.upDownBars.down, 'downBar'),
      };
      drawUpDownBars(
        ctx, index => valueFor(first, index), index => valueFor(last, index), pointCount, toX,
        yMapFor(first), yMapFor(last), slotWidth, upDownBars, ptToPx,
        // Empty upBars/downBars paint is application-defined. The retained
        // Office observation is limited to classic Style 2; other styles keep
        // the geometry/model but do not receive a guessed white/black paint.
        chart.legacyChartStyle === 2 ? {
          lineColor: '000000', lineWidthEmu: 9525,
          upFillColor: 'FFFFFF', downFillColor: '000000',
        } : undefined,
        shapeRotationDeg,
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

function axisCrossingValue(
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

function categoryAxisCrossingValue(chart: ChartModel, min: number, max: number): number {
  return axisCrossingValue(chart.catAxisCrossesAt, chart.catAxisCrosses, min, max);
}

export function renderLineChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);

  const legacyGrouping = chart.chartType === 'stackedLinePct'
    ? 'percentStacked' : chart.chartType === 'stackedLine' ? 'stacked' : 'standard';
  const lineGroups = chart.plotGroups?.filter(group => group.kind === 'line') ?? [{
    kind: 'line' as const,
    seriesStart: 0,
    seriesCount: chart.series.length,
    categoryAxis: 'primary' as const,
    valueAxis: 'primary' as const,
    seriesAxis: 'none' as const,
    grouping: legacyGrouping,
  }];
  const stackedBySeries = new Array<boolean>(chart.series.length).fill(false);
  const percentBySeries = new Array<boolean>(chart.series.length).fill(false);
  const secondaryBySeries = new Array<boolean>(chart.series.length).fill(false);
  const percentTotalsBySeries = new Array<number[] | null>(chart.series.length).fill(null);
  const plottedValues = chart.series.map(() => new Array<number>(n).fill(0));
  const allPercentByAxis = new Map<string, boolean>();
  for (const group of lineGroups) {
    const axis = group.valueAxis;
    allPercentByAxis.set(
      axis,
      (allPercentByAxis.get(axis) ?? true) && group.grouping === 'percentStacked',
    );
  }
  for (const group of lineGroups) {
    const grouping = group.grouping ?? 'standard';
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const pct = grouping === 'percentStacked';
    const members = chart.series.slice(group.seriesStart, group.seriesStart + group.seriesCount);
    const percentMultiplier = pct && allPercentByAxis.get(group.valueAxis) === true ? 100 : 1;
    const totals = pct
      ? cats.map((_, categoryIndex) => members.reduce(
          (sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0,
        ) || 1)
      : null;
    for (let offset = 0; offset < members.length; offset++) {
      const seriesIndex = group.seriesStart + offset;
      stackedBySeries[seriesIndex] = stacked;
      percentBySeries[seriesIndex] = pct;
      secondaryBySeries[seriesIndex] = group.valueAxis === 'secondary';
      percentTotalsBySeries[seriesIndex] = totals;
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        const raw = members[offset].values[categoryIndex] ?? 0;
        if (!stacked) {
          plottedValues[seriesIndex][categoryIndex] = raw;
          continue;
        }
        const prior = offset === 0 ? 0 : plottedValues[seriesIndex - 1][categoryIndex];
        const contribution = pct && totals
          ? raw / totals[categoryIndex] * percentMultiplier
          : raw;
        plottedValues[seriesIndex][categoryIndex] = prior + contribution;
      }
    }
  }
  const plotted = (seriesIndex: number, categoryIndex: number): number =>
    plottedValues[seriesIndex]?.[categoryIndex] ?? 0;
  const primaryGroups = lineGroups.filter(group => group.valueAxis !== 'secondary');
  const axisIsPercent = primaryGroups.length > 0
    && primaryGroups.every(group => group.grouping === 'percentStacked');
  // How null cells are plotted (`<c:dispBlanksAs>`, §21.2.2.42). Default "gap"
  // preserves the historical line break (byte-stable). "zero" treats a null as
  // 0; "span" bridges the neighbours with a straight line (skip the null but
  // keep the run going). Only unstacked charts see nulls — a stacked sum already
  // reads null as 0 — so the value only steers the unstacked path below.
  const dispBlanks = chart.dispBlanksAs ?? 'gap';

  // Combo line charts may bind some series to a SECONDARY value axis drawn on
  // the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with axPos="r"). `sec`
  // is non-null only when the axis is declared AND at least one series opts in;
  // secondary series are then excluded from the PRIMARY scale and mapped through
  // the secondary one. Stacked line charts stack ALL series onto the primary
  // axis (a percentStacked/stacked secondary combo is not an Office construct),
  // so the split only applies to plain (unstacked) line charts. When `sec` is
  // null every series stays on the primary axis, identical to the pre-CH7 path.
  const secondaryGroups = lineGroups.filter(group => group.valueAxis === 'secondary');
  const secondaryAxisIsPercent = secondaryGroups.length > 0
    && secondaryGroups.every(group => group.grouping === 'percentStacked');
  const sec = chart.secondaryValAxis && chart.series.some(
    (series, index) => secondaryBySeries[index] || (
      chart.plotGroups == null && series.useSecondaryAxis === true
    ),
  )
    ? chart.secondaryValAxis
    : null;
  const seriesIndexByIdentity = new Map(chart.series.map((series, index) => [series, index]));
  const isSecondarySeries = (series: ChartSeries): boolean => {
    const index = seriesIndexByIdentity.get(series) ?? -1;
    return sec != null && (secondaryBySeries[index]
      || (chart.plotGroups == null && series.useSecondaryAxis === true));
  };

  // Resolve the primary extent before frame placement. An authored
  // `layoutTarget="outer"` rectangle includes the value-axis labels, so its
  // conversion to the inner plot rectangle needs the width of the formatted
  // tick labels. This is the same extent used again for the final scale below.
  let dataMin = Infinity; let dataMax = -Infinity;
  for (let ci = 0; ci < n; ci++) {
    for (let si = 0; si < chart.series.length; si++) {
      if (isSecondarySeries(chart.series[si])) continue;
      if (!stackedBySeries[si] && chart.series[si].values[ci] == null) continue;
      const v = plotted(si, ci);
      dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v);
    }
  }
  for (let si = 0; si < chart.series.length; si++) {
    const series = chart.series[si];
    if (isSecondarySeries(series)) continue;
    forEachErrorBarEndpoint(
      series,
      'y',
      index => series.values[index] == null ? null : plotted(si, index),
      value => {
        dataMin = Math.min(dataMin, value);
        dataMax = Math.max(dataMax, value);
      },
    );
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  const isLogAxis = chart.valAxisLogBase != null && chart.valAxisLogBase >= 2;
  if (chart.valMin != null) dataMin = axisIsPercent ? chart.valMin * 100 : chart.valMin;
  else if (axisIsPercent && dataMin > 0 && !isLogAxis) dataMin = 0;
  if (chart.valMax != null) dataMax = axisIsPercent ? chart.valMax * 100 : chart.valMax;
  else if (axisIsPercent && dataMax < 0) dataMax = 0;

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  let titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  let titleFontPx = titleBand.fontPx;
  let titleTopPad = titleBand.topPad;
  let titleH = titleBand.bandH;
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  // Axis-title bands use the real title font (XML @sz when set), independent of
  // the tick-label sizes above, so 18pt titles get a wide enough gutter.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  // Vertical pads (independent of the right gutter) so an estimated plot height
  // is known before the secondary-axis scale + right-gutter measurement — the
  // same up-front ordering the bar renderer uses. The top adds half a value-axis
  // label so the topmost gridline label rides above the plot; the bottom reserves
  // PowerPoint's full category-label band (gap + line-height + margin).
  let padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  // Secondary value-axis scale (shared helper). Its axis is the vertical right
  // edge, so its length is the plot height. Null when there is no secondary axis.
  const secScale = computeSecondaryAxis(
    sec,
    chart.series,
    phEst / ptToPx,
    'y',
    secondaryAxisIsPercent,
    false,
    (_series, index) => secondaryBySeries[index]
      || (chart.plotGroups == null && chart.series[index].useSecondaryAxis === true),
    (series, pointIndex, seriesIndex) => !stackedBySeries[seriesIndex]
      && series.values[pointIndex] == null
      ? null : plotted(seriesIndex, pointIndex),
  );
  // Right-edge gutter for the secondary tick labels + rotated title. Measured
  // with the SAME font/format the axis is drawn with so the reserve matches the
  // painted labels (mirrors the bar renderer). Zero when there is no secondary
  // axis, so `pad.r` is unchanged on the common single-axis path.
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = chartFontCss(
      secFontPx,
      chartFontFamily(chart, sec.fontFace, 'minor'),
      false,
      sec.fontItalic ?? false,
    );
    let wmax = 0;
    for (const value of secScale.majorLines) {
      wmax = Math.max(wmax, ctx.measureText(formatAxisTickWithUnits(value, sec.formatCode ?? null, chart.date1904, sec.displayUnits)).width);
    }
    secLabelBandW = wmax + 18;
    ctx.font = prevFont;
  }
  const secTitleBandW = sec && sec.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  const titleLeftBandW = legLeftW + Math.max(
    valAxFontPx * 2.2 + 10 + valTitleW,
    dataTableHeaderW,
  );
  const titleRightBandW = legRightW + w * 0.05 + secLabelBandW + secTitleBandW;

  const provisionalPlan = planValueAxis(chart, dataMin, dataMax, phEst / ptToPx, axisIsPercent);
  let primaryLabelWidth = 0;
  if (
    !chart.valAxisHidden
    && chart.valAxisTickLabelPos !== 'none'
    && chart.plotAreaManualLayout != null
    && chart.plotAreaManualLayout.layoutTarget !== 'inner'
  ) {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    for (const value of provisionalPlan.majorLines) {
      primaryLabelWidth = Math.max(
        primaryLabelWidth,
        ctx.measureText(formatPrimaryValueAxisTick(chart, value, axisIsPercent)).width,
      );
    }
    ctx.font = previousFont;
  }
  // Pad based on actual label metrics rather than magic percents so an explicit
  // <c:txPr sz="1000"> (10pt) correctly compresses the plot area.
  const pad = {
    t: padT,
    r: titleRightBandW,
    b: padB,
    l: titleLeftBandW,
  };

  const manualOuterInsets = chartManualOuterAxisInsets({
    valAxisHidden: chart.valAxisHidden,
    catAxisHidden: chart.catAxisHidden,
    valLabelWidth: primaryLabelWidth,
    valLabelFontPx: valAxFontPx,
    catLabelFontPx: catAxFontPx,
    valLabelGapPx: chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(valAxFontPx)
      : 6,
    catLabelGapPx: chart.catAxisFontSizeHpt != null
      ? categoryLabelOffsetPx(
        categoryTickLabelGapPx(catAxFontPx),
        chart.catAxisLabelOffsetPercent,
      )
      : categoryLabelOffsetPx(5, chart.catAxisLabelOffsetPercent),
    outerTextMarginPx: AXIS_OUTER_TEXT_MARGIN_PT * ptToPx,
    valTitleBandW: valTitleW,
    catTitleBandH: catTitleH,
    secondaryBandW: secLabelBandW + secTitleBandW,
  });

  let lineFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  // The automatic title box follows the inner plot width. Resolve that width
  // from the same frame that owns the axis/legend gutters, then run one stable
  // vertical-layout pass if wrapping adds title lines.
  const plotWidthTitleBand = measuredCartesianTitleBand(
    ctx,
    chart,
    lineFrame.plotRect.pw,
    h,
    ptToPx,
  );
  if (Math.abs(plotWidthTitleBand.bandH - titleBand.bandH) > 0.01) {
    titleBand = plotWidthTitleBand;
    titleFontPx = titleBand.fontPx;
    titleTopPad = titleBand.topPad;
    titleH = titleBand.bandH;
    padT = titleH + legTopH + valAxFontPx / 2 + 2;
    pad.t = padT;
    lineFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
      titleBand,
      legendSideReserveFrac: 0.22,
      legendReserve: leg,
      pad,
      honorPlotAreaManualLayout: true,
      manualOuterInsets,
    });
  }
  const { px0, py0, pw } = lineFrame.plotRect;
  let { ph } = lineFrame.plotRect;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleTopPad, titleFontPx,
  );
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // Value axis is vertical → its length is the plot height (axis-length-aware
  // auto major unit, same model as the bar/column renderer). `planValueAxis`
  // folds in the CH6 major unit / logBase / orientation; with none set it is
  // byte-identical to the old `valueAxisScale` + linear `toY`.
  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx, axisIsPercent);
  if (plan.max - plan.min === 0) return;

  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;
  const primaryCategoryAxisY = toY(
    categoryAxisCrossingValue(chart, plan.min, plan.max),
  );
  const secondaryCategoryAxisY = sec && secScale
    ? toYSecondary(axisCrossingValue(
      chart.secondaryCatAxis?.crossesAt,
      chart.secondaryCatAxis?.crosses,
      secScale.min,
      secScale.max,
    ))
    : primaryCategoryAxisY;
  const categoryAxisYFor = (series: ChartSeries): number =>
    isSecondarySeries(series) ? secondaryCategoryAxisY : primaryCategoryAxisY;
  const primaryCatLine = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const primaryValLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const primaryCatTickColor = chart.catAxisLineColor != null ? primaryCatLine.color : undefined;
  const primaryCatTickWidth = chart.catAxisLineWidthEmu != null ? primaryCatLine.width : undefined;
  const primaryValTickColor = chart.valAxisLineColor != null ? primaryValLine.color : undefined;
  const primaryValTickWidth = chart.valAxisLineWidthEmu != null ? primaryValLine.width : undefined;
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  // crossBetween="between" (default) insets the first/last category by half a
  // step so points aren't flush against the axes. "midCat" anchors them.
  // A `maxMin` category orientation (§21.2.2.130) mirrors the index left↔right.
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const toX = dateAxisPlan
    ? (i0: number) => px0 + dateAxisPlan.positions[i0]! * pw
    : between
      ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
      : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    // Resolved gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or default).
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    // Minor gridlines first (under the majors), then major gridlines + ticks +
    // labels. Minor lines are only populated when the file declares them.
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, primaryValTickColor, primaryValTickWidth, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        const gap = chart.valAxisFontSizeHpt != null
          ? valueTickLabelGapPx(valAxFontPx)
          : 6;
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, axisIsPercent), px0 - gap, gy);
      }
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of plan.minorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, toY(value), primaryValTickColor, primaryValTickWidth, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }


  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks across the plot height. Off by default
  // (byte-stable). Shared placement with the bar renderer via
  // `categoryGridlineFractions`.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    const previousDash = cg.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (cg.dash.length > 0) ctx.setLineDash(cg.dash);
    const fractions = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
      : catGridlineFractions(chart, n);
    for (const frac of fractions) {
      const gx = px0 + frac * pw;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (cg.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // Axis lines: bottom (category) + left (value). Both default to visible
  // unless hidden explicitly. Office treats `<c:spPr><a:ln><a:noFill>` as
  // suppressing the rule and tick marks while retaining labels/gridlines.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, primaryCategoryAxisY, px0 + pw, primaryCategoryAxisY,
      primaryCatLine.color, primaryCatLine.width, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      primaryValLine.color, primaryValLine.width, chart.valAxisLineDash,
    );
  }

  // CT_LineChart owns drop lines, high-low lines, and up/down bars at the
  // group level (ECMA-376 §21.2.2 EG_LineChartShared / CT_LineChart). Paint
  // this background geometry before the series so authored lines and markers
  // remain on top, as in Office. Group provenance prevents a second line group
  // in a combo chart from inheriting the first group's decorations.
  const decorationSlotWidth = dateAxisPlan
    ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw
    : between ? pw / n : n > 1 ? pw / (n - 1) : pw;
  const lineSeriesIndex = new Map(
    chart.series.map((series, seriesIndex) => [series, seriesIndex]),
  );
  drawLineGroupDecorations(
    ctx, chart, n, toX, yMapFor,
    categoryAxisYFor,
    (series, index) => {
      const seriesIndex = lineSeriesIndex.get(series);
      return seriesIndex != null ? plotted(seriesIndex, index) : null;
    },
    decorationSlotWidth, ptToPx, shapeRotationDeg, 'background',
  );

  // Line width and marker size come from OOXML in points (<a:ln w=EMU> /
  // <c:marker><c:size val=pt>). Omitted series strokes keep the PowerPoint
  // defaults (2.25pt line, 5pt marker diameter) scaled to the viewport.
  const lineWidthPx = Math.max(1, 2.25 * ptToPx);
  const markerR = Math.max(2, 2.5 * ptToPx);
  const dataLabelPx = axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
  // Data labels are a chart-wide foreground layer. Painting them inside the
  // per-series loop lets a later series line (or trendline) cross labels that
  // belong to an earlier series. Excel keeps every series label above all
  // series geometry, so collect the label painters and flush them only after
  // every line, error bar, marker, and trendline has been painted.
  const deferredDataLabels: Array<() => void> = [];
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const seriesStacked = stackedBySeries[si];
    const seriesPercentTotals = percentTotalsBySeries[si];
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const color = chartColor(si, s);
    // Secondary series ride their own vertical scale; primary series (and every
    // series when there is no secondary axis) map through the primary `toY`.
    const yOf = yMapFor(s);
    ctx.strokeStyle = s.lineColor ? `#${s.lineColor}` : color;
    ctx.lineWidth = s.lineWidthEmu != null
      ? axisLineWidthPx(s.lineWidthEmu, ptToPx)
      : lineWidthPx;
    ctx.setLineDash(dashPatternForPreset(s.chartexStyle?.lineDash ?? undefined, ctx.lineWidth));
    ctx.lineCap = s.chartexStyle?.lineCap === 'rnd'
      ? 'round'
      : s.chartexStyle?.lineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = s.chartexStyle?.lineJoin === 'round' || s.chartexStyle?.lineJoin === 'bevel'
      ? s.chartexStyle.lineJoin
      : 'miter';
    ctx.beginPath();
    // Collect runs of consecutive present points (a null breaks the line into a
    // fresh run; stacked charts have no nulls in the plotted sum). Each run is
    // stroked as a polyline or a smooth spline (§21.2.2.194) via appendCurve.
    // For a non-smooth series this emits the exact prior moveTo/lineTo sequence
    // (byte-stable); smooth swaps the straight segments for a Bézier curve.
    const smooth = s.smooth === true;
    let run: Array<{ x: number; y: number }> = [];
    const flushRun = (): void => {
      if (run.length === 0) return;
      ctx.moveTo(run[0].x, run[0].y);
      appendCurve(ctx, run, smooth);
      run = [];
    };
    for (let ci = 0; ci < n; ci++) {
      // Unstacked null handling per dispBlanksAs (§21.2.2.42): "gap" flushes the
      // run (line breaks — the historical default); "span" skips the null but
      // keeps the run open (neighbours join directly); "zero" plots it at 0
      // (plotted() reads a null as 0). Stacked charts never have plotted nulls.
      if (!seriesStacked && s.values[ci] == null) {
        if (dispBlanks === 'gap') { flushRun(); continue; }
        if (dispBlanks === 'span') continue;
        // "zero": fall through and push a point at value 0.
      }
      run.push({ x: toX(ci), y: yOf(plotted(si, ci)) });
    }
    flushRun();
    if (s.lineHidden !== true) ctx.stroke();

    // Error bars (`<c:errBars>`, §21.2.2.20) — drawn under the markers so the
    // dots overlay the bar tips. Only fires for series that carry them.
    const plottedOf = (ci: number): number => plotted(si, ci);
    for (const eb of s.errBars ?? []) {
      drawCategoryErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, color,
      );
    }

    ctx.fillStyle = color;
    // ECMA-376 §21.2.2.32 — when the series resolves to no marker, skip the
    // data-point dots but keep data labels. Markers / labels pin to the plotted
    // (cumulative) value so they ride the stacked line, not the raw datum.
    const seriesMarkersVisible = s.showMarker !== false && s.markerSymbol !== 'none';
    const drawMarkers = seriesMarkersVisible || hasVisiblePointMarkerOverride(s);
    // Series carrying explicit `<c:marker>` detail route through drawMarker
    // (symbol/size/fill/line + per-point `<c:dPt>` overrides). Series without
    // any detail keep the historical fixed-circle fast path unchanged
    // (byte-stable). `markerSymbol: "none"` is caught by the showMarker gate.
    const hasMarkerDetail = seriesHasMarkerDetail(s);
    // Per-point / series-level data labels (`<c:dLbl idx>` / `<c:dLbls>`) take
    // precedence over the family's simple `showDataLabels` value dump. Merely
    // decide the route here; painting is deferred until all series geometry is
    // complete so labels remain the chart-wide foreground layer.
    const perPointLabels = (s.dataLabelOverrides?.length ?? 0) > 0 || s.seriesDataLabels != null;
    if (perPointLabels) {
      deferredDataLabels.push(() => {
        drawCategoryDataLabels(
          ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false,
          // Mirror the marker loop's gate just below: stacked series never see a
          // plotted null (a stacked sum already reads null as 0), and unstacked
          // "zero" mode plots the null at 0 — both cases get a label too.
          seriesStacked || dispBlanks === 'zero',
          chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
          // §21.2.2.48 `<c:dLblPos>` precedence: per-point/series positions win,
          // else the chart-level position, else the line-chart default `'r'`.
          chart.dataLabelPosition ?? 'r',
          // Automatic endpoint labels may occupy the chart gutter outside the
          // plot area (the plot's manual layout often reserves that space).
          // Keep vertical clipping aligned to the plot, but clamp horizontally
          // to the chart rectangle so `l`/`r` remain outside the end markers.
          { x, y: py0, w, h: ph },
          { x, y, w, h },
          percentBySeries[si] && seriesPercentTotals
            ? ci => (s.values[ci] ?? 0) / seriesPercentTotals[ci]
            : undefined,
          ci => {
            if (!drawMarkers) return 0;
            const dpt = pointOverrides.get(ci);
            if (!hasMarkerDetail && !pointHasMarkerDetail(dpt)) return markerR;
            const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
            if (symbol === 'none') return 0;
            return ((dpt?.markerSize ?? s.markerSize ?? 5) / 2) * ptToPx;
          },
          face => chartFontFamily(chart, face, 'minor'),
          isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
          ci => dataLabelLegendKey(si, ci),
          value => dataLabelWithinAxisMaximum(
            chart, value,
            isSecondarySeries(s) && secScale ? secScale.max : plan.max,
          ),
          shapeRotationDeg,
        );
      });
    }
    for (let ci = 0; ci < n; ci++) {
      // A null point gets a marker/label only in "zero" mode (plotted at 0);
      // "gap"/"span" leave the hole empty.
      if (!seriesStacked && s.values[ci] == null && dispBlanks !== 'zero') continue;
      const pv = plotted(si, ci);
      if (drawMarkers) {
        const dpt = pointOverrides.get(ci);
        if (hasMarkerDetail || pointHasMarkerDetail(dpt)) {
          const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
          if (symbol !== 'none') {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = markerFillColorFor(s, dpt, ci, color);
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            const lineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawMarker(
              ctx, toX(ci), yOf(pv), symbol, sizePt, fill, line, ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
            );
          }
        } else {
          ctx.beginPath(); ctx.arc(toX(ci), yOf(pv), markerR, 0, Math.PI * 2); ctx.fill();
        }
      }
    }

    if (chart.showDataLabels && !perPointLabels) {
      deferredDataLabels.push(() => {
        for (let ci = 0; ci < n; ci++) {
          if (!seriesStacked && s.values[ci] == null && dispBlanks !== 'zero') continue;
          const pv = plotted(si, ci);
          // §21.2.2.48 `<c:dLblPos>`: the family-level value dump honors the
          // chart-level position (else the line default `'r'`). The marker gap
          // stays directional while the whole label layer is painted last.
          const labelText = effectiveDataLabelText({
            showValue: true,
            sourceValue: s.values[ci] ?? 0,
            valueDivisor: displayUnitDivisor(
              isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
            ),
            formatCode: chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
            date1904: chart.date1904,
          });
          drawDataLabelText(
            ctx, toX(ci), yOf(pv), labelText,
            chart.dataLabelPosition ?? 'r', dataLabelPx,
            chart.dataLabelFontColor ?? undefined, chart.dataLabelFontBold ?? false,
            chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
            drawMarkers ? markerR + 1 : 2,
            { x: px0, y: py0, w: pw, h: ph },
          );
        }
      });
    }

    // Trendlines (`<c:trendline>`, §21.2.2.211) over this series' points —
    // drawn on top of the line/markers, dashed, in the series color unless the
    // trendline declares its own `<a:ln>`.
    drawSeriesTrendlines(
      ctx, s, color, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  drawLineGroupDecorations(
    ctx, chart, n, toX, yMapFor,
    categoryAxisYFor,
    (series, index) => {
      const seriesIndex = lineSeriesIndex.get(series);
      return seriesIndex != null ? plotted(seriesIndex, index) : null;
    },
    decorationSlotWidth, ptToPx, shapeRotationDeg, 'foreground',
  );

  for (const drawLabels of deferredDataLabels) drawLabels();

  if (!chart.catAxisHidden) {
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      catAxFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Tick marks and labels have independent authored skip intervals
    // (§21.2.2.205/§21.2.2.206). When tickLblSkip is absent every non-empty
    // cached category is paintable; sparse caches deliberately use blank
    // indices to author intervals such as every second year.
    const tickInterval = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    const majorTickXs = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => px0 + tick.fraction * pw)
      : Array.from({ length: Math.ceil(n / tickInterval) }, (_, index) => toX(index * tickInterval));
    for (const tx of majorTickXs) {
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, tx, primaryCatTickColor, primaryCatTickWidth, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none' && dateAxisPlan) {
      for (const tick of dateAxisPlan.minorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMinorTickMark, 'cat', primaryCategoryAxisY,
          px0 + tick.fraction * pw, primaryCatTickColor, primaryCatTickWidth,
          false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
        );
      }
    }
    const showLabels = !hasDataTable && catLabelsVisible(chart);
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const rotRad = catLabelRotationRad(chart);
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        label: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        x: px0 + tick.fraction * pw,
        categoryIndex: -1,
      }))
      : Array.from({ length: Math.ceil(n / labelInterval) }, (_, index) => {
        const ci = index * labelInterval;
        return {
          label: formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904),
          x: toX(ci),
          categoryIndex: ci,
        };
      });
    for (const entry of labelEntries) {
      const anchor = entry.categoryIndex < 0
        ? null
        : categoryLabelAnchorFraction(
          entry.categoryIndex,
          n,
          isCrossBetween(chart),
          catAxisReversed(chart),
          chart.catAxisLabelAlignment,
        );
      const tx = anchor ? px0 + anchor.fraction * pw : entry.x;
      if (!showLabels) continue;
      ctx.textAlign = anchor?.textAlign ?? 'center';
      ctx.fillStyle = catLabelColor;
      // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
      // category-axis numFmt; string categories pass through unchanged.
      const label = entry.label;
      if (!label) continue;
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(catAxFontPx)
          : 5,
        chart.catAxisLabelOffsetPercent,
      );
      const labelPosition = chart.catAxisTickLabelPos ?? 'nextTo';
      const labelAxisY = labelPosition === 'nextTo'
        ? primaryCategoryAxisY
        : labelPosition === 'high' ? py0 : py0 + ph;
      drawRotatedCatLabel(ctx, label, tx, labelAxisY + gap, rotRad);
    }
  }

  // Secondary value axis (right edge) — drawn after the series + category labels
  // so it sits atop the plot, mirroring the bar renderer's ordering.
  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r, px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  if (dataTableLayout) {
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, py0 + ph, pw, x + legLeftW, ptToPx,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Stock chart (ECMA-376 §21.2.2.198)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * High-low-close (and open-high-low-close) stock chart. Series order is fixed
 * by the spec: a 3-series chart is High, Low, Close; a 4-series chart is Open,
 * High, Low, Close. For each category we draw:
 *   - a thin vertical "hi-lo line" from the Low value to the High value
 *     (`<c:hiLowLines>`, §21.2.2.60) — always, when hiLowLines is present;
 *   - the Close series marker at its value (a short tick / dot);
 *   - the Open series marker (4-series only).
 * The value axis, date/category axis, title and legend reuse the shared
 * Cartesian scaffolding (identical to the line renderer). Four-series charts
 * also draw `<c:upDownBars>` (§21.2.2.227) between the first and last
 * authored series. In a four-series stock chart those are Open and Close; the
 * schema also permits the element on a three-series High/Low/Close chart.
 */
function renderStockChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);

  // Fixed spec series roles by position. With 4 series the first is Open; the
  // last three are always High, Low, Close. Fewer than 3 series can't form a
  // hi-lo-close plot, so fall back to plotting each series' markers only.
  const stockGroup = chart.plotGroups?.find(group => group.kind === 'stock');
  const series = stockGroup
    ? chart.series.slice(stockGroup.seriesStart, stockGroup.seriesStart + stockGroup.seriesCount)
    : chart.series;
  const lineOverlaySeries = chart.plotGroups == null
    ? []
    : chart.plotGroups
        .filter(group => group.kind === 'line')
        .flatMap(group => chart.series.slice(
          group.seriesStart, group.seriesStart + group.seriesCount,
        ));
  const scaleSeries = [...series, ...lineOverlaySeries];
  const sourceSeriesIndices = new Map(chart.series.map((entry, index) => [entry, index]));
  const hasOpen = series.length >= 4;
  const openIdx = hasOpen ? 0 : -1;
  const highIdx = hasOpen ? 1 : 0;
  const lowIdx = hasOpen ? 2 : 1;
  const closeIdx = hasOpen ? 3 : 2;
  const highS = series[highIdx];
  const lowS = series[lowIdx];
  const closeS = series[closeIdx] as ChartSeries | undefined;
  const openS = openIdx >= 0 ? series[openIdx] : undefined;
  const upDownStartS = series[0] as ChartSeries | undefined;
  const upDownEndS = series.at(-1) as ChartSeries | undefined;
  const sec = chart.secondaryValAxis && scaleSeries.some(stockSeries =>
    stockSeries.useSecondaryAxis === true
  ) ? chart.secondaryValAxis : null;
  const isSecondarySeries = (stockSeries: ChartSeries): boolean =>
    sec != null && stockSeries.useSecondaryAxis === true;

  // ── Shared Cartesian frame (mirrors renderLineChart's band computation) ──
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legBottomH, legTopH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;

  const phEst = h - padT - padB;
  const secScale = computeSecondaryAxis(sec, scaleSeries, phEst / ptToPx);
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      secFontPx,
      chartFontFamily(chart, sec.fontFace, 'minor'),
      sec.fontBold ?? false,
      sec.fontItalic ?? false,
    );
    let maxLabelWidth = 0;
    for (const value of secScale.majorLines) {
      maxLabelWidth = Math.max(maxLabelWidth, ctx.measureText(formatAxisTickWithUnits(
        value, sec.formatCode ?? null, chart.date1904, sec.displayUnits,
      )).width);
    }
    secLabelBandW = maxLabelWidth + 18;
    ctx.font = previousFont;
  }
  const secTitleBandW = sec?.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: legLeftW + Math.max(valAxFontPx * 2.2 + 10 + valTitleW, dataTableHeaderW),
  };

  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  const stockFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  const { px0, py0, pw } = stockFrame.plotRect;
  let { ph } = stockFrame.plotRect;
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // ── Value-axis extent: across every series' plotted values (the hi-lo line
  // needs both the low and high extremes). Authored bounds are retained and
  // omitted bounds flow through the shared automatic planner. ──
  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of scaleSeries) {
    if (isSecondarySeries(s)) continue;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      dataMin = Math.min(dataMin, v);
      dataMax = Math.max(dataMax, v);
    }
  }
  for (const stockSeries of scaleSeries) {
    if (isSecondarySeries(stockSeries)) continue;
    forEachErrorBarEndpoint(
      stockSeries,
      'y',
      index => stockSeries.values[index] ?? null,
      value => {
        dataMin = Math.min(dataMin, value);
        dataMax = Math.max(dataMax, value);
      },
    );
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (chart.valMin != null) dataMin = chart.valMin;
  if (chart.valMax != null) dataMax = chart.valMax;

  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx);
  if (plan.max - plan.min === 0) return;
  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;
  const toYSecondary = secScale?.makeToY(py0, ph) ?? toY;
  const toYFor = (stockSeries: ChartSeries): ((value: number) => number) =>
    isSecondarySeries(stockSeries) ? toYSecondary : toY;

  // Category X mapping — stock charts use crossBetween="between" by default so
  // the first/last hi-lo line isn't flush against the axes (matches Excel).
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  const toX = dateAxisPlan
    ? (i0: number) => px0 + dateAxisPlan.positions[i0]! * pw
    : between
      ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
      : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  // ── Value axis: gridlines + ticks + labels (identical to the line renderer) ──
  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, undefined, undefined, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), px0 - 6, gy);
      }
    }
    for (const v of plan.minorTicks) {
      drawAxisTick(
        ctx, chart.valAxisMinorTickMark, 'val', px0, toY(v),
        undefined, undefined, false, chart.valAxisLineHidden, 'minor', ptToPx,
        chart.valAxisLineDash,
      );
    }
  }

  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Axis rules (bottom = category, left = value).
  const stockCatLine = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const stockValLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0 + ph, px0 + pw, py0 + ph,
      stockCatLine.color, stockCatLine.width, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      stockValLine.color, stockValLine.width, chart.valAxisLineDash,
    );
  }

  // CT_StockChart/dropLines uses the same chart-line paint contract as line
  // and area charts. A stock drop line connects the category axis to the
  // envelope of every finite stock value at that category.
  if (chart.stockDropLines) {
    const linked = chartStyleRoleLine(chart, chart.stockDropLines, 'dropLine');
    const dropLineStyle = {
      ...linked,
      color: linked.color ?? (linked.paintAuthored === true
        ? null : chart.stockAutomaticStyle?.lineColor),
      widthEmu: linked.widthEmu ?? chart.stockAutomaticStyle?.lineWidthEmu,
    };
    if ((dropLineStyle.paintAuthored !== true || dropLineStyle.color != null)
      && (dropLineStyle.color != null || dropLineStyle.widthEmu != null
      || dropLineStyle.dash != null) && applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      drawDropLineEnvelopes(
        ctx,
        series,
        n,
        toX,
        stockSeries => toYFor(stockSeries),
        () => py0 + ph,
        (stockSeries, index) => stockSeries.values[index] ?? null,
      );
    }
  }

  // ── Hi-lo lines: vertical Low↔High per category. CT_StockChart makes
  // `<c:hiLowLines>` optional, so absence must remain absence; only a present
  // element receives linked or bounded automatic paint. ──
  const drawHiLo = chart.stockHiLowLines === true && highS != null && lowS != null;
  if (drawHiLo && highS && lowS) {
    const directStyle = chart.stockHiLowLineStyle ?? {
      color: chart.stockHiLowLineColor ?? null,
    };
    const linkedStyle = chartStyleRoleLine(chart, directStyle, 'hiLoLine');
    const lineStyle = {
      ...linkedStyle,
      color: linkedStyle.color ?? (linkedStyle.paintAuthored === true
        ? null : chart.stockAutomaticStyle?.lineColor),
      widthEmu: linkedStyle.widthEmu ?? chart.stockAutomaticStyle?.lineWidthEmu,
    };
    if ((lineStyle.paintAuthored !== true || lineStyle.color != null)
      && (lineStyle.color != null || lineStyle.widthEmu != null || lineStyle.dash != null)
      && applyDecorationLineStyle(ctx, lineStyle, ptToPx)) {
      for (let ci = 0; ci < n; ci++) {
        const hi = highS.values[ci];
        const lo = lowS.values[ci];
        if (hi == null || lo == null) continue;
        const cx = toX(ci);
        const highToY = toYFor(highS);
        const lowToY = toYFor(lowS);
        ctx.beginPath();
        ctx.moveTo(cx, highToY(hi));
        ctx.lineTo(cx, lowToY(lo));
        ctx.stroke();
      }
    }
  }

  // ── Close (and Open) markers. A stock chart's close is drawn as a short tick.
  // If the series carries an explicit `<c:marker>` (symbol/size/fill), honor it;
  // otherwise draw a left/right tick in the series color. ──
  const drawStockTick = (
    s: ChartSeries | undefined,
    seriesIndex: number,
    side: 'left' | 'right' | 'both',
  ): void => {
    if (!s) return;
    const color = chartColor(seriesIndex, s);
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const seriesMarkerVisible = s.markerSymbol != null && s.markerSymbol !== 'none'
      && seriesHasMarkerDetail(s);
    const tickLen = Math.max(3, (pw / n) * 0.22);
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      const cx = toX(ci);
      const cy = toYFor(s)(v);
      const point = pointOverrides.get(ci);
      if (point?.markerSymbol === 'none' || (point?.markerSymbol == null && s.markerSymbol === 'none')) {
        continue;
      }
      const hasExplicitMarker = seriesMarkerVisible
        || (point?.markerSymbol != null && point.markerSymbol !== 'none');
      if (hasExplicitMarker) {
        const symbol = point?.markerSymbol ?? s.markerSymbol ?? 'circle';
        drawMarker(
          ctx, cx, cy, symbol as string,
          point?.markerSize ?? s.markerSize ?? 3,
          markerFillColorFor(s, point, ci, color),
          point?.markerLine ?? s.markerLine ?? null,
          ptToPx,
          (point?.markerLineWidthEmu ?? s.markerLineWidthEmu) != null
            ? axisLineWidthPx(
                (point?.markerLineWidthEmu ?? s.markerLineWidthEmu) as number,
                ptToPx,
              )
            : undefined,
          markerFillPaintFor(s, point, ci), shapeRotationDeg,
        );
        continue;
      }
      // Horizontal tick: close ticks to the RIGHT of the line, open ticks to the
      // LEFT (Excel's open-high-low-close convention). `both` centers it.
      ctx.strokeStyle = color;
      ctx.lineWidth = Math.max(1, 0.75 * ptToPx);
      ctx.beginPath();
      const x0 = side === 'right' ? cx : side === 'left' ? cx - tickLen : cx - tickLen / 2;
      const x1 = side === 'right' ? cx + tickLen : side === 'left' ? cx : cx + tickLen / 2;
      ctx.moveTo(x0, cy);
      ctx.lineTo(x1, cy);
      ctx.stroke();
    }
  };
  // Office accepts a line group after a stock group. Stock decorations retain
  // ownership of the stock slice; the later line group is a normal category
  // line overlay on its resolved value axis rather than becoming a fifth stock
  // role.
  for (const lineSeries of lineOverlaySeries) {
    const chartIndex = sourceSeriesIndices.get(lineSeries) ?? 0;
    const color = chartColor(Math.max(0, chartIndex), lineSeries);
    const yOf = toYFor(lineSeries);
    const pointOverrides = indexPointOverrides(lineSeries.dataPointOverrides);
    if (lineSeries.lineHidden !== true) {
      const structuredLine = chartExStyleLinePaintDecision(
        chart, lineSeries.chartexStyle, chartIndex, chart.series.length,
      );
      const resolvedLine = structuredLine === undefined
        ? lineSeries.lineColor ? `#${lineSeries.lineColor}` : color
        : structuredLine == null
          ? null
          : resolveFill(structuredLine, ctx, px0, py0, pw, ph, shapeRotationDeg);
      if (resolvedLine != null) {
        ctx.save();
        ctx.strokeStyle = resolvedLine;
      ctx.lineWidth = lineSeries.lineWidthEmu != null
        ? axisLineWidthPx(lineSeries.lineWidthEmu, ptToPx)
        : Math.max(1, 2.25 * ptToPx);
      ctx.setLineDash(drawingmlLineDashArray(
        lineSeries.chartexStyle?.lineCustomDash,
        lineSeries.chartexStyle?.lineDash,
        ctx.lineWidth,
      ));
      ctx.lineCap = lineSeries.chartexStyle?.lineCap === 'rnd'
        ? 'round' : lineSeries.chartexStyle?.lineCap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = lineSeries.chartexStyle?.lineJoin === 'round'
        || lineSeries.chartexStyle?.lineJoin === 'bevel'
        ? lineSeries.chartexStyle.lineJoin : 'miter';
      ctx.beginPath();
      let run: Array<{ x: number; y: number }> = [];
      const flushRun = (): void => {
        if (run.length === 0) return;
        ctx.moveTo(run[0].x, run[0].y);
        appendCurve(ctx, run, lineSeries.smooth === true);
        run = [];
      };
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        const value = lineSeries.values[categoryIndex];
        if (value == null) { flushRun(); continue; }
        run.push({ x: toX(categoryIndex), y: yOf(value) });
      }
      flushRun();
      ctx.stroke();
        ctx.restore();
      }
    }
    const seriesMarkersVisible = lineSeries.showMarker !== false
      && lineSeries.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(lineSeries)) {
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        const value = lineSeries.values[categoryIndex];
        if (value == null) continue;
        const point = pointOverrides.get(categoryIndex);
        const symbol = effectiveMarkerSymbol(
          lineSeries, point, 'circle', seriesMarkersVisible,
        );
        if (symbol === 'none') continue;
        drawMarker(
          ctx, toX(categoryIndex), yOf(value), symbol,
          point?.markerSize ?? lineSeries.markerSize ?? 5,
          markerFillColorFor(lineSeries, point, categoryIndex, color),
          point?.markerLine ?? lineSeries.markerLine ?? null,
          ptToPx,
          (point?.markerLineWidthEmu ?? lineSeries.markerLineWidthEmu) != null
            ? axisLineWidthPx(
                (point?.markerLineWidthEmu ?? lineSeries.markerLineWidthEmu) as number,
                ptToPx,
              )
            : undefined,
          markerFillPaintFor(lineSeries, point, categoryIndex),
          shapeRotationDeg,
        );
      }
    }
    drawSeriesTrendlines(
      ctx, lineSeries, color, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  // ── First/last-series up-down bars (§21.2.2.218/227). In DrawingML these
  // decorations follow the line series. Excel paints their opaque bodies over
  // both the high-low rule and the owning series lines, leaving plot geometry
  // visible only outside each body. Explicit stock marker glyphs are replayed
  // after the bodies below, so a marker at a bar endpoint remains fully visible.
  if (chart.stockUpDownBars && upDownStartS && upDownEndS) {
    const directStyle = chart.stockUpDownBarStyle ?? {
      gapWidthPercent: 150,
      up: {},
      down: {},
    };
    const style = {
      ...directStyle,
      up: chartStyleRoleBarPaint(chart, directStyle.up, 'upBar'),
      down: chartStyleRoleBarPaint(chart, directStyle.down, 'downBar'),
    };
    const slotWidth = dateAxisPlan
      ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw
      : between ? pw / n : n > 1 ? pw / (n - 1) : pw;
    drawUpDownBars(
      ctx,
      index => upDownStartS.values[index] ?? null,
      index => upDownEndS.values[index] ?? null,
      n, toX, toYFor(upDownStartS), toYFor(upDownEndS),
      slotWidth, style, ptToPx, chart.stockAutomaticStyle ?? undefined, shapeRotationDeg,
    );
  }

  // Marker/tick glyphs are the foreground annotation of a stock datum. Desktop
  // Excel paints them after up/down-bar bodies; otherwise a circular High/Low
  // marker intersecting the body is clipped to a semicircle.
  drawStockTick(openS, openIdx, 'left');
  if (highS?.markerSymbol != null || (highS && hasVisiblePointMarkerOverride(highS))) {
    drawStockTick(highS, highIdx, 'both');
  }
  if (lowS?.markerSymbol != null || (lowS && hasVisiblePointMarkerOverride(lowS))) {
    drawStockTick(lowS, lowIdx, 'both');
  }
  drawStockTick(closeS, closeIdx, 'right');

  // CT_LineSer error bars remain attached to their authored stock series.
  // Stock uses a category X axis, so only Y-direction bars have data-unit
  // geometry; the same shared category-series painter is used by line/area.
  for (const stockSeries of scaleSeries) {
    const seriesIndex = sourceSeriesIndices.get(stockSeries) ?? 0;
    const color = chartColor(seriesIndex, stockSeries);
    for (const errorBars of stockSeries.errBars ?? []) {
      drawCategoryErrorBars(
        ctx,
        stockSeries,
        chartStyleRoleErrorBar(chart, errorBars),
        n,
        toX,
        toYFor(stockSeries),
        index => stockSeries.values[index] ?? 0,
        color,
      );
    }
  }

  // If fewer than 3 series (not a real hi-lo-close), still plot each series'
  // markers so nothing is silently dropped.
  if (series.length < 3) {
    for (let si = 0; si < series.length; si++) {
      drawStockTick(series[si], si, 'both');
    }
  }

  // CT_StockChart owns CT_LineSer children, so the same series/default and
  // per-point dLbl contracts used by ordinary category-line charts apply here
  // as well. Paint labels after stock glyphs/error bars so their callout boxes
  // and text remain on top of the plot geometry.
  for (const stockSeries of scaleSeries) {
    const seriesIndex = sourceSeriesIndices.get(stockSeries) ?? 0;
    drawSeriesDataLabels(
      ctx,
      stockSeries,
      cats,
      true,
      toX,
      toYFor(stockSeries),
      ph,
      ptToPx,
      chart.date1904,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      chart.dataLabelPosition ?? 'r',
      { x: px0, y: py0, w: pw, h: ph },
      r,
      face => chartFontFamily(chart, face, 'minor'),
      isSecondarySeries(stockSeries) ? sec?.displayUnits : chart.valAxisDisplayUnits,
      pointIndex => dataLabelLegendKey(seriesIndex, pointIndex),
      value => dataLabelWithinAxisMaximum(
        chart, value, isSecondarySeries(stockSeries) ? secScale?.max ?? plan.max : plan.max,
      ),
      shapeRotationDeg,
    );
  }

  // ── Category (date) axis labels — same path as the line renderer. ──
  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      catAxFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    const catSlotMaxPx = dateAxisPlan
      ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw - 4
      : (pw / n) * labelInterval - 4;
    const showLabels = !hasDataTable && catLabelsVisible(chart);
    const rotRad = catLabelRotationRad(chart);
    const labelEntries = dateAxisPlan && dateAxisPlan.majorTicks.length > 0
      ? dateAxisPlan.majorTicks.map(tick => ({
        label: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        x: px0 + tick.fraction * pw,
        categoryIndex: -1,
      }))
      : Array.from({ length: Math.ceil(n / labelInterval) }, (_, index) => {
        const ci = index * labelInterval;
        return {
          label: formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904),
          x: toX(ci),
          categoryIndex: ci,
        };
      });
    for (const entry of labelEntries) {
      const anchor = entry.categoryIndex < 0
        ? null
        : categoryLabelAnchorFraction(
          entry.categoryIndex,
          n,
          isCrossBetween(chart),
          catAxisReversed(chart),
          chart.catAxisLabelAlignment,
        );
      const tx = anchor ? px0 + anchor.fraction * pw : entry.x;
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx, stockCatLine.color, stockCatLine.width, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      if (!showLabels) continue;
      ctx.textAlign = anchor?.textAlign ?? 'center';
      ctx.fillStyle = catLabelColor;
      const label = entry.label;
      const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
      drawRotatedCatLabel(
        ctx,
        elideToWidth(ctx, label, budget),
        tx,
        py0 + ph + categoryLabelOffsetPx(5, chart.catAxisLabelOffsetPercent),
        rotRad,
      );
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none' && dateAxisPlan) {
      for (const tick of dateAxisPlan.minorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMinorTickMark, 'cat', py0 + ph,
          px0 + tick.fraction * pw, undefined, undefined,
          false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
        );
      }
    }
  }

  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r,
      px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  if (dataTableLayout) {
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, py0 + ph, pw, x + legLeftW, ptToPx,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Surface / contour chart (ECMA-376 §21.2.2.204)
// ═══════════════════════════════════════════════════════════════════════════

interface SurfaceVertex extends ThreeDScenePoint { value: number }

const surfaceCellTriangleIndices = (
  values: readonly [number, number, number, number],
): readonly [readonly [number, number, number], readonly [number, number, number]] =>
  values[0] + values[2] > values[1] + values[3]
    ? [[0, 1, 2], [0, 2, 3]]
    : [[0, 1, 3], [1, 2, 3]];

const MAX_SURFACE_PAINT_POLYGONS = 200_000;
const SURFACE_SCENE_DEPTH_SCALE = 1.25;

function clipSurfacePolygon(
  polygon: SurfaceVertex[],
  threshold: number,
  keepAbove: boolean,
): SurfaceVertex[] {
  if (polygon.length === 0) return [];
  const output: SurfaceVertex[] = [];
  const inside = (vertex: SurfaceVertex): boolean =>
    keepAbove ? vertex.value >= threshold : vertex.value <= threshold;
  let previous = polygon[polygon.length - 1];
  let previousInside = inside(previous);
  for (const current of polygon) {
    const currentInside = inside(current);
    if (currentInside !== previousInside) {
      const denominator = current.value - previous.value;
      const fraction = denominator === 0 ? 0 : (threshold - previous.value) / denominator;
      output.push({
        x: previous.x + (current.x - previous.x) * fraction,
        y: previous.y + (current.y - previous.y) * fraction,
        depth: previous.depth + (current.depth - previous.depth) * fraction,
        value: threshold,
      });
    }
    if (currentInside) output.push(current);
    previous = current;
    previousInside = currentInside;
  }
  return output;
}

function renderSurfaceChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const categories = chartCategories(chart);
  const rows = chart.series;
  const columnCount = categories.length;
  const rowCount = rows.length;
  if (columnCount < 2 || rowCount < 2) return;

  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const row of rows) {
    for (let column = 0; column < columnCount; column++) {
      const value = row.values[column];
      if (value == null || !Number.isFinite(value)) continue;
      dataMin = Math.min(dataMin, value);
      dataMax = Math.max(dataMax, value);
    }
  }
  if (!Number.isFinite(dataMin) || !Number.isFinite(dataMax)) return;
  if (chart.valMin != null) dataMin = chart.valMin;
  if (chart.valMax != null) dataMax = chart.valMax;

  // Excel supplies the standard oblique perspective values for omitted
  // classic-Surface view fields. S1-S5 carry no c:view3D and isolate that
  // effective camera: parallel source-grid edges converge in the Office
  // output, so substituting a right-angle/orthographic camera loses the visible
  // depth. Keep these field defaults local to the Surface family; every
  // authored view3D field remains authoritative.
  const surfaceView = {
    ...(chart.threeD ?? {}),
    rotationX: chart.threeD?.rotationX ?? 15,
    rotationY: chart.threeD?.rotationY ?? 20,
    rightAngleAxes: chart.threeD?.rightAngleAxes ?? false,
    perspective: chart.threeD?.perspective ?? 30,
  };
  const observedAutomaticSurfaceCamera = isObservedAutomaticSurfaceCamera(surfaceView);
  const perspectiveTangentGain = surfacePerspectiveTangentGain(surfaceView);

  // Office bases an omitted Surface major unit on the projected value-axis
  // length. Probe the shared camera before reserving the legend: this is
  // independent of the eventual band count and correctly degenerates to the
  // compact five-interval class for a 90-degree contour view.
  const axisProbe = planChartThreeDProjection(surfaceView, r, {
    sceneDepthScale: SURFACE_SCENE_DEPTH_SCALE,
    perspectiveTangentGain,
  });
  let projectedValueAxisLenPt: number | undefined;
  if (axisProbe) {
    const probeX = axisProbe.topology.axisX === 'min'
      ? axisProbe.front.x : axisProbe.front.x + axisProbe.front.w;
    const probeBottom = axisProbe.project(
      probeX, axisProbe.front.y + axisProbe.front.h, axisProbe.topology.nearDepth,
    );
    const probeTop = axisProbe.project(
      probeX, axisProbe.front.y, axisProbe.topology.nearDepth,
    );
    projectedValueAxisLenPt = Math.hypot(
      probeTop.x - probeBottom.x,
      probeTop.y - probeBottom.y,
    ) / ptToPx;
  }
  const automaticSurfaceUnit = chart.valAxisMajorUnit == null
    ? automaticSurfaceMajorUnit(dataMin, dataMax, projectedValueAxisLenPt)
    : null;
  const provisionalAxis = planValueAxis(
    automaticSurfaceUnit == null
      ? chart
      : { ...chart, valAxisMajorUnit: automaticSurfaceUnit },
    dataMin,
    dataMax,
    projectedValueAxisLenPt,
  );
  // Surface value bands follow the shared value-axis plan. OOXML does not
  // define an Office-compatible automatic band count or implicit lighting, so
  // the renderer deliberately does not invent either here.
  const step = provisionalAxis.step;
  if (!(step > 0) || !Number.isFinite(step)) return;
  // Surface bands terminate on the first major boundary containing the data;
  // unlike ordinary value axes, Office does not append one headroom interval
  // when the maximum already lands on a boundary (S1/S4/S5 and their scaled
  // axis mirrors). Authored bounds remain authoritative.
  const surfaceMin = chart.valMin ?? provisionalAxis.min;
  const surfaceMax = chart.valMax ?? Math.max(
    surfaceMin + step,
    surfaceMin + Math.ceil((dataMax - surfaceMin) / step) * step,
  );
  const surfaceSpan = surfaceMax - surfaceMin;
  if (!(surfaceSpan > 0) || !Number.isFinite(surfaceSpan)) return;
  const rawBandCount = Math.ceil(surfaceSpan / step);
  const rawMajorLineCount = Math.floor(surfaceSpan / step + 1e-9) + 1;
  const triangleCount = (columnCount - 1) * (rowCount - 1) * 2;
  if (
    !Number.isSafeInteger(rawBandCount)
    || !Number.isSafeInteger(rawMajorLineCount)
    || rawBandCount < 1
    || rawMajorLineCount < 2
    || triangleCount < 1
    || rawBandCount > MAX_AXIS_TICKS
    || rawMajorLineCount > MAX_AXIS_TICKS
    || rawBandCount > Math.floor(MAX_SURFACE_PAINT_POLYGONS / triangleCount)
    || rawMajorLineCount > MAX_SURFACE_PAINT_POLYGONS
  ) return;
  const bandCount = rawBandCount;
  const surfaceFrac = valAxisReversed(chart)
    ? (value: number): number => 1 - (value - surfaceMin) / surfaceSpan
    : (value: number): number => (value - surfaceMin) / surfaceSpan;
  const surfaceMajorLines = Array.from(
    { length: rawMajorLineCount },
    (_, index) => surfaceMin + index * step,
  );
  const wireframeSplitFractions = (start: number, end: number): number[] => {
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    const firstBoundary = Math.max(
      1,
      Math.floor((low - surfaceMin) / step) + 1,
    );
    const lastBoundary = Math.min(
      bandCount - 1,
      Math.ceil((high - surfaceMin) / step) - 1,
    );
    const fractions = [0];
    if (start !== end) {
      for (let boundary = firstBoundary; boundary <= lastBoundary; boundary++) {
        fractions.push((surfaceMin + boundary * step - start) / (end - start));
      }
    }
    fractions.push(1);
    fractions.sort((left, right) => left - right);
    return fractions;
  };
  const bandColors = Array.from({ length: bandCount }, (_, index) =>
    legacyPattern2Color(
      chart.themeAccentColors ?? [],
      index,
      bandCount,
      // Office's omitted classic chart style uses Pattern 2. Keep this
      // compatibility default scoped to renderer-generated surface bands;
      // authored series/point paint remains parser-owned.
      chart.legacyChartStyle ?? 2,
    ) ?? (rows[index]?.color ? `#${rows[index].color}` : chartColor(index, rows[index])),
  );
  const bandFormats = new Map((chart.surfaceBandFormats ?? []).map(format => [format.idx, format]));
  const linkedBandStyle = chart.chartStyleRoles?.dataPoint3D;
  const linkedWireframeStyle = chart.chartStyleRoles?.dataPointWireframe;
  const styleHasLinePaint = (style: ChartSeries['chartexStyle']): boolean =>
    style?.lineNoStyle !== true && (
      style?.linePaintAuthored === true
      || style?.lineHidden === true
      || (style?.lineColors?.length ?? 0) > 0
      || (style?.linePaints?.length ?? 0) > 0
    );
  const bandFillRecipes = chart.surfaceWireframe === true ? [] : Array.from(
    { length: bandCount }, (_, index) => {
      const format = bandFormats.get(index);
      let decision: Fill | null | undefined;
      if (format?.fillHidden === true) decision = null;
      else if (format?.fill) decision = format.fill;
      else decision = chartStyleFillDecision(format?.style, index);
      return decision === undefined
        ? chartStyleFillDecision(linkedBandStyle, index)
        : decision;
    },
  );
  const bandLineRecipes = chart.surfaceWireframe === true ? [] : Array.from(
    { length: bandCount }, (_, index) => {
      const format = bandFormats.get(index);
      let decision: ChartModel['plotAreaLineFill'] | null | undefined;
      if (format?.lineHidden === true) decision = null;
      else if (format?.lineColor) decision = { fillType: 'solid', color: format.lineColor };
      else decision = chartStyleLineDecision(format?.style, index);
      return decision === undefined
        ? chartStyleLineDecision(linkedBandStyle, index)
        : decision;
    },
  );
  interface SurfaceWireframeLineStyle {
    paint: ChartModel['plotAreaLineFill'] | null | undefined;
    lineWidthEmu: number | null | undefined;
    lineDash: string | null | undefined;
    lineCustomDash: ChartModel['plotAreaLineCustomDash'];
    lineCap: string | null | undefined;
    lineJoin: string | null | undefined;
    lineCompound: string | null | undefined;
  }
  const styleGeometry = (
    direct: ChartSeries['chartexStyle'],
    fallback: SurfaceWireframeLineStyle,
  ): Omit<SurfaceWireframeLineStyle, 'paint'> => {
    const effectiveDirect = direct?.lineNoStyle === true ? undefined : direct;
    const dashAuthored = effectiveDirect?.lineDashAuthored === true
      || effectiveDirect?.lineDash != null
      || effectiveDirect?.lineCustomDash != null;
    return {
      lineWidthEmu: effectiveDirect?.lineWidthEmu ?? fallback.lineWidthEmu,
      lineDash: dashAuthored ? effectiveDirect?.lineDash : fallback.lineDash,
      lineCustomDash: dashAuthored
        ? effectiveDirect?.lineCustomDash : fallback.lineCustomDash,
      lineCap: effectiveDirect?.lineCap ?? fallback.lineCap,
      lineJoin: effectiveDirect?.lineJoin ?? fallback.lineJoin,
      lineCompound: effectiveDirect?.lineCompound ?? fallback.lineCompound,
    };
  };
  // Office's Surface wireframe uses a fixed dataPointWireframe reference as
  // the mesh default. A relative palette has no single chart-wide index, so it
  // remains fail-closed rather than silently choosing entry zero.
  const fixedWireframeLineIndex = linkedWireframeStyle?.lineColorIndex;
  let linkedWireframePaint: ChartModel['plotAreaLineFill'] | null | undefined;
  if (!styleHasLinePaint(linkedWireframeStyle)) linkedWireframePaint = undefined;
  else if (linkedWireframeStyle?.lineHidden === true) {
    linkedWireframePaint = chartStyleLineDecision(linkedWireframeStyle, 0);
  } else {
    linkedWireframePaint = fixedWireframeLineIndex != null
      ? chartStyleLineDecision(linkedWireframeStyle, fixedWireframeLineIndex)
      : null;
  }
  const emptyWireframeStyle: SurfaceWireframeLineStyle = {
    paint: undefined,
    lineWidthEmu: undefined,
    lineDash: undefined,
    lineCustomDash: undefined,
    lineCap: undefined,
    lineJoin: undefined,
    lineCompound: undefined,
  };
  const linkedGeometry = linkedWireframeStyle?.lineNoStyle === true
    ? emptyWireframeStyle
    : {
      paint: linkedWireframePaint,
      lineWidthEmu: linkedWireframeStyle?.lineWidthEmu,
      lineDash: linkedWireframeStyle?.lineDash,
      lineCustomDash: linkedWireframeStyle?.lineCustomDash,
      lineCap: linkedWireframeStyle?.lineCap,
      lineJoin: linkedWireframeStyle?.lineJoin,
      lineCompound: linkedWireframeStyle?.lineCompound,
    };
  const firstSurfaceSeries = rows[0];
  let directSeriesPaint: ChartModel['plotAreaLineFill'] | null | undefined;
  if (firstSurfaceSeries?.lineHidden === true) directSeriesPaint = null;
  else if (firstSurfaceSeries?.lineColor != null) {
    directSeriesPaint = { fillType: 'solid', color: firstSurfaceSeries.lineColor };
  } else {
    directSeriesPaint = chartStyleLineDecision(firstSurfaceSeries?.chartexStyle, 0);
  }
  const baseGeometry = styleGeometry(firstSurfaceSeries?.chartexStyle, linkedGeometry);
  const baseWireframeStyle: SurfaceWireframeLineStyle = {
    paint: directSeriesPaint === undefined ? linkedWireframePaint : directSeriesPaint,
    ...baseGeometry,
    lineWidthEmu: firstSurfaceSeries?.lineWidthEmu ?? baseGeometry.lineWidthEmu,
  };
  if (baseWireframeStyle.lineCompound != null) baseWireframeStyle.paint = null;
  const directBandLineDecisions = Array.from({ length: bandCount }, (_, index) => {
    const format = bandFormats.get(index);
    if (!format) return undefined;
    if (format.lineHidden === true) return null;
    if (format.lineColor != null) return { fillType: 'solid' as const, color: format.lineColor };
    return chartStyleLineDecision(format.style, index);
  });
  const wireframeLineStyles = directBandLineDecisions.map((directPaint, index) => {
    const format = bandFormats.get(index);
    const geometry = styleGeometry(format?.style, baseWireframeStyle);
    const style: SurfaceWireframeLineStyle = {
      paint: directPaint === undefined ? baseWireframeStyle.paint : directPaint,
      ...geometry,
      lineWidthEmu: format?.lineWidthEmu ?? geometry.lineWidthEmu,
    };
    if (style.lineCompound != null) style.paint = null;
    return style;
  });
  const sameWireframeStyle = (
    left: SurfaceWireframeLineStyle,
    right: SurfaceWireframeLineStyle,
  ): boolean => left.paint !== undefined
    && left.paint === right.paint
    && left.lineWidthEmu === right.lineWidthEmu
    && left.lineDash === right.lineDash
    && left.lineCustomDash === right.lineCustomDash
    && left.lineCap === right.lineCap
    && left.lineJoin === right.lineJoin
    && left.lineCompound === right.lineCompound;
  interface SurfaceWireframeBandRun {
    from: number;
    to: number;
    band: number;
  }
  const wireframeBandRuns = (start: number, end: number): SurfaceWireframeBandRun[] => {
    const fractions = wireframeSplitFractions(start, end);
    const runs: SurfaceWireframeBandRun[] = [];
    for (let index = 0; index < fractions.length - 1; index++) {
      const from = fractions[index];
      const to = fractions[index + 1];
      const midpoint = start + (end - start) * ((from + to) / 2);
      const band = Math.max(0, Math.min(
        bandCount - 1,
        Math.floor((midpoint - surfaceMin) / step),
      ));
      const previous = runs[runs.length - 1];
      if (previous
        && directBandLineDecisions[previous.band] === undefined
        && directBandLineDecisions[band] === undefined
        && sameWireframeStyle(
          wireframeLineStyles[previous.band], wireframeLineStyles[band],
        )) previous.to = to;
      else runs.push({ from, to, band });
    }
    return runs;
  };
  if (chart.surfaceWireframe === true) {
    let wireframeSegmentCount = 0;
    const chargeEdge = (start: number | null, end: number | null): boolean => {
      if (start == null || end == null || !Number.isFinite(start) || !Number.isFinite(end)) {
        return true;
      }
      wireframeSegmentCount += wireframeBandRuns(start, end).length;
      return wireframeSegmentCount <= MAX_SURFACE_PAINT_POLYGONS;
    };
    for (let row = 0; row < rowCount; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        if (!chargeEdge(rows[row].values[column], rows[row].values[column + 1])) return;
      }
    }
    for (let column = 0; column < columnCount; column++) {
      for (let row = 0; row < rowCount - 1; row++) {
        if (!chargeEdge(rows[row].values[column], rows[row + 1].values[column])) return;
      }
    }
    // A wireframe Surface contains both the source row/column mesh and the
    // contour at each value-band boundary. Charge the latter before any
    // projection or paint allocation, using the same cell triangulation as
    // the renderer below.
    for (let row = 0; row < rowCount - 1; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        const values = [
          rows[row].values[column],
          rows[row].values[column + 1],
          rows[row + 1].values[column + 1],
          rows[row + 1].values[column],
        ];
        if (values.some(value => value == null || !Number.isFinite(value))) continue;
        const finiteValues = values as [number, number, number, number];
        for (const indices of surfaceCellTriangleIndices(finiteValues)) {
          const triangleMin = Math.min(...indices.map(index => finiteValues[index]));
          const triangleMax = Math.max(...indices.map(index => finiteValues[index]));
          const firstBoundary = Math.max(
            1,
            Math.floor((triangleMin - surfaceMin) / step) + 1,
          );
          const lastBoundary = Math.min(
            bandCount - 1,
            Math.ceil((triangleMax - surfaceMin) / step) - 1,
          );
          if (lastBoundary < firstBoundary) continue;
          wireframeSegmentCount += lastBoundary - firstBoundary + 1;
          if (wireframeSegmentCount > MAX_SURFACE_PAINT_POLYGONS) return;
        }
      }
    }
  }
  const usesBaseWireframeLine = directBandLineDecisions.some(
    decision => decision === undefined,
  );
  const surfaceFacePaints = [
    { surface: chart.threeD?.floor, role: 'floor' as const },
    { surface: chart.threeD?.sideWall, role: 'wall' as const },
    { surface: chart.threeD?.backWall, role: 'wall' as const },
  ].map(({ surface, role }) => chartThreeDSurfacePaint(chart, surface, role));
  let surfacePaintComponents = 0;
  for (const recipe of [
    ...bandFillRecipes,
    ...bandLineRecipes,
    ...(chart.surfaceWireframe === true
      && usesBaseWireframeLine
      ? [baseWireframeStyle.paint]
      : []),
    ...(chart.surfaceWireframe === true
      ? directBandLineDecisions.filter(decision => decision !== undefined)
      : []),
    ...surfaceFacePaints.flatMap(paint => [paint.fill, paint.line]),
  ]) {
    if (recipe == null) continue;
    const components = markerPaintComponents(recipe);
    if ((recipe.fillType === 'gradient'
        && components > MAX_CANVAS_MARKER_GRADIENT_STOPS)
      || components > MAX_CANVAS_MARKER_PAINT_COMPONENTS - surfacePaintComponents) return;
    surfacePaintComponents += components;
  }
  const bandLabels = Array.from({ length: bandCount }, (_, index) => {
    const lower = surfaceMin + index * step;
    const upper = Math.min(surfaceMax, lower + step);
    return `${formatChartVal(lower)}-${formatChartVal(upper)}`;
  });
  const legendChart: ChartModel = {
    ...chart,
    series: bandLabels.map((name, index) => ({
      name,
      color: bandColors[index].replace(/^#/, ''),
      values: [],
    })),
  };
  // A top-down contour lists the highest band first. In the ordinary oblique
  // Surface view Office lists low-to-high. S1-S5 and the 90° contour boundary
  // isolate this to the authored/effective camera elevation, not legend side.
  if (Math.abs(surfaceView.rotationX) === 90) legendChart.series.reverse();
  const legend = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    legend, chart.legendOverlay === true,
  );
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const seriesAxis = chart.threeD?.seriesAxis;
  const seriesFontPx = chartTextFontSizePx(seriesAxis?.fontSizeHpt, ptToPx) ?? catFontPx;
  const pad = {
    t: titleBand.bandH + legTopH + seriesFontPx / 2,
    r: legRightW + seriesFontPx * 3.2 + 12,
    b: catAxisLabelBandH(catFontPx, chart.catAxisLabelOffsetPercent) + legBottomH,
    l: legLeftW + catFontPx * 1.5,
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: legend,
    pad,
    honorPlotAreaManualLayout: true,
  });
  const { px0, py0, pw, ph } = frame.plotRect;
  if (!(pw > 0) || !(ph > 0)) return;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleBand.topPad, titleBand.fontPx,
  );
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  let projection = planChartThreeDProjection(surfaceView, { x: px0, y: py0, w: pw, h: ph }, {
    // Surface rows occupy a real series axis rather than the compact prism
    // slab used by 3-D columns. The boundary corpus fits that grid with the
    // same depth occupancy as standard (depth-arranged) cartesian series.
    sceneDepthScale: SURFACE_SCENE_DEPTH_SCALE,
    perspectiveTangentGain,
  });
  if (!projection) return;
  projection = fitChartThreeDProjectionToWallThickness(
    projection,
    chart.threeD ?? {},
    { x: px0, y: py0, w: pw, h: ph },
  );
  const useObservedAutomaticMaterial = observedAutomaticSurfaceCamera || (
    Math.abs(surfaceView.rotationX) === 90
    && surfaceView.rotationY === 0
    && surfaceView.rightAngleAxes === false
    && surfaceView.perspective === 0
  );
  const { front } = projection;
  const categoryReversed = chart.catAxisOrientation === 'maxMin';
  const categoryBetween = isCrossBetween(chart);
  const toX = (column: number): number => front.x + categoryPositionFraction(
    column, columnCount, categoryBetween, categoryReversed,
  ) * front.w;
  const seriesReversed = seriesAxis?.orientation === 'maxMin';
  // Surface rows are values on `<c:serAx>`, whose schema has no
  // `<c:crossBetween>`. Office places the first and last rows on the series
  // axis endpoints, unlike category points centred by value-axis
  // crossBetween="between". Reuse the shared endpoint placement so reversal
  // and the single-row midpoint remain consistent with other ordinal axes.
  const toDepth = (row: number): number => categoryPositionFraction(
    row, rowCount, false, seriesReversed,
  );
  const toValueY = (value: number): number =>
    front.y + front.h - surfaceFrac(value) * front.h;
  interface SurfacePaint {
    points: Array<{ x: number; y: number }>;
    scenePoints: SurfaceVertex[];
    band: number;
    depth: number;
  }
  const paints: SurfacePaint[] = [];
  interface SurfaceWireframeSegment {
    points: [{ x: number; y: number }, { x: number; y: number }];
    band: number;
  }
  const wireframeSegments: SurfaceWireframeSegment[] = [];
  const appendWireframeEdge = (start: SurfaceVertex, end: SurfaceVertex): void => {
    const pointAt = (fraction: number): SurfaceVertex => ({
      x: start.x + (end.x - start.x) * fraction,
      y: start.y + (end.y - start.y) * fraction,
      depth: start.depth + (end.depth - start.depth) * fraction,
      value: start.value + (end.value - start.value) * fraction,
    });
    for (const run of wireframeBandRuns(start.value, end.value)) {
      const from = pointAt(run.from);
      const to = pointAt(run.to);
      wireframeSegments.push({
        points: [
          projection.project(from.x, from.y, from.depth),
          projection.project(to.x, to.y, to.depth),
        ],
        band: run.band,
      });
    }
  };
  const appendWireframeContours = (triangle: readonly SurfaceVertex[]): void => {
    const triangleMin = Math.min(...triangle.map(vertex => vertex.value));
    const triangleMax = Math.max(...triangle.map(vertex => vertex.value));
    const firstBoundary = Math.max(
      1,
      Math.floor((triangleMin - surfaceMin) / step) + 1,
    );
    const lastBoundary = Math.min(
      bandCount - 1,
      Math.ceil((triangleMax - surfaceMin) / step) - 1,
    );
    for (let boundary = firstBoundary; boundary <= lastBoundary; boundary++) {
      const threshold = surfaceMin + boundary * step;
      const intersections: SurfaceVertex[] = [];
      const addIntersection = (vertex: SurfaceVertex): void => {
        if (intersections.some(existing =>
          Math.abs(existing.x - vertex.x) < 1e-9
          && Math.abs(existing.y - vertex.y) < 1e-9
          && Math.abs(existing.depth - vertex.depth) < 1e-9
        )) return;
        intersections.push(vertex);
      };
      for (let index = 0; index < triangle.length; index++) {
        const start = triangle[index];
        const end = triangle[(index + 1) % triangle.length];
        if (start.value === threshold) addIntersection(start);
        if ((start.value < threshold && end.value > threshold)
          || (start.value > threshold && end.value < threshold)) {
          const fraction = (threshold - start.value) / (end.value - start.value);
          addIntersection({
            x: start.x + (end.x - start.x) * fraction,
            y: start.y + (end.y - start.y) * fraction,
            depth: start.depth + (end.depth - start.depth) * fraction,
            value: threshold,
          });
        }
      }
      if (intersections.length !== 2) continue;
      wireframeSegments.push({
        points: [
          projection.project(
            intersections[0].x, intersections[0].y, intersections[0].depth,
          ),
          projection.project(
            intersections[1].x, intersections[1].y, intersections[1].depth,
          ),
        ],
        // A boundary closes the band below it. This keeps c:bandFmt indexing
        // consistent with the lower-inclusive clipping used by filled Surface.
        band: boundary - 1,
      });
    }
  };
  const paintTriangle = (triangle: SurfaceVertex[]): void => {
    if (chart.surfaceWireframe === true) {
      appendWireframeContours(triangle);
      return;
    }
    const triangleMin = Math.min(...triangle.map(vertex => vertex.value));
    const triangleMax = Math.max(...triangle.map(vertex => vertex.value));
    const firstBand = Math.max(0, Math.floor((triangleMin - surfaceMin) / step));
    const lastBand = Math.min(
      bandCount - 1,
      Math.floor((triangleMax - surfaceMin) / step),
    );
    for (let band = firstBand; band <= lastBand; band++) {
      const lower = surfaceMin + band * step;
      const upper = band === bandCount - 1 ? surfaceMax : lower + step;
      let polygon = clipSurfacePolygon(triangle, lower, true);
      polygon = clipSurfacePolygon(polygon, upper, false);
      if (polygon.length < 3) continue;
      const points = polygon.map(vertex => projection.project(vertex.x, vertex.y, vertex.depth));
      paints.push({
        points,
        scenePoints: polygon,
        band,
        depth: polygon.reduce(
          (sum, vertex) => sum + projection.cameraDepth(vertex.x, vertex.y, vertex.depth),
          0,
        ) / polygon.length,
      });
    }
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(px0, py0, pw, ph);
  ctx.clip();

  const strokeScenePath = (
    scenePoints: readonly ThreeDScenePoint[],
    color: string,
    width: number,
    dash: string | null | undefined,
    unbounded = false,
  ): void => {
    if (scenePoints.length < 2) return;
    const points = unbounded
      ? scenePoints.map(point => projection.projectUnbounded(point.x, point.y, point.depth))
      : scenePoints.map(point => projection.project(point.x, point.y, point.depth));
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    for (let index = 1; index < points.length; index++) ctx.lineTo(points[index].x, points[index].y);
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.setLineDash(pptxPresetDashArray(dash ?? 'solid', width));
    ctx.stroke();
  };
  const farDepth = projection.topology.farDepth;
  const nearDepth = projection.topology.nearDepth;
  const floorY = front.y + front.h;
  const wallTopY = front.y;
  const farX = projection.topology.farX === 'min' ? front.x : front.x + front.w;
  const surfaceSlabs = (
    ['floor', 'sideWall', 'backWall'] as const
  ).map(kind => {
    const surface = chart.threeD?.[kind];
    return planChartThreeDSurfaceGeometry(projection, kind, surface?.thicknessPercent);
  });
  const surfaceKinds = ['floor', 'sideWall', 'backWall'] as const;
  const strokeSurfaceGridRule = (
    slabIndex: number,
    coordinate: 'x' | 'y',
    fraction: number,
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    const slab = surfaceSlabs[slabIndex];
    const kind = surfaceKinds[slabIndex];
    for (const segment of planChartThreeDSurfaceGridSegments(
      slab,
      kind,
      coordinate,
      fraction,
    )) {
      if (slab.thickness > 0
        && !projection.cameraFacing(slab.faces[segment.faceIndex])) continue;
      strokeScenePath(segment.scenePoints, color, width, dash, true);
    }
  };
  const strokeAuthoredValueSurfaceRules = (
    values: readonly number[],
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    for (const value of values) {
      const fraction = surfaceFrac(value);
      strokeSurfaceGridRule(1, 'y', fraction, color, width, dash);
      strokeSurfaceGridRule(2, 'y', fraction, color, width, dash);
    }
  };
  const strokeAuthoredCategorySurfaceRules = (
    fractions: readonly number[],
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    for (const fraction of fractions) {
      strokeSurfaceGridRule(0, 'x', fraction, color, width, dash);
      strokeSurfaceGridRule(2, 'x', fraction, color, width, dash);
    }
  };
  // Keep the pre-thickness Surface3D path byte-stable when all three values
  // are omitted/zero. Its existing floor plane is family-owned; positive
  // thickness opts into the shared camera-aware CT_Surface slabs.
  const surfaceFaceGroups = surfaceSlabs.some(slab => slab.thickness > 0)
    ? surfaceSlabs.map(slab => slab.faces
      .filter(face => slab.thickness === 0 || projection.cameraFacing(face))
      .map(face => face.map(point =>
        projection.projectUnbounded(point.x, point.y, point.depth)
      )))
    : [
      [
        projection.project(front.x, floorY, nearDepth),
        projection.project(front.x + front.w, floorY, nearDepth),
        projection.project(front.x + front.w, floorY, farDepth),
        projection.project(front.x, floorY, farDepth),
      ],
      [
        projection.project(farX, floorY, nearDepth),
        projection.project(farX, floorY, farDepth),
        projection.project(farX, wallTopY, farDepth),
        projection.project(farX, wallTopY, nearDepth),
      ],
      [
        projection.project(front.x, floorY, farDepth),
        projection.project(front.x + front.w, floorY, farDepth),
        projection.project(front.x + front.w, wallTopY, farDepth),
        projection.project(front.x, wallTopY, farDepth),
      ],
    ].map(face => [face]);
  for (let index = 0; index < surfaceFaceGroups.length; index++) {
    const faces = surfaceFaceGroups[index];
    if (!faces.length) continue;
    const points = faces.flat();
    const effective = surfaceFacePaints[index];
    const imageFill = effective.fill?.fillType === 'image' ? effective.fill : null;
    if (imageFill) {
      const image = chartImageFillSource(imageFill);
      const surface = chart.threeD?.[surfaceKinds[index]];
      if (image) {
        const project = (point: ThreeDScenePoint) =>
          projection.projectUnbounded(point.x, point.y, point.depth);
        paintChartThreeDSurfacePicture(
          ctx, imageFill, image, surface, surfaceKinds[index],
          surfaceSlabs[index], surfaceSlabs[index].faces
            .map((face, faceIndex) => ({ face, faceIndex }))
            .filter(({ face }) => surfaceSlabs[index].thickness === 0
              || projection.cameraFacing(face))
            .map(({ faceIndex }) => faceIndex),
          project, surfaceSpan,
        );
      }
    }
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const fill = imageFill
      ? null
      : effective.fill?.fillType === 'solid'
      ? `#${effective.fill.color}`
      : effective.fill
        ? resolveFill(effective.fill, ctx, minX, minY, maxX - minX, maxY - minY)
        : null;
    const line = effective.line?.fillType === 'solid'
      ? `#${effective.line.color}`
      : effective.line
        ? resolveFill(effective.line, ctx, minX, minY, maxX - minX, maxY - minY)
        : null;
    const width = effective.lineWidthEmu != null
      ? axisLineWidthPx(effective.lineWidthEmu, ptToPx) : 1;
    for (const face of faces) {
      ctx.beginPath();
      ctx.moveTo(face[0].x, face[0].y);
      for (let pointIndex = 1; pointIndex < face.length; pointIndex++) {
        ctx.lineTo(face[pointIndex].x, face[pointIndex].y);
      }
      ctx.closePath();
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fill();
      }
      if (line) {
        ctx.strokeStyle = line;
        ctx.lineWidth = width;
        ctx.setLineDash(drawingmlLineDashArray(
          effective.lineCustomDash,
          effective.lineDash,
          width,
        ));
        ctx.lineCap = effective.lineCap === 'rnd'
          ? 'round' : effective.lineCap === 'sq' ? 'square' : 'butt';
        ctx.lineJoin = effective.lineJoin === 'round' || effective.lineJoin === 'bevel'
          ? effective.lineJoin : 'miter';
        ctx.stroke();
      }
    }
  }
  if (chart.valAxisMinorGridlines === true) {
    const line = valMinorGridStroke(chart, ptToPx);
    strokeAuthoredValueSurfaceRules(
      provisionalAxis.minorLines.filter(value => value >= surfaceMin && value <= surfaceMax),
      line.color,
      line.width,
      chart.valAxisMinorGridlineDash,
    );
  }
  if (drawValMajorGridlines(chart)) {
    const line = resolveGridline(
      chart.valAxisGridlineColor,
      chart.valAxisGridlineWidthEmu,
      ptToPx,
    );
    if (chart.valAxisMajorGridlines === true) {
      strokeAuthoredValueSurfaceRules(
        surfaceMajorLines,
        line.color,
        line.width,
        chart.valAxisGridlineDash,
      );
    } else {
      for (const value of surfaceMajorLines) {
        const fraction = surfaceFrac(value);
        const gridY = toValueY(value);
        if (surfaceSlabs[2].thickness > 0) {
          strokeSurfaceGridRule(
            2, 'y', fraction, line.color, line.width, chart.valAxisGridlineDash,
          );
        } else {
          strokeScenePath([
            { x: front.x, y: gridY, depth: farDepth },
            { x: front.x + front.w, y: gridY, depth: farDepth },
          ], line.color, line.width, chart.valAxisGridlineDash);
        }
        if (surfaceSlabs[1].thickness > 0) {
          strokeSurfaceGridRule(
            1, 'y', fraction, line.color, line.width, chart.valAxisGridlineDash,
          );
        } else {
          strokeScenePath([
            { x: farX, y: gridY, depth: nearDepth },
            { x: farX, y: gridY, depth: farDepth },
          ], line.color, line.width, chart.valAxisGridlineDash);
        }
      }
    }
  }
  if (chart.catAxisMinorGridlines === true) {
    const line = catMinorGridStroke(chart, ptToPx);
    strokeAuthoredCategorySurfaceRules(
      categoryMinorGridlineFractions(columnCount, categoryBetween),
      line.color,
      line.width,
      chart.catAxisMinorGridlineDash,
    );
  }
  if (chart.catAxisMajorGridlines) {
    const line = resolveGridline(
      chart.catAxisGridlineColor,
      chart.catAxisGridlineWidthEmu,
      ptToPx,
    );
    // Gridlines mark category boundaries under crossBetween="between"; data
    // points remain at the interval centres. Using `toX(column)` here would
    // incorrectly draw lines through the 25%/75% data points of a two-column
    // Surface instead of the 0%/50%/100% boundaries authored by the axis.
    strokeAuthoredCategorySurfaceRules(
      catGridlineFractions(chart, columnCount),
      line.color,
      line.width,
      chart.catAxisGridlineDash,
    );
  }

  for (let row = 0; row < rowCount - 1; row++) {
    for (let column = 0; column < columnCount - 1; column++) {
      const values = [
        rows[row].values[column],
        rows[row].values[column + 1],
        rows[row + 1].values[column + 1],
        rows[row + 1].values[column],
      ];
      if (values.some(value => value == null || !Number.isFinite(value))) continue;
      const rowColumn = {
        x: toX(column), y: toValueY(values[0] as number), depth: toDepth(row),
        value: values[0] as number,
      };
      const rowColumnNext = {
        x: toX(column + 1), y: toValueY(values[1] as number), depth: toDepth(row),
        value: values[1] as number,
      };
      const rowNextColumnNext = {
        x: toX(column + 1), y: toValueY(values[2] as number), depth: toDepth(row + 1),
        value: values[2] as number,
      };
      const rowNextColumn = {
        x: toX(column), y: toValueY(values[3] as number), depth: toDepth(row + 1),
        value: values[3] as number,
      };
      // Office's filled Surface keeps the upper of the two possible cell
      // diagonals. The 2x2 plane/saddle/reversed-saddle boundaries isolate all
      // three cases: equal opposing sums retain the source-grid B-D diagonal;
      // otherwise the opposing pair with the larger value sum forms the ridge.
      // This is an application rendering rule — OOXML stores only the matrix.
      const vertices = [
        rowColumn, rowColumnNext, rowNextColumnNext, rowNextColumn,
      ] as const;
      for (const indices of surfaceCellTriangleIndices(values as [number, number, number, number])) {
        paintTriangle(indices.map(index => vertices[index]));
      }
    }
  }
  if (chart.surfaceWireframe === true) {
    for (let row = 0; row < rowCount; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        const startValue = rows[row].values[column];
        const endValue = rows[row].values[column + 1];
        if (startValue == null || endValue == null
          || !Number.isFinite(startValue) || !Number.isFinite(endValue)) continue;
        appendWireframeEdge(
          {
            x: toX(column), y: toValueY(startValue), depth: toDepth(row), value: startValue,
          },
          {
            x: toX(column + 1), y: toValueY(endValue), depth: toDepth(row), value: endValue,
          },
        );
      }
    }
    for (let column = 0; column < columnCount; column++) {
      for (let row = 0; row < rowCount - 1; row++) {
        const startValue = rows[row].values[column];
        const endValue = rows[row + 1].values[column];
        if (startValue == null || endValue == null
          || !Number.isFinite(startValue) || !Number.isFinite(endValue)) continue;
        appendWireframeEdge(
          {
            x: toX(column), y: toValueY(startValue), depth: toDepth(row), value: startValue,
          },
          {
            x: toX(column), y: toValueY(endValue), depth: toDepth(row + 1), value: endValue,
          },
        );
      }
    }
  }
  paints.sort((left, right) => left.depth - right.depth);
  const bandBounds = Array.from({ length: bandCount }, () => ({
    minX: Number.POSITIVE_INFINITY,
    minY: Number.POSITIVE_INFINITY,
    maxX: Number.NEGATIVE_INFINITY,
    maxY: Number.NEGATIVE_INFINITY,
  }));
  for (const paint of paints) {
    const bounds = bandBounds[paint.band];
    for (const point of paint.points) {
      bounds.minX = Math.min(bounds.minX, point.x);
      bounds.minY = Math.min(bounds.minY, point.y);
      bounds.maxX = Math.max(bounds.maxX, point.x);
      bounds.maxY = Math.max(bounds.maxY, point.y);
    }
  }
  for (const segment of wireframeSegments) {
    const bounds = bandBounds[segment.band];
    for (const point of segment.points) {
      bounds.minX = Math.min(bounds.minX, point.x);
      bounds.minY = Math.min(bounds.minY, point.y);
      bounds.maxX = Math.max(bounds.maxX, point.x);
      bounds.maxY = Math.max(bounds.maxY, point.y);
    }
  }
  type SurfaceCanvasPaint = string | CanvasGradient | CanvasPattern | null | undefined;
  const resolveBandPaint = (
    recipe: Fill | null | undefined,
    band: number,
  ): SurfaceCanvasPaint => {
    if (recipe == null) return recipe;
    if (recipe.fillType === 'solid') return `#${recipe.color}`;
    const bounds = bandBounds[band];
    if (!Number.isFinite(bounds.minX) || !Number.isFinite(bounds.minY)
      || !Number.isFinite(bounds.maxX) || !Number.isFinite(bounds.maxY)) return null;
    return resolveFill(
      recipe,
      ctx,
      bounds.minX,
      bounds.minY,
      bounds.maxX - bounds.minX,
      bounds.maxY - bounds.minY,
    );
  };
  // A c:bandFmt styles one logical band. Resolve its DrawingML recipes once
  // against the complete projected band bounds, then reuse the Canvas paint
  // for every clipped polygon instead of replaying gradient stops per face.
  const resolvedBandFills = bandFillRecipes.map(resolveBandPaint);
  const resolvedBandLines = bandLineRecipes.map(resolveBandPaint);
  for (const paint of paints) {
    const format = bandFormats.get(paint.band);
    ctx.beginPath();
    ctx.moveTo(paint.points[0].x, paint.points[0].y);
    for (let index = 1; index < paint.points.length; index++) {
      ctx.lineTo(paint.points[index].x, paint.points[index].y);
    }
    ctx.closePath();
    const bandFill = resolvedBandFills[paint.band];
    if (bandFill !== null) {
      ctx.fillStyle = bandFill ?? scaleHexColor(
        bandColors[paint.band],
        useObservedAutomaticMaterial
          ? surfaceMaterialFactor(projection.cameraNormal(paint.scenePoints))
          : 1,
      );
      ctx.fill();
    }
    const bandLine = resolvedBandLines[paint.band];
    if (bandLine != null) {
      const directGeometry = format?.style;
      const linkedGeometry = linkedBandStyle?.lineNoStyle === true
        ? undefined : linkedBandStyle;
      ctx.strokeStyle = bandLine;
      const widthEmu = format?.lineWidthEmu
        ?? directGeometry?.lineWidthEmu ?? linkedGeometry?.lineWidthEmu;
      ctx.lineWidth = widthEmu != null
        ? axisLineWidthPx(widthEmu, ptToPx)
        : 1;
      ctx.setLineDash(drawingmlLineDashArray(
        directGeometry?.lineCustomDash ?? linkedGeometry?.lineCustomDash,
        directGeometry?.lineDash ?? linkedGeometry?.lineDash,
        ctx.lineWidth,
      ));
      const cap = directGeometry?.lineCap ?? linkedGeometry?.lineCap;
      const join = directGeometry?.lineJoin ?? linkedGeometry?.lineJoin;
      ctx.lineCap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = join === 'round' || join === 'bevel' ? join : 'miter';
      ctx.stroke();
    }
  }
  if (chart.surfaceWireframe === true) {
    const baseWireframeLine = !usesBaseWireframeLine
      ? undefined
      : baseWireframeStyle.paint?.fillType === 'solid'
      ? `#${baseWireframeStyle.paint.color}`
      : baseWireframeStyle.paint
        ? resolveFill(baseWireframeStyle.paint, ctx, px0, py0, pw, ph)
        : baseWireframeStyle.paint;
    const resolvedWireframeLines = directBandLineDecisions.map((decision, band) =>
      decision === undefined ? baseWireframeLine : resolveBandPaint(decision, band)
    );
    for (const segment of wireframeSegments) {
      const style = wireframeLineStyles[segment.band];
      const line = resolvedWireframeLines[segment.band];
      if (line === null) continue;
      ctx.beginPath();
      ctx.moveTo(segment.points[0].x, segment.points[0].y);
      ctx.lineTo(segment.points[1].x, segment.points[1].y);
      ctx.strokeStyle = line ?? bandColors[segment.band];
      ctx.lineWidth = style.lineWidthEmu != null
        ? axisLineWidthPx(style.lineWidthEmu, ptToPx)
        : Math.max(1, 0.75 * ptToPx);
      ctx.setLineDash(drawingmlLineDashArray(
        style.lineCustomDash,
        style.lineDash,
        ctx.lineWidth,
      ));
      const cap = style.lineCap;
      const join = style.lineJoin;
      ctx.lineCap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = join === 'round' || join === 'bevel' ? join : 'miter';
      ctx.stroke();
    }
  }
  ctx.restore();

  const sceneCenter = projection.project(
    front.x + front.w / 2,
    front.y + front.h / 2,
    0.5,
  );
  const drawSurfaceAxisTick = (
    mode: string | null | undefined,
    point: { x: number; y: number },
    axisStart: { x: number; y: number },
    axisEnd: { x: number; y: number },
    color: string,
    lineWidth: number,
    lineHidden: boolean,
    level: 'major' | 'minor',
    dash?: string | null,
  ): void => {
    if (lineHidden || mode == null || mode === 'none') return;
    const dx = axisEnd.x - axisStart.x;
    const dy = axisEnd.y - axisStart.y;
    const axisLength = Math.hypot(dx, dy);
    if (!(axisLength > 1e-6)) return;
    let normalX = -dy / axisLength;
    let normalY = dx / axisLength;
    const midpointX = (axisStart.x + axisEnd.x) / 2;
    const midpointY = (axisStart.y + axisEnd.y) / 2;
    if ((midpointX - sceneCenter.x) * normalX
      + (midpointY - sceneCenter.y) * normalY < 0) {
      normalX = -normalX;
      normalY = -normalY;
    }
    const length = axisTickLengthPx(level, lineWidth, ptToPx);
    const sideLength = mode === 'cross' ? length / 2 : length;
    const outer = mode === 'out' || mode === 'cross' ? sideLength : 0;
    const inner = mode === 'in' || mode === 'cross' ? sideLength : 0;
    strokeAxisSegment(
      ctx,
      point.x + normalX * outer,
      point.y + normalY * outer,
      point.x - normalX * inner,
      point.y - normalY * inner,
      color,
      lineWidth,
      dash,
    );
  };

  const categoryAxisStart = projection.project(front.x, floorY, nearDepth);
  const categoryAxisEnd = projection.project(front.x + front.w, floorY, nearDepth);
  const surfaceCatLineWidth = chart.catAxisLineWidthEmu != null
    ? axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx)
    : 1;
  strokeAxisSegment(
    ctx,
    categoryAxisStart.x,
    categoryAxisStart.y,
    categoryAxisEnd.x,
    categoryAxisEnd.y,
    chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#000000',
    surfaceCatLineWidth,
    chart.catAxisLineDash,
  );
  const categoryAxisColor = chart.catAxisLineColor
    ? `#${chart.catAxisLineColor}` : '#000000';
  const categoryAxisSuppressed = chart.catAxisHidden || chart.catAxisLineHidden === true;
  const categoryTickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
  for (let column = 0; column < columnCount; column += categoryTickSkip) {
    drawSurfaceAxisTick(
      chart.catAxisMajorTickMark,
      projection.project(toX(column), floorY, nearDepth),
      categoryAxisStart,
      categoryAxisEnd,
      categoryAxisColor,
      surfaceCatLineWidth,
      categoryAxisSuppressed,
      'major',
      chart.catAxisLineDash,
    );
  }
  if (chart.catAxisMinorTickMark != null && chart.catAxisMinorTickMark !== 'none') {
    for (let column = 0; column < columnCount - 1; column++) {
      const fraction = (
        categoryPositionFraction(column, columnCount, categoryBetween, categoryReversed)
        + categoryPositionFraction(column + 1, columnCount, categoryBetween, categoryReversed)
      ) / 2;
      drawSurfaceAxisTick(
        chart.catAxisMinorTickMark,
        projection.project(front.x + fraction * front.w, floorY, nearDepth),
        categoryAxisStart,
        categoryAxisEnd,
        categoryAxisColor,
        surfaceCatLineWidth,
        categoryAxisSuppressed,
        'minor',
        chart.catAxisLineDash,
      );
    }
  }
  ctx.font = chartFontCss(
    catFontPx,
    chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#000000';
  ctx.textBaseline = 'top';
  for (let column = 0; column < columnCount; column++) {
    const anchor = categoryLabelAnchorFraction(
      column,
      columnCount,
      isCrossBetween(chart),
      catAxisReversed(chart),
      chart.catAxisLabelAlignment,
    );
    const point = projection.project(front.x + anchor.fraction * front.w, floorY, nearDepth);
    ctx.textAlign = anchor.textAlign;
    ctx.fillText(
      categories[column] ?? '',
      point.x,
      point.y + categoryLabelOffsetPx(8, chart.catAxisLabelOffsetPercent),
    );
  }

  if (!seriesAxis?.hidden) {
    const seriesAxisWidth = seriesAxis?.lineWidthEmu != null
      ? axisLineWidthPx(seriesAxis.lineWidthEmu, ptToPx)
      : 1;
    const seriesAxisMinPoint = projection.project(front.x, floorY, 0.5);
    const seriesAxisMaxPoint = projection.project(front.x + front.w, floorY, 0.5);
    const seriesAxisX = seriesAxisMinPoint.x >= seriesAxisMaxPoint.x
      ? front.x : front.x + front.w;
    const seriesStart = projection.project(seriesAxisX, floorY, nearDepth);
    const seriesEnd = projection.project(seriesAxisX, floorY, farDepth);
    strokeAxisSegment(
      ctx, seriesStart.x, seriesStart.y, seriesEnd.x, seriesEnd.y,
      seriesAxis?.lineColor ? `#${seriesAxis.lineColor}` : '#000000',
      seriesAxisWidth, seriesAxis?.lineDash,
    );
    const seriesAxisColor = seriesAxis?.lineColor
      ? `#${seriesAxis.lineColor}` : '#000000';
    const seriesTickSkip = Math.max(1, Math.floor(seriesAxis?.tickMarkSkip ?? 1));
    for (let row = 0; row < rowCount; row += seriesTickSkip) {
      drawSurfaceAxisTick(
        seriesAxis?.majorTickMark,
        projection.project(seriesAxisX, floorY, toDepth(row)),
        seriesStart,
        seriesEnd,
        seriesAxisColor,
        seriesAxisWidth,
        seriesAxis?.lineHidden === true,
        'major',
        seriesAxis?.lineDash,
      );
    }
    if (seriesAxis?.minorTickMark != null && seriesAxis.minorTickMark !== 'none') {
      for (let row = 0; row < rowCount - 1; row++) {
        drawSurfaceAxisTick(
          seriesAxis.minorTickMark,
          projection.project(seriesAxisX, floorY, (toDepth(row) + toDepth(row + 1)) / 2),
          seriesStart,
          seriesEnd,
          seriesAxisColor,
          seriesAxisWidth,
          seriesAxis.lineHidden === true,
          'minor',
          seriesAxis.lineDash,
        );
      }
    }
    ctx.font = chartFontCss(
      seriesFontPx,
      chartFontFamily(chart, seriesAxis?.fontFace, 'minor'),
      seriesAxis?.fontBold ?? false,
      seriesAxis?.fontItalic ?? false,
    );
    ctx.fillStyle = seriesAxis?.fontColor ? `#${seriesAxis.fontColor}` : '#000000';
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
    for (let row = 0; row < rowCount; row++) {
      const point = projection.project(seriesAxisX, floorY, toDepth(row));
      ctx.fillText(rows[row].name, point.x + 8, point.y);
    }
  }

  if (!chart.valAxisHidden) {
    const valueAxisX = projection.topology.axisX === 'min' ? front.x : front.x + front.w;
    const valueAxisBottom = projection.project(valueAxisX, front.y + front.h, nearDepth);
    const valueAxisTop = projection.project(valueAxisX, front.y, nearDepth);
    if (Math.hypot(valueAxisTop.x - valueAxisBottom.x, valueAxisTop.y - valueAxisBottom.y) > 4) {
      const valueAxisWidth = chart.valAxisLineWidthEmu != null
        ? axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx)
        : 1;
      strokeAxisSegment(
        ctx, valueAxisBottom.x, valueAxisBottom.y, valueAxisTop.x, valueAxisTop.y,
        chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#000000',
        valueAxisWidth, chart.valAxisLineDash,
      );
      const valFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
      ctx.font = chartFontCss(
        valFontPx,
        chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
        chart.valAxisFontBold ?? false,
        chart.valAxisFontItalic ?? false,
      );
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#000000';
      const left = (valueAxisBottom.x + valueAxisTop.x) / 2 < px0 + pw / 2;
      ctx.textAlign = left ? 'right' : 'left';
      ctx.textBaseline = 'middle';
      for (const value of surfaceMajorLines) {
        const point = projection.project(valueAxisX, toValueY(value), nearDepth);
        ctx.fillText(
          formatChartValWithCode(value, chart.valAxisFormatCode, chart.date1904),
          point.x + (left ? -6 : 6),
          point.y,
        );
      }
    }
  }
  drawLegendForLayout(
    ctx,
    legendChart,
    legend,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    titleBand.bandH + 2,
    ptToPx,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  // A plot area can contain both `<c:areaChart>` and `<c:lineChart>` groups.
  // Only the ordered area-group series participate in the filled stack; line
  // series share the axes but remain independent overlays (§21.2.2.145).
  const areaSeries = chart.series
    .map((series, chartIndex) => ({ series, chartIndex }))
    .filter(({ series }) => series.seriesType == null || series.seriesType === 'area');
  const lineSeries = chart.series
    .map((series, chartIndex) => ({ series, chartIndex }))
    .filter(({ series }) => series.seriesType === 'line');
  if (areaSeries.length === 0 && lineSeries.length === 0) return;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const legacyGrouping = chart.chartType === 'stackedAreaPct'
    ? 'percentStacked' : chart.chartType === 'stackedArea' ? 'stacked' : 'standard';
  const sourceAreaGroups = chart.plotGroups?.filter(group => group.kind === 'area') ?? [{
    kind: 'area' as const,
    seriesStart: 0,
    seriesCount: areaSeries.length,
    categoryAxis: 'primary' as const,
    valueAxis: 'primary' as const,
    seriesAxis: 'none' as const,
    grouping: legacyGrouping,
  }];
  const areaIndexAtChartIndex = new Array<number>(chart.series.length).fill(-1);
  for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
    areaIndexAtChartIndex[areaSeries[areaIndex].chartIndex] = areaIndex;
  }
  const areaGroupMembers = sourceAreaGroups.map(group => {
    const areaIndices: number[] = [];
    const end = Math.min(chart.series.length, group.seriesStart + group.seriesCount);
    for (let chartIndex = group.seriesStart; chartIndex < end; chartIndex++) {
      const areaIndex = areaIndexAtChartIndex[chartIndex];
      if (areaIndex >= 0) areaIndices.push(areaIndex);
    }
    return { group, areaIndices };
  });
  const stackedByArea = new Array<boolean>(areaSeries.length).fill(false);
  const percentByArea = new Array<boolean>(areaSeries.length).fill(false);
  const percentTotalsByArea = new Array<number[] | null>(areaSeries.length).fill(null);
  const areaBaseValues = areaSeries.map(() => new Array<number>(n).fill(0));
  const areaTopValues = areaSeries.map(() => new Array<number>(n).fill(0));
  const axisPlanningGroups = chart.plotGroups?.filter(group =>
    (group.kind === 'area' || group.kind === 'line') && group.seriesCount > 0
  ) ?? sourceAreaGroups;
  const allPercentByAreaAxis = new Map<string, boolean>();
  for (const group of axisPlanningGroups) {
    const axis = group.valueAxis;
    allPercentByAreaAxis.set(
      axis,
      (allPercentByAreaAxis.get(axis) ?? true) && group.grouping === 'percentStacked',
    );
  }
  for (const { group, areaIndices } of areaGroupMembers) {
    const grouping = group.grouping ?? 'standard';
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const pct = grouping === 'percentStacked';
    const multiplier = pct && allPercentByAreaAxis.get(group.valueAxis) === true ? 100 : 1;
    const totals = pct
      ? cats.map((_, categoryIndex) => areaIndices.reduce(
          (sum, areaIndex) => sum + Math.abs(
            areaSeries[areaIndex].series.values[categoryIndex] ?? 0,
          ), 0,
        ) || 1)
      : null;
    for (const areaIndex of areaIndices) {
      stackedByArea[areaIndex] = stacked;
      percentByArea[areaIndex] = pct;
      percentTotalsByArea[areaIndex] = totals;
    }
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      let positive = 0;
      let negative = 0;
      for (const areaIndex of areaIndices) {
        const raw = areaSeries[areaIndex].series.values[categoryIndex] ?? 0;
        const contribution = pct && totals ? raw / totals[categoryIndex] * multiplier : raw;
        const base = contribution >= 0 ? positive : negative;
        areaBaseValues[areaIndex][categoryIndex] = stacked ? base : 0;
        areaTopValues[areaIndex][categoryIndex] = stacked ? base + contribution : contribution;
        if (stacked) {
          if (contribution >= 0) positive += contribution;
          else negative += contribution;
        }
      }
    }
  }
  const primaryAxisGroups = axisPlanningGroups.filter(group => group.valueAxis !== 'secondary');
  const axisIsPercent = primaryAxisGroups.length > 0
    && primaryAxisGroups.every(group => group.grouping === 'percentStacked');

  // Combo area charts may bind some series to a SECONDARY value axis on the
  // right (ECMA-376 §21.2.2.*). As with line, this applies only to plain
  // (unstacked) area — a stacked/percentStacked secondary combo is not an Office
  // construct. `sec` is null (single-axis, byte-identical to pre-CH7) unless the
  // axis is declared AND a series opts in; secondary series are then excluded
  // from the primary extent and mapped through the secondary scale.
  const areaIndexBySeries = new Map(areaSeries.map((entry, index) => [entry.series, index]));
  const seriesChartIndex = new Map(chart.series.map((series, index) => [series, index]));
  const seriesUsesSecondary = (series: ChartSeries): boolean => {
    const chartIndex = seriesChartIndex.get(series) ?? -1;
    return plotGroupBySeries[chartIndex]?.valueAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  };
  const secondaryAxisGroups = axisPlanningGroups.filter(group => group.valueAxis === 'secondary');
  const secondaryAxisIsPercent = secondaryAxisGroups.length > 0
    && secondaryAxisGroups.every(group => group.grouping === 'percentStacked');
  const sec = chart.secondaryValAxis && chart.series.some(series => seriesUsesSecondary(series))
    ? chart.secondaryValAxis
    : null;
  const isSecondarySeries = (series: ChartSeries): boolean => {
    return sec != null && seriesUsesSecondary(series);
  };

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  // Vertical pads first so the estimated plot height is known before the
  // secondary-axis scale + right-gutter measurement (same ordering as bar/line).
  // Top: title band + half a value-axis label above the top gridline. Bottom:
  // PowerPoint's category-label band (gap + line-height + margin).
  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  const secScale = computeSecondaryAxis(
    sec,
    chart.series,
    phEst / ptToPx,
    'y',
    secondaryAxisIsPercent,
    false,
    series => seriesUsesSecondary(series),
    (series, pointIndex) => {
      const areaIndex = areaIndexBySeries.get(series);
      if (areaIndex == null) return series.values[pointIndex] ?? null;
      return series.values[pointIndex] == null
        ? null : areaTopValues[areaIndex][pointIndex] ?? null;
    },
  );
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = `${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    let wmax = 0;
    for (const value of secScale.majorLines) {
      wmax = Math.max(wmax, ctx.measureText(formatAxisTickWithUnits(value, sec.formatCode ?? null, chart.date1904, sec.displayUnits)).width);
    }
    secLabelBandW = wmax + 18;
    ctx.font = prevFont;
  }
  const secTitleBandW = sec && sec.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  // Resolve the primary extent before frame placement so an authored
  // `layoutTarget="outer"` can be converted to the inner data rectangle using
  // the actual formatted tick-label width. The outer rectangle includes axis
  // labels and ticks (ECMA-376 §21.2.2.89); treating its left edge as `px0`
  // pushes the labels outside chart space.
  const computeAreaDataExtent = (): { min: number; max: number } => {
    let min = Infinity;
    let max = -Infinity;
    for (let ci = 0; ci < n; ci++) {
      for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
        const { series } = areaSeries[areaIndex];
        if (isSecondarySeries(series) || series.values[ci] == null) continue;
        min = Math.min(min, areaBaseValues[areaIndex][ci], areaTopValues[areaIndex][ci]);
        max = Math.max(max, areaBaseValues[areaIndex][ci], areaTopValues[areaIndex][ci]);
      }
      for (const { series } of lineSeries) {
        if (isSecondarySeries(series)) continue;
        const value = series.values[ci];
        if (value == null) continue;
        min = Math.min(min, value);
        max = Math.max(max, value);
      }
    }
    if (!isFinite(min) || !isFinite(max)) return { min: 0, max: 1 };
    if (axisIsPercent) return { min: min < 0 ? -100 : 0, max: max > 0 ? 100 : 0 };
    return { min, max };
  };
  let areaExtent = computeAreaDataExtent();
  if (!axisIsPercent) {
    const includeEndpoint = (value: number): void => {
      areaExtent = {
        min: Math.min(areaExtent.min, value),
        max: Math.max(areaExtent.max, value),
      };
    };
    for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
      const { series } = areaSeries[areaIndex];
      if (isSecondarySeries(series)) continue;
      forEachErrorBarEndpoint(
        series,
        'y',
        index => {
          if (series.values[index] == null) return null;
          return areaTopValues[areaIndex][index];
        },
        includeEndpoint,
      );
    }
    for (const { series } of lineSeries) {
      if (isSecondarySeries(series)) continue;
      forEachErrorBarEndpoint(series, 'y', index => series.values[index] ?? null, includeEndpoint);
    }
  }
  const provisionalScale = planValueAxis(
    chart,
    areaExtent.min,
    areaExtent.max,
    phEst / ptToPx,
    axisIsPercent,
  );
  const manualValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxFontPx
    : Math.max(8, Math.min(11, phEst / 20));
  let primaryLabelWidth = 0;
  if (
    !chart.valAxisHidden
    && chart.plotAreaManualLayout != null
    && chart.plotAreaManualLayout.layoutTarget !== 'inner'
  ) {
    const prevFont = ctx.font;
    ctx.font = chartFontCss(
      manualValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    for (const value of provisionalScale.majorLines) {
      primaryLabelWidth = Math.max(
        primaryLabelWidth,
        ctx.measureText(formatPrimaryValueAxisTick(chart, value, axisIsPercent)).width,
      );
    }
    ctx.font = prevFont;
  }
  const manualOuterInsets = chartManualOuterAxisInsets({
    valAxisHidden: chart.valAxisHidden,
    catAxisHidden: chart.catAxisHidden,
    valLabelWidth: primaryLabelWidth,
    valLabelFontPx: manualValTickFontPx,
    catLabelFontPx: catAxFontPx,
    valLabelGapPx: chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(manualValTickFontPx)
      : 6,
    catLabelGapPx: chart.catAxisFontSizeHpt != null
      ? categoryLabelOffsetPx(
        categoryTickLabelGapPx(catAxFontPx),
        chart.catAxisLabelOffsetPercent,
      )
      : categoryLabelOffsetPx(3, chart.catAxisLabelOffsetPercent),
    outerTextMarginPx: AXIS_OUTER_TEXT_MARGIN_PT * ptToPx,
    valTitleBandW: valTitleW,
    catTitleBandH: catTitleH,
    secondaryBandW: secLabelBandW + secTitleBandW,
  });

  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: legLeftW + Math.max(w * 0.12 + valTitleW, dataTableHeaderW),
  };

  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  const areaFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  const { px0, py0, pw } = areaFrame.plotRect;
  let { ph } = areaFrame.plotRect;
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // Primary extent from the PRIMARY series only (secondary series live on their
  // own axis). When `sec` is null every series is primary, byte-identical to
  // the pre-CH7 path.
  // Value axis is vertical → length = plot height (axis-length-aware auto major unit). An
  // explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the auto step.
  const areaPlan = planValueAxis(
    chart, areaExtent.min, areaExtent.max, ph / ptToPx, axisIsPercent,
  );

  // crossBetween="between" (Office's default; ECMA-376 §21.2.2.32 leaves the
  // default application-defined) gives each category a band of width pw/n and
  // plots its point at the band CENTER, leaving a half-band margin before the
  // first and after the last category — matching PowerPoint's Jan…Dec inset.
  // "midCat" anchors points on the category dividers (flush to the axes).
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  const toX = dateAxisPlan
    ? (index: number) => px0 + dateAxisPlan.positions[index]! * pw
    : between
      ? (index: number) => {
        const i = catRev ? n - 1 - index : index;
        return px0 + ((i + 0.5) / n) * pw;
      }
      : (index: number) => {
        const i = catRev ? n - 1 - index : index;
        return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
      };
  const toY = (v: number) => py0 + ph - areaPlan.frac(v) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;
  const primaryCategoryAxisY = toY(
    categoryAxisCrossingValue(chart, areaPlan.min, areaPlan.max),
  );
  const secondaryCategoryAxisY = sec && secScale
    ? toYSecondary(axisCrossingValue(
      chart.secondaryCatAxis?.crossesAt,
      chart.secondaryCatAxis?.crosses,
      secScale.min,
      secScale.max,
    ))
    : primaryCategoryAxisY;
  const categoryAxisYFor = (series: ChartSeries): number =>
    isSecondarySeries(series) ? secondaryCategoryAxisY : primaryCategoryAxisY;

  // Axis line colour/weight from `<c:*Ax><c:spPr><a:ln>` (EMU → px at scale),
  // mirroring the bar/line renderers. Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);

  // Value-axis MAJOR gridlines are drawn UNDER the series (before the fills), so
  // an opaque/translucent area occludes the gridlines inside its region —
  // matching Office vector observations in which opaque area fill occludes
  // gridlines below its top edge. This mirrors the bar/line/stock/
  // scatter/waterfall/box renderers, which already stroke gridlines first. The
  // axis rules, tick marks and value/category labels stay AFTER the series (drawn
  // further below) so they sit atop the plot. `<c:valAx><c:majorGridlines>` is on
  // by default (`drawValMajorGridlines`); `<c:minorGridlines>` only when declared.
  if (!chart.valAxisHidden) {
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    // Minor gridlines (`<c:valAx><c:minorGridlines>`, §21.2.2.129) drawn first,
    // UNDER the majors and the series when the file declares them. An omitted
    // minor unit uses the shared automatic major/5 fallback.
    if (chart.valAxisMinorGridlines) {
      for (const v of areaPlan.minorLines) {
        strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
      }
    }
    if (drawValMajorGridlines(chart)) {
      for (const v of areaPlan.majorLines) {
        strokeValueGridlineH(ctx, px0, pw, toY(v), v === 0, grid);
      }
    }
  }
  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }
  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks, also under the fills. Off by default
  // (byte-stable when the file omits them).
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    const previousDash = cg.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (cg.dash.length > 0) ctx.setLineDash(cg.dash);
    const fractions = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
      : catGridlineFractions(chart, n);
    for (const frac of fractions) {
      const gx = px0 + frac * pw;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (cg.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // Draw the series area fills ON TOP of the gridlines laid down above.
  // In a stacked area chart, series order is the stacking order: series 0 is
  // adjacent to the category axis, then series 1, and so on (CT_AreaChart's
  // ordered `ser` sequence). Standard areas use the same document paint order:
  // a later series is painted later and therefore overlays earlier series,
  // matching Excel's vector output at their intersections.
  const seriesOrder = areaGroupMembers.flatMap(({ areaIndices }) => areaIndices);
  const plottedAreaValue = (areaIndex: number, categoryIndex: number): number =>
    areaTopValues[areaIndex]?.[categoryIndex] ?? 0;
  for (const areaIndex of seriesOrder) {
    const { series: s, chartIndex } = areaSeries[areaIndex];
    const color = chartColor(chartIndex, s);
    const baseY = py0 + ph;
    // Unstacked secondary series ride their own vertical scale; the stacked
    // branch is never reached with a secondary axis (`sec` is null when
    // stacked), so its `toY` mapping stays the primary one.
    const yOf = yMapFor(s);

    // Smooth (`<c:ser><c:smooth>`, §21.2.2.194) curves the top edge through the
    // points; the baseline connection stays straight. Non-smooth keeps the exact
    // prior moveTo/lineTo sequence (byte-stable) — appendCurve with smooth=false
    // emits identical lineTo calls.
    //
    // NB: `CT_AreaSer` (§A.5.1) has no `<c:smooth>` child (only `CT_LineSer` /
    // `CT_ScatterSer` do), so `extract_series_smooth` never sets `s.smooth` for
    // a real area series and this branch is dead against actual chart XML —
    // it only fires for a model constructed directly (tests / other producers).
    // Kept for symmetry with the line renderer above rather than dropped.
    const smooth = s.smooth === true;
    ctx.beginPath();
    if (stackedByArea[areaIndex]) {
      const topPts = [];
      for (let ci = 0; ci < n; ci++) {
        topPts.push({ x: toX(ci), y: toY(areaTopValues[areaIndex][ci]) });
      }
      ctx.moveTo(topPts[0].x, topPts[0].y);
      appendCurve(ctx, topPts, smooth);
      for (let ci = n - 1; ci >= 0; ci--) {
        ctx.lineTo(toX(ci), toY(areaBaseValues[areaIndex][ci]));
      }
    } else {
      const topPts = [];
      for (let ci = 0; ci < n; ci++) topPts.push({ x: toX(ci), y: yOf(s.values[ci] ?? 0) });
      ctx.moveTo(toX(0), baseY);
      ctx.lineTo(topPts[0].x, topPts[0].y);
      appendCurve(ctx, topPts, smooth);
      ctx.lineTo(toX(n - 1), baseY);
    }
    ctx.closePath();
    // `<a:solidFill>` is opaque unless the DrawingML color itself carries an
    // alpha transform. The shared model currently carries an opaque resolved
    // hex, so do not invent translucency for area series.
    ctx.fillStyle = color;
    ctx.fill();
    if (s.lineHidden !== true) {
      ctx.strokeStyle = s.lineColor ? `#${s.lineColor}` : color;
      ctx.lineWidth = s.lineWidthEmu ? axisLineWidthPx(s.lineWidthEmu, ptToPx) : 1.5;
      ctx.setLineDash([]);
      ctx.stroke();
    }
  }

  // `CT_AreaChart` includes `dropLines` through `EG_AreaChartShared`
  // (ECMA-376 Part 1, dml-chart.xsd). Office vector output establishes one
  // drop line per category, spanning the extrema of the category-axis crossing
  // and every plotted point in the owning group. This matters for a standard
  // multi-series area chart (one envelope line, not one line per series) and
  // for an interior crossing (the line spans points on both sides). Paint after
  // the opaque area fills so the authored geometry remains visible, but before
  // point markers and labels.
  const decorationAreaGroupMembers = new Map<number, Array<{ series: ChartSeries; areaIndex: number }>>();
  for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
    const series = areaSeries[areaIndex].series;
    const groupIndex = series.areaGroupIndex ?? 0;
    const members = decorationAreaGroupMembers.get(groupIndex) ?? [];
    members.push({ series, areaIndex });
    decorationAreaGroupMembers.set(groupIndex, members);
  }
  for (const decoration of chart.areaGroupDecorations ?? []) {
    if (!decoration.dropLines) {
      continue;
    }
    const dropLineStyle = chartStyleRoleLine(chart, decoration.dropLines, 'dropLine');
    if (!applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      continue;
    }
    const members = decorationAreaGroupMembers.get(decoration.groupIndex) ?? [];
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      let minY = Infinity;
      let maxY = -Infinity;
      let hasPoint = false;
      for (const member of members) {
        if (member.series.values[categoryIndex] == null) continue;
        const pointY = yMapFor(member.series)(
          plottedAreaValue(member.areaIndex, categoryIndex),
        );
        const axisY = categoryAxisYFor(member.series);
        if (!Number.isFinite(pointY) || !Number.isFinite(axisY)) continue;
        minY = Math.min(minY, pointY, axisY);
        maxY = Math.max(maxY, pointY, axisY);
        hasPoint = true;
      }
      if (!hasPoint || Math.abs(maxY - minY) < 0.01) continue;
      ctx.beginPath();
      ctx.moveTo(toX(categoryIndex), minY);
      ctx.lineTo(toX(categoryIndex), maxY);
      ctx.stroke();
    }
  }

  // Markers, error bars, and per-point data labels for area series. Drawn in a
  // SEPARATE forward pass (after all fills) so the fill loop above stays
  // byte-identical, and each block fires ONLY for series carrying the relevant
  // fields — an area chart with no marker/errBar/dLbl detail draws exactly as
  // before. The plotted top-of-band value matches where the fill's top edge sat
  // (cumulative for stacked). ECMA-376 §21.2.2.32 / §21.2.2.20 / §21.2.2.45.
  //
  // NB: an area chart's filled region has always read a blank cell as 0
  // (`?? 0`), so `<c:dispBlanksAs>` (§21.2.2.42) is a no-op for the area family
  // here — breaking or spanning a *filled* region is not modeled, and changing
  // the default would break byte-stability. dispBlanksAs steers the line family
  // (where "gap" is the historical default).
  {
    const areaMarkerR = Math.max(2, 2.5 * ptToPx);
    // Top of each series' band per category (stacked); the raw value otherwise.
    // Rebuilt independently of the fill loop's mutated stackBase. The ordered
    // series sequence stacks forward, so band si reaches Σ_{k=0..si}.
    for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
      const { series: s, chartIndex } = areaSeries[areaIndex];
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      const color = chartColor(chartIndex, s);
      const yOf = yMapFor(s);
      const plottedOf = (ci: number): number => plottedAreaValue(areaIndex, ci);
      const seriesPercentTotals = percentTotalsByArea[areaIndex];
      // Error bars first (markers overlay their tips).
      for (const eb of s.errBars ?? []) {
        drawCategoryErrorBars(
          ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, color,
        );
      }
      // Markers only when the series opts in (`<c:marker>` symbol/size/… — area
      // charts default to NO markers, so nothing fires without explicit detail).
      const seriesMarkersVisible = (s.showMarker === true || seriesHasMarkerDetail(s))
        && s.markerSymbol !== 'none';
      if (seriesMarkersVisible || hasVisiblePointMarkerOverride(s)) {
        for (let ci = 0; ci < n; ci++) {
          if (s.values[ci] == null) continue;
          const dpt = pointOverrides.get(ci);
          const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
          if (symbol === 'none') continue;
          const px = toX(ci); const py = yOf(plottedOf(ci));
          if (seriesHasMarkerDetail(s) || pointHasMarkerDetail(dpt)) {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = markerFillColorFor(s, dpt, ci, color);
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            const lineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawMarker(
              ctx, px, py, symbol, sizePt, fill, line, ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
            );
          } else {
            ctx.fillStyle = color;
            ctx.beginPath(); ctx.arc(px, py, areaMarkerR, 0, Math.PI * 2); ctx.fill();
          }
        }
      }
      // Per-point / series-level data labels. Area's filled region has always
      // read a blank cell as 0 (`?? 0`, see the topValue/plottedOf comment
      // above), so every category index is a "plotted" point here regardless
      // of dispBlanksAs — pass true unconditionally (byte-stable: unchanged
      // from before this parameter existed).
      drawCategoryDataLabels(
        ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false, true,
        chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
        // §21.2.2.48 `<c:dLblPos>` precedence: chart-level position, else the
        // area-chart default `'ctr'` (centered on the point, ECMA-376 default
        // for the areaChart group).
        chart.dataLabelPosition ?? 'ctr',
        { x: px0, y: py0, w: pw, h: ph },
        { x, y, w, h },
        percentByArea[areaIndex] && seriesPercentTotals
          ? ci => (s.values[ci] ?? 0) / seriesPercentTotals[ci]
          : undefined,
        undefined,
        face => chartFontFamily(chart, face, 'minor'),
        isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
        ci => dataLabelLegendKey(chartIndex, ci),
        value => dataLabelWithinAxisMaximum(
          chart, value,
          isSecondarySeries(s) && secScale ? secScale.max : areaPlan.max,
        ),
        shapeRotationDeg,
      );
    }
  }

  // Paint `<c:lineChart>` groups after the area fills, using the same category
  // and value-axis transforms. They do not alter the area stack. This is the
  // OOXML combo-chart z-order: later chart groups overlay earlier ones.
  for (const { series: s, chartIndex } of lineSeries) {
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const color = chartColor(chartIndex, s);
    const stroke = s.lineColor ? `#${s.lineColor}` : color;
    const yOf = yMapFor(s);
    if (s.lineHidden !== true) {
      ctx.strokeStyle = stroke;
      ctx.lineWidth = s.lineWidthEmu ? axisLineWidthPx(s.lineWidthEmu, ptToPx) : Math.max(1, 2.25 * ptToPx);
      ctx.setLineDash([]);
      ctx.beginPath();
      let run: Array<{ x: number; y: number }> = [];
      const flushRun = (): void => {
        if (run.length === 0) return;
        ctx.moveTo(run[0].x, run[0].y);
        appendCurve(ctx, run, s.smooth === true);
        run = [];
      };
      for (let ci = 0; ci < n; ci++) {
        const value = s.values[ci];
        if (value == null) {
          if ((chart.dispBlanksAs ?? 'gap') === 'gap') flushRun();
          if ((chart.dispBlanksAs ?? 'gap') !== 'zero') continue;
        }
        run.push({ x: toX(ci), y: yOf(value ?? 0) });
      }
      flushRun();
      ctx.stroke();
    }

    const plottedOf = (ci: number): number => s.values[ci] ?? 0;
    for (const eb of s.errBars ?? []) {
      drawCategoryErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, stroke,
      );
    }
    const seriesMarkersVisible = (s.showMarker === true || seriesHasMarkerDetail(s))
      && s.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(s)) {
      for (let ci = 0; ci < n; ci++) {
        const value = s.values[ci];
        if (value == null) continue;
        const dpt = pointOverrides.get(ci);
        const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        drawMarker(
          ctx, toX(ci), yOf(value), symbol, dpt?.markerSize ?? s.markerSize ?? 5,
          markerFillColorFor(s, dpt, ci, stroke),
          dpt?.markerLine ?? s.markerLine ?? null, ptToPx,
          (dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu) != null
            ? axisLineWidthPx((dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu) as number, ptToPx)
            : undefined,
          markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
        );
      }
    }
    drawCategoryDataLabels(
      ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false, false,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'), chart.dataLabelPosition ?? 'r',
      { x: px0, y: py0, w: pw, h: ph },
      { x, y, w, h },
      undefined,
      undefined,
      face => chartFontFamily(chart, face, 'minor'),
      isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
      ci => dataLabelLegendKey(chartIndex, ci),
      value => dataLabelWithinAxisMaximum(
        chart, value,
        isSecondarySeries(s) && secScale ? secScale.max : areaPlan.max,
      ),
      shapeRotationDeg,
    );
    drawSeriesTrendlines(
      ctx, s, stroke, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  // Value-axis tick marks + labels. The gridlines themselves were already laid
  // down UNDER the series (above the fill loop); here we only add the tick marks
  // and the value labels, which belong ON TOP of the plot.
  if (!chart.valAxisHidden) {
    const drawnValTickFontPx = chart.valAxisFontSizeHpt != null
      ? valAxFontPx
      : Math.max(8, Math.min(11, ph / 20));
    ctx.font = chartFontCss(
      drawnValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    for (const v of areaPlan.majorLines) {
      const gy = toY(v);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
      ctx.textAlign = 'right';
      const gap = chart.valAxisFontSizeHpt != null
        ? valueTickLabelGapPx(drawnValTickFontPx)
        : 6;
      ctx.fillText(formatPrimaryValueAxisTick(chart, v, axisIsPercent), px0 - gap, gy);
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of areaPlan.minorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, toY(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }
  // Category-axis baseline + value-axis rule. Office treats
  // `<c:*Ax><c:spPr><a:ln><a:noFill>` as suppressing the rule and tick marks
  // while labels/gridlines remain. The value rule is drawn only when the file
  // gives it a colour, matching the bar/line renderers.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, primaryCategoryAxisY, px0 + pw, primaryCategoryAxisY,
      catLineColor, catLineW, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      valLineColor, valLineW, chart.valAxisLineDash,
    );
  }
  // Category-axis major tick marks. With crossBetween="between" PowerPoint
  // draws them at the band BOUNDARIES (n+1 dividers); "midCat" ticks centers.
  if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
    const tickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    if (dateAxisPlan) {
      for (const tick of dateAxisPlan.majorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY,
          px0 + tick.fraction * pw, catLineColor, catLineW,
          false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash,
        );
      }
    } else if (between) {
      for (let ci = 0; ci <= n; ci += tickSkip) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, px0 + (ci / n) * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      }
    } else {
      for (let ci = 0; ci < n; ci += tickSkip) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, toX(ci), catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      }
    }
  }
  if (
    !chart.catAxisHidden
    && chart.catAxisMinorTickMark
    && chart.catAxisMinorTickMark !== 'none'
    && dateAxisPlan
  ) {
    for (const tick of dateAxisPlan.minorTicks) {
      drawAxisTick(
        ctx, chart.catAxisMinorTickMark, 'cat', primaryCategoryAxisY,
        px0 + tick.fraction * pw, catLineColor, catLineW,
        false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
      );
    }
  }

  if (!hasDataTable && !chart.catAxisHidden) {
    const drawnCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, pw / n * 0.8));
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      drawnCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Category labels are controlled by the authored `<c:tickLblSkip>` interval.
    // Do not add an automatic collision interval: sparse category caches often
    // deliberately leave alternating entries empty to obtain a two-year label
    // cadence, and a computed interval starting at index 0 can discard every
    // non-empty label. Excel paints the authored sparse labels even when the
    // final pair overlaps.
    // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
    // category-axis numFmt before measuring and drawing; string categories
    // pass through unchanged.
    const authoredSkip = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        label: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        x: px0 + tick.fraction * pw,
        categoryIndex: -1,
      }))
      : Array.from({ length: Math.ceil(n / authoredSkip) }, (_, index) => {
        const ci = index * authoredSkip;
        return {
          label: formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904),
          x: toX(ci),
          categoryIndex: ci,
        };
      });
    for (const entry of labelEntries) {
      const label = entry.label;
      if (!label) continue;
      const anchor = entry.categoryIndex < 0
        ? null
        : categoryLabelAnchorFraction(
          entry.categoryIndex,
          n,
          isCrossBetween(chart),
          catAxisReversed(chart),
          chart.catAxisLabelAlignment,
        );
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(drawnCatTickFontPx)
          : 3,
        chart.catAxisLabelOffsetPercent,
      );
      ctx.textAlign = anchor?.textAlign ?? 'center';
      const labelPosition = chart.catAxisTickLabelPos ?? 'nextTo';
      const labelAxisY = labelPosition === 'nextTo'
        ? primaryCategoryAxisY
        : labelPosition === 'high' ? py0 : py0 + ph;
      ctx.fillText(label, anchor ? px0 + anchor.fraction * pw : entry.x, labelAxisY + gap);
    }
  }

  // Secondary value axis (right edge) — drawn after the fills + category labels
  // so it sits atop the plot, mirroring the bar/line ordering.
  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r, px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  if (dataTableLayout) {
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, py0 + ph, pw, x + legLeftW, ptToPx,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

/** Inside-radius fraction (of the outer radius) for a SOLID pie's `ctr` / `inEnd`
 *  / `bestFit` data labels (§21.2.2.48). PowerPoint places these near the rim,
 *  not at the disc mid-radius: four observed slice sizes place labels at
 *  0.878 / 0.888 / 0.887 / 0.912·outerR — a flat near-rim
 *  constant independent of slice angle (see the `labelR` comment in
 *  {@link drawPieRichLabels}). 0.88 is the empirical fit; it is an approximation
 *  of an undocumented PowerPoint layout, not a spec-defined geometry. Doughnut
 *  labels use the exact ring midpoint instead and never consult this. */
const PIE_CTR_LABEL_RADIUS_FRAC = 0.88;

interface OfPiePoint {
  sourceIndex: number;
  value: number;
}

/** ECMA-376 §21.2.2.126 pie-of-pie / bar-of-pie. Source point identity stays
 * attached to every detail item; only the primary plot receives one aggregate
 * slice. Automatic geometry is deliberately compact and parameterized solely
 * by the authored gap and second-plot size. */
function renderOfPieChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const source = chart.series[0];
  if (!source) return;
  const secondarySet = planOfPieSecondaryIndices(chart.ofPie, source.values);
  if (secondarySet == null || secondarySet.size === 0) {
    renderPieChart(
      ctx, { ...chart, chartType: 'pie' }, r, false, ptToPx, shapeRotationDeg,
    );
    return;
  }
  const primary: OfPiePoint[] = [];
  const secondary: OfPiePoint[] = [];
  for (let sourceIndex = 0; sourceIndex < source.values.length; sourceIndex++) {
    const sourceValue = source.values[sourceIndex];
    const value = sourceValue == null ? 0 : Math.abs(sourceValue);
    if (!(value > 0) || !Number.isFinite(value)) continue;
    (secondarySet.has(sourceIndex) ? secondary : primary).push({ sourceIndex, value });
  }
  if (secondary.length === 0) {
    renderPieChart(
      ctx, { ...chart, chartType: 'pie' }, r, false, ptToPx, shapeRotationDeg,
    );
    return;
  }

  const legendChart: ChartModel = { ...chart, chartType: 'pie' };
  const legend = measuredLegendReserve(ctx, legendChart, r.w, r.h, 0.28, ptToPx);
  const frame = computeChartFrame(chart, r.x, r.y, r.w, r.h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.28,
    legendReserve: legend,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(
    ctx, chart, r.x, r.y, r.w, r.h,
    r.y + frame.title.topPad, frame.title.fontPx,
  );
  const { px0, py0, pw, ph } = frame.plotRect;
  if (!(pw > 0) || !(ph > 0)) return;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const options = chart.ofPie;
  const sizeRatio = Math.max(0.05, Math.min(2, (options?.secondPieSizePercent ?? 75) / 100));
  const gapUnits = Math.max(0, options?.gapWidthPercent ?? 150) / 100;
  const mainRadius = Math.min(ph * 0.44, (pw * 0.9) / (2 + 2 * sizeRatio + gapUnits));
  const detailRadius = mainRadius * sizeRatio;
  if (!(mainRadius > 0) || !(detailRadius > 0)) return;
  const usedW = 2 * mainRadius + gapUnits * mainRadius + 2 * detailRadius;
  const left = px0 + (pw - usedW) / 2;
  const mainCx = left + mainRadius;
  const detailCx = left + 2 * mainRadius + gapUnits * mainRadius + detailRadius;
  const cy = py0 + ph / 2;
  const secondaryTotal = secondary.reduce((sum, point) => sum + point.value, 0);
  const mainPoints: OfPiePoint[] = [
    ...primary,
    { sourceIndex: secondary[0].sourceIndex, value: secondaryTotal },
  ];

  const drawPie = (
    points: OfPiePoint[], cx: number, radius: number,
  ): { aggregateStart: number; aggregateEnd: number } => {
    const total = points.reduce((sum, point) => sum + point.value, 0);
    let angle = -Math.PI / 2;
    let aggregateStart = angle;
    let aggregateEnd = angle;
    for (let index = 0; index < points.length; index++) {
      const point = points[index];
      const sweep = total > 0 ? point.value / total * Math.PI * 2 : 0;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, angle, angle + sweep);
      ctx.closePath();
      ctx.fillStyle = pieSliceColor(point.sourceIndex, source, chart.varyColors !== false);
      ctx.fill();
      ctx.strokeStyle = '#fff';
      ctx.lineWidth = 1;
      ctx.stroke();
      if (index === points.length - 1) {
        aggregateStart = angle;
        aggregateEnd = angle + sweep;
      }
      angle += sweep;
    }
    return { aggregateStart, aggregateEnd };
  };

  const aggregateAngles = drawPie(mainPoints, mainCx, mainRadius);
  let connectorTop = cy - detailRadius;
  let connectorBottom = cy + detailRadius;
  if ((options?.type ?? 'pie') === 'bar') {
    let top = cy - detailRadius;
    const barW = detailRadius;
    for (const point of secondary) {
      const height = secondaryTotal > 0 ? 2 * detailRadius * point.value / secondaryTotal : 0;
      ctx.fillStyle = pieSliceColor(point.sourceIndex, source, chart.varyColors !== false);
      ctx.fillRect(detailCx - barW / 2, top, barW, height);
      ctx.strokeStyle = '#fff';
      ctx.lineWidth = 1;
      ctx.strokeRect(detailCx - barW / 2, top, barW, height);
      top += height;
    }
    connectorTop = cy - detailRadius;
    connectorBottom = cy + detailRadius;
  } else {
    drawPie(secondary, detailCx, detailRadius);
  }

  if (options?.seriesLines ?? true) {
    ctx.strokeStyle = '#808080';
    ctx.lineWidth = Math.max(1, 0.75 * ptToPx);
    ctx.setLineDash([]);
    const fromTop = {
      x: mainCx + Math.cos(aggregateAngles.aggregateStart) * mainRadius,
      y: cy + Math.sin(aggregateAngles.aggregateStart) * mainRadius,
    };
    const fromBottom = {
      x: mainCx + Math.cos(aggregateAngles.aggregateEnd) * mainRadius,
      y: cy + Math.sin(aggregateAngles.aggregateEnd) * mainRadius,
    };
    ctx.beginPath(); ctx.moveTo(fromTop.x, fromTop.y); ctx.lineTo(detailCx - detailRadius, connectorTop); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(fromBottom.x, fromBottom.y); ctx.lineTo(detailCx - detailRadius, connectorBottom); ctx.stroke();
  }
  if (legend) {
    drawLegendForLayout(
      ctx, legendChart, legend,
      r.x, r.y, r.w, r.h, px0, py0, pw, ph, frame.title.bandH + 2, ptToPx,
    );
  }
}

function renderPieChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  isDoughnut: boolean,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;
  const legendChart: ChartModel = {
    ...chart,
    series: [{ ...s, categories: cats }],
  };

  // Shared frame (radial form). Pie uses title pads 0.035 / 0.035; its legend
  // labels categories (one row per slice) so it reserves a wider 0.28 side band
  // (vs the default 0.22). The h*0.02 gap below the title/legend before centring
  // is the shared radial gap. Params keep pixels unchanged.
  const pieLeg = measuredLegendReserve(ctx, legendChart, w, h, 0.28, ptToPx);
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.28,
    legendReserve: pieLeg,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  const titleFontPx = frame.title.fontPx;
  const titleH = frame.title.bandH;
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, titleFontPx);

  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(
    ctx, chart, plotLeft, plotTop, pw, ph, ptToPx, shapeRotationDeg,
  );
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  const outerR = Math.min(pw, ph) * 0.42;

  // §21.2.2.52 firstSliceAng: the first slice begins `firstSliceAngle` degrees
  // clockwise from 12 o'clock. Canvas 0 rad points right (+x) and its angles
  // grow clockwise (y-down), so 12 o'clock is −90°. Default 0 keeps the
  // historical −90° start (byte-stable for files without the element).
  const startAngle = -Math.PI / 2 + ((chart.firstSliceAngle ?? 0) * Math.PI) / 180;

  // §21.2.2.60 holeSize (doughnut only): hole diameter as 1–90% of the outer
  // diameter. The ECMA schema default is 10%, but a real doughnut always writes
  // an explicit holeSize (Office emits 50–75%); 50% is the historical inner
  // radius, so an absent holeSize keeps the prior look (byte-stable). Pie has
  // no hole (innerR = 0).
  const holePct = isDoughnut ? Math.max(1, Math.min(90, chart.holeSize ?? 50)) : 0;

  // Concentric rings. Doughnut plots EVERY series as a ring (outermost =
  // series[0]); pie plots only series[0]. The band from the hole radius to the
  // outer radius is split evenly across the rings. A single-series doughnut is
  // byte-identical to the prior single-ring geometry.
  const rings = isDoughnut ? chart.series : [s];
  const pointOverridesBySeries = new Map(
    rings.map(series => [series, indexPointOverrides(series.dataPointOverrides)] as const),
  );
  const innerR = outerR * (holePct / 100);
  const ringBand = (outerR - innerR) / rings.length;

  // Explosion offset for slice `i` of series `ser`: move the slice out from the
  // center along its mid-angle by `explosion`% of the outer radius. §21.2.2.61
  // only defines `explosion` as an unbounded `xsd:unsignedInt` "amount the data
  // point shall be moved from the center of the pie" — the 0-100-as-percent
  // interpretation is a de-facto Office convention (the Point Explosion UI
  // slider), not a spec-mandated range (see `ChartDataPointOverride.explosion`
  // in types/chart.ts). Absent / zero explosion → no offset (byte-stable).
  const explodeOffset = (ser: ChartSeries, i: number): number => {
    const e = pointOverridesBySeries.get(ser)?.get(i)?.explosion ?? ser.explosion ?? 0;
    return e > 0 ? (e / 100) * outerR : 0;
  };

  // The legacy `showDataLabels` percent label (drawn INLINE per slice on the
  // outer ring, exactly as before) is used only when the series has no rich
  // `<c:dLbls>` definition; the rich labels are drawn in a separate pass after
  // all slices. Keeping the legacy path inline preserves the historical
  // draw-call order for a plain pie/doughnut (byte-stable).
  const richDef: ChartSeriesDataLabels = s.seriesDataLabels ?? {
    showVal: false,
    showCatName: false,
    showSerName: false,
    showPercent: false,
  };
  // A point-level dLbl is independently authored and must not depend on a
  // series dLbls default existing. Even a delete-only override participates in
  // rich dispatch so the legacy chart-wide percent path cannot resurrect it.
  const hasRichLabels = s.seriesDataLabels != null || (s.dataLabelOverrides?.length ?? 0) > 0;
  const legacyLabels = chart.showDataLabels && !hasRichLabels;
  const dLblFont = chartFontFamily(
    chart, richDef.fontFace ?? chart.dataLabelFontFace, 'minor',
  );

  for (let ring = 0; ring < rings.length; ring++) {
    const rs = rings[ring];
    const rVals = rs.values.map(v => Math.abs(v ?? 0));
    const rTotal = rVals.reduce((a, b) => a + b, 0);
    if (rTotal === 0) continue;
    // Ring 0 is the OUTERMOST band; deeper rings step inward toward the hole.
    const rOuter = outerR - ring * ringBand;
    const rInner = rOuter - ringBand;

    let angle = startAngle;
    for (let i = 0; i < rVals.length; i++) {
      const slice = (rVals[i] / rTotal) * Math.PI * 2;
      const color = pieSliceColor(i, rs, chart.varyColors !== false);
      const midAngle = angle + slice / 2;
      const off = explodeOffset(rs, i);
      const ox = off > 0 ? Math.cos(midAngle) * off : 0;
      const oy = off > 0 ? Math.sin(midAngle) * off : 0;
      ctx.beginPath();
      if (rInner > 0.01) {
        // Annular slice (doughnut ring): outer arc CW, inner arc CCW.
        ctx.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
        ctx.arc(cx2 + ox, cy2 + oy, rInner, angle + slice, angle, true);
      } else {
        // Solid wedge (pie, or the innermost pie-like ring).
        ctx.moveTo(cx2 + ox, cy2 + oy);
        ctx.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
      }
      ctx.closePath();
      ctx.fillStyle = color; ctx.fill();
      const point = pointOverridesBySeries.get(rs)?.get(i);
      const lineHidden = point?.lineHidden ?? rs.lineHidden;
      const lineColor = point?.lineColor ?? rs.lineColor;
      if (lineHidden !== true && lineColor) {
        const lineWidthEmu = point?.lineWidthEmu ?? rs.lineWidthEmu;
        const lineWidth = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx) : Math.max(.5, ptToPx * .75);
        ctx.save();
        ctx.strokeStyle = `#${lineColor}`;
        ctx.lineWidth = lineWidth;
        ctx.setLineDash(dashPatternForPreset(
          point?.lineDash ?? rs.chartexStyle?.lineDash ?? undefined,
          lineWidth,
        ));
        ctx.lineCap = rs.chartexStyle?.lineCap === 'rnd'
          ? 'round' : rs.chartexStyle?.lineCap === 'sq' ? 'square' : 'butt';
        ctx.lineJoin = rs.chartexStyle?.lineJoin === 'round'
          || rs.chartexStyle?.lineJoin === 'bevel'
          ? rs.chartexStyle.lineJoin : 'miter';
        ctx.stroke();
        ctx.restore();
      }

      // Legacy percent label — outer ring only, drawn inline (byte-stable).
      if (legacyLabels && ring === 0 && slice > 0.15) {
        const labelR = outerR * (isDoughnut ? 0.75 : 0.6);
        const lx2 = cx2 + ox + Math.cos(midAngle) * labelR;
        const ly2 = cy2 + oy + Math.sin(midAngle) * labelR;
        const pct2 = Math.round((rVals[i] / rTotal) * 100);
        const lsz = Math.max(8, outerR * 0.1);
        ctx.font = `bold ${lsz}px ${dLblFont}`;
        ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(`${pct2}%`, lx2, ly2);
      }

      angle += slice;
    }
  }

  // Rich data labels (`<c:dLbls>`: showVal / showCatName / showSerName /
  // showPercent + dLblPos, §21.2.2.35), drawn on the OUTER ring after all
  // slices. Only runs when a rich definition is present; the plain percent
  // labels above are byte-identical to the pre-CH8 pie.
  if (hasRichLabels) {
    const outerRingInnerR = isDoughnut ? outerR - ringBand : 0;
    drawPieRichLabels(
      ctx, chart, richDef, s, cats, vals, total,
      cx2, cy2, outerR, outerRingInnerR, startAngle, dLblFont, ptToPx,
      plotLeft, plotTop, pw, ph,
      x, y, w, h,
      shapeRotationDeg,
    );
  }

  if (pieLeg) {
    // Pie/doughnut legends are category-driven: one row per slice, each colored
    // exactly like its slice (`pieSliceColor`). `buildLegendEntries` derives the
    // rows from the real series, so pass it through unchanged (with the resolved
    // category labels attached). The previous pseudo-series collapsed all
    // swatches to one color because it folded the series-level fill (`s.color`)
    // into every entry while the slices used the per-index palette.
    drawLegendForLayout(
      ctx, legendChart, pieLeg,
      x, y, w, h, plotLeft, plotTop, pw, ph, titleH + 2,
      ptToPx,
    );
  }
}

/** Draw the rich outer-ring data labels for a pie / doughnut from a series-level
 *  `<c:dLbls>` (§21.2.2.35: showVal / showCatName / showSerName / showPercent +
 *  dLblPos). Only called when such a definition exists; the plain percent-label
 *  path stays inline in the slice loop (byte-stable). `font` is the pre-resolved
 *  data-label CSS font-family.
 *
 *  When the `<c:dLbls>` carries a callout-box shape (`<c:spPr>` → `def.labelBox`,
 *  §21.2.2.197) the labels are drawn Word-style: each is a boxed callout placed
 *  OUTSIDE its slice at the slice mid-angle, with adjacent boxes pushed apart to
 *  avoid overlap (`bestFit`), and a leader line back to the rim for any box that
 *  ends up far from its slice. Plain `outEnd` labels use the same outside-rim
 *  invariant without painting a box; the inside positions retain their radial
 *  layout. */
function drawPieRichLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number, innerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  plotX: number, plotY: number, plotW: number, plotH: number,
  chartX: number, chartY: number, chartW: number, chartH: number,
  shapeRotationDeg: number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(
    { ...chart, series: [{ ...s, categories: cats }] },
    ptToPx,
  );
  const calloutIndices = new Set<number>();
  for (let index = 0; index < vals.length; index++) {
    const override = overridesByIndex.get(index);
    if (dataLabelIsDeleted(def, override)) continue;
    const labelBox = mergeChartLabelBoxes(override?.labelBox, def.labelBox);
    // A border-only label shape is an outline around the ordinary radial
    // label; it does not opt into Word's filled boxed-callout layout. Treating
    // any visible border as a callout invented leaders for Excel bestFit pie
    // labels even though the authored labels remain on their slices.
    const hasVisibleCalloutFill = labelBox?.fillHidden !== true
      && (labelBox?.fill != null || labelBox?.fillPaint != null);
    if (hasVisibleCalloutFill) {
      calloutIndices.add(index);
    }
  }
  // Boxed labels have their own paint/collision pass, but only those points are
  // dispatched to it. A per-point spPr must not turn every sibling into a
  // callout or change the sibling's authored radial position.
  if (calloutIndices.size > 0) {
    drawPieCalloutLabels(
      ctx, chart, def, s, cats, vals, total, cx2, cy2, outerR, innerR, startAngle,
      font, ptToPx, plotX, plotW, plotY, plotH, chartX, chartY, chartW, chartH,
      calloutIndices, overridesByIndex, shapeRotationDeg,
    );
  }

  const outsideLabels: PieOutsideLabel[] = [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    if (calloutIndices.has(i)) continue;
    // A per-point `<c:dLbl idx>` (§21.2.2.47) overrides the series-level
    // `<c:dLbls>` (§21.2.2.49) for this one slice. Its show-flags, font color /
    // size / bold, and position each fall back to the series default when the
    // point declares none. A point may set `showCatName=0 showPercent=1` plus
    // white text while the series default is
    // `showCatName=1` black — so honoring the per-point flags is what makes the
    // labels render as white percent-only (matching PowerPoint / the PDF).
    const ov = overridesByIndex.get(i);
    // A genuinely deleted label (`<c:delete val="1">`, §21.2.2.43) is skipped.
    // A style/flag-only `<c:dLbl>` (no `<c:tx>`) is NOT a delete. Such slices
    // can carry `text: ""` with white/percent-only flag overrides, so we
    // key off the explicit `deleted` flag, never the empty text.
    if (dataLabelIsDeleted(def, ov)) continue;
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    const showLegendKey = ov?.showLegendKey ?? def.showLegendKey ?? false;
    // §21.2.2.35 label composition. A per-point custom `<c:tx>` (non-empty
    // override text) wins outright; otherwise compose from the resolved flags.
    // Positioning is handled below (§21.2.2.48 `dLblPos`). Percent is the
    // slice's share of the total.
    const text = effectiveDataLabelText({
      customText: ov?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: (cats[i] ?? '').toString(),
      seriesName: s.name,
      sourceValue: vals[i],
      percentRatio: vals[i] / total,
      formatCode: ov?.formatCode ?? def.formatCode ?? s.valFormatCode ?? null,
      percentFormatCode: ov?.formatCode ?? def.formatCode ?? '0%',
      date1904: chart.date1904 ?? false,
      separator: ov?.separator ?? def.separator,
    });
    const legendKey = showLegendKey ? dataLabelLegendKey(0, i) : undefined;
    if (!text && !legendKey) continue;
    const pos = ov?.position ?? def.position ?? 'bestFit';
    const outside = pos === 'outEnd';
    const sizeHpt = ov?.fontSizeHpt ?? def.fontSizeHpt;
    const sizePx = chartTextFontSizePx(sizeHpt, ptToPx) ?? Math.max(8, outerR * 0.1);
    const bold = ov?.fontBold ?? def.fontBold;
    const fontColor = ov?.fontColor ?? def.fontColor;
    const labelFont = (ov?.fontFace ?? def.fontFace)
      ? chartFontFamily(chart, ov?.fontFace ?? def.fontFace, 'minor')
      : font;
    const textStyle = effectiveDataLabelTextStyle(ov, def);
    const rich = customRichDataLabelOptions(
      chart, ov, ptToPx, labelFont, bold ?? false, textStyle,
    );
    const automaticLabelR = innerR > 0.01
      ? (innerR + outerR) / 2
      : outerR * PIE_CTR_LABEL_RADIUS_FRAC;
    if (ov?.manualLayout) {
      ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
      drawBoundedDataLabelText(
        ctx,
        text,
        {
          kind: 'point',
          x: cx2 + Math.cos(midAngle) * automaticLabelR,
          y: cy2 + Math.sin(midAngle) * automaticLabelR,
          position: 'ctr',
        },
        { x: plotX, y: plotY, w: plotW, h: plotH },
        sizePx,
        fontColor ? `#${fontColor}` : '#fff',
        ov.manualLayout,
        { x: chartX, y: chartY, w: chartW, h: chartH },
        rich,
        legendKey,
        textStyle,
        ptToPx,
        mergeChartLabelBoxes(ov?.labelBox, def.labelBox),
        shapeRotationDeg,
      );
      continue;
    }
    if (outside) {
      ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
      const richBlock = rich
        ? resolveRichDataLabelBlock(ctx, rich, sizePx, fontColor ? `#${fontColor}` : '#333')
        : null;
      const lineHeight = sizePx * 1.15;
      const fittedLines = richBlock ? [] : fitStyledDataLabelLines(
        text, Math.max(0, chartW - sizePx), Math.max(0, chartH - sizePx),
        lineHeight, value => ctx.measureText(value).width, textStyle,
      );
      if (rich && !richBlock) continue;
      if (!richBlock && fittedLines.length === 0 && !legendKey) continue;
      const textW = richBlock?.width ?? fittedLines.reduce(
        (max, line) => Math.max(max, ctx.measureText(line).width), 0,
      );
      const textH = richBlock?.height
        ?? (sizePx + Math.max(0, fittedLines.length - 1) * lineHeight);
      const keyW = legendKey
        ? (legendSwatchWidths([legendKey.entry], sizePx, ptToPx)[0] ?? 0)
        : 0;
      const keyH = legendKey ? legendSwatchHeight(legendKey.entry, sizePx, ptToPx) : 0;
      outsideLabels.push(createPieOutsideLabel(
        fittedLines, midAngle, cx2, cy2, outerR,
        Math.min(keyW + (text ? LEGEND_SWATCH_TEXT_GAP : 0) + textW, Math.max(0, chartW - sizePx)),
        Math.min(Math.max(keyH, textH), Math.max(0, chartH - sizePx)),
        lineHeight, sizePx, bold ?? false,
        fontColor ? `#${fontColor}` : '#333',
        labelFont,
        richBlock ?? undefined,
        legendKey,
        textStyle,
        ptToPx,
      ));
      continue;
    }
    // §21.2.2.48 ST_DLblPos radial placement. The spec enumerates the positions
    // (bestFit / ctr / inEnd / outEnd …) but gives no geometry, so the inside
    // radii below reproduce the bounded Office vector observations for solid
    // pie and doughnut labels:
    //
    //   • DOUGHNUT (innerR > 0), ctr / inEnd / bestFit → the RING midpoint
    //     (innerR + outerR)/2. Verified on the 55%-hole doughnut: labels sit at
    //     0.772–0.778·outerR ≈ (0.55+1)/2 = 0.775. Byte-stable — unchanged.
    //   • SOLID pie (innerR ≈ 0), ctr / inEnd / bestFit → ≈0.88·outerR, NOT the
    //     disc mid-radius. Measured label-centroid ratios across the 54/27/14/5%
    //     slices were 0.878 / 0.888 / 0.887 / 0.912 (center + outer radius from a
    //     least-squares rim fit, residual std 0.43pt), i.e. a flat near-rim
    //     constant independent of slice angle — so it is a fixed fraction, not a
    //     sector centroid. The 5% sliver rides marginally further out in
    //     PowerPoint; we do not model that per-slice nudge. This is an empirical
    //     approximation of an undocumented PowerPoint layout, not a spec formula.
    const labelR = automaticLabelR;
    const lx2 = cx2 + Math.cos(midAngle) * labelR;
    const ly2 = cy2 + Math.sin(midAngle) * labelR;
    ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
    const tangentialCapacity = 2 * labelR * Math.sin(Math.min(Math.PI, Math.abs(slice)) / 2)
      - sizePx;
    const radialCapacity = innerR > 0.01
      ? outerR - innerR - sizePx
      : outerR - sizePx;
    if (!(tangentialCapacity > 0) || !(radialCapacity > 0)) continue;
    const sliceBounds = dataLabelRectIntersection(
      {
        x: lx2 - tangentialCapacity / 2,
        y: ly2 - radialCapacity / 2,
        w: tangentialCapacity,
        h: radialCapacity,
      },
      { x: plotX, y: plotY, w: plotW, h: plotH },
    );
    if (!sliceBounds) continue;
    drawBoundedDataLabelText(
      ctx,
      text,
      { kind: 'point', x: lx2, y: ly2, position: 'ctr' },
      sliceBounds,
      sizePx,
      fontColor ? `#${fontColor}` : '#fff',
      undefined,
      { x: chartX, y: chartY, w: chartW, h: chartH },
      rich,
      legendKey,
      textStyle,
      ptToPx,
      mergeChartLabelBoxes(ov?.labelBox, def.labelBox),
      shapeRotationDeg,
    );
  }

  drawPieOutsideLabels(ctx, outsideLabels, chartX, chartY, chartW, chartH);
}

/** Plain `<c:dLblPos val="outEnd">` label block. */
interface PieOutsideLabel {
  lines: string[];
  rich?: RichDataLabelBlock;
  legendKey?: DataLabelLegendKey;
  boxW: number;
  boxH: number;
  unrotatedW: number;
  unrotatedH: number;
  textStyle: DataLabelTextStyle;
  ptToPx: number;
  lineHeight: number;
  fontPx: number;
  bold: boolean;
  fontColor: string;
  font: string;
  cxBox: number;
  cyBox: number;
}

function pointToRectDistance(
  px: number, py: number,
  rectCx: number, rectCy: number,
  halfW: number, halfH: number,
): number {
  const dx = Math.max(Math.abs(rectCx - px) - halfW, 0);
  const dy = Math.max(Math.abs(rectCy - py) - halfH, 0);
  return Math.hypot(dx, dy);
}

/** Find the first point on a slice-midpoint ray whose complete visible label
 * rectangle clears the pie. This restores the release geometry without
 * reintroducing collision moves or their synthetic leader lines. */
function outsideLabelRadialDistance(
  midAngle: number,
  outerR: number,
  halfW: number,
  halfH: number,
  clearance: number,
): number {
  const ux = Math.cos(midAngle);
  const uy = Math.sin(midAngle);
  const target = outerR + clearance;
  let low = 0;
  let high = target + Math.hypot(halfW, halfH);
  for (let i = 0; i < 32; i++) {
    const mid = (low + high) / 2;
    const distance = pointToRectDistance(0, 0, ux * mid, uy * mid, halfW, halfH);
    if (distance >= target) high = mid;
    else low = mid;
  }
  return high;
}

function createPieOutsideLabel(
  lines: string[],
  midAngle: number,
  pieCx: number,
  pieCy: number,
  outerR: number,
  boxW: number,
  boxH: number,
  lineHeight: number,
  fontPx: number,
  bold: boolean,
  fontColor: string,
  font: string,
  rich?: RichDataLabelBlock,
  legendKey?: DataLabelLegendKey,
  textStyle: DataLabelTextStyle = {},
  ptToPx = 1,
): PieOutsideLabel {
  const visibleRotated = rotatedDataLabelSize(
    boxW, boxH, textStyle.textRotation, textStyle.textVerticalMode,
  );
  const insets = dataLabelInsets(textStyle, ptToPx);
  const unrotatedW = boxW + insets.left + insets.right;
  const unrotatedH = boxH + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
  );
  boxW = rotated.w;
  boxH = rotated.h;
  // `outEnd` requires the visible label content, rather than only its anchor,
  // to clear the pie. Keep the release-era radial geometry while leaving each
  // label on its authored slice-midpoint ray; no collision movement means no
  // synthetic leader line is introduced.
  const distance = outsideLabelRadialDistance(
    midAngle, outerR, visibleRotated.w / 2, visibleRotated.h / 2, fontPx * 0.5,
  );
  const cxBox = pieCx + Math.cos(midAngle) * distance;
  const cyBox = pieCy + Math.sin(midAngle) * distance;
  return {
    lines, rich, legendKey,
    boxW, boxH, unrotatedW, unrotatedH, textStyle, ptToPx,
    lineHeight, fontPx, bold, fontColor, font,
    cxBox, cyBox,
  };
}

/** Paint automatic plain outEnd labels at their authored slice-midpoint anchors.
 * Rich callout labels retain their separate bounded collision/leader resolver. */
function drawPieOutsideLabels(
  ctx: CanvasRenderingContext2D,
  labels: PieOutsideLabel[],
  boundsX: number,
  boundsY: number,
  boundsW: number,
  boundsH: number,
): void {
  if (labels.length === 0) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(boundsX, boundsY, boundsW, boundsH);
  ctx.clip();

  for (const label of labels) {
    const insets = dataLabelInsets(label.textStyle, label.ptToPx);
    const rotated = rotatedDataLabelSize(
      label.unrotatedW, label.unrotatedH,
      label.textStyle.textRotation, label.textStyle.textVerticalMode,
    );
    const textCx = label.cxBox + (insets.left - insets.right) / 2;
    const textCy = label.cyBox + (insets.top - insets.bottom) / 2;
    const innerWidth = Math.max(0, label.unrotatedW - insets.left - insets.right);
    const paintAlign = dataLabelCanvasTextAlign(label.textStyle, 'center');
    const textAnchorX = paintAlign === 'left'
      ? label.cxBox - label.unrotatedW / 2 + insets.left
      : paintAlign === 'right'
        ? label.cxBox + label.unrotatedW / 2 - insets.right
        : textCx;
    ctx.save();
    if (rotated.radians !== 0) {
      ctx.translate(label.cxBox, label.cyBox);
      ctx.rotate(rotated.radians);
      ctx.translate(-label.cxBox, -label.cyBox);
    }
    if (!label.legendKey) {
      if (label.rich) {
        paintRichDataLabelBlock(
          ctx, label.rich, textAnchorX, textCy, paintAlign, 'middle', innerWidth,
        );
        ctx.restore();
        continue;
      }
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.bold ? 'bold ' : ''}${label.fontPx}px ${label.font}`;
      ctx.fillStyle = label.fontColor;
      ctx.textAlign = paintAlign;
      ctx.textBaseline = 'middle';
      const baselineShift = (label.textStyle.fontBaseline ?? 0) * label.fontPx;
      const firstY = textCy - ((label.lines.length - 1) * label.lineHeight) / 2 - baselineShift;
      if (!(label.textStyle.fontPaintAuthored === true
        && (label.textStyle.fontHidden === true || label.textStyle.fontColor == null))) {
        for (let i = 0; i < label.lines.length; i++) {
          ctx.fillText(label.lines[i], textAnchorX, firstY + i * label.lineHeight);
        }
      }
      ctx.restore();
      continue;
    }
    ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.bold ? 'bold ' : ''}${label.fontPx}px ${label.font}`;
    const keyWidth = label.legendKey
      ? (legendSwatchWidths([label.legendKey.entry], label.fontPx, label.legendKey.ptToPx)[0] ?? 0)
      : 0;
    const keyHeight = label.legendKey
      ? legendSwatchHeight(label.legendKey.entry, label.fontPx, label.legendKey.ptToPx)
      : 0;
    const textWidth = label.rich?.width ?? label.lines.reduce(
      (max, line) => Math.max(max, ctx.measureText(line).width), 0,
    );
    const gap = label.legendKey && (label.rich || label.lines.length > 0)
      ? LEGEND_SWATCH_TEXT_GAP
      : 0;
    const contentWidth = keyWidth + gap + textWidth;
    const contentLeft = textCx - contentWidth / 2;
    if (label.legendKey) {
      drawLegendSwatch(
        ctx,
        label.legendKey.entry.swatchStyle,
        label.legendKey.entry.color,
        contentLeft,
        textCy - keyHeight / 2,
        keyWidth,
        keyHeight,
        label.legendKey.entry.marker,
        label.legendKey.entry.fillPaint,
        label.legendKey.entry.outlineColor,
        label.legendKey.entry.outlineWidthEmu,
        label.legendKey.entry.outlineDash,
        label.legendKey.entry.outlineCap,
        label.legendKey.entry.outlineJoin,
        label.legendKey.ptToPx,
        label.legendKey.shapeRotationDeg,
      );
    }
    if (label.rich) {
      paintRichDataLabelBlock(
        ctx, label.rich, contentLeft + keyWidth + gap, textCy, 'left', 'middle',
      );
      ctx.restore();
      continue;
    }
    ctx.fillStyle = label.fontColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const baselineShift = (label.textStyle.fontBaseline ?? 0) * label.fontPx;
    const firstY = textCy - ((label.lines.length - 1) * label.lineHeight) / 2 - baselineShift;
    if (!(label.textStyle.fontPaintAuthored === true
      && (label.textStyle.fontHidden === true || label.textStyle.fontColor == null))) {
      for (let i = 0; i < label.lines.length; i++) {
        ctx.fillText(label.lines[i], contentLeft + keyWidth + gap, firstY + i * label.lineHeight);
      }
    }
    ctx.restore();
  }
  ctx.restore();
}

/** One laid-out pie callout label: its wrapped text lines, box rectangle, the
 *  rim anchor point on its slice, and the resolved per-point style. */
interface PieCalloutLabel {
  lines: string[];
  rich?: RichDataLabelBlock;
  legendKey?: DataLabelLegendKey;
  lineHeight: number;
  /** Slice mid-angle (canvas radians) — the leader-line target direction. */
  midAngle: number;
  /** Rim anchor point (on the outer arc at `midAngle`). */
  rimX: number;
  rimY: number;
  /** Half-height of the text block (px) — box grows symmetrically around cy. */
  boxW: number;
  boxH: number;
  unrotatedW: number;
  unrotatedH: number;
  /** Box centre (mutated by the collision pass). */
  cxBox: number;
  cyBox: number;
  /** true when the label sits on the left half (box hangs to the left). */
  leftSide: boolean;
  fontColor: string;
  box?: ChartLabelBox;
  fontPx: number;
  bold: boolean;
  font: string;
  textStyle: DataLabelTextStyle;
  ptToPx: number;
  /** An authored inside position keeps the box at its slice anchor. */
  inside: boolean;
  /** Explicit per-point manual layout is never moved by the auto collision pass. */
  manualClip?: DataLabelRect;
}

/** Word-style boxed pie/doughnut callout labels (`bestFit`). Each label is a
 *  filled+bordered rectangle placed just outside its slice at the slice
 *  mid-angle; adjacent boxes on the same side are pushed vertically apart so
 *  they do not overlap, and a leader line is drawn back to the rim for any box
 *  whose gap from the rim exceeds a small threshold. Style (box fill/border,
 *  leader colour/width, per-point font colour and box overrides) all comes from
 *  the parsed model — no empirical constants beyond the layout paddings, which
 *  are geometry (not spec values). */
function drawPieCalloutLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number, innerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  boundsX: number, boundsW: number, boundsY: number, boundsH: number,
  chartX: number, chartY: number, chartW: number, chartH: number,
  indices: ReadonlySet<number>,
  overridesByIndex: ReadonlyMap<number, ChartDataLabelOverride>,
  shapeRotationDeg: number,
): void {
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  const findOverride = (i: number): ChartDataLabelOverride | undefined =>
    overridesByIndex.get(i);

  // Base font size: series default (hpt → px) or a radius-relative fallback.
  const baseFontPx = chartTextFontSizePx(def.fontSizeHpt, ptToPx)
    ?? Math.max(9, outerR * 0.09);

  const seriesBox = def.labelBox;

  // ── Build each label: wrapped lines + measured box + rim anchor ──────────
  const labels: PieCalloutLabel[] = [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    if (slice <= 0) continue;
    if (!indices.has(i)) continue;

    const ov = findOverride(i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the label; a per-point
    // *styling / flag* override is NOT a delete even though
    // it also has `text === ""` — key off the explicit `deleted` flag.
    if (dataLabelIsDeleted(def, ov)) continue;

    // §21.2.2.35 composition, with per-point `<c:dLbl>` show-flags (§21.2.2.47)
    // overriding the series defaults for this slice. Word stacks category name
    // and percent on separate lines, so each `show*` part is
    // its own line rather than space-joined.
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    const showLegendKey = ov?.showLegendKey ?? def.showLegendKey ?? false;
    // Per-point overrides (font colour/size/bold + box), else series defaults.
    const fontPx = chartTextFontSizePx(ov?.fontSizeHpt, ptToPx) ?? baseFontPx;
    const bold = ov?.fontBold ?? def.fontBold ?? false;
    const labelFont = (ov?.fontFace ?? def.fontFace)
      ? chartFontFamily(chart, ov?.fontFace ?? def.fontFace, 'minor')
      : font;
    const fontColor = ov?.fontColor ? `#${ov.fontColor}` : (def.fontColor ? `#${def.fontColor}` : '#000');
    const box = mergeChartLabelBoxes(ov?.labelBox, seriesBox);
    const position = ov?.position ?? def.position ?? 'bestFit';

    const text = effectiveDataLabelText({
      customText: ov?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: (cats[i] ?? '').toString(),
      seriesName: s.name,
      sourceValue: vals[i],
      percentRatio: vals[i] / total,
      formatCode: ov?.formatCode ?? def.formatCode ?? s.valFormatCode ?? null,
      percentFormatCode: ov?.formatCode ?? def.formatCode ?? '0%',
      date1904: chart.date1904 ?? false,
      separator: ov?.separator ?? def.separator,
      defaultSeparator: '\n',
    });
    const legendKey = showLegendKey ? dataLabelLegendKey(0, i) : undefined;
    if (!text && !legendKey) continue;
    const textStyle = effectiveDataLabelTextStyle(ov, def);
    const richOptions = customRichDataLabelOptions(
      chart, ov, ptToPx, labelFont, bold, textStyle,
    );

    const authoredInsets = textStyle.textBodyAuthored === true
      || textStyle.textLInsEmu != null || textStyle.textTInsEmu != null
      || textStyle.textRInsEmu != null || textStyle.textBInsEmu != null;
    const bodyInsets = dataLabelInsets(textStyle, ptToPx);
    const padLeft = authoredInsets ? bodyInsets.left : Math.max(4, fontPx * 0.45);
    const padRight = authoredInsets ? bodyInsets.right : Math.max(4, fontPx * 0.45);
    const padTop = authoredInsets ? bodyInsets.top : Math.max(2, fontPx * 0.28);
    const padBottom = authoredInsets ? bodyInsets.bottom : Math.max(2, fontPx * 0.28);
    const lineGap = fontPx * 0.22;
    const lineH = fontPx + lineGap;
    ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontPx}px ${labelFont}`;
    const rich = richOptions
      ? resolveRichDataLabelBlock(ctx, richOptions, fontPx, fontColor)
      : null;
    if (richOptions && !rich) continue;
    let lines = rich ? [] : fitStyledDataLabelLines(
      text,
      Math.max(0, boundsW - padLeft - padRight),
      Math.max(0, boundsH - padTop - padBottom),
      lineH,
      value => ctx.measureText(value).width,
      textStyle,
    );
    if (!rich && lines.length === 0 && !legendKey) continue;
    let textW = rich?.width ?? 0;
    if (!rich) for (const ln of lines) textW = Math.max(textW, ctx.measureText(ln).width);
    const keyW = legendKey
      ? (legendSwatchWidths([legendKey.entry], fontPx, ptToPx)[0] ?? 0)
      : 0;
    const keyH = legendKey ? legendSwatchHeight(legendKey.entry, fontPx, ptToPx) : 0;
    const keyGap = legendKey && text ? LEGEND_SWATCH_TEXT_GAP : 0;
    let unrotatedW = keyW + keyGap + textW + padLeft + padRight;
    let unrotatedH = Math.max(
      keyH, rich?.height ?? (lines.length > 0 ? lines.length * lineH - lineGap : 0),
    ) + padTop + padBottom;
    let rotated = rotatedDataLabelSize(
      unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
    );
    let boxW = Math.min(rotated.w, boundsW);
    let boxH = Math.max(keyH, rich?.height ?? (lines.length > 0 ? lines.length * lineH - lineGap : 0));
    boxH = Math.min(rotated.h, boundsH);

    const rimX = cx2 + Math.cos(midAngle) * outerR;
    const rimY = cy2 + Math.sin(midAngle) * outerR;
    let leftSide = Math.cos(midAngle) < 0;

    // Initial box centre: outside the rim along the mid-angle. The gap scales
    // with the box so small slices get pulled further out (Word `bestFit`).
    const outGap = Math.max(boxW, boxH) * 0.55 + outerR * 0.06;
    let cxBox = rimX + Math.cos(midAngle) * outGap;
    let cyBox = rimY + Math.sin(midAngle) * outGap;
    let manualClip: DataLabelRect | undefined;
    let inside = false;
    if (ov?.manualLayout) {
      const manual = resolveDataLabelPlacement(
        { kind: 'point', x: cxBox, y: cyBox, position: 'ctr' },
        { x: boundsX, y: boundsY, w: boundsW, h: boundsH },
        { w: boxW, h: boxH },
        fontPx,
        ov.manualLayout,
        { x: chartX, y: chartY, w: chartW, h: chartH },
      );
      if (!manual) continue;
      boxW = manual.rect.w;
      boxH = manual.rect.h;
      unrotatedW = boxW;
      unrotatedH = boxH;
      if (!rich) {
        lines = fitStyledDataLabelLines(
          text,
          Math.max(0, boxW - padLeft - padRight - keyW - keyGap),
          Math.max(0, boxH - padTop - padBottom),
          lineH,
          value => ctx.measureText(value).width,
          textStyle,
        );
        if (lines.length === 0 && !legendKey) continue;
      }
      cxBox = manual.rect.x + manual.rect.w / 2;
      cyBox = manual.rect.y + manual.rect.h / 2;
      leftSide = cxBox < cx2;
      manualClip = manual.clip;
    } else if (position !== 'bestFit' && position !== 'outEnd') {
      const labelR = innerR > 0.01
        ? (innerR + outerR) / 2
        : outerR * PIE_CTR_LABEL_RADIUS_FRAC;
      const labelX = cx2 + Math.cos(midAngle) * labelR;
      const labelY = cy2 + Math.sin(midAngle) * labelR;
      const tangentialCapacity = 2 * labelR
        * Math.sin(Math.min(Math.PI, Math.abs(slice)) / 2) - fontPx;
      const radialCapacity = innerR > 0.01
        ? outerR - innerR - fontPx
        : outerR - fontPx;
      const sliceBounds = dataLabelRectIntersection(
        {
          x: labelX - tangentialCapacity / 2,
          y: labelY - radialCapacity / 2,
          w: tangentialCapacity,
          h: radialCapacity,
        },
        { x: boundsX, y: boundsY, w: boundsW, h: boundsH },
      );
      if (!sliceBounds) continue;
      if (!rich) {
        lines = fitStyledDataLabelLines(
          text,
          Math.max(0, sliceBounds.w - padLeft - padRight - keyW - keyGap),
          Math.max(0, sliceBounds.h - padTop - padBottom),
          lineH,
          value => ctx.measureText(value).width,
          textStyle,
        );
        if (lines.length === 0 && !legendKey) continue;
        textW = lines.reduce((width, line) => Math.max(width, ctx.measureText(line).width), 0);
        unrotatedW = keyW + keyGap + textW + padLeft + padRight;
        unrotatedH = Math.max(keyH, lines.length > 0 ? lines.length * lineH - lineGap : 0)
          + padTop + padBottom;
        rotated = rotatedDataLabelSize(
          unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
        );
        boxW = rotated.w;
        boxH = rotated.h;
      } else {
        boxW = Math.min(boxW, sliceBounds.w);
        boxH = Math.min(boxH, sliceBounds.h);
      }
      const anchorPosition = position === 'inBase' || position === 'inEnd'
        ? 'ctr'
        : position;
      const placement = resolveDataLabelPlacement(
        { kind: 'point', x: labelX, y: labelY, position: anchorPosition },
        sliceBounds,
        { w: boxW, h: boxH },
        fontPx,
      );
      if (!placement) continue;
      cxBox = placement.textAlign === 'left'
        ? placement.x + boxW / 2
        : placement.textAlign === 'right'
          ? placement.x - boxW / 2
          : placement.x;
      cyBox = placement.textBaseline === 'top'
        ? placement.y + boxH / 2
        : placement.textBaseline === 'bottom'
          ? placement.y - boxH / 2
          : placement.y;
      leftSide = cxBox < cx2;
      manualClip = placement.clip;
      inside = true;
    }

    labels.push({
      lines, rich: rich ?? undefined, legendKey, lineHeight: lineH, midAngle, rimX, rimY,
      boxW, boxH, unrotatedW, unrotatedH, cxBox, cyBox,
      leftSide, fontColor, box, fontPx, bold, font: labelFont, textStyle, ptToPx,
      inside, manualClip,
    });
  }

  // ── Collision pass (bestFit): split into left/right columns and push boxes
  //    apart vertically so their rectangles do not overlap. Word lays labels
  //    out radially then de-overlaps; this greedy top-down separation +
  //    within-bounds fit-back is a faithful, deterministic approximation (no
  //    sample-specific tuning). ──
  const topLimit = boundsY + 2;
  const bottomLimit = boundsY + boundsH - 2;
  const band = bottomLimit - topLimit;
  const separate = (col: PieCalloutLabel[]): void => {
    if (col.length === 0) return;
    col.sort((a, b) => a.cyBox - b.cyBox);
    // Total height the boxes need when stacked edge-to-edge with a 3px gap
    // between them: the sum of box heights plus the inter-box gaps.
    let stackH = 0;
    for (const l of col) stackH += l.boxH;
    stackH += (col.length - 1) * 3;

    if (stackH > band) {
      // More label than plot: the boxes cannot all fit with the full 3px gaps
      // inside the plot rect. Distribute them so the FIRST box top sits at
      // topLimit and the LAST box bottom sits at bottomLimit, spacing the
      // in-between boxes by an equal step. This keeps the whole column WITHIN
      // [topLimit, bottomLimit] — never spilling past the bottom — which is the
      // overflow #767 guarded against. When the boxes are short enough to fit
      // (sumBoxH ≤ band) the step is a positive gap (no overlap); only a genuine
      // over-pack (sumBoxH > band, i.e. more labels than the plot can hold)
      // forces the boxes to touch/slightly overlap rather than escape the frame.
      const sumBoxH = col.reduce((a, l) => a + l.boxH, 0);
      const n = col.length;
      if (n === 1) {
        col[0].cyBox = Math.min(Math.max(col[0].cyBox, topLimit + col[0].boxH / 2), bottomLimit - col[0].boxH / 2);
        return;
      }
      // Equal gap so first-top = topLimit and last-bottom = bottomLimit:
      //   topLimit + ΣboxH + (n−1)·gap = bottomLimit  ⇒  gap = (band − ΣboxH)/(n−1)
      const gap = (band - sumBoxH) / (n - 1); // may be negative when over-packed
      let cursor = topLimit;
      for (const l of col) {
        l.cyBox = cursor + l.boxH / 2;
        cursor += l.boxH + gap;
      }
      return;
    }

    // Fits: push each box below the previous one by at least their combined half
    // heights (+ a small gap) so rectangles never overlap.
    for (let k = 1; k < col.length; k++) {
      const prev = col[k - 1];
      const cur = col[k];
      const minGap = (prev.boxH + cur.boxH) / 2 + 3;
      if (cur.cyBox - prev.cyBox < minGap) cur.cyBox = prev.cyBox + minGap;
    }
    // The overlap push above is one-directional (boxes only move DOWN), so a
    // bottom-heavy initial layout can now overrun EITHER bound. Because we are
    // in the fits case (stackH ≤ band) the rigid column is shorter than the
    // band, so a single slide brings BOTH ends inside [topLimit, bottomLimit] at
    // once. Slide up by any bottom overflow, then — symmetrically — down by any
    // top underflow. Sliding the whole column down cannot re-cross the bottom
    // because the column fits, so this two-step slide is a true round-trip
    // clamp (the earlier code capped the down-slide against a bottom "room" that
    // the prior up-slide had already zeroed, so a top underflow of ~100px was
    // left uncorrected — #767 was asymmetric, guarding only the bottom edge).
    const bottomOverflow = (col[col.length - 1].cyBox + col[col.length - 1].boxH / 2) - bottomLimit;
    if (bottomOverflow > 0) for (const l of col) l.cyBox -= bottomOverflow;
    const topUnderflow = topLimit - (col[0].cyBox - col[0].boxH / 2);
    if (topUnderflow > 0) for (const l of col) l.cyBox += topUnderflow;
  };
  separate(labels.filter(l => !l.manualClip && !l.leftSide));
  separate(labels.filter(l => !l.manualClip && l.leftSide));

  // Final round-trip clamp (both edges): guarantee no box escapes the plot rect
  // vertically, independent of which separate() branch ran. In the fits case the
  // symmetric slide above already lands every box inside [topLimit, bottomLimit];
  // in the over-packed case the equal-step distribution pins the first top to
  // topLimit and last bottom to bottomLimit. This per-box clamp is therefore a
  // no-op on the current paths, but makes the "no box leaves the frame at either
  // end" invariant explicit and robust to future layout changes. Clamp top FIRST
  // then bottom so a box taller than the band (degenerate) pins to the TOP edge
  // rather than escaping upward.
  for (const l of labels) {
    if (l.manualClip) continue;
    l.cyBox = Math.max(topLimit + l.boxH / 2, l.cyBox);
    l.cyBox = Math.min(bottomLimit - l.boxH / 2, l.cyBox);
  }

  // Horizontal clamp: keep each box fully inside the chart rect.
  const leftLimit = boundsX + 2;
  const rightLimit = boundsX + boundsW - 2;
  for (const l of labels) {
    if (l.manualClip) continue;
    const half = l.boxW / 2;
    if (l.cxBox - half < leftLimit) l.cxBox = leftLimit + half;
    if (l.cxBox + half > rightLimit) l.cxBox = rightLimit - half;
  }

  // ── Draw leader lines first (under the boxes), then boxes + text ─────────
  ctx.save();
  ctx.beginPath();
  ctx.rect(boundsX, boundsY, boundsW, boundsH);
  ctx.clip();
  const leader = chartStyleRoleLeaderLine(chart, def);
  const leaderColor = leader.color ? `#${leader.color}` : '#a6a6a6';
  const leaderPx = leader.widthEmu
    ? Math.max(0.5, (leader.widthEmu / EMU_PER_PT) * ptToPx)
    : 1;
  ctx.setLineDash(dashPatternForPreset(leader.dash ?? undefined, leaderPx));

  for (const l of labels) {
    // The box edge nearest the pie centre — where a leader line should meet.
    const edgeX = l.cxBox + (l.leftSide ? l.boxW / 2 : -l.boxW / 2);
    const edgeY = l.cyBox;
    // Distance from the box's inner edge to its slice rim. When the box abuts
    // the slice the leader is redundant; draw one only past a small threshold.
    const dx = edgeX - l.rimX;
    const dy = edgeY - l.rimY;
    const dist = Math.hypot(dx, dy);
    if (!l.inside && def.showLeaderLines && leader.hidden !== true && dist > l.fontPx * 0.9) {
      ctx.beginPath();
      ctx.moveTo(l.rimX, l.rimY);
      ctx.lineTo(edgeX, edgeY);
      ctx.strokeStyle = leaderColor;
      ctx.lineWidth = leaderPx;
      ctx.stroke();
    }
  }

  for (const l of labels) {
    if (l.manualClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(l.manualClip.x, l.manualClip.y, l.manualClip.w, l.manualClip.h);
      ctx.clip();
    }
    const bx = l.cxBox - l.boxW / 2;
    const by = l.cyBox - l.boxH / 2;
    // Box fill + border (§21.2.2.197 spPr). Fill may carry an 8-digit RGBA hex
    // (e.g. a 90%-opacity white) — valid canvas fillStyle.
    paintChartLabelBox(
      ctx, l.box, { x: bx, y: by, w: l.boxW, h: l.boxH }, ptToPx,
      shapeRotationDeg,
    );
    const authoredInsets = l.textStyle.textBodyAuthored === true
      || l.textStyle.textLInsEmu != null || l.textStyle.textTInsEmu != null
      || l.textStyle.textRInsEmu != null || l.textStyle.textBInsEmu != null;
    const bodyInsets = dataLabelInsets(l.textStyle, l.ptToPx);
    const padLeft = authoredInsets ? bodyInsets.left : Math.max(4, l.fontPx * 0.45);
    const padRight = authoredInsets ? bodyInsets.right : Math.max(4, l.fontPx * 0.45);
    const padTop = authoredInsets ? bodyInsets.top : Math.max(2, l.fontPx * 0.28);
    const padBottom = authoredInsets ? bodyInsets.bottom : Math.max(2, l.fontPx * 0.28);
    const rotated = rotatedDataLabelSize(
      l.unrotatedW, l.unrotatedH,
      l.textStyle.textRotation, l.textStyle.textVerticalMode,
    );
    const contentCx = l.cxBox + (padLeft - padRight) / 2;
    const contentCy = l.cyBox + (padTop - padBottom) / 2;
    const innerLeft = bx + padLeft;
    const innerRight = bx + l.boxW - padRight;
    const innerWidth = Math.max(0, innerRight - innerLeft);
    const paintAlign = dataLabelCanvasTextAlign(l.textStyle, 'center');
    const alignedX = paintAlign === 'left' ? innerLeft
      : paintAlign === 'right' ? innerRight : contentCx;
    const anchoredCenterY = (contentHeight: number): number =>
      (l.textStyle.textVerticalAnchor
        ?? (l.textStyle.textBodyAuthored === true ? 't' : 'ctr')) === 't'
        ? by + padTop + contentHeight / 2
        : (l.textStyle.textVerticalAnchor
          ?? (l.textStyle.textBodyAuthored === true ? 't' : 'ctr')) === 'b'
          ? by + l.boxH - padBottom - contentHeight / 2
          : contentCy;
    const alignedGroupLeft = (contentWidth: number): number =>
      paintAlign === 'left' ? innerLeft
        : paintAlign === 'right' ? innerRight - contentWidth
          : contentCx - contentWidth / 2;
    ctx.save();
    ctx.beginPath();
    ctx.rect(bx, by, l.boxW, l.boxH);
    ctx.clip();
    if (rotated.radians !== 0) {
      ctx.translate(l.cxBox, l.cyBox);
      ctx.rotate(rotated.radians);
      ctx.translate(-l.cxBox, -l.cyBox);
    }
    // Text: centred, stacked lines. A custom rich body uses the same bounded
    // inline block that measured the box, keeping measurement and paint exact.
    if (!l.legendKey) {
      const textHeight = l.rich?.height
        ?? Math.max(0, l.lines.length * l.lineHeight - (l.lineHeight - l.fontPx));
      const textCenterY = anchoredCenterY(textHeight);
      if (l.rich) {
        paintRichDataLabelBlock(
          ctx, l.rich, alignedX, textCenterY, paintAlign, 'middle', innerWidth,
        );
        ctx.restore();
        if (l.manualClip) ctx.restore();
        continue;
      }
      ctx.font = `${l.textStyle.fontItalic ? 'italic ' : ''}${l.bold ? 'bold ' : ''}${l.fontPx}px ${l.font}`;
      ctx.fillStyle = l.fontColor;
      ctx.textAlign = paintAlign;
      ctx.textBaseline = 'middle';
      const lineGap = l.lineHeight - l.fontPx;
      const baselineShift = (l.textStyle.fontBaseline ?? 0) * l.fontPx;
      const blockTop = textCenterY
        - (l.lines.length * l.lineHeight - lineGap) / 2 + l.fontPx / 2 - baselineShift;
      if (!(l.textStyle.fontPaintAuthored === true
        && (l.textStyle.fontHidden === true || l.textStyle.fontColor == null))) {
        for (let li = 0; li < l.lines.length; li++) {
          ctx.fillText(l.lines[li], alignedX, blockTop + li * l.lineHeight);
        }
      }
      ctx.restore();
      if (l.manualClip) ctx.restore();
      continue;
    }
    const keyWidth = l.legendKey
      ? (legendSwatchWidths([l.legendKey.entry], l.fontPx, l.legendKey.ptToPx)[0] ?? 0)
      : 0;
    const keyHeight = l.legendKey
      ? legendSwatchHeight(l.legendKey.entry, l.fontPx, l.legendKey.ptToPx)
      : 0;
    const keyGap = l.legendKey && (l.rich || l.lines.length > 0) ? LEGEND_SWATCH_TEXT_GAP : 0;
    const textWidth = l.rich?.width ?? l.lines.reduce(
      (width, line) => Math.max(width, ctx.measureText(line).width), 0,
    );
    const groupWidth = keyWidth + keyGap + textWidth;
    const groupHeight = Math.max(
      keyHeight,
      l.rich?.height ?? Math.max(0, l.lines.length * l.lineHeight - (l.lineHeight - l.fontPx)),
    );
    const groupCenterY = anchoredCenterY(groupHeight);
    const contentLeft = alignedGroupLeft(groupWidth);
    if (l.legendKey) {
      drawLegendSwatch(
        ctx,
        l.legendKey.entry.swatchStyle,
        l.legendKey.entry.color,
        contentLeft,
        groupCenterY - keyHeight / 2,
        keyWidth,
        keyHeight,
        l.legendKey.entry.marker,
        l.legendKey.entry.fillPaint,
        l.legendKey.entry.outlineColor,
        l.legendKey.entry.outlineWidthEmu,
        l.legendKey.entry.outlineDash,
        l.legendKey.entry.outlineCap,
        l.legendKey.entry.outlineJoin,
        l.legendKey.ptToPx,
        l.legendKey.shapeRotationDeg,
      );
    }
    if (l.rich) {
      paintRichDataLabelBlock(
        ctx, l.rich, contentLeft + keyWidth + keyGap, groupCenterY, 'left', 'middle',
        textWidth,
      );
      ctx.restore();
      if (l.manualClip) ctx.restore();
      continue;
    }
    ctx.font = `${l.textStyle.fontItalic ? 'italic ' : ''}${l.bold ? 'bold ' : ''}${l.fontPx}px ${l.font}`;
    ctx.fillStyle = l.fontColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const lineGap = l.lineHeight - l.fontPx;
    const baselineShift = (l.textStyle.fontBaseline ?? 0) * l.fontPx;
    const blockTop = groupCenterY
      - (l.lines.length * l.lineHeight - lineGap) / 2 + l.fontPx / 2 - baselineShift;
    if (!(l.textStyle.fontPaintAuthored === true
      && (l.textStyle.fontHidden === true || l.textStyle.fontColor == null))) {
      for (let li = 0; li < l.lines.length; li++) {
        ctx.fillText(l.lines[li], contentLeft + keyWidth + keyGap, blockTop + li * l.lineHeight);
      }
    }
    ctx.restore();
    if (l.manualClip) ctx.restore();
  }
  ctx.restore();
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  // Shared frame (radial form). Radar uses title pads 0.035 / 0.035 and the
  // default 0.22 side-legend reserve (unlike pie's 0.28). Params keep pixels
  // unchanged.
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  const titleFontPx = frame.title.fontPx;
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, titleFontPx);

  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(
    ctx, chart, plotLeft, plotTop, pw, ph, ptToPx, shapeRotationDeg,
  );
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  // An explicitly sized `layoutTarget="inner"` rectangle defines the data
  // region itself (ECMA-376 §21.2.2.88), so the outer radar ring is the
  // largest circle inscribed in it. Automatic, position-only, and outer
  // layouts keep the existing label reserve.
  const manualLayout = chart.plotAreaManualLayout;
  const hasExplicitInnerSize = manualLayout?.layoutTarget === 'inner'
    && manualLayout.w != null
    && manualLayout.h != null
    && Number.isFinite(manualLayout.w)
    && Number.isFinite(manualLayout.h)
    && frame.plotAreaManualLayoutApplied;
  const rd = hasExplicitInnerSize
    ? Math.min(pw, ph) / 2
    : Math.min(pw, ph) * 0.38;

  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of chart.series) for (const v of s.values) {
    if (v == null) continue;
    dataMin = Math.min(dataMin, v);
    dataMax = Math.max(dataMax, v);
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (dataMax === 0) dataMax = 1;
  const needsMinorTicks = chart.valAxisMinorTickMark != null
    && chart.valAxisMinorTickMark !== 'none';
  // An explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the automatic
  // ring step. Omission uses the radar-specific radial density observed across
  // small/ordinary/large boundary charts, then the shared bounded planner.
  const radarLog = chart.valAxisLogBase != null
    && Number.isFinite(chart.valAxisLogBase)
    && chart.valAxisLogBase >= 2;
  const radarMajorUnit = chart.valAxisMajorUnit ?? (radarLog
    ? null
    : automaticRadarMajorUnit(
        chart.valMin ?? dataMin,
        chart.valMax ?? dataMax,
        rd / ptToPx,
      ));
  const radarAxisPlan = planNumericValueAxis({
    dataMin,
    dataMax,
    explicitMin: chart.valMin,
    explicitMax: chart.valMax,
    axisLenPt: rd / ptToPx,
    axisOrientation: 'vertical',
    majorUnit: radarMajorUnit,
    minorUnit: chart.valAxisMinorUnit,
    needMinor: chart.valAxisMinorGridlines === true || needsMinorTicks,
    logBase: chart.valAxisLogBase,
    reversed: valAxisReversed(chart),
  });
  const radarFrac = (value: number): number => clamp(radarAxisPlan.fraction(value), 0, 1);

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  // Rings sit on the value-axis MAJOR ticks — i.e. at value `ri * step`, whose
  // radius is proportional to the value (`v / axMax`). Deriving the radius from
  // the value (not `ri / rings`) keeps the rings on the major-unit multiples
  // even when `axMax` is not an exact multiple of `step` (e.g. an explicit
  // `<c:majorUnit>` §21.2.2.103 that doesn't divide the auto-rounded max).
  const ringValues = radarAxisPlan.majorTicks.filter(value => radarFrac(value) > 0);
  const strokeRing = (value: number): void => {
    const rr = radarFrac(value) * rd;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  };
  if (chart.valAxisMinorGridlines) {
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    ctx.strokeStyle = minorGrid.color;
    ctx.lineWidth = minorGrid.width;
    const previousDash = minorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (minorGrid.dash.length > 0) ctx.setLineDash(minorGrid.dash);
    for (const value of radarAxisPlan.minorTicks) strokeRing(value);
    if (minorGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }
  if (!chart.valAxisHidden && drawValMajorGridlines(chart)) {
    const majorGrid = valGridStroke(chart, ptToPx);
    ctx.strokeStyle = majorGrid.color;
    ctx.lineWidth = majorGrid.width;
    const previousDash = majorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (majorGrid.dash.length > 0) ctx.setLineDash(majorGrid.dash);
    for (const ringValue of ringValues) strokeRing(ringValue);
    if (majorGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }

  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * rd, cy2 + Math.sin(a) * rd); ctx.stroke();
  }

  // Radial tick labels on the top (12 o'clock) spoke — Excel places the value
  // axis there for radar charts. Respect <c:valAx><c:delete val="1"/> when the
  // caller hides the axis, and skip the 0-label at the center to avoid
  // overlapping the origin point.
  if (!chart.valAxisHidden) {
    const valAxPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
    ctx.font = chartFontCss(
      valAxPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (const v of ringValues) {
      const rr = radarFrac(v) * rd;
      const y = cy2 - rr;
      const valAxisLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
      drawAxisTick(
        ctx, chart.valAxisMajorTickMark, 'val', cx2, y,
        valAxisLine.color, valAxisLine.width, false, chart.valAxisLineHidden, 'major', ptToPx,
        chart.valAxisLineDash,
      );
      if (chart.valAxisTickLabelPos !== 'none') {
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), cx2 - 3, y);
      }
    }
    if (needsMinorTicks) {
      const valAxisLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
      for (const value of radarAxisPlan.minorTicks) {
        drawAxisTick(
          ctx,
          chart.valAxisMinorTickMark,
          'val',
          cx2,
          cy2 - radarFrac(value) * rd,
          valAxisLine.color,
          valAxisLine.width,
          false,
          chart.valAxisLineHidden,
          'minor',
          ptToPx,
          chart.valAxisLineDash,
        );
      }
    }
  }

  const radarCatFontPx = chart.catAxisFontSizeHpt != null
    ? axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx)
    : Math.max(8, Math.min(11, rd * 0.2));
  ctx.font = chartFontCss(
    radarCatFontPx,
    chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#444';
  ctx.textBaseline = 'middle';
  // Spoke labels radiate from just outside the ring. Cap each at the room
  // between its anchor and the nearest horizontal plot edge so long category
  // names are elided instead of overrunning the chart frame. Left/right-aligned
  // labels extend toward one edge; centered (top/bottom) labels straddle the
  // anchor, so give them twice the smaller side.
  const plotLeftX = cx2 - pw / 2;
  const plotRightX = cx2 + pw / 2;
  if (!chart.catAxisHidden && catLabelsVisible(chart)) for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const radialLabelOffset = categoryLabelOffsetPx(12, chart.catAxisLabelOffsetPercent);
    const lx = cx2 + Math.cos(a) * (rd + radialLabelOffset);
    const ly = cy2 + Math.sin(a) * (rd + radialLabelOffset);
    const authoredAlignment = chart.catAxisLabelAlignment;
    const align: CanvasTextAlign = authoredAlignment === 'l'
      ? 'left'
      : authoredAlignment === 'r'
        ? 'right'
        : authoredAlignment === 'ctr'
          ? 'center'
          : Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.textAlign = align;
    const maxPx =
      align === 'right' ? lx - plotLeftX
        : align === 'left' ? plotRightX - lx
          : 2 * Math.min(plotRightX - lx, lx - plotLeftX);
    // §21.2.2.71: format numeric-serial categories via the category-axis
    // numFmt; string spoke labels pass through unchanged.
    const label = formatCategoryLabel((cats[i] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
    ctx.fillText(elideToWidth(ctx, label, maxPx), lx, ly);
  }

  // ECMA-376 §21.2.3.10 c:radarStyle — "filled" closes the polygon with a
  // translucent area fill; "standard" / "marker" (and default) draw the
  // line only. Markers come from per-series `<c:marker>` (which can
  // override the chart-type style by setting `<c:symbol val="none"/>`);
  // A chart may set radarStyle="marker" while every series carries
  // `<c:marker><c:symbol val="none"/>`, in which case Office draws
  // lines only — no dots.
  const filled = markersSuppressedByChartStyle(
    'radar', chart.chartType, chart.scatterStyle, chart.radarStyle,
  );
  const markerRadius = Math.max(2, rd * 0.025);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    // Build the per-spoke point list, leaving holes where the series has
    // no value (`<c:val>` ptCount > pts implies missing indices), so Office draws an open polyline
    // from idx 1 to idx 10 without bridging back through the top spoke).
    const pts: Array<[number, number] | null> = [];
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v == null) { pts.push(null); continue; }
      const frac = radarFrac(v);
      const a = spoke(i);
      pts.push([cx2 + Math.cos(a) * rd * frac, cy2 + Math.sin(a) * rd * frac]);
    }

    // Stroke the polyline, breaking on holes (no synthetic 0-fill).
    ctx.beginPath();
    let pen = false;
    for (const pt of pts) {
      if (pt == null) { pen = false; continue; }
      if (!pen) { ctx.moveTo(pt[0], pt[1]); pen = true; }
      else { ctx.lineTo(pt[0], pt[1]); }
    }
    // Only close the polygon when there are no gaps. With a hole anywhere
    // the radar is an open path (matches Excel's "skip missing point").
    const allPresent = pts.every(p => p != null);
    if (filled && allPresent) {
      ctx.closePath();
      ctx.fillStyle = hexToRgba(color, 0.25); ctx.fill();
    } else if (allPresent) {
      ctx.closePath();
    }
    if (s.lineHidden !== true) {
      const previousDash = ctx.getLineDash ? ctx.getLineDash() : [];
      const previousCap = ctx.lineCap;
      const previousJoin = ctx.lineJoin;
      ctx.strokeStyle = s.lineColor ? `#${s.lineColor}` : color;
      ctx.lineWidth = s.lineWidthEmu != null
        ? axisLineWidthPx(s.lineWidthEmu, ptToPx)
        : 2;
      ctx.setLineDash(dashPatternForPreset(s.chartexStyle?.lineDash ?? undefined, ctx.lineWidth));
      ctx.lineCap = s.chartexStyle?.lineCap === 'rnd'
        ? 'round'
        : s.chartexStyle?.lineCap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = s.chartexStyle?.lineJoin === 'round' || s.chartexStyle?.lineJoin === 'bevel'
        ? s.chartexStyle.lineJoin
        : 'miter';
      ctx.stroke();
      ctx.setLineDash(previousDash);
      ctx.lineCap = previousCap;
      ctx.lineJoin = previousJoin;
    }

    // Markers: honor the per-series marker_symbol. When the series
    // explicitly carries `<c:marker><c:symbol val="none"/>`, the parser
    // sets showMarker=false — respect that even for radarStyle="marker"
    // charts (the chart-level style is the default; series overrides win).
    const seriesMarkersVisible = !filled && s.showMarker !== false && s.markerSymbol !== 'none';
    if (!filled && (seriesMarkersVisible || hasVisiblePointMarkerOverride(s))) {
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      for (let pointIndex = 0; pointIndex < pts.length; pointIndex++) {
        const pt = pts[pointIndex];
        if (pt == null) continue;
        const point = pointOverrides.get(pointIndex);
        const symbol = effectiveMarkerSymbol(s, point, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        const size = point?.markerSize ?? s.markerSize ?? Math.max(4, markerRadius * 2 / ptToPx);
        const fill = markerFillColorFor(s, point, pointIndex, color);
        const line = point?.markerLine ?? s.markerLine ?? null;
        const lineWidth = point?.markerLineWidthEmu ?? s.markerLineWidthEmu;
        drawMarker(
          ctx, pt[0], pt[1], symbol, size, fill, line, ptToPx,
          lineWidth != null ? axisLineWidthPx(lineWidth, ptToPx) : 1,
          markerFillPaintFor(s, point, pointIndex), shapeRotationDeg,
        );
      }
    }
  }

  drawLegendForLayout(
    ctx, chart, leg,
    x, y, w, h,
    plotLeft, plotTop, pw, ph, frame.title.bandH + 2,
    ptToPx,
  );
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
function scatterXValue(cats: string[], index: number, useIndexX: boolean): number | null {
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
type BubbleGroupSettings = Pick<
  ChartModel, 'bubbleScale' | 'bubbleSizeRepresents' | 'showNegativeBubbles'
>;

function bubbleSizeMagnitude(chart: BubbleGroupSettings, value: number): number {
  return chart.bubbleSizeRepresents === 'w' ? value : Math.sqrt(value);
}

type ScatterSeriesLayer = {
  series: ChartSeries;
  seriesIndex: number;
  fallbackColor: string;
  cats: string[];
  pointOverrides: Map<number, NonNullable<ChartSeries['dataPointOverrides']>[number]>;
};

function scatterPointFill(
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
function bubblePointFill(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  index: number,
  fallbackColor: string,
  bubble3D = bubblePointIsThreeD(series, point),
): { color: string; paint: Fill | null | undefined } {
  const bubbleSize = series.bubbleSizes?.[index];
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
  const directPoint = chartExStylePaintDecision(
    chart, point?.chartexStyle, index, series.values.length,
  );
  if (directPoint !== undefined) {
    return {
      color: directPoint?.fillType === 'solid' ? directPoint.color : fallbackColor,
      paint: directPoint,
    };
  }
  if (point?.fillHidden === true) return { color: '00000000', paint: null };
  if (point?.color != null) return { color: point.color, paint: undefined };
  const pointColor = series.dataPointColors?.[index];
  if (pointColor != null) return { color: pointColor, paint: undefined };

  const directSeries = chartExStylePaintDecision(
    chart, series.chartexStyle, index, series.values.length,
  );
  if (directSeries !== undefined) {
    return {
      color: directSeries?.fillType === 'solid' ? directSeries.color : fallbackColor,
      paint: directSeries,
    };
  }
  if (series.color != null) return { color: series.color, paint: undefined };
  const linkedPoint = chartExStylePaintDecision(
    chart, chart.chartStyleRoles?.dataPoint, index, series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : fallbackColor,
      paint: linkedPoint,
    };
  }
  return {
    color: scatterPointFill(series, point, index, fallbackColor),
    paint: markerFillPaintFor(series, point, index),
  };
}

function bubblePointLine(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  index: number,
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
  const linkedStyle = chart.chartStyleRoles?.dataPoint;
  const linkedGeometry = linkedStyle?.lineNoStyle === true ? undefined : linkedStyle;
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
  const pointPaint = chartExStyleLinePaintDecision(
    chart, pointStyle, index, series.values.length,
  );
  if (pointPaint !== undefined) {
    return {
      color: pointPaint?.fillType === 'solid' ? pointPaint.color : point?.lineColor ?? null,
      paint: pointPaint,
      ...geometry,
    };
  }
  if (point?.lineHidden === true) {
    return { color: null, paint: null, ...geometry };
  }
  if (point?.lineColor != null) {
    return { color: point.lineColor, paint: undefined, ...geometry };
  }

  const seriesPaint = chartExStyleLinePaintDecision(
    chart, seriesStyle, index, series.values.length,
  );
  if (seriesPaint !== undefined) {
    return {
      color: seriesPaint?.fillType === 'solid' ? seriesPaint.color : series.lineColor ?? null,
      paint: seriesPaint,
      ...geometry,
    };
  }
  if (series.lineHidden === true) {
    return { color: null, paint: null, ...geometry };
  }
  if (series.lineColor != null) {
    return { color: series.lineColor, paint: undefined, ...geometry };
  }
  const linkedPoint = chartExStyleLinePaintDecision(
    chart, linkedStyle, index, series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : null,
      paint: linkedPoint,
      ...geometry,
    };
  }
  const bubbleSize = series.bubbleSizes?.[index];
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

function makeScatterSeriesLayer(
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
function bubbleSizeToDiameterScale(
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
function drawScatterSeriesLayer(
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

  for (const { series: s, fallbackColor, cats } of layers) {
    const automaticPointStyle = style === 'marker'
      && hasFilteredScatterAutomaticPointStyle(s);
    const drawLines = automaticPointStyle || groupDrawsLines;
    const drawSmooth = (!automaticPointStyle && style === 'marker' && !isBubble)
      || groupDrawsSmooth;
    if ((drawLines || drawSmooth) && s.lineHidden !== true) {
      const pts: Array<{ x: number; y: number }> = [];
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci];
        if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        pts.push({ x: toX(xv), y: toY(yv) });
      }
      if (pts.length >= 2) {
        ctx.save();
        if (automaticPointStyle && s.dataPointColors?.some(Boolean)) {
          ctx.lineWidth = 1.5;
          for (let i = 1; i < pts.length; i++) {
            ctx.strokeStyle = `#${s.dataPointColors[i] ?? s.color ?? fallbackColor.replace(/^#/, '')}`;
            ctx.beginPath();
            ctx.moveTo(pts[i - 1].x, pts[i - 1].y);
            ctx.lineTo(pts[i].x, pts[i].y);
            ctx.stroke();
          }
        } else {
          ctx.strokeStyle = s.color ? `#${s.color}` : fallbackColor;
          ctx.lineWidth = 1.5;
          ctx.beginPath();
          ctx.moveTo(pts[0].x, pts[0].y);
          if (drawSmooth && pts.length >= 3) {
            for (let i = 0; i < pts.length - 1; i++) {
              const p0 = pts[i - 1] ?? pts[i];
              const p1 = pts[i];
              const p2 = pts[i + 1];
              const p3 = pts[i + 2] ?? p2;
              ctx.bezierCurveTo(
                p1.x + (p2.x - p0.x) / 6,
                p1.y + (p2.y - p0.y) / 6,
                p2.x - (p3.x - p1.x) / 6,
                p2.y - (p3.y - p1.y) / 6,
                p2.x,
                p2.y,
              );
            }
          } else {
            for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
          }
          ctx.stroke();
        }
        ctx.restore();
      }
    }
  }

  for (const { series: s, fallbackColor, cats, pointOverrides } of layers) {
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
          ? bubblePointFill(chart, s, dpt, ci, fallbackColor)
          : null;
        const fill = bubbleFill?.color ?? scatterPointFill(s, dpt, ci, fallbackColor);
        // Bubble geometry is the series shape itself, so its outline comes from
        // `<c:ser><c:spPr><a:ln>` rather than a `<c:marker>` block. Ordinary
        // scatter markers continue to use markerLine only.
        const bubbleLine = isBubble ? bubblePointLine(chart, s, dpt, ci) : null;
        const line = isBubble
          ? bubbleLine!.color
          : dpt?.markerLine ?? s.markerLine ?? null;
        const markerLineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
        const bubbleLineWidthEmu = bubbleLine?.widthEmu;
        const lineWidthEmu = isBubble ? bubbleLineWidthEmu : markerLineWidthEmu;
        const lineWidthPx = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx)
          : undefined;
        drawMarker(
          ctx, toX(xv), toY(yv), symbol, sizePt, fill, line, ptToPx, lineWidthPx,
          isBubble ? bubbleFill!.paint : markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
          isBubble ? bubbleLine!.paint : undefined,
          isBubble ? bubbleLine!.dash : undefined,
          isBubble ? bubbleLine!.customDash : undefined,
          isBubble ? bubbleLine!.cap : undefined,
          isBubble ? bubbleLine!.join : undefined,
          isBubble ? bubblePointIsThreeD(s, dpt) : false,
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

function renderScatterChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const entries = chart.series.map((series, index) => ({ series, index }));
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const usesSecondaryX = ({ series, index }: (typeof entries)[number]): boolean =>
    plotGroupBySeries[index]?.categoryAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  const usesSecondaryY = ({ series, index }: (typeof entries)[number]): boolean =>
    plotGroupBySeries[index]?.valueAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  const primaryXEntries = entries.filter(entry => !usesSecondaryX(entry));
  const secondaryXEntries = entries.filter(usesSecondaryX);
  const primaryYEntries = entries.filter(entry => !usesSecondaryY(entry));
  const secondaryYEntries = entries.filter(usesSecondaryY);
  const primaryEntries = entries.filter(entry => !usesSecondaryX(entry) && !usesSecondaryY(entry));
  const secondaryEntries = entries.filter(entry => usesSecondaryX(entry) && usesSecondaryY(entry));
  const secondaryX = secondaryXEntries.length > 0 ? chart.secondaryCatAxis : null;
  const secondaryY = secondaryYEntries.length > 0 ? chart.secondaryValAxis : null;

  const numericXValues = (entries: Array<{ series: ChartSeries; index: number }>): number[] => {
    const values: number[] = [];
    for (const { series } of entries) {
      const cats = series.categories ?? chart.categories;
      for (const category of cats) {
        const value = parseFloat(category);
        if (Number.isFinite(value)) values.push(value);
      }
    }
    return values;
  };
  const allNumericX = numericXValues(entries);
  const useIndexX = allNumericX.length === 0;
  const textBubbleOrdinalMax = entries.length === 1
    && entries[0].series.bubbleXSourceIsString === true
    ? entries[0].series.values.length + 1
    : null;
  const pairedExtents = (
    entries: Array<{ series: ChartSeries; index: number }>,
  ): { x: { min: number; max: number }; y: { min: number; max: number } } => {
    const xs: number[] = [];
    const ys: number[] = [];
    for (const { series } of entries) {
      const cats = series.categories ?? chart.categories;
      for (let index = 0; index < series.values.length; index++) {
        const yValue = series.values[index];
        if (yValue == null) continue;
        const xValue = scatterXValue(cats, index, useIndexX);
        if (xValue == null) continue;
        xs.push(xValue);
        ys.push(yValue);
      }
      forEachErrorBarEndpoint(
        series,
        'x',
        index => series.values[index] == null ? null : scatterXValue(cats, index, useIndexX),
        value => xs.push(value),
      );
      forEachErrorBarEndpoint(
        series,
        'y',
        index => {
          const xValue = scatterXValue(cats, index, useIndexX);
          return xValue == null ? null : series.values[index] ?? null;
        },
        value => ys.push(value),
      );
    }
    if (useIndexX && xs.length === 0) {
      let count = 0;
      for (const { series } of entries) count = Math.max(count, series.values.length);
      for (let index = 0; index < count; index++) xs.push(index);
    }
    return { x: finiteDataExtent(xs), y: finiteDataExtent(ys) };
  };
  const primaryExtent = {
    x: pairedExtents(primaryXEntries.length > 0 ? primaryXEntries : secondaryXEntries).x,
    y: pairedExtents(primaryYEntries.length > 0 ? primaryYEntries : secondaryYEntries).y,
  };
  const secondaryExtent = {
    x: pairedExtents(secondaryXEntries).x,
    y: pairedExtents(secondaryYEntries).y,
  };
  // Shared frame bands. Title + bottom axis-label bands follow PowerPoint's
  // chart auto-layout (font-proportional, pinned to the demo slide-5 line-chart
  // PDF); see cartesianTitleBand / catAxisLabelBandH in layout.ts. Scatter's X
  // axis is a numeric value axis, so the bottom band holds its single line of
  // X-value labels (sized like any value-axis label). Default 0.22 side-legend
  // reserve unchanged.
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const xAxLabelFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const yAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  // Title placement — manual layout overrides the auto position.
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  // Plot area placement: honor `<c:plotArea><c:manualLayout>` when present.
  // ECMA-376: layoutTarget="inner" (default) describes the inner plot rect
  // (no axes / labels); "outer" includes axes. For scatter we treat both
  // identically (the inner padding stays the same). The pad is pure arithmetic
  // and is ignored by computeChartFrame when the manual layout applies.
  const provisionalSecondaryY = secondaryY
    ? planNumericValueAxis({
        dataMin: secondaryExtent.y.min,
        dataMax: secondaryExtent.y.max,
        explicitMin: secondaryY.min,
        explicitMax: secondaryY.max,
        axisLenPt: Math.max(1, h * 0.7 / ptToPx),
        axisOrientation: 'vertical',
        majorUnit: secondaryY.majorUnit,
        minorUnit: secondaryY.minorUnit,
        needMinor: secondaryY.minorGridlines === true
          || (secondaryY.minorTickMark != null && secondaryY.minorTickMark !== 'none'),
        logBase: secondaryY.logBase,
        reversed: secondaryY.orientation === 'maxMin',
      })
    : null;
  let secondaryYLabelWidth = 0;
  if (secondaryY && provisionalSecondaryY && !secondaryY.hidden && secondaryY.tickLabelPos !== 'none') {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      chartTextFontSizePx(secondaryY.fontSizeHpt, ptToPx) ?? yAxLabelFontPx,
      chartFontFamily(chart, secondaryY.fontFace, 'minor'),
      secondaryY.fontBold ?? false,
      secondaryY.fontItalic ?? false,
    );
    for (const value of provisionalSecondaryY.majorTicks) {
      secondaryYLabelWidth = Math.max(
        secondaryYLabelWidth,
        ctx.measureText(formatAxisTickWithUnits(value, secondaryY.formatCode, chart.date1904, secondaryY.displayUnits)).width,
      );
    }
    secondaryYLabelWidth += valueTickLabelGapPx(yAxLabelFontPx) + 4;
    ctx.font = previousFont;
  }
  const secondaryXLabelHeight = secondaryX && !secondaryX.hidden && secondaryX.tickLabelPos !== 'none'
    ? (chartTextFontSizePx(secondaryX.fontSizeHpt, ptToPx) ?? xAxLabelFontPx)
      + categoryTickLabelGapPx(xAxLabelFontPx) + 2
    : 0;
  const pad = {
    t: titleBand.bandH + legTopH + yAxLabelFontPx / 2 + 2 + secondaryXLabelHeight,
    r: legRightW + w * 0.05 + secondaryYLabelWidth,
    b: (chart.catAxisHidden ? h * 0.04 : catAxisLabelBandH(xAxLabelFontPx)) + catTitleH + legBottomH,
    l: (chart.valAxisHidden ? w * 0.04 : w * 0.12) + valTitleW + legLeftW,
  };
  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  if (pw <= 0 || ph <= 0) return;

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // X / Y data extents. Secondary-group points have their own independent
  // top/right value axes and therefore must not stretch the primary scales.
  let { min: xMin, max: xMax } = primaryExtent.x;
  let { min: yMin, max: yMax } = primaryExtent.y;
  // Apply explicit `<c:valAx><c:scaling><c:min/max>` and `<c:catAx>` scaling.
  // Omitted Y bounds, including point ranges, flow through the shared planner.
  if (chart.valMin != null) yMin = chart.valMin;
  if (chart.valMax != null) yMax = chart.valMax;
  const yNeedsMinor = chart.valAxisMinorGridlines === true
    || (chart.valAxisMinorTickMark != null && chart.valAxisMinorTickMark !== 'none');
  const yAxisPlan = planNumericValueAxis({
    dataMin: yMin,
    dataMax: yMax,
    explicitMin: chart.valMin,
    explicitMax: chart.valMax,
    axisLenPt: ph / ptToPx,
    axisOrientation: 'vertical',
    majorUnit: chart.valAxisMajorUnit,
    minorUnit: chart.valAxisMinorUnit,
    needMinor: yNeedsMinor,
    logBase: chart.valAxisLogBase,
    reversed: valAxisReversed(chart),
  });
  yMin = yAxisPlan.min;
  yMax = yAxisPlan.max;
  const xNeedsMinor = chart.catAxisMinorGridlines === true
    || (chart.catAxisMinorTickMark != null && chart.catAxisMinorTickMark !== 'none');
  const xAxisPlan = planNumericValueAxis({
    dataMin: xMin,
    dataMax: xMax,
    // Office gives a string-backed lone bubble series one empty ordinal slot
    // on each side (four points => 0..5). Authored axis bounds still win.
    explicitMin: chart.catAxisMin ?? (textBubbleOrdinalMax == null ? null : 0),
    explicitMax: chart.catAxisMax ?? textBubbleOrdinalMax,
    axisLenPt: pw / ptToPx,
    axisOrientation: 'horizontal',
    majorUnit: chart.catAxisMajorUnit,
    minorUnit: chart.catAxisMinorUnit,
    needMinor: xNeedsMinor,
    logBase: chart.catAxisLogBase,
    reversed: catAxisReversed(chart),
  });
  xMin = xAxisPlan.min;
  xMax = xAxisPlan.max;

  const secondaryXPlan = secondaryX
    ? planNumericValueAxis({
        dataMin: secondaryExtent.x.min,
        dataMax: secondaryExtent.x.max,
        explicitMin: secondaryX.min,
        explicitMax: secondaryX.max,
        axisLenPt: pw / ptToPx,
        axisOrientation: 'horizontal',
        majorUnit: secondaryX.majorUnit,
        minorUnit: secondaryX.minorUnit,
        needMinor: secondaryX.minorGridlines === true
          || (secondaryX.minorTickMark != null && secondaryX.minorTickMark !== 'none'),
        logBase: secondaryX.logBase,
        reversed: secondaryX.orientation === 'maxMin',
      })
    : null;
  const secondaryYPlan = secondaryY
    ? planNumericValueAxis({
        dataMin: secondaryExtent.y.min,
        dataMax: secondaryExtent.y.max,
        explicitMin: secondaryY.min,
        explicitMax: secondaryY.max,
        axisLenPt: ph / ptToPx,
        axisOrientation: 'vertical',
        majorUnit: secondaryY.majorUnit,
        minorUnit: secondaryY.minorUnit,
        needMinor: secondaryY.minorGridlines === true
          || (secondaryY.minorTickMark != null && secondaryY.minorTickMark !== 'none'),
        logBase: secondaryY.logBase,
        reversed: secondaryY.orientation === 'maxMin',
      })
    : null;

  const toX = (v: number) => px0 + xAxisPlan.fraction(v) * pw;
  const toY = (v: number) => py0 + ph - yAxisPlan.fraction(v) * ph;
  const toSecondaryX = (v: number) => px0 + (secondaryXPlan?.fraction(v) ?? 0) * pw;
  const toSecondaryY = (v: number) => py0 + ph - (secondaryYPlan?.fraction(v) ?? 0) * ph;
  const xStep = xAxisPlan.majorUnit;
  const yMajorTicks = yAxisPlan.majorTicks;
  const yMinorTicks = yAxisPlan.minorTicks;
  const xMajorTicks = xAxisPlan.majorTicks;
  const xMinorTicks = xAxisPlan.minorTicks;

  // Each scatter axis is a numeric value axis. Its crossing coordinate comes
  // from the opposite axis's scale (§21.2.2.31 / §21.2.2.32): autoZero uses
  // zero when the range contains it, while min/max pin the rule to an edge.
  let xAxisY = py0 + ph;
  if (chart.catAxisCrossesAt != null) {
    xAxisY = clamp(toY(chart.catAxisCrossesAt), py0, py0 + ph);
  } else {
    const crosses = chart.catAxisCrosses ?? 'autoZero';
    if (crosses === 'autoZero' && yMin < 0 && yMax > 0) xAxisY = clamp(toY(0), py0, py0 + ph);
    else if (crosses === 'max') xAxisY = py0;
  }

  let yAxisX = px0;
  if (chart.valAxisCrossesAt != null) {
    yAxisX = clamp(toX(chart.valAxisCrossesAt), px0, px0 + pw);
  } else {
    const crosses = chart.valAxisCrosses ?? 'autoZero';
    if (crosses === 'autoZero' && xMin < 0 && xMax > 0) yAxisX = clamp(toX(0), px0, px0 + pw);
    else if (crosses === 'max') yAxisX = px0 + pw;
  }

  // Y-axis gridlines + labels + major tick marks. Scatter has no baseline
  // special-case, so it strokes every gridline in the resolved color/width.
  const grid = valGridStroke(chart, ptToPx);
  if (!chart.valAxisHidden) {
    const yTickFontPx = chart.valAxisFontSizeHpt != null
      ? axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx)
      : Math.max(8, Math.min(11, ph / 20));
    const yTickGap = chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(yTickFontPx)
      : 4;
    ctx.font = chartFontCss(
      yTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    const yAxisLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : undefined;
    const yAxisLineWidth = axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx);
    const yMajorTickOutset = chart.valAxisLineHidden
      ? 0
      : axisTickOutwardExtentPx(chart.valAxisMajorTickMark, 'major', yAxisLineWidth, ptToPx);
    if (chart.valAxisMinorGridlines) {
      const minorGrid = valMinorGridStroke(chart, ptToPx);
      for (const value of yMinorTicks) {
        strokeValueGridlineH(ctx, px0, pw, toY(value), false, minorGrid);
      }
    }
    for (const v of yMajorTicks) {
      const gy = toY(v);
      ctx.strokeStyle = grid.color; ctx.lineWidth = grid.width;
      if (drawValMajorGridlines(chart)) {
        const previousDash = grid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (grid.dash.length > 0) ctx.setLineDash(grid.dash);
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        if (grid.dash.length > 0) ctx.setLineDash(previousDash);
      }
      if (chart.valAxisTickLabelPos !== 'none') {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        const labelPos = chart.valAxisTickLabelPos ?? 'nextTo';
        let labelX: number;
        if (labelPos === 'high') {
          ctx.textAlign = 'left'; labelX = px0 + pw + yTickGap;
        } else if (labelPos === 'low') {
          ctx.textAlign = 'right'; labelX = px0 - yTickGap;
        } else {
          ctx.textAlign = 'right'; labelX = yAxisX - yMajorTickOutset - yTickGap;
        }
        ctx.textBaseline = 'middle';
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), labelX, gy);
      }
      // Scatter keeps its own undefined colour default (→ drawAxisTick's '#888'),
      // so only the width formula is shared. `axisLineWidthPx`'s 1 px fallback is
      // equivalent to undefined here (drawAxisTick treats both as a hairline).
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', yAxisX, gy, yAxisLineColor, yAxisLineWidth, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of yMinorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', yAxisX, toY(value), yAxisLineColor, yAxisLineWidth, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }

  // A scatter chart's horizontal axis is represented by the shared category-
  // axis fields in the model even though OOXML stores it as a second valAx.
  // Its major gridlines therefore run vertically through each numeric X tick.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart) && xStep > 0) {
    const xGrid = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = xGrid.color;
    ctx.lineWidth = xGrid.width;
    const previousDash = xGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (xGrid.dash.length > 0) ctx.setLineDash(xGrid.dash);
    for (const v of xMajorTicks) {
      const gx = toX(v);
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (xGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }
  if (!chart.catAxisHidden && chart.catAxisMinorGridlines && xStep > 0) {
    const xGrid = catMinorGridStroke(chart, ptToPx);
    const previousDash = xGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    ctx.strokeStyle = xGrid.color;
    ctx.lineWidth = xGrid.width;
    if (xGrid.dash.length > 0) ctx.setLineDash(xGrid.dash);
    for (const value of xMinorTicks) {
      const gx = toX(value);
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (xGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // X-axis line (the timeline ruler in Gantt-style scatter charts depends
  // on this line's stroke). Tick labels are skipped when the category axis
  // is hidden via `<c:delete val="1"/>`. Office treats
  // `<c:catAx><c:spPr><a:ln><a:noFill>` as suppressing the rule and tick
  // marks. Color and weight come from
  // `<c:catAx><c:spPr><a:ln>` when present; default otherwise.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.save();
    ctx.lineCap = 'butt';
    strokeAxisSegment(
      ctx, px0, xAxisY, px0 + pw, xAxisY,
      chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#888',
      axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx), chart.catAxisLineDash,
    );
    ctx.restore();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.save();
    strokeAxisSegment(
      ctx, yAxisX, py0, yAxisX, py0 + ph,
      chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#888',
      axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx), chart.valAxisLineDash,
    );
    ctx.restore();
  }

  // X-axis tick labels (catAxis), formatted via catAxisFormatCode (typically
  // a date code like "m/d/yyyy"). Skipped when catAxisHidden. Drawn just
  // at the authored high/low plot edge or next to the crossing axis. Major
  // tick marks remain attached to the axis rule so `<c:majorTickMark val="cross">` produces
  // the crossing ruler look that templates like the Vertex42 timeline
  // depend on.
  if (!chart.catAxisHidden) {
    const tickFontPx = chart.catAxisFontSizeHpt != null
      ? axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx)
      : Math.max(8, Math.min(11, ph / 20));
    const tickGap = chart.catAxisFontSizeHpt != null
      ? categoryTickLabelGapPx(tickFontPx)
      : 4;
    ctx.font = chartFontCss(
      tickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center';
    const labelPos = chart.catAxisTickLabelPos ?? 'nextTo';
    const lineWidth = axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx);
    const xAxisLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : undefined;
    const tickOutset = chart.catAxisLineHidden
      ? 0
      : axisTickOutwardExtentPx(chart.catAxisMajorTickMark, 'major', lineWidth, ptToPx);
    const labelY = labelPos === 'low'
      ? py0 + ph + tickGap
      : labelPos === 'high' ? py0 - tickGap : xAxisY + tickOutset + tickGap;
    ctx.textBaseline = labelPos === 'high' ? 'bottom' : 'top';
    for (const v of xMajorTicks) {
      const gx = toX(v);
      if (labelPos !== 'none') {
        ctx.fillText(formatAxisTickWithUnits(v, chart.catAxisFormatCode, chart.date1904, chart.catAxisDisplayUnits), gx, labelY);
      }
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', xAxisY, gx, xAxisLineColor, lineWidth, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none') {
      for (const value of xMinorTicks) {
        drawAxisTick(ctx, chart.catAxisMinorTickMark, 'cat', xAxisY, toX(value), xAxisLineColor, lineWidth, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
      }
    }
  }

  // Office preserves scatter/bubble group order for overlapping geometry. A
  // group still paints its own line/error/marker/label phases, but the next
  // source group is composited after it. This differs from bar/area/line,
  // whose cross-family layering is application-defined and handled by their
  // dedicated combo path.
  const drawNumericEntries = (
    entries: Array<{ series: ChartSeries; index: number }>,
    isBubble: boolean,
    style: string,
    xMap: (value: number) => number,
    yMap: (value: number) => number,
    maximum: number,
    displayUnits?: ChartDisplayUnits | null,
    bubbleSettings?: BubbleGroupSettings,
  ): void => {
    if (entries.length === 0) return;
    drawScatterSeriesLayer(
      ctx, chart, entries, useIndexX, xMap, yMap, r,
      px0, py0, pw, ph, ptToPx, isBubble, style,
      { x, y, w, h }, maximum, displayUnits, shapeRotationDeg, bubbleSettings,
    );
  };
  if (chart.plotGroups == null) {
    const isBubble = chart.chartType === 'bubble';
    const style = isBubble ? 'marker' : (chart.scatterStyle ?? 'marker');
    drawNumericEntries(
      primaryEntries, isBubble, style, toX, toY,
      yAxisPlan.max, chart.valAxisDisplayUnits,
    );
    if (secondaryEntries.length > 0 && secondaryXPlan && secondaryYPlan) {
      drawNumericEntries(
        secondaryEntries, isBubble, style, toSecondaryX, toSecondaryY,
        secondaryYPlan.max, secondaryY?.displayUnits,
      );
    }
  } else {
    for (const group of chart.plotGroups) {
      if (group.kind !== 'scatter' && group.kind !== 'bubble') continue;
      const entries = chart.series
        .slice(group.seriesStart, group.seriesStart + group.seriesCount)
        .map((series, offset) => ({ series, index: group.seriesStart + offset }));
      if (entries.length === 0) continue;
      const isBubble = group.kind === 'bubble';
      const usesSecondaryX = group.categoryAxis === 'secondary';
      const usesSecondaryY = group.valueAxis === 'secondary';
      drawNumericEntries(
        entries, isBubble,
        isBubble ? 'marker' : (group.scatterStyle ?? chart.scatterStyle ?? 'marker'),
        usesSecondaryX ? toSecondaryX : toX,
        usesSecondaryY ? toSecondaryY : toY,
        usesSecondaryY && secondaryYPlan ? secondaryYPlan.max : yAxisPlan.max,
        usesSecondaryY ? secondaryY?.displayUnits : chart.valAxisDisplayUnits,
        isBubble ? {
          bubbleScale: group.bubbleScale ?? chart.bubbleScale,
          bubbleSizeRepresents: group.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
          showNegativeBubbles: group.showNegativeBubbles ?? chart.showNegativeBubbles,
        } : undefined,
      );
    }
  }

  // The second CT_ScatterChart group owns an independent top X and right Y
  // value-axis pair. Both axes use the same numeric planner and authored unit
  // formatting as the primary pair; only their screen edge differs.
  if (secondaryX && secondaryXPlan && !secondaryX.hidden) {
    const line = resolveAxisLine(secondaryX.lineColor, secondaryX.lineWidthEmu, ptToPx);
    if (!secondaryX.lineHidden) {
      strokeAxisSegment(
        ctx, px0, py0, px0 + pw, py0,
        line.color, line.width, secondaryX.lineDash,
      );
    }
    const fontPx = chartTextFontSizePx(secondaryX.fontSizeHpt, ptToPx) ?? xAxLabelFontPx;
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, secondaryX.fontFace, 'minor'),
      secondaryX.fontBold ?? false,
      secondaryX.fontItalic ?? false,
    );
    ctx.fillStyle = secondaryX.fontColor ? `#${secondaryX.fontColor}` : '#555';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    const tickOutset = secondaryX.lineHidden
      ? 0
      : axisTickOutwardExtentPx(secondaryX.majorTickMark, 'major', line.width, ptToPx);
    for (const value of secondaryXPlan.majorTicks) {
      const sx = toSecondaryX(value);
      if (secondaryX.tickLabelPos !== 'none') {
        ctx.fillText(
          formatAxisTickWithUnits(value, secondaryX.formatCode, chart.date1904, secondaryX.displayUnits),
          sx,
          py0 - tickOutset - categoryTickLabelGapPx(fontPx),
        );
      }
      drawAxisTick(
        ctx, secondaryX.majorTickMark, 'cat', py0, sx, line.color, line.width,
        true, secondaryX.lineHidden, 'major', ptToPx, secondaryX.lineDash,
      );
    }
    if (secondaryX.minorTickMark && secondaryX.minorTickMark !== 'none') {
      for (const value of secondaryXPlan.minorTicks) {
        drawAxisTick(
          ctx, secondaryX.minorTickMark, 'cat', py0, toSecondaryX(value), line.color,
          line.width, true, secondaryX.lineHidden, 'minor', ptToPx,
          secondaryX.lineDash,
        );
      }
    }
  }
  if (secondaryY && secondaryYPlan) {
    const scale: SecondaryAxisScale = {
      min: secondaryYPlan.min,
      max: secondaryYPlan.max,
      step: secondaryYPlan.majorUnit,
      majorLines: secondaryYPlan.majorTicks,
      minorTicks: secondaryYPlan.minorTicks,
      makeToY: () => toSecondaryY,
    };
    drawSecondaryValueAxis(
      ctx, chart, secondaryY, scale, toSecondaryY, r,
      px0, py0, pw, ph, ptToPx,
      chartTextFontSizePx(secondaryY.fontSizeHpt, ptToPx) ?? yAxLabelFontPx,
      secondaryYLabelWidth,
      chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555',
      chart.date1904,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleBand.bandH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// Three fixed gradients with 4 + 5 + 6 stops. Keep the work count aligned
// with the complete material so a bubble is admitted or rejected atomically.
const BUBBLE_3D_MATERIAL_COMPONENTS = 15;

/** ECMA-376 §21.2.2.21 only enables `bubble3D`; it does not define a lighting
 * material. Paint the bounded application-defined material observed in desktop
 * Excel vector output. A single radial envelope cannot independently
 * express the diffuse highlight, right/lower falloff, and narrow lower
 * reflected-light band, so those three components are composited in order.
 * The recipe is normalized to bubble-local coordinates and is therefore shared
 * by every colour, size, and host transform rather than fitted per sample.
 * `source-atop` preserves the authored fill alpha on every pass. */
function paintBubble3DMaterial(
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
): void {
  const sizePx = Math.max(2, sizePt * ptToPx);
  const half = sizePx / 2;
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
function drawSeriesErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true) return;
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
function drawSeriesDataLabels(
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

function drawDataLabelText(
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
function customRichDataLabelOptions(
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
function drawBoundedDataLabelWithLegendKey(
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
    entry.outlineColor,
    entry.outlineWidthEmu,
    entry.outlineDash,
    entry.outlineCap,
    entry.outlineJoin,
    ptToPx,
    shapeRotationDeg,
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

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/** Append `pts` to the CURRENT path starting from `pts[0]` (which the caller has
 *  already `moveTo`'d, or the first point is the current pen position). When
 *  `smooth` and there are ≥3 points, draw a Catmull-Rom → cubic-Bézier curve
 *  through the points (tangents from neighbours, the same formula scatter uses,
 *  ECMA-376 §21.2.2.194); otherwise straight `lineTo` segments. The caller owns
 *  `beginPath`/`moveTo`/`stroke`/`fill` so this composes into both the line
 *  stroke and the area fill's top edge. */
function appendCurve(
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

function dashPatternForPreset(preset: string | undefined, lineWidth = 1): number[] {
  const scale = Number.isFinite(lineWidth) && lineWidth > 0 ? lineWidth : 1;
  return pptxPresetDashArray(preset ?? 'solid', scale);
}

function dashPatternForLine(
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
function drawCategoryErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  fallbackColor: string,
): void {
  if (eb.hidden === true || eb.dir === 'x') return; // no data-unit X scale on a category axis
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
function drawBarErrorBars(
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
  if (eb.hidden === true) return;
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
function drawCategoryDataLabels(
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

interface ResolvedChartExLabel {
  text: string;
  showLegendKey: boolean;
  position?: string;
  fontColor?: string;
  fontSizeHpt?: number;
  fontBold?: boolean;
  fontFace?: string;
  manualLayout?: ChartDataLabelOverride['manualLayout'];
  labelBox?: ChartLabelBox;
  richRuns?: ChartDataLabelOverride['richRuns'];
  textStyle: DataLabelTextStyle;
}

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
export function resolveChartExLabel(
  chart: ChartModel,
  series: ChartSeries | null | undefined,
  index: number,
  category: string,
  value: number,
  defaults: {
    visible: boolean;
    showVal: boolean;
    showCatName: boolean;
    showSerName?: boolean;
    showPercent?: boolean;
  },
  overrideLookup: ReadonlyMap<number, NonNullable<ChartSeries['dataLabelOverrides']>[number]>,
  valueOption?: boolean | number,
  valueDisplayUnits?: ChartDisplayUnits | null,
): ResolvedChartExLabel | null {
  if (!series) return null;
  const definition = series.seriesDataLabels;
  const override = overrideLookup.get(index);
  if (dataLabelIsDeleted(definition, override)) return null;
  if (!definition && !override && !defaults.visible) return null;
  const suppressValue = typeof valueOption === 'boolean' ? valueOption : false;
  const percentRatio = typeof valueOption === 'number' ? valueOption : undefined;
  const showVal = !suppressValue
    && (override?.showVal ?? definition?.showVal ?? defaults.showVal);
  const showCatName = override?.showCatName ?? definition?.showCatName ?? defaults.showCatName;
  const showSerName = override?.showSerName
    ?? definition?.showSerName
    ?? defaults.showSerName
    ?? false;
  const showPercent = override?.showPercent
    ?? definition?.showPercent
    ?? defaults.showPercent
    ?? false;
  const showLegendKey = override?.showLegendKey ?? definition?.showLegendKey ?? false;
  const authoredFormatCode = override?.formatCode
    ?? definition?.formatCode
    ?? chart.dataLabelFormatCode
    ?? null;
  const text = effectiveDataLabelText({
    customText: override?.text,
    showCategory: showCatName,
    showSeries: showSerName,
    showValue: showVal,
    showPercent,
    category,
    seriesName: series.name,
    sourceValue: value,
    valueDivisor: displayUnitDivisor(valueDisplayUnits),
    percentRatio,
    formatCode: authoredFormatCode ?? series.valFormatCode ?? null,
    percentFormatCode: authoredFormatCode ?? '0%',
    date1904: chart.date1904,
    separator: override?.separator ?? definition?.separator,
  });
  if (!text && !showLegendKey) return null;
  return {
    text,
    showLegendKey,
    position: override?.position ?? definition?.position,
    fontColor: override?.fontColor ?? definition?.fontColor,
    fontSizeHpt: override?.fontSizeHpt ?? definition?.fontSizeHpt,
    fontBold: override?.fontBold ?? definition?.fontBold,
    fontFace: override?.fontFace ?? definition?.fontFace,
    manualLayout: override?.manualLayout,
    labelBox: mergeChartLabelBoxes(override?.labelBox, definition?.labelBox),
    // Rich text is authoritative only for a non-empty custom point label.
    richRuns: override?.text ? override.richRuns : undefined,
    textStyle: effectiveDataLabelTextStyle(override, definition),
  };
}

export function chartExStyleColor(
  _chart: ChartModel,
  style: ChartExStyle | null | undefined,
  kind: 'fill' | 'line',
  index: number,
  _count: number,
): string | null {
  return chartStyleColor(style, kind, index);
}

function chartExPaletteColor(
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

function chartExSemanticFill(chart: ChartModel, index: number, count: number): string {
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

function chartExStyleFillPaint(
  style: ChartExStyle | null | undefined,
  index: number,
): Fill | null {
  return chartStyleFillPaint(style, index);
}

function chartExStyleLinePaint(
  style: ChartExStyle | null | undefined,
  index: number,
): ChartModel['plotAreaLineFill'] {
  return chartStyleLinePaint(style, index);
}

/** Line-paint counterpart of chartExStylePaintDecision. `undefined` means the
 * layer did not author a line paint, while `null` is an authored noFill or an
 * authored-but-unresolved paint that must suppress lower-precedence color. */
function chartExStyleLinePaintDecision(
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
function chartExStylePaintDecision(
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
  const local = chartExStylePaintDecision(chart, localStyle, index, count);
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
  const local = localStyle?.fillHidden
    ? undefined
    : chartExStylePaintDecision(chart, localStyle, index, count);
  if (local !== undefined) return local;
  if (localStyle && legacyColor) return { fillType: 'solid', color: legacyColor };
  const linked = chartExStylePaintDecision(chart, linkedStyle, index, count);
  if (linked !== undefined) return linked;
  if (legacyColor) return { fillType: 'solid', color: legacyColor };
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

interface ResolvedChartExLineStyle {
  visible: boolean;
  color: string;
  widthEmu: number | null;
  dash: string | null;
  cap: string | null;
  join: string | null;
}

type ChartExSeriesStyleCarrier = Pick<
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
  const resolved = (
    style: ChartExStyle | null | undefined,
    fallback: string,
  ): ResolvedChartExLineStyle => ({
    visible: style?.lineHidden !== true,
    color: chartExStyleColor(chart, style, 'line', index, count) ?? fallback,
    widthEmu: style?.lineWidthEmu ?? null,
    dash: style?.lineDash ?? null,
    cap: style?.lineCap ?? null,
    join: style?.lineJoin ?? null,
  });
  const local = series?.chartexStyle;
  const localHasLine = local?.lineHidden != null
    || local?.lineColors?.some(Boolean)
    || local?.lineWidthEmu != null
    || local?.lineDash != null
    || local?.lineCap != null
    || local?.lineJoin != null;
  if (localHasLine && !(local?.lineHidden && local.lineNoStyle)) {
    const line = resolved(local, series?.lineColor ?? fallbackColor);
    // Classic `<c:ser><c:spPr><a:ln>` keeps its legacy color/width fields for
    // the ordinary chart API while dash/cap/join share the structured carrier
    // introduced for ChartEx. Merge those authored properties instead of
    // letting the presence of a dash reset the same line to the 1px fallback.
    line.widthEmu ??= series?.lineWidthEmu ?? null;
    return line;
  }
  const hasLegacyLocalLine = series?.lineHidden != null
    || series?.lineColor != null
    || series?.lineWidthEmu != null;
  if (hasLegacyLocalLine) {
    return {
      visible: series?.lineHidden !== true,
      color: series?.lineColor ?? fallbackColor,
      widthEmu: series?.lineWidthEmu ?? null,
      dash: null,
      cap: null,
      join: null,
    };
  }
  if (linkedStyle?.lineNoStyle && options.linkedNoStyleFallback) {
    return resolved(null, fallbackColor);
  }
  return resolved(linkedStyle, fallbackColor);
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
  ctx.setLineDash(dashPatternForPreset(line.dash ?? undefined, ctx.lineWidth));
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
      lineDash: inheritPlotOutline ? line.dash : null,
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
const MAX_CANVAS_MARKER_GRADIENT_STOPS = MAX_CHART_PAINT_RECIPE_COMPONENTS;
const MAX_CANVAS_MARKER_PAINT_COMPONENTS = MAX_CHART_PAINT_COMPONENTS;
const MAX_CANVAS_LABEL_GRADIENT_STOPS = MAX_CANVAS_MARKER_GRADIENT_STOPS;
const MAX_CANVAS_LABEL_PAINT_COMPONENTS = MAX_CANVAS_MARKER_PAINT_COMPONENTS;

const CLASSIC_THREE_D_FAMILIES = new Set([
  'pie',
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
]);

function classicDataLabelPointIsPainted(
  chart: ChartModel,
  series: ChartSeries,
  family: string,
  index: number,
  scatterHasNumericX: boolean,
): boolean {
  if (family === 'surface') return false;
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
  if ((family === 'pie' || family === 'doughnut') && value <= 0) return false;
  if (family === 'scatter') {
    if (!scatterHasNumericX) return true;
    const category = (series.categories ?? chart.categories)[index];
    return category != null && Number.isFinite(Number.parseFloat(category));
  }
  return true;
}

function markerKeyPaintSizesPx(
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
    for (let index = 0; index < pointCount; index++) {
      if (!classicMarkerPointIsPainted(
        chart, series, family, index, scatterHasNumericX, markerContext,
      )) continue;
      if (isBubble
        && (bubbleScale <= 0
          || visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]) == null)) continue;
      const point = overrides.get(index);
      const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
      if (!markerSymbolConsumesFill(symbol)) continue;
      const bubblePaint = isBubble
        ? bubblePointFill(chart, series, point, index, chartColor(seriesIndex, series))
        : null;
      const paint = isBubble
        ? bubblePaint!.paint
        : markerFillPaintFor(series, point, index);
      let sizePx = Math.max(2, (point?.markerSize ?? series.markerSize ?? 5) * ptToPx);
      if (family === 'scatter' && isBubble) {
        const size = visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]);
        sizePx = size == null ? 0 : bubbleSizeMagnitude(bubbleSettings!, size) * bubbleScale;
      } else if (family === 'radar' && point?.markerSize == null && series.markerSize == null
        && chartRect) {
        sizePx = Math.max(4 * ptToPx, Math.min(chartRect.w, chartRect.h) * 0.025);
      }
      if (!chargePaint(paint, 1, sizePx)) return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      if (isBubble) {
        const linePaint = bubblePointLine(chart, series, point, index).paint;
        if (!chargePaint(linePaint, 1, sizePx)) {
          return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
        }
        if (bubblePointIsThreeD(series, point) && bubblePaint!.paint !== null
          && !chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS)) {
          return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
        }
      }
    }

    const seriesKeySymbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
    if (markerSymbolConsumesFill(seriesKeySymbol) && seriesLegendMarkerIsVisible(
      effectiveChartType, effectiveScatterStyle, series, effectiveRadarStyle,
    )) {
      const legendEntryDeleted = deletedLegendEntries.has(seriesIndex);
      const labelKeys = dataLabelLegendKeyCount(
        chart, series, family, pointCount, scatterHasNumericX, markerContext,
      );
      const keySizes = markerKeyPaintSizesPx(chart, series, ptToPx);
      const bubbleKeyFill = isBubble
        ? bubblePointFill(chart, series, undefined, seriesIndex, chartColor(seriesIndex, series))
        : null;
      const keyPaint = isBubble ? bubbleKeyFill!.paint : seriesMarkerFillPaint(series);
      const bubbleKeyLine = isBubble ? bubblePointLine(chart, series, undefined, seriesIndex) : null;
      const chargeKey = (repetitions: number, sizePx: number): boolean =>
        chargePaint(keyPaint, repetitions, sizePx)
        && (!isBubble || chargePaint(bubbleKeyLine!.paint, repetitions, sizePx))
        && (!isBubble || !bubblePointIsThreeD(series, undefined)
          || bubbleKeyFill!.paint === null
          || chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS, repetitions));
      if ((chart.showLegend && !legendEntryDeleted
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

function chartLabelBoxPaintComponents(box: ChartLabelBox | null | undefined): number | null {
  let total = 0;
  for (const paint of [box?.fillPaint, box?.borderFill]) {
    if (!paint) continue;
    const components = markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CANVAS_LABEL_GRADIENT_STOPS) {
      return null;
    }
    total += components;
  }
  return total;
}

function dataLabelHasContent(
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
): number | null {
  // ChartEx hierarchy labels are expanded only by the optional ChartEx
  // renderer, which owns their resource preflight with the hierarchy model.
  if (chart.chartexSunburst || chart.chartexTreemap) return null;
  let total = 0;
  const charge = (box: ChartLabelBox | null | undefined): boolean => {
    const components = chartLabelBoxPaintComponents(box);
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
  for (const series of chart.series) {
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
        chart, series, family, index, scatterHasNumericX,
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
function classicThreeDWorkCount(
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

function drawChartTextBoxes(
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
const CHART_SPACE_CORNER_RADIUS_PT = 10;

function chartSpaceRoundedPath(
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

/**
 * Render a chart (background frame + dispatch on `chartType`).
 * `rect` is in pixel coordinates on the target canvas.
 */
function renderChartImpl(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  /**
   * Pixels per point at the caller's current display scale. For PPTX at
   * 960px/12192000EMU the value is ~1.05; xlsx's sheet view renders at
   * device-px where 1pt≈1.333. Used to size title/axis labels whose
   * XML-specified sizes are in OOXML hundredths of a point.
   */
  ptToPx: number = PT_TO_PX,
  /**
   * Rotation already applied by the host frame transform. DrawingML gradient
   * fills with `rotWithShape="0"` counter-rotate by this amount.
   */
  shapeRotationDeg = 0,
  /** Optional 3-D renderer. Without it, the canonical 2-D family remains
   * visible and no mesh/camera implementation enters the static render path. */
  threeD?: ChartThreeDRenderer,
  /** Optional offline Region Map renderer. */
  regionMap?: ChartRegionMapRenderer,
  /** Host-warmed image cache used by picture-fill availability preflight. */
  imageLookup?: ChartImageLookup,
  /** Optional Microsoft ChartEx family renderer. */
  chartEx?: ChartExRenderer,
): void {
  // The per-family renderers (and the early-return/default text paths below)
  // mutate shared canvas state — textAlign, textBaseline, font, fillStyle,
  // etc. — without restoring it. Callers (docx/pptx draw chart shapes inline
  // with surrounding text; xlsx happens to wrap the call in its own
  // save/clip/restore) must not observe those mutations afterward. Wrapping
  // the whole body in a single save/restore here fixes it once for every
  // caller instead of requiring each call site to remember to do so.
  ctx.save();
  try {
    // Refuse oversized caller-supplied classic models before visibility/style
    // projections allocate replacement series, override, or trendline arrays.
    // Parsed packages are already bounded, but the public ChartModel contract
    // can also be constructed directly by an application.
    const sourceStructureCount = sourceChartStructureCount(chart);
    if (rejectOversizedCanvasChart(ctx, rect, sourceStructureCount)) return;
    chart = applyPlotVisibleOnly(chart);
    chart = applyLinkedChartStyleRoles(chart);
    const { x, y, w, h } = rect;
    const rounded = chart.roundedCorners === true;
    const cornerRadius = rounded ? CHART_SPACE_CORNER_RADIUS_PT * ptToPx : 0;
    if (rounded) {
      chartSpaceRoundedPath(ctx, x, y, w, h, cornerRadius);
      ctx.clip();
    }
    // Only fill the outer chartSpace when chartBg is set; a null means noFill
    // (transparent) per OOXML, so the underlying slide/sheet shows through.
    if (chart.chartFillHidden === true) {
      // Direct or linked `noFill`: retain the host surface beneath the chart.
    } else if (chart.chartFill?.fillType === 'image') {
      paintChartImageFill(
        ctx, chart.chartFill, x, y, w, h, ptToPx, shapeRotationDeg,
      );
    } else if (chart.chartFill) {
      const fill = resolveFill(chart.chartFill, ctx, x, y, w, h, shapeRotationDeg);
      if (fill) ctx.fillStyle = fill;
      if (fill) ctx.fillRect(x, y, w, h);
    } else if (chart.chartBg) {
      ctx.fillStyle = `#${chart.chartBg}`;
      ctx.fillRect(x, y, w, h);
    }

    // Explicit chart border — drawn only when DrawingML declares a paintable
    // line. Width comes from
    // `<a:ln@w>` (EMU → pt → px); absent width falls back to a 1px hairline.
    if (chart.chartBorderHidden !== true
      && (chart.chartBorderLineFill || chart.chartBorderColor)) {
      ctx.save();
      const stroke = chart.chartBorderLineFill
        ? resolveFill(chart.chartBorderLineFill, ctx, x, y, w, h, shapeRotationDeg)
        : chart.chartBorderColor ? `#${chart.chartBorderColor}` : null;
      if (!stroke) {
        ctx.restore();
      } else {
        ctx.strokeStyle = stroke;
      // `<a:ln>` with no `@w` means width 0 per ECMA-376 §20.1.2.2.24, i.e. invisible;
      // but Excel renders a fill-without-width line as a ~hairline, so we draw 1px to
      // match the app rather than dropping a declared border.
      const totalLineWidth = chart.chartBorderWidthEmu
        ? Math.max(0.5, chart.chartBorderWidthEmu / EMU_PER_PT) * ptToPx
        : 1;
      ctx.setLineDash(dashPatternForLine(
        chart.chartBorderCustomDash, chart.chartBorderDash, totalLineWidth,
      ));
      ctx.lineCap = chart.chartBorderCap === 'rnd'
        ? 'round' : chart.chartBorderCap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = chart.chartBorderJoin === 'round' || chart.chartBorderJoin === 'bevel'
        ? chart.chartBorderJoin : 'miter';
      // Inset by half the line width so the full stroke stays inside the rect.
      strokeChartFrameRect(
        ctx, x, y, w, h, totalLineWidth, chart.chartBorderCompound,
        rounded ? cornerRadius : 0,
      );
        ctx.restore();
      }
    }

    // chartEx box-and-whisker / sunburst / treemap carry their data in the structured
    // `chartexBox` / `chartexSunburst` / `chartexTreemap` fields, not the flat `series` array, so the
    // empty-series "(no data)" guard must not fire for them.
    const hasChartexData = chart.chartexBox != null || chart.chartexSunburst != null
      || chart.chartexTreemap != null || chart.chartexRegionMap != null;
    if (chart.series.length === 0 && !hasChartexData) {
      ctx.fillStyle = '#888';
      ctx.font = '12px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('(no data)', x + w / 2, y + h / 2);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    const classicPointCount = classicCanvasPointCount(chart);
    const classicMarkerPaintWork = (classicPointCount != null || chart.chartexBox != null)
      && (classicPointCount ?? 0) <= MAX_CANVAS_CHART_POINTS
      ? classicMarkerPaintWorkCount(chart, imageLookup, ptToPx, rect) : null;
    const classicThreeDWork = classicThreeDWorkCount(chart, threeD);
    const classicLabelPaintWork = chartLabelPaintWorkCount(chart, threeD);
    if (
      (classicPointCount != null || classicMarkerPaintWork != null
        || classicThreeDWork != null || classicLabelPaintWork != null)
      && rejectOversizedCanvasChart(
        ctx,
        rect,
        Math.max(
          classicPointCount ?? 0,
          classicMarkerPaintWork != null
            && classicMarkerPaintWork > MAX_CANVAS_MARKER_PAINT_COMPONENTS
            ? MAX_CANVAS_CHART_POINTS + 1 : 0,
          classicThreeDWork ?? 0,
          classicLabelPaintWork != null
            && classicLabelPaintWork > MAX_CANVAS_LABEL_PAINT_COMPONENTS
            ? MAX_CANVAS_CHART_POINTS + 1 : 0,
        ),
      )
    ) {
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    const plotDispatch = classicPlotDispatch(chart);
    if (plotDispatch === 'unsupported') {
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('Unsupported chart', x + w / 2, y + h / 2);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    if (plotDispatch !== 'legacy') {
      switch (plotDispatch) {
        case 'bar-combo':
          renderBarChart(ctx, chart, rect, ptToPx, {}, shapeRotationDeg);
          break;
        case 'line-groups':
          renderLineChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
          break;
        case 'area-groups':
          renderAreaChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
          break;
        case 'scatter-bubble':
          renderScatterChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
          break;
        case 'stock-line':
          renderStockChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
          break;
      }
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    // Classic 3-D groups keep their canonical 2-D family name in the shared
    // model, while `threeD` carries the authored view/depth contract.  Consume
    // that contract before ordinary dispatch; 2-D charts return false and keep
    // their existing byte-stable family paths.
    if (threeD?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }
    if (regionMap?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }
    if (chartEx?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    switch (chart.chartType) {
      case 'clusteredBar':
      case 'clusteredBarH':
      case 'stackedBar':
      case 'stackedBarH':
      case 'stackedBarPct':
      case 'stackedBarHPct':
        renderBarChart(ctx, chart, rect, ptToPx, {}, shapeRotationDeg); break;
      case 'line':
      case 'stackedLine':
      case 'stackedLinePct':
        renderLineChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'area':
      case 'stackedArea':
      case 'stackedAreaPct':
        renderAreaChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'pie':
        renderPieChart(ctx, chart, rect, false, ptToPx, shapeRotationDeg); break;
      case 'ofPie':
        renderOfPieChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'doughnut':
        renderPieChart(ctx, chart, rect, true, ptToPx, shapeRotationDeg); break;
      case 'radar':
        renderRadarChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'scatter':
      case 'bubble':
        renderScatterChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'stock':
        renderStockChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      case 'surface':
      case 'surface3D':
        renderSurfaceChart(ctx, chart, rect, ptToPx, shapeRotationDeg); break;
      default:
        ctx.fillStyle = '#888';
        ctx.font = '11px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        // The public model can carry a future layout identifier of arbitrary
        // length. Preserve that identifier in the model, but keep the
        // fail-closed paint path constant-work instead of shaping attacker-
        // controlled text that is not part of the rendered document.
        ctx.fillText('Unsupported chart', x + w / 2, y + h / 2);
    }
    drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
    drawChartTextBoxes(ctx, chart, rect, ptToPx);
  } finally {
    ctx.restore();
  }
}

export function renderChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number = PT_TO_PX,
  shapeRotationDeg = 0,
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  imageLookup?: ChartImageLookup,
  chartEx?: ChartExRenderer,
): void {
  withChartImageLookup(imageLookup, () => {
    renderChartImpl(
      ctx, chart, rect, ptToPx, shapeRotationDeg, threeD, regionMap, imageLookup, chartEx,
    );
  });
}
