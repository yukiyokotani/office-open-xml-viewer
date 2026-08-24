// Shared chart-frame layout (Phase 4 A1). Before this module, every chart
// family (bar/line/area/pie/radar/scatter/waterfall) recomputed the same
// outer-frame structure — title band → legend reserve → axis-title / label
// gutters → plot rectangle — inline, and the constants had DRIFTED between
// families (e.g. the title top pad is `h*0.02` for bar but `h*0.045` for line,
// the side legend reserve is `w*0.22` everywhere except pie's `w*0.28`).
//
// `computeChartFrame` centralises the frame math. It does NOT unify the drifted
// constants: each family passes its current numbers via `FrameParams`, so the
// pixels are byte-for-byte unchanged. The drift is now visible in ONE place
// (each call site's params object) as a prerequisite for a later, VRT-gated
// decision to converge them. See docs/dev-notes / the A1 report for the table.
//
// Two frame shapes are produced:
//   • cartesian (bar/line/area/scatter): a `plotRect` derived from a
//     `{t,r,b,l}` pad, itself built from the shared title/legend/axis-title
//     bands plus family-specific extras the caller resolves and passes in.
//   • radial (pie/radar): no pad; the plot is centred in the space left after
//     the title and legend bands, and the caller reads `plotRect` + `center`.
//
// The per-family MARK drawing (bars, lines, slices, points) stays in
// renderer.ts; only the frame is shared.

import type { ChartManualLayout, ChartModel } from '../types/chart';
import { EMU_PER_PT } from '../units.js';
import { categoryLabelOffsetPx } from './category-spacing';

// ─── Public types ────────────────────────────────────────────────────────────

/** Which side the legend occupies, or null when hidden. Mirrors the private
 *  `LegendSide` in renderer.ts (kept in sync; not exported from there). */
export type ChartLegendSide = 'r' | 'l' | 't' | 'b';

/** Legend band reserved out of the chart space. `reserveW` > 0 for left/right
 *  placement, `reserveH` > 0 for top/bottom. `null` = no legend. */
export interface ChartLegendReserve {
  side: ChartLegendSide;
  reserveW: number;
  reserveH: number;
}

/** Canvas-measured legend entry widths used by automatic packing. The renderer
 * resolves the authored/theme font and swatch geometry, while this module owns
 * the pure reserve calculation shared by every chart family. */
export interface ChartLegendMetrics {
  itemWidths: readonly number[];
  rowHeight: number;
  itemGap: number;
  horizontalPadding: number;
  verticalPadding: number;
}

/** Greedily pack source-order legend items into rows that fit `maxWidth`.
 * A single over-wide item owns its row and is width-capped by the painter. */
export function packLegendRows(
  itemWidths: readonly number[],
  maxWidth: number,
  itemGap: number,
): number[][] {
  const availableWidth = Math.max(1, maxWidth);
  const rows: number[][] = [];
  let row: number[] = [];
  let rowWidth = 0;
  for (let index = 0; index < itemWidths.length; index++) {
    const itemWidth = Math.min(availableWidth, Math.max(0, itemWidths[index]));
    const nextWidth = row.length === 0
      ? itemWidth
      : rowWidth + itemGap + itemWidth;
    if (row.length > 0 && nextWidth > availableWidth) {
      rows.push(row);
      row = [index];
      rowWidth = itemWidth;
    } else {
      row.push(index);
      rowWidth = nextWidth;
    }
  }
  if (row.length > 0) rows.push(row);
  return rows;
}

/** Per-side reserved legend widths/heights, split out of a
 *  {@link ChartLegendReserve} for convenient consumption. Exactly one of the
 *  four is non-zero (or all zero when there is no legend). */
export interface ChartLegendBands {
  legRightW: number;
  legLeftW: number;
  legTopH: number;
  legBottomH: number;
}

/** Title band metrics. `bandH` is the total vertical space the title reserves
 *  (`fontPx + topPad + bottomPad`, or 0 when there is no title). `topPad` is
 *  also the y-offset at which the title text is drawn (`y + topPad`). */
export interface ChartTitleBand {
  fontPx: number;
  topPad: number;
  bottomPad: number;
  bandH: number;
}

/** Axis-title bands (cartesian only). `catBandH` is reserved at the bottom
 *  (under the tick labels), `valBandW` at the left; both 0 when the respective
 *  title is absent. `catFontPx`/`valFontPx` are the title font sizes used both
 *  to size the bands and to draw the titles. Identical shape to the private
 *  `AxisTitleLayout` in renderer.ts. */
export interface ChartAxisTitleBands {
  catFontPx: number;
  valFontPx: number;
  catBandH: number;
  valBandW: number;
}

/** The computed plot rectangle (inner data region), in canvas px. */
export interface ChartPlotRect {
  px0: number;
  py0: number;
  pw: number;
  ph: number;
}

/** Full resolved frame for a chart. `plotRect` is the inner data region; the
 *  bands describe the gutters reserved around it. `center` is set for radial
 *  charts (pie/radar) as the plot-rect centre. */
export interface ChartFrame {
  title: ChartTitleBand;
  legend: ChartLegendReserve | null;
  legendBands: ChartLegendBands;
  axisTitles: ChartAxisTitleBands;
  plotRect: ChartPlotRect;
  center: { cx: number; cy: number };
}

// ─── Title band ──────────────────────────────────────────────────────────────

/** Product fallback for a chart title whose authored rich text and linked
 *  Chart Style both omit a size. OOXML does not define an automatic size. */
const DEFAULT_CHART_TITLE_SIZE_PT = 14;

/** Validate DrawingML `ST_TextFontSize` at the public-model boundary and
 * convert hundredths of a point to CSS pixels. Parsers apply the same
 * 100..400000 schema bound, but renderers also accept hand-authored models. */
export function chartTextFontSizePx(
  sizeHpt: number | null | undefined,
  ptToPx: number,
): number | null {
  return typeof sizeHpt === 'number'
    && Number.isFinite(sizeHpt)
    && sizeHpt >= 100
    && sizeHpt <= 400_000
    ? (sizeHpt / 100) * ptToPx
    : null;
}

/** Chart title font size (px). Honor the parser-resolved size first (authored
 *  rich text, then linked Chart Style); otherwise use one deterministic 14pt
 *  fallback across classic and ChartEx chart families. */
export function chartTitleFontPx(chart: ChartModel, _h: number, ptToPx: number): number {
  return chartTextFontSizePx(chart.titleFontSizeHpt, ptToPx)
    ?? DEFAULT_CHART_TITLE_SIZE_PT * ptToPx;
}

/** Fraction of the title font size used as the band's TOP pad — the gap from
 *  the band top down to the title's draw origin (`textBaseline='top'` box top).
 *
 *  This is FONT-proportional, not chart-height-proportional, on purpose. The
 *  title's inset above its glyphs is a property of the type (PowerPoint centers
 *  the title text in a slot sized to the font), so tying it to `h` made the
 *  title ride higher in tall frames and lower in short ones — the same title at
 *  the same point size landed at a different fraction of its own height. With
 *  `textBaseline='top'` the glyph cap-top sits ~0.19×font below the draw origin
 *  (the box-top → cap-top gap intrinsic to the face), so a top pad of ~0.62×font
 *  places the cap-top at ~0.81×font from the band top, matching PowerPoint's
 *  rendered chart titles in the bounded Office vector corpus.
 *
 *  For an already-resolved `fontPx`, the band's TOTAL height (`bandH`) is
 *  unchanged by this top/bottom-pad redistribution — see
 *  {@link chartTitleBand}. A different resolved title size may still change the
 *  band and therefore the plot rectangle below it. */
export const TITLE_TOP_PAD_FONT_FRAC = 0.62;

/** Resolve the title band from the family's top/bottom pad FRACTIONS (of `h`).
 *  These fractions still set the band's TOTAL height (`bandH = fontPx +
 *  h*topPadFrac + h*bottomPadFrac`), which every family's plot layout depends on.
 *  Given the same resolved `fontPx`, changing only the top/bottom-pad split keeps
 *  that total byte-identical; resolving a different font size changes `bandH`.
 *
 *  What changed: the title's vertical placement WITHIN the band. `topPad` (the
 *  draw offset) is now a FONT-proportional inset ({@link TITLE_TOP_PAD_FONT_FRAC}
 *  × fontPx) rather than the old `h * topPadFrac`, fixing the title riding at a
 *  different fraction of its height in tall vs short frames. `bottomPad` becomes
 *  the remainder so `bandH` is preserved exactly. The font inset is clamped to
 *  `[0, bandH - fontPx]` so a shallow band never pushes the title past the plot.
 *
 *  When the chart has no title the band collapses to zero (matching the
 *  `chart.title ? … : 0` guards inline). */
export function chartTitleBand(
  chart: ChartModel,
  h: number,
  ptToPx: number,
  topPadFrac: number,
  bottomPadFrac: number,
): ChartTitleBand {
  if (!chart.title && !chart.titlePresent) return { fontPx: 0, topPad: 0, bottomPad: 0, bandH: 0 };
  const fontPx = chartTitleFontPx(chart, h, ptToPx);
  // For a given resolved font size, the family fractions preserve the total
  // band height while only the top/bottom-pad split changes.
  const bandH = fontPx + h * topPadFrac + h * bottomPadFrac;
  // Font-proportional top inset, clamped so the title never overflows the band.
  const topPad = Math.min(Math.max(0, bandH - fontPx), fontPx * TITLE_TOP_PAD_FONT_FRAC);
  const bottomPad = bandH - fontPx - topPad;
  return { fontPx, topPad, bottomPad, bandH };
}

// ─── Legend reserve ──────────────────────────────────────────────────────────

/** Resolve legend placement from `<c:legendPos>`. Returns null when hidden.
 *  Verbatim from renderer.ts `legendLayout`, except the side reserve FRACTION
 *  is a parameter (`sideReserveFrac`) so pie can request its wider 0.28 band
 *  while every other family keeps 0.22. When measured metrics are supplied,
 *  top/bottom reserve complete packed rows (bounded to 30% of chart height)
 *  and sides reserve the measured entry width within the same 30% bound. */
export function chartLegendReserve(
  chart: ChartModel,
  w: number,
  h: number,
  sideReserveFrac: number,
  metrics?: ChartLegendMetrics,
): ChartLegendReserve | null {
  if (!chart.showLegend) return null;
  const pos = chart.legendPos ?? 'r';
  const side: ChartLegendSide = pos === 'l' ? 'l' : pos === 't' ? 't' : pos === 'b' ? 'b' : 'r';
  if (side === 'r' || side === 'l') {
    if (metrics) {
      const minWidth = Math.min(80, w * 0.3);
      // Once Canvas metrics are available, size the band from the content
      // rather than reusing the legacy family fraction as an upper bound. The
      // 30% safety cap still leaves the majority of the frame to the plot, but
      // permits a long authored series name to wrap into complete words.
      const maxWidth = w * 0.3;
      const measuredWidth = Math.max(0, ...metrics.itemWidths) + metrics.horizontalPadding;
      return {
        side,
        reserveW: Math.min(maxWidth, Math.max(minWidth, measuredWidth)),
        reserveH: 0,
      };
    }
    return { side, reserveW: Math.max(80, w * sideReserveFrac), reserveH: 0 };
  }
  if (metrics) {
    const availableWidth = Math.max(1, w - metrics.horizontalPadding);
    const rowCount = packLegendRows(
      metrics.itemWidths,
      availableWidth,
      metrics.itemGap,
    ).length;
    const desiredHeight = rowCount * metrics.rowHeight + metrics.verticalPadding;
    return {
      side,
      reserveW: 0,
      reserveH: Math.min(h * 0.3, desiredHeight),
    };
  }
  return { side, reserveW: 0, reserveH: Math.max(18, h * 0.08) };
}

/** Split a legend reserve into the four per-side bands (three of which are 0).
 *  Matches the `leg?.side === 'r' ? leg.reserveW : 0` idiom repeated inline. */
export function chartLegendBands(
  leg: ChartLegendReserve | null,
  overlay = false,
): ChartLegendBands {
  // §21.2.2.132: an overlay legend keeps its automatic paint rectangle but
  // contributes no reserved band to the plot layout.
  if (overlay) return { legRightW: 0, legLeftW: 0, legTopH: 0, legBottomH: 0 };
  return {
    legRightW: leg?.side === 'r' ? leg.reserveW : 0,
    legLeftW: leg?.side === 'l' ? leg.reserveW : 0,
    legTopH: leg?.side === 't' ? leg.reserveH : 0,
    legBottomH: leg?.side === 'b' ? leg.reserveH : 0,
  };
}

// ─── Axis-title bands ────────────────────────────────────────────────────────

/** Product fallback for an axis title whose run/style omits `a:rPr@sz`.
 *  OOXML does not define application auto-layout text metrics; the fixed 10pt
 *  value is the intentionally small compatibility policy recorded in #1228. */
export const AXIS_TITLE_FALLBACK_PT = 10;

/** Axis-title font size (px). Authored/style size is authoritative; omission
 *  is fixed at 10pt and therefore invariant across chart dimensions. */
export function axisTitleFontPx(
  sizeHpt: number | null | undefined,
  ptToPx: number,
): number {
  // ST_TextFontSize is 100..400000 hundredths of a point. Keep non-conforming
  // parser output and direct public-model inputs from creating negative or
  // unbounded layout bands; invalid values use the same product fallback as
  // omission. `ptToPx` is a host scale and is validated by the host renderer.
  return chartTextFontSizePx(sizeHpt, ptToPx) ?? AXIS_TITLE_FALLBACK_PT * ptToPx;
}

/** Convert local top/bottom DrawingML text-body insets into the screen extent
 * perpendicular to an axis title's baseline. Rotation changes which screen
 * dimension reserves the already-resolved local inset sum. */
export function axisTitleVerticalInsetPx(
  insetEmu: number | null | undefined,
  ptToPx: number,
): number {
  const inset = insetEmu != null && Number.isFinite(insetEmu) ? insetEmu : 0;
  return Math.max(0, inset) / EMU_PER_PT * ptToPx;
}

export type ChartAxisTitleSide = 'left' | 'right' | 'horizontal';

/** Resolve the title's paint rotation. DrawingML `ST_Angle` is expressed in
 *  60000ths of a degree; any explicit `bodyPr@rot`/`bodyPr@vert` value has
 *  priority. With no authoring, both vertical value-axis titles read
 *  bottom-to-top, matching Excel's primary and secondary axes,
 *  and a top/bottom value axis stays horizontal. */
export function axisTitleRotationRad(
  side: ChartAxisTitleSide,
  authoredRotation: number | null | undefined,
  authoredVerticalMode?: ChartModel['catAxisTitleVerticalMode'],
): number {
  let authoredDegrees = 0;
  let hasAuthoredOrientation = false;
  if (authoredVerticalMode != null) {
    hasAuthoredOrientation = true;
    // Canvas cannot reproduce East-Asian upright-glyph or WordArt stacking in
    // this single-line chart-title painter. Preserve those modes in the model
    // and approximate their vertical flow explicitly instead of silently
    // treating them as horizontal or applying the automatic side fallback.
    switch (authoredVerticalMode) {
      case 'horz':
        break;
      case 'vert270':
        authoredDegrees -= 90;
        break;
      case 'vert':
      case 'wordArtVert':
      case 'eaVert':
      case 'mongolianVert':
      case 'wordArtVertRtl':
        authoredDegrees += 90;
        break;
    }
  }
  if (authoredRotation != null && Number.isFinite(authoredRotation)) {
    authoredDegrees += authoredRotation / 60_000;
    hasAuthoredOrientation = true;
  }
  if (hasAuthoredOrientation) {
    return authoredDegrees * Math.PI / 180;
  }
  if (side === 'left') return -Math.PI / 2;
  if (side === 'right') return -Math.PI / 2;
  return 0;
}

/** Margin (px) between the chart's outer edge and an axis title. Verbatim from
 *  renderer.ts `axisTitleMargin`. */
export function axisTitleMargin(dim: number): number {
  return Math.max(8, dim * 0.02);
}

/** Axis-title bands (cat = bottom, val = left). Verbatim from renderer.ts
 *  `axisTitleLayout`: reserve `fontPx + margin + 4` on the side whose title is
 *  present, else 0. Identical across bar/line/area/scatter. */
export function chartAxisTitleBands(
  chart: ChartModel,
  w: number,
  h: number,
  ptToPx: number,
): ChartAxisTitleBands {
  const catFontPx = axisTitleFontPx(chart.catAxisTitleFontSizeHpt, ptToPx);
  const valFontPx = axisTitleFontPx(chart.valAxisTitleFontSizeHpt, ptToPx);
  const catTextInsetPx = axisTitleVerticalInsetPx(
    chart.catAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  const valTextInsetPx = axisTitleVerticalInsetPx(
    chart.valAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  return {
    catFontPx,
    valFontPx,
    catBandH: chart.catAxisTitle
      ? catFontPx + catTextInsetPx + axisTitleMargin(h) + 4
      : 0,
    valBandW: chart.valAxisTitle
      ? valFontPx + valTextInsetPx + axisTitleMargin(w) + 4
      : 0,
  };
}

// ─── PowerPoint auto-layout plot bands (cartesian) ───────────────────────────
//
// The reserves below match PowerPoint's chart AUTO-layout (`<c:plotArea><c:layout/>`
// with no `<c:manualLayout>`). ECMA-376 does not specify the auto-layout geometry
// — it only says the plot area is positioned automatically — so these constants
// model the RUNTIME behavior PowerPoint applies, pinned to the rendered ground
// truth. The load-bearing pin is the observed PLOT/frame ratio: a classic line
// chart places the plot rect at 0.611 of the frame height. The remaining
// 0.389 splits into the top reserve above the plot (title band + the gap down to
// the first gridline ≈ 0.236) and the bottom reserve (category-label band ≈
// 0.154). The title BAND itself is ≈ 0.200 of the frame — 0.236 is the top pad,
// not the band. User-approved to match the PDF.
//
// They are expressed as multiples of the relevant TEXT size (title font / axis
// label font), not of the chart height, because PowerPoint sizes each reserved
// band to the text it holds: a chart's title point size and axis-label point size
// are fixed by the XML regardless of the chart's pixel size, so a band tied to the
// frame height would ride at a different fraction of its own text in tall vs short
// frames (the same reasoning the title top-pad uses — see TITLE_TOP_PAD_FONT_FRAC).

/** Total vertical band a chart TITLE reserves, as a multiple of the title font
 *  size. PowerPoint centers the title text in a slot with air above and below;
 *  `2.25 × fontPx` reserves that slot. The reserve is pinned via the plot/frame
 *  ratio (0.611 in the measured classic-line case, see the block comment above);
 *  at that frame size the title BAND works out to ≈ 0.200 of the frame. (The
 *  0.236 figure sometimes quoted is the TOP PAD — band plus the gap down to the
 *  first gridline — not the band itself.) Replaces the old `fontPx + h·(top+bottom)`
 *  mix, whose h-proportional pad made the band collapse to a much smaller fraction
 *  of the frame on large charts (e.g. the xlsx demo charts) — a different fraction
 *  of the same-point title per frame size. */
export const TITLE_BAND_FONT_FRAC = 2.25;

/** Total vertical band a single row of horizontal CATEGORY tick labels reserves
 *  below the plot, as a multiple of the category-axis label font size. Models
 *  PowerPoint's reserve = axis-to-label gap (≈0.4×) + one label line-height
 *  (≈1.35×, ascent+descent+leading) + bottom outer margin (≈1.0×) = 2.75×fontPx.
 *  Pinned so the demo slide-5 line chart's category band lands at 0.154 of the
 *  frame (PowerPoint PDF). The old `fontPx + 12` (a fixed 12px gap) under-reserved
 *  this — 0.106 of the frame on slide-5, and it did not scale, shrinking to a few
 *  percent of the frame on the larger xlsx demo charts. */
export const CAT_AXIS_LABEL_BAND_FONT_FRAC = 2.75;

/** Font-proportional TITLE band for a cartesian chart (bar/line/area/scatter).
 *  Replaces the frac-based {@link chartTitleBand} for these families: the total
 *  band height is `titleFontPx × TITLE_BAND_FONT_FRAC` (independent of the chart
 *  height) so the title reserves the same fraction of its own text regardless of
 *  frame size. `topPad` (the draw offset) keeps the font-proportional inset from
 *  {@link TITLE_TOP_PAD_FONT_FRAC}, clamped inside the band. Collapses to zero
 *  when there is no title. The radial families (pie/radar) keep {@link
 *  chartTitleBand} via `computeChartFrame`, so this change does not touch them. */
export function cartesianTitleBand(
  chart: ChartModel,
  h: number,
  ptToPx: number,
): ChartTitleBand {
  if (!chart.title && !chart.titlePresent) return { fontPx: 0, topPad: 0, bottomPad: 0, bandH: 0 };
  const fontPx = chartTitleFontPx(chart, h, ptToPx);
  const bandH = fontPx * TITLE_BAND_FONT_FRAC;
  const topPad = Math.min(Math.max(0, bandH - fontPx), fontPx * TITLE_TOP_PAD_FONT_FRAC);
  const bottomPad = bandH - fontPx - topPad;
  return { fontPx, topPad, bottomPad, bandH };
}

/** Total bottom band (px) reserved for one row of horizontal category tick
 *  labels. `catAxFontPx` is the resolved category-axis label font size. Callers
 *  add the axis-title band and any bottom-legend reserve on top of this.
 *  See {@link CAT_AXIS_LABEL_BAND_FONT_FRAC}. */
export function catAxisLabelBandH(
  catAxFontPx: number,
  labelOffsetPercent?: number | null,
): number {
  const defaultGap = categoryTickLabelGapPx(catAxFontPx);
  const offsetGap = categoryLabelOffsetPx(defaultGap, labelOffsetPercent);
  return catAxFontPx * CAT_AXIS_LABEL_BAND_FONT_FRAC + offsetGap - defaultGap;
}

/** Office's default distance from an axis rule to one line of tick-label text,
 * expressed relative to the resolved label font. These are paint metrics, not
 * chart-size percentages, so the authored point size remains stable at zoom. */
export function categoryTickLabelGapPx(fontPx: number): number {
  return fontPx * (5 / 6);
}

export function valueTickLabelGapPx(fontPx: number): number {
  return fontPx;
}

/** Excel/PowerPoint keep 1.5 pt of clear chart space outside the tick-label
 * ink when resolving an authored outer plot rectangle. This is distinct from
 * the rule-to-label gap. The value is consistent across the horizontal and
 * vertical axes in the Office vector exports used to verify outer layouts. */
export const AXIS_OUTER_TEXT_MARGIN_PT = 1.5;

/** Measured conversion from a `layoutTarget="outer"` plot-area rectangle to
 * its inner data rectangle (ECMA-376 §21.2.2.89). The outer rectangle includes
 * tick marks, tick labels, axis titles, and Office's outer text clearance.
 * Keeping this one-line label geometry shared prevents the
 * bar, line, and area families from assigning different inner plot rectangles
 * to the same authored layout. */
export function chartManualOuterAxisInsets(metrics: Readonly<{
  valAxisHidden: boolean;
  catAxisHidden: boolean;
  valLabelWidth: number;
  valLabelFontPx: number;
  catLabelFontPx: number;
  valLabelGapPx?: number;
  catLabelGapPx?: number;
  outerTextMarginPx?: number;
  valTitleBandW: number;
  catTitleBandH: number;
  secondaryBandW?: number;
}>): ChartPad {
  const outerMargin = metrics.outerTextMarginPx ?? 0;
  return {
    t: metrics.valAxisHidden ? 0 : metrics.valLabelFontPx / 2 + outerMargin,
    r: (metrics.secondaryBandW ?? 0) > 0
      ? (metrics.secondaryBandW ?? 0) + outerMargin
      : 0,
    b: metrics.catAxisHidden
      ? 0
      : metrics.catLabelFontPx
        + (metrics.catLabelGapPx ?? categoryTickLabelGapPx(metrics.catLabelFontPx))
        + metrics.catTitleBandH
        + outerMargin,
    l: metrics.valAxisHidden
      ? 0
      : metrics.valLabelWidth
        + (metrics.valLabelGapPx ?? valueTickLabelGapPx(metrics.valLabelFontPx))
        + metrics.valTitleBandW
        + outerMargin,
  };
}

// ─── Frame parameters + computeChartFrame ────────────────────────────────────

/** A resolved `{t,r,b,l}` plot pad (canvas px). The caller builds this from the
 *  frame's shared bands plus its own extras (measured value-label gutter,
 *  secondary-axis bands, magic fractions), so `computeChartFrame` stays
 *  agnostic to per-family pad formulas while still owning the rect arithmetic. */
export interface ChartPad {
  t: number;
  r: number;
  b: number;
  l: number;
}

/** Parameters that drive {@link computeChartFrame}. Exactly one of `pad`
 *  (cartesian) or `radialGapFrac` (radial) selects the frame shape.
 *
 *  Title band: provide EITHER a pre-computed `titleBand` (the cartesian families
 *  fold {@link cartesianTitleBand} into their `pad` and pass the SAME band here,
 *  so `frame.title` matches the real reserved band) OR the frac pair
 *  `titleTopPadFrac` / `titleBottomPadFrac` (the radial families let
 *  `computeChartFrame` build the frac-based {@link chartTitleBand}). `titleBand`
 *  wins when both are set.
 *
 *  - `legendSideReserveFrac`: side (l/r) legend reserve fraction of `w`.
 *  - `pad`: fully-resolved cartesian plot pad. Its presence means "cartesian".
 *  - `plotAreaManualLayout`: honored (overrides `pad`) when present with w/h,
 *    matching the `<c:plotArea><c:manualLayout>` handling inline today.
 *  - `radialGapFrac`: for pie/radar, the extra `h * frac` gap subtracted below
 *    the title/legend before centring the plot. Presence means "radial". */
export interface FrameParams {
  /** Pre-computed title band (cartesian). When set, `frame.title` is exactly
   *  this — matching the band the caller already folded into `pad.t` — instead
   *  of a frac-derived band that would disagree with the plot rect. */
  titleBand?: ChartTitleBand;
  titleTopPadFrac?: number;
  titleBottomPadFrac?: number;
  legendSideReserveFrac: number;
  /** Pre-measured automatic legend reserve. Canvas callers pass this so the
   * pure frame uses the same measured bands the painter consumes. */
  legendReserve?: ChartLegendReserve | null;
  pad?: ChartPad;
  radialGapFrac?: number;
  honorPlotAreaManualLayout?: boolean;
  /** Insets from an authored `layoutTarget="outer"` rectangle to its inner
   * data region. The caller measures the real axis-label/tick bands and passes
   * them here; `layoutTarget="inner"` ignores these insets. ECMA-376
   * §21.2.2.89 defines the outer target as including tick marks and axis
   * labels, while CT_LayoutTarget defaults an omitted `val` to `outer`. */
  manualOuterInsets?: ChartPad;
}

export interface ManualLayoutRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/**
 * Resolve CT_ManualLayout against an element's automatic position.
 *
 * ECMA-376 §21.2.2.229/232/235 and CT_LayoutMode define factor dimensions as
 * chart-space fractions and edge dimensions as right/bottom coordinates.
 * MS-OI29500 §2.1.1587 further defines factor x/y as offsets from the default
 * element position, again in chart-space fractions. CT_LayoutMode@val defaults
 * to factor, including a present mode element with an omitted val attribute.
 */
export function resolveManualLayoutRect(
  manual: ChartManualLayout,
  chartRect: ManualLayoutRect,
  defaultRect: ManualLayoutRect,
): ManualLayoutRect | null {
  const xMode = manual.xMode || 'factor';
  const yMode = manual.yMode || 'factor';
  const wMode = manual.wMode || 'factor';
  const hMode = manual.hMode || 'factor';
  const x = xMode === 'edge'
    ? chartRect.x + manual.x * chartRect.w
    : defaultRect.x + manual.x * chartRect.w;
  const y = yMode === 'edge'
    ? chartRect.y + manual.y * chartRect.h
    : defaultRect.y + manual.y * chartRect.h;
  const rightOrWidth = manual.w == null
    ? defaultRect.w
    : wMode === 'edge'
      ? chartRect.x + manual.w * chartRect.w - x
      : manual.w * chartRect.w;
  const bottomOrHeight = manual.h == null
    ? defaultRect.h
    : hMode === 'edge'
      ? chartRect.y + manual.h * chartRect.h - y
      : manual.h * chartRect.h;
  if (![x, y, rightOrWidth, bottomOrHeight].every(Number.isFinite) ||
      rightOrWidth <= 0 || bottomOrHeight <= 0) return null;
  return { x, y, w: rightOrWidth, h: bottomOrHeight };
}

/**
 * Compute a chart's outer frame: title band, legend reserve, axis-title bands,
 * and the plot rectangle. This is the single home for the frame geometry that
 * every family previously duplicated.
 *
 * The shared bands (title/legend/axis-title) are always computed. The plot rect
 * is then resolved in one of two ways:
 *   • cartesian — `params.pad` supplies the resolved `{t,r,b,l}` insets;
 *     `px0=x+pad.l`, `pw=w-pad.l-pad.r`, etc. A `<c:plotArea><c:manualLayout>`
 *     overrides it when `honorPlotAreaManualLayout` is set.
 *   • radial — `params.radialGapFrac` is set; the plot fills the space left
 *     after the title/legend bands minus a `h*gap`, and `center` is its middle.
 *
 * NB: this function performs NO drawing and reads NO `ctx`; the family passes in
 * any ctx-measured value (e.g. the column value-label gutter) already folded
 * into `pad`. That keeps the frame math pure and unit-testable.
 */
export function computeChartFrame(
  chart: ChartModel,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx: number,
  params: FrameParams,
): ChartFrame {
  // Cartesian callers pass the SAME band they folded into `pad.t` so `frame.title`
  // agrees with the plot rect; radial callers pass the frac pair and let us build
  // the frac-based band. `titleBand` wins when both are set.
  const title =
    params.titleBand ??
    chartTitleBand(chart, h, ptToPx, params.titleTopPadFrac ?? 0, params.titleBottomPadFrac ?? 0);
  const legend = params.legendReserve !== undefined
    ? params.legendReserve
    : chartLegendReserve(chart, w, h, params.legendSideReserveFrac);
  const legendBands = chartLegendBands(legend, chart.legendOverlay === true);
  const axisTitles = chartAxisTitleBands(chart, w, h, ptToPx);

  let px0: number, py0: number, pw: number, ph: number;

  // First compute the automatic inner plot rectangle. Factor-mode x/y are
  // offsets from this default position, so manual layout cannot be resolved
  // correctly before the automatic frame exists.
  if (params.radialGapFrac != null) {
    // Radial (pie/radar): centre the plot in the leftover space. Verbatim from
    // the pie/radar inline math.
    const gap = h * params.radialGapFrac;
    pw = w - legendBands.legRightW - legendBands.legLeftW;
    ph = h - title.bandH - legendBands.legTopH - legendBands.legBottomH - gap;
    px0 = x + legendBands.legLeftW;
    py0 = y + title.bandH + legendBands.legTopH + gap;
  } else {
    const pad = params.pad;
    if (!pad) {
      throw new Error('computeChartFrame: cartesian frame requires params.pad');
    }
    px0 = x + pad.l;
    py0 = y + pad.t;
    pw = w - pad.l - pad.r;
    ph = h - pad.t - pad.b;
  }

  const pml = params.honorPlotAreaManualLayout ? chart.plotAreaManualLayout : null;
  if (pml) {
    const inset = pml.layoutTarget === 'inner'
      ? { t: 0, r: 0, b: 0, l: 0 }
      : (params.manualOuterInsets ?? { t: 0, r: 0, b: 0, l: 0 });
    const defaultTarget = pml.layoutTarget === 'inner'
      ? { x: px0, y: py0, w: pw, h: ph }
      : {
          x: px0 - inset.l,
          y: py0 - inset.t,
          w: pw + inset.l + inset.r,
          h: ph + inset.t + inset.b,
        };
    const resolved = resolveManualLayoutRect(
      pml,
      { x, y, w, h },
      defaultTarget,
    );
    if (resolved && resolved.w > inset.l + inset.r && resolved.h > inset.t + inset.b) {
      px0 = resolved.x + inset.l;
      py0 = resolved.y + inset.t;
      pw = resolved.w - inset.l - inset.r;
      ph = resolved.h - inset.t - inset.b;
    }
  }

  return {
    title,
    legend,
    legendBands,
    axisTitles,
    plotRect: { px0, py0, pw, ph },
    center: { cx: px0 + pw / 2, cy: py0 + ph / 2 },
  };
}
