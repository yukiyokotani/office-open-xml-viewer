// ===== Unified chart model =====
// Shared by @silurus/ooxml-pptx and @silurus/ooxml-xlsx.
//
// Parser JSON from each format is adapted into `ChartModel` and then passed
// to `renderChart` in @silurus/ooxml-core. This keeps a single source of
// truth for chart rendering across PPTX / XLSX (and future DrawingML charts
// in DOCX).

import type {
  DrawingMLCustomDashSegment,
  Fill,
  GradientFill,
  PatternFill,
  SolidFill,
} from './common';

export interface ChartSeries {
  name: string;
  /** Effective ChartEx `CT_Series@formatIdx` used to select linked Chart Style
   * formatting. When the attribute is omitted the parser stores the series'
   * original document-order index, before hidden series are removed. */
  chartexFormatIdx?: number | null;
  /** Hex without '#'. null = fall back to palette. */
  color: string | null;
  /**
   * `<c:ser><c:spPr><a:pattFill>` DrawingML pattern fill. When present it
   * paints the series marks and matching legend key; `color` remains the
   * solid/fallback series colour.
   */
  fillPattern?: PatternFill | null;
  /** `<c:invertIfNegative>` — negative bars use the alternate series paint. */
  invertIfNegative?: boolean | null;
  /** Application-generated negative marker style for an otherwise unformatted
   * classic bar/column series. Kept separate from authored
   * `<c:invertIfNegative>` so the wire model preserves element provenance. */
  automaticNegativeStyle?: boolean | null;
  /**
   * Alternate DrawingML fill from `c14:invertSolidFillFmt/c14:spPr`.
   * Kept as a shared fill recipe so solid, gradient, and pattern alternates
   * render identically in XLSX, DOCX, and PPTX chart hosts.
   */
  invertedFill?: SolidFill | GradientFill | PatternFill | null;
  /** Alternate negative fill explicitly authored as `<a:noFill>`. */
  invertedFillHidden?: boolean | null;
  /** Whether `c14:invertSolidFillFmt` directly authored a fill choice. */
  invertedFillAuthored?: boolean | null;
  /** Effective directly-authored outline for a negative inverted bar. */
  invertedLineColor?: string | null;
  invertedLineWidthEmu?: number | null;
  invertedLineHidden?: boolean | null;
  /** Whether the alternate `c14:spPr` directly contains `<a:ln>`. */
  invertedLineAuthored?: boolean | null;
  /** ChartEx `<cx:series><cx:spPr>` local shape paint. Positive fill/line
   * properties override linked style roles; series noFill remains distinct
   * from a data-point noFill. */
  chartexStyle?: ChartExElementStyle | null;
  /** `<c:ser><c:spPr><a:ln><a:solidFill>` series outline/stroke color (hex,
   * no '#'). For bar/column charts this is the border around each bar. */
  lineColor?: string | null;
  /** Explicit series outline/stroke width from `<a:ln@w>` (EMU). */
  lineWidthEmu?: number | null;
  /** Per-series 3-D bar shape (`CT_BarSer/c:shape`). Overrides the chart-group
   * `bar3DChart/c:shape`; omission inherits the group/schema default `box`. */
  threeDShape?: 'box' | 'cylinder' | 'cone' | 'coneToMax' | 'pyramid' | 'pyramidToMax' | string | null;
  /** Numeric values; null = missing data point. */
  values: (number | null)[];
  /** Host-resolved visibility provenance for the cells supplying each plotted
   * value (and X/bubble-size cells for scatter/bubble). `true` means at least
   * one required source cell is in a hidden row or column. The shared renderer
   * consumes this only when `ChartModel.plotVisibleOnly` is explicitly true. */
  sourceHidden?: boolean[] | null;
  /**
   * Per-data-point colors (pie / doughnut). Hex without '#'. null inside the
   * array = use palette for that slice. Omit entirely for non-pie series.
   */
  dataPointColors?: (string | null)[] | null;
  /**
   * `<c:pieChart|doughnutChart><c:ser><c:explosion val>` default pull-out
   * amount for every slice in this series. A point-level `dPt/explosion`
   * overrides it for that slice.
   */
  explosion?: number | null;
  /**
   * Per-data-point data-label text colors. Used by chartEx (`<cx:dataLabel idx>`)
   * to override label colour per bar — sample-2's waterfall paints negative
   * △ values in red while positive values stay black. Null inside the array =
   * fall back to the chart-level `dataLabelFontColor`.
   */
  dataLabelColors?: (string | null)[] | null;
  /**
   * Series-level data-label text colour (`<c:ser><c:dLbls><c:txPr>…solidFill`,
   * ECMA-376 §21.2.2.216). Hex without '#'. Stacked-bar charts colour each
   * segment's label independently (e.g. white on the dark segment, black on
   * the light one), which a single chart-level `dataLabelFontColor` can't
   * express. Takes precedence over `dataLabelFontColor`; null = no override.
   */
  labelColor?: string | null;
  /**
   * Mixed chart: per-series chart type override. Currently only "line" (XLSX
   * and PPTX combo charts) is honoured; other values are treated as the
   * chart's primary type.
   */
  seriesType?: string | null;
  /** Zero-based document-order index of the owning classic line-chart group.
   * Group decorations resolve through `ChartModel.lineGroupDecorations`. */
  lineGroupIndex?: number | null;
  /** Zero-based document-order index of the owning classic area-chart group.
   * Group drop lines resolve through `ChartModel.areaGroupDecorations`. */
  areaGroupIndex?: number | null;
  /** Zero-based document-order index of the owning classic `<c:barChart>` or
   * `<c:bar3DChart>` group. Separate groups sharing an axis overlay rather than
   * becoming additional members of one cluster. */
  barGroupIndex?: number | null;
  /** Direct `<c:barDir>` on the owning bar-chart group. */
  barGroupDirection?: 'bar' | 'col' | string | null;
  /** Direct `<c:grouping>` on the owning bar-chart group. */
  barGroupGrouping?: 'standard' | 'clustered' | 'stacked' | 'percentStacked' | string | null;
  /** Group-local `<c:gapWidth>`; distinct bar groups may author different
   * widths in the same combo chart. */
  barGroupGapWidth?: number | null;
  /** Group-local `<c:overlap>` for the owning bar group. */
  barGroupOverlap?: number | null;
  /**
   * Combo chart: this series is plotted against the SECONDARY value axis
   * (`ChartModel.secondaryValAxis`) — the `<c:valAx>` with `axPos="r"` /
   * `<c:crosses val="max">`. When false/absent the series uses the primary
   * (left) value-axis scale. PowerPoint's "Revenue vs. gross margin" combo
   * (sample-14 slide-8) puts the margin line on a 0–100% secondary axis.
   */
  useSecondaryAxis?: boolean | null;
  /**
   * Scatter-only X values (as strings). When null the series uses
   * `ChartModel.categories` as X.
   */
  categories?: string[] | null;
  /** Bubble-only provenance for a string-backed `<c:xVal>` source. Excel
   * exposes such a lone bubble series as one legend entry per point while a
   * numeric X source keeps the ordinary one-entry-per-series legend. */
  bubbleXSourceIsString?: boolean | null;
  /**
   * Resolved marker visibility for line/scatter series. ECMA-376 §21.2.2.32
   * `<c:marker><c:symbol>` defaults to "none" for line charts unless the
   * chart-level `<c:marker val="1"/>` or a per-series symbol opts in. When
   * undefined/null the renderer uses its own default (visible) so callers
   * that don't parse markers (e.g. pptx today) keep their existing behavior.
   */
  showMarker?: boolean | null;
  /**
   * Excel number-format code for this series' values (ECMA-376 §21.2.2.37,
   * `<c:val>/<c:numRef>/<c:formatCode>`). Used to format data labels when the
   * chart-level `<c:dLbls><c:numFmt>` is not set. null = no series-level code.
   */
  valFormatCode?: string | null;
  /** Number format from the series category/X cache. */
  catFormatCode?: string | null;
  /** Built-in worksheet number-format ID of the category source. Kept
   * separately because ID 14 is locale-sensitive, unlike a literal m/d/yy. */
  catFormatBuiltinId?: number | null;
  /** Per-point category/X formats from `<c:pt@formatCode>`. */
  catFormatCodes?: (string | null)[] | null;
  /**
   * `<c:marker><c:symbol val>` (ECMA-376 §21.2.2.32) — point marker shape.
   * One of "circle"|"square"|"diamond"|"triangle"|"x"|"plus"|"star"|
   * "dot"|"dash"|"picture"|"none". null = renderer default (circle when
   * showMarker is true).
   */
  markerSymbol?: string | null;
  /** Host-resolved automatic scatter marker when no `<c:marker>` symbol is authored. */
  automaticMarkerSymbol?: string | null;
  /**
   * `<c:marker><c:size val>` (ECMA-376 §21.2.2.34) — marker side length in
   * points. null = renderer default (~5 pt).
   */
  markerSize?: number | null;
  /** `<c:marker><c:spPr>` fill as 6/8-digit resolved hex (no `#`); transparent
   *  `00000000` represents an explicit `<a:noFill/>`. */
  markerFill?: string | null;
  /** Structured `<c:marker><c:spPr>` fill. Direct solid, gradient, and pattern
   * paints share the DrawingML fill model used by shapes and chart regions. */
  markerFillPaint?: Fill | null;
  /** A direct marker fill was authored even when its color could not be
   * resolved or its grammar is not represented by the current marker model.
   * This keeps a less-specific linked Chart Style from replacing it. */
  markerFillPaintAuthored?: boolean | null;
  /** `<c:marker><c:spPr><a:ln><a:solidFill>` resolved hex (no `#`). */
  markerLine?: string | null;
  /** `<c:marker><c:spPr><a:ln w>` marker-outline width in EMU. */
  markerLineWidthEmu?: number | null;
  /**
   * Per-data-point overrides (ECMA-376 §21.2.2.39 `<c:dPt>`). Keyed by point
   * index. Any unset field falls back to the series-level value.
   */
  dataPointOverrides?: ChartDataPointOverride[] | null;
  /**
   * Per-data-point custom labels (ECMA-376 §21.2.2.45 `<c:dLbl idx>`).
   * `text` is the resolved plain string — `<a:fld type="CELLRANGE">`
   * placeholders are already substituted at parse time. An empty string
   * means the point's label was deleted with `<c:delete val="1"/>` and
   * the renderer should skip it.
   */
  dataLabelOverrides?: ChartDataLabelOverride[] | null;
  /**
   * Series-level `<c:dLbls>` block (showVal / showSerName / position).
   * Applied to every point lacking its own `<c:dLbl>` override.
   */
  seriesDataLabels?: ChartSeriesDataLabels | null;
  /**
   * `<c:errBars>` per-series error bars (ECMA-376 §21.2.2.20). Up to two
   * (one per direction). Plus / minus deltas are absolute per-point values
   * regardless of `errValType`.
   */
  errBars?: ChartErrBars[] | null;
  /**
   * `<c:bubbleSize>` per-point sizes for bubble charts (ECMA-376 §21.2.2.4).
   * Drives marker radius — renderer treats the values as areas (radius
   * scales by sqrt) so visual area is proportional to value, matching
   * Excel. null / empty array = uniform marker size. Ignored for non-bubble
   * series.
   */
  bubbleSizes?: (number | null)[] | null;
  /** `<c:bubbleChart><c:bubble3D>` copied from this series' owning group.
   * Kept per series so multiple bubble groups cannot leak defaults. */
  bubble3DGroupDefault?: boolean | null;
  /** Direct `<c:bubbleChart><c:ser><c:bubble3D>` override. */
  bubble3D?: boolean | null;
  /**
   * `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series flag
   * requesting a smoothed (spline) curve through the points instead of straight
   * segments. Only consulted for the line and area families (scatter carries its
   * smoothing in `ChartModel.scatterStyle`). null/undefined/false = straight
   * polyline (the default; byte-stable for series that never set it).
   */
  smooth?: boolean | null;
  /**
   * `<c:ser><c:trendline>` per-series trendlines (ECMA-376 §21.2.2.211,
   * `CT_Trendline`). A series can carry several (e.g. a linear fit + a moving
   * average). null/undefined/empty = no trendline (the default; byte-stable for
   * series that never declare one).
   */
  trendLines?: ChartTrendline[] | null;
  /**
   * `<c:ser><c:spPr><a:ln><a:noFill/>` (ECMA-376 §21.2.2.198 CT_ShapeProperties
   * → DrawingML §20.1.2.2.24 CT_LineProperties). true when the series connecting
   * line is explicitly turned OFF. For a scatter/line series this OVERRIDES the
   * chart-group `<c:scatterStyle>` (§21.2.2.42) / line default — Excel and
   * PowerPoint draw markers only (no connecting line) even when the group style
   * is `lineMarker`. null/undefined = no explicit line-off, so the group default
   * governs (byte-stable for series that carry a paintable line).
   */
  lineHidden?: boolean | null;
}

/**
 * `<c:ser><c:trendline>` (ECMA-376 §21.2.2.211). A regression/smoothing curve
 * fitted to the series' data points.
 */
export interface ChartTrendline {
  /** Optional authored `<c:name>` shown in the legend. */
  name?: string | null;
  /**
   * `<c:trendlineType val>` (§21.2.2.213, `ST_TrendlineType` §21.2.3.50):
   * "linear" | "exp" | "log" | "power" | "poly" | "movingAvg". The renderer
   * draws all six forms through a bounded data-space fitter.
   */
  trendlineType: string;
  /** `<c:order val>` — polynomial order (`poly`, default 2). */
  order?: number | null;
  /** `<c:period val>` — moving-average window (`movingAvg`, default 2). */
  period?: number | null;
  /** `<c:forward val>` — units to extend the line past the last point. */
  forward?: number | null;
  /** `<c:backward val>` — units to extend the line before the first point. */
  backward?: number | null;
  /** `<c:intercept val>` — forced y-intercept (linear/exp). null = free fit. */
  intercept?: number | null;
  /** `<c:dispRSqr val="1">` — show the R² value. */
  dispRSqr?: boolean | null;
  /** `<c:dispEq val="1">` — show the fit equation. */
  dispEq?: boolean | null;
  /** `<c:trendlineLbl><c:layout><c:manualLayout>` authored label placement. */
  labelManualLayout?: ChartManualLayout | null;
  /** Explicit `<c:trendlineLbl><c:tx>` text, when present. */
  labelText?: string | null;
  /** Bounded formatted runs from `<c:trendlineLbl><c:tx><c:rich>`. */
  labelRichRuns?: ChartTextRun[] | null;
  /** `<c:trendlineLbl><c:numFmt formatCode>` for generated equation/R² values. */
  labelFormatCode?: string | null;
  /** `<c:trendlineLbl><c:numFmt sourceLinked>` authored linkage state. */
  labelFormatSourceLinked?: boolean | null;
  /** Trendline-label `<c:txPr>` run properties. */
  labelFontSizeHpt?: number | null;
  labelFontBold?: boolean | null;
  labelFontItalic?: boolean | null;
  labelFontColor?: string | null;
  /** A label-text fill choice was authored, even when it cannot be resolved. */
  labelFontPaintAuthored?: boolean | null;
  /** Direct `<a:noFill/>` on the trendline-label text. */
  labelFontHidden?: boolean | null;
  labelFontFace?: string | null;
  labelFontLanguage?: string | null;
  /** Normalized baseline shift (`0.3` = 30% of the effective font size). */
  labelFontBaseline?: number | null;
  labelTextRotation?: number | null;
  labelTextWrap?: string | null;
  labelTextVerticalAnchor?: string | null;
  labelTextVerticalMode?: string | null;
  labelTextLInsEmu?: number | null;
  labelTextTInsEmu?: number | null;
  labelTextRInsEmu?: number | null;
  labelTextBInsEmu?: number | null;
  /** A trendline-label `bodyPr` layer was authored. */
  labelTextBodyAuthored?: boolean | null;
  /** `<c:trendlineLbl><c:spPr>` fill/border. */
  labelBox?: ChartLabelBox | null;
  /** First authored paragraph alignment from `<c:trendlineLbl><c:txPr>`. */
  labelTextAlign?: string | null;
  /** `<c:spPr><a:ln><a:solidFill>` trendline color (hex without '#'). null =
   *  inherit the series color. */
  lineColor?: string | null;
  /** `<c:spPr><a:ln w>` trendline width in EMU. */
  lineWidthEmu?: number | null;
  /** `<c:spPr><a:ln><a:prstDash val>` DrawingML dash preset. */
  lineDash?: string | null;
  /** `<c:spPr><a:ln><a:noFill/>` — suppress the trendline stroke. */
  lineHidden?: boolean | null;
}

export interface ChartDataPointOverride {
  idx: number;
  /** Resolved fill hex (no `#`). */
  color?: string;
  /** Direct point `<a:noFill/>`; suppresses series/style fill fallback. */
  fillHidden?: boolean;
  /** Direct `<c:dPt><c:spPr>` DrawingML shape paint. Bubble charts consume
   * this carrier for gradient, pattern, picture, and unresolved fill
   * provenance; other classic families retain their established point model. */
  chartexStyle?: ChartExElementStyle | null;
  /** Direct point outline color (no `#`). */
  lineColor?: string;
  /** Direct point outline width in EMU. */
  lineWidthEmu?: number;
  /** Direct point DrawingML preset dash. */
  lineDash?: string;
  /** Direct point `<a:ln><a:noFill/>`; suppresses outline fallback. */
  lineHidden?: boolean;
  markerSymbol?: string;
  markerSize?: number;
  markerFill?: string;
  /** Direct point-marker structured fill. */
  markerFillPaint?: Fill | null;
  /** Direct point marker-fill provenance; retained independently from the
   * resolved paint so unsupported/unresolved paint still wins precedence. */
  markerFillPaintAuthored?: boolean | null;
  markerLine?: string;
  /** Direct point marker-outline width in EMU. */
  markerLineWidthEmu?: number;
  /** Direct `<c:dPt><c:bubble3D>` override. Only bubble charts consume it. */
  bubble3D?: boolean | null;
  /**
   * `<c:dPt><c:explosion val>` (ECMA-376 §21.2.2.61) — the amount this
   * pie/doughnut slice is moved out from the center. The schema type is
   * `CT_UnsignedInt` (unbounded `xsd:unsignedInt`); the spec text only says
   * "the amount the data point shall be moved from the center of the pie"
   * and does not itself define units or a 0–100 range. We treat it as a
   * de-facto percentage of the outer radius (0–100 typical), matching
   * Office's UI (the Point Explosion slider caps at 100%) rather than a
   * spec-mandated bound. undefined/absent = 0 (no explosion, flush with the
   * ring). Only consulted by the pie/doughnut renderer.
   */
  explosion?: number;
}

export interface ChartDataLabelOverride {
  idx: number;
  /** Empty string = label deleted (skip drawing). */
  text: string;
  /** Bounded DrawingML runs from a custom `<c:dLbl><c:tx><c:rich>` body.
   * Runs in one paragraph remain inline; a dedicated run containing `\n`
   * marks an authored paragraph break. The parser caps this payload at 4096
   * Unicode scalars and four lines.
   * Inline run paint is consumed by the shared bounded label path for classic
   * line/area/scatter/bubble, bar/column, and pie/doughnut (including callouts).
   * Family-specific layout still owns each label's anchor, capacity and clip.
   * undefined keeps the established plain-label path. */
  richRuns?: ChartTextRun[];
  /** "l"|"r"|"t"|"b"|"ctr"|"outEnd"|"bestFit". undefined = inherit. */
  position?: string;
  fontColor?: string;
  /** A point-label text fill was authored, including unsupported paint. */
  fontPaintAuthored?: boolean;
  /** Direct `<a:noFill/>` on point-label text. */
  fontHidden?: boolean;
  fontSizeHpt?: number;
  /** Effective per-point `<a:latin typeface>`; undefined inherits the series. */
  fontFace?: string;
  /** `<a:defRPr b="1">` inside the per-idx rich text. */
  fontBold?: boolean;
  fontItalic?: boolean;
  fontLanguage?: string;
  /** Normalized baseline shift (`0.3` = 30% of the effective font size). */
  fontBaseline?: number;
  textRotation?: number;
  textWrap?: string;
  textVerticalAnchor?: string;
  textVerticalMode?: string;
  textLInsEmu?: number;
  textTInsEmu?: number;
  textRInsEmu?: number;
  textBInsEmu?: number;
  /** A point-label `bodyPr` layer was authored. */
  textBodyAuthored?: boolean;
  /** First effective DrawingML paragraph alignment. */
  textAlign?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | string;
  /** Per-point number format; undefined inherits the series default. */
  formatCode?: string;
  /** Per-point component separator; undefined inherits the series default. */
  separator?: string;
  /** Authored `<c:dLbl><c:layout><c:manualLayout>` geometry. Automatic label
   *  placement is used when absent; explicit layout takes precedence when set. */
  manualLayout?: ChartManualLayout;
  /** Per-point callout box (`<c:dLbl><c:spPr>`, ECMA-376 §21.2.2.47/§21.2.2.197):
   *  overrides the series-default box for this one slice. */
  labelBox?: ChartLabelBox;
  /**
   * Per-point label-content flags (`<c:dLbl>` §21.2.2.47 carries the same
   * show-flag group as the series `<c:dLbls>` §21.2.2.49: §21.2.2.189
   * `<c:showVal>`, §21.2.2.177 `<c:showCatName>`, §21.2.2.180 `<c:showSerName>`,
   * §21.2.2.187 `<c:showPercent>`). When present they OVERRIDE the series-level
   * defaults for that one point (e.g. sample-14 slide-7's pie sets
   * `showCatName=0 showPercent=1` per slice while the series default is
   * `showCatName=1`, so each label is percent only). undefined = inherit the
   * series default for that flag.
   */
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  /** `<c:showBubbleSize>` for this point; undefined inherits the series. */
  showBubbleSize?: boolean;
  /** `<c:showLegendKey>` for this point; undefined inherits the series. */
  showLegendKey?: boolean;
  /**
   * `<c:dLbl><c:delete val="1"/>` (ECMA-376 §21.2.2.43) — the point's label is
   * removed. Distinguishes a genuine delete from a `<c:dLbl>` that only carries
   * style / flag overrides with no `<c:tx>` (both otherwise present as
   * `text === ''`). true = skip the label; undefined/absent = not deleted.
   */
  deleted?: boolean;
}

/** Callout-box style for a pie/doughnut data label — the white (or themed)
 *  rounded rectangle with a thin border Word draws around a `bestFit` label
 *  placed outside its slice. From the label's `<c:spPr>` (§21.2.2.197). All
 *  fields optional: absent → transparent / unbordered. Mirror of Rust
 *  `ChartLabelBox`. */
export interface ChartLabelBox {
  /** `<a:solidFill>` resolved hex (no `#`). Box background. */
  fill?: string;
  fillPaint?: SolidFill | GradientFill | PatternFill | null;
  fillHidden?: boolean | null;
  fillPaintAuthored?: boolean | null;
  /** `<a:ln><a:solidFill>` resolved hex (no `#`). Border stroke. */
  borderColor?: string;
  borderFill?: SolidFill | GradientFill | PatternFill | null;
  /** `<a:ln w>` border width in EMU (12700 EMU = 1 pt). */
  borderWidthEmu?: number;
  borderHidden?: boolean | null;
  borderPaintAuthored?: boolean | null;
  borderDash?: string | null;
  borderDashAuthored?: boolean | null;
  borderCustomDash?: ChartLineDashSegment[] | null;
  borderCap?: string | null;
  borderJoin?: string | null;
  borderCompound?: string | null;
}

export interface ChartSeriesDataLabels {
  /** Series-level `<c:dLbls><c:delete>` collection visibility. */
  deleted?: boolean | null;
  showVal: boolean;
  showCatName: boolean;
  showSerName: boolean;
  showPercent: boolean;
  /** Show the corresponding `<c:bubbleSize>` value in a bubble data label. */
  showBubbleSize?: boolean;
  /** Show the effective series/point legend key beside each data label. */
  showLegendKey?: boolean;
  position?: string;
  fontColor?: string;
  fontPaintAuthored?: boolean;
  fontHidden?: boolean;
  formatCode?: string;
  /** `<c:dLbls><c:separator>` (§21.2.2.170), including authored line breaks. */
  separator?: string;
  /** Series-level bold default for data labels. */
  fontBold?: boolean;
  fontItalic?: boolean;
  fontLanguage?: string;
  /** Normalized baseline shift (`0.3` = 30% of the effective font size). */
  fontBaseline?: number;
  /** Series-level font size for data labels (OOXML hundredths of a point). */
  fontSizeHpt?: number;
  /** Series-level `<c:dLbls><c:txPr>…<a:latin typeface>` font face. */
  fontFace?: string;
  textRotation?: number;
  textWrap?: string;
  textVerticalAnchor?: string;
  textVerticalMode?: string;
  textLInsEmu?: number;
  textTInsEmu?: number;
  textRInsEmu?: number;
  textBInsEmu?: number;
  /** A series-default `bodyPr` layer was authored. */
  textBodyAuthored?: boolean;
  /** Series-default DrawingML paragraph alignment. */
  textAlign?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | string;
  /** Series-default callout box (`<c:dLbls><c:spPr>`, ECMA-376 §21.2.2.49/
   *  §21.2.2.197). When present the pie/doughnut renderer draws Word's boxed
   *  callout layout (box + optional leader line) instead of plain text. */
  labelBox?: ChartLabelBox;
  /** `<c:dLbls><c:showLeaderLines val>` (§21.2.2.183) — draw leader lines from
   *  a pulled-away label back to its slice. Default false. */
  showLeaderLines?: boolean;
  /** `<c:leaderLines><c:spPr><a:ln><a:solidFill>` (§21.2.2.92) resolved hex
   *  (no `#`). undefined → renderer uses a neutral grey. */
  leaderLineColor?: string;
  /** `<c:leaderLines><c:spPr><a:ln w>` leader-line width in EMU. */
  leaderLineWidthEmu?: number;
  /** Explicit `<a:noFill/>` on the leader-line stroke. */
  leaderLineHidden?: boolean;
  /** DrawingML preset dash for the leader-line stroke. */
  leaderLineDash?: string;
}

export interface ChartErrBars {
  /** "x" | "y". */
  dir: string;
  /** "plus" | "minus" | "both". */
  barType: string;
  plus: (number | null)[];
  minus: (number | null)[];
  noEndCap: boolean;
  /** Resolved hex (no `#`). */
  color?: string;
  lineWidthEmu?: number;
  /** "solid"|"dash"|"dot"|"dashDot"|... */
  dash?: string;
  /** Explicit `<c:errBars><c:spPr><a:ln><a:noFill/>`. */
  hidden?: boolean;
}

/**
 * Canonical chart type vocabulary. Embeds direction (`H` = horizontal) and
 * grouping (`Pct` = percent-stacked) so renderers do not need to inspect
 * separate `barDir`/`grouping` fields.
 */
export type ChartType =
  | 'line' | 'stackedLine' | 'stackedLinePct'
  | 'clusteredBar' | 'clusteredBarH'
  | 'stackedBar' | 'stackedBarH'
  | 'stackedBarPct' | 'stackedBarHPct'
  | 'area' | 'stackedArea' | 'stackedAreaPct'
  | 'pie' | 'doughnut'
  | 'scatter' | 'bubble' | 'radar' | 'waterfall'
  | 'stock'
  | 'surface' | 'surface3D'
  // chartEx (MS 2014 chartex ext) layouts CH15 renders.
  | 'boxWhisker' | 'sunburst' | 'treemap'
  | string;

/** Exact classic chart-group element retained from `<c:plotArea>` source
 * order. These names deliberately do not fold 3-D or bubble groups into a
 * canonical 2-D family. */
export type ChartPlotGroupKind =
  | 'area' | 'area3D' | 'line' | 'line3D' | 'stock' | 'radar'
  | 'scatter' | 'pie' | 'pie3D' | 'doughnut' | 'bar' | 'bar3D'
  | 'ofPie' | 'surface' | 'surface3D' | 'bubble';

/** Resolved ownership of one axis role for a classic plot group. */
export type ChartPlotGroupAxisSlot = 'primary' | 'secondary' | 'none' | 'unresolved';

/**
 * Bounded source-order metadata for one direct classic chart-group child of
 * `<c:plotArea>`. Series are stored once in `ChartModel.series`; this record
 * owns a contiguous slice and therefore avoids a second scene graph or
 * group-by-series copying.
 */
export interface ChartPlotGroup {
  kind: ChartPlotGroupKind;
  seriesStart: number;
  seriesCount: number;
  categoryAxis: ChartPlotGroupAxisSlot;
  valueAxis: ChartPlotGroupAxisSlot;
  seriesAxis: ChartPlotGroupAxisSlot;
  /** Authored `axId` values in group-child order, retained as provenance. */
  axisIds?: string[] | null;
  grouping?: string | null;
  barDirection?: string | null;
  scatterStyle?: string | null;
  radarStyle?: string | null;
  gapWidth?: number | null;
  overlap?: number | null;
  bubbleScale?: number | null;
  bubbleSizeRepresents?: 'area' | 'w' | null;
  showNegativeBubbles?: boolean | null;
}

/** Backward-compatible chart name for the shared DrawingML dash atom. */
export type ChartLineDashSegment = DrawingMLCustomDashSegment;

/** Effective paint for one role in an Office 2013+ Chart Style part. */
export interface ChartExElementStyle {
  /** Linked Chart Style text defaults (`fontRef` + `defRPr`). */
  fontSizeHpt?: number | null;
  fontBold?: boolean | null;
  fontItalic?: boolean | null;
  fontColor?: string | null;
  fontPaintAuthored?: boolean | null;
  fontHidden?: boolean | null;
  fontFace?: string | null;
  /** Authored BCP-47 language from linked `defRPr`; never inferred from text. */
  fontLanguage?: string | null;
  /** Linked `defRPr@baseline`, normalized to a fraction (`0.3` = 30%). */
  fontBaseline?: number | null;
  /** Linked `bodyPr` text-body defaults. Direct chart text properties win. */
  textRotation?: number | null;
  textWrap?: string | null;
  textVerticalAnchor?: string | null;
  textVerticalMode?: string | null;
  textLInsEmu?: number | null;
  textTInsEmu?: number | null;
  textRInsEmu?: number | null;
  textBInsEmu?: number | null;
  textBodyAuthored?: boolean | null;
  /**
   * Per-color-style-index DrawingML fill recipes after `phClr` substitution.
   * Uses the same shared fill model as DrawingML shapes and cell-adjacent
   * drawing content, including relationship-backed picture fills.
   */
  fillPaints?: Array<Fill | null> | null;
  /** Per-color-style-index fills after `phClr` substitution and transforms. */
  fillColors?: Array<string | null> | null;
  fillHidden?: boolean | null;
  /** Linked fill was authored even when its paint is currently unsupported. */
  fillPaintAuthored?: boolean | null;
  /** Linked Chart Style uses `NoStyle`, not an authored no-fill paint. */
  fillNoStyle?: boolean | null;
  /** Per-color-style-index outlines after `phClr` substitution/transforms. */
  lineColors?: Array<string | null> | null;
  /** Structured outline paints after `phClr` substitution/transforms. */
  linePaints?: Array<SolidFill | GradientFill | PatternFill | null> | null;
  /** A linked outline was authored even when its paint is unsupported. */
  linePaintAuthored?: boolean | null;
  lineWidthEmu?: number | null;
  lineHidden?: boolean | null;
  /** Linked Chart Style uses `NoStyle`, not an authored no-fill outline. */
  lineNoStyle?: boolean | null;
  lineDash?: string | null;
  /** A preset/custom dash choice was authored, even if its value is absent. */
  lineDashAuthored?: boolean | null;
  /** DrawingML `<a:custDash>` atoms; presence overrides `lineDash`. */
  lineCustomDash?: ChartLineDashSegment[] | null;
  lineCap?: string | null;
  lineJoin?: string | null;
  /** Parsed compound-line token. Chart-frame painters use the bounded rail
   * ratios observed in Office vector output for `dbl`, `thinThick`,
   * `thickThin`, and `tri`. */
  lineCompound?: string | null;
  /** Fixed zero-based Chart Colors index; absent means relative (`auto`). */
  fillColorIndex?: number | null;
  /** Fixed zero-based Chart Colors index; absent means relative (`auto`). */
  lineColorIndex?: number | null;
}

/**
 * Paint-bearing `CT_ChartStyle` roles from MS-ODRAWXML §2.8.3.1. Marker
 * layout and `extLst` have their own non-style-entry grammar and therefore do
 * not appear in this map.
 */
export type ChartStyleRole =
  | 'axisTitle' | 'categoryAxis' | 'chartArea' | 'dataLabel'
  | 'dataLabelCallout' | 'dataPoint' | 'dataPoint3D' | 'dataPointLine'
  | 'dataPointMarker' | 'dataPointWireframe' | 'dataTable' | 'downBar'
  | 'dropLine' | 'errorBar' | 'floor' | 'gridlineMajor' | 'gridlineMinor'
  | 'hiLoLine' | 'leaderLine' | 'legend' | 'plotArea' | 'plotArea3D'
  | 'seriesAxis' | 'seriesLine' | 'title' | 'trendline'
  | 'trendlineLabel' | 'upBar' | 'valueAxis' | 'wall';

/** Authored low-to-high formatting for one classic surface-chart band. */
export interface ChartSurfaceBandFormat {
  idx: number;
  /** Direct outline geometry/paint and fill-authorship provenance. The one
   * authoritative direct fill recipe is carried by `fill`. */
  style?: ChartExElementStyle | null;
  fill?: SolidFill | GradientFill | PatternFill | null;
  fillHidden?: boolean | null;
  lineColor?: string | null;
  lineWidthEmu?: number | null;
  lineHidden?: boolean | null;
}

/** `<c:plotArea><c:dTable>` (`CT_DTable`) for classic DrawingML charts. */
export interface ChartDataTable {
  showHorizontalBorder: boolean;
  showVerticalBorder: boolean;
  showOutline: boolean;
  showKeys: boolean;
  fontSizeHpt?: number | null;
  fontFace?: string | null;
  fontColor?: string | null;
  fontBold?: boolean | null;
  fontItalic?: boolean | null;
  /** Resolved solid compatibility projection of `<c:dTable><c:spPr>`. */
  fillColor?: string | null;
  /** Direct DrawingML fill recipe. Preserved even where Office's application-
   * defined data-table paint extent has not been established for that recipe. */
  fill?: SolidFill | GradientFill | PatternFill | null;
  /** Explicit `<a:noFill>` on the data-table shape properties. */
  fillHidden?: boolean | null;
  /** True when `spPr` authored any DrawingML fill child, including one whose
   * paint recipe this implementation cannot resolve. */
  fillPaintAuthored?: boolean | null;
  lineColor?: string | null;
  lineWidthEmu?: number | null;
  lineDash?: string | null;
  lineHidden?: boolean | null;
}

export interface ChartModel {
  /** @internal Application font fallbacks prepared for this document. */
  providerFontRoutes?: import('../fonts/provider.js').FontFamilyRoutes;
  chartType: ChartType;
  title: string | null;
  /** Formatted DrawingML runs for a legacy chart title. `title` remains the
   * plain-text compatibility value used by callers that do not need styling. */
  titleRichRuns?: ChartTextRun[] | null;
  /** Direct chart title element exists; an empty title still reserves its band. */
  titlePresent?: boolean;
  categories: string[];
  /** Host-resolved visibility of the shared category source, aligned by point
   * index. Kept separate from the category strings so authored chart caches
   * remain authoritative while XLSX can still supply row/column visibility. */
  categorySourceHidden?: boolean[] | null;
  /**
   * `<c:multiLvlStrCache>` category levels, deepest/leaf level first. Sparse
   * outer levels retain empty entries so each non-empty label marks the start
   * of its category span.
   */
  categoryLevels?: string[][] | null;
  series: ChartSeries[];
  /** Ordered classic chart groups. Absent keeps legacy public models on the
   * existing single-family compatibility path. */
  plotGroups?: ChartPlotGroup[] | null;
  /** Text boxes in the Chart Drawing part referenced by `<c:userShapes>`.
   *  Coordinates are chart-space fractions from `<cdr:relSizeAnchor>`. */
  chartTextBoxes?: ChartTextBox[] | null;
  /**
   * §21.2.2.227 `<c:varyColors val="1"/>` on a SINGLE-series bar/column chart:
   * color each data point (bar) from the theme/palette sequence and list one
   * legend entry per point, matching Office. Pie/doughnut also preserve their
   * effective authored value so explicit false inherits the single series fill.
   */
  varyColors?: boolean | null;
  /** Show data labels on bars / points / slices. */
  showDataLabels: boolean;
  /** Explicit Y-axis minimum (OOXML `<c:valAx><c:min>`). */
  valMin: number | null;
  /** Explicit Y-axis maximum (OOXML `<c:valAx><c:max>`). */
  valMax: number | null;
  catAxisTitle: string | null;
  valAxisTitle: string | null;
  /** `<c:catAx><c:delete val="1"/>`. */
  catAxisHidden: boolean;
  /** `<c:valAx><c:delete val="1"/>`. */
  valAxisHidden: boolean;
  /** `<c:catAx><c:spPr><a:ln><a:noFill>` — Office-compatible suppression of
   *  the axis rule and tick marks. Labels and gridlines remain independent.
   *  Distinct from `catAxisHidden` (which removes everything via
   *  `<c:delete val="1"/>`). */
  catAxisLineHidden: boolean;
  /** `<c:valAx><c:spPr><a:ln><a:noFill>` — Office-compatible suppression of
   *  the axis rule and tick marks; labels and gridlines remain independent. */
  valAxisLineHidden: boolean;
  /** Hex without '#'. From `<c:plotArea><c:spPr><a:solidFill>`. */
  plotAreaBg: string | null;
  /** Structured `<c:plotArea><c:spPr>` fill. Solid fills are also mirrored in
   * `plotAreaBg` for wire compatibility. */
  plotAreaFill?: Fill | null;
  /** Explicit plot-area `noFill`; prevents linked style fallback. */
  plotAreaFillHidden?: boolean | null;
  /** A direct plot-area fill paint was authored, even when unresolved. */
  plotAreaFillPaintAuthored?: boolean | null;
  /** `plotAreaBg` is a host automatic fallback, not direct formatting. */
  plotAreaFillAutomatic?: boolean | null;
  /** Direct plot-area outline paint and width. */
  plotAreaLineColor?: string | null;
  plotAreaLineFill?: SolidFill | GradientFill | PatternFill | null;
  plotAreaLineWidthEmu?: number | null;
  plotAreaLineDash?: string | null;
  plotAreaLineDashAuthored?: boolean | null;
  plotAreaLineCustomDash?: ChartLineDashSegment[] | null;
  plotAreaLineCap?: string | null;
  plotAreaLineJoin?: string | null;
  plotAreaLineCompound?: string | null;
  /** Explicit plot-area outline `noFill`. */
  plotAreaLineHidden?: boolean | null;
  /** A direct plot-area line paint was authored, even when unresolved. */
  plotAreaLinePaintAuthored?: boolean | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** Structured non-solid `<c:chartSpace><c:spPr>` fill. Solid fills retain
   *  the legacy `chartBg` representation; gradient/pattern use the shared
   *  DrawingML fill model. */
  chartFill?: Fill | null;
  /** Explicit chart-area `noFill`; prevents host-default or linked fallback. */
  chartFillHidden?: boolean | null;
  /** A direct chart-area fill paint was authored, even when unresolved. */
  chartFillPaintAuthored?: boolean | null;
  /** `<c:chartSpace><c:roundedCorners>`; a bare element is true. Omission is
   *  preserved and renders the ordinary rectangular chart area. */
  roundedCorners?: boolean | null;
  /** `<c:chart><c:plotVisOnly>` (§21.2.2.146). A bare element is true; omission
   * is retained rather than inventing an application default. */
  plotVisibleOnly?: boolean | null;
  /** True when `<c:legend>` is declared in the chart XML. False = no legend. */
  showLegend: boolean;
  /** Optional category/series table authored below the cartesian plot. */
  dataTable?: ChartDataTable | null;
  /** `<c:legend><c:legendPos val>` — "r"|"l"|"t"|"b"|"tr". null = default (r). */
  legendPos: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** `<c:legend><c:overlay>` — when true, the legend is painted over the chart
   *  instead of reserving space in the automatic plot layout. */
  legendOverlay?: boolean | null;
  /** Indexed `<c:legendEntry>` overrides. Entries retain source order; `idx`
   *  addresses the effective series- or point-driven legend entry. */
  legendEntries?: ChartLegendEntryOverride[] | null;
  /** `<c:catAx><c:crossBetween val="..."/>`. "between" inserts 0.5-step padding
   *  on each end of the category axis; "midCat" anchors endpoints to the axes. */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. ECMA-376 default is "cross". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:valAx | catAx><c:minorTickMark>`. An omitted element is preserved as
   *  undefined/null: the renderer applies the host default (none in ordinary
   *  2-D charts; cross for the value axis in classic 3-D charts). A present
   *  element without `val` uses CT_TickMark's schema default `cross`. */
  valAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  catAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  /** Title font size in OOXML hundredths of a point (1600 = 16pt). null = default. */
  titleFontSizeHpt: number | null;
  /** Title font color as a hex string without '#' (e.g. "1B4332"). null = default. */
  titleFontColor: string | null;
  /** Title font family from `<a:latin typeface>` (ECMA-376 §20.1.4.2.24). null = default. */
  titleFontFace: string | null;
  /** `<c:catAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  catAxisFontSizeHpt: number | null;
  /** `<c:valAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  valAxisFontSizeHpt: number | null;
  /** `<c:catAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#').
   *  null = renderer default. Lets templates color category labels gray. */
  catAxisFontColor?: string | null;
  /** `<c:valAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  valAxisFontColor?: string | null;
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
  /** `<c:dLbls|cx:dataLabels>` text bold flag. null = chart-style default. */
  dataLabelFontBold?: boolean | null;
  /** Waterfall subtotal category indices. */
  subtotalIndices: number[];
  /** `<c:legend><c:manualLayout>` absolute placement fractions of the chart
   *  space (ECMA-376 §21.2.2.31). Overrides the default side-based legend
   *  rectangle while still letting `legendPos` decide which side of the plot
   *  gets the reserved band. null = use default layout. */
  legendManualLayout?: LegendManualLayout | null;
  /**
   * `<c:valAx><c:numFmt@formatCode>` — format code applied to value-axis tick
   * labels (ECMA-376 §21.2.2.21). null = plain numeric formatting.
   */
  valAxisFormatCode?: string | null;
  /** `<c:valAx><c:dispUnits>` display-only divisor and optional label. Series
   * values and plot geometry stay in their authored units. */
  valAxisDisplayUnits?: ChartDisplayUnits | null;
  /** Display units for a numeric horizontal axis (scatter/bubble). */
  catAxisDisplayUnits?: ChartDisplayUnits | null;
  /**
   * `<c:barChart><c:gapWidth>` — space between category groups as a
   * percentage of bar width (ECMA-376 §21.2.2.13). Default per spec is 150.
   * null = renderer default.
   */
  barGapWidth?: number | null;
  /**
   * `<c:barChart><c:overlap>` — signed percentage overlap between bars in the
   * same category cluster (ECMA-376 §21.2.2.25). Negative = gap, positive =
   * overlap, 0 = flush. Range [-100, 100]. null = renderer default (0).
   */
  barOverlap?: number | null;
  /**
   * `<c:dLbls><c:dLblPos>` — data label position (ECMA-376 §21.2.2.16).
   * "ctr"|"inBase"|"inEnd"|"outEnd"|"l"|"r"|"t"|"b"|"bestFit" etc.
   */
  dataLabelPosition?: string | null;
  /** Hex (no `#`) for data label text, resolved from `<c:dLbls><c:txPr>`. */
  dataLabelFontColor?: string | null;
  /**
   * `<c:dLbls><c:numFmt@formatCode>` — chart-level override for data label
   * number format (ECMA-376 §21.2.2.35). When absent, `valFormatCode` on each
   * series is used.
   */
  dataLabelFormatCode?: string | null;
  /** `<c:title>...defRPr@b>` chart title bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:catAx><c:txPr>...defRPr@b>` X-axis tick label bold flag. */
  catAxisFontBold?: boolean | null;
  /** `<c:catAx><c:txPr>...defRPr@i>` X-axis tick label italic flag. */
  catAxisFontItalic?: boolean | null;
  /** `<c:valAx><c:txPr>...defRPr@b>` Y-axis tick label bold flag. */
  valAxisFontBold?: boolean | null;
  /** `<c:valAx><c:txPr>...defRPr@i>` Y-axis tick label italic flag. */
  valAxisFontItalic?: boolean | null;
  /** `<c:catAx><c:title>` run-prop font size (hpt). Distinct from
   *  `catAxisFontSizeHpt` (tick labels). null = renderer default. */
  catAxisTitleFontSizeHpt?: number | null;
  /** `<c:catAx><c:title>` run-prop bold flag. null = not bold. */
  catAxisTitleFontBold?: boolean | null;
  /** `<c:catAx><c:title>` run-prop italic flag. */
  catAxisTitleFontItalic?: boolean | null;
  /** `<c:catAx><c:title>` run-prop color (hex without '#'). null = default. */
  catAxisTitleFontColor?: string | null;
  /** Authored `<c:catAx><c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle`
   *  units (60000ths of a degree). Applied independently from `vert`. */
  catAxisTitleRotation?: number | null;
  /** Authored DrawingML `bodyPr@vert`. Omission uses the side-based product
   *  fallback; explicit modes remain distinguishable from horizontal text. */
  catAxisTitleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:catAx><c:title><c:layout><c:manualLayout>`. */
  catAxisTitleManualLayout?: ChartManualLayout | null;
  /** Effective sum of the DrawingML top/bottom text insets for the
   * category-axis title, including CT_TextBodyProperties defaults. */
  catAxisTitleTextVerticalInsetEmu?: number | null;
  /** `<c:valAx><c:title>` run-prop font size (hpt). null = renderer default. */
  valAxisTitleFontSizeHpt?: number | null;
  /** `<c:valAx><c:title>` run-prop bold flag. null = not bold. */
  valAxisTitleFontBold?: boolean | null;
  /** `<c:valAx><c:title>` run-prop italic flag. */
  valAxisTitleFontItalic?: boolean | null;
  /** `<c:valAx><c:title>` run-prop color (hex without '#'). null = default. */
  valAxisTitleFontColor?: string | null;
  /** Authored `<c:valAx><c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle`
   *  units (60000ths of a degree). */
  valAxisTitleRotation?: number | null;
  /** Authored DrawingML `bodyPr@vert`. */
  valAxisTitleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:valAx><c:title><c:layout><c:manualLayout>`. */
  valAxisTitleManualLayout?: ChartManualLayout | null;
  /** Effective sum of the DrawingML top/bottom text insets for the
   * value-axis title. */
  valAxisTitleTextVerticalInsetEmu?: number | null;
  // ── Chart text font faces (CH10) ─────────────────────────────────────────
  // Each is the `<a:latin typeface>` (ECMA-376 §20.1.4.2.24) resolved from the
  // element's `<c:txPr>`. When absent the renderer falls back to the theme
  // body/heading font (`themeMinorFontLatin` / `themeMajorFontLatin`) and
  // finally to the built-in sans-serif, so a chart that specifies no faces is
  // byte-stable. Faces mirror the existing color/size/bold groups.
  /** `<c:catAx><c:txPr>…<a:latin typeface>` tick-label font. */
  catAxisFontFace?: string | null;
  /** `<c:valAx><c:txPr>…<a:latin typeface>` tick-label font. */
  valAxisFontFace?: string | null;
  /** `<c:catAx><c:title>…<a:latin typeface>` axis-title font. */
  catAxisTitleFontFace?: string | null;
  /** `<c:valAx><c:title>…<a:latin typeface>` axis-title font. */
  valAxisTitleFontFace?: string | null;
  /** `<c:dLbls><c:txPr>…<a:latin typeface>` data-label font. */
  dataLabelFontFace?: string | null;
  /** `<c:legend><c:txPr>…<a:latin typeface>` legend font. */
  legendFontFace?: string | null;
  /** `<c:legend><c:txPr>…<a:solidFill>` legend text color (hex without '#'). */
  legendFontColor?: string | null;
  /** `<c:legend><c:txPr>` legend font size (OOXML hundredths of a point). */
  legendFontSizeHpt?: number | null;
  /** `<c:legend><c:txPr>…defRPr@b` legend bold flag. */
  legendFontBold?: boolean | null;
  /** `<c:legend><c:spPr>` explicit frame fill (hex without '#'). */
  legendFillColor?: string | null;
  /** Structured `<c:legend><c:spPr>` fill. Solid fills are also mirrored in
   * `legendFillColor` for wire compatibility. */
  legendFill?: Fill | null;
  /** Explicit `<c:legend><c:spPr><a:noFill/>`; prevents linked style fallback. */
  legendFillHidden?: boolean | null;
  /** A direct legend fill paint was authored, even when its color could not be resolved. */
  legendFillPaintAuthored?: boolean | null;
  /** `<c:legend><c:spPr><a:ln>` explicit frame stroke (hex without '#'). */
  legendLineColor?: string | null;
  legendLineFill?: SolidFill | GradientFill | PatternFill | null;
  /** `<c:legend><c:spPr><a:ln@w>` frame stroke width in EMU. */
  legendLineWidthEmu?: number | null;
  legendLineDash?: string | null;
  legendLineDashAuthored?: boolean | null;
  legendLineCustomDash?: ChartLineDashSegment[] | null;
  legendLineCap?: string | null;
  legendLineJoin?: string | null;
  legendLineCompound?: string | null;
  /** Explicit `<c:legend><c:spPr><a:ln><a:noFill/>`; prevents linked fallback. */
  legendLineHidden?: boolean | null;
  /** A direct legend line paint was authored, even when its color could not be resolved. */
  legendLinePaintAuthored?: boolean | null;
  /**
   * Theme font-scheme faces (`<a:fontScheme>`, ECMA-376 §20.1.4.2). Latin
   * heading (majorFont) and body (minorFont) typefaces, used as the fallback
   * for any chart text element whose own `<c:txPr>` supplies no `<a:latin>`.
   * null when the theme is not threaded to the chart (then the renderer's
   * built-in sans-serif remains, byte-stable). Axis titles / chart title use
   * the major (heading) face; tick labels / data labels / legend use the
   * minor (body) face — matching Office's default chart text styling.
   */
  themeMajorFontLatin?: string | null;
  themeMinorFontLatin?: string | null;
  /** Explicit chart border color (hex without '#') from
   *  `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr>`. Only set when the
   *  XML explicitly declares a paintable line; null otherwise (no default
   *  border is drawn). */
  chartBorderColor?: string | null;
  chartBorderLineFill?: SolidFill | GradientFill | PatternFill | null;
  /** `<c:chartSpace><c:spPr><a:ln@w>` border width in EMU. null = 1px hairline
   *  when a color is present. */
  chartBorderWidthEmu?: number | null;
  chartBorderDash?: string | null;
  chartBorderDashAuthored?: boolean | null;
  chartBorderCustomDash?: ChartLineDashSegment[] | null;
  chartBorderCap?: string | null;
  chartBorderJoin?: string | null;
  chartBorderCompound?: string | null;
  /** Explicit chart-area border `noFill`. */
  chartBorderHidden?: boolean | null;
  /** A direct chart-area line paint was authored, even when unresolved. */
  chartBorderPaintAuthored?: boolean | null;
  /**
   * `<c:catAx><c:crosses val>` (`autoZero` | `min` | `max`). Drives the Y
   * coordinate where the X axis is drawn. Default `autoZero` puts the X
   * axis at y=0 — that's how Excel "Project Timeline" templates split
   * milestones (positive Y) above and tasks (negative Y) below the axis.
   */
  catAxisCrosses?: string | null;
  /** `<c:catAx><c:crossesAt val>` — explicit numeric override for the
   *  crossing point. Takes precedence over `catAxisCrosses`. */
  catAxisCrossesAt?: number | null;
  valAxisCrosses?: string | null;
  valAxisCrossesAt?: number | null;
  /** Axis line color (hex without `#`) and width in EMU from
   *  `<c:catAx|valAx><c:spPr><a:ln>`. */
  catAxisLineColor?: string | null;
  catAxisLineWidthEmu?: number | null;
  catAxisLineDash?: string | null;
  /** A direct `<c:catAx><c:spPr><a:ln>` paint was authored. */
  catAxisLinePaintAuthored?: boolean | null;
  valAxisLineColor?: string | null;
  valAxisLineWidthEmu?: number | null;
  valAxisLineDash?: string | null;
  /** A direct `<c:valAx><c:spPr><a:ln>` paint was authored. */
  valAxisLinePaintAuthored?: boolean | null;
  /**
   * `<c:catAx><c:numFmt@formatCode>` (or scatter X-axis valAx). When set,
   * the renderer formats X-axis tick labels with this code (e.g. dates).
   */
  catAxisFormatCode?: string | null;
  /**
   * `<c:catAx><c:scaling><c:min/max>` — explicit X-axis range. Used by
   * scatter / bubble charts whose X axis is numeric. null = derive from
   * data extents.
   */
  catAxisMin?: number | null;
  catAxisMax?: number | null;
  /**
   * `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27) absolute
   * placement for the chart title.
   */
  titleManualLayout?: ChartManualLayout | null;
  /**
   * `<c:plotArea><c:layout><c:manualLayout>` absolute placement for the
   * plot area. `layoutTarget="inner"` describes the inner plot rect (no axes /
   * labels); `outer` describes the outer rect (axes included) and is the schema
   * default when the element or its `val` is omitted.
   */
  plotAreaManualLayout?: ChartManualLayout | null;
  /**
   * `<c:scatterChart><c:scatterStyle val>` (ECMA-376 §21.2.2.42). Drives
   * whether scatter charts connect points with lines and whether those
   * lines are smoothed. Values: "marker" (markers only — Excel default
   * "Scatter"), "line" / "lineMarker" (straight segments), "smooth" /
   * "smoothMarker" (cubic Bézier through points), "lineNoMarker",
   * "smoothNoMarker". null = renderer default ("marker"). Only consulted
   * for `chartType === "scatter"`; bubble ignores it.
   */
  scatterStyle?: string | null;
  /**
   * `<c:bubbleChart><c:bubbleScale val>` (ECMA-376 §21.2.2.21), 0–300 percent
   * of the default bubble diameter. null/undefined uses the schema default 100.
   */
  bubbleScale?: number | null;
  /**
   * `<c:bubbleChart><c:sizeRepresents val>` (ECMA-376 §21.2.2.193,
   * ST_SizeRepresents §21.2.3.43). `area` (the schema default) makes bubble
   * area proportional to the value; `w` makes the radius proportional.
   */
  bubbleSizeRepresents?: 'area' | 'w' | null;
  /**
   * `<c:bubbleChart><c:showNegBubbles val>` (ECMA-376 §21.2.2.179).
   * Absent means false; a present CT_Boolean without `val` means true.
   */
  showNegativeBubbles?: boolean | null;
  /**
   * `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10). Controls
   * whether radar series render as line + markers ("standard" / "marker")
   * or as a closed polygon with area fill ("filled"). null = default
   * ("standard" — line, no fill). Only consulted for `chartType === "radar"`.
   */
  radarStyle?: string | null;
  /**
   * Secondary value axis for combo charts (bar + line). When present, series
   * with `useSecondaryAxis` are plotted against this axis's independent scale
   * and the axis is drawn on the right edge of the plot. null/absent = single
   * value axis (the common case). See {@link SecondaryValueAxis}.
   */
  secondaryValAxis?: SecondaryValueAxis | null;
  /** Secondary horizontal axis: either the numeric `<c:valAx>` used by an
   *  overlaid scatter/bubble group, or the top/right `<c:catAx>` paired with a
   *  secondary bar/column group. */
  secondaryCatAxis?: SecondaryValueAxis | null;
  /**
   * `<c:date1904>` (ECMA-376 §21.2.2.38). When true the chart's serial
   * date-times resolve against the 1904 date system (base 1904-01-01) instead
   * of the default 1900 system. Threaded to the date formatters for date-axis
   * category labels and value-axis tick labels. Omitted/false ⇒ 1900 system.
   * Note: per §21.2.2.38 the element's `val` defaults to true when present but
   * the attribute is omitted, so `<c:date1904/>` alone means date1904=true.
   */
  date1904?: boolean;
  /**
   * `<c:doughnutChart><c:holeSize val>` (ECMA-376 §21.2.2.60,
   * `ST_HoleSizePercent` §21.2.3.55) — the doughnut hole diameter as a
   * percentage 1–90 of the outer diameter. Ignored for pie (which has no
   * hole). null/undefined = use the renderer's doughnut default when the
   * element is absent. Note the ECMA `CT_HoleSize` schema default is 10%, but
   * a real doughnut file always writes an explicit `<c:holeSize>` (Excel /
   * PowerPoint emit 50–75%); the renderer falls back to 50% only for the
   * pathological absent case.
   */
  holeSize?: number | null;
  /**
   * `<c:pieChart | doughnutChart><c:firstSliceAng val>` (ECMA-376 §21.2.2.52,
   * `ST_FirstSliceAng` §21.2.3.15) — the angle in degrees (0–360, clockwise
   * from the 12 o'clock position) at which the first slice begins.
   * null/undefined = 0 (start at 12 o'clock), which matches the renderer's
   * historical fixed −90° (canvas up) start.
   */
  firstSliceAngle?: number | null;
  /**
   * `<c:chartSpace><c:chart><c:dispBlanksAs val>` (ECMA-376 §21.2.2.42,
   * `ST_DispBlanksAs` §21.2.3.10) — how blank (null) cells are plotted on
   * line/area charts:
   *   - "gap"  → leave a gap (break the line). The renderer's historical
   *              behavior and the model default when the element is absent.
   *   - "zero" → plot the blank as the value 0 (the point drops to the axis).
   *   - "span" → skip the blank but connect its neighbours with a straight
   *              line (bridge the gap).
   * Note the XSD `@val` default is "zero" (applies when `<c:dispBlanksAs/>` is
   * present but the attribute is omitted); when the ELEMENT is absent entirely
   * Office falls back to "gap", which is what we model as the default. Only
   * consulted for the line and area families. null/undefined = "gap".
   */
  dispBlanksAs?: string | null;
  /**
   * `<c:chart><c:showDLblsOverMax>` (ECMA-376 §21.2.2.180). When true,
   * labels whose plotted value exceeds the effective value-axis maximum remain
   * visible. A missing element or explicit false suppresses those labels.
   */
  showDataLabelsOverMax?: boolean | null;
  // ── Axis scale model (CH6) ───────────────────────────────────────────────
  // Gridline presence, manual major/minor units, log scale and orientation.
  // Every field is byte-stable when absent: the renderer keeps its historical
  // "value gridlines always on, category gridlines off, linear minMax axis"
  // behavior unless one of these is explicitly set.
  /**
   * `<c:valAx><c:majorGridlines>` presence (ECMA-376 §21.2.2.100). `false` when
   * the value axis exists but omits the element (Office suppresses value
   * gridlines). null/undefined ⇒ the renderer's historical always-on value
   * gridlines (byte-stable). `true` is redundant with the default but honored.
   */
  valAxisMajorGridlines?: boolean | null;
  /**
   * `<c:catAx><c:majorGridlines>` presence (§21.2.2.100). `true` turns on
   * category-axis gridlines (Office omits them by default). null/undefined/false
   * ⇒ no category gridlines (the historical default, byte-stable).
   */
  catAxisMajorGridlines?: boolean | null;
  /**
   * `<c:valAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
   * color (hex without `#`) — ECMA-376 §21.2.2.100. When set, the value-axis
   * major gridlines are stroked in this color instead of the renderer's faint
   * `#e0e0e0` default (e.g. sample-1 slide 5's `accent3` gridlines). null/absent
   * ⇒ the historical default (byte-stable).
   */
  valAxisGridlineColor?: string | null;
  /**
   * `<c:valAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU. When
   * set, the value-axis gridline stroke width is derived from this (floored so a
   * hairline stays visible). null/absent ⇒ the renderer's 0.5 px default.
   */
  valAxisGridlineWidthEmu?: number | null;
  /** `<c:valAx><c:majorGridlines>...<a:prstDash val>` dash preset. */
  valAxisGridlineDash?: string | null;
  /**
   * `<c:catAx><c:majorGridlines><c:spPr><a:ln><a:solidFill>` resolved gridline
   * color (hex without `#`). Only meaningful when {@link catAxisMajorGridlines}
   * is on. null/absent ⇒ the faint default.
   */
  catAxisGridlineColor?: string | null;
  /** `<c:catAx><c:majorGridlines><c:spPr><a:ln w>` gridline width in EMU. */
  catAxisGridlineWidthEmu?: number | null;
  /** `<c:catAx><c:majorGridlines>...<a:prstDash val>` dash preset. */
  catAxisGridlineDash?: string | null;
  /** `<c:valAx><c:minorGridlines>` presence (§21.2.2.109). Only drawn when a
   *  minor step is resolvable (see {@link valAxisMinorUnit}). */
  valAxisMinorGridlines?: boolean | null;
  /** Authored value-axis minor-gridline paint. */
  valAxisMinorGridlineColor?: string | null;
  valAxisMinorGridlineWidthEmu?: number | null;
  valAxisMinorGridlineDash?: string | null;
  /** `<c:catAx|valAx><c:minorGridlines>` on the horizontal/scatter axis. */
  catAxisMinorGridlines?: boolean | null;
  /** Authored horizontal-axis minor-gridline paint. */
  catAxisMinorGridlineColor?: string | null;
  catAxisMinorGridlineWidthEmu?: number | null;
  catAxisMinorGridlineDash?: string | null;
  /**
   * `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit distance between major
   * gridlines/ticks, overriding the Excel-style auto "nice" step. null/undefined
   * ⇒ auto step (byte-stable).
   */
  valAxisMajorUnit?: number | null;
  /** `<c:valAx><c:minorUnit val>` (§21.2.2.112) — explicit minor step. Drives
   *  minor gridlines/ticks when present. When omitted but either feature is
   *  requested, the renderer uses the automatic major unit divided by five. */
  valAxisMinorUnit?: number | null;
  /** Numeric horizontal-axis major step (scatter/bubble `<c:valAx>`). */
  catAxisMajorUnit?: number | null;
  /** Numeric horizontal-axis minor step (scatter/bubble `<c:valAx>`). */
  catAxisMinorUnit?: number | null;
  /**
   * Whether the authored category axis is a classic date axis (`<c:dateAx>`).
   * Date axes position numeric serial categories in authored base-time-unit
   * calendar slots rather than assigning every source point an ordinal slot.
   */
  catAxisIsDate?: boolean | null;
  /** `<c:dateAx><c:baseTimeUnit val>` (`days` by schema default). */
  catAxisBaseTimeUnit?: 'days' | 'months' | 'years' | string | null;
  /** `<c:dateAx><c:majorTimeUnit val>` (`days` by schema default). */
  catAxisMajorTimeUnit?: 'days' | 'months' | 'years' | string | null;
  /** `<c:dateAx><c:minorTimeUnit val>` (`days` by schema default). */
  catAxisMinorTimeUnit?: 'days' | 'months' | 'years' | string | null;
  /** `<c:catAx><c:noMultiLvlLbl>` suppresses outer category levels. */
  catAxisNoMultiLevelLabels?: boolean | null;
  /**
   * `<c:valAx><c:scaling><c:logBase val>` (§21.2.2.98, `ST_LogBase` §21.2.3.25)
   * — logarithmic value-axis base (>= 2). When set, values map to pixels in log
   * space and gridlines fall on powers of the base. null/undefined ⇒ linear
   * (byte-stable).
   */
  valAxisLogBase?: number | null;
  /** Numeric horizontal-axis logarithmic base (scatter/bubble second valAx). */
  catAxisLogBase?: number | null;
  /**
   * `<c:valAx><c:scaling><c:orientation val>` (§21.2.2.130, `ST_Orientation`
   * §21.2.3.30) — "minMax" (normal) | "maxMin" (reversed, so the value axis runs
   * top→bottom max→min). null/undefined/"minMax" ⇒ normal (byte-stable).
   */
  valAxisOrientation?: 'minMax' | 'maxMin' | string | null;
  /** `<c:catAx><c:scaling><c:orientation val>` — "maxMin" reverses the category
   *  axis left↔right. null/"minMax" ⇒ normal. */
  catAxisOrientation?: 'minMax' | 'maxMin' | string | null;
  /**
   * `<c:catAx><c:tickLblPos val>` (§21.2.2.207, `ST_TickLblPos` §21.2.3.47) —
   * "nextTo" (default) | "low" | "high" | "none". "none" hides the category tick
   * labels. null/undefined ⇒ nextTo (byte-stable).
   */
  catAxisTickLabelPos?: string | null;
  /** `<c:catAx><c:tickLblSkip val>` category-label interval. */
  catAxisTickLabelSkip?: number | null;
  /** `<c:catAx><c:tickMarkSkip val>` category-tick interval. */
  catAxisTickMarkSkip?: number | null;
  /** `<c:catAx><c:lblAlgn val>` tick-label text alignment. */
  catAxisLabelAlignment?: 'l' | 'ctr' | 'r' | string | null;
  /** `<c:catAx|dateAx><c:lblOffset val>` normalized 0–1000 percentage. */
  catAxisLabelOffsetPercent?: number | null;
  /** `<c:valAx><c:tickLblPos val>` (§21.2.2.207). "none" hides value tick labels. */
  valAxisTickLabelPos?: string | null;
  /**
   * `<c:catAx><c:txPr><a:bodyPr rot>` (DrawingML `ST_Angle`, 60000ths of a
   * degree) — category tick-label rotation. e.g. -2700000 = -45°. null/undefined
   * /0 ⇒ horizontal labels (byte-stable).
   */
  catAxisLabelRotation?: number | null;
  /** Group-owned decorations for classic line-chart groups. */
  lineGroupDecorations?: ChartLineGroupDecorations[] | null;
  /** Group-owned drop lines for classic area-chart groups. */
  areaGroupDecorations?: ChartAreaGroupDecorations[] | null;
  /** Group-owned series lines for classic bar-chart groups. */
  barGroupDecorations?: ChartBarGroupDecorations[] | null;
  // ── Stock chart (CH13, §21.2.2.198) ──────────────────────────────────────
  /** `<c:stockChart><c:dropLines>` direct DrawingML line paint. */
  stockDropLines?: ChartDecorationLineStyle | null;
  /** `<c:stockChart><c:hiLowLines>` direct DrawingML line paint. */
  stockHiLowLineStyle?: ChartDecorationLineStyle | null;
  /**
   * `<c:stockChart><c:hiLowLines>` presence (ECMA-376 §21.2.2.60). When true
   * the stock renderer draws a vertical line spanning each category's low↔high
   * value. Only set for `chartType === "stock"`; null/undefined on every other
   * chart type (byte-stable).
   */
  stockHiLowLines?: boolean | null;
  /**
   * `<c:hiLowLines><c:spPr><a:ln><a:solidFill>` resolved color (hex, no `#`).
   * null leaves direct paint omitted; linked/automatic layers resolve later.
   */
  stockHiLowLineColor?: string | null;
  /**
   * `<c:stockChart><c:upDownBars>` presence (ECMA-376 §21.2.2.227). Parsed so a
   * stock file carrying up/down bars draws them between the first and last
   * series (Open↔Close for four-series, High↔Close for three-series charts).
   * null/undefined when absent.
   */
  stockUpDownBars?: boolean | null;
  /** Parsed `<c:upDownBars>` geometry and direct up/down bar paint. */
  stockUpDownBarStyle?: ChartStockUpDownBarStyle | null;
  /**
   * Bounded Office automatic paint for otherwise-empty stock decorations.
   * This is resolved by the parser from the legacy chart style and theme; it
   * stays separate from authored/linked paint so precedence remains explicit.
   */
  stockAutomaticStyle?: {
    lineColor: string;
    lineWidthEmu: number;
    upFillColor: string;
    downFillColor: string;
  } | null;
  /** `<c:surfaceChart|surface3DChart><c:wireframe>` effective boolean. */
  surfaceWireframe?: boolean | null;
  /** `<c:bandFmts>` indexed low-to-high (§21.2.2.13/14). */
  surfaceBandFormats?: ChartSurfaceBandFormat[] | null;
  /** Legacy `<c:style@val>` (1..48). Automatic classic-chart palettes use
   * this authored style independently of any ChartEx sidecar. */
  legacyChartStyle?: number | null;
  /** Theme `accent1..accent6`, retained for renderer-generated classic chart
   * objects such as automatic surface value bands. */
  themeAccentColors?: string[] | null;
  /** Pie-of-pie / bar-of-pie secondary-plot contract (§21.2.2.126). */
  ofPie?: ChartOfPie | null;
  /** Authored 3D chart-space view and group depth controls. */
  threeD?: ChartThreeD | null;
  // ── chartEx structured layouts (CH15, MS 2014 chartex ext) ────────────────
  /**
   * Structured box-and-whisker data (`chartType === 'boxWhisker'`). Present
   * ONLY for boxWhisker charts; null/absent otherwise so the flat
   * `categories`/`series` model the other chartEx renderers consume is
   * untouched. The renderer computes quartiles / mean / whiskers / outliers.
   */
  chartexBox?: ChartexBoxWhisker | null;
  /**
   * Structured sunburst hierarchy (`chartType === 'sunburst'`). Present ONLY
   * for sunburst charts; null/absent otherwise.
   */
  chartexSunburst?: ChartexSunburst | null;
  /**
   * Structured treemap hierarchy (`chartType === 'treemap'`). Present ONLY
   * for treemap charts; null/absent otherwise.
   */
  chartexTreemap?: ChartexTreemap | null;
  /** Structured geospatial rows (`chartType === 'regionMap'`). */
  chartexRegionMap?: ChartexRegionMap | null;
  /** ChartEx histogram controls; raw observations remain in `series[0]`. */
  chartexHistogramBinning?: ChartexHistogramBinning | null;
  /**
   * Theme accent palette (`accent1..6`, hex without '#') for chartEx charts
   * that color by branch/series index (boxWhisker series and
   * sunburst/treemap branches).
   * null/absent when the resolver supplies no default palette (pptx); the
   * renderer then falls back to its own `CHART_PALETTE`.
   */
  chartexAccents?: string[] | null;
  /** Total color set resolved from the linked Chart Colors part. */
  chartexColorPalette?: Array<string | null> | null;
  /** `<cs:colorStyle meth>`; unknown methods have `cycle` semantics. */
  chartexColorStyleMethod?: string | null;
  /**
   * Bounded effective paint recipes for every paint-bearing linked Chart
   * Style role. Direct chart formatting remains in its existing fields and
   * takes precedence in the renderer; this map is a linked-style fallback.
   */
  chartStyleRoles?: Partial<Record<ChartStyleRole, ChartExElementStyle>> | null;
  /** Total resolved color set associated with the linked Chart Style part. */
  chartStyleColorPalette?: Array<string | null> | null;
  /** `<cs:colorStyle meth>` used when selecting a linked role color. */
  chartStyleColorMethod?: string | null;
  /** Linked `dataPointMarkerLayout@size`, in points (2..72). */
  chartStyleMarkerSizePt?: number | null;
  /** Linked `dataPointMarkerLayout@symbol` for marker-bearing data points. */
  chartStyleMarkerSymbol?: string | null;
  /** Effective `<cs:dataPoint>` style. */
  chartexDataPointStyle?: ChartExElementStyle | null;
  /** Effective `<cs:dataPointLine>` style for whiskers/median/connectors. */
  chartexDataPointLineStyle?: ChartExElementStyle | null;
  /** Effective `<cs:seriesLine>` style for waterfall connector lines. */
  chartexSeriesLineStyle?: ChartExElementStyle | null;
  /** Effective `<cs:dataPointMarker>` style for raw/outlier/mean markers. */
  chartexDataPointMarkerStyle?: ChartExElementStyle | null;
  /** Legacy ChartEx alias for `chartStyleMarkerSizePt`. */
  chartexMarkerSizePt?: number | null;
  /** Legacy ChartEx alias for `chartStyleMarkerSymbol`. */
  chartexMarkerSymbol?: string | null;
  /** `<cx:series><cx:layoutPr><cx:visibility connectorLines>` for waterfall. */
  chartexConnectorLines?: boolean | null;
}

export interface ChartStockBarPaint {
  fillColor?: string | null;
  fill?: SolidFill | GradientFill | PatternFill | null;
  /** Direct/linked fill owns this component even when it cannot be resolved. */
  fillPaintAuthored?: boolean | null;
  fillHidden?: boolean | null;
  lineColor?: string | null;
  /** Direct/linked outline owns this component even when it cannot be resolved. */
  linePaintAuthored?: boolean | null;
  lineWidthEmu?: number | null;
  lineDash?: string | null;
  lineCap?: string | null;
  lineJoin?: string | null;
  lineHidden?: boolean | null;
}

export interface ChartStockUpDownBarStyle {
  /** `<c:gapWidth val>`, percent of one bar width; omission defaults to 150. */
  gapWidthPercent: number;
  up: ChartStockBarPaint;
  down: ChartStockBarPaint;
}

export interface ChartDecorationLineStyle {
  color?: string | null;
  /** Direct/linked outline owns this component even when it cannot be resolved. */
  paintAuthored?: boolean | null;
  widthEmu?: number | null;
  dash?: string | null;
  cap?: string | null;
  join?: string | null;
  hidden?: boolean | null;
}

export interface ChartLineGroupDecorations {
  /** Zero-based document-order index among classic line-chart groups. */
  groupIndex: number;
  /** `<c:dropLines>` direct line paint; object presence means geometry exists. */
  dropLines?: ChartDecorationLineStyle | null;
  /** `<c:hiLowLines>` direct line paint; object presence means geometry exists. */
  hiLowLines?: ChartDecorationLineStyle | null;
  /** `<c:upDownBars>` geometry and direct bar paint. */
  upDownBars?: ChartStockUpDownBarStyle | null;
}

export interface ChartAreaGroupDecorations {
  /** Zero-based document-order index among classic area-chart groups. */
  groupIndex: number;
  /** `<c:dropLines>` direct line paint; object presence means geometry exists. */
  dropLines?: ChartDecorationLineStyle | null;
}

export interface ChartBarGroupDecorations {
  /** Zero-based document-order index among classic bar-chart groups. */
  groupIndex: number;
  /** Every authored `<c:serLines>` direct line paint, in document order. */
  seriesLines?: ChartDecorationLineStyle[] | null;
}

export interface ChartOfPie {
  type: 'pie' | 'bar';
  splitType: 'auto' | 'cust' | 'percent' | 'pos' | 'val';
  /** Whether `<c:splitType>` was authored rather than omitted. */
  splitTypeAuthored?: boolean | null;
  splitPos?: number | null;
  /** Whether `<c:splitPos>` was authored, including an invalid/missing value. */
  splitPosAuthored?: boolean | null;
  customSplitIndices?: number[] | null;
  /** Secondary pie diameter relative to the primary pie, 5–200%. */
  secondPieSizePercent: number;
  /** Gap between the primary and secondary plots, as a percent. */
  gapWidthPercent: number;
  seriesLines: boolean;
}

/** Supported DrawingML paint authored on a 3-D chart surface (`floor`,
 * `sideWall`, or `backWall`): solid/no fill and basic line properties. Each
 * surface is a real face of the shared projected scene, not a renderer
 * decoration. */
export interface ChartThreeDSurface {
  /** Full direct `<c:spPr>` paint/line recipe. */
  style?: ChartExElementStyle | null;
  fillColor?: string | null;
  fillHidden?: boolean | null;
  lineColor?: string | null;
  lineWidthEmu?: number | null;
  lineDash?: string | null;
  lineHidden?: boolean | null;
  /** `<c:thickness val>`, normalized to percent when authored. */
  thicknessPercent?: number | null;
  pictureOptions?: ChartThreeDPictureOptions | null;
}

export interface ChartThreeDPictureOptions {
  applyToFront?: boolean | null;
  applyToSides?: boolean | null;
  applyToEnd?: boolean | null;
  pictureFormat?: 'stretch' | 'stack' | 'stackScale' | string | null;
  /** Whether `<c:pictureFormat>` was authored, including an unsupported value. */
  pictureFormatAuthored?: boolean | null;
  pictureStackUnit?: number | null;
  /** Whether `<c:pictureStackUnit>` was authored, including an invalid value. */
  pictureStackUnitAuthored?: boolean | null;
}

export interface ChartThreeD {
  /** Whether `<c:view3D>` itself was authored. */
  view3DPresent?: boolean | null;
  rotationX?: number | null;
  rotationXAuthored?: boolean | null;
  rotationY?: number | null;
  rotationYAuthored?: boolean | null;
  heightPercent?: number | null;
  heightPercentAuthored?: boolean | null;
  depthPercent?: number | null;
  depthPercentAuthored?: boolean | null;
  perspective?: number | null;
  perspectiveAuthored?: boolean | null;
  rightAngleAxes?: boolean | null;
  rightAngleAxesAuthored?: boolean | null;
  gapDepthPercent?: number | null;
  gapDepthPercentAuthored?: boolean | null;
  shape?: string | null;
  /** `<c:bar3DChart><c:grouping val>` (§21.2.2.77). `standard` uses the
   *  series/depth axis; `clustered` uses adjacent category-axis slots. */
  barGrouping?: 'standard' | 'clustered' | 'stacked' | 'percentStacked' | string | null;
  /** `<c:serAx>` (§21.2.2.175), the series/depth axis of a standard 3-D bar. */
  seriesAxis?: ChartThreeDSeriesAxis | null;
  /** `<c:floor>` (§21.2.2.69), including direct DrawingML fill/line paint. */
  floor?: ChartThreeDSurface | null;
  /** `<c:sideWall>` (§21.2.2.191), including direct DrawingML fill/line paint. */
  sideWall?: ChartThreeDSurface | null;
  /** `<c:backWall>` (§21.2.2.11), including direct DrawingML fill/line paint. */
  backWall?: ChartThreeDSurface | null;
}

export interface ChartThreeDSeriesAxis {
  title?: string | null;
  hidden: boolean;
  orientation?: 'minMax' | 'maxMin' | string | null;
  tickLabelPos?: string | null;
  tickLabelSkip?: number | null;
  tickMarkSkip?: number | null;
  majorTickMark: string;
  /** `<c:serAx><c:minorTickMark>`; omission means no minor tick marks. */
  minorTickMark?: string | null;
  fontColor?: string | null;
  fontSizeHpt?: number | null;
  fontBold?: boolean | null;
  fontItalic?: boolean | null;
  fontFace?: string | null;
  lineColor?: string | null;
  lineWidthEmu?: number | null;
  lineDash?: string | null;
  /** A direct `<c:serAx><c:spPr><a:ln>` paint was authored. */
  linePaintAuthored?: boolean | null;
  lineHidden: boolean;
  titleFontSizeHpt?: number | null;
  titleFontBold?: boolean | null;
  titleFontItalic?: boolean | null;
  titleFontColor?: string | null;
  titleFontFace?: string | null;
  titleRotation?: number | null;
  titleVerticalMode?: ChartModel['catAxisTitleVerticalMode'];
  titleManualLayout?: ChartManualLayout | null;
}

/** A formatted DrawingML run inside a chart-relative text box. */
export interface ChartTextRun {
  text: string;
  fontSizeHpt?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  color?: string | null;
  /** This run/default authored a text-fill choice, resolved or otherwise. */
  colorPaintAuthored?: boolean | null;
  /** Effective direct `<a:noFill/>` for this run. */
  colorHidden?: boolean | null;
  fontFace?: string | null;
  language?: string | null;
  /** DrawingML baseline shift normalized to a fraction of the font size. */
  baseline?: number | null;
  /** Effective `<a:pPr algn>` for the paragraph that owns this run. */
  paragraphAlign?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | string | null;
}

/** One DrawingML paragraph inside a chart-relative text box. */
export interface ChartTextParagraph {
  runs: ChartTextRun[];
  align?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | string | null;
}

/** A text shape anchored to a fractional rectangle in the chart space. */
export interface ChartTextBox {
  x: number;
  y: number;
  w: number;
  h: number;
  paragraphs: ChartTextParagraph[];
  verticalAnchor?: 't' | 'ctr' | 'b' | 'just' | 'dist' | string | null;
  /** DrawingML `<a:bodyPr wrap>`; absent uses the application-default square wrap. */
  wrap?: 'none' | 'square' | string | null;
  /** DrawingML text-body insets in EMU. Omitted wire values use the
   * ECMA-376 defaults: lIns/rIns=91440, tIns/bIns=45720. */
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
}

/**
 * One box-and-whisker series (chartEx `boxWhisker`, MS 2014 chartex ext). Each
 * `<cx:series>` references its own raw sample points via `<cx:dataId>`; the
 * parser groups them by category and threads the `<cx:layoutPr>` flags. The
 * renderer derives the statistics.
 */
export interface ChartexBoxSeries {
  /** Series display name (`<cx:tx><cx:v>`). */
  name: string;
  /** Effective ChartEx `CT_Series@formatIdx`; omitted authoring resolves to
   * the original document-order series index. */
  chartexFormatIdx?: number | null;
  /** Explicit `<cx:series><cx:spPr>` fill (hex, no '#'). null = resolve the
   *  Chart Style / linked Chart Colors, then fall back to the theme palette. */
  color?: string | null;
  /** Explicit `<cx:series><cx:spPr><a:ln>` outline color (hex, no '#'). */
  lineColor?: string | null;
  /** Explicit series outline width from `<a:ln@w>` (EMU). */
  lineWidthEmu?: number | null;
  /** Lossless ChartEx series-local `<cx:spPr>` paint. */
  chartexStyle?: ChartExElementStyle | null;
  /** Raw sample values grouped by category (outer = category index parallel to
   *  {@link ChartexBoxWhisker.categories}, inner = the points in that group). */
  valuesByCategory: number[][];
  /** `<cx:visibility meanMarker>` — draw the mean `×`. */
  meanMarker: boolean;
  /** `<cx:visibility meanLine>` — draw a mean connector line across categories. */
  meanLine: boolean;
  /** `<cx:visibility outliers>` — draw outlier points. */
  showOutliers: boolean;
  /** `<cx:visibility nonoutliers>` — draw the interior (non-outlier) sample
   *  points as dots on top of the box. */
  showNonoutliers: boolean;
  /** `<cx:statistics quartileMethod>` — "exclusive" (Excel default) | "inclusive". */
  quartileMethod: string;
}

/** A chartEx box-and-whisker chart: unique categories + one series per column. */
export interface ChartexBoxWhisker {
  /** True when the source omits a category dimension and each ChartEx series
   * is itself one category/box. Such diagonal data must not be clustered a
   * second time by the total series count. */
  oneBoxPerSeries?: boolean;
  /** Unique category labels in first-seen order. */
  categories: string[];
  /** One entry per `<cx:series>`. */
  series: ChartexBoxSeries[];
}

/**
 * One row of a chartEx `sunburst`: the branch→…→leaf label chain (empty
 * trailing segments trimmed) and its size value.
 */
export interface ChartexSunburstRow {
  /** Label chain root→leaf. */
  path: string[];
  /** `<cx:numDim type="size">` value attaching to the deepest node in `path`. */
  size: number;
}

/** A chartEx sunburst: the flat rows the renderer folds into a ring tree. */
export interface ChartexSunburst {
  rows: ChartexSunburstRow[];
}

/** A chartEx treemap hierarchy and its requested parent-label presentation. */
export interface ChartexTreemap {
  rows: ChartexSunburstRow[];
  /** `<cx:parentLabelLayout val>`: `banner`, `overlapping`, or `none`. */
  parentLabelLayout?: string | null;
}

/** One source row of a ChartEx geospatial series. */
export interface ChartexRegionMapRow {
  /** Authored category text; no parser-side geocoding or alias rewriting. */
  label: string;
  /** Optional stable geography identity from `strDim@type="entityId"`. */
  entityId?: string | null;
  /** `numDim@type="colorVal"`. */
  value?: number | null;
}

/** Authored ChartEx geography metadata. Opaque provider cache bytes are not
 * exposed; only their presence/provider are retained. */
export interface ChartexGeography {
  projectionType?: 'mercator' | 'miller' | 'robinson' | 'albers' | string | null;
  viewedRegionType?: string | null;
  cultureLanguage?: string | null;
  cultureRegion?: string | null;
  attribution?: string | null;
  cacheProvider?: string | null;
  cachePresent: boolean;
}

export interface ChartexValueColorStop {
  kind: 'extremeValue' | 'number' | 'percent' | string;
  value?: number | null;
}

export interface ChartexRegionMapColors {
  /** CT_ValueColorPositions@count; omitted is the schema default 2. */
  stopCount?: 2 | 3 | null;
  minColor?: string | null;
  midColor?: string | null;
  maxColor?: string | null;
  minPosition?: ChartexValueColorStop | null;
  midPosition?: ChartexValueColorStop | null;
  maxPosition?: ChartexValueColorStop | null;
}

/** Data-only ChartEx Region Map model. Geography lookup and finite offline
 * geometry are owned by the core renderer, never by a parser host. */
export interface ChartexRegionMap {
  rows: ChartexRegionMapRow[];
  regionLabelLayout?: 'none' | 'bestFitOnly' | 'showAll' | null;
  geography?: ChartexGeography | null;
  colors?: ChartexRegionMapColors | null;
}

/** ChartEx `CT_Binning` controls retained for histogram aggregation. */
export interface ChartexHistogramBinning {
  binSize?: number | null;
  binCount?: number | null;
  intervalClosed?: 'l' | 'r' | null;
  underflow?: number | null;
  overflow?: number | null;
}

/**
 * A secondary value axis (combo charts). Mirrors the primary value-axis
 * properties but lives in its own object so the flat primary-axis fields stay
 * untouched. Parsed from the right-hand `<c:valAx>` (`axPos="r"`,
 * `<c:crosses val="max">`).
 */
export interface SecondaryValueAxis {
  /** `<c:scaling><c:min val>`. null = derive from the series data. */
  min: number | null;
  /** `<c:scaling><c:max val>`. null = derive from the series data. */
  max: number | null;
  /** `<c:title>` plain text. null = no title. */
  title: string | null;
  /** `<c:delete val="1"/>` — hide labels/ticks entirely. */
  hidden: boolean;
  /** `<c:numFmt formatCode>` for tick labels. */
  formatCode?: string | null;
  /** `<c:dispUnits>` for this auxiliary value axis. */
  displayUnits?: ChartDisplayUnits | null;
  /** `<c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  fontColor?: string | null;
  /** `<c:txPr>` tick-label font size (hpt). */
  fontSizeHpt?: number | null;
  /** `<c:txPr>` tick-label italic flag. */
  fontItalic?: boolean | null;
  /** `<c:txPr>` tick-label bold flag. */
  fontBold?: boolean | null;
  /** `<c:txPr>…<a:latin typeface>` tick-label font face. */
  fontFace?: string | null;
  /** `<c:spPr><a:ln><a:solidFill>` axis-line color (hex without '#'). */
  lineColor?: string | null;
  /** `<c:spPr><a:ln w>` axis-line width in EMU. */
  lineWidthEmu?: number | null;
  /** `<c:spPr><a:ln><a:prstDash val>` axis-line dash preset. */
  lineDash?: string | null;
  /** `<c:spPr><a:ln><a:noFill>` — Office-compatible suppression of the
   *  secondary axis rule and tick marks; labels and gridlines remain. */
  lineHidden: boolean;
  /** `<c:majorTickMark>` — "cross" (default) | "out" | "in" | "none". */
  majorTickMark: string;
  /** `<c:minorTickMark>` — omitted means no minor ticks. */
  minorTickMark?: string | null;
  /** `<c:minorGridlines>` independently requests plot-area lines. */
  minorGridlines?: boolean;
  minorGridlineColor?: string | null;
  minorGridlineWidthEmu?: number | null;
  minorGridlineDash?: string | null;
  /** `<c:majorGridlines>` presence and authored line paint. */
  majorGridlines?: boolean;
  majorGridlineColor?: string | null;
  majorGridlineWidthEmu?: number | null;
  majorGridlineDash?: string | null;
  /**
   * `<c:valAx><c:majorUnit val>` (§21.2.2.103) — explicit distance between
   * major ticks/gridlines on THIS secondary axis, overriding the Excel-style
   * auto "nice" step. null/undefined ⇒ auto step (byte-stable). Symmetric with
   * {@link ChartModel.valAxisMajorUnit} on the primary axis.
   */
  majorUnit?: number | null;
  /** `<c:valAx><c:minorUnit val>` explicit minor-tick step; omitted minor ticks
   *  use this axis's automatic major unit divided by five. */
  minorUnit?: number | null;
  /** `<c:scaling><c:logBase>`; null means linear. */
  logBase?: number | null;
  /** `<c:scaling><c:orientation>`. */
  orientation?: 'minMax' | 'maxMin' | string | null;
  /** `<c:tickLblPos>`; `none` hides tick labels without hiding gridlines. */
  tickLabelPos?: string | null;
  /** `<c:catAx><c:lblAlgn>` when used as `secondaryCatAxis`. */
  labelAlignment?: 'l' | 'ctr' | 'r' | null;
  /** `<c:catAx|dateAx><c:lblOffset>` when used as `secondaryCatAxis`. */
  labelOffsetPercent?: number | null;
  /** `<c:catAx><c:tickLblSkip>` when used as `secondaryCatAxis`. */
  tickLabelSkip?: number | null;
  /** `<c:catAx><c:tickMarkSkip>` when used as `secondaryCatAxis`. */
  tickMarkSkip?: number | null;
  /** `<c:crosses>` / `<c:crossesAt>` retained for axis placement. */
  crosses?: string | null;
  crossesAt?: number | null;
  /** `<c:title>` run-prop font size (hpt). */
  titleFontSizeHpt?: number | null;
  /** `<c:title>` run-prop bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:title>` run-prop italic flag. */
  titleFontItalic?: boolean | null;
  /** `<c:title>` run-prop color (hex without '#'). */
  titleFontColor?: string | null;
  titleFontFace?: string | null;
  /** Authored `<c:title>` DrawingML `bodyPr@rot` in raw `ST_Angle` units. */
  titleRotation?: number | null;
  /** Authored `<c:title>` DrawingML `bodyPr@vert`. */
  titleVerticalMode?:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null;
  /** `<c:title><c:layout><c:manualLayout>` for this auxiliary axis. */
  titleManualLayout?: ChartManualLayout | null;
}

/** ECMA-376 §21.2.2.45 display-unit scaling. The divisor changes displayed
 * axis-associated values (ticks and generated `showVal` data-label text),
 * never the source value or value-to-pixel mapping. */
export interface ChartDisplayUnits {
  divisor: number;
  /** Authored `ST_BuiltInUnit` token; absent for `<c:custUnit>`. */
  builtInUnit?: string | null;
  /** `<c:dispUnitsLbl>`; absent means the scale is applied without a label. */
  label?: ChartDisplayUnitsLabel | null;
}

/** ECMA-376 §21.2.2.46 display-unit label, independently styled and laid out
 * from the axis title. */
export interface ChartDisplayUnitsLabel {
  /** Explicit `<c:tx>` text. null/undefined uses the unit's automatic name. */
  text?: string | null;
  manualLayout?: ChartManualLayout | null;
  fontSizeHpt?: number | null;
  fontBold?: boolean | null;
  fontItalic?: boolean | null;
  fontColor?: string | null;
  fontFace?: string | null;
  /** DrawingML `bodyPr@rot`, in 60000ths of a degree. */
  rotation?: number | null;
  boxStyle?: ChartLabelBox | null;
}

/**
 * `<c:manualLayout>` block. Fractions are of the chart-space rect.
 * `xMode`/`yMode`: "edge" = absolute fraction from top-left, "factor" =
 * fraction offset from default position.
 */
export interface ChartManualLayout {
  xMode?: string;
  yMode?: string;
  wMode?: string;
  hMode?: string;
  layoutTarget?: string;
  x: number;
  y: number;
  w?: number;
  h?: number;
}

export interface LegendManualLayout {
  /** `"edge"` = `x`/`y` are fractions from top-left of chart space;
   *  `"factor"` = fractions offset from the default position. */
  xMode?: string;
  yMode?: string;
  wMode?: string;
  hMode?: string;
  /** Fractions of chart space width/height. */
  x: number;
  y: number;
  w?: number;
  h?: number;
}

/** Classic-chart `<c:legendEntry>` (§21.2.2.94 / CT_LegendEntry). */
export interface ChartLegendEntryOverride {
  idx: number;
  /** `<c:delete>`; a bare element is true. */
  deleted?: boolean | null;
  /** Entry-local `<c:txPr>` properties. Omitted properties inherit the
   *  chart-level legend text style. */
  fontFace?: string | null;
  fontColor?: string | null;
  fontSizeHpt?: number | null;
  fontBold?: boolean | null;
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
