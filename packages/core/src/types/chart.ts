// ===== Unified chart model =====
// Shared by @silurus/ooxml-pptx and @silurus/ooxml-xlsx.
//
// Parser JSON from each format is adapted into `ChartModel` and then passed
// to `renderChart` in @silurus/ooxml-core. This keeps a single source of
// truth for chart rendering across PPTX / XLSX (and future DrawingML charts
// in DOCX).

export interface ChartSeries {
  name: string;
  /** Hex without '#'. null = fall back to palette. */
  color: string | null;
  /** Numeric values; null = missing data point. */
  values: (number | null)[];
  /**
   * Per-data-point colors (pie / doughnut). Hex without '#'. null inside the
   * array = use palette for that slice. Omit entirely for non-pie series.
   */
  dataPointColors?: (string | null)[] | null;
  /**
   * Mixed chart: per-series chart type override. Currently only "line" (XLSX)
   * is honoured; other values are treated as the chart's primary type.
   */
  seriesType?: string | null;
  /**
   * Scatter-only X values (as strings). When null the series uses
   * `ChartModel.categories` as X.
   */
  categories?: string[] | null;
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
  | string;

export interface ChartModel {
  chartType: ChartType;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
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
  /** Hex without '#'. From `<c:plotArea><c:spPr><a:solidFill>`. */
  plotAreaBg: string | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** True when `<c:legend>` is declared in the chart XML. False = no legend. */
  showLegend: boolean;
  /** `<c:legend><c:legendPos val>` — "r"|"l"|"t"|"b"|"tr". null = default (r). */
  legendPos: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** `<c:catAx><c:crossBetween val="..."/>`. "between" inserts 0.5-step padding
   *  on each end of the category axis; "midCat" anchors endpoints to the axes. */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. ECMA-376 default is "cross". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
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
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
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
}

export interface LegendManualLayout {
  /** `"edge"` = `x`/`y` are fractions from top-left of chart space;
   *  `"factor"` = fractions offset from the default position. */
  xMode: string;
  yMode: string;
  /** Fractions of chart space width/height. */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
