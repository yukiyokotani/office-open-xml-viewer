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
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
