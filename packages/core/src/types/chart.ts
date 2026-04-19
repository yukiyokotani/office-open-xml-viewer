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
  /** Waterfall subtotal category indices. */
  subtotalIndices: number[];
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
