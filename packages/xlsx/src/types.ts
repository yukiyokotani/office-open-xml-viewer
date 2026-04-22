export interface Workbook {
  sheets: SheetMeta[];
}

export interface SheetMeta {
  name: string;
  sheetId: number;
  rId: string;
}

export interface MergeCell {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface Worksheet {
  name: string;
  rows: Row[];
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  defaultColWidth: number;
  defaultRowHeight: number;
  mergeCells: MergeCell[];
  freezeRows: number;
  freezeCols: number;
  conditionalFormats: ConditionalFormat[];
  images: ImageAnchor[];
  charts: ChartAnchor[];
  /** Whether to display zero values (ECMA-376 §18.3.1.94). Defaults to true. */
  showZeros?: boolean;
  /** Whether to draw default grid lines (ECMA-376 §18.3.1.83
   *  `<sheetView showGridLines>`). Mirrors the Excel "View → Gridlines"
   *  checkbox. Defaults to true. */
  showGridlines?: boolean;
  /** Sheet tab color (ECMA-376 §18.3.1.79). */
  tabColor?: string | null;
  /** AutoFilter header range (ECMA-376 §18.3.1.2). */
  autoFilter?: CellRange | null;
  /** Hyperlinks in this worksheet (ECMA-376 §18.3.1.47). */
  hyperlinks?: Hyperlink[];
  /** A1-style cell refs of commented cells (ECMA-376 §18.7.3). Rendered as a
   *  small red triangle in each cell's top-right corner. */
  commentRefs?: string[];
  /** Defined names in scope for this sheet (ECMA-376 §18.2.5). Used by
   *  conditional-formatting `expression` rules that call named ranges
   *  (e.g. `task_start`, `today`). */
  definedNames?: DefinedName[];
  /** Excel Tables on this sheet (ECMA-376 §18.5). The renderer overlays a
   *  built-in style (bold header, banded rows) on the given ranges. */
  tables?: TableInfo[];
}

export interface TableInfo {
  range: CellRange;
  styleName: string;
  headerRowCount: number;
  totalsRowCount: number;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  /** Accent color resolved by the parser from the built-in style name against
   *  the file's theme accents (e.g. `TableStyleLight18` → accent3). */
  accentColor: string;
}

export interface DefinedName {
  name: string;
  formula: string;
}

// ─── Chart types ─────────────────────────────────────────────────────────────

/**
 * XLSX parser's raw chart series (includes XLSX-only `seriesType` for mixed
 * charts). Adapted to `ChartSeries` from @silurus/ooxml-core before rendering.
 */
export interface XlsxChartSeries {
  name: string;
  /** Chart sub-type for this series (allows mixed charts). */
  seriesType: string;
  categories: string[];
  values: (number | null)[];
  /** Explicit fill color hex (from c:spPr). Undefined = use palette. */
  color?: string | null;
  /** Marker visibility resolved from `<c:marker>`/chart-level default
   *  (ECMA-376 §21.2.2.32). */
  showMarker?: boolean;
}

/**
 * XLSX parser's raw chart output. Retains parser-native `barDir` + `grouping`
 * which the renderer combines into a canonical `ChartModel.chartType` (e.g.
 * `clusteredBarH`, `stackedBarPct`) at render time.
 */
export interface ChartData {
  /** Primary chart type: "bar"|"line"|"area"|"pie"|"doughnut"|"radar"|"scatter" */
  chartType: string;
  /** "col" (vertical bars) | "row" (horizontal bars) */
  barDir: string;
  /** "clustered"|"stacked"|"standard"|"percentStacked" */
  grouping: string;
  title: string | null;
  categories: string[];
  series: XlsxChartSeries[];
  /** Whether data labels are enabled (c:dLbls showVal/showPercent). */
  showDataLabels?: boolean;
  /** Category axis title (c:catAx/c:title). */
  catAxisTitle?: string | null;
  /** Value axis title (c:valAx/c:title). */
  valAxisTitle?: string | null;
  /** True when <c:legend> is present. Absence means the legend is hidden. */
  showLegend?: boolean;
  /** `<c:legendPos val>` — "r"|"l"|"t"|"b"|"tr". null/undefined = default ("r"). */
  legendPos?: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** Chart title font size in OOXML hundredths of a point (e.g. 1400 = 14pt). */
  titleFontSizeHpt?: number | null;
  /** Chart title font color as a hex string without '#' (srgbClr only). */
  titleFontColor?: string | null;
  /** Chart title font family from `<a:latin typeface>` (ECMA-376 §20.1.4.2.24). */
  titleFontFace?: string | null;
  /** Category axis tick-label font size in hpt (ECMA-376 §21.2.2.17 c:txPr). */
  catAxisFontSizeHpt?: number | null;
  /** Value axis tick-label font size in hpt. */
  valAxisFontSizeHpt?: number | null;
}

export interface ChartAnchor {
  fromCol: number; fromColOff: number;
  fromRow: number; fromRowOff: number;
  toCol: number;   toColOff: number;
  toRow: number;   toRowOff: number;
  chart: ChartData;
}

/**
 * Image anchored to a rectangle of cells (EMU offsets within the anchor cells).
 * 914400 EMU = 1 inch, 9525 EMU = 1 px @ 96 DPI.
 */
export interface ImageAnchor {
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  /** Data URL (data:image/png;base64,...) */
  dataUrl: string;
}

export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface Hyperlink {
  col: number;
  row: number;
  url: string | null;
}

export interface ConditionalFormat {
  sqref: CellRange[];
  rules: CfRule[];
}

export type CfRule =
  | { type: 'cellIs'; operator: string; formulas: string[]; dxfId: number | null; priority: number }
  | { type: 'expression'; formula: string; dxfId: number | null; priority: number; stopIfTrue: boolean }
  | { type: 'colorScale'; stops: CfStop[]; priority: number }
  | { type: 'dataBar'; color: string; min: CfValue; max: CfValue; priority: number; gradient: boolean }
  | { type: 'top10'; top: boolean; percent: boolean; rank: number; dxfId: number | null; priority: number }
  | { type: 'aboveAverage'; aboveAverage: boolean; dxfId: number | null; priority: number }
  | { type: 'iconSet'; iconSet: string; cfvos: CfValue[]; reverse: boolean; priority: number; customIcons?: CfIcon[] }
  | { type: 'other'; kind: string; priority: number };

export interface CfIcon {
  iconSet: string;
  iconId: number;
}

export interface CfStop {
  kind: string;
  value: string | null;
  color: string;
}

export interface CfValue {
  kind: string;
  value: string | null;
}

export interface Row {
  index: number;
  height: number | null;
  cells: Cell[];
}

export interface Cell {
  col: number;
  row: number;
  colRef: string;
  value: CellValue;
  styleIndex: number;
  /** Raw `<f>` formula text (ECMA-376 §18.3.1.40), when present. The renderer
   *  uses this to recompute volatile functions (TODAY, NOW) at display time
   *  so the cached `<v>` — frozen when the file was last saved — doesn't
   *  show a stale date. */
  formula?: string;
}

export type CellValue =
  | { type: 'empty' }
  | { type: 'text'; text: string; runs?: Run[] }
  | { type: 'number'; number: number }
  | { type: 'bool'; bool: boolean }
  | { type: 'error'; error: string };

export interface Run {
  text: string;
  font?: RunFont;
}

export interface RunFont {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  size?: number;
  color?: string | null;
  name?: string | null;
}

export interface SharedString {
  text: string;
  runs?: Run[];
}

export interface NumFmt {
  numFmtId: number;
  formatCode: string;
}

export interface Styles {
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  cellXfs: CellXf[];
  numFmts: NumFmt[];
  dxfs: Dxf[];
}

export interface Dxf {
  font: Font | null;
  fill: Fill | null;
  border: Border | null;
}

export interface Font {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  size: number;
  color: string | null;
  name: string | null;
}

export interface Fill {
  patternType: string;
  fgColor: string | null;
  bgColor: string | null;
  /** Set when the style's `<fill>` was a `<gradientFill>`; patternType stays "none". */
  gradient?: GradientFillSpec | null;
}

export interface GradientFillSpec {
  /** "linear" (default) or "path". */
  gradientType: string;
  /** Rotation in degrees for linear gradients (0 = left→right). */
  degree: number;
  /** Path-gradient bounding box (0..1) — unused for linear. */
  left: number;
  right: number;
  top: number;
  bottom: number;
  stops: { position: number; color: string }[];
}

export interface Border {
  left: BorderEdge | null;
  right: BorderEdge | null;
  top: BorderEdge | null;
  bottom: BorderEdge | null;
  diagonalUp?: BorderEdge | null;
  diagonalDown?: BorderEdge | null;
}

export interface BorderEdge {
  style: string;
  color: string | null;
}

export interface CellXf {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  alignH: string | null;
  alignV: string | null;
  wrapText: boolean;
  /** Indentation level (each level ≈ 3 characters, ECMA-376 §18.8.44) */
  indent?: number;
  /** Text rotation: 1–90 = counter-clockwise °, 91–180 = (val−90)° clockwise, 255 = stacked */
  textRotation?: number;
  shrinkToFit?: boolean;
}

export interface ParsedWorkbook {
  workbook: Workbook;
  styles: Styles;
  sharedStrings: SharedString[];
}

export interface ViewportRange {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

export interface RenderViewportOptions {
  width?: number;
  height?: number;
  dpr?: number;
  defaultFontFamily?: string;
  defaultFontSize?: number;
  scrollOffsetX?: number;
  scrollOffsetY?: number;
  freezeRows?: number;
  freezeCols?: number;
  /** Scale factor applied to all cell/header dimensions (default 1). */
  cellScale?: number;
  /** Pre-loaded Image elements keyed by their dataUrl (for ImageAnchor rendering). */
  loadedImages?: Map<string, HTMLImageElement>;
}

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; data: ArrayBuffer }
  | { type: 'parseSheet'; data: ArrayBuffer; sheetIndex: number; sheetName: string };

export type WorkerResponse =
  | { type: 'parsed'; workbook: ParsedWorkbook }
  | { type: 'parsedSheet'; worksheet: Worksheet }
  | { type: 'error'; message: string };
