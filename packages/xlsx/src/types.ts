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
  /** Sheet tab color (ECMA-376 §18.3.1.79). */
  tabColor?: string | null;
  /** AutoFilter header range (ECMA-376 §18.3.1.2). */
  autoFilter?: CellRange | null;
  /** Hyperlinks in this worksheet (ECMA-376 §18.3.1.47). */
  hyperlinks?: Hyperlink[];
}

// ─── Chart types ─────────────────────────────────────────────────────────────

export interface ChartSeries {
  name: string;
  /** Chart sub-type for this series (allows mixed charts). */
  seriesType: string;
  categories: string[];
  values: (number | null)[];
}

export interface ChartData {
  /** Primary chart type: "bar"|"line"|"area"|"pie"|"doughnut"|"radar"|"scatter" */
  chartType: string;
  /** "col" (vertical bars) | "row" (horizontal bars) */
  barDir: string;
  /** "clustered"|"stacked"|"standard"|"percentStacked" */
  grouping: string;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
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
  | { type: 'expression'; formula: string; dxfId: number | null; priority: number }
  | { type: 'colorScale'; stops: CfStop[]; priority: number }
  | { type: 'dataBar'; color: string; min: CfValue; max: CfValue; priority: number }
  | { type: 'top10'; top: boolean; percent: boolean; rank: number; dxfId: number | null; priority: number }
  | { type: 'aboveAverage'; aboveAverage: boolean; dxfId: number | null; priority: number }
  | { type: 'iconSet'; iconSet: string; cfvos: CfValue[]; reverse: boolean; priority: number }
  | { type: 'other'; kind: string; priority: number };

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
