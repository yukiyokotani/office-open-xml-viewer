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
  | { type: 'text'; text: string }
  | { type: 'number'; number: number }
  | { type: 'bool'; bool: boolean }
  | { type: 'error'; error: string };

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
}

export interface Font {
  bold: boolean;
  italic: boolean;
  underline: boolean;
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
}

export interface ParsedWorkbook {
  workbook: Workbook;
  styles: Styles;
  sharedStrings: string[];
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
}

export type WorkerRequest =
  | { type: 'parse'; data: ArrayBuffer }
  | { type: 'parseSheet'; data: ArrayBuffer; sheetIndex: number; sheetName: string };

export type WorkerResponse =
  | { type: 'parsed'; workbook: ParsedWorkbook }
  | { type: 'parsedSheet'; worksheet: Worksheet }
  | { type: 'error'; message: string };
