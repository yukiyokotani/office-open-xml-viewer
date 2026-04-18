// ===== Output JSON model (mirrors Rust types) =====

export interface Document {
  section: SectionProps;
  body: BodyElement[];
  headers: HeadersFooters;
  footers: HeadersFooters;
}

export interface HeadersFooters {
  default: HeaderFooter | null;
  first: HeaderFooter | null;
  even: HeaderFooter | null;
}

export interface HeaderFooter {
  body: BodyElement[];
}

export interface SectionProps {
  pageWidth: number;   // pt
  pageHeight: number;  // pt
  marginTop: number;   // pt
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  headerDistance: number;   // pt — top of page to header
  footerDistance: number;   // pt — bottom of page to footer
  titlePage: boolean;
  evenAndOddHeaders: boolean;
}

export type BodyElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable
  | { type: 'pageBreak' };

export interface DocParagraph {
  alignment: 'left' | 'center' | 'right' | 'justify';
  indentLeft: number;   // pt
  indentRight: number;  // pt
  indentFirst: number;  // pt
  spaceBefore: number;  // pt
  spaceAfter: number;   // pt
  lineSpacing: LineSpacing | null;
  numbering: NumberingInfo | null;
  runs: DocRun[];
}

export interface LineSpacing {
  value: number;   // multiplier (auto/atLeast) or pt (exact)
  rule: 'auto' | 'exact' | 'atLeast';
}

export interface NumberingInfo {
  numId: number;
  level: number;
  format: string;
  text: string;       // resolved bullet text, e.g. "1." or "•"
  indentLeft: number; // pt
  tab: number;        // pt
}

export type DocRun =
  | { type: 'text' } & TextRun
  | { type: 'image' } & ImageRun
  | { type: 'break'; breakType: 'line' | 'page' | 'column' }
  | { type: 'field' } & FieldRun;

export interface FieldRun {
  /** "page" | "numPages" | "other" */
  fieldType: string;
  instruction: string;
  fallbackText: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  background: string | null;
}

export interface TextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  isLink: boolean;
  background: string | null;
}

export interface ImageRun {
  dataUrl: string;
  widthPt: number;
  heightPt: number;
  /** true = wp:anchor (absolute positioned), false/undefined = wp:inline (flows with text) */
  anchor?: boolean;
  /** X offset in pt (anchor only) */
  anchorXPt?: number;
  /** Y offset in pt (anchor only) */
  anchorYPt?: number;
  /**
   * If true, anchorXPt is relative to the left margin — add section.marginLeft to get page X.
   * If false/absent, anchorXPt is already page-absolute.
   */
  anchorXFromMargin?: boolean;
  /**
   * If true, anchorYPt is relative to the paragraph's top Y in the renderer.
   * If false/absent, anchorYPt is already page-absolute.
   */
  anchorYFromPara?: boolean;
  /**
   * When set, the renderer replaces all pixels of this hex color (e.g. "FFFFFF") with full
   * transparency. Implements a:clrChange (make-background-transparent).
   */
  colorReplaceFrom?: string;
}

// ===== Table =====

export interface DocTable {
  colWidths: number[];  // pt
  rows: DocTableRow[];
  borders: TableBorders;
  cellMarginTop: number;
  cellMarginBottom: number;
  cellMarginLeft: number;
  cellMarginRight: number;
}

export interface TableBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
  insideH: BorderSpec | null;
  insideV: BorderSpec | null;
}

export interface BorderSpec {
  width: number;   // pt
  color: string | null;
  style: string;
}

export interface DocTableRow {
  cells: DocTableCell[];
  rowHeight: number | null;  // pt
  isHeader: boolean;
}

export interface DocTableCell {
  content: DocParagraph[];
  colSpan: number;
  vMerge: boolean | null;  // null=no merge, true=start, false=continuation
  borders: CellBorders;
  background: string | null;
  vAlign: 'top' | 'center' | 'bottom';
  widthPt: number | null;
}

export interface CellBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; data: ArrayBuffer };

export type WorkerResponse =
  | { type: 'parsed'; document: Document }
  | { type: 'error'; message: string };

// ===== Public API types =====

export interface RenderPageOptions {
  /** Canvas CSS width in px; height is auto-computed from page aspect ratio */
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
}
