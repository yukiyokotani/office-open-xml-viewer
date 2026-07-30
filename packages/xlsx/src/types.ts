import type { ChartModel, MathNode, SpaceLine } from '@silurus/ooxml-core';

export interface Workbook {
  sheets: SheetMeta[];
  /** Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28).
   *  `true` selects the 1904 date system (Mac-authored workbooks); serial
   *  dates are resolved against the 1904 epoch (§18.17.4.1). Omitted from the
   *  parser JSON when false (default 1900 date system). */
  date1904?: boolean;
  /** #773 partial degradation: a WORKBOOK-LEVEL degradation that leaves every
   *  sheet openable. Set when a shared workbook part was PRESENT but corrupt —
   *  most commonly `xl/sharedStrings.xml` (§18.4.9): a broken shared-string table
   *  silently blanks every string cell across ALL sheets, so unlike a per-sheet
   *  break it can't be attributed to one placeholder sheet. Tagged with the
   *  offending part (e.g. `"xl/sharedStrings.xml: <detail>"`) so the loss is
   *  surfaced instead of silent, while every sheet still renders its non-string
   *  content. Absent (`undefined`) when every shared part read cleanly. Also set
   *  (`"(zip container): <detail>"`) for a whole-container degradation (#774). */
  parseError?: string;
}

/** Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19 `ST_SheetState`). */
export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

export interface SheetMeta {
  name: string;
  sheetId: number;
  rId: string;
  /** Sheet tab color (`<sheetPr><tabColor>`, ECMA-376 §18.3.1.93) resolved to
   *  `#RRGGBB`. Surfaced at workbook-list time so tabs can be painted up front.
   *  Absent when the sheet declares no tab color. */
  tabColor?: string | null;
  /** Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19 `ST_SheetState`).
   *  Absent ⇒ visible. `'veryHidden'` sheets are revealable only
   *  programmatically in Excel. Read via `XlsxWorkbook.isHidden` /
   *  `XlsxWorkbook.sheetVisibility`. */
  visibility?: 'hidden' | 'veryHidden';
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
  /** Per-column outline (grouping) depth 0-7 (ECMA-376 §18.3.1.13
   *  `<col outlineLevel>`), keyed by 1-based column index. Present only for
   *  grouped columns; absent (⇒ level 0) on outline-free sheets. */
  colOutlineLevels?: Record<number, number>;
  /** Per-column `<col collapsed>` (§18.3.1.13): `true` on a summary column whose
   *  one-level-deeper detail columns are collapsed. Only `true` entries. */
  colCollapsed?: Record<number, boolean>;
  /** Per-column `<col hidden>` (§18.3.1.13): `true` when the column is hidden
   *  (e.g. a collapsed outline hides its detail columns). Distinct from
   *  `colWidths[c] === 0`. Only `true` entries. */
  colHidden?: Record<number, boolean>;
  defaultColWidth: number;
  defaultRowHeight: number;
  mergeCells: MergeCell[];
  freezeRows: number;
  freezeCols: number;
  conditionalFormats: ConditionalFormat[];
  images: ImageAnchor[];
  charts: ChartAnchor[];
  /** Grouped shapes from `<xdr:grpSp>` inside twoCellAnchors
   *  (ECMA-376 §20.5.2.17). Each anchor holds leaf shapes pre-flattened
   *  with normalized [0,1] geometry relative to the anchor rect. */
  shapeGroups?: ShapeAnchor[];
  /** Whether to display zero values (ECMA-376 §18.3.1.94). Defaults to true. */
  showZeros?: boolean;
  /** Whether to draw default grid lines (ECMA-376 §18.3.1.83
   *  `<sheetView showGridLines>`). Mirrors the Excel "View → Gridlines"
   *  checkbox. Defaults to true. */
  showGridlines?: boolean;
  /** Whether the sheet grid is laid out right-to-left, mirroring the entire
   *  grid so column A sits on the right (ECMA-376 §18.3.1.87
   *  `<sheetView rightToLeft>`). Defaults to false. */
  rightToLeft?: boolean;
  /** Outline display flags from `<sheetPr><outlinePr>` (ECMA-376 §18.3.1.61).
   *  Absent when the sheet declares no `<outlinePr>`; consumers apply the
   *  schema defaults (`summaryBelow` / `summaryRight` both `true`). Decides
   *  which side of a group the summary row/column (and its +/- toggle) sits. */
  outlinePr?: OutlinePr;
  /** Sheet tab color (ECMA-376 §18.3.1.79). */
  tabColor?: string | null;
  /** AutoFilter header range (ECMA-376 §18.3.1.2). */
  autoFilter?: CellRange | null;
  /** Hyperlinks in this worksheet (ECMA-376 §18.3.1.47). */
  hyperlinks?: Hyperlink[];
  /** A1-style cell refs of commented cells (ECMA-376 §18.7.3). Rendered as a
   *  small red triangle in each cell's top-right corner. */
  commentRefs?: string[];
  /** Full-fidelity comment bodies (cell ref + author + plain text) for every
   *  `<comment>` in `xl/commentsN.xml` (ECMA-376 §18.7). Parallel to
   *  {@link commentRefs} (one entry per ref). Consume this to read the note
   *  text; the renderer uses {@link commentRefs} for the red indicator and the
   *  viewer surfaces these bodies in an Excel-style hover popup. */
  comments?: XlsxComment[];
  /** Data-validation rules on this sheet (ECMA-376 §18.3.1.32–33). Exposed for
   *  tooling. The viewer draws a list-dropdown arrow on the active cell when the
   *  selection intersects a `list`-type rule's `sqref` (display only — opening
   *  the list / picking a value is out of scope for a viewer). */
  dataValidations?: DataValidation[];
  /** Defined names in scope for this sheet (ECMA-376 §18.2.5). Used by
   *  conditional-formatting `expression` rules that call named ranges
   *  (e.g. `task_start`, `today`). */
  definedNames?: DefinedName[];
  /** Excel Tables on this sheet (ECMA-376 §18.5). The renderer overlays a
   *  built-in style (bold header, banded rows) on the given ranges. */
  tables?: TableInfo[];
  /** Pivot / table slicers (Office 2010+ extension). Each anchor carries a
   *  caption and the saved item list (with selection flags) so the renderer
   *  can draw a static button bank without the live pivot engine. */
  slicers?: SlicerAnchor[];
  /** Metadata-only pivot facts (ECMA-376 §18.10). Consumers must treat saved
   * worksheet cell values and styles as authoritative. */
  pivotTables?: PivotTableMetadata[];
  /** Pivot parts skipped because their XML, identity, or location was invalid. */
  pivotDiagnostics?: PivotDiagnostic[];
  /** Sparkline groups (Office 2010+ extension `x14:sparklineGroup`).
   *  Cross-sheet `<xm:f>` data references are resolved to numeric values at
   *  parse time, and theme + tint colors are flattened to `#RRGGBB`. */
  sparklineGroups?: SparklineGroup[];
  /** Family name of the workbook's Normal-style font, resolved by the parser
   *  from `<cellStyleXfs>[0].fontId` → `<fonts>[fontId].name.val`. The
   *  renderer uses this together with `defaultFontSize` to compute the Max
   *  Digit Width for column-width pixel conversion (ECMA-376 §18.3.1.13).
   *  Workbook-wide value, denormalized onto every worksheet. */
  defaultFontFamily?: string;
  /** Point size of the workbook's Normal-style font (`<fonts>[N].sz.val`). */
  defaultFontSize?: number;
  /** Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28),
   *  denormalized onto every worksheet by the parser so the cell formatter can
   *  resolve serial dates (§18.17.4.1) without a workbook back-reference.
   *  `true` = 1904 date system. Omitted (⇒ false) for the default 1900 system. */
  date1904?: boolean;
  /** RB7 partial degradation: set when THIS sheet's part could not be
   *  read/parsed. The workbook still opens with the OTHER sheets intact; this one
   *  is an empty placeholder (`rows` empty) whose `parseError` names the offending
   *  part (e.g. `"xl/worksheets/sheet3.xml: <detail>"`). Absent (`undefined`) for
   *  every healthy sheet. The renderer paints a visible error overlay. */
  parseError?: string;
}

export interface PivotTableMetadata {
  name: string;
  /** Declarative identity fact only; cache resolution uses the pivot part relationship. */
  cacheId: number;
  location: PivotLocation;
  /** Signed field indexes; `-2` is the Values pseudo-field sentinel. */
  rowFields: number[];
  columnFields: number[];
  pageFields: PivotPageField[];
  dataFields: PivotDataField[];
  /** Absent when the cache definition could not be parsed; false includes the schema default. */
  refreshOnLoad?: boolean;
  /** Cache invalidity flag; absent when the cache definition did not parse. */
  cacheInvalid?: boolean;
  cacheDefinitionPart?: string;
  cacheSource?: PivotCacheSource;
  status: PivotMetadataStatus;
  extensionUris?: string[];
}

export interface PivotLocation extends CellRange {
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
}

export interface PivotPageField {
  field: number;
  item?: number;
  name?: string;
}

export interface PivotDataField {
  field: number;
  subtotal?: string;
  rawSubtotal?: string;
  name?: string;
}

export type PivotCacheSource =
  | { kind: 'worksheet'; sheet?: string; reference?: string; name?: string; relationshipId?: string }
  | { kind: 'external' }
  | { kind: 'consolidation' }
  | { kind: 'scenario' };

export type PivotMetadataStatus =
  | { state: 'complete' }
  | { state: 'partial'; reasons: PivotPartialReason[] };

export type PivotPartialReason =
  | { kind: 'missingCacheRelationship' }
  | { kind: 'malformedCacheRelationships' }
  | { kind: 'unreadableCacheRelationships' }
  | { kind: 'externalCacheRelationship' }
  | { kind: 'ambiguousCacheRelationship' }
  | { kind: 'unreadableCacheDefinition' }
  | { kind: 'malformedCacheDefinition' }
  | { kind: 'malformedField'; field: string }
  | { kind: 'unsupportedCacheSource'; sourceType: string }
  | { kind: 'unresolvedWorksheetSourceRelationship' }
  | { kind: 'unsupportedSemanticFeature'; feature: string };

export interface PivotDiagnostic {
  part: string;
  reason:
    | { kind: 'unreadableWorksheetRelationships' }
    | { kind: 'malformedWorksheetRelationships' }
    | { kind: 'malformedPivotRelationship' }
    | { kind: 'externalPivotRelationship' }
    | { kind: 'unreadablePart' }
    | { kind: 'malformedXml' }
    | { kind: 'missingIdentity' }
    | { kind: 'invalidLocation' };
}

export interface SparklineGroup {
  /** `line` (default) | `column` | `stem` (win-loss). */
  kind: 'line' | 'column' | 'stem';
  markers: boolean;
  high: boolean;
  low: boolean;
  first: boolean;
  last: boolean;
  negative: boolean;
  /** Show the horizontal axis line through 0 when data crosses it. */
  displayXAxis: boolean;
  /** `gap` (default) | `zero` | `span`. */
  displayEmptyCellsAs: string;
  /** `individual` (default) | `group` | `custom`. */
  minAxisType: string;
  maxAxisType: string;
  manualMin?: number;
  manualMax?: number;
  /** Stroke weight in pt for `line`. ECMA-376 default 0.75. */
  lineWeight: number;
  /** Resolved RGB hex strings (theme/tint already flattened by the parser). */
  colorSeries?: string;
  colorNegative?: string;
  colorAxis?: string;
  colorMarkers?: string;
  colorFirst?: string;
  colorLast?: string;
  colorHigh?: string;
  colorLow?: string;
  sparklines: Sparkline[];
}

export interface Sparkline {
  /** 1-based row of the destination cell (`<xm:sqref>`). */
  row: number;
  /** 1-based column of the destination cell. */
  col: number;
  /** Numeric values resolved from the `<xm:f>` range. `null` for empty
   *  / non-numeric cells; honored as gaps at render time. */
  values: (number | null)[];
}

export interface SlicerAnchor {
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  caption: string;
  items: SlicerItem[];
  /** Resolved custom slicer style. Absent means use the built-in fallback. */
  style?: SlicerStyle;
}

export interface SlicerItem {
  name: string;
  selected: boolean;
}

export interface SlicerStyle {
  whole?: SlicerElementStyle;
  header?: SlicerElementStyle;
  selectedItemWithData?: SlicerElementStyle;
  unselectedItemWithData?: SlicerElementStyle;
}

export interface SlicerElementStyle {
  fontColor?: string;
  fontSize?: number;
  fontBold?: boolean;
  fontFamily?: string;
  fillColor?: string;
  borderColor?: string;
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
  /** `true` when `styleName` is defined in the file's `<tableStyles>` block,
   *  i.e. a *custom* style (ECMA-376 §18.5.1.2). The renderer draws such tables
   *  strictly from their declared element dxfs and must NOT apply the accent
   *  approximation (banding / synthesized rules / header) reserved for built-in
   *  style names whose definitions are absent from the file. */
  isCustom?: boolean;
  /** Dxf index for the `wholeTable` element of a custom `<tableStyle>`
   *  (ECMA-376 §18.8.83). When set, its border/fill apply to every cell of
   *  the table as a base layer. Undefined for built-in style names. */
  wholeTableDxf?: number;
  /** Dxf index for the `headerRow` element of a custom `<tableStyle>` —
   *  provides header background, font color/weight, and vertical separators. */
  headerRowDxf?: number;
  /** Dxf index for the `totalRow` element (ECMA-376 §18.18.93). */
  totalRowDxf?: number;
  /** Dxf index for the `firstColumn` element. */
  firstColumnDxf?: number;
  /** Dxf index for the `lastColumn` element. */
  lastColumnDxf?: number;
  /** Dxf index for `firstRowStripe` (band1 horizontal) — odd banded-row stripe. */
  band1HorizontalDxf?: number;
  /** Dxf index for `secondRowStripe` (band2 horizontal) — even banded-row stripe. */
  band2HorizontalDxf?: number;
  /** Per-column DXF references (ECMA-376 §18.5.1.3 `tableColumn`). Index by
   *  `cellCol - range.left`. The renderer can use these to apply column-level
   *  overlays for named-style tables; for files where Excel pre-bakes the
   *  column DXF result into the cell `xf` (the common case), reading `xf` is
   *  sufficient and these fields are informational. */
  columns: TableColumnInfo[];
}

/** Per-column DXF references inside a `<table>` element
 *  (ECMA-376 §18.5.1.3 `tableColumn`). */
export interface TableColumnInfo {
  /** `<tableColumn dataDxfId>` — applied to data cells in this column. */
  dataDxfId?: number;
  /** `<tableColumn headerRowDxfId>` — applied to the header cell of this column. */
  headerRowDxfId?: number;
  /** `<tableColumn totalsRowDxfId>` — applied to the totals cell of this column. */
  totalsRowDxfId?: number;
}

export interface DefinedName {
  name: string;
  formula: string;
}

/** One cell comment. Sourced from the classic notes file `xl/commentsN.xml`
 *  (ECMA-376 §18.7) when present, otherwise from the Office-365 threaded
 *  comments part `xl/threadedComments/` (MS-XLSX schema
 *  `…/spreadsheetml/2018/threadedcomments`, `personId` resolved via
 *  `xl/persons/`). `text` is the joined plain text — every `<r><t>` run for
 *  classic notes, every reply in the thread (newline-joined) for threaded
 *  comments; rich-text formatting is dropped. 1:1 with the Rust `XlsxComment`
 *  (serde camelCase). */
export interface XlsxComment {
  /** A1-style cell reference (`@ref` on the comment element). */
  cellRef: string;
  /** Resolved author name — the `<authors>` entry (classic) or the `<person>`
   *  `displayName` (threaded). Absent when unresolved. */
  author?: string;
  /** Concatenated plain text of every run / threaded reply. */
  text: string;
}

/** One `<dataValidation>` rule (ECMA-376 §18.3.1.33). `type` is the constraint
 *  class (`list` | `whole` | `decimal` | `date` | `time` | `textLength` |
 *  `custom`); `operator` qualifies it (`between` | `notBetween` | `equal` | …).
 *  `formula1` / `formula2` are the operands (for `list`, `formula1` is the
 *  comma-separated literal list or a range/named reference). 1:1 with the Rust
 *  `DataValidation` (serde camelCase). */
export interface DataValidation {
  /** Affected cell ranges, verbatim from `@sqref` (space-separated A1 refs). */
  sqref: string;
  /** Constraint class. Absent means the spec default (`none`, no constraint). */
  validationType?: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  /** `@allowBlank` — empty input is permitted. */
  allowBlank?: boolean;
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  errorMessage?: string;
}

// ─── Chart types ─────────────────────────────────────────────────────────────
//
// The chart payload on `ChartAnchor` is the canonical {@link ChartModel} emitted
// by the Rust parser (`ooxml_common::chart::ChartModel`) — a single source of
// truth shared with the pptx parser and the core renderer. The former
// XLSX-local chart interfaces (`ChartData`, `XlsxChartSeries`, `ManualLayout`,
// etc.) that duplicated the core sub-types are gone; the parser now applies the
// old `adaptChartData` defaults in Rust before emit. The core sub-types are
// re-exported below (with back-compat aliases) so downstream code keeps a stable
// import surface.
export type {
  ChartModel,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartErrBars,
  ChartManualLayout,
  LegendManualLayout,
} from '@silurus/ooxml-core';
import type {
  ChartSeries as CoreChartSeries,
  ChartSeriesDataLabels as CoreSeriesDataLabels,
  ChartDataLabelOverride as CoreDataLabelOverride,
  ChartDataPointOverride as CoreDataPointOverride,
  ChartErrBars as CoreErrBars,
  ChartManualLayout as CoreManualLayout,
} from '@silurus/ooxml-core';
/**
 * @deprecated Chart series are now the core {@link ChartModel}'s `ChartSeries`.
 * Kept as an alias for backward-compatible imports.
 */
export type XlsxChartSeries = CoreChartSeries;
/** @deprecated Use `ChartSeriesDataLabels` from @silurus/ooxml-core. */
export type SeriesDataLabels = CoreSeriesDataLabels;
/** @deprecated Use `ChartDataLabelOverride` from @silurus/ooxml-core. */
export type DataLabelOverride = CoreDataLabelOverride;
/** @deprecated Use `ChartDataPointOverride` from @silurus/ooxml-core. */
export type DataPointOverride = CoreDataPointOverride;
/** @deprecated Use `ChartErrBars` from @silurus/ooxml-core. */
export type ErrBars = CoreErrBars;
/** @deprecated Use `ChartManualLayout` from @silurus/ooxml-core. */
export type ManualLayout = CoreManualLayout;

export interface ChartAnchor {
  fromCol: number; fromColOff: number;
  fromRow: number; fromRowOff: number;
  toCol: number;   toColOff: number;
  toRow: number;   toRowOff: number;
  /** The chart payload, already in the canonical {@link ChartModel} shape the
   *  Rust parser emits. The parser adapts its internal parse structure into
   *  `ChartModel` (formerly the TS `adaptChartData`); this is passed straight
   *  to `renderChart`. */
  chart: ChartModel;
}

export interface ShapeAnchor {
  fromCol: number; fromColOff: number;
  fromRow: number; fromRowOff: number;
  toCol: number;   toColOff: number;
  toRow: number;   toRowOff: number;
  /** `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). With `"oneCell"` the
   *  renderer uses `nativeExtCx`/`nativeExtCy` as the on-sheet size, since
   *  Excel preserves the group's saved EMU extent regardless of cell
   *  resizing ("Move but don't size with cells"). Absent ⇒ default `"twoCell"`. */
  editAs?: string;
  /** Saved EMU extent of the top-level grpSp (or the stand-alone sp/pic).
   *  Authoritative when `editAs === "oneCell"`. 0 = unavailable. */
  nativeExtCx: number;
  nativeExtCy: number;
  shapes: ShapeInfo[];
}

export interface ShapeInfo {
  /** Normalized [0,1] position/size relative to the anchor rect. */
  x: number; y: number; w: number; h: number;
  /** Rotation in degrees, clockwise. */
  rot: number;
  /** Effective DrawingML reflection after composing parent groups. */
  flipH?: boolean;
  flipV?: boolean;
  fillColor?: string;
  strokeColor?: string;
  /** Stroke width in EMU. 0 = no stroke. */
  strokeWidth: number;
  geom: ShapeGeom;
  /** Optional text body (`<xdr:txBody>`, ECMA-376 §20.5.2.34). Present for
   *  text boxes (`txBox="1"`) and any other shape that carries visible text. */
  text?: ShapeText;
}

export interface ShapeText {
  /** `<a:bodyPr@anchor>` — vertical alignment of the text block within the
   *  shape rect. `t` (top, default), `ctr` (middle), `b` (bottom). */
  anchor: string;
  /** `<a:bodyPr@wrap>` — `square` (wrap to width) | `none`. */
  wrap: string;
  /** `<a:bodyPr>` autofit child — `'sp'` (`spAutoFit`), `'norm'` (`normAutofit`),
   *  or `'none'` (`noAutofit`/absent). ECMA-376 §21.1.2.1.1-.3. Always present
   *  (default `'none'`), mirroring the core `TextBody.autoFit`. */
  autoFit?: string;
  /** `<a:normAutofit@fontScale>` — stored font-shrink fraction (e.g. 0.625 for
   *  `fontScale="62500"`). Null/absent when unset. Modeled for parity with
   *  pptx; the xlsx renderer does not currently apply it. */
  fontScale?: number | null;
  /** `<a:normAutofit@lnSpcReduction>` — stored line-spacing reduction fraction
   *  (e.g. 0.20 for `lnSpcReduction="20000"`). Null/absent when unset. */
  lnSpcReduction?: number | null;
  /** `<a:bodyPr@lIns>` — left text inset in EMU (ECMA-376 §21.1.2.1.1). Always
   *  present; the spec default 91440 EMU (7.2 pt) when the attribute is absent.
   *  Same EMU convention as `ShapeParagraph.marL`. */
  lIns: number;
  /** `<a:bodyPr@tIns>` — top text inset in EMU. Default 45720 (3.6 pt). */
  tIns: number;
  /** `<a:bodyPr@rIns>` — right text inset in EMU. Default 91440 (7.2 pt). */
  rIns: number;
  /** `<a:bodyPr@bIns>` — bottom text inset in EMU. Default 45720 (3.6 pt). */
  bIns: number;
  paragraphs: ShapeParagraph[];
}

export interface ShapeParagraph {
  /** `<a:pPr@algn>` — `l` (default) | `ctr` | `r` | `just` | `dist`. */
  align: string;
  /** `<a:pPr@rtl>` — whether the paragraph reads right-to-left
   *  (ECMA-376 §21.1.2.2.7). Omitted (undefined) when false. */
  rtl?: boolean;
  /** `<a:pPr@marL>` — left margin in EMU (ECMA-376 §21.1.2.2.7,
   *  `CT_TextParagraphProperties`). Direct attribute only — xlsx text boxes
   *  have no lstStyle/level cascade. Omitted (undefined) when unset. */
  marL?: number;
  /** `<a:pPr@marR>` — right margin in EMU (ECMA-376 §21.1.2.2.7).
   *  Omitted (undefined) when unset. */
  marR?: number;
  /** `<a:pPr@indent>` — first-line indent in EMU (negative = hanging),
   *  ECMA-376 §21.1.2.2.7. Omitted (undefined) when unset. */
  indent?: number;
  /** `<a:pPr>/<a:lnSpc>` line spacing (ECMA-376 §21.1.2.2.5). Direct-only;
   *  omitted when unset. */
  spaceLine?: SpaceLine | null;
  runs: ShapeTextRun[];
}

/** A run within a shape paragraph — tagged union mirroring the Rust enum
 *  (matches the pptx `TextRun` shape): styled text, a soft line break, or an
 *  OMML equation. Excel stores "Insert > Equation" as OMML inside the shared
 *  DrawingML `<xdr:txBody>` grammar (ECMA-376 §22.1), like PowerPoint. */
export type ShapeTextRun =
  | {
      type: 'text';
      text: string;
      bold: boolean;
      italic: boolean;
      /** Font size in points (already converted from `<a:rPr@sz>` 100ths-of-a-pt).
       *  0 = inherit (renderer falls back to its default). */
      size: number;
      color?: string;
      fontFace?: string;
      /** East-Asian typeface (`<a:ea@typeface>`, ECMA-376 §21.1.2.3.1). The
       *  common Japanese encoding sets Meiryo here while leaving `<a:latin>`
       *  default; the renderer floors the line box by this face's design line
       *  too (see `drawShapeText`). Undefined when the run declares no `<a:ea>`. */
      fontFaceEa?: string;
      /** Complex-script typeface (`<a:cs@typeface>`, ECMA-376 §21.1.2.3.1).
       *  Parsed/modeled but NOT used in the line-box floor: the cs face renders
       *  only complex-script glyphs (Arabic/Hebrew/Thai), so flooring the whole
       *  line box by it would over-grow Latin/CJK runs (deferred to per-glyph
       *  handling — see `drawShapeText`). Undefined when the run declares no
       *  `<a:cs>`. */
      fontFaceCs?: string;
    }
  | { type: 'break' }
  | {
      type: 'math';
      /** OMML AST (shared `MathNode` model) for the equation. */
      nodes: MathNode[];
      /** true = block (`m:oMathPara`), false = inline (`m:oMath`). */
      display: boolean;
      /** Point size when the run carries an explicit `rPr@sz`; else inherit. */
      fontSize?: number;
      color?: string;
    };

export type ShapeGeom =
  | {
      type: 'preset';
      name: string;
      /** Adjust handles from `<a:avLst><a:gd>` in `adj1..adj8` order
       *  (ECMA-376 §19.5.31.3 / §20.1.9.5). `null` entries mean "use the
       *  preset's declared default". Omitted entirely when the shape has no
       *  `<a:avLst>`. Consumed by the shared `renderPresetShape` engine. */
      adj?: (number | null)[];
    }
  | { type: 'custom'; paths: PathInfo[] }
  /** Bitmap (or vector) picture leaf inside a `<xdr:grpSp>`. `imagePath` is the
   *  zip path of the drawing's relationship target — the blip's raster `r:embed`
   *  fallback, or the SVG itself when no raster is embedded — and `mimeType` its
   *  MIME. Bytes are fetched lazily by path; nothing is inlined as base64. */
  | {
      type: 'image';
      imagePath: string;
      /** MIME type of the blip at {@link imagePath} (e.g. `image/png`, or
       *  `image/svg+xml` for the SVG-only fallback). */
      mimeType: string;
      /** Vector original from the Microsoft `asvg:svgBlip` extension
       *  (MS-ODRAWXML), as a zip path. Prefer this over `imagePath` (the raster
       *  fallback, or the SVG itself when no raster blip is embedded). Absent
       *  when the picture carries no svgBlip extension. Its MIME is always
       *  `image/svg+xml` and is owned by the SVG decoder. */
      svgImagePath?: string;
      /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop on the leaf pic
       *  (signed fractions from each edge; visible region `[l, t, 1-r,
       *  1-b]`). Absent ⇒ the whole blip fills the leaf rect. Honored identically
       *  to the top-level {@link ImageAnchor.srcRect} (raster only). */
      srcRect?: { l: number; t: number; r: number; b: number };
      /** ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` opacity fraction (0..1) on the
       *  leaf pic. Absent ⇒ opaque. Applied via `globalAlpha`. */
      alpha?: number;
      /** ECMA-376 §20.1.8.23 `<a:duotone>` recolour on the leaf pic. Absent ⇒
       *  no effect. */
      duotone?: Duotone;
    };

/** ECMA-376 §20.1.8.23 `<a:duotone>` image effect, resolved to its two endpoint
 *  colours (mirrors the shared Rust `ooxml_common::blip::Duotone`). `clr1` is the
 *  dark endpoint (luminance 0), `clr2` the light endpoint (luminance 1); both are
 *  6-char uppercase hex WITHOUT a leading `#`, with per-colour transforms already
 *  applied by the parser. */
export interface Duotone {
  clr1: string;
  clr2: string;
}

export interface PathInfo {
  w: number;
  h: number;
  commands: PathCmd[];
}

export type PathCmd =
  | { op: 'moveTo'; x: number; y: number }
  | { op: 'lineTo'; x: number; y: number }
  | { op: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x3: number; y3: number }
  | { op: 'quadBezTo'; x1: number; y1: number; x2: number; y2: number }
  /** ECMA-376 §20.1.9.3. stAng/swAng in 60000ths of a degree. wr/hr in
   *  the path's own coordinate units. Pen position is the arc start. */
  | { op: 'arcTo'; wr: number; hr: number; stAng: number; swAng: number }
  | { op: 'close' };

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
  /** `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). `"oneCell"` instructs the
   *  renderer to use `nativeExtCx`/`nativeExtCy` as the size and ignore the
   *  `to` anchor (Excel's "Move but don't size with cells"). Absent ⇒ default
   *  `"twoCell"`. */
  editAs?: string;
  /** `<xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>` in EMU — the picture's saved
   *  size. Authoritative when `editAs === "oneCell"`. 0 = unavailable. */
  nativeExtCx: number;
  nativeExtCy: number;
  /** Zip path of the blip inside the package (e.g. `xl/media/image1.png`). The
   *  blip's own `r:embed` raster fallback when an svgBlip extension is present;
   *  otherwise the only source. Falls back to the SVG part itself when the
   *  picture has no raster `r:embed`. Bytes are fetched lazily by path. */
  imagePath: string;
  /** MIME type of the blip at {@link ImageAnchor.imagePath} (e.g. `image/png`,
   *  or `image/svg+xml` for the SVG-only fallback). */
  mimeType: string;
  /** Vector original from the Microsoft `asvg:svgBlip` extension (MS-ODRAWXML),
   *  as a zip path. Preferred over `imagePath` (the raster fallback, or the SVG
   *  itself when no raster blip is embedded). Absent when the picture carries no
   *  svgBlip extension. Its MIME is always `image/svg+xml` and is owned by the
   *  SVG decoder. */
  svgImagePath?: string;
  /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop. Each edge inset is a
   *  signed fraction of the source bitmap, measured from its edge, so the visible
   *  source region is `[l, t, 1-r, 1-b]`. Absent (the common case) ⇒ the whole
   *  blip fills the anchor rect; when present, the renderer draws only the
   *  cropped sub-rectangle (raster only — a metafile is rasterized to the
   *  display box, so its crop can't be honored faithfully and is skipped). */
  srcRect?: { l: number; t: number; r: number; b: number };
  /** ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` — the blip's overall opacity as a
   *  fraction (0..1). Absent ⇒ opaque. The renderer sets `ctx.globalAlpha` so the
   *  picture composites over the cells beneath it (e.g. a pink translucent photo
   *  over a matching cell fill). */
  alpha?: number;
  /** ECMA-376 §20.1.8.23 `<a:duotone>` recolour effect. Absent (the common case)
   *  ⇒ no effect. When present, the renderer remaps the image along the
   *  `clr1`→`clr2` luminance ramp before drawing. */
  duotone?: Duotone;
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
  /** External target (ECMA-376 §18.3.1.47 `r:id`, resolved via worksheet rels).
   *  `null` for a purely internal hyperlink. */
  url: string | null;
  /** Internal target (§18.3.1.47 `location`): a defined name or a cell reference
   *  such as `Sheet1!A1`. Present when the hyperlink navigates within the
   *  workbook rather than to an external URL. */
  location?: string | null;
  /** Optional display text (§18.3.1.47 `display`). Not used for rendering. */
  display?: string | null;
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
  | { type: 'aboveAverage'; aboveAverage: boolean; equalAverage?: boolean; stdDev?: number; dxfId: number | null; priority: number }
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
  /** Outline (grouping) depth 0-7 (ECMA-376 §18.3.1.73 `<row outlineLevel>`).
   *  Omitted on the wire when `0` (ungrouped); read as `outlineLevel ?? 0`. */
  outlineLevel?: number;
  /** `<row collapsed>` (§18.3.1.73): `true` on a summary row whose
   *  one-level-deeper detail rows are collapsed. Omitted when false. */
  collapsed?: boolean;
  /** `<row hidden>` (§18.3.1.73): `true` when the row is hidden — most often
   *  because a collapsed outline hides its detail rows. Distinct from
   *  `height === 0`. Omitted on the wire when false. */
  hidden?: boolean;
}

/** `<sheetPr><outlinePr>` flags (ECMA-376 §18.3.1.61). Both default to `true`. */
export interface OutlinePr {
  /** `true` (default) ⇒ a group's summary row sits *below* its detail rows;
   *  `false` ⇒ above. */
  summaryBelow: boolean;
  /** `true` (default) ⇒ a group's summary column sits to the *right* of its
   *  detail columns; `false` ⇒ to the left. */
  summaryRight: boolean;
}

export interface Cell {
  col: number;
  row: number;
  value: CellValue;
  /** Style index into the styles table. Omitted on the wire when `0` (the
   *  common unstyled case), so read it as `styleIndex ?? 0`. */
  styleIndex?: number;
  /** Raw `<f>` formula text (ECMA-376 §18.3.1.40), when present. The renderer
   *  uses this to recompute volatile functions (TODAY, NOW) at display time
   *  so the cached `<v>` — frozen when the file was last saved — doesn't
   *  show a stale date. */
  formula?: string;
  /** Whether this cell displays its phonetic hint (furigana). The parser
   *  resolves it as `cell/@ph ?? row/@ph ?? false` — the per-cell `<c ph>`
   *  (ECMA-376 §18.3.1.4) wins when present (an explicit `ph="0"` overrides an
   *  enabled row), otherwise the row-level `<row ph>` (§18.3.1.73) is inherited,
   *  otherwise the schema default (false). Omitted on the wire when false, so
   *  read as `showPhonetic ?? false`. A cell whose String Item carries `<rPh>`
   *  runs still shows NO furigana unless the resolved value is true. */
  showPhonetic?: boolean;
}

export type CellValue =
  | { type: 'empty' }
  | {
      type: 'text';
      text: string;
      runs?: Run[];
      /** ECMA-376 §18.4.6 phonetic runs (furigana) carried over from the
       *  resolved String Item. Present for inline strings, and populated by
       *  {@link resolveSharedStrings} for shared-string cells. Absent when the
       *  string has no furigana. */
      phoneticRuns?: PhoneticRun[];
      /** ECMA-376 §18.4.3 phonetic display properties (font index / char set /
       *  alignment) for the furigana above. Absent when the `<si>` had no
       *  `<phoneticPr>`. */
      phoneticPr?: PhoneticProperties;
    }
  | { type: 'number'; number: number }
  | { type: 'bool'; bool: boolean }
  | { type: 'error'; error: string }
  /** Shared-string reference into `ParsedWorkbook.sharedStrings` (ECMA-376
   *  §18.4.8). Resolved to `{ type: 'text', ... }` by the workbook before the
   *  renderer (or any other consumer) sees it, so downstream code never
   *  encounters this variant. */
  | { type: 'shared'; si: number };

/** ECMA-376 §18.4.6 `<rPh sb=".." eb="..">` — one furigana run. `sb`/`eb` are
 *  zero-based character offsets into the base text; the hint `text` is shown
 *  over base characters `[sb, eb)`. */
export interface PhoneticRun {
  /** Zero-based start character offset into the base text (inclusive). */
  sb: number;
  /** Zero-based end character offset into the base text (exclusive). */
  eb: number;
  /** The phonetic hint text (e.g. the katakana reading). */
  text: string;
}

/** ECMA-376 §18.18.57 ST_PhoneticType — the East-Asian character set the
 *  furigana is displayed in. Absent on {@link PhoneticProperties} defaults to
 *  `'fullwidthKatakana'` per the CT_PhoneticPr schema. */
export type PhoneticType =
  | 'fullwidthKatakana'
  | 'halfwidthKatakana'
  | 'Hiragana'
  | 'noConversion';

/** ECMA-376 §18.18.56 ST_PhoneticAlignment — how the furigana is aligned over
 *  the base text. Absent on {@link PhoneticProperties} defaults to `'left'`. */
export type PhoneticAlignment = 'left' | 'center' | 'distributed' | 'noControl';

/** ECMA-376 §18.4.3 `<phoneticPr>` — phonetic display properties. */
export interface PhoneticProperties {
  /** Zero-based index into `Styles.fonts` (§18.18.32 ST_FontId). Out of bounds
   *  falls back to font 0 (§18.4.3). Drives the furigana font size / family. */
  fontId: number;
  /** §18.18.57 — absent means `'fullwidthKatakana'` (schema default). */
  type?: PhoneticType;
  /** §18.18.56 — absent means `'left'` (schema default). */
  alignment?: PhoneticAlignment;
}

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
  /**
   * Underline style when not the default single line. ECMA-376 §18.4.13
   * (`ST_UnderlineValues`): "double" | "singleAccounting" | "doubleAccounting".
   * Absent means single (when `underline` is true) or no underline.
   */
  underlineStyle?: string;
  /**
   * ECMA-376 §18.4.6 (`ST_VerticalAlignRun`): "superscript" | "subscript".
   * Absent leaves the run on the baseline.
   */
  vertAlign?: 'superscript' | 'subscript';
}

export interface SharedString {
  text: string;
  runs?: Run[];
  /** ECMA-376 §18.4.6 phonetic runs (furigana). Absent when the `<si>` has no
   *  `<rPh>`. */
  phoneticRuns?: PhoneticRun[];
  /** ECMA-376 §18.4.3 phonetic display properties. Absent when the `<si>` has
   *  no `<phoneticPr>`. */
  phoneticPr?: PhoneticProperties;
}

export interface NumFmt {
  numFmtId: number;
  formatCode: string;
}

export interface Styles {
  fonts: CellFont[];
  fills: CellFill[];
  borders: Border[];
  cellXfs: CellXf[];
  numFmts: NumFmt[];
  dxfs: Dxf[];
}

export interface Dxf {
  font: CellFont | null;
  fill: CellFill | null;
  border: Border | null;
  /** Number format override from the dxf (ECMA-376 §18.8.17). When a
   *  conditional-formatting rule matches, this numFmt replaces the cell's own
   *  style numFmt for rendering — e.g. switching a calendar cell from `d` to
   *  `m"月"d"日"` on the first day of each month. */
  numFmt?: NumFmt | null;
}

export interface CellFont {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  size: number;
  color: string | null;
  name: string | null;
  /** ECMA-376 §18.4.13 ST_UnderlineValues — see RunFont.underlineStyle. */
  underlineStyle?: string;
  /** ECMA-376 §18.4.6 ST_VerticalAlignRun on a cell-level <font>. */
  vertAlign?: 'superscript' | 'subscript';
}

export interface CellFill {
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
  /** Inner horizontal rule between rows inside a region
   *  (ECMA-376 §18.8.40 `tableStyleElement/border/horizontal`).
   *  Only set on table-style dxfs; absent on cell-level borders. */
  horizontal?: BorderEdge | null;
  /** Inner vertical rule between columns inside a region. */
  vertical?: BorderEdge | null;
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
  /** `<alignment readingOrder>` (ECMA-376 §18.8.1) — 0 = context (default),
   *  1 = LTR, 2 = RTL. Drives canvas `direction`. */
  readingOrder?: number;
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

/** Emitted once per cell that has text, with the cell's canvas-pixel bounds. */
export interface XlsxTextRunInfo {
  /** Worksheet name used by Office CLI's `/Sheet/A1` addressing. */
  sheetName: string;
  /** Stable A1 cell reference within {@link sheetName}. */
  cellRef: string;
  text: string;
  /** Canvas CSS-pixel x of the cell's top-left corner. */
  x: number;
  /** Canvas CSS-pixel y of the cell's top-left corner. */
  y: number;
  /** Cell width in canvas CSS pixels. */
  width: number;
  /** Cell height in canvas CSS pixels. */
  height: number;
  row: number;
  col: number;
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
  /** Pre-decoded image sources keyed by their zip `imagePath` (for ImageAnchor
   *  and group-leaf image rendering). */
  loadedImages?: Map<string, CanvasImageSource | null>;
  /** Fetch an embedded image's bytes by zip path, wrapped in a Blob of the given
   *  MIME (twin of pptx/docx `fetchImage`). The orchestrator decodes these into
   *  {@link loadedImages} before the synchronous draw. Supplied by
   *  {@link XlsxWorkbook} (routing through the worker) or the render worker
   *  (reading its retained buffer). Absent ⇒ no images are decoded. */
  fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  /** Called once per cell that contains text, with canvas-pixel position and cell address. */
  onTextRun?: (info: XlsxTextRunInfo) => void;
  /** Highlighted row range for selected row headers (1-indexed inclusive).
   *  `strong: true` → light blue + blue border (rows / cols / all selection modes).
   *  `strong: false` → slightly darker grey (cells selection mode). */
  selectedRowRange?: { start: number; end: number; strong: boolean } | null;
  /** Same shape as selectedRowRange, for column headers. */
  selectedColRange?: { start: number; end: number; strong: boolean } | null;
}

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number }
  /** Parse one sheet lazily. Deliberately carries NO `data`: the worker already
   *  retained the whole-workbook buffer on the preceding `parse`, so re-sending
   *  it here would structured-clone the entire file per sheet switch for no
   *  gain. `parseSheet` is therefore only valid AFTER a `parse` (a `parseSheet`
   *  with no retained buffer is a protocol violation). */
  | {
      type: 'parseSheet';
      id: number;
      sheetIndex: number;
      sheetName: string;
      maxZipEntryBytes?: number;
      maxZipTotalBytes?: number;
      maxZipEntries?: number;
    }
  /** Pull one embedded image's raw bytes by zip path from the buffer the worker
   *  retained at parse time. Twin of pptx/docx `extractImage`; xlsx uses the
   *  `type` discriminant. */
  | { type: 'extractImage'; id: number; path: string }
  /** Project the retained archive to GitHub-flavoured markdown
   *  (`XlsxArchive.to_markdown`, the handle already opened at `parse` — no
   *  re-copy of the file). Twin of `extractImage`: the archive stays in the
   *  worker, only the string crosses back. */
  | { type: 'toMarkdown'; id: number };

export type WorkerResponse =
  // The workbook index / worksheet cross the worker boundary as raw UTF-8 JSON
  // bytes (transferred, not cloned); the main thread does the single
  // `TextDecoder.decode` + `JSON.parse` into a `ParsedWorkbook` / `Worksheet`.
  // See `parse_xlsx` (Rust) for why.
  | { type: 'parsed'; id: number; workbookJson: ArrayBuffer }
  | { type: 'parsedSheet'; id: number; worksheetJson: ArrayBuffer }
  | { type: 'imageExtracted'; id: number; bytes: ArrayBuffer }
  | { type: 'markdownRendered'; id: number; markdown: string }
  | { type: 'error'; id: number; message: string };
