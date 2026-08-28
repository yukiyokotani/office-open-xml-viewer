import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  DocxTextRun,
  FieldRun,
  HeadersFooters,
  ImageRun,
  NumberingInfo,
  LineNumbering,
  ChartRun,
  ShapeRun,
  DocTable,
  TextPath,
  TblpPr,
} from './types.js';
import type {
  NumberingMarkerShapeInput,
  FloatingTablePositionInput,
  SourceRef,
  TableFormatInput,
  TableRowExceptionInput,
  TableRowHeightInput,
  VmlTextPathAcquisitionInput,
} from './layout/types.js';
import type { MathOccurrence } from './layout/resources.js';
import { anchorOccurrenceKey, chartResourceKey, mathResourceKey, sourceKey } from './layout/source-key.js';
import { mathFallbackText } from './layout/math-fallback-text.js';
import type {
  ComplexFieldBoundaryInput,
  TextFontSlotPresence,
  TextFontSlots,
  UnavailableDrawingAcquisitionRun,
} from './layout/text.js';
import type { ParagraphAcquisitionInput, ParagraphAcquisitionRun, ParagraphLayoutSource } from './layout/text.js';
import type { AnchorAcquisitionInput, InternalAnchorRunWire } from './layout/anchor-input.js';
import {
  paragraphTypographyAcquisitionInput,
  runTypographyAcquisitionInput,
  type InternalParagraphTypographyWire,
  type InternalRunTypographyWire,
} from './layout/typography-input.js';
import { deepFreezePlainData, snapshotPlainData } from './layout/plain-data.js';
import {
  normalizeTextBoxInput,
  type TextBoxAcquisitionInput,
} from './layout/textbox-input.js';
import {
  effectiveTableWidthKind,
  projectTableColumnLayoutInput,
  tableDxaPtFromLexical,
  tableWidthConstraintFromLexical,
  type CellIntrinsicWidths,
  type TableAcquisitionInput,
  type TableCellLayoutAcquisitionWire,
  type TableLayoutAcquisitionWire,
  type TableLayoutSource,
  type TableMarginAcquisitionWire,
  type TablePropertyExceptionAcquisitionWire,
  type TableRowHeightAcquisitionWire,
  type TableRowLayoutAcquisitionWire,
  type TableSourceAcquisitionInput,
  type TableSourceSemanticInput,
  type TableWidthAcquisitionWire,
} from './layout/table-source-acquisition.js';
import {
  sectionPageBox,
  type BodySectionIndexInput,
  type BodySectionOccurrence,
} from './layout/context.js';
import { normalizeAdjacentTables } from './layout/adjacent-tables.js';
import { isWrapFloat } from './float-layout.js';
import type {
  BodyLayoutAcquisitionInput,
  BodyLayoutInput,
  BodyLayoutSequenceEntryFor,
} from './layout/body-layout-input.js';
import { projectBodyLayoutInput } from './layout/body-layout-input.js';
import type { BodyAcquisitionInputProjections } from './layout/acquisition-input-projections.js';
import type { DocumentTypographySettingsInput } from './layout-context.js';
import { isInklessParagraph } from './layout/paragraph-visibility.js';
import {
  wordTableCellSpacingValuePt,
  wordTableMarginValuePt,
  wordTableRowHeightRule,
} from './layout/table-compatibility.js';
import {
  mapParseDiagnostics,
  type ParseDiagnosticWire,
} from './layout/diagnostics.js';

export type {
  CellIntrinsicWidths,
  TableAcquisitionInput,
  TableCellLayoutAcquisitionWire,
  TableLayoutAcquisitionWire,
  TableLayoutKindAcquisitionWire,
  TableMarginAcquisitionWire,
  TablePropertyExceptionAcquisitionWire,
  TableRowHeightAcquisitionWire,
  TableRowLayoutAcquisitionWire,
  TableSourceAcquisitionInput,
  TableSourceSemanticInput,
  TableWidthAcquisitionWire,
} from './layout/table-source-acquisition.js';

export interface InternalRunFontSlots {
  readonly direct: TextFontSlots;
  readonly theme: TextFontSlots;
  readonly themePresent: TextFontSlotPresence;
}

/** Parser-emitted metadata intentionally kept outside the stable public model.
 * Ordinary text and field results share these resolved WordprocessingML axes. */
export interface InternalRunSlotMetadata {
  fontFamilyHighAnsi?: string | null;
  fontSlots?: InternalRunFontSlots;
  fontFamilyEastAsia?: string | null;
  fontHint?: 'default' | 'eastAsia' | 'cs';
  rtl?: boolean;
  cs?: boolean;
  fontFamilyCs?: string | null;
  fontSizeCs?: number;
  boldCs?: boolean;
  italicCs?: boolean;
  langBidi?: string;
  langEastAsia?: string;
}

interface InternalNoBreakHyphenWire {
  readonly __noBreakBefore?: boolean;
  readonly __noBreakAfter?: boolean;
  /** UTF-16 offsets immediately after the injected U+002D glyph. */
  readonly __noBreakHyphenOffsets?: readonly number[];
}

/** Effective parser-owned run facts used by non-content glyphs such as list
 * markers and paragraph marks. Kept off the stable public document model. */
export interface InternalRunFontFacts extends InternalRunSlotMetadata {
  fontFamily?: string | null;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  kerning?: number;
}

export interface InternalNumberingInfo extends NumberingInfo {
  fontFacts?: InternalRunFontFacts;
}

export interface InternalDocParagraph extends DocParagraph {
  numbering: InternalNumberingInfo | null;
  paragraphMarkFontFacts?: InternalRunFontFacts;
  readonly __complexFieldBoundaries?: readonly InternalComplexFieldBoundaryWire[];
  __runRevisions?: readonly (DocRun['revision'] | null)[];
}

type UnavailableDrawingRunWire = Omit<
  UnavailableDrawingAcquisitionRun,
  'anchorAcquisitionInput'
> & InternalAnchorRunWire;

interface UnavailableDrawingSidecarEntry {
  /** Insertion point in the stable public `runs` array. Equal insertion points
   * retain parser order. */
  readonly publicRunIndex: number;
  readonly run: Readonly<UnavailableDrawingRunWire>;
}

/** Parser-only flow occurrences removed from the stable public document model.
 * The normalized paragraph identity is the ownership boundary: acquisition can
 * reconstruct the authored run order, while callers observing/serializing
 * `DocxDocumentModel` see only the declared `DocRun` union. */
const unavailableDrawingSidecars = new WeakMap<
  DocParagraph,
  readonly UnavailableDrawingSidecarEntry[]
>();

function unavailableDrawingEntries(
  paragraph: Readonly<DocParagraph>,
): readonly UnavailableDrawingSidecarEntry[] {
  return unavailableDrawingSidecars.get(paragraph as DocParagraph) ?? [];
}

function paragraphRunsWithUnavailableDrawings(
  paragraph: Readonly<DocParagraph>,
): readonly Readonly<DocRun | UnavailableDrawingRunWire>[] {
  const entries = unavailableDrawingEntries(paragraph);
  if (entries.length === 0) {
    return paragraph.runs as readonly Readonly<DocRun>[];
  }
  const runs: Readonly<DocRun | UnavailableDrawingRunWire>[] = [];
  let entryIndex = 0;
  for (let publicRunIndex = 0; publicRunIndex <= paragraph.runs.length; publicRunIndex += 1) {
    while (entries[entryIndex]?.publicRunIndex === publicRunIndex) {
      runs.push(entries[entryIndex]!.run);
      entryIndex += 1;
    }
    if (publicRunIndex < paragraph.runs.length) {
      runs.push(paragraph.runs[publicRunIndex]!);
    }
  }
  return runs;
}

function paragraphHasUnavailableDrawing(paragraph: Readonly<DocParagraph>): boolean {
  return unavailableDrawingEntries(paragraph).length > 0;
}

export interface InternalComplexFieldBoundaryWire {
  readonly occurrenceId: number;
  readonly boundary: 'start' | 'end';
  readonly runIndex: number;
  readonly fieldType: 'ref' | 'pageRef' | 'other';
  readonly instruction: string;
  readonly hyperlinkAnchor?: string;
}

type TextOnlyMetadata = Pick<
  DocxTextRun,
  | 'ruby' | 'revision' | 'hyperlink' | 'hyperlinkAnchor'
  | 'underlineStyle' | 'underlineColor' | 'colorAuto' | 'border'
  | 'snapToGrid' | 'charSpacing' | 'charScale' | 'fitTextVal' | 'fitTextId'
  | 'position' | 'kerning' | 'eastAsianVert' | 'eastAsianVertCompress'
>;

export type InternalTextRun = DocxTextRun & InternalRunSlotMetadata;
export type InternalFieldRun = FieldRun & Partial<TextOnlyMetadata> & InternalRunSlotMetadata;
export type InternalTextBearingRun = InternalTextRun | InternalFieldRun;
export type InternalMathRun = Extract<DocRun, { type: 'math' }> & {
  readonly source: SourceRef;
  readonly resourceKey: string;
};

export interface InternalDocxDocumentModel extends DocxDocumentModel {
  fontFamilyCharsets?: Record<string, string>;
  readonly diagnostics?: readonly ParseDiagnosticWire[];
  readonly __pageLayoutSettings?: Readonly<{
    mirrorMargins?: boolean;
    gutterAtTop?: boolean;
    bookFoldPrinting?: boolean;
    bookFoldRevPrinting?: boolean;
    printTwoOnOne?: boolean;
  }>;
  readonly __noteLayoutSettings?: Readonly<{
    footnotePosition?: string;
    endnotePosition?: string;
  }>;
  readonly __documentTypographySettings?: Readonly<{
    normalStyleFontSizePt?: number;
  }>;
}

export function documentTypographySettingsInput(
  doc: DocxDocumentModel,
): DocumentTypographySettingsInput {
  const authored = (doc as InternalDocxDocumentModel)
    .__documentTypographySettings?.normalStyleFontSizePt;
  return snapshotPlainData({
    normalStyleFontSizePt:
      typeof authored === 'number' && Number.isFinite(authored) && authored > 0
        ? authored
        : 10,
  }, 'DOCX document typography settings input');
}

export interface DocumentPageLayoutSettingsInput {
  readonly mirrorMargins: boolean;
  readonly gutterAtTop: boolean;
  readonly bookFoldPrinting: boolean;
  readonly bookFoldRevPrinting: boolean;
  readonly printTwoOnOne: boolean;
}

export function documentPageLayoutSettingsInput(
  doc: DocxDocumentModel,
): DocumentPageLayoutSettingsInput {
  const settings = (doc as InternalDocxDocumentModel).__pageLayoutSettings;
  return snapshotPlainData({
    mirrorMargins: settings?.mirrorMargins === true,
    gutterAtTop: settings?.gutterAtTop === true,
    bookFoldPrinting: settings?.bookFoldPrinting === true,
    bookFoldRevPrinting: settings?.bookFoldRevPrinting === true,
    printTwoOnOne: settings?.printTwoOnOne === true,
  }, 'DOCX page layout settings input');
}

export interface DocumentNoteLayoutSettingsInput {
  readonly footnotePosition: string;
  readonly endnotePosition: string;
}

export function documentNoteLayoutSettingsInput(
  doc: DocxDocumentModel,
): DocumentNoteLayoutSettingsInput {
  const settings = (doc as InternalDocxDocumentModel).__noteLayoutSettings;
  return snapshotPlainData({
    // §17.11.21/.22 defaults when document-wide w:pos is absent.
    footnotePosition: settings?.footnotePosition ?? 'pageBottom',
    endnotePosition: settings?.endnotePosition ?? 'docEnd',
  }, 'DOCX note layout settings input');
}

interface InternalSectionPlacementWire {
  readonly sectionId: string;
  readonly sectionBidi?: boolean;
  readonly vAlign?: string | null;
  readonly lineNumbering?: LineNumbering | null;
  readonly docGridType?: string | null;
  readonly docGridLinePitch?: number | null;
  readonly docGridCharSpace?: number | null;
  readonly gutterPt?: number | null;
  readonly rtlGutter?: boolean | null;
  readonly pageBordersAuthored?: boolean;
  readonly pageBorders?: import('./types.js').PageBorders | null;
  readonly pageGeometry?: Readonly<{
    pageWidth?: number | null;
    pageHeight?: number | null;
    marginTop?: number | null;
    marginRight?: number | null;
    marginBottom?: number | null;
    marginLeft?: number | null;
    headerDistance?: number | null;
    footerDistance?: number | null;
  }> | null;
}

type InternalSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement?: InternalSectionPlacementWire;
};

type InternalSectionProps = DocxDocumentModel['section'] & {
  readonly __sectionPlacement?: InternalSectionPlacementWire;
};

/** Immutable parser-private section placement input. It is intentionally kept
 * outside BodyElement/DocxDocumentModel's stable declaration surface. */
export interface SectionPlacementInput {
  readonly sectionId: string;
  readonly sectionBidi: boolean;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly docGridType: string | null;
  readonly docGridLinePitch: number | null;
  readonly docGridCharSpace: number | null;
  readonly gutterPt: number | null;
  readonly rtlGutter: boolean | null;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: Readonly<import('./types.js').PageBorders> | null;
  readonly pageGeometry: InternalSectionPlacementWire['pageGeometry'];
}

interface DocumentSectionPlacementInputs {
  readonly endingSections: ReadonlyMap<number, SectionPlacementInput>;
  readonly finalSection: SectionPlacementInput;
}

function normalizeSectionGeometryWire(
  geometry: InternalSectionPlacementWire['pageGeometry'],
): Readonly<Partial<import('./types.js').SectionGeom>> {
  if (!geometry) return Object.freeze({});
  return Object.freeze(Object.fromEntries(Object.entries(geometry).filter(
    (entry): entry is [string, number] => typeof entry[1] === 'number',
  )) as Partial<import('./types.js').SectionGeom>);
}

function sectionPlacementFacts(input: SectionPlacementInput): import('./layout/context.js').SectionPlacementFacts {
  return Object.freeze({
    ...input,
    pageGeometry: normalizeSectionGeometryWire(input.pageGeometry),
  });
}

const sectionPlacementInputsByDocument = new WeakMap<object, DocumentSectionPlacementInputs>();
// The paginator seam receives two independent identity inputs. Nest the cache so
// reusing one body with a different final SectionProps cannot inherit stale final
// placement facts. Each entry remains an acquisition-time snapshot: subsequent
// caller mutation of the same SectionProps object does not rewrite retained data.
const sectionPlacementInputsByBody = new WeakMap<
  object,
  WeakMap<object, DocumentSectionPlacementInputs>
>();

interface InternalTable extends DocTable {
  readonly __tableLayout?: TableLayoutAcquisitionWire;
}

type InternalTableRow = DocTable['rows'][number] & {
  readonly __tableRowLayout?: TableRowLayoutAcquisitionWire;
};

type InternalTableCell = DocTable['rows'][number]['cells'][number] & {
  readonly __tableCellLayout?: TableCellLayoutAcquisitionWire;
};

const tableAcquisitionInputs = new WeakMap<object, TableAcquisitionInput>();
const tableFormatInputs = new WeakMap<object, TableFormatInput>();
const tableSourceAcquisitionInputs = new WeakMap<object, TableSourceAcquisitionInput>();

/** Snapshot serde-only table facts once at the parser/model boundary. Layout
 * receives only clone-safe immutable data, while hand-built public `DocTable`
 * values remain supported through aligned null entries and their public fields. */
export function tableAcquisitionInput(table: TableLayoutSource): TableAcquisitionInput {
  const cached = tableAcquisitionInputs.get(table);
  if (cached) return cached;
  const internal = table as Readonly<InternalTable>;
  const input = snapshotPlainData({
    table: internal.__tableLayout ?? null,
    rows: table.rows.map((row) => {
      const internalRow = row as Readonly<InternalTableRow>;
      return {
        row: internalRow.__tableRowLayout ?? null,
        cells: row.cells.map(
          (cell) => (cell as Readonly<InternalTableCell>).__tableCellLayout ?? null,
        ),
      };
    }),
  }, 'DOCX table acquisition input') as TableAcquisitionInput;
  tableAcquisitionInputs.set(table, input);
  return input;
}

const finiteOrNull = (value: number | null | undefined): number | null => (
  value != null && Number.isFinite(value) ? value : null
);

/** Detach only public semantics consumed by the column algorithm. Content and
 * nested tables remain owned by the document repository, preventing a table
 * fact per nesting level from retaining the same subtree repeatedly. */
function tableColumnSemanticInput(
  table: TableLayoutSource,
): TableSourceSemanticInput {
  return snapshotPlainData({
    // The stable TypeScript shape requires colWidths, but historical rich
    // textbox fixtures and JS callers may omit it. The pre-extraction column
    // path treated an absent parser grid as an empty compatibility grid.
    colWidths: (table.colWidths ?? []).map((width) => (
      Number.isFinite(width) && width >= 0 ? width : 0
    )),
    layout: table.layout ?? null,
    widthPt: finiteOrNull(table.widthPt),
    widthPct: finiteOrNull(table.widthPct),
    rows: table.rows.map((row) => ({
      gridBefore: finiteOrNull(row.gridBefore) ?? 0,
      gridAfter: finiteOrNull(row.gridAfter) ?? 0,
      cells: row.cells.map((cell) => ({
        colSpan: finiteOrNull(cell.colSpan) ?? 1,
        widthPt: finiteOrNull(cell.widthPt),
        widthPct: finiteOrNull(cell.widthPct),
      })),
    })),
  }, 'DOCX table column semantic input') as TableSourceSemanticInput;
}

/** Complete immutable table source fact. Layout can retain and replay this
 * value without a parser-model object, WeakMap, or acquisition callback. */
export function tableSourceAcquisitionInput(
  table: TableLayoutSource,
): TableSourceAcquisitionInput {
  const cached = tableSourceAcquisitionInputs.get(table);
  if (cached) return cached;
  const input = deepFreezePlainData({
    semantic: tableColumnSemanticInput(table),
    lexical: tableAcquisitionInput(table),
    format: tableFormatInput(table),
  }) as TableSourceAcquisitionInput;
  tableSourceAcquisitionInputs.set(table, input);
  return input;
}

/** Effective flow classification acquired before tblpPr defaults erase the
 * distinction governed by `word-effective-floating-table-positioning`. */
export function tableParticipatesInOrdinaryFlow(table: TableLayoutSource): boolean {
  return tableSourceAcquisitionInput(table).format.ordinaryFlow;
}

/** Positioning payload only when `word-effective-floating-table-positioning`
 * treats the authored tblpPr as effective. */
export function effectiveTablePositioning(table: TableLayoutSource): TblpPr | null {
  return tableSourceAcquisitionInput(table).format.positioning === null
    ? null
    : (table.tblpPr ?? null);
}

function floatingTablePositionInput(positioning: TblpPr): FloatingTablePositionInput {
  return {
    leftFromTextPt: positioning.leftFromText,
    rightFromTextPt: positioning.rightFromText,
    topFromTextPt: positioning.topFromText,
    bottomFromTextPt: positioning.bottomFromText,
    horzAnchor: positioning.horzAnchor,
    horzSpecified: positioning.horzSpecified,
    vertAnchor: positioning.vertAnchor,
    xPt: positioning.tblpX,
    yPt: positioning.tblpY,
    ...(positioning.tblpXSpec == null ? {} : { xAlign: positioning.tblpXSpec }),
    ...(positioning.tblpYSpec == null ? {} : { yAlign: positioning.tblpYSpec }),
  };
}

function finiteTableLexicalNumber(value: string | null, allowPercent: boolean): number | null {
  if (value === null) return null;
  const lexical = value.trim();
  const numeric = allowPercent && lexical.endsWith('%') ? lexical.slice(0, -1) : lexical;
  if (numeric.length === 0) return null;
  const parsed = Number(numeric);
  return Number.isFinite(parsed) ? parsed : null;
}

function tableTwipsValuePt(value: string | null | undefined): number | null {
  const parsed = finiteTableLexicalNumber(value ?? null, false);
  return parsed === null ? null : parsed / 20;
}

function normalizedTableHeightRule(rule: string): TableRowHeightInput['rule'] {
  if (rule === 'exact' || rule === 'atLeast') return rule;
  return 'auto';
}

function privateTableRowHeight(height: TableRowHeightAcquisitionWire): TableRowHeightInput {
  // `word-omitted-row-height-rule-at-least`: authored presence is semantic
  // input, not a parser implementation detail.
  return {
    rule: wordTableRowHeightRule(
      normalizedTableHeightRule(height.rule),
      height.ruleAuthored,
    ),
    valuePt: tableTwipsValuePt(height.value),
  };
}

function publicTableRowHeight(row: TableLayoutSource['rows'][number]): TableRowHeightInput | null {
  if (row.rowHeight === null || !Number.isFinite(row.rowHeight)) return null;
  // The stable public model predates authored-presence retention. Keep its
  // compatibility fallback at the model boundary, never in the layout solver.
  const normalized = normalizedTableHeightRule(row.rowHeightRule);
  return {
    rule: normalized === 'auto' ? 'atLeast' : normalized,
    valuePt: row.rowHeight,
  };
}

function wordTableCellSpacingPt(
  ...widths: readonly (TableWidthAcquisitionWire | null | undefined)[]
): number | null {
  for (const width of widths) {
    if (!width) continue;
    // `word-table-cell-spacing-scope-shadow` resolves authored pct/auto/nil at
    // this precedence scope instead of exposing a lower scope.
    const kind = effectiveTableWidthKind(width);
    const valuePt = tableDxaPtFromLexical(width);
    const resolved = wordTableCellSpacingValuePt(kind, valuePt);
    if (resolved !== null) return resolved;
  }
  return null;
}

type TableMarginScope = 'cell' | 'exception' | 'table' | 'style';
type TableMarginEdge = 'top' | 'bottom' | 'start' | 'end';

function wordTableMarginPt(
  width: TableWidthAcquisitionWire | null | undefined,
  scope: TableMarginScope,
  edge: TableMarginEdge,
): number | null {
  if (!width) return null;
  const kind = effectiveTableWidthKind(width);
  return wordTableMarginValuePt({
    kind,
    dxaValuePt: kind === 'dxa' ? tableTwipsValuePt(width.value ?? '0') : null,
    scope,
    edge,
  });
}

function effectiveTableCellMargins(
  table: TableLayoutSource,
  cell: TableLayoutSource['rows'][number]['cells'][number],
  hasPrivateCellWire: boolean,
  cellMargins: TableMarginAcquisitionWire | null | undefined,
  exceptionMargins: TableMarginAcquisitionWire | null | undefined,
  tableMargins: TableMarginAcquisitionWire | null | undefined,
  styleMargins: TableMarginAcquisitionWire | null | undefined,
): TableFormatInput['rows'][number]['cells'][number]['marginsPt'] {
  const bidi = table.bidiVisual === true;
  const physical = (
    margins: TableMarginAcquisitionWire | null | undefined,
    edge: 'left' | 'right',
  ): Readonly<{ width: TableWidthAcquisitionWire | null | undefined; edge: 'start' | 'end' }> => {
    const logicalEdge = edge === 'left'
      ? (bidi ? 'end' : 'start')
      : (bidi ? 'start' : 'end');
    return { width: margins?.[edge] ?? margins?.[logicalEdge], edge: logicalEdge };
  };
  const firstMargin = (
    edge: TableMarginEdge,
    ...candidates: readonly Readonly<{
      width: TableWidthAcquisitionWire | null | undefined;
      scope: TableMarginScope;
      edge?: TableMarginEdge;
    }>[]
  ): number | null => {
    for (const candidate of candidates) {
      const value = wordTableMarginPt(candidate.width, candidate.scope, candidate.edge ?? edge);
      if (value !== null) return value;
    }
    return null;
  };
  const cellLeft = physical(cellMargins, 'left');
  const exceptionLeft = physical(exceptionMargins, 'left');
  const tableLeft = physical(tableMargins, 'left');
  const styleLeft = physical(styleMargins, 'left');
  const cellRight = physical(cellMargins, 'right');
  const exceptionRight = physical(exceptionMargins, 'right');
  const tableRight = physical(tableMargins, 'right');
  const styleRight = physical(styleMargins, 'right');
  const publicCellMargin = (value: number | null | undefined): number | null => (
    !hasPrivateCellWire && value != null && Number.isFinite(value) ? value : null
  );
  return {
    top: firstMargin('top',
      { width: cellMargins?.top, scope: 'cell' },
    ) ?? publicCellMargin(cell.marginTop) ?? firstMargin('top',
      { width: exceptionMargins?.top, scope: 'exception' },
      { width: tableMargins?.top, scope: 'table' },
      { width: styleMargins?.top, scope: 'style' },
    ) ?? table.cellMarginTop,
    bottom: firstMargin('bottom',
      { width: cellMargins?.bottom, scope: 'cell' },
    ) ?? publicCellMargin(cell.marginBottom) ?? firstMargin('bottom',
      { width: exceptionMargins?.bottom, scope: 'exception' },
      { width: tableMargins?.bottom, scope: 'table' },
      { width: styleMargins?.bottom, scope: 'style' },
    ) ?? table.cellMarginBottom,
    left: firstMargin(cellLeft.edge,
      { ...cellLeft, scope: 'cell' },
    ) ?? publicCellMargin(cell.marginLeft) ?? firstMargin(exceptionLeft.edge,
      { ...exceptionLeft, scope: 'exception' },
      { ...tableLeft, scope: 'table' },
      { ...styleLeft, scope: 'style' },
    ) ?? table.cellMarginLeft,
    right: firstMargin(cellRight.edge,
      { ...cellRight, scope: 'cell' },
    ) ?? publicCellMargin(cell.marginRight) ?? firstMargin(exceptionRight.edge,
      { ...exceptionRight, scope: 'exception' },
      { ...tableRight, scope: 'table' },
      { ...styleRight, scope: 'style' },
    ) ?? table.cellMarginRight,
  };
}

function normalizedTableRowException(
  exception: TablePropertyExceptionAcquisitionWire | null | undefined,
): TableRowExceptionInput | null {
  if (!exception) return null;
  const indentKind = exception.indent ? effectiveTableWidthKind(exception.indent) : null;
  return {
    preferredWidthAuthored: exception.preferredWidth != null,
    preferredWidth: tableWidthConstraintFromLexical(exception.preferredWidth),
    layout: exception.layout?.kind === 'fixed' || exception.layout?.kind === 'autofit'
      ? exception.layout.kind
      : null,
    justification: exception.justification,
    indentAuthored: exception.indent != null && (indentKind === 'dxa' || indentKind === 'nil'),
    indentPt: indentKind === 'nil'
      ? 0
      : tableDxaPtFromLexical(exception.indent),
    borders: exception.borders,
  };
}

/** Resolve parser-private and public-compatibility table formatting once. */
export function tableFormatInput(table: TableLayoutSource): TableFormatInput {
  const cached = tableFormatInputs.get(table);
  if (cached) return cached;
  const acquisition = tableAcquisitionInput(table);
  const ordinaryFlow = acquisition.table?.ordinaryFlow ?? table.tblpPr == null;
  const rows = table.rows.map((row, rowIndex) => {
    const rowWire = acquisition.rows[rowIndex]?.row ?? null;
    const exception = rowWire?.exception ?? null;
    return {
      height: rowWire?.height ? privateTableRowHeight(rowWire.height) : publicTableRowHeight(row),
      cantSplit: row.cantSplit === true,
      repeatedHeader: row.isHeader === true,
      cellSpacingPt: wordTableCellSpacingPt(
        rowWire?.cellSpacing,
        exception?.cellSpacing,
        acquisition.table?.cellSpacing,
        rowWire?.styleCellSpacing,
      ) ?? 0,
      justification: rowWire?.justification ?? exception?.justification ?? null,
      exception: normalizedTableRowException(exception),
      cells: row.cells.map((cell, cellIndex) => ({
        marginsPt: effectiveTableCellMargins(
          table,
          cell,
          acquisition.rows[rowIndex]?.cells[cellIndex] !== null
            && acquisition.rows[rowIndex]?.cells[cellIndex] !== undefined,
          acquisition.rows[rowIndex]?.cells[cellIndex]?.margins,
          exception?.cellMargins,
          acquisition.table?.cellMargins,
          rowWire?.styleCellMargins,
        ),
      })),
    };
  });
  const input = snapshotPlainData({
    effectiveStyleId: acquisition.table?.effectiveStyleId ?? null,
    ordinaryFlow,
    // §17.4.37 membership is decided by the parser; carry its facts verbatim.
    logicalSequenceId: acquisition.table?.logicalSequenceId ?? null,
    logicalRowOffset: acquisition.table?.logicalRowOffset ?? 0,
    logicalTotalRows: acquisition.table?.logicalTotalRows ?? 0,
    positioning: ordinaryFlow || table.tblpPr == null
      ? null
      : floatingTablePositionInput(table.tblpPr),
    rows,
    // `word-first-row-table-exception-scope` applies these selected facts
    // table-wide.
    firstRowException: rows[0]?.exception ?? null,
  }, 'DOCX table format input') as TableFormatInput;
  tableFormatInputs.set(table, input);
  return input;
}

export function adjacentTableSequenceInput(
  body: readonly BodyElement[],
): readonly import('./layout/adjacent-tables.js').AdjacentTableSequenceInput[] {
  return Object.freeze(body.map((element) => {
    if (element.type !== 'table') return Object.freeze({ element, table: null });
    const format = tableFormatInput(element);
    // A hand-built public table has no acquisition wire and therefore no
    // parser-owned logical identity; it can never join a logical sequence.
    if (format.logicalSequenceId == null) return Object.freeze({ element, table: null });
    return Object.freeze({
      element,
      table: Object.freeze({
        logicalSequenceId: format.logicalSequenceId,
        logicalRowOffset: format.logicalRowOffset ?? 0,
        logicalTotalRows: format.logicalTotalRows ?? 0,
        rowCount: element.rows.length,
      }),
    });
  }));
}

type AcquiredBodySectionReference = Readonly<{
  sectionOccurrenceId: string;
  startType: string;
}>;

const bodySourceAt = (bodyIndex: number): SourceRef => Object.freeze({
  story: 'body', storyInstance: 'body', path: Object.freeze([bodyIndex]),
});

export interface PublicAnchorBridge {
  readonly occurrenceId: string;
  readonly pageOwned: boolean;
}

const PUBLIC_HOST_RELATIVE_ANCHOR_FRAMES = new Set(['paragraph', 'line', 'character']);

/** Compatibility projection for hand-built public anchor runs. Parser-owned
 * DrawingML facts always win and therefore never enter this fallback bridge. */
export function publicAnchorBridge(
  run: Readonly<DocRun>,
  source: SourceRef,
  runIndex: number,
): PublicAnchorBridge | null {
  if (
    (run.type !== 'shape' && run.type !== 'image' && run.type !== 'chart')
    || anchorAcquisitionInput(run) !== undefined
    || !isWrapFloat(run.wrapMode)
    || (run.type !== 'shape' && !run.anchor)
    || run.widthPt <= 0
    || run.heightPt <= 0
  ) return null;
  const horizontalReference = run.anchorXRelativeFrom
    ?? (run.anchorXFromMargin ? 'margin' : 'page');
  const verticalReference = run.anchorYRelativeFrom
    ?? (run.anchorYFromPara ? 'paragraph' : 'page');
  const paragraphId = `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
  return Object.freeze({
    occurrenceId: run.type === 'shape'
      ? `public-shape:${paragraphId}:${runIndex}`
      : `public-anchor:${paragraphId}:${runIndex}`,
    pageOwned: !PUBLIC_HOST_RELATIVE_ANCHOR_FRAMES.has(horizontalReference)
      && !PUBLIC_HOST_RELATIVE_ANCHOR_FRAMES.has(verticalReference),
  });
}

function acquiredBodyParagraph(paragraph: DocParagraph, source: SourceRef) {
  const hostRelative = new Set(['paragraph', 'line', 'character']);
  const pageOwnedAnchorOccurrenceIds = Object.freeze([...new Set(
    paragraphRunsWithUnavailableDrawings(paragraph).flatMap((run, runIndex) => {
      const internalRun = run as Readonly<DocRun | UnavailableDrawingRunWire>;
      if (
        run.type !== 'shape'
        && run.type !== 'image'
        && run.type !== 'chart'
        && internalRun.type !== 'unavailableDrawing'
      ) return [];
      const acquisition = anchorAcquisitionInput(internalRun);
      if (!acquisition) {
        const bridge = internalRun.type === 'unavailableDrawing'
          ? null
          : publicAnchorBridge(internalRun as Readonly<DocRun>, source, runIndex);
        return bridge?.pageOwned ? [bridge.occurrenceId] : [];
      }
      if (acquisition.horizontal.relativeFromStatus !== 'valid'
        || acquisition.vertical.relativeFromStatus !== 'valid'
        || acquisition.horizontal.relativeFrom === null
        || acquisition.vertical.relativeFrom === null
        || acquisition.wrap.kind === 'none'
        || hostRelative.has(acquisition.horizontal.relativeFrom)
        || hostRelative.has(acquisition.vertical.relativeFrom)) return [];
      return [anchorOccurrenceKey(source, acquisition.occurrenceId)];
    }),
  )]);
  return Object.freeze({
    kind: 'paragraph' as const,
    source,
    pageBreakBefore: paragraph.pageBreakBefore === true,
    keepLines: paragraph.keepLines === true,
    keepNext: paragraph.keepNext === true,
    widowControl: paragraph.widowControl !== false,
    spaceBeforePt: paragraph.spaceBefore ?? 0,
    spaceAfterPt: paragraph.spaceAfter ?? 0,
    contextualSpacing: paragraph.contextualSpacing === true,
    styleId: paragraph.styleId ?? null,
    inkless: !paragraphHasUnavailableDrawing(paragraph) && isInklessParagraph(paragraph),
    ...(pageOwnedAnchorOccurrenceIds.length === 0 ? {} : { pageOwnedAnchorOccurrenceIds }),
  });
}

function acquiredBodyTable(source: SourceRef) {
  return Object.freeze({
    kind: 'table' as const,
    source,
  });
}

function bodyLayoutSequenceInput(
  body: readonly BodyElement[],
  sectionAtMarker: (bodyIndex: number) => AcquiredBodySectionReference,
): readonly BodyLayoutSequenceEntryFor<AcquiredBodySectionReference>[] {
  let bodyIndex = 0;
  return Object.freeze(normalizeAdjacentTables(adjacentTableSequenceInput(body)).map((entry) => {
    if (entry.kind === 'adjacent-table-group') {
      const firstIndex = bodyIndex;
      bodyIndex += entry.tables.length;
      return Object.freeze({
        kind: 'adjacent-table-group' as const,
        logicalSequenceId: entry.logicalSequenceId,
        source: bodySourceAt(firstIndex),
        tables: Object.freeze(entry.tables.map((table, tableIndex) => Object.freeze({
          ...acquiredBodyTable(bodySourceAt(firstIndex + tableIndex)),
          rowCount: table.rows.length,
        }))),
      });
    }
    const element = entry.element;
    const entryBodyIndex = bodyIndex;
    const source = bodySourceAt(entryBodyIndex);
    bodyIndex += 1;
    if (element.type === 'paragraph') {
      return element.markVanish === true
        && !paragraphHasUnavailableDrawing(element)
        && isInklessParagraph(element)
        ? Object.freeze({ kind: 'consume-source' as const, source, reason: 'hidden-paragraph' as const })
        : Object.freeze({ kind: 'body-block' as const, block: acquiredBodyParagraph(element, source) });
    }
    if (element.type === 'table') {
      return Object.freeze({
        kind: 'body-block' as const,
        block: acquiredBodyTable(source),
      });
    }
    if (element.type === 'pageBreak' || element.type === 'columnBreak') {
      return Object.freeze({
        kind: 'authored-break' as const,
        source,
        break: element.type === 'pageBreak' ? 'page' as const : 'column' as const,
        ...(element.type === 'pageBreak' && element.parity !== undefined
          ? { parity: element.parity }
          : {}),
        ...(element.type === 'pageBreak' && element.sameParagraphAsPrevious === true
          ? { sameSourceParagraphAsPrevious: true }
          : {}),
      });
    }
    if (element.type === 'sectionBreak') {
      return Object.freeze({
        kind: 'begin-section' as const,
        source,
        section: sectionAtMarker(entryBodyIndex),
      });
    }
    throw new Error(`Unsupported body layout source at ${entryBodyIndex}`);
  }));
}

/** Project normalized parser/model facts into the pure §17.18.87 solver contract. */
export function tableColumnLayoutInput(
  table: TableLayoutSource,
  availableWidthPt: number,
  intrinsicWidths: (cell: TableLayoutSource['rows'][number]['cells'][number]) => CellIntrinsicWidths,
  maximumWidthPt: number | null = availableWidthPt,
): import('./layout/types.js').TableColumnLayoutInput {
  const source = tableSourceAcquisitionInput(table);
  return projectTableColumnLayoutInput(
    source,
    availableWidthPt,
    (rowIndex, cellIndex) => intrinsicWidths(table.rows[rowIndex]!.cells[cellIndex]!),
    maximumWidthPt,
  );
}

function setBodySectionPlacementInputs(
  body: readonly BodyElement[],
  finalSection: DocxDocumentModel['section'] | undefined,
  inputs: DocumentSectionPlacementInputs,
): void {
  if (!finalSection || typeof finalSection !== 'object') return;
  let byFinalSection = sectionPlacementInputsByBody.get(body);
  if (!byFinalSection) {
    byFinalSection = new WeakMap<object, DocumentSectionPlacementInputs>();
    sectionPlacementInputsByBody.set(body, byFinalSection);
  }
  byFinalSection.set(finalSection, inputs);
}

function projectSectionPlacementInputs(doc: InternalDocxDocumentModel): DocumentSectionPlacementInputs {
  const endingSections = new Map<number, SectionPlacementInput>();
  let ordinal = 0;
  doc.body.forEach((element, bodyIndex) => {
    if (element.type !== 'sectionBreak') return;
    const wire = (element as InternalSectionBreak).__sectionPlacement;
    endingSections.set(bodyIndex, snapshotPlainData({
      sectionId: wire?.sectionId ?? `section:${ordinal}`,
      sectionBidi: wire?.sectionBidi === true,
      vAlign: wire?.vAlign ?? null,
      lineNumbering: wire?.lineNumbering ?? null,
      docGridType: wire?.docGridType ?? null,
      docGridLinePitch: wire?.docGridLinePitch ?? null,
      docGridCharSpace: wire?.docGridCharSpace ?? null,
      gutterPt: wire?.gutterPt ?? null,
      rtlGutter: wire?.rtlGutter ?? null,
      pageBordersAuthored: wire?.pageBordersAuthored ?? false,
      pageBorders: wire?.pageBorders ?? null,
      pageGeometry: wire?.pageGeometry ?? element.geom ?? {},
    }, 'DOCX ending-section placement input'));
    ordinal += 1;
  });
  const finalWire = (doc.section as InternalSectionProps | undefined)?.__sectionPlacement;
  return Object.freeze({
    endingSections,
    finalSection: snapshotPlainData({
      sectionId: finalWire?.sectionId ?? `section:${ordinal}`,
      sectionBidi: finalWire?.sectionBidi === true,
      // Resource-only entry points (for example image preloading) historically
      // accept a partial document projection with no section. Section placement
      // is irrelevant there, so preserve that compatibility with neutral facts.
      vAlign: finalWire?.vAlign ?? doc.section?.vAlign ?? null,
      lineNumbering: finalWire?.lineNumbering ?? doc.section?.lineNumbering ?? null,
      docGridType: finalWire?.docGridType ?? doc.section?.docGridType ?? null,
      docGridLinePitch: finalWire?.docGridLinePitch ?? doc.section?.docGridLinePitch ?? null,
      docGridCharSpace: finalWire?.docGridCharSpace ?? doc.section?.docGridCharSpace ?? null,
      gutterPt: finalWire?.gutterPt ?? null,
      rtlGutter: finalWire?.rtlGutter ?? null,
      pageBordersAuthored: finalWire?.pageBordersAuthored ?? doc.section?.pageBorders != null,
      pageBorders: finalWire?.pageBorders ?? doc.section?.pageBorders ?? null,
      pageGeometry: finalWire?.pageGeometry
        ?? (doc.section ? sectionPageBox(doc.section) : {}),
    }, 'DOCX final-section placement input'),
  });
}

/** Resolve the section which owns body content beginning at `startIndex`.
 * Non-final section facts come from the next terminating SectionBreak; the
 * body-level sectPr owns the final section. */
export function sectionPlacementInputFrom(
  doc: InternalDocxDocumentModel,
  startIndex: number,
): SectionPlacementInput {
  let inputs = sectionPlacementInputsByDocument.get(doc);
  if (!inputs) {
    inputs = projectSectionPlacementInputs(doc);
    sectionPlacementInputsByDocument.set(doc, inputs);
  }
  for (let index = startIndex; index < doc.body.length; index += 1) {
    if (doc.body[index]?.type !== 'sectionBreak') continue;
    return inputs.endingSections.get(index) ?? inputs.finalSection;
  }
  return inputs.finalSection;
}

/** Body-array keyed twin used by the paginator, whose stable public signature
 * receives body + final SectionProps rather than the document wrapper. */
export function sectionPlacementInputFromBody(
  body: readonly BodyElement[],
  finalSection: DocxDocumentModel['section'],
  startIndex: number,
): SectionPlacementInput {
  let inputs = sectionPlacementInputsByBody.get(body)?.get(finalSection);
  if (!inputs) {
    const synthetic = { body, section: finalSection } as InternalDocxDocumentModel;
    inputs = projectSectionPlacementInputs(synthetic);
    setBodySectionPlacementInputs(body, finalSection, inputs);
  }
  for (let index = startIndex; index < body.length; index += 1) {
    if (body[index]?.type !== 'sectionBreak') continue;
    return inputs.endingSections.get(index) ?? inputs.finalSection;
  }
  return inputs.finalSection;
}

const EMPTY_SECTION_HEADERS_FOOTERS: HeadersFooters = Object.freeze({
  default: null,
  first: null,
  even: null,
});

/**
 * Acquire every §17.6.18 paragraph-owned occurrence and the final §17.6.17
 * occurrence before layout. Equal-looking sections remain distinct entries;
 * layout receives no document-model scanning capability or parser-private wire.
 */
export function bodySectionIndexInput(doc: DocxDocumentModel): BodySectionIndexInput {
  type PendingOccurrence = Omit<BodySectionOccurrence, 'geometry' | 'gutterPt'> & {
    readonly authoredGeometry: Readonly<Partial<import('./types.js').SectionGeom>>;
    readonly authoredGutterPt: number | null;
  };
  const pending: PendingOccurrence[] = [];
  const placements = projectSectionPlacementInputs(doc as InternalDocxDocumentModel);
  let startBodyIndex = 0;

  doc.body.forEach((element, bodyIndex) => {
    if (element.type !== 'sectionBreak') return;
    const placement = placements.endingSections.get(bodyIndex) ?? placements.finalSection;
    const ordinal = pending.length;
    pending.push({
      sectionOccurrenceId: placement.sectionId,
      ordinal,
      startBodyIndex,
      endBodyIndex: bodyIndex,
      markerBodyIndex: bodyIndex,
      final: false,
      startType: element.kind ?? 'nextPage',
      columns: element.columns ?? null,
      authoredGeometry: normalizeSectionGeometryWire(placement.pageGeometry),
      textDirection: element.textDirection ?? null,
      pageNumType: element.pageNumType ?? null,
      headers: element.headers ?? EMPTY_SECTION_HEADERS_FOOTERS,
      footers: element.footers ?? EMPTY_SECTION_HEADERS_FOOTERS,
      titlePage: element.titlePage ?? false,
      sectionBidi: placement.sectionBidi,
      vAlign: placement.vAlign,
      lineNumbering: placement.lineNumbering,
      docGridType: placement.docGridType,
      docGridLinePitch: placement.docGridLinePitch,
      docGridCharSpace: placement.docGridCharSpace,
      authoredGutterPt: placement.gutterPt,
      rtlGutter: placement.rtlGutter === true,
      pageBordersAuthored: placement.pageBordersAuthored,
      pageBorders: placement.pageBorders,
      placement: sectionPlacementFacts(placement),
    });
    startBodyIndex = bodyIndex + 1;
  });

  const placement = placements.finalSection;
  pending.push({
    sectionOccurrenceId: placement.sectionId,
    ordinal: pending.length,
    startBodyIndex,
    endBodyIndex: doc.body.length - 1,
    markerBodyIndex: null,
    final: true,
    startType: doc.section.sectionStart ?? 'nextPage',
    columns: doc.section.columns ?? null,
    authoredGeometry: placement.pageGeometry == null
      ? sectionPageBox(doc.section)
      : normalizeSectionGeometryWire(placement.pageGeometry),
    textDirection: doc.section.textDirection ?? null,
    pageNumType: doc.section.pageNumType ?? null,
    headers: doc.headers ?? EMPTY_SECTION_HEADERS_FOOTERS,
    footers: doc.footers ?? EMPTY_SECTION_HEADERS_FOOTERS,
    titlePage: doc.section.titlePage,
    sectionBidi: placement.sectionBidi,
    vAlign: placement.vAlign,
    lineNumbering: placement.lineNumbering,
    docGridType: placement.docGridType,
    docGridLinePitch: placement.docGridLinePitch,
    docGridCharSpace: placement.docGridCharSpace,
    authoredGutterPt: placement.gutterPt,
    rtlGutter: placement.rtlGutter === true,
    pageBordersAuthored: placement.pageBordersAuthored,
    pageBorders: placement.pageBorders,
    placement: sectionPlacementFacts(placement),
  });

  const occurrences = new Array<BodySectionOccurrence>(pending.length);
  const documentGeometry = sectionPageBox(doc.section);
  let followingGeometry: import('./types.js').SectionGeom | null = null;
  let followingGutterPt: number | null = null;
  for (let index = pending.length - 1; index >= 0; index -= 1) {
    const occurrence = pending[index]!;
    // pgSz/pgMar are optional children and ECMA-376 does not define a Letter /
    // one-inch default for an omitted page box. Preserve the renderer's
    // established document-level fallback for non-continuous sections. Per
    // §17.18.77, a continuous section instead inherits omitted page-level facts
    // from the following section. Authored fields remain occurrence-local.
    const fallback: import('./types.js').SectionGeom =
      occurrence.startType === 'continuous' && followingGeometry !== null
        ? followingGeometry
        : documentGeometry;
    const authored = occurrence.authoredGeometry;
    const geometry: import('./types.js').SectionGeom = {
      pageWidth: authored.pageWidth ?? fallback.pageWidth,
      pageHeight: authored.pageHeight ?? fallback.pageHeight,
      marginTop: authored.marginTop ?? fallback.marginTop,
      marginRight: authored.marginRight ?? fallback.marginRight,
      marginBottom: authored.marginBottom ?? fallback.marginBottom,
      marginLeft: authored.marginLeft ?? fallback.marginLeft,
      headerDistance: authored.headerDistance ?? fallback.headerDistance,
      footerDistance: authored.footerDistance ?? fallback.footerDistance,
    };
    const gutterPt: number = occurrence.authoredGutterPt
      ?? (occurrence.startType === 'continuous' ? followingGutterPt : null)
      ?? 0;
    const {
      authoredGeometry: _authoredGeometry,
      authoredGutterPt: _authoredGutterPt,
      ...facts
    } = occurrence;
    occurrences[index] = { ...facts, geometry, gutterPt };
    followingGeometry = geometry;
    followingGutterPt = gutterPt;
  }

  return snapshotPlainData({
    bodyLength: doc.body.length,
    occurrences,
  }, 'DOCX body section index input') as BodySectionIndexInput;
}

/** Consume parser-owned document nodes and private settings into one clone-safe
 * structural value before layout resolves section contexts. */
export function bodyLayoutAcquisitionInput(doc: DocxDocumentModel): BodyLayoutAcquisitionInput {
  const sectionIndex = bodySectionIndexInput(doc);
  const incomingByMarker = new Map<number, BodySectionOccurrence>();
  for (const occurrence of sectionIndex.occurrences) {
    if (occurrence.startBodyIndex === 0) continue;
    incomingByMarker.set(occurrence.startBodyIndex - 1, occurrence);
  }
  const sequence = bodyLayoutSequenceInput(doc.body, (markerBodyIndex) => {
    const occurrence = incomingByMarker.get(markerBodyIndex);
    if (!occurrence) throw new Error(`Missing incoming body section at ${markerBodyIndex}`);
    return Object.freeze({
      sectionOccurrenceId: occurrence.sectionOccurrenceId,
      startType: occurrence.startType,
    });
  });
  return snapshotPlainData({
    sectionIndex,
    evenAndOddHeaders: doc.section.evenAndOddHeaders,
    endnoteIds: (doc.endnotes ?? []).map((note) => note.id),
    noteLayoutSettings: documentNoteLayoutSettingsInput(doc),
    pageLayoutSettings: documentPageLayoutSettingsInput(doc),
    parserDiagnostics: mapParseDiagnostics(
      (doc as InternalDocxDocumentModel).diagnostics,
      doc.body.length,
    ),
    sequence,
  }, 'DOCX body layout acquisition input') as BodyLayoutAcquisitionInput;
}

/** Resolved transitional VML facts emitted by the parser in addition to the
 * stable public `TextPath` surface. CT_Path owns `textPathOk`; CT_TextPath owns
 * the remaining switches. They stay private because they are acquisition
 * policy, not consumer-facing document content. */
export interface InternalVmlTextPath extends TextPath {
  textPathOk?: boolean;
  on?: boolean;
  fitShape?: boolean;
  fitPath?: boolean;
  trim?: boolean;
  xScale?: boolean;
  fontSizePt?: number;
}

export interface InternalUnsupportedTextBoxBlock {
  readonly type: 'unsupportedTextBoxBlock';
  readonly qName: string;
  readonly sourcePath: readonly number[];
}

export type InternalTextBoxBlock =
  | Extract<BodyElement, { type: 'paragraph' | 'table' }>
  | InternalUnsupportedTextBoxBlock;

export interface InternalShapeRun extends ShapeRun {
  textPath?: InternalVmlTextPath | null;
  textBoxContent?: InternalTextBoxBlock[];
}

export interface NormalizedDocumentInput {
  readonly document: InternalDocxDocumentModel;
  readonly mathOccurrences: readonly MathOccurrence[];
  readonly fontFamilyCharsets: Readonly<Record<string, string>>;
  readonly bodyLayoutInput: BodyLayoutInput;
  readonly bodyModelGateway: Readonly<{
    acquisitionInputs: BodyAcquisitionInputProjections;
    bodySectionIndex: BodySectionIndexInput;
    effectiveTablePositioning: typeof effectiveTablePositioning;
    publicAnchorBridge: typeof publicAnchorBridge;
  }>;
}

/** Snapshot VML WordArt semantics at the parser/model boundary. The retained
 * acquisition layer consumes this clone-safe value and never needs to inspect
 * parser-only extensions on the public ShapeRun object. */
export function vmlTextPathAcquisitionInput(
  shape: Readonly<ShapeRun>,
): Readonly<VmlTextPathAcquisitionInput> | undefined {
  const textPath = (shape as Readonly<InternalShapeRun>).textPath;
  if (!textPath) return undefined;
  return snapshotPlainData({
    string: textPath.string,
    ...(textPath.fontFamily !== undefined ? { fontFamily: textPath.fontFamily } : {}),
    bold: textPath.bold ?? false,
    italic: textPath.italic ?? false,
    ...(textPath.textPathOk !== undefined ? { textPathOk: textPath.textPathOk } : {}),
    ...(textPath.on !== undefined ? { on: textPath.on } : {}),
    ...(textPath.fitShape !== undefined ? { fitShape: textPath.fitShape } : {}),
    ...(textPath.fitPath !== undefined ? { fitPath: textPath.fitPath } : {}),
    ...(textPath.trim !== undefined ? { trim: textPath.trim } : {}),
    ...(textPath.xScale !== undefined ? { xScale: textPath.xScale } : {}),
    ...(textPath.fontSizePt !== undefined ? { fontSizePt: textPath.fontSizePt } : {}),
  }, 'DOCX VML text path acquisition input');
}

/** Snapshot the parser's complete CT_TxbxContent block sequence without
 * widening the stable public ShapeRun contract. Layout migration consumes this
 * plain-data boundary; unsupported schema-permitted children stay ordered and
 * diagnostic instead of disappearing. */
export function textBoxContentAcquisitionInput(
  shape: Readonly<ShapeRun>,
): readonly InternalTextBoxBlock[] | undefined {
  const content = (shape as Readonly<InternalShapeRun>).textBoxContent;
  if (content === undefined) return undefined;
  return snapshotPlainData(
    content,
    'DOCX text box content acquisition input',
  ) as readonly InternalTextBoxBlock[];
}

/** Project parser-only anchor facts without widening the public run contract.
 * Parser-produced malformed input remains distinguishable from a hand-built
 * public run because required-but-missing values are explicit nulls. */
export function anchorAcquisitionInput(
  run: Readonly<object>,
): Readonly<AnchorAcquisitionInput> | undefined {
  const wire = (run as Readonly<InternalAnchorRunWire>).__anchorAcquisition;
  if (wire === undefined) return undefined;
  return snapshotPlainData(wire, 'DOCX anchor acquisition input');
}

/** Snapshot the parser's effective numbering-level rPr into the plain retained
 * layout contract. This is the parser-model/layout boundary: layout code never
 * dereferences the private parser extension itself. */
export function numberingMarkerShapeInput(
  num: NumberingInfo,
  fallbackFontSizePt: number,
): NumberingMarkerShapeInput {
  const facts = internalNumberingInfo(num).fontFacts;
  const complexScript = facts?.rtl === true || facts?.cs === true;
  const fontSizePt = complexScript
    ? (facts?.fontSizeCs ?? facts?.fontSize ?? fallbackFontSizePt)
    : (facts?.fontSize ?? fallbackFontSizePt);
  const ascii = facts?.fontFamily ?? num.fontFamily ?? null;
  const fallbackFonts: TextFontSlots = {
    ascii,
    highAnsi: facts?.fontFamilyHighAnsi ?? ascii,
    eastAsia: facts?.fontFamilyEastAsia ?? num.fontFamilyEastAsia ?? ascii,
    complexScript: facts?.fontFamilyCs ?? ascii,
  };
  const slots = facts?.fontSlots;
  return Object.freeze({
    fontSizePt,
    fonts: Object.freeze({ ...(slots?.direct ?? fallbackFonts) }),
    themeFonts: slots?.theme ? Object.freeze({ ...slots.theme }) : undefined,
    themeFontPresence: slots?.themePresent
      ? Object.freeze({ ...slots.themePresent })
      : undefined,
    weight: (complexScript ? (facts?.boldCs ?? false) : (facts?.bold ?? false)) ? 700 : 400,
    style: (complexScript ? (facts?.italicCs ?? false) : (facts?.italic ?? false))
      ? 'italic'
      : 'normal',
    complexScript,
    fontHint: facts?.fontHint,
    eastAsiaLanguage: facts?.langEastAsia,
    kerning: facts?.kerning == null ? undefined : fontSizePt >= facts.kerning,
  });
}

/** Project effective numbering-level run properties before a shape crosses the
 * parser/layout boundary. Public hand-built ShapeRun values use the normalizer's
 * compatibility fallback; parser-created shapes retain the full resolved slot
 * and theme facts without exposing their private wire object to layout. */
export function textBoxAcquisitionInput(
  shape: Readonly<ShapeRun>,
  source: SourceRef,
): TextBoxAcquisitionInput {
  const content = (shape as InternalShapeRun).textBoxContent;
  return content === undefined
      ? snapshotPlainData({
        kind: 'compatibility',
        source,
        paragraphs: normalizeTextBoxInput(shape, source, numberingMarkerShapeInput),
      }, 'DOCX public text box acquisition input')
    : snapshotPlainData({
        kind: 'complete',
        source,
        blockCount: content.length,
      }, 'DOCX complete text box acquisition input');
}

/** Snapshot private paragraph-mark rPr facts at the parser boundary. Retained
 * line layout receives only this plain immutable service input. */
export function paragraphMarkShapeInput(
  paragraph: ParagraphLayoutSource,
): NumberingMarkerShapeInput | undefined {
  const facts = internalParagraph(paragraph).paragraphMarkFontFacts;
  if (!facts) return undefined;
  const complexScript = facts.rtl === true || facts.cs === true;
  const fallbackFontSizePt = paragraph.runs.find(
    (run): run is Extract<DocRun, { type: 'text' | 'field' }> => run.type === 'text' || run.type === 'field',
  )?.fontSize ?? paragraph.defaultFontSize ?? 10;
  const fontSizePt = complexScript
    ? (facts.fontSizeCs ?? facts.fontSize ?? fallbackFontSizePt)
    : (facts.fontSize ?? fallbackFontSizePt);
  const ascii = facts.fontFamily ?? paragraph.defaultFontFamily ?? null;
  const fallbackFonts: TextFontSlots = {
    ascii,
    highAnsi: facts.fontFamilyHighAnsi ?? ascii,
    eastAsia: facts.fontFamilyEastAsia ?? paragraph.defaultFontFamilyEastAsia ?? ascii,
    complexScript: facts.fontFamilyCs ?? ascii,
  };
  return Object.freeze({
    fontSizePt,
    fonts: Object.freeze({ ...(facts.fontSlots?.direct ?? fallbackFonts) }),
    themeFonts: facts.fontSlots?.theme ? Object.freeze({ ...facts.fontSlots.theme }) : undefined,
    themeFontPresence: facts.fontSlots?.themePresent
      ? Object.freeze({ ...facts.fontSlots.themePresent }) : undefined,
    weight: (complexScript ? (facts.boldCs ?? false) : (facts.bold ?? false)) ? 700 : 400,
    style: (complexScript ? (facts.italicCs ?? false) : (facts.italic ?? false)) ? 'italic' : 'normal',
    complexScript,
    fontHint: facts.fontHint,
    eastAsiaLanguage: facts.langEastAsia,
    kerning: facts.kerning == null ? undefined : fontSizePt >= facts.kerning,
  });
}

/** Immutable all-run snapshot for the retained paragraph acquisition kernel. */
export function paragraphAcquisitionInput(
  paragraph: ParagraphLayoutSource,
  source: SourceRef,
): ParagraphAcquisitionInput {
  const parserParagraph = paragraph as unknown as DocParagraph;
  // Table pagination may have attached legacy cache stamps containing live font
  // resolver functions. They are renderer state, not parser/model facts, and must
  // never cross the retained acquisition boundary.
  const {
    layoutLines: _layoutLines,
    lineSlice: _lineSlice,
    runs: _runs,
    paragraphMarkFontFacts: _privateParagraphMarkFontFacts,
    __paragraphTypographyAcquisition: _privateParagraphTypography,
    __complexFieldBoundaries: _privateComplexFieldBoundaries,
    __runRevisions: _privateRunRevisions,
    ...semanticParagraph
  } = paragraph as DocParagraph & Record<string, unknown> & {
    __paragraphTypographyAcquisition?: InternalParagraphTypographyWire;
    __complexFieldBoundaries?: readonly InternalComplexFieldBoundaryWire[];
    __runRevisions?: readonly (DocRun['revision'] | null)[];
  };
  const typographyInput = paragraphTypographyAcquisitionInput(parserParagraph);
  const complexFieldBoundaries = (
    paragraph as InternalDocParagraph
  ).__complexFieldBoundaries?.map((boundary): ComplexFieldBoundaryInput => ({
    occurrenceKey: [
      'complex-field',
      source.story,
      source.storyInstance,
      source.path.slice(0, -1).join('.'),
      String(boundary.occurrenceId),
    ].join(':'),
    boundary: boundary.boundary,
    runIndex: boundary.runIndex,
    fieldType: boundary.fieldType,
    instruction: boundary.instruction,
    ...(boundary.hyperlinkAnchor === undefined
      ? {}
      : { hyperlinkAnchor: boundary.hyperlinkAnchor }),
  }));
  const numbering = semanticParagraph.numbering as (NumberingInfo & { fontFacts?: InternalRunFontFacts }) | null;
  const canonicalNumbering = numbering == null ? null : (({
    fontFacts: _privateNumberingFontFacts,
    ...retainedNumbering
  }) => retainedNumbering)(numbering);
  const snapshot = structuredClone({
    ...semanticParagraph,
    numbering: canonicalNumbering,
  }) as Omit<DocParagraph, 'runs'>;
  const sidecarEntries = unavailableDrawingEntries(parserParagraph);
  const runPairs: Array<Readonly<{
    run: DocRun | UnavailableDrawingRunWire;
    originalRun: DocRun | UnavailableDrawingRunWire;
  }>> = [];
  if (sidecarEntries.length === 0) {
    parserParagraph.runs.forEach((run, runIndex) => {
      runPairs.push({ run, originalRun: parserParagraph.runs[runIndex]! });
    });
  } else {
    let entryIndex = 0;
    for (let publicRunIndex = 0; publicRunIndex <= parserParagraph.runs.length; publicRunIndex += 1) {
      while (sidecarEntries[entryIndex]?.publicRunIndex === publicRunIndex) {
        const originalRun = sidecarEntries[entryIndex]!.run;
        runPairs.push({
          run: originalRun,
          originalRun,
        });
        entryIndex += 1;
      }
      if (publicRunIndex < parserParagraph.runs.length) {
        runPairs.push({
          run: parserParagraph.runs[publicRunIndex]!,
          originalRun: parserParagraph.runs[publicRunIndex]!,
        });
      }
    }
  }
  const runs = runPairs.map(({ run, originalRun }, runIndex): ParagraphAcquisitionRun => {
    const internalRun = run as UnavailableDrawingRunWire;
    if (internalRun.type === 'unavailableDrawing') {
      const localAnchorInput = anchorAcquisitionInput(originalRun);
      const anchorInput = localAnchorInput === undefined
        ? undefined
        : snapshotPlainData({
            ...localAnchorInput,
            occurrenceId: anchorOccurrenceKey(source, localAnchorInput.occurrenceId),
          }, 'DOCX scoped unavailable drawing anchor acquisition input');
      const { __anchorAcquisition: _privateAnchor, ...retainedRun } = internalRun;
      return Object.freeze({
        ...retainedRun,
        ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
      });
    }
    if (run.type === 'math') {
      const runRef: SourceRef = Object.freeze({ ...source, path: Object.freeze([...source.path, runIndex]) });
      const internal = run as Partial<InternalMathRun>;
      return Object.freeze({
        type: 'math',
        display: run.display,
        fontSize: run.fontSize,
        ...(run.jc === undefined ? {} : { jc: run.jc }),
        source: internal.source ?? runRef,
        resourceKey: internal.resourceKey ?? mathResourceKey(runRef, run.display ? 'display' : 'inline'),
        fallbackText: mathFallbackText(run.nodes),
      });
    }
    if (run.type === 'anchorHost') {
      const internal = run as typeof run & { __anchorOccurrenceId?: string };
      const { __anchorOccurrenceId, ...host } = internal;
      return Object.freeze({
        ...host,
        ...(__anchorOccurrenceId === undefined
          ? {}
          : { anchorOccurrenceId: anchorOccurrenceKey(source, __anchorOccurrenceId) }),
      }) as ParagraphAcquisitionRun;
    }
    if (run.type === 'shape' || run.type === 'image' || run.type === 'chart') {
      const localAnchorInput = anchorAcquisitionInput(originalRun);
      const anchorInput = localAnchorInput === undefined
        ? undefined
        : snapshotPlainData({
            ...localAnchorInput,
            occurrenceId: anchorOccurrenceKey(source, localAnchorInput.occurrenceId),
          }, 'DOCX scoped anchor acquisition input');
      const { __anchorAcquisition: _privateAnchor, ...publicRun } = run as typeof run & InternalAnchorRunWire;
      if (run.type !== 'shape') {
        const retainedRun = run.type === 'chart'
          ? (({ chart: _chartPayload, ...marker }) => ({
              ...marker,
              resourceKey: chartResourceKey({ ...source, path: [...source.path, runIndex] }),
            }))(publicRun as Extract<DocRun, { type: 'chart' }>)
          : publicRun;
        return Object.freeze({
          ...structuredClone(retainedRun),
          ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
        }) as ParagraphAcquisitionRun;
      }
      const originalShape = originalRun as ShapeRun;
      const vmlTextPathInput = vmlTextPathAcquisitionInput(originalShape);
      const shapeSource: SourceRef = Object.freeze({
        ...source,
        path: Object.freeze([...source.path, runIndex]),
      });
      const textBoxInput = textBoxAcquisitionInput(originalShape, {
        story: 'textbox',
        storyInstance: `${shapeSource.story}:${shapeSource.storyInstance}:${shapeSource.path.join('.')}`,
        path: [],
      });
      const {
        textBoxContent: _privateTextBoxContent,
        textBlocks: _publicCompatibilityTextBlocks,
        textPath: _privateVmlTextPath,
        ...retainedShape
      } = publicRun as InternalShapeRun;
      return Object.freeze({
        type: 'shape' as const,
        ...structuredClone(retainedShape),
        ...(vmlTextPathInput === undefined ? {} : { vmlTextPathInput }),
        ...((textBoxInput.kind === 'complete'
          ? textBoxInput.blockCount
          : textBoxInput.paragraphs.length) === 0 ? {} : { textBoxInput }),
        ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
      }) as ParagraphAcquisitionRun;
    }
    if (run.type === 'text' || run.type === 'field') {
      const originalTextRun = originalRun as Extract<DocRun, { type: 'text' | 'field' }>;
      const runTypographyInput = runTypographyAcquisitionInput(originalTextRun);
      const {
        __typographyAcquisition: _privateRunTypography,
        __noBreakBefore: noBreakBefore,
        __noBreakAfter: noBreakAfter,
        __noBreakHyphenOffsets: noBreakHyphenOffsets,
        ...publicRun
      } = run as typeof run & InternalNoBreakHyphenWire & {
        __typographyAcquisition?: InternalRunTypographyWire;
      };
      const noBreakRanges = run.type === 'text'
        ? noBreakHyphenOffsets
          ?.filter((end) => Number.isInteger(end) && end > 0 && end <= run.text.length)
          .map((end) => Object.freeze({ start: end - 1, end }))
        : undefined;
      return Object.freeze({
        ...structuredClone(publicRun),
        ...(noBreakBefore === true ? { noBreakBefore: true } : {}),
        ...(noBreakAfter === true ? { noBreakAfter: true } : {}),
        ...(noBreakRanges?.length ? { noBreakRanges: Object.freeze(noBreakRanges) } : {}),
        ...(runTypographyInput === undefined ? {} : { typographyInput: runTypographyInput }),
      }) as ParagraphAcquisitionRun;
    }
    return Object.freeze(structuredClone(run)) as ParagraphAcquisitionRun;
  });
  return deepFreezePlainData({
    ...snapshot,
    runs: runs as readonly ParagraphAcquisitionRun[],
    ...(complexFieldBoundaries?.length
      ? { complexFieldBoundaries }
      : {}),
    numberingMarkerShapeInput: paragraph.numbering
      ? numberingMarkerShapeInput(
          paragraph.numbering,
          parserParagraph.runs.find(
            (run): run is Extract<DocRun, { type: 'text' | 'field' }> =>
              run.type === 'text' || run.type === 'field',
          )?.fontSize ?? paragraph.defaultFontSize ?? 10,
        )
      : undefined,
    paragraphMarkShapeInput: paragraphMarkShapeInput(paragraph),
    ...(typographyInput === undefined ? {} : { typographyInput }),
  }) as unknown as ParagraphAcquisitionInput;
}

/** Pure structural normalization for stable math addressing and parser-only
 * acquisition sidecars. Only affected ancestry is shallow-cloned; the caller's
 * parser model is untouched and the returned public model contains only the
 * declared `DocRun` union. */
export function normalizeInternalDocumentModel(doc: DocxDocumentModel): NormalizedDocumentInput {
  return normalizeInternalDocumentModelWithOwnership(doc, false);
}

/** Destructive equivalent for an exclusively builder-owned parser graph. It
 * preserves the same normalized contract while replacing changed ancestry in
 * place, avoiding another complete body-index graph at the stream terminal. */
export function normalizeOwnedInternalDocumentModel(doc: DocxDocumentModel): NormalizedDocumentInput {
  return normalizeInternalDocumentModelWithOwnership(doc, true);
}

function normalizeInternalDocumentModelWithOwnership(
  doc: DocxDocumentModel,
  consumeOwned: boolean,
): NormalizedDocumentInput {
  const occurrences: MathOccurrence[] = [];
  const normalizeElement = (
    element: BodyElement,
    story: SourceRef['story'],
    storyInstance: string,
    path: number[],
  ): BodyElement => {
    if (element.type === 'paragraph') {
      const internalParagraph = element as InternalDocParagraph;
      const runRevisions = internalParagraph.__runRevisions ?? [];
      let runsChanged = internalParagraph.__runRevisions !== undefined;
      const runs: DocRun[] = [];
      const unavailableDrawings: UnavailableDrawingSidecarEntry[] = [];
      const hasEmbeddedUnavailableDrawing = (
        element.runs as unknown as readonly Readonly<{ type: string }>[]
      ).some((run) => run.type === 'unavailableDrawing');
      paragraphRunsWithUnavailableDrawings(element).forEach((rawRun, runIndex) => {
        const revision = runRevisions[runIndex] ?? undefined;
        const rawRevision = (rawRun as Readonly<{ revision?: DocRun['revision'] }>).revision;
        const run = revision === undefined || rawRevision !== undefined
          ? rawRun
          : { ...rawRun, revision };
        if (run !== rawRun) runsChanged = true;
        if (run.type === 'unavailableDrawing') {
          unavailableDrawings.push(Object.freeze({
            publicRunIndex: runs.length,
            run: snapshotPlainData(
              run,
              'DOCX unavailable drawing parser sidecar',
            ) as Readonly<UnavailableDrawingRunWire>,
          }));
          if (hasEmbeddedUnavailableDrawing) runsChanged = true;
          return;
        }
        if (run.type === 'math') {
          runsChanged = true;
          const source: SourceRef = Object.freeze({
            story,
            storyInstance,
            path: Object.freeze([...path, runIndex]),
          });
          const resourceKey = mathResourceKey(source, run.display ? 'display' : 'inline');
          occurrences.push(Object.freeze({
            nodes: run.nodes,
            display: run.display,
            source,
            resourceKey,
          }));
          runs.push(Object.freeze({ ...run, source, resourceKey }) as InternalMathRun);
          return;
        }
        if (run.type !== 'shape') {
          runs.push(run);
          return;
        }
        const shape = run as InternalShapeRun;
        const content = shape.textBoxContent;
        if (content === undefined) {
          runs.push(run);
          return;
        }
        const shapeSource: SourceRef = {
          story,
          storyInstance,
          path: [...path, runIndex],
        };
        const textBoxStoryInstance = `${shapeSource.story}:${shapeSource.storyInstance}:${shapeSource.path.join('.')}`;
        let contentChanged = false;
        const textBoxContent = (consumeOwned ? content : new Array<InternalTextBoxBlock>(content.length));
        content.forEach((block, blockIndex): void => {
          if (block.type === 'unsupportedTextBoxBlock') {
            textBoxContent[blockIndex] = block;
            return;
          }
          const normalized = normalizeElement(
            block,
            'textbox',
            textBoxStoryInstance,
            [blockIndex],
          ) as Extract<BodyElement, { type: 'paragraph' | 'table' }>;
          if (normalized !== block) contentChanged = true;
          textBoxContent[blockIndex] = normalized;
        });
        if (!contentChanged) {
          runs.push(run);
          return;
        }
        runsChanged = true;
        if (consumeOwned) {
          shape.textBoxContent = textBoxContent;
          runs.push(run as DocRun);
        } else {
          runs.push({ ...run, textBoxContent } as DocRun);
        }
      });
      let paragraph: Extract<BodyElement, { type: 'paragraph' }>;
      if (consumeOwned) {
        if (runsChanged) Object.assign(element, { runs });
        delete (element as InternalDocParagraph).__runRevisions;
        paragraph = element;
      } else if (runsChanged) {
        const { __runRevisions: _privateRunRevisions, ...publicParagraph } = internalParagraph;
        paragraph = { ...publicParagraph, runs } as Extract<BodyElement, { type: 'paragraph' }>;
      } else {
        paragraph = element;
      }
      if (unavailableDrawings.length > 0) {
        unavailableDrawingSidecars.set(
          paragraph,
          Object.freeze(unavailableDrawings),
        );
      }
      return paragraph;
    }
    if (element.type === 'table') {
      if (consumeOwned) {
        element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          cell.content = normalizeBody(
            cell.content as BodyElement[], story, storyInstance, [...path, rowIndex, cellIndex],
          ) as typeof cell.content;
        }));
        return element;
      }
      let tableChanged = false;
      const rows = element.rows.map((row, rowIndex) => {
        let rowChanged = false;
        const cells = row.cells.map((cell, cellIndex) => {
          const content = normalizeBody(
            cell.content as BodyElement[], story, storyInstance, [...path, rowIndex, cellIndex],
          );
          if (content === cell.content) return cell;
          rowChanged = true;
          return { ...cell, content: content as typeof cell.content };
        });
        if (!rowChanged) return row;
        tableChanged = true;
        return { ...row, cells };
      });
      return tableChanged ? { ...element, rows } as BodyElement : element;
    }
    if (element.type !== 'sectionBreak') return element;
    const elementIndex = path.at(-1) ?? 0;
    let sectionChanged = false;
    const normalizeParts = (
      parts: HeadersFooters | undefined,
      partStory: 'header' | 'footer',
    ): HeadersFooters | undefined => {
      if (!parts) return parts;
      let result = parts;
      for (const kind of ['default', 'first', 'even'] as const) {
        const part = parts[kind];
        if (!part) continue;
        const nextBody = normalizeBody(part.body, partStory, `section:${elementIndex}:${kind}`);
        if (nextBody === part.body) continue;
        if (result === parts) result = { ...parts };
        result[kind] = { ...part, body: nextBody };
        sectionChanged = true;
      }
      return result;
    };
    const headers = normalizeParts(element.headers, 'header');
    const footers = normalizeParts(element.footers, 'footer');
    if (!sectionChanged) return element;
    return { ...element, headers, footers };
  };
  const normalizeBody = (
    body: BodyElement[],
    story: SourceRef['story'],
    storyInstance: string,
    prefix: number[] = [],
  ): BodyElement[] => {
    if (consumeOwned) {
      for (let elementIndex = 0; elementIndex < body.length; elementIndex += 1) {
        body[elementIndex] = normalizeElement(
          body[elementIndex]!,
          story,
          storyInstance,
          [...prefix, elementIndex],
        );
      }
      return body;
    }
    let changed = false;
    const normalized = body.map((element, elementIndex): BodyElement => {
      const next = normalizeElement(
        element,
        story,
        storyInstance,
        [...prefix, elementIndex],
      );
      if (next !== element) changed = true;
      return next;
    });
    return changed ? normalized : body;
  };
  const normalizeParts = (
    parts: HeadersFooters | undefined,
    story: 'header' | 'footer',
  ): HeadersFooters => {
    if (!parts) return { default: null, first: null, even: null };
    let result = parts;
    for (const kind of ['default', 'first', 'even'] as const) {
      const part = parts[kind];
      if (!part) continue;
      const body = normalizeBody(part.body, story, kind);
      if (body === part.body) continue;
      if (result === parts) result = { ...parts };
      result[kind] = { ...part, body };
    }
    return result;
  };
  const body = normalizeBody(doc.body, 'body', 'body');
  const headers = normalizeParts(doc.headers, 'header');
  const footers = normalizeParts(doc.footers, 'footer');
  const normalizeNotes = <T extends { id: string; content: BodyElement[] }>(
    notes: T[] | undefined,
    story: 'footnote' | 'endnote',
  ): T[] | undefined => {
    if (!notes) return notes;
    if (consumeOwned) {
      for (const note of notes) note.content = normalizeBody(note.content, story, note.id);
      return notes;
    }
    let changed = false;
    const normalized = notes.map((note) => {
      const content = normalizeBody(note.content, story, note.id);
      if (content === note.content) return note;
      changed = true;
      return { ...note, content };
    });
    return changed ? normalized : notes;
  };
  const footnotes = normalizeNotes(doc.footnotes, 'footnote');
  const endnotes = normalizeNotes(doc.endnotes, 'endnote');
  const changed = body !== doc.body || headers !== doc.headers || footers !== doc.footers
    || footnotes !== doc.footnotes || endnotes !== doc.endnotes;
  const document = (changed
    ? { ...doc, body, headers, footers, footnotes, endnotes }
    : doc) as InternalDocxDocumentModel;
  const sectionPlacementInputs = projectSectionPlacementInputs(document);
  sectionPlacementInputsByDocument.set(document, sectionPlacementInputs);
  setBodySectionPlacementInputs(document.body, document.section, sectionPlacementInputs);
  let bodyLayoutAcquisition: BodyLayoutAcquisitionInput | undefined;
  const acquiredBodyLayout = (): BodyLayoutAcquisitionInput => (
    bodyLayoutAcquisition ??= bodyLayoutAcquisitionInput(document)
  );
  const acquisitionInputs = documentScopedBodyAcquisitionInputProjections();
  return Object.freeze({
    document,
    mathOccurrences: Object.freeze(occurrences),
    fontFamilyCharsets: Object.freeze({ ...(internalDocumentModel(document).fontFamilyCharsets ?? {}) }),
    get bodyLayoutInput() {
      return projectBodyLayoutInput(acquiredBodyLayout());
    },
    bodyModelGateway: Object.freeze({
      acquisitionInputs,
      get bodySectionIndex() {
        return acquiredBodyLayout().sectionIndex;
      },
      effectiveTablePositioning,
      publicAnchorBridge,
    }),
  });
}

/** Return the stable public model while registering parser-private acquisition
 * sidecars on the returned object identities for same-realm rendering. */
export function normalizeDocxDocumentModel(doc: DocxDocumentModel): DocxDocumentModel {
  return normalizeInternalDocumentModel(doc).document;
}

export function internalFieldRun(run: FieldRun): InternalFieldRun {
  return run as InternalFieldRun;
}

export function internalTextRun(run: DocxTextRun): InternalTextRun {
  return run as InternalTextRun;
}

export function internalNumberingInfo(numbering: NumberingInfo): InternalNumberingInfo {
  return numbering as InternalNumberingInfo;
}

export function internalParagraph(paragraph: ParagraphLayoutSource): InternalDocParagraph {
  return paragraph as unknown as InternalDocParagraph;
}

export function internalDocumentModel(doc: DocxDocumentModel): InternalDocxDocumentModel {
  return doc as InternalDocxDocumentModel;
}

/** One explicit parser-owned projection record composed by the renderer and
 * consumed by parser-independent retained acquisition. Function identities are
 * preserved: this record introduces no compatibility wrappers. */
export const bodyAcquisitionInputProjections = Object.freeze({
  numberingMarkerShapeInput,
  paragraphMarkShapeInput,
  tableFormatInput,
  tableColumnLayoutInput,
  tableParticipatesInOrdinaryFlow,
  paragraphAcquisitionInput,
}) satisfies BodyAcquisitionInputProjections;

/**
 * Parser-fact projections are independent of pagination location, continuation,
 * wrap exclusions, and every other acquisition-time input. Keep their cache on
 * the normalized document gateway so repeated whole-document convergence passes
 * reuse the same immutable snapshot without sharing facts between documents.
 *
 * Paragraph object identity prevents equal-looking hand-built paragraphs from
 * aliasing. The canonical source key keeps source-scoped anchor, field, math,
 * and text-box identities distinct while accepting equivalent SourceRef values.
 */
function documentScopedBodyAcquisitionInputProjections(): BodyAcquisitionInputProjections {
  const paragraphInputs = new WeakMap<
    object,
    Map<string, ParagraphAcquisitionInput>
  >();
  const cachedParagraphAcquisitionInput: BodyAcquisitionInputProjections['paragraphAcquisitionInput'] = (
    paragraph,
    source,
  ) => {
    let bySource = paragraphInputs.get(paragraph);
    if (!bySource) {
      bySource = new Map();
      paragraphInputs.set(paragraph, bySource);
    }
    const key = sourceKey(source);
    const retained = bySource.get(key);
    if (retained) return retained;
    const projected = paragraphAcquisitionInput(paragraph, source);
    bySource.set(key, projected);
    return projected;
  };
  return Object.freeze({
    ...bodyAcquisitionInputProjections,
    paragraphAcquisitionInput: cachedParagraphAcquisitionInput,
  });
}
