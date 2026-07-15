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
  TableBorders,
  TextPath,
  TblpPr,
} from './types.js';
import type {
  NumberingMarkerShapeInput,
  FloatingTablePositionInput,
  SourceRef,
  TableColumnLayoutInput,
  TableFormatInput,
  TablePreferredWidthConstraint,
  TableRowExceptionInput,
  TableRowHeightInput,
  VmlTextPathAcquisitionInput,
} from './layout/types.js';
import type { MathOccurrence } from './layout/resources.js';
import { anchorOccurrenceKey, mathResourceKey } from './layout/source-key.js';
import type { TextFontSlotPresence, TextFontSlots } from './layout/text.js';
import type { ParagraphAcquisitionInput, ParagraphAcquisitionRun } from './layout/text.js';
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
  type NormalizedTextBoxParagraphInput,
} from './layout/textbox-input.js';
import { tableCellHorizontalSpacingInsets } from './layout/table-columns.js';
import {
  sectionGeometry,
  type BodySectionIndexInput,
  type BodySectionOccurrence,
} from './layout/context.js';

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
}

interface InternalSectionPlacementWire {
  readonly sectionId: string;
  readonly vAlign?: string | null;
  readonly lineNumbering?: LineNumbering | null;
}

type InternalSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement?: InternalSectionPlacementWire;
};

/** Immutable parser-private section placement input. It is intentionally kept
 * outside BodyElement/DocxDocumentModel's stable declaration surface. */
export interface SectionPlacementInput {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
}

interface DocumentSectionPlacementInputs {
  readonly endingSections: ReadonlyMap<number, SectionPlacementInput>;
  readonly finalSection: SectionPlacementInput;
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

/** Lexical CT_TblWidth facts. Element absence is represented by the owning
 * nullable field; null attributes retain malformed/partial authored OOXML. */
export interface TableWidthAcquisitionWire {
  readonly kind: string | null;
  readonly value: string | null;
}

export interface TableLayoutKindAcquisitionWire {
  readonly kind: string | null;
}

export interface TableMarginAcquisitionWire {
  readonly top?: TableWidthAcquisitionWire | null;
  readonly bottom?: TableWidthAcquisitionWire | null;
  readonly start?: TableWidthAcquisitionWire | null;
  readonly end?: TableWidthAcquisitionWire | null;
  readonly left?: TableWidthAcquisitionWire | null;
  readonly right?: TableWidthAcquisitionWire | null;
}

export interface TableLayoutAcquisitionWire {
  readonly effectiveStyleId: string | null;
  /** Word-effective flow participation after [MS-OI29500] 2.1.162(b-c), not
   * the lexical presence/absence of the public `tblpPr` compatibility field. */
  readonly ordinaryFlow: boolean;
  readonly grid: {
    readonly authored: boolean;
    readonly columns: readonly { readonly width: string | null }[];
    readonly requiredColumnCount: number;
  };
  readonly preferredWidth: TableWidthAcquisitionWire | null;
  readonly layout: TableLayoutKindAcquisitionWire | null;
  readonly cellSpacing: TableWidthAcquisitionWire | null;
  readonly cellMargins?: TableMarginAcquisitionWire | null;
}

export interface TableRowHeightAcquisitionWire {
  readonly value: string | null;
  readonly rule: string;
  readonly ruleAuthored: boolean;
}

export interface TablePropertyExceptionAcquisitionWire {
  readonly preferredWidth: TableWidthAcquisitionWire | null;
  readonly layout: TableLayoutKindAcquisitionWire | null;
  readonly justification: string | null;
  readonly indent: TableWidthAcquisitionWire | null;
  readonly borders: TableBorders | null;
  readonly cellMargins: TableMarginAcquisitionWire | null;
  readonly cellSpacing: TableWidthAcquisitionWire | null;
}

export interface TableRowLayoutAcquisitionWire {
  readonly height: TableRowHeightAcquisitionWire | null;
  readonly justification: string | null;
  readonly beforeWidth: TableWidthAcquisitionWire | null;
  readonly afterWidth: TableWidthAcquisitionWire | null;
  readonly cellSpacing: TableWidthAcquisitionWire | null;
  readonly styleCellSpacing?: TableWidthAcquisitionWire | null;
  readonly styleCellMargins?: TableMarginAcquisitionWire | null;
  readonly exception: TablePropertyExceptionAcquisitionWire | null;
}

export interface TableCellLayoutAcquisitionWire {
  readonly preferredWidth: TableWidthAcquisitionWire | null;
  readonly margins: TableMarginAcquisitionWire | null;
}

interface InternalTable extends DocTable {
  readonly __tableLayout?: TableLayoutAcquisitionWire;
}

type InternalTableRow = DocTable['rows'][number] & {
  readonly __tableRowLayout?: TableRowLayoutAcquisitionWire;
};

type InternalTableCell = DocTable['rows'][number]['cells'][number] & {
  readonly __tableCellLayout?: TableCellLayoutAcquisitionWire;
};

export interface TableAcquisitionInput {
  readonly table: TableLayoutAcquisitionWire | null;
  readonly rows: readonly {
    readonly row: TableRowLayoutAcquisitionWire | null;
    readonly cells: readonly (TableCellLayoutAcquisitionWire | null)[];
  }[];
}

const tableAcquisitionInputs = new WeakMap<object, TableAcquisitionInput>();
const tableFormatInputs = new WeakMap<object, TableFormatInput>();

/** Snapshot serde-only table facts once at the parser/model boundary. Layout
 * receives only clone-safe immutable data, while hand-built public `DocTable`
 * values remain supported through aligned null entries and their public fields. */
export function tableAcquisitionInput(table: Readonly<DocTable>): TableAcquisitionInput {
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

/** Word-effective flow classification acquired before tblpPr defaults erase the
 * distinction between authored positioning and an ignored positioning payload. */
export function tableParticipatesInOrdinaryFlow(table: Readonly<DocTable>): boolean {
  return tableFormatInput(table).ordinaryFlow;
}

/** Positioning payload only when Word treats the authored tblpPr as effective. */
export function effectiveTablePositioning(table: Readonly<DocTable>): TblpPr | null {
  return tableFormatInput(table).positioning === null ? null : (table.tblpPr ?? null);
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

type TableLexicalWidth = Readonly<{
  kind: string | null;
  value: string | null;
}>;

function finiteTableLexicalNumber(value: string | null, allowPercent: boolean): number | null {
  if (value === null) return null;
  const lexical = value.trim();
  const numeric = allowPercent && lexical.endsWith('%') ? lexical.slice(0, -1) : lexical;
  if (numeric.length === 0) return null;
  const parsed = Number(numeric);
  return Number.isFinite(parsed) ? parsed : null;
}

function effectiveTableWidthKind(width: TableLexicalWidth): string {
  // ECMA-376 §17.4.87 makes the measurement syntax authoritative when it
  // contradicts @type. Keeping this resolution here prevents each consumer
  // from assigning different semantics to the same CT_TblWidth value.
  return width.value?.trim().endsWith('%') ? 'pct' : (width.kind ?? 'dxa');
}

function tableWidthConstraintFromLexical(
  width: TableLexicalWidth | null | undefined,
): TablePreferredWidthConstraint | null {
  if (!width) return null;
  const lexicalValue = width.value?.trim() ?? '';
  // §17.4.87: omitted type defaults to dxa and omitted w defaults to zero.
  const kind = effectiveTableWidthKind(width);
  if (kind === 'dxa') {
    const value = finiteTableLexicalNumber(width.value ?? '0', false);
    return value === null ? null : { kind: 'dxa', value: value / 20 };
  }
  if (kind !== 'pct') return null;
  const value = finiteTableLexicalNumber(width.value ?? '0', true);
  if (value === null) return null;
  return {
    kind: 'pct',
    value: lexicalValue.endsWith('%') ? value / 100 : value / 5000,
  };
}

function tableDxaPtFromLexical(width: TableLexicalWidth | null | undefined): number | null {
  const constraint = tableWidthConstraintFromLexical(width);
  return constraint?.kind === 'dxa' ? constraint.value : null;
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
  // ECMA-376 §17.4.80 defaults omitted hRule to auto, but Word intentionally
  // differs: [MS-OI29500] 2.1.180 treats an omitted hRule as atLeast. Authored
  // presence is therefore semantic input, not a parser implementation detail.
  return {
    rule: height.ruleAuthored ? normalizedTableHeightRule(height.rule) : 'atLeast',
    valuePt: tableTwipsValuePt(height.value),
  };
}

function publicTableRowHeight(row: Readonly<DocTable['rows'][number]>): TableRowHeightInput | null {
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
    // Word resolves authored pct/auto spacing to zero at that precedence scope
    // instead of exposing a lower scope ([MS-OI29500] 2.1.152–154).
    const kind = effectiveTableWidthKind(width);
    if (kind === 'pct' || kind === 'auto' || kind === 'nil') return 0;
    const valuePt = tableDxaPtFromLexical(width);
    if (valuePt !== null) return valuePt;
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
  if (kind === 'dxa') return tableTwipsValuePt(width.value ?? '0');
  // Individual-cell margin exceptions ignore pct/auto/nil. The default
  // leading/trailing table margin differs in Word: pct/auto becomes zero
  // ([MS-OI29500] 2.1.125/.146), while nil keeps ST_TblWidth's zero meaning.
  if (scope === 'cell' || scope === 'exception') return null;
  if (edge === 'start' || edge === 'end') {
    if (kind === 'pct' || kind === 'auto' || kind === 'nil') return 0;
  }
  // Word ignores nil top/bottom margin elements ([MS-OI29500] 2.1.116/.177,
  // also referenced by the corresponding default-margin sections).
  return null;
}

function effectiveTableCellMargins(
  table: Readonly<DocTable>,
  cell: Readonly<DocTable['rows'][number]['cells'][number]>,
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
export function tableFormatInput(table: Readonly<DocTable>): TableFormatInput {
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
    positioning: ordinaryFlow || table.tblpPr == null
      ? null
      : floatingTablePositionInput(table.tblpPr),
    rows,
    // Word applies selected first-row tblPrEx values table-wide
    // ([MS-OI29500] 2.1.156/.158/.167).
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
    return Object.freeze({
      element,
      table: Object.freeze({
        effectiveStyleId: format.effectiveStyleId,
        ordinaryFlow: format.ordinaryFlow,
      }),
    });
  }));
}

export interface CellIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

function publicTableCellConstraint(
  cell: DocTable['rows'][number]['cells'][number],
): TablePreferredWidthConstraint | null {
  if (cell.widthPt != null) return { kind: 'dxa', value: cell.widthPt };
  if (cell.widthPct != null) return { kind: 'pct', value: cell.widthPct / 5000 };
  return null;
}

function tablePreferredWidthPt(
  table: DocTable,
  input: TableAcquisitionInput,
  availableWidthPt: number,
  firstRowException: TableRowExceptionInput | null,
): number | null {
  const exception = firstRowException?.preferredWidth ?? null;
  if (firstRowException?.preferredWidthAuthored) {
    // [MS-OI29500] 2.1.167 applies a first-row tblPrEx/tblW to the whole
    // table. Authored auto/nil/zero values therefore shadow the parent tblW
    // without becoming an invented physical length.
    if (exception?.kind === 'dxa') return exception.value > 0 ? exception.value : null;
    if (exception?.kind === 'pct') {
      return exception.value > 0 ? exception.value * availableWidthPt : null;
    }
    return null;
  }
  const lexical = tableWidthConstraintFromLexical(input.table?.preferredWidth);
  if (lexical?.kind === 'dxa') return lexical.value > 0 ? lexical.value : null;
  if (lexical?.kind === 'pct') return lexical.value > 0 ? lexical.value * availableWidthPt : null;
  if (table.widthPt != null && table.widthPt > 0) return table.widthPt;
  if (table.widthPct != null && table.widthPct > 0) return table.widthPct / 5000 * availableWidthPt;
  return null;
}

function tableGridWidthsPt(table: DocTable, input: TableAcquisitionInput): number[] {
  const grid = input.table?.grid;
  if (!grid) return table.colWidths.map((width) => Math.max(0, width));
  const count = Math.max(grid.requiredColumnCount, grid.columns.length);
  return Array.from({ length: count }, (_unused, column) => {
    const points = tableTwipsValuePt(grid.columns[column]?.width ?? null);
    return points === null ? 0 : Math.max(0, points);
  });
}

function skippedTableWidthConstraint(
  width: TableLexicalWidth | null | undefined,
  availableWidthPt: number,
): TablePreferredWidthConstraint | null {
  const constraint = tableWidthConstraintFromLexical(width);
  if (constraint?.kind !== 'pct') return constraint;
  return { kind: 'dxa', value: Math.max(0, constraint.value) * Math.max(0, availableWidthPt) };
}

/** Project normalized parser/model facts into the pure §17.18.87 solver contract. */
export function tableColumnLayoutInput(
  table: Readonly<DocTable>,
  availableWidthPt: number,
  intrinsicWidths: (cell: Readonly<DocTable['rows'][number]['cells'][number]>) => CellIntrinsicWidths,
  maximumWidthPt: number = availableWidthPt,
): TableColumnLayoutInput {
  const acquisition = tableAcquisitionInput(table);
  const format = tableFormatInput(table);
  const gridWidthsPt = tableGridWidthsPt(table as DocTable, acquisition);
  const layoutKind = format.firstRowException?.layout === 'fixed'
    ? 'fixed'
    : (acquisition.table?.layout?.kind ?? table.layout);
  const authoredGridCount = acquisition.table?.grid.authored
    ? acquisition.table.grid.columns.length
    : null;
  const normalizedBeforeSpans = table.rows.map((row) => {
    const requested = Math.max(0, row.gridBefore ?? 0);
    return authoredGridCount !== null && requested > authoredGridCount ? 0 : requested;
  });
  const contentGridCount = Math.max(
    authoredGridCount ?? 0,
    acquisition.table?.grid.requiredColumnCount ?? 0,
    ...table.rows.map((row, rowIndex) => (
      (normalizedBeforeSpans[rowIndex] ?? 0)
      + row.cells.reduce((total, cell) => total + Math.max(1, cell.colSpan), 0)
    )),
  );
  return {
    layout: layoutKind === 'fixed' ? 'fixed' : 'autofit',
    availableWidthPt: Math.max(0, maximumWidthPt),
    gridWidthsPt,
    tablePreferredWidthPt: tablePreferredWidthPt(
      table as DocTable,
      acquisition,
      availableWidthPt,
      format.firstRowException,
    ),
    rows: table.rows.map((row, rowIndex) => {
      const rowInput = acquisition.rows[rowIndex];
      const beforeSpan = normalizedBeforeSpans[rowIndex] ?? 0;
      const requestedAfterSpan = Math.max(0, row.gridAfter ?? 0);
      const occupiedColumns = beforeSpan
        + row.cells.reduce((total, cell) => total + Math.max(1, cell.colSpan), 0);
      const afterSpan = authoredGridCount !== null
        && occupiedColumns + requestedAfterSpan > contentGridCount
        ? 0
        : requestedAfterSpan;
      let columnStart = beforeSpan;
      return {
        before: beforeSpan > 0 ? {
          columnSpan: beforeSpan,
          preferredWidth: skippedTableWidthConstraint(rowInput?.row?.beforeWidth, availableWidthPt),
        } : null,
        after: afterSpan > 0 ? {
          columnSpan: afterSpan,
          preferredWidth: skippedTableWidthConstraint(rowInput?.row?.afterWidth, availableWidthPt),
        } : null,
        cells: row.cells.map((cell, cellIndex) => {
          const wire = rowInput?.cells[cellIndex] ?? null;
          const span = Math.max(1, cell.colSpan);
          const intrinsic = layoutKind === 'fixed'
            ? { minWidthPt: 0, maxWidthPt: 0 }
            : intrinsicWidths(cell);
          const spacingInsets = tableCellHorizontalSpacingInsets(
            format.rows[rowIndex]?.cellSpacingPt ?? 0,
            columnStart,
            span,
            gridWidthsPt.length,
          );
          const horizontalSpacingPt = spacingInsets.startPt + spacingInsets.endPt;
          const result = {
            columnStart,
            columnSpan: span,
            preferredWidth: tableWidthConstraintFromLexical(wire?.preferredWidth)
              ?? publicTableCellConstraint(cell),
            minContentWidthPt: Math.max(0, intrinsic.minWidthPt) + horizontalSpacingPt,
            maxContentWidthPt:
              Math.max(intrinsic.minWidthPt, intrinsic.maxWidthPt) + horizontalSpacingPt,
          };
          columnStart += span;
          return result;
        }),
      };
    }),
  };
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
      vAlign: wire?.vAlign ?? null,
      lineNumbering: wire?.lineNumbering ?? null,
    }, 'DOCX ending-section placement input'));
    ordinal += 1;
  });
  return Object.freeze({
    endingSections,
    finalSection: snapshotPlainData({
      sectionId: `section:${ordinal}`,
      // Resource-only entry points (for example image preloading) historically
      // accept a partial document projection with no section. Section placement
      // is irrelevant there, so preserve that compatibility with neutral facts.
      vAlign: doc.section?.vAlign ?? null,
      lineNumbering: doc.section?.lineNumbering ?? null,
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
  const inheritedGeometry = sectionGeometry(doc.section);
  const occurrences: BodySectionOccurrence[] = [];
  let startBodyIndex = 0;

  doc.body.forEach((element, bodyIndex) => {
    if (element.type !== 'sectionBreak') return;
    const placement = sectionPlacementInputFromBody(doc.body, doc.section, bodyIndex);
    const ordinal = occurrences.length;
    occurrences.push({
      sectionOccurrenceId: placement.sectionId,
      ordinal,
      startBodyIndex,
      endBodyIndex: bodyIndex,
      markerBodyIndex: bodyIndex,
      final: false,
      startType: element.kind ?? 'nextPage',
      columns: element.columns ?? null,
      // A paragraph-owned sectPr can omit pgSz/pgMar; inherit the final
      // physical page box rather than manufacturing defaults after acquisition.
      geometry: element.geom ?? inheritedGeometry,
      textDirection: element.textDirection ?? null,
      pageNumType: element.pageNumType ?? null,
      headers: element.headers ?? EMPTY_SECTION_HEADERS_FOOTERS,
      footers: element.footers ?? EMPTY_SECTION_HEADERS_FOOTERS,
      titlePage: element.titlePage ?? false,
      vAlign: placement.vAlign,
      lineNumbering: placement.lineNumbering,
      placement,
    });
    startBodyIndex = bodyIndex + 1;
  });

  const placement = sectionPlacementInputFromBody(doc.body, doc.section, doc.body.length);
  occurrences.push({
    sectionOccurrenceId: placement.sectionId,
    ordinal: occurrences.length,
    startBodyIndex,
    endBodyIndex: doc.body.length - 1,
    markerBodyIndex: null,
    final: true,
    startType: doc.section.sectionStart ?? 'nextPage',
    columns: doc.section.columns ?? null,
    geometry: inheritedGeometry,
    textDirection: doc.section.textDirection ?? null,
    pageNumType: doc.section.pageNumType ?? null,
    headers: doc.headers ?? EMPTY_SECTION_HEADERS_FOOTERS,
    footers: doc.footers ?? EMPTY_SECTION_HEADERS_FOOTERS,
    titlePage: doc.section.titlePage,
    vAlign: placement.vAlign,
    lineNumbering: placement.lineNumbering,
    placement,
  });

  return snapshotPlainData({
    bodyLength: doc.body.length,
    occurrences,
  }, 'DOCX body section index input') as BodySectionIndexInput;
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

export interface InternalShapeRun extends ShapeRun {
  textPath?: InternalVmlTextPath | null;
}

export interface NormalizedDocumentInput {
  readonly document: InternalDocxDocumentModel;
  readonly mathOccurrences: readonly MathOccurrence[];
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

/** Project parser-only anchor facts without widening the public run contract.
 * Parser-produced malformed input remains distinguishable from a hand-built
 * public run because required-but-missing values are explicit nulls. */
export function anchorAcquisitionInput(
  run: Readonly<DocRun | ShapeRun | ImageRun | ChartRun>,
): Readonly<AnchorAcquisitionInput> | undefined {
  const wire = (run as Readonly<DocRun & InternalAnchorRunWire>).__anchorAcquisition;
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
): readonly NormalizedTextBoxParagraphInput[] {
  return normalizeTextBoxInput(shape, source, numberingMarkerShapeInput);
}

/** Snapshot private paragraph-mark rPr facts at the parser boundary. Retained
 * line layout receives only this plain immutable service input. */
export function paragraphMarkShapeInput(
  paragraph: DocParagraph,
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
  paragraph: DocParagraph,
  source: SourceRef,
): ParagraphAcquisitionInput {
  // Table pagination may have attached legacy cache stamps containing live font
  // resolver functions. They are renderer state, not parser/model facts, and must
  // never cross the retained acquisition boundary.
  const {
    layoutLines: _layoutLines,
    lineSlice: _lineSlice,
    __paragraphTypographyAcquisition: _privateParagraphTypography,
    ...semanticParagraph
  } = paragraph as DocParagraph & Record<string, unknown> & {
    __paragraphTypographyAcquisition?: InternalParagraphTypographyWire;
  };
  const typographyInput = paragraphTypographyAcquisitionInput(paragraph);
  const snapshot = structuredClone(semanticParagraph) as DocParagraph;
  const runs = snapshot.runs.map((run, runIndex): ParagraphAcquisitionRun => {
    if (run.type === 'math') {
      const runRef: SourceRef = Object.freeze({ ...source, path: Object.freeze([...source.path, runIndex]) });
      const internal = run as Partial<InternalMathRun>;
      return Object.freeze({
        ...run,
        source: internal.source ?? runRef,
        resourceKey: internal.resourceKey ?? mathResourceKey(runRef, run.display ? 'display' : 'inline'),
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
      const originalRun = paragraph.runs[runIndex] as DocRun;
      const localAnchorInput = anchorAcquisitionInput(originalRun);
      const anchorInput = localAnchorInput === undefined
        ? undefined
        : snapshotPlainData({
            ...localAnchorInput,
            occurrenceId: anchorOccurrenceKey(source, localAnchorInput.occurrenceId),
          }, 'DOCX scoped anchor acquisition input');
      const { __anchorAcquisition: _privateAnchor, ...publicRun } = run as typeof run & InternalAnchorRunWire;
      if (run.type !== 'shape') {
        return Object.freeze({
          ...publicRun,
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
      return Object.freeze({
        ...publicRun,
        ...(vmlTextPathInput === undefined ? {} : { vmlTextPathInput }),
        ...(textBoxInput.length === 0 ? {} : { textBoxInput }),
        ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
      }) as ParagraphAcquisitionRun;
    }
    if (run.type === 'text' || run.type === 'field') {
      const originalRun = paragraph.runs[runIndex] as Extract<DocRun, { type: 'text' | 'field' }>;
      const runTypographyInput = runTypographyAcquisitionInput(originalRun);
      const {
        __typographyAcquisition: _privateRunTypography,
        ...publicRun
      } = run as typeof run & { __typographyAcquisition?: InternalRunTypographyWire };
      return Object.freeze({
        ...publicRun,
        ...(runTypographyInput === undefined ? {} : { typographyInput: runTypographyInput }),
      }) as ParagraphAcquisitionRun;
    }
    return Object.freeze({ ...run }) as ParagraphAcquisitionRun;
  });
  return deepFreezePlainData({
    ...snapshot,
    runs: runs as ParagraphAcquisitionRun[],
    numberingMarkerShapeInput: paragraph.numbering
      ? numberingMarkerShapeInput(
          paragraph.numbering,
          paragraph.runs.find(
            (run): run is Extract<DocRun, { type: 'text' | 'field' }> =>
              run.type === 'text' || run.type === 'field',
          )?.fontSize ?? paragraph.defaultFontSize ?? 10,
        )
      : undefined,
    paragraphMarkShapeInput: paragraphMarkShapeInput(paragraph),
    ...(typographyInput === undefined ? {} : { typographyInput }),
  }) as unknown as ParagraphAcquisitionInput;
}

/** Pure structural normalization for stable math addressing. Only ancestry that
 * contains a math run is shallow-cloned; the caller's parser model is untouched. */
export function normalizeInternalDocumentModel(doc: DocxDocumentModel): NormalizedDocumentInput {
  const occurrences: MathOccurrence[] = [];
  const normalizeBody = (
    body: BodyElement[],
    story: SourceRef['story'],
    storyInstance: string,
    prefix: number[] = [],
  ): BodyElement[] => {
    let changed = false;
    const normalized = body.map((element, elementIndex): BodyElement => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        let runsChanged = false;
        const runs = element.runs.map((run, runIndex): DocRun => {
          if (run.type !== 'math') return run;
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
          return Object.freeze({ ...run, source, resourceKey }) as InternalMathRun;
        });
        if (!runsChanged) return element;
        changed = true;
        return { ...element, runs };
      }
      if (element.type === 'table') {
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
        if (!tableChanged) return element;
        changed = true;
        return { ...element, rows } as BodyElement;
      }
      if (element.type !== 'sectionBreak') return element;
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
      changed = true;
      return { ...element, headers, footers };
    });
    return changed ? normalized : body;
  };
  const normalizeParts = (
    parts: HeadersFooters,
    story: 'header' | 'footer',
  ): HeadersFooters => {
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
  return Object.freeze({
    document,
    mathOccurrences: Object.freeze(occurrences),
  });
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

export function internalParagraph(paragraph: DocParagraph): InternalDocParagraph {
  return paragraph as InternalDocParagraph;
}

export function internalDocumentModel(doc: DocxDocumentModel): InternalDocxDocumentModel {
  return doc as InternalDocxDocumentModel;
}
