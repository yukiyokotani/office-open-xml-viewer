import type {
  TableColumnLayoutInput,
  TableFormatInput,
  TablePreferredWidthConstraint,
} from './types.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import {
  divideExactLengthKey,
  exactLengthKeyFromDecimal,
  exactLengthKeyFromNumber,
  exactLengthKeyToNumber,
  multiplyExactLengthKeys,
} from './exact-length.js';
import type { DocTable, FramePr } from '../types.js';
import type { DeepReadonly } from './types.js';

export type TableLayoutSource = DeepReadonly<DocTable>;

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
  readonly cellFrame?: DeepReadonly<FramePr> | null;
  readonly ordinaryFlow: boolean;
  readonly logicalSequenceId?: string | null;
  readonly logicalRowOffset?: number;
  readonly logicalTotalRows?: number;
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
  readonly borders: import('../types.js').TableBorders | null;
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

export interface TableAcquisitionInput {
  readonly table: TableLayoutAcquisitionWire | null;
  readonly rows: readonly Readonly<{
    readonly row: TableRowLayoutAcquisitionWire | null;
    readonly cells: readonly (TableCellLayoutAcquisitionWire | null)[];
  }>[];
}

/** Frozen parser-boundary facts consumed by table layout. No parser object,
 * callback, or identity-only sidecar is retained in this contract. */
export interface TableSourceSemanticInput {
  readonly colWidths: readonly number[];
  readonly layout: string | null;
  readonly widthPt: number | null;
  readonly widthPct: number | null;
  readonly rows: readonly Readonly<{
    gridBefore: number;
    gridAfter: number;
    cells: readonly Readonly<{
      colSpan: number;
      widthPt: number | null;
      widthPct: number | null;
      noWrap?: boolean;
    }>[];
  }>[];
}

export interface TableSourceAcquisitionInput {
  /** Only public fields consumed by column projection. Cell contents and
   * nested tables belong to the SourceRef repository and are never duplicated
   * into this fact. */
  readonly semantic: TableSourceSemanticInput;
  readonly lexical: TableAcquisitionInput;
  readonly format: TableFormatInput;
}

export interface CellIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
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

export function effectiveTableWidthKind(width: TableLexicalWidth): string {
  // ECMA-376 §17.4.87 makes the measurement syntax authoritative when it
  // contradicts @type.
  return width.value?.trim().endsWith('%') ? 'pct' : (width.kind ?? 'dxa');
}

export function tableWidthConstraintFromLexical(
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

export function tableDxaPtFromLexical(
  width: TableLexicalWidth | null | undefined,
): number | null {
  const constraint = tableWidthConstraintFromLexical(width);
  return constraint?.kind === 'dxa' ? constraint.value : null;
}

/** WORD_TABLE_CELL_ZERO_NIL_WIDTH_AUTO (table-compatibility.ts): lexical
 * tcW zero and nil mean auto. Project this once before intrinsic measurement
 * and carry the same preference into the solver; raw widthPt=0 is not dxa
 * protection. Without lexical facts, public zero widths have the same meaning. */
export function projectEffectiveCellPreferredWidth(
  cell: Readonly<{ widthPt: number | null; widthPct?: number | null }>,
  lexicalWidth?: TableWidthAcquisitionWire | null,
): TablePreferredWidthConstraint | null {
  const constraint = lexicalWidth
    ? tableWidthConstraintFromLexical(lexicalWidth)
    : cell.widthPt != null ? { kind: 'dxa' as const, value: cell.widthPt }
    : cell.widthPct != null ? { kind: 'pct' as const, value: cell.widthPct / 5000 }
    : null;
  return constraint?.value === 0 ? null : constraint;
}

/** WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE: an authored first-row tblPrEx/tblW governs the
 * whole table, including auto/nil/zero clearing the body's preference. Both
 * occurrence ceiling selection and column projection use this resolver. */
export function projectEffectiveTablePreferredWidthPt(
  input: TableSourceAcquisitionInput,
  availableWidthPt: number,
): number | null {
  const exception = input.format.firstRowException?.preferredWidth ?? null;
  if (input.format.firstRowException?.preferredWidthAuthored) {
    if (exception?.kind === 'dxa') return exception.value > 0 ? exception.value : null;
    if (exception?.kind === 'pct') {
      return exception.value > 0 ? exception.value * availableWidthPt : null;
    }
    return null;
  }
  const lexical = tableWidthConstraintFromLexical(input.lexical.table?.preferredWidth);
  if (lexical?.kind === 'dxa') return lexical.value > 0 ? lexical.value : null;
  if (lexical?.kind === 'pct') return lexical.value > 0 ? lexical.value * availableWidthPt : null;
  if (input.semantic.widthPt != null && input.semantic.widthPt > 0) return input.semantic.widthPt;
  if (input.semantic.widthPct != null && input.semantic.widthPct > 0) {
    return input.semantic.widthPct / 5000 * availableWidthPt;
  }
  return null;
}

const tableGridPointFactors: Readonly<Record<string, string>> = Object.freeze({
  pt: '1/1', in: '72/1', cm: '3600/127', mm: '360/127', pc: '12/1', pi: '12/1',
});
const xsdUnsignedLongMaximum = '18446744073709551615';

function xsdUnsignedLongLexical(value: string): string | null {
  const collapsed = value.replace(/[\u0009\u000a\u000d\u0020]+/g, ' ').replace(/^ | $/g, '');
  const match = /^([+-]?)([0-9]+)$/.exec(collapsed);
  if (!match) return null;
  const [, sign, authoredDigits] = match;
  if (sign === '-' && /[1-9]/.test(authoredDigits!)) return null;
  const digits = authoredDigits!.replace(/^0+/, '') || '0';
  if (digits.length > xsdUnsignedLongMaximum.length
    || (digits.length === xsdUnsignedLongMaximum.length && digits > xsdUnsignedLongMaximum)) {
    return null;
  }
  return collapsed;
}

interface GridTrack { readonly key: string | null; readonly widthPt: number }
const definitionalZeroTrack: GridTrack = { key: '0/1', widthPt: 0 };

function exactTrack(key: string): GridTrack {
  const widthPt = exactLengthKeyToNumber(key);
  return Number.isFinite(widthPt) ? { key, widthPt } : definitionalZeroTrack;
}

function degradedTrack(magnitude: string, factor: string): GridTrack {
  const magnitudeF64 = Number(magnitude);
  if (!Number.isFinite(magnitudeF64)) return definitionalZeroTrack;
  const exactMagnitude = exactLengthKeyFromNumber(magnitudeF64);
  const widthPt = exactMagnitude === null
    ? 0
    : exactLengthKeyToNumber(multiplyExactLengthKeys(exactMagnitude, factor));
  return Number.isFinite(widthPt) ? { key: null, widthPt } : definitionalZeroTrack;
}

function tableGridTrack(value: string | null | undefined): GridTrack {
  if (value == null) return definitionalZeroTrack;
  const unsignedLong = xsdUnsignedLongLexical(value);
  if (unsignedLong !== null) {
    const twips = exactLengthKeyFromDecimal(unsignedLong);
    return twips === null ? definitionalZeroTrack : exactTrack(divideExactLengthKey(twips, 20n));
  }
  const universal = /^([0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$/.exec(value);
  if (!universal) return definitionalZeroTrack;
  const factor = tableGridPointFactors[universal[2]!]!;
  const magnitude = exactLengthKeyFromDecimal(universal[1]!);
  return magnitude === null
    ? degradedTrack(universal[1]!, factor)
    : exactTrack(multiplyExactLengthKeys(magnitude, factor));
}

function tableGridLayout(input: TableSourceAcquisitionInput): Readonly<{
  widthsPt: readonly number[];
  widthKeys: readonly (string | null)[];
}> {
  const grid = input.lexical.table?.grid;
  if (!grid) {
    const tracks = input.semantic.colWidths.map((width): GridTrack => (
      Number.isFinite(width) && width >= 0
        ? { widthPt: width, key: exactLengthKeyFromNumber(width) ?? '0/1' }
        : definitionalZeroTrack
    ));
    return {
      widthsPt: tracks.map((track) => track.widthPt),
      widthKeys: tracks.map((track) => track.key),
    };
  }
  const count = Math.max(grid.requiredColumnCount, grid.columns.length);
  const tracks = Array.from({ length: count }, (_unused, column) => (
    tableGridTrack(grid.columns[column]?.width ?? null)
  ));
  return {
    widthsPt: tracks.map((track) => track.widthPt),
    widthKeys: tracks.map((track) => track.key),
  };
}

function skippedTableWidthConstraint(
  width: TableLexicalWidth | null | undefined,
  availableWidthPt: number,
): TablePreferredWidthConstraint | null {
  const constraint = tableWidthConstraintFromLexical(width);
  if (constraint?.kind !== 'pct') return constraint;
  return { kind: 'dxa', value: Math.max(0, constraint.value) * Math.max(0, availableWidthPt) };
}

/** Pure projection from frozen source facts into the §17.18.87 solver input. */
export function projectTableColumnLayoutInput(
  input: TableSourceAcquisitionInput,
  availableWidthPt: number,
  intrinsicWidths: (
    rowIndex: number, cellIndex: number, preferredWidth: TablePreferredWidthConstraint | null,
  ) => CellIntrinsicWidths,
  maximumWidthPt: number | null = availableWidthPt,
): TableColumnLayoutInput {
  const table = input.semantic;
  const { widthsPt: gridWidthsPt, widthKeys: gridWidthKeys } = tableGridLayout(input);
  const layoutKind = input.format.firstRowException?.layout === 'fixed'
    ? 'fixed'
    : (input.lexical.table?.layout?.kind ?? table.layout);
  const authoredGridCount = input.lexical.table?.grid.authored
    ? input.lexical.table.grid.columns.length
    : null;
  const normalizedBeforeSpans = table.rows.map((row) => {
    const requested = Math.max(0, row.gridBefore ?? 0);
    return authoredGridCount !== null && requested > authoredGridCount ? 0 : requested;
  });
  const contentGridCount = Math.max(
    authoredGridCount ?? 0,
    input.lexical.table?.grid.requiredColumnCount ?? 0,
    ...table.rows.map((row, rowIndex) => (
      (normalizedBeforeSpans[rowIndex] ?? 0)
      + row.cells.reduce((total, cell) => total + Math.max(1, cell.colSpan), 0)
    )),
  );
  return {
    layout: layoutKind === 'fixed' ? 'fixed' : 'autofit',
    availableWidthPt: maximumWidthPt === null ? null : Math.max(0, maximumWidthPt),
    gridWidthsPt,
    gridWidthKeys,
    tablePreferredWidthPt: projectEffectiveTablePreferredWidthPt(input, availableWidthPt),
    rows: table.rows.map((row, rowIndex) => {
      const rowInput = input.lexical.rows[rowIndex];
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
          const preferredWidth = projectEffectiveCellPreferredWidth(cell, wire?.preferredWidth);
          const intrinsic = layoutKind === 'fixed'
            ? { minWidthPt: 0, maxWidthPt: 0 }
            : intrinsicWidths(rowIndex, cellIndex, preferredWidth);
          const spacingInsets = tableCellHorizontalSpacingInsets(
            input.format.rows[rowIndex]?.cellSpacingPt ?? 0,
            columnStart,
            span,
            gridWidthsPt.length,
          );
          const horizontalSpacingPt = spacingInsets.startPt + spacingInsets.endPt;
          const margins = input.format.rows[rowIndex]?.cells[cellIndex]?.marginsPt;
          const horizontalMarginsPt = Math.max(0, margins?.left ?? 0)
            + Math.max(0, margins?.right ?? 0);
          const result = {
            columnStart,
            columnSpan: span,
            preferredWidth,
            ...(cell.noWrap ? { noWrap: true } : {}),
            ...(horizontalMarginsPt > 0 ? { horizontalMarginsPt } : {}),
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
