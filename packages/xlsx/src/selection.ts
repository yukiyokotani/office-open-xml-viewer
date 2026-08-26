import { parseA1 } from './a1.js';
import { MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from './internal/grid-geometry.js';
import type { ViewerCommentThreadContext } from '@silurus/ooxml-core';

export interface CellAddress {
  row: number;
  col: number;
}

export type XlsxSelectionArea =
  | Readonly<{
      kind: 'cells';
      top: number;
      left: number;
      bottom: number;
      right: number;
    }>
  | Readonly<{ kind: 'rows'; firstRow: number; lastRow: number }>
  | Readonly<{ kind: 'columns'; firstColumn: number; lastColumn: number }>
  | Readonly<{ kind: 'sheet' }>;

/**
 * Canonical worksheet selection state.
 *
 * This follows SpreadsheetML's separation of `sqref`, `activeCell`, and
 * `activeCellId` (ECMA-376 §18.3.1.78). `extensionAnchor` is viewer interaction
 * state: Shift/drag extends from it, independently of the active cell.
 */
export interface XlsxSelectionState {
  readonly areas: readonly XlsxSelectionArea[];
  readonly activeAreaIndex: number;
  readonly activeCell: CellAddress;
  readonly extensionAnchor: CellAddress;
}

/** One selection-setting entry point: an A1 area, full canonical state, or null. */
export type XlsxSelectionInput = string | XlsxSelectionState | null;

export interface XlsxSelectionContextOptions {
  /** Maximum populated cells returned. Default 1,000; hard maximum 10,000. */
  readonly maxCells?: number;
  /** Maximum returned UTF-16 code units. Range context defaults to 1 Mi and has
   * an 8 Mi hard cap; element context retains its snapshot budget and has a
   * 65,536-character hard cap. Viewer change notifications request 65,536. */
  readonly maxTextCharacters?: number;
}

export interface XlsxSelectionContextCell {
  readonly address: CellAddress;
  /** Formatted text shown by the Viewer, including number/date formatting. */
  readonly displayText: string;
  readonly valueType: 'empty' | 'text' | 'number' | 'bool' | 'error' | 'shared';
  /** Detached scalar value; rich-text formatting stays outside the AI context. */
  readonly value: string | number | boolean | null;
  readonly formula?: string;
  /** Authored note or threaded comment attached to this selected cell. */
  readonly comment?: ViewerCommentThreadContext;
}

/**
 * Serializable, resource-bounded context for read-only AI/MCP integrations.
 * Selection state remains the UI authority; this snapshot adds displayed cell
 * content without exposing mutable workbook internals.
 */
export interface XlsxRangeSelectionContext {
  readonly format: 'xlsx';
  /** Discriminant shared with DOCX/PPTX selection-context snapshots. */
  readonly kind: 'range';
  readonly sheetIndex: number;
  readonly sheetName: string;
  readonly selection: XlsxSelectionState;
  /** Sum of area sizes before overlap removal; always safe to compute. */
  readonly coordinateCountUpperBound: number;
  /** Populated cells only, in worksheet order. */
  readonly cells: readonly XlsxSelectionContextCell[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('cells' | 'text')[];
  readonly maxCells: number;
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

export interface XlsxElementAnchorMarker {
  /** One-based worksheet row/column containing the DrawingML marker. */
  readonly row: number;
  readonly col: number;
  /** Marker offsets inside the cell, in DrawingML EMU. */
  readonly offsetX: number;
  readonly offsetY: number;
}

/**
 * Detached context for the topmost rendered worksheet object established by a
 * click. The Viewer outlines this focus for clarity, but it is not an editable
 * Excel object selection and intentionally exposes no mutable drawing model.
 */
export interface XlsxElementContext {
  readonly format: 'xlsx';
  readonly kind: 'element';
  readonly sheetIndex: number;
  readonly sheetName: string;
  readonly elementType: 'chart' | 'image' | 'shape';
  /** Index in the matching immutable worksheet collection for this snapshot. */
  readonly elementIndex: number;
  /** Leaf index inside a shape group; present only for `elementType: "shape"`. */
  readonly shapeIndex?: number;
  readonly anchor: Readonly<{
    from: XlsxElementAnchorMarker;
    to: XlsxElementAnchorMarker;
  }>;
  readonly text?: string;
  readonly mimeType?: string;
  readonly seriesCount?: number;
  readonly shapeCount?: number;
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

export type XlsxSelectionContext = XlsxRangeSelectionContext | XlsxElementContext;

export const MAX_SELECTION_AREAS = 128;
export const MAX_SELECTION_CONTEXT_CELLS = 10_000;
export const MAX_SELECTION_CONTEXT_TEXT_CHARACTERS = 8 * 1_024 * 1_024;

export function selectionCoordinateCountUpperBound(state: XlsxSelectionState): number {
  return state.areas.reduce((sum, area) => sum + (
    area.kind === 'cells'
      ? (area.bottom - area.top + 1) * (area.right - area.left + 1)
      : area.kind === 'rows'
        ? (area.lastRow - area.firstRow + 1) * MAX_WORKSHEET_COL
        : area.kind === 'columns'
          ? (area.lastColumn - area.firstColumn + 1) * MAX_WORKSHEET_ROW
          : MAX_WORKSHEET_ROW * MAX_WORKSHEET_COL
  ), 0);
}

function integerInRange(value: number, max: number): boolean {
  return Number.isInteger(value) && value >= 1 && value <= max;
}

function normalizeArea(area: XlsxSelectionArea): XlsxSelectionArea {
  switch (area.kind) {
    case 'cells': {
      if (
        !integerInRange(area.top, MAX_WORKSHEET_ROW) ||
        !integerInRange(area.bottom, MAX_WORKSHEET_ROW) ||
        !integerInRange(area.left, MAX_WORKSHEET_COL) ||
        !integerInRange(area.right, MAX_WORKSHEET_COL)
      ) throw new RangeError('Cell selection bounds must be inside the XLSX grid.');
      return {
        kind: 'cells',
        top: Math.min(area.top, area.bottom),
        left: Math.min(area.left, area.right),
        bottom: Math.max(area.top, area.bottom),
        right: Math.max(area.left, area.right),
      };
    }
    case 'rows':
      if (
        !integerInRange(area.firstRow, MAX_WORKSHEET_ROW) ||
        !integerInRange(area.lastRow, MAX_WORKSHEET_ROW)
      ) throw new RangeError('Row selection bounds must be inside the XLSX grid.');
      return {
        kind: 'rows',
        firstRow: Math.min(area.firstRow, area.lastRow),
        lastRow: Math.max(area.firstRow, area.lastRow),
      };
    case 'columns':
      if (
        !integerInRange(area.firstColumn, MAX_WORKSHEET_COL) ||
        !integerInRange(area.lastColumn, MAX_WORKSHEET_COL)
      ) throw new RangeError('Column selection bounds must be inside the XLSX grid.');
      return {
        kind: 'columns',
        firstColumn: Math.min(area.firstColumn, area.lastColumn),
        lastColumn: Math.max(area.firstColumn, area.lastColumn),
      };
    case 'sheet':
      return { kind: 'sheet' };
  }
}

export function areaContainsCell(area: XlsxSelectionArea, cell: CellAddress): boolean {
  switch (area.kind) {
    case 'cells':
      return cell.row >= area.top && cell.row <= area.bottom &&
        cell.col >= area.left && cell.col <= area.right;
    case 'rows':
      return cell.row >= area.firstRow && cell.row <= area.lastRow;
    case 'columns':
      return cell.col >= area.firstColumn && cell.col <= area.lastColumn;
    case 'sheet':
      return true;
  }
}

function normalizeCell(cell: CellAddress, name: string): CellAddress {
  if (
    !integerInRange(cell.row, MAX_WORKSHEET_ROW) ||
    !integerInRange(cell.col, MAX_WORKSHEET_COL)
  ) throw new RangeError(`${name} must be inside the XLSX grid.`);
  return { row: cell.row, col: cell.col };
}

/** Validate, normalize, and detach caller-owned selection objects. */
export function normalizeSelectionState(state: XlsxSelectionState): XlsxSelectionState {
  if (!Array.isArray(state.areas) || state.areas.length === 0) {
    throw new TypeError('A selection must contain at least one area.');
  }
  if (state.areas.length > MAX_SELECTION_AREAS) {
    throw new RangeError(`A selection may contain at most ${MAX_SELECTION_AREAS} areas.`);
  }
  if (!Number.isInteger(state.activeAreaIndex) ||
      state.activeAreaIndex < 0 || state.activeAreaIndex >= state.areas.length) {
    throw new RangeError('activeAreaIndex must identify an area in the selection.');
  }
  const normalizedAreas = state.areas.map(normalizeArea);
  const areas: XlsxSelectionArea[] = [];
  const areaIndices = new Map<string, number>();
  let activeAreaIndex = 0;
  for (let index = 0; index < normalizedAreas.length; index++) {
    const area = normalizedAreas[index];
    const key = area.kind === 'cells'
      ? `c:${area.top}:${area.left}:${area.bottom}:${area.right}`
      : area.kind === 'rows'
        ? `r:${area.firstRow}:${area.lastRow}`
        : area.kind === 'columns'
          ? `k:${area.firstColumn}:${area.lastColumn}`
          : 's';
    let canonicalIndex = areaIndices.get(key);
    if (canonicalIndex === undefined) {
      canonicalIndex = areas.length;
      areaIndices.set(key, canonicalIndex);
      areas.push(area);
    }
    if (index === state.activeAreaIndex) activeAreaIndex = canonicalIndex;
  }
  const activeCell = normalizeCell(state.activeCell, 'activeCell');
  const extensionAnchor = normalizeCell(state.extensionAnchor, 'extensionAnchor');
  const activeArea = areas[activeAreaIndex];
  if (!areaContainsCell(activeArea, activeCell)) {
    throw new RangeError('activeCell must be inside the active selection area.');
  }
  if (!areaContainsCell(activeArea, extensionAnchor)) {
    throw new RangeError('extensionAnchor must be inside the active selection area.');
  }
  return { areas, activeAreaIndex, activeCell, extensionAnchor };
}

function columnNumber(letters: string): number | null {
  let col = 0;
  for (const char of letters) col = col * 26 + char.charCodeAt(0) - 64;
  return col >= 1 && col <= MAX_WORKSHEET_COL ? col : null;
}

/** Parse one contiguous Excel-style reference. Multi-area state uses the structured form. */
export function selectionStateFromReference(reference: string): XlsxSelectionState | null {
  const ref = reference.trim().toUpperCase();
  const rowRange = /^\$?(\d+):\$?(\d+)$/.exec(ref);
  if (rowRange) {
    const first = Number(rowRange[1]);
    const last = Number(rowRange[2]);
    if (!integerInRange(first, MAX_WORKSHEET_ROW) || !integerInRange(last, MAX_WORKSHEET_ROW)) return null;
    const firstRow = Math.min(first, last);
    const lastRow = Math.max(first, last);
    return {
      areas: [{ kind: 'rows', firstRow, lastRow }],
      activeAreaIndex: 0,
      activeCell: { row: firstRow, col: 1 },
      extensionAnchor: { row: firstRow, col: 1 },
    };
  }
  const columnRange = /^\$?([A-Z]+):\$?([A-Z]+)$/.exec(ref);
  if (columnRange) {
    const first = columnNumber(columnRange[1]);
    const last = columnNumber(columnRange[2]);
    if (first === null || last === null) return null;
    const firstColumn = Math.min(first, last);
    const lastColumn = Math.max(first, last);
    return {
      areas: [{ kind: 'columns', firstColumn, lastColumn }],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: firstColumn },
      extensionAnchor: { row: 1, col: firstColumn },
    };
  }
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const first = parseA1(parts[0]);
  const last = parts.length === 2 ? parseA1(parts[1]) : first;
  if (!first || !last) return null;
  const area: XlsxSelectionArea = {
    kind: 'cells',
    top: Math.min(first.row, last.row),
    left: Math.min(first.col, last.col),
    bottom: Math.max(first.row, last.row),
    right: Math.max(first.col, last.col),
  };
  // A reference describes an area, not Excel's ActiveCell. Use the normalized
  // upper-left deterministically; structured state expresses another ActiveCell.
  const upperLeft = { row: area.top, col: area.left };
  return { areas: [area], activeAreaIndex: 0, activeCell: upperLeft, extensionAnchor: upperLeft };
}

export function selectionStatesEqual(a: XlsxSelectionState | null, b: XlsxSelectionState | null): boolean {
  if (a === b) return true;
  if (!a || !b || a.activeAreaIndex !== b.activeAreaIndex || a.areas.length !== b.areas.length) return false;
  if (a.activeCell.row !== b.activeCell.row || a.activeCell.col !== b.activeCell.col ||
      a.extensionAnchor.row !== b.extensionAnchor.row || a.extensionAnchor.col !== b.extensionAnchor.col) return false;
  return a.areas.every((area, index) => JSON.stringify(area) === JSON.stringify(b.areas[index]));
}
