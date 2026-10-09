import type { Worksheet } from '../types.js';

export type RowResizeRange = Readonly<{ first: number; last: number; height: number }>;
type Sizes = { columns?: Map<number, number>; rows?: readonly RowResizeRange[] };
const sizes = new WeakMap<Worksheet, Sizes>();
const emptyRows: readonly RowResizeRange[] = Object.freeze([]);

// XLSX view-only sizing context. Keep the two resize representations under one
// projection owner: mutable column maps need copying, while frozen row runs can
// be shared. Parser maps and admission/resource policy stay separate.
function retain(ws: Worksheet, context: Sizes): void {
  if (context.columns || context.rows) sizes.set(ws, context);
  else sizes.delete(ws);
}

export function columnCssWidths(ws: Worksheet): ReadonlyMap<number, number> | undefined {
  return sizes.get(ws)?.columns;
}

export function getColumnCssWidth(ws: Worksheet, index: number): number | undefined {
  return sizes.get(ws)?.columns?.get(index);
}

export function setColumnCssWidth(ws: Worksheet, index: number, cssPx: number | null): boolean {
  const context = sizes.get(ws), map = context?.columns;
  if (cssPx === null || !Number.isFinite(cssPx) || cssPx < 0) {
    if (!map?.delete(index)) return false;
    if (map.size === 0) {
      delete (context as Sizes).columns;
      retain(ws, context as Sizes);
    }
    return true;
  }
  if (map?.get(index) === cssPx) return false;
  if (map) map.set(index, cssPx);
  else sizes.set(ws, { ...context, columns: new Map([[index, cssPx]]) });
  return true;
}

export function rowResizeRanges(ws: Worksheet): readonly RowResizeRange[] {
  return sizes.get(ws)?.rows ?? emptyRows;
}

// View-only resource policy: bound interval fragmentation, not selected row
// count. One interval can resize all 1,048,576 worksheet rows. Hidden rows are
// excluded at gesture start; subsequent size-0 point edits still win in geometry.
export const MAX_ROW_RESIZE_INTERVALS = 16_384;

/** Validate the private wire channel before any projection mutation. */
export function validateRowResizeRanges(ranges: readonly RowResizeRange[]): void {
  if (ranges.length > MAX_ROW_RESIZE_INTERVALS) throw new RangeError('Too many row resize intervals.');
  let previous = 0;
  for (const range of ranges) {
    if (!Number.isInteger(range.first) || !Number.isInteger(range.last)
      || range.first <= previous || range.last < range.first || range.last > 1_048_576
      || !Number.isFinite(range.height) || range.height <= 0) {
      throw new RangeError('Invalid row resize interval.');
    }
    previous = range.last;
  }
}

/** The immutable ranges are projection metadata, never parser/model facts. */
export function setRowResizeRanges(ws: Worksheet, ranges: readonly RowResizeRange[]): void {
  validateRowResizeRanges(ranges);
  const context = { ...sizes.get(ws) };
  if (ranges.length) context.rows = Object.isFrozen(ranges) && ranges.every(Object.isFrozen)
    ? ranges : Object.freeze(ranges.map(r => Object.freeze({ ...r })));
  else delete context.rows;
  retain(ws, context);
}

/** Fresh viewer, automatic-height and worker projections copy all sizing
 * metadata together. Later edits cannot mutate the source's context or map. */
export function inheritWorksheetSizeContext(source: Worksheet, projection: Worksheet): void {
  const context = sizes.get(source);
  if (!context) sizes.delete(projection);
  else sizes.set(projection, {
    ...context,
    columns: context.columns && new Map(context.columns),
  });
}

/** Column-only compatibility seam; preserve any existing row context. */
export function inheritColumnCssWidths(source: Worksheet, projection: Worksheet): void {
  const columns = sizes.get(source)?.columns, context = { ...sizes.get(projection) };
  if (columns) context.columns = new Map(columns);
  else delete context.columns;
  retain(projection, context);
}
