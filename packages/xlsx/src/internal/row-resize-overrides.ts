import type { Worksheet } from '../types.js';

export type RowResizeRange = Readonly<{ first: number; last: number; height: number }>;
export type RowBandRange = Readonly<{ first: number; last: number }>;
export type RowResizePreview = Readonly<{ apply(height: number): void; rollback(): void }>;

// View-only resource policy: bound interval fragmentation, not selected row
// count. One interval can resize all 1,048,576 worksheet rows. Hidden rows are
// excluded at gesture start; subsequent size-0 point edits still win in geometry.
export const MAX_ROW_RESIZE_INTERVALS = 16_384;
const rows = new WeakMap<Worksheet, readonly RowResizeRange[]>();

export function rowResizeRanges(ws: Worksheet): readonly RowResizeRange[] {
  return rows.get(ws) ?? [];
}

export function rowResizeContains(ranges: readonly RowBandRange[], index: number): boolean {
  let low = 0, high = ranges.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (ranges[middle].first <= index) low = middle + 1;
    else high = middle;
  }
  return low > 0 && ranges[low - 1].last >= index;
}

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
  if (ranges.length) rows.set(ws, Object.isFrozen(ranges) && ranges.every(Object.isFrozen)
    ? ranges : Object.freeze(ranges.map(r => Object.freeze({ ...r }))));
  else rows.delete(ws);
}

export function inheritRowResizeRanges(source: Worksheet, target: Worksheet): void {
  const ranges = rows.get(source);
  if (ranges) rows.set(target, ranges);
}

/** Overlay a uniform gesture on a canonical prior set. Sweep boundaries rather
 * than row ordinals; later overlapping edits win and equal neighbours coalesce.
 * Build and validate the complete candidate before installing any state. */
export function replaceRowResizeRanges(
  prior: readonly RowResizeRange[], targets: readonly RowBandRange[], height: number,
): readonly RowResizeRange[] {
  const boundaries = [...new Set([...prior, ...targets].flatMap(r => [r.first, r.last + 1]))]
    .sort((a, b) => a - b);
  const next: RowResizeRange[] = [];
  let oldIndex = 0, targetIndex = 0;
  for (let i = 0; i + 1 < boundaries.length; i++) {
    const first = boundaries[i], last = boundaries[i + 1] - 1;
    while (oldIndex < prior.length && prior[oldIndex].last < first) oldIndex++;
    while (targetIndex < targets.length && targets[targetIndex].last < first) targetIndex++;
    const target = targets[targetIndex], old = prior[oldIndex];
    const value = target && target.first <= first ? height
      : old && old.first <= first ? old.height : undefined;
    if (value === undefined) continue;
    const previous = next.at(-1);
    if (previous && previous.last + 1 === first && previous.height === value) {
      next[next.length - 1] = { first: previous.first, last, height: value };
    } else {
      if (next.length === MAX_ROW_RESIZE_INTERVALS) throw new RangeError('Too many row resize intervals.');
      next.push({ first, last, height: value });
    }
  }
  validateRowResizeRanges(next);
  return Object.freeze(next.map(r => Object.freeze(r)));
}
