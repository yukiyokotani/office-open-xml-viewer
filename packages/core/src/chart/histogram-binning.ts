import type { ChartexHistogramBinning } from '../types/chart';
import { formatChartValWithCode } from './chart-number-format.js';
import { hostNumberSymbols } from '../number-symbols.js';

/** Matches the shared parser's maximum retained ChartEx cache width. */
export const MAX_HISTOGRAM_INPUT_POINTS = 1_048_576;

/**
 * Histogram output is a Canvas primitive plan, not a lossless data cache.
 * Keep it comfortably below the general 10,000-mark Canvas ceiling so an
 * authored microscopic bin size cannot expand a compact source into a large
 * synchronous paint.
 */
export const MAX_HISTOGRAM_BINS = 512;

export type HistogramBinPlan =
  | { kind: 'bins'; categories: string[]; counts: number[] }
  | { kind: 'tooManyInputPoints' };

function finiteOrNull(value: number | null | undefined): number | null {
  return value != null && Number.isFinite(value) ? value : null;
}

/** Number of decimal places in the shortest round-trip form of `value`
 *  (Infinity when it is written with an exponent). */
function decimalPlaces(value: number): number {
  const text = String(Math.abs(value));
  if (text.includes('e')) return Number.POSITIVE_INFINITY;
  const dot = text.indexOf('.');
  return dot < 0 ? 0 : text.length - dot - 1;
}

/**
 * Remove binary accumulation noise (3.9000000000000004 -> 3.9) relative to
 * the bin geometry: round to the decimal grid of the width and first edge when
 * that grid is exactly representable, otherwise to 15 significant digits.
 * Integers below 1e15 are never altered, so 15-digit or epoch-millisecond
 * edges stay distinct.
 */
function makeEdgeCleaner(grid: number): (value: number) => number {
  return (value) => {
    if (!Number.isFinite(value) || value === 0) return value === 0 ? 0 : value;
    if (grid <= 15 && Math.abs(value) * 10 ** grid < 9e15) {
      return Number(value.toFixed(grid));
    }
    return Number(value.toPrecision(15));
  };
}

function makeBoundaryLabel(
  formatCode: string | null | undefined,
  date1904: boolean,
): (value: number) => string {
  const general = !formatCode || formatCode.trim().toLowerCase() === 'general';
  return (value) => {
    const v = value === 0 ? 0 : value;
    // General: shortest round-trip decimal. Otherwise the value dimension's
    // authored number format, as PowerPoint applies it to bin edges.
    // The decimal point is the host locale's (number-symbols.ts); the edge
    // separator stays ", ". Excel for Windows under fr-FR draws General
    // edges with bin width 0.5 as `[0,1, 0,6]`, `(0,6, 1,1]` (synthetic
    // control, 2026-11-28).
    return general ? String(v).replace('.', hostNumberSymbols().decimal) : formatChartValWithCode(v, formatCode, date1904);
  };
}

/** Office's automatic width: 3.5 * sample sigma / n^(1/3), 2 significant digits
 *  (measured against PowerPoint 16.113). sigma and n span every finite point;
 *  n = 1 or sigma = 0 gives 5. Streams over the source without copying it. */
function automaticWidth(source: readonly (number | null | undefined)[]): number {
  let n = 0;
  let mean = 0;
  let m2 = 0;
  for (const v of source) {
    if (v == null || !Number.isFinite(v)) continue;
    n++;
    const delta = v - mean;
    mean += delta / n;
    m2 += delta * (v - mean);
  }
  if (n <= 1) return 5;
  const sigma = Math.sqrt(m2 / (n - 1));
  if (!(sigma > 0)) return 5;
  const width = Number(((3.5 * sigma) / Math.cbrt(n)).toPrecision(2));
  return width > 0 ? width : 5;
}

/**
 * Aggregate raw ChartEx histogram observations into a bounded bar plan.
 *
 * MS-ODRAWXML defines authored bin size/count and interval boundaries but not
 * automatic bin selection. Omission follows the rule measured from PowerPoint:
 * width = 3.5 * sample sigma / n^(1/3) rounded to 2 significant digits, first
 * edge = minimum (or underflow), count = ceil(range / width). Authored plans
 * beyond the Canvas bound are coarsened over the same domain instead of
 * allocating an unbounded counts array.
 *
 * Labels use Office's text form: `[a, b]` for the first bin, `(a, b]` after it
 * (r-closed); `[a, b)` with a closed last bin (l-closed); `≤ u` / `> o` for
 * r-closed underflow/overflow. The l-closed forms `< u` / `≥ o` are unmeasured
 * and kept symmetric.
 */
export function planHistogramBins(
  source: readonly (number | null | undefined)[],
  options: ChartexHistogramBinning,
  formatCode?: string | null,
  date1904 = false,
): HistogramBinPlan {
  if (source.length > MAX_HISTOGRAM_INPUT_POINTS) {
    return { kind: 'tooManyInputPoints' };
  }
  const boundaryLabel = makeBoundaryLabel(options.edgeFormatCode ?? formatCode, date1904);

  // CT_Binning@intervalClosed is optional with no schema default
  // (MS-ODRAWXML 2.24.3.7 / 2.24.4.12 ST_IntervalClosedSide: "l" | "r"), and
  // omission is unmeasured against Office; keep the historical left-closed
  // reading rather than guess.
  const intervalClosed = options.intervalClosed === 'r' ? 'r' : 'l';
  let underflow = finiteOrNull(options.underflow);
  let overflow = finiteOrNull(options.overflow);
  if (underflow != null && overflow != null && underflow >= overflow) {
    underflow = null;
    overflow = null;
  }

  const isUnderflow = (value: number): boolean => underflow != null
    && (intervalClosed === 'r' ? value <= underflow : value < underflow);
  const isOverflow = (value: number): boolean => overflow != null
    && (intervalClosed === 'r' ? value > overflow : value >= overflow);

  let regularMin = Number.POSITIVE_INFINITY;
  let regularMax = Number.NEGATIVE_INFINITY;
  let regularCount = 0;
  let underflowCount = 0;
  let overflowCount = 0;
  for (const sourceValue of source) {
    const value = finiteOrNull(sourceValue);
    if (value == null) continue;
    if (isUnderflow(value)) {
      underflowCount++;
    } else if (isOverflow(value)) {
      overflowCount++;
    } else {
      regularCount++;
      regularMin = Math.min(regularMin, value);
      regularMax = Math.max(regularMax, value);
    }
  }
  if (regularCount + underflowCount + overflowCount === 0) {
    return { kind: 'bins', categories: [], counts: [] };
  }

  const categories: string[] = [];
  const counts: number[] = [];
  if (underflow != null) {
    categories.push(`${intervalClosed === 'r' ? '≤' : '<'} ${boundaryLabel(underflow)}`);
    counts.push(underflowCount);
  }

  if (regularCount > 0) {
    const lower = underflow ?? regularMin;
    const upper = overflow ?? regularMax;
    const range = upper - lower;
    if (!Number.isFinite(range)) {
      categories.push(`[${boundaryLabel(lower)}, ${boundaryLabel(upper)}]`);
      counts.push(regularCount);
    } else {
      const authoredSize = finiteOrNull(options.binSize);
      const authoredCount = options.binCount != null && Number.isFinite(options.binCount) && options.binCount > 0
        ? Math.max(1, Math.floor(options.binCount))
        : null;
      const sizeMode = authoredSize != null && authoredSize > 0;
      // Width-driven planning: authored size, or the automatic width.
      const widthDriven = sizeMode || authoredCount == null;
      let nominalWidth = 0;
      let requestedCount: number;
      if (widthDriven) {
        nominalWidth = sizeMode ? authoredSize : automaticWidth(source);
        requestedCount = 1;
        if (range > 0 && Number.isFinite(nominalWidth) && nominalWidth > 0) {
          // ceil(range / width), decided on the cleaned decimal edges so that
          // 3.0000000000000004 widths of float noise never add a bin.
          const cleanEdgeAt = makeEdgeCleaner(Math.max(decimalPlaces(nominalWidth), decimalPlaces(lower)));
          const at = (k: number): number => cleanEdgeAt(lower + nominalWidth * k);
          let k = Math.max(1, Math.ceil(range / nominalWidth));
          if (Number.isFinite(k) && k <= MAX_HISTOGRAM_BINS * 4) {
            while (k > 1 && at(k - 1) >= upper) k--;
            while (at(k) < upper && k <= MAX_HISTOGRAM_BINS * 4) k++;
          }
          requestedCount = k;
        }
        if (!Number.isFinite(requestedCount)) requestedCount = 1;
      } else {
        requestedCount = authoredCount;
      }
      const binCount = range <= 0 ? 1 : Math.min(MAX_HISTOGRAM_BINS, requestedCount);
      const usesWidth = widthDriven && requestedCount <= MAX_HISTOGRAM_BINS
        && Number.isFinite(nominalWidth) && nominalWidth > 0;
      const width = usesWidth ? nominalWidth : range === 0 ? 1 : range / binCount;
      const regularCounts = new Array<number>(binCount).fill(0);
      const grid = usesWidth
        ? Math.max(decimalPlaces(width), decimalPlaces(lower))
        : Number.POSITIVE_INFINITY;
      const clean = makeEdgeCleaner(grid);
      const edges: number[] = [];
      for (let k = 0; k <= binCount; k++) {
        const raw = usesWidth ? lower + width * k : lower + (range * k) / binCount;
        edges.push(clean(overflow == null ? raw : Math.min(raw, overflow)));
      }
      let strict = true;
      for (let k = 1; k <= binCount; k++) if (!(edges[k] > edges[k - 1])) strict = false;
      for (const sourceValue of source) {
        const value = finiteOrNull(sourceValue);
        if (value == null) continue;
        if (isUnderflow(value) || isOverflow(value)) continue;
        let index: number;
        if (strict) {
          // Compare against the cleaned decimal boundaries themselves.
          // r-closed: first bin whose upper edge >= value; l-closed: last bin
          // whose lower edge <= value. The top edge stays in the last bin.
          let lo = 0;
          let hi = binCount - 1;
          if (intervalClosed === 'r') {
            while (lo < hi) {
              const mid = (lo + hi) >> 1;
              if (value <= edges[mid + 1]) hi = mid; else lo = mid + 1;
            }
          } else {
            while (lo < hi) {
              const mid = (lo + hi + 1) >> 1;
              if (value >= edges[mid]) lo = mid; else hi = mid - 1;
            }
          }
          index = lo;
        } else {
          // Degenerate (subnormal / non-resolvable) edges: normalize by the
          // whole range, since dividing a subnormal range by binCount can
          // underflow the width to zero.
          const position = range === 0 ? 0 : ((value - lower) / range) * binCount;
          const rawIndex = intervalClosed === 'r' ? Math.ceil(position) - 1 : Math.floor(position);
          index = Math.max(0, Math.min(binCount - 1, rawIndex));
        }
        regularCounts[index]++;
      }
      for (let index = 0; index < binCount; index++) {
        const start = boundaryLabel(edges[index]);
        const end = boundaryLabel(edges[index + 1]);
        const last = index === binCount - 1;
        const leftOpen = intervalClosed === 'r' && (underflow != null || index > 0);
        const rightOpen = intervalClosed === 'l' && !(last && overflow == null);
        categories.push(`${leftOpen ? '(' : '['}${start}, ${end}${rightOpen ? ')' : ']'}`);
        counts.push(regularCounts[index]);
      }
    }
  }

  if (overflow != null) {
    categories.push(`${intervalClosed === 'r' ? '>' : '≥'} ${boundaryLabel(overflow)}`);
    counts.push(overflowCount);
  }
  return { kind: 'bins', categories, counts };
}
