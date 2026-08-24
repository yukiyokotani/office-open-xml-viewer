/** Which format owns the omitted category-axis gap policy. */
export type CategoryGapPolicy = 'legacy' | 'chartex';
export type CategoryLabelAlignment = 'l' | 'ctr' | 'r' | string | null | undefined;

export interface CategoryLabelAnchor {
  fraction: number;
  textAlign: CanvasTextAlign;
}

/**
 * Resolve `<c:catAx|dateAx><c:lblOffset>` against the renderer's established
 * default rule-to-label gap. ECMA-376 §21.2.3.23 defines the authored value as
 * a percentage of that default, not as a chart-size or font-size percentage.
 */
export function categoryLabelOffsetPx(
  defaultGapPx: number,
  offsetPercent?: number | null,
): number {
  if (offsetPercent == null || !Number.isFinite(offsetPercent)) return defaultGapPx;
  return defaultGapPx * Math.max(0, Math.min(1000, offsetPercent)) / 100;
}

/** Position a category-like data point inside its authored axis interval. */
export function categoryPositionFraction(
  index: number,
  count: number,
  between: boolean,
  reversed = false,
): number {
  const last = Math.max(0, count - 1);
  const safeIndex = Number.isFinite(index) ? Math.max(0, Math.min(last, index)) : 0;
  const fraction = between
    ? (safeIndex + 0.5) / Math.max(1, count)
    : count === 1 ? 0.5 : safeIndex / last;
  return reversed ? 1 - fraction : fraction;
}

/** Category-axis major gridline positions. A between-category axis owns all
 * interval boundaries; a mid-category axis owns the category centres. */
export function categoryGridlineFractions(count: number, between: boolean): number[] {
  if (count <= 0) return [];
  const fractions: number[] = [];
  const last = between ? count : count - 1;
  for (let index = 0; index <= last; index++) {
    fractions.push(between ? index / count : count === 1 ? 0.5 : index / (count - 1));
  }
  return fractions;
}

/** Category-axis minor gridline positions. Office places minor rules halfway
 * between the major rules: at category centres for a between-category axis,
 * and at the interior midpoints between category centres for a mid-category
 * axis. */
export function categoryMinorGridlineFractions(count: number, between: boolean): number[] {
  if (count <= 0) return [];
  if (between) {
    return Array.from({ length: count }, (_, index) => (index + 0.5) / count);
  }
  if (count <= 1) return [];
  return Array.from(
    { length: count - 1 },
    (_, index) => (index + 0.5) / (count - 1),
  );
}

/**
 * Resolve the physical text anchor inside one category label cell.
 *
 * `lblAlgn` aligns tick-label text inside the interval owned by the category;
 * it does not move the category data point. `between` axes own a full equal
 * interval, while `midCat` axes own the half-intervals around endpoint ticks.
 * Axis reversal is applied before choosing physical left/right so `l` and `r`
 * remain visual text alignment, not logical minimum/maximum alignment.
 */
export function categoryLabelAnchorFraction(
  index: number,
  count: number,
  between: boolean,
  reversed: boolean,
  alignment: CategoryLabelAlignment,
): CategoryLabelAnchor {
  // Omission preserves the chart family's established tick anchor. In
  // particular, a mid-category axis labels the endpoint tick itself; treating
  // omission as an authored `ctr` would instead move the end labels halfway
  // into their neighbouring intervals.
  if (alignment == null) {
    return {
      fraction: categoryPositionFraction(index, count, between, reversed),
      textAlign: 'center',
    };
  }
  const last = Math.max(0, count - 1);
  const safeIndex = Number.isFinite(index) ? Math.max(0, Math.min(last, index)) : 0;
  let start: number;
  let end: number;
  if (count <= 1) {
    start = 0;
    end = 1;
  } else if (between) {
    start = safeIndex / count;
    end = (safeIndex + 1) / count;
  } else {
    start = safeIndex === 0 ? 0 : (safeIndex - 0.5) / last;
    end = safeIndex === last ? 1 : (safeIndex + 0.5) / last;
  }
  if (reversed) {
    const reversedStart = 1 - end;
    end = 1 - start;
    start = reversedStart;
  }
  if (alignment === 'l') return { fraction: start, textAlign: 'left' };
  if (alignment === 'r') return { fraction: end, textAlign: 'right' };
  return { fraction: (start + end) / 2, textAlign: 'center' };
}

/**
 * Resolve the gap between category bodies as a percentage of one body.
 *
 * Classic `<c:barChart>` keeps the ECMA-376 default of 150%. ChartEx
 * `<cx:catScaling gapWidth>` has no schema default, so the supported ordinal
 * layouts share a small deterministic 33% fallback. An authored value has
 * already been normalized by the parser and is always authoritative.
 */
export function resolveCategoryGapWidthPercent(
  authoredPercent: number | null | undefined,
  policy: CategoryGapPolicy,
): number {
  if (authoredPercent != null && Number.isFinite(authoredPercent)) {
    return Math.max(0, Math.min(500, authoredPercent));
  }
  return policy === 'legacy' ? 150 : 33;
}
