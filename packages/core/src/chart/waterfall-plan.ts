export interface WaterfallPaintSite {
  start: number;
  end: number;
  isSub: boolean;
  isPos: boolean;
  hasValue: boolean;
  paintSlot: boolean;
  semanticIndex: 0 | 1 | 2;
}

export interface WaterfallPlan {
  bars: WaterfallPaintSite[];
  rawMin: number;
  rawMax: number;
  cumulativeOverflow: boolean;
}

/** Pure Waterfall source projection shared by paint, effect budgeting, and
 * structured-paint preflight. Missing numeric points retain their category
 * slot; non-finite authored values never become Canvas paint sites. */
export function planWaterfallPaintSites(
  values: ReadonlyArray<number | null>,
  categoryCount: number,
  subtotalIndices: ReadonlyArray<number>,
): WaterfallPlan {
  const count = Math.max(values.length, categoryCount);
  const subtotals = new Set(subtotalIndices);
  let cumulativeOverflow = false;
  const safeAdd = (left: number, right: number): number => {
    const sum = left + right;
    if (Number.isFinite(sum)) return sum;
    cumulativeOverflow = true;
    return sum < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
  };
  let running = 0;
  let rawMax = Number.NEGATIVE_INFINITY;
  let rawMin = 0;
  const bars: WaterfallPaintSite[] = [];
  for (let index = 0; index < count; index++) {
    const authoredValue = values[index];
    const hasValue = authoredValue != null && Number.isFinite(authoredValue);
    const paintSlot = authoredValue == null || hasValue;
    const value = hasValue ? authoredValue as number : 0;
    const isSub = subtotals.has(index);
    if (isSub) {
      const bar: WaterfallPaintSite = {
        start: 0, end: value, isSub: true, isPos: true, hasValue, paintSlot,
        semanticIndex: 2,
      };
      bars.push(bar);
      if (paintSlot) {
        rawMax = Math.max(rawMax, bar.start, bar.end);
        rawMin = Math.min(rawMin, bar.start, bar.end);
      }
      if (hasValue) running = value;
    } else {
      const next = safeAdd(running, value);
      const start = value >= 0 ? running : next;
      const end = value >= 0 ? next : running;
      const bar: WaterfallPaintSite = {
        start, end, isSub: false, isPos: value >= 0, hasValue, paintSlot,
        semanticIndex: value >= 0 ? 0 : 1,
      };
      bars.push(bar);
      if (paintSlot) {
        rawMax = Math.max(rawMax, start, end);
        rawMin = Math.min(rawMin, start, end);
      }
      if (hasValue) running = next;
    }
  }
  return { bars, rawMin, rawMax, cumulativeOverflow };
}
