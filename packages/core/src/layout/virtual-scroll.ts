/**
 * Virtualization range math for the continuous-scroll viewers (DocxScrollViewer /
 * PptxScrollViewer). Pure and DOM-free — the viewer owns the scroll surface and
 * the slot pool (design §6); this only answers "given the heights, the gap, and
 * the current scrollTop, which item indices must be mounted, where does each
 * item sit, and how tall is the whole scroll region?". Prefix-sum offsets +
 * binary search of the first visible index. See design §5.1
 * (docs/dev-notes/2026-07-01-scroll-viewer-design.md).
 */
export interface VisibleWindow {
  /** First index to mount (inclusive, includes overscan). `start > end` ⇒ nothing
   *  to mount (empty input, or a 0-height viewport whose top sits exactly on an
   *  item boundary) — mount loops over `[start, end]` naturally run zero times. */
  start: number;
  /** Last index to mount (inclusive, includes overscan). Empty input ⇒ -1. */
  end: number;
  /** First item intersecting the viewport top (EXCLUDES overscan) — for
   *  onVisiblePageChange. A viewport top strictly inside the gap BETWEEN items i
   *  and i+1 is attributed to item i (gap = trailing padding of the preceding
   *  item — the standard virtualization convention; mount-safe, and flips to i+1
   *  at `offsets[i+1]`). Exception: the flip happens up to
   *  TOP_EDGE_ROUNDING_EPSILON (1 CSS px) EARLY, because fractional offsets meet
   *  a browser-snapped integer scrollTop at exactly that boundary — see the
   *  constant's doc. */
  topIndex: number;
  /** `leading` + Σ heights + (n-1)*gap + `trailing` (gap between items only, none
   *  after the last) → spacer height. With no padding this reduces to
   *  Σ heights + (n-1)*gap. */
  totalHeight: number;
}

export interface VisibleRange extends VisibleWindow {
  /** Top offset (px) of every item i: `leading` + Σ heights[0..i-1] + i*gap.
   *  length = heights.length. With no padding (`leading` 0) this reduces to the
   *  bare prefix-sum. Cached callers retain this array across scroll queries. */
  offsets: number[];
}

/** Scale-dependent prefix geometry. Build it when item sizes change, then use
 * {@link computeVisibleWindow} for O(log n) scroll queries without rebuilding
 * or reallocating the document-length offsets array. */
export interface VirtualScrollGeometry {
  readonly offsets: number[];
  readonly totalHeight: number;
}

/** Optional leading/trailing padding (px) added OUTSIDE the item run — the desk
 *  margin a PDF reader leaves above the first item and below the last. Distinct
 *  from `gap`, which only sits BETWEEN adjacent items. Both default 0, so an
 *  omitted `pad` is exactly the pre-padding behaviour (fully backward-compatible). */
export interface VisibleRangePad {
  /** px above the FIRST item (shifts every offset down by this amount). Default 0. */
  leading?: number;
  /** px below the LAST item (added to totalHeight only). Default 0. */
  trailing?: number;
}

/** Clamp `v` to `[lo, hi]` (hi < lo yields lo — only reached when n === 0, guarded upstream). */
function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/**
 * Rounding tolerance at the top-edge boundary, in CSS px. Item offsets are
 * fractional (pt × a fractional fit/zoom scale, e.g. 1193.4), but browsers
 * snap a programmatic `scrollTop` to an integer, landing strictly BELOW the
 * target item's offset — without this tolerance `scrollToPage(k)` reports
 * `topIndex` k−1 even though item k is what the viewport top shows.
 *
 * Exactly 1 px: it covers the integer snap and nothing more. A viewport top
 * deeper inside the gap still belongs to the preceding item (the user is
 * reading its tail), so a wider band — e.g. gap/2 — would misattribute
 * genuine scroll positions, not just rounding artifacts.
 */
const TOP_EDGE_ROUNDING_EPSILON = 1;

/** Build the variable-height prefix geometry once per layout/scale revision. */
export function createVirtualScrollGeometry(
  heights: readonly number[],
  gap: number,
  pad?: VisibleRangePad,
): VirtualScrollGeometry {
  const n = heights.length;
  if (n === 0) return { offsets: [], totalHeight: 0 };
  const leading = pad?.leading ?? 0;
  const trailing = pad?.trailing ?? 0;
  const offsets = new Array<number>(n);
  let acc = 0;
  for (let i = 0; i < n; i++) {
    offsets[i] = leading + acc + i * gap;
    acc += heights[i];
  }
  return {
    offsets,
    totalHeight: leading + acc + (n - 1) * gap + trailing,
  };
}

/** Query cached variable-height geometry. Only the two binary searches depend
 * on the current scroll position; the returned range reuses `geometry.offsets`. */
export function computeVisibleWindow(
  geometry: VirtualScrollGeometry,
  scrollTop: number,
  viewportHeight: number,
  overscan: number,
): VisibleRange {
  const offsets = geometry.offsets;
  const n = offsets.length;
  if (n === 0) {
    return { start: 0, end: -1, topIndex: 0, offsets, totalHeight: 0 };
  }

  // The top edge gets TOP_EDGE_ROUNDING_EPSILON slack: a scrollTop snapped by
  // the browser to an integer just below a fractional item start still counts
  // as that item. The bottom edge needs no such slack — a sub-pixel difference
  // there never changes which items intersect the viewport.
  const topEdge = scrollTop + TOP_EDGE_ROUNDING_EPSILON;
  let lo = 0;
  let hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (offsets[mid] <= topEdge) lo = mid + 1;
    else hi = mid;
  }
  const topIndex = clamp(lo - 1, 0, n - 1);

  const bottom = scrollTop + viewportHeight;
  lo = 0;
  hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (offsets[mid] < bottom) lo = mid + 1;
    else hi = mid;
  }
  const lastVisible = clamp(lo - 1, 0, n - 1);
  return {
    start: clamp(topIndex - overscan, 0, n - 1),
    end: clamp(lastVisible + overscan, 0, n - 1),
    topIndex,
    offsets,
    totalHeight: geometry.totalHeight,
  };
}

/** O(1) visible-window arithmetic for uniform slides. Unlike the compatibility
 * `computeVisibleRange` API, this deliberately materializes no offsets array. */
export function computeUniformVisibleWindow(
  itemCount: number,
  itemHeight: number,
  gap: number,
  scrollTop: number,
  viewportHeight: number,
  overscan: number,
  pad?: VisibleRangePad,
): VisibleWindow {
  if (itemCount === 0) return { start: 0, end: -1, topIndex: 0, totalHeight: 0 };
  const leading = pad?.leading ?? 0;
  const trailing = pad?.trailing ?? 0;
  const stride = itemHeight + gap;
  // Same boundary rule as computeVisibleWindow: the +1 px slack only ever
  // flips the single boundary the browser's integer scrollTop snap can cross,
  // because a real slide stride is always ≫ 1 px.
  const topEdge = scrollTop + TOP_EDGE_ROUNDING_EPSILON;
  const topIndex = clamp(
    topEdge < leading
      ? 0
      : stride > 0 ? Math.floor((topEdge - leading) / stride) : itemCount - 1,
    0,
    itemCount - 1,
  );
  const bottom = scrollTop + viewportHeight;
  const lastVisible = clamp(
    bottom <= leading
      ? 0
      : stride > 0 ? Math.ceil((bottom - leading) / stride) - 1 : itemCount - 1,
    0,
    itemCount - 1,
  );
  return {
    start: clamp(topIndex - overscan, 0, itemCount - 1),
    end: clamp(lastVisible + overscan, 0, itemCount - 1),
    topIndex,
    totalHeight: leading + itemCount * itemHeight + (itemCount - 1) * gap + trailing,
  };
}

/**
 * Compute which item slots to mount for a vertical virtualized scroll region.
 *
 * @param heights        per-item extent along the scroll axis (px), in order.
 * @param gap            px between adjacent items (contributes BETWEEN items only,
 *                       never before the first nor after the last).
 * @param scrollTop      current scroll offset (px); negative is treated as 0 via
 *                       the search / clamps.
 * @param viewportHeight visible height of the scroll surface (px).
 * @param overscan       extra items kept mounted beyond the viewport on each side.
 * @param pad            optional {@link VisibleRangePad} desk margin OUTSIDE the
 *                       item run: `leading` px above the first item shifts every
 *                       offset down; `trailing` px below the last item extends
 *                       totalHeight. Both default 0 — an omitted `pad` is exactly
 *                       the pre-padding behaviour (backward-compatible). A viewport
 *                       top inside the leading pad (scrollTop < leading, so below
 *                       every offset) yields topIndex 0 via the existing clamp.
 * @returns a {@link VisibleRange}. Empty `heights` ⇒
 *          `{ start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 }`
 *          (an empty mount range: `start > end`). NOTE: with n === 0 the padding
 *          is DELIBERATELY NOT applied — an empty document shows no desk padding,
 *          consistent with the viewers' empty-doc no-op contract (they never mount
 *          a spacer for a zero-item document). `pad` only takes effect once there
 *          is at least one item.
 */
export function computeVisibleRange(
  heights: number[],
  gap: number,
  scrollTop: number,
  viewportHeight: number,
  overscan: number,
  pad?: VisibleRangePad,
): VisibleRange {
  return computeVisibleWindow(
    createVirtualScrollGeometry(heights, gap, pad),
    scrollTop,
    viewportHeight,
    overscan,
  );
}
