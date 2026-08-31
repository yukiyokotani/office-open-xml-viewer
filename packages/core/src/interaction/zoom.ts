/** Multiplicative scale change for one conventional wheel step. Purely an
 * interaction-feel constant (no ECMA-376 bearing). */
const ZOOM_FACTOR_PER_WHEEL_STEP = 1.1;

/** WheelEvent delta units that amount to one zoom step. Pixel-mode trackpad
 * events stay proportional up to this boundary; larger discrete mouse-wheel
 * deltas are capped to one step per event. Line-mode browsers commonly emit
 * three lines and page mode one page for the same physical wheel action. */
const PIXELS_PER_WHEEL_STEP = 10;
const LINES_PER_WHEEL_STEP = 3;
const PAGES_PER_WHEEL_STEP = 1;

function wheelSteps(deltaY: number, deltaMode: number): number {
  let steps: number;
  switch (deltaMode) {
    case 1:
      steps = deltaY / LINES_PER_WHEEL_STEP;
      break;
    case 2:
      steps = deltaY / PAGES_PER_WHEEL_STEP;
      break;
    default:
      steps = deltaY / PIXELS_PER_WHEEL_STEP;
      break;
  }
  return Math.max(-1, Math.min(1, steps));
}

/**
 * New scale for a wheel/pinch zoom event. The step is *exponential* in a
 * `deltaMode`-normalized and bounded wheel distance rather than a fixed per-event
 * increment, which fixes three problems with a sign-only `scale ± 0.1`:
 *
 *  - A trackpad pinch arrives as a high-frequency stream of small-`deltaY`
 *    wheel events; a fixed per-event increment compounds across dozens of
 *    events per gesture and zooms wildly. Because `exp(-k·a)·exp(-k·b) =
 *    exp(-k·(a+b))`, fine-grained events combine according to the summed
 *    gesture distance rather than the event count.
 *  - It is multiplicative, so a step feels proportional at every zoom level
 *    (the old additive `+0.1` was huge at 20% and tiny at 400%), and exactly
 *    symmetric: zooming in then out by the same delta returns to the start.
 *  - Browsers report wheel deltas in pixels, lines, or pages. Normalizing
 *    `deltaMode` and capping a single event gives each conventional mouse-wheel
 *    step the same gentle 10% change instead of letting a common pixel delta of
 *    100 jump by ~2.72×. Small trackpad deltas remain proportional.
 *
 * Negative `deltaY` (scroll up / pinch out) zooms in. The result is unclamped
 * and unsnapped; the caller clamps to its `[zoomMin, zoomMax]` range (and, for
 * XlsxViewer, snaps to whole percent). Shared by XlsxViewer and the two
 * continuous-scroll viewers (design §5.2).
 */
export function zoomStepScale(
  currentScale: number,
  deltaY: number,
  deltaMode: number = 0,
): number {
  return currentScale * Math.pow(ZOOM_FACTOR_PER_WHEEL_STEP, -wheelSteps(deltaY, deltaMode));
}

/** Clamp range for {@link anchoredZoomOffset}: the scroll offset must stay within
 *  `[0, maxScroll]` (the browser clamps native scrollLeft/scrollTop the same way).
 *  `maxScroll` is the largest legal scroll on that axis at the NEW scale (usually
 *  `totalContentPx − viewportPx`, floored at 0). */
export interface ScrollClamp {
  /** Largest legal scroll offset on this axis at the new scale (≥ 0). A degenerate
   *  `maxScroll < 0` (content shorter than the viewport) is treated as 0. */
  maxScroll: number;
}

/**
 * The new scroll offset (one axis) that keeps the content point currently under a
 * gesture anchor pinned under that same anchor across a zoom — i.e. pointer-
 * anchored ("zoom toward the cursor") zoom, instead of the viewport-top / origin
 * anchoring a naive rescale gives.
 *
 * DERIVATION. Work in the SCALING content's own pixels:
 *
 *   - `scrollOld` — the current NATIVE scroll offset (scaled px, ≥ 0).
 *   - `anchor`    — the pointer's offset INTO the scaling region (scaled px).
 *                   The content pixel under the pointer is `c = scrollOld + anchor`.
 *   - A zoom multiplies every content pixel by `ratio = scaleNew / scaleOld`, so
 *     that SAME logical content sits at `c' = c · ratio` at the new scale.
 *   - To keep it under the (unchanged, screen-fixed) pointer, we need
 *     `scrollNew + anchor = c'`, hence:
 *
 *         scrollNew = (scrollOld + anchor) · (scaleNew / scaleOld) − anchor
 *
 * LEAD-INS — how the caller derives `anchor` from a raw pointer position `x`
 * (measured from the scroll viewport's origin). The two kinds behave differently:
 *
 *   - FIXED lead-in (scale-INVARIANT space before the scaling content, e.g. the
 *     scroll viewers' left desk gutter `padL`, where screen-x of content pixel c
 *     is `padL + c − scroll`): subtract it from the ANCHOR ONLY — `anchor = x −
 *     padL` — and keep `scrollOld` and the clamp in NATIVE scroll space. Do NOT
 *     also shift the scroll by ±padL around the call: that runs the fixed gutter
 *     through the zoom ratio and over-compensates by `padL·(ratio−1)` per step.
 *   - SCALING lead-in (a lead-in drawn at ×scale, e.g. XlsxViewer's fixed-screen
 *     header + frozen band of UNSCALED extent K, where the logical content under
 *     the pointer is `(x + scroll)/scale − K`): the `K·scale` terms cancel out of
 *     the invariance equation entirely, so pass the RAW pointer — `anchor = x`.
 *
 * IDENTITY: `scaleNew === scaleOld` ⇒ `ratio = 1` ⇒ `scrollNew = scrollOld` (no
 * drift on a no-op zoom). ROUND-TRIP: zooming out by the inverse ratio returns the
 * original offset exactly (the map is affine and invertible). ANCHOR = 0 (pointer
 * at the region start) reduces to the plain `scrollOld · ratio` origin-anchored
 * rescale.
 *
 * CLAMP: the result is clamped to `[0, clamp.maxScroll]` — the same range the
 * browser pins native scroll to. AT A CONTENT EDGE the anchor cannot always be
 * honoured (there is no scroll offset that would place the pinned point under the
 * pointer without exposing off-content area); the clamp then wins and the point
 * drifts, which is the natural, expected behaviour (design §5.3).
 *
 * Pure; no DOM. Shared by the two continuous-scroll viewers (both axes; fixed
 * gutter subtracted from the horizontal anchor) and XlsxViewer (both axes; its
 * header + frozen band is a SCALING lead-in, so the raw pointer is the anchor).
 */
export function anchoredZoomOffset(
  scrollOld: number,
  anchor: number,
  scaleOld: number,
  scaleNew: number,
  clamp: ScrollClamp,
): number {
  // Guard a degenerate old scale (never happens for a live viewer, but keeps the
  // function total): fall back to no movement rather than dividing by zero.
  const ratio = scaleOld > 0 ? scaleNew / scaleOld : 1;
  const want = (scrollOld + anchor) * ratio - anchor;
  const max = clamp.maxScroll > 0 ? clamp.maxScroll : 0;
  return want < 0 ? 0 : want > max ? max : want;
}
