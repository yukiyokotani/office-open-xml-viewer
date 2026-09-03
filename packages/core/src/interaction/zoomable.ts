/**
 * IX9 — the shared zoom API contract for every viewer (DocxViewer, PptxViewer,
 * DocxScrollViewer, PptxScrollViewer, XlsxViewer).
 *
 * This module owns ONLY the pure, DOM-free pieces of the contract: the type
 * ({@link ZoomableViewer}), the discrete zoom-step ladder ({@link nextZoomStep} /
 * {@link prevZoomStep}), the fit-to-content scale math ({@link fitScale}), and the
 * range clamp ({@link clampScale}). Each viewer implements the interface with its
 * own scale field and re-render path; this keeps ONE definition of "what a zoom
 * factor means" and "what the +/- steps are" across all five, so a host can drive
 * any viewer through the same six calls without special-casing the format.
 *
 * SCALE SEMANTICS (the contract): a scale of `1` means 100% — the content at its
 * natural size (a docx page at `widthPt × PT_TO_PX`, a pptx slide at
 * `slideWidth / EMU_PER_PX`, an xlsx grid at `cellScale` 1). `getScale()` and
 * `setScale(n)` speak this user-facing factor for EVERY viewer.
 *
 * KNOWN FAMILY DIFFERENCE — the INITIAL scale right after load (deliberate,
 * documented rather than papered over): the single-canvas viewers (DocxViewer /
 * PptxViewer) and XlsxViewer start at `1` (or the effective factor implied by an
 * explicit `width` option); the continuous-scroll viewers (DocxScrollViewer /
 * PptxScrollViewer) AUTO-FIT to the container on first layout, so their
 * `getScale()` right after load reports the fit-to-width BASE factor (≠ 1 unless
 * the container happens to match the natural width). The unit is identical — only
 * the starting point differs, because fit-to-width is the natural resting state
 * of a continuous document viewer.
 *
 * PRE-LOAD `setScale` (family-unified, IX9 F1): a `setScale` called before the
 * content is loaded / before the layout is established is LATCHED — never
 * silently dropped — and applied once the viewer establishes its scale (the
 * single-canvas viewers honour it on the first render; the scroll viewers apply
 * it right after the base fit establishes, firing `onScaleChange` at application
 * time). `getScale()` reports the latched factor while it is pending.
 *
 * API SHAPE (idiomatic default — the integrator MAY veto; see the IX9 PR): a
 * six-method surface plus one change notification (`onScaleChange`). Deliberately
 * NO new UI here — the contract is API only (design decision IX9 §4). Touch-pinch
 * (IX8) is out of scope.
 */

/**
 * The zoom contract every viewer satisfies. All scales are the user-facing factor
 * where `1` = 100% (see the module note). `fitWidth`/`fitPage` are async because a
 * fit re-renders at the new scale; the getters/steppers resolve synchronously.
 */
export interface ZoomableViewer {
  /** The current zoom factor (`1` = 100%). Never throws — returns the default
   *  (`1`) before anything is loaded, or the latched pending factor when a
   *  pre-load `setScale` is waiting to be applied (see the module note). */
  getScale(): number;
  /** Set the absolute zoom factor (`1` = 100%), clamped to the viewer's
   *  effective zoom range. A continuous-scroll viewer extends its floor down to
   *  a smaller width-fit base when necessary, so its opening fit stays reachable
   *  after zooming in. Re-renders and fires `onScaleChange`
   *  when the clamped value actually changes. Called BEFORE the content is
   *  loaded / the layout is established, the (clamped) factor is LATCHED and
   *  applied once the viewer establishes its scale — family-unified semantics
   *  (IX9 F1): never silently dropped by any viewer. */
  setScale(scale: number): void | Promise<void>;
  /** Step up to the next larger rung of the shared zoom ladder (25 %→400 %),
   *  clamped to `zoomMax`. Equivalent to `setScale(nextZoomStep(getScale()))`. */
  zoomIn(): void | Promise<void>;
  /** Step down to the next smaller ladder rung, then to a continuous viewer's
   *  width-fit floor when that floor is below both `zoomMin` and the ladder. */
  zoomOut(): void | Promise<void>;
  /** Fit the content's WIDTH to the container (the common "fit width" / "fit
   *  page width" verb). Sets the scale so one page/slide/sheet-column-run spans
   *  the available width, then re-renders. Resolves once the fit render settles.
   *
   *  PERSISTENCE is viewer-implementation-dependent (deliberate, by family): the
   *  single-canvas viewers (DocxViewer / PptxViewer) and XlsxViewer apply the fit
   *  ONE-SHOT — they observe no container resizes, so a later resize does NOT
   *  re-fit (call `fitWidth()` again after a layout change). The continuous-
   *  scroll viewers (DocxScrollViewer / PptxScrollViewer) re-fit their width-fit
   *  base on every container resize, so a `fitWidth()` there effectively
   *  PERSISTS across resizes (the resize re-fit preserves the width-fit state). */
  fitWidth(): void | Promise<void>;
  /** Fit the WHOLE content (width AND height) inside the container, so an entire
   *  page/slide is visible without scrolling. Sets the scale to the smaller of the
   *  width- and height-fit factors, then re-renders.
   *
   *  PERSISTENCE is viewer-implementation-dependent, and — unlike `fitWidth` —
   *  a page fit does NOT persist across container resizes on ANY viewer: the
   *  single-canvas viewers and XlsxViewer observe no resizes at all (one-shot),
   *  and the continuous-scroll viewers' resize handler re-applies the WIDTH fit
   *  (preserving the zoom multiplier), not the page fit. Re-invoke `fitPage()`
   *  after a layout change to re-fit. */
  fitPage(): void | Promise<void>;
}

/**
 * The discrete zoom rungs the +/- steppers snap through, as user-facing factors.
 * The Excel / Acrobat family (25, 33, 50, 67, 75, 90, 100, 110, 125, 150, 175,
 * 200, 250, 300, 400 %) — one consistent ascending series (a superset of Excel's
 * status-bar presets and Acrobat's toolbar presets), so `zoomIn`/`zoomOut` feel
 * familiar in either lineage. Frozen (readonly) so a caller cannot mutate the
 * shared ladder. Values are factors, not percents (0.25 … 4).
 */
export const ZOOM_STEP_LADDER: readonly number[] = Object.freeze([
  0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4,
]);

/** Small epsilon so a `scale` sitting exactly on (or a floating-point hair away
 *  from) a ladder rung steps to the NEXT rung, not back to the one it is already
 *  on. 0.5 % of a factor is far below the gap between any two adjacent rungs. */
const LADDER_EPS = 0.005;

/**
 * The next ladder rung strictly greater than `scale` (the "+" step). A `scale`
 * already at or above the top rung returns the top rung (the caller then clamps
 * to `zoomMax`, which may be higher or lower than the ladder top). A `scale`
 * between rungs jumps UP to the next rung — so an arbitrary wheel-zoomed value
 * snaps back onto the ladder on the first "+" press. Pure.
 */
export function nextZoomStep(scale: number): number {
  for (const rung of ZOOM_STEP_LADDER) {
    if (rung > scale + LADDER_EPS) return rung;
  }
  return ZOOM_STEP_LADDER[ZOOM_STEP_LADDER.length - 1];
}

/**
 * The next ladder rung strictly less than `scale` (the "−" step). Mirror of
 * {@link nextZoomStep}: a `scale` at or below the bottom rung returns the bottom
 * rung by default. Callers whose valid floor is below the ladder (for example a
 * continuous viewer's width fit) may pass it as `floor`; once no lower rung
 * exists, the step returns that floor instead. A between-rungs value jumps DOWN
 * to the next rung. Pure.
 */
export function prevZoomStep(
  scale: number,
  floor: number = ZOOM_STEP_LADDER[0],
): number {
  for (let i = ZOOM_STEP_LADDER.length - 1; i >= 0; i--) {
    const rung = ZOOM_STEP_LADDER[i];
    if (rung < scale - LADDER_EPS) return rung;
  }
  return Math.min(floor, ZOOM_STEP_LADDER[0]);
}

/** Clamp `scale` to `[min, max]`. Pure; `max < min` yields `min` (degenerate
 *  bounds resolve to the floor, matching the viewers' inline clamps). */
export function clampScale(scale: number, min: number, max: number): number {
  return scale < min ? min : scale > max ? max : scale;
}

/** Content + container extents (px, same axis) for a fit computation. */
export interface FitInput {
  /** Natural content width at 100% (CSS px). */
  contentWidth: number;
  /** Natural content height at 100% (CSS px). */
  contentHeight: number;
  /** Available container width (CSS px). */
  containerWidth: number;
  /** Available container height (CSS px). Only used by {@link fitScale} in
   *  `'page'` mode; a `'width'` fit ignores it. */
  containerHeight: number;
}

/** Which dimension a fit targets: `'width'` fits the width only (height scrolls);
 *  `'page'` fits both so the whole page is visible. */
export type FitMode = 'width' | 'page';

/**
 * The user-facing zoom factor that fits the content to the container.
 *
 * - `'width'`: `containerWidth / contentWidth` — the page spans the width, height
 *   overflows into the scroll region.
 * - `'page'`: `min(containerWidth / contentWidth, containerHeight / contentHeight)`
 *   — the whole page fits inside the box (letterboxed on the looser axis).
 *
 * Returns `0` when any input needed for the chosen mode is non-positive (an
 * unlaid-out container or empty content) so the caller can DEFER — the same
 * zero-as-defer convention the scroll viewers already use for their base fit.
 * Pure; no DOM, no clamping (the caller clamps to `[zoomMin, zoomMax]`).
 */
export function fitScale(input: FitInput, mode: FitMode): number {
  const { contentWidth, contentHeight, containerWidth, containerHeight } = input;
  if (contentWidth <= 0 || containerWidth <= 0) return 0;
  const widthFit = containerWidth / contentWidth;
  if (mode === 'width') return widthFit;
  if (contentHeight <= 0 || containerHeight <= 0) return 0;
  const heightFit = containerHeight / contentHeight;
  return Math.min(widthFit, heightFit);
}
