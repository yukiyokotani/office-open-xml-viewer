import { EMU_PER_PX, zoomStepScale, anchoredZoomOffset, nextZoomStep, prevZoomStep, fitScale, type FindHighlightColors, type FindMatch, type FindMatchesOptions, type HyperlinkTarget, type OoxmlResourceMetrics, type ViewerContextMenuEvent, type ZoomableViewer, openExternalHyperlink } from '@silurus/ooxml-core';
import {
  computeUniformVisibleWindow,
  type VisibleWindow,
} from '@silurus/ooxml-core/internal/virtual-scroll';
import {
  createCanvasElementOutlineLayer,
  CanvasViewerErrorRouter,
  renderCanvasElementOutline,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import { eventTargetsDataAttributeWithin } from '@silurus/ooxml-core/internal/dom-interaction-boundary';
import type { ReadOnlyCommentMarginGeometry } from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import { PptxPresentation, type LoadOptions, type RenderSlideOptions } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import type { PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { PptxFindController, type PptxMatchLocation } from './find';
import { buildPptxHighlightLayer } from './find-highlight-layer';
import {
  createPptxCommentSelectionContext,
  readPptxTextSelectionContext,
} from './selection-context';
import type {
  PptxElementBounds,
  PptxElementContext,
  PptxSelectionContext,
  PptxSelectionContextOptions,
} from './element-selection';
import {
  limitPptxElementContext,
  MAX_ELEMENT_TEXT_CHARACTERS,
} from './element-selection';
import type { PptxCommentsOptions } from './comment-margin';
import { pptxCommentOccurrenceKey } from './comment-occurrence';
import type { PptxComment } from './types';
import { renderPptxFocusedSlide } from './focused-view-runtime';
import {
  subscribePptxLayout,
  type PptxLayoutPublication,
} from './presentation-layout-events';

/**
 * Debounce window (ms) after the last `setScale` in a zoom burst before the
 * full-resolution settle re-render is dispatched (design §7 "Flicker-free zoom").
 *
 * This is a UI-INTERACTION-FEEL policy constant, NOT an ECMA-376 / ISO-29500
 * value: it exists only so a rapid wheel/pinch gesture (which fires dozens of
 * `setScale` calls) coalesces into a single high-res render at the end instead of
 * re-rendering per tick. Each `setScale` shows an immediate CSS preview (the
 * existing bitmap stretched) and resets this timer; the settle fires once the
 * gesture pauses for `ZOOM_SETTLE_MS`. Lower = snappier but more redundant renders
 * mid-gesture; higher = fewer renders but a longer soft-preview tail. Deliberately
 * duplicated per viewer (a one-line timing constant, not shared logic).
 */
const ZOOM_SETTLE_MS = 150;

/**
 * Default CSS `box-shadow` painted on every slide canvas — the soft drop shadow a
 * presentation viewer casts under each slide (matches the Examples/recipe look,
 * which the scroll viewer now reproduces with zero config). See
 * {@link PptxScrollViewerOptions.pageShadow}.
 */
const DEFAULT_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';
const COMMENT_MARGIN_GAP_PX = 12;
type PptxCommentUiRuntime = typeof import('./comment-ui-runtime.js');
let pptxCommentUiRuntimePromise: Promise<PptxCommentUiRuntime> | undefined;

function loadPptxCommentUiRuntime(): Promise<PptxCommentUiRuntime> {
  return pptxCommentUiRuntimePromise ??= import('./comment-ui-runtime.js');
}
// Presentation slides have a wider natural canvas than a DOCX page. Give the
// built-in PPTX review margin its own baseline so it does not become
// disproportionately small when the composite slide + margin is fit to width.
const COMMENT_MARGIN_WIDTH_PX = 440;
const COMMENT_MARGIN_FONT_SIZE_PX = 20;
const borrowedPresentationOption = Symbol('PptxScrollViewer.borrowedPresentation');
type InternalPptxScrollViewerOptions = PptxScrollViewerOptions & {
  [borrowedPresentationOption]?: PptxPresentation;
};

/**
 * Options for {@link PptxScrollViewer}. Only the `width` and `dpr` per-slide
 * render knobs apply to this virtualized Viewer; it owns text-run collection,
 * media controls, and hidden-slide dimming itself.
 */
export interface PptxScrollViewerOptions extends Pick<RenderSlideOptions, 'width' | 'dpr'>, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive slides. Default 16. */
  gap?: number;
  /** Desk padding (px) ABOVE the FIRST slide — the margin a presentation viewer
   *  leaves between the top of the scroll surface and the first slide. Default:
   *  `gap` (uniform desk rhythm — the first slide sits the same distance from the
   *  top as slides sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST slide — the margin below the final slide.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Desk gutter (px) to the LEFT of the slides — the horizontal margin between
   *  the left edge of the scroll surface and a slide sitting flush-left (i.e. once
   *  zoomed wide enough that centering no longer applies). Default: `gap` (uniform
   *  desk rhythm — the horizontal gutters match the vertical ones). It also shrinks
   *  the container-derived FIT width so a slide sits inside the gutters at 100%
   *  (an EXPLICIT `opts.width` is the slide's CSS-width contract and is NOT reduced;
   *  the gutters still apply around placement). Pass `0` for a flush-left layout. */
  paddingLeft?: number;
  /** Desk gutter (px) to the RIGHT of the slides. Default: `gap`. Shrinks the
   *  container-derived fit width symmetrically with `paddingLeft`. Pass `0` for a
   *  flush-right layout. */
  paddingRight?: number;
  /** Slides kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /** Per-slide transparent text-selection overlay. IX6 — works in BOTH render
   *  modes: in worker mode the per-run geometry is collected off-thread and
   *  shipped back beside the slide bitmap, so the overlay is populated identically
   *  to main mode (no more empty overlay / one-time warning). */
  enableTextSelection?: boolean;
  /** Show the built-in read-only comments. Pass options to configure them. Default false. */
  comments?: boolean | PptxCommentsOptions;
  /** Enable read-only slide-element selection with a non-editable outline. Default false. */
  enableElementSelection?: boolean;
  /** Straight-line hit tolerance in CSS pixels. Default 6. */
  elementHitTolerance?: number;
  /** Emits bounded, detached text or element context for read-only AI/MCP use. */
  onSelectionContextChange?: (context: PptxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the text or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<PptxSelectionContext>) => void;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /**
   * Enable interactive audio/video playback. When true, mounted slides render
   * through {@link PptxPresentation.presentSlide} only while they are inside the
   * real viewport plus {@link mediaOverscan}. Other mounted slides retain static
   * canvases and selectable text without allocating media blobs or RAF loops.
   * Default false.
   */
  enableMediaPlayback?: boolean;
  /**
   * Slides beyond the real viewport that may keep interactive media handles.
   * Independent from {@link overscan}, so an integration may mount every text
   * overlay for browser-native Find while media resources stay bounded. Default 1.
   */
  mediaOverscan?: number;
  /** Minimum zoom scale — a DIMENSIONLESS multiplier over the 96-dpi natural
   *  slide size (10% = 0.1), matching `DocxScrollViewer`. Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale (dimensionless multiplier, 400% = 4). Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * Re-fit the presentation to the container width when the container is
   * resized. Default true. Set false to preserve the current absolute scale,
   * including an explicit pre-load `setScale(1)`, independently of the viewport
   * width. Explicit `fitWidth()` and `fitPage()` calls remain available.
   */
  refitOnResize?: boolean;
  /**
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between slides — the gray a presentation viewer paints around the
   * slide. Applied to the viewer-owned scroll host. The slides themselves are
   * always drawn on their own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * CSS `box-shadow` painted on every slide CANVAS (not the wrapper — the
   * text-selection overlay must not cast its own shadow). The soft drop shadow a
   * presentation viewer leaves under each slide.
   *
   * - Default (`undefined`): `'0 1px 3px rgba(0,0,0,0.2)'` — the recipe look, so
   *   the scroll viewer reproduces the Examples appearance with zero config.
   * - `false`: NO shadow (flat slides).
   * - A custom string is applied verbatim. A spread-only ring such as
   *   `'0 0 0 1px #c8ccd0'` gives a crisp 1px BORDER look — and because
   *   `box-shadow` never affects layout (unlike `border`, which would grow the
   *   box and shift every offset), a border and a drop shadow are the SAME knob
   *   here rather than two competing options.
   */
  pageShadow?: string | false;
  /** Fires when the top-most visible slide changes. `topIndex` from
   *  `computeVisibleRange` (the first slide intersecting the viewport top,
   *  EXCLUDING overscan). */
  onVisibleSlideChange?: (topIndex: number, total: number, layoutComplete: boolean) => void;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100% = a slide
   *  at its natural EMU→px size): from {@link PptxScrollViewer.setScale},
   *  `zoomIn`/`zoomOut`, `fitWidth`/`fitPage`, a Ctrl/⌘+wheel gesture, or a
   *  container-resize re-fit (when `refitOnResize` is enabled). Named
   *  `onScaleChange` to match the single-canvas viewers so all five share one
   *  notification shape. */
  onScaleChange?: (scale: number) => void;
  /** Receives asynchronous Viewer-managed failures that cannot be observed by
   *  awaiting the method that started them. `load()` failures always reject and
   *  are not also delivered here. Virtualized per-slot render failures (both
   *  main `renderSlide` and worker `renderSlideToBitmap` rejections) and
   *  embedded-media fetch/decode/playback failures invoke it. A failed slide is
   *  left blank rather than crashing the loop.
   *  Without an `onError`, failures are logged via `console.error` so they are
   *  never fully silent. Stable cases can be narrowed with `OoxmlError`,
   *  `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   *  this package. Other failures remain `Error` values; a `code` of
   *  `parser-crashed` identifies a recognized WASM trap, not a reliably
   *  classified OOM. */
  onError?: (err: Error) => void;
  /**
   * IX1 (design decision — NOT user-confirmed, integrator may veto). Fires on a
   * hyperlink click in any mounted slide's text overlay (requires
   * {@link enableTextSelection}). Default when omitted: external →
   * {@link openExternalHyperlink} (new tab, sanitised, noopener); internal
   * slide-jump → {@link scrollToSlide} once the action resolves to a slide index
   * via {@link PptxPresentation.resolveInternalTarget} (a jump that resolves to
   * no reachable slide is a safe no-op). When provided, the viewer calls this
   * instead and takes NO default action.
   */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the hyperlink machinery is not wired at all: the overlay's link
   *  spans are non-interactive, so there is no pointer cursor, no title tooltip,
   *  no default navigation (external new-tab / internal slide jump), and
   *  `onHyperlinkClick` is never called. Links still render exactly as authored
   *  but are inert, like plain text. */
  enableHyperlinks?: boolean;
}

/** One mounted slide. `canvas` is the drawn slide; `textLayer` the optional
 *  per-slide selection overlay (both render modes — IX6 ships the worker's run
 *  geometry back beside the bitmap). `renderedSlide` guards against
 *  re-rendering a recycled slot for a slide whose render is still in flight. */
interface SlideSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  highlightLayer: HTMLDivElement;
  elementLayer: HTMLDivElement | null;
  loadingLayer: HTMLSpanElement;
  commentMarkerLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentDecorationLayer: HTMLDivElement | null;
  commentElementBounds: readonly PptxElementBounds[];
  commentGeometry: ReadOnlyCommentMarginGeometry | null;
  commentAnchorSlide: number;
  commentAnchorGeneration: number;
  /** slide index this slot is currently rendering / has rendered, or -1 when free. */
  renderedSlide: number;
  /** The `_scale` at which this slot's on-screen canvas bitmap (and text overlay)
   *  were last rendered, or -1 when unrendered. The flicker-free CSS preview
   *  (design §7) stretches that bitmap to the new layout size on `setScale` and
   *  scales the text overlay by `newScale / renderedScale`; the debounced settle
   *  re-render then repaints at the new scale and updates this to match. */
  renderedScale: number;
  /** Shared single-canvas generation and worker-bitmap ownership primitive. */
  dispatcher: StaticCanvasRenderDispatcher;
  /** Interactive media handle for the canvas currently mounted in this slot. */
  presentationHandle: PresentationHandle | null;
  /** Whether this slot currently owns or awaits an interactive media handle. */
  mediaInteractive: boolean;
  /** Per-slot paint generation across static and interactive canvas ownership. */
  renderGeneration: number;
  /**
   * Per-slot async generation. Unlike the viewer render epoch, this also changes
   * when a pooled slot is recycled and immediately reused for the same slide
   * index, so a late presentSlide() result can never attach to the new owner.
   */
  presentationGeneration: number;
}

export class PptxScrollViewer implements ZoomableViewer {
  private readonly _presentationOwner: TerminalResourceOwner<PptxPresentation>;
  private get _pres(): PptxPresentation | null { return this._presentationOwner.current; }
  private readonly _borrowed: boolean;
  private readonly _opts: PptxScrollViewerOptions;
  private readonly _errorRouter: CanvasViewerErrorRouter;
  private readonly _container: HTMLElement;
  private readonly _wrapper: HTMLDivElement;
  private readonly _scrollHost: HTMLDivElement;
  private readonly _spacer: HTMLDivElement;
  /** Resolved render mode. When an engine is borrowed the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `PptxPresentation.load`. */
  private _mode: 'main' | 'worker';

  /** Dimensionless zoom multiplier over the 96-dpi natural slide size (mirrors
   *  `DocxScrollViewer`, whose `_scale` multiplies `widthPt × PT_TO_PX`). The
   *  natural (1×) slide width in CSS px is `slideEmu / EMU_PER_PX`; the base fit
   *  sets `_scale` so that natural width maps to the container width, and zoom
   *  multiplies it further (design §7). */
  private _scale = 1;
  /** Whether the base fit scale has been established. Set true the first time
   *  `relayout()` resolves a positive base scale. We use an explicit flag rather
   *  than a `_scale === 1` sentinel because a fit scale of exactly 1 is a valid
   *  established state (a 1× fit would otherwise be re-fit forever). */
  private _scaleEstablished = false;
  /**
   * IX9 F1 — a `setScale` factor requested BEFORE the base fit is established
   * (pre-load, or a zero-width container), already clamped to
   * `[zoomMin, zoomMax]`, or `null` when none is pending. The single-canvas
   * viewers latch a pre-load `setScale` and honour it on the first render; the
   * scroll viewers used to silently DROP it — the family-unified semantics are
   * "latch and apply once the layout establishes". `relayout()` applies (and
   * clears) this right after establishing the base, firing `onScaleChange` at
   * application time; `getScale()` reports it while pending so the caller sees
   * the same value a single-canvas viewer would show.
   */
  private _pendingScale: number | null = null;
  /** Live slots keyed by slide index. */
  private readonly _slots = new Map<number, SlideSlot>();
  /** Recyclable detached slots (canvas + textLayer reused across slides). */
  private readonly _free: SlideSlot[] = [];
  /** Uniform slide height at the current scale. Keeping the scalar avoids both
   * the document-length height and offset arrays in every scroll query. */
  private _uniformSlideHeight = 0;
  private _lastRange: VisibleWindow | null = null;
  private _lastTopIndex = -1;
  private _lastReportedTotal = -1;
  private _lastReportedLayoutComplete: boolean | null = null;
  private _layoutUnsubscribe: (() => void) | null = null;
  private _scrollListener: (() => void) | null = null;
  private _selectionChangeListener: (() => void) | null = null;
  private _selectionContextKey = 'null';
  private _elementClickListener: ((event: MouseEvent) => void) | null = null;
  private _contextMenuListener: ((event: MouseEvent) => void) | null = null;
  private _commentOutsidePointerListener: ((event: PointerEvent) => void) | null = null;
  private _elementContext: PptxElementContext | null = null;
  private _activeCommentId: string | null = null;
  private _activeCommentSlide: number | null = null;
  private _commentNavigationGeneration = 0;
  private _commentUi: PptxCommentUiRuntime | null = null;
  private _commentGeometryScheduled = false;
  private _commentGeometryFrame: number | null = null;
  private readonly _pendingCommentGeometry = new Map<number, {
    readonly slot: SlideSlot;
    readonly connectorsOnly: boolean;
  }>();
  private _hasComments = false;
  /** Horizontal origin used only for a reachable left-hand review rail. */
  private _reviewOriginPx = 0;
  /** Opening prefix already inspected for authored comments. Progressive
   * presentations make metadata authoritative one slide at a time, so a
   * negative scan is provisional until this frontier reaches slideCount. */
  private _commentScanFrontier = 0;
  private readonly _layoutWaiters = new Set<() => void>();
  private _layoutFailed = false;
  private _elementHitGeneration = 0;
  private readonly _elementHitTolerance: number;
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  /** Worker mode: slide indices whose bitmap render is currently dispatched to the
   *  engine. Coalesces a scroll storm — we never dispatch a second render for a
   *  slide whose first is still in flight — and lets us drop slides that scrolled
   *  out of the window before dispatch (design §11 worker coalescing).
   *
   *  T4 ZOOM HAZARD (RESOLVED by the render epoch below): coalescing keys on slide
   *  INDEX only, with no notion of the scale a dispatch was made at. Once
   *  `setScale` can change the zoom mid-flight, an in-flight bitmap dispatched at
   *  the OLD scale can still pass the on-resolution identity check if the SAME
   *  slot object is re-mounted for slide `i` (the pool reuses slot objects, so
   *  `_slots.get(i) === slot && slot.renderedSlide === i` can hold for an old
   *  dispatch), and get painted at the WRONG resolution. We fix this with a render
   *  epoch (`_renderEpoch`): each dispatch captures the epoch, and on resolution a
   *  moved epoch ⇒ STALE (close + re-dispatch the live slot). See
   *  `_renderSlotBitmap`. */
  private readonly _slideInFlight = new Set<number>();
  /** Render generation, bumped on every effective `setScale` (and the resize
   *  re-fit in `_onResize`, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private _renderEpoch = 0;
  /** Pending settle-render timer handle (design §7 mechanism 2). Set by
   *  `_scheduleSettle` after each `setScale`, reset on the next one so a burst
   *  dispatches ONE settle at the end, and cleared in `destroy()`. `ReturnType`
   *  of `setTimeout` (a number in the DOM, a Timeout object in node) so the type
   *  is host-agnostic. */
  private _settleTimer: ReturnType<typeof setTimeout> | null = null;
  private _wheelListener: ((e: WheelEvent) => void) | null = null;
  /** Gesture-only pointer anchor for the NEXT `setScale`, in scrollHost-viewport
   *  px (`{ x, y }` from the wheel event, relative to the scroll host's top-left).
   *  Set by the Ctrl/⌘+wheel handler right before it calls `setScale` so the zoom
   *  pivots on the cursor ("zoom toward the pointer") in BOTH axes; consumed and
   *  cleared by `setScale`. `null` for every non-gesture source (the public
   *  `setScale`, the +/- steppers, `fitWidth`/`fitPage`, the resize re-fit), which
   *  keep the historical viewport-TOP re-anchor so their behaviour is unchanged. */
  private _pendingZoomAnchor: { x: number; y: number } | null = null;
  /** Observes the container so a width change re-fits the base scale. Disconnected
   *  in `destroy()`. */
  private _resizeObserver: ResizeObserver | null = null;
  /** The base fit scale at the last established/re-fit layout. `_onResize` divides
   *  `_scale` by this to recover the current zoom multiplier so a width change
   *  re-fits the base while preserving the user's zoom (design §11). */
  private _prevBase = 0;
  /** The fit width (px) the base scale was last established at. Lets `_onResize`
   *  skip the re-fit when only the height changed (a ResizeObserver fires on ANY
   *  box change, but only a WIDTH change alters the fit-to-width base scale). */
  private _lastFitWidth = 0;
  /** Resolved slide-canvas `box-shadow` (design: the recipe drop shadow by
   *  default). Resolved ONCE with `??` — NOT `||` — so `pageShadow: false`
   *  survives as the "no shadow" sentinel (a `||` would treat `false` as absent
   *  and wrongly re-apply the default). Applied by `_applyPageShadow` at EVERY
   *  canvas-creation site (`_acquireSlot` and the double-buffer spare in
   *  `_settleSlot`) so a recycled/re-mounted slot and a settle-swapped spare all
   *  carry it. */
  private readonly _pageShadow: string | false;
  private readonly _find = new PptxFindController(
    () => this.slideCount,
    (slide) => this._collectSlideRuns(slide),
  );
  private _findGeneration = 0;
  private _findActive = false;
  private _findMeasureCtx: CanvasRenderingContext2D | null | undefined;

  /**
   * Create a Scroll Viewer that borrows an already-loaded presentation.
   *
   * The presentation's render mode is authoritative. The returned Viewer
   * cannot load another source, and destroying it leaves the caller-owned
   * presentation open. The initial virtual window is laid out during
   * construction.
   */
  static fromPresentation(
    container: HTMLElement,
    presentation: PptxPresentation,
    opts: Omit<PptxScrollViewerOptions, keyof LoadOptions> = {},
  ): Omit<PptxScrollViewer, 'load'> {
    return new PptxScrollViewer(container, {
      ...opts,
      [borrowedPresentationOption]: presentation,
    } as InternalPptxScrollViewerOptions);
  }

  constructor(container: HTMLElement, opts: PptxScrollViewerOptions = {}) {
    // A <canvas> is an HTMLElement too, so the type system cannot stop a caller
    // used to the pager API (PptxViewer takes a canvas) from passing one — but
    // canvas children never render, so the viewer would come up silently blank.
    // Fail loudly with the fix instead. (tagName, not instanceof: cross-realm safe.)
    if (container.tagName === 'CANVAS') {
      throw new Error(
        'PptxScrollViewer takes a container element (e.g. a <div>), not a <canvas> — ' +
          'the viewer creates and manages its own canvases. Pass a block container; ' +
          'for the single-slide canvas API use PptxViewer.',
      );
    }
    this._container = container;
    this._opts = opts;
    this._errorRouter = new CanvasViewerErrorRouter('PptxScrollViewer', opts.onError);
    const elementHitTolerance = opts.elementHitTolerance ?? 6;
    if (!Number.isFinite(elementHitTolerance) || elementHitTolerance < 0) {
      throw new RangeError('elementHitTolerance must be a finite non-negative number.');
    }
    this._elementHitTolerance = elementHitTolerance;
    // `??` (not `||`): a caller's explicit `false` must disable the shadow, not
    // fall through to the default.
    this._pageShadow = opts.pageShadow ?? DEFAULT_PAGE_SHADOW;
    const borrowedPresentation = (opts as InternalPptxScrollViewerOptions)[borrowedPresentationOption];
    this._borrowed = borrowedPresentation !== undefined;
    if (borrowedPresentation) {
      this._presentationOwner = new TerminalResourceOwner('PptxScrollViewer', borrowedPresentation, false);
      this._mode = resolveCanvasViewerMode('PptxScrollViewer', opts.mode, borrowedPresentation);
      this._scanAvailableComments(borrowedPresentation, false);
    } else {
      this._presentationOwner = new TerminalResourceOwner('PptxScrollViewer');
      this._mode = resolveCanvasViewerMode('PptxScrollViewer', opts.mode, undefined);
    }

    // container → wrapper → scrollHost → spacer  (design §6)
    this._wrapper = document.createElement('div');
    this._wrapper.style.cssText = 'position:relative;width:100%;height:100%;overflow:hidden;';
    this._scrollHost = document.createElement('div');
    // Reserve the classic vertical scrollbar gutter before content overflows.
    // Together with `_fitWidthPx` reading this scrollport's clientWidth, this
    // prevents a vertical scrollbar from stealing width after the initial fit
    // and creating a small, unintended horizontal overflow.
    this._scrollHost.style.cssText = 'position:absolute;inset:0;overflow:auto;';
    this._scrollHost.style.scrollbarGutter = 'stable';
    // The "desk" behind/between slides. Undefined ⇒ transparent (container shows
    // through); slides keep their own white canvas regardless.
    if (opts.background) this._scrollHost.style.background = opts.background;
    this._spacer = document.createElement('div');
    this._spacer.style.cssText = 'position:absolute;top:0;left:0;width:1px;height:0;pointer-events:none;';
    this._scrollHost.appendChild(this._spacer);
    this._wrapper.appendChild(this._scrollHost);
    this._container.appendChild(this._wrapper);

    if (this._commentsEnabled()) {
      void loadPptxCommentUiRuntime().then((commentUi) => {
        if (this._destroyed) return;
        this._commentUi = commentUi;
        for (const [slide, slot] of this._slots) this._redrawSlotComments(slide, slot);
      }).catch((error) => this._reportRenderError(error));
    }

    if (opts.enableTextSelection && (opts.onSelectionContextChange || opts.enableElementSelection)) {
      this._selectionChangeListener = () => this._emitSelectionContextChange();
      this._wrapper.ownerDocument.addEventListener('selectionchange', this._selectionChangeListener);
    }
    if (opts.enableElementSelection) {
      this._elementClickListener = (event) => {
        void this._onElementClick(event).catch((error) => this._reportRenderError(error));
      };
      this._scrollHost.addEventListener('click', this._elementClickListener);
    }
    if (opts.onContextMenu) {
      this._contextMenuListener = (event) => this._onContextMenu(event);
      this._scrollHost.addEventListener('contextmenu', this._contextMenuListener);
    }

    this._scrollListener = () => this._onScroll();
    this._scrollHost.addEventListener('scroll', this._scrollListener);

    if (opts.comments) {
      this._commentOutsidePointerListener = (event) => {
        if (eventTargetsDataAttributeWithin(event, this._wrapper, 'ooxmlCommentId')) return;
        if (this._activeCommentId === null) return;
        this._activeCommentId = null;
        this._activeCommentSlide = null;
        for (const [slide, slot] of this._slots) this._redrawSlotComments(slide, slot);
        this._emitSelectionContextChange();
      };
      this._wrapper.ownerDocument.addEventListener('pointerdown', this._commentOutsidePointerListener);
    }

    // Ctrl/Cmd+wheel zoom (design §7). Bare wheel is left untouched so the
    // scrollHost scrolls natively. `enableZoom:false` installs no handler at all.
    // `{ passive: false }` is required because we call preventDefault() to stop
    // the browser's own ctrl+wheel page zoom.
    if (this._opts.enableZoom !== false) {
      this._wheelListener = (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) return; // bare wheel scrolls natively
        e.preventDefault();
        if (e.deltaY === 0) return;
        // Pointer-anchored zoom: pivot on the cursor, not the viewport top. Record
        // the pointer in scrollHost-viewport px (subtract the host's on-screen
        // origin) so `setScale` can keep the content point under the cursor fixed.
        // A malformed event (no clientX/Y) yields a non-finite anchor; drop it so
        // `setScale` falls back to the historical viewport-top re-anchor.
        const rect = this._scrollHost.getBoundingClientRect();
        const ax = e.clientX - rect.left;
        const ay = e.clientY - rect.top;
        this._pendingZoomAnchor =
          Number.isFinite(ax) && Number.isFinite(ay) ? { x: ax, y: ay } : null;
        this.setScale(zoomStepScale(this._scale, e.deltaY));
      };
      this._scrollHost.addEventListener('wheel', this._wheelListener as EventListener, {
        passive: false,
      });
    }

    // Re-fit the base scale on a container resize (design §11). A container that
    // is 0-wide at construction (a common flexbox/tab layout) establishes its
    // scale on the first non-zero resize — the zero-width deferral is completed
    // here. `ResizeObserver` may be absent in a non-DOM host; guard for it.
    if (typeof ResizeObserver !== 'undefined') {
      this._resizeObserver = new ResizeObserver(() => this._onResize());
      this._resizeObserver.observe(this._container);
    }

    if (this._borrowed) {
      this._bindLayoutPresentation(borrowedPresentation!);
      // A borrowed engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path re-runs it once width appears).
      this.relayout();
    }
  }

  /**
   * Load a PPTX from URL or ArrayBuffer and render the first window.
   * Unsupported on a Viewer created by {@link fromPresentation}; the caller
   * already owns the parsed engine.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
    if (this._borrowed) {
      throw new Error(
        'PptxScrollViewer.load() is unsupported on a Viewer created by fromPresentation(); ' +
          'the borrowed presentation is already loaded.',
      );
    }
    // SC20 atomic swap: a self-loaded viewer OWNS its engine, so a re-load must
    // not orphan the previous one.
    // Retain it locally and free it only after the new engine loads — a FAILED
    // re-load then keeps the current deck rendered rather than going blank. (The
    // borrowed path returned above can never reach here, so this only ever frees
    // an engine we created.)
    let selectionInvalidated = false;
    try {
      const pres = await this._presentationOwner.replace(() => PptxPresentation.load(source, {
        password: this._opts.password,
        useGoogleFonts: this._opts.useGoogleFonts,
        maxZipEntryBytes: this._opts.maxZipEntryBytes,
        resourceLimits: this._opts.resourceLimits,
        debug: this._opts.debug,
        onResourceMetrics: this._opts.onResourceMetrics,
        workerTimeoutMs: this._opts.workerTimeoutMs,
        wasmUrl: this._opts.wasmUrl,
        math: this._opts.math,
        threeD: this._opts.threeD,
        regionMap: this._opts.regionMap,
        chartEx: this._opts.chartEx,
        mode: this._mode,
        progressiveLayout: this._opts.progressiveLayout,
        onLayoutProgress: this._opts.onLayoutProgress,
        onLayoutPartial: this._opts.onLayoutPartial,
        onLayoutComplete: this._opts.onLayoutComplete,
      }), (ownedPresentation) => {
        // Invalidate before TerminalResourceOwner installs the candidate and
        // destroys the prior worker, whose pending hit requests reject on close.
        this._invalidateElementSelection(false);
        selectionInvalidated = true;
        this._invalidateFind();
        this._findActive = false;
        this._activeCommentId = null;
        this._activeCommentSlide = null;
        this._hasComments = false;
        this._commentScanFrontier = 0;
        this._beginCommentNavigation();
        this._unbindLayoutPresentation();
        if (ownedPresentation) {
          for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
          this._lastTopIndex = -1;
        }
      });
      if (!pres) return;
      if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
      // A successful reload replaces the selection surface. Retire hit tests
      // issued against the old engine and notify that its element focus ended.
      this._invalidateFind();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentSlide = null;
      this._hasComments = false;
      this._commentScanFrontier = 0;
      this._scanAvailableComments(pres, false);
      this._bindLayoutPresentation(pres);
      // Lay out + mount the first window now that the engine exists (mirrors the
      // borrowed-engine path in the constructor). relayout() is idempotent and
      // defers under a zero-width container — `_onResize` re-runs it once width
      // appears.
      const initialRenders: Promise<void>[] = [];
      this._relayout(initialRenders);
      await Promise.all(initialRenders);
    } catch (err) {
      if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
      throw err instanceof Error ? err : new Error(String(err));
    }
    // Notify only after the replacement has committed and relayout completed;
    // consumer callback failures are not presentation/render failures.
    if (selectionInvalidated && !this._destroyed) this._emitSelectionContextChange();
  }

  get slideCount(): number {
    return this._pres?.slideCount ?? 0;
  }

  /** Number of opening slides currently paintable under progressive layout. */
  get availableSlideCount(): number {
    return this._pres?.availableSlideCount ?? this.slideCount;
  }

  /** True only after every slide became paintable successfully. */
  get layoutComplete(): boolean {
    return this._pres?.layoutComplete ?? true;
  }

  /** Wait until every slide is paintable; rejects if progressive preparation fails. */
  async waitUntilLayoutComplete(): Promise<void> {
    await this._errorRouter.ownBackgroundLifecycle(async () => {
      await this._pres?.waitUntilLayoutComplete?.();
    });
  }

  /** Uniform slide width in CSS px at the current scale. `_scale` is a
   *  dimensionless multiplier over the natural 96-dpi width (`slideEmu /
   *  EMU_PER_PX`), mirroring docx's `widthPt × PT_TO_PX × _scale`. */
  private _slideWidthPx(): number {
    return (this._pres!.slideWidth / EMU_PER_PX) * this._scale;
  }

  /** Uniform slide height in CSS px at the current scale. */
  private _slideHeightPx(): number {
    return (this._pres!.slideHeight / EMU_PER_PX) * this._scale;
  }

  /** The fit width (px), deferring when the container is unlaid-out. An EXPLICIT
   *  `opts.width` is the slide's CSS-width contract and is returned UNCHANGED (the
   *  gutters still apply around placement, not to the width). The container-derived
   *  default instead targets `containerWidth − padL − padR` so a slide sits INSIDE
   *  the horizontal gutters at 100%. A non-positive result (gutters wider than the
   *  container) is treated as unlaid-out — the same deferral as a zero-width box. */
  private _fitWidthPx(): number {
    if (this._opts.width && this._opts.width > 0) return this._opts.width;
    // Fit to the real scrollport, not its outer container: a non-overlay vertical
    // scrollbar reduces scrollHost.clientWidth but leaves container.clientWidth
    // unchanged. The container is only a fallback for synthetic / not-yet-laid-
    // out hosts where the absolutely positioned scrollport still reports zero.
    const cw = this._scrollHost.clientWidth || this._container.clientWidth;
    if (cw <= 0) return 0; // 0 ⇒ defer (design §11 zero-width deferral)
    const { left, right } = this._padH();
    const available = cw - left - right;
    if (available <= 0) return 0;
    // Fit the authored slide itself. Review cards are an adjacent horizontal
    // surface and may extend the horizontal scroll range, but must never change
    // slide scale or vertical scroll geometry when progressive metadata reveals
    // a later comment.
    return available;
  }

  private _commentMarginExtent(): number {
    return this._hasCommentMargin()
      ? (COMMENT_MARGIN_GAP_PX + COMMENT_MARGIN_WIDTH_PX) * this._commentZoom()
      : 0;
  }

  private _hasCommentMargin(): boolean {
    return this._commentsEnabled() && this._hasComments &&
      this._commentsOptions()?.cards !== false;
  }

  /** Comment chrome uses the same absolute zoom as the rendered presentation. */
  private _commentZoom(): number {
    return this._scaleEstablished ? this._scale : 1;
  }

  private _commentsEnabled(): boolean {
    return this._opts.comments === true || typeof this._opts.comments === 'object';
  }

  private _commentsOptions(): PptxCommentsOptions | undefined {
    return typeof this._opts.comments === 'object' ? this._opts.comments : undefined;
  }

  private _commentSide(): 'left' | 'right' {
    const requested = this._commentsOptions()?.side;
    if (requested === 'left' || requested === 'right') return requested;
    const computedDirection = this._container.ownerDocument.defaultView?.getComputedStyle?.(
      this._container,
    ).direction;
    const direction = computedDirection || this._container.dir || this._container.style.direction;
    return direction === 'rtl' ? 'left' : 'right';
  }

  private _syncCommentMarginGeometry(margin: HTMLDivElement | null): void {
    if (!margin) return;
    const zoom = this._commentZoom();
    const offset = `calc(100% + ${COMMENT_MARGIN_GAP_PX * zoom}px)`;
    margin.style.left = this._commentSide() === 'right' ? offset : '';
    margin.style.right = this._commentSide() === 'left' ? offset : '';
    margin.style.width = `${COMMENT_MARGIN_WIDTH_PX * zoom}px`;
    margin.style.fontSize = `${COMMENT_MARGIN_FONT_SIZE_PX}px`;
    margin.dataset.ooxmlCommentZoom = String(zoom);
  }

  /** Inspect only the prefix whose slide metadata is authoritative. When a
   * later publication reveals the first comment, enable the already-present
   * review layers and horizontal extent without replacing the authored canvas
   * or changing fit/vertical geometry. */
  private _scanAvailableComments(
    presentation: PptxPresentation,
    rebuildMountedSurface: boolean,
  ): void {
    if (!this._commentsEnabled() || this._hasComments) return;
    const includeResolved = this._commentsOptions()?.includeResolved === true;
    const available = Math.min(presentation.availableSlideCount, presentation.slideCount);
    for (let i = this._commentScanFrontier; i < available; i++) {
      if (!presentation.getComments(i).some((comment) => includeResolved ||
        (comment.status !== 'resolved' && comment.status !== 'closed'))) continue;
      this._commentScanFrontier = available;
      this._hasComments = true;
      if (rebuildMountedSurface) this._refreshDiscoveredComments();
      return;
    }
    this._commentScanFrontier = Math.max(this._commentScanFrontier, available);
  }

  private _refreshDiscoveredComments(): void {
    // Comment-enabled slots own empty review layers from birth. Revealing later
    // metadata therefore updates only those layers: the painted canvas,
    // dispatcher, slide scale, stride, scrollTop, and spacer height stay intact.
    this._syncSpacerWidth();
    for (const [slide, slot] of this._slots) this._redrawSlotComments(slide, slot);
  }

  /** Base scale: the DIMENSIONLESS multiplier that fits the (uniform) slide
   *  width to the fit-width. `natural = slideWidthEmu / EMU_PER_PX` is the 96-dpi
   *  CSS-px width; `base = fitWidth / natural` (mirrors docx's `w / (widthPt ×
   *  PT_TO_PX)`). Returns 0 when the container has no width yet (deferral). */
  private _baseScale(): number {
    if (!this._pres || this._pres.slideCount === 0) return 0;
    const w = this._fitWidthPx();
    const naturalW = this._pres.slideWidth / EMU_PER_PX;
    if (w <= 0 || naturalW <= 0) return 0;
    return w / naturalW; // dimensionless multiplier over the natural width
  }

  /**
   * Recompute per-slide heights + the spacer and re-mount the visible window.
   *
   * The viewer already calls this automatically after `load()`, a borrowed
   * engine, a container resize, and a zoom, so most integrations never need it.
   * It is public as a deliberate escape hatch: if the host mutates the layout in
   * a way the `ResizeObserver` cannot observe (e.g. a CSS change on an ancestor
   * that resizes the container without a box-size event, or a font that finishes
   * loading after first paint), call `relayout()` to force a re-fit. Idempotent —
   * safe to call repeatedly, and a no-op while the container has zero width (the
   * fit is deferred until width appears, design §11).
   */
  relayout(): void {
    this._relayout();
  }

  /** Synchronous geometry/layout pass. When `initialRenders` is supplied by
   * load(), newly-mounted slot Promises are collected for direct rejection
   * instead of being routed through the background onError channel. */
  private _relayout(initialRenders?: Promise<void>[]): void {
    if (!this._pres) return;
    // Establish the base fit scale on the first layout that has a positive
    // width. Zoom (T4) layers its own multiplier on top of this; here we only
    // set the base. An explicit `_scaleEstablished` flag (NOT a `_scale === 1`
    // sentinel) so a legitimate 1× fit is not re-fit on every relayout.
    if (!this._scaleEstablished) {
      const base = this._baseScale();
      if (base > 0) {
        this._scale = base;
        this._prevBase = base;
        this._lastFitWidth = this._fitWidthPx();
        this._scaleEstablished = true;
        // IX9 F1: apply a setScale latched BEFORE establishment (pre-load / a
        // zero-width container), now that the base exists. Applied here — before
        // heights/spacer/mount below — so the first window renders directly at
        // the requested factor (no intermediate base-scale frame). `_prevBase`
        // stays the true base so a later resize re-fit preserves the implied
        // zoom multiplier. onScaleChange fires at application time (the latch
        // itself was silent), and only when the pending factor actually moved
        // the scale off the base fit.
        if (this._pendingScale !== null) {
          const pending = this._pendingScale;
          this._pendingScale = null;
          if (pending !== this._scale) {
            this._scale = pending;
            this._opts.onScaleChange?.(pending);
          }
        }
      } else {
        return; // container has no width yet — retry on the next resize
      }
    }
    this._recomputeHeights();
    this._syncSpacer();
    this._mountVisible(initialRenders);
  }

  /** Refresh the uniform scale-dependent height without allocating per-slide state. */
  private _recomputeHeights(): void {
    this._uniformSlideHeight = this._slideHeightPx();
  }

  private _gap(): number {
    return this._opts.gap ?? 16;
  }

  private _overscan(): number {
    return this._opts.overscan ?? 1;
  }

  /** Media lifecycle window, deliberately independent from text/canvas overscan. */
  private _mediaOverscan(): number {
    return this._opts.mediaOverscan ?? 1;
  }

  /** Desk padding fed to `computeVisibleRange`: `paddingTop`/`paddingBottom`,
   *  each defaulting to `gap` (uniform rhythm). Resolved here (not stored) to
   *  mirror `_gap()`/`_overscan()`, and consumed at EVERY `computeVisibleRange`
   *  call site so the padded offsets are the single source of geometry. */
  private _pad(): { leading: number; trailing: number } {
    const gap = this._gap();
    return { leading: this._opts.paddingTop ?? gap, trailing: this._opts.paddingBottom ?? gap };
  }

  /** Horizontal desk gutters: `paddingLeft`/`paddingRight`, each defaulting to
   *  `gap` (uniform rhythm — the horizontal gutters match the vertical padding).
   *  Consumed by `_fitWidthPx` (to shrink the container-derived fit), by
   *  `_positionSlot` (the flush-left floor), and by `_syncSpacerWidth` (the spacer
   *  width). Resolved here (not stored) to mirror `_gap()`/`_pad()`. */
  private _padH(): { left: number; right: number } {
    const gap = this._gap();
    return { left: this._opts.paddingLeft ?? gap, right: this._opts.paddingRight ?? gap };
  }

  private _slideOffset(index: number): number {
    return this._pad().leading + index * (this._uniformSlideHeight + this._gap());
  }

  /** Index of the slide spanning content-offset `y`, preserving the historical
   * convention that an inter-slide gap belongs to the preceding slide. */
  private _slideIndexAtOffset(y: number): number {
    return computeUniformVisibleWindow(
      this._pres?.slideCount ?? 0,
      this._uniformSlideHeight,
      this._gap(),
      y,
      0,
      0,
      this._pad(),
    ).topIndex;
  }

  private _rangeAt(scrollTop: number, overscan: number): VisibleWindow {
    return computeUniformVisibleWindow(
      this._pres?.slideCount ?? 0,
      this._uniformSlideHeight,
      this._gap(),
      scrollTop,
      this._scrollHost.clientHeight,
      overscan,
      this._pad(),
    );
  }

  private _range(): VisibleWindow {
    return this._rangeAt(this._scrollHost.scrollTop, this._overscan());
  }

  private _mediaRange(): VisibleWindow {
    return this._rangeAt(this._scrollHost.scrollTop, this._mediaOverscan());
  }

  private _rangeContains(r: VisibleWindow, index: number): boolean {
    return index >= r.start && index <= r.end;
  }

  private _syncSpacer(): void {
    const r = this._range();
    this._lastRange = r;
    this._spacer.style.height = `${r.totalHeight}px`;
    this._syncSpacerWidth();
  }

  /** Horizontal scroll extent: the (uniform deck-wide) slide width plus both
   *  gutters. A spacer NARROWER than the container never creates a scrollbar
   *  (scrollWidth = max(clientWidth, content)), so it is always safe to set — it
   *  only matters when a zoomed-in slide grows past the viewport, where it gives
   *  the gutters something to scroll to on either side. Called from `_syncSpacer`
   *  and after every scale change (zoom / resize re-fit) so the extent tracks the
   *  current slide px width. */
  private _syncSpacerWidth(): void {
    const { left, right } = this._padH();
    const marginExtent = this._commentMarginExtent();
    const next = this._commentSide() === 'left' ? marginExtent : 0;
    const delta = next - this._reviewOriginPx;
    const targetScrollLeft = Math.max(0, this._scrollHost.scrollLeft + delta);
    // Establish the new native scroll range before applying compensation.
    // Browsers clamp scrollLeft to the current range at assignment time.
    this._spacer.style.width = `${this._slideWidthPx() + marginExtent + left + right}px`;
    if (delta === 0) return;
    this._reviewOriginPx = next;
    (this._scrollHost.style as CSSStyleDeclaration & Record<string, string>)[
      '--ooxml-review-origin-x'
    ] = `${next}px`;
    this._scrollHost.scrollLeft = targetScrollLeft;
  }

  private _onScroll(): void {
    if (!this._pres || !this._scaleEstablished) return;
    this._mountVisible(undefined, false);
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(
    initialRenders?: Promise<void>[],
    repositionExisting = true,
  ): void {
    if (!this._pres || this._pres.slideCount === 0) return;
    const r = this._range();
    const mediaRange = this._opts.enableMediaPlayback ? this._mediaRange() : null;
    this._lastRange = r;

    // Detach slots that left [start, end] into the free pool.
    for (const [idx, slot] of [...this._slots]) {
      if (idx < r.start || idx > r.end) {
        this._recycleSlot(idx, slot);
      }
    }
    // Mount any missing index in the window.
    for (let i = r.start; i <= r.end; i++) {
      if (!this._slots.has(i)) {
        const slot = this._acquireSlot();
        this._positionSlot(slot, i, r);
        this._slots.set(i, slot);
        this._redrawSlotComments(i, slot);
        const render = this._renderSlot(
          i,
          slot,
          !!mediaRange && this._rangeContains(mediaRange, i),
          initialRenders === undefined,
        );
        // Progressive load resolves once the opening paintable prefix is on
        // screen. Later mounted slots keep their loading UI and finish from the
        // presentation's availability wait; they must not hold `load()` open.
        if (initialRenders && render && i < this.availableSlideCount) initialRenders.push(render);
      } else if (repositionExisting) {
        // Re-position (offsets shift after a spacer/height change).
        this._positionSlot(this._slots.get(i)!, i, r);
      }
    }
    if (mediaRange) this._syncMediaPlayback(mediaRange);
    // onVisibleSlideChange fires ONLY when the top visible slide actually changes
    // (change-only latch; `_lastTopIndex` starts at -1 so the first layout fires
    // once for slide 0). Every mount path — scroll, zoom, resize re-fit, and
    // scrollToSlide — funnels through here, so navigation never double-fires.
    this._emitVisibleSlideChange(r);
  }

  /** Apply the resolved slide-canvas shadow (design: recipe drop shadow by
   *  default, `false` ⇒ none). Single source so `_acquireSlot` and the
   *  double-buffer spare in `_settleSlot` stay in lock-step — a spare that missed
   *  this would lose the shadow on the settle swap. `box-shadow` never affects
   *  layout, so this is safe to (re)set on a live/pooled canvas without shifting
   *  any offset. */
  private _applyPageShadow(canvas: HTMLCanvasElement): void {
    if (this._pageShadow !== false) canvas.style.boxShadow = this._pageShadow;
  }

  private _acquireSlot(): SlideSlot {
    const reused = this._free.pop();
    if (reused) {
      // _recycleSlot already reset renderedSlide to -1 before pooling this slot.
      this._scrollHost.appendChild(reused.wrapper);
      return reused;
    }
    // `left` is set explicitly per mount by `_positionSlot` (JS centering with a
    // left-gutter floor), so no CSS auto-centering (`left:0;right:0;margin:0 auto`)
    // here — it would fight the explicit `left`.
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:absolute;';
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(canvas);
    wrapper.appendChild(canvas);
    let textLayer: HTMLDivElement | null = null;
    if (this._opts.enableTextSelection) {
      textLayer = document.createElement('div');
      textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      wrapper.appendChild(textLayer);
    }
    const highlightLayer = document.createElement('div');
    highlightLayer.style.cssText =
      'position:absolute;top:0;left:0;width:100%;height:100%;' +
      'overflow:hidden;pointer-events:none;';
    wrapper.appendChild(highlightLayer);
    const loadingLayer = document.createElement('span');
    loadingLayer.style.cssText = [
      'position:absolute',
      'top:0',
      'right:0',
      'bottom:0',
      'left:0',
      'display:none',
      'align-items:center',
      'justify-content:center',
      'background:rgba(255,255,255,0.72)',
      'pointer-events:none',
      'z-index:4',
    ].join(';');
    loadingLayer.setAttribute('role', 'status');
    loadingLayer.setAttribute('aria-live', 'polite');
    loadingLayer.setAttribute('aria-label', 'Loading slide');
    const progress = document.createElement('progress');
    progress.setAttribute('aria-hidden', 'true');
    loadingLayer.appendChild(progress);
    wrapper.appendChild(loadingLayer);
    let commentMarkerLayer: HTMLDivElement | null = null;
    let commentMargin: HTMLDivElement | null = null;
    let commentDecorationLayer: HTMLDivElement | null = null;
    if (this._commentsEnabled()) {
      commentMarkerLayer = document.createElement('div');
      commentMarkerLayer.style.cssText =
        'position:absolute;inset:0;overflow:hidden;pointer-events:none;';
      wrapper.appendChild(commentMarkerLayer);
      if (this._commentsOptions()?.cards !== false) {
        commentMargin = document.createElement('div');
        commentMargin.style.cssText =
          'position:absolute;top:0;height:100%;box-sizing:border-box;' +
          'overflow-x:hidden;overflow-y:auto;pointer-events:auto;';
        this._syncCommentMarginGeometry(commentMargin);
        if (this._commentsOptions()?.connectors !== undefined) {
          commentDecorationLayer = document.createElement('div');
          commentDecorationLayer.style.cssText =
            'position:absolute;top:0;left:0;overflow:visible;pointer-events:none;';
          wrapper.appendChild(commentDecorationLayer);
        }
        wrapper.appendChild(commentMargin);
      }
    }
    const elementLayer = createCanvasElementOutlineLayer(
      wrapper,
      this._opts.enableElementSelection === true,
    );
    this._scrollHost.appendChild(wrapper);
    const slot: SlideSlot = {
      wrapper,
      canvas,
      textLayer,
      highlightLayer,
      elementLayer,
      loadingLayer,
      commentMarkerLayer,
      commentMargin,
      commentDecorationLayer,
      commentElementBounds: Object.freeze([]),
      commentGeometry: null,
      commentAnchorSlide: -1,
      commentAnchorGeneration: 0,
      renderedSlide: -1,
      renderedScale: -1,
      dispatcher: new StaticCanvasRenderDispatcher(
        canvas,
        this._mode === 'worker' && !this._opts.enableMediaPlayback,
      ),
      presentationHandle: null,
      mediaInteractive: false,
      renderGeneration: 0,
      presentationGeneration: 0,
    };
    return slot;
  }

  private _recycleSlot(idx: number, slot: SlideSlot): void {
    this._slots.delete(idx);
    slot.renderGeneration++;
    // Invalidate pending presentSlide() calls before releasing the current
    // handle. A pending handle destroys itself when it resolves and observes the
    // changed generation.
    slot.presentationGeneration++;
    slot.presentationHandle?.destroy();
    slot.presentationHandle = null;
    slot.mediaInteractive = false;
    slot.dispatcher.destroy();
    if (!this._destroyed) {
      slot.dispatcher = new StaticCanvasRenderDispatcher(
        slot.canvas,
        this._mode === 'worker' && !this._opts.enableMediaPlayback,
      );
    }
    // Clear the per-slot text overlay so a slot sitting in the free pool holds no
    // stale spans. buildPptxTextLayer also clears on its next build, but an
    // unrendered pooled slot never gets that build, and the detached spans would
    // otherwise linger; drop them here.
    if (slot.textLayer) {
      slot.textLayer.innerHTML = '';
      // Drop any preview transform so a pooled slot re-used for another slide does
      // not inherit a stale scale() before its overlay is rebuilt.
      slot.textLayer.style.transform = '';
      slot.textLayer.style.transformOrigin = '';
    }
    slot.highlightLayer.innerHTML = '';
    slot.highlightLayer.style.transform = '';
    slot.highlightLayer.style.transformOrigin = '';
    slot.loadingLayer.style.display = 'none';
    if (slot.commentMarkerLayer) {
      slot.commentMarkerLayer.replaceChildren();
      slot.commentMarkerLayer.style.visibility = '';
    }
    if (slot.commentMargin) {
      this._commentUi?.disposeReadOnlyCommentMargin(slot.commentMargin);
      if (!this._commentUi) slot.commentMargin.replaceChildren();
      slot.commentMargin.style.visibility = '';
    }
    if (slot.commentDecorationLayer) {
      this._commentUi?.disposeReadOnlyCommentDecoration(slot.commentDecorationLayer);
      if (!this._commentUi) slot.commentDecorationLayer.replaceChildren();
      slot.commentDecorationLayer.style.visibility = '';
    }
    slot.commentElementBounds = Object.freeze([]);
    slot.commentGeometry = null;
    slot.commentAnchorSlide = -1;
    slot.commentAnchorGeneration++;
    renderCanvasElementOutline(slot.elementLayer, null);
    // `_previewSlot` pins an explicit CSS height while stretching the current
    // bitmap during a zoom burst. A slot can leave the visible range before the
    // debounced settle replaces that canvas, so do not carry the old-scale height
    // into its next slide. Main-mode renderSlide intentionally updates only the
    // CSS width and lets height follow the backing-store aspect ratio.
    slot.canvas.style.height = '';
    slot.renderedSlide = -1;
    slot.renderedScale = -1;
    slot.wrapper.remove();
    this._free.push(slot);
  }

  private _positionSlot(slot: SlideSlot, i: number, _r: VisibleWindow): void {
    slot.wrapper.dataset.slideIndex = String(i);
    slot.wrapper.style.top = `${this._slideOffset(i)}px`;
    const wpx = this._slideWidthPx();
    slot.wrapper.style.width = `${wpx}px`;
    slot.wrapper.style.height = `${this._slideHeightPx()}px`;
    this._syncCommentMarginGeometry(slot.commentMargin);
    if (slot.commentDecorationLayer) {
      const marginExtent = this._commentMarginExtent();
      slot.commentDecorationLayer.style.left = this._commentSide() === 'left'
        ? `${-marginExtent}px`
        : '0px';
      slot.commentDecorationLayer.style.width = `${wpx + marginExtent}px`;
      slot.commentDecorationLayer.style.height = `${this._slideHeightPx()}px`;
    }
    this._redrawElementOutlineForSlot(i, slot);
    // Horizontal placement (replaces the old CSS `left:0;right:0;margin:0 auto`
    // auto-centering, which cannot honour a left gutter). Centre the slide in the
    // scroll viewport, but never let its left edge cross the left gutter: when the
    // slide is narrower than the viewport it is centred (`(cw − sw)/2 > padL`); once
    // zoomed wider than the viewport the centre would go negative, so the floor
    // pins it at `padL` and the overflow scrolls right. Formula deliberately
    // duplicated per viewer (one line; not hoisted to core).
    const { left: padL } = this._padH();
    const authoredLeft = Math.max(padL, (this._scrollHost.clientWidth - wpx) / 2);
    slot.wrapper.style.left = this._commentSide() === 'left' && this._commentsEnabled()
      ? `calc(${authoredLeft}px + var(--ooxml-review-origin-x, 0px))`
      : `${authoredLeft}px`;
  }

  /** Device-pixel ratio for a render (opts override → window → 1). */
  private _dpr(): number {
    return this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
  }

  /**
   * Render slide `i` into `slot`. Routes strictly on the constructor-resolved
   * `_mode` (design §11 — no probing, no silent mis-pathing): `main` ⇒ paint the
   * slot's canvas directly via `renderSlide`; `worker` ⇒ transfer an ImageBitmap
   * from `renderSlideToBitmap`.
   *
   * Slot-identity guard: a slot recycled to a DIFFERENT slide while a previous
   * render is in flight must not repaint the stale slide. `slot.renderedSlide`
   * tracks the slide this slot is committed to; we stamp it up-front and bail on
   * resolution if it changed (the engine's own token guard is per-canvas; this is
   * the viewer's per-slot slide-identity check).
   *
   * Render epoch (main path): pixel staleness after a mid-flight `setScale` is
   * already handled by the engine's per-canvas token (the newer renderSlide on the
   * same canvas wins) — `setScale` recycles + re-mounts, and the re-mount always
   * re-dispatches `renderSlide` (renderedSlide reset to -1), so a fresh render is
   * always issued. But the viewer-side side effects of a STALE resolution — the
   * text-layer build (its run geometry is at the OLD scale) and the renderedSlide
   * bookkeeping — must NOT run, or a superseded render would rebuild the overlay
   * with stale x/y/w/h (the pool reuses slot objects, so the identity check alone
   * can pass for an old-epoch resolution). We gate them on the captured epoch.
   */
  private _renderSlot(
    i: number,
    slot: SlideSlot,
    mediaInteractive = false,
    reportErrors = true,
  ): Promise<void> | null {
    if (!this._pres) return null;
    // Slot-identity guard: this slot is already rendering / has rendered slide i.
    if (slot.renderedSlide === i) return null;
    if (i >= this.availableSlideCount && !this.layoutComplete) {
      // A virtual slot is only a stable placeholder until its metadata is
      // published. Do not start a Presentation wait/render here: the slot may be
      // recycled long before that slide becomes available, amplifying work and
      // retaining resources for an off-screen page. Layout publication below
      // dispatches only the placeholders that are still mounted.
      slot.renderedSlide = -1;
      slot.mediaInteractive = false;
      slot.loadingLayer.style.display = 'flex';
      return null;
    }
    slot.renderedSlide = i;
    const renderGeneration = ++slot.renderGeneration;
    slot.loadingLayer.style.display =
      i >= this.availableSlideCount && !this.layoutComplete ? 'flex' : 'none';

    const dpr = this._dpr();
    const widthPx = this._slideWidthPx();
    const epoch = this._renderEpoch;
    const scale = this._scale;
    const dispatcher = slot.dispatcher;
    const generation = dispatcher.begin();

    if (this._opts.enableMediaPlayback && mediaInteractive) {
      slot.mediaInteractive = true;
      return this._trackSlotLoading(
        i,
        slot,
        renderGeneration,
        this._renderInteractiveSlot(i, slot, widthPx, dpr, scale, epoch, reportErrors),
      );
    }
    slot.mediaInteractive = false;

    if (this._mode === 'worker') {
      return this._trackSlotLoading(
        i,
        slot,
        renderGeneration,
        this._renderSlotBitmap(
          i,
          slot,
          widthPx,
          dpr,
          scale,
          renderGeneration,
          dispatcher,
          generation,
          reportErrors,
        ),
      );
    }

    // Main mode: render straight onto the slot's canvas.
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive;
    const onTextRun = wantRuns ? (r: PptxTextRunInfo) => runs.push(r) : undefined;
    const canvas = slot.canvas;
    return this._trackSlotLoading(i, slot, renderGeneration, renderPptxFocusedSlide(this._pres, canvas, i, 'main', {
      width: widthPx, // this slide's own px width → uniform px-per-EMU scale (§7)
      dpr,
      onTextRun,
    })
      .then(() => {
        // Stale if the epoch moved (a setScale rescaled mid-flight — the run
        // geometry is at the old scale), or a recycle re-purposed this slot for a
        // different slide / freed it. Either way: skip the (stale) overlay build.
        // The engine's per-canvas token already discards the superseded pixels.
        if (
          renderGeneration !== slot.renderGeneration ||
          !dispatcher.isCurrent(generation) ||
          canvas !== slot.canvas ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedSlide !== i
        ) return;
        // This fresh render defines the scale the on-screen bitmap now lives at,
        // so a subsequent zoom preview stretches from HERE.
        slot.renderedScale = scale;
        if (wantOverlay && slot.textLayer) {
          // buildPptxTextLayer takes NUMBERS (not strings) for width/height. The
          // overlay must match the slot's CSS box, NOT the canvas backing store:
          // renderSlide sets `canvas.width = cssWidth × dpr`, so on a retina (dpr 2)
          // display the backing store is 2× the CSS box. Passing it would size the
          // overlay 2× too large (overflowing the wrapper + inflating the scroll
          // area). Pass the CSS px directly — the uniform slide width/height at the
          // current scale (rounded).
          buildPptxTextLayer(slot.textLayer, runs, Math.round(widthPx), Math.round(this._slideHeightPx()), this._hyperlinkHandler(), i);
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitSlotComments(i, slot);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        const isCurrent =
          renderGeneration === slot.renderGeneration &&
          dispatcher.isCurrent(generation) &&
          canvas === slot.canvas &&
          epoch === this._renderEpoch &&
          this._slots.get(i) === slot &&
          slot.renderedSlide === i;
        if (!isCurrent) return;
        if (reportErrors) this._reportRenderError(err);
        else throw err;
      }));
  }

  private async _trackSlotLoading(
    slideIndex: number,
    slot: SlideSlot,
    renderGeneration: number,
    render: Promise<void>,
  ): Promise<void> {
    try {
      await render;
    } finally {
      if (
        renderGeneration === slot.renderGeneration &&
        this._slots.get(slideIndex) === slot &&
        slot.renderedSlide === slideIndex
      ) {
        slot.loadingLayer.style.display = 'none';
      }
    }
  }

  private _bindLayoutPresentation(presentation: PptxPresentation): void {
    this._unbindLayoutPresentation();
    let initial = true;
    this._layoutUnsubscribe = subscribePptxLayout(
      presentation,
      () => ({
        availableSlides: presentation.availableSlideCount,
        slideCount: presentation.slideCount,
        exact: presentation.layoutComplete,
        complete: presentation.layoutComplete,
      }),
      (publication) => {
        if (initial) {
          initial = false;
          return;
        }
        this._onLayoutPublication(presentation, publication);
      },
      (error) => this._reportRenderError(error),
    );
  }

  private _unbindLayoutPresentation(): void {
    this._layoutUnsubscribe?.();
    this._layoutUnsubscribe = null;
    this._layoutFailed = false;
    this._wakeLayoutWaiters();
  }

  private _onLayoutPublication(
    presentation: PptxPresentation,
    publication: PptxLayoutPublication,
  ): void {
    if (this._destroyed || presentation !== this._pres) return;
    this._wakeLayoutWaiters();
    if (publication.error !== undefined) {
      this._layoutFailed = true;
      this._errorRouter.reportBackground(
        publication.error,
        this._opts.onLayoutComplete !== undefined,
      );
      return;
    }
    this._scanAvailableComments(presentation, true);
    const mediaRange = this._opts.enableMediaPlayback ? this._mediaRange() : null;
    for (const [slideIndex, slot] of this._slots) {
      if (slideIndex >= presentation.availableSlideCount || slot.renderedSlide === slideIndex) continue;
      void this._renderSlot(
        slideIndex,
        slot,
        !!mediaRange && this._rangeContains(mediaRange, slideIndex),
      );
    }
    if (this._lastRange) this._emitVisibleSlideChange(this._lastRange);
  }

  private _wakeLayoutWaiters(): void {
    for (const resolve of this._layoutWaiters) resolve();
    this._layoutWaiters.clear();
  }

  /** Supersede pending comment navigation and release availability waits now. */
  private _beginCommentNavigation(): number {
    const generation = ++this._commentNavigationGeneration;
    this._wakeLayoutWaiters();
    return generation;
  }

  private async _waitForSlideMetadata(
    presentation: PptxPresentation,
    slideIndex: number,
    generation: number,
  ): Promise<boolean> {
    return await this._errorRouter.ownBackgroundLifecycle(async () => {
      while (
        !this._destroyed &&
        generation === this._commentNavigationGeneration &&
        presentation === this._pres &&
        slideIndex >= presentation.availableSlideCount &&
        !presentation.layoutComplete &&
        !this._layoutFailed
      ) {
        await new Promise<void>((resolve) => this._layoutWaiters.add(resolve));
      }
      if (this._destroyed || presentation !== this._pres) return false;
      if (presentation.layoutComplete || this._layoutFailed) {
        await presentation.waitUntilLayoutComplete?.();
      }
      if (generation !== this._commentNavigationGeneration) return false;
      return slideIndex < presentation.availableSlideCount;
    });
  }

  private _emitVisibleSlideChange(range: VisibleWindow): void {
    if (!this._pres) return;
    const total = this._pres.slideCount;
    const complete = this.layoutComplete;
    if (
      range.topIndex === this._lastTopIndex &&
      total === this._lastReportedTotal &&
      complete === this._lastReportedLayoutComplete
    ) return;
    this._lastTopIndex = range.topIndex;
    this._lastReportedTotal = total;
    this._lastReportedLayoutComplete = complete;
    this._opts.onVisibleSlideChange?.(range.topIndex, total, complete);
  }

  /**
   * Render one mounted slot through the stateful media presentation API.
   *
   * This path intentionally bypasses bitmaprenderer even for worker-backed
   * presentations: presentSlide() renders the base off-thread and composites
   * interactive video on a main-thread 2D canvas. The slot generation closes the
   * same-index recycle/reload hole that a viewer-wide render epoch cannot detect.
   */
  private _renderInteractiveSlot(
    i: number,
    slot: SlideSlot,
    widthPx: number,
    dpr: number,
    scale: number,
    epoch: number,
    reportErrors = true,
  ): Promise<void> {
    if (!this._pres) return Promise.resolve();
    const generation = ++slot.presentationGeneration;
    slot.presentationHandle?.destroy();
    slot.presentationHandle = null;
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive;
    const onTextRun = wantRuns ? (r: PptxTextRunInfo) => runs.push(r) : undefined;

    return this._pres
      .presentSlide(slot.canvas, i, {
        width: widthPx,
        dpr,
        onTextRun,
        onError: (error) => {
          if (generation === slot.presentationGeneration) this._reportRenderError(error);
        },
      })
      .then((handle) => {
        if (
          generation !== slot.presentationGeneration ||
          !slot.mediaInteractive ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedSlide !== i
        ) {
          handle.destroy();
          return;
        }
        slot.presentationHandle = handle;
        slot.renderedScale = scale;
        if (wantOverlay && slot.textLayer) {
          buildPptxTextLayer(
            slot.textLayer,
            runs,
            Math.round(widthPx),
            Math.round(this._slideHeightPx()),
            this._hyperlinkHandler(),
            i,
          );
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitSlotComments(i, slot);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        if (generation !== slot.presentationGeneration) return;
        if (reportErrors) this._reportRenderError(err);
        else throw err;
      });
  }

  /**
   * Reconcile stateful media handles against the small media lifecycle range,
   * independently from the (potentially whole-deck) mounted text range.
   */
  private _syncMediaPlayback(mediaRange = this._mediaRange()): void {
    if (!this._opts.enableMediaPlayback) return;
    for (const [i, slot] of this._slots) {
      const shouldBeInteractive = this._rangeContains(mediaRange, i);
      if (shouldBeInteractive === slot.mediaInteractive) continue;
      if (shouldBeInteractive) {
        slot.mediaInteractive = true;
        // Render onto a spare canvas. A statically-rendered worker slot already
        // owns a bitmaprenderer context, which cannot be changed to the 2D
        // context presentSlide() needs.
        this._settleInteractiveSlot(
          i,
          slot,
          this._slideWidthPx(),
          this._dpr(),
          this._scale,
          this._renderEpoch,
        );
      } else {
        // Stop playback/RAF and invalidate a pending presentSlide immediately.
        // Keep the last painted canvas and text overlay as the static offscreen
        // representation; if it becomes active again, the spare-canvas upgrade
        // redraws it at the current scale.
        slot.mediaInteractive = false;
        slot.presentationGeneration++;
        slot.presentationHandle?.destroy();
        slot.presentationHandle = null;
      }
    }
  }

  /** Route an async render failure to `onError`, or `console.error` when none is
   *  set (so failures are never fully silent), and never after teardown. */
  private _reportRenderError(err: unknown): void {
    this._errorRouter.report(err);
  }

  /**
   * Worker-mode slot render: dispatch `renderSlideToBitmap`, transfer the result
   * via a per-slot `bitmaprenderer` context, and manage the ImageBitmap lifecycle.
   *
   * Coalescing / drop-stale (design §11):
   *  - Skip if slide `i` is already in flight (a scroll storm won't double-dispatch).
   *  - Skip if slide `i` already left the mounted window before dispatch.
   *  - On resolution, if `slot` is no longer THIS slide's live slot (it recycled to
   *    another slide, or slide `i` re-mounted onto a DIFFERENT slot while this render
   *    was in flight), close the orphan bitmap and skip the paint. In that
   *    re-mount case a live slot for `i` still awaits a render, so once we clear
   *    the in-flight guard we re-dispatch it — a slide that recycled and re-mounted
   *    mid-flight must never stay blank.
   *  - RENDER EPOCH: the dispatch captures `this._renderEpoch`. `setScale` bumps
   *    the epoch, so a resolution whose captured epoch ≠ the live epoch is STALE
   *    even when the SAME slot object is still mounted for slide `i` (the pool
   *    reuses slot objects, so the identity check alone can't catch a zoom that
   *    happened mid-flight). A moved epoch ⇒ close the orphan + re-dispatch the
   *    live slot at the new scale, never paint the old-scale bitmap.
   *
   * Do NOT pass `dim` or `skipMediaControls` to `renderSlideToBitmap`. The scroll
   * viewer never dims slides (design §8.2 / Delta 6); passing neither means the
   * static play-badge renders on media slides (matching `PptxViewer`'s
   * non-media-playback path) — acceptable for v1.
   */
  private async _renderSlotBitmap(
    i: number,
    slot: SlideSlot,
    widthPx: number,
    dpr: number,
    scale: number,
    renderGeneration = ++slot.renderGeneration,
    dispatcher = slot.dispatcher,
    generation = dispatcher.begin(),
    reportErrors = true,
  ): Promise<void> {
    if (this._slideInFlight.has(i)) return; // coalesce: already dispatched
    // Drop-stale before dispatch: if this slide already scrolled out of the
    // mounted window, don't dispatch at all.
    if (this._slots.get(i) !== slot) return;
    const epoch = this._renderEpoch;
    this._slideInFlight.add(i);
    // Capture the actual canvas/context pair at dispatch. A media promotion swaps
    // slot.canvas to a 2D canvas while this worker bitmap is in flight; reading
    // the mutable slot fields after await would clear that new interactive canvas.
    const canvas = slot.canvas;
    // Whether this invocation actually painted its slot. When it did NOT (stale
    // epoch or moved identity), the `finally` may need to re-dispatch a live slot.
    let painted = false;
    // IX6 — harvest the slide's run geometry alongside the bitmap so the
    // worker-mode selection overlay is built from the SAME data main mode uses.
    // The runs ride back beside the bitmap (one round-trip), collected only when
    // an overlay is actually wanted.
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive;
    const runs: PptxTextRunInfo[] = [];
    try {
      const bmp = await renderPptxFocusedSlide(this._pres!, canvas, i, 'worker', {
        width: widthPx,
        dpr,
        onTextRun: wantRuns ? (r) => runs.push(r) : undefined,
      });
      // Stale if EITHER (a) the epoch moved (a setScale rescaled mid-flight, so
      // this bitmap is at a superseded resolution — this catches the case where
      // the SAME slot object is re-mounted for slide `i`, which the identity check
      // below cannot), or (b) the slot recycled to a different slide / slide `i`
      // re-mounted onto a DIFFERENT slot. Either way: close + skip the paint.
      if (
        renderGeneration !== slot.renderGeneration ||
        !dispatcher.isCurrent(generation) ||
        canvas !== slot.canvas ||
        epoch !== this._renderEpoch ||
        this._slots.get(i) !== slot ||
        slot.renderedSlide !== i
      ) {
        bmp.close();
        return;
      }
      const size = {
        cssWidth: Math.round(bmp.width / dpr),
        cssHeight: Math.round(bmp.height / dpr),
      };
      // Interactive media later paints through `presentSlide()` on the same
      // pooled canvas, so that canvas must remain a 2D canvas. A browser canvas
      // cannot switch back from `bitmaprenderer` after the first acquisition.
      const committed = this._opts.enableMediaPlayback
        ? dispatcher.commitBitmapTo2d(generation, bmp, size)
        : dispatcher.commitBitmap(generation, bmp, size);
      if (!committed) return;
      // This bitmap now defines the scale the on-screen canvas lives at, so a
      // later zoom preview stretches from HERE (design §7 renderedScale).
      slot.renderedScale = scale;
      // IX6 — build the selection overlay from the runs the worker just shipped.
      // Reached only past the staleness gate, so the geometry matches THIS paint.
      // Clear any preview transform first (a settle lands at the current scale, so
      // the `scale()` from `_previewSlot` is stale) — mirrors the main path. The
      // overlay is sized to the slot's CSS box (Math.round of the uniform slide
      // width/height at the current scale), NOT the dpr-scaled backing store.
      if (slot.textLayer) {
        slot.textLayer.style.transform = '';
        slot.textLayer.style.transformOrigin = '';
        if (wantOverlay) {
          buildPptxTextLayer(
            slot.textLayer,
            runs,
            Math.round(widthPx),
            Math.round(this._slideHeightPx()),
            this._hyperlinkHandler(),
            i,
          );
        }
      }
      if (wantRuns) this._refreshFindRuns(i, runs);
      this._commitSlotComments(i, slot);
      this._redrawSlotHighlights(i, slot);
      painted = true;
    } catch (err) {
      const isCurrent =
        renderGeneration === slot.renderGeneration &&
        dispatcher.isCurrent(generation) &&
        canvas === slot.canvas &&
        epoch === this._renderEpoch &&
        this._slots.get(i) === slot &&
        slot.renderedSlide === i;
      if (isCurrent) {
        if (reportErrors) this._reportRenderError(err);
        else throw err;
      }
    } finally {
      this._slideInFlight.delete(i);
      // Re-dispatch ONLY when this invocation went stale — a LIVE slot for slide
      // `i` still awaits a correct render and the reason we didn't paint was
      // staleness, not a render failure. The two staleness cases:
      //  - IDENTITY MOVED (`live !== slot`): slide `i` re-mounted onto a DIFFERENT
      //    slot while we ran (the re-mount's own dispatch was coalesced away by
      //    the in-flight guard), so the live slot has no render in flight.
      //  - EPOCH MOVED (`epoch !== this._renderEpoch`): a `setScale` bumped the
      //    epoch mid-flight, so this bitmap was at a superseded scale. The live
      //    slot may be the SAME object reused from the pool, which the identity
      //    test alone would miss — the epoch test catches the same-slot case.
      // NO RETRY ON PLAIN REJECTION: when the slot is still live at the same epoch
      // and we simply failed (`renderSlideToBitmap` rejected or the transfer threw),
      // `!painted` holds but BOTH staleness tests are false, so we do NOT
      // re-dispatch. Retrying a plain failure would loop unbounded (reject →
      // re-dispatch → reject → …); the onError contract is that "a failed slide is
      // left blank" (see PptxScrollViewerOptions.onError), so we leave it blank.
      // Bounded epoch-then-reject: an epoch-moved re-dispatch captures the NEW
      // epoch, so if that fresh render then rejects at the still-current epoch,
      // both tests are false and it stops — no unbounded retry.
      const live = this._slots.get(i);
      if (
        !painted &&
        live &&
        (
          live !== slot ||
          epoch !== this._renderEpoch ||
          renderGeneration !== live.renderGeneration ||
          !dispatcher.isCurrent(generation)
        ) &&
        !this._slideInFlight.has(i) &&
        !this._destroyed &&
        !(this._opts.enableMediaPlayback && live.mediaInteractive)
      ) {
        // live.renderedSlide === i already (set by _renderSlot on mount); the fresh
        // dispatch runs at the CURRENT epoch/scale via _slideWidthPx().
        void this._renderSlotBitmap(i, live, this._slideWidthPx(), this._dpr(), this._scale);
      }
    }
  }

  /**
   * Set the absolute (dimensionless) zoom scale — a multiplier over the 96-dpi
   * natural slide size, matching `DocxScrollViewer` — clamped inline to
   * `[zoomMin ?? 0.1, zoomMax ?? 4]` (absolute bounds, XlsxViewer convention — NOT
   * multiples of the base fit; design §3 keeps the clamp in the viewer, not core),
   * then re-anchor VERTICALLY so the slide currently under the viewport top stays
   * fixed. A no-op when the clamped scale is unchanged. Called BEFORE the deck is
   * loaded / the base fit is established, the clamped factor is LATCHED (IX9 F1,
   * family-unified with the single-canvas viewers) and applied by `relayout()`
   * once the layout establishes — `onScaleChange` fires then.
   *
   * FLICKER-FREE (design §7): this does NOT re-render the visible slides inline.
   * It shows an immediate CSS preview (stretch the existing bitmaps, scale the
   * overlays) and DEBOUNCES a full-resolution settle re-render for ZOOM_SETTLE_MS,
   * so a wheel/pinch burst never blanks a slide and coalesces into one crisp render.
   *
   * Re-anchor (written from scratch — XlsxViewer only re-anchors horizontally):
   * capture `top = topIndex` and the intra-slide fraction `intraFrac` from the
   * CURRENT range BEFORE rescale; after recomputing heights at the new scale,
   * `newScrollTop = offsets'[top] + intraFrac × heights'[top]`, clamped to
   * `[0, totalHeight' − viewportHeight]`. Because a slide's height scales linearly
   * with `_scale`, the same fractional position maps exactly to the new geometry.
   *
   * CAVEAT — base fit below the floor: `relayout()` sets `_scale = base` WITHOUT
   * clamping to `[zoomMin, zoomMax]`. If the base fit is below `zoomMin` (a wide
   * slide in a narrow container), the initial scale sits under the floor, but once
   * the user zooms via `setScale` the clamp pins the minimum to `zoomMin`, so they
   * can no longer return below the floor to the original base fit through this API.
   */
  setScale(scale: number): void {
    const zoomMin = this._opts.zoomMin ?? 0.1;
    const zoomMax = this._opts.zoomMax ?? 4;
    const next = Math.min(zoomMax, Math.max(zoomMin, scale));
    // Consume the gesture-only pointer anchor (Ctrl/⌘+wheel set it just above)
    // FIRST — before every early return — so a gesture whose setScale ends up a
    // NO-OP (already pinned at zoomMin/zoomMax) or latches pre-establishment can
    // never leak a stale anchor into a later non-gesture setScale (slider,
    // steppers, fitWidth/fitPage, resize re-fit, public API), which must keep
    // the historical viewport-TOP anchoring. `null` for every non-gesture source.
    const gestureAnchor = this._pendingZoomAnchor;
    this._pendingZoomAnchor = null;
    if (!this._pres || this._pres.slideCount === 0 || !this._scaleEstablished) {
      // IX9 F1 (family-unified pre-load semantics): a setScale before the deck is
      // loaded / before the base fit is established is LATCHED, not dropped —
      // matching the single-canvas viewers, which honour a pre-load setScale on
      // their first render. relayout() applies it right after establishing the
      // base and fires onScaleChange there (at application time).
      this._pendingScale = next;
      return;
    }
    if (next === this._scale) return;
    const prevScale = this._scale;
    const anchorY = gestureAnchor ? gestureAnchor.y : 0;

    // Capture the VERTICAL anchor from the CURRENT layout, before rescale, as a
    // (slide index, intra-slide fraction) pair. Anchoring on a slide — not on the
    // raw scrollTop — is what keeps the re-anchor exact despite the scale-INVARIANT
    // desk padding and inter-slide gaps (only the slide heights scale, so a whole-
    // scroll linear rescale would drift by the padding). The point we pin is the
    // content under the pointer: content-y = scrollTop + anchorY (anchorY 0 ⇒ the
    // viewport top, the historical behaviour).
    const scrollTop0 = this._scrollHost.scrollTop;
    const anchorContentY = scrollTop0 + anchorY;
    // Which slide does that content-y fall in? `computeVisibleRange` attributes a
    // point in the trailing gap to the slide ABOVE it, so clamp the fraction to
    // [0,1] to pin the slide rather than drift into the gap.
    const top = this._slideIndexAtOffset(anchorContentY);
    const h0 = this._uniformSlideHeight;
    let intraFrac = h0 > 0 ? (anchorContentY - this._slideOffset(top)) / h0 : 0;
    intraFrac = Math.min(1, Math.max(0, intraFrac));

    // HORIZONTAL anchor (gesture only — a non-gesture setScale leaves scrollLeft
    // untouched, matching the historical behaviour). The slide's left edge sits at
    // the scale-INVARIANT left gutter `padL` when it overflows the viewport (see
    // `_positionSlot`): screen-x of content pixel c is `padL + c − scrollLeft`,
    // so the pointer's offset INTO the scaling region is `x − padL` and the
    // scroll offset itself already lives in the region's own px.
    const padL = this._padH().left;
    const scrollLeft0 = this._scrollHost.scrollLeft || 0;

    // Bump the render epoch BEFORE recycling/re-dispatching so any in-flight
    // render dispatched at the old scale is recognised as stale on resolution.
    this._renderEpoch++;

    // Rescale, recompute heights, resize the spacer to the new total height.
    this._scale = next;
    this._recomputeHeights();
    const r1 = this._rangeAt(0, this._overscan());
    this._spacer.style.height = `${r1.totalHeight}px`;
    // The slide px width changed with the scale, so the horizontal extent moves too.
    this._syncSpacerWidth();

    // Pin the same fractional position of the same slide under the pointer (or the
    // viewport top for a non-gesture zoom): the on-screen y of that content point
    // must stay at `anchorY`, so newScrollTop = newContentY − anchorY.
    const maxTop = Math.max(0, r1.totalHeight - this._scrollHost.clientHeight);
    const newContentY = this._slideOffset(top) + intraFrac * this._uniformSlideHeight;
    // The leading desk padding is fixed viewport space: none of it scales. If the
    // anchor is still inside that padding, keep the native scroll offset instead
    // of snapping to slide 0's offset and hiding the margin after a programmatic
    // zoom at scrollTop 0.
    const reanchoredTop = anchorContentY < this._slideOffset(0)
      ? scrollTop0
      : newContentY - anchorY;
    this._scrollHost.scrollTop = Math.min(maxTop, Math.max(0, reanchoredTop));

    // Re-anchor horizontally for a gesture zoom. `padL` is a FIXED (non-scaling)
    // gutter, so it is subtracted from the ANCHOR only — the scroll offset stays
    // in NATIVE space with the browser's own [0, maxLeft] clamp. (Shifting the
    // scroll by ±padL as well would run the fixed gutter through the zoom ratio
    // and over-compensate by padL·(ratio−1) per step: with `screen = padL + c −
    // scrollLeft`, pinning c·ratio under the pointer x gives exactly
    // `scrollLeft' = ratio·(scrollLeft + (x−padL)) − (x−padL)`.) Skipped entirely
    // for a non-gesture setScale so slider/stepper/API/resize is unchanged.
    if (gestureAnchor) {
      const maxLeft = Math.max(0, (this._spacer.offsetWidth || 0) - this._scrollHost.clientWidth);
      this._scrollHost.scrollLeft = anchoredZoomOffset(
        scrollLeft0,
        gestureAnchor.x - padL,
        prevScale,
        next,
        { maxScroll: maxLeft },
      );
    }

    // FLICKER-FREE ZOOM (design §7). Do NOT recycle + re-render in-window slots
    // (that blanks each visible slide to white every tick). Instead:
    //  1. CSS-PREVIEW the currently-mounted slots at the new geometry — reposition
    //     the wrapper, stretch the existing canvas bitmap via style.width/height
    //     (soft but never blank), and scale the text overlay by the ratio between
    //     the new scale and the scale the overlay was built at.
    //  2. DEBOUNCE a full-resolution settle re-render: schedule it ZOOM_SETTLE_MS
    //     after the LAST setScale so a wheel/pinch burst coalesces into one render.
    this._previewVisible();
    this._scheduleSettle();
    // IX9 change notification. Only reached when `next` differs from the prior
    // scale (early-returned above), so every source — the public setScale, the
    // ladder steppers, fitWidth/fitPage, Ctrl-wheel, and the _onResize re-fit
    // (which routes through here) — notifies through this one hook.
    this._opts.onScaleChange?.(next);
  }

  // ─── IX9 zoom contract (ZoomableViewer) ───────────────────────────────────

  /** IX9 {@link ZoomableViewer} — the current zoom factor, where `1` = 100% (a
   *  slide at its natural EMU→px size). This is the viewer's absolute `_scale`
   *  (`slideWidth/EMU_PER_PX × _scale` is the drawn width), so it reads `1` at
   *  true 100% and, after the initial fit-to-width, the base fit factor. Before
   *  the fit is established it reports a latched pre-load `setScale` (IX9 F1) if
   *  one is pending — matching what a single-canvas viewer would show — else `1`. */
  getScale(): number {
    if (this._scaleEstablished) return this._scale;
    return this._pendingScale ?? 1;
  }

  /** IX9 {@link ZoomableViewer} — step up to the next rung of the shared zoom
   *  ladder above the current factor (clamped to `zoomMax` by {@link setScale}). */
  zoomIn(): void {
    this.setScale(nextZoomStep(this.getScale()));
  }

  /** IX9 {@link ZoomableViewer} — step down to the next lower ladder rung. */
  zoomOut(): void {
    this.setScale(prevZoomStep(this.getScale()));
  }

  /**
   * IX9 {@link ZoomableViewer} — fit a slide's WIDTH to the container (the classic
   * continuous-scroll "fit width"). Sets the scale to the width-fit base for the
   * current container, then re-anchors + re-renders via {@link setScale}. Defers
   * (no-op) while the container is unlaid-out. The `zoomMin`/`zoomMax` clamp still
   * applies, so a fit below `zoomMin` pins to `zoomMin`.
   */
  fitWidth(): void {
    this._fit('width');
  }

  /**
   * IX9 {@link ZoomableViewer} — fit a WHOLE slide (width and height) inside the
   * container so one slide is visible without scrolling; takes the tighter of the
   * width/height fit. Uses the deck-wide (uniform) slide size. Defers while
   * unlaid-out.
   */
  fitPage(): void {
    this._fit('page');
  }

  /** Shared fit for {@link fitWidth}/{@link fitPage}: the width-fit factor is the
   *  established base (`_baseScale`); the page-fit additionally bounds by the
   *  container height against the (uniform) slide height. Applies via
   *  {@link setScale} so the flicker-free re-anchor / settle path and
   *  `onScaleChange` all run. */
  private _fit(mode: 'width' | 'page'): void {
    if (!this._pres || this._pres.slideCount === 0) return;
    const scale = fitScale(
      {
        contentWidth: this._pres.slideWidth / EMU_PER_PX,
        contentHeight: this._pres.slideHeight / EMU_PER_PX,
        containerWidth: this._fitWidthPx(),
        containerHeight: this._scrollHost.clientHeight,
      },
      mode,
    );
    if (scale <= 0) return; // unlaid-out — defer
    this.setScale(scale);
  }

  /**
   * CSS preview of the visible window at the current `_scale` (design §7
   * mechanism 1), WITHOUT re-rendering. Slots leaving the window recycle normally;
   * slots ENTERING the window mount fresh (rendered at the current scale directly,
   * so they never need a preview); slots that STAY are repositioned and their
   * canvas + text overlay are CSS-transformed to the new size (the device buffer
   * is untouched — that is the whole point: no synchronous clear, no blank frame).
   */
  private _previewVisible(): void {
    if (!this._pres || this._pres.slideCount === 0) return;
    const r = this._range();
    const mediaRange = this._opts.enableMediaPlayback ? this._mediaRange() : null;
    this._lastRange = r;

    // Recycle slots that left [start, end].
    for (const [idx, slot] of [...this._slots]) {
      if (idx < r.start || idx > r.end) this._recycleSlot(idx, slot);
    }
    // For every index in the window: mount fresh if missing (renders at the current
    // scale), or CSS-preview if already mounted (no re-render, no device resize).
    for (let i = r.start; i <= r.end; i++) {
      const existing = this._slots.get(i);
      if (!existing) {
        const slot = this._acquireSlot();
        this._positionSlot(slot, i, r);
        this._slots.set(i, slot);
        this._redrawSlotComments(i, slot);
        this._renderSlot(i, slot, !!mediaRange && this._rangeContains(mediaRange, i));
      } else {
        this._previewSlot(existing, i, r);
      }
    }
    if (mediaRange) this._syncMediaPlayback(mediaRange);
    // Fire onVisibleSlideChange only when the top slide actually changed.
    this._emitVisibleSlideChange(r);
  }

  /**
   * CSS-preview a single already-mounted slot at the new geometry (design §7): the
   * wrapper is repositioned + sized (via `_positionSlot`), the canvas bitmap is
   * STRETCHED to the new CSS size (no `canvas.width` — the device buffer, and thus
   * the drawn pixels, are left intact, just scaled by the browser), and the text
   * overlay is scaled by `newScale / renderedScale` so it tracks the stretched
   * slide. `renderedScale <= 0` means the slot's first render hasn't resolved yet
   * (nothing to stretch); the pending render captured the current scale, so it
   * lands correct and no preview is needed.
   */
  private _previewSlot(slot: SlideSlot, i: number, r: VisibleWindow): void {
    this._positionSlot(slot, i, r);
    // Stretch the existing bitmap to the new CSS box (device buffer untouched).
    slot.canvas.style.width = `${this._slideWidthPx()}px`;
    slot.canvas.style.height = `${this._slideHeightPx()}px`;
    if (slot.textLayer && slot.renderedScale > 0) {
      const ratio = this._scale / slot.renderedScale;
      slot.textLayer.style.transformOrigin = '0 0';
      slot.textLayer.style.transform = `scale(${ratio})`;
    }
    if (slot.renderedScale > 0) {
      const ratio = this._scale / slot.renderedScale;
      const previewScale = Math.round(ratio * 1_000_000) / 1_000_000;
      if (slot.commentMargin) this._commentUi?.previewReadOnlyCommentMargin(slot.commentMargin, ratio);
      for (const marker of slot.commentMarkerLayer?.children ?? []) {
        if ((marker as HTMLElement).dataset.ooxmlCommentMarker === undefined) continue;
        (marker as HTMLElement).style.transform = `translate(-50%,-50%) scale(${previewScale})`;
      }
      if (slot.commentMarkerLayer) slot.commentMarkerLayer.style.visibility = '';
      if (slot.commentMargin) slot.commentMargin.style.visibility = '';
      if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = '';
      return;
    }
    // No committed geometry exists during the first render, so there is nothing
    // trustworthy to preview yet.
    if (slot.commentMarkerLayer) slot.commentMarkerLayer.style.visibility = 'hidden';
    if (slot.commentMargin) slot.commentMargin.style.visibility = 'hidden';
    if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = 'hidden';
  }

  /** (Re)schedule the debounced settle re-render (design §7 mechanism 2). Resets
   *  the timer on every call so a burst of `setScale` dispatches ONE settle
   *  ZOOM_SETTLE_MS after the LAST call. Cleared in `destroy()`. */
  private _scheduleSettle(): void {
    if (this._settleTimer !== null) clearTimeout(this._settleTimer);
    this._settleTimer = setTimeout(() => {
      this._settleTimer = null;
      this._settleRender();
    }, ZOOM_SETTLE_MS);
  }

  /** Full-resolution settle re-render of the visible window (design §7 mechanisms
   *  2+3). Re-renders each mounted slot at the current scale via the double-buffer
   *  swap (main) / same-canvas transfer (worker). Both modes rebuild the text
   *  overlay from the fresh render's run geometry (IX6 — worker mode collects the
   *  runs off-thread via `_renderSlotBitmap`) and clear the preview transform.
   *  Dispatched at the CURRENT epoch; the existing epoch gate discards it if a
   *  later `setScale` supersedes it mid-render. */
  private _settleRender(): void {
    if (this._destroyed || !this._pres || this._pres.slideCount === 0) return;
    const mediaRange = this._opts.enableMediaPlayback ? this._mediaRange() : null;
    for (const [i, slot] of [...this._slots]) {
      // Whole-deck text mounting must not turn a zoom settle into a whole-deck
      // media rebuild. Offscreen static canvases keep their CSS preview and are
      // redrawn through presentSlide only when they enter the bounded media range.
      if (mediaRange && !this._rangeContains(mediaRange, i)) continue;
      // Skip slots already at the current scale (a slot that entered the window
      // during the burst mounted fresh at the current scale — nothing to settle).
      if (slot.renderedScale === this._scale) continue;
      this._settleSlot(i, slot);
    }
  }

  /**
   * Settle-render one slot at the current scale (design §7 mechanism 3).
   *
   * WORKER: re-dispatch the bitmap render into the SAME canvas. The worker path
   * sizes the device buffer and `transferFromImageBitmap`s it in ONE synchronous
   * step (no await between `canvas.width = …` and the transfer), so the browser
   * never composites an intermediate blank frame — no spare canvas is needed. The
   * `renderedScale === _scale` gate in `_settleRender` plus the epoch gate inside
   * `_renderSlotBitmap` keep this correct and idempotent.
   *
   * MAIN: `renderSlide` synchronously sets `canvas.width = …` (which CLEARS the
   * backing store to blank) BEFORE its first await and paints AFTER — so rendering
   * into the on-screen canvas would flash it white. Render into a SPARE off-DOM
   * canvas instead; only once it resolves at the current epoch do we swap it into
   * the wrapper (replacing the old canvas). The old canvas keeps showing the
   * stretched preview until the instant of the swap — blank-free.
   */
  private _settleSlot(i: number, slot: SlideSlot): void {
    if (!this._pres) return;
    const dpr = this._dpr();
    const widthPx = this._slideWidthPx();
    const scale = this._scale;
    const epoch = this._renderEpoch;

    if (this._opts.enableMediaPlayback && slot.mediaInteractive) {
      this._settleInteractiveSlot(i, slot, widthPx, dpr, scale, epoch);
      return;
    }
    if (this._opts.enableMediaPlayback) return;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
      return;
    }

    // Main mode: double-buffer. Render into a spare canvas kept off-DOM. The
    // spare REPLACES the on-screen canvas on swap, so it must carry the slide
    // shadow too — otherwise a settle would silently drop it.
    const spare = document.createElement('canvas');
    const renderGeneration = ++slot.renderGeneration;
    spare.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(spare);
    const spareDispatcher = new StaticCanvasRenderDispatcher(spare, false);
    const generation = spareDispatcher.begin();
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive;
    const onTextRun = wantRuns ? (r: PptxTextRunInfo) => runs.push(r) : undefined;
    renderPptxFocusedSlide(this._pres, spare, i, 'main', {
      width: widthPx,
      dpr,
      onTextRun,
    })
      .then(() => {
        // Discard if superseded: a later setScale bumped the epoch (this spare is
        // at a stale scale), or the slot recycled / moved to another slide. Drop
        // the spare (it is off-DOM, so GC reclaims it) and do NOT swap.
        if (
          renderGeneration !== slot.renderGeneration ||
          !spareDispatcher.isCurrent(generation) ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedSlide !== i
        ) {
          spareDispatcher.destroy();
          return;
        }
        // Swap the freshly-painted spare in for the old (stretched-preview) canvas.
        // The old canvas was the only child that showed content; replacing it in
        // one DOM op means the screen goes from preview → crisp with no blank tick.
        const old = slot.canvas;
        slot.dispatcher.destroy();
        slot.wrapper.insertBefore(spare, old);
        old.remove();
        slot.canvas = spare;
        slot.dispatcher = spareDispatcher;
        slot.renderedScale = scale;
        // Rebuild the overlay at the full resolution and CLEAR the preview
        // transform (the crisp render no longer needs the scale()).
        if (slot.textLayer) {
          slot.textLayer.style.transform = '';
          slot.textLayer.style.transformOrigin = '';
          if (wantOverlay) {
            // buildPptxTextLayer takes NUMBERS: pass the CSS box (uniform slide
            // width/height at the current scale), NOT the retina backing store.
            buildPptxTextLayer(slot.textLayer, runs, Math.round(widthPx), Math.round(this._slideHeightPx()), this._hyperlinkHandler(), i);
          }
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitSlotComments(i, slot);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        if (
          renderGeneration === slot.renderGeneration &&
          spareDispatcher.isCurrent(generation) &&
          epoch === this._renderEpoch &&
          this._slots.get(i) === slot &&
          slot.renderedSlide === i
        ) this._reportRenderError(err);
        spareDispatcher.destroy();
      });
  }

  /**
   * Settle an interactive slide onto a spare 2D canvas, then atomically swap it
   * in and destroy the handle tied to the retired canvas. Pending handles from a
   * superseded zoom/recycle are destroyed as soon as they resolve.
   */
  private _settleInteractiveSlot(
    i: number,
    slot: SlideSlot,
    widthPx: number,
    dpr: number,
    scale: number,
    epoch: number,
  ): void {
    if (!this._pres) return;
    const generation = ++slot.presentationGeneration;
    const spare = document.createElement('canvas');
    spare.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(spare);
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive;
    const onTextRun = wantRuns ? (r: PptxTextRunInfo) => runs.push(r) : undefined;

    this._pres
      .presentSlide(spare, i, {
        width: widthPx,
        dpr,
        onTextRun,
        onError: (error) => {
          if (generation === slot.presentationGeneration) this._reportRenderError(error);
        },
      })
      .then((handle) => {
        if (
          generation !== slot.presentationGeneration ||
          !slot.mediaInteractive ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedSlide !== i
        ) {
          handle.destroy();
          return;
        }

        const oldCanvas = slot.canvas;
        const oldHandle = slot.presentationHandle;
        slot.dispatcher.destroy();
        slot.wrapper.insertBefore(spare, oldCanvas);
        oldCanvas.remove();
        slot.canvas = spare;
        slot.dispatcher = new StaticCanvasRenderDispatcher(spare, false);
        slot.presentationHandle = handle;
        slot.renderedScale = scale;
        oldHandle?.destroy();

        if (slot.textLayer) {
          slot.textLayer.style.transform = '';
          slot.textLayer.style.transformOrigin = '';
          if (wantOverlay) {
            buildPptxTextLayer(
              slot.textLayer,
              runs,
              Math.round(widthPx),
              Math.round(this._slideHeightPx()),
              this._hyperlinkHandler(),
              i,
            );
          }
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitSlotComments(i, slot);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        if (generation === slot.presentationGeneration) this._reportRenderError(err);
      });
  }

  /**
   * Scroll so slide `index`'s top edge sits at the viewport top. Clamps `index` to
   * `[0, slideCount-1]` (the pager convention) and the resulting scrollTop to
   * `[0, totalHeight − viewportHeight]` so the last slides don't scroll past the
   * end. A no-op when nothing is loaded or the deck is empty.
   *
   * `opts.behavior` ('auto' | 'smooth', default 'auto') is honoured via
   * `scrollHost.scrollTo({ top, behavior })` when the host supports it (a real
   * browser); the stub-DOM has no `scrollTo`, so the fallback sets `scrollTop`
   * directly (which is what the tests assert). We then call `_mountVisible` once.
   *
   * MOUNTING CAVEAT: synchronous mounting of the target slide is guaranteed only on
   * the DEFAULT/'auto' path — there `scrollTop` has already jumped to `top`, so the
   * `_mountVisible` call reads the final scroll position and the target slide's slots
   * exist immediately. With `behavior: 'smooth'` the scroll animates ASYNCHRONOUSLY:
   * `scrollTop` is still near the old position when `_mountVisible` runs, so the
   * target slide mounts lazily via the animation's subsequent `scroll` events, not
   * from this call.
   */
  scrollToSlide(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    if (!this._pres || this._pres.slideCount === 0 || !this._scaleEstablished) return;
    const clamped = Math.max(0, Math.min(index, this._pres.slideCount - 1));
    const r = this._rangeAt(0, this._overscan());
    const target = this._slideOffset(clamped);
    const maxTop = Math.max(0, r.totalHeight - this._scrollHost.clientHeight);
    const top = Math.min(maxTop, Math.max(0, target));
    const host = this._scrollHost as HTMLDivElement & {
      scrollTo?: (opts: { top: number; behavior?: 'auto' | 'smooth' }) => void;
    };
    if (typeof host.scrollTo === 'function') {
      host.scrollTo({ top, behavior: opts?.behavior ?? 'auto' });
    } else {
      this._scrollHost.scrollTop = top;
    }
    this._mountVisible();
  }

  private _scrollToSlideCommentTarget(
    slide: number,
    comment: Readonly<PptxComment>,
    opts?: { behavior?: 'auto' | 'smooth' },
    resolvedBounds?: Readonly<{ x: number; y: number; width: number; height: number }>,
  ): boolean {
    if (!this._pres) return false;
    const slot = this._slots.get(slide);
    const boundsById = new Map(
      (slot?.commentElementBounds ?? []).map((entry) => [entry.elementId, entry.bounds]),
    );
    const anchored = resolvedBounds ?? (comment.anchors ?? []).flatMap((anchor) => {
      if ((anchor.type !== 'drawingElement' && anchor.type !== 'textRange') || !anchor.elementId) {
        return [];
      }
      const bounds = boundsById.get(anchor.elementId);
      return bounds ? [bounds] : [];
    })[0];
    const anchors = comment.anchors ?? [];
    const hasPosition = Number.isFinite(comment.x) && Number.isFinite(comment.y) && (
      anchors.length === 0 || anchors.some((anchor) => anchor.type === 'slide')
    );
    if (!anchored && !hasPosition) return false;
    const x = anchored
      ? anchored.x + (hasPosition ? comment.x as number : anchored.width)
      : comment.x as number;
    const y = anchored
      ? anchored.y + (hasPosition ? comment.y as number : 0)
      : comment.y as number;
    const width = this._slideWidthPx();
    const { left: paddingLeft } = this._padH();
    const slideLeft = Math.max(
      paddingLeft,
      (this._scrollHost.clientWidth - width) / 2,
    ) + this._reviewOriginPx;
    const range = this._rangeAt(0, this._overscan());
    const maxTop = Math.max(0, range.totalHeight - this._scrollHost.clientHeight);
    const spacerWidth = this._spacer.offsetWidth || Number.parseFloat(this._spacer.style.width) || 0;
    const maxLeft = Math.max(0, spacerWidth - this._scrollHost.clientWidth);
    const targetX = x / EMU_PER_PX * this._scale;
    const targetY = y / EMU_PER_PX * this._scale;
    const top = Math.min(maxTop, Math.max(
      0,
      this._slideOffset(slide) + targetY - this._scrollHost.clientHeight / 2,
    ));
    const left = Math.min(maxLeft, Math.max(
      0,
      slideLeft + targetX - this._scrollHost.clientWidth / 2,
    ));
    const host = this._scrollHost as HTMLDivElement & {
      scrollTo?: (options: {
        top: number;
        left: number;
        behavior?: 'auto' | 'smooth';
      }) => void;
    };
    if (typeof host.scrollTo === 'function') {
      host.scrollTo({ top, left, behavior: opts?.behavior ?? 'auto' });
    } else {
      this._scrollHost.scrollTop = top;
      this._scrollHost.scrollLeft = left;
    }
    this._mountVisible();
    return true;
  }

  private async _resolveSlideCommentElementBounds(
    slide: number,
    comment: Readonly<PptxComment>,
  ): Promise<Readonly<{ x: number; y: number; width: number; height: number }> | undefined> {
    const presentation = this._pres;
    if (!presentation) return undefined;
    const elementIds = (comment.anchors ?? []).flatMap((anchor) =>
      (anchor.type === 'drawingElement' || anchor.type === 'textRange') && anchor.elementId
        ? [anchor.elementId]
        : []);
    if (elementIds.length === 0) return undefined;
    const cached = new Map(
      (this._slots.get(slide)?.commentElementBounds ?? [])
        .map((entry) => [entry.elementId, entry.bounds]),
    );
    const cachedTarget = elementIds.flatMap((elementId) => {
      const bounds = cached.get(elementId);
      return bounds ? [bounds] : [];
    })[0];
    if (cachedTarget) return cachedTarget;
    const bounds = await presentation.getElementBoundsByIds(slide, elementIds);
    return elementIds.flatMap((elementId) => {
      const entry = bounds.find((candidate) => candidate.elementId === elementId);
      return entry ? [entry.bounds] : [];
    })[0];
  }

  /**
   * Reveal one authored comment occurrence from an application-owned list and
   * scroll its anchored element or authored slide point into view.
   * `commentIndex` is its index in `presentation.getComments(slideIndex)`.
   * Returns `false` when either index does not identify a comment.
   */
  async goToComment(
    slideIndex: number,
    commentIndex: number,
    opts?: { behavior?: 'auto' | 'smooth' },
  ): Promise<boolean> {
    if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
    const presentation = this._pres;
    if (!presentation || !Number.isInteger(slideIndex) || !Number.isInteger(commentIndex)) {
      return false;
    }
    if (slideIndex < 0 || slideIndex >= presentation.slideCount || commentIndex < 0) return false;
    const generation = this._beginCommentNavigation();
    if (slideIndex >= presentation.availableSlideCount && !presentation.layoutComplete) {
      const available = await this._waitForSlideMetadata(presentation, slideIndex, generation);
      if (!available) return false;
    }
    if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
    if (generation !== this._commentNavigationGeneration || presentation !== this._pres) {
      return false;
    }
    const comment = presentation.getComments(slideIndex)[commentIndex];
    if (!comment) return false;

    const bounds = await this._resolveSlideCommentElementBounds(slideIndex, comment);
    if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
    if (generation !== this._commentNavigationGeneration || presentation !== this._pres) {
      return false;
    }
    const anchors = comment.anchors ?? [];
    const hasSlidePoint = Number.isFinite(comment.x) && Number.isFinite(comment.y) && (
      anchors.length === 0 || anchors.some((anchor) => anchor.type === 'slide')
    );
    if (!bounds && !hasSlidePoint) return false;
    this.scrollToSlide(slideIndex, opts);
    if (!this._scrollToSlideCommentTarget(slideIndex, comment, opts, bounds)) return false;
    this._activeCommentId = pptxCommentOccurrenceKey(comment, commentIndex, slideIndex);
    this._activeCommentSlide = slideIndex;
    this._elementContext = null;
    for (const [mountedSlide, slot] of this._slots) {
      this._redrawSlotComments(mountedSlide, slot);
    }
    this._emitSelectionContextChange();
    return true;
  }

  /** Search the complete presentation, including slides outside the
   * virtualized mounted window. Matching is case-insensitive by default. */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<PptxMatchLocation>[]> {
    const presentation = this._pres;
    if (!presentation) return [];
    const generation = ++this._findGeneration;
    this._findActive = query.length > 0;
    if (query.length === 0) {
      this._find.invalidate();
      this._redrawHighlights();
      return [];
    }
    if (!presentation.layoutComplete) {
      await this._errorRouter.ownBackgroundLifecycle(
        () => presentation.waitUntilLayoutComplete(),
      );
    }
    if (this._destroyed || generation !== this._findGeneration || presentation !== this._pres) return [];
    const matches = await this._errorRouter.ownAwaitable(() => this._find.find(query, opts));
    if (this._destroyed || generation !== this._findGeneration || presentation !== this._pres) return [];
    this._redrawHighlights();
    return matches;
  }

  /** Activate and reveal the next match, wrapping at the end. */
  async findNext(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** Activate and reveal the previous match, wrapping at the beginning. */
  async findPrev(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** Clear the current query and every mounted highlight. */
  clearFind(): void {
    this._findActive = false;
    this._invalidateFind();
    this._redrawHighlights();
  }

  private _invalidateFind(): void {
    this._findGeneration++;
    this._find.invalidate();
  }

  private async _activateMatch(
    match: FindMatch<PptxMatchLocation> | null,
  ): Promise<FindMatch<PptxMatchLocation> | null> {
    if (match) this.scrollToSlide(match.location.slide);
    this._redrawHighlights();
    return match;
  }

  private async _collectSlideRuns(slide: number): Promise<PptxTextRunInfo[]> {
    if (!this._pres) return [];
    return this._pres.collectSlideRuns(slide, this._slideWidthPx());
  }

  private _redrawHighlights(): void {
    for (const [slide, slot] of this._slots) this._redrawSlotHighlights(slide, slot);
  }

  private _refreshFindRuns(slide: number, runs: PptxTextRunInfo[]): void {
    if (this._findActive) this._find.setSlideRuns(slide, runs);
  }

  private _redrawSlotComments(slide: number, slot: SlideSlot): void {
    if (!this._pres || !slot.commentMarkerLayer) return;
    const commentUi = this._commentUi;
    if (!commentUi) {
      slot.commentMarkerLayer.replaceChildren();
      slot.commentMargin?.replaceChildren();
      slot.commentDecorationLayer?.replaceChildren();
      slot.commentGeometry = null;
      return;
    }
    slot.commentGeometry = commentUi.buildPptxCommentMargin(
      slot.commentMarkerLayer,
      slot.commentMargin,
      this._pres.getComments(slide),
      slot.commentElementBounds,
      slide,
      this._pres.slideWidth,
      this._pres.slideHeight,
      this._activeCommentId,
      (id, active) => {
        const next = active ? id : this._activeCommentId === id ? null : this._activeCommentId;
        if (next === this._activeCommentId) return;
        this._activeCommentId = next;
        this._activeCommentSlide = next ? slide : null;
        this._elementContext = null;
        for (const [mountedSlide, mountedSlot] of this._slots) {
          this._redrawSlotComments(mountedSlide, mountedSlot);
        }
        this._emitSelectionContextChange();
      },
      this._commentZoom(),
      COMMENT_MARGIN_WIDTH_PX,
      this._commentsOptions()?.markers !== false,
      this._commentsOptions()?.includeResolved === true,
      slot.commentDecorationLayer
        ? () => this._scheduleCommentGeometry(slide, slot, false)
        : undefined,
      slot.commentDecorationLayer
        ? () => this._scheduleCommentGeometry(slide, slot, true)
        : undefined,
    );
    this._redrawSlotCommentConnectors(slide, slot);
  }

  private _redrawSlotCommentConnectors(slide: number, slot: SlideSlot): void {
    const layer = slot.commentDecorationLayer;
    const margin = slot.commentMargin;
    const geometry = slot.commentGeometry;
    const connectorOptions = this._commentsOptions()?.connectors;
    if (!layer || !margin || !geometry || !connectorOptions) return;
    const width = this._slideWidthPx();
    const height = this._slideHeightPx();
    const side = this._commentSide();
    const marginExtent = this._commentMarginExtent();
    const commentUi = this._commentUi;
    if (!commentUi) return;
    commentUi.buildReadOnlyCommentDecoration(
      layer,
      Object.freeze({
        surfaceBounds: Object.freeze({
          x: side === 'left' ? -marginExtent : 0,
          y: 0,
          width: width + marginExtent,
          height,
        }),
        contentBounds: Object.freeze({ x: 0, y: 0, width, height }),
        side,
        threads: commentUi.projectReadOnlyCommentMarginScroll(geometry, margin.scrollTop),
      }),
      {
        route: connectorOptions.route ?? 'bezier',
        stroke: connectorOptions.stroke ?? 'solid',
        color: connectorOptions.color,
        activeColor: connectorOptions.activeColor,
      },
    );
  }

  /** Commit comment geometry and reveal every comment layer in one settled frame. */
  private _commitSlotComments(slide: number, slot: SlideSlot): void {
    this._ensureSlotCommentAnchors(slide, slot);
    this._redrawSlotComments(slide, slot);
    if (slot.commentMarkerLayer) slot.commentMarkerLayer.style.visibility = '';
    if (slot.commentMargin) slot.commentMargin.style.visibility = '';
    if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = '';
  }

  private _ensureSlotCommentAnchors(slide: number, slot: SlideSlot): void {
    const presentation = this._pres;
    if (!presentation || slot.commentAnchorSlide === slide) return;
    slot.commentAnchorSlide = slide;
    slot.commentElementBounds = Object.freeze([]);
    const elementIds = [...new Set(presentation.getComments(slide).flatMap((comment) =>
      (comment.anchors ?? []).flatMap((anchor) =>
        (anchor.type === 'drawingElement' || anchor.type === 'textRange') && anchor.elementId
          ? [anchor.elementId]
          : [])))];
    if (elementIds.length === 0) return;
    const generation = ++slot.commentAnchorGeneration;
    void presentation.getElementBoundsByIds(slide, elementIds).then((bounds) => {
      if (this._destroyed || generation !== slot.commentAnchorGeneration ||
          presentation !== this._pres || this._slots.get(slide) !== slot ||
          slot.commentAnchorSlide !== slide) return;
      slot.commentElementBounds = bounds;
      this._redrawSlotComments(slide, slot);
      const active = presentation.getComments(slide).find((comment, index) =>
        pptxCommentOccurrenceKey(comment, index, slide) === this._activeCommentId);
      if (active) this._scrollToSlideCommentTarget(slide, active);
    }).catch((error: unknown) => {
      if (!this._destroyed && generation === slot.commentAnchorGeneration) {
        this._reportRenderError(error);
      }
    });
  }

  /** Coalesce card measurement and scroll-only connector projection into one
   * geometry refresh per animation frame. A full refresh dominates a pending
   * connector-only refresh for the same slot. */
  private _scheduleCommentGeometry(
    slide: number,
    slot: SlideSlot,
    connectorsOnly = false,
  ): void {
    const current = this._pendingCommentGeometry.get(slide);
    this._pendingCommentGeometry.set(slide, {
      slot,
      connectorsOnly: current?.slot === slot
        ? current.connectorsOnly && connectorsOnly
        : connectorsOnly,
    });
    if (this._commentGeometryScheduled) return;
    this._commentGeometryScheduled = true;
    const flush = (): void => {
      this._commentGeometryScheduled = false;
      this._commentGeometryFrame = null;
      const pending = [...this._pendingCommentGeometry];
      this._pendingCommentGeometry.clear();
      if (this._destroyed) return;
      for (const [pendingSlide, entry] of pending) {
        const { slot: pendingSlot, connectorsOnly: pendingConnectorsOnly } = entry;
        if (
          this._slots.get(pendingSlide) === pendingSlot &&
          pendingSlot.renderedScale === this._scale
        ) {
          if (pendingConnectorsOnly) {
            this._redrawSlotCommentConnectors(pendingSlide, pendingSlot);
          } else {
            this._redrawSlotComments(pendingSlide, pendingSlot);
          }
        }
      }
    };
    const ownerWindow = this._wrapper.ownerDocument.defaultView;
    if (ownerWindow?.requestAnimationFrame) {
      this._commentGeometryFrame = ownerWindow.requestAnimationFrame(flush);
    }
    else queueMicrotask(flush);
  }

  private _redrawSlotHighlights(slide: number, slot: SlideSlot): void {
    if (!this._findActive) {
      slot.highlightLayer.innerHTML = '';
      return;
    }
    const runs = this._find.slideRuns(slide);
    if (!runs) {
      slot.highlightLayer.innerHTML = '';
      return;
    }
    buildPptxHighlightLayer(
      slot.highlightLayer,
      runs,
      this._find.slideHighlights(slide),
      this._slideWidthPx(),
      this._slideHeightPx(),
      (font) => this._measureForFind(font),
      this._opts.findHighlightColors,
    );
  }

  private _measureForFind(font: string): (text: string) => number {
    if (this._findMeasureCtx === undefined) {
      const canvas = document.createElement('canvas');
      this._findMeasureCtx = canvas.getContext('2d');
    }
    const ctx = this._findMeasureCtx;
    if (!ctx || typeof ctx.measureText !== 'function') return (text) => text.length;
    ctx.font = font;
    return (text) => ctx.measureText(text).width;
  }

  /**
   * IX1 hyperlink click dispatch (mirrors {@link PptxViewer._onHyperlinkClick}).
   * When the integrator supplies `opts.onHyperlinkClick` it OWNS the click (no
   * default). Otherwise: an external link opens in a new tab via the shared,
   * scheme-sanitised {@link openExternalHyperlink}; an internal slide jump scrolls
   * to the target slide via {@link scrollToSlide} once the action resolves to a
   * slide index (a jump resolving to no reachable slide is a safe no-op).
   */
  /**
   * IX1 — the click handler passed to the text-layer overlay, or `undefined` when
   * `enableHyperlinks` is `false`. This is the single gate that disables hyperlink
   * interactivity: {@link buildPptxTextLayer} renders link runs exactly like plain
   * runs when no handler is supplied, so no hit region, cursor, tooltip, listener,
   * or navigation is wired (a custom `onHyperlinkClick` is suppressed too). When
   * enabled, the returned handler dispatches through {@link _onHyperlinkClick}.
   */
  private _hyperlinkHandler(): ((target: HyperlinkTarget) => void) | undefined {
    if (this._opts.enableHyperlinks === false) return undefined;
    return (t) => this._onHyperlinkClick(t);
  }

  private _onHyperlinkClick(target: HyperlinkTarget): void {
    const enriched = this._resolveInternalSlideIndex(target);
    if (this._opts.onHyperlinkClick) {
      this._opts.onHyperlinkClick(enriched);
      return;
    }
    if (enriched.kind === 'external') {
      openExternalHyperlink(enriched.url);
      return;
    }
    if (enriched.slideIndex !== undefined) this.scrollToSlide(enriched.slideIndex);
  }

  /** Populate an internal {@link HyperlinkTarget}'s `slideIndex` from its `ref`
   *  via the engine's stamped part names. Relative `hlinkshowjump` verbs are
   *  resolved against the slide currently at the viewport top
   *  (`_range().topIndex`); a `../slides/slideN.xml` part target resolves through
   *  the part-name map. An already-set index, an external target, and an
   *  unresolvable ref all pass through unchanged (safe no-op). */
  private _resolveInternalSlideIndex(target: HyperlinkTarget): HyperlinkTarget {
    if (target.kind !== 'internal' || target.slideIndex !== undefined) return target;
    const idx = this._pres?.resolveInternalTarget(target.ref, this._range().topIndex);
    return idx === undefined ? target : { ...target, slideIndex: idx };
  }

  /**
   * Re-fit the base scale on a container resize while PRESERVING the current zoom
   * multiplier (design §11), then re-anchor + re-render. A `ResizeObserver` fires
   * on any box change, but only a WIDTH change alters the fit-to-width base scale;
   * a height-only change skips the re-fit yet STILL re-mounts the visible window
   * (via `_mountVisible`), because a taller viewport reveals rows that were below
   * the fold and would otherwise stay blank until the next scroll. Empty/unloaded
   * ⇒ no-op; a still-zero width ⇒ defer.
   *
   * Zero-width recovery: a container that was 0-wide at construction never
   * established a scale (`_scaleEstablished` is false), so the first non-zero
   * resize establishes it here via `relayout()` — completing the T2 deferral.
   *
   * Re-fit math (zoom multiplier preserved):
   *   mult      = _scale / _prevBase            (the user's zoom over the old base)
   *   newScale  = newBase × mult
   * Routing through `setScale(newScale)` bumps `_renderEpoch` (resize IS an epoch
   * event — T4 banner) and re-anchors + CSS-previews + debounces a settle re-render
   * of every slot at the new geometry, exactly like a zoom (design §7 flicker-free
   * path — a rapid ResizeObserver burst therefore also coalesces into one settle).
   * `setScale`'s clamp/no-op guards apply: an unchanged newScale (identical width)
   * is a no-op there — so we short-circuit BEFORE it when the fit-width is
   * unchanged (mounting the revealed window without a needless re-render), and
   * after it we call `_mountVisible` again to cover the case where the clamp made
   * `setScale` no-op yet the viewport still grew.
   */
  private _onResize(): void {
    if (!this._pres || this._pres.slideCount === 0) return;
    // Zero-width recovery: first non-zero layout establishes the base scale.
    if (!this._scaleEstablished) {
      this.relayout();
      return;
    }
    if (this._opts.refitOnResize === false) {
      // Fixed-scale hosts (for example VS Code previews) must not turn a pane
      // resize into an implicit zoom. Recompute the visible window and horizontal
      // centering only; slide geometry and rendered bitmaps remain valid.
      this._lastFitWidth = this._fitWidthPx();
      this._mountVisible();
      return;
    }
    const newBase = this._baseScale();
    if (newBase <= 0) return; // still unlaid-out — wait for the next resize
    const newFitWidth = this._fitWidthPx();
    if (newFitWidth === this._lastFitWidth) {
      // Height-only change (or any resize that leaves the fit-width identical):
      // the base scale is unchanged, so there is no re-fit to do — but a taller
      // viewport now exposes rows that were below the fold. `_mountVisible`
      // recomputes the visible range from the CURRENT clientHeight and mounts the
      // newly-revealed slides; without it those rows stay blank until the user
      // scrolls (which recomputes the range). No epoch bump — the geometry
      // (and every mounted slot's px size) is unchanged, so cached canvases are
      // still valid; we only add the missing slots.
      this._mountVisible();
      return;
    }
    this._lastFitWidth = newFitWidth;
    // Preserve the zoom multiplier across the re-fit: newScale = newBase × mult.
    const mult = this._prevBase > 0 ? this._scale / this._prevBase : 1;
    this._prevBase = newBase;
    // Route through setScale so the epoch bumps and the re-anchor/force-re-render
    // path runs identically to a zoom.
    //
    // zoomMin RATCHET (design §8.2 caveat, see setScale JSDoc): `zoomMin`/`zoomMax`
    // are ABSOLUTE dimensionless bounds, but the re-fit base (`newBase × mult`) is
    // computed UNCLAMPED. A resize that transits the scale below `zoomMin` (a wide
    // slide in a container that briefly narrows) is clamped UP by `setScale`,
    // which permanently inflates the implied multiplier even with zero user zoom —
    // the next re-fit reads back the clamped `_scale` as `mult`. This is bounded and
    // converges (the clamp floor is fixed), but it means the preserved multiplier can
    // drift above 1 purely from resize transits below the floor. Accepted consequence
    // of using absolute bounds (§8.2) with an unclamped relayout base.
    this.setScale(newBase * mult);
    // `setScale` no-ops when the clamped scale is unchanged (e.g. already pinned at
    // a clamp boundary), which would skip its preview + settle. A width+height
    // growth that ends up clamped to the same scale must still reveal the taller
    // viewport's rows, so mount here too. Idempotent when `setScale` ran: the
    // window is already mounted and every present slot is a re-position no-op.
    this._mountVisible();
  }

  get topVisibleSlide(): number {
    return this._lastRange?.topIndex ?? 0;
  }

  /** @internal test hook: slide indices currently mounted. */
  mountedSlideIndicesForTest(): number[] {
    return [...this._slots.keys()];
  }

  /** @internal test hook: slots currently owning or awaiting media handles. */
  interactiveSlideIndicesForTest(): number[] {
    return [...this._slots]
      .filter(([, slot]) => slot.mediaInteractive)
      .map(([index]) => index);
  }

  /** @internal test hook: the current absolute (dimensionless) zoom scale. */
  scaleForTest(): number {
    return this._scale;
  }

  /** @internal test hook: the base fit scale (pre-zoom) at the current width. */
  baseScaleForTest(): number {
    return this._baseScale();
  }

  /** @internal test hook: the current render epoch (bumped on setScale + resize). */
  renderEpochForTest(): number {
    return this._renderEpoch;
  }

  /** @internal test hook: fire the observed resize path (a real host drives this
   *  via the constructor's ResizeObserver). */
  resizeForTest(): void {
    this._onResize();
  }

  /** @internal test hook: the content point (slide index + intra-slide fraction)
   *  currently under viewport-y `y` (px from the scroll host top). Lets a test
   *  capture "what is under the cursor" before a zoom and re-query its on-screen
   *  y afterwards to assert the pointer-anchored invariant. */
  contentAtViewportYForTest(y: number): { slide: number; frac: number } {
    const contentY = this._scrollHost.scrollTop + y;
    const slide = this._slideIndexAtOffset(contentY);
    const h = this._uniformSlideHeight;
    const frac = h > 0 ? Math.min(1, Math.max(0, (contentY - this._slideOffset(slide)) / h)) : 0;
    return { slide, frac };
  }

  /** @internal test hook: inverse of {@link contentAtViewportYForTest} — the
   *  current viewport-y (px from the scroll host top) of the content point at
   *  (`slide`, intra-slide `frac`). */
  viewportYOfForTest(slide: number, frac: number): number {
    const contentY = this._slideOffset(slide) + frac * this._uniformSlideHeight;
    return contentY - this._scrollHost.scrollTop;
  }

  /** Return the owning engine's latest content-free package-usage snapshot. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this._pres) throw new Error('Presentation not loaded');
    return await this._pres.getResourceMetrics();
  }

  /** Return the current mounted browser text selection with PPTX source locators. */
  getSelectionContext(options: PptxSelectionContextOptions = {}): PptxSelectionContext | null {
    if (this._destroyed) throw new Error('PptxScrollViewer is destroyed');
    if (this._pres && this._activeCommentId !== null && this._activeCommentSlide !== null) {
      const comments = this._pres.getComments(this._activeCommentSlide);
      const commentIndex = comments.findIndex((comment, index) =>
        pptxCommentOccurrenceKey(comment, index, this._activeCommentSlide as number) ===
          this._activeCommentId);
      const entry = comments[commentIndex];
      if (entry && commentIndex >= 0) {
        return createPptxCommentSelectionContext(
          entry,
          this._activeCommentSlide,
          commentIndex,
          this._activeCommentId,
          options,
        );
      }
    }
    const text = this._opts.enableTextSelection
      ? readPptxTextSelectionContext(
          this._wrapper,
          this._wrapper.ownerDocument?.getSelection?.() ?? null,
          options,
        )
      : null;
    return text ?? (this._elementContext
      ? limitPptxElementContext(
          this._elementContext,
          options.maxTextCharacters,
        )
      : null);
  }

  private _emitSelectionContextChange(): void {
    const context = this.getSelectionContext();
    if (context?.kind === 'text') {
      this._elementHitGeneration++;
      this._elementContext = null;
      this._redrawElementOutlines();
    }
    const key = JSON.stringify(context);
    if (key === this._selectionContextKey) return;
    this._selectionContextKey = key;
    this._opts.onSelectionContextChange?.(context ? structuredClone(context) : null);
  }

  private _setElementContext(context: PptxElementContext | null): void {
    this._elementContext = context ? structuredClone(context) : null;
    this._redrawElementOutlines();
    this._emitSelectionContextChange();
  }

  private _invalidateElementSelection(notify = true): void {
    this._elementHitGeneration++;
    this._elementContext = null;
    this._redrawElementOutlines();
    if (notify) this._emitSelectionContextChange();
  }

  private _redrawElementOutlines(): void {
    for (const [slide, slot] of this._slots) this._redrawElementOutlineForSlot(slide, slot);
  }

  private _redrawElementOutlineForSlot(slideIndex: number, slot: SlideSlot): void {
    const context = this._elementContext;
    const presentation = this._pres;
    if (!context || !presentation || context.slideIndex !== slideIndex) {
      renderCanvasElementOutline(slot.elementLayer, null);
      return;
    }
    renderCanvasElementOutline(slot.elementLayer, {
      x: context.bounds.x / presentation.slideWidth,
      y: context.bounds.y / presentation.slideHeight,
      width: context.bounds.width / presentation.slideWidth,
      height: context.bounds.height / presentation.slideHeight,
      rotation: context.bounds.rotation,
    });
  }

  private async _onElementClick(event: MouseEvent): Promise<void> {
    if (this._destroyed || event.defaultPrevented || event.button !== 0) return;
    await this._resolveContextAt(event);
  }

  private _onContextMenu(event: MouseEvent): void {
    let context: Promise<PptxSelectionContext | null> | undefined;
    this._opts.onContextMenu?.({
      originalEvent: event,
      getContext: () => context ??= this._resolveContextAt(event),
    });
  }

  private async _resolveContextAt(event: MouseEvent): Promise<PptxSelectionContext | null> {
    const presentation = this._pres;
    if (this._destroyed || !presentation) return null;
    if (this._opts.enableTextSelection && readPptxTextSelectionContext(
      this._wrapper,
      this._wrapper.ownerDocument?.getSelection?.() ?? null,
    )) {
      this._emitSelectionContextChange();
      return this._destroyed ? null : this.getSelectionContext();
    }
    if (!this._opts.enableElementSelection) return this.getSelectionContext();
    const target = event.target as Node | null;
    const entry = [...this._slots].find(([, slot]) => target !== null && slot.wrapper.contains(target));
    if (!entry) {
      this._invalidateElementSelection();
      return null;
    }
    const [slideIndex, slot] = entry;
    const rect = slot.canvas.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      this._invalidateElementSelection();
      return null;
    }
    const localX = event.clientX - rect.left;
    const localY = event.clientY - rect.top;
    if (localX < 0 || localY < 0 || localX > rect.width || localY > rect.height) {
      this._invalidateElementSelection();
      return null;
    }
    const generation = ++this._elementHitGeneration;
    const point = {
      x: localX / rect.width * presentation.slideWidth,
      y: localY / rect.height * presentation.slideHeight,
    };
    let context: PptxElementContext | null;
    try {
      context = await presentation.getElementContextAt(slideIndex, point, {
        tolerance: this._elementHitTolerance / rect.width * presentation.slideWidth,
        maxTextCharacters: MAX_ELEMENT_TEXT_CHARACTERS,
      });
    } catch (error) {
      if (this._destroyed || generation !== this._elementHitGeneration ||
        presentation !== this._pres) return null;
      throw error;
    }
    if (this._destroyed || generation !== this._elementHitGeneration || presentation !== this._pres) return null;
    this._setElementContext(context);
    return this._destroyed ? null : this.getSelectionContext();
  }

  /**
   * Tear down the viewer: remove the DOM subtree and (only for a self-loaded
   * engine) destroy the engine. A borrowed engine is left intact — the caller
   * owns its lifecycle. Per-slot worker ImageBitmaps are closed on recycle.
   */
  destroy(): void {
    if (this._destroyed) return;
    this._destroyed = true;
    this._beginCommentNavigation();
    this._errorRouter.close();
    this._invalidateFind();
    this._findActive = false;
    this._unbindLayoutPresentation();
    if (this._selectionChangeListener) {
      this._wrapper.ownerDocument.removeEventListener('selectionchange', this._selectionChangeListener);
      this._selectionChangeListener = null;
    }
    this._elementHitGeneration++;
    if (this._elementClickListener) {
      this._scrollHost.removeEventListener('click', this._elementClickListener);
      this._elementClickListener = null;
    }
    if (this._contextMenuListener) {
      this._scrollHost.removeEventListener('contextmenu', this._contextMenuListener);
      this._contextMenuListener = null;
    }
    if (this._commentOutsidePointerListener) {
      this._wrapper.ownerDocument.removeEventListener(
        'pointerdown',
        this._commentOutsidePointerListener,
      );
      this._commentOutsidePointerListener = null;
    }
    if (this._commentGeometryFrame !== null) {
      this._wrapper.ownerDocument.defaultView?.cancelAnimationFrame?.(this._commentGeometryFrame);
      this._commentGeometryFrame = null;
    }
    this._commentGeometryScheduled = false;
    this._pendingCommentGeometry.clear();
    this._elementContext = null;
    if (this._scrollListener) {
      this._scrollHost.removeEventListener('scroll', this._scrollListener);
      this._scrollListener = null;
    }
    if (this._wheelListener) {
      this._scrollHost.removeEventListener('wheel', this._wheelListener as EventListener);
      this._wheelListener = null;
    }
    this._resizeObserver?.disconnect();
    this._resizeObserver = null;
    // Cancel a pending settle so no re-render is dispatched after teardown
    // (design §7 mechanism 2). `_destroyed` also guards `_settleRender`, but
    // clearing the timer avoids the wasted wake-up and keeps fake-timer tests
    // deterministic.
    if (this._settleTimer !== null) {
      clearTimeout(this._settleTimer);
      this._settleTimer = null;
    }
    for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
    this._free.length = 0;
    this._presentationOwner.close();
    this._wrapper.remove();
  }
}
