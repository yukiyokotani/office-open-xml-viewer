import { openExternalHyperlink, PT_TO_PX, zoomStepScale, anchoredZoomOffset, nextZoomStep, prevZoomStep, fitScale } from '@silurus/ooxml-core';
import type { FindHighlightColors, FindMatch, FindMatchesOptions, HyperlinkTarget, OoxmlResourceMetrics, ViewerContextMenuEvent, ZoomableViewer } from '@silurus/ooxml-core';
import {
  computeVisibleWindow,
  createVirtualScrollGeometry,
  type VirtualScrollGeometry,
  type VisibleRange,
} from '@silurus/ooxml-core/internal/virtual-scroll';
import {
  createCanvasElementOutlineLayer,
  CanvasViewerErrorRouter,
  renderCanvasElementOutline,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import { READ_ONLY_COMMENT_MARGIN_WIDTH_PX } from '@silurus/ooxml-core/internal/read-only-comment-contract';
import { eventTargetsDataAttributeWithin } from '@silurus/ooxml-core/internal/dom-interaction-boundary';
import type { ReadOnlyCommentMarginGeometry } from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import {
  activeDocxLayoutViewOf,
  selectDocxLayoutView,
  subscribeDocxLayoutView,
  type DocxLayoutViewPublication,
} from './document-layout-view.js';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import { DocxFindController, type DocxMatchLocation } from './find';
import { buildDocxHighlightLayer } from './find-highlight-layer';
import type { RenderPageOptions } from './types';
import {
  createDocxCommentSelectionContext,
  readDocxTextSelectionContext,
  type DocxElementContext,
  type DocxSelectionContext,
  type DocxSelectionContextOptions,
} from './selection-context';
import {
  limitDocxElementContext,
  MAX_DOCX_ELEMENT_TEXT_CHARACTERS,
} from './element-context';
import type { DocxCommentsOptions } from './comment-margin';
import { resolveCommentAnchorRuns } from './comments';
import { renderDocxFocusedPage } from './focused-view-runtime';
import {
  subscribeDocxLayout,
  type DocxLayoutPublication,
} from './document-layout-events.js';

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
 * Default CSS `box-shadow` painted on every page canvas — the soft drop shadow a
 * PDF reader casts under each sheet (matches the Examples/recipe look, which the
 * scroll viewer now reproduces with zero config). See
 * {@link DocxScrollViewerOptions.pageShadow}.
 */
const DEFAULT_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';
const COMMENT_MARGIN_GAP_PX = 12;
const COMMENT_MARGIN_FONT_SIZE_PX = 13;
const borrowedDocumentOption = Symbol('DocxScrollViewer.borrowedDocument');
type DocxCommentUiRuntime = typeof import('./comment-ui-runtime.js');
let docxCommentUiRuntimePromise: Promise<DocxCommentUiRuntime> | undefined;

function loadDocxCommentUiRuntime(): Promise<DocxCommentUiRuntime> {
  return docxCommentUiRuntimePromise ??= import('./comment-ui-runtime.js');
}

type InternalDocxScrollViewerOptions = DocxScrollViewerOptions & {
  [borrowedDocumentOption]?: DocxDocument;
};

/**
 * Options for {@link DocxScrollViewer}. Extends `RenderPageOptions` (per-page
 * render knobs, minus `onTextRun`) and `LoadOptions` (parse/worker knobs). See
 * design §8.1.
 *
 * `onTextRun` is omitted deliberately: the viewer drives it internally per
 * mounted slot to build the optional per-page selection overlay (gated by
 * `enableTextSelection`), so exposing it here would let a caller's callback be
 * silently overridden.
 */
export interface DocxScrollViewerOptions extends Omit<RenderPageOptions, 'onTextRun'>, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive pages. Default 16. */
  gap?: number;
  /** Desk padding (px) ABOVE the FIRST page — the margin a PDF reader leaves
   *  between the top of the scroll surface and the first sheet. Default: `gap`
   *  (uniform desk rhythm — the first page sits the same distance from the top as
   *  pages sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST page — the margin below the final sheet.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Desk gutter (px) to the LEFT of the pages — the horizontal margin between the
   *  left edge of the scroll surface and a page sitting flush-left (i.e. once
   *  zoomed wide enough that centering no longer applies). Default: `gap` (uniform
   *  desk rhythm — the horizontal gutters match the vertical ones). It also shrinks
   *  the container-derived FIT width so a page sits inside the gutters at 100%
   *  (an EXPLICIT `opts.width` is the page's CSS-width contract and is NOT reduced;
   *  the gutters still apply around placement). Pass `0` for a flush-left layout. */
  paddingLeft?: number;
  /** Desk gutter (px) to the RIGHT of the pages. Default: `gap`. Shrinks the
   *  container-derived fit width symmetrically with `paddingLeft`. Pass `0` for a
   *  flush-right layout. */
  paddingRight?: number;
  /** Pages kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /**
   * Paint the document's opening pages as soon as they are laid out, instead of
   * waiting for the whole document.
   *
   * A large document otherwise shows nothing until every page has been
   * paginated. With this on, the viewer mounts the first pages within a few
   * hundred milliseconds regardless of document length, the scrollbar grows as
   * the rest of the layout arrives, and the viewer relays out once it lands.
   * Published pages are provisional: later header/footer, field, anchored-object,
   * or section convergence can still replace them. Mounted pages repaint when
   * the authoritative layout lands.
   *
   * `pageCount` therefore starts small and grows; {@link findText} waits for the
   * full layout internally. Ignored by `fromDocument`, whose document is already
   * loaded. Works in both render modes and in either tracked-changes view; in
   * `mode: 'worker'` it additionally keeps the remaining pagination off the
   * main thread.
   */
  progressiveLayout?: boolean;
  /**
   * Lay the document out in slices, keeping the thread responsive while a large
   * document paginates. `load()` still resolves only once layout is complete —
   * use {@link progressiveLayout} to paint before then.
   */
  sliceLayout?: boolean;
  /** Per-page transparent text-selection overlay. IX6 — works in BOTH render
   *  modes: in worker mode the per-run geometry is collected off-thread and
   *  shipped back beside the page bitmap, so the overlay is populated identically
   *  to main mode (no more empty overlay / one-time warning). */
  enableTextSelection?: boolean;
  /** Show the built-in read-only comments. Pass options to configure them. Default false. */
  comments?: boolean | DocxCommentsOptions;
  /**
   * Enable read-only selection of mounted pictures, charts, and shapes. The
   * selected object exposes element context and receives a non-editable outline.
   */
  enableElementSelection?: boolean;
  /** Emits bounded, detached text or element context suitable for read-only AI/MCP use. */
  onSelectionContextChange?: (context: DocxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the text or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<DocxSelectionContext>) => void;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /** Minimum zoom scale (px-per-pt multiplier floor). A smaller width-fit base
   * remains reachable as the effective minimum. Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale. Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * Re-fit the document to the container width when the container is resized.
   * Default true. Set false to preserve the current absolute scale, including
   * an explicit pre-load `setScale(1)`, independently of the viewport width.
   * Explicit `fitWidth()` and `fitPage()` calls remain available.
   */
  refitOnResize?: boolean;
  /**
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between pages — the gray a PDF reader paints around the sheet.
   * Applied to the viewer-owned scroll host. The pages themselves are always
   * drawn on the document's own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * CSS `box-shadow` painted on every page CANVAS (not the wrapper — the
   * text-selection overlay must not cast its own shadow). The soft drop shadow a
   * PDF reader leaves under each sheet.
   *
   * - Default (`undefined`): `'0 1px 3px rgba(0,0,0,0.2)'` — the recipe look, so
   *   the scroll viewer reproduces the Examples appearance with zero config.
   * - `false`: NO shadow (flat pages).
   * - A custom string is applied verbatim. A spread-only ring such as
   *   `'0 0 0 1px #c8ccd0'` gives a crisp 1px BORDER look — and because
   *   `box-shadow` never affects layout (unlike `border`, which would grow the
   *   box and shift every offset), a border and a drop shadow are the SAME knob
   *   here rather than two competing options.
   */
  pageShadow?: string | false;
  /** Fires when the top-most visible page OR the document's page count changes.
   *  `topIndex` from `computeVisibleRange` (the first page intersecting the
   *  viewport top, EXCLUDING overscan).
   *
   *  `layoutComplete` is false while progressive layout is still running, and
   *  `total` is then the pages laid out SO FAR, not the document's total — a
   *  "page X of Y" indicator should mark it provisional (Word shows an
   *  unsettled count the same way during background repagination). The count
   *  is watched as well as the index precisely so that indicator updates when
   *  the rest of the document arrives without the user scrolling. */
  onVisiblePageChange?: (
    topIndex: number,
    total: number,
    layoutComplete: boolean,
  ) => void;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100% = a page
   *  at its natural pt→px size): from {@link DocxScrollViewer.setScale},
   *  `zoomIn`/`zoomOut`, `fitWidth`/`fitPage`, a Ctrl/⌘+wheel gesture, or a
   *  container-resize re-fit (when `refitOnResize` is enabled). Named
   *  `onScaleChange` to match the single-canvas viewers so all five share one
   *  notification shape. */
  onScaleChange?: (scale: number) => void;
  /** IX1 (design decision — NOT user-confirmed, integrator may veto). Called when
   *  a hyperlink run is clicked. When omitted, the default is: external → open in a
   *  new tab via core `openExternalHyperlink` (sanitised, noopener,noreferrer);
   *  internal → jump to the page whose text contains the bookmark (best-effort). */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the hyperlink machinery is not wired at all: no overlay hit region
   *  is installed for link runs, so there is no pointer cursor, no title tooltip,
   *  no default navigation (external new-tab / internal bookmark jump), and
   *  `onHyperlinkClick` is never called. Links still render exactly as authored
   *  but are inert, like plain text. */
  enableHyperlinks?: boolean;
  /** Receives asynchronous Viewer-managed failures that cannot be observed by
   *  awaiting the method that started them. `load()` failures always reject and
   *  are not also delivered here. Virtualized per-slot render failures (both
   *  main `renderPage` and worker `renderPageToBitmap` rejections) invoke it; a
   *  failed page is left blank rather than crashing the loop. Without an
   *  `onError`, render failures are logged via
   *  `console.error` so they are never fully silent. Stable cases can be
   *  narrowed with `OoxmlError`, `OoxmlResourceLimitError`, or
   *  `OoxmlDecodedImageLimitError` re-exported by this package. Other failures
   *  remain `Error` values; a `code` of `parser-crashed` identifies a recognized
   *  WASM trap, not a reliably classified OOM. */
  onError?: (err: Error) => void;
}

/** One mounted page. `canvas` is the drawn page; `textLayer` the optional
 *  per-page selection overlay (both render modes — IX6 ships the worker's run
 *  geometry back beside the bitmap). `renderedPage` guards against
 *  re-rendering a recycled slot for a page whose render is still in flight. */
interface PageSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  highlightLayer: HTMLDivElement;
  elementLayer: HTMLDivElement | null;
  commentTintLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentDecorationLayer: HTMLDivElement | null;
  commentRuns: readonly Readonly<DocxTextRunInfo>[];
  commentGeometry: ReadOnlyCommentMarginGeometry | null;
  /** page index this slot is currently rendering / has rendered, or -1 when free. */
  renderedPage: number;
  /** The `_scale` at which this slot's on-screen canvas bitmap (and text overlay)
   *  were last rendered, or -1 when unrendered. The flicker-free CSS preview
   *  (design §7) stretches that bitmap to the new layout size on `setScale` and
   *  scales the text overlay by `newScale / renderedScale`; the debounced settle
   *  re-render then repaints at the new scale and updates this to match. */
  renderedScale: number;
  /** Shared single-canvas generation and worker-bitmap ownership primitive. */
  dispatcher: StaticCanvasRenderDispatcher;
}

export class DocxScrollViewer implements ZoomableViewer {
  private readonly _documentOwner: TerminalResourceOwner<DocxDocument>;
  private get _doc(): DocxDocument | null { return this._documentOwner.current; }
  private readonly _borrowed: boolean;
  private readonly _opts: DocxScrollViewerOptions;
  private readonly _errorRouter: CanvasViewerErrorRouter;
  private readonly _container: HTMLElement;
  private readonly _wrapper: HTMLDivElement;
  private readonly _scrollHost: HTMLDivElement;
  private readonly _spacer: HTMLDivElement;
  /** Resolved render mode. When an engine is borrowed the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `DocxDocument.load`. */
  private _mode: 'main' | 'worker';

  /** px-per-pt zoom multiplier. Base fit maps the widest page's width to the
   *  container width (or opts.width). Zoom multiplies this (design §7). */
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
  /** Live slots keyed by page index. */
  private readonly _slots = new Map<number, PageSlot>();
  /** Recyclable detached slots (canvas + textLayer reused across pages). */
  private readonly _free: PageSlot[] = [];
  /** Cached per-page heights in px at the current scale (index-aligned). */
  private _heights: number[] = [];
  /** Prefix offsets rebuilt only when scale/page geometry changes. Pure scroll
   * queries binary-search this cache instead of walking every document page. */
  private _scrollGeometry: VirtualScrollGeometry = { offsets: [], totalHeight: 0 };
  private _lastRange: VisibleRange | null = null;
  private _lastTopIndex = -1;
  /** Second half of the visible-page latch: a document that grows under the
   *  viewport changes `total` without changing `topIndex`. */
  private _lastReportedTotal = -1;
  /** Completion is observable callback state too: an authoritative publication
   * can replace a provisional layout without changing its page count. */
  private _lastReportedLayoutComplete: boolean | null = null;
  /** Subscription to the document currently installed by `_documentOwner`.
   *  Failed and stale acquisitions never replace it, so they cannot revoke the
   *  retained document's authority to publish background layout progress. */
  private _layoutUnsubscribe: (() => void) | null = null;
  /** Page prefix currently represented by the native scroll extent. */
  private _presentedPageCount = 0;
  private _scrollListener: (() => void) | null = null;
  private _selectionChangeListener: (() => void) | null = null;
  private _selectionContextKey = 'null';
  private _elementClickListener: ((event: MouseEvent) => void) | null = null;
  private _contextMenuListener: ((event: MouseEvent) => void) | null = null;
  private _commentOutsidePointerListener: ((event: PointerEvent) => void) | null = null;
  private _elementContext: DocxElementContext | null = null;
  private _activeCommentId: string | null = null;
  private _activeCommentPage: number | null = null;
  private _commentUi: DocxCommentUiRuntime | null = null;
  private readonly _commentPageById = new Map<string, number>();
  /** Page-run requests shared by comment navigation. The scale stamp prevents
   * geometry collected before a zoom from being reused for target scrolling. */
  private readonly _commentRunsByPage = new Map<number, {
    readonly scale: number;
    readonly runs: Promise<readonly Readonly<DocxTextRunInfo>[]>;
  }>();
  /** Pages whose runs have already been joined to every authored comment anchor. */
  private readonly _commentIndexedPages = new Set<number>();
  /** First page not yet included in the shared comment-page index. */
  private _commentScanFrontier = 0;
  /** Latest list-navigation request. Older async scans may populate caches but
   * must never restore their scroll/selection after a newer click. */
  private _commentNavigationGeneration = 0;
  /** Latest default internal-link navigation; later clicks supersede work that
   * is still waiting for the authoritative bookmark projection. */
  private _internalHyperlinkGeneration = 0;
  private _commentAnchorRangesForMargin: ReturnType<DocxDocument['commentAnchorRanges']> | null = null;
  private _commentAnchorIds: ReadonlySet<string> = new Set();
  /** Horizontal origin used only for a reachable left-hand review rail. */
  private _reviewOriginPx = 0;
  private _commentGeometryScheduled = false;
  private _commentGeometryFrame: number | null = null;
  private readonly _pendingCommentGeometry = new Map<number, {
    readonly slot: PageSlot;
    readonly connectorsOnly: boolean;
  }>();
  private _elementHitGeneration = 0;
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  /** Throwaway 2D context reused to measure text for the §17.3.2.10 縦中横 overlay
   *  clamp (#836). Lazily created; `null` when canvas metrics are unavailable
   *  (headless), in which case the overlay degrades to the un-clamped span. */
  private _measureCtx: CanvasRenderingContext2D | null | undefined;
  /** Worker mode: page indices whose bitmap render is currently dispatched to the
   *  engine. Coalesces a scroll storm — we never dispatch a second render for a
   *  page whose first is still in flight — and lets us drop pages that scrolled
   *  out of the window before dispatch (design §11 worker coalescing).
   *
   *  T4 ZOOM HAZARD (RESOLVED by the render epoch below): coalescing keys on page
   *  INDEX only, with no notion of the scale a dispatch was made at. Once
   *  `setScale` can change the zoom mid-flight, an in-flight bitmap dispatched at
   *  the OLD scale can still pass the on-resolution identity check if the SAME
   *  slot object is re-mounted for page `i` (the pool reuses slot objects, so
   *  `_slots.get(i) === slot && slot.renderedPage === i` can hold for an old
   *  dispatch), and get painted at the WRONG resolution. We fix this with a render
   *  epoch (`_renderEpoch`): each dispatch captures the epoch, and on resolution a
   *  moved epoch ⇒ STALE (close + re-dispatch the live slot). See
   *  `_renderSlotBitmap`. */
  private readonly _bitmapInFlight = new Set<number>();
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
  /** Resolved page-canvas `box-shadow` (design: the recipe drop shadow by
   *  default). Resolved ONCE with `??` — NOT `||` — so `pageShadow: false`
   *  survives as the "no shadow" sentinel (a `||` would treat `false` as absent
   *  and wrongly re-apply the default). Applied by `_applyPageShadow` at EVERY
   *  canvas-creation site (`_acquireSlot` and the double-buffer spare in
   *  `_refreshSlotAtomically`) so a recycled/re-mounted slot and a swapped spare all
   *  carry it. */
  private readonly _pageShadow: string | false;
  private readonly _find = new DocxFindController(
    () => this.pageCount,
    (page) => this._collectPageRuns(page),
  );
  private _findActive = false;
  /** Covers the pre-search progressive wait before DocxFindController.find()
   * can establish its own cancellation generation. */
  private _findRequestGeneration = 0;
  /** ECMA-376 §17.13.5 — current tracked-change view. A LAYOUT axis (deletions
   *  change line breaking and pagination), so it selects which retained layout
   *  variant the viewer reads geometry from; toggle it with
   *  {@link setShowTrackedChanges}. */
  private _showTrackedChanges: boolean;
  /** Canonical epoch milliseconds for the document-global field-date layout
   * axis. Kept beside `_showTrackedChanges` because borrowed documents can
   * change either axis after construction. */
  private _currentDate: Date | number | undefined;
  private _layoutViewGeneration = 0;
  private _layoutViewPublicationGeneration = 0;

  /**
   * Create a Scroll Viewer that borrows an already-loaded document.
   *
   * The document's render mode and active layout view are authoritative. The
   * returned Viewer cannot load another source, and destroying it leaves the
   * caller-owned document open. The initial virtual window is laid out during
   * construction.
   */
  static fromDocument(
    container: HTMLElement,
    document: DocxDocument,
    opts: Omit<DocxScrollViewerOptions, keyof LoadOptions> = {},
  ): Omit<DocxScrollViewer, 'load'> {
    const layoutView = activeDocxLayoutViewOf(document);
    return new DocxScrollViewer(container, {
      ...opts,
      // The caller-owned document has already selected these pagination axes.
      // Seed the viewer from that authoritative state: LoadOptions are omitted
      // from this factory's opts and its local defaults can therefore disagree.
      currentDate: layoutView.currentDate,
      showTrackedChanges: layoutView.showTrackedChanges,
      [borrowedDocumentOption]: document,
    } as InternalDocxScrollViewerOptions);
  }

  constructor(container: HTMLElement, opts: DocxScrollViewerOptions = {}) {
    // A <canvas> is an HTMLElement too, so the type system cannot stop a caller
    // used to the pager API (DocxViewer takes a canvas) from passing one — but
    // canvas children never render, so the viewer would come up silently blank.
    // Fail loudly with the fix instead. (tagName, not instanceof: cross-realm safe.)
    if (container.tagName === 'CANVAS') {
      throw new Error(
        'DocxScrollViewer takes a container element (e.g. a <div>), not a <canvas> — ' +
          'the viewer creates and manages its own canvases. Pass a block container; ' +
          'for the single-page canvas API use DocxViewer.',
      );
    }
    this._container = container;
    this._opts = opts;
    this._errorRouter = new CanvasViewerErrorRouter('DocxScrollViewer', opts.onError);
    this._showTrackedChanges = opts.showTrackedChanges === true;
    this._currentDate = opts.currentDate;
    // `??` (not `||`): a caller's explicit `false` must disable the shadow, not
    // fall through to the default.
    this._pageShadow = opts.pageShadow ?? DEFAULT_PAGE_SHADOW;
    const borrowedDocument = (opts as InternalDocxScrollViewerOptions)[borrowedDocumentOption];
    this._borrowed = borrowedDocument !== undefined;
    if (borrowedDocument) {
      this._documentOwner = new TerminalResourceOwner('DocxScrollViewer', borrowedDocument, false);
      this._mode = resolveCanvasViewerMode('DocxScrollViewer', opts.mode, borrowedDocument);
    } else {
      this._documentOwner = new TerminalResourceOwner('DocxScrollViewer');
      this._mode = resolveCanvasViewerMode('DocxScrollViewer', opts.mode, undefined);
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
    // The "desk" behind/between pages. Undefined ⇒ transparent (container shows
    // through); pages keep their own white canvas regardless.
    if (opts.background) this._scrollHost.style.background = opts.background;
    this._spacer = document.createElement('div');
    this._spacer.style.cssText = 'position:absolute;top:0;left:0;width:1px;height:0;pointer-events:none;';
    this._scrollHost.appendChild(this._spacer);
    this._wrapper.appendChild(this._scrollHost);
    this._container.appendChild(this._wrapper);

    if (this._commentsEnabled()) {
      void loadDocxCommentUiRuntime().then((commentUi) => {
        if (this._destroyed) return;
        this._commentUi = commentUi;
        for (const [page, slot] of this._slots) this._redrawSlotComments(page, slot);
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
        this._activeCommentPage = null;
        for (const [page, slot] of this._slots) this._redrawSlotComments(page, slot);
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
        this.setScale(zoomStepScale(this._scale, e.deltaY, e.deltaMode));
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
      this._bindLayoutDocument(borrowedDocument!);
      // A borrowed engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path re-runs it once width appears).
      this.relayout();
    }
  }

  /**
   * Load a DOCX from URL or ArrayBuffer and render the first window.
   * Unsupported on a Viewer created by {@link fromDocument}; the caller already
   * owns the parsed engine.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
    if (this._borrowed) {
      throw new Error(
        'DocxScrollViewer.load() is unsupported on a Viewer created by fromDocument(); ' +
          'the borrowed document is already loaded.',
      );
    }
    // SC20 atomic swap: a self-loaded viewer OWNS its engine, so a re-load must
    // not orphan the previous one.
    // Retain it locally and free it only after the new engine loads — a FAILED
    // re-load then keeps the current document rendered rather than going blank.
    // (The borrowed path returned above can never reach here, so this only ever
    // frees an engine we created.)
    let elementInvalidated = false;
    try {
      const doc = await this._documentOwner.replace(() => DocxDocument.load(source, {
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
        tiff: this._opts.tiff,
        mode: this._mode,
        // The variant the viewer will render. Without these, load builds the
        // final view while every render asks for the markup view, and the first
        // paint pays a full synchronous repagination.
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        ...(this._currentDate === undefined
          ? {}
          : { currentDate: this._currentDate }),
        ...(this._opts.progressiveLayout ? { progressiveLayout: true } : {}),
        ...(this._opts.sliceLayout ? { sliceLayout: true } : {}),
        onLayoutProgress: this._opts.onLayoutProgress,
        onLayoutPartial: this._opts.onLayoutPartial,
        onLayoutComplete: this._opts.onLayoutComplete,
      }), (ownedDocument) => {
        this._invalidateElementContext(false);
        elementInvalidated = true;
        this._findRequestGeneration++;
        this._find.invalidate();
        this._findActive = false;
        this._activeCommentId = null;
        this._activeCommentPage = null;
        this._resetCommentNavigation();
        this._unbindLayoutDocument();
        if (ownedDocument) {
          // Recycle before the old worker is terminated. Every captured slot
          // dispatcher then becomes stale before its expected rejection lands.
          for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
          this._lastTopIndex = -1;
          this._lastReportedTotal = -1;
          this._lastReportedLayoutComplete = null;
        }
      });
      if (!doc) return;
      if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
      this._bindLayoutDocument(doc);
      this._find.invalidate();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentPage = null;
      this._resetCommentNavigation();
      // Lay out + mount the first window now that the engine exists (mirrors the
      // borrowed-engine path in the constructor). relayout() is idempotent and
      // defers under a zero-width container — `_onResize` re-runs it once width
      // appears.
      const initialRenders: Promise<void>[] = [];
      this._relayout(initialRenders);
      await Promise.all(initialRenders);
    } catch (err) {
      // Superseded loads own no error reporting — the winning load (or destroy())
      // is the outcome the caller awaits; swallow this stale rejection.
      if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
      throw err instanceof Error ? err : new Error(String(err));
    }
    if (elementInvalidated && !this._destroyed) this._emitSelectionContextChange();
  }

  get pageCount(): number {
    return this._doc?.pageCount ?? 0;
  }

  /**
   * Whether every page has been laid out.
   *
   * False while progressive pagination is pending and remains false if that
   * background work fails. While false, {@link pageCount} is provisional;
   * {@link waitUntilLayoutComplete} distinguishes pending work from failure.
   */
  get layoutComplete(): boolean {
    return this._doc?.layoutComplete ?? true;
  }

  /**
   * Resolve once the whole document is laid out.
   *
   * Await this before anything that must see every page — a total page count,
   * printing, export. {@link findText} does so internally. Resolves immediately
   * unless progressive layout actually deferred work, and rejects if that
   * background pagination fails.
   */
  async waitUntilLayoutComplete(): Promise<void> {
    // Optional-called because an INJECTED engine (fromDocument) may predate this
    // method; a document that cannot defer layout is already complete.
    await this._errorRouter.ownBackgroundLifecycle(async () => {
      await this._doc?.waitUntilLayoutComplete?.();
    });
  }

  private _bindLayoutDocument(doc: DocxDocument): void {
    this._unbindLayoutDocument();
    this._layoutViewPublicationGeneration = 0;
    this._presentedPageCount = doc.pageCount;
    const unsubscribeView = subscribeDocxLayoutView(
      doc,
      (publication) => this._onLayoutViewPublication(doc, publication),
      (error) => this._reportRenderError(error),
    );
    let initial = true;
    const unsubscribeLayout = subscribeDocxLayout(
      doc,
      () => ({
        pageCount: doc.pageCount,
        exact: doc.layoutComplete,
        complete: doc.layoutComplete,
      }),
      (publication) => {
        if (initial) {
          initial = false;
          return;
        }
        this._onLayoutPublication(doc, publication);
      },
      (error) => this._reportRenderError(error),
    );
    this._layoutUnsubscribe = () => {
      unsubscribeLayout();
      unsubscribeView();
    };
  }

  private _unbindLayoutDocument(): void {
    this._layoutUnsubscribe?.();
    this._layoutUnsubscribe = null;
    this._presentedPageCount = 0;
  }

  private _onLayoutPublication(doc: DocxDocument, publication: DocxLayoutPublication): void {
    if (this._destroyed || doc !== this._doc) return;
    if (publication.error !== undefined) {
      this._errorRouter.reportBackground(
        publication.error,
        this._opts.onLayoutComplete !== undefined,
      );
      return;
    }
    this._find.invalidate();
    this._refreshCommentSurface();
    // Publish every paintable prefix immediately. Atomic canvas replacement
    // below prevents blank frames while an existing page refreshes, so there is
    // no need to hide scroll growth until the reader reaches the old tail.
    this._applyLayoutPublication(publication);
  }

  private _onLayoutViewPublication(
    doc: DocxDocument,
    publication: DocxLayoutViewPublication,
  ): void {
    if (
      this._destroyed
      || doc !== this._doc
      || publication.generation <= this._layoutViewPublicationGeneration
    ) return;
    this._layoutViewPublicationGeneration = publication.generation;
    if (publication.requester === this) return;
    this._layoutViewGeneration++;
    this._showTrackedChanges = publication.view.showTrackedChanges;
    this._currentDate = publication.view.currentDate;
    this._find.invalidate();
    this._applyLayoutPublication({
      pageCount: doc.pageCount,
      exact: true,
      complete: doc.layoutComplete,
    });
  }

  /** Refresh only the empty review layers created with each comment-enabled
   * slot. Progressive anchor discovery never detaches a painted page canvas. */
  private _refreshCommentSurface(): void {
    if (!this._commentsEnabled() || this._slots.size === 0) return;
    this._syncSpacerWidth();
    for (const [page, slot] of this._slots) this._redrawSlotComments(page, slot);
  }

  /** Admit one layout publication to scroll geometry and refresh every mounted
   * page atomically. The old canvas remains visible while main mode paints an
   * off-DOM replacement; worker mode already commits ImageBitmap pixels in one
   * synchronous transfer. */
  private _applyLayoutPublication(publication: DocxLayoutPublication): void {
    const mounted = [...this._slots];
    this._presentedPageCount = publication.pageCount;
    // Supersede spares/bitmaps dispatched for an earlier publication. The
    // newest layout is the only one allowed to swap into a live slot.
    this._renderEpoch++;
    this.relayout();
    for (const [page, slot] of mounted) {
      if (page >= this._presentedPageCount || this._slots.get(page) !== slot) continue;
      this._refreshSlotAtomically(page, slot);
    }
  }

  /** CSS px width of page `i` at the current scale. */
  private _pageWidthPx(i: number): number {
    return this._doc!.pageSize(i).widthPt * PT_TO_PX * this._scale;
  }

  /** CSS px height of page `i` at the current scale. */
  private _pageHeightPx(i: number): number {
    return this._doc!.pageSize(i).heightPt * PT_TO_PX * this._scale;
  }

  /** The fit width (px), deferring when the container is unlaid-out. An EXPLICIT
   *  `opts.width` is the page's CSS-width contract and is returned UNCHANGED (the
   *  gutters still apply around placement, not to the width). The container-derived
   *  default instead targets `containerWidth − padL − padR` so a page sits INSIDE
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
    // Fit the authored page itself. Review cards are an adjacent horizontal
    // surface, so their late discovery never changes page scale or vertical
    // scroll extent.
    return available;
  }

  private _hasCommentMargin(): boolean {
    return this._hasDisplayableComments() && this._commentsOptions()?.cards !== false;
  }

  private _hasDisplayableComments(): boolean {
    if (!this._commentsEnabled()) return false;
    const doc = this._doc;
    if (!doc) return false;
    const anchorRanges = doc.commentAnchorRanges();
    if (this._commentAnchorRangesForMargin !== anchorRanges) {
      this._commentAnchorRangesForMargin = anchorRanges;
      this._commentAnchorIds = new Set(
        anchorRanges.map((anchor) => anchor.commentId),
      );
    }
    if (this._commentAnchorIds.size === 0) return false;
    const includeResolved = this._commentsOptions()?.includeResolved === true;
    return doc.comments.some((comment) =>
      this._commentAnchorIds.has(comment.id) &&
      comment.parentId === undefined &&
      (includeResolved || comment.resolved !== true));
  }

  private _commentMarginExtent(): number {
    return this._hasCommentMargin()
      ? (COMMENT_MARGIN_GAP_PX + READ_ONLY_COMMENT_MARGIN_WIDTH_PX) * this._commentZoom()
      : 0;
  }

  /** Comment chrome uses the same absolute zoom as the rendered document. */
  private _commentZoom(): number {
    return this._scaleEstablished ? this._scale : 1;
  }

  private _commentsEnabled(): boolean {
    return this._opts.comments === true || typeof this._opts.comments === 'object';
  }

  private _commentsOptions(): DocxCommentsOptions | undefined {
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
    // Empty absolutely positioned review cards otherwise enlarge scrollWidth
    // even though no displayable comment contributes a declared margin extent.
    margin.style.display = this._hasCommentMargin() ? '' : 'none';
    const zoom = this._commentZoom();
    const offset = `calc(100% + ${COMMENT_MARGIN_GAP_PX * zoom}px)`;
    margin.style.left = this._commentSide() === 'right' ? offset : '';
    margin.style.right = this._commentSide() === 'left' ? offset : '';
    margin.style.width = `${READ_ONLY_COMMENT_MARGIN_WIDTH_PX * zoom}px`;
    margin.style.fontSize = `${COMMENT_MARGIN_FONT_SIZE_PX}px`;
    margin.dataset.ooxmlCommentZoom = String(zoom);
  }

  /** Widest authored page width. DOCX sections may differ by a fraction of a
   * point or switch orientation, and fit-width must cover the same extent as the
   * horizontal spacer. */
  private _widestPageWidthPt(): number {
    if (!this._doc) return 0;
    let widthPt = 0;
    const pageCount = this._presentedPageCount || this._doc.pageCount;
    for (let i = 0; i < pageCount; i++) {
      const pageWidthPt = this._doc.pageSize(i).widthPt;
      if (pageWidthPt > widthPt) widthPt = pageWidthPt;
    }
    return widthPt;
  }

  /** Base scale: widest page's width fit to the fit-width. Returns 0 when the
   *  container has no width yet (deferral). */
  private _baseScale(): number {
    if (!this._doc || this._doc.pageCount === 0) return 0;
    const w = this._fitWidthPx();
    if (w <= 0) return 0;
    const widestWpt = this._widestPageWidthPt();
    if (widestWpt <= 0) return 0;
    return w / (widestWpt * PT_TO_PX);
  }

  /**
   * Recompute per-page heights + the spacer and re-mount the visible window.
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
    if (!this._doc) return;
    // Non-progressive/authoritative engines have no pending prefix boundary.
    // Keep the historical relayout escape hatch able to observe an injected
    // engine whose final page count changed between calls.
    if (this._doc.layoutComplete !== false) {
      this._presentedPageCount = this._doc.pageCount;
    }
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
    } else {
      // Progressive pagination or a layout-view switch can reveal a page wider
      // than the one(s) used for the previous base. Re-fit even when the
      // container itself did not resize, preserving the user's zoom multiplier.
      const base = this._baseScale();
      if (base > 0 && base !== this._prevBase) {
        const mult = this._prevBase > 0 ? this._scale / this._prevBase : 1;
        this._prevBase = base;
        this.setScale(base * mult);
      }
    }
    this._recomputeHeights();
    this._syncSpacer();
    this._mountVisible(initialRenders);
    // A progressive publication may grow page count without changing the
    // current top slot. Republish logical/decoration state even when the page's
    // pixels and collected runs are already exact, so logical anchors and
    // built-in card/connector geometry cannot remain latched to an earlier
    // prefix.
    for (const [page, slot] of this._slots) {
      // A newly mounted slot is stamped with `renderedPage` before its async
      // paint completes.  Its comment runs are not authoritative until that
      // paint commits and records `renderedScale`; the render completion path
      // publishes them.  Relayout only republishes already-painted slots.
      if (slot.renderedPage === page && slot.renderedScale >= 0) {
        this._redrawSlotComments(page, slot);
      }
    }
  }

  private _recomputeHeights(): void {
    const n = Math.min(this._presentedPageCount, this._doc!.pageCount);
    const h = new Array<number>(n);
    for (let i = 0; i < n; i++) h[i] = this._pageHeightPx(i);
    this._heights = h;
    this._scrollGeometry = createVirtualScrollGeometry(h, this._gap(), this._pad());
  }

  private _gap(): number {
    return this._opts.gap ?? 16;
  }

  private _overscan(): number {
    return this._opts.overscan ?? 1;
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
   *  `_positionSlot` (the flush-left floor), and by `_syncSpacer` (the spacer
   *  width). Resolved here (not stored) to mirror `_gap()`/`_pad()`. */
  private _padH(): { left: number; right: number } {
    const gap = this._gap();
    return { left: this._opts.paddingLeft ?? gap, right: this._opts.paddingRight ?? gap };
  }

  /** Index of the page whose slot spans content-offset `y` (largest `i` with
   *  `offsets[i] <= y`), for the pointer-anchored zoom re-anchor. Mirrors the
   *  `topIndex` search `computeVisibleRange` runs for the scrollTop, but for an
   *  ARBITRARY content-y (the pointer, not the viewport top). Clamped into
   *  `[0, n-1]`; a `y` below the first page (inside the leading pad) yields 0. */
  private _pageIndexAtOffset(r: VisibleRange, y: number): number {
    const { offsets } = r;
    let lo = 0;
    let hi = offsets.length - 1;
    let idx = 0;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (offsets[mid] <= y) {
        idx = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return idx;
  }

  private _range(): VisibleRange {
    return computeVisibleWindow(
      this._scrollGeometry,
      this._scrollHost.scrollTop,
      this._scrollHost.clientHeight,
      this._overscan(),
    );
  }

  private _syncSpacer(): void {
    const r = this._range();
    this._lastRange = r;
    this._spacer.style.height = `${r.totalHeight}px`;
    this._syncSpacerWidth();
  }

  /** Horizontal scroll extent: the widest page (docx pages can differ in width)
   *  plus both gutters. A spacer NARROWER than the container never creates a
   *  scrollbar (scrollWidth = max(clientWidth, content)), so it is always safe to
   *  set — it only matters when a zoomed-in page grows past the viewport, where it
   *  gives the gutters something to scroll to on either side. Max over per-page
   *  widths so the extent covers the widest page in the document. Called from
   *  `_syncSpacer` and after every scale change (zoom / resize re-fit) so the
   *  extent tracks the current page px width. */
  private _syncSpacerWidth(): void {
    const { left, right } = this._padH();
    let maxW = 0;
    for (let i = 0; i < this._heights.length; i++) {
      const w = this._pageWidthPx(i);
      if (w > maxW) maxW = w;
    }
    const marginExtent = this._commentMarginExtent();
    const next = this._commentSide() === 'left' ? marginExtent : 0;
    const delta = next - this._reviewOriginPx;
    const targetScrollLeft = Math.max(0, this._scrollHost.scrollLeft + delta);
    // Establish the new native scroll range before applying compensation.
    // Browsers clamp scrollLeft to the current range at assignment time.
    this._spacer.style.width = `${maxW + marginExtent + left + right}px`;
    if (delta === 0) return;
    this._reviewOriginPx = next;
    (this._scrollHost.style as CSSStyleDeclaration & Record<string, string>)[
      '--ooxml-review-origin-x'
    ] = `${next}px`;
    this._scrollHost.scrollLeft = targetScrollLeft;
  }

  private _onScroll(): void {
    if (!this._doc || !this._scaleEstablished) return;
    this._mountVisible(undefined, false);
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(
    initialRenders?: Promise<void>[],
    repositionExisting = true,
  ): void {
    if (!this._doc || this._doc.pageCount === 0) return;
    const r = this._range();
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
        const render = this._renderSlot(i, slot, initialRenders === undefined);
        if (initialRenders && render) initialRenders.push(render);
      } else if (repositionExisting) {
        // Re-position (offsets shift after a spacer/height change).
        const slot = this._slots.get(i)!;
        this._positionSlot(slot, i, r);
        // `_renderSlot` returns null when the slot already holds page `i`, so
        // ordinary scrolling and resizing never re-render a mounted page.
        // Layout replacement and zoom settle use `_refreshSlotAtomically`
        // separately so the live canvas remains painted until its replacement
        // is ready.
        const render = this._renderSlot(i, slot, initialRenders === undefined);
        if (initialRenders && render) initialRenders.push(render);
      }
    }
    this._emitVisiblePageChange(r);
  }

  /**
   * Fire `onVisiblePageChange`, but only on an actual change.
   *
   * The latch is the (topIndex, total, complete) tuple. Watching the index alone
   * was enough while a document's page count was fixed at load; under
   * progressive layout the count can grow while the user sits at the top, and
   * the authoritative publication can retain the same count while changing
   * `complete` to true. Every emit path funnels through here so unchanged state
   * still never double-fires.
   */
  private _emitVisiblePageChange(r: VisibleRange): void {
    if (!this._doc) return;
    const total = this._doc.pageCount;
    const complete = this.layoutComplete;
    if (
      r.topIndex === this._lastTopIndex &&
      total === this._lastReportedTotal &&
      complete === this._lastReportedLayoutComplete
    ) return;
    this._lastTopIndex = r.topIndex;
    this._lastReportedTotal = total;
    this._lastReportedLayoutComplete = complete;
    this._opts.onVisiblePageChange?.(r.topIndex, total, complete);
  }

  /** Apply the resolved page-canvas shadow (design: recipe drop shadow by
   *  default, `false` ⇒ none). Single source so `_acquireSlot` and the
   *  double-buffer spare in `_refreshSlotAtomically` stay in lock-step — a spare
   *  that missed this would lose the shadow on the settle swap. `box-shadow`
   *  never affects layout, so this is safe to (re)set on a live/pooled canvas
   *  without shifting any offset. */
  private _applyPageShadow(canvas: HTMLCanvasElement): void {
    if (this._pageShadow !== false) canvas.style.boxShadow = this._pageShadow;
  }

  private _acquireSlot(): PageSlot {
    const reused = this._free.pop();
    if (reused) {
      // _recycleSlot already reset renderedPage to -1 before pooling this slot.
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
    let commentTintLayer: HTMLDivElement | null = null;
    let commentMargin: HTMLDivElement | null = null;
    let commentDecorationLayer: HTMLDivElement | null = null;
    if (this._commentsEnabled()) {
      commentTintLayer = document.createElement('div');
      commentTintLayer.style.cssText =
        'position:absolute;inset:0;overflow:hidden;pointer-events:none;';
      wrapper.appendChild(commentTintLayer);
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
    const slot: PageSlot = {
      wrapper,
      canvas,
      textLayer,
      highlightLayer,
      elementLayer,
      commentTintLayer,
      commentMargin,
      commentDecorationLayer,
      commentRuns: Object.freeze([]),
      commentGeometry: null,
      renderedPage: -1,
      renderedScale: -1,
      dispatcher: new StaticCanvasRenderDispatcher(canvas, this._mode === 'worker'),
    };
    return slot;
  }

  private _recycleSlot(idx: number, slot: PageSlot): void {
    this._slots.delete(idx);
    slot.dispatcher.destroy();
    if (!this._destroyed) {
      slot.dispatcher = new StaticCanvasRenderDispatcher(slot.canvas, this._mode === 'worker');
    }
    // Clear the per-slot text overlay so a slot sitting in the free pool holds no
    // stale spans. buildDocxTextLayer also clears on its next build, but an
    // unrendered pooled slot never gets that build, and the detached spans would
    // otherwise linger; drop them here.
    if (slot.textLayer) {
      slot.textLayer.innerHTML = '';
      // Drop any preview transform so a pooled slot re-used for another page does
      // not inherit a stale scale() before its overlay is rebuilt.
      this._clearTextLayerPreview(slot.textLayer);
    }
    slot.highlightLayer.innerHTML = '';
    slot.highlightLayer.style.transform = '';
    slot.highlightLayer.style.transformOrigin = '';
    if (slot.commentTintLayer) {
      slot.commentTintLayer.replaceChildren();
      slot.commentTintLayer.style.transform = '';
      slot.commentTintLayer.style.transformOrigin = '';
      slot.commentTintLayer.style.visibility = '';
    }
    if (slot.commentMargin) {
      this._commentUi?.disposeReadOnlyCommentMargin(slot.commentMargin);
      if (!this._commentUi) slot.commentMargin.replaceChildren();
      slot.commentMargin.style.visibility = '';
    }
    if (slot.commentDecorationLayer) {
      this._commentUi?.disposeReadOnlyCommentDecoration(slot.commentDecorationLayer);
      if (!this._commentUi) slot.commentDecorationLayer.replaceChildren();
    }
    slot.commentRuns = Object.freeze([]);
    slot.commentGeometry = null;
    renderCanvasElementOutline(slot.elementLayer, null);
    slot.renderedPage = -1;
    slot.renderedScale = -1;
    slot.wrapper.remove();
    this._free.push(slot);
  }

  private _positionSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    slot.wrapper.style.top = `${r.offsets[i]}px`;
    const wpx = this._pageWidthPx(i);
    const hpx = this._pageHeightPx(i);
    slot.wrapper.style.width = `${wpx}px`;
    slot.wrapper.style.height = `${hpx}px`;
    this._syncCommentMarginGeometry(slot.commentMargin);
    if (slot.commentDecorationLayer) {
      const marginExtent = this._commentMarginExtent();
      slot.commentDecorationLayer.style.left = this._commentSide() === 'left'
        ? `${-marginExtent}px`
        : '0px';
      slot.commentDecorationLayer.style.width = `${wpx + marginExtent}px`;
      slot.commentDecorationLayer.style.height = `${hpx}px`;
    }
    this._redrawElementOutlineForSlot(i, slot);
    // Horizontal placement (replaces the old CSS `left:0;right:0;margin:0 auto`
    // auto-centering, which cannot honour a left gutter). Centre the page in the
    // scroll viewport, but never let its left edge cross the left gutter: when the
    // page is narrower than the viewport it is centred (`(cw − pw)/2 > padL`); once
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
   * Render page `i` into `slot`. Routes strictly on the constructor-resolved
   * `_mode` (design §11 — no probing, no silent mis-pathing): `main` ⇒ paint the
   * slot's canvas directly via `renderPage`; `worker` ⇒ transfer an ImageBitmap
   * from `renderPageToBitmap`.
   *
   * Slot-identity guard: a slot recycled to a DIFFERENT page while a previous
   * render is in flight must not repaint the stale page. `slot.renderedPage`
   * tracks the page this slot is committed to; we stamp it up-front and bail on
   * resolution if it changed (the engine's own token guard is per-canvas; this is
   * the viewer's per-slot page-identity check).
   *
   * Render epoch (main path): pixel staleness after a mid-flight `setScale` is
   * already handled by the engine's per-canvas token (the newer renderPage on the
   * same canvas wins) — `setScale` recycles + re-mounts, and the re-mount always
   * re-dispatches `renderPage` (renderedPage reset to -1), so a fresh render is
   * always issued. But the viewer-side side effects of a STALE resolution — the
   * text-layer build (its run geometry is at the OLD scale) and the renderedPage
   * bookkeeping — must NOT run, or a superseded render would rebuild the overlay
   * with stale x/y/w/h (the pool reuses slot objects, so the identity check alone
   * can pass for an old-epoch resolution). We gate them on the captured epoch.
   */
  private _renderSlot(i: number, slot: PageSlot, reportErrors = true): Promise<void> | null {
    if (!this._doc) return null;
    // Slot-identity guard: this slot is already rendering / has rendered page i.
    if (slot.renderedPage === i) return null;
    slot.renderedPage = i;

    const dpr = this._dpr();
    const widthPx = this._pageWidthPx(i);
    const epoch = this._renderEpoch;
    const scale = this._scale;
    const dispatcher = slot.dispatcher;
    const generation = dispatcher.begin();

    if (this._mode === 'worker') {
      return this._renderSlotBitmap(
        i,
        slot,
        widthPx,
        dpr,
        scale,
        dispatcher,
        generation,
        reportErrors,
      );
    }

    // Main mode: render straight onto the slot's canvas.
    const runs: DocxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive || !!slot.commentTintLayer;
    const onTextRun = wantRuns ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    let render: Promise<void>;
    try {
      render = renderDocxFocusedPage(this._doc, slot.canvas, i, 'main', {
        width: widthPx, // this page's own px width → uniform px-per-pt scale (§7)
        dpr,
        imageResources: this._opts.imageResources,
        defaultTextColor: this._opts.defaultTextColor,
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        onTextRun,
      });
    } catch (error) {
      if (reportErrors) {
        this._reportRenderError(error);
        return Promise.resolve();
      }
      return Promise.reject(error);
    }
    return render
      .then(() => {
        // Stale if the epoch moved (a setScale rescaled mid-flight — the run
        // geometry is at the old scale), or a recycle re-purposed this slot for a
        // different page / freed it. Either way: skip the (stale) overlay build.
        // The engine's per-canvas token already discards the superseded pixels.
        if (
          !dispatcher.isCurrent(generation) ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedPage !== i
        ) return;
        // This fresh render defines the scale the on-screen bitmap now lives at,
        // so a subsequent zoom preview stretches from HERE.
        slot.renderedScale = scale;
        if (wantOverlay && slot.textLayer) {
          const { width, height } = this._canvasCssPx(slot.canvas);
          buildDocxTextLayer(
            slot.textLayer,
            runs,
            width,
            height,
            this._hyperlinkHandler(),
            (font) => this._measureForFont(font),
            i,
          );
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitCommentRuns(i, slot, runs);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        const isCurrent =
          dispatcher.isCurrent(generation) &&
          epoch === this._renderEpoch &&
          this._slots.get(i) === slot &&
          slot.renderedPage === i;
        if (!isCurrent) return;
        if (reportErrors) this._reportRenderError(err);
        else throw err;
      });
  }

  /**
   * IX1/IX-nav — the click handler passed to the text-layer overlay. When the
   * caller supplied `onHyperlinkClick`, it fully owns the behaviour (the default
   * is suppressed). Otherwise the built-in default is: an external link opens in
   * a new tab through core `openExternalHyperlink` (URL sanitised against the
   * safe scheme allowlist, `noopener,noreferrer`); an internal `<w:anchor>` link
   * resolves its bookmark name to its destination page via
   * {@link DocxDocument.getBookmarkPage} (ECMA-376 §17.16.23) and scrolls there
   * with {@link scrollToPage}. An anchor naming no known bookmark is a safe no-op
   * rather than a scroll to a guessed page.
   *
   * IX1 — returns `undefined` when `enableHyperlinks` is `false`, the single gate
   * that disables hyperlink interactivity: {@link buildDocxTextLayer} treats a
   * missing handler as "render link runs like plain runs", so no hit region,
   * cursor, tooltip, listener, or navigation is wired (a custom
   * `onHyperlinkClick` is suppressed too).
   */
  private _hyperlinkHandler(): ((target: HyperlinkTarget) => void) | undefined {
    if (this._opts.enableHyperlinks === false) return undefined;
    const custom = this._opts.onHyperlinkClick;
    if (custom) return custom;
    return (target: HyperlinkTarget): void => {
      if (target.kind === 'external') {
        openExternalHyperlink(target.url);
        return;
      }
      const doc = this._doc;
      if (!doc) return;
      const generation = ++this._internalHyperlinkGeneration;
      void this._navigateInternalHyperlink(doc, target.ref, generation)
        .catch((error) => this._reportRenderError(error));
    };
  }

  private async _navigateInternalHyperlink(
    doc: DocxDocument,
    ref: string,
    generation: number,
  ): Promise<void> {
    if (!doc.layoutComplete) await doc.waitUntilLayoutComplete();
    if (this._destroyed || this._doc !== doc || generation !== this._internalHyperlinkGeneration) {
      return;
    }
    const page = doc.getBookmarkPage(ref);
    if (page !== undefined) this.scrollToPage(page);
  }

  /** A width-measurer primed with a run's `font` — used ONLY to clamp a §17.3.2.10
   *  縦中横 selection span to its drawn one-em cell (#836). Mirrors DocxViewer's
   *  `_measureForFont`. Returns a length-based fallback when canvas metrics are
   *  unavailable so the caller still gets a callable (the overlay then sees scale
   *  1 and leaves the span un-clamped). */
  private _measureForFont(font: string): (s: string) => number {
    if (this._measureCtx === undefined) {
      const c = document.createElement('canvas');
      this._measureCtx = c.getContext('2d');
    }
    const ctx = this._measureCtx;
    if (!ctx || typeof ctx.measureText !== 'function') return (s) => s.length;
    ctx.font = font;
    return (s) => ctx.measureText(s).width;
  }

  /** A canvas's intended CSS box in px (the % denominators the overlay builders
   *  expect). Reads the inline `style.width`/`height` set by the render path,
   *  falling back to the backing-store size when unset; tolerates the `px` suffix. */
  private _canvasCssPx(canvas: HTMLCanvasElement): { width: number; height: number } {
    return {
      width: parseFloat(canvas.style.width) || canvas.width,
      height: parseFloat(canvas.style.height) || canvas.height,
    };
  }

  /** Route an async render failure to `onError`, or `console.error` when none is
   *  set (so failures are never fully silent), and never after teardown. */
  private _reportRenderError(err: unknown): void {
    this._errorRouter.report(err);
  }

  /**
   * Worker-mode slot render: dispatch `renderPageToBitmap`, transfer the result
   * via a per-slot `bitmaprenderer` context, and manage the ImageBitmap lifecycle.
   *
   * Coalescing / drop-stale (design §11):
   *  - Skip if page `i` is already in flight (a scroll storm won't double-dispatch).
   *  - Skip if page `i` already left the mounted window before dispatch.
   *  - On resolution, if `slot` is no longer THIS page's live slot (it recycled to
   *    another page, or page `i` re-mounted onto a DIFFERENT slot while this render
   *    was in flight), close the orphan bitmap and skip the paint. In that
   *    re-mount case a live slot for `i` still awaits a render, so once we clear
   *    the in-flight guard we re-dispatch it — a page that recycled and re-mounted
   *    mid-flight must never stay blank.
   *  - RENDER EPOCH: the dispatch captures `this._renderEpoch`. `setScale` bumps
   *    the epoch, so a resolution whose captured epoch ≠ the live epoch is STALE
   *    even when the SAME slot object is still mounted for page `i` (the pool
   *    reuses slot objects, so the identity check alone can't catch a zoom that
   *    happened mid-flight). A moved epoch ⇒ close the orphan + re-dispatch the
   *    live slot at the new scale, never paint the old-scale bitmap.
   */
  private async _renderSlotBitmap(
    i: number,
    slot: PageSlot,
    widthPx: number,
    dpr: number,
    scale: number,
    dispatcher = slot.dispatcher,
    generation = dispatcher.begin(),
    reportErrors = true,
  ): Promise<void> {
    if (this._bitmapInFlight.has(i)) return; // coalesce: already dispatched
    // Drop-stale before dispatch: if this page already scrolled out of the
    // mounted window, don't dispatch at all.
    if (this._slots.get(i) !== slot) return;
    const epoch = this._renderEpoch;
    // Logical CSS geometry is independent from the worker bitmap's backing
    // dimensions. The renderer may reduce the bitmap to stay inside the browser
    // canvas area limit; the page must still occupy its requested layout box.
    const heightPx = this._pageHeightPx(i);
    this._bitmapInFlight.add(i);
    // Whether this invocation actually painted its slot. When it did NOT (stale
    // epoch or moved identity), the `finally` may need to re-dispatch a live slot.
    let painted = false;
    // IX6 — harvest the page's run geometry alongside the bitmap so the
    // worker-mode selection overlay is built from the SAME data main mode uses.
    // The runs ride back beside the bitmap (one round-trip), collected only when
    // an overlay is actually wanted.
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive || !!slot.commentTintLayer;
    const runs: DocxTextRunInfo[] = [];
    try {
      const bmp = await renderDocxFocusedPage(this._doc!, slot.canvas, i, 'worker', {
        width: widthPx,
        dpr,
        imageResources: this._opts.imageResources,
        defaultTextColor: this._opts.defaultTextColor,
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        onTextRun: wantRuns ? (r) => runs.push(r) : undefined,
      });
      // Stale if EITHER (a) the epoch moved (a setScale rescaled mid-flight, so
      // this bitmap is at a superseded resolution — this catches the case where
      // the SAME slot object is re-mounted for page `i`, which the identity check
      // below cannot), or (b) the slot recycled to a different page / page `i`
      // re-mounted onto a DIFFERENT slot. Either way: close + skip the paint.
      if (
        !dispatcher.isCurrent(generation) ||
        epoch !== this._renderEpoch ||
        this._slots.get(i) !== slot ||
        slot.renderedPage !== i
      ) {
        bmp.close();
        return;
      }
      if (!dispatcher.commitBitmap(generation, bmp, {
        cssWidth: widthPx,
        cssHeight: heightPx,
      })) return;
      // This bitmap now defines the scale the on-screen canvas lives at, so a
      // later zoom preview stretches from HERE (design §7 renderedScale).
      slot.renderedScale = scale;
      // IX6 — build the selection overlay from the runs the worker just shipped.
      // Reached only past the staleness gate, so the geometry matches THIS paint
      // (same epoch guard the main-mode path relies on for stale-scale safety).
      // Clear any preview transform first: a settle re-render lands at the
      // current scale, so the overlay's `scale()` from `_previewSlot` is stale
      // and the rebuilt spans already sit at the crisp geometry (mirrors the
      // main-mode `_refreshSlotAtomically` clear).
      if (slot.textLayer) {
        this._clearTextLayerPreview(slot.textLayer);
        if (wantOverlay) {
          const { width, height } = this._canvasCssPx(slot.canvas);
          buildDocxTextLayer(
            slot.textLayer,
            runs,
            width,
            height,
            this._hyperlinkHandler(),
            (font) => this._measureForFont(font),
            i,
          );
        }
      }
      if (wantRuns) this._refreshFindRuns(i, runs);
      this._commitCommentRuns(i, slot, runs);
      this._redrawSlotHighlights(i, slot);
      painted = true;
    } catch (err) {
      const isCurrent =
        dispatcher.isCurrent(generation) &&
        epoch === this._renderEpoch &&
        this._slots.get(i) === slot &&
        slot.renderedPage === i;
      if (isCurrent) {
        if (reportErrors) this._reportRenderError(err);
        else throw err;
      }
    } finally {
      this._bitmapInFlight.delete(i);
      // Re-dispatch ONLY when this invocation went stale — a LIVE slot for page
      // `i` still awaits a correct render and the reason we didn't paint was
      // staleness, not a render failure. The two staleness cases:
      //  - IDENTITY MOVED (`live !== slot`): page `i` re-mounted onto a DIFFERENT
      //    slot while we ran (the re-mount's own dispatch was coalesced away by
      //    the in-flight guard), so the live slot has no render in flight.
      //  - EPOCH MOVED (`epoch !== this._renderEpoch`): a `setScale` bumped the
      //    epoch mid-flight, so this bitmap was at a superseded scale. The live
      //    slot may be the SAME object reused from the pool, which the identity
      //    test alone would miss — the epoch test catches the same-slot case.
      // NO RETRY ON PLAIN REJECTION: when the slot is still live at the same epoch
      // and we simply failed (`renderPageToBitmap` rejected or the transfer threw),
      // `!painted` holds but BOTH staleness tests are false, so we do NOT
      // re-dispatch. Retrying a plain failure would loop unbounded (reject →
      // re-dispatch → reject → …); the onError contract is that "a failed page is
      // left blank" (see DocxScrollViewerOptions.onError), so we leave it blank.
      // Bounded epoch-then-reject: an epoch-moved re-dispatch captures the NEW
      // epoch, so if that fresh render then rejects at the still-current epoch,
      // both tests are false and it stops — no unbounded retry.
      const live = this._slots.get(i);
      if (
        !painted &&
        live &&
        (live !== slot || epoch !== this._renderEpoch || !dispatcher.isCurrent(generation)) &&
        !this._bitmapInFlight.has(i) &&
        !this._destroyed
      ) {
        // live.renderedPage === i already (set by _renderSlot on mount); the fresh
        // dispatch runs at the CURRENT epoch/scale via _pageWidthPx(i). Keep the
        // replacement in this Promise chain: load() awaits the render originally
        // mounted for the opening window, and must therefore follow superseding
        // epochs until the render that can actually commit has finished. Callers
        // that intentionally fire-and-forget this method still do so at their
        // outer call site.
        const nextDispatcher = live.dispatcher;
        await this._renderSlotBitmap(
          i,
          live,
          this._pageWidthPx(i),
          this._dpr(),
          this._scale,
          nextDispatcher,
          nextDispatcher.begin(),
          reportErrors,
        );
      }
    }
  }

  /**
   * Set the absolute px-per-pt zoom scale, clamped inline to
   * `[zoomMin ?? 0.1, zoomMax ?? 4]` (absolute bounds, XlsxViewer convention — NOT
   * multiples of the base fit; design §3 keeps the clamp in the viewer, not core),
   * then re-anchor VERTICALLY so the page currently under the viewport top stays
   * fixed. A no-op when the clamped scale is unchanged. Called BEFORE the doc is
   * loaded / the base fit is established, the clamped factor is LATCHED (IX9 F1,
   * family-unified with the single-canvas viewers) and applied by `relayout()`
   * once the layout establishes — `onScaleChange` fires then.
   *
   * FLICKER-FREE (design §7): this does NOT re-render the visible pages inline.
   * It shows an immediate CSS preview (stretch the existing bitmaps, scale the
   * overlays) and DEBOUNCES a full-resolution settle re-render for ZOOM_SETTLE_MS,
   * so a wheel/pinch burst never blanks a page and coalesces into one crisp render.
   *
   * Re-anchor (written from scratch — XlsxViewer only re-anchors horizontally):
   * capture `top = topIndex` and the intra-page fraction `intraFrac` from the
   * CURRENT range BEFORE rescale; after recomputing heights at the new scale,
   * `newScrollTop = offsets'[top] + intraFrac × heights'[top]`, clamped to
   * `[0, totalHeight' − viewportHeight]`. Because a page's height scales linearly
   * with `_scale`, the same fractional position maps exactly to the new geometry.
   *
   * When the width-fit base is below `zoomMin` (for example a poster-sized page
   * in a narrow container), that fit becomes the effective floor. The opening
   * view already permits the smaller scale, so keeping it reachable avoids a
   * one-way zoom ratchet after the user zooms in.
   */
  setScale(scale: number): void {
    const zoomMin = this._effectiveZoomMin();
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
    if (!this._doc || this._doc.pageCount === 0 || !this._scaleEstablished) {
      // IX9 F1 (family-unified pre-load semantics): a setScale before the doc is
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
    // (page index, intra-page fraction) pair. Anchoring on a page — not on the raw
    // scrollTop — is what keeps the re-anchor exact despite the scale-INVARIANT
    // desk padding and inter-page gaps (only the page heights scale, so a whole-
    // scroll linear rescale would drift by the padding). The point we pin is the
    // content under the pointer: content-y = scrollTop + anchorY (anchorY 0 ⇒ the
    // viewport top, the historical behaviour).
    const r0 = this._range();
    const scrollTop0 = this._scrollHost.scrollTop;
    const anchorContentY = scrollTop0 + anchorY;
    // Which page does that content-y fall in? `computeVisibleRange` attributes a
    // point in the trailing gap to the page ABOVE it, so clamp the fraction to
    // [0,1] to pin the page rather than drift into the gap.
    const top = this._pageIndexAtOffset(r0, anchorContentY);
    const h0 = this._heights[top] || 0;
    let intraFrac = h0 > 0 ? (anchorContentY - r0.offsets[top]) / h0 : 0;
    intraFrac = Math.min(1, Math.max(0, intraFrac));

    // HORIZONTAL anchor (gesture only — a non-gesture setScale leaves scrollLeft
    // untouched, matching the historical behaviour). The page's left edge sits at
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
    const r1 = computeVisibleWindow(
      this._scrollGeometry,
      0,
      this._scrollHost.clientHeight,
      this._overscan(),
    );
    this._spacer.style.height = `${r1.totalHeight}px`;
    // The page px width changed with the scale, so the horizontal extent moves too.
    this._syncSpacerWidth();

    // Pin the same fractional position of the same page under the pointer (or the
    // viewport top for a non-gesture zoom): the on-screen y of that content point
    // must stay at `anchorY`, so newScrollTop = newContentY − anchorY.
    const maxTop = Math.max(0, r1.totalHeight - this._scrollHost.clientHeight);
    const newContentY = (r1.offsets[top] ?? 0) + intraFrac * (this._heights[top] || 0);
    // The leading desk padding is fixed viewport space: none of it scales. If the
    // anchor is still inside that padding, keep the native scroll offset instead
    // of snapping to page 0's offset and hiding the margin after a programmatic
    // zoom at scrollTop 0.
    const reanchoredTop = anchorContentY < (r0.offsets[0] ?? 0)
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
    // (that blanks each visible page to white every tick). Instead:
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
   *  page at its natural pt→px width). This is the viewer's absolute `_scale`
   *  (`widthPt × PT_TO_PX × _scale` is the drawn width), so it reads `1` at true
   *  100% and, after the initial fit-to-width, the base fit factor. Before the
   *  fit is established it reports a latched pre-load `setScale` (IX9 F1) if one
   *  is pending — matching what a single-canvas viewer would show — else `1`. */
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
    this.setScale(prevZoomStep(this.getScale(), this._effectiveZoomMin()));
  }

  /** The configured floor, extended down to the current width fit when needed.
   * Before layout establishes there is no fit exception, so pre-load requests
   * retain the documented `[zoomMin, zoomMax]` clamp. */
  private _effectiveZoomMin(): number {
    const configured = this._opts.zoomMin ?? 0.1;
    return this._scaleEstablished && this._prevBase > 0
      ? Math.min(configured, this._prevBase)
      : configured;
  }

  /**
   * IX9 {@link ZoomableViewer} — fit a page's WIDTH to the container (the classic
   * continuous-scroll "fit width"). Sets the scale to the width-fit base for the
   * current container and the widest page, then re-anchors + re-renders via
   * {@link setScale}. Defers (no-op) while the container is unlaid-out. A width
   * fit below `zoomMin` becomes the effective floor so it remains reachable.
   */
  fitWidth(): void {
    this._fit('width');
  }

  /**
   * IX9 {@link ZoomableViewer} — fit a WHOLE page (width and height) inside the
   * container so one page is visible without scrolling; takes the tighter of the
   * width/height fit. Uses the FIRST page's size; unlike document-wide
   * `fitWidth()`, this is a one-page fit target. Defers while unlaid-out.
   */
  fitPage(): void {
    this._fit('page');
  }

  /** Shared fit for {@link fitWidth}/{@link fitPage}: width fit uses the widest
   *  page so it agrees with the horizontal spacer; page fit continues to target
   *  the first page's own width and height. Applies via {@link setScale} so the
   *  flicker-free re-anchor / settle path and `onScaleChange` all run. */
  private _fit(mode: 'width' | 'page'): void {
    if (!this._doc || this._doc.pageCount === 0) return;
    const size = this._doc.pageSize(0);
    const widthPt = mode === 'width' ? this._widestPageWidthPt() : size.widthPt;
    const scale = fitScale(
      {
        contentWidth: widthPt * PT_TO_PX,
        contentHeight: size.heightPt * PT_TO_PX,
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
    if (!this._doc || this._doc.pageCount === 0) return;
    const r = this._range();
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
        this._renderSlot(i, slot);
      } else {
        this._previewSlot(existing, i, r);
      }
    }
    this._emitVisiblePageChange(r);
  }

  /**
   * CSS-preview a single already-mounted slot at the new geometry (design §7): the
   * wrapper is repositioned + sized (via `_positionSlot`), the canvas bitmap is
   * STRETCHED to the new CSS size (no `canvas.width` — the device buffer, and thus
   * the drawn pixels, are left intact, just scaled by the browser), and the text
   * overlay is scaled by `newScale / renderedScale` so it tracks the stretched
   * page. `renderedScale <= 0` means the slot's first render hasn't resolved yet
   * (nothing to stretch); the pending render captured the current scale, so it
   * lands correct and no preview is needed.
   */
  private _previewSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    this._positionSlot(slot, i, r);
    // Stretch the existing bitmap to the new CSS box (device buffer untouched).
    slot.canvas.style.width = `${this._pageWidthPx(i)}px`;
    slot.canvas.style.height = `${this._pageHeightPx(i)}px`;
    if (slot.renderedScale > 0) {
      const ratio = this._scale / slot.renderedScale;
      if (slot.textLayer) {
        slot.textLayer.style.transformOrigin = '0 0';
        // The wrapper has already moved to the new layout dimensions. Keep the
        // overlay's box at its last committed dimensions while scaling it;
        // otherwise width/height:100% resolve against the NEW wrapper and the
        // transform applies the zoom ratio a second time. Besides mis-sizing the
        // overlay, that transformed border box becomes scrollable overflow. In
        // worker mode each page settles independently, so the document extent
        // would then shrink page-by-page and could strand the viewport over blank
        // space. The transformed box below is exactly the current page box:
        // (new / ratio) × ratio = new.
        slot.textLayer.style.width = `${this._pageWidthPx(i) / ratio}px`;
        slot.textLayer.style.height = `${this._pageHeightPx(i) / ratio}px`;
        slot.textLayer.style.transform = `scale(${ratio})`;
      }
      if (slot.commentMargin) this._commentUi?.previewReadOnlyCommentMargin(slot.commentMargin, ratio);
      for (const marker of slot.commentTintLayer?.children ?? []) {
        if ((marker as HTMLElement).dataset.ooxmlCommentMarker === undefined) continue;
        (marker as HTMLElement).style.transform = `translate(-50%,-50%) scale(${ratio})`;
      }
      if (slot.commentTintLayer) slot.commentTintLayer.style.visibility = '';
      if (slot.commentMargin) slot.commentMargin.style.visibility = '';
      if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = '';
      return;
    }
    // No committed geometry exists during the first render, so there is nothing
    // trustworthy to preview yet.
    if (slot.commentTintLayer) slot.commentTintLayer.style.visibility = 'hidden';
    if (slot.commentMargin) slot.commentMargin.style.visibility = 'hidden';
    if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = 'hidden';
  }

  /** Restore a text overlay after its transient CSS zoom preview. */
  private _clearTextLayerPreview(layer: HTMLDivElement): void {
    layer.style.transform = '';
    layer.style.transformOrigin = '';
    layer.style.width = '100%';
    layer.style.height = '100%';
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
    if (this._destroyed || !this._doc || this._doc.pageCount === 0) return;
    for (const [i, slot] of [...this._slots]) {
      // Skip slots already at the current scale (a slot that entered the window
      // during the burst mounted fresh at the current scale — nothing to settle).
      if (slot.renderedScale === this._scale) continue;
      this._refreshSlotAtomically(i, slot);
    }
  }

  /**
   * Refresh one mounted slot without exposing an intermediate blank frame.
   *
   * WORKER: re-dispatch the bitmap render into the SAME canvas. The worker path
   * sizes the device buffer and `transferFromImageBitmap`s it in ONE synchronous
   * step (no await between `canvas.width = …` and the transfer), so the browser
   * never composites an intermediate blank frame — no spare canvas is needed. The
   * `renderedScale === _scale` gate in `_settleRender` plus the epoch gate inside
   * `_renderSlotBitmap` keep this correct and idempotent.
   *
   * MAIN: `renderPage` (via renderDocumentToCanvas) synchronously sets
   * `canvas.width = …` (which CLEARS the backing store to blank) BEFORE its first
   * await and paints AFTER — so rendering into the on-screen canvas would flash it
   * white. Render into a SPARE off-DOM canvas instead; only once it resolves at the
   * current epoch do we swap it into the wrapper (replacing the old canvas, which is
   * DISCARDED — the pooled unit is the slot, not the canvas). The old canvas keeps
   * showing the stretched preview until the instant of the swap — blank-free.
   */
  private _refreshSlotAtomically(i: number, slot: PageSlot): void {
    if (!this._doc) return;
    const dpr = this._dpr();
    const widthPx = this._pageWidthPx(i);
    const scale = this._scale;
    const epoch = this._renderEpoch;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
      return;
    }

    // Main mode: double-buffer. Render into a spare canvas kept off-DOM. The
    // spare REPLACES the on-screen canvas on swap, so it must carry the page
    // shadow too — otherwise a settle would silently drop it.
    const spare = document.createElement('canvas');
    spare.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(spare);
    const spareDispatcher = new StaticCanvasRenderDispatcher(spare, false);
    const generation = spareDispatcher.begin();
    const runs: DocxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const wantRuns = wantOverlay || this._findActive || !!slot.commentTintLayer;
    const onTextRun = wantRuns ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    renderDocxFocusedPage(this._doc, spare, i, 'main', {
      width: widthPx,
      dpr,
      imageResources: this._opts.imageResources,
      defaultTextColor: this._opts.defaultTextColor,
      currentDate: this._currentDate,
      ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
      onTextRun,
    })
      .then(() => {
        // Discard if superseded: a later setScale bumped the epoch (this spare is
        // at a stale scale), or the slot recycled / moved to another page. Drop
        // the spare (it is off-DOM, so GC reclaims it) and do NOT swap.
        if (
          !spareDispatcher.isCurrent(generation) ||
          epoch !== this._renderEpoch ||
          this._slots.get(i) !== slot ||
          slot.renderedPage !== i
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
          this._clearTextLayerPreview(slot.textLayer);
          if (wantOverlay) {
            const { width, height } = this._canvasCssPx(spare);
            buildDocxTextLayer(
              slot.textLayer,
              runs,
              width,
              height,
              this._hyperlinkHandler(),
              (font) => this._measureForFont(font),
              i,
            );
          }
        }
        if (wantRuns) this._refreshFindRuns(i, runs);
        this._commitCommentRuns(i, slot, runs);
        this._redrawSlotHighlights(i, slot);
      })
      .catch((err: unknown) => {
        if (
          spareDispatcher.isCurrent(generation) &&
          epoch === this._renderEpoch &&
          this._slots.get(i) === slot &&
          slot.renderedPage === i
        ) this._reportRenderError(err);
        spareDispatcher.destroy();
      });
  }

  // ─── §17.13.5 tracked-changes view toggle ─────────────────────────────────

  /**
   * ECMA-376 §17.13.5 — switch between the final view (`false`, the default:
   * deletions hidden) and the markup view (`true`: author-coloured revision
   * decoration + margin change bars) at runtime. Every mounted page
   * re-renders against the selected layout variant; find results are
   * invalidated because the visible text differs between the views.
   */
  async setShowTrackedChanges(value: boolean): Promise<void> {
    const generation = ++this._layoutViewGeneration;
    const doc = this._doc;
    if (this._showTrackedChanges === value) {
      if (doc) await selectDocxLayoutView(doc, {
        showTrackedChanges: value,
        currentDate: this._currentDate,
      }, this);
      return;
    }
    // The markup view is a different retained layout with its own pagination,
    // so move the document's active variant before reading any geometry from
    // it — page count and page heights are about to change.
    const selected = doc
      ? await selectDocxLayoutView(doc, {
          showTrackedChanges: value,
          currentDate: this._currentDate,
        }, this)
      : true;
    if (!selected) return;
    if (this._destroyed || generation !== this._layoutViewGeneration || doc !== this._doc) return;
    this._showTrackedChanges = value;
    this._find.invalidate();
    // Re-render every mounted slot at the new variant, and relayout: heights,
    // spacer and mount window all follow the new page count, and a shrinking
    // document must recycle slots that are now out of range rather than ask for
    // pages that no longer exist.
    this._applyLayoutPublication({
      pageCount: doc?.pageCount ?? 0,
      exact: true,
      complete: doc?.layoutComplete !== false,
    });
  }

  /**
   * Scroll so page `index`'s top edge sits at the viewport top. Clamps `index` to
   * `[0, pageCount-1]` (the pager convention) and the resulting scrollTop to
   * `[0, totalHeight − viewportHeight]` so the last pages don't scroll past the
   * end. A no-op when nothing is loaded or the document is empty.
   *
   * `opts.behavior` ('auto' | 'smooth', default 'auto') is honoured via
   * `scrollHost.scrollTo({ top, behavior })` when the host supports it (a real
   * browser); the stub-DOM has no `scrollTo`, so the fallback sets `scrollTop`
   * directly (which is what the tests assert). We then call `_mountVisible` once.
   *
   * MOUNTING CAVEAT: synchronous mounting of the target page is guaranteed only on
   * the DEFAULT/'auto' path — there `scrollTop` has already jumped to `top`, so the
   * `_mountVisible` call reads the final scroll position and the target page's slots
   * exist immediately. With `behavior: 'smooth'` the scroll animates ASYNCHRONOUSLY:
   * `scrollTop` is still near the old position when `_mountVisible` runs, so the
   * target page mounts lazily via the animation's subsequent `scroll` events, not
   * from this call.
   */
  scrollToPage(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    if (!this._doc || this._doc.pageCount === 0 || !this._scaleEstablished) return;
    const clamped = Math.max(0, Math.min(index, this._doc.pageCount - 1));
    // Recompute offsets from the current heights (independent of scrollTop).
    const r = computeVisibleWindow(
      this._scrollGeometry,
      0,
      this._scrollHost.clientHeight,
      this._overscan(),
    );
    const target = r.offsets[clamped] ?? 0;
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

  private _scrollToPageTarget(
    page: number,
    target: Readonly<{ x: number; y: number; w: number; h: number }>,
    opts?: { behavior?: 'auto' | 'smooth' },
  ): void {
    const range = computeVisibleWindow(
      this._scrollGeometry,
      0,
      this._scrollHost.clientHeight,
      this._overscan(),
    );
    const pageWidth = this._pageWidthPx(page);
    const { left: paddingLeft } = this._padH();
    const pageLeft = Math.max(
      paddingLeft,
      (this._scrollHost.clientWidth - pageWidth) / 2,
    ) + this._reviewOriginPx;
    const maxTop = Math.max(0, range.totalHeight - this._scrollHost.clientHeight);
    const spacerWidth = this._spacer.offsetWidth || Number.parseFloat(this._spacer.style.width) || 0;
    const maxLeft = Math.max(0, spacerWidth - this._scrollHost.clientWidth);
    const top = Math.min(maxTop, Math.max(
      0,
      (range.offsets[page] ?? 0) + target.y + target.h / 2 - this._scrollHost.clientHeight / 2,
    ));
    const left = Math.min(maxLeft, Math.max(
      0,
      pageLeft + target.x + target.w / 2 - this._scrollHost.clientWidth / 2,
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
  }

  private _resetCommentNavigation(): void {
    this._commentNavigationGeneration++;
    this._commentPageById.clear();
    this._commentRunsByPage.clear();
    this._commentIndexedPages.clear();
    this._commentScanFrontier = 0;
  }

  private _advanceCommentScanFrontier(): void {
    while (this._commentIndexedPages.has(this._commentScanFrontier)) {
      this._commentScanFrontier++;
    }
  }

  /** Join one page's retained run geometry to every comment while it is already
   * in hand. This makes the page scan shared by all application-owned list rows
   * instead of repeating a document-prefix scan for each clicked comment. */
  private _indexCommentPages(
    page: number,
    runs: readonly Readonly<DocxTextRunInfo>[],
    anchors: ReturnType<DocxDocument['commentAnchorRanges']>,
  ): void {
    if (this._commentIndexedPages.has(page)) return;
    for (const anchor of anchors) {
      if (this._commentPageById.has(anchor.commentId)) continue;
      if (resolveCommentAnchorRuns(anchor, runs).length > 0) {
        this._commentPageById.set(anchor.commentId, page);
      }
    }
    this._commentIndexedPages.add(page);
    this._advanceCommentScanFrontier();
  }

  /** Collect run geometry at most once per page and scale. Concurrent navigation
   * requests share the in-flight Promise; a zoom retries rather than committing
   * coordinates captured at the superseded scale. */
  private async _commentRunsForPage(
    page: number,
    doc: DocxDocument,
  ): Promise<readonly Readonly<DocxTextRunInfo>[] | null> {
    while (!this._destroyed && this._doc === doc) {
      const scale = this._scale;
      let entry = this._commentRunsByPage.get(page);
      if (!entry || entry.scale !== scale) {
        const runs = doc.collectPageRuns(page, {
          width: this._pageWidthPx(page),
          currentDate: this._currentDate,
          ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        });
        entry = { scale, runs };
        this._commentRunsByPage.set(page, entry);
      }
      try {
        const runs = await entry.runs;
        if (this._destroyed || this._doc !== doc) return null;
        if (this._scale !== scale) continue;
        return runs;
      } catch (error) {
        if (this._commentRunsByPage.get(page) === entry) {
          this._commentRunsByPage.delete(page);
        }
        throw error;
      }
    }
    return null;
  }

  /**
   * Reveal a top-level authored comment by its DOCX comment id. This is the
   * navigation primitive for an application-owned comment list: the Viewer
   * resolves the comment's page lazily, caches that stable page index, scrolls
   * the first anchored text run into view, and selects the thread.
   *
   * Returns `false` for an unknown id or a comment with no rendered anchor.
   */
  async goToComment(
    commentId: string,
    opts?: { pageIndex?: number; behavior?: 'auto' | 'smooth' },
  ): Promise<boolean> {
    if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
    const doc = this._doc;
    if (!doc || !doc.comments.some((comment) =>
      comment.id === commentId && comment.parentId === undefined)) return false;
    const generation = ++this._commentNavigationGeneration;
    const startedWithProvisionalLayout = !doc.layoutComplete;
    let anchors = doc.commentAnchorRanges().filter((anchor) => anchor.commentId === commentId);
    const requestedPage = opts?.pageIndex;
    if (requestedPage !== undefined && (!Number.isInteger(requestedPage) || requestedPage < 0)) {
      return false;
    }
    let page = requestedPage ?? this._commentPageById.get(commentId);
    let targetRun: Readonly<DocxTextRunInfo> | undefined;
    const scanAvailablePages = async (): Promise<number | undefined> => {
      const allAnchors = doc.commentAnchorRanges();
      while (page === undefined && this._commentScanFrontier < doc.pageCount) {
        const index = this._commentScanFrontier;
        const runs = await this._commentRunsForPage(index, doc);
        if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
        if (this._doc !== doc || generation !== this._commentNavigationGeneration || !runs) {
          return undefined;
        }
        this._indexCommentPages(index, runs, allAnchors);
        page = this._commentPageById.get(commentId);
      }
      return page;
    };

    if (requestedPage === undefined && page === undefined && anchors.length > 0) {
      await scanAvailablePages();
    }

    // A progressive prefix can prove that a comment is present without yet
    // proving its page. Worker partials may expose no anchor projection at all.
    // In either case, "not in the prefix" is not "does not exist": finish the
    // canonical layout, discard provisional page joins, and retry once.
    const needsAuthoritativeLayout = startedWithProvisionalLayout && (
      (requestedPage !== undefined && requestedPage >= doc.pageCount) ||
      (requestedPage === undefined && page === undefined)
    );
    if (needsAuthoritativeLayout) {
      await this._errorRouter.ownBackgroundLifecycle(() => doc.waitUntilLayoutComplete());
      if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
      if (this._doc !== doc || generation !== this._commentNavigationGeneration) return false;
      anchors = doc.commentAnchorRanges().filter((anchor) => anchor.commentId === commentId);
      this._commentPageById.clear();
      this._commentRunsByPage.clear();
      this._commentIndexedPages.clear();
      this._commentScanFrontier = 0;
      page = requestedPage;
      if (requestedPage === undefined && anchors.length > 0) await scanAvailablePages();
    }
    if (anchors.length === 0) return false;
    if (requestedPage !== undefined && requestedPage >= doc.pageCount) return false;
    if (page === undefined) return false;
    const runs = await this._commentRunsForPage(page, doc);
    if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
    if (this._doc !== doc || generation !== this._commentNavigationGeneration || !runs) return false;
    targetRun = anchors.flatMap((anchor) => resolveCommentAnchorRuns(anchor, runs))[0];
    if (!targetRun) return false;

    this._activeCommentId = commentId;
    this._activeCommentPage = page;
    this._elementContext = null;
    this._scrollToPageTarget(page, targetRun, opts);
    for (const [mountedPage, slot] of this._slots) {
      this._redrawSlotComments(mountedPage, slot);
    }
    this._emitSelectionContextChange();
    return true;
  }

  /** Search the complete document, including pages outside the virtualized
   * mounted window. Matching is case-insensitive by default. */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<DocxMatchLocation>[]> {
    if (!this._doc) return [];
    const generation = ++this._findRequestGeneration;
    // Search spans every page, so a progressively-loaded document has to finish
    // laying out first — otherwise the search silently covers only the pages
    // that happen to exist yet. Guarded on the document actually being
    // incomplete so that an ordinary document still starts its find
    // synchronously: `findText()` followed immediately by `clearFind()` must
    // cancel the find, which it cannot do if the find has not begun.
    if (this._doc.layoutComplete === false) {
      const doc = this._doc;
      await this._errorRouter.ownBackgroundLifecycle(
        () => doc.waitUntilLayoutComplete(),
      );
      if (this._destroyed || this._doc !== doc || generation !== this._findRequestGeneration) {
        return [];
      }
    }
    this._findActive = query.length > 0;
    const matches = await this._errorRouter.ownAwaitable(
      () => this._find.find(query, opts),
    );
    this._redrawHighlights();
    return matches;
  }

  /** Activate and reveal the next match, wrapping at the end. */
  async findNext(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** Activate and reveal the previous match, wrapping at the beginning. */
  async findPrev(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** Clear the current query and every mounted highlight. */
  clearFind(): void {
    this._findRequestGeneration++;
    this._findActive = false;
    this._find.invalidate();
    this._redrawHighlights();
  }

  private async _activateMatch(
    match: FindMatch<DocxMatchLocation> | null,
  ): Promise<FindMatch<DocxMatchLocation> | null> {
    if (match) this.scrollToPage(match.location.page);
    this._redrawHighlights();
    return match;
  }

  private async _collectPageRuns(page: number): Promise<DocxTextRunInfo[]> {
    if (!this._doc) return [];
    return this._doc.collectPageRuns(page, {
      width: this._pageWidthPx(page),
      currentDate: this._currentDate,
      ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
    });
  }

  private _redrawHighlights(): void {
    for (const [page, slot] of this._slots) this._redrawSlotHighlights(page, slot);
  }

  private _refreshFindRuns(page: number, runs: DocxTextRunInfo[]): void {
    if (this._findActive) this._find.setPageRuns(page, runs);
  }

  private _commitCommentRuns(
    page: number,
    slot: PageSlot,
    runs: readonly Readonly<DocxTextRunInfo>[],
  ): void {
    if (!slot.commentTintLayer) return;
    slot.commentRuns = Object.freeze([...runs]);
    this._commentRunsByPage.set(page, {
      scale: this._scale,
      runs: Promise.resolve(slot.commentRuns),
    });
    // Only extend the index in document order. A header/footer anchor can repeat
    // on later pages; accepting an arbitrary mounted page first would make
    // `goToComment()` skip the authored range's earliest rendered occurrence.
    if (this._doc && page === this._commentScanFrontier) {
      this._indexCommentPages(page, slot.commentRuns, this._doc.commentAnchorRanges());
    }
    slot.commentTintLayer.style.transform = '';
    slot.commentTintLayer.style.transformOrigin = '';
    // Rebuild against the committed run geometry while the transient preview is
    // still hidden, then reveal tint, cards, and connectors together.
    this._redrawSlotComments(page, slot);
    slot.commentTintLayer.style.visibility = '';
    if (slot.commentMargin) slot.commentMargin.style.visibility = '';
    if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = '';
  }

  private _redrawSlotComments(page: number, slot: PageSlot): void {
    if (!this._doc || !slot.commentTintLayer) return;
    this._syncCommentMarginGeometry(slot.commentMargin);
    const commentUi = this._commentUi;
    if (!commentUi) {
      slot.commentTintLayer.replaceChildren();
      slot.commentMargin?.replaceChildren();
      slot.commentDecorationLayer?.replaceChildren();
      slot.commentGeometry = null;
      return;
    }
    slot.commentGeometry = commentUi.buildDocxCommentMargin(
      slot.commentTintLayer,
      slot.commentMargin,
      slot.commentRuns,
      { comments: this._doc.comments, anchors: this._doc.commentAnchorRanges() },
      this._pageWidthPx(page),
      this._pageHeightPx(page),
      this._activeCommentId,
      (id, active) => {
        const next = active ? id : this._activeCommentId === id ? null : this._activeCommentId;
        if (next === this._activeCommentId) return;
        this._activeCommentId = next;
        this._activeCommentPage = next ? page : null;
        this._elementContext = null;
        for (const [mountedPage, mountedSlot] of this._slots) {
          this._redrawSlotComments(mountedPage, mountedSlot);
        }
        this._emitSelectionContextChange();
      },
      this._commentZoom(),
      READ_ONLY_COMMENT_MARGIN_WIDTH_PX,
      this._commentsOptions()?.markers !== false,
      this._commentsOptions()?.includeResolved === true,
      slot.commentDecorationLayer
        ? () => this._scheduleCommentGeometry(page, slot, false)
        : undefined,
      slot.commentDecorationLayer
        ? () => this._scheduleCommentGeometry(page, slot, true)
        : undefined,
    );
    this._redrawSlotCommentConnectors(page, slot);
  }

  private _redrawSlotCommentConnectors(page: number, slot: PageSlot): void {
    const layer = slot.commentDecorationLayer;
    const margin = slot.commentMargin;
    const geometry = slot.commentGeometry;
    const connectorOptions = this._commentsOptions()?.connectors;
    if (!layer || !margin || !geometry || !connectorOptions) return;
    const width = this._pageWidthPx(page);
    const height = this._pageHeightPx(page);
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

  /** Coalesce card measurement and scroll-only connector projection into one
   * geometry refresh per animation frame. A full refresh dominates a pending
   * connector-only refresh for the same slot. */
  private _scheduleCommentGeometry(
    page: number,
    slot: PageSlot,
    connectorsOnly = false,
  ): void {
    const current = this._pendingCommentGeometry.get(page);
    this._pendingCommentGeometry.set(page, {
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
      for (const [pendingPage, entry] of pending) {
        const { slot: pendingSlot, connectorsOnly: pendingConnectorsOnly } = entry;
        if (
          this._slots.get(pendingPage) === pendingSlot &&
          pendingSlot.renderedScale === this._scale
        ) {
          if (pendingConnectorsOnly) {
            this._redrawSlotCommentConnectors(pendingPage, pendingSlot);
          } else {
            this._redrawSlotComments(pendingPage, pendingSlot);
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

  private _redrawSlotHighlights(page: number, slot: PageSlot): void {
    if (!this._findActive) {
      slot.highlightLayer.innerHTML = '';
      return;
    }
    const runs = this._find.pageRuns(page);
    if (!runs) {
      slot.highlightLayer.innerHTML = '';
      return;
    }
    buildDocxHighlightLayer(
      slot.highlightLayer,
      runs,
      this._find.pageHighlights(page),
      this._pageWidthPx(page),
      this._pageHeightPx(page),
      (font) => this._measureForFont(font),
      this._opts.findHighlightColors,
    );
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
    if (!this._doc || this._doc.pageCount === 0) return;
    // Zero-width recovery: first non-zero layout establishes the base scale.
    if (!this._scaleEstablished) {
      this.relayout();
      return;
    }
    if (this._opts.refitOnResize === false) {
      // Fixed-scale hosts (for example VS Code previews) must not turn a pane
      // resize into an implicit zoom. Recompute the visible window and horizontal
      // centering only; page geometry and rendered bitmaps remain valid.
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
      // newly-revealed pages; without it those rows stay blank until the user
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
    // `_prevBase` is updated before `setScale`, so a newly smaller width fit also
    // becomes the effective floor for this resize. That preserves the multiplier
    // without ratcheting an initially fitted oversized page up to zoomMin.
    this.setScale(newBase * mult);
    // `setScale` no-ops when the clamped scale is unchanged (e.g. already pinned at
    // a clamp boundary), which would skip its preview + settle. A width+height
    // growth that ends up clamped to the same scale must still reveal the taller
    // viewport's rows, so mount here too. Idempotent when `setScale` ran: the
    // window is already mounted and every present slot is a re-position no-op.
    this._mountVisible();
  }

  get topVisiblePage(): number {
    return this._lastRange?.topIndex ?? 0;
  }

  /** @internal test hook: page indices currently mounted. */
  mountedPageIndicesForTest(): number[] {
    return [...this._slots.keys()];
  }

  /** @internal test hook: the current absolute px-per-pt scale. */
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

  /** @internal test hook: the content point (page index + intra-page fraction)
   *  currently under viewport-y `y` (px from the scroll host top). Lets a test
   *  capture "what is under the cursor" before a zoom and re-query its on-screen
   *  y afterwards to assert the pointer-anchored invariant. */
  contentAtViewportYForTest(y: number): { page: number; frac: number } {
    const r = this._range();
    const contentY = this._scrollHost.scrollTop + y;
    const page = this._pageIndexAtOffset(r, contentY);
    const h = this._heights[page] || 0;
    const frac = h > 0 ? Math.min(1, Math.max(0, (contentY - r.offsets[page]) / h)) : 0;
    return { page, frac };
  }

  /** @internal test hook: inverse of {@link contentAtViewportYForTest} — the
   *  current viewport-y (px from the scroll host top) of the content point at
   *  (`page`, intra-page `frac`). */
  viewportYOfForTest(page: number, frac: number): number {
    const r = this._range();
    const contentY = (r.offsets[page] ?? 0) + frac * (this._heights[page] || 0);
    return contentY - this._scrollHost.scrollTop;
  }

  /** Return the owning engine's latest content-free package-usage snapshot. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this._doc) throw new Error('Document not loaded');
    return await this._doc.getResourceMetrics();
  }

  /** Return the current mounted text selection or clicked drawing context. */
  getSelectionContext(options: DocxSelectionContextOptions = {}): DocxSelectionContext | null {
    if (this._destroyed) throw new Error('DocxScrollViewer is destroyed');
    const comment = this._doc && this._activeCommentId !== null && this._activeCommentPage !== null
      ? createDocxCommentSelectionContext(
          this._doc.comments,
          this._doc.commentAnchorRanges(),
          this._activeCommentId,
          this._activeCommentPage,
          options,
        )
      : null;
    if (comment) return comment;
    const text = this._opts.enableTextSelection
      ? readDocxTextSelectionContext(
          this._wrapper,
          this._wrapper.ownerDocument?.getSelection?.() ?? null,
          options,
        )
      : null;
    return text ?? (this._elementContext
      ? limitDocxElementContext(this._elementContext, options.maxTextCharacters)
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

  private _setElementContext(context: DocxElementContext | null): void {
    this._elementContext = context ? structuredClone(context) : null;
    this._redrawElementOutlines();
    this._emitSelectionContextChange();
  }

  private _invalidateElementContext(notify = true): void {
    this._elementHitGeneration++;
    this._elementContext = null;
    this._redrawElementOutlines();
    if (notify) this._emitSelectionContextChange();
  }

  private _redrawElementOutlines(): void {
    for (const [page, slot] of this._slots) this._redrawElementOutlineForSlot(page, slot);
  }

  private _redrawElementOutlineForSlot(pageIndex: number, slot: PageSlot): void {
    const context = this._elementContext;
    const doc = this._doc;
    if (!context || !doc || context.pageIndex !== pageIndex) {
      renderCanvasElementOutline(slot.elementLayer, null);
      return;
    }
    const page = doc.pageSize(pageIndex);
    renderCanvasElementOutline(slot.elementLayer, {
      x: context.bounds.xPt / page.widthPt,
      y: context.bounds.yPt / page.heightPt,
      width: context.bounds.widthPt / page.widthPt,
      height: context.bounds.heightPt / page.heightPt,
    });
  }

  private async _onElementClick(event: MouseEvent): Promise<void> {
    if (this._destroyed || event.defaultPrevented || event.button !== 0) return;
    await this._resolveContextAt(event);
  }

  private _onContextMenu(event: MouseEvent): void {
    let context: Promise<DocxSelectionContext | null> | undefined;
    this._opts.onContextMenu?.({
      originalEvent: event,
      getContext: () => context ??= this._resolveContextAt(event),
    });
  }

  private async _resolveContextAt(event: MouseEvent): Promise<DocxSelectionContext | null> {
    const doc = this._doc;
    if (this._destroyed || !doc) return null;
    if (this._opts.enableTextSelection && readDocxTextSelectionContext(
      this._wrapper,
      this._wrapper.ownerDocument?.getSelection?.() ?? null,
    )) {
      this._emitSelectionContextChange();
      return this._destroyed ? null : this.getSelectionContext();
    }
    if (!this._opts.enableElementSelection) return this.getSelectionContext();
    const target = event.target as Node | null;
    const entry = [...this._slots].find(([, slot]) =>
      target !== null && slot.wrapper.contains(target));
    if (!entry) {
      this._invalidateElementContext();
      return null;
    }
    const [pageIndex, slot] = entry;
    const rect = slot.canvas.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      this._invalidateElementContext();
      return null;
    }
    const localX = event.clientX - rect.left;
    const localY = event.clientY - rect.top;
    if (localX < 0 || localY < 0 || localX > rect.width || localY > rect.height) {
      this._invalidateElementContext();
      return null;
    }
    const generation = ++this._elementHitGeneration;
    const pageSize = doc.pageSize(pageIndex);
    let context: DocxElementContext | null;
    try {
      context = await doc.getElementContextAt(pageIndex, {
        xPt: localX / rect.width * pageSize.widthPt,
        yPt: localY / rect.height * pageSize.heightPt,
      }, {
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        maxTextCharacters: MAX_DOCX_ELEMENT_TEXT_CHARACTERS,
      });
    } catch (error) {
      if (this._destroyed || generation !== this._elementHitGeneration || doc !== this._doc) {
        return null;
      }
      throw error;
    }
    if (this._destroyed || generation !== this._elementHitGeneration || doc !== this._doc) return null;
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
    this._findRequestGeneration++;
    this._errorRouter.close();
    this._unbindLayoutDocument();
    this._layoutViewGeneration++;
    this._resetCommentNavigation();
    this._find.invalidate();
    this._findActive = false;
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
    this._documentOwner.close();
    this._wrapper.remove();
  }
}
