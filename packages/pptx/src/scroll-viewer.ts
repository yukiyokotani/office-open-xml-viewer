import { EMU_PER_PX, type FindMatch, type FindMatchesOptions, type HyperlinkTarget, type OoxmlResourceMetrics, type ZoomableViewer, openExternalHyperlink } from '@silurus/ooxml-core';
import {
  computeUniformVisibleWindow,
  type VisibleWindow,
} from '@silurus/ooxml-core/internal/virtual-scroll';
import {
  createCanvasElementOutlineLayer,
  CanvasViewerErrorRouter,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import { ScrollViewerShell } from '@silurus/ooxml-core/internal/scroll-viewer-shell';
import { HighlightLayerController } from '@silurus/ooxml-core/internal/highlight-layer-controller';
import { BitmapSlotRenderer, type BitmapSlotHooks } from '@silurus/ooxml-core/internal/bitmap-slot-renderer';
import { MainSlotRenderer } from '@silurus/ooxml-core/internal/main-slot-renderer';
import { SlotLayerController } from '@silurus/ooxml-core/internal/slot-layer-controller';
import { ScrollNavigationController } from '@silurus/ooxml-core/internal/scroll-navigation-controller';
import { DEFAULT_SCROLL_PAGE_SHADOW, ScrollViewportPolicy } from '@silurus/ooxml-core/internal/scroll-viewport-policy';
import { VisibleUnitEvents } from '@silurus/ooxml-core/internal/visible-unit-events';
import { ScrollLoadController } from '@silurus/ooxml-core/internal/scroll-load-controller';
import { DEFAULT_ZOOM_SETTLE_MS, SlotScroller, clearTextLayerPreview, createSlotHost, createCommentSlotLayers } from '@silurus/ooxml-core/internal/slot-scroller';
import { CommentMarginController } from '@silurus/ooxml-core/internal/comment-margin-controller';
import { ScrollZoomController } from '@silurus/ooxml-core/internal/scroll-zoom-controller';
import { SelectionContextController } from '@silurus/ooxml-core/internal/selection-context-controller';
import { CommentOverlayController } from '@silurus/ooxml-core/internal/comment-overlay-controller';
import { PptxPresentation, type LoadOptions } from './presentation';
import type { PptxScrollSlot as SlideSlot } from './scroll-slot';
import type { PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { PptxFindController, type PptxMatchLocation } from './find';
import { buildPptxHighlightLayer } from './find-highlight-layer';
import {
  createPptxCommentSelectionContext,
  readPptxTextSelectionContext,
} from './selection-context';
import type {
  PptxElementContext,
  PptxSelectionContext,
  PptxSelectionContextOptions,
} from './element-selection';
import {
  limitPptxElementContext,
  MAX_ELEMENT_TEXT_CHARACTERS,
} from './element-selection';
import type { PptxCommentsOptions } from './comment-margin';
import type { PptxScrollViewerOptions } from './scroll-viewer-options';
import { pptxCommentOccurrenceKey } from './comment-occurrence';
import { renderPptxFocusedSlide } from './focused-view-runtime';
import { PartFailureDelivery, prepareErrorPlaceholder } from '@silurus/ooxml-core/internal/part-failures';
import { pptxPartFailureLog, type PptxPartFailure } from './part-failures';
import { PptxScrollLayoutController } from './scroll-layout-controller';
import { createPptxLoadingIndicator } from './loading-indicator';
import { PptxScrollMediaController } from './scroll-media-controller';
import { PptxScrollCommentNavigation } from './scroll-comment-navigation';

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

export type { PptxScrollViewerOptions } from './scroll-viewer-options';

export class PptxScrollViewer implements ZoomableViewer {
  private readonly _presentationOwner: TerminalResourceOwner<PptxPresentation>;
  private get _pres(): PptxPresentation | null { return this._presentationOwner.current; }
  private readonly _borrowed: boolean;
  private readonly _opts: PptxScrollViewerOptions;
  private readonly _errorRouter: CanvasViewerErrorRouter;
  private readonly _container: HTMLElement;
  private readonly _shell: ScrollViewerShell;
  private get _wrapper(): HTMLDivElement { return this._shell.wrapper; }
  private get _scrollHost(): HTMLDivElement { return this._shell.scrollHost; }
  private get _spacer(): HTMLDivElement { return this._shell.spacer; }
  /** Resolved render mode. When an engine is borrowed the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `PptxPresentation.load`. */
  private _mode: 'main' | 'worker';

  private readonly _viewport = new ScrollViewportPolicy({
    options: () => this._opts,
    container: () => this._container,
    scrollHost: () => this._scrollHost,
  });
  private readonly _zoom = new ScrollZoomController({
    scrollHost: () => this._scrollHost,
    spacer: () => this._spacer,
    count: () => this._pres?.slideCount ?? 0,
    zoomMin: () => this._opts.zoomMin ?? 0.1,
    zoomMax: () => this._opts.zoomMax ?? 4,
    baseScale: () => this._baseScale(),
    fitWidthPx: () => this._viewport.fitWidth(),
    fitContentSize: () => this._pres ? {
      width: this._pres.slideWidth / EMU_PER_PX,
      height: this._pres.slideHeight / EMU_PER_PX,
    } : null,
    indexAt: (y) => this._slideIndexAtOffset(y),
    offset: (index) => this._slideOffset(index),
    height: () => this._uniformSlideHeight,
    totalHeight: () => this._rangeAt(0, this._viewport.overscan()).totalHeight,
    recomputeHeights: () => this._recomputeHeights(),
    syncSpacerWidth: () => this._syncSpacerWidth(),
    padLeft: () => this._viewport.horizontalPadding().left,
    invalidateRender: () => { this._renderEpoch++; },
    preview: () => this._previewVisible(),
    scheduleSettle: () => this._scheduleSettle(),
    onScaleChange: (scale) => this._opts.onScaleChange?.(scale),
    relayout: () => this.relayout(),
    mountVisible: () => this._mountVisible(),
    refitOnResize: () => this._opts.refitOnResize !== false,
  });
  private get _scale(): number { return this._zoom.scale; }
  private get _scaleEstablished(): boolean { return this._zoom.established; }
  private readonly _scroller = new SlotScroller<SlideSlot, VisibleWindow>({
    spacer: () => this._spacer,
    count: () => this._pres?.slideCount ?? 0,
    range: () => this._range(),
    createSlot: () => this._createSlot(),
    attachSlot: (slot) => this._scrollHost.appendChild(slot.wrapper),
    resetSlot: (index, slot) => this._resetSlot(index, slot),
    positionSlot: (index, slot, range) => this._positionSlot(slot, index, range),
    renderSlot: (index, slot, reportErrors) => this._renderSlot(
      index, slot,
      this._opts.enableMediaPlayback === true && this._rangeContains(this._mediaRange(), index),
      reportErrors,
    ),
    previewSlot: (index, slot, range) => this._previewSlot(slot, index, range),
    settleSlot: (index, slot) => this._settleSlot(index, slot),
    renderedScale: (slot) => slot.renderedScale,
    scale: () => this._scale,
    syncSpacerWidth: () => this._syncSpacerWidth(),
    onRange: (range) => this._emitVisibleSlideChange(range),
    // PPTX keeps media handles only near the viewport. Mounted static slides
    // retain their canvas and can upgrade when the media window reaches them.
    onNewSlot: (index, slot) => this._redrawSlotComments(index, slot),
    afterMount: () => { if (this._opts.enableMediaPlayback) this._media.sync(); },
    awaitInitialRender: (index) => index < this.availableSlideCount,
    shouldSettle: (index) => !this._opts.enableMediaPlayback || this._rangeContains(this._mediaRange(), index),
  });
  private readonly _slots = this._scroller.slots;
  private readonly _layout = new PptxScrollLayoutController({
    current: () => this._pres,
    destroyed: () => this._destroyed,
    report: (error) => this._reportRenderError(error),
    reportBackground: (error) => this._errorRouter.reportBackground(
      error, this._opts.onLayoutComplete !== undefined),
    wakeComments: () => this._commentNavigation.wake(),
    resetComments: () => this._commentNavigation.resetLayout(),
    failComments: () => this._commentNavigation.failLayout(),
    scanComments: (pres) => this._scanAvailableComments(pres, true),
    renderAvailable: (pres) => {
      const mediaRange = this._opts.enableMediaPlayback ? this._mediaRange() : null;
      for (const [slideIndex, slot] of this._slots) {
        if (slideIndex >= pres.availableSlideCount || slot.renderedSlide === slideIndex) continue;
        void this._renderSlot(slideIndex, slot,
          !!mediaRange && this._rangeContains(mediaRange, slideIndex));
      }
    },
    emitVisible: () => {
      if (this._scroller.lastRange) this._emitVisibleSlideChange(this._scroller.lastRange);
    },
  });

  private readonly _renderHooks = {
    slots: () => this._slots,
    epoch: () => this._renderEpoch,
    scale: () => this._scale,
    slotIndex: (slot) => slot.renderedSlide,
    token: (slot) => slot.renderGeneration,
    nextToken: (slot) => ++slot.renderGeneration,
    wantRuns: (slot) => !!(this._opts.enableTextSelection && slot.textLayer) || this._findActive,
    reportError: (error) => this._reportRenderError(error),
  } satisfies Pick<BitmapSlotHooks<SlideSlot, PptxTextRunInfo>,
    'slots' | 'epoch' | 'scale' | 'slotIndex' | 'token' | 'nextToken' | 'wantRuns' | 'reportError'>;
  private readonly _bitmap = new BitmapSlotRenderer<SlideSlot, PptxTextRunInfo>({
    ...this._renderHooks,
    inFlight: () => this._scroller.inFlight,
    destroyed: () => this._destroyed,
    width: () => this._slideWidthPx(),
    dpr: () => this._viewport.dpr(),
    canRetry: (slot) => !(this._opts.enableMediaPlayback && slot.mediaInteractive),
    render: (slide, canvas, width, dpr, onTextRun) =>
      renderPptxFocusedSlide(this._pres!, canvas, slide, 'worker', {
        width, dpr, imageResources: this._opts.imageResources, onTextRun,
        ...(this._opts.errorPlaceholder === undefined ? {} : { errorPlaceholder: this._opts.errorPlaceholder }),
      }),
    commitBitmap: (_slide, _slot, dispatcher, generation, bitmap, width) => {
      const size = { cssWidth: Math.round(width), cssHeight: Math.round(this._slideHeightPx()) };
      return this._opts.enableMediaPlayback
        ? dispatcher.commitBitmapTo2d(generation, bitmap, size)
        : dispatcher.commitBitmap(generation, bitmap, size);
    },
    commitRuns: (slide, slot, runs, width, wantedRuns) =>
      this._commitRenderedRuns(slide, slot, runs, width, wantedRuns, true),
  });
  private readonly _main = new MainSlotRenderer<SlideSlot, PptxTextRunInfo>({
    ...this._renderHooks,
    render: (slide, canvas, width, dpr, onTextRun, settled) =>
      renderPptxFocusedSlide(this._pres!, canvas, slide, 'main', {
        width, dpr,
        ...(settled ? {} : { imageResources: this._opts.imageResources }),
        onTextRun,
        ...(this._opts.errorPlaceholder === undefined ? {} : { errorPlaceholder: this._opts.errorPlaceholder }),
      }),
    commitRuns: (slide, slot, runs, _canvas, width, wantedRuns, settled) =>
      this._commitRenderedRuns(slide, slot, runs, width, wantedRuns, settled),
    shadow: () => this._pageShadow,
  });
  private readonly _media = new PptxScrollMediaController<SlideSlot>({
    presentation: () => this._pres,
    slots: () => this._slots,
    epoch: () => this._renderEpoch,
    scale: () => this._scale,
    dpr: () => this._viewport.dpr(),
    width: () => this._slideWidthPx(),
    height: () => this._slideHeightPx(),
    imageResources: () => this._opts.imageResources,
    errorPlaceholder: () => this._opts.errorPlaceholder,
    textSelection: () => this._opts.enableTextSelection === true,
    findActive: () => this._findActive,
    mediaEnabled: () => this._opts.enableMediaPlayback === true,
    mediaRange: () => this._mediaRange(),
    rangeContains: (range, slide) => this._rangeContains(range, slide),
    shadow: () => this._pageShadow,
    hyperlinkHandler: () => this._hyperlinkHandler(),
    refreshFindRuns: (slide, runs) => this._highlights.refreshRuns(slide, runs),
    redrawHighlights: (slide, slot) => this._highlights.redrawSlot(slide, slot),
    commitComments: (slide, slot) => this._commitSlotComments(slide, slot),
    clearTextPreview: (layer) => this._clearTextLayerPreview(layer),
    reportError: (error) => this._reportRenderError(error),
  });
  private readonly _commentNavigation = new PptxScrollCommentNavigation<SlideSlot>({
    presentation: () => this._pres,
    destroyed: () => this._destroyed,
    slots: () => this._slots,
    scale: () => this._scale,
    scrollPoint: (slide, x, y, options) => this._navigation.scrollToPoint(slide, x, y, options),
    scrollToSlide: (slide, options) => this.scrollToSlide(slide, options),
    select: (commentId, slide) => {
      this._activeCommentId = commentId;
      this._activeCommentSlide = slide;
      this._selection.clearElementContext();
      for (const [mountedSlide, slot] of this._slots) this._redrawSlotComments(mountedSlide, slot);
      this._selection.emitChange();
    },
    ownBackground: (operation) => this._errorRouter.ownBackgroundLifecycle(operation),
  });
  private readonly _navigation = new ScrollNavigationController({
    host: () => this._scrollHost,
    spacer: () => this._spacer,
    count: () => this.slideCount,
    established: () => this._scaleEstablished,
    offset: (slide) => this._slideOffset(slide),
    height: () => this._uniformSlideHeight,
    indexAt: (y) => this._slideIndexAtOffset(y),
    totalHeight: () => this._rangeAt(0, this._viewport.overscan()).totalHeight,
    width: () => this._slideWidthPx(),
    padLeft: () => this._viewport.horizontalPadding().left,
    marginOrigin: () => this._commentMargin.originPx,
    mount: () => this._mountVisible(),
  });
  private readonly _selection = new SelectionContextController<PptxSelectionContext, PptxElementContext, PptxPresentation, SlideSlot>({
    wrapper: () => this._wrapper,
    scrollHost: () => this._scrollHost,
    slots: () => this._slots,
    resource: () => this._pres,
    destroyed: () => this._destroyed,
    textSelectionEnabled: () => this._opts.enableTextSelection === true,
    elementSelectionEnabled: () => this._opts.enableElementSelection === true,
    textSelected: () => readPptxTextSelectionContext(
      this._wrapper, this._wrapper.ownerDocument?.getSelection?.() ?? null,
    ) !== null,
    getContext: () => this.getSelectionContext(),
    hitTest: (presentation, slideIndex, xRatio, yRatio, canvasWidth) =>
      presentation.getElementContextAt(slideIndex, {
        x: xRatio * presentation.slideWidth,
        y: yRatio * presentation.slideHeight,
      }, {
        tolerance: this._elementHitTolerance / canvasWidth * presentation.slideWidth,
        maxTextCharacters: MAX_ELEMENT_TEXT_CHARACTERS,
      }),
    outline: (presentation, slideIndex, context) => {
      if (context.slideIndex !== slideIndex) return null;
      return {
        x: context.bounds.x / presentation.slideWidth,
        y: context.bounds.y / presentation.slideHeight,
        width: context.bounds.width / presentation.slideWidth,
        height: context.bounds.height / presentation.slideHeight,
        rotation: context.bounds.rotation,
      };
    },
    onChange: (context) => this._opts.onSelectionContextChange?.(context),
    onContextMenu: (event, getContext) => this._opts.onContextMenu?.({ originalEvent: event, getContext }),
    reportError: (error) => this._reportRenderError(error),
  });
  /** Uniform slide height at the current scale. Keeping the scalar avoids both
   * the document-length height and offset arrays in every scroll query. */
  private _uniformSlideHeight = 0;
  private readonly _visibleEvents = new VisibleUnitEvents(
    (index, total, complete) => this._opts.onVisibleSlideChange?.(index, total, complete),
  );
  private readonly _loader = new ScrollLoadController<PptxPresentation>({
    name: () => 'PptxScrollViewer',
    borrowed: () => this._borrowed,
    borrowedMessage: () => 'PptxScrollViewer.load() is unsupported on a Viewer created by fromPresentation(); ' +
      'the borrowed presentation is already loaded.',
    destroyed: () => this._destroyed,
    owner: () => this._presentationOwner,
    acquire: (source) => PptxPresentation.load(source, {
      password: this._opts.password,
      useGoogleFonts: this._opts.useGoogleFonts,
      fonts: this._opts.fonts,
      cjkFallback: this._opts.cjkFallback,
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
      ...(this._opts.legacyPpt === undefined ? undefined : { legacyPpt: this._opts.legacyPpt }),
      progressiveLayout: this._opts.progressiveLayout,
      onLayoutProgress: this._opts.onLayoutProgress,
      onLayoutPartial: this._opts.onLayoutPartial,
      onLayoutComplete: this._opts.onLayoutComplete,
    }),
    beforeReplace: (previous) => {
      this._selection.invalidateElementContext(false);
      this._invalidateFind();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentSlide = null;
      this._hasComments = false;
      this._commentScanFrontier = 0;
      this._commentNavigation.begin();
      this._partFailureDelivery.attach(null);
      this._layout.unbind();
      if (previous) {
        for (const [index, slot] of [...this._slots]) this._recycleSlot(index, slot);
        this._visibleEvents.resetIndex();
      }
    },
    afterReplace: (pres) => {
      this._invalidateFind();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentSlide = null;
      this._hasComments = false;
      this._commentScanFrontier = 0;
      this._scanAvailableComments(pres, false);
      this._layout.bind(pres);
      this._partFailureDelivery.attach(pptxPartFailureLog(pres));
    },
    mountOpeningWindow: async () => {
      const initialRenders: Promise<void>[] = [];
      this._relayout(initialRenders);
      await Promise.all(initialRenders);
    },
    selectionChanged: () => this._selection.emitChange(),
  });
  private _activeCommentId: string | null = null;
  private _activeCommentSlide: number | null = null;
  private _commentUi: PptxCommentUiRuntime | null = null;
  private _hasComments = false;
  private readonly _commentMargin = new CommentMarginController({
    container: () => this._container,
    scrollHost: () => this._scrollHost,
    spacer: () => this._spacer,
    enabled: () => this._commentsEnabled(),
    cards: () => this._commentsOptions()?.cards !== false,
    hasDisplayableComments: () => this._hasComments,
    requestedSide: () => this._commentsOptions()?.side,
    zoom: () => this._scaleEstablished ? this._scale : 1,
    gapPx: COMMENT_MARGIN_GAP_PX,
    widthPx: COMMENT_MARGIN_WIDTH_PX,
    fontSizePx: COMMENT_MARGIN_FONT_SIZE_PX,
  });
  private readonly _layers = new SlotLayerController<SlideSlot>({
    markerLayer: (slot) => slot.commentMarkerLayer,
    syncMargin: (margin) => this._commentMargin.syncMargin(margin),
    marginSide: () => this._commentMargin.side(),
    marginExtent: () => this._commentMargin.extent(),
    commentsEnabled: () => this._commentsEnabled(),
    previewMargin: (margin, ratio) => this._commentUi?.previewReadOnlyCommentMargin(margin, ratio),
    disposeMargin: (margin) => {
      this._commentUi?.disposeReadOnlyCommentMargin(margin);
      if (!this._commentUi) margin.replaceChildren();
    },
    disposeDecoration: (layer) => {
      this._commentUi?.disposeReadOnlyCommentDecoration(layer);
      if (!this._commentUi) layer.replaceChildren();
    },
    redrawOutline: (unit, slot) => this._selection.redrawOutlineForSlot(unit, slot),
    markerRatio: (ratio) => Math.round(ratio * 1_000_000) / 1_000_000,
    resetMarkerTransform: () => false,
    resetDecorationVisibility: () => true,
  });
  private readonly _commentOverlay = new CommentOverlayController<SlideSlot>({
    slots: () => this._slots,
    scale: () => this._scale,
    destroyed: () => this._destroyed,
    ownerWindow: () => this._wrapper.ownerDocument.defaultView,
    width: () => this._slideWidthPx(),
    height: () => this._slideHeightPx(),
    side: () => this._commentMargin.side(),
    marginExtent: () => this._commentMargin.extent(),
    connectorOptions: () => this._commentsOptions()?.connectors,
    runtime: () => this._commentUi,
    redrawComments: (slide, slot) => this._redrawSlotComments(slide, slot),
  });
  /** Opening prefix already inspected for authored comments. Progressive
   * presentations make metadata authoritative one slide at a time, so a
   * negative scan is provisional until this frontier reaches slideCount. */
  private _commentScanFrontier = 0;
  private readonly _elementHitTolerance: number;
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  private readonly _partFailureDelivery: PartFailureDelivery<PptxPartFailure>;
  /** Render generation, bumped on every effective `setScale` (and the resize
   *  re-fit in the zoom controller, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private get _renderEpoch(): number { return this._scroller.renderEpoch; }
  private set _renderEpoch(value: number) { this._scroller.renderEpoch = value; }
  /** Resolved once so explicit `false` disables the shared slot/spare shadow. */
  private readonly _pageShadow: string | false;
  private readonly _find = new PptxFindController(
    () => this.slideCount,
    (slide) => this._collectSlideRuns(slide),
  );
  private _findGeneration = 0;
  private _findActive = false;
  private readonly _highlights = new HighlightLayerController<SlideSlot, PptxTextRunInfo, PptxMatchLocation>({
    slots: () => this._slots,
    active: () => this._findActive,
    runs: (unit) => this._find.slideRuns(unit),
    setRuns: (unit, runs) => this._find.setSlideRuns(unit, runs),
    paint: (unit, slot, runs, measure) => buildPptxHighlightLayer(
      slot.highlightLayer, runs, this._find.slideHighlights(unit),
      this._slideWidthPx(), this._slideHeightPx(),
      measure, this._opts.findHighlightColors,
    ),
    reveal: (unit) => this.scrollToSlide(unit),
    unitOf: (location) => location.slide,
  });

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
    prepareErrorPlaceholder(opts.errorPlaceholder);
    this._partFailureDelivery = new PartFailureDelivery('PptxScrollViewer', opts.onPartFailure, (failure) => failure.slideIndex);
    const elementHitTolerance = opts.elementHitTolerance ?? 6;
    if (!Number.isFinite(elementHitTolerance) || elementHitTolerance < 0) {
      throw new RangeError('elementHitTolerance must be a finite non-negative number.');
    }
    this._elementHitTolerance = elementHitTolerance;
    // `??` (not `||`): a caller's explicit `false` must disable the shadow, not
    // fall through to the default.
    this._pageShadow = opts.pageShadow ?? DEFAULT_SCROLL_PAGE_SHADOW;
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

    this._shell = new ScrollViewerShell(container, {
      background: opts.background,
      comments: !!opts.comments,
      onScroll: () => this._onScroll(),
      onOutsideComment: () => {
        if (this._activeCommentId === null) return;
        this._activeCommentId = null;
        this._activeCommentSlide = null;
        for (const [index, slot] of this._slots) this._redrawSlotComments(index, slot);
        this._selection.emitChange();
      },
    });

    if (this._commentsEnabled()) {
      void loadPptxCommentUiRuntime().then((commentUi) => {
        if (this._destroyed) return;
        this._commentUi = commentUi;
        for (const [slide, slot] of this._slots) this._redrawSlotComments(slide, slot);
      }).catch((error) => this._reportRenderError(error));
    }

    this._selection.bind(!!opts.onSelectionContextChange, !!opts.onContextMenu);

    this._zoom.bind(this._container, this._scrollHost, this._opts.enableZoom !== false);

    if (this._borrowed) {
      this._layout.bind(borrowedPresentation!);
      this._partFailureDelivery.attach(pptxPartFailureLog(borrowedPresentation!));
      // A borrowed engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path re-runs it once width appears).
      this.relayout();
    }
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    // #1974: a reload stops delivering the committed engine's part failures at
    // once; the engine committed when it settles is delivered again.
    const failureLoad = this._partFailureDelivery.beginLoad();
    try {
      await this._loader.load(source);
    } finally {
      const committed = this._destroyed ? null : this._pres;
      this._partFailureDelivery.endLoad(failureLoad, committed ? pptxPartFailureLog(committed) : null);
    }
  }

  get slideCount(): number {
    return this._pres?.slideCount ?? 0;
  }

  /** Slides of the loaded presentation whose part failed to parse (#1974);
   *  `[]` before load and after destroy. */
  get partFailures(): readonly PptxPartFailure[] {
    return this._destroyed ? [] : this._pres?.partFailures ?? [];
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

  private _commentsEnabled(): boolean {
    return this._opts.comments === true || typeof this._opts.comments === 'object';
  }

  private _commentsOptions(): PptxCommentsOptions | undefined {
    return typeof this._opts.comments === 'object' ? this._opts.comments : undefined;
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
    const w = this._viewport.fitWidth();
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
    if (!this._scaleEstablished) {
      if (!this._zoom.establishBase()) return;
    }
    this._recomputeHeights();
    this._syncSpacer();
    this._mountVisible(initialRenders);
  }

  /** Refresh the uniform scale-dependent height without allocating per-slide state. */
  private _recomputeHeights(): void {
    this._uniformSlideHeight = this._slideHeightPx();
  }

  /** Media lifecycle window, deliberately independent from text/canvas overscan. */
  private _mediaOverscan(): number {
    return this._opts.mediaOverscan ?? 1;
  }

  private _slideOffset(index: number): number {
    return this._viewport.verticalPadding().leading + index * (this._uniformSlideHeight + this._viewport.gap());
  }

  /** Index of the slide spanning content-offset `y`, preserving the historical
   * convention that an inter-slide gap belongs to the preceding slide. */
  private _slideIndexAtOffset(y: number): number {
    return computeUniformVisibleWindow(
      this._pres?.slideCount ?? 0,
      this._uniformSlideHeight,
      this._viewport.gap(),
      y,
      0,
      0,
      this._viewport.verticalPadding(),
    ).topIndex;
  }

  private _rangeAt(scrollTop: number, overscan: number): VisibleWindow {
    return computeUniformVisibleWindow(
      this._pres?.slideCount ?? 0,
      this._uniformSlideHeight,
      this._viewport.gap(),
      scrollTop,
      this._scrollHost.clientHeight,
      overscan,
      this._viewport.verticalPadding(),
    );
  }

  private _range(): VisibleWindow {
    return this._rangeAt(this._scrollHost.scrollTop, this._viewport.overscan());
  }

  private _mediaRange(): VisibleWindow {
    return this._rangeAt(this._scrollHost.scrollTop, this._mediaOverscan());
  }

  private _rangeContains(r: VisibleWindow, index: number): boolean {
    return index >= r.start && index <= r.end;
  }

  private _syncSpacer(): void { this._scroller.syncSpacer(); }

  /** Horizontal scroll extent: the (uniform deck-wide) slide width plus both
   *  gutters. A spacer NARROWER than the container never creates a scrollbar
   *  (scrollWidth = max(clientWidth, content)), so it is always safe to set — it
   *  only matters when a zoomed-in slide grows past the viewport, where it gives
   *  the gutters something to scroll to on either side. Called from `_syncSpacer`
   *  and after every scale change (zoom / resize re-fit) so the extent tracks the
   *  current slide px width. */
  private _syncSpacerWidth(): void {
    const { left, right } = this._viewport.horizontalPadding();
    this._commentMargin.syncSpacerWidth(this._slideWidthPx(), left, right);
  }

  private _onScroll(): void {
    if (!this._pres || !this._scaleEstablished) return;
    this._mountVisible(undefined, false);
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(initialRenders?: Promise<void>[], repositionExisting = true): void {
    this._scroller.mount(initialRenders, repositionExisting);
  }

  private _createSlot(): SlideSlot {
    // The common canvas, selection and highlight stack is owned by core.
    const { wrapper, canvas, textLayer, highlightLayer } = createSlotHost(
      this._scrollHost, this._opts.enableTextSelection === true, this._pageShadow,
    );
    const loadingIndicator = createPptxLoadingIndicator(wrapper);
    const { markerLayer: commentMarkerLayer, margin: commentMargin, decorationLayer: commentDecorationLayer } =
      createCommentSlotLayers(
        wrapper,
        this._commentsEnabled(),
        this._commentsOptions()?.cards !== false,
        this._commentsOptions()?.connectors !== undefined,
        (margin) => this._commentMargin.syncMargin(margin),
      );
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
      loadingIndicator,
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
    this._scroller.recycleSlot(idx, slot);
  }

  private _resetSlot(_idx: number, slot: SlideSlot): void {
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
    this._layers.reset(slot);
    slot.loadingIndicator.setLoading(false);
    slot.commentElementBounds = Object.freeze([]);
    slot.commentGeometry = null;
    slot.commentAnchorSlide = -1;
    slot.commentAnchorGeneration++;
    // `_previewSlot` pins an explicit CSS height while stretching the current
    // bitmap during a zoom burst. A slot can leave the visible range before the
    // debounced settle replaces that canvas, so do not carry the old-scale height
    // into its next slide. Main-mode renderSlide intentionally updates only the
    // CSS width and lets height follow the backing-store aspect ratio.
    slot.canvas.style.height = '';
    slot.renderedSlide = -1;
    slot.renderedScale = -1;
    slot.wrapper.remove();
  }

  private _positionSlot(slot: SlideSlot, i: number, _r: VisibleWindow): void {
    slot.wrapper.dataset.slideIndex = String(i);
    this._layers.position(i, slot, this._slideOffset(i), this._slideWidthPx(), this._slideHeightPx(),
      this._scrollHost.clientWidth, this._viewport.horizontalPadding().left);
  }

  /** Static slot dispatch is owned by core; media remains a PPTX hook. */
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
      slot.loadingIndicator.setLoading(true);
      return null;
    }
    slot.renderedSlide = i;
    const renderGeneration = ++slot.renderGeneration;
    slot.loadingIndicator.setLoading(i >= this.availableSlideCount && !this.layoutComplete);

    const dpr = this._viewport.dpr();
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
        this._media.renderInteractive(i, slot, widthPx, dpr, scale, epoch, reportErrors),
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

    return this._trackSlotLoading(i, slot, renderGeneration,
      this._main.render(i, slot, widthPx, dpr, renderGeneration, dispatcher, generation, reportErrors));
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
        slot.loadingIndicator.setLoading(false);
      }
    }
  }

  private _emitVisibleSlideChange(range: VisibleWindow): void {
    if (this._pres) this._visibleEvents.publish(range, this._pres.slideCount, this.layoutComplete);
  }

  /** Route an async render failure to `onError`, or `console.error` when none is
   *  set (so failures are never fully silent), and never after teardown. */
  private _reportRenderError(err: unknown): void {
    this._errorRouter.report(err);
  }

  private _commitRenderedRuns(
    slide: number, slot: SlideSlot, runs: PptxTextRunInfo[], width: number,
    wantedRuns: boolean, clearPreview: boolean,
  ): void {
    if (slot.textLayer) {
      if (clearPreview) this._clearTextLayerPreview(slot.textLayer);
      if (this._opts.enableTextSelection) {
        buildPptxTextLayer(slot.textLayer, runs,
          Math.round(width), Math.round(this._slideHeightPx()),
          this._hyperlinkHandler(), slide);
      }
    }
    if (wantedRuns) this._highlights.refreshRuns(slide, runs);
    this._commitSlotComments(slide, slot);
    this._highlights.redrawSlot(slide, slot);
  }

  private _renderSlotBitmap(
    i: number, slot: SlideSlot, widthPx: number, dpr: number, scale: number,
    renderGeneration = ++slot.renderGeneration,
    dispatcher = slot.dispatcher, generation = dispatcher.begin(), reportErrors = true,
  ): Promise<void> {
    return this._bitmap.render(
      i, slot, widthPx, dpr, scale, renderGeneration, dispatcher, generation, reportErrors,
    );
  }

  /** Keep the public zoom facade while core owns scale, fit and anchoring. */
  setScale(scale: number): void { this._zoom.setScale(scale); }
  getScale(): number { return this._zoom.getScale(); }
  zoomIn(): void { this._zoom.zoomIn(); }
  zoomOut(): void { this._zoom.zoomOut(); }
  fitWidth(): void { this._zoom.fit('width'); }
  fitPage(): void { this._zoom.fit('page'); }

  private _previewVisible(): void { this._scroller.preview(); }

  private _previewSlot(slot: SlideSlot, i: number, r: VisibleWindow): void {
    this._positionSlot(slot, i, r);
    this._layers.preview(slot, this._slideWidthPx(), this._slideHeightPx(), this._scale);
  }

  private _clearTextLayerPreview(layer: HTMLDivElement): void {
    clearTextLayerPreview(layer);
  }

  private _scheduleSettle(): void { this._scroller.scheduleSettle(DEFAULT_ZOOM_SETTLE_MS); }

  private _settleSlot(i: number, slot: SlideSlot): void {
    if (!this._pres) return;
    const dpr = this._viewport.dpr();
    const widthPx = this._slideWidthPx();
    const scale = this._scale;
    const epoch = this._renderEpoch;

    if (this._opts.enableMediaPlayback && slot.mediaInteractive) {
      this._media.settleInteractive(i, slot, widthPx, dpr, scale, epoch);
      return;
    }
    if (this._opts.enableMediaPlayback) return;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
      return;
    }

    this._main.settle(i, slot, widthPx, dpr);
  }

  scrollToSlide(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    this._navigation.scrollToUnit(index, opts);
  }

  /** Reveal an authored occurrence through the PPTX comment navigation adapter. */
  async goToComment(
    slideIndex: number, commentIndex: number,
    opts?: { behavior?: 'auto' | 'smooth' },
  ): Promise<boolean> {
    return this._commentNavigation.goToComment(slideIndex, commentIndex, opts);
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
      // A progressive deck must clear existing highlights without waiting for
      // slides that have not finished preparing. DOCX retains its layout wait.
      this._find.invalidate();
      this._highlights.redrawAll();
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
    this._highlights.redrawAll();
    return matches;
  }

  /** Activate and reveal the next match, wrapping at the end. */
  async findNext(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._highlights.activate(this._find.next());
  }

  /** Activate and reveal the previous match, wrapping at the beginning. */
  async findPrev(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._highlights.activate(this._find.prev());
  }

  /** Clear the current query and every mounted highlight. */
  clearFind(): void {
    this._findActive = false;
    this._invalidateFind();
    this._highlights.redrawAll();
  }

  private _invalidateFind(): void {
    this._findGeneration++;
    this._find.invalidate();
  }

  private async _collectSlideRuns(slide: number): Promise<PptxTextRunInfo[]> {
    if (!this._pres) return [];
    return this._pres.collectSlideRuns(slide, this._slideWidthPx());
  }

  private _redrawSlotComments(slide: number, slot: SlideSlot): void {
    if (!this._pres || !slot.commentMarkerLayer) return;
    this._commentMargin.syncMargin(slot.commentMargin);
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
        this._selection.clearElementContext();
        for (const [mountedSlide, mountedSlot] of this._slots) {
          this._redrawSlotComments(mountedSlide, mountedSlot);
        }
        this._selection.emitChange();
      },
      this._commentMargin.zoom(),
      COMMENT_MARGIN_WIDTH_PX,
      this._commentsOptions()?.markers !== false,
      this._commentsOptions()?.includeResolved === true,
      slot.commentDecorationLayer
        ? () => this._commentOverlay.schedule(slide, slot, false)
        : undefined,
      slot.commentDecorationLayer
        ? () => this._commentOverlay.schedule(slide, slot, true)
        : undefined,
    );
    this._commentOverlay.drawConnectors(slide, slot);
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
      if (active) this._commentNavigation.scrollToTarget(slide, active);
    }).catch((error: unknown) => {
      if (!this._destroyed && generation === slot.commentAnchorGeneration) {
        this._reportRenderError(error);
      }
    });
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

  get topVisibleSlide(): number {
    return this._scroller.lastRange?.topIndex ?? 0;
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
    return text ?? (this._selection.elementContext
      ? limitPptxElementContext(
          this._selection.elementContext,
          options.maxTextCharacters,
        )
      : null);
  }

  /**
   * Tear down the viewer: remove the DOM subtree and (only for a self-loaded
   * engine) destroy the engine. A borrowed engine is left intact — the caller
   * owns its lifecycle. Per-slot worker ImageBitmaps are closed on recycle.
   */
  destroy(): void {
    if (this._destroyed) return;
    this._destroyed = true;
    this._commentNavigation.begin();
    this._errorRouter.close();
    this._partFailureDelivery.close();
    this._invalidateFind();
    this._findActive = false;
    this._layout.unbind();
    this._selection.destroy();
    this._media.destroy();
    this._commentNavigation.destroy();
    this._highlights.destroy();
    this._commentOverlay.destroy();
    this._selection.clearElementContext();
    // Cancel a pending settle so no re-render is dispatched after teardown
    // (design §7 mechanism 2). Clearing the timer avoids a wasted wake-up and
    // keeps fake-timer tests deterministic.
    this._scroller.destroy();
    this._zoom.destroy();
    this._commentMargin.destroy();
    this._presentationOwner.close();
    this._shell.destroy();
  }
}
