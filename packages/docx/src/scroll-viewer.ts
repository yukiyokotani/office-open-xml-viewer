import { updateNativeReadingNotice, noReadingNotices } from './native-reading-notice.js';
import { openExternalHyperlink, PT_TO_PX } from '@silurus/ooxml-core';
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
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import { READ_ONLY_COMMENT_MARGIN_WIDTH_PX } from '@silurus/ooxml-core/internal/read-only-comment-contract';
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
import { DocxDocument, docxViewerLoadSignal } from './document';
import type { DocxViewerLoadControl, LoadOptions } from './document';
import {
  activeDocxLayoutViewOf,
  reconcilePendingDocxLayoutView,
  selectDocxLayoutView,
} from './document-layout-view.js';
import type { DocxTextRunInfo } from './renderer';
import type { DocxScrollSlot as PageSlot } from './scroll-slot';
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
import type { DocxScrollViewerOptions } from './scroll-viewer-options';
import { DocxScrollCommentNavigation } from './scroll-comment-navigation';
import { renderDocxFocusedPage } from './focused-view-runtime';
import type { DocxLayoutPublication } from './document-layout-events.js';
import { DocxScrollLayoutController } from './scroll-layout-controller';

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

export type { DocxScrollViewerOptions } from './scroll-viewer-options';

export class DocxScrollViewer implements ZoomableViewer {
  private _pendingLoadAbort: AbortController | null = null;
  private _pendingRequestedView: boolean | undefined;
  private _pendingViewChanged: (() => void) | null = null;
  private readonly _documentOwner: TerminalResourceOwner<DocxDocument>;
  private get _doc(): DocxDocument | null { return this._documentOwner.current; }
  private readonly _borrowed: boolean;
  private readonly _opts: DocxScrollViewerOptions;
  private readonly _errorRouter: CanvasViewerErrorRouter;
  private readonly _container: HTMLElement;
  private readonly _shell: ScrollViewerShell;
  private get _wrapper(): HTMLDivElement { return this._shell.wrapper; }
  private get _scrollHost(): HTMLDivElement { return this._shell.scrollHost; }
  private get _spacer(): HTMLDivElement { return this._shell.spacer; }
  /** Resolved render mode. When an engine is borrowed the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `DocxDocument.load`. */
  private _mode: 'main' | 'worker';

  private readonly _viewport = new ScrollViewportPolicy({
    options: () => this._opts,
    container: () => this._container,
    scrollHost: () => this._scrollHost,
  });
  private readonly _zoom = new ScrollZoomController({
    scrollHost: () => this._scrollHost,
    spacer: () => this._spacer,
    count: () => this._doc?.pageCount ?? 0,
    zoomMin: () => this._opts.zoomMin ?? 0.1,
    zoomMax: () => this._opts.zoomMax ?? 4,
    baseScale: () => this._baseScale(),
    fitWidthPx: () => this._viewport.fitWidth(),
    fitContentSize: (mode) => {
      if (!this._doc) return null;
      const size = this._doc.pageSize(0);
      return {
        width: (mode === 'width' ? this._widestPageWidthPt() : size.widthPt) * PT_TO_PX,
        height: size.heightPt * PT_TO_PX,
      };
    },
    indexAt: (y) => this._pageIndexAtOffset(this._range(), y),
    offset: (index) => this._scrollGeometry.offsets[index] ?? 0,
    height: (index) => this._heights[index] || 0,
    totalHeight: () => this._scrollGeometry.totalHeight,
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
  private readonly _scroller = new SlotScroller<PageSlot, VisibleRange>({
    spacer: () => this._spacer,
    count: () => this._doc?.pageCount ?? 0,
    range: () => this._range(),
    createSlot: () => this._createSlot(),
    attachSlot: (slot) => this._scrollHost.appendChild(slot.wrapper),
    resetSlot: (index, slot) => this._resetSlot(index, slot),
    positionSlot: (index, slot, range) => this._positionSlot(slot, index, range),
    renderSlot: (index, slot, reportErrors) => this._renderSlot(index, slot, reportErrors),
    previewSlot: (index, slot, range) => this._previewSlot(slot, index, range),
    settleSlot: (index, slot) => this._refreshSlotAtomically(index, slot),
    renderedScale: (slot) => slot.renderedScale,
    scale: () => this._scale,
    syncSpacerWidth: () => this._syncSpacerWidth(),
    onRange: (range) => this._emitVisiblePageChange(range),
    // An authoritative DOCX layout publication may replace an already mounted
    // page without changing its index. _renderSlot checks whether it is current.
    onExistingSlot: (index, slot, reportErrors) => this._renderSlot(index, slot, reportErrors),
  });
  private readonly _slots = this._scroller.slots;
  private readonly _layout = new DocxScrollLayoutController({
    current: () => this._doc,
    destroyed: () => this._destroyed,
    report: (error) => this._reportRenderError(error),
    reportBackground: (error) => {
      if (this._doc?.pageCount === 0 && this._doc._readingPublicationOwned) this._clearReadingPages();
      this._errorRouter.reportBackground(error, this._opts.onLayoutComplete !== undefined);
    },
    invalidateFind: () => this._find.invalidate(),
    refreshComments: () => this._refreshCommentSurface(),
    adoptView: (publication) => {
      if (publication.requester === this) return;
      this._layoutViewGeneration++;
      this._showTrackedChanges = publication.view.showTrackedChanges;
      this._currentDate = publication.view.currentDate;
      this._find.invalidate();
      this._syncReadingNotice();
      this._layout.apply({
        pageCount: this._doc!.pageCount, exact: true, complete: this._doc!.layoutComplete,
      });
    },
    relayout: () => this.relayout(),
    invalidateRender: () => { this._renderEpoch++; },
    mounted: () => this._slots,
    stillMounted: (page, slot) => this._slots.get(page) === slot,
    refreshSlot: (page, slot) => this._refreshSlotAtomically(page, slot as PageSlot),
  });

  private readonly _renderHooks = {
    slots: () => this._slots,
    epoch: () => this._renderEpoch,
    scale: () => this._scale,
    slotIndex: (slot) => slot.renderedPage,
    token: () => 0,
    nextToken: () => 0,
    wantRuns: (slot) => !!(this._opts.enableTextSelection && slot.textLayer) ||
      this._findActive || !!slot.commentTintLayer,
    reportError: (error) => this._reportRenderError(error),
  } satisfies Pick<BitmapSlotHooks<PageSlot, DocxTextRunInfo>,
    'slots' | 'epoch' | 'scale' | 'slotIndex' | 'token' | 'nextToken' | 'wantRuns' | 'reportError'>;
  private readonly _bitmap = new BitmapSlotRenderer<PageSlot, DocxTextRunInfo>({
    ...this._renderHooks,
    inFlight: () => this._scroller.inFlight,
    destroyed: () => this._destroyed,
    width: (page) => this._pageWidthPx(page),
    dpr: () => this._viewport.dpr(),
    canRetry: () => true,
    render: (page, canvas, width, dpr, onTextRun) =>
      renderDocxFocusedPage(this._doc!, canvas, page, 'worker', {
        width, dpr,
        imageResources: this._opts.imageResources,
        defaultTextColor: this._opts.defaultTextColor,
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        onTextRun,
      }),
    commitBitmap: (page, slot, dispatcher, generation, bitmap, width) => {
      const doc = this._doc, publication = doc?._readingPublicationToken;
      try { return dispatcher.commitBitmap(generation, bitmap, {
        cssWidth: width, cssHeight: this._pageHeightPx(page),
      }); } catch (error) {
        if (doc === this._doc && publication != null) doc?._invalidateReadingLayout(error, publication);
        throw error;
      }
    },
    commitRuns: (page, slot, runs, _width, wantedRuns) =>
      this._commitRenderedRuns(page, slot, runs, slot.canvas, wantedRuns, true),
  });
  private readonly _main = new MainSlotRenderer<PageSlot, DocxTextRunInfo>({
    ...this._renderHooks,
    render: (page, canvas, width, dpr, onTextRun) =>
      renderDocxFocusedPage(this._doc!, canvas, page, 'main', {
        width, dpr,
        imageResources: this._opts.imageResources,
        defaultTextColor: this._opts.defaultTextColor,
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        onTextRun,
      }),
    commitRuns: (page, slot, runs, canvas, _width, wantedRuns, settled) =>
      this._commitRenderedRuns(page, slot, runs, canvas, wantedRuns, settled),
    shadow: () => this._pageShadow,
  });
  private readonly _navigation = new ScrollNavigationController({
    host: () => this._scrollHost,
    spacer: () => this._spacer,
    count: () => this.pageCount,
    established: () => this._scaleEstablished,
    offset: (page) => this._scrollGeometry.offsets[page] ?? 0,
    height: (page) => this._heights[page] || 0,
    indexAt: (y) => this._pageIndexAtOffset(this._range(), y),
    totalHeight: () => this._scrollGeometry.totalHeight,
    width: (page) => this._pageWidthPx(page),
    padLeft: () => this._viewport.horizontalPadding().left,
    marginOrigin: () => this._commentMargin.originPx,
    mount: () => this._mountVisible(),
  });
  private readonly _selection = new SelectionContextController<DocxSelectionContext, DocxElementContext, DocxDocument, PageSlot>({
    wrapper: () => this._wrapper,
    scrollHost: () => this._scrollHost,
    slots: () => this._slots,
    resource: () => this._doc,
    destroyed: () => this._destroyed,
    textSelectionEnabled: () => this._opts.enableTextSelection === true,
    elementSelectionEnabled: () => this._opts.enableElementSelection === true,
    textSelected: () => readDocxTextSelectionContext(
      this._wrapper, this._wrapper.ownerDocument?.getSelection?.() ?? null,
    ) !== null,
    getContext: () => this.getSelectionContext(),
    hitTest: (doc, pageIndex, xRatio, yRatio) => {
      const size = doc.pageSize(pageIndex);
      return doc.getElementContextAt(pageIndex, {
        xPt: xRatio * size.widthPt, yPt: yRatio * size.heightPt,
      }, {
        currentDate: this._currentDate,
        ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
        maxTextCharacters: MAX_DOCX_ELEMENT_TEXT_CHARACTERS,
      });
    },
    outline: (doc, pageIndex, context) => {
      if (context.pageIndex !== pageIndex) return null;
      const size = doc.pageSize(pageIndex);
      return {
        x: context.bounds.xPt / size.widthPt,
        y: context.bounds.yPt / size.heightPt,
        width: context.bounds.widthPt / size.widthPt,
        height: context.bounds.heightPt / size.heightPt,
      };
    },
    onChange: (context) => this._opts.onSelectionContextChange?.(context),
    onContextMenu: (event, getContext) => this._opts.onContextMenu?.({ originalEvent: event, getContext }),
    reportError: (error) => this._reportRenderError(error),
  });
  /** Cached per-page heights in px at the current scale (index-aligned). */
  private _heights: number[] = [];
  /** Prefix offsets rebuilt only when scale/page geometry changes. Pure scroll
   * queries binary-search this cache instead of walking every document page. */
  private _scrollGeometry: VirtualScrollGeometry = { offsets: [], totalHeight: 0 };
  private readonly _visibleEvents = new VisibleUnitEvents(
    (index, total, complete) => this._opts.onVisiblePageChange?.(index, total, complete),
  );
  private readonly _loader = new ScrollLoadController<DocxDocument, AbortController>({
    name: () => 'DocxScrollViewer',
    borrowed: () => this._borrowed,
    borrowedMessage: () => 'DocxScrollViewer.load() is unsupported on a Viewer created by fromDocument(); ' +
      'the borrowed document is already loaded.',
    destroyed: () => this._destroyed,
    owner: () => this._documentOwner,
    beginLoad: () => {
      const inheritedRequestedView = this._pendingRequestedView;
      this._pendingLoadAbort?.abort();
      const loadAbort = new AbortController();
      this._pendingLoadAbort = loadAbort;
      this._pendingRequestedView = inheritedRequestedView;
      return loadAbort;
    },
    isCurrentLoad: (loadAbort) => this._pendingLoadAbort === loadAbort,
    finishLoad: (loadAbort) => {
      if (this._pendingLoadAbort !== loadAbort) return;
      this._pendingLoadAbort = null;
      this._pendingRequestedView = undefined;
      this._pendingViewChanged = null;
    },
    acquire: async (source, loadAbort) => {
      const inheritedRequestedView = this._pendingRequestedView;
      const loaded = await DocxDocument.load(source, {
        password: this._opts.password,
        useGoogleFonts: this._opts.useGoogleFonts,
        allowFootnoteContinuation: this._opts.allowFootnoteContinuation,
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
        // Match the first paint to the requested view even while sliced layout
        // is in progress. An explicit false overrides a model-source default.
        ...(inheritedRequestedView !== undefined
          ? { showTrackedChanges: inheritedRequestedView }
          : this._opts.modelSources === undefined
            ? (this._showTrackedChanges ? { showTrackedChanges: true } : {})
            : (this._requestedShowTrackedChanges === undefined
              ? undefined
              : { showTrackedChanges: this._requestedShowTrackedChanges })),
        ...(this._currentDate === undefined ? {} : { currentDate: this._currentDate }),
        ...(this._opts.modelSources === undefined ? undefined : { modelSources: this._opts.modelSources }),
        ...(this._opts.progressiveLayout ? { progressiveLayout: true } : {}),
        ...(this._opts.sliceLayout === undefined ? {} : { sliceLayout: this._opts.sliceLayout }),
        [docxViewerLoadSignal]: {
          signal: loadAbort.signal,
          requestedView: () => this._pendingRequestedView,
          subscribeViewChange: (listener: () => void) => {
            this._pendingViewChanged = listener;
            return () => {
              if (this._pendingViewChanged === listener) this._pendingViewChanged = null;
            };
          },
        } satisfies DocxViewerLoadControl,
        onLayoutProgress: this._opts.onLayoutProgress,
        onLayoutPartial: this._opts.onLayoutPartial,
        onLayoutComplete: this._opts.onLayoutComplete,
      } as LoadOptions);
      // Reconcile a view change that races the final sliced-layout probe
      // before the resource owner commits the candidate.
      await reconcilePendingDocxLayoutView(
        loaded, loadAbort.signal, () => this._pendingRequestedView,
        () => this._currentDate, this,
      );
      return loaded;
    },
    beforeReplace: (previous) => {
      this._readingNoticeElement = updateNativeReadingNotice(this._container, this._readingNoticeElement, noReadingNotices);
      this._selection.invalidateElementContext(false);
      this._findRequestGeneration++;
      this._find.invalidate();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentPage = null;
      this._commentNavigation.reset();
      this._layout.unbind();
      if (previous) {
        for (const [index, slot] of [...this._slots]) this._recycleSlot(index, slot);
        this._visibleEvents.reset();
      }
    },
    afterReplace: (doc) => {
      if (this._pendingRequestedView !== undefined) {
        this._showTrackedChanges = this._pendingRequestedView;
        this._requestedShowTrackedChanges = this._pendingRequestedView;
      }
      if (this._opts.modelSources !== undefined) {
        this._showTrackedChanges = activeDocxLayoutViewOf(doc).showTrackedChanges;
      }
      this._syncReadingNotice();
      this._layout.bind(doc);
      this._find.invalidate();
      this._findActive = false;
      this._activeCommentId = null;
      this._activeCommentPage = null;
      this._commentNavigation.reset();
    },
    mountOpeningWindow: async () => {
      const doc = this._doc;
      const readingOwned = doc?._readingPublicationOwned === true;
      const mount = async (): Promise<void> => {
        const initialRenders: Promise<void>[] = [];
        try {
          this._relayout(initialRenders);
          await Promise.all(initialRenders);
          // Slot recycling can suppress its obsolete dispatcher rejection.
          // The same live reading owner still retains the terminal cause.
          if (readingOwned && doc && !this._destroyed && doc === this._doc)
            await doc.waitUntilLayoutComplete();
        } catch (error) {
          if (readingOwned && (this._destroyed || doc !== this._doc)) return;
          throw error;
        }
      };
      if (readingOwned) await this._errorRouter.ownBackgroundLifecycle(mount);
      else await mount();
    },
    selectionChanged: () => this._selection.emitChange(),
  });
  private _layoutUnsubscribe: (() => void) | null = null;
  /** Page prefix currently represented by the native scroll extent. */
  private _presentedPageCount = 0;
  private _activeCommentId: string | null = null;
  private _activeCommentPage: number | null = null;
  private _commentUi: DocxCommentUiRuntime | null = null;
  /** Latest default internal-link navigation; later clicks supersede work that
   * is still waiting for the authoritative bookmark projection. */
  private _internalHyperlinkGeneration = 0;
  private _commentAnchorRangesForMargin: ReturnType<DocxDocument['commentAnchorRanges']> | null = null;
  private _commentAnchorIds: ReadonlySet<string> = new Set();
  private readonly _commentMargin = new CommentMarginController({
    container: () => this._container,
    scrollHost: () => this._scrollHost,
    spacer: () => this._spacer,
    enabled: () => this._commentsEnabled(),
    cards: () => this._commentsOptions()?.cards !== false,
    hasDisplayableComments: () => this._hasDisplayableComments(),
    requestedSide: () => this._commentsOptions()?.side,
    zoom: () => this._scaleEstablished ? this._scale : 1,
    gapPx: COMMENT_MARGIN_GAP_PX,
    widthPx: READ_ONLY_COMMENT_MARGIN_WIDTH_PX,
    fontSizePx: COMMENT_MARGIN_FONT_SIZE_PX,
  });
  private readonly _layers = new SlotLayerController<PageSlot>({
    markerLayer: (slot) => slot.commentTintLayer,
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
    markerRatio: (ratio) => ratio,
    resetMarkerTransform: () => true,
    resetDecorationVisibility: () => false,
  });
  private readonly _commentNavigation = new DocxScrollCommentNavigation({
    document: () => this._doc,
    destroyed: () => this._destroyed,
    scale: () => this._scale,
    pageWidth: (page) => this._pageWidthPx(page),
    currentDate: () => this._currentDate,
    showTrackedChanges: () => this._showTrackedChanges,
    waitForLayout: (doc) => this._errorRouter.ownBackgroundLifecycle(
      () => doc.waitUntilLayoutComplete()),
    select: (commentId, page, target, options) => {
      this._activeCommentId = commentId;
      this._activeCommentPage = page;
      this._selection.clearElementContext();
      this._scrollToPageTarget(page, target, options);
      for (const [mountedPage, slot] of this._slots) this._redrawSlotComments(mountedPage, slot);
      this._selection.emitChange();
    },
  });
  private readonly _commentOverlay = new CommentOverlayController<PageSlot>({
    slots: () => this._slots,
    scale: () => this._scale,
    destroyed: () => this._destroyed,
    ownerWindow: () => this._wrapper.ownerDocument.defaultView,
    width: (page) => this._pageWidthPx(page),
    height: (page) => this._pageHeightPx(page),
    side: () => this._commentMargin.side(),
    marginExtent: () => this._commentMargin.extent(),
    connectorOptions: () => this._commentsOptions()?.connectors,
    runtime: () => this._commentUi,
    redrawComments: (page, slot) => this._redrawSlotComments(page, slot),
  });
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  /** Throwaway 2D context reused to measure text for the §17.3.2.10 縦中横 overlay
   *  clamp (#836). Lazily created; `null` when canvas metrics are unavailable
   *  (headless), in which case the overlay degrades to the un-clamped span. */
  /** Render generation, bumped on every effective `setScale` (and the resize
   *  re-fit in `_onResize`, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private get _renderEpoch(): number { return this._scroller.renderEpoch; }
  private set _renderEpoch(value: number) { this._scroller.renderEpoch = value; }
  private get _prevBase(): number { return this._zoom.prevBase; }
  private set _prevBase(value: number) { this._zoom.prevBase = value; }
  /** Resolved once so explicit `false` disables the shared slot/spare shadow. */
  private readonly _pageShadow: string | false;
  private readonly _find = new DocxFindController(
    () => this.pageCount,
    (page) => this._collectPageRuns(page),
  );
  private _findActive = false;
  private readonly _highlights = new HighlightLayerController<PageSlot, DocxTextRunInfo, DocxMatchLocation>({
    slots: () => this._slots,
    active: () => this._findActive,
    runs: (unit) => this._find.pageRuns(unit),
    setRuns: (unit, runs) => this._find.setPageRuns(unit, runs),
    paint: (unit, slot, runs, measure) => buildDocxHighlightLayer(
      slot.highlightLayer, runs, this._find.pageHighlights(unit),
      this._pageWidthPx(unit), this._pageHeightPx(unit),
      measure, this._opts.findHighlightColors,
    ),
    reveal: (unit) => this.scrollToPage(unit),
    unitOf: (location) => location.page,
  });
  /** Covers the pre-search progressive wait before DocxFindController.find()
   * can establish its own cancellation generation. */
  private _findRequestGeneration = 0;
  /** ECMA-376 §17.13.5 — current tracked-change view. A LAYOUT axis (deletions
   *  change line breaking and pagination), so it selects which retained layout
   *  variant the viewer reads geometry from; toggle it with
   *  {@link setShowTrackedChanges}. */
  private _showTrackedChanges: boolean;
  /** The tracked-change view requested by the caller or by
   *  {@link setShowTrackedChanges}; `undefined` lets each loaded document's
   *  own view default apply. Sent tri-state on every load. */
  declare private _requestedShowTrackedChanges: boolean | undefined;
  /** Canonical epoch milliseconds for the document-global field-date layout
   * axis. Kept beside `_showTrackedChanges` because borrowed documents can
   * change either axis after construction. */
  private _currentDate: Date | number | undefined;
  private _layoutViewGeneration = 0;

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
    if (opts.modelSources !== undefined) this._requestedShowTrackedChanges = opts.showTrackedChanges;
    this._currentDate = opts.currentDate;
    // `??` (not `||`): a caller's explicit `false` must disable the shadow, not
    // fall through to the default.
    this._pageShadow = opts.pageShadow ?? DEFAULT_SCROLL_PAGE_SHADOW;
    const borrowedDocument = (opts as InternalDocxScrollViewerOptions)[borrowedDocumentOption];
    this._borrowed = borrowedDocument !== undefined;
    if (borrowedDocument) {
      this._documentOwner = new TerminalResourceOwner('DocxScrollViewer', borrowedDocument, false);
      this._mode = resolveCanvasViewerMode('DocxScrollViewer', opts.mode, borrowedDocument);
    } else {
      this._documentOwner = new TerminalResourceOwner('DocxScrollViewer');
      this._mode = resolveCanvasViewerMode('DocxScrollViewer', opts.mode, undefined);
    }

    this._shell = new ScrollViewerShell(container, {
      background: opts.background,
      comments: !!opts.comments,
      onScroll: () => this._onScroll(),
      onOutsideComment: () => {
        if (this._activeCommentId === null) return;
        this._activeCommentId = null;
        this._activeCommentPage = null;
        for (const [index, slot] of this._slots) this._redrawSlotComments(index, slot);
        this._selection.emitChange();
      },
    });

    if (this._commentsEnabled()) {
      void loadDocxCommentUiRuntime().then((commentUi) => {
        if (this._destroyed) return;
        this._commentUi = commentUi;
        for (const [page, slot] of this._slots) this._redrawSlotComments(page, slot);
      }).catch((error) => this._reportRenderError(error));
    }

    this._selection.bind(!!opts.onSelectionContextChange, !!opts.onContextMenu);

    this._zoom.bind(this._container, this._scrollHost, this._opts.enableZoom !== false);

    if (this._borrowed) {
      this._layout.bind(borrowedDocument!);
      // A borrowed engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path re-runs it once width appears).
      this.relayout();
    }
  }

  async load(source: string | ArrayBuffer): Promise<void> {
    await this._loader.load(source);
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

  /** Refresh only review layers; a layout publication never detaches a painted canvas. */
  private _refreshCommentSurface(): void {
    if (!this._commentsEnabled() || this._slots.size === 0) return;
    this._syncSpacerWidth();
    for (const [page, slot] of this._slots) this._redrawSlotComments(page, slot);
  }

  /** CSS px width of page `i` at the current scale. */
  private _pageWidthPx(i: number): number {
    return this._doc!.pageSize(i).widthPt * PT_TO_PX * this._scale;
  }

  /** CSS px height of page `i` at the current scale. */
  private _pageHeightPx(i: number): number {
    return this._doc!.pageSize(i).heightPt * PT_TO_PX * this._scale;
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

  private _commentsEnabled(): boolean {
    return this._opts.comments === true || typeof this._opts.comments === 'object';
  }

  private _commentsOptions(): DocxCommentsOptions | undefined {
    return typeof this._opts.comments === 'object' ? this._opts.comments : undefined;
  }


  /** Widest authored page width. DOCX sections may differ by a fraction of a
   * point or switch orientation, and fit-width must cover the same extent as the
   * horizontal spacer. */
  private _widestPageWidthPt(): number {
    if (!this._doc) return 0;
    let widthPt = 0;
    const pageCount = this._layout.presentedPageCount || this._doc.pageCount;
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
    const w = this._viewport.fitWidth();
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
      this._layout.presentedPageCount = this._doc.pageCount;
    }
    if (!this._scaleEstablished) {
      if (!this._zoom.establishBase()) return;
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
    const n = Math.min(this._layout.presentedPageCount, this._doc!.pageCount);
    const h = new Array<number>(n);
    for (let i = 0; i < n; i++) h[i] = this._pageHeightPx(i);
    this._heights = h;
    this._scrollGeometry = createVirtualScrollGeometry(h, this._viewport.gap(), this._viewport.verticalPadding());
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
      this._viewport.overscan(),
    );
  }

  private _syncSpacer(): void { this._scroller.syncSpacer(); }

  /** Horizontal scroll extent: the widest page (docx pages can differ in width)
   *  plus both gutters. A spacer NARROWER than the container never creates a
   *  scrollbar (scrollWidth = max(clientWidth, content)), so it is always safe to
   *  set — it only matters when a zoomed-in page grows past the viewport, where it
   *  gives the gutters something to scroll to on either side. Max over per-page
   *  widths so the extent covers the widest page in the document. Called from
   *  `_syncSpacer` and after every scale change (zoom / resize re-fit) so the
   *  extent tracks the current page px width. */
  private _syncSpacerWidth(): void {
    const { left, right } = this._viewport.horizontalPadding();
    let maxW = 0;
    for (let i = 0; i < this._heights.length; i++) {
      const w = this._pageWidthPx(i);
      if (w > maxW) maxW = w;
    }
    this._commentMargin.syncSpacerWidth(maxW, left, right);
  }

  private _onScroll(): void {
    if (!this._doc || !this._scaleEstablished) return;
    this._mountVisible(undefined, false);
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(initialRenders?: Promise<void>[], repositionExisting = true): void {
    this._scroller.mount(initialRenders, repositionExisting);
  }

  private _emitVisiblePageChange(range: VisibleRange): void {
    if (this._doc) this._visibleEvents.publish(range, this._doc.pageCount, this.layoutComplete);
  }

  private _createSlot(): PageSlot {
    // The common canvas, selection and highlight stack is owned by core.
    const { wrapper, canvas, textLayer, highlightLayer } = createSlotHost(
      this._scrollHost, this._opts.enableTextSelection === true, this._pageShadow,
    );
    const { markerLayer: commentTintLayer, margin: commentMargin, decorationLayer: commentDecorationLayer } =
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
    this._scroller.recycleSlot(idx, slot);
  }

  private _resetSlot(_idx: number, slot: PageSlot): void {
    slot.dispatcher.destroy();
    if (!this._destroyed) {
      slot.dispatcher = new StaticCanvasRenderDispatcher(slot.canvas, this._mode === 'worker');
    }
    this._layers.reset(slot);
    slot.commentRuns = Object.freeze([]);
    slot.commentGeometry = null;
    slot.renderedPage = -1;
    slot.renderedScale = -1;
    slot.wrapper.remove();
  }

  private _positionSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    this._layers.position(i, slot, r.offsets[i], this._pageWidthPx(i), this._pageHeightPx(i),
      this._scrollHost.clientWidth, this._viewport.horizontalPadding().left);
  }

  /** Static slot dispatch is owned by core. */
  private _renderSlot(i: number, slot: PageSlot, reportErrors = true): Promise<void> | null {
    if (!this._doc) return null;
    this._syncReadingNotice();
    // Slot-identity guard: this slot is already rendering / has rendered page i.
    if (slot.renderedPage === i) return null;
    slot.renderedPage = i;

    const dpr = this._viewport.dpr();
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

    return this._main.render(i, slot, widthPx, dpr, 0, dispatcher, generation, reportErrors);
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

  /** A canvas's intended CSS box in px (the % denominators the overlay builders
   *  expect). Reads the inline `style.width`/`height` set by the render path,
   *  falling back to the backing-store size when unset; tolerates the `px` suffix. */
  private _canvasCssPx(canvas: HTMLCanvasElement): { width: number; height: number } {
    return {
      width: parseFloat(canvas.style.width) || canvas.width,
      height: parseFloat(canvas.style.height) || canvas.height,
    };
  }

  private _commitRenderedRuns(
    page: number, slot: PageSlot, runs: DocxTextRunInfo[], canvas: HTMLCanvasElement,
    wantedRuns: boolean, clearPreview: boolean,
  ): void {
    if (slot.textLayer) {
      if (clearPreview) this._clearTextLayerPreview(slot.textLayer);
      if (this._opts.enableTextSelection) {
        const { width, height } = this._canvasCssPx(canvas);
        buildDocxTextLayer(slot.textLayer, runs, width, height,
          this._hyperlinkHandler(), (font) => this._highlights.measure(font), page);
      }
    }
    if (wantedRuns) this._highlights.refreshRuns(page, runs);
    this._commitCommentRuns(page, slot, runs);
    this._highlights.redrawSlot(page, slot);
  }

  /** Route an async render failure to `onError`, or `console.error` when none is
   *  set (so failures are never fully silent), and never after teardown. */
  private _readingNoticeElement: HTMLElement | null = null;
  private _syncReadingNotice(): void {
    const notices = this._doc?.readingNotices ?? noReadingNotices;
    if (this._readingNoticeElement && notices.length === 0) this._clearReadingPages();
    this._readingNoticeElement = updateNativeReadingNotice(this._container, this._readingNoticeElement, notices);
  }
  private _clearReadingPages(): void {
    this._readingNoticeElement = updateNativeReadingNotice(this._container, this._readingNoticeElement, noReadingNotices);
    this._renderEpoch++;
    for (const [index, slot] of [...this._slots]) this._recycleSlot(index, slot);
    this._find.invalidate(); this._visibleEvents.reset();
  }
  private _reportRenderError(err: unknown): void {
    // Failure revocation belongs to the document operation's captured epoch,
    // never whichever document happens to be current when reporting settles.
    this._errorRouter.report(err);
  }

  private _renderSlotBitmap(
    i: number, slot: PageSlot, widthPx: number, dpr: number, scale: number,
    dispatcher = slot.dispatcher, generation = dispatcher.begin(), reportErrors = true,
  ): Promise<void> {
    return this._bitmap.render(i, slot, widthPx, dpr, scale, 0, dispatcher, generation, reportErrors);
  }

  /** Keep the public zoom facade while core owns scale, fit and anchoring. */
  setScale(scale: number): void { this._zoom.setScale(scale); }
  getScale(): number { return this._zoom.getScale(); }
  zoomIn(): void { this._zoom.zoomIn(); }
  zoomOut(): void { this._zoom.zoomOut(); }
  fitWidth(): void { this._zoom.fit('width'); }
  fitPage(): void { this._zoom.fit('page'); }

  private _previewVisible(): void { this._scroller.preview(); }

  private _previewSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    this._positionSlot(slot, i, r);
    this._layers.preview(slot, this._pageWidthPx(i), this._pageHeightPx(i), this._scale);
  }

  /** Restore a text overlay after its transient CSS zoom preview. */
  private _clearTextLayerPreview(layer: HTMLDivElement): void {
    clearTextLayerPreview(layer);
  }

  private _scheduleSettle(): void { this._scroller.scheduleSettle(DEFAULT_ZOOM_SETTLE_MS); }

  private _refreshSlotAtomically(i: number, slot: PageSlot): void {
    if (!this._doc) return;
    const dpr = this._viewport.dpr();
    const widthPx = this._pageWidthPx(i);
    const scale = this._scale;
    const epoch = this._renderEpoch;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
      return;
    }

    this._main.settle(i, slot, widthPx, dpr);
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
    if (this._pendingLoadAbort) {
      const changed = this._pendingRequestedView !== value;
      this._pendingRequestedView = value;
      // Keep an explicit false even when it matches the default, but avoid
      // notifying the in-flight paginator again for the same request.
      if (changed) this._pendingViewChanged?.();
    }
    const doc = this._doc;
    // Explicitness is independent of the current value: false before load
    // must win over a model source's true view default.
    if (this._opts.modelSources !== undefined) this._requestedShowTrackedChanges = value;
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
    if (this._opts.modelSources !== undefined) this._requestedShowTrackedChanges = value;
    this._find.invalidate();
    this._syncReadingNotice();
    // Re-render every mounted slot at the new variant, and relayout: heights,
    // spacer and mount window all follow the new page count, and a shrinking
    // document must recycle slots that are now out of range rather than ask for
    // pages that no longer exist.
    this._layout.apply({
      pageCount: doc?.pageCount ?? 0,
      exact: true,
      complete: doc?.layoutComplete !== false,
    });
  }

  scrollToPage(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    this._navigation.scrollToUnit(index, opts);
  }

  private _scrollToPageTarget(page: number,
    target: Readonly<{ x: number; y: number; w: number; h: number }>,
    opts?: { behavior?: 'auto' | 'smooth' }): void {
    this._navigation.scrollToPoint(page, target.x + target.w / 2,
      target.y + target.h / 2, opts);
  }

  /** Reveal an authored comment through the DOCX anchor navigation adapter. */
  async goToComment(
    commentId: string,
    opts?: { pageIndex?: number; behavior?: 'auto' | 'smooth' },
  ): Promise<boolean> {
    return this._commentNavigation.goToComment(commentId, opts);
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
    this._highlights.redrawAll();
    return matches;
  }

  /** Activate and reveal the next match, wrapping at the end. */
  async findNext(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._highlights.activate(this._find.next());
  }

  /** Activate and reveal the previous match, wrapping at the beginning. */
  async findPrev(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._highlights.activate(this._find.prev());
  }

  /** Clear the current query and every mounted highlight. */
  clearFind(): void {
    this._findRequestGeneration++;
    this._findActive = false;
    this._find.invalidate();
    this._highlights.redrawAll();
  }

  private async _collectPageRuns(page: number): Promise<DocxTextRunInfo[]> {
    if (!this._doc) return [];
    return this._doc.collectPageRuns(page, {
      width: this._pageWidthPx(page),
      currentDate: this._currentDate,
      ...(this._showTrackedChanges ? { showTrackedChanges: true } : {}),
    });
  }

  private _commitCommentRuns(
    page: number,
    slot: PageSlot,
    runs: readonly Readonly<DocxTextRunInfo>[],
  ): void {
    if (!slot.commentTintLayer) return;
    slot.commentRuns = Object.freeze([...runs]);
    this._commentNavigation.commitRuns(page, slot.commentRuns);
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
    this._commentMargin.syncMargin(slot.commentMargin);
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
        this._selection.clearElementContext();
        for (const [mountedPage, mountedSlot] of this._slots) {
          this._redrawSlotComments(mountedPage, mountedSlot);
        }
        this._selection.emitChange();
      },
      this._commentMargin.zoom(),
      READ_ONLY_COMMENT_MARGIN_WIDTH_PX,
      this._commentsOptions()?.markers !== false,
      this._commentsOptions()?.includeResolved === true,
      slot.commentDecorationLayer
        ? () => this._commentOverlay.schedule(page, slot, false)
        : undefined,
      slot.commentDecorationLayer
        ? () => this._commentOverlay.schedule(page, slot, true)
        : undefined,
    );
    this._commentOverlay.drawConnectors(page, slot);
  }



  private _onResize(): void { this._zoom.onResize(); }

  get topVisiblePage(): number {
    return this._scroller.lastRange?.topIndex ?? 0;
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
    const point = this._navigation.contentAtViewportY(y);
    return { page: point.unit, frac: point.frac };
  }

  /** @internal test hook: inverse of contentAtViewportYForTest. */
  viewportYOfForTest(page: number, frac: number): number {
    return this._navigation.viewportYOf(page, frac);
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
    return text ?? (this._selection.elementContext
      ? limitDocxElementContext(this._selection.elementContext, options.maxTextCharacters)
      : null);
  }

  /**
   * Tear down the viewer: remove the DOM subtree and (only for a self-loaded
   * engine) destroy the engine. A borrowed engine is left intact — the caller
   * owns its lifecycle. Per-slot worker ImageBitmaps are closed on recycle.
   */
  destroy(): void {
    if (this._destroyed) return;
    this._readingNoticeElement = updateNativeReadingNotice(this._container, this._readingNoticeElement, noReadingNotices);
    this._destroyed = true;
    this._pendingLoadAbort?.abort();
    this._pendingLoadAbort = null;
    this._pendingViewChanged = null;
    this._findRequestGeneration++;
    this._errorRouter.close();
    this._layout.unbind();
    this._layoutViewGeneration++;
    this._commentNavigation.reset();
    this._find.invalidate();
    this._findActive = false;
    this._selection.destroy();
    this._highlights.destroy();
    this._commentOverlay.destroy();
    this._selection.clearElementContext();
    // Cancel a pending settle so no re-render is dispatched after teardown
    // (design §7 mechanism 2). Clearing the timer avoids a wasted wake-up and
    // keeps fake-timer tests deterministic.
    this._scroller.destroy();
    this._zoom.destroy();
    this._commentMargin.destroy();
    this._documentOwner.close();
    this._shell.destroy();
  }
}
