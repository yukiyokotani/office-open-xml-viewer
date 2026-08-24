import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { RenderPageOptions } from './types';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import { buildDocxHighlightLayer, type DocxHighlightMatch } from './find-highlight-layer';
import { DocxFindController, type DocxMatchLocation } from './find';
import { openExternalHyperlink, PT_TO_PX, nextZoomStep, prevZoomStep, clampScale, fitScale } from '@silurus/ooxml-core';
import type { FindHighlightColors, HyperlinkTarget, FindMatch, FindMatchesOptions, OoxmlResourceMetrics, ViewerContextMenuEvent, ZoomableViewer } from '@silurus/ooxml-core';
import {
  CallerCanvasMount,
  CanvasOverlayHost,
  CanvasViewerErrorRouter,
  renderCanvasElementOutline,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import { invalidateDocxRenderTarget } from './paint/canvas-document';
import {
  readDocxTextSelectionContext,
  type DocxElementContext,
  type DocxSelectionContext,
  type DocxSelectionContextOptions,
} from './selection-context';
import {
  limitDocxElementContext,
  MAX_DOCX_ELEMENT_TEXT_CHARACTERS,
} from './element-context';

const borrowedDocumentOption = Symbol('DocxViewer.borrowedDocument');
type InternalDocxViewerOptions = DocxViewerOptions & {
  [borrowedDocumentOption]?: DocxDocument;
};

export interface DocxViewerOptions extends Omit<RenderPageOptions, 'onTextRun'>, LoadOptions {
  container?: HTMLElement;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on document content.
   */
  enableTextSelection?: boolean;
  /**
   * Enable read-only selection of rendered pictures, charts, and shapes. The
   * selected object exposes element context and receives a non-editable outline.
   * Default false; hit-testing runs only for clicks when enabled.
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
  /** Called when a page finishes rendering. */
  onPageChange?: (index: number, total: number) => void;
  /** IX9 zoom contract ({@link ZoomableViewer}) — the clamp range for
   *  {@link DocxViewer.setScale} / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage`,
   *  as user-facing zoom factors (`1` = 100% = the page at its natural pt→px
   *  size). Defaults 0.1–4 (10%–400%), matching the other viewers. */
  zoomMin?: number;
  zoomMax?: number;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100%): from
   *  {@link DocxViewer.setScale}, `zoomIn`/`zoomOut`, or `fitWidth`/`fitPage`.
   *  Named `onScaleChange` to match the pptx/xlsx viewers so all five share one
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
   *  (their colour/underline are painted on the canvas) but are inert, like plain
   *  text. Set it to disable clickable links entirely — e.g. in a preview where
   *  navigation must not leave the current view. */
  enableHyperlinks?: boolean;
  /**
   * Receives asynchronous Viewer-managed failures that cannot be observed by
   * awaiting the method that started them. Failures from `load()`, including
   * its initial render, always reject that Promise and are not also delivered
   * here. Later event-driven render failures invoke this callback, or fall back
   * to `console.error` when omitted.
   *
   * Stable cases can be narrowed with `OoxmlError`,
   * `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   * this package. Other failures remain `Error` values; do not parse message
   * text as an API. A `code` of `parser-crashed` identifies a recognized WASM
   * trap, not a reliably classified OOM.
   */
  onError?: (err: Error) => void;
}

export class DocxViewer implements ZoomableViewer {
  private readonly _documentOwner: TerminalResourceOwner<DocxDocument>;
  private get _doc(): DocxDocument | null { return this._documentOwner.current; }
  private readonly _borrowed: boolean;
  private readonly _hostWindow: Window & typeof globalThis;
  private _currentPage = 0;
  /**
   * IX9 explicit zoom factor (`1` = 100% = the page at its natural pt→px width),
   * or `null` when the caller has never invoked a zoom method. `null` preserves
   * the pre-IX9 render path EXACTLY: the page renders at `opts.width` (or its
   * natural width when that is unset), so default rendering is byte-identical. The
   * first `setScale`/`zoomIn`/`zoomOut`/`fitWidth`/`fitPage` call latches a number
   * here, after which `_renderPage` derives the canvas width from it instead.
   */
  private _scale: number | null = null;
  private _canvas: HTMLCanvasElement;
  private _wrapper: HTMLDivElement;
  private readonly _canvasMount: CallerCanvasMount;
  private _textLayer: HTMLDivElement | null = null;
  /** IX2 — the find-highlight overlay layer. Always created (independent of
   *  `enableTextSelection`): highlights ride the same positioned-DOM overlay
   *  mechanism as the selection layer but are visible boxes, not transparent
   *  spans. Sits above the text layer so a highlight shows over a link's hit
   *  region without stealing its clicks (`pointer-events:none`). */
  private _highlightLayer: HTMLDivElement | null = null;
  private _elementLayer: HTMLDivElement | null = null;
  /** IX2 — find state (per-page runs, matches, active cursor). */
  private _find: DocxFindController;
  /** A 2d context used only to measure text for highlight geometry (its own
   *  1×1 offscreen canvas, so measuring never touches the visible canvas). */
  private _measureCtx: CanvasRenderingContext2D | null = null;
  private _opts: DocxViewerOptions;
  private readonly _mode: 'main' | 'worker';
  private readonly _renderDispatcher: StaticCanvasRenderDispatcher;
  private readonly _errorRouter: CanvasViewerErrorRouter;
  private _destroyed = false;
  private _selectionChangeListener: (() => void) | null = null;
  private _selectionContextKey = 'null';
  private _elementContext: DocxElementContext | null = null;
  private _elementHitGeneration = 0;
  private _elementClickListener: ((event: MouseEvent) => void) | null = null;
  private _contextMenuListener: ((event: MouseEvent) => void) | null = null;
  /**
   * Create a Viewer that borrows an already-loaded document.
   *
   * The document's render mode is authoritative. The returned Viewer cannot
   * load another source, and destroying it leaves the caller-owned document
   * open. Call {@link goToPage} to render the initial page.
   */
  static fromDocument(
    canvas: HTMLCanvasElement,
    document: DocxDocument,
    opts: Omit<DocxViewerOptions, keyof LoadOptions> = {},
  ): Omit<DocxViewer, 'load'> {
    return new DocxViewer(canvas, {
      ...opts,
      [borrowedDocumentOption]: document,
    } as InternalDocxViewerOptions);
  }

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;
    const borrowedDocument = (opts as InternalDocxViewerOptions)[borrowedDocumentOption];
    this._borrowed = borrowedDocument !== undefined;
    this._mode = resolveCanvasViewerMode('DocxViewer', opts.mode, borrowedDocument);
    this._documentOwner = new TerminalResourceOwner(
      'DocxViewer',
      borrowedDocument ?? null,
      false,
    );
    const hostWindow = canvas.ownerDocument?.defaultView ??
      (typeof window !== 'undefined' ? window : null);
    if (!hostWindow) throw new Error('DocxViewer requires a canvas with an active Window');
    this._hostWindow = hostWindow;

    this._canvasMount = new CallerCanvasMount(canvas, {
      wrapperCssText: 'position:relative;display:inline-block;vertical-align:top;',
      forceDisplayBlock: true,
    });
    this._wrapper = this._canvasMount.wrapper;
    this._renderDispatcher = new StaticCanvasRenderDispatcher(canvas, this._mode === 'worker');
    this._errorRouter = new CanvasViewerErrorRouter('DocxViewer', opts.onError);
    const overlays = new CanvasOverlayHost(
      this._wrapper,
      opts.enableTextSelection === true,
      opts.enableElementSelection === true,
    );
    this._textLayer = overlays.textLayer;
    this._highlightLayer = overlays.highlightLayer;
    this._elementLayer = overlays.elementLayer;
    if (this._textLayer && (opts.onSelectionContextChange || opts.enableElementSelection)) {
      this._selectionChangeListener = () => this._emitSelectionContextChange();
      this._wrapper.ownerDocument.addEventListener('selectionchange', this._selectionChangeListener);
    }
    if (opts.enableElementSelection) {
      this._elementClickListener = (event) => {
        void this._onElementClick(event).catch((error) => this._reportRenderError(error));
      };
      this._wrapper.addEventListener('click', this._elementClickListener);
    }
    if (opts.onContextMenu) {
      this._contextMenuListener = (event) => this._onContextMenu(event);
      this._wrapper.addEventListener('contextmenu', this._contextMenuListener);
    }

    this._find = new DocxFindController(
      () => this.pageCount,
      (page) => this._collectPageRuns(page),
    );
  }

  /**
   * Load a DOCX from URL or ArrayBuffer and render the first page.
   *
   * Parse, load, and initial-render failures always reject this Promise.
   * `onError` is reserved for later Viewer-managed work that has no directly
   * awaitable method result, so one failure is never delivered twice.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this._destroyed) throw new Error('DocxViewer is destroyed');
    if (this._borrowed) {
      throw new Error(
        'DocxViewer.load() is unsupported on a Viewer created by fromDocument(); ' +
          'the borrowed document is already loaded.',
      );
    }
    // SC20 atomic swap: retain the previous engine locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // engine's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current document + its rendered page intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old engine is freed the moment the new model arrives).
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
        mode: this._mode,
      }), () => {
        // Invalidate operations owned by the old document before its worker is
        // terminated, so their expected rejection cannot surface as a reload
        // failure for the winning document.
        this._invalidateElementContext(false);
        elementInvalidated = true;
        this._renderDispatcher.begin();
        this._find.invalidate();
      });
      if (!doc) return;
      if (this._destroyed) throw new Error('DocxViewer is destroyed');
      this._currentPage = 0;
      // A new document invalidates any prior find state (cached runs / matches).
      this._find.invalidate();
      await this._render();
    } catch (err) {
      if (this._destroyed) throw new Error('DocxViewer is destroyed');
      throw err instanceof Error ? err : new Error(String(err));
    }
    if (elementInvalidated && !this._destroyed) this._emitSelectionContextChange();
  }

  get pageCount(): number {
    return this._doc?.pageCount ?? 0;
  }

  get currentPage(): number {
    return this._currentPage;
  }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement {
    return this._canvas;
  }

  async goToPage(index: number): Promise<void> {
    if (!this._doc) return;
    const clamped = Math.max(0, Math.min(index, this.pageCount - 1));
    const changed = clamped !== this._currentPage;
    if (changed) this._invalidateElementContext(false);
    this._currentPage = clamped;
    await this._render();
    if (changed && !this._destroyed) this._emitSelectionContextChange();
  }

  async nextPage(): Promise<void> { await this.goToPage(this._currentPage + 1); }
  async prevPage(): Promise<void> { await this.goToPage(this._currentPage - 1); }

  // ─── IX9 zoom contract (ZoomableViewer) ───────────────────────────────────

  /** Natural (100%) CSS-px width of the current page — `widthPt × PT_TO_PX`.
   *  This is the scale-1 reference every zoom factor multiplies. 0 when nothing
   *  is loaded. */
  private _naturalWidthPx(): number {
    if (!this._doc || this._doc.pageCount === 0) return 0;
    return this._doc.pageSize(this._currentPage).widthPt * PT_TO_PX;
  }

  /**
   * The width (CSS px) `_renderPage` renders the current page at, honouring the
   * zoom state. `_scale === null` (no zoom method ever called) ⇒ the pre-IX9
   * value `opts.width` verbatim (byte-identical default: `undefined` lets the
   * renderer use the page's natural width). Once a factor latched ⇒
   * `naturalWidth × scale` (rounded), so the on-screen page is exactly `scale ×`
   * its natural size regardless of the original `opts.width`.
   */
  private _renderWidth(): number | undefined {
    if (this._scale === null) return this._opts.width;
    const natural = this._naturalWidthPx();
    if (natural <= 0) return this._opts.width; // unloaded — fall back, defer
    return Math.round(natural * this._scale);
  }

  /** IX9 {@link ZoomableViewer} — the current zoom factor (`1` = 100%). Before
   *  any zoom method is called this is the EFFECTIVE scale implied by the current
   *  render width: `opts.width / naturalWidth`, or `1` when `opts.width` is unset
   *  (the page renders at its natural size) or nothing is loaded. */
  getScale(): number {
    if (this._scale !== null) return this._scale;
    const natural = this._naturalWidthPx();
    if (natural <= 0) return 1;
    return this._opts.width && this._opts.width > 0 ? this._opts.width / natural : 1;
  }

  private _zoomMin(): number { return this._opts.zoomMin ?? 0.1; }
  private _zoomMax(): number { return this._opts.zoomMax ?? 4; }

  /**
   * IX9 {@link ZoomableViewer} — set the absolute zoom factor (`1` = 100% = the
   * page at its natural pt→px width), clamped to `[zoomMin, zoomMax]`, and
   * re-render the current page at the new size. Fires `onScaleChange` when the
   * clamped factor actually changes. Resolves once the re-render settles. A no-op
   * (but still latches the scale) when nothing is loaded.
   */
  async setScale(scale: number): Promise<void> {
    const next = clampScale(scale, this._zoomMin(), this._zoomMax());
    const changed = next !== this.getScale();
    this._scale = next;
    await this._render();
    if (changed) this._opts.onScaleChange?.(next);
  }

  /** IX9 {@link ZoomableViewer} — step up to the next rung of the shared zoom
   *  ladder (clamped to `zoomMax`). */
  async zoomIn(): Promise<void> { await this.setScale(nextZoomStep(this.getScale())); }

  /** IX9 {@link ZoomableViewer} — step down to the next lower ladder rung. */
  async zoomOut(): Promise<void> { await this.setScale(prevZoomStep(this.getScale())); }

  /**
   * IX9 {@link ZoomableViewer} — fit the current page's WIDTH to the host
   * container (the element the canvas lives in, or `opts.container` if supplied),
   * then re-render. Defers (no-op) when nothing is loaded or the container is
   * unlaid-out. Routes through {@link setScale}, so the factor is clamped and
   * `onScaleChange` fires.
   */
  async fitWidth(): Promise<void> { await this._fit('width'); }

  /**
   * IX9 {@link ZoomableViewer} — fit the WHOLE current page (width and height)
   * inside the container so it is visible without scrolling; takes the tighter of
   * the width/height fit. Defers when unloaded / unlaid-out.
   */
  async fitPage(): Promise<void> { await this._fit('page'); }

  /** Shared fit for {@link fitWidth}/{@link fitPage}: measure the natural page
   *  size + the container box, ask core's pure `fitScale`, apply via setScale. */
  private async _fit(mode: 'width' | 'page'): Promise<void> {
    if (!this._doc || this._doc.pageCount === 0) return;
    const size = this._doc.pageSize(this._currentPage);
    const container = this._fitContainer();
    if (!container) return;
    const scale = fitScale(
      {
        contentWidth: size.widthPt * PT_TO_PX,
        contentHeight: size.heightPt * PT_TO_PX,
        containerWidth: container.clientWidth,
        containerHeight: container.clientHeight,
      },
      mode,
    );
    if (scale <= 0) return; // unlaid-out / empty — defer
    await this.setScale(scale);
  }

  /** The element a fit measures against: the explicit `opts.container`, else the
   *  host the wrapper was inserted into (`_wrapper.parentElement`). `null` when
   *  the canvas was mounted detached (no host to fit to). */
  private _fitContainer(): { clientWidth: number; clientHeight: number } | null {
    return this._opts.container ?? this._wrapper.parentElement ?? null;
  }

  /**
   * IX2 — find every occurrence of `query` in the document and highlight them
   * all (a soft box per match, drawn on the highlight overlay over the drawn
   * glyphs). Returns every match in document order, each tagged with its
   * `{ page }` (0-based). Case-insensitive by default (browser find-in-page);
   * pass `{ caseSensitive: true }` to match case exactly.
   *
   * Scans all pages, so a large document renders each page once (offscreen) to
   * read its text (the visible page reuses its on-screen render). IX6 — works in
   * BOTH `mode: 'main'` and `mode: 'worker'`: in worker mode each page's run
   * geometry is collected off-thread and shipped back, so find returns the same
   * matches on the same code path. An empty query clears the find and returns `[]`.
   */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<DocxMatchLocation>[]> {
    if (!this._doc) return [];
    const matches = await this._find.find(query, opts);
    // Redraw the current page's highlights (matches on it become visible without
    // navigating). Cheap DOM geometry — no page re-render.
    this._redrawHighlights();
    return matches;
  }

  /**
   * IX2 — move to the next match (wrap-around from last to first), navigating to
   * its page if needed, and draw it in the distinct active-match colour. Returns
   * the now-active match, or `null` when there are no matches. Call
   * {@link findText} first.
   */
  async findNext(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** IX2 — move to the previous match (wrap-around from first to last). */
  async findPrev(): Promise<FindMatch<DocxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
    this._find.invalidate();
    this._redrawHighlights();
  }

  /** Navigate to the active match's page (if not already there) and redraw the
   *  highlights so the active box shows in the emphasis colour. */
  private async _activateMatch(
    match: FindMatch<DocxMatchLocation> | null,
  ): Promise<FindMatch<DocxMatchLocation> | null> {
    if (!match) {
      this._redrawHighlights();
      return null;
    }
    if (match.location.page !== this._currentPage) {
      // goToPage re-renders, which rebuilds the highlight layer for the new page.
      await this.goToPage(match.location.page);
    } else {
      this._redrawHighlights();
    }
    return match;
  }

  /** Rebuild the highlight overlay for the current page from cached runs
   *  (no page re-render). */
  private _redrawHighlights(): void {
    const runs = this._find.pageRuns(this._currentPage) ?? [];
    this._buildHighlightLayer(runs);
  }

  /** Latest content-free resource metrics for the loaded document. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this._doc) throw new Error('Document not loaded');
    return await this._doc.getResourceMetrics();
  }

  /** Return the current browser text selection or clicked drawing context. */
  getSelectionContext(options: DocxSelectionContextOptions = {}): DocxSelectionContext | null {
    if (this._destroyed) throw new Error('DocxViewer is destroyed');
    const text = this._textLayer
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
      this._redrawElementOutline();
    }
    const key = JSON.stringify(context);
    if (key === this._selectionContextKey) return;
    this._selectionContextKey = key;
    this._opts.onSelectionContextChange?.(context ? structuredClone(context) : null);
  }

  private _setElementContext(context: DocxElementContext | null): void {
    this._elementContext = context ? structuredClone(context) : null;
    this._redrawElementOutline();
    this._emitSelectionContextChange();
  }

  private _invalidateElementContext(notify = true): void {
    this._elementHitGeneration++;
    this._elementContext = null;
    this._redrawElementOutline();
    if (notify) this._emitSelectionContextChange();
  }

  private _redrawElementOutline(): void {
    const context = this._elementContext;
    const doc = this._doc;
    if (!context || !doc || context.pageIndex !== this._currentPage) {
      renderCanvasElementOutline(this._elementLayer, null);
      return;
    }
    const page = doc.pageSize(context.pageIndex);
    renderCanvasElementOutline(this._elementLayer, {
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
    if (this._textLayer && readDocxTextSelectionContext(
      this._wrapper,
      this._wrapper.ownerDocument?.getSelection?.() ?? null,
    )) {
      this._emitSelectionContextChange();
      return this._destroyed ? null : this.getSelectionContext();
    }
    if (!this._opts.enableElementSelection) return this.getSelectionContext();
    const rect = this._canvas.getBoundingClientRect();
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
    const pageIndex = this._currentPage;
    const pageSize = doc.pageSize(pageIndex);
    let context: DocxElementContext | null;
    try {
      context = await doc.getElementContextAt(pageIndex, {
        xPt: localX / rect.width * pageSize.widthPt,
        yPt: localY / rect.height * pageSize.heightPt,
      }, {
        currentDate: this._opts.currentDate,
        maxTextCharacters: MAX_DOCX_ELEMENT_TEXT_CHARACTERS,
      });
    } catch (error) {
      if (this._destroyed || generation !== this._elementHitGeneration ||
        pageIndex !== this._currentPage || doc !== this._doc) return null;
      throw error;
    }
    if (this._destroyed || generation !== this._elementHitGeneration ||
      pageIndex !== this._currentPage || doc !== this._doc) return null;
    this._setElementContext(context);
    return this._destroyed ? null : this.getSelectionContext();
  }

  /**
   * Terminate the parser worker and release resources.
   *
   * The caller-owned `<canvas>` is returned to the DOM position it held before
   * the constructor was called (same parent, same next-sibling) and its inline
   * `display` is restored, so the canvas can be reused — e.g. to construct a new
   * viewer on the same element. If the canvas was passed detached (no parent) it
   * is simply removed from the internal wrapper. Safe to call more than once.
   */
  destroy(): void {
    if (this._destroyed) return;
    this._destroyed = true;
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). Bump the load generation
    // too so a load() still in flight is treated as superseded and its engine is
    // cleaned up rather than installed onto a torn-down viewer.
    this._errorRouter.close();
    this._renderDispatcher.destroy();
    invalidateDocxRenderTarget(this._canvas);
    this._documentOwner.close();
    // IX2 — drop the find state (matches + cached runs) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer.
    this._find.invalidate();
    if (this._selectionChangeListener) {
      this._wrapper.ownerDocument.removeEventListener('selectionchange', this._selectionChangeListener);
      this._selectionChangeListener = null;
    }
    this._elementHitGeneration++;
    if (this._elementClickListener) {
      this._wrapper.removeEventListener('click', this._elementClickListener);
      this._elementClickListener = null;
    }
    if (this._contextMenuListener) {
      this._wrapper.removeEventListener('contextmenu', this._contextMenuListener);
      this._contextMenuListener = null;
    }
    this._elementContext = null;
    this._canvasMount.restore();
  }

  private async _render(): Promise<void> {
    const generation = this._renderDispatcher.begin();
    try {
      await this._renderPage(generation);
    } catch (err) {
      if (!this._renderDispatcher.isCurrent(generation)) return;
      throw err;
    }
  }

  /** Route a render failure to `onError`, or `console.error` when none is given
   *  (never fully silent), and never after teardown. Mirrors the scroll viewers'
   *  `_reportRenderError`. */
  private _reportRenderError(err: unknown): void {
    this._errorRouter.report(err);
  }

  private async _renderPage(generation: number): Promise<void> {
    if (!this._doc) return;
    const isWorker = this._mode === 'worker';
    // IX9: the width to render at. When no zoom method was ever called
    // (`_scale === null`) this is exactly `opts.width` (pre-IX9 path, byte-
    // identical default); once a zoom latched a factor it is `naturalWidth ×
    // scale`.
    const renderWidth = this._renderWidth();
    // Collect runs unconditionally (not just when a text layer exists): the
    // find-highlight overlay needs the current page's run geometry too, and
    // caching them here means find() reuses the visible render for this page
    // instead of re-rendering it offscreen. IX6 — in worker mode the runs ride
    // back beside the bitmap, so both modes populate the same `runs` array,
    // at the zoom-aware `renderWidth` (the geometry follows setScale).
    const runs: DocxTextRunInfo[] = [];
    const onTextRun = (r: DocxTextRunInfo) => runs.push(r);
    if (isWorker) {
      const dpr = this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
      // Only serializable render options may cross to the worker — spreading the
      // full viewer opts would postMessage non-cloneable values (the math
      // engine, callbacks, container element) and throw a DataCloneError. The
      // `onTextRun` callback stays main-thread; the proxy invokes it with the
      // worker's returned runs (IX6).
      const bmp = await this._doc.renderPageToBitmap(this._currentPage, {
        width: renderWidth,
        dpr: this._opts.dpr,
        defaultTextColor: this._opts.defaultTextColor,
        currentDate: this._opts.currentDate,
        onTextRun,
      });
      // The bitmap is sized in device px; mirror the main renderer by setting
      // the CSS size to the logical (÷dpr) dimensions so it isn't 2× on HiDPI.
      if (!this._renderDispatcher.commitBitmap(generation, bmp, {
        cssWidth: Math.round(bmp.width / dpr),
        cssHeight: Math.round(bmp.height / dpr),
      })) return;
    } else {
      await this._doc.renderPage(this._canvas, this._currentPage, { ...this._opts, width: renderWidth, onTextRun });
      if (!this._renderDispatcher.isCurrent(generation)) return;
    }
    // IX6 — identical overlay build for both modes: the run geometry the worker
    // shipped is the same shape `onTextRun` emits in main mode.
    if (this._textLayer) {
      this._buildTextLayer(this._textLayer, runs);
    }
    // Feed the just-rendered page's runs to the find controller so highlight
    // geometry matches exactly what was drawn, then (re)draw the highlights.
    this._find.setPageRuns(this._currentPage, runs);
    this._buildHighlightLayer(runs);
    this._opts.onPageChange?.(this._currentPage, this.pageCount);
  }

  /** Draw the find-highlight boxes for the current page from its runs. Clears
   *  the overlay when there is no active find. */
  private _buildHighlightLayer(runs: DocxTextRunInfo[]): void {
    const layer = this._highlightLayer;
    if (!layer) return;
    const { width, height } = this._canvasCssPx();
    const highlights: DocxHighlightMatch[] = this._find.pageHighlights(this._currentPage);
    buildDocxHighlightLayer(
      layer,
      runs,
      highlights,
      width,
      height,
      (font) => this._measureForFont(font),
      this._opts.findHighlightColors,
    );
  }

  /** The canvas's intended CSS box in px (the % denominators the overlay builders
   *  expect). Reads the inline `style.width`/`height` set by the render path
   *  (which mirror the render's logical size), falling back to the backing-store
   *  dimensions when unset. Parsing tolerates the trailing `px`. */
  private _canvasCssPx(): { width: number; height: number } {
    const w = parseFloat(this._canvas.style.width) || this._canvas.width;
    const h = parseFloat(this._canvas.style.height) || this._canvas.height;
    return { width: w, height: h };
  }

  /** A width-measurer primed with `font`, backed by a private 1×1 canvas so it
   *  never disturbs the visible canvas's context state. */
  private _measureForFont(font: string): (s: string) => number {
    if (!this._measureCtx) {
      const c = document.createElement('canvas');
      this._measureCtx = c.getContext('2d');
    }
    const ctx = this._measureCtx;
    if (!ctx) return (s) => s.length; // measurement unavailable (headless w/o canvas)
    ctx.font = font;
    return (s) => ctx.measureText(s).width;
  }

  /** Render a page to a throwaway offscreen canvas purely to collect its runs
   *  (text + geometry) for search, without touching the visible canvas. Used by
   *  the find controller for pages other than the one on screen. */
  private async _collectPageRuns(page: number): Promise<DocxTextRunInfo[]> {
    if (!this._doc) return [];
    // IX6 — `collectPageRuns` renders the page (off-thread in worker mode, to a
    // throwaway offscreen canvas in main mode) and returns just its run
    // geometry. The find controller only calls this for pages OTHER than the one
    // on screen (the visible page's runs are cached by _renderPage). Pass the
    // same geometry-affecting options as the visible render — including the
    // IX9 zoom-aware `_renderWidth()`, so the harvested geometry matches what a
    // navigation to that page would draw at the current scale (worker mode
    // postMessages these — no callbacks/engine).
    return this._doc.collectPageRuns(page, {
      width: this._renderWidth(),
      currentDate: this._opts.currentDate,
    });
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
    const { width, height } = this._canvasCssPx();
    buildDocxTextLayer(
      layer,
      runs,
      width,
      height,
      this._hyperlinkHandler(),
      // §17.3.2.10 縦中横 (#836) — the same measurer the highlight overlay uses,
      // so a tate-chu-yoko selection span is clamped to its drawn one-em cell.
      (font) => this._measureForFont(font),
      this._currentPage,
    );
  }

  /**
   * IX1/IX-nav — the click handler passed to the text-layer overlay. When the
   * caller supplied `onHyperlinkClick`, it fully owns the behaviour (the default
   * is suppressed). Otherwise the built-in default is: an external link opens in
   * a new tab through core `openExternalHyperlink` (URL sanitised against the
   * safe scheme allowlist, `noopener,noreferrer`); an internal `<w:anchor>` link
   * resolves its bookmark name to a page via
   * {@link DocxDocument.getBookmarkPage} (ECMA-376 §17.16.23) and jumps there
   * with {@link goToPage}. An anchor naming no known bookmark is a safe no-op
   * rather than a jump to a guessed page.
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
        openExternalHyperlink(target.url, undefined, this._hostWindow);
        return;
      }
      // Internal anchor (IX-nav): map the bookmark name to its destination page
      // and navigate. `undefined` ⇒ no bookmark of that name ⇒ inert.
      const page = this._doc?.getBookmarkPage(target.ref);
      if (page !== undefined) {
        void this.goToPage(page).catch((error) => this._reportRenderError(error));
      }
    };
  }
}
