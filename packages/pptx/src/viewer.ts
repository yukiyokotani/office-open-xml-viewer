import {
  invalidatePptxRenderTarget,
  type RenderOptions,
  type PptxTextRunInfo,
} from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { buildPptxHighlightLayer, type PptxHighlightMatch } from './find-highlight-layer';
import { PptxFindController, type PptxMatchLocation } from './find';
import { PptxPresentation, type LoadOptions } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import { countVisible } from './hidden';
import type { DimOptions } from './types';
import {
  type HyperlinkTarget,
  type FindHighlightColors,
  type FindMatch,
  type FindMatchesOptions,
  type OoxmlResourceMetrics,
  type ViewerContextMenuEvent,
  type ZoomableViewer,
  EMU_PER_PX,
  openExternalHyperlink,
  nextZoomStep,
  prevZoomStep,
  clampScale,
  fitScale,
} from '@silurus/ooxml-core';
import {
  CallerCanvasMount,
  CanvasOverlayHost,
  CanvasViewerErrorRouter,
  renderCanvasElementOutline,
  resolveCanvasViewerMode,
  StaticCanvasRenderDispatcher,
  TerminalResourceOwner,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import {
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
import { renderPptxFocusedSlide } from './focused-view-runtime';
import {
  subscribePptxLayout,
  type PptxLayoutPublication,
} from './presentation-layout-events';

const borrowedPresentationOption = Symbol('PptxViewer.borrowedPresentation');
type InternalPptxViewerOptions = PptxViewerOptions & {
  [borrowedPresentationOption]?: PptxPresentation;
};

/** How {@link PptxViewer} presents hidden slides (`<p:sld show="0">`). */
export type HiddenSlideMode = 'show' | 'skip' | 'dim';

/** Default `'dim'` overlay: 60% white (hidden content shows at 40%). */
const DEFAULT_HIDDEN_DIM: DimOptions = { color: '#ffffff', opacity: 0.6 };

export interface PptxViewerOptions extends Pick<RenderOptions, 'width' | 'dpr'>, LoadOptions {
  /** Called when a slide finishes rendering or progressive availability changes. */
  onSlideChange?: (index: number, total: number, layoutComplete: boolean) => void;
  /**
   * Receives asynchronous Viewer-managed failures that cannot be observed by
   * awaiting the method that started them. Failures from `load()`, including
   * its initial render, always reject that Promise and are not also delivered
   * here. Later event-driven render or media failures invoke this callback, or
   * fall back to `console.error` when omitted.
   *
   * Stable cases can be narrowed with `OoxmlError`,
   * `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   * this package. Other failures remain `Error` values; do not parse message
   * text as an API. A `code` of `parser-crashed` identifies a recognized WASM
   * trap, not a reliably classified OOM.
   */
  onError?: (err: Error) => void;
  /** IX9 zoom contract ({@link ZoomableViewer}) — the clamp range for
   *  {@link PptxViewer.setScale} / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage`,
   *  as user-facing zoom factors (`1` = 100% = the slide at its natural
   *  EMU→px size). Defaults 0.1–4 (10%–400%), matching the other viewers. */
  zoomMin?: number;
  zoomMax?: number;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100%): from
   *  {@link PptxViewer.setScale}, `zoomIn`/`zoomOut`, or `fitWidth`/`fitPage`.
   *  Named `onScaleChange` to match the docx/xlsx viewers so all five share one
   *  notification shape. */
  onScaleChange?: (scale: number) => void;
  /**
   * Enable interactive audio/video playback. When true, slides are rendered
   * via {@link PptxPresentation.presentSlide} so media elements become
   * clickable and the viewer draws its own play/pause chrome. When false
   * (default) the viewer renders a static slide with a non-interactive play
   * badge over media posters.
   */
  enableMediaPlayback?: boolean;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on slide content.
   */
  enableTextSelection?: boolean;
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
   * How hidden slides (`<p:sld show="0">`, §19.3.1.38) are presented:
   * - `'show'` (default): drawn like any other slide.
   * - `'skip'`: sequential navigation (`nextSlide`/`prevSlide`, initial load)
   *   jumps over them; absolute indices are unchanged, and an explicit
   *   `goToSlide(i)` to a hidden slide is still honored.
   * - `'dim'`: drawn under a translucent overlay (PowerPoint thumbnail look).
   *
   * Named to match the {@link PptxViewer.hiddenSlideMode} getter and
   * {@link PptxViewer.setHiddenSlideMode} setter.
   */
  hiddenSlideMode?: HiddenSlideMode;
  /**
   * Overrides for the `'dim'` overlay. Merged over the default
   * `{ color: '#ffffff', opacity: 0.6 }`. A `Partial<DimOptions>` so it stays
   * in sync if {@link DimOptions} gains a field.
   */
  hiddenSlideDim?: Partial<DimOptions>;
  /**
   * IX1 (design decision — NOT user-confirmed, integrator may veto). Fires on a
   * hyperlink click (a text run whose `<a:rPr>` carried an `<a:hlinkClick>`;
   * requires {@link enableTextSelection} so the overlay spans exist). Default
   * when omitted: external → {@link openExternalHyperlink} (new tab, sanitised,
   * noopener); internal slide-jump → {@link goToSlide} once the action resolves
   * to a slide index via {@link PptxPresentation.resolveInternalTarget} (a jump
   * that resolves to no reachable slide is a safe no-op). When provided, the
   * viewer calls this instead and takes NO default action.
   */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the hyperlink machinery is not wired at all: the overlay's link
   *  spans are non-interactive, so there is no pointer cursor, no title tooltip,
   *  no default navigation (external new-tab / internal slide jump), and
   *  `onHyperlinkClick` is never called. Links still render exactly as authored
   *  (theme `hlink` colour + underline are painted on the canvas) but are inert,
   *  like plain text. */
  enableHyperlinks?: boolean;
}

/**
 * Opinionated single-canvas PPTX viewer.
 *
 * Accepts a caller-supplied `<canvas>` element and wraps it in a positioned
 * container for the optional text-selection overlay.  The wrapper is inserted
 * into the canvas's existing parent (reparent), so the canvas stays at its
 * original position in the DOM.
 *
 * For custom layouts (multi-canvas, thumbnails, scroll view) use PptxPresentation directly.
 */
export class PptxViewer implements ZoomableViewer {
  private readonly canvas: HTMLCanvasElement;
  private readonly wrapper: HTMLDivElement;
  private readonly canvasMount: CallerCanvasMount;
  /**
   * IX9 explicit zoom factor (`1` = 100% = the slide at its natural EMU→px
   * width), or `null` when the caller has never invoked a zoom method. `null`
   * preserves the pre-IX9 render path EXACTLY: the slide renders at `opts.width`
   * (or `canvas.offsetWidth || 960` when unset), so default rendering is
   * byte-identical. The first zoom call latches a number here, after which
   * {@link _targetWidth} derives the render width from it.
   */
  private _scale: number | null = null;
  private textLayer: HTMLDivElement | null = null;
  /** IX2 — the find-highlight overlay layer (always created, above the text
   *  layer, `pointer-events:none`). */
  private highlightLayer: HTMLDivElement | null = null;
  private elementLayer: HTMLDivElement | null = null;
  /** IX2 — find state (per-slide runs, matches, active cursor). */
  private _find: PptxFindController;
  private _findGeneration = 0;
  /** Private 2d context for measuring highlight text (own 1×1 canvas). */
  private _measureCtx: CanvasRenderingContext2D | null = null;
  private readonly presentationOwner: TerminalResourceOwner<PptxPresentation>;
  private get engine(): PptxPresentation | null { return this.presentationOwner.current; }
  private readonly borrowed: boolean;
  private readonly hostWindow: Window & typeof globalThis;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;
  private _renderedSlide = -1;
  private _hiddenMode: HiddenSlideMode;
  private handle: PresentationHandle | null = null;
  private readonly _mode: 'main' | 'worker';
  private readonly renderDispatcher: StaticCanvasRenderDispatcher;
  private readonly errorRouter: CanvasViewerErrorRouter;
  private destroyed = false;
  private selectionChangeListener: (() => void) | null = null;
  private selectionContextKey = 'null';
  private elementClickListener: ((event: MouseEvent) => void) | null = null;
  private contextMenuListener: ((event: MouseEvent) => void) | null = null;
  private elementContext: PptxElementContext | null = null;
  private elementHitGeneration = 0;
  private readonly elementHitTolerance: number;
  private readonly _loadingLayer: HTMLSpanElement;
  private _layoutUnsubscribe: (() => void) | null = null;
  private readonly _layoutWaiters = new Set<() => void>();
  private _layoutFailed = false;
  private _navigationGeneration = 0;
  private _renderProgressGeneration = 0;
  private _lastReportedSlide = -1;
  private _lastReportedTotal = -1;
  private _lastReportedAvailable = -1;
  private _lastReportedLayoutComplete: boolean | null = null;
  /**
   * Create a Viewer that borrows an already-loaded presentation.
   *
   * The presentation's render mode is authoritative. The returned Viewer
   * cannot load another source, and destroying it leaves the caller-owned
   * presentation open. Call {@link goToSlide} to render the initial slide.
   */
  static fromPresentation(
    canvas: HTMLCanvasElement,
    presentation: PptxPresentation,
    opts: Omit<PptxViewerOptions, keyof LoadOptions> = {},
  ): Omit<PptxViewer, 'load'> {
    return new PptxViewer(canvas, {
      ...opts,
      [borrowedPresentationOption]: presentation,
    } as InternalPptxViewerOptions);
  }

  constructor(canvas: HTMLCanvasElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;
    this.canvas = canvas;
    const borrowedPresentation = (opts as InternalPptxViewerOptions)[borrowedPresentationOption];
    this.borrowed = borrowedPresentation !== undefined;
    this._mode = resolveCanvasViewerMode('PptxViewer', opts.mode, borrowedPresentation);
    this.presentationOwner = new TerminalResourceOwner(
      'PptxViewer',
      borrowedPresentation ?? null,
      false,
    );
    const hostWindow = canvas.ownerDocument?.defaultView ??
      (typeof window !== 'undefined' ? window : null);
    if (!hostWindow) throw new Error('PptxViewer requires a canvas with an active Window');
    this.hostWindow = hostWindow;
    const elementHitTolerance = opts.elementHitTolerance ?? 6;
    if (!Number.isFinite(elementHitTolerance) || elementHitTolerance < 0) {
      throw new RangeError('elementHitTolerance must be a finite non-negative number.');
    }
    this.elementHitTolerance = elementHitTolerance;
    this._hiddenMode = opts.hiddenSlideMode ?? 'show';

    this.canvasMount = new CallerCanvasMount(canvas, {
      wrapperCssText: 'position:relative;display:inline-block;vertical-align:top;',
      forceDisplayBlock: true,
    });
    this.wrapper = this.canvasMount.wrapper;
    this.renderDispatcher = new StaticCanvasRenderDispatcher(
      canvas,
      this._mode === 'worker' && !opts.enableMediaPlayback,
    );
    this.errorRouter = new CanvasViewerErrorRouter('PptxViewer', opts.onError);
    const overlays = new CanvasOverlayHost(
      this.wrapper,
      opts.enableTextSelection === true,
      opts.enableElementSelection === true,
    );
    this.textLayer = overlays.textLayer;
    this.highlightLayer = overlays.highlightLayer;
    this.elementLayer = overlays.elementLayer;
    this._loadingLayer = this.wrapper.ownerDocument.createElement('span');
    this._loadingLayer.style.cssText = [
      'position:absolute',
      'inset:0',
      'display:none',
      'align-items:center',
      'justify-content:center',
      'background:rgba(255,255,255,0.72)',
      'pointer-events:none',
      'z-index:4',
    ].join(';');
    this._loadingLayer.setAttribute('role', 'status');
    this._loadingLayer.setAttribute('aria-live', 'polite');
    this._loadingLayer.setAttribute('aria-label', 'Loading slide');
    const progress = this.wrapper.ownerDocument.createElement('progress');
    progress.setAttribute('aria-hidden', 'true');
    this._loadingLayer.appendChild(progress);
    this.wrapper.insertBefore(this._loadingLayer, this.elementLayer);
    if (this.textLayer && (opts.onSelectionContextChange || opts.enableElementSelection)) {
      this.selectionChangeListener = () => this._emitSelectionContextChange();
      this.wrapper.ownerDocument.addEventListener('selectionchange', this.selectionChangeListener);
    }
    if (opts.enableElementSelection) {
      this.elementClickListener = (event) => {
        void this._onElementClick(event).catch((error) => this._reportRenderError(error));
      };
      this.wrapper.addEventListener('click', this.elementClickListener);
    }
    if (opts.onContextMenu) {
      this.contextMenuListener = (event) => this._onContextMenu(event);
      this.wrapper.addEventListener('contextmenu', this.contextMenuListener);
    }

    this._find = new PptxFindController(
      () => this.slideCount,
      (slide) => this._collectSlideRuns(slide),
    );
    if (borrowedPresentation) this._bindLayoutPresentation(borrowedPresentation);
  }

  /**
   * Load a PPTX from URL or ArrayBuffer and render the first slide.
   *
   * Parse, load, and initial-render failures always reject this Promise.
   * `onError` is reserved for later Viewer-managed work that has no directly
   * awaitable method result, so one failure is never delivered twice.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this.destroyed) throw new Error('PptxViewer is destroyed');
    if (this.borrowed) {
      throw new Error(
        'PptxViewer.load() is unsupported on a Viewer created by fromPresentation(); ' +
          'the borrowed presentation is already loaded.',
      );
    }
    // SC20 atomic swap: retain the previous engine locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // engine's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current engine + its rendered slide intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old engine is freed the moment the new model arrives).
    let selectionInvalidated = false;
    try {
      const engine = await this.presentationOwner.replace(() => PptxPresentation.load(source, {
        password: this.opts.password,
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
        resourceLimits: this.opts.resourceLimits,
        debug: this.opts.debug,
        onResourceMetrics: this.opts.onResourceMetrics,
        workerTimeoutMs: this.opts.workerTimeoutMs,
        wasmUrl: this.opts.wasmUrl,
        math: this.opts.math,
        threeD: this.opts.threeD,
        regionMap: this.opts.regionMap,
        chartEx: this.opts.chartEx,
        mode: this._mode,
        progressiveLayout: this.opts.progressiveLayout,
        onLayoutProgress: this.opts.onLayoutProgress,
        onLayoutPartial: this.opts.onLayoutPartial,
        onLayoutComplete: this.opts.onLayoutComplete,
      }), () => {
        // Retire old-engine hit promises before install() destroys that engine:
        // a worker bridge may reject them synchronously during destroy, and its
        // microtask must already observe the new selection generation.
        this._invalidateElementSelection(false);
        selectionInvalidated = true;
        this.renderDispatcher.begin();
        this._invalidateFind();
        this.handle?.destroy();
        this.handle = null;
        this._unbindLayoutPresentation();
      });
      if (!engine) return;
      if (this.destroyed) throw new Error('PptxViewer is destroyed');
      // The loaded presentation is a new selection surface. Invalidate both a
      // retained element focus and every hit-test promise issued against the
      // previous engine before rendering the replacement deck.
      // Discard the stale slide's media handle before swapping engines so its RAF
      // loop / object URLs don't outlive the replaced presentation.
      this._bindLayoutPresentation(engine);
      const navigationGeneration = this._beginNavigation();
      this.currentSlide = await this._initialSlide(navigationGeneration);
      if (navigationGeneration !== this._navigationGeneration || engine !== this.engine) return;
      this._renderedSlide = -1;
      // A new presentation invalidates any prior find state.
      this._invalidateFind();
      await this.renderCurrentSlide();
    } catch (err) {
      if (this.destroyed) throw new Error('PptxViewer is destroyed');
      throw err instanceof Error ? err : new Error(String(err));
    }
    // Consumer selection callbacks are outside the engine/render error path and
    // cannot prevent the replacement deck from being committed and rendered.
    if (selectionInvalidated && !this.destroyed) this._emitSelectionContextChange();
  }

  /** Navigate to a specific slide (0-indexed). */
  async goToSlide(index: number): Promise<void> {
    const generation = this._beginNavigation();
    await this._goToSlide(index, generation);
  }

  private async _goToSlide(index: number, generation: number): Promise<void> {
    if (generation !== this._navigationGeneration) return;
    if (!this.engine || this.slideCount === 0) return;
    const next = Math.max(0, Math.min(index, this.slideCount - 1));
    const changed = next !== this.currentSlide;
    if (changed) this._invalidateElementSelection(false);
    this.currentSlide = next;
    await this.renderCurrentSlide();
    // Navigation and rendering complete before application notification; a
    // consumer callback failure cannot strand the Viewer on the prior slide.
    if (changed && !this.destroyed) this._emitSelectionContextChange();
  }

  async nextSlide(): Promise<void> {
    const generation = this._beginNavigation();
    const next = await this._step(1, generation);
    await this._goToSlide(next, generation);
  }

  async prevSlide(): Promise<void> {
    const generation = this._beginNavigation();
    const next = await this._step(-1, generation);
    await this._goToSlide(next, generation);
  }

  /** Next index for sequential nav: skip mode jumps over hidden slides. */
  private async _step(dir: 1 | -1, generation: number): Promise<number> {
    const engine = this.engine;
    const from = this.currentSlide;
    if (this._hiddenMode !== 'skip' || !engine) return from + dir;
    for (let i = from + dir; i >= 0 && i < this.slideCount; i += dir) {
      if (i >= engine.availableSlideCount) {
        if (!engine.layoutComplete) this._setLoading(true);
        if (!await this._waitForSlide(engine, i, () => generation === this._navigationGeneration)) {
          return from;
        }
      }
      if (!engine.isHidden(i)) return i;
    }
    return from;
  }

  /** Initial slide for load() / mode switch: skip mode lands on a visible one. */
  private async _initialSlide(generation: number): Promise<number> {
    const engine = this.engine;
    if (this._hiddenMode !== 'skip' || !engine || this.slideCount === 0) return 0;
    if (
      engine.availableSlideCount === 0 &&
      !await this._waitForSlide(engine, 0, () => generation === this._navigationGeneration)
    ) return 0;
    if (!engine.isHidden(0)) return 0;
    const forward = await this._step(1, generation);
    return forward !== 0 ? forward : 0;
  }

  /** Resolved `'dim'` overlay (defaults merged with the `hiddenSlideDim` option). */
  private _dim(): DimOptions {
    return {
      color: this.opts.hiddenSlideDim?.color ?? DEFAULT_HIDDEN_DIM.color,
      opacity: this.opts.hiddenSlideDim?.opacity ?? DEFAULT_HIDDEN_DIM.opacity,
    };
  }

  /**
   * Switch the hidden-slide mode at runtime and re-render. Entering `'skip'`
   * while on a hidden slide advances to the nearest visible slide.
   */
  async setHiddenSlideMode(mode: HiddenSlideMode): Promise<void> {
    const generation = this._beginNavigation();
    this._hiddenMode = mode;
    let next = this.currentSlide;
    if (mode === 'skip' && this.engine) {
      const engine = this.engine;
      if (
        this.currentSlide >= engine.availableSlideCount &&
        !await this._waitForSlide(engine, this.currentSlide, () => generation === this._navigationGeneration)
      ) return;
      if (engine.isHidden(this.currentSlide)) {
        next = await this._step(1, generation);
        if (next === this.currentSlide) next = await this._step(-1, generation);
      }
    }
    if (generation !== this._navigationGeneration) return;
    const changed = next !== this.currentSlide;
    if (changed) this._invalidateElementSelection(false);
    this.currentSlide = next;
    await this.renderCurrentSlide();
    if (changed && !this.destroyed) this._emitSelectionContextChange();
  }

  /** The current hidden-slide mode. */
  get hiddenSlideMode(): HiddenSlideMode { return this._hiddenMode; }

  /** Number of non-hidden slides (absolute `slideCount` is unchanged). During
   * progressive loading this is provisional until {@link layoutComplete}. */
  get visibleSlideCount(): number {
    if (!this.engine) return 0;
    const engine = this.engine;
    return countVisible((i) => engine.isHidden(i), this.slideCount);
  }

  get slideIndex(): number { return this.currentSlide; }
  get slideCount(): number { return this.engine?.slideCount ?? 0; }
  /** Number of opening slides currently paintable under progressive layout. */
  get availableSlideCount(): number { return this.engine?.availableSlideCount ?? this.slideCount; }
  /** Whether all slides are paintable. */
  get layoutComplete(): boolean { return this.engine?.layoutComplete ?? true; }
  /** Wait until all slides are paintable. */
  async waitUntilLayoutComplete(): Promise<void> {
    await this.errorRouter.ownBackgroundLifecycle(async () => {
      await this.engine?.waitUntilLayoutComplete?.();
    });
  }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5). Passthrough to {@link PptxPresentation.getNotes}:
   * 0-based index, returns `null` when the slide has no notes part, the index
   * is out of range, or nothing is loaded yet. During progressive loading the
   * answer is authoritative only below {@link availableSlideCount}; await
   * {@link waitUntilLayoutComplete} before scanning the whole deck.
   */
  getNotes(slideIndex: number): string | null {
    return this.engine?.getNotes(slideIndex) ?? null;
  }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement { return this.canvas; }

  // ─── IX9 zoom contract (ZoomableViewer) ───────────────────────────────────

  /** Natural (100%) CSS-px width of a slide — `slideWidth(EMU) / EMU_PER_PX`.
   *  0 when nothing is loaded. The scale-1 reference every zoom factor
   *  multiplies. */
  private _naturalWidthPx(): number {
    const emu = this.engine?.slideWidth ?? 0;
    return emu > 0 ? emu / EMU_PER_PX : 0;
  }

  /**
   * The width (CSS px) the render paths draw the slide at, honouring the zoom
   * state. `_scale === null` (no zoom method ever called) ⇒ the pre-IX9 value
   * `opts.width ?? (canvas.offsetWidth || 960)` verbatim (byte-identical
   * default). Once a factor latched ⇒ `naturalWidth × scale` (rounded), so the
   * slide is exactly `scale ×` its natural size regardless of `opts.width`.
   */
  private _targetWidth(): number {
    if (this._scale === null) return this.opts.width ?? (this.canvas.offsetWidth || 960);
    const natural = this._naturalWidthPx();
    if (natural <= 0) return this.opts.width ?? (this.canvas.offsetWidth || 960);
    return Math.round(natural * this._scale);
  }

  /** IX9 {@link ZoomableViewer} — the current zoom factor (`1` = 100%). Before
   *  any zoom method is called this is the EFFECTIVE scale implied by the render
   *  width: `targetWidth / naturalWidth`, or `1` when nothing is loaded. */
  getScale(): number {
    if (this._scale !== null) return this._scale;
    const natural = this._naturalWidthPx();
    if (natural <= 0) return 1;
    return this._targetWidth() / natural;
  }

  private _zoomMin(): number { return this.opts.zoomMin ?? 0.1; }
  private _zoomMax(): number { return this.opts.zoomMax ?? 4; }

  /**
   * IX9 {@link ZoomableViewer} — set the absolute zoom factor (`1` = 100% = the
   * slide at its natural EMU→px width), clamped to `[zoomMin, zoomMax]`, and
   * re-render the current slide at the new size. Fires `onScaleChange` when the
   * clamped factor actually changes. Resolves once the re-render settles.
   */
  async setScale(scale: number): Promise<void> {
    const next = clampScale(scale, this._zoomMin(), this._zoomMax());
    const changed = next !== this.getScale();
    this._scale = next;
    await this.renderCurrentSlide();
    if (changed) this.opts.onScaleChange?.(next);
  }

  /** IX9 {@link ZoomableViewer} — step up to the next rung of the shared zoom
   *  ladder (clamped to `zoomMax`). */
  async zoomIn(): Promise<void> { await this.setScale(nextZoomStep(this.getScale())); }

  /** IX9 {@link ZoomableViewer} — step down to the next lower ladder rung. */
  async zoomOut(): Promise<void> { await this.setScale(prevZoomStep(this.getScale())); }

  /**
   * IX9 {@link ZoomableViewer} — fit the current slide's WIDTH to the host
   * container (the element the canvas lives in), then re-render. Defers (no-op)
   * when nothing is loaded or the container is unlaid-out. Routes through
   * {@link setScale}.
   */
  async fitWidth(): Promise<void> { await this._fit('width'); }

  /**
   * IX9 {@link ZoomableViewer} — fit the WHOLE current slide (width and height)
   * inside the container so it is fully visible; takes the tighter of the
   * width/height fit. Defers when unloaded / unlaid-out.
   */
  async fitPage(): Promise<void> { await this._fit('page'); }

  /** Shared fit for {@link fitWidth}/{@link fitPage}: measure the natural slide
   *  size + the container box, ask core's pure `fitScale`, apply via setScale. */
  private async _fit(mode: 'width' | 'page'): Promise<void> {
    if (!this.engine) return;
    const container = this.wrapper.parentElement;
    if (!container) return;
    const scale = fitScale(
      {
        contentWidth: this.engine.slideWidth / EMU_PER_PX,
        contentHeight: this.engine.slideHeight / EMU_PER_PX,
        containerWidth: container.clientWidth,
        containerHeight: container.clientHeight,
      },
      mode,
    );
    if (scale <= 0) return; // unlaid-out / empty — defer
    await this.setScale(scale);
  }

  private async renderCurrentSlide(): Promise<void> {
    const engine = this.engine;
    if (!engine) return;
    const slide = this.currentSlide;
    const progressGeneration = ++this._renderProgressGeneration;
    this._setLoading(slide >= this.availableSlideCount && !this.layoutComplete);
    const generation = this.renderDispatcher.begin();
    try {
      if (slide >= engine.availableSlideCount) {
        const ready = await this._waitForSlide(
          engine,
          slide,
          () => progressGeneration === this._renderProgressGeneration &&
            this.renderDispatcher.isCurrent(generation) &&
            engine === this.engine &&
            slide === this.currentSlide,
        );
        if (!ready) return;
      }
      const dim = this._hiddenMode === 'dim' && engine.isHidden(slide)
        ? this._dim()
        : undefined;
      const targetWidth = this._targetWidth();
      const dpr = this.opts.dpr ?? (window.devicePixelRatio || 1);
      const scale = targetWidth / engine.slideWidth;
      const cssHeight = Math.round(engine.slideHeight * scale);
      this.canvas.style.width = `${targetWidth}px`;
      this.canvas.style.height = `${cssHeight}px`;

      this.handle?.destroy();
      this.handle = null;

      const isWorker = this._mode === 'worker';
      // Collect runs unconditionally (not just when a text layer exists): the
      // find-highlight overlay needs the current slide's run geometry too, and
      // caching them lets find() reuse the visible render for this slide. IX6 —
      // in worker mode the runs ride back beside the bitmap (via the proxy's
      // `onTextRun`), so both modes populate the same `runs` array.
      const runs: PptxTextRunInfo[] = [];
      const onTextRun = (r: PptxTextRunInfo) => runs.push(r);

      if (this.opts.enableMediaPlayback) {
        // presentSlide supports both modes (worker: base off-thread, video
        // overlay composited on the main thread).
        const handle = await engine.presentSlide(this.canvas, slide, {
          width: targetWidth,
          dpr,
          dim,
          onTextRun,
          onError: (error) => {
            if (this.renderDispatcher.isCurrent(generation)) this._reportRenderError(error);
          },
        });
        if (!this.renderDispatcher.isCurrent(generation)) {
          handle.destroy();
          return;
        }
        this.handle = handle;
      } else if (isWorker) {
        const bmp = await renderPptxFocusedSlide(
          this.engine,
          this.canvas,
          slide,
          'worker',
          { width: targetWidth, dpr, dim, onTextRun },
        );
        if (!this.renderDispatcher.commitBitmap(generation, bmp)) return;
      } else {
        await renderPptxFocusedSlide(
          this.engine,
          this.canvas,
          slide,
          'main',
          { width: targetWidth, dpr, onTextRun, dim },
        );
        if (!this.renderDispatcher.isCurrent(generation)) return;
      }
      this._renderedSlide = slide;
      this._emitSlideChange(true);
      // IX6 — identical overlay build for both modes: the run geometry the worker
      // shipped is the same shape `onTextRun` emits in main mode.
      if (this.textLayer) this._buildTextLayer(this.textLayer, runs, targetWidth, cssHeight);
      // Feed the just-rendered slide's runs to the find controller (geometry
      // matches what was drawn) and (re)draw its highlights.
      this._find.setSlideRuns(slide, runs);
      this._buildHighlightLayer(runs, targetWidth, cssHeight);
    } catch (err) {
      // Superseded paint failures are stale, but a same-presentation terminal
      // layout failure still belongs to the public navigation Promise that was
      // waiting for it. Do not turn cancellation into silent data loss.
      if (!this.renderDispatcher.isCurrent(generation) &&
          !(engine === this.engine && this._layoutFailed)) return;
      throw err;
    } finally {
      if (progressGeneration === this._renderProgressGeneration) this._setLoading(false);
    }
  }

  private _bindLayoutPresentation(presentation: PptxPresentation): void {
    this._unbindLayoutPresentation();
    this._layoutFailed = false;
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
    this._navigationGeneration++;
    this._renderProgressGeneration++;
    this._wakeLayoutWaiters();
    this._setLoading(false);
  }

  /** Supersede every pending navigation and wake its availability wait now. */
  private _beginNavigation(): number {
    const generation = ++this._navigationGeneration;
    this._wakeLayoutWaiters();
    return generation;
  }

  private _onLayoutPublication(
    presentation: PptxPresentation,
    publication: PptxLayoutPublication,
  ): void {
    if (this.destroyed || presentation !== this.engine) return;
    this._wakeLayoutWaiters();
    if (publication.error !== undefined) {
      this._layoutFailed = true;
      this.errorRouter.reportBackground(
        publication.error,
        this.opts.onLayoutComplete !== undefined,
      );
      return;
    }
    if (this._renderedSlide !== this.currentSlide) return;
    this._emitSlideChange();
  }

  private async _waitForSlide(
    presentation: PptxPresentation,
    slide: number,
    isCurrent: () => boolean,
  ): Promise<boolean> {
    return await this.errorRouter.ownBackgroundLifecycle(async () => {
      while (
        !this.destroyed &&
        isCurrent() &&
        presentation === this.engine &&
        slide >= presentation.availableSlideCount &&
        !presentation.layoutComplete &&
        !this._layoutFailed
      ) {
        await new Promise<void>((resolve) => this._layoutWaiters.add(resolve));
      }
      if (this.destroyed || presentation !== this.engine) return false;
      if (presentation.layoutComplete || this._layoutFailed) {
        await presentation.waitUntilLayoutComplete?.();
      }
      if (!isCurrent()) return false;
      return slide < presentation.availableSlideCount;
    });
  }

  private _wakeLayoutWaiters(): void {
    for (const resolve of this._layoutWaiters) resolve();
    this._layoutWaiters.clear();
  }

  private _emitSlideChange(force = false): void {
    const total = this.slideCount;
    const available = this.availableSlideCount;
    const complete = this.layoutComplete;
    if (!force &&
      this.currentSlide === this._lastReportedSlide &&
      total === this._lastReportedTotal &&
      available === this._lastReportedAvailable &&
      complete === this._lastReportedLayoutComplete
    ) return;
    this._lastReportedSlide = this.currentSlide;
    this._lastReportedTotal = total;
    this._lastReportedAvailable = available;
    this._lastReportedLayoutComplete = complete;
    this.opts.onSlideChange?.(this.currentSlide, total, complete);
  }

  private _setLoading(loading: boolean): void {
    this._loadingLayer.style.display = loading ? 'flex' : 'none';
  }

  /** Draw the find-highlight boxes for the current slide from its runs. */
  private _buildHighlightLayer(runs: PptxTextRunInfo[], cssWidth: number, cssHeight: number): void {
    const layer = this.highlightLayer;
    if (!layer) return;
    const highlights: PptxHighlightMatch[] = this._find.slideHighlights(this.currentSlide);
    buildPptxHighlightLayer(
      layer,
      runs,
      highlights,
      cssWidth,
      cssHeight,
      (font) => this._measureForFont(font),
      this.opts.findHighlightColors,
    );
  }

  /** A width-measurer primed with `font`, backed by a private 1×1 canvas. */
  private _measureForFont(font: string): (s: string) => number {
    if (!this._measureCtx) {
      const c = document.createElement('canvas');
      this._measureCtx = c.getContext('2d');
    }
    const ctx = this._measureCtx;
    if (!ctx) return (s) => s.length;
    ctx.font = font;
    return (s) => ctx.measureText(s).width;
  }

  /** IX6 — collect a slide's runs for search without touching the visible
   *  canvas. Delegates to `collectSlideRuns`, which works in BOTH modes (worker:
   *  off-thread, ships only the runs; main: throwaway offscreen canvas). Used for
   *  slides other than the one on screen. */
  private async _collectSlideRuns(slide: number): Promise<PptxTextRunInfo[]> {
    if (!this.engine) return [];
    // IX9 — collect at the zoom-aware width so the harvested geometry matches
    // what a navigation to that slide would draw at the current scale.
    return this.engine.collectSlideRuns(slide, this._targetWidth());
  }

  /**
   * IX2 — find every occurrence of `query` across all slides and highlight them
   * (a soft box per match on the highlight overlay). Returns every match in
   * document order, each tagged with its `{ slide }` (0-based). Case-insensitive
   * by default; pass `{ caseSensitive: true }` for an exact match.
   *
   * Scans all slides (each rendered once offscreen to read its text; the visible
   * slide reuses its on-screen render). IX6 — works in BOTH `mode: 'main'` and
   * `mode: 'worker'`: in worker mode each slide's run geometry is collected
   * off-thread and shipped back, so find returns the same matches on the same
   * code path. An empty query clears the find.
   */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<PptxMatchLocation>[]> {
    const engine = this.engine;
    if (!engine) return [];
    const generation = ++this._findGeneration;
    if (query.length === 0) {
      this._find.invalidate();
      this._redrawHighlights();
      return [];
    }
    if (!engine.layoutComplete) {
      await this.errorRouter.ownBackgroundLifecycle(
        () => engine.waitUntilLayoutComplete(),
      );
    }
    if (this.destroyed || generation !== this._findGeneration || engine !== this.engine) return [];
    const matches = await this.errorRouter.ownAwaitable(() => this._find.find(query, opts));
    if (this.destroyed || generation !== this._findGeneration || engine !== this.engine) return [];
    this._redrawHighlights();
    return matches;
  }

  /**
   * IX2 — move to the next match (wrap-around), navigating to its slide if
   * needed, and draw it in the active-match colour. Returns the now-active
   * match, or `null` when there are none. Call {@link findText} first.
   */
  async findNext(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** IX2 — move to the previous match (wrap-around). */
  async findPrev(): Promise<FindMatch<PptxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
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
    if (!match) {
      this._redrawHighlights();
      return null;
    }
    if (match.location.slide !== this.currentSlide) {
      // goToSlide re-renders, rebuilding the highlight layer for the new slide.
      await this.goToSlide(match.location.slide);
    } else {
      this._redrawHighlights();
    }
    return match;
  }

  /** Rebuild the highlight overlay for the current slide from cached runs. */
  private _redrawHighlights(): void {
    const runs = this._find.slideRuns(this.currentSlide) ?? [];
    const targetWidth = this._targetWidth();
    const cssHeight = this.engine
      ? Math.round(this.engine.slideHeight * (targetWidth / this.engine.slideWidth))
      : 0;
    this._buildHighlightLayer(runs, targetWidth, cssHeight);
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: PptxTextRunInfo[], cssWidth: number, cssHeight: number): void {
    buildPptxTextLayer(
      layer, runs, cssWidth, cssHeight, this._hyperlinkHandler(), this.currentSlide,
    );
  }

  /**
   * IX1 — the click handler passed to the text-layer overlay, or `undefined` when
   * `enableHyperlinks` is `false`. This is the single gate that disables hyperlink
   * interactivity: {@link buildPptxTextLayer} renders link runs exactly like plain
   * runs when no handler is supplied, so no hit region, cursor, tooltip, listener,
   * or navigation is wired (a custom `onHyperlinkClick` is suppressed too). When
   * enabled, the returned handler dispatches through {@link _onHyperlinkClick}.
   */
  private _hyperlinkHandler(): ((target: HyperlinkTarget) => void) | undefined {
    if (this.opts.enableHyperlinks === false) return undefined;
    return (t) => this._onHyperlinkClick(t);
  }

  /**
   * IX1/IX-nav hyperlink click dispatch. An internal target is first *enriched*
   * with its resolved 0-based `slideIndex` (via
   * {@link PptxPresentation.resolveInternalTarget}, relative to the current
   * slide) so a jump verb / slide-part ref arrives already mapped — this is the
   * field that was previously always `undefined`. When the integrator supplies
   * `opts.onHyperlinkClick` it OWNS the (enriched) click and takes NO default
   * action. Otherwise the viewer's default policy applies: an external link
   * opens in a new tab via the shared, scheme-sanitised
   * {@link openExternalHyperlink}; an internal slide jump navigates via
   * {@link goToSlide} to the resolved index (a target that resolves to no
   * reachable slide is a safe no-op).
   */
  private _onHyperlinkClick(target: HyperlinkTarget): void {
    const enriched = this._resolveInternalSlideIndex(target);
    if (this.opts.onHyperlinkClick) {
      this.opts.onHyperlinkClick(enriched);
      return;
    }
    if (enriched.kind === 'external') {
      openExternalHyperlink(enriched.url, undefined, this.hostWindow);
      return;
    }
    if (enriched.slideIndex !== undefined) {
      void this.goToSlide(enriched.slideIndex).catch((error) => this._reportRenderError(error));
    }
  }

  /** Populate an internal {@link HyperlinkTarget}'s `slideIndex` from its `ref`
   *  (a `ppaction://hlinkshowjump?jump=…` verb resolved relative to the current
   *  slide, or a `../slides/slideN.xml` part target resolved through the stamped
   *  part-name map — no filename-suffix heuristic). Any already-set `slideIndex`
   *  is kept; an external target and an unresolvable ref pass through unchanged so
   *  the caller no-ops safely. */
  private _resolveInternalSlideIndex(target: HyperlinkTarget): HyperlinkTarget {
    if (target.kind !== 'internal' || target.slideIndex !== undefined) return target;
    const idx = this.engine?.resolveInternalTarget(target.ref, this.currentSlide);
    return idx === undefined ? target : { ...target, slideIndex: idx };
  }

  /** PD14 render-error contract: route a render failure to `onError`, or
   *  `console.error` when none is given (never fully silent), and never after
   *  teardown. Mirrors the scroll viewers' `_reportRenderError` so all three
   *  single-canvas viewers agree. */
  private _reportRenderError(err: unknown): void {
    this.errorRouter.report(err);
  }

  /** Latest content-free resource metrics for the loaded presentation. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this.engine) throw new Error('Presentation not loaded');
    return await this.engine.getResourceMetrics();
  }

  /** Return the current browser text selection with PPTX source locators. */
  getSelectionContext(options: PptxSelectionContextOptions = {}): PptxSelectionContext | null {
    if (this.destroyed) throw new Error('PptxViewer is destroyed');
    const text = this.textLayer
      ? readPptxTextSelectionContext(
          this.wrapper,
          this.wrapper.ownerDocument?.getSelection?.() ?? null,
          options,
        )
      : null;
    return text ?? (this.elementContext
      ? limitPptxElementContext(
          this.elementContext,
          options.maxTextCharacters,
        )
      : null);
  }

  private _emitSelectionContextChange(): void {
    const context = this.getSelectionContext();
    // Native text selection becomes the sole current focus. Do not resurrect a
    // previously clicked element when that browser selection later collapses.
    if (context?.kind === 'text') {
      this.elementHitGeneration++;
      this.elementContext = null;
      this._redrawElementOutline();
    }
    const key = JSON.stringify(context);
    if (key === this.selectionContextKey) return;
    this.selectionContextKey = key;
    this.opts.onSelectionContextChange?.(context ? structuredClone(context) : null);
  }

  private _setElementContext(context: PptxElementContext | null): void {
    this.elementContext = context ? structuredClone(context) : null;
    this._redrawElementOutline();
    this._emitSelectionContextChange();
  }

  private _invalidateElementSelection(notify = true): void {
    this.elementHitGeneration++;
    this.elementContext = null;
    this._redrawElementOutline();
    if (notify) this._emitSelectionContextChange();
  }

  private _redrawElementOutline(): void {
    const context = this.elementContext;
    const engine = this.engine;
    if (!context || !engine || context.slideIndex !== this.currentSlide) {
      renderCanvasElementOutline(this.elementLayer, null);
      return;
    }
    renderCanvasElementOutline(this.elementLayer, {
      x: context.bounds.x / engine.slideWidth,
      y: context.bounds.y / engine.slideHeight,
      width: context.bounds.width / engine.slideWidth,
      height: context.bounds.height / engine.slideHeight,
      rotation: context.bounds.rotation,
    });
  }

  private async _onElementClick(event: MouseEvent): Promise<void> {
    if (this.destroyed || event.defaultPrevented || event.button !== 0) return;
    await this._resolveContextAt(event);
  }

  private _onContextMenu(event: MouseEvent): void {
    let context: Promise<PptxSelectionContext | null> | undefined;
    this.opts.onContextMenu?.({
      originalEvent: event,
      getContext: () => context ??= this._resolveContextAt(event),
    });
  }

  private async _resolveContextAt(event: MouseEvent): Promise<PptxSelectionContext | null> {
    const engine = this.engine;
    if (this.destroyed || !engine) return null;
    if (this.textLayer && readPptxTextSelectionContext(
      this.wrapper,
      this.wrapper.ownerDocument?.getSelection?.() ?? null,
    )) {
      // selectionchange is task-delivered and may not have run yet. Establish
      // text precedence synchronously so an older pending element hit cannot
      // survive a select/click/collapse sequence in the same task.
      this._emitSelectionContextChange();
      return this.destroyed ? null : this.getSelectionContext();
    }
    if (!this.opts.enableElementSelection) return this.getSelectionContext();
    const rect = this.canvas.getBoundingClientRect();
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
    const generation = ++this.elementHitGeneration;
    const slideIndex = this.currentSlide;
    const point = {
      x: localX / rect.width * engine.slideWidth,
      y: localY / rect.height * engine.slideHeight,
    };
    let context: PptxElementContext | null;
    try {
      context = await engine.getElementContextAt(slideIndex, point, {
        tolerance: this.elementHitTolerance / rect.width * engine.slideWidth,
        maxTextCharacters: MAX_ELEMENT_TEXT_CHARACTERS,
      });
    } catch (error) {
      if (this.destroyed || generation !== this.elementHitGeneration ||
        slideIndex !== this.currentSlide || engine !== this.engine) return null;
      throw error;
    }
    if (this.destroyed || generation !== this.elementHitGeneration ||
      slideIndex !== this.currentSlide || engine !== this.engine) return null;
    // Keep consumer callback exceptions outside the engine-error path. They are
    // application failures, not presentation/render failures.
    this._setElementContext(context);
    return this.destroyed ? null : this.getSelectionContext();
  }

  /**
   * Clean up the viewer and terminate the background worker.
   *
   * The caller-owned `<canvas>` is returned to the DOM position it held before
   * the constructor was called (same parent, same next-sibling) and its inline
   * `display` is restored, so the canvas can be reused — e.g. to construct a new
   * viewer on the same element. If the canvas was passed detached (no parent) it
   * is simply removed from the internal wrapper. Safe to call more than once.
   */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). Bump the load generation
    // too so a load() still in flight is treated as superseded and its engine is
    // cleaned up rather than installed onto a torn-down viewer.
    this.errorRouter.close();
    this.renderDispatcher.destroy();
    invalidatePptxRenderTarget(this.canvas);
    this.handle?.destroy();
    this.handle = null;
    this._unbindLayoutPresentation();
    this.presentationOwner.close();
    // IX2 — drop the find state (matches + cached runs) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer.
    this._invalidateFind();
    if (this.selectionChangeListener) {
      this.wrapper.ownerDocument.removeEventListener('selectionchange', this.selectionChangeListener);
      this.selectionChangeListener = null;
    }
    this.elementHitGeneration++;
    if (this.elementClickListener) {
      this.wrapper.removeEventListener('click', this.elementClickListener);
      this.elementClickListener = null;
    }
    if (this.contextMenuListener) {
      this.wrapper.removeEventListener('contextmenu', this.contextMenuListener);
      this.contextMenuListener = null;
    }
    this.elementContext = null;
    this.canvasMount.restore();
  }
}
