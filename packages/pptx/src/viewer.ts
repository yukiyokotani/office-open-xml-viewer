import type { RenderOptions, PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { buildPptxHighlightLayer, type PptxHighlightMatch } from './find-highlight-layer';
import { PptxFindController, type PptxMatchLocation } from './find';
import { PptxPresentation, type LoadOptions } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import { nextVisibleIndex, resolveVisibleIndex, countVisible } from './hidden';
import type { DimOptions } from './types';
import {
  type HyperlinkTarget,
  type FindMatch,
  type FindMatchesOptions,
  type ZoomableViewer,
  EMU_PER_PX,
  openExternalHyperlink,
  nextZoomStep,
  prevZoomStep,
  clampScale,
  fitScale,
} from '@silurus/ooxml-core';

/** How {@link PptxViewer} presents hidden slides (`<p:sld show="0">`). */
export type HiddenSlideMode = 'show' | 'skip' | 'dim';

/** Default `'dim'` overlay: 60% white (hidden content shows at 40%). */
const DEFAULT_HIDDEN_DIM: DimOptions = { color: '#ffffff', opacity: 0.6 };

export interface PptxViewerOptions extends RenderOptions, LoadOptions {
  /** Called when a slide finishes rendering */
  onSlideChange?: (index: number, total: number) => void;
  /** Called on parse or render errors */
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
  /**
   * IX9 explicit zoom factor (`1` = 100% = the slide at its natural EMU→px
   * width), or `null` when the caller has never invoked a zoom method. `null`
   * preserves the pre-IX9 render path EXACTLY: the slide renders at `opts.width`
   * (or `canvas.offsetWidth || 960` when unset), so default rendering is
   * byte-identical. The first zoom call latches a number here, after which
   * {@link _targetWidth} derives the render width from it.
   */
  private _scale: number | null = null;
  /** The canvas's DOM position BEFORE the constructor reparented it into
   *  {@link wrapper}, captured so {@link destroy} can return the caller-owned
   *  canvas to exactly where it was. `null` parent = canvas was passed
   *  detached. */
  private readonly _originalParent: Node | null;
  private readonly _originalNextSibling: Node | null;
  /** The canvas's inline `display` before the constructor forced `block`
   *  (empty string if it was unset), restored on {@link destroy}. */
  private readonly _originalDisplay: string;
  private textLayer: HTMLDivElement | null = null;
  /** IX2 — the find-highlight overlay layer (always created, above the text
   *  layer, `pointer-events:none`). */
  private highlightLayer: HTMLDivElement | null = null;
  /** IX2 — find state (per-slide runs, matches, active cursor). */
  private _find: PptxFindController;
  /** Private 2d context for measuring highlight text (own 1×1 canvas). */
  private _measureCtx: CanvasRenderingContext2D | null = null;
  private engine: PptxPresentation | null = null;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;
  private _hiddenMode: HiddenSlideMode;
  private handle: PresentationHandle | null = null;
  private readonly _mode: 'main' | 'worker';
  /** The canvas's bitmaprenderer context, used only by the static worker-mode
   *  render path. The media-playback path keeps a 2d context (via presentSlide),
   *  so this is obtained only when worker mode renders without media playback. */
  private _bitmapCtx: ImageBitmapRenderingContext | null = null;
  /** Set by {@link destroy} (first line). Guards {@link _reportRenderError} so a
   *  render rejection that lands AFTER teardown is swallowed rather than surfaced
   *  to an `onError` / `console.error` on a dead viewer — parity with the scroll
   *  viewers' `_destroyed` flag. */
  private _destroyed = false;
  /**
   * Concurrent-load latch (generation token). Every {@link load} increments this
   * and captures the value; after its engine finishes loading it re-checks the
   * live value and BAILS (destroying its own just-loaded engine) if a newer
   * `load()` has since started. Without it, two overlapping `load(A)`/`load(B)`
   * calls race the WASM parse / worker init, and whichever RESOLVES last wins the
   * swap — even the stale `load(A)` resolving after `load(B)`; the loser's freshly
   * created engine (never installed, or installed then overwritten) then leaks its
   * worker + pinned WASM allocation. The latch composes with SC20: the check runs
   * AFTER the new engine loads but BEFORE the field assignment and
   * `previous?.destroy()`, so a superseded load never touches `this.engine` nor
   * frees the current (newer) engine. {@link destroy} also bumps it so a load in
   * flight at teardown is treated as superseded and its engine cleaned up.
   */
  private _loadGen = 0;

  constructor(canvas: HTMLCanvasElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;
    this.canvas = canvas;
    this._mode = opts.mode ?? 'main';
    this._hiddenMode = opts.hiddenSlideMode ?? 'show';

    const parent = canvas.parentElement;
    // Capture the canvas's DOM position and inline display BEFORE reparenting so
    // destroy() can put the caller-owned canvas back exactly where it was.
    this._originalParent = parent;
    this._originalNextSibling = canvas.nextSibling;
    this._originalDisplay = canvas.style.display;
    this.wrapper = document.createElement('div');
    // vertical-align:top removes the inline-block baseline descender gap that
    // otherwise lets the host container's background show through below the
    // canvas (~6 px on default font metrics).
    this.wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    // Force `display:block` on the canvas so it does not inherit the inline
    // baseline of an enclosing wrapper, which would otherwise leave a 4–6px
    // descender gap between the canvas bottom and the wrapper bottom — the
    // host container's background would show through that strip.
    if (!canvas.style.display) canvas.style.display = 'block';
    if (parent) parent.insertBefore(this.wrapper, canvas);
    this.wrapper.appendChild(canvas);

    // Static worker-mode rendering paints worker-produced bitmaps via a
    // bitmaprenderer context (grabbed once — a canvas holds one context type for
    // its lifetime). The media-playback path uses presentSlide, which keeps a 2d
    // context, so skip bitmaprenderer there.
    if (this._mode === 'worker' && !opts.enableMediaPlayback) {
      this._bitmapCtx = canvas.getContext('bitmaprenderer');
    }

    if (opts.enableTextSelection) {
      this.textLayer = document.createElement('div');
      this.textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this.wrapper.appendChild(this.textLayer);
    }

    // IX2 — find-highlight overlay layer, appended last (stacks above the text
    // layer). `pointer-events:none` keeps selection + link clicks working
    // through it. IX6 — populated in BOTH render modes (worker mode ships the
    // run geometry back beside the bitmap).
    this.highlightLayer = document.createElement('div');
    this.highlightLayer.style.cssText =
      'position:absolute;top:0;left:0;width:100%;height:100%;overflow:hidden;pointer-events:none;';
    this.wrapper.appendChild(this.highlightLayer);

    this._find = new PptxFindController(
      () => this.slideCount,
      (slide) => this._collectSlideRuns(slide),
    );
  }

  /**
   * Load a PPTX from URL or ArrayBuffer and render the first slide.
   *
   * Error contract (shared by all three viewers):
   * - Parse/load failure (the underlying `PptxPresentation.load()` call itself
   *   rejects): if an `onError` callback was provided it is invoked and `load`
   *   resolves normally; if not, the error is rethrown so it is never silently
   *   swallowed.
   * - Render failure (the first slide fails to draw AFTER a successful
   *   parse/load): routed to the shared `_reportRenderError` contract (`onError`
   *   if provided, else `console.error` — never silent) and `load` still
   *   RESOLVES, matching every subsequent navigation call.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    // SC20 atomic swap: retain the previous engine locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // engine's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current engine + its rendered slide intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old engine is freed the moment the new model arrives).
    const gen = ++this._loadGen;
    const previous = this.engine;
    try {
      const engine = await PptxPresentation.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
        workerTimeoutMs: this.opts.workerTimeoutMs,
        wasmUrl: this.opts.wasmUrl,
        math: this.opts.math,
        mode: this._mode,
      });
      if (gen !== this._loadGen) {
        // A newer load() (or destroy()) started while this one was in flight — we
        // lost the concurrent-load race. Destroy the engine we just loaded (it was
        // never installed) and leave the winning load's engine + SC20 swap
        // untouched: do NOT touch `this.engine`/`this.handle` and do NOT destroy
        // `previous` (irrelevant to the winner; possibly already stale).
        engine.destroy();
        return;
      }
      // Discard the stale slide's media handle before swapping engines so its RAF
      // loop / object URLs don't outlive the replaced presentation.
      this.handle?.destroy();
      this.handle = null;
      this.engine = engine;
      previous?.destroy();
      this.currentSlide = this._initialSlide();
      // A new presentation invalidates any prior find state.
      this._find.invalidate();
      await this.renderCurrentSlide();
    } catch (err) {
      // Superseded loads own no error reporting — the winning load (or destroy())
      // is the outcome the caller awaits; swallow this stale rejection.
      if (gen !== this._loadGen) return;
      const e = err instanceof Error ? err : new Error(String(err));
      if (this.opts.onError) {
        this.opts.onError(e);
        return;
      }
      throw e;
    }
  }

  /** Navigate to a specific slide (0-indexed). */
  async goToSlide(index: number): Promise<void> {
    if (!this.engine || this.slideCount === 0) return;
    this.currentSlide = Math.max(0, Math.min(index, this.slideCount - 1));
    await this.renderCurrentSlide();
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this._step(1));
  }

  async prevSlide(): Promise<void> {
    await this.goToSlide(this._step(-1));
  }

  /** Next index for sequential nav: skip mode jumps over hidden slides. */
  private _step(dir: 1 | -1): number {
    if (this._hiddenMode === 'skip' && this.engine) {
      return nextVisibleIndex(this.currentSlide, dir, (i) => this.engine!.isHidden(i), this.slideCount);
    }
    return this.currentSlide + dir;
  }

  /** Initial slide for load() / mode switch: skip mode lands on a visible one. */
  private _initialSlide(): number {
    if (this._hiddenMode === 'skip' && this.engine) {
      return resolveVisibleIndex(0, (i) => this.engine!.isHidden(i), this.slideCount);
    }
    return 0;
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
    this._hiddenMode = mode;
    if (mode === 'skip' && this.engine) {
      this.currentSlide = resolveVisibleIndex(
        this.currentSlide,
        (i) => this.engine!.isHidden(i),
        this.slideCount,
      );
    }
    await this.renderCurrentSlide();
  }

  /** The current hidden-slide mode. */
  get hiddenSlideMode(): HiddenSlideMode { return this._hiddenMode; }

  /** Number of non-hidden slides (absolute `slideCount` is unchanged). */
  get visibleSlideCount(): number {
    if (!this.engine) return 0;
    const engine = this.engine;
    return countVisible((i) => engine.isHidden(i), this.slideCount);
  }

  get slideIndex(): number { return this.currentSlide; }
  get slideCount(): number { return this.engine?.slideCount ?? 0; }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5). Passthrough to {@link PptxPresentation.getNotes}:
   * 0-based index, returns `null` when the slide has no notes part, the index
   * is out of range, or nothing is loaded yet.
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
    if (!this.engine) return;
    const dim =
      this._hiddenMode === 'dim' && this.engine.isHidden(this.currentSlide)
        ? this._dim()
        : undefined;
    const targetWidth = this._targetWidth();
    const dpr = this.opts.dpr ?? (window.devicePixelRatio || 1);

    const scale = targetWidth / this.engine.slideWidth;
    const cssHeight = Math.round(this.engine.slideHeight * scale);
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

    try {
      if (this.opts.enableMediaPlayback) {
        // presentSlide supports both modes (worker: base off-thread, video
        // overlay composited on the main thread).
        this.handle = await this.engine.presentSlide(this.canvas, this.currentSlide, {
          width: targetWidth,
          dpr,
          dim,
          onTextRun,
        });
      } else if (isWorker) {
        const bmp = await this.engine.renderSlideToBitmap(this.currentSlide, { width: targetWidth, dpr, dim, onTextRun });
        this.canvas.width = bmp.width;
        this.canvas.height = bmp.height;
        this._bitmapCtx?.transferFromImageBitmap(bmp);
      } else {
        await this.engine.renderSlide(this.canvas, this.currentSlide, { width: targetWidth, dpr, onTextRun, dim });
      }
      this.opts.onSlideChange?.(this.currentSlide, this.slideCount);
    } catch (err) {
      this._reportRenderError(err);
    }

    // IX6 — identical overlay build for both modes: the run geometry the worker
    // shipped is the same shape `onTextRun` emits in main mode.
    if (this.textLayer) {
      this._buildTextLayer(this.textLayer, runs, targetWidth, cssHeight);
    }
    // Feed the just-rendered slide's runs to the find controller (geometry
    // matches what was drawn) and (re)draw its highlights.
    this._find.setSlideRuns(this.currentSlide, runs);
    this._buildHighlightLayer(runs, targetWidth, cssHeight);
  }

  /** Draw the find-highlight boxes for the current slide from its runs. */
  private _buildHighlightLayer(runs: PptxTextRunInfo[], cssWidth: number, cssHeight: number): void {
    const layer = this.highlightLayer;
    if (!layer) return;
    const highlights: PptxHighlightMatch[] = this._find.slideHighlights(this.currentSlide);
    buildPptxHighlightLayer(layer, runs, highlights, cssWidth, cssHeight, (font) =>
      this._measureForFont(font),
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
    if (!this.engine) return [];
    const matches = await this._find.find(query, opts);
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
    this._find.invalidate();
    this._redrawHighlights();
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
    buildPptxTextLayer(layer, runs, cssWidth, cssHeight, this._hyperlinkHandler());
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
      openExternalHyperlink(enriched.url);
      return;
    }
    if (enriched.slideIndex !== undefined) void this.goToSlide(enriched.slideIndex);
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
    if (this._destroyed) return;
    const e = err instanceof Error ? err : new Error(String(err));
    if (this.opts.onError) this.opts.onError(e);
    else console.error('[ooxml] PptxViewer render failed:', e);
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
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). Bump the load generation
    // too so a load() still in flight is treated as superseded and its engine is
    // cleaned up rather than installed onto a torn-down viewer.
    this._destroyed = true;
    this._loadGen++;
    this.handle?.destroy();
    this.handle = null;
    this.engine?.destroy();
    // IX2 — drop the find state (matches + cached runs) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer.
    this._find.invalidate();
    // Return the caller-owned canvas to its original DOM slot before discarding
    // the wrapper. insertBefore still works if the original parent was itself
    // detached; when there was no original parent the canvas is left detached
    // (just pulled out of the wrapper). The recorded next-sibling may have been
    // removed or moved by the caller since construction — insertBefore throws
    // NotFoundError for a reference that is no longer a child of the parent, so
    // fall back to appending at the end in that case.
    if (this._originalParent) {
      const ref =
        this._originalNextSibling && this._originalNextSibling.parentNode === this._originalParent
          ? this._originalNextSibling
          : null;
      this._originalParent.insertBefore(this.canvas, ref);
    } else if (this.canvas.parentNode) {
      this.canvas.parentNode.removeChild(this.canvas);
    }
    this.canvas.style.display = this._originalDisplay;
    this.wrapper.remove();
  }
}
