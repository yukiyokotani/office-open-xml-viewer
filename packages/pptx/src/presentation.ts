import type { DimOptions, PptxComment } from './types';
import { renderSlide, dropImageBitmapCache, type TextRunCallback, type PptxTextRunInfo } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import {
  buildSlidePartIndex,
  resolveInternalSlideTarget,
  type SlidePartNames,
} from './slide-nav';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  isHTMLCanvas,
  dropSvgImageCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  OoxmlResourceLimitError,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type OoxmlResourceMetrics,
  workerRendererDescriptors,
} from '@silurus/ooxml-core';
import {
  deserializeWorkerError,
  disposeRejectedLoad,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  HARD_MAX_PPTX_CACHED_SLIDES,
  HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  readLatestOoxmlResourceMetrics,
  parseResourceLimitError,
  PULL_SESSION_PROTOCOL,
  type NormalizedOoxmlResourcePolicy,
  type PullSessionResponse,
  type WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { PPTX_GOOGLE_FONTS } from './google-fonts';
import {
  findPreflightMimeType,
  normalizePresentationBootstrap,
  normalizePresentationPreflight,
  type PresentationPreflight,
} from './presentation-preflight';
import { PptxSlideRepository } from './slide-repository';
import {
  isPptxSlidePullResponse,
  PptxSlidePullClient,
} from './slide-pull-client';
import type {
  PptxWorkerRequest,
  PptxWorkerResponse,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';
import {
  hitTestPptxSlideContext,
  type PptxElementContextOptions,
  type PptxElementContext,
  type PptxSlidePoint,
} from './element-selection';

/** Options for {@link PptxPresentation.load}. */
export type LoadOptions = CoreLoadOptions & {
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link PptxPresentation.renderSlideToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   */
  mode?: 'main' | 'worker';
};

/** Options for {@link PptxPresentation.renderSlideToBitmap}. */
export interface RenderSlideToBitmapOptions {
  /** Slide width in CSS pixels. Defaults to 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio (workers have none). */
  dpr?: number;
  /**
   * Skip the static media play-badge so a live overlay can draw its own
   * controls. Used internally by {@link PptxPresentation.presentSlide}.
   * @internal
   */
  skipMediaControls?: boolean;
  /** Translucent overlay drawn over the finished slide (hidden-slide dimming). */
  dim?: DimOptions;
  /**
   * IX6 — receives the slide's text-run geometry (the same stream `renderSlide`
   * emits in main mode). Stays main-thread (never crosses the wire); in worker
   * mode the proxy invokes it with the runs the worker shipped back beside the
   * bitmap, so a caller builds the selection / find overlay on the SAME code
   * path in both modes.
   */
  onTextRun?: TextRunCallback;
}

/** Options for rendering a single slide onto a canvas. */
export interface RenderSlideOptions {
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: TextRunCallback;
  /**
   * Skip drawing the play badge overlay on media elements. Used internally by
   * {@link PptxPresentation.presentSlide} so its interactive handle can draw
   * its own play/pause chrome without duplication.
   */
  skipMediaControls?: boolean;
  /** Translucent overlay drawn over the finished slide (hidden-slide dimming). */
  dim?: DimOptions;
}

/** Options for {@link PptxPresentation.presentSlide}. */
export interface PresentSlideOptions extends Omit<RenderSlideOptions, 'skipMediaControls'> {
  /**
   * Called for embedded-media decode and playback failures that occur after
   * the presentation handle has been returned. Initial rendering and media
   * acquisition failures reject `presentSlide()` instead.
   */
  onError?: (error: Error) => void;
}

/**
 * Headless PPTX rendering engine.
 *
 * Parses `.pptx` archives in a background worker (WASM) but renders slides
 * synchronously on the main thread, so the canvas shares the document's
 * `FontFaceSet` — avoiding subtle wrap differences between system fallback
 * fonts and theme-declared webfonts (e.g. Nunito Sans).
 *
 * Construct via the static `load` factory. A single instance can drive any
 * number of canvases (scroll view, thumbnail grid, master-detail, etc.).
 *
 * @example
 * const pres = await PptxPresentation.load(buffer);
 * await pres.renderSlide(canvas, 0, { width: 960 });
 */
export class PptxPresentation {
  private _metrics: OoxmlResourceMetricsSession | null = null;
  private readonly _worker: Worker;
  private readonly _bridge: WorkerBridge<
    PptxWorkerResponse | RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>
  >;
  private _mode: 'main' | 'worker' = 'main';
  private _preflight: PresentationPreflight | null = null;
  private _slides: PptxSlideRepository | null = null;
  private _slidePullClient: PptxSlidePullClient | null = null;
  /** First fatal package/model violation for this presentation generation. */
  private _resourceFailure: OoxmlResourceLimitError | null = null;
  /** Lazily-built `partName → slide index` map for internal hyperlink slide
   *  jumps (IX-nav). Cleared on {@link destroy}; built on first
   *  {@link getSlideIndexByPartName}/{@link resolveInternalTarget} from the
   *  common compact preflight in either render mode. */
  private _slidePartIndex: Map<string, number> | null = null;
  /** One bounded retained-byte owner shared by images and media. */
  private readonly _rawParts = new BoundedRawPartCache({
    maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
    maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
  });
  /** Google-Fonts `FontFace` objects this deck preloaded into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with
   *  its own FontFaceSet). Released in {@link destroy} so they do not leak into
   *  the shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in
   *  core, so a web font shared with another open deck survives until both go). */
  private _googleFontFaces: FontFace[] = [];
  /** One stable closure per instance: the decoded-bitmap and SVG caches key on
   *  this identity to scope decodes per deck (so two open decks never swap
   *  images for a shared zip path like ppt/media/image1.png). Reusing the same
   *  reference across every render also lets those caches hit across slides. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImage(path, mime);
  private readonly _fetchMedia = (path: string): Promise<Blob> => this.getMedia(path);
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderSlide` / `presentSlide` reuses it — equations render when present,
   *  and are skipped when omitted. */
  private _math: MathRenderer | undefined;
  private _threeD: ChartThreeDRenderer | undefined;
  private _regionMap: ChartRegionMapRenderer | undefined;
  private _chartEx: ChartExRenderer | undefined;

  private constructor(worker: Worker, mode: 'main' | 'worker', wasmUrlOverride?: string | URL) {
    this._worker = worker;
    this._mode = mode;
    this._bridge = new WorkerBridge<
      PptxWorkerResponse | RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>
    >(this._worker, {
      // Every response carries an id (no `ready` handshake — the worker `await`s
      // its own init promise before each request, docx/xlsx pattern).
      correlate: (msg) =>
        'protocol' in msg && msg.protocol === PULL_SESSION_PROTOCOL
          ? msg.requestId
          : 'id' in msg
            ? msg.id
            : undefined,
      // Pull errors are correlated protocol values consumed by
      // BoundedPullSession; only ordinary worker errors reject here.
      toError: (msg) =>
        !('protocol' in msg) && msg.kind === 'error'
          ? deserializeWorkerError(msg)
          : undefined,
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
    this._bridge.post({ kind: 'init', wasmUrl } satisfies PptxWorkerRequest);
  }

  private _assertResourceHealthy(): void {
    if (this._resourceFailure) throw this._resourceFailure;
  }

  private _rethrowWithResourceFailure(error: unknown): never {
    const typed = error instanceof OoxmlResourceLimitError
      ? error
      : parseResourceLimitError(error);
    if (typed) {
      this._resourceFailure ??= typed;
      throw this._resourceFailure;
    }
    throw error;
  }

  /** Parse a PPTX from URL or ArrayBuffer. */
  static async load(
    source: string | ArrayBuffer,
    opts: LoadOptions = {},
  ): Promise<PptxPresentation> {
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'pptx',
      mode,
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    try {
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    // Resolve the container on the main thread — before spinning up the worker.
    // A normal ZIP passes through unchanged; an Agile-encrypted CFB is decrypted
    // when `opts.password` is supplied ([MS-OFFCRYPTO]); a password-protected
    // file without a password, or a legacy-binary / unknown CFB, becomes a typed
    // OoxmlError (whose `instanceof` would not survive the worker boundary).
    buffer = toArrayBuffer(await resolveOoxmlContainer(buffer, opts.password));
    metrics.setSourceBytes(buffer.byteLength);
    metrics.checkpoint('container ready');
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const rendererDescriptors = mode === 'worker' ? workerRendererDescriptors(opts) : undefined;
    let pres: PptxPresentation | undefined;
    try {
      pres = new PptxPresentation(worker, mode, opts.wasmUrl);
      pres._metrics = metrics;
      if (opts.math && mode === 'worker' && !rendererDescriptors?.math) {
        console.warn(
          "[ooxml] a custom math renderer cannot cross the worker boundary; equations will be skipped in mode: 'worker'. Use the math renderer from @silurus/ooxml/math.",
        );
      }
      if (opts.threeD && mode === 'worker' && !rendererDescriptors?.threeD) {
        console.warn(
          "[ooxml] a custom 3-D chart renderer cannot cross the worker boundary; charts use their 2-D family fallback in mode: 'worker'. Use the renderer from @silurus/ooxml/three-d.",
        );
      }
      pres._math = mode === 'worker' ? undefined : opts.math;
      pres._threeD = mode === 'worker' ? undefined : opts.threeD;
      if (opts.regionMap && mode === 'worker' && !rendererDescriptors?.regionMap) {
        console.warn(
          "[ooxml] a custom Region Map renderer cannot cross the worker boundary; geospatial charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/region-map.",
        );
      }
      pres._regionMap = mode === 'worker' ? undefined : opts.regionMap;
      if (opts.chartEx && mode === 'worker' && !rendererDescriptors?.chartEx) {
        console.warn(
          "[ooxml] a custom ChartEx renderer cannot cross the worker boundary; ChartEx charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/chart-ex.",
        );
      }
      pres._chartEx = mode === 'worker' ? undefined : opts.chartEx;
      await pres._parse(
        buffer,
        resourceOptions.policy,
        mode === 'worker' ? !!opts.useGoogleFonts : false,
        opts.workerTimeoutMs,
        (usage) => metrics.observeUsage(usage),
        rendererDescriptors,
      );
      metrics.checkpoint('presentation preflight ready');
      if (mode === 'main' && opts.useGoogleFonts && pres._preflight) {
        pres._googleFontFaces = await preloadGoogleFonts(
          pres._preflight.fontPreloadNames,
          PPTX_GOOGLE_FONTS,
        );
      }
      metrics.succeed({ slides: pres.slideCount });
      return pres;
    } catch (error) {
      const rejectedPresentation = pres;
      disposeRejectedLoad(worker, rejectedPresentation ? () => rejectedPresentation.destroy() : undefined);
      throw error;
    }
    } catch (error) {
      metrics.fail(error);
      throw error;
    }
  }

  private async _parse(
    buffer: ArrayBuffer,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    useGoogleFonts = false,
    timeoutMs?: number,
    onUsage?: (usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void,
    renderers?: WorkerRendererDescriptors,
  ): Promise<void> {
    const response = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ kind: 'parse', id, buffer, resourcePolicy, useGoogleFonts, renderers } satisfies RenderWorkerRequest)
          : ({ kind: 'parse', id, buffer, resourcePolicy } satisfies PptxWorkerRequest),
      [buffer],
      { timeoutMs },
    );
    if (this._mode === 'worker') {
      const ready = response as Extract<RenderWorkerResponse, { kind: 'presentationReady' }>;
      if (ready.usage) onUsage?.(ready.usage);
      this._preflight = normalizePresentationPreflight(
        (response as Extract<RenderWorkerResponse, { kind: 'presentationReady' }>).preflight,
      );
      return;
    }

    const bootstrap = normalizePresentationBootstrap(
      (response as Extract<PptxWorkerResponse, { kind: 'presentationOpened' }>).bootstrap,
    );
    this._slidePullClient = new PptxSlidePullClient({
      slideCount: bootstrap.slideCount,
      transport: this._bridge.transport(isPptxSlidePullResponse),
      open: async (slideIndex, identity, operationTimeoutMs) => {
        await this._bridge.request(
          (id) => ({
            kind: 'openSlideSession',
            id,
            slideIndex,
            ...identity,
          }) satisfies PptxWorkerRequest,
          undefined,
          { timeoutMs: operationTimeoutMs },
        );
      },
      onUsage,
    });

    // Preflight deliberately drains one transferred unit at a time without
    // decoding it in Window. The worker prepares compact facts from the same
    // unit and publishes them only when this consumer ACKs it.
    for (let slideIndex = 0; slideIndex < bootstrap.slideCount; slideIndex += 1) {
      await this._slidePullClient.load(slideIndex, false, timeoutMs);
    }
    const finished = await this._bridge.request(
      (id) => ({ kind: 'finishPresentationPreflight', id }) satisfies PptxWorkerRequest,
      undefined,
      { timeoutMs },
    );
    this._preflight = normalizePresentationPreflight(
      (finished as Extract<PptxWorkerResponse, { kind: 'presentationPreflightReady' }>).preflight,
    );
    this._slides = new PptxSlideRepository({
      slideCount: this._preflight.slideCount,
      maxCachedSlides: HARD_MAX_PPTX_CACHED_SLIDES,
      maxCachedStructuralBytes: HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
      loadSlide: async (slideIndex) => {
        const slide = await this._slidePullClient?.load(slideIndex, true);
        if (!slide) throw new Error('PPTX slide pull client is unavailable');
        return slide;
      },
    });
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._preflight?.slideCount ?? 0; }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._preflight?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._preflight?.slideHeight ?? 0; }

  /** The render mode this engine was loaded with ('main' | 'worker'). A fact for
   *  integrators and the scroll viewer: an injected engine's mode decides whether
   *  slides render via renderSlide (main) or renderSlideToBitmap (worker) — no
   *  probing (design §11: no silent mis-pathing). */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5 — Notes Slide). Returns the notes-body text as a single
   * string (paragraphs joined with `\n`), or `null` when the slide has no
   * notes part. The notes are parsed at {@link load} time, so this is a
   * synchronous lookup.
   *
   * `slideIndex` is 0-based. Unlike navigation methods it is *not* clamped:
   * an out-of-range or non-integer index returns `null` rather than the notes
   * of the nearest slide (so a tool iterating by index gets an honest "no
   * notes" instead of a duplicated neighbour).
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * for (let i = 0; i < pres.slideCount; i++) {
   *   const notes = pres.getNotes(i);
   *   if (notes) console.log(`Slide ${i + 1} notes:`, notes);
   * }
   */
  getNotes(slideIndex: number): string | null {
    return Number.isInteger(slideIndex)
      ? (this._preflight?.slides[slideIndex]?.notes ?? null)
      : null;
  }

  /** Read-only slide comments in authored order. Classic and modern comments
   * share this compact mode-independent projection; modern replies remain
   * nested under their root. Use it for fully custom UI, or opt into the
   * ScrollViewer's marker-and-card view. Returns `[]` for an invalid or
   * comment-free slide. */
  getComments(slideIndex: number): readonly Readonly<PptxComment>[] {
    return Number.isInteger(slideIndex)
      ? (this._preflight?.slides[slideIndex]?.comments ?? [])
      : [];
  }

  /**
   * Whether the slide at `slideIndex` (0-based, absolute) is marked hidden
   * (`<p:sld show="0">`, ECMA-376 §19.3.1.38). Like {@link getNotes} the index
   * is NOT clamped — out-of-range / non-integer ⇒ `false`. This is a *fact*
   * about the model; deciding what to do with a hidden slide (skip / dim) is the
   * caller's policy (see {@link PptxViewer}'s `hiddenSlideMode` modes).
   */
  isHidden(slideIndex: number): boolean {
    return Number.isInteger(slideIndex)
      ? (this._preflight?.slides[slideIndex]?.hidden ?? false)
      : false;
  }

  /** The compact preflight's per-slide `partName` array (`sldIdLst` order). */
  private _partNames(): SlidePartNames {
    return (this._preflight?.slides ?? []).map((slide) => slide.partName);
  }

  /** Lazily build (and cache) the `partName → index` map. Nulled by
   *  {@link destroy} so a reused reference never serves a stale deck's indices. */
  private _partIndex(): Map<string, number> {
    if (!this._slidePartIndex) this._slidePartIndex = buildSlidePartIndex(this._partNames());
    return this._slidePartIndex;
  }

  /**
   * Resolve a slide's OPC part name (e.g. `ppt/slides/slide3.xml`) to its
   * 0-based index in `sldIdLst` order, or `undefined` when no slide has that
   * part name. This is the map an internal hyperlink slide jump
   * (`<a:hlinkClick action="ppaction://hlinksldjump" r:id>`, ECMA-376
   * §21.1.2.3.5) resolves against: the click's rel Target names a slide part, and
   * this turns it into the index a viewer can navigate to. Works in both `main`
   * and `worker` mode through the same compact preflight contract.
   */
  getSlideIndexByPartName(partName: string): number | undefined {
    return this._partIndex().get(partName);
  }

  /**
   * Resolve an internal hyperlink target string to a 0-based slide index, or
   * `undefined` when it names no reachable slide. Handles both
   * `<a:hlinkClick @action>` classes (§21.1.2.3.5):
   *
   *   - a **relative** show jump — `ppaction://hlinkshowjump?jump=firstslide |
   *     lastslide | nextslide | previousslide` — resolved arithmetically from
   *     `currentIndex` (clamped at the deck ends);
   *   - a **specific** slide-part jump — `ppaction://hlinksldjump`, whose
   *     resolved target is a slide-rel part name like `../slides/slide3.xml` —
   *     resolved through {@link getSlideIndexByPartName}.
   *
   * `ref` is the internal reference a `HyperlinkTarget` of kind `'internal'`
   * carries: the raw `ppaction://…` action string for a relative jump, or the
   * resolved slide-part target string for a specific jump. A viewer's
   * `onHyperlinkClick` default calls this with `ref` and the current slide, then
   * navigates to the returned index.
   *
   * @param ref          the internal action/target string.
   * @param currentIndex the 0-based slide the jump is relative to (default 0).
   */
  resolveInternalTarget(ref: string, currentIndex = 0): number | undefined {
    return resolveInternalSlideTarget(ref, this._partIndex(), currentIndex);
  }

  /** Render a slide onto the given canvas. */
  async renderSlide(
    canvas: HTMLCanvasElement | OffscreenCanvas,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<void> {
    this._assertResourceHealthy();
    try {
      if (this._mode === 'worker') {
        throw new Error(
          "renderSlide(canvas) is unavailable in mode: 'worker'; use renderSlideToBitmap() and paint it via an ImageBitmapRenderingContext",
        );
      }
      const compact = this._preflight;
      const repository = this._slides;
      if (!compact || !repository) throw new Error('Presentation not loaded');
      const dpr = opts.dpr ?? defaultDpr();
      const width = opts.width ?? ((isHTMLCanvas(canvas) ? canvas.offsetWidth : 0) || 960);
      await repository.withSlide(slideIndex, (slide) => {
        // A render may have waited behind another consumer after its public
        // entrance check. Re-check the presentation poison at the ownership
        // boundary before a cached Slide becomes observable.
        this._assertResourceHealthy();
        return renderSlide(
          canvas,
          slide,
          compact.slideWidth,
          compact.slideHeight,
          {
            width,
            dpr,
            defaultTextColor: compact.defaultTextColor,
            majorFont: compact.majorFont,
            minorFont: compact.minorFont,
            hlinkColor: compact.hlinkColor,
            fetchMedia: this._fetchMedia,
            fetchImage: this._fetchImage,
            skipMediaControls: opts.skipMediaControls,
            dim: opts.dim,
            math: this._math,
            threeD: this._threeD,
            regionMap: this._regionMap,
            chartEx: this._chartEx,
          },
          opts.onTextRun,
        );
      });
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * Render a slide and return it as an ImageBitmap. Works in both modes; in
   * worker mode the entire render runs off the main thread. Paint with:
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   */
  async renderSlideToBitmap(
    slideIndex: number,
    opts: RenderSlideToBitmapOptions = {},
  ): Promise<ImageBitmap> {
    this._assertResourceHealthy();
    try {
      const width = opts.width ?? 960;
      const dpr = opts.dpr ?? defaultDpr();
      if (this._mode === 'worker') {
        if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
          throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
        }
        const res = await this._bridge.request(
          (id) => ({ kind: 'renderSlide', id, slideIndex, width, dpr, skipMediaControls: opts.skipMediaControls, dim: opts.dim }) satisfies RenderWorkerRequest,
        );
        const rendered = res as Extract<RenderWorkerResponse, { kind: 'slideRendered' }>;
        // IX6 — replay the worker's run geometry to the caller's collector so the
        // selection / find overlay is built on the same path as main mode.
        if (opts.onTextRun) for (const r of rendered.runs) opts.onTextRun(r);
        return rendered.bitmap;
      }
      const off = new OffscreenCanvas(1, 1);
      await this.renderSlide(off, slideIndex, {
        width,
        dpr,
        skipMediaControls: opts.skipMediaControls,
        dim: opts.dim,
        onTextRun: opts.onTextRun,
      });
      return off.transferToImageBitmap();
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * IX6 — collect a slide's text-run geometry (`PptxTextRunInfo[]`) without
   * painting a visible canvas. Works in BOTH modes: worker mode renders the
   * slide off-thread and ships only the runs (no bitmap transfer); main mode
   * renders to a throwaway offscreen canvas. Used by the find controller to scan
   * every slide for matches. Run geometry is in CSS px (independent of dpr) and
   * dimming does not move glyphs, so only `width` is threaded — matching the
   * historical main-mode `_collectSlideRuns`.
   */
  async collectSlideRuns(slideIndex: number, width = 960): Promise<PptxTextRunInfo[]> {
    this._assertResourceHealthy();
    try {
      if (this._mode === 'worker') {
        if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
          throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
        }
        const res = await this._bridge.request(
          (id) => ({ kind: 'collectRuns', id, slideIndex, width }) satisfies RenderWorkerRequest,
        );
        return (res as Extract<RenderWorkerResponse, { kind: 'runsCollected' }>).runs;
      }
      const runs: PptxTextRunInfo[] = [];
      const off = new OffscreenCanvas(1, 1);
      await this.renderSlide(off, slideIndex, { width, onTextRun: (r) => runs.push(r) });
      return runs;
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * Return a compact, detached snapshot of the topmost element whose transformed
   * frame contains a point in slide EMU coordinates. Straight lines use the
   * explicit tolerance. Works in both render modes and exposes no archive paths,
   * mutable element model, or editor tree position.
   */
  async getElementContextAt(
    slideIndex: number,
    point: PptxSlidePoint,
    options: PptxElementContextOptions = {},
  ): Promise<PptxElementContext | null> {
    this._assertResourceHealthy();
    if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
    try {
      if (this._mode === 'worker') {
        const response = await this._bridge.request(
          (id) => ({ kind: 'hitTestElement', id, slideIndex, point, options }) satisfies RenderWorkerRequest,
        );
        return (response as Extract<RenderWorkerResponse, { kind: 'elementHit' }>).context;
      }
      if (!this._slides) throw new Error('Presentation not loaded');
      return await this._slides.withSlide(slideIndex, (slide) =>
        hitTestPptxSlideContext(slideIndex, slide, point, options));
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * Extract raw media bytes for a zip path referenced by {@link MediaElement}.
   * Results share a count- and byte-bounded cache with embedded images.
   */
  async getMedia(mediaPath: string): Promise<Blob> {
    this._assertResourceHealthy();
    try {
      const mimeType = this._findMimeTypeForPath(mediaPath);
      return await this._rawParts.get(mediaPath, mimeType, async () => {
        const res = await this._bridge.request(
          (id) => ({ kind: 'extractMedia', id, path: mediaPath }) satisfies PptxWorkerRequest,
        );
        const bytes = (res as Extract<PptxWorkerResponse, { kind: 'mediaExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      });
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  private _findMimeTypeForPath(mediaPath: string): string {
    return this._preflight ? findPreflightMimeType(this._preflight, mediaPath) : '';
  }

  /**
   * Extract raw bytes for an embedded image by zip path (e.g.
   * "ppt/media/image1.png"), wrapped in a Blob of the given MIME type. Mirrors
   * {@link getMedia}; results are cached by path for the lifetime of this
   * instance within a common count and byte budget. The renderer routes its `fetchImage` option here so images are
   * decoded lazily rather than inlined as base64 at parse time.
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    this._assertResourceHealthy();
    try {
      return await this._rawParts.get(imagePath, mimeType, async () => {
        const res = await this._bridge.request(
          (id) => ({ kind: 'extractImage', id, path: imagePath }) satisfies PptxWorkerRequest,
        );
        const bytes = (res as Extract<PptxWorkerResponse, { kind: 'imageExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      });
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /** Return a fresh content-free metrics snapshot, including lazy slide and
   * media work completed since load. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    const metrics = this._metrics;
    if (!metrics) throw new Error('Presentation not loaded');
    return readLatestOoxmlResourceMetrics(metrics, async (timeoutMs) => {
      const response = await this._bridge.request(
        (id) => ({ kind: 'resourceUsage', id }) satisfies PptxWorkerRequest,
        undefined,
        { timeoutMs },
      );
      return (response as Extract<PptxWorkerResponse, { kind: 'resourceUsage' }>).usage;
    });
  }

  /**
   * Project the presentation to GitHub-flavoured markdown: title slides become
   * `#` headings, body shapes become nested bullets at each paragraph's `lvl`,
   * tables become pipe tables, charts become summarised bullets, and speaker
   * notes and comments are collated. Positioning, animations, images, and
   * drawing detail are discarded — the projection is meant for AI ingestion and
   * full-text search, not layout.
   *
   * Runs entirely in the worker off the archive opened at {@link load} (no
   * re-copy of the file, no re-parse of the model on the main thread), so it
   * works in BOTH `mode: 'main'` and `mode: 'worker'`.
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * const md = await pres.toMarkdown();
   */
  async toMarkdown(): Promise<string> {
    this._assertResourceHealthy();
    try {
      const res = await this._bridge.request(
        (id) => ({ kind: 'toMarkdown', id }) satisfies PptxWorkerRequest,
      );
      return (res as Extract<PptxWorkerResponse, { kind: 'markdownRendered' }>).markdown;
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * Render a slide and attach canvas-native playback controls for any
   * embedded audio/video. Returns a {@link PresentationHandle} that owns the
   * RAF loop, media elements, and object URLs. Unlike {@link renderSlide}, this
   * method is stateful — always call `handle.destroy()` when leaving the slide.
   */
  async presentSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: PresentSlideOptions = {},
  ): Promise<PresentationHandle> {
    this._assertResourceHealthy();
    try {
      if (!this._preflight) {
        throw new Error('Presentation not loaded');
      }
    if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? (canvas.offsetWidth || 960);

    const drawBase =
      this._mode === 'worker'
        ? async () => {
            // Whole slide rendered off-thread; the handle snapshots this paint
            // into its own base copy, so the bitmap can be closed right after.
            // IX6 — the run geometry rides back beside the bitmap, so a media
            // slide is as selectable/searchable in worker mode as in main mode.
            const bmp = await this.renderSlideToBitmap(slideIndex, { width, dpr, skipMediaControls: true, dim: opts.dim, onTextRun: opts.onTextRun });
            canvas.width = bmp.width;
            canvas.height = bmp.height;
            // Set only the CSS width and let height follow the intrinsic aspect
            // ratio — mirrors the main renderer (renderer.ts), which avoids an
            // explicit style.height that could fight the ratio.
            canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
            if (!canvas.style.display) canvas.style.display = 'block';
            const ctx = canvas.getContext('2d');
            if (!ctx) throw new Error('2D context not available');
            ctx.drawImage(bmp, 0, 0);
            bmp.close();
          }
        : () =>
            this.renderSlide(canvas, slideIndex, {
              width,
              dpr,
              skipMediaControls: true,
              dim: opts.dim,
              onTextRun: opts.onTextRun,
            });

    const mediaElements = this._preflight.slides[slideIndex]?.mediaElements ?? [];

      return await createPresentationHandle(canvas, mediaElements, {
        width,
        dpr,
        slideWidthEmu: this.slideWidth,
        fetchMedia: this._fetchMedia,
        fetchImage: this._fetchImage,
        drawBase,
        onError: opts.onError,
      });
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._slidePullClient?.cancelAll();
    this._bridge.terminate();
    this._slides?.clear();
    this._slides = null;
    this._slidePullClient = null;
    this._preflight = null;
    this._resourceFailure = null;
    this._slidePartIndex = null;
    this._rawParts.clear();
    // Release the Google-Fonts substitutes this deck preloaded into the shared
    // FontFaceSet (main mode). Refcounted in core: a web font also used by another
    // open deck stays until that one is destroyed too. Without this, every opened
    // deck left its Google FontFace objects in `document.fonts` forever (SPA leak).
    if (this._googleFontFaces.length > 0) {
      unloadGoogleFonts(this._googleFontFaces);
      this._googleFontFaces = [];
    }
    // Release this deck's decoded raster bitmaps (GPU-backed), duotone-recoloured
    // rasters, and SVG object URLs promptly; all three caches are keyed by
    // `_fetchImage`.
    dropImageBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
  }
}
