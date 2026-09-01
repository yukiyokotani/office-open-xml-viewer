import type { DimOptions, Presentation, PptxComment, Slide } from './types';
import {
  renderSlideWithEmbeddedFonts,
  dropImageBitmapCache,
  type TextRunCallback,
  type PptxTextRunInfo,
} from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import {
  buildSlidePartIndex,
  resolveInternalSlideTarget,
  type SlidePartNames,
} from './slide-nav';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  unregisterEmbeddedFonts,
  WorkerBridge,
  defaultDpr,
  isHTMLCanvas,
  dropSvgImageCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  OoxmlResourceLimitError,
  type LoadOptions as CoreLoadOptions,
  type ProgressiveLayoutPartial,
  type ProgressiveLayoutProgress,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type TiffRenderer,
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
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { ProgressiveLayoutObserverNotifier } from '@silurus/ooxml-core/internal/progressive-layout-observers';
import { PPTX_GOOGLE_FONTS } from './google-fonts';
import {
  findPreflightMimeType,
  normalizePresentationBootstrap,
  normalizePresentationPreflight,
  normalizePresentationPreflightPrefix,
  PresentationPreflightBuilder,
  type PresentationPreflight,
} from './presentation-preflight';
import { PptxSlideRepository } from './slide-repository';
import { excludeEmbeddedFontFamilies, loadEmbeddedFonts } from './embedded-fonts';
import {
  isPptxSlidePullResponse,
  PptxSlidePullClient,
} from './slide-pull-client';
import type {
  PptxWorkerRequest,
  PptxWorkerResponse,
  PresentationBootstrap,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';
import {
  findPptxElementBoundsByIds,
  hitTestPptxSlideContext,
  type PptxElementContextOptions,
  type PptxElementContext,
  type PptxElementBounds,
  type PptxSlidePoint,
} from './element-selection';
import { publishPptxLayout } from './presentation-layout-events';
import { yieldToHostTaskQueue } from './worker-task-scheduler';

/** Options for {@link PptxPresentation.load}. */
export type LoadOptions = CoreLoadOptions & {
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link PptxPresentation.renderSlideToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   */
  mode?: 'main' | 'worker';
  /**
   * Resolve `load()` when the opening slide is paintable, then prepare the
   * remaining slides in the background. Mirrors DOCX progressive layout:
   * {@link PptxPresentation.layoutComplete} becomes false until the same
   * sequential preflight reaches the end, and
   * {@link PptxPresentation.waitUntilLayoutComplete} observes the final result.
   *
   * Unlike DOCX pagination, the PPTX bootstrap already knows the final slide
   * count and uniform dimensions. {@link PptxPresentation.slideCount} is
   * therefore final from first paint; only
   * {@link PptxPresentation.availableSlideCount} grows. This keeps a
   * ScrollViewer's extent and scrollbar stable while later slides prepare.
   */
  progressiveLayout?: boolean;
  /** Called as the sequential preflight commits slides. Observer failures are isolated. */
  onLayoutProgress?: (progress: Readonly<ProgressiveLayoutProgress>) => void;
  /** Called for each additional paintable prefix after `load()` resolves. Observer failures are isolated. */
  onLayoutPartial?: (progress: Readonly<ProgressiveLayoutPartial>) => void;
  /** Called once background preflight completes, or with its failure. Only
   * fires when progressive loading actually deferred work after `load()`.
   * Observer failures are isolated. */
  onLayoutComplete?: (error?: unknown) => void;
};

interface Deferred<T> {
  readonly promise: Promise<T>;
  resolve(value: T | PromiseLike<T>): void;
  reject(reason?: unknown): void;
}

function deferred<T>(): Deferred<T> {
  let resolve!: Deferred<T>['resolve'];
  let reject!: Deferred<T>['reject'];
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

interface ProgressiveLoad {
  readonly onProgress?: LoadOptions['onLayoutProgress'];
  readonly onPartial?: LoadOptions['onLayoutPartial'];
  readonly onComplete?: LoadOptions['onLayoutComplete'];
  readonly firstPublication: Deferred<void>;
  published: boolean;
  deferred: boolean;
  settled: boolean;
}

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
  private _bootstrap: PresentationBootstrap | null = null;
  private _preflight: PresentationPreflight | null = null;
  /** Paintable prefix under progressiveLayout; final slideCount is bootstrap-owned. */
  private _availableSlideCount = 0;
  private readonly _layoutLifecycle = new ProgressiveLayoutLifecycle();
  private readonly _layoutObservers = new ProgressiveLayoutObserverNotifier();
  private _layoutCompletion: Promise<void> | null = null;
  private _parseRequestId: number | null = null;
  private _progressive: ProgressiveLoad | null = null;
  private _progressiveWatchdog: ReturnType<typeof setTimeout> | undefined;
  private _progressiveWatchdogMs: number | undefined;
  private readonly _layoutWaiters = new Set<() => void>();
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
  /** Embedded Font parts registered into the main-thread FontFaceSet. */
  private _embeddedFontFaces: FontFace[] = [];
  private _embeddedFontAliases: ReadonlyMap<string, string> = new Map();
  private _embeddedFontAuthoredFamilies: ReadonlyMap<string, string> = new Map();
  private _destroyed = false;
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
  private _tiff: TiffRenderer | undefined;

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
      onUnsolicited: (msg) => this._onWorkerLayoutPush(msg),
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
      if (opts.tiff && mode === 'worker' && !rendererDescriptors?.tiff) {
        console.warn(
          "[ooxml] a custom TIFF codec cannot cross the worker boundary; TIFF images will be skipped in mode: 'worker'. Use the codec from @silurus/ooxml/tiff.",
        );
      }
      pres._tiff = mode === 'worker' ? undefined : opts.tiff;
      const progressive = opts.progressiveLayout
        ? {
            onProgress: opts.onLayoutProgress,
            onPartial: opts.onLayoutPartial,
            onComplete: opts.onLayoutComplete,
            firstPublication: deferred<void>(),
            published: false,
            deferred: false,
            settled: false,
          } satisfies ProgressiveLoad
        : undefined;
      await pres._parse(
        buffer,
        resourceOptions.policy,
        !!opts.useGoogleFonts,
        opts.workerTimeoutMs,
        (usage) => metrics.observeUsage(usage),
        rendererDescriptors,
        progressive,
      );
      metrics.checkpoint('presentation preflight ready');
      if (mode === 'main' && opts.useGoogleFonts && pres._preflight && !progressive) {
        pres._googleFontFaces = await preloadGoogleFonts(
          excludeEmbeddedFontFamilies(
            pres._preflight.fontPreloadNames,
            pres._embeddedFontAliases,
          ),
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
    progressive?: ProgressiveLoad,
  ): Promise<void> {
    if (progressive) {
      this._progressive = progressive;
      if (this._mode === 'worker') {
        await this._parseWorkerProgressively(
          buffer, resourcePolicy, useGoogleFonts, timeoutMs, onUsage, renderers, progressive,
        );
      } else {
        await this._parseMainProgressively(
          buffer, resourcePolicy, useGoogleFonts, timeoutMs, onUsage, progressive,
        );
      }
      return;
    }
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
      this._bootstrap = this._preflight;
      this._availableSlideCount = this._preflight.slideCount;
      return;
    }

    const bootstrap = normalizePresentationBootstrap(
      (response as Extract<PptxWorkerResponse, { kind: 'presentationOpened' }>).bootstrap,
    );
    this._bootstrap = bootstrap;
    // Font extraction is independent of the retained slide cursor, so start it
    // from bootstrap metadata while the worker preflights the slide sequence.
    // Rendering still awaits `_parse`, which does not return until both finish.
    const embeddedFontLoad = loadEmbeddedFonts(
      bootstrap.embeddedFonts,
      (path) => this.getFontBytes(path),
    ).then((loaded) => {
      if (this._destroyed) unregisterEmbeddedFonts(loaded.faces);
      else {
        this._embeddedFontFaces = loaded.faces;
        this._embeddedFontAliases = loaded.aliases;
        this._embeddedFontAuthoredFamilies = loaded.authoredFamilies;
      }
    });
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
    let finished: Extract<PptxWorkerResponse, { kind: 'presentationPreflightReady' }>;
    try {
      for (let slideIndex = 0; slideIndex < bootstrap.slideCount; slideIndex += 1) {
        await this._slidePullClient.load(slideIndex, false, timeoutMs);
      }
      finished = await this._bridge.request(
        (id) => ({ kind: 'finishPresentationPreflight', id }) satisfies PptxWorkerRequest,
        undefined,
        { timeoutMs },
      ) as Extract<PptxWorkerResponse, { kind: 'presentationPreflightReady' }>;
      await embeddedFontLoad;
    } catch (error) {
      void embeddedFontLoad.catch(() => undefined);
      throw error;
    }
    this._preflight = normalizePresentationPreflight(
      (finished as Extract<PptxWorkerResponse, { kind: 'presentationPreflightReady' }>).preflight,
    );
    this._availableSlideCount = this._preflight.slideCount;
    this._slides = new PptxSlideRepository({
      slideCount: this._preflight.slideCount,
      maxCachedSlides: HARD_MAX_PPTX_CACHED_SLIDES,
      maxCachedStructuralBytes: HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
      loadSlide: async (slideIndex) => {
        const slide = await this._slidePullClient?.load(slideIndex, true, timeoutMs);
        if (!slide) throw new Error('PPTX slide pull client is unavailable');
        return slide;
      },
    });
  }

  private async _parseMainProgressively(
    buffer: ArrayBuffer,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    useGoogleFonts: boolean,
    timeoutMs: number | undefined,
    onUsage: ((usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void) | undefined,
    progressive: ProgressiveLoad,
  ): Promise<void> {
    const response = await this._bridge.request(
      (id) => ({
        kind: 'parse', id, buffer, resourcePolicy, progressiveLayout: true,
      }) satisfies PptxWorkerRequest,
      [buffer],
      { timeoutMs },
    );
    const bootstrap = normalizePresentationBootstrap(
      (response as Extract<PptxWorkerResponse, { kind: 'presentationOpened' }>).bootstrap,
    );
    this._bootstrap = bootstrap;
    const embeddedFontLoad = loadEmbeddedFonts(
      bootstrap.embeddedFonts,
      (path) => this.getFontBytes(path),
    ).then((loaded) => {
      if (this._destroyed) unregisterEmbeddedFonts(loaded.faces);
      else {
        this._embeddedFontFaces = loaded.faces;
        this._embeddedFontAliases = loaded.aliases;
        this._embeddedFontAuthoredFamilies = loaded.authoredFamilies;
      }
    });
    this._slidePullClient = this._createSlidePullClient(bootstrap.slideCount, timeoutMs, onUsage);
    this._slides = new PptxSlideRepository({
      slideCount: bootstrap.slideCount,
      maxCachedSlides: HARD_MAX_PPTX_CACHED_SLIDES,
      maxCachedStructuralBytes: HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
      loadSlide: async (slideIndex) => {
        const slide = await this._slidePullClient?.load(slideIndex, true, timeoutMs);
        if (!slide) throw new Error('PPTX slide pull client is unavailable');
        return slide;
      },
    });
    const builder = new PresentationPreflightBuilder(bootstrap);
    const loadedGoogleFonts = new Set<string>();
    const ensureFonts = async (): Promise<void> => {
      await embeddedFontLoad;
      if (!useGoogleFonts) return;
      const requested = excludeEmbeddedFontFamilies(
        builder.currentFontPreloadNames,
        this._embeddedFontAliases,
      ).filter((name): name is string => !!name && !loadedGoogleFonts.has(name));
      if (requested.length === 0) return;
      for (const name of requested) loadedGoogleFonts.add(name);
      this._googleFontFaces.push(...await preloadGoogleFonts(requested, PPTX_GOOGLE_FONTS));
    };
    const full = (async () => {
      for (let slideIndex = 0; slideIndex < bootstrap.slideCount; slideIndex += 1) {
        await this._slides!.withSlide(slideIndex, (slide) => {
          builder.addSlide(slide);
        });
        await ensureFonts();
        this._applyProgressivePrefix(builder.snapshot(), progressive);
        if (slideIndex === 0 && progressive.deferred) {
          // Match the worker-mode acknowledgement gate: once the opening slide
          // is publishable, let load() continuations enqueue its paint/resource
          // work before preflight starts pulling the next slide.
          await yieldToHostTaskQueue();
        }
      }
      await embeddedFontLoad;
      this._finishProgressiveLayout(builder.finish(), progressive);
    })();
    this._layoutCompletion = full.then(
      () => undefined,
      (error) => this._failProgressiveLayout(error, progressive),
    );
    await progressive.firstPublication.promise;
  }

  private async _parseWorkerProgressively(
    buffer: ArrayBuffer,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    useGoogleFonts: boolean,
    timeoutMs: number | undefined,
    onUsage: ((usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void) | undefined,
    renderers: WorkerRendererDescriptors | undefined,
    progressive: ProgressiveLoad,
  ): Promise<void> {
    this._progressiveWatchdogMs = timeoutMs;
    const parsed = this._bridge.request(
      (id) => {
        this._parseRequestId = id;
        return {
          kind: 'parse', id, buffer, resourcePolicy, useGoogleFonts, renderers,
          progressiveLayout: true,
        } satisfies RenderWorkerRequest;
      },
      [buffer],
      // Healthy progressive work may exceed this interval while continuing to
      // publish slides. Measure silence between publications instead of using
      // an absolute deadline for the authoritative final response.
      { timeoutMs: false },
    );
    this._rearmProgressiveWatchdog();
    this._layoutCompletion = parsed.then(
      (response) => {
        this._parseRequestId = null;
        const ready = response as Extract<RenderWorkerResponse, { kind: 'presentationReady' }>;
        if (ready.usage) onUsage?.(ready.usage);
        this._finishProgressiveLayout(
          normalizePresentationPreflight(ready.preflight),
          progressive,
        );
      },
      (error) => {
        this._parseRequestId = null;
        this._failProgressiveLayout(error, progressive);
      },
    );
    await progressive.firstPublication.promise;
  }

  private _createSlidePullClient(
    slideCount: number,
    timeoutMs: number | undefined,
    onUsage: ((usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void) | undefined,
  ): PptxSlidePullClient {
    return new PptxSlidePullClient({
      slideCount,
      transport: this._bridge.transport(isPptxSlidePullResponse),
      open: async (slideIndex, identity, operationTimeoutMs) => {
        await this._bridge.request(
          (id) => ({
            kind: 'openSlideSession', id, slideIndex, ...identity,
          }) satisfies PptxWorkerRequest,
          undefined,
          { timeoutMs: operationTimeoutMs ?? timeoutMs },
        );
      },
      onUsage,
    });
  }

  private _onWorkerLayoutPush(
    response: PptxWorkerResponse | RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>,
  ): void {
    if (
      !('kind' in response) ||
      response.kind !== 'presentationLayoutPartial' ||
      response.forId !== this._parseRequestId ||
      !this._progressive
    ) return;
    try {
      this._rearmProgressiveWatchdog();
      if (response.usage) this._metrics?.observeUsage(response.usage);
      if (response.bootstrap) this._bootstrap = normalizePresentationBootstrap(response.bootstrap);
      const bootstrap = this._bootstrap;
      if (!bootstrap) throw new Error('PPTX progressive worker published before bootstrap');
      const prior = this._preflight?.slides ?? [];
      if (response.availableSlides !== prior.length + 1 || response.slide.index !== prior.length) {
        throw new Error('PPTX progressive worker published a non-sequential slide');
      }
      this._applyProgressivePrefix(
        normalizePresentationPreflightPrefix({
          ...bootstrap,
          slides: [...prior, response.slide],
          fontPreloadNames: response.fontPreloadNames,
        }),
        this._progressive,
      );
      // Acknowledge only after Window crosses a task boundary. load() promise
      // continuations can enqueue the opening-slide render first; Worker message
      // ordering then handles that request before this ACK releases preflight of
      // the next slide.
      void yieldToHostTaskQueue().then(() => {
        if (this._destroyed || this._parseRequestId !== response.forId) return;
        this._bridge.post({
          kind: 'continuePresentationPreflight',
          forId: response.forId,
          availableSlides: response.availableSlides,
        } satisfies RenderWorkerRequest);
      }).catch((error) => {
        if (this._destroyed || this._parseRequestId !== response.forId || !this._progressive) return;
        this._failProgressiveLayout(error, this._progressive);
        this._bridge.terminate();
      });
    } catch (error) {
      this._failProgressiveLayout(error, this._progressive);
      // The worker is blocked on this publication's acknowledgement. A
      // malformed prefix cannot be acknowledged safely, so terminate the
      // bridge to reject the pending parse request and settle `_layoutCompletion`.
      this._bridge.terminate();
    }
  }

  private _applyProgressivePrefix(
    prefix: PresentationPreflight,
    progressive: ProgressiveLoad,
  ): void {
    if (progressive.settled || this._destroyed) return;
    this._preflight = prefix;
    this._availableSlideCount = prefix.slides.length;
    if (!progressive.published && prefix.slides.length === prefix.slideCount) {
      // Nothing remains deferred. Keep the prefix internal until the
      // authoritative finish path publishes one complete snapshot; callers of
      // load() must not observe a provisional lifecycle for a one-slide deck.
      progressive.published = true;
      progressive.deferred = false;
      return;
    }
    this._layoutLifecycle.begin();
    this._wakeLayoutWaiters();
    this._layoutObservers.notify(
      'onLayoutProgress', progressive.onProgress, { committedUnits: this._availableSlideCount },
    );
    publishPptxLayout(this, {
      availableSlides: this._availableSlideCount,
      slideCount: this.slideCount,
      exact: false,
      complete: false,
    });
    if (!progressive.published) {
      progressive.published = true;
      progressive.deferred = prefix.slides.length < prefix.slideCount;
      progressive.firstPublication.resolve();
      return;
    }
    this._layoutObservers.notify('onLayoutPartial', progressive.onPartial, {
      availableUnits: this._availableSlideCount,
      totalUnits: this.slideCount,
      exact: false,
    });
  }

  private _finishProgressiveLayout(
    preflight: PresentationPreflight,
    progressive: ProgressiveLoad,
  ): void {
    if (progressive.settled || this._destroyed) return;
    progressive.settled = true;
    this._clearProgressiveWatchdog();
    this._preflight = preflight;
    this._bootstrap ??= preflight;
    this._availableSlideCount = preflight.slideCount;
    this._layoutLifecycle.succeed();
    this._wakeLayoutWaiters();
    progressive.firstPublication.resolve();
    publishPptxLayout(this, {
      availableSlides: this._availableSlideCount,
      slideCount: this.slideCount,
      exact: true,
      complete: true,
    });
    if (progressive.deferred) {
      this._layoutObservers.notify('onLayoutComplete', progressive.onComplete);
    }
  }

  private _failProgressiveLayout(error: unknown, progressive: ProgressiveLoad): void {
    if (progressive.settled) return;
    progressive.settled = true;
    this._clearProgressiveWatchdog();
    if (this._destroyed) return;
    if (!progressive.published) {
      progressive.firstPublication.reject(error);
      return;
    }
    const layoutError = this._layoutLifecycle.fail(error);
    this._wakeLayoutWaiters();
    publishPptxLayout(this, {
      availableSlides: this._availableSlideCount,
      slideCount: this.slideCount,
      exact: false,
      complete: false,
      error: layoutError,
    });
    this._layoutObservers.notify('onLayoutComplete', progressive.onComplete, layoutError);
  }

  private _wakeLayoutWaiters(): void {
    for (const resolve of this._layoutWaiters) resolve();
    this._layoutWaiters.clear();
  }

  private _rearmProgressiveWatchdog(): void {
    if (this._progressiveWatchdogMs === undefined) return;
    clearTimeout(this._progressiveWatchdog);
    this._progressiveWatchdog = setTimeout(() => {
      const progressive = this._progressive;
      const silenceMs = this._progressiveWatchdogMs;
      if (!progressive || progressive.settled || silenceMs === undefined || this._destroyed) return;
      const error = new Error(`worker layout produced no progress for ${silenceMs}ms`);
      this._failProgressiveLayout(error, progressive);
      this._bridge.terminate();
    }, this._progressiveWatchdogMs);
  }

  private _clearProgressiveWatchdog(): void {
    clearTimeout(this._progressiveWatchdog);
    this._progressiveWatchdog = undefined;
    this._progressiveWatchdogMs = undefined;
  }

  private async _waitForSlide(slideIndex: number): Promise<void> {
    while (
      !this._destroyed &&
      slideIndex >= this._availableSlideCount &&
      !this._layoutLifecycle.settled
    ) {
      await new Promise<void>((resolve) => this._layoutWaiters.add(resolve));
    }
    if (slideIndex >= this._availableSlideCount) await this.waitUntilLayoutComplete();
  }

  private _assertSlideIndex(slideIndex: number): void {
    if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._bootstrap?.slideCount ?? this._preflight?.slideCount ?? 0; }

  /** Slides whose compact facts and full model can currently be painted. */
  get availableSlideCount(): number { return this._availableSlideCount; }

  /** True only when every slide is paintable; remains false after background failure. */
  get layoutComplete(): boolean { return this._layoutLifecycle.complete; }

  /** Wait until all slides are paintable; rethrows a post-load background failure. */
  async waitUntilLayoutComplete(): Promise<void> {
    if (this._layoutCompletion) await this._layoutCompletion;
    this._layoutLifecycle.throwIfFailed();
  }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._bootstrap?.slideWidth ?? this._preflight?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._bootstrap?.slideHeight ?? this._preflight?.slideHeight ?? 0; }

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
   * notes part. This is a synchronous lookup. During progressive loading its
   * answer is authoritative only for `slideIndex < availableSlideCount`; await
   * {@link waitUntilLayoutComplete} before scanning the whole deck.
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
   * comment-free slide. During progressive loading, `[]` for an unavailable
   * slide means "not known yet"; inspect only `slideIndex < availableSlideCount`
   * or await {@link waitUntilLayoutComplete} before a whole-deck scan. */
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
   * caller's policy (see {@link PptxViewer}'s `hiddenSlideMode` modes). During
   * progressive loading this fact is authoritative only below
   * {@link availableSlideCount}; await completion before scanning every slide.
   */
  isHidden(slideIndex: number): boolean {
    return Number.isInteger(slideIndex)
      ? (this._preflight?.slides[slideIndex]?.hidden ?? false)
      : false;
  }

  /** The compact preflight's per-slide `partName` array (`sldIdLst` order). */
  private _partNames(): SlidePartNames {
    return (this._bootstrap?.slides ?? this._preflight?.slides ?? [])
      .map((slide) => slide.partName);
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
      this._assertSlideIndex(slideIndex);
      await this._waitForSlide(slideIndex);
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
        return renderSlideWithEmbeddedFonts(
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
            embeddedFontAliases: this._embeddedFontAliases,
            embeddedFontAuthoredFamilies: this._embeddedFontAuthoredFamilies,
            fetchMedia: this._fetchMedia,
            fetchImage: this._fetchImage,
            skipMediaControls: opts.skipMediaControls,
            dim: opts.dim,
            math: this._math,
            threeD: this._threeD,
            regionMap: this._regionMap,
            chartEx: this._chartEx,
            tiff: this._tiff,
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
      this._assertSlideIndex(slideIndex);
      await this._waitForSlide(slideIndex);
      const width = opts.width ?? 960;
      const dpr = opts.dpr ?? defaultDpr();
      if (this._mode === 'worker') {
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
      this._assertSlideIndex(slideIndex);
      await this._waitForSlide(slideIndex);
      if (this._mode === 'worker') {
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
    this._assertSlideIndex(slideIndex);
    try {
      await this._waitForSlide(slideIndex);
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

  /** Resolve DrawingML element ids to immutable slide geometry in one lazy
   * slide read. Modern-comment UIs use this to honor authored anchors; custom
   * UIs can use the same primitive without receiving the full slide model. */
  async getElementBoundsByIds(
    slideIndex: number,
    elementIds: readonly string[],
  ): Promise<readonly PptxElementBounds[]> {
    this._assertResourceHealthy();
    this._assertSlideIndex(slideIndex);
    const ids = Object.freeze(elementIds.filter((id) => typeof id === 'string' && id.length > 0));
    if (ids.length === 0) return Object.freeze([]);
    try {
      await this._waitForSlide(slideIndex);
      if (this._mode === 'worker') {
        const response = await this._bridge.request(
          (id) => ({
            kind: 'resolveElementBounds', id, slideIndex, elementIds: ids,
          }) satisfies RenderWorkerRequest,
        );
        return (response as Extract<RenderWorkerResponse, {
          kind: 'elementBoundsResolved';
        }>).bounds;
      }
      if (!this._slides) throw new Error('Presentation not loaded');
      return await this._slides.withSlide(slideIndex, (slide) =>
        findPptxElementBoundsByIds(slide, ids));
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

  private async getFontBytes(fontPath: string): Promise<Uint8Array> {
    this._assertResourceHealthy();
    try {
      const response = await this._bridge.request(
        (id) => ({ kind: 'extractFont', id, path: fontPath }) satisfies PptxWorkerRequest,
      );
      return new Uint8Array(
        (response as Extract<PptxWorkerResponse, { kind: 'fontExtracted' }>).bytes,
      );
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
      this._assertSlideIndex(slideIndex);
      await this._waitForSlide(slideIndex);
      if (!this._preflight) {
        throw new Error('Presentation not loaded');
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

  /**
   * Assemble a detached editor {@link Presentation} JSON model from this
   * loaded package: theme fields from preflight plus every slide pulled through
   * the main-thread repository.
   *
   * Intended as the bootstrap input for `@maxgent/ooxml-pptx-editor` so hosts
   * do not need a second parser JSON source. Slides are `structuredClone`d so
   * later optimistic edits cannot mutate the render cache. Unavailable in
   * `mode: 'worker'` — the same constraint as {@link replaceSlides}.
   *
   * Loading every slide into memory is intentional for editor startup; the
   * ordinary viewer path remains pull/LRU based.
   */
  async toEditorPresentation(): Promise<Presentation> {
    this._assertResourceHealthy();
    if (this._mode === 'worker') {
      throw new Error(
        "toEditorPresentation is unavailable in mode: 'worker'; use mode: 'main' for editor bootstrap",
      );
    }
    if (!this._preflight || !this._slides) throw new Error('Presentation not loaded');

    try {
      await this.waitUntilLayoutComplete();
      const compact = this._preflight;
      const repository = this._slides;
      if (!compact || !repository) throw new Error('Presentation not loaded');
      const slides: Slide[] = [];
      for (let slideIndex = 0; slideIndex < compact.slideCount; slideIndex += 1) {
        slides.push(await repository.withSlide(slideIndex, (slide) => {
          this._assertResourceHealthy();
          return structuredClone(slide);
        }));
      }

      const presentation: Presentation = {
        slideWidth: compact.slideWidth,
        slideHeight: compact.slideHeight,
        slides,
        defaultTextColor: compact.defaultTextColor,
        majorFont: compact.majorFont,
        minorFont: compact.minorFont,
      };
      if (compact.hlinkColor != null) presentation.hlinkColor = compact.hlinkColor;
      if (compact.folHlinkColor != null) presentation.folHlinkColor = compact.folHlinkColor;
      return presentation;
    } catch (error) {
      this._rethrowWithResourceFailure(error);
    }
  }

  /**
   * Replace in-memory slide models used by subsequent main-thread renders.
   *
   * Intended for editor optimistic updates: keep the loaded package's media /
   * theme / fetch plumbing, but paint the caller's mutated {@link Slide}
   * snapshots. Unavailable in `mode: 'worker'` — the worker owns its own slide
   * projections and cannot accept main-thread model patches.
   *
   * @internal Used by `@maxgent/ooxml-pptx-editor`; not part of the public
   * `@maxgent/ooxml/pptx` API.
   */
  replaceSlides(
    replacements: ReadonlyArray<{ readonly index: number; readonly slide: Slide }>,
  ): void {
    this._assertResourceHealthy();
    if (this._mode === 'worker') {
      throw new Error(
        "replaceSlides is unavailable in mode: 'worker'; use mode: 'main' for editor-driven slide updates",
      );
    }
    const repository = this._slides;
    if (!repository) throw new Error('Presentation not loaded');
    for (const { index, slide } of replacements) {
      repository.replace(index, slide);
    }
  }

  /**
   * Replace the complete in-memory slide list used by subsequent main-thread renders.
   *
   * @internal Used by `@maxgent/ooxml-pptx-editor` for optimistic slide insertion
   * and rollback; not part of the public `@maxgent/ooxml/pptx` API.
   */
  replaceSlideList(slides: readonly Slide[]): void {
    this._assertResourceHealthy();
    if (this._mode === 'worker') {
      throw new Error(
        "replaceSlideList is unavailable in mode: 'worker'; use mode: 'main' for editor-driven slide updates",
      );
    }
    const current = this._preflight;
    const repository = this._slides;
    if (!current || !repository) throw new Error('Presentation not loaded');

    const nextSlides = [...slides];
    const nextBootstrap = normalizePresentationBootstrap({
      ...(this._bootstrap ?? current),
      slideCount: nextSlides.length,
      slides: nextSlides.map((slide, index) => ({
        index,
        ...(slide.partName === undefined ? {} : { partName: slide.partName }),
      })),
    });
    const nextPreflight = normalizePresentationPreflight({
      ...current,
      slideCount: nextSlides.length,
      slides: nextSlides.map((slide, index) => ({
        index,
        ...(slide.partName === undefined ? {} : { partName: slide.partName }),
        notes: slide.notes ?? null,
        hidden: slide.hidden ?? false,
        mediaElements: slide.elements.filter((element) => element.type === 'media'),
      })),
    });

    repository.replaceAll(nextSlides);
    this._bootstrap = nextBootstrap;
    this._preflight = nextPreflight;
    this._availableSlideCount = nextSlides.length;
    this._slidePartIndex = null;
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._destroyed = true;
    this._clearProgressiveWatchdog();
    this._slidePullClient?.cancelAll();
    this._bridge.terminate();
    this._slides?.clear();
    this._slides = null;
    this._slidePullClient = null;
    this._bootstrap = null;
    this._preflight = null;
    this._availableSlideCount = 0;
    this._layoutLifecycle.succeed();
    this._layoutCompletion = null;
    this._progressive = null;
    this._parseRequestId = null;
    this._wakeLayoutWaiters();
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
    if (this._embeddedFontFaces.length > 0) {
      unregisterEmbeddedFonts(this._embeddedFontFaces);
      this._embeddedFontFaces = [];
    }
    this._embeddedFontAliases = new Map();
    this._embeddedFontAuthoredFamilies = new Map();
    // Release this deck's decoded raster bitmaps (GPU-backed), duotone-recoloured
    // rasters, and SVG object URLs promptly; all three caches are keyed by
    // `_fetchImage`.
    dropImageBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
  }
}
