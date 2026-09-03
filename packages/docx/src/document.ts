import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  releaseOwnedBitmap,
  unloadGoogleFonts,
  unloadLocalFontMetrics,
  unregisterEmbeddedFonts,
  WorkerBridge,
  defaultDpr,
  dropSvgImageCache,
  dropDecodedBitmapCache,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type ProgressiveLayoutPartial,
  type ProgressiveLayoutProgress,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type OoxmlResourceMetrics,
  workerRendererDescriptors,
} from '@silurus/ooxml-core';
import { resolveOfficeInputWithOptionalConversion } from '@silurus/ooxml-core/internal/legacy-office-conversion';
import {
  deserializeWorkerError,
  disposeRejectedLoad,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  normalizeLoadResourceOptions,
  OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS,
  OoxmlResourceMetricsSession,
  readLatestOoxmlResourceMetrics,
  PULL_SESSION_PROTOCOL,
  respondToWorkerSvgDecodeRequest,
  type NormalizedOoxmlResourcePolicy,
  type WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { ProgressiveLayoutObserverNotifier } from '@silurus/ooxml-core/internal/progressive-layout-observers';
import type { DocxDocumentModel, RenderPageOptions, WorkerRequest, WorkerResponse, DocComment, DocNote, DocRevision } from './types';
import { renderLayoutSourceToCanvas, documentHasMath, prepareMathRuns, type DocxTextRunInfo } from './renderer';
import { createLayoutServices } from './layout-runtime.js';
import { buildBookmarkPageMap } from './bookmark-nav';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import { loadDocxLocalFontMetrics } from './local-font-metrics';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
  layoutVariantStoreOf,
} from './layout/runtime-state.js';
import {
  type LayoutSourceStore,
} from './layout/layout-source-store.js';
import type { DeepReadonly, DocumentLayout } from './layout/types.js';
import { snapshotPlainData } from './layout/plain-data.js';
import type {
  DocumentLayoutPartial,
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
  WireRenderPageOptions,
} from './worker-protocol';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import {
  reviewProjectionIndexForDocument,
  type ReviewProjectionIndex,
} from './layout/text-index.js';
import {
  hitTestSelectedDocxElementContext,
  type DocxElementContextOptions,
} from './element-context.js';
import type { DocxElementContext, DocxPagePoint } from './selection-context.js';
import {
  collectLayoutSourceCommentRangesIfPresent,
  resolveDocxCommentThreads,
  type CommentAnchorRange,
  type ResolvedDocxCommentThread,
  type ResolveDocxCommentThreadsOptions,
} from './comments.js';
import {
  collectLayoutSourceRevisionRangesIfPresent,
  type RevisionAnchorRange,
} from './revisions.js';
import {
  isDocumentPullResponse,
  materializeDocumentPullAdapterSession,
} from './document-pull-client.js';
import { layoutDocumentInputAsync } from './layout/document.js';
import { layoutDocumentProgressively } from './layout/progressive.js';
import { PaginationAbortError } from './layout/pagination-scheduler.js';
import { normalizeLayoutOptions, type LayoutOptions } from './layout/options.js';
import type { LayoutVariantStore } from './layout/variant-store.js';
import { publishDocxLayout } from './document-layout-events.js';
import {
  docxLayoutViewRequester,
  publishDocxLayoutView,
} from './document-layout-view.js';

/** Options for {@link DocxDocument.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `resourceLimits`, and the
 *  deprecated `maxZipEntryBytes` alias) with the opt-in math engine. */
export interface LoadOptions extends CoreLoadOptions {
  /**
   * Opt-in OMML equation engine. Import it from the separate `@silurus/ooxml/math`
   * entry and pass it in: `import { math } from '@silurus/ooxml/math'`. When
   * omitted, equations are skipped and the ~3 MB engine is not loaded or evaluated.
   */
  math?: MathRenderer;
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse, paginate AND render inside the worker; use
   * {@link DocxDocument.renderPageToBitmap} and paint the returned ImageBitmap
   * via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas. Documents
   * needing DOM-only OpenType vertical glyph selection transparently continue
   * in main mode; {@link DocxDocument.mode} reports the effective mode. Built-in
   * optional renderers use the same injection options in both modes. Custom
   * renderer objects use their documented fallback in worker mode.
   */
  mode?: 'main' | 'worker';
  /**
   * Lay the document out in slices instead of one blocking call, releasing the
   * thread between body entries.
   *
   * Layout is unchanged — this only spreads it over event-loop turns — so the
   * document that eventually loads is byte-identical either way. What changes is
   * that the main thread stays responsive while a large document paginates.
   * In `mode: 'worker'`, pagination is already off the UI thread; this option
   * has no additional effect unless progressive layout is enabled. `load()`
   * still resolves only once layout is complete.
   *
   * Off by default: for small documents the slicing is pure overhead, and
   * existing callers should not silently change scheduling behaviour.
   */
  sliceLayout?: boolean;
  /**
   * Called as layout progresses, with the number of pages committed so far.
   *
   * Monotonic within a pagination pass, but NOT across passes: pagination
   * restarts from page zero for each convergence pass (header/footer reserve
   * measurement, PAGE/NUMPAGES feedback), so this can go down. Treat a decrease
   * as "layout is re-deriving pages you were already told about" — it is
   * progress reporting, not a page count.
   *
   * Implies {@link sliceLayout} in main mode, since a blocking layout cannot
   * deliver progress the UI could act on. Worker mode reports this callback
   * when {@link progressiveLayout} is enabled; an ordinary worker load does not
   * publish intermediate pagination progress. Observer failures are reported
   * and isolated from the layout result.
   */
  onLayoutProgress?: (progress: Readonly<ProgressiveLayoutProgress>) => void;
  /**
   * Resolve `load()` as soon as the document's opening pages are ready, and
   * finish laying the rest out in the background.
   *
   * A large document otherwise shows nothing at all until every page has been
   * paginated. With this on, the canonical paginator publishes composed
   * checkpoints while continuing from the same suspended session, so the
   * viewer can paint without replaying a growing source prefix.
   * {@link DocxDocument.pageCount} therefore starts small and grows, the way
   * Word's page count settles during background repagination.
   *
   * Early pages are provisional because later header/footer, field or anchored
   * object convergence can still repaginate them. The finished layout is
   * byte-identical to a blocking load either way.
   *
   * Await {@link DocxDocument.waitUntilLayoutComplete} before anything that needs the
   * whole document: page-count-sensitive UI, text search, bookmark navigation,
   * printing or export.
   *
   * Works in BOTH `mode: 'main'` and `mode: 'worker'`, and for any layout
   * variant: `showTrackedChanges` and an explicit {@link currentDate} each
   * select their own pagination, and the worker previews, paginates and reports
   * metadata for whichever one the load chose. In worker mode the pagination
   * stays in the render worker, so the opening pages arrive without a blank
   * wait AND the remaining layout never competes for main-thread frames.
   *
   * Off by default: `load()` resolving before the document is fully laid out is
   * a behaviour change existing callers should opt into.
   */
  progressiveLayout?: boolean;
  /**
   * Called once the full layout has replaced the provisional one, or with the
   * failure if background layout threw. Only fires when
   * {@link progressiveLayout} actually deferred work. Observer failures are
   * reported and isolated from the layout result.
   */
  onLayoutComplete?: (error?: unknown) => void;
  /**
   * Called each time progressive layout publishes more pages, before the
   * authoritative layout replaces them.
   *
   * `availableUnits` is the pages available so far. `totalUnits` is omitted
   * because pagination does not know the final page count yet. `exact` is
   * currently always false because a later convergence pass can still replace
   * those pages. Observer failures are reported and isolated from the layout
   * result.
   */
  onLayoutPartial?: (progress: Readonly<ProgressiveLayoutPartial>) => void;
  /**
   * Lay the document out for the tracked-change markup view (ECMA-376
   * §17.13.5) rather than its final state.
   *
   * This is a LAYOUT input, not a paint flag: showing deletions changes line
   * breaking and pagination, so it selects a different retained layout with its
   * own page count. Passing it here makes load build the variant that will
   * actually be rendered — otherwise the first render misses the cached layout
   * and repaginates the whole document synchronously, which on a large reviewed
   * document is seconds of frozen UI. Viewers that expose a markup toggle
   * should pass their initial state here and call
   * {@link DocxDocument.setLayoutView} when it changes.
   */
  showTrackedChanges?: boolean;
  /**
   * Date used to resolve DATE/TIME fields, and part of the layout variant key
   * for the same reason as {@link showTrackedChanges}. Omitted means "now at
   * load time", which is what the renderer defaults to.
   */
  currentDate?: Date | number;
}

/** Options for {@link DocxDocument.collectPageRuns}. */
export type CollectPageRunsOptions = Pick<
  RenderPageOptions,
  'width' | 'currentDate' | 'showTrackedChanges'
>;

/** Options for {@link DocxDocument.getCommentThreads}. */
export interface DocxPageCommentThreadsOptions
  extends CollectPageRunsOptions, ResolveDocxCommentThreadsOptions {}

/** IX6 — options for {@link DocxDocument.renderPageToBitmap}: the serializable
 *  render knobs plus an OPTIONAL `onTextRun`. The callback stays main-thread (it
 *  never crosses the wire); in worker mode the proxy invokes it with the runs
 *  the worker shipped back beside the bitmap, so a caller gets the selection /
 *  find geometry on the same path in both modes. */
export type RenderPageToBitmapOptions = Omit<RenderPageOptions, 'onTextRun'> & {
  onTextRun?: (run: DocxTextRunInfo) => void;
};

interface ReviewSnapshot {
  readonly comments: readonly Readonly<DocComment>[];
  readonly revisions: readonly Readonly<DocRevision>[];
}

/** Live state of a worker-mode progressive load: the caller's callbacks, the
 *  abort controller `destroy()` trips, and the deferred that lets `load()`
 *  resolve on the first published prefix instead of the finished document. */
interface WorkerProgressiveLoad {
  readonly onPartial?: LoadOptions['onLayoutPartial'];
  readonly onComplete?: LoadOptions['onLayoutComplete'];
  readonly onProgress?: LoadOptions['onLayoutProgress'];
  /** The immutable view selected by the parse request. Later publications from
   *  this load may update host geometry only while this view remains active. */
  readonly layoutOptions: LayoutOptions;
  readonly abort: AbortController;
  /** Settles when the worker publishes its first prefix, or — when it publishes
   *  none — when the authoritative `parsedMeta` lands. Either way it is what
   *  `load()` waits on. */
  readonly firstPublication: Deferred<void>;
  published: boolean;
  /** Exactly one of authoritative success, parse failure, watchdog failure, or
   * teardown may terminate this worker progressive session. */
  settled: boolean;
}

interface Deferred<T> {
  readonly promise: Promise<T>;
  readonly resolve: (value: T) => void;
  readonly reject: (error: unknown) => void;
}

/** LayoutOptions are already normalized; absent and explicit false select the
 * same retained variant. */
function sameLayoutView(
  left: LayoutOptions | null,
  right: LayoutOptions,
): boolean {
  return left !== null
    && left.currentDateMs === right.currentDateMs
    && (left.showTrackedChanges === true) === (right.showTrackedChanges === true);
}

function deferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

/** Identity-stable empties for the two anchor projections.
 *  `commentAnchorRanges()` is polled per draw and consumers cache on identity
 *  (`DocxScrollViewer._hasDisplayableComments`), so returning a fresh `[]` each
 *  call would rebuild that cache on every frame — most visibly during the
 *  progressive window, where worker metadata carries no anchor ranges yet. */
const NO_COMMENT_ANCHOR_RANGES: readonly CommentAnchorRange[] = Object.freeze([]);
const NO_REVISION_ANCHOR_RANGES: readonly RevisionAnchorRange[] = Object.freeze([]);

const EMPTY_REVIEW_SNAPSHOT: ReviewSnapshot = Object.freeze({
  comments: Object.freeze([]),
  revisions: Object.freeze([]),
});

function snapshotReviewData(
  comments: readonly DocComment[],
  revisions: readonly DocRevision[],
): ReviewSnapshot {
  // Runtime sealing is recursive; the public type stays shallowly readonly so
  // existing consumers can read nested paragraph arrays without a new type.
  return snapshotPlainData({ comments, revisions }, 'DOCX review metadata') as unknown as ReviewSnapshot;
}

export class DocxDocument {
  private _metrics: OoxmlResourceMetricsSession | null = null;
  private _document: DocxDocumentModel | null = null;
  private _source: LayoutSourceStore | null = null;
  private _meta: DocumentMeta | null = null;
  /** One immutable review snapshot backs both the public detached records and
   * lazy anchor projections. It is captured once per load in both render modes. */
  private _review: ReviewSnapshot = EMPTY_REVIEW_SNAPSHOT;
  /** Separates successful completion from terminal background failure. */
  private readonly _layoutLifecycle = new ProgressiveLayoutLifecycle();
  private readonly _layoutObservers = new ProgressiveLayoutObserverNotifier();
  /** Settles when background layout finishes; never rejects (the lifecycle
   *  retains a failure and `waitUntilLayoutComplete` re-throws it). */
  private _layoutCompletion: Promise<void> | null = null;
  /** Cancels background layout when the document is destroyed or replaced. */
  private _layoutAbort: AbortController | null = null;
  /** Lazily-built `bookmarkName → 0-based page index` map for internal hyperlink
   *  anchors (IX-nav). Built on first {@link getBookmarkPage} from the paginated
   *  pages (main) or the worker meta's `bookmarkPages` (worker). Nulled by
   *  {@link destroy} so a reused reference never serves a stale document. */
  private _bookmarkPages: Map<string, number> | null = null;
  /** Lazily-computed §17.13.4 comment anchor ranges (main mode; worker mode
   *  reads them from the meta). Nulled by {@link destroy} with the other
   *  per-document caches. */
  private _commentAnchorRanges: readonly CommentAnchorRange[] | null = null;
  /** Lazily-computed §17.13.5 revision source ranges (main mode; worker mode
   * reads them from metadata). */
  private _revisionAnchorRanges: readonly RevisionAnchorRange[] | null = null;
  /** Shared one-pass index for comment and revision anchor projections. */
  private _reviewProjectionIndex: ReviewProjectionIndex | null = null;
  /** Correlation id of the in-flight worker `parse`. Progressive pushes name it
   *  in `forId`; ids are monotonic and never reused, so a push carrying any
   *  other id belongs to a parse this document has already moved past. */
  private _parseRequestId: number | null = null;
  /** Set for the duration of a worker-mode progressive parse. */
  private _progressive: WorkerProgressiveLoad | null = null;
  /** Silence watchdog for a progressive parse. The parse request itself gives up
   *  its deadline (a background layout legitimately outlives any fixed timeout),
   *  so liveness is tracked here instead and re-armed by every worker push. */
  private _parseWatchdog: ReturnType<typeof setTimeout> | undefined;
  private _parseWatchdogMs: number | undefined;
  /** Latest runtime layout-view request. Worker replies may arrive out of
   * order; only this generation may atomically install its variant + metadata. */
  private _layoutViewGeneration = 0;
  private _mode: 'main' | 'worker' = 'main';
  private _threeD: ChartThreeDRenderer | undefined;
  private _regionMap: ChartRegionMapRenderer | undefined;
  private _chartEx: ChartExRenderer | undefined;
  private _tiff: import('@silurus/ooxml-core').TiffRenderer | undefined;
  private _worker: Worker;
  private _bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private readonly _rawParts = new BoundedRawPartCache({
    maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
    maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
  });
  /** Embedded `FontFace` objects this document registered into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with
   *  its own FontFaceSet). Released in {@link destroy} so they do not leak into
   *  the shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in
   *  core, so a font shared with another open document survives until both go). */
  private _embeddedFontFaces: FontFace[] = [];
  /** Google-Fonts `FontFace` objects this document preloaded into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with its
   *  own FontFaceSet). Released in {@link destroy} so they do not leak into the
   *  shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in core,
   *  so a web font shared with another open document survives until both go). */
  private _googleFontFaces: FontFace[] = [];
  /** Exact local faces used for version-adaptive Office line metrics. */
  private _localMetricFontFaces: FontFace[] = [];
  /** One stable closure per instance: core's path-keyed SVG cache namespaces on
   *  this identity, so two open documents never swap a shared zip path (e.g.
   *  word/media/image1.svg). Reusing one reference also lets the SVG cache hit
   *  across page renders. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImage(path, mime);

  private constructor(
    worker: Worker,
    mode: 'main' | 'worker',
    defaultCurrentDateMs: number,
    wasmUrlOverride?: string | URL,
  ) {
    this._worker = worker;
    this._mode = mode;
    attachDocumentLayoutRuntime(this, defaultCurrentDateMs);
    this._bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this._worker, {
      correlate: (res) =>
        'protocol' in res && res.protocol === PULL_SESSION_PROTOCOL
          ? res.requestId
          : 'id' in res
            ? res.id
            : undefined,
      // Progressive layout pushes carry `forId` rather than `id`, so `correlate`
      // above returns undefined for them and they arrive here instead of
      // resolving the still-pending `parse`. That is the whole mechanism.
      onUnsolicited: (res) => this._onWorkerLayoutPush(res),
      toError: (res) => {
        if ('protocol' in res || res.type !== 'error') return undefined;
        // Reconstruct every shared typed error first (resource quota, decoded
        // image quota, pull credit, container errors), then preserve DOCX-only
        // pagination diagnostics as supplemental fields.
        return Object.assign(deserializeWorkerError(res), {
          ...(res.reason !== undefined ? { reason: res.reason } : {}),
          ...(res.outgoingColumnIndex !== undefined
            ? { outgoingColumnIndex: res.outgoingColumnIndex }
            : {}),
          ...(res.outgoingColumnCount !== undefined
            ? { outgoingColumnCount: res.outgoingColumnCount }
            : {}),
          ...(res.incomingColumnCount !== undefined
            ? { incomingColumnCount: res.incomingColumnCount }
            : {}),
        });
      },
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
    this._bridge.post({ type: 'init', wasmUrl } satisfies WorkerRequest);
  }

  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<DocxDocument> {
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const defaultCurrentDateMs = Date.now();
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'docx',
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
    // Resolve the container on the main thread before spinning up the worker.
    // Container errors remain typed OoxmlError instances here; `instanceof`
    // would not survive the worker boundary.
    buffer = toArrayBuffer(await resolveOfficeInputWithOptionalConversion(
      buffer,
      'docx',
      opts.legacyConversion,
      opts.password,
    ));
    metrics.setSourceBytes(buffer.byteLength);
    metrics.checkpoint('container ready');
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const rendererDescriptors = mode === 'worker' ? workerRendererDescriptors(opts) : undefined;
    const workerProgressive = mode === 'worker' && !!opts.progressiveLayout;
    let doc: DocxDocument | undefined;
    try {
      doc = new DocxDocument(worker, mode, defaultCurrentDateMs, opts.wasmUrl);
      doc._metrics = metrics;
      // The variant the caller will actually render, recorded for BOTH render
      // modes and recorded BEFORE the parse: geometry accessors and the
      // per-call option fill-in (`_withActiveView`) read it, the wire options
      // for every render/collect/hit-test request are filled from it, and in
      // worker mode the parse itself now carries it so the worker paginates
      // — and reports metadata for — this same variant rather than the
      // default one.
      const loadRuntime = documentLayoutRuntimeOf(doc);
      const initialLayoutOptions = normalizeLayoutOptions(
        opts.currentDate,
        loadRuntime.defaultCurrentDateMs,
        opts.showTrackedChanges === true,
      );
      loadRuntime.activeLayoutOptions = initialLayoutOptions;
      // In worker mode the worker preloads fonts before paginating (pagination
      // measures text), so the flag is forwarded; in main mode fonts are loaded
      // here after parse, before the lazy first pagination.
      await doc._parse(
        buffer,
        resourceOptions.policy,
        mode === 'worker' ? !!opts.useGoogleFonts : false,
        opts.workerTimeoutMs,
        (usage) => metrics.observeUsage(usage),
        rendererDescriptors,
        workerProgressive
          ? {
              onPartial: opts.onLayoutPartial,
              onComplete: opts.onLayoutComplete,
              onProgress: opts.onLayoutProgress,
              layoutOptions: initialLayoutOptions,
              abort: new AbortController(),
              firstPublication: deferred<void>(),
              published: false,
              settled: false,
            }
          : undefined,
      );
      if (mode === 'worker' && doc._mode === 'main') {
        metrics.setMode('main');
        console.warn(
          "[ooxml] mode: 'worker' fell back to main-thread rendering because this document requires DOM OpenType vertical glyph selection.",
        );
      }
      if (opts.math && doc._mode === 'worker' && !rendererDescriptors?.math) {
        console.warn(
          "[ooxml] a custom math renderer cannot cross the worker boundary; equations will be skipped in mode: 'worker'. Use the math renderer from @silurus/ooxml/math.",
        );
      }
      if (opts.threeD && doc._mode === 'worker' && !rendererDescriptors?.threeD) {
        console.warn(
          "[ooxml] a custom 3-D chart renderer cannot cross the worker boundary; charts use their 2-D family fallback in mode: 'worker'. Use the renderer from @silurus/ooxml/three-d.",
        );
      }
      doc._threeD = doc._mode === 'worker' ? undefined : opts.threeD;
      if (opts.regionMap && doc._mode === 'worker' && !rendererDescriptors?.regionMap) {
        console.warn(
          "[ooxml] a custom Region Map renderer cannot cross the worker boundary; geospatial charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/region-map.",
        );
      }
      doc._regionMap = doc._mode === 'worker' ? undefined : opts.regionMap;
      if (opts.chartEx && doc._mode === 'worker' && !rendererDescriptors?.chartEx) {
        console.warn(
          "[ooxml] a custom ChartEx renderer cannot cross the worker boundary; ChartEx charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/chart-ex.",
        );
      }
      doc._chartEx = doc._mode === 'worker' ? undefined : opts.chartEx;
      if (opts.tiff && doc._mode === 'worker' && !rendererDescriptors?.tiff) {
        console.warn(
          "[ooxml] a custom TIFF codec cannot cross the worker boundary; recognized TIFF images will use an unavailable-image placeholder in mode: 'worker'. Use the codec from @silurus/ooxml/tiff to display them.",
        );
      }
      doc._tiff = doc._mode === 'worker' ? undefined : opts.tiff;
      if (doc._mode === 'main' && opts.useGoogleFonts && doc._document) {
        doc._googleFontFaces = await preloadGoogleFonts(
          docxFontPreloadNames(doc._document),
          DOCX_GOOGLE_FONTS,
        );
      }
      // ECMA-376 §17.8.1 / §17.8.3 — register the document's embedded fonts (via
      // the worker's zip-entry extraction) before the lazy first pagination, so
      // text measures/draws with the authored typeface. Worker mode does this
      // inside the worker (before it paginates); here it runs on the main thread.
      if (doc._mode === 'main' && doc._document?.embeddedFonts?.length) {
        const loadingDocument = doc;
        doc._embeddedFontFaces = await loadEmbeddedFonts(
          doc._document,
          (p) => loadingDocument.getFontBytes(p),
        );
      }
      let localMetrics: Awaited<ReturnType<typeof loadDocxLocalFontMetrics>> | undefined;
      if (doc._mode === 'main' && doc._document) {
        localMetrics = await loadDocxLocalFontMetrics(doc._document);
        doc._localMetricFontFaces = localMetrics.faces;
      }
      // Equations are converted + rasterized before pagination (which reads their
      // extents synchronously). Requires the opt-in `math` engine; without it,
      // equations are skipped (and the engine asset is never bundled). Worker
      // mode performs the same preparation with the renderer's imported engine.
      let preparedMath;
      if (doc._mode === 'main' && opts.math && doc._document && documentHasMath(doc._document)) {
        preparedMath = await prepareMathRuns(doc._document, opts.math);
      }
      if (doc._mode === 'main' && doc._document && doc._source) {
        const layoutDocument = doc;
        const runtime = documentLayoutRuntimeOf(doc);
        runtime.services = createLayoutServices(doc._source, {
          localMetrics: localMetrics?.metrics,
          useGoogleFonts: !!opts.useGoogleFonts,
          embeddedFaces: doc._embeddedFontFaces,
          googleFaces: doc._googleFontFaces,
          mathResources: preparedMath?.records,
          mathDrawables: preparedMath?.drawables,
        });
        const services = runtime.services;
        const retained = retainRenderWorkerDocumentLayout(
          doc._source,
          services,
          runtime.defaultCurrentDateMs,
        );
        // Worker mode must build this layout to return parsedMeta. Main mode does
        // the same work here so layout failures reject load() in both modes.
        //
        // Sliced when asked: the same pagination generator, drained across
        // event-loop turns instead of in one blocking call, then deposited in
        // the variant store so every later synchronous render selects it
        // normally. The layout is identical either way.
        // A fatally-unparseable document is served a synthetic error page by the
        // variant store's builder rather than being paginated at all; neither
        // slicing nor previewing may route around that substitution.
        const deferrable = doc._source.fatalParse === null;
        // The variant the caller will actually render, recorded as the active
        // view above BEFORE any geometry read (the metrics snapshot below
        // reads `pageCount`), so that priming, the store lookup on first
        // render, and every geometry accessor all agree on one key.
        const layoutOptions = runtime.activeLayoutOptions;
        if (!layoutOptions) throw new Error('Active layout view was not recorded at load');
        const scheduler = {
          onProgress: opts.onLayoutProgress
            ? (committedPages: number) => layoutDocument._layoutObservers.notify(
              'onLayoutProgress', opts.onLayoutProgress, { committedUnits: committedPages },
            )
            : undefined,
        };
        if (deferrable && opts.progressiveLayout) {
          const store = retained.layoutVariants;
          // Narrowed once: the closures below outlive this block's control flow.
          const progressiveDocument = doc;
          const abort = new AbortController();
          progressiveDocument._layoutAbort = abort;
          // The opening checkpoint is itself laid out in scheduler slices, so the
          // first publication arrives asynchronously — this deferred is what
          // load() resolves on, exactly as the worker path's firstPublication.
          const firstPublication = deferred<void>();
          let publishedLayout: DeepReadonly<DocumentLayout> | null = null;
          let ownsPublication = true;
          const full = layoutDocumentProgressively(
            doc._source.bodyLayoutInput,
            services,
            layoutOptions,
            {
              scheduler: { ...scheduler, signal: abort.signal },
              onPreview: (preview) => {
                if (!ownsPublication) return;
                const first = publishedLayout === null;
                const retainedPreview = progressiveDocument._replaceMainLayoutPublication(
                  store,
                  layoutOptions,
                  publishedLayout,
                  preview.layout,
                );
                if (retainedPreview === null) {
                  // A view change may evict this variant and a later geometry
                  // read may rebuild it authoritatively. The old drain has then
                  // lost its publication token and must never overwrite it.
                  ownsPublication = false;
                  return;
                }
                publishedLayout = retainedPreview;
                if (!first) {
                  // This background session still owns its store entry, but a
                  // runtime view switch may have selected a different entry.
                  // Never describe the inactive session as current document
                  // geometry; switching back reads its latest primed prefix.
                  if (!progressiveDocument._isLayoutViewActive(layoutOptions)) return;
                  // A later checkpoint: more pages are now available.
                  publishDocxLayout(progressiveDocument, {
                    pageCount: preview.layout.pages.length,
                    exact: preview.exact,
                    complete: false,
                  });
                  progressiveDocument._layoutObservers.notify('onLayoutPartial', opts.onLayoutPartial, {
                    availableUnits: preview.layout.pages.length,
                    exact: preview.exact,
                  });
                } else {
                  progressiveDocument._layoutLifecycle.begin();
                  publishDocxLayout(progressiveDocument, {
                    pageCount: preview.layout.pages.length,
                    exact: preview.exact,
                    complete: false,
                  });
                  firstPublication.resolve();
                }
              },
            },
          ).then((layout) => {
            // Replace only the exact prefix this drain last published. A newer
            // synchronous rebuild of the same variant owns the key otherwise.
            if (ownsPublication) {
              const authoritative = progressiveDocument._replaceMainLayoutPublication(
                store,
                layoutOptions,
                publishedLayout,
                layout,
              );
              if (authoritative === null) ownsPublication = false;
            }
            progressiveDocument._layoutLifecycle.succeed();
            publishDocxLayout(progressiveDocument, {
              // A runtime view switch can complete before this original
              // session. Publish the geometry the document actually exposes,
              // not the just-finished inactive session's page count.
              pageCount: progressiveDocument.pageCount,
              exact: true,
              complete: true,
            });
            if (publishedLayout !== null) {
              progressiveDocument._layoutObservers.notify(
                'onLayoutComplete', opts.onLayoutComplete,
              );
            }
            // Nothing was published: there was nothing to show early, so
            // load() resolves here, on the layout that would have been built
            // anyway. Resolving an already-resolved deferred is a no-op.
            firstPublication.resolve();
          });
          // Never awaited raw: once a publication resolves load(), a later
          // failure can no longer reject it and must surface through
          // waitUntilLayoutComplete() instead of as an unhandled rejection.
          progressiveDocument._layoutCompletion = full.catch((error: unknown) => {
            if (publishedLayout === null) {
              // Nothing was shown early, so this is still load()'s own
              // rejection — including an abort, which for an un-resolved
              // load() means the caller's await must not hang forever.
              firstPublication.reject(error);
              return;
            }
            // An aborted drain means the document was destroyed or replaced,
            // not that layout failed. Settle quietly: there is nobody left to
            // tell, and `waitUntilLayoutComplete` must not reject for it.
            if (error instanceof PaginationAbortError) {
              progressiveDocument._layoutLifecycle.succeed();
              return;
            }
            const layoutError = progressiveDocument._layoutLifecycle.fail(error);
            publishDocxLayout(progressiveDocument, {
              pageCount: progressiveDocument.pageCount,
              exact: false,
              complete: false,
              error: layoutError,
            });
            progressiveDocument._layoutObservers.notify(
              'onLayoutComplete', opts.onLayoutComplete, layoutError,
            );
          });
          await firstPublication.promise;
        } else if (deferrable && (opts.sliceLayout || opts.onLayoutProgress)) {
          const layout = await layoutDocumentInputAsync(
            doc._source.bodyLayoutInput,
            services,
            layoutOptions,
            scheduler,
          );
          retained.layoutVariants.prime(layoutOptions, layout);
        } else {
          // Build the variant that will be rendered, not the default one.
          retained.layoutVariants.layoutFor(layoutOptions);
        }
      }
      // This final snapshot includes eager embedded-font extraction performed
      // after the parse response. Telemetry is strictly best-effort: a worker
      // failure or a silent worker may omit the newest counters, but must not
      // turn an otherwise successful load into a rejection or an endless wait.
      await doc._resourceUsage(
        opts.workerTimeoutMs ?? OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS,
      ).then(
        (usage) => metrics.observeUsage(usage),
        () => undefined,
      );
      metrics.checkpoint('model and layout ready');
      metrics.succeed({ pages: doc.pageCount });
      return doc;
    } catch (error) {
      const rejectedDocument = doc;
      disposeRejectedLoad(worker, rejectedDocument ? () => rejectedDocument.destroy() : undefined);
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
    progressive?: WorkerProgressiveLoad,
  ): Promise<void> {
    if (progressive) {
      await this._parseProgressively(
        buffer,
        resourcePolicy,
        useGoogleFonts,
        timeoutMs,
        onUsage,
        renderers,
        progressive,
      );
      return;
    }
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ type: 'parse', id, data: buffer, resourcePolicy, useGoogleFonts, defaultCurrentDateMs: documentLayoutRuntimeOf(this).defaultCurrentDateMs, ...this._parseViewFields(), renderers } satisfies RenderWorkerRequest)
          : ({ type: 'parse', id, data: buffer, resourcePolicy } satisfies WorkerRequest),
      [buffer],
      { timeoutMs },
    );
    if ('protocol' in res) {
      throw new Error('DOCX parse open returned a pull-protocol response');
    }
    if (this._mode === 'worker') {
      if ('usage' in res && res.usage) onUsage?.(res.usage);
      if (res.type === 'mainThreadVerticalFallback') {
        const adapted = await materializeDocumentPullAdapterSession(
          this._bridge.transport(isDocumentPullResponse),
          res,
          { timeoutMs, onUsage },
        );
        this._source = adapted.source;
        this._document = adapted.document;
        this._meta = null;
        this._mode = 'main';
      } else {
        this._meta = (res as Extract<RenderWorkerResponse, { type: 'parsedMeta' }>).meta;
      }
    } else {
      const identity = res as Extract<WorkerResponse, { type: 'documentSessionOpened' }>;
      const adapted = await materializeDocumentPullAdapterSession(
        this._bridge.transport(isDocumentPullResponse),
        identity,
        { timeoutMs, onUsage },
      );
      this._source = adapted.source;
      this._document = adapted.document;
    }
    this._review = snapshotReviewData(
      this._meta?.comments ?? this._document?.comments ?? [],
      this._meta?.revisions ?? this._document?.revisions ?? [],
    );
  }

  /**
   * Route an uncorrelated worker push.
   *
   * Only pushes naming the in-flight `parse` are acted on: request ids are
   * monotonic and never reused, so a `forId` from any other parse belongs to a
   * load this document has already moved past.
   */
  private _onWorkerLayoutPush(res: WorkerResponse | RenderWorkerResponse): void {
    if (respondToWorkerSvgDecodeRequest(
      (message, transfer) => (
        this._worker.postMessage as (value: unknown, transfer?: Transferable[]) => void
      )(message, transfer),
      res,
    )) return;
    if (!('forId' in res) || res.forId !== this._parseRequestId) return;
    // Any push proves the worker is alive and working, which is the only thing
    // the watchdog is asking about.
    this._rearmParseWatchdog();
    const progressive = this._progressive;
    if (progressive?.settled) return;
    if (res.type === 'layoutProgress') {
      this._layoutObservers.notify(
        'onLayoutProgress', progressive?.onProgress, { committedUnits: res.committedPages },
      );
      return;
    }
    if (res.type !== 'layoutPartial' || !progressive) return;
    // The worker continues its load-time variant after setLayoutView() builds a
    // different one. The original session may keep priming its own store entry,
    // but its pushes no longer own the host's synchronous geometry.
    const ownsActiveView = !progressive.published
      || this._isLayoutViewActive(progressive.layoutOptions);
    if (ownsActiveView) this._applyLayoutPartial(res.partial);
    if (progressive.published) {
      if (!ownsActiveView) return;
      publishDocxLayout(this, {
        pageCount: res.partial.pageCount,
        exact: res.partial.exact,
        complete: false,
      });
      this._layoutObservers.notify('onLayoutPartial', progressive.onPartial, {
        availableUnits: res.partial.pageCount,
        exact: res.partial.exact,
      });
      return;
    }
    // The first publication is what `load()` has been waiting for. It fires no
    // `onLayoutPartial`: the caller is about to see these pages as the loaded
    // document, not as an extension of one — same rule main mode follows.
    progressive.published = true;
    this._layoutLifecycle.begin();
    publishDocxLayout(this, {
      pageCount: res.partial.pageCount,
      exact: res.partial.exact,
      complete: false,
    });
    progressive.firstPublication.resolve();
  }

  /**
   * Fold a provisional prefix into the worker metadata the geometry accessors
   * read.
   *
   * Replaces rather than mutates: `pageCount` and `pageSizes` have to move
   * together, and a reader that caught one without the other would size a
   * scroll extent against the wrong page list.
   */
  private _applyLayoutPartial(partial: DocumentLayoutPartial): void {
    // `review` rides on the first publication only; later ones keep what the
    // first established.
    const review = partial.review ?? this._meta;
    this._meta = {
      pageCount: partial.pageCount,
      pageSizes: partial.pageSizes,
      bookmarkPages: partial.bookmarkPages,
      revisions: review?.revisions ?? [],
      comments: review?.comments ?? [],
      footnotes: review?.footnotes ?? [],
      endnotes: review?.endnotes ?? [],
      commentAnchorRanges: partial.commentAnchorRanges,
      revisionAnchorRanges: partial.revisionAnchorRanges,
    };
    if (partial.review) {
      this._review = snapshotReviewData(partial.review.comments, partial.review.revisions);
    }
    this._invalidateLayoutDerivedCaches();
  }

  /** Install the authoritative metadata and close the progressive window. */
  private _onAuthoritativeMeta(meta: DocumentMeta): void {
    this._clearParseWatchdog();
    const progressive = this._progressive;
    if (progressive?.settled) return;
    if (!progressive || this._isLayoutViewActive(progressive.layoutOptions) || !this._meta) {
      this._meta = meta;
    } else {
      // Model-derived data belongs to the document and can be refreshed from
      // the final parse response. Layout-derived fields must stay on the
      // runtime-selected variant installed by setLayoutView().
      const selected = this._meta;
      this._meta = {
        ...meta,
        pageCount: selected.pageCount,
        pageSizes: selected.pageSizes,
        bookmarkPages: selected.bookmarkPages,
        commentAnchorRanges: selected.commentAnchorRanges,
        revisionAnchorRanges: selected.revisionAnchorRanges,
      };
    }
    this._invalidateLayoutDerivedCaches();
    this._review = snapshotReviewData(meta.comments, meta.revisions);
    if (!progressive) return;
    progressive.settled = true;
    this._layoutLifecycle.succeed();
    publishDocxLayout(this, {
      pageCount: this.pageCount,
      exact: true,
      complete: true,
    });
    // A load whose worker published nothing resolves here instead — there was
    // never anything to show early, so `load()` waited for the real document.
    progressive.firstPublication.resolve();
    if (progressive.published) {
      this._layoutObservers.notify('onLayoutComplete', progressive.onComplete);
    }
  }

  /** Bookmark pages and the review anchor projections are derived from the
   *  layout, so they belong to the publication that produced them. */
  private _invalidateLayoutDerivedCaches(): void {
    this._bookmarkPages = null;
    this._commentAnchorRanges = null;
    this._revisionAnchorRanges = null;
    this._reviewProjectionIndex = null;
  }

  /** Atomically advance one main-thread layout publication and invalidate only
   * the derived projections that belong to the layout the document currently
   * exposes. An inactive progressive variant may continue draining in the
   * background, but it must neither stale nor evict the active variant's
   * bookmark/review caches. */
  private _replaceMainLayoutPublication(
    store: LayoutVariantStore,
    options: LayoutOptions,
    expected: DeepReadonly<DocumentLayout> | null,
    next: DeepReadonly<DocumentLayout>,
  ): DeepReadonly<DocumentLayout> | null {
    const retained = store.replaceIfCurrent(options, expected, next);
    if (retained !== null && this._isLayoutViewActive(options)) {
      this._invalidateLayoutDerivedCaches();
    }
    return retained;
  }

  /** Whether synchronous geometry and paint currently select this normalized
   *  layout view. Optional false and absent are the same layout key. */
  private _isLayoutViewActive(options: LayoutOptions): boolean {
    return sameLayoutView(documentLayoutRuntimeOf(this).activeLayoutOptions, options);
  }

  private _rearmParseWatchdog(): void {
    if (this._parseWatchdogMs === undefined) return;
    clearTimeout(this._parseWatchdog);
    this._parseWatchdog = setTimeout(() => this._onParseWentSilent(), this._parseWatchdogMs);
  }

  private _clearParseWatchdog(): void {
    clearTimeout(this._parseWatchdog);
    this._parseWatchdog = undefined;
    this._parseWatchdogMs = undefined;
  }

  /** The worker produced neither a publication nor a heartbeat within the
   *  caller's timeout. A background layout may legitimately run for a long
   *  time, but it may not go quiet: treat silence as a wedged worker. */
  private _onParseWentSilent(): void {
    const progressive = this._progressive;
    const silenceMs = this._parseWatchdogMs;
    this._clearParseWatchdog();
    const error = new Error(
      `worker layout produced no progress for ${silenceMs}ms`,
    );
    if (progressive) this._failWorkerProgressive(progressive, error);
    this._bridge.terminate();
  }

  /** Single terminal failure authority for worker progressive layout. The
   * watchdog terminates the worker after calling this method; the resulting
   * bridge rejection is therefore a duplicate and is ignored by `settled`. */
  private _failWorkerProgressive(
    progressive: WorkerProgressiveLoad,
    cause: unknown,
  ): void {
    if (progressive.settled) return;
    progressive.settled = true;
    this._clearParseWatchdog();
    if (!progressive.published) {
      progressive.firstPublication.reject(cause);
      return;
    }
    const layoutError = this._layoutLifecycle.fail(cause);
    publishDocxLayout(this, {
      pageCount: this.pageCount,
      exact: false,
      complete: false,
      error: layoutError,
    });
    this._layoutObservers.notify('onLayoutComplete', progressive.onComplete, layoutError);
  }

  /** The variant fields the worker `parse` carries, derived from the SAME
   *  recorded active options every render request is filled from — so the
   *  worker cannot end up paginating one view and painting another. Omitted
   *  when the load selects the document's default view, keeping the wire shape
   *  identical to what pre-variant builds sent. */
  private _parseViewFields(): Pick<
    Extract<RenderWorkerRequest, { type: 'parse' }>,
    'currentDateMs' | 'showTrackedChanges'
  > {
    const runtime = documentLayoutRuntimeOf(this);
    const active = runtime.activeLayoutOptions;
    if (!active) return {};
    return {
      ...(active.currentDateMs === runtime.defaultCurrentDateMs
        ? {}
        : { currentDateMs: active.currentDateMs }),
      ...(active.showTrackedChanges === true ? { showTrackedChanges: true } : {}),
    };
  }

  /**
   * Worker-mode progressive parse: resolve as soon as the worker publishes a
   * paintable prefix, and let the authoritative layout land in the background.
   *
   * The `parse` request deliberately gives up its own deadline. Deferring
   * `parsedMeta` until the whole document is paginated would otherwise turn
   * `workerTimeoutMs` into "the entire layout must finish within it" — failing
   * on exactly the large documents progressive layout exists for. Liveness is
   * enforced by the silence watchdog instead, which is what that timeout was
   * really protecting against.
   */
  private async _parseProgressively(
    buffer: ArrayBuffer,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    useGoogleFonts: boolean,
    timeoutMs: number | undefined,
    onUsage: ((usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void) | undefined,
    renderers: WorkerRendererDescriptors | undefined,
    progressive: WorkerProgressiveLoad,
  ): Promise<void> {
    this._progressive = progressive;
    this._layoutAbort = progressive.abort;
    this._parseWatchdogMs = timeoutMs;
    const parsed = this._bridge.request(
      (id) => {
        this._parseRequestId = id;
        return {
          type: 'parse',
          id,
          data: buffer,
          resourcePolicy,
          useGoogleFonts,
          defaultCurrentDateMs: documentLayoutRuntimeOf(this).defaultCurrentDateMs,
          ...this._parseViewFields(),
          renderers,
          progressiveLayout: true,
        } satisfies RenderWorkerRequest;
      },
      [buffer],
      { timeoutMs: false },
    );
    // Retained rather than awaited: once a publication resolves load(), a later
    // failure can no longer reject it, and must surface through
    // waitUntilLayoutComplete() instead of as an unhandled rejection.
    this._layoutCompletion = parsed.then(
      async (res) => {
        this._parseRequestId = null;
        if ('protocol' in res) {
          throw new Error('DOCX parse open returned a pull-protocol response');
        }
        if ('usage' in res && res.usage) onUsage?.(res.usage);
        if (res.type === 'mainThreadVerticalFallback') {
          // This document needs DOM OpenType vertical glyph selection, so the
          // worker never paginated it and never published anything. Demote to
          // main mode exactly as the ordinary parse does; load()'s main-mode
          // block then lays it out — progressively — on this thread.
          this._clearParseWatchdog();
          progressive.settled = true;
          this._progressive = null;
          this._layoutAbort = null;
          const adapted = await materializeDocumentPullAdapterSession(
            this._bridge.transport(isDocumentPullResponse),
            res,
            { timeoutMs, onUsage },
          );
          this._source = adapted.source;
          this._document = adapted.document;
          this._meta = null;
          this._mode = 'main';
          this._review = snapshotReviewData(
            adapted.document.comments ?? [],
            adapted.document.revisions ?? [],
          );
          progressive.firstPublication.resolve();
          return;
        }
        this._onAuthoritativeMeta(
          (res as Extract<RenderWorkerResponse, { type: 'parsedMeta' }>).meta,
        );
      },
      (error: unknown) => {
        this._parseRequestId = null;
        this._clearParseWatchdog();
        if (progressive.settled) return;
        // Destroyed or replaced mid-layout: the worker was terminated on
        // purpose, there is nobody left to tell, and waitUntilLayoutComplete() must
        // not reject for it. Exactly how main mode treats PaginationAbortError.
        if (progressive.abort.signal.aborted) {
          progressive.settled = true;
          this._layoutLifecycle.succeed();
          progressive.firstPublication.resolve();
          return;
        }
        this._failWorkerProgressive(progressive, error);
      },
    );
    this._rearmParseWatchdog();
    await progressive.firstPublication.promise;
  }

  destroy(): void {
    // Stop background layout first: without this, a destroyed document's
    // remaining pagination kept consuming main-thread slices to completion for
    // a viewer that no longer exists.
    this._layoutAbort?.abort();
    this._layoutAbort = null;
    // In worker mode the abort above only marks intent; terminate() is what
    // actually stops the worker's background pagination. The pending `parse`
    // then rejects with 'Worker terminated', which the progressive rejection
    // handler recognises as a deliberate teardown and settles quietly.
    this._clearParseWatchdog();
    this._parseRequestId = null;
    this._progressive = null;
    this._layoutViewGeneration++;
    this._bridge.terminate();
    this._document = null;
    this._source = null;
    this._meta = null;
    this._review = EMPTY_REVIEW_SNAPSHOT;
    documentLayoutRuntimeOf(this).services = null;
    this._bookmarkPages = null;
    this._commentAnchorRanges = null;
    this._revisionAnchorRanges = null;
    this._reviewProjectionIndex = null;
    this._rawParts.clear();
    // Release the embedded fonts this document added to the shared FontFaceSet
    // (main mode). Refcounted in core: a font also used by another open document
    // stays until that one is destroyed too. Without this, every opened document
    // left its FontFace objects in `document.fonts` forever (SPA memory leak).
    if (this._embeddedFontFaces.length > 0) {
      unregisterEmbeddedFonts(this._embeddedFontFaces);
      this._embeddedFontFaces = [];
    }
    // Release the Google-Fonts substitutes this document preloaded into the
    // shared FontFaceSet (main mode). Same refcount contract as the embedded
    // fonts: a web font also used by another open document stays until that one
    // is destroyed too. Without this, every opened document left its Google
    // FontFace objects in `document.fonts` forever (SPA memory leak).
    if (this._googleFontFaces.length > 0) {
      unloadGoogleFonts(this._googleFontFaces);
      this._googleFontFaces = [];
    }
    if (this._localMetricFontFaces.length > 0) {
      unloadLocalFontMetrics(this._localMetricFontFaces);
      this._localMetricFontFaces = [];
    }
    // Release both image owners keyed by this document's stable loader: the
    // shared decoded owner (base + derived colour surfaces) and the SVG lookup
    // owner. SVG object URLs are revoked immediately after decode; dropping its
    // lookup releases retained decoded elements and prevents stale reuse.
    dropDecodedBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
  }

  /**
   * Extract raw bytes for an embedded image by zip path (e.g.
   * `word/media/image1.png`), wrapped in a Blob of the given MIME type. Routes
   * through the persistent worker via the `extractImage` message (twin of
   * pptx's `getImage`/`getMedia`); results are cached by path for the lifetime
   * of this instance. The renderer's `fetchImage` option points here so images
   * are decoded lazily rather than inlined as base64 at parse time.
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    return this._rawParts.get(imagePath, mimeType, () => this._bridge
      .request((id) => ({ type: 'extractImage', id, path: imagePath }) satisfies WorkerRequest)
      .then((res) => {
        const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      }));
  }

  /**
   * Extract raw bytes for an embedded font part by zip path (e.g.
   * `word/fonts/font1.odttf`). Routes through the SAME persistent-worker
   * `extractImage` message as {@link getImage} — `DocxArchive.extract_image`
   * reads ANY zip entry, not just media — returning the raw (still obfuscated)
   * `.odttf` bytes rather than a Blob. Consumed by {@link loadEmbeddedFonts},
   * which de-obfuscates (ECMA-376 §17.8.1) and registers each as a FontFace.
   */
  async getFontBytes(partPath: string): Promise<Uint8Array> {
    const res = await this._bridge.request(
      (id) => ({ type: 'extractImage', id, path: partPath }) satisfies WorkerRequest,
    );
    const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
    return new Uint8Array(bytes);
  }

  private async _resourceUsage(
    timeoutMs: number,
  ): Promise<import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot> {
    const res = await this._bridge.request(
      (id) => ({ type: 'resourceUsage', id }) satisfies WorkerRequest,
      undefined,
      { timeoutMs },
    );
    return (res as Extract<WorkerResponse, { type: 'resourceUsage' }>).usage;
  }

  /** Return a fresh content-free metrics snapshot, including lazy archive work
   * completed since load. Collection is always active; `debug` only controls
   * console presentation. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    const metrics = this._metrics;
    if (!metrics) throw new Error('Document not loaded');
    return readLatestOoxmlResourceMetrics(metrics, (timeoutMs) => this._resourceUsage(timeoutMs));
  }

  /**
   * Project the document to GitHub-flavoured markdown: headings (from
   * `<w:outlineLvl>`), bullet / numbered lists, tables (with vMerge
   * continuation), and rich-text formatting (bold / italic / strikethrough /
   * hyperlink), with footnotes / endnotes / comments collated at the end.
   * Positioning, section properties, fonts, and drawing shapes are discarded —
   * the projection is meant for AI ingestion and full-text search, not layout.
   *
   * Runs entirely in the worker off the archive opened at {@link load} (no
   * re-copy of the file, no re-parse of the model on the main thread), so it
   * works in BOTH `mode: 'main'` and `mode: 'worker'`.
   *
   * @example
   * const doc = await DocxDocument.load(buffer);
   * const md = await doc.toMarkdown();
   */
  async toMarkdown(): Promise<string> {
    const res = await this._bridge.request(
      (id) => ({ type: 'toMarkdown', id }) satisfies WorkerRequest,
    );
    return (res as Extract<WorkerResponse, { type: 'markdownRendered' }>).markdown;
  }

  get pageCount(): number {
    if (this._meta) return this._meta.pageCount;
    if (!this._document) return 0;
    return this._getLayout()?.pages.length ?? 0;
  }

  /**
   * Whether every page has been laid out.
   *
   * True only after the authoritative layout succeeds. Under
   * `progressiveLayout`, it is false between `load()` resolving on the opening
   * pages and successful completion, and remains false after a background
   * failure. While false, {@link pageCount} is not an authoritative total;
   * {@link waitUntilLayoutComplete} distinguishes pending work from failure.
   */
  get layoutComplete(): boolean {
    return this._layoutLifecycle.complete;
  }

  /**
   * Resolve once the whole document is laid out.
   *
   * Await this before anything that must see every page — text search, bookmark
   * navigation, page-count UI, print or export. Resolves immediately for an
   * ordinary load. Re-throws a background layout failure, which cannot reject
   * `load()` because that already resolved.
   */
  async waitUntilLayoutComplete(): Promise<void> {
    if (this._layoutCompletion) await this._layoutCompletion;
    this._layoutLifecycle.throwIfFailed();
  }

  /** The render mode this engine was loaded with ('main' | 'worker'). A fact for
   *  integrators and the scroll viewer: an injected engine's mode decides whether
   *  pages render via renderPage (main) or renderPageToBitmap (worker) — no
   *  probing (design §11: no silent mis-pathing). */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }

  /**
   * The raw parsed document model. Available only in `mode: 'main'`; in
   * `mode: 'worker'` the model stays in the worker and this throws.
   */
  get document(): DocxDocumentModel {
    if (this._meta && !this._document) {
      throw new Error(
        "the raw document model stays in the worker in mode: 'worker'; use mode: 'main' if you need direct model access",
      );
    }
    if (!this._document) throw new Error('Document not loaded');
    return this._document;
  }

  /**
   * ECMA-376 §17.13.4 — the document's comments (`word/comments.xml`), each with
   * id / author / initials / date / plain-text body. Use this low-level API to
   * build a custom review panel or export; {@link DocxScrollViewer} can also
   * provide an opt-in read-only margin. Returns `[]` when the document has no
   * comments part. The same data is also reachable via `document.comments`.
   */
  get comments(): readonly Readonly<DocComment>[] {
    return this._reviewSnapshot().comments;
  }

  /** ECMA-376 §17.13.5 body-story revision events in document order. Available
   * in both main and worker modes. Rendering uses the accepted final state by
   * default; `showTrackedChanges` opts into the markup projection. Consumers may
   * use these detached records in their own review UI. */
  get revisions(): readonly Readonly<DocRevision>[] {
    return this._reviewSnapshot().revisions;
  }

  /**
   * ECMA-376 §17.13.4 — the comment anchors resolved to per-paragraph run
   * intervals in document order (`commentRangeStart`/`End` pairs, plus point
   * anchors). Each range carries the exact story/path identity used by
   * `DocxTextRunInfo.source`, covering body, headers, footers, notes, and text
   * boxes. Join it to rendered geometry with `sourceRunIndex`. Mode-agnostic:
   * main mode resolves lazily from the retained source; worker mode returns the
   * same projection in metadata. During progressive loading it covers the
   * currently available page prefix; await {@link waitUntilLayoutComplete}
   * before treating an empty or partial result as whole-document authority.
   */
  commentAnchorRanges(): readonly CommentAnchorRange[] {
    if (this._meta) return this._meta.commentAnchorRanges ?? NO_COMMENT_ANCHOR_RANGES;
    if (!this._document || !this._source) return [];
    const comments = this._reviewSnapshot().comments;
    if (comments.length === 0) return NO_COMMENT_ANCHOR_RANGES;
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    // The ACTIVE variant's layout: on a document loaded with an explicit
    // `currentDate` or the markup view, reading `defaultLayout` here would
    // silently build a whole variant nobody is rendering — and index text runs
    // that do not match the visible ones.
    const store = layoutVariantStoreOf(services);
    if (!store) throw new Error('Document layout variant store is not initialized');
    const active = runtime.activeLayoutOptions;
    const layout = active ? store.layoutFor(active) : store.defaultLayout;
    this._reviewProjectionIndex ??= reviewProjectionIndexForDocument(layout);
    const projectionOptions = (this._layoutLifecycle?.complete ?? true)
      ? undefined
      : {
          completedSourceKeys: this._reviewProjectionIndex.completedSourceKeys,
        };
    this._commentAnchorRanges ??= collectLayoutSourceCommentRangesIfPresent(
      comments,
      this._source,
      this._reviewProjectionIndex.renderedRunIndex,
      projectionOptions,
    );
    return this._commentAnchorRanges;
  }

  /** ECMA-376 §17.13.5 revision containers resolved to normalized source-run
   * intervals. Join them to `collectPageRuns()` with
   * `resolveRevisionAnchorRuns()`. Deletions and move sources use the nearest
   * deterministic final-state text position because accepted rendering gives
   * their own content no geometry. Mode-agnostic. During progressive loading
   * this projects the available prefix; await {@link waitUntilLayoutComplete}
   * before a whole-document review scan. */
  revisionAnchorRanges(): readonly RevisionAnchorRange[] {
    if (this._meta) return this._meta.revisionAnchorRanges ?? NO_REVISION_ANCHOR_RANGES;
    if (!this._document || !this._source) return [];
    const revisions = this._reviewSnapshot().revisions;
    if (revisions.length === 0) return NO_REVISION_ANCHOR_RANGES;
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    // Same active-variant selection as commentAnchorRanges above.
    const store = layoutVariantStoreOf(services);
    if (!store) throw new Error('Document layout variant store is not initialized');
    const active = runtime.activeLayoutOptions;
    const layout = active ? store.layoutFor(active) : store.defaultLayout;
    this._reviewProjectionIndex ??= reviewProjectionIndexForDocument(layout);
    const projectionOptions = (this._layoutLifecycle?.complete ?? true)
      ? undefined
      : {
          completedSourceKeys: this._reviewProjectionIndex.completedSourceKeys,
        };
    this._revisionAnchorRanges ??= collectLayoutSourceRevisionRangesIfPresent(
      revisions,
      this._source,
      this._reviewProjectionIndex.renderedRunIndex,
      projectionOptions,
    );
    return this._revisionAnchorRanges;
  }

  /** Object.create-based focused tests and older deserialized instances can
   * bypass field initializers; preserve the same one-time immutable contract. */
  private _reviewSnapshot(): ReviewSnapshot {
    if (this._review) return this._review;
    this._review = snapshotReviewData(
      this._meta?.comments ?? this._document?.comments ?? [],
      this._meta?.revisions ?? this._document?.revisions ?? [],
    );
    return this._review;
  }

  /**
   * ECMA-376 §17.11.10 — the document's footnotes (`word/footnotes.xml`),
   * excluding the reserved separator entries. Each note carries its `id` and
   * block-level `content`; use {@link noteText} for the plain-text body. These
   * ARE drawn at the bottom of the page that holds their reference; this getter
   * additionally exposes them as data. Returns `[]` when absent.
   */
  get footnotes(): DocNote[] {
    return this._meta?.footnotes ?? this._document?.footnotes ?? [];
  }

  /**
   * ECMA-376 §17.11.4 — the document's endnotes (`word/endnotes.xml`). Same
   * shape as {@link footnotes}; rendered at the end of the document. Returns
   * `[]` when absent.
   */
  get endnotes(): DocNote[] {
    return this._meta?.endnotes ?? this._document?.endnotes ?? [];
  }

  private _getLayout(): DeepReadonly<DocumentLayout> | null {
    if (!this._document) return null;
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    const store = layoutVariantStoreOf(services);
    if (!store) throw new Error('Document layout variant store is not initialized');
    // The ACTIVE variant, not the default one: a tracked-changes viewer paints
    // the markup layout, so its page count and page geometry must come from
    // that same layout. Reading the default here also silently paginated the
    // whole document a second time for a variant nobody was viewing.
    const active = runtime.activeLayoutOptions;
    return active ? store.layoutFor(active) : store.defaultLayout;
  }

  /**
   * Select the layout variant this document is viewed as.
   *
   * `showTrackedChanges` and an explicit `currentDate` each select a different
   * retained layout with its own pagination, so the geometry accessors
   * ({@link pageCount}, {@link pageSize}, bookmark lookup) have to follow
   * whichever one the renderer is actually painting. Viewers call this when the
   * user toggles the markup view.
   *
   * In worker mode the switch resolves after the worker has produced matching
   * metadata; until then the previous variant remains active, so synchronous
   * geometry can never describe different pages from paint. Main mode records
   * the variant synchronously and builds it on the next geometry read.
   */
  async setLayoutView(
    view: Readonly<{ showTrackedChanges?: boolean; currentDate?: Date | number }> = {},
  ): Promise<void> {
    // Request ownership is explicit input metadata. It is captured before a
    // worker request yields, stripped by normalization, and never serialized.
    const requester = (view as Readonly<{
      [docxLayoutViewRequester]?: object;
    }>)[docxLayoutViewRequester];
    const runtime = documentLayoutRuntimeOf(this);
    const next = normalizeLayoutOptions(
      view.currentDate,
      runtime.defaultCurrentDateMs,
      view.showTrackedChanges === true,
    );
    const generation = ++this._layoutViewGeneration;
    const active = runtime.activeLayoutOptions;
    if (sameLayoutView(active, next)) return;
    if (this._mode === 'worker' && this._meta) {
      const res = await this._bridge.request(
        (id) => ({
          type: 'selectLayoutView',
          id,
          currentDateMs: next.currentDateMs,
          showTrackedChanges: next.showTrackedChanges === true,
        }) satisfies RenderWorkerRequest,
      );
      if (generation !== this._layoutViewGeneration) return;
      const variant = res as Extract<RenderWorkerResponse, { type: 'layoutViewSelected' }>;
      // Install selection and geometry in one turn. Until this point every
      // synchronous getter and option fill-in remains on the previous variant.
      runtime.activeLayoutOptions = next;
      this._meta = { ...this._meta, ...variant.meta };
      this._invalidateLayoutDerivedCaches();
      publishDocxLayoutView(this, requester);
      return;
    }

    runtime.activeLayoutOptions = next;
    // Bookmark pages and the review anchor caches are derived from the
    // layout, so they belong to the variant that produced them.
    this._invalidateLayoutDerivedCaches();
    publishDocxLayoutView(this, requester);
  }

  /** Lazily build (and cache) the `bookmarkName → page index` map from either
   *  the worker meta (worker mode) or the paginated pages (main mode). */
  private _getBookmarkPages(): Map<string, number> | null {
    if (this._bookmarkPages) return this._bookmarkPages;
    if (this._meta) {
      this._bookmarkPages = new Map(this._meta.bookmarkPages);
      return this._bookmarkPages;
    }
    const layout = this._getLayout();
    if (!layout) return null;
    this._bookmarkPages = buildBookmarkPageMap(layout);
    return this._bookmarkPages;
  }

  /**
   * ECMA-376 §17.13.6.2 / §17.16.23 — resolve a bookmark name (a
   * `<w:hyperlink w:anchor>` internal-link target) to the 0-based index of the
   * page its `<w:bookmarkStart w:name>` destination falls on, or `undefined`
   * when the document has no bookmark of that name. When a bookmark's paragraph
   * spans a page break, the page where it *begins* is returned.
   *
   * This is the map an internal-hyperlink click resolves against: a viewer's
   * `onHyperlinkClick` default (or an integrator) turns the anchor into a page
   * and calls {@link DocxViewer.goToPage} (or scrolls the scroll viewer to it).
   * Works in BOTH `main` and `worker` mode (the map rides along in the worker
   * meta, built from the same paginated pages as `pageSizes`). During
   * progressive loading, `undefined` means only "not in the available prefix";
   * await {@link waitUntilLayoutComplete} before concluding it is absent.
   */
  getBookmarkPage(bookmarkName: string): number | undefined {
    return this._getBookmarkPages()?.get(bookmarkName);
  }

  /**
   * ECMA-376 §17.6.13 / §17.6.11 — the page size (pt) of page `pageIndex`, per
   * section (a mixed portrait/landscape document returns different sizes per page).
   * Available in BOTH modes: worker mode reads the worker-built `pageSizes` meta;
   * main mode reads the paginated pages' stamped geometry. Returns the body-level
   * section size for an out-of-range index (clamped) or a page with no stamped
   * geometry. `{ 0, 0 }` means "not loaded" (before `load()` resolves or after
   * `destroy()`). Returns a fresh object per call — safe to mutate.
   * The recommended way to ask "how big is page i?" for layout.
   */
  pageSize(pageIndex: number): { widthPt: number; heightPt: number } {
    if (this._meta) {
      const sizes = this._meta.pageSizes;
      const clamped = Math.max(0, Math.min(pageIndex, sizes.length - 1));
      const s = sizes[clamped];
      // Copy — never alias the meta's stored object (a caller mutating the
      // return value must not corrupt subsequent reads; main mode below already
      // builds a fresh object per call).
      return s ? { widthPt: s.widthPt, heightPt: s.heightPt } : { widthPt: 0, heightPt: 0 };
    }
    if (!this._document) return { widthPt: 0, heightPt: 0 };
    const layout = this._getLayout();
    if (!layout || layout.pages.length === 0) return { widthPt: 0, heightPt: 0 };
    const clamped = Math.max(0, Math.min(pageIndex, layout.pages.length - 1));
    const geometry = layout.pages[clamped]!.geometry;
    return { widthPt: geometry.widthPt, heightPt: geometry.heightPt };
  }

  /**
   * Fill omitted view axes from the load-time active variant.
   *
   * `load({ currentDate, showTrackedChanges })` primes and records the variant
   * the caller will render precisely so the first render does not synchronously
   * repaginate; a per-call selection derived from omitted options would pick
   * the DEFAULT variant instead, paying that repagination anyway and letting
   * paint disagree with the geometry accessors (which follow the active
   * variant). An explicitly passed value still wins — including
   * `showTrackedChanges: false`, which selects the final view regardless of
   * the loaded variant.
   */
  private _withActiveView<
    T extends { currentDate?: Date | number; showTrackedChanges?: boolean },
  >(opts: T): T {
    const active = documentLayoutRuntimeOf(this).activeLayoutOptions;
    if (!active) return opts;
    const filled = { ...opts };
    if (opts.currentDate === undefined) filled.currentDate = active.currentDateMs;
    if (opts.showTrackedChanges === undefined && active.showTrackedChanges === true) {
      filled.showTrackedChanges = true;
    }
    return filled;
  }

  renderPage(
    target: HTMLCanvasElement | OffscreenCanvas,
    pageIndex: number,
    opts: RenderPageOptions = {},
  ): Promise<void> {
    if (this._mode === 'worker') {
      throw new Error(
        "renderPage(canvas) is unavailable in mode: 'worker'; use renderPageToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this._source) throw new Error('Document not loaded');
    return renderLayoutSourceToCanvas(this._source, target, pageIndex, {
      ...this._withActiveView(opts),
      // Lazy image bytes: the renderer fetches each embedded blip on demand by
      // zip path (decoded only when drawn) instead of reading inlined base64.
      fetchImage: this._fetchImage,
      layoutServices: documentLayoutRuntimeOf(this).services ?? undefined,
      defaultCurrentDateMs: documentLayoutRuntimeOf(this).defaultCurrentDateMs,
      threeD: this._threeD,
      regionMap: this._regionMap,
      chartEx: this._chartEx,
      tiff: this._tiff,
    });
  }

  /**
   * Render a page and return it as an ImageBitmap. Works in both modes; in
   * worker mode the render runs entirely off the main thread. Paint with:
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   *
   * IX6 — an optional `onTextRun` in `opts` receives the page's text-run
   * geometry (the same stream `renderPage` emits in main mode), so a caller can
   * build the selection / find overlay from a worker-rendered page on the SAME
   * code path as main mode. In worker mode the runs ride back beside the bitmap
   * (one round-trip, no second render).
   */
  async renderPageToBitmap(
    pageIndex: number,
    opts: RenderPageToBitmapOptions = {},
  ): Promise<ImageBitmap> {
    const { onTextRun, ...wire } = opts;
    const wireOpts: WireRenderPageOptions = {
      ...this._withActiveView(wire),
      dpr: wire.dpr ?? defaultDpr(),
    };
    if (this._mode === 'worker') {
      // The selected date variant may have a different page count than default
      // metadata, so the worker validates against the layout it actually paints.
      const res = await this._bridge.request(
        (id) => ({ type: 'renderPage', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      const rendered = res as Extract<RenderWorkerResponse, { type: 'pageRendered' }>;
      try {
        if (onTextRun) for (const r of rendered.runs) onTextRun(r);
      } catch (error) {
        releaseOwnedBitmap(rendered.bitmap);
        throw error;
      }
      return rendered.bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderPage(off, pageIndex, { ...wireOpts, onTextRun });
    return off.transferToImageBitmap();
  }

  /**
   * Collect a page's text-run geometry (`DocxTextRunInfo[]`) directly from the
   * retained layout. Works in BOTH modes without painting or constructing a
   * Canvas; worker mode ships only the projected plain-data runs. Used by the
   * find controller to scan every page. Geometry uses the same selected layout
   * variant and width scale as `renderPage`; DPR does not change CSS pixels.
   */
  async collectPageRuns(
    pageIndex: number,
    opts: CollectPageRunsOptions = {},
  ): Promise<DocxTextRunInfo[]> {
    const wireOpts: WireRenderPageOptions = { ...this._withActiveView(opts) };
    if (this._mode === 'worker') {
      // Keep collection validation on the same selected worker layout as paint.
      const res = await this._bridge.request(
        (id) => ({ type: 'collectRuns', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'runsCollected' }>).runs;
    }
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    return textRunsForSelectedPage(services, pageIndex, {
      currentDate: wireOpts.currentDate,
      defaultCurrentDateMs: runtime.defaultCurrentDateMs,
      width: wireOpts.width,
      showTrackedChanges: wireOpts.showTrackedChanges,
    });
  }

  /**
   * Resolve the comment threads that have rendered anchor geometry on one page.
   * A range that crosses pages, or a repeating header/footer anchor, appears on
   * every page where it is rendered; each result contains only that page's
   * rectangles. Unanchored comments remain available through {@link comments}
   * but are not returned here.
   */
  async getCommentThreads(
    pageIndex: number,
    options: DocxPageCommentThreadsOptions = {},
  ): Promise<readonly Readonly<ResolvedDocxCommentThread>[]> {
    const { includeResolved, ...runOptions } = options;
    const runs = await this.collectPageRuns(pageIndex, runOptions);
    return resolveDocxCommentThreads(
      this.comments,
      this.commentAnchorRanges(),
      runs,
      { includeResolved },
    );
  }

  /**
   * Hit-test a rendered picture, chart, or shape on demand, including inline
   * resources. No element index is maintained during render, scroll, or pointer
   * movement; worker mode runs the same projection against the worker-owned
   * layout and returns only detached context.
   */
  async getElementContextAt(
    pageIndex: number,
    point: DocxPagePoint,
    opts: DocxElementContextOptions = {},
  ): Promise<DocxElementContext | null> {
    const viewOpts = this._withActiveView(opts);
    if (this._mode === 'worker') {
      const res = await this._bridge.request(
        (id) => ({ type: 'hitTestElement', id, pageIndex, point, opts: viewOpts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'elementHit' }>).context;
    }
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    return hitTestSelectedDocxElementContext(services, pageIndex, point, {
      ...viewOpts,
      defaultCurrentDateMs: runtime.defaultCurrentDateMs,
    });
  }
}
