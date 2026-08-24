import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  unloadLocalFontMetrics,
  unregisterEmbeddedFonts,
  WorkerBridge,
  defaultDpr,
  dropSvgImageCache,
  dropDecodedBitmapCache,
  resolveOoxmlContainer,
  toArrayBuffer,
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
  normalizeLoadResourceOptions,
  OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS,
  OoxmlResourceMetricsSession,
  readLatestOoxmlResourceMetrics,
  PULL_SESSION_PROTOCOL,
  type NormalizedOoxmlResourcePolicy,
  type WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import type { DocxDocumentModel, RenderPageOptions, WorkerRequest, WorkerResponse, DocComment, DocNote } from './types';
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
import type {
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
  WireRenderPageOptions,
} from './worker-protocol';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import {
  hitTestSelectedDocxElementContext,
  type DocxElementContextOptions,
} from './element-context.js';
import type { DocxElementContext, DocxPagePoint } from './selection-context.js';
import {
  collectDocumentCommentRanges,
  type CommentAnchorRange,
} from './comment-margin-layout.js';
import {
  isDocumentPullResponse,
  materializeDocumentPullAdapterSession,
} from './document-pull-client.js';

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
}

/** Options for {@link DocxDocument.collectPageRuns}. */
export type CollectPageRunsOptions = Pick<
  RenderPageOptions,
  'width' | 'currentDate' | 'showTrackedChanges'
>;

/** IX6 — options for {@link DocxDocument.renderPageToBitmap}: the serializable
 *  render knobs plus an OPTIONAL `onTextRun`. The callback stays main-thread (it
 *  never crosses the wire); in worker mode the proxy invokes it with the runs
 *  the worker shipped back beside the bitmap, so a caller gets the selection /
 *  find geometry on the same path in both modes. */
export type RenderPageToBitmapOptions = Omit<RenderPageOptions, 'onTextRun'> & {
  onTextRun?: (run: DocxTextRunInfo) => void;
};

export class DocxDocument {
  private _metrics: OoxmlResourceMetricsSession | null = null;
  private _document: DocxDocumentModel | null = null;
  private _source: LayoutSourceStore | null = null;
  private _meta: DocumentMeta | null = null;
  /** Lazily-built `bookmarkName → 0-based page index` map for internal hyperlink
   *  anchors (IX-nav). Built on first {@link getBookmarkPage} from the paginated
   *  pages (main) or the worker meta's `bookmarkPages` (worker). Nulled by
   *  {@link destroy} so a reused reference never serves a stale document. */
  private _bookmarkPages: Map<string, number> | null = null;
  /** Lazily-computed §17.13.4 comment anchor ranges (main mode; worker mode
   *  reads them from the meta). Nulled by {@link destroy} with the other
   *  per-document caches. */
  private _commentAnchorRanges: readonly CommentAnchorRange[] | null = null;
  private _mode: 'main' | 'worker' = 'main';
  private _threeD: ChartThreeDRenderer | undefined;
  private _regionMap: ChartRegionMapRenderer | undefined;
  private _chartEx: ChartExRenderer | undefined;
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
    let doc: DocxDocument | undefined;
    try {
      doc = new DocxDocument(worker, mode, defaultCurrentDateMs, opts.wasmUrl);
      doc._metrics = metrics;
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
        retained.layoutVariants.defaultLayout;
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
  ): Promise<void> {
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ type: 'parse', id, data: buffer, resourcePolicy, useGoogleFonts, defaultCurrentDateMs: documentLayoutRuntimeOf(this).defaultCurrentDateMs, renderers } satisfies RenderWorkerRequest)
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
  }

  destroy(): void {
    this._bridge.terminate();
    this._document = null;
    this._source = null;
    this._meta = null;
    documentLayoutRuntimeOf(this).services = null;
    this._bookmarkPages = null;
    this._commentAnchorRanges = null;
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
   * id / author / initials / date / plain-text body. Comments are a data-only
   * API: they are NOT drawn on the page (authoring applications may display
   * them in a margin pane / balloons, which this viewer does not reproduce).
   * Use this to build a review
   * panel, export an annotation list, etc. Returns `[]` when the document has no
   * comments part. The same data is also reachable via `document.comments`.
   */
  get comments(): DocComment[] {
    return this._meta?.comments ?? this._document?.comments ?? [];
  }

  /**
   * ECMA-376 §17.13.4 — the comment anchors resolved to per-paragraph run
   * intervals in document order (`commentRangeStart`/`End` pairs, plus
   * zero-length boundaries for reference-only comments). Join each range to
   * rendered geometry via `DocxTextRunInfo.source.path` + `sourceRunIndex`.
   * Mode-agnostic: main mode walks the body model lazily (cached per
   * document); worker mode reads the ranges the worker computed with the same
   * pure walker. Returns `[]` when the document has no comment anchors.
   */
  commentAnchorRanges(): readonly CommentAnchorRange[] {
    if (this._meta) return this._meta.commentAnchorRanges ?? [];
    if (!this._document) return [];
    this._commentAnchorRanges ??= collectDocumentCommentRanges(this._document.body);
    return this._commentAnchorRanges;
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
    const services = documentLayoutRuntimeOf(this).services;
    if (!services) throw new Error('Document layout services are not initialized');
    const store = layoutVariantStoreOf(services);
    if (!store) throw new Error('Document layout variant store is not initialized');
    return store.defaultLayout;
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
   * meta, built from the same paginated pages as `pageSizes`).
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
      ...opts,
      // Lazy image bytes: the renderer fetches each embedded blip on demand by
      // zip path (decoded only when drawn) instead of reading inlined base64.
      fetchImage: this._fetchImage,
      layoutServices: documentLayoutRuntimeOf(this).services ?? undefined,
      defaultCurrentDateMs: documentLayoutRuntimeOf(this).defaultCurrentDateMs,
      threeD: this._threeD,
      regionMap: this._regionMap,
      chartEx: this._chartEx,
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
    const wireOpts: WireRenderPageOptions = { ...wire, dpr: wire.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      // The selected date variant may have a different page count than default
      // metadata, so the worker validates against the layout it actually paints.
      const res = await this._bridge.request(
        (id) => ({ type: 'renderPage', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      const rendered = res as Extract<RenderWorkerResponse, { type: 'pageRendered' }>;
      if (onTextRun) for (const r of rendered.runs) onTextRun(r);
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
    const wireOpts: WireRenderPageOptions = { ...opts };
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
    if (this._mode === 'worker') {
      const res = await this._bridge.request(
        (id) => ({ type: 'hitTestElement', id, pageIndex, point, opts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'elementHit' }>).context;
    }
    const runtime = documentLayoutRuntimeOf(this);
    const services = runtime.services;
    if (!services) throw new Error('Document layout services are not initialized');
    return hitTestSelectedDocxElementContext(services, pageIndex, point, {
      ...opts,
      defaultCurrentDateMs: runtime.defaultCurrentDateMs,
    });
  }
}
