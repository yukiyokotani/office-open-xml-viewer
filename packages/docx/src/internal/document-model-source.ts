import { resolveCjkFallback, type CjkLang } from '@silurus/ooxml-core';
import {
  preloadGoogleFonts,
  loadOfficeFontFallbacks,
  unloadOfficeFontFallbacks,
  releaseOwnedBitmap,
  unloadGoogleFonts,
  unregisterEmbeddedFonts,
  WorkerBridge,
  defaultDpr,
  dropSvgImageCache,
  dropDecodedBitmapCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  type AdmittedModelSourceLoad,
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
import type { DocxDocumentModel, RenderPageOptions, WorkerRequest, WorkerResponse, DocComment, DocNote, DocRevision } from '../types';
import { renderLayoutSourceToCanvas, documentHasMath, prepareMathRuns, type DocxTextRunInfo } from '../renderer';
import { createLayoutServices } from '../layout-runtime.js';
import { buildBookmarkPageMap } from '../bookmark-nav';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames, docxOfficeFontFallbackRequests } from '../google-fonts';
import { loadEmbeddedFonts } from '../embedded-fonts';
import { loadBundledCalibri, unloadBundledOfficeFonts } from '../bundled-office-fonts.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
  layoutVariantStoreOf,
} from '../layout/runtime-state.js';
import {
  type LayoutSourceStore,
} from '../layout/layout-source-store.js';
import type { DeepReadonly, DocumentLayout } from '../layout/types.js';
import { snapshotPlainData } from '../layout/plain-data.js';
import type {
  DocumentLayoutPartial,
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
  WireRenderPageOptions,
} from '../worker-protocol';
import { retainRenderWorkerDocumentLayout } from '../render-worker-layout.js';
import { textRunsForSelectedPage } from '../text-run-projection.js';
import {
  reviewProjectionIndexForDocument,
  type ReviewProjectionIndex,
} from '../layout/text-index.js';
import {
  hitTestSelectedDocxElementContext,
  type DocxElementContextOptions,
} from '../element-context.js';
import type { DocxElementContext, DocxPagePoint } from '../selection-context.js';
import {
  collectLayoutSourceCommentRangesIfPresent,
  resolveDocxCommentThreads,
  type CommentAnchorRange,
  type ResolvedDocxCommentThread,
  type ResolveDocxCommentThreadsOptions,
} from '../comments.js';
import {
  collectLayoutSourceRevisionRangesIfPresent,
  type RevisionAnchorRange,
} from '../revisions.js';
import {
  isDocumentPullResponse,
  materializeDocumentPullAdapterSession,
} from '../document-pull-client.js';
import { layoutDocumentInputAsync } from '../layout/document.js';
import { layoutDocumentProgressively } from '../layout/progressive.js';
import { PaginationAbortError } from '../layout/pagination-scheduler.js';
import { normalizeLayoutOptions, type LayoutOptions } from '../layout/options.js';
import type { LayoutVariantStore } from '../layout/variant-store.js';
import { publishDocxLayout } from '../document-layout-events.js';
import {
  docxLayoutViewRequester,
  publishDocxLayoutView,
} from '../document-layout-view.js';


import { DocxDocument, type DocxViewerLoadControl, type LoadOptions } from '../document.js';
import { selectModelSource, beginModelSourceLoad } from '@silurus/ooxml-core/internal/model-source';
/** Parse-request fields for an application-selected model source. */
function modelSourceFields(
  load: AdmittedModelSourceLoad,
): { source: AdmittedModelSourceLoad['module']; sourceTransfer?: readonly Transferable[]; sourceOwnerUrl: string } {
  // The source owner is an opt-in ESM sidecar. An inline OOXML worker contains
  // only this URL and imports the sidecar after a source is selected.
  const sourceOwnerUrl = new URL(
    import.meta.env.DEV ? './worker-document-source.ts' : './docx-source-worker.mjs',
    import.meta.url,
  ).href;
  return load.transfer.length > 0
    ? { source: load.module, sourceTransfer: load.transfer, sourceOwnerUrl }
    : { source: load.module, sourceOwnerUrl };
}

interface Deferred<T> { promise: Promise<T>; resolve(value: T): void; reject(error: unknown): void; }
function deferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

type SourceDocxFriend = Pick<DocxDocument, keyof DocxDocument> & Record<
  '_metrics' | '_cjkFallback' | '_parse' | '_mode' | '_threeD' | '_regionMap' |
  '_chartEx' | '_tiff' | '_document' | '_embeddedFontFaces' | '_officeFontFaces' |
  '_bundledOfficeFontFaces' | '_bundledOfficeFontUrls' |
  '_googleFontFaces' | '_source' | '_layoutObservers' | '_layoutAbort' |
  '_replaceMainLayoutPublication' | '_isLayoutViewActive' | '_layoutLifecycle' |
  '_layoutCompletion' | '_resourceUsage' | '_progressive',
  any
>;

function adoptSourceView(doc: SourceDocxFriend, showTrackedChanges: boolean | undefined): void {
    if (showTrackedChanges === undefined) return;
    const runtime = documentLayoutRuntimeOf(doc);
    const active = runtime.activeLayoutOptions;
    if (!active || (active.showTrackedChanges === true) === showTrackedChanges) return;
    const next = normalizeLayoutOptions(active.currentDateMs, runtime.defaultCurrentDateMs, showTrackedChanges);
    runtime.activeLayoutOptions = next;
    if (doc._progressive) {
      (doc._progressive as { layoutOptions: LayoutOptions }).layoutOptions = next;
    }
}

export async function loadDocxModelSource(
  source: string | ArrayBuffer,
  opts: LoadOptions = {},
  control?: DocxViewerLoadControl,
): Promise<DocxDocument> {
    const signal = control?.signal;
    const checkAbort = () => { if (signal?.aborted) throw new PaginationAbortError(); };
    checkAbort();
    const cjkFallback = resolveCjkFallback(opts.cjkFallback);
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
      const res = await fetch(source, { signal });
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    checkAbort();
    // An application-supplied model source claims its input from the raw bytes
    // before OOXML container resolution; without `modelSources` nothing here
    // runs and the OOXML path below is unchanged.
    const selected = selectModelSource(opts.modelSources, 'docx', new Uint8Array(buffer));
    if (!selected) return DocxDocument.load(buffer, { ...opts, modelSources: undefined });
    const sourceLoad = beginModelSourceLoad(selected, 'docx');
    try {
    // Resolve the container on the main thread before spinning up the worker.
    // Container errors remain typed OoxmlError instances here; `instanceof`
    // would not survive the worker boundary.
    metrics.setSourceBytes(buffer.byteLength);
    metrics.checkpoint('container ready');
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker = mode === 'worker'
      ? (await import('../render-worker-source-host')).createRenderWorker()
      : new (await import('../worker-source.ts?worker&inline')).default();
    let doc: SourceDocxFriend | undefined;
    let publicDoc: DocxDocument | undefined;
    let disposed = false;
    const abortDocument = () => {
      if (disposed || !doc) return;
      disposed = true;
      doc.destroy();
    };
    const wiredWorker = sourceWorker(worker, sourceLoad, opts, (view) => { if (doc) adoptSourceView(doc, view); });
    const rendererDescriptors = mode === 'worker' ? workerRendererDescriptors(opts) : undefined;
    const workerProgressive = mode === 'worker' && !!opts.progressiveLayout;
    try {
      publicDoc = Reflect.construct(DocxDocument, [wiredWorker, mode, defaultCurrentDateMs, opts.wasmUrl]) as DocxDocument;
      // TypeScript's private members have no cross-module friend access;
      // the constructor above creates the real public class instance.
      doc = publicDoc as unknown as SourceDocxFriend;
      signal?.addEventListener('abort', abortDocument, { once: true });
      checkAbort();
      doc._metrics = metrics;
      doc._cjkFallback = cjkFallback;
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
      doc._bundledOfficeFontUrls = opts.useBundledOfficeFonts
        ? (await import('../assets/carlito/urls.js').catch(() => undefined))?.CARLITO_URLS
        : undefined;
      // In worker mode the worker preloads fonts before paginating (pagination
      // measures text), so the flag is forwarded; in main mode fonts are loaded
      // here after parse, before the lazy first pagination.
      await doc._parse(
        buffer,
        resourceOptions.policy,
        mode === 'worker' ? !!opts.useGoogleFonts : false,
        mode === 'worker' ? !!opts.useBundledOfficeFonts : false,
        opts.workerTimeoutMs,
        (usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => metrics.observeUsage(usage),
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
      // ECMA-376 §17.8.1 / §17.8.3 — register the document's embedded fonts (via
      // the worker's zip-entry extraction) before the lazy first pagination, so
      // text measures/draws with the authored typeface. Worker mode does this
      // inside the worker (before it paginates); here it runs on the main thread.
      let embeddedMetrics: Awaited<ReturnType<typeof loadEmbeddedFonts>>['metrics'] | undefined;
      let embeddedRoutes: Awaited<ReturnType<typeof loadEmbeddedFonts>>['routes'] | undefined;
      if (doc._mode === 'main' && doc._document?.embeddedFonts?.length) {
        const loadingDocument = doc;
        const loadedEmbedded = await loadEmbeddedFonts(
          doc._document,
          (p) => loadingDocument.getFontBytes(p),
        );
        // A canceled load may finish registering faces after destroy() has
        // already drained this document's earlier acquisitions.
        if (signal?.aborted) {
          unregisterEmbeddedFonts(loadedEmbedded.faces);
          throw new PaginationAbortError();
        }
        doc._embeddedFontFaces = loadedEmbedded.faces;
        embeddedMetrics = loadedEmbedded.metrics;
        embeddedRoutes = loadedEmbedded.routes;
      }
      const officeFonts = doc._mode === 'main' && doc._document
        ? await loadOfficeFontFallbacks(docxOfficeFontFallbackRequests(doc._document).filter((request) =>
            !embeddedRoutes?.some((route) => route.requestedFamily.toLowerCase() === request.family.toLowerCase()
              && route.weight === (request.weight ?? 400) && route.style === (request.style ?? 'normal'))))
        : { faces: [], routes: {} };
      if (signal?.aborted) {
        unloadOfficeFontFallbacks(officeFonts.faces);
        throw new PaginationAbortError();
      }
      doc._officeFontFaces = officeFonts.faces;
      const officeRequests = doc._mode === 'main' && doc._document
        ? docxOfficeFontFallbackRequests(doc._document) : [];
      const resolvedOfficeTuples = new Set([
        ...(embeddedRoutes ?? []).map((route) =>
          `${route.requestedFamily.toLowerCase()}:${route.weight}:${route.style}`),
        ...Object.values(officeFonts.routes).map((route) =>
          `${route.requestedFamily.toLowerCase()}:${route.weight}:${route.style}`),
      ]);
      const bundledFonts = doc._mode === 'main' && opts.useBundledOfficeFonts
        ? await loadBundledCalibri(officeRequests, resolvedOfficeTuples, doc._bundledOfficeFontUrls)
        : { faces: [], routes: [] };
      if (signal?.aborted) {
        unloadBundledOfficeFonts(bundledFonts.faces);
        throw new PaginationAbortError();
      }
      doc._bundledOfficeFontFaces = [...bundledFonts.faces];
      if (doc._mode === 'main' && opts.useGoogleFonts && doc._document) {
        // A proven local Calibri face already resolves this authored family;
        // avoid the optional Google Fonts substitution for the same request.
        const names = docxFontPreloadNames(doc._document, cjkFallback).filter((name) =>
          name?.toLowerCase() !== 'calibri' || (!('calibri' in officeFonts.routes)
            && bundledFonts.routes.length === 0));
        const googleFaces = await preloadGoogleFonts(names, DOCX_GOOGLE_FONTS);
        if (signal?.aborted) {
          unloadGoogleFonts(googleFaces);
          throw new PaginationAbortError();
        }
        doc._googleFontFaces = googleFaces;
      }
      // Equations are converted + rasterized before pagination (which reads their
      // extents synchronously). Requires the opt-in `math` engine; without it,
      // equations are skipped (and the engine asset is never bundled). Worker
      // mode performs the same preparation with the renderer's imported engine.
      let preparedMath;
      if (doc._mode === 'main' && opts.math && doc._document && documentHasMath(doc._document)) {
        preparedMath = await prepareMathRuns(doc._document, opts.math);
      }
      checkAbort();
      if (doc._mode === 'main' && doc._document && doc._source) {
        const layoutDocument = doc;
        const runtime = documentLayoutRuntimeOf(doc);
        runtime.services = createLayoutServices(doc._source, {
          fontMetrics: embeddedMetrics,
          useGoogleFonts: !!opts.useGoogleFonts,
          cjkFallback,
          embeddedRoutes,
          officeRoutes: [...Object.values(officeFonts.routes), ...bundledFonts.routes],
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
        // Sliced by default in main mode: the same pagination generator, drained across
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
            // The terminal success callback fires exactly once per load,
            // whether or not any partial was published — consumers must not
            // have to infer completion from document speed.
            progressiveDocument._layoutObservers.notify(
              'onLayoutComplete', opts.onLayoutComplete,
            );
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
        } else if (deferrable && (opts.sliceLayout !== false || opts.onLayoutProgress)) {
          for (;;) {
            checkAbort();
            const requestedView = control?.requestedView();
            const currentOptions = requestedView === undefined
              ? layoutOptions
              : normalizeLayoutOptions(opts.currentDate, runtime.defaultCurrentDateMs, requestedView);
            runtime.activeLayoutOptions = currentOptions;
            const abort = new AbortController();
            let viewChanged = false;
            const unsubscribe = control?.subscribeViewChange(() => {
              // Re-sending the in-flight view during progress is not a layout
              // change. A distinct request still cancels this slice once.
              const requested = control.requestedView();
              if (viewChanged || requested === undefined
                || (requested === true) === (currentOptions.showTrackedChanges === true)) return;
              viewChanged = true;
              abort.abort();
            });
            doc._layoutAbort = abort;
            try {
              const layout = await layoutDocumentInputAsync(
                doc._source.bodyLayoutInput,
                services,
                currentOptions,
                { ...scheduler, signal: abort.signal },
              );
              checkAbort();
              if (viewChanged) continue;
              retained.layoutVariants.prime(currentOptions, layout);
              break;
            } catch (error) {
              if (error instanceof PaginationAbortError && viewChanged && !signal?.aborted) continue;
              throw error;
            } finally {
              unsubscribe?.();
              doc._layoutAbort = null;
            }
          }
        } else {
          // Build the variant that will be rendered, not the default one.
          retained.layoutVariants.layoutFor(layoutOptions);
        }
      }
      // This final snapshot includes eager embedded-font extraction performed
      // after the parse response. Telemetry is strictly best-effort: a worker
      // failure or a silent worker may omit the newest counters, but must not
      // turn an otherwise successful load into a rejection or an endless wait.
      checkAbort();
      await doc._resourceUsage(
        opts.workerTimeoutMs ?? OOXML_RESOURCE_METRICS_PROBE_TIMEOUT_MS,
      ).then(
        (usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot | undefined) => metrics.observeUsage(usage),
        () => undefined,
      );
      metrics.checkpoint('model and layout ready');
      checkAbort();
      metrics.succeed({ pages: doc.pageCount });
      sourceLoad.release();
      return publicDoc;
    } catch (error) {
      disposeRejectedLoad(worker, doc ? abortDocument : undefined);
      throw error;
    } finally {
      signal?.removeEventListener('abort', abortDocument);
    }
    } finally {
      sourceLoad.release();
    }
    } catch (error) {
      metrics.fail(error);
      throw error;
    }
  }

/** Decorate only the selected-source worker; the ordinary bridge stays unchanged. */
function sourceWorker(
  worker: Worker,
  load: AdmittedModelSourceLoad,
  opts: LoadOptions,
  adoptView: (value: boolean | undefined) => void,
): Worker {
  const listeners = new Map<EventListenerOrEventListenerObject, EventListener>();
  return new Proxy(worker, {
    get(target, key) {
      if (key === 'postMessage') return (message: unknown, transfer?: Transferable[]) => {
        if (typeof message === 'object' && message !== null) {
          const wire = message as { type?: string };
          if (wire.type === 'init') return;
          if (wire.type === 'parse') {
            target.postMessage({ ...wire, ...modelSourceFields(load),
              ...(opts.showTrackedChanges === undefined ? {} : { showTrackedChanges: opts.showTrackedChanges }) },
            [...(transfer ?? []), ...load.transfer]);
            return;
          }
        }
        target.postMessage(message, transfer ?? []);
      };
      if (key === 'addEventListener') return (
        type: string, listener: EventListenerOrEventListenerObject, options?: AddEventListenerOptions,
      ) => {
        if (type !== 'message') return target.addEventListener(type, listener, options);
        const wrapped: EventListener = (event) => {
          const wire = (event as MessageEvent).data as {
            type?: string;
            viewDefaults?: { showTrackedChanges?: boolean };
            showTrackedChanges?: boolean;
          };
          if (opts.showTrackedChanges === undefined
            && (wire.type === 'documentSessionOpened' || wire.type === 'mainThreadVerticalFallback')) {
            adoptView(wire.viewDefaults?.showTrackedChanges);
          }
          if (wire.type === 'parsedMeta' || wire.type === 'layoutPartial') {
            adoptView(wire.showTrackedChanges);
          }
          if (typeof listener === 'function') listener(event);
          else listener.handleEvent(event);
        };
        listeners.set(listener, wrapped);
        target.addEventListener('message', wrapped, options);
      };
      if (key === 'removeEventListener') return (
        type: string, listener: EventListenerOrEventListenerObject, options?: EventListenerOptions,
      ) => target.removeEventListener(type, type === 'message' ? (listeners.get(listener) ?? listener) : listener, options);
      const value = Reflect.get(target, key, target);
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}
