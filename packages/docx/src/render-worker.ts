/**
 * Render-capable worker entry: parse → font preload → paginate, all
 * worker-side; renders pages into an OffscreenCanvas and replies with
 * transferable ImageBitmaps. Used by DocxDocument.load(src, { mode: 'worker' });
 * the slim parse-only worker.ts stays untouched so main-mode users pay no
 * bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 */
import init, { DocxArchive, reinit } from './wasm/docx_parser.js';
import {
  decodeDataUrl,
  preloadGoogleFonts,
  unloadGoogleFonts,
  unregisterEmbeddedFonts,
  unloadLocalFontMetrics,
  WasmParserHost,
  dropDecodedBitmapCache,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  PULL_SESSION_PROTOCOL,
  resourcePolicyForWasm,
  serializeWorkerError,
  loadWorkerRenderers,
  isWorkerSvgDecodeResponse,
  postOwnedImageBitmap,
  WorkerSvgDecodeClient,
  type LoadedWorkerRenderers,
  type PullSessionResponse,
  type WasmInitInput,
  type WorkerSvgDecodeResponse,
} from '@silurus/ooxml-core/worker';
import { prepareMathRuns, renderLayoutSourceToCanvas } from './renderer';
import { createLayoutServices } from './layout-runtime.js';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import { loadDocxLocalFontMetrics } from './local-font-metrics';
import type {
  RenderWorkerResponse,
  RenderWorkerWireRequest,
  DocumentLayoutPartial,
  DocumentMeta,
} from './worker-protocol';
import { layoutSourceModelAdapterFromOwnedModel } from './layout-source-model-adapter.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';
import {
  retainRenderWorkerDocumentLayout,
  type RetainedRenderWorkerDocumentLayout,
} from './render-worker-layout.js';
import { paginateRenderWorkerDocumentProgressively } from './render-worker-progressive.js';
import { PaginationAbortError } from './layout/pagination-scheduler.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import { hitTestSelectedDocxElementContext } from './element-context.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';
import {
  renderWorkerLayoutMeta,
  projectRenderWorkerLayoutMeta,
  type RenderWorkerReviewIndexInput,
} from './render-worker-metadata.js';
import { materializeDocumentPullOwnedModelsSession } from './document-pull-client.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  isDocumentPullCommand,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';
import { WorkerDocumentSourceOwner } from './internal/worker-document-source.js';

// RB6: self-poison + auto-respawn. A trap during parse (or an in-worker image /
// embedded-font read) recycles the instance so the next document renders on
// clean linear memory. The host owns the `DocxArchive` handle (`host.archive`).
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
const sourceOwner = new WorkerDocumentSourceOwner(host);
const documentPull = new DocumentPullWorker(
  () => sourceOwner.cursor(),
  (operation) => sourceOwner.execute(operation),
);
let documentGeneration = 0;
let parseGeneration = 0;
let ooxmlWasmInput: WasmInitInput | undefined;
let fallbackPull: DocumentPullWorker | null = null;
let doc: RetainedRenderWorkerDocumentLayout | null = null;
/** Compact model-derived inputs needed to re-project variant-specific review
 * anchor geometry. The complete parser/public model is not retained. */
let reviewIndexInput: RenderWorkerReviewIndexInput | null = null;
/** Cancels a still-running progressive drain when a new `parse` supersedes it.
 *  The host's `destroy()` terminates the worker outright, so this covers only
 *  the re-parse path — where the worker survives and would otherwise keep
 *  paginating a document nobody is holding any more. */
let layoutAbort: AbortController | null = null;
/** Floor between `layoutProgress` posts. `onProgress` fires at EVERY pagination
 *  suspension point; forwarding each one would flood the wire with messages the
 *  host can only render one of per frame. */
const LAYOUT_PROGRESS_POST_INTERVAL_MS = 100;
let renderers: LoadedWorkerRenderers = {};
let localMetricFontFaces: FontFace[] = [];
let googleFontFaces: FontFace[] = [];
let embeddedFontFaces: FontFace[] = [];
function releaseRetainedFonts(): void {
  const local = localMetricFontFaces;
  const google = googleFontFaces;
  const embedded = embeddedFontFaces;
  localMetricFontFaces = [];
  googleFontFaces = [];
  embeddedFontFaces = [];
  try { unloadLocalFontMetrics(local); } catch {}
  try { unloadGoogleFonts(google); } catch {}
  try { unregisterEmbeddedFonts(embedded); } catch {}
}
const rawParts = new BoundedRawPartCache({
  maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
  maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
});

const rawPost = (
  msg: unknown,
  transfer?: Transferable[],
) => (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);
const post = (
  msg: RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>,
  transfer?: Transferable[],
) => rawPost(msg, transfer);
const svgDecodeClient = new WorkerSvgDecodeClient(rawPost);

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  return rawParts.get(path, mimeType, async () => {
    const bytes = sourceOwner.execute((loaded) => loaded.extract_image(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  });
}

self.onmessage = async (e: MessageEvent<RenderWorkerWireRequest | WorkerSvgDecodeResponse>) => {
  const req = e.data;
  if (isWorkerSvgDecodeResponse(req)) {
    svgDecodeClient.accept(req);
    return;
  }
  if (isDocumentPullCommand(req)) {
    try {
      if (!fallbackPull) throw new Error('DOCX vertical fallback session is not open');
      await fallbackPull.dispatch(req, post);
    } catch (error) {
      post({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'error',
        sessionId: req.sessionId,
        operationId: req.operationId,
        generation: req.generation,
        requestId: req.requestId,
        error: serializeWorkerError(error),
      });
    }
    return;
  }
  if (req.type === 'init') {
    ooxmlWasmInput = decodeDataUrl(req.wasmUrl) ?? req.wasmUrl;
    return;
  }
  const id = req.id;
  let requestedGeneration: number | undefined;
  try {
    if (req.type !== 'parse' && sourceOwner.cursor()) {
      sourceOwner.execute((retained) => retained.assert_healthy());
    }
    if (req.type === 'parse') {
      requestedGeneration = ++parseGeneration;
      layoutAbort?.abort();
      layoutAbort = null;
      await documentPull.reset();
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      const previousFallback = fallbackPull;
      await previousFallback?.reset();
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      if (fallbackPull === previousFallback) fallbackPull = null;
      sourceOwner.closeNative();
      host.run(() => host.disposeArchive());
      doc = null;
      reviewIndexInput = null;
      releaseRetainedFonts();
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      rawParts.clear();
      const requestedRenderers = await loadWorkerRenderers(req.renderers);
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      renderers = requestedRenderers;
      // A re-parse starts a fresh document: also drop the shared decoded owner
      // (base raster + derived colour surfaces) and SVG lookup owner, symmetric
      // with DocxDocument.destroy(). The worker's `getImage`
      // closure is a stable module-level identity, so without this a new document
      // sharing a zip path (e.g. word/media/image1.png) would be served the
      // previous file's decoded surface. Symmetric across docx/pptx/xlsx render
      // workers (issue #781).
      dropDecodedBitmapCache(getImage);
      dropSvgImageCache(getImage);
      const bytes = new Uint8Array(req.data);
      // Both sources drain the same pull/ACK state machine locally. OOXML
      // construction/calls use host.run; the direct owner applies its own WASM
      // trap boundary. Neither route creates a monolithic model JSON value.
      if (req.source) {
        await sourceOwner.openNative(bytes, req.source);
        if (requestedGeneration !== parseGeneration) {
          throw new Error('render-worker parse was superseded');
        }
      } else {
        if (ooxmlWasmInput === undefined) throw new Error('DOCX WASM input was not configured');
        host.setWasmInput(ooxmlWasmInput);
        await host.ensureReady();
        if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
        const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
        host.run(() => {
          const archive = new DocxArchive(bytes, maxEntry, maxTotal, maxEntries);
          host.setArchive(archive);
        });
      }
      documentGeneration += 1;
      const identity = {
        sessionId: documentGeneration,
        operationId: documentGeneration,
        generation: documentGeneration,
      };
      documentPull.open(identity);
      let pulledModels: Awaited<ReturnType<typeof materializeDocumentPullOwnedModelsSession>>;
      let resourceUsage: OoxmlResourceUsageSnapshot | undefined;
      try {
        pulledModels = await materializeDocumentPullOwnedModelsSession(
          createLocalDocumentPullTransport(documentPull),
          identity,
          { onUsage: (usage) => { resourceUsage = usage; } },
        );
      } finally {
        if (requestedGeneration === parseGeneration) {
          await documentPull.reset().catch(() => undefined);
        }
      }
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      if (documentRequiresDomVerticalGlyphLayout(pulledModels.document)) {
        // The normalized public model deliberately omits parser-only sidecars
        // such as unavailable-drawing geometry. Stream the untouched parser
        // model to the main-thread normalization boundary one body block at a
        // time, without exposing those facts through `DocxDocument.document`.
        const fallbackArchive = new MaterializedDocumentCursorArchive(pulledModels.document);
        fallbackPull = new DocumentPullWorker(() => fallbackArchive);
        documentGeneration += 1;
        const fallbackIdentity = {
          sessionId: documentGeneration,
          operationId: documentGeneration,
          generation: documentGeneration,
        };
        fallbackPull.open(fallbackIdentity);
        post({ type: 'mainThreadVerticalFallback', id, ...fallbackIdentity, usage: resourceUsage });
        return;
      }
      const adapted = layoutSourceModelAdapterFromOwnedModel(
        pulledModels.document,
        pulledModels.ownedLayoutDocument,
      );
      const source = adapted.source;
      const model = adapted.document;
      const requestedReviewIndexInput = {
        comments: model.comments ?? [],
        revisions: model.revisions ?? [],
      };
      let googleFaces: FontFace[] = [];
      let embeddedFaces: FontFace[] = [];
      let localFaces: FontFace[] = [];
      let fontsTransferred = false;
      try {
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land before canonical layout —
        // same ordering the main-mode load() guarantees.
        googleFaces = await preloadGoogleFonts(
          docxFontPreloadNames(model),
          DOCX_GOOGLE_FONTS,
        );
        if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      }
      // ECMA-376 §17.8.1 / §17.8.3 — register embedded fonts into the worker's
      // FontFaceSet (self.fonts) before pagination measures text. Bytes are read
      // straight from the retained archive (extract_image reads any zip entry).
      if (model.embeddedFonts?.length) {
        embeddedFaces = await loadEmbeddedFonts(model, async (p) => {
          if (requestedGeneration !== parseGeneration) {
            throw new Error('render-worker parse was superseded');
          }
          return sourceOwner.execute((loaded) => loaded.extract_image(p));
        });
        if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      }
      const localMetrics = await loadDocxLocalFontMetrics(model);
      localFaces = localMetrics.faces;
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      let preparedMath: Awaited<ReturnType<typeof prepareMathRuns>> | undefined;
      preparedMath = requestedRenderers.math && source.mathOccurrences.length > 0
        ? await prepareMathRuns(model, requestedRenderers.math)
        : undefined;
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      const layoutServices = createLayoutServices(source, {
        localMetrics: localMetrics.metrics,
        useGoogleFonts: !!req.useGoogleFonts,
        embeddedFaces,
        googleFaces,
        mathResources: preparedMath?.records,
        mathDrawables: preparedMath?.drawables,
      });
      const requestedDoc = retainRenderWorkerDocumentLayout(
        source,
        layoutServices,
        req.defaultCurrentDateMs,
      );
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      localMetricFontFaces = localFaces;
      googleFontFaces = googleFaces;
      embeddedFontFaces = embeddedFaces;
      fontsTransferred = true;
      reviewIndexInput = requestedReviewIndexInput;
      doc = requestedDoc;
      } finally {
        if (!fontsTransferred) {
          try { unloadLocalFontMetrics(localFaces); } catch {}
          try { unloadGoogleFonts(googleFaces); } catch {}
          try { unregisterEmbeddedFonts(embeddedFaces); } catch {}
        }
      }
      // The variant this load will actually be viewed as. Everything below —
      // the progressive prefix, the authoritative layout, and the metadata the
      // host's geometry accessors read — is built for THIS view, so a
      // tracked-changes or explicit-date load no longer reports a page count
      // belonging to a pagination nobody is going to paint.
      const layoutOptions = normalizeLayoutOptions(
        req.currentDateMs,
        req.defaultCurrentDateMs,
        req.showTrackedChanges,
      );
      // Progressive layout: publish the opening pages long before the whole
      // document is paginated, so the host can resolve load() and paint while
      // the rest is still being laid out. Every publication primes the variant
      // store first, so a `renderPage` for a just-announced page is served from
      // the same store the authoritative layout will later replace.
      //
      // The guard mirrors main mode's `deferrable`: a fatally-unparseable
      // document is served a synthetic error page by the variant store's
      // builder, and neither previewing nor slicing may route around that.
      if (req.progressiveLayout && source.fatalParse === null) {
        // The parsed model is the source of review data, so the first
        // publication carries it: the host has nothing else to answer
        // `comments` / `revisions` from until `parsedMeta` arrives, and load()
        // is about to resolve on that first publication.
        let review: DocumentLayoutPartial['review'] | undefined = {
          revisions: model.revisions ?? [],
          comments: model.comments ?? [],
          footnotes: model.footnotes ?? [],
          endnotes: model.endnotes ?? [],
        };
        let lastProgressMs = 0;
        const abort = new AbortController();
        layoutAbort = abort;
        await paginateRenderWorkerDocumentProgressively(doc, source, {
          publish: (publication) => {
            if (requestedGeneration !== parseGeneration) return;
            post({
              type: 'layoutPartial',
              forId: id,
              partial: review ? { ...publication, review } : publication,
            });
            review = undefined;
          },
          progress: (committedPages) => {
            if (requestedGeneration !== parseGeneration) return;
            const now = Date.now();
            if (now - lastProgressMs < LAYOUT_PROGRESS_POST_INTERVAL_MS) return;
            lastProgressMs = now;
            post({ type: 'layoutProgress', forId: id, committedPages });
          },
        }, layoutOptions, abort.signal, reviewIndexInput);
        if (layoutAbort === abort) layoutAbort = null;
      }
      if (requestedGeneration !== parseGeneration) throw new Error('render-worker parse was superseded');
      // Usually a cache hit: the progressive drive above primed this exact
      // variant, so this reads the authoritative layout back rather than
      // paginating a second time. Without progressive layout it is the
      // blocking build, as it always was.
      const layout = doc.layoutVariants.layoutFor(layoutOptions);
      const meta: DocumentMeta = {
        revisions: model.revisions ?? [],
        comments: model.comments ?? [],
        footnotes: model.footnotes ?? [],
        endnotes: model.endnotes ?? [],
        ...projectRenderWorkerLayoutMeta(layout, source, reviewIndexInput),
      };
      if (sourceOwner.kind === 'ooxml') {
        const loadedArchive = sourceOwner.ooxml('resource usage');
        resourceUsage = decodeOoxmlResourceUsage(host.run(() => loadedArchive.resource_usage()));
      }
      post({ type: 'parsedMeta', id, meta, usage: resourceUsage });
      return;
    }
    if (req.type === 'selectLayoutView') {
      if (!doc || !reviewIndexInput) throw new Error('Document not loaded');
      post({
        type: 'layoutViewSelected',
        id,
        meta: renderWorkerLayoutMeta(
          doc,
          reviewIndexInput,
          req.currentDateMs,
          req.showTrackedChanges,
        ),
      });
      return;
    }
    if (req.type === 'renderPage') {
      if (!doc) throw new Error('Document not loaded');
      const canvas = new OffscreenCanvas(1, 1); // renderer resizes it
      const source = layoutSourceStoreOf(doc.layoutServices);
      if (!source) throw new Error('Document layout source is not initialized');
      await renderLayoutSourceToCanvas(source, canvas, req.pageIndex, {
        ...req.opts,
        fetchImage: getImage,
        svgDecoder: svgDecodeClient.decode,
        layoutServices: doc.layoutServices,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
        threeD: renderers.threeD,
        regionMap: renderers.regionMap,
        chartEx: renderers.chartEx,
        tiff: renderers.tiff,
      });
      const runs = textRunsForSelectedPage(doc.layoutServices, req.pageIndex, {
        ...req.opts,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
      });
      const bitmap = canvas.transferToImageBitmap();
      postOwnedImageBitmap(post, { type: 'pageRendered', id, bitmap, runs });
      return;
    }
    if (req.type === 'collectRuns') {
      if (!doc) throw new Error('Document not loaded');
      const runs = textRunsForSelectedPage(doc.layoutServices, req.pageIndex, {
        ...req.opts,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
      });
      post({ type: 'runsCollected', id, runs });
      return;
    }
    if (req.type === 'hitTestElement') {
      if (!doc) throw new Error('Document not loaded');
      const context = hitTestSelectedDocxElementContext(
        doc.layoutServices,
        req.pageIndex,
        req.point,
        { ...req.opts, defaultCurrentDateMs: doc.defaultCurrentDateMs },
      );
      post({ type: 'elementHit', id, context });
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      // wasm-bindgen returns an owned full-span Uint8Array; transfer its
      // standalone buffer directly, matching the parse worker contract.
      const bytes = sourceOwner.execute((archive) => archive.extract_image(req.path).buffer as ArrayBuffer);
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'resourceUsage') {
      const archive = sourceOwner.ooxml('resource usage');
      const usage = decodeOoxmlResourceUsage(host.run(() => archive.resource_usage()));
      post({ type: 'resourceUsage', id, usage });
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = sourceOwner.ooxml('markdown conversion');
      const markdown = host.run(() => archive.to_markdown());
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
  } catch (err) {
    // Non-parse pagination probes may still abort silently. A superseded parse
    // receives a terminal response for its old correlation id, but generation
    // checks above prevent it from publishing or closing the newer source.
    if (err instanceof PaginationAbortError && requestedGeneration === undefined) return;
    if (requestedGeneration !== undefined && requestedGeneration === parseGeneration) {
      // Cleanup must not replace the parse failure or suppress its terminal
      // response. A stale parse never touches the newer generation's owner.
      try { sourceOwner.closeNative(); } catch {}
      releaseRetainedFonts();
      doc = null;
      reviewIndexInput = null;
    }
    const error = err instanceof Error ? err : new Error(String(err));
    const details = error as Error & {
      code?: string;
      reason?: string;
      outgoingColumnIndex?: number;
      outgoingColumnCount?: number;
      incomingColumnCount?: number;
    };
    post({
      type: 'error',
      id,
      ...serializeWorkerError(error),
      ...(details.code !== undefined ? { code: details.code } : {}),
      ...(details.reason !== undefined ? { reason: details.reason } : {}),
      ...(details.outgoingColumnIndex !== undefined
        ? { outgoingColumnIndex: details.outgoingColumnIndex }
        : {}),
      ...(details.outgoingColumnCount !== undefined
        ? { outgoingColumnCount: details.outgoingColumnCount }
        : {}),
      ...(details.incomingColumnCount !== undefined
        ? { incomingColumnCount: details.incomingColumnCount }
        : {}),
    });
  }
};
