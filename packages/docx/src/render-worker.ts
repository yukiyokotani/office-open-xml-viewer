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
  type LoadedWorkerRenderers,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { prepareMathRuns, renderLayoutSourceToCanvas } from './renderer';
import { createLayoutServices } from './layout-runtime.js';
import { buildBookmarkPageMap } from './bookmark-nav';
import { collectLayoutSourceCommentRangesIfPresent } from './comments.js';
import { collectLayoutSourceRevisionRangesIfPresent } from './revisions.js';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import { loadDocxLocalFontMetrics } from './local-font-metrics';
import type {
  RenderWorkerResponse,
  RenderWorkerWireRequest,
  DocumentMeta,
} from './worker-protocol';
import { layoutSourceModelAdapterFromOwnedModel } from './layout-source-model-adapter.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';
import {
  retainRenderWorkerDocumentLayout,
  type RetainedRenderWorkerDocumentLayout,
} from './render-worker-layout.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import { textRunSourceIndexForDocument } from './layout/text-index.js';
import { hitTestSelectedDocxElementContext } from './element-context.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';
import { materializeDocumentPullOwnedModelsSession } from './document-pull-client.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  isDocumentPullCommand,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';

// RB6: self-poison + auto-respawn. A trap during parse (or an in-worker image /
// embedded-font read) recycles the instance so the next document renders on
// clean linear memory. The host owns the `DocxArchive` handle (`host.archive`).
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
const documentPull = new DocumentPullWorker(
  () => host.archive,
  (operation) => host.run(() => {
    const archive = host.archive;
    if (!archive) throw new Error('No docx loaded');
    return operation(archive);
  }),
);
let documentGeneration = 0;
let fallbackPull: DocumentPullWorker | null = null;
let doc: RetainedRenderWorkerDocumentLayout | null = null;
let renderers: LoadedWorkerRenderers = {};
let localMetricFontFaces: FontFace[] = [];
const rawParts = new BoundedRawPartCache({
  maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
  maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
});

const post = (
  msg: RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>,
  transfer?: Transferable[],
) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  return rawParts.get(path, mimeType, async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('No docx loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  });
}

self.onmessage = async (e: MessageEvent<RenderWorkerWireRequest>) => {
  const req = e.data;
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
    host.setWasmInput(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.type !== 'parse' && host.archive) {
      const retained = host.archive;
      host.run(() => retained.assert_healthy());
    }
    if (req.type === 'parse') {
      await documentPull.reset();
      await fallbackPull?.reset();
      fallbackPull = null;
      doc = null;
      if (localMetricFontFaces.length > 0) {
        unloadLocalFontMetrics(localMetricFontFaces);
        localMetricFontFaces = [];
      }
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      rawParts.clear();
      renderers = await loadWorkerRenderers(req.renderers);
      // A re-parse starts a fresh document: also drop the shared decoded owner
      // (base raster + derived colour surfaces) and SVG lookup owner, symmetric
      // with DocxDocument.destroy(). The worker's `getImage`
      // closure is a stable module-level identity, so without this a new document
      // sharing a zip path (e.g. word/media/image1.png) would be served the
      // previous file's decoded surface. Symmetric across docx/pptx/xlsx render
      // workers (issue #781).
      dropDecodedBitmapCache(getImage);
      dropSvgImageCache(getImage);
      const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
      const bytes = new Uint8Array(req.data);
      // Construction and every later cursor call run under `host.run`. Render
      // mode drains the same pull/ACK state machine locally, so it avoids both a
      // monolithic Rust model JSON value and an unnecessary Worker transfer.
      host.run(() => {
        const archive = new DocxArchive(bytes, maxEntry, maxTotal, maxEntries);
        host.setArchive(archive);
      });
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
        await documentPull.reset().catch(() => undefined);
      }
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
      let googleFaces: FontFace[] = [];
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land before canonical layout —
        // same ordering the main-mode load() guarantees.
        googleFaces = await preloadGoogleFonts(
          docxFontPreloadNames(model),
          DOCX_GOOGLE_FONTS,
        );
      }
      // ECMA-376 §17.8.1 / §17.8.3 — register embedded fonts into the worker's
      // FontFaceSet (self.fonts) before pagination measures text. Bytes are read
      // straight from the retained archive (extract_image reads any zip entry).
      let embeddedFaces: FontFace[] = [];
      if (model.embeddedFonts?.length) {
        embeddedFaces = await loadEmbeddedFonts(model, async (p) => {
          const loaded = host.archive;
          if (!loaded) throw new Error('No docx loaded');
          return host.run(() => loaded.extract_image(p));
        });
      }
      const localMetrics = await loadDocxLocalFontMetrics(model);
      localMetricFontFaces = localMetrics.faces;
      const preparedMath = renderers.math && source.mathOccurrences.length > 0
        ? await prepareMathRuns(model, renderers.math)
        : undefined;
      const layoutServices = createLayoutServices(source, {
        localMetrics: localMetrics.metrics,
        useGoogleFonts: !!req.useGoogleFonts,
        embeddedFaces,
        googleFaces,
        mathResources: preparedMath?.records,
        mathDrawables: preparedMath?.drawables,
      });
      doc = retainRenderWorkerDocumentLayout(
        source,
        layoutServices,
        req.defaultCurrentDateMs,
      );
      const layout = doc.layoutVariants.defaultLayout;
      const pageSizes = layout.pages.map((page) => ({
        widthPt: page.geometry.widthPt,
        heightPt: page.geometry.heightPt,
      }));
      const renderedRunIndex = textRunSourceIndexForDocument(layout);
      const meta: DocumentMeta = {
        pageCount: layout.pages.length,
        revisions: model.revisions ?? [],
        comments: model.comments ?? [],
        footnotes: model.footnotes ?? [],
        endnotes: model.endnotes ?? [],
        pageSizes,
        bookmarkPages: [...buildBookmarkPageMap(layout)],
        // §17.13.4 — story-aware anchor ranges use the exact retained source
        // identities shipped by projected text runs in both modes.
        commentAnchorRanges: collectLayoutSourceCommentRangesIfPresent(
          model.comments,
          source,
          renderedRunIndex,
        ),
        revisionAnchorRanges: collectLayoutSourceRevisionRangesIfPresent(
          model.revisions,
          source,
          renderedRunIndex,
        ),
      };
      const loadedArchive = host.archive;
      if (!loadedArchive) throw new Error('No docx loaded');
      resourceUsage = decodeOoxmlResourceUsage(
        host.run(() => loadedArchive.resource_usage()),
      );
      post({ type: 'parsedMeta', id, meta, usage: resourceUsage });
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
        layoutServices: doc.layoutServices,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
        threeD: renderers.threeD,
        regionMap: renderers.regionMap,
        chartEx: renderers.chartEx,
      });
      const runs = textRunsForSelectedPage(doc.layoutServices, req.pageIndex, {
        ...req.opts,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
      });
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'pageRendered', id, bitmap, runs }, [bitmap]);
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
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      // wasm-bindgen returns an owned full-span Uint8Array; transfer its
      // standalone buffer directly, matching the parse worker contract.
      const bytes = host.run(() => archive.extract_image(req.path).buffer as ArrayBuffer);
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'resourceUsage') {
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      const usage = decodeOoxmlResourceUsage(host.run(() => archive.resource_usage()));
      post({ type: 'resourceUsage', id, usage });
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      const markdown = host.run(() => archive.to_markdown());
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
  } catch (err) {
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
