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
  dropBitmapCacheByPath,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types';
import { renderDocumentToCanvas, dropColorReplacedCache } from './renderer';
import { createLayoutServices } from './layout-runtime.js';
import { buildBookmarkPageMap } from './bookmark-nav';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import { loadDocxLocalFontMetrics } from './local-font-metrics';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';
import { normalizeInternalDocumentModel } from './parser-model.js';
import {
  retainRenderWorkerDocumentLayout,
  type RetainedRenderWorkerDocumentLayout,
} from './render-worker-layout.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';

// RB6: self-poison + auto-respawn. A trap during parse (or an in-worker image /
// embedded-font read) recycles the instance so the next document renders on
// clean linear memory. The host owns the `DocxArchive` handle (`host.archive`).
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let doc: RetainedRenderWorkerDocumentLayout | null = null;
let localMetricFontFaces: FontFace[] = [];
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('No docx loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;
  if (req.type === 'init') {
    host.setWasmUrl(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.type === 'parse') {
      doc = null;
      if (localMetricFontFaces.length > 0) {
        unloadLocalFontMetrics(localMetricFontFaces);
        localMetricFontFaces = [];
      }
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      imageCache.clear();
      // A re-parse starts a fresh document: also drop the shared, per-`getImage`
      // decoded caches (base raster, a:clrChange/duotone recolour, SVG object
      // URLs), symmetric with DocxDocument.destroy(). The worker's `getImage`
      // closure is a stable module-level identity, so without this a new document
      // sharing a zip path (e.g. word/media/image1.png) would be served the
      // previous file's decoded bitmap, and the GPU/URL handles would linger past
      // the LRU cap. Symmetric across docx/pptx/xlsx render workers (issue #781).
      dropBitmapCacheByPath(getImage);
      dropColorReplacedCache(getImage);
      dropSvgImageCache(getImage);
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const maxTotal =
        typeof req.maxZipTotalBytes === 'number' && req.maxZipTotalBytes > 0
          ? BigInt(req.maxZipTotalBytes)
          : undefined;
      const maxEntries =
        typeof req.maxZipEntries === 'number' && req.maxZipEntries > 0
          ? BigInt(req.maxZipEntries)
          : undefined;
      const bytes = new Uint8Array(req.data);
      // Construction + `parse()` run under `host.run` so a trap in EITHER poisons
      // + recycles the instance (and frees the archive). `setArchive` frees any
      // prior handle first — the re-parse dispose. `parse()` throws on
      // parse/serialize failure (Result<Vec<u8>, JsValue>); the outer try/catch
      // converts a graceful failure into an error response. Render mode consumes
      // the model in-worker, so decode + parse it here (one decode, no
      // passthrough).
      const parsedModel = host.run(() => {
        const archive = new DocxArchive(bytes, max, maxTotal, maxEntries);
        host.setArchive(archive);
        return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel;
      });
      const model = normalizeInternalDocumentModel(parsedModel).document;
      if (documentRequiresDomVerticalGlyphLayout(model)) {
        // The normalized public model deliberately omits parser-only sidecars
        // such as unavailable-drawing geometry. Send the untouched parser wire
        // to the main-thread fallback so its normalization boundary can rebuild
        // those identity-owned acquisition facts without exposing them through
        // `DocxDocument.document`.
        const encoded = new TextEncoder().encode(JSON.stringify(parsedModel));
        const documentJson = encoded.buffer.slice(
          encoded.byteOffset,
          encoded.byteOffset + encoded.byteLength,
        ) as ArrayBuffer;
        post({ type: 'mainThreadVerticalFallback', id, documentJson }, [documentJson]);
        return;
      }
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
          return new Uint8Array(host.run(() => loaded.extract_image(p))).slice();
        });
      }
      const localMetrics = await loadDocxLocalFontMetrics(model);
      localMetricFontFaces = localMetrics.faces;
      const layoutServices = createLayoutServices(model, {
        localMetrics: localMetrics.metrics,
        useGoogleFonts: !!req.useGoogleFonts,
        embeddedFaces,
        googleFaces,
      });
      doc = retainRenderWorkerDocumentLayout(
        model,
        layoutServices,
        req.defaultCurrentDateMs,
      );
      const layout = doc.layoutVariants.defaultLayout;
      const pageSizes = layout.pages.map((page) => ({
        widthPt: page.geometry.widthPt,
        heightPt: page.geometry.heightPt,
      }));
      const meta: DocumentMeta = {
        pageCount: layout.pages.length,
        comments: model.comments ?? [],
        footnotes: model.footnotes ?? [],
        endnotes: model.endnotes ?? [],
        pageSizes,
        bookmarkPages: [...buildBookmarkPageMap(layout)],
      };
      post({ type: 'parsedMeta', id, meta });
      return;
    }
    if (req.type === 'renderPage') {
      if (!doc) throw new Error('Document not loaded');
      const canvas = new OffscreenCanvas(1, 1); // renderer resizes it
      await renderDocumentToCanvas(doc.model, canvas, req.pageIndex, {
        ...req.opts,
        fetchImage: getImage,
        layoutServices: doc.layoutServices,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
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
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      const raw = host.run(() => archive.extract_image(req.path));
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
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
      message: error.message,
      errorName: error.name,
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
