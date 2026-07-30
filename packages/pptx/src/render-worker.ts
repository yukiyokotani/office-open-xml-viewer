/**
 * Render-capable worker entry: a superset of `worker.ts` that keeps the parsed
 * Presentation worker-side and renders slides into an OffscreenCanvas,
 * replying with transferable ImageBitmaps. Used by
 * `PptxPresentation.load(src, { mode: 'worker' })`; the slim parse-only
 * `worker.ts` stays untouched so main-mode users pay no bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 * A re-parse while a render is suspended mid-await would let that render
 * resume against the new model — callers must not interleave them.
 */
import type { MediaElement, Presentation } from './types';
import init, { PptxArchive, reinit } from './wasm/pptx_parser.js';
import { renderSlide, type PptxTextRunInfo } from './renderer';
import { selectNotes } from './notes';
import { findMimeTypeForPath } from './media-mime';
import { PPTX_GOOGLE_FONTS, pptxFontPreloadNames } from './google-fonts';
import {
  preloadGoogleFonts,
  decodeDataUrl,
  WasmParserHost,
  dropBitmapCacheByPath,
  dropDuotoneBitmapCache,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import type { RenderWorkerRequest, RenderWorkerResponse, PresentationMeta } from './worker-protocol';

// RB6: same self-poison + auto-respawn as the parse-only worker. A trap during
// parse (or an in-worker media/image read) recycles the instance so the next
// document renders on clean linear memory instead of a corrupted heap. The host
// owns the `PptxArchive` handle (`host.archive`): copies the file into WASM ONCE
// and opens it ONCE; the in-worker `getMedia` / `getImage` closures then read
// bytes by zip path straight from the retained archive. Freed + replaced on a
// re-parse, freed + nulled by the host on a trap.
const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let pres: Presentation | null = null;
/** Settled before any render when `useGoogleFonts` was requested. The resolved
 *  value (the preloaded FontFace[]) is unused here: the worker owns its own
 *  FontFaceSet (`self.fonts`) and terminates with it, so there is nothing to
 *  release — only the sequencing (fonts landed before first paint) matters. */
let fontsLoaded: Promise<unknown> = Promise.resolve();
const mediaCache = new Map<string, Promise<Blob>>();
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

function getMedia(path: string): Promise<Blob> {
  const hit = mediaCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    const loaded = host.archive;
    if (!loaded || !pres) throw new Error('No pptx loaded');
    const bytes = host.run(() => loaded.extract_media(path));
    return new Blob([new Uint8Array(bytes).slice()], { type: findMimeTypeForPath(pres, path) });
  })();
  mediaCache.set(path, p);
  return p;
}

/** In-worker image-byte loader (twin of {@link getMedia}). The renderer's
 *  `fetchImage` routes here in worker mode, so image bytes are decoded straight
 *  from the retained archive with no main-thread round-trip. Mime travels on the
 *  element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('No pptx loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    // Retain the init lifecycle in the host (docx/xlsx pattern) rather than a
    // `ready` flag + handshake. Every request below `await`s `ensureReady()`, so
    // a REJECTED init rejects the request (the outer catch posts an `error`
    // response the bridge turns into a rejected `load()` / render), never a
    // silent hang on a main-side wait. After a trap, `ensureReady()` respawns.
    host.setWasmUrl(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  try {
    await host.ensureReady();
    if (req.kind === 'parse') {
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's media/image.
      mediaCache.clear();
      imageCache.clear();
      // A re-parse starts a fresh document: also drop the shared, per-`getImage`
      // decoded caches (base raster, a:duotone recolour, SVG object URLs),
      // symmetric with PptxPresentation.destroy(). The worker's `getImage` closure
      // is a stable module-level identity, so without this a new deck sharing a
      // zip path (e.g. ppt/media/image1.png) would be served the previous file's
      // decoded bitmap, and the GPU/URL handles would linger past the LRU cap.
      // Symmetric across docx/pptx/xlsx render workers (issue #781).
      dropBitmapCacheByPath(getImage);
      dropDuotoneBitmapCache(getImage);
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
      const bytes = new Uint8Array(req.buffer);
      // Construction + `parse()` run under `host.run` so a trap in EITHER poisons
      // + recycles the instance (and frees the archive). `setArchive` frees any
      // prior handle first — the re-parse dispose. `parse()` returns UTF-8 JSON
      // bytes (Result<Vec<u8>, JsValue>). Render mode consumes the model
      // in-worker, so decode + parse here (one decode, no passthrough).
      pres = host.run(() => {
        const archive = new PptxArchive(bytes, max, maxTotal, maxEntries);
        host.setArchive(archive);
        return JSON.parse(new TextDecoder().decode(archive.parse())) as Presentation;
      });
      if (req.useGoogleFonts) {
        // Kick the preload now so it overlaps with main-thread work; renders
        // await `fontsLoaded` so text never rasterizes with fallback metrics.
        fontsLoaded = preloadGoogleFonts(
          pptxFontPreloadNames(pres),
          PPTX_GOOGLE_FONTS,
        );
      }
      const meta: PresentationMeta = {
        slideCount: pres.slides.length,
        slideWidth: pres.slideWidth,
        slideHeight: pres.slideHeight,
        majorFont: pres.majorFont ?? null,
        minorFont: pres.minorFont ?? null,
        notes: pres.slides.map((_, i) => selectNotes(pres!.slides, i)),
        mediaElements: pres.slides.map((s) =>
          s.elements.filter((el): el is MediaElement => el.type === 'media')),
        hidden: pres.slides.map((s) => s.hidden ?? false),
        partNames: pres.slides.map((s) => s.partName),
      };
      post({ kind: 'parsedMeta', id: req.id, meta });
      return;
    }

    if (req.kind === 'renderSlide') {
      if (!pres) throw new Error('No pptx loaded');
      const slide = pres.slides[req.slideIndex];
      if (!slide) throw new Error(`Slide index ${req.slideIndex} out of range (count: ${pres.slides.length})`);
      await fontsLoaded;
      const canvas = new OffscreenCanvas(1, 1); // renderSlide resizes it
      // IX6 — collect the run geometry the same render emits so the main thread
      // can build its selection / find overlay without a second render. The
      // callback runs worker-side; only the resulting plain array crosses back.
      const runs: PptxTextRunInfo[] = [];
      await renderSlide(canvas, slide, pres.slideWidth, pres.slideHeight, {
        width: req.width,
        dpr: req.dpr,
        defaultTextColor: pres.defaultTextColor,
        majorFont: pres.majorFont,
        minorFont: pres.minorFont,
        hlinkColor: pres.hlinkColor ?? null,
        fetchMedia: getMedia,
        fetchImage: getImage,
        skipMediaControls: req.skipMediaControls,
        dim: req.dim,
        // math intentionally omitted: MathJax needs a DOM <script>; worker
        // mode skips equations (documented in the design spec).
      }, (r) => runs.push(r));
      const bitmap = canvas.transferToImageBitmap();
      post({ kind: 'slideRendered', id: req.id, bitmap, runs }, [bitmap]);
      return;
    }
    if (req.kind === 'collectRuns') {
      // IX6 — render a slide purely to harvest its text-run geometry (find scans
      // every slide). The bitmap is discarded worker-side; only `runs` crosses
      // the wire. Same renderer / width as the main-mode `_collectSlideRuns`, so
      // the geometry is identical to what a `renderSlide` of the same slide draws
      // (no dpr / dim needed: run geometry is in CSS px, independent of dpr, and
      // dimming does not move glyphs — matching main-mode `_collectSlideRuns`).
      if (!pres) throw new Error('No pptx loaded');
      const slide = pres.slides[req.slideIndex];
      if (!slide) throw new Error(`Slide index ${req.slideIndex} out of range (count: ${pres.slides.length})`);
      await fontsLoaded;
      const canvas = new OffscreenCanvas(1, 1);
      const runs: PptxTextRunInfo[] = [];
      await renderSlide(canvas, slide, pres.slideWidth, pres.slideHeight, {
        width: req.width,
        defaultTextColor: pres.defaultTextColor,
        majorFont: pres.majorFont,
        minorFont: pres.minorFont,
        hlinkColor: pres.hlinkColor ?? null,
        fetchMedia: getMedia,
        fetchImage: getImage,
      }, (r) => runs.push(r));
      post({ kind: 'runsCollected', id: req.id, runs });
      return;
    }

    if (req.kind === 'extractMedia') {
      const blob = await getMedia(req.path);
      const bytes = await blob.arrayBuffer();
      post({ kind: 'mediaExtracted', id: req.id, bytes }, [bytes]);
      return;
    }

    if (req.kind === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this message arm exists only for protocol parity with worker.ts. Raw
      // bytes are read straight from the retained archive (no mime needed for a
      // byte transfer).
      const archive = host.archive;
      if (!archive) throw new Error('No pptx loaded');
      const raw = host.run(() => archive.extract_image(req.path));
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ kind: 'imageExtracted', id: req.id, bytes }, [bytes]);
      return;
    }

    if (req.kind === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = host.archive;
      if (!archive) throw new Error('No pptx loaded');
      const markdown = host.run(() => archive.to_markdown());
      post({ kind: 'markdownRendered', id: req.id, markdown });
      return;
    }
  } catch (err) {
    if ('id' in req) {
      post({ kind: 'error', id: req.id, message: err instanceof Error ? err.message : String(err) });
    }
  }
};
