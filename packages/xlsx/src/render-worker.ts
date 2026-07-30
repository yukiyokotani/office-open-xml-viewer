/**
 * Render-capable worker entry: parse → font preload → (lazy) per-sheet parse →
 * render, all worker-side; renders a sheet viewport into an OffscreenCanvas and
 * replies with a transferable ImageBitmap. Used by
 * XlsxWorkbook.load(src, { mode: 'worker' }); the slim parse-only worker.ts
 * stays untouched so main-mode users pay no bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders. A
 * re-`parse` resets all per-document caches so a reused worker never serves
 * stale sheets / images.
 */
import init, { XlsxArchive, reinit } from './wasm/xlsx_parser.js';
import {
  decodeDataUrl,
  preloadGoogleFonts,
  WasmParserHost,
  dropBitmapCacheByPath,
  dropDuotoneBitmapCache,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import { resolveSharedStrings } from './shared-strings.js';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { applySizeOverrides } from './worker-protocol.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';

// RB6: self-poison + auto-respawn. A trap during parse / per-sheet parse / image
// read recycles the instance so the next workbook renders on clean linear
// memory. The host owns the `XlsxArchive` handle (`host.archive`): copies the
// file into WASM ONCE; the workbook / sharedStrings / theme parts are parsed
// ONCE and reused on every `parse_sheet`. Freed + replaced on a re-parse, freed +
// nulled by the host on a trap.
const host = new WasmParserHost<XlsxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let workbook: ParsedWorkbook | null = null;
/** Settled before any render when `useGoogleFonts` was requested. The resolved
 *  value (the preloaded FontFace[]) is unused here: the worker owns its own
 *  FontFaceSet (`self.fonts`) and terminates with it, so there is nothing to
 *  release — only the sequencing (fonts landed before first paint) matters. */
let fontsLoaded: Promise<unknown> = Promise.resolve();
const sheetCache = new Map<number, Worksheet>();
// Synchronous-lookup map for the draw pass, keyed by imageCacheKey(path, duotone).
// A pure lookup layer — the decoded bitmaps are owned by the shared, per-`getImage`
// core caches; this is refreshed each frame by prefetchImages and dropped (along
// with those shared caches) on re-parse.
const imageCache = new Map<string, CanvasImageSource | null>();
// Fetched image *bytes* (as Blobs) keyed by zip path. Twin of the docx render
// worker's `imageCache`; kept separate from the decoded-source `imageCache`
// above. Cleared on re-parse so a reused worker never serves a stale file's
// image.
const imageBlobCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** In-worker image-byte loader (twin of the docx render-worker `getImage`). The
 *  orchestrator's `fetchImage` routes here in worker mode, so image bytes are
 *  read straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageBlobCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('Workbook not loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageBlobCache.set(path, p);
  return p;
}

/** Lazily parse one worksheet directly via WASM and cache it — the worker-side
 *  equivalent of XlsxWorkbook.getWorksheet (same `parse_sheet` call, same
 *  decode cache). The handle's own shared-part cache means repeat sheet switches
 *  no longer re-parse the workbook / sharedStrings / theme. Shared by the
 *  `renderViewport` handler. */
function parseSheetLocally(sheetIndex: number): Worksheet {
  const cached = sheetCache.get(sheetIndex);
  if (cached) return cached;
  const loaded = host.archive;
  if (!workbook || !loaded) throw new Error('Workbook not loaded');
  const sheetMeta = workbook.workbook.sheets[sheetIndex];
  if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);
  // `parse_sheet` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); the
  // render worker consumes the model in-worker, so decode + parse here. Guarded
  // so a trap on ONE sheet recycles the instance instead of wedging the workbook.
  const json = host.run(() => loaded.parse_sheet(sheetIndex, sheetMeta.name));
  const ws = JSON.parse(new TextDecoder().decode(json)) as Worksheet;
  // Resolve `{ type: 'shared', si }` cells against the dedup'd sharedStrings
  // table so the renderer only ever sees fully-materialized text (worker-mode
  // twin of workbook.ts getWorksheet).
  resolveSharedStrings(ws, workbook.sharedStrings);
  sheetCache.set(sheetIndex, ws);
  return ws;
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
      // A re-parse starts a fresh document: drop any cached sheets / images so
      // we never serve stale data from a previous load. `imageCache` is now a
      // pure lookup map into the shared, per-`getImage` core caches (base raster,
      // duotone recolour, SVG); clearing it drops lookup references, and dropping
      // the three shared caches keyed by the module-level `getImage` closure
      // releases the GPU-backed ImageBitmaps and SVG object URLs AND prevents the
      // next document from being served a stale bitmap for an identically-named
      // zip path. Symmetric with XlsxWorkbook.destroy() and the docx/pptx render
      // workers (issue #781).
      sheetCache.clear();
      imageCache.clear();
      dropBitmapCacheByPath(getImage);
      dropDuotoneBitmapCache(getImage);
      dropSvgImageCache(getImage);
      imageBlobCache.clear();
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
      // Construction + `parse()` run under `host.run` so a trap in EITHER poisons
      // + recycles the instance (and frees the archive). `setArchive` frees any
      // prior handle first — the re-parse dispose. `parse()` returns UTF-8 JSON
      // bytes (Result<Vec<u8>, JsValue>); decode + parse the workbook index here
      // (consumed in-worker, then a light copy is sent to the proxy as an object).
      workbook = host.run(() => {
        const archive = new XlsxArchive(new Uint8Array(req.data), max, maxTotal, maxEntries);
        host.setArchive(archive);
        return JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
      });
      if (req.useGoogleFonts) {
        // Mirror XlsxWorkbook._load exactly: queue Google Fonts substitutes for
        // every styled font name, plus the generic Arabic fallbacks. Fonts must
        // land before rendering (which measures text), so we keep the promise
        // and await it in the renderViewport handler.
        fontsLoaded = preloadGoogleFonts(xlsxFontPreloadNames(workbook), XLSX_GOOGLE_FONTS);
      }
      post({ type: 'parsed', id, workbook });
      return;
    }
    if (req.type === 'parseSheet') {
      // Worker-side equivalent of the slim worker.ts `parseSheet` arm, so
      // XlsxWorkbook.getWorksheet / resolveValidationList resolve in worker
      // mode instead of hanging forever. Reuses parseSheetLocally (which parses
      // from the worker's stored rawData and cache); the main-mode message's
      // `sheetName` field is ignored. Reply shape matches worker.ts.
      if (!workbook) throw new Error('Workbook not loaded');
      const worksheet = parseSheetLocally(req.sheetIndex);
      post({ type: 'parsedSheet', id, worksheet });
      return;
    }
    if (req.type === 'renderViewport') {
      if (!workbook) throw new Error('Workbook not loaded');
      await fontsLoaded;
      const ws = parseSheetLocally(req.sheetIndex);
      // Converge this worker-local sheet to the main-thread model before
      // drawing: view-only size mutations (outline collapse/expand, drag
      // resize #567) happen on the main thread's Worksheet copy and would
      // otherwise never reach this cache — the gutter/overlays would update
      // while the grid bitmap stayed stale. The override map is cumulative and
      // idempotent, so re-applying it every frame (and after a worker-side
      // re-parse rebuilt the cache from the file) always converges.
      const { sizeOverrides, ...renderOpts } = req.opts;
      applySizeOverrides(ws, sizeOverrides);
      const canvas = new OffscreenCanvas(1, 1); // orchestrator resizes it
      await renderWorksheetViewport(
        { ws, styles: workbook.styles, imageCache },
        canvas,
        req.viewport,
        // Supply the in-worker byte loader so embedded images decode straight
        // from the retained rawData (no main-thread round-trip).
        { ...renderOpts, fetchImage: getImage },
      );
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'viewportRendered', id, bitmap }, [bitmap]);
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      const archive = host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      const raw = host.run(() => archive.extract_image(req.path));
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      const markdown = host.run(() => archive.to_markdown());
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
