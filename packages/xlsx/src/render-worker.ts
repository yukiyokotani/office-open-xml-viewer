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
  dropDecodedBitmapCache,
  dropSvgImageCache,
  type FontFamilyRoutes,
} from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import {
  decodeOoxmlResourceUsage,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  resourcePolicyForWasm,
  serializeWorkerError,
  loadWorkerRenderers,
  isWorkerSvgDecodeResponse,
  postOwnedImageBitmap,
  WorkerSvgDecodeClient,
  FontProviderClient,
  isWorkerFontResponse,
  type LoadedWorkerRenderers,
  type PullSessionCommand,
  type PullSessionResponse,
  type WorkerSvgDecodeResponse,
} from '@silurus/ooxml-core/worker';
import { workerRenderDeps } from './worker-render-deps.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import { resolveSharedStringRows } from './shared-strings.js';
import {
  addWorksheetCacheUsage,
  assertWorksheetCacheUsage,
  type WorksheetCacheUsage,
} from './worksheet-resource-limits.js';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { WorksheetViewProjectionCache } from './worker-protocol.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { readXlsxArchiveBootstrap } from './internal/archive-bootstrap.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';
import { isWorksheetPullCommand, WorksheetPullWorker } from './worksheet-pull-worker.js';
import {
  applyWorkbookFontRoutes,
  applyWorksheetFontRoutes,
  xlsxFontProviderNames,
  xlsxWorksheetFontProviderNames,
} from './provider-fonts.js';

// RB6: self-poison + auto-respawn. A trap during parse / per-sheet parse / image
// read recycles the instance so the next workbook renders on clean linear
// memory. The host owns the `XlsxArchive` handle (`host.archive`): copies the
// file into WASM ONCE; the workbook / sharedStrings / theme parts are parsed
// ONCE and reused by every worksheet cursor. Freed + replaced on a new workbook,
// freed + nulled by the host on a trap.
const host = new WasmParserHost<XlsxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let workbook: ParsedWorkbook | null = null;
let renderers: LoadedWorkerRenderers = {};
/** Settled before any render when `useGoogleFonts` was requested. The resolved
 *  value (the preloaded FontFace[]) is unused here: the worker owns its own
 *  FontFaceSet (`self.fonts`) and terminates with it, so there is nothing to
 *  release — only the sequencing (fonts landed before first paint) matters. */
let fontsLoaded: Promise<unknown> = Promise.resolve();
let providerFontRoutes: FontFamilyRoutes = {};
let providerEnabled = false;
let generation = 0;
const sheetCache = new Map<number, Worksheet>();
const viewProjectionCache = new WorksheetViewProjectionCache();
const sheetCacheUsage = new Map<number, WorksheetCacheUsage>();
let retainedSheetUsage: WorksheetCacheUsage = {
  rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
};
// Fetched image *bytes* (as Blobs) keyed by zip path. Twin of the docx render
// worker's raw cache. Cleared on re-parse so a reused worker never serves a
// stale file's image.
const rawParts = new BoundedRawPartCache({
  maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
  maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
});
// Keep the renderer and its orchestrator behind explicit module boundaries. The
// production worker is flattened into one self-contained asset, and a static
// function import can otherwise be hoisted past the initializers of shared draw
// dependencies that are also reached by optional renderers. Awaiting the modules
// preserves ESM initialization order, so the shared border dash tables and
// pattern-fill caches exist before the first bordered or filled cell is stroked.
const rendererModule = import('./renderer.js');
const orchestratorModule = import('./render-orchestrator.js');
const worksheetPull = new WorksheetPullWorker(
  () => host.archive,
  (sheetIndex, worksheet, measured, resourceUsage) => {
    const previous = sheetCache.get(sheetIndex);
    const previousUsage = sheetCacheUsage.get(sheetIndex);
    const nextUsage = addWorksheetCacheUsage(retainedSheetUsage, measured, previousUsage);
    assertWorksheetCacheUsage(
      nextUsage,
      'get-worksheet-worker',
      undefined,
      resourceUsage,
    );
    sheetCache.set(sheetIndex, worksheet);
    applyWorksheetFontRoutes(worksheet, providerFontRoutes);
    return {
      commit: () => {
        retainedSheetUsage = nextUsage;
        sheetCacheUsage.set(sheetIndex, measured);
      },
      rollback: () => {
        if (sheetCache.get(sheetIndex) !== worksheet) return;
        if (previous) sheetCache.set(sheetIndex, previous);
        else sheetCache.delete(sheetIndex);
      },
    };
  },
  (operation) => {
    const archive = host.archive;
    if (!archive) throw new Error('Workbook not loaded');
    return host.run(() => operation(archive));
  },
  (rows) => {
    if (workbook) resolveSharedStringRows(rows, workbook.sharedStrings);
  },
);

const rawPost = (msg: unknown, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);
const post = (msg: RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) =>
  rawPost(msg, transfer);
const svgDecodeClient = new WorkerSvgDecodeClient(rawPost);
const fontProvider = new FontProviderClient(rawPost);

/** In-worker image-byte loader (twin of the docx render-worker `getImage`). The
 *  orchestrator's `fetchImage` routes here in worker mode, so image bytes are
 *  read straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  return rawParts.get(path, mimeType, async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('Workbook not loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  });
}

self.onmessage = async (e: MessageEvent<
  RenderWorkerRequest | PullSessionCommand<number> | WorkerSvgDecodeResponse
>) => {
  const req = e.data;
  if (isWorkerFontResponse(req)) {
    await fontProvider.accept(req);
    return;
  }
  if (isWorkerSvgDecodeResponse(req)) {
    svgDecodeClient.accept(req);
    return;
  }
  if (isWorksheetPullCommand(req)) {
    await worksheetPull.dispatchSafely(req, post);
    return;
  }
  if (req.type === 'init') {
    host.setWasmInput(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }
  if (req.type === 'releaseViewProjection') {
    viewProjectionCache.release(req.projectionId);
    return;
  }
  const id = req.id;
  if (req.type === 'openSheetSession') worksheetPull.reserveOpen(req);
  try {
    if (req.type === 'openSheetSession') {
      await host.ensureReady();
      if (host.archive) {
        const retained = host.archive;
        host.run(() => retained.assert_healthy());
      }
      await worksheetPull.open(req.sheetIndex, req.sheetName, req);
      await worksheetPull.postOpenedSafely(
        req,
        () => post({
          type: 'sheetSessionOpened',
          id,
          sessionId: req.sessionId,
          operationId: req.operationId,
          generation: req.generation,
        }),
        (error) => post({ type: 'error', id, ...serializeWorkerError(error) }),
      );
      return;
    }
    if (req.type === 'parse') await worksheetPull.reset();
    await worksheetPull.run(async () => {
    await host.ensureReady();
    if (req.type !== 'parse' && host.archive) {
      const retained = host.archive;
      host.run(() => retained.assert_healthy());
    }
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
      fontProvider.reset();
      providerFontRoutes = {};
      providerEnabled = !!req.useFontProvider;
      generation += 1;
      fontsLoaded = Promise.resolve();
      viewProjectionCache.clear();
      sheetCacheUsage.clear();
      retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
      dropDecodedBitmapCache(getImage);
      dropSvgImageCache(getImage);
      rawParts.clear();
      renderers = await loadWorkerRenderers(req.renderers);
      const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
      // Construction + `parse()` run under `host.run` so a trap in EITHER poisons
      // + recycles the instance (and frees the archive). `setArchive` frees any
      // prior handle first — the re-parse dispose. `parse()` returns UTF-8 JSON
      // bytes (Result<Vec<u8>, JsValue>); decode + parse the workbook index here
      // (consumed in-worker, then a light copy is sent to the proxy as an object).
      const bootstrap = readXlsxArchiveBootstrap(
        () => host.run(() => {
          const archive = new XlsxArchive(
            new Uint8Array(req.data),
            maxEntry,
            maxTotal,
            maxEntries,
          );
          host.setArchive(archive);
          return JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
        }),
        () => host.run(() => host.archive!.resource_usage()),
      );
      workbook = bootstrap.workbook;
      if (req.useFontProvider) {
        providerFontRoutes = await fontProvider.resolve(
          xlsxFontProviderNames(workbook),
          generation,
        );
        applyWorkbookFontRoutes(workbook, providerFontRoutes);
      }
      if (req.useGoogleFonts) {
        // Mirror XlsxWorkbook._load exactly: queue Google Fonts substitutes for
        // every styled font name, plus the generic Arabic fallbacks. Fonts must
        // land before rendering (which measures text), so we keep the promise
        // and await it in the renderViewport handler.
        fontsLoaded = preloadGoogleFonts(xlsxFontPreloadNames(workbook), XLSX_GOOGLE_FONTS);
      }
      post({ type: 'parsed', id, workbook, usage: bootstrap.usage });
      return;
    }
    if (req.type === 'renderViewport') {
      if (!workbook) throw new Error('Workbook not loaded');
      await fontsLoaded;
      const { inheritSheetRenderCache, markAutoRowHeightsPrepared } = await rendererModule;
      const { renderWorksheetViewport } = await orchestratorModule;
      const ws = sheetCache.get(req.sheetIndex);
      if (!ws) throw new Error('Worksheet is not loaded through its pull session');
      if (providerEnabled) {
        providerFontRoutes = {
          ...providerFontRoutes,
          ...await fontProvider.resolve(xlsxWorksheetFontProviderNames(ws), generation),
        };
        applyWorksheetFontRoutes(ws, providerFontRoutes);
      }
      // Apply view-only size mutations to a render-local projection. Multiple
      // viewers may share this worker cache while retaining different outline
      // and resize state, so the cached worksheet itself must stay unchanged.
      const { sizeOverrides, ...renderOpts } = req.opts;
      const projected = viewProjectionCache.resolve(
        ws,
        req.sheetIndex,
        req.viewProjection,
        sizeOverrides,
      );
      const renderWorksheet = projected.worksheet;
      if (projected.created) inheritSheetRenderCache(ws, renderWorksheet);
      if (req.viewProjection?.autoRowHeightsPrepared) {
        markAutoRowHeightsPrepared(renderWorksheet);
      }
      const maximumDigitWidth = req.layoutMetrics?.maximumDigitWidth;
      if (maximumDigitWidth !== undefined) {
        if (!Number.isFinite(maximumDigitWidth) || maximumDigitWidth <= 0) {
          throw new Error('XLSX maximum digit width must be a finite positive number');
        }
        GridGeometry.forWorksheet(renderWorksheet, maximumDigitWidth);
      }
      const canvas = new OffscreenCanvas(1, 1); // orchestrator resizes it
      await renderWorksheetViewport(
        workerRenderDeps(renderWorksheet, workbook.styles, renderers),
        canvas,
        req.viewport,
        // Supply the in-worker byte loader so embedded images decode straight
        // from the retained archive (no main-thread round-trip).
        { ...renderOpts, fetchImage: getImage },
        svgDecodeClient.decode,
      );
      const bitmap = canvas.transferToImageBitmap();
      postOwnedImageBitmap(post, { type: 'viewportRendered', id, bitmap });
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      const archive = host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      // wasm-bindgen returns an owned full-span Uint8Array; transfer its
      // standalone buffer directly, matching the parse worker contract.
      const bytes = host.run(() => archive.extract_image(req.path).buffer as ArrayBuffer);
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'resourceUsage') {
      const archive = host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      const usage = host.run(() => decodeOoxmlResourceUsage(archive.resource_usage()));
      post({ type: 'resourceUsage', id, usage });
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
    });
  } catch (err) {
    if (req.type === 'openSheetSession') worksheetPull.abandonOpen(req.sessionId);
    try {
      post({ type: 'error', id, ...serializeWorkerError(err) });
    } catch {
      // Preserve cleanup and avoid an unhandled async worker rejection when
      // even the plain fallback response cannot be delivered.
    }
  }
};
