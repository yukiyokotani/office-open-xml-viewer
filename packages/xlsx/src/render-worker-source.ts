import type { CjkLang } from '@silurus/ooxml-core';
import { xlsxCjkFallback } from './google-fonts.js';
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
  loadOfficeFontFallbacks,
  unloadOfficeFontFallbacks,
  type OfficeFontFallbackRoute,
  WasmParserHost,
  dropDecodedBitmapCache,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import {
  decodeOoxmlResourceUsage,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  resourcePolicyForWasm,
  DEFAULT_XLSX_WORKSHEET_POLICY,
  normalizeXlsxWorksheetPolicy,
  xlsxWorksheetPolicyForWasm,
  type NormalizedXlsxWorksheetPolicy,
  serializeWorkerError,
  loadWorkerRenderers,
  isWorkerSvgDecodeResponse,
  postOwnedImageBitmap,
  WorkerSvgDecodeClient,
  type LoadedWorkerRenderers,
  type PullSessionCommand,
  type PullSessionResponse,
  type WorkerSvgDecodeResponse,
} from '@silurus/ooxml-core/worker';
import { workerRenderDeps } from './worker-render-deps.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames, xlsxOfficeFontRequests, xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import { officeRequestKey } from './shape-office-line.js';
import { resolveSharedStringRows } from './shared-strings.js';
import {
  addWorksheetCacheUsage,
  assertWorksheetCacheUsage,
  measureWorksheet,
  type WorksheetCacheUsage,
} from './worksheet-resource-limits.js';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { bindWorksheetPolicy } from './worksheet-policy-context.js';
import { WorksheetViewProjectionCache } from './worker-protocol.js';
import { evictWorkerWorksheets } from './internal/worksheet-cache.js';
import { readXlsxArchiveBootstrap, type XlsxArchiveBootstrap } from './internal/archive-bootstrap-source.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';
import { isWorksheetPullCommand, WorksheetPullWorker } from './worksheet-pull-source-worker.js';
import type { WorkerWorksheetSourceOwner, WorkerWorksheetArchive } from './internal/worker-worksheet-source.js';

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
let source: WorkerWorksheetSourceOwner<XlsxArchive> | undefined;
function executeArchive<T>(operation: (archive: WorkerWorksheetArchive) => T): T {
  if (source) return source.execute(operation);
  const archive = host.archive;
  if (!archive) throw new Error('Workbook not loaded');
  return host.run(() => operation(archive));
}
function sourceUsage(): Uint8Array | undefined {
  return source ? source.resourceUsage() : host.run(() => host.archive?.resource_usage());
}
let cjkFallback: CjkLang = 'jp';
let workbook: ParsedWorkbook | null = null;
let archiveBacked = false;
/** Normalized worksheet policy of the current document. Replaced only when a
 *  new document begins, after the previous pull session was reset. */
let worksheetPolicy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY;
let renderers: LoadedWorkerRenderers = {};
/** Settled before any render when `useGoogleFonts` was requested. The resolved
 *  value (the preloaded FontFace[]) is unused here: the worker owns its own
 *  FontFaceSet (`self.fonts`) and terminates with it, so there is nothing to
 *  release — only the sequencing (fonts landed before first paint) matters. */
let fontsLoaded: Promise<unknown> = Promise.resolve();
let officeFontFaces: FontFace[] = [];
let officeFontRoutes: Record<string, OfficeFontFallbackRoute> = {};
let checkedOfficeTupleSet = new Set<string>();
let googleSubstitutes = false;
let officeSheetLoads = new WeakMap<Worksheet, Promise<void>>();
let officeSheetLoadQueue: Promise<void> = Promise.resolve();

function startFontLoad(parsed: ParsedWorkbook, useGoogleFonts: boolean): void {
  googleSubstitutes = useGoogleFonts;
  fontsLoaded = Promise.all([
    useGoogleFonts
      ? preloadGoogleFonts(xlsxFontPreloadNames(parsed, cjkFallback), XLSX_GOOGLE_FONTS)
      : Promise.resolve([]),
    loadOfficeFontFallbacks(xlsxOfficeFontRequests(parsed)),
  ]).then(([, office]) => {
    officeFontFaces = office.faces;
    officeFontRoutes = office.routes;
    checkedOfficeTupleSet = new Set(office.checked);
  });
}
const sheetCache = new Map<number, Worksheet>();
const provisionalSheets = new Map<number, Worksheet>();
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
const delimitedTextModule = import('./delimited-text.js');
const worksheetPull = new WorksheetPullWorker(
  () => source?.cursor() ?? host.archive,
  (sheetIndex, worksheet, measured, resourceUsage) => {
    const previous = sheetCache.get(sheetIndex);
    const previousUsage = sheetCacheUsage.get(sheetIndex);
    const nextUsage = addWorksheetCacheUsage(
      retainedSheetUsage, measured, previousUsage, worksheetPolicy,
    );
    assertWorksheetCacheUsage(
      nextUsage,
      'get-worksheet-worker',
      undefined,
      resourceUsage,
      worksheetPolicy,
    );
    bindWorksheetPolicy(worksheet, worksheetPolicy);
    sheetCache.set(sheetIndex, worksheet);
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
    return executeArchive(operation);
  },
  (rows) => {
    if (workbook) resolveSharedStringRows(rows, workbook.sharedStrings);
  },
  {
    preview: (sheetIndex, worksheet) => {
      bindWorksheetPolicy(worksheet, worksheetPolicy);
      provisionalSheets.set(sheetIndex, worksheet);
      sheetCache.set(sheetIndex, worksheet);
    },
    stop: (sheetIndex) => {
      const preview = provisionalSheets.get(sheetIndex);
      provisionalSheets.delete(sheetIndex);
      if (preview && sheetCache.get(sheetIndex) === preview) sheetCache.delete(sheetIndex);
    },
  },
  () => worksheetPolicy,
);

const rawPost = (msg: unknown, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);
const post = (msg: RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) =>
  rawPost(msg, transfer);
const svgDecodeClient = new WorkerSvgDecodeClient(rawPost);

/** In-worker image-byte loader (twin of the docx render-worker `getImage`). The
 *  orchestrator's `fetchImage` routes here in worker mode, so image bytes are
 *  read straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  return rawParts.get(path, mimeType, async () => {
    const loaded = source?.cursor() ?? host.archive;
    if (!loaded) throw new Error('Workbook not loaded');
    const bytes = executeArchive((archive) => archive.extract_image(path));
    return new Blob([bytes as BlobPart], { type: mimeType });
  });
}

self.onmessage = async (e: MessageEvent<
  RenderWorkerRequest | PullSessionCommand<number> | WorkerSvgDecodeResponse
>) => {
  const req = e.data;

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
  if (req.type === 'evictWorksheets') {
    try {
      retainedSheetUsage = evictWorkerWorksheets(
        req.sheetIndices, sheetCache, sheetCacheUsage, retainedSheetUsage, viewProjectionCache,
        worksheetPolicy,
      );
      post({ type: 'worksheetsEvicted', id: req.id });
    } catch (error) {
      post({ type: 'error', id: req.id, ...serializeWorkerError(error) });
    }
    return;
  }
  const id = req.id;
  if (req.type === 'openSheetSession') worksheetPull.reserveOpen(req);
  try {
    if (req.type === 'openSheetSession') {
      if (!archiveBacked) throw new Error('Worksheet is already materialized');
      if (!source) await host.ensureReady();
      if (source?.cursor() ?? host.archive) executeArchive((archive) => archive.assert_healthy());
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
    // Normalize before any host/source/reset effect so invalid options leave
    // the current document (and any source claim) untouched.
    const requestPolicy = req.type === 'parse' || req.type === 'parseDelimitedText'
      ? normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: req.worksheetPolicy?.worksheet })
      : undefined;
    if (req.type === 'parse' || req.type === 'parseDelimitedText') {
      await worksheetPull.reset();
    }
    const runRequest = async (): Promise<void> => {
    if ((req.type === 'parse' && !req.source)
      || (req.type !== 'parseDelimitedText' && archiveBacked && !source)) await host.ensureReady();
    if (req.type !== 'parse' && req.type !== 'parseDelimitedText' && (source?.cursor() ?? host.archive)) {
      executeArchive((archive) => archive.assert_healthy());
    }
    if (req.type === 'parse' || req.type === 'parseDelimitedText') {
      // A re-parse starts a fresh document: drop any cached sheets / images so
      // we never serve stale data from a previous load. `imageCache` is now a
      // pure lookup map into the shared, per-`getImage` core caches (base raster,
      // duotone recolour, SVG); clearing it drops lookup references, and dropping
      // the three shared caches keyed by the module-level `getImage` closure
      // releases the GPU-backed ImageBitmaps and SVG object URLs AND prevents the
      // next document from being served a stale bitmap for an identically-named
      // zip path. Symmetric with XlsxWorkbook.destroy() and the docx/pptx render
      // workers (issue #781).
      cjkFallback = req.cjkFallback ?? 'jp';
      await fontsLoaded;
      await officeSheetLoadQueue;
      unloadOfficeFontFallbacks(officeFontFaces);
      officeFontFaces = [];
      officeFontRoutes = {};
      checkedOfficeTupleSet = new Set();
      officeSheetLoads = new WeakMap();
      officeSheetLoadQueue = Promise.resolve();
      sheetCache.clear();
      viewProjectionCache.clear();
      sheetCacheUsage.clear();
      retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
      dropDecodedBitmapCache(getImage);
      dropSvgImageCache(getImage);
      rawParts.clear();
      if (requestPolicy) worksheetPolicy = requestPolicy;
      renderers = await loadWorkerRenderers(req.renderers);
      if (req.type === 'parseDelimitedText') {
        source?.closeModelSource();
        source = undefined;
        host.disposeArchive();
        archiveBacked = false;
        const { parseDelimitedWorksheet } = await delimitedTextModule;
        const parsed = parseDelimitedWorksheet(req.data, req.options, worksheetPolicy);
        bindWorksheetPolicy(parsed.worksheet, worksheetPolicy);
        const measured = measureWorksheet(parsed.worksheet, worksheetPolicy);
        assertWorksheetCacheUsage(
          measured, 'parse-delimited-text-worker', undefined, undefined, worksheetPolicy,
        );
        workbook = parsed.workbook;
        cjkFallback = xlsxCjkFallback(workbook, cjkFallback);
        retainedSheetUsage = measured;
        sheetCache.set(0, parsed.worksheet);
        sheetCacheUsage.set(0, measured);
        startFontLoad(parsed.workbook, !!req.useGoogleFonts);
        const worksheetJson = new TextEncoder()
          .encode(JSON.stringify(parsed.worksheet)).buffer as ArrayBuffer;
        post({
          type: 'delimitedTextParsed',
          id,
          workbook,
          worksheetJson,
        }, [worksheetJson]);
        return;
      }
      archiveBacked = true;
      source?.closeModelSource();
      source = undefined;
      let bootstrap: XlsxArchiveBootstrap<ParsedWorkbook>;
      if (req.source) {
        host.disposeArchive();
        if (!req.sourceOwnerUrl) throw new TypeError('XLSX source owner URL is missing');
        const { WorkerWorksheetSourceOwner } = await import(/* @vite-ignore */ req.sourceOwnerUrl) as typeof import('./internal/worker-worksheet-source.js');
        const owner = new WorkerWorksheetSourceOwner(host);
        source = owner;
        // This worker owns the renderer, so it measures the model source's
        // Normal font with the same computeMdw that sizes the painted grid.
        const { computeMdw } = await rendererModule;
        await owner.openModelSource(
          new Uint8Array(req.data),
          req.source,
          (font) => computeMdw(
            font.family,
            font.sizePt,
            undefined,
            !!req.useGoogleFonts,
            font.bold ? 700 : 400,
            font.italic ? 'italic' : 'normal',
          ),
          req.sourceTransfer,
        );
        bootstrap = readXlsxArchiveBootstrap(
          () => JSON.parse(new TextDecoder().decode(
            owner.execute((archive) => archive.parse()),
          )) as ParsedWorkbook,
          () => owner.resourceUsage(),
        );
      } else {
        const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
        // Keep construction and parse in the same host.run as the ordinary
        // OOXML worker on main. Both operations share one trap boundary.
        bootstrap = readXlsxArchiveBootstrap(
          () => host.run(() => {
            const archive = new XlsxArchive(
              new Uint8Array(req.data), maxEntry, maxTotal, maxEntries,
            );
            host.setArchive(archive);
            archive.set_worksheet_limits(...xlsxWorksheetPolicyForWasm(worksheetPolicy));
            return JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
          }),
          () => host.run(() => host.archive!.resource_usage()),
        );
      }
      workbook = bootstrap.workbook;
      const maximumDigitWidth = source?.maximumDigitWidth;
      if (maximumDigitWidth !== undefined) workbook.layoutMetrics = { maximumDigitWidth };
      cjkFallback = xlsxCjkFallback(workbook, cjkFallback);
      startFontLoad(workbook, !!req.useGoogleFonts);
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
      let sheetFonts = officeSheetLoads.get(ws);
      if (!sheetFonts) {
        // Different sheets may be rendered concurrently. Serialize their
        // worksheet-only probes so a completed missing tuple is tried once,
        // while a tuple omitted by a preflight budget can be retried later.
        sheetFonts = officeSheetLoadQueue.then(async () => {
          const extraRequests = xlsxWorksheetOfficeFontRequests(ws).filter((request) => {
            const key = officeRequestKey(request);
            return !checkedOfficeTupleSet.has(key) && !(key in officeFontRoutes);
          });
          if (extraRequests.length === 0) return;
          const extra = await loadOfficeFontFallbacks(extraRequests);
          officeFontFaces.push(...extra.faces);
          for (const key of extra.checked) {
            if (checkedOfficeTupleSet.has(key)) continue;
            checkedOfficeTupleSet.add(key);
          }
          Object.assign(officeFontRoutes, extra.routes);
        });
        officeSheetLoadQueue = sheetFonts.catch(() => {});
        officeSheetLoads.set(ws, sheetFonts);
      }
      await sheetFonts;
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
      // The orchestrator resizes it. A caller-transferred surface keeps the
      // caller's inherited canvas language; a request without one keeps a
      // local surface.
      const canvas = req.canvas ?? new OffscreenCanvas(1, 1);
      await renderWorksheetViewport(
        { ...workerRenderDeps(renderWorksheet, workbook.styles, renderers), cjkFallback },
        canvas,
        req.viewport,
        // Supply the in-worker byte loader so embedded images decode straight
        // from the retained archive (no main-thread round-trip). Pass the
        // viewer's MDW through the render bind: seeding GridGeometry before
        // that bind is ineffective because the worker's FontFaceSet can
        // invalidate it on first use.
        { ...renderOpts, authoritativeMdw: req.layoutMetrics?.maximumDigitWidth
          ?? workbook?.layoutMetrics?.maximumDigitWidth,
          officeFontRoutes, googleSubstitutes, fetchImage: getImage },
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
      const archive = source?.cursor() ?? host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      // wasm-bindgen returns an owned full-span Uint8Array; transfer its
      // standalone buffer directly, matching the parse worker contract.
      const bytes = executeArchive((current) => source
        ? source.copyBytes(current.extract_image(req.path))
        : current.extract_image(req.path).buffer as ArrayBuffer);
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'resourceUsage') {
      const archive = source?.cursor() ?? host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      const bytes = sourceUsage();
      post({
        type: 'resourceUsage',
        id,
        usage: bytes === undefined ? undefined : decodeOoxmlResourceUsage(bytes),
      });
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = source?.cursor() ?? host.archive;
      if (!archive) throw new Error('Workbook not loaded');
      const markdown = source ? source.toMarkdown() : host.run(() => host.archive!.to_markdown());
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
    };
    if (req.type === 'renderViewport' && provisionalSheets.has(req.sheetIndex)) {
      await runRequest();
    } else {
      await worksheetPull.run(runRequest);
    }
  } catch (err) {
    if (req.type === 'openSheetSession') worksheetPull.abandonOpen(req.sessionId);
    if (req.type === 'parse') {
      try { source?.closeModelSource(); } catch {}
    }
    try {
      post({ type: 'error', id, ...serializeWorkerError(err) });
    } catch {
      // Preserve cleanup and avoid an unhandled async worker rejection when
      // even the plain fallback response cannot be delivered.
    }
  }
};
