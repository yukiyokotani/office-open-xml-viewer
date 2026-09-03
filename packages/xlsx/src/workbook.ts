import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/xlsx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  dropDecodedBitmapCache,
  dropSvgImageCache,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type TiffRenderer,
  OoxmlResourceLimitError,
  type OoxmlResourceMetrics,
  workerRendererDescriptors,
} from '@silurus/ooxml-core';
import { resolveOfficeInputWithOptionalConversion } from '@silurus/ooxml-core/internal/legacy-office-conversion';
import {
  deserializeWorkerError,
  disposeRejectedLoad,
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  readLatestOoxmlResourceMetrics,
  normalizeResourcePolicy,
  type NormalizedOoxmlResourcePolicy,
  PULL_SESSION_PROTOCOL,
  type PullSessionResponse,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
  respondToWorkerSvgDecodeRequest,
} from '@silurus/ooxml-core/worker';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, XlsxRenderViewportOptions, WorkerRequest, WorkerResponse, Cell, SheetVisibility, XlsxComment } from './types.js';
import { selectSheetVisibility } from './sheet-visibility.js';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import { formatCellValue } from './number-format.js';
import {
  addWorksheetUsage,
  addWorksheetCacheUsage,
  assertWorksheetCacheUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  completeWorksheetUsage,
  measureRows,
  measureWorksheet,
  type WorksheetCacheUsage,
  type WorksheetModelUsage,
} from './worksheet-resource-limits.js';
import {
  parseListFormula,
  resolveListValues,
  type ResolvedList,
} from './validation-list.js';
import type {
  RenderWorkerRequest,
  RenderWorkerResponse,
  WireRenderViewportOptions,
} from './worker-protocol.js';
import {
  createSizeOverriddenWorksheet,
  extractViewerRenderContext,
} from './worker-protocol.js';
import {
  isXlsxWorksheetPullResponse,
  XlsxWorksheetPullClient,
} from './worksheet-pull-client.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { applyAutoRowHeights, inheritSheetRenderCache } from './renderer.js';
import {
  assertDelimitedTextSourceBytes,
  resolveDelimitedTextOptions,
  type ResolvedDelimitedTextOptions,
  type XlsxSheetLoadOptions,
} from './delimited-text.js';
import { readDelimitedTextResponse } from './delimited-text-source.js';
import type {
  DelimitedTextParseRequest,
  DelimitedTextParseResponse,
} from './delimited-text-protocol.js';

/** Public options for {@link XlsxWorkbook.renderViewportToBitmap}. Viewer-only
 * worksheet projection state is intentionally not part of this contract. */
export type RenderViewportToBitmapOptions = Omit<
  XlsxRenderViewportOptions,
  'onTextRun'
> & { width: number; height: number };

/** @internal Viewer-only hook for retaining web fonts in the canvas document. */
export const retainXlsxViewerFonts = Symbol('retain-xlsx-viewer-fonts');
/** @internal Resolve display-time row auto-fit on a viewer-owned projection. */
export const prepareXlsxViewerRowHeights = Symbol('prepare-xlsx-viewer-row-heights');
/** @internal Release worker-side viewer projection cache entries. */
export const releaseXlsxViewerProjection = Symbol('release-xlsx-viewer-projection');
/** @internal XlsxSheetViewer-only source dispatcher. */
export const loadXlsxSheetSource = Symbol('load-xlsx-sheet-source');

interface RetainedFontSet {
  refs: number;
  faces: FontFace[] | null;
  readonly loading: Promise<FontFace[]>;
}

/** Options for {@link XlsxWorkbook.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `resourceLimits`, the
 *  deprecated `maxZipEntryBytes` alias, and `math`) with worker rendering. */
export interface LoadOptions extends CoreLoadOptions {
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link XlsxWorkbook.renderViewportToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   * Built-in optional renderers use the same injection options in both modes
   * and are reconstructed inside the worker. Custom renderer objects use their
   * documented fallback in worker mode.
   */
  mode?: 'main' | 'worker';
}

type WorkbookBridge = WorkerBridge<
  WorkerResponse | RenderWorkerResponse | DelimitedTextParseResponse |
  PullSessionResponse<ArrayBuffer, number>
>;

export class XlsxWorkbook {
  private metrics: OoxmlResourceMetricsSession | null = null;
  private bridge: WorkbookBridge | null = null;
  private delimitedTextBacked = false;
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  /** One materialization per sheet at a time. This becomes the ownership seam
   * for the bounded worksheet cursor: concurrent callers share one cursor and
   * one eventual mutable compatibility object instead of doubling peak work. */
  private sheetLoads = new Map<number, Promise<Worksheet>>();
  /** Cache of fetched image *bytes* (as Blobs) keyed by zip path, populated by
   *  {@link XlsxWorkbook.getImage}. Twin of pptx/docx's per-instance
   *  raw-part owner; decoded sources are owned separately by core. */
  private readonly rawParts = new BoundedRawPartCache({
    maxEntries: HARD_MAX_RAW_PART_CACHE_ENTRIES,
    maxBytes: HARD_MAX_RAW_PART_CACHE_BYTES,
  });
  /** Public archive-queue reservations. Kept separate so an active render does
   * not await a same-path load that is queued behind that render. */
  private queuedImageLoads = new Map<string, Promise<Blob>>();
  /** One stable closure per instance: core's path-keyed SVG cache namespaces on
   *  this identity, so two open workbooks never swap a shared zip path (e.g.
   *  xl/media/image1.svg). Reusing one reference also lets the SVG cache hit
   *  across viewport renders. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImageWithinArchiveOperation(path, mime);
  private resourcePolicy: NormalizedOoxmlResourcePolicy | null = null;
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderViewport` call reuses it — equations in shapes render when present,
   *  and are skipped when omitted. */
  private math: MathRenderer | undefined;
  /** Optional synchronous 3-D chart renderer. Worker mode reconstructs the
   * built-in implementation from its serializable identity. */
  private threeD: ChartThreeDRenderer | undefined;
  /** Optional synchronous Region Map renderer. Worker mode reconstructs the
   * built-in implementation from its serializable identity. */
  private regionMap: ChartRegionMapRenderer | undefined;
  /** Optional Microsoft ChartEx renderer. */
  private chartEx: ChartExRenderer | undefined;
  /** Optional TIFF codec. Worker mode reconstructs the built-in implementation. */
  private tiff: TiffRenderer | undefined;
  /** Web-font registrations are per FontFaceSet. Same-origin child windows have
   * their own set even when they share this workbook instance. */
  private googleFontNames: string[] = [];
  private readonly retainedFontSets = new Map<FontFaceSet, RetainedFontSet>();
  private fontsDestroyed = false;
  private _mode: 'main' | 'worker' = 'main';
  private generation = 0;
  private archiveOperationTail: Promise<void> = Promise.resolve();
  private worksheetPullClient: XlsxWorksheetPullClient | null = null;
  private workerTimeoutMs: number | undefined;
  private retainedSheetUsage: WorksheetCacheUsage = {
    rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
  };
  /** First fatal model/package violation. Compatibility materialization happens
   * on main, so this latch is the document-level poison boundary for every
   * later public operation on the same workbook instance. */
  private resourceFailure: OoxmlResourceLimitError | null = null;

  private constructor(
    worker: Worker | null,
    mode: 'main' | 'worker',
    wasmUrlOverride?: string | URL,
    initializeWasm = true,
  ) {
    this._mode = mode;
    if (!worker) return;
    this.bridge = new WorkerBridge<
      WorkerResponse | RenderWorkerResponse | PullSessionResponse<ArrayBuffer, number>
    >(worker, {
      correlate: (res) =>
        'protocol' in res && res.protocol === PULL_SESSION_PROTOCOL
          ? res.requestId
          : 'id' in res
            ? res.id
            : undefined,
      // Pull `kind:error` is a correlated protocol value consumed by
      // BoundedPullSession. Ordinary `type:error` retains WorkerBridge's
      // historical rejection behavior.
      toError: (res) =>
        'type' in res && res.type === 'error' ? deserializeWorkerError(res) : undefined,
      onUnsolicited: (res) => {
        respondToWorkerSvgDecodeRequest(
          (message, transfer) => (
            worker.postMessage as (value: unknown, transfer?: Transferable[]) => void
          )(message, transfer),
          res,
        );
      },
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    if (initializeWasm) {
      const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
      this.bridge.post({ type: 'init', wasmUrl } satisfies WorkerRequest);
    }
  }

  /** The render mode this loaded workbook owns. Injected viewers use this fact
   *  to select direct-canvas or worker-bitmap rendering without probing. */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }

  /** @internal XlsxSheetViewer-only source dispatcher. */
  static async [loadXlsxSheetSource](
    source: string | ArrayBuffer,
    opts: LoadOptions,
    sourceOptions: XlsxSheetLoadOptions = {},
  ): Promise<XlsxWorkbook> {
    if (sourceOptions.format === undefined || sourceOptions.format === 'xlsx') {
      return await this.load(source, opts);
    }
    if (
      sourceOptions.format === 'csv'
      || sourceOptions.format === 'tsv'
      || sourceOptions.format === 'delimited-text'
    ) {
      return await this.loadDelimitedText(source, opts, sourceOptions);
    }
    throw new TypeError('Unsupported XlsxSheetViewer source format');
  }

  private static async loadDelimitedText(
    source: string | ArrayBuffer,
    opts: LoadOptions,
    sourceOptions: Exclude<XlsxSheetLoadOptions, Readonly<{ format?: 'xlsx' }>>,
  ): Promise<XlsxWorkbook> {
    const delimited = resolveDelimitedTextOptions(sourceOptions);
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode,
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    try {
      if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
        throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
      }
      const callerBuffer = typeof source === 'string' ? undefined : source;
      let buffer: ArrayBuffer;
      if (typeof source === 'string') {
        const response = await fetch(source);
        if (!response.ok) {
          throw new Error(`Failed to fetch: ${response.status} ${response.statusText}`);
        }
        buffer = await readDelimitedTextResponse(response);
      } else {
        buffer = source;
      }
      // Reject caller-owned buffers before `slice(0)` doubles their retained
      // memory, and keep this admission gate in front of worker creation.
      assertDelimitedTextSourceBytes(buffer.byteLength);
      metrics.setSourceBytes(buffer.byteLength);
      metrics.checkpoint('source ready');
      const worker = mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : (await import('./delimited-text-worker-host')).createDelimitedTextWorker();
      let workbook: XlsxWorkbook | undefined;
      try {
        const loaded = new XlsxWorkbook(worker, mode, undefined, false);
        workbook = loaded;
        loaded.metrics = metrics;
        await loaded._loadDelimitedText(
          callerBuffer === buffer ? buffer.slice(0) : buffer,
          opts,
          resourceOptions.policy,
          delimited,
        );
        if (mode === 'main') {
          loaded.bridge?.terminate();
          loaded.bridge = null;
        }
        metrics.checkpoint('worksheet ready');
        metrics.succeed({ sheets: 1 });
        return loaded;
      } catch (error) {
        const rejectedWorkbook = workbook;
        disposeRejectedLoad(
          worker,
          rejectedWorkbook ? () => rejectedWorkbook.destroy() : undefined,
        );
        throw error;
      }
    } catch (error) {
      metrics.fail(error);
      throw error;
    }
  }

  /** Parse an XLSX from a URL or ArrayBuffer. */
  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<XlsxWorkbook> {
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode,
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    try {
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    // Resolve the bytes first, then resolve the container on the main thread —
    // before spinning up the worker. A normal ZIP passes through unchanged; an
    // Agile-encrypted CFB is decrypted when `opts.password` is supplied
    // ([MS-OFFCRYPTO]); a password-protected file without a password, or a
    // legacy-binary / unknown CFB, becomes a typed OoxmlError (whose `instanceof`
    // would not survive the worker boundary). The resolved buffer is handed to
    // `_load` so a URL source is not fetched twice.
    const callerBuffer = typeof source === 'string' ? undefined : source;
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    buffer = toArrayBuffer(await resolveOfficeInputWithOptionalConversion(
      buffer,
      'xlsx',
      opts.legacyConversion,
      opts.password,
    ));
    const preserveCallerBuffer = buffer === callerBuffer;
    metrics.setSourceBytes(buffer.byteLength);
    metrics.checkpoint('container ready');
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    let wb: XlsxWorkbook | undefined;
    try {
      wb = new XlsxWorkbook(worker, mode, opts.wasmUrl);
      wb.metrics = metrics;
      await wb._load(
        buffer,
        opts,
        resourceOptions.policy,
        (usage) => metrics.observeUsage(usage),
        preserveCallerBuffer,
      );
      metrics.checkpoint('workbook index ready');
      metrics.succeed({ sheets: wb.sheetCount });
      return wb;
    } catch (error) {
      const rejectedWorkbook = wb;
      disposeRejectedLoad(worker, rejectedWorkbook ? () => rejectedWorkbook.destroy() : undefined);
      throw error;
    }
    } catch (error) {
      metrics.fail(error);
      throw error;
    }
  }

  // `load()` always resolves a URL/string source to an ArrayBuffer (via
  // resolveOoxmlContainer, so decryption sees the container before the render
  // worker is constructed) before calling `_load`, so this only ever receives
  // an ArrayBuffer — no separate string-source fetch branch is needed here.
  private async _load(
    data: ArrayBuffer,
    opts: LoadOptions = {},
    resourcePolicy: NormalizedOoxmlResourcePolicy = normalizeResourcePolicy(opts),
    onUsage?: (usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void,
    preserveCallerBuffer = false,
  ): Promise<void> {
    const bridge = this.requireBridge();
    this.resourceFailure = null;
    this.retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
    this.sheetCache.clear();
    await this.worksheetPullClient?.cancelAll('closed');
    this.worksheetPullClient = null;
    this.generation = (this.generation ?? 0) + 1;
    this.resourcePolicy = resourcePolicy;
    this.workerTimeoutMs = opts.workerTimeoutMs;
    this.math = this._mode === 'worker' ? undefined : opts.math;
    this.threeD = this._mode === 'worker' ? undefined : opts.threeD;
    this.regionMap = this._mode === 'worker' ? undefined : opts.regionMap;
    this.chartEx = this._mode === 'worker' ? undefined : opts.chartEx;
    this.tiff = this._mode === 'worker' ? undefined : opts.tiff;
    const rendererDescriptors = this._mode === 'worker'
      ? workerRendererDescriptors(opts)
      : undefined;
    if (opts.math && this._mode === 'worker' && !rendererDescriptors?.math) {
      console.warn(
        "[ooxml] a custom math renderer cannot cross the worker boundary; equations will be skipped in mode: 'worker'. Use the math renderer from @silurus/ooxml/math.",
      );
    }
    if (opts.threeD && this._mode === 'worker' && !rendererDescriptors?.threeD) {
      console.warn(
        "[ooxml] a custom 3-D chart renderer cannot cross the worker boundary; charts use their 2-D family fallback in mode: 'worker'. Use the renderer from @silurus/ooxml/three-d.",
      );
    }
    if (opts.regionMap && this._mode === 'worker' && !rendererDescriptors?.regionMap) {
      console.warn(
        "[ooxml] a custom Region Map renderer cannot cross the worker boundary; geospatial charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/region-map.",
      );
    }
    if (opts.chartEx && this._mode === 'worker' && !rendererDescriptors?.chartEx) {
      console.warn(
        "[ooxml] a custom ChartEx renderer cannot cross the worker boundary; ChartEx charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/chart-ex.",
      );
    }
    if (opts.tiff && this._mode === 'worker' && !rendererDescriptors?.tiff) {
      console.warn(
        "[ooxml] a custom TIFF codec cannot cross the worker boundary; recognized TIFF images will use an unavailable-image placeholder in mode: 'worker'. Use the codec from @silurus/ooxml/tiff to display them.",
      );
    }
    // In worker mode the worker preloads fonts before its first render
    // (rendering measures text), so the flag is forwarded; in main mode fonts
    // are loaded here after parse.
    // Preserve XLSX's historical caller-owned ArrayBuffer contract only when
    // the resolved ZIP is literally the caller's buffer. URL and decrypted
    // buffers are library-owned and can transfer directly without a peak copy.
    const workerData = preserveCallerBuffer ? data.slice(0) : data;
    const parsed = await bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({
              type: 'parse',
              id,
              data: workerData,
              resourcePolicy,
              useGoogleFonts: !!opts.useGoogleFonts,
              renderers: rendererDescriptors,
            } satisfies RenderWorkerRequest)
          : ({
              type: 'parse',
              id,
              data: workerData,
              resourcePolicy,
            } satisfies WorkerRequest),
      [workerData],
      { timeoutMs: opts.workerTimeoutMs },
    );
    // Both modes carry the light, workbook-level ParsedWorkbook back, so
    // sheetNames / tabColors / resolveValidationList keep working. In parse mode
    // it arrives as transferred UTF-8 JSON bytes — decode + parse once here.
    if (this._mode === 'worker') {
      const response = parsed as Extract<RenderWorkerResponse, { type: 'parsed' }>;
      this.parsedWorkbook = response.workbook;
      if (response.usage) onUsage?.(response.usage);
    } else {
      const { workbookJson, usage } = parsed as Extract<WorkerResponse, { type: 'parsed' }>;
      if (usage) onUsage?.(usage);
      this.parsedWorkbook = JSON.parse(
        new TextDecoder().decode(new Uint8Array(workbookJson)),
      ) as ParsedWorkbook;
    }
    const parsedWorkbook = this.parsedWorkbook;
    if (!parsedWorkbook) throw new Error('XLSX worker returned no workbook metadata');
    this.ensureWorksheetPullClient();
    // #773: a workbook-level degradation (a present-but-corrupt shared part such
    // as `xl/sharedStrings.xml`, which blanks every string cell across all sheets)
    // still opens the workbook, but must not be SILENT. Surface it once at load —
    // the model also carries it on `workbook.parseError` for callers that inspect
    // it. Per-sheet placeholders (a broken worksheet) already surface via the
    // sheet-grid overlay, so they are not re-logged here.
    const workbookError = parsedWorkbook.workbook.parseError;
    if (workbookError) {
      console.warn(`[ooxml] xlsx opened with a degraded part: ${workbookError}`);
    }
    if (opts.useGoogleFonts) {
      // The composite viewer computes hit/scroll/overlay geometry on the main
      // realm even when paint runs in a worker. Register the same fallback
      // faces in both realms before any worksheet geometry snapshot is made so
      // ECMA-376 MDW is identical across paint and interaction.
      this.googleFontNames = [...xlsxFontPreloadNames(parsedWorkbook)];
      if (typeof document !== 'undefined' && document.fonts) {
        await this.retainFontsInSet(document.fonts);
      }
    }
  }

  private async _loadDelimitedText(
    data: ArrayBuffer,
    opts: LoadOptions,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    options: ResolvedDelimitedTextOptions,
  ): Promise<void> {
    const bridge = this.requireBridge();
    this.delimitedTextBacked = true;
    this.resourcePolicy = resourcePolicy;
    this.workerTimeoutMs = opts.workerTimeoutMs;
    this.generation++;
    this.math = this._mode === 'worker' ? undefined : opts.math;
    this.threeD = this._mode === 'worker' ? undefined : opts.threeD;
    this.regionMap = this._mode === 'worker' ? undefined : opts.regionMap;
    this.chartEx = this._mode === 'worker' ? undefined : opts.chartEx;
    this.tiff = this._mode === 'worker' ? undefined : opts.tiff;
    const rendererDescriptors = this._mode === 'worker'
      ? workerRendererDescriptors(opts)
      : undefined;
    const response = await bridge.request(
      (id) => ({
        type: 'parseDelimitedText',
        id,
        data,
        options,
        useGoogleFonts: !!opts.useGoogleFonts,
        renderers: rendererDescriptors,
      } satisfies DelimitedTextParseRequest),
      [data],
      { timeoutMs: opts.workerTimeoutMs },
    ) as Extract<DelimitedTextParseResponse, { type: 'delimitedTextParsed' }>;
    const worksheet = JSON.parse(
      new TextDecoder().decode(response.worksheetJson),
    ) as Worksheet;
    const sheets = response.workbook.workbook.sheets;
    if (sheets.length !== 1 || sheets[0]?.name !== worksheet.name) {
      throw new Error('Delimited text worker returned inconsistent worksheet metadata');
    }
    const measured = measureWorksheet(worksheet);
    assertWorksheetModelUsage(measured, 'load-delimited-text', undefined);
    assertWorksheetJsonBytes(measured.jsonBytes, 'load-delimited-text', undefined);
    assertWorksheetCacheUsage(measured, 'load-delimited-text', undefined);
    this.parsedWorkbook = response.workbook;
    this.sheetCache.set(0, worksheet);
    this.retainedSheetUsage = measured;

    if (opts.useGoogleFonts) {
      this.googleFontNames = [...xlsxFontPreloadNames(response.workbook)];
      if (typeof document !== 'undefined' && document.fonts) {
        await this.retainFontsInSet(document.fonts);
      }
    }
  }

  private async retainFontsInSet(fontSet: FontFaceSet): Promise<() => void> {
    if (this.googleFontNames.length === 0 || this.fontsDestroyed) return () => undefined;
    let retained = this.retainedFontSets.get(fontSet);
    if (retained) {
      retained.refs++;
    } else {
      const loading = preloadGoogleFonts(this.googleFontNames, XLSX_GOOGLE_FONTS, fontSet);
      retained = { refs: 1, faces: null, loading };
      this.retainedFontSets.set(fontSet, retained);
      loading.then((faces) => {
        retained!.faces = faces;
        if (this.fontsDestroyed) unloadGoogleFonts(faces);
      });
    }
    await retained.loading;
    let released = false;
    return () => {
      if (released) return;
      released = true;
      const current = this.retainedFontSets.get(fontSet);
      if (current !== retained) return;
      current.refs--;
      if (current.refs > 0) return;
      this.retainedFontSets.delete(fontSet);
      if (current.faces) unloadGoogleFonts(current.faces);
      else current.loading.then(unloadGoogleFonts);
    };
  }

  /** @internal Retain required faces in the document that owns a viewer canvas. */
  async [retainXlsxViewerFonts](targetDocument: Document): Promise<() => void> {
    return await this.retainFontsInSet(targetDocument.fonts);
  }

  /** @internal Fonts are retained before this hook runs, so Canvas text
   * measurement observes the same faces that the subsequent paint uses. */
  [prepareXlsxViewerRowHeights](worksheet: Worksheet, ctx: CanvasRenderingContext2D): void {
    if (!this.parsedWorkbook) return;
    applyAutoRowHeights(ctx, worksheet, this.parsedWorkbook.styles);
  }

  get sheetNames(): string[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.name) ?? [];
  }

  get sheetCount(): number {
    return this.parsedWorkbook?.workbook.sheets.length ?? 0;
  }

  /** Per-sheet tab colors (`#RRGGBB`) parallel to {@link sheetNames}.
   *  `null` for sheets that declare no tab color. */
  get tabColors(): (string | null)[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.tabColor ?? null) ?? [];
  }

  /**
   * Full visibility fact for the sheet at `sheetIndex` (0-based):
   * `'visible'` | `'hidden'` | `'veryHidden'` (`<sheet state>`, ECMA-376
   * §18.2.19). NOT clamped — out-of-range / non-integer ⇒ `'visible'`. This is a
   * *fact*; deciding what to do with a hidden sheet (hide/skip/dim its tab) is
   * {@link XlsxViewer}'s policy. `'veryHidden'` is revealable only
   * programmatically in Excel; it is surfaced distinctly here.
   */
  sheetVisibility(sheetIndex: number): SheetVisibility {
    return selectSheetVisibility(this.parsedWorkbook?.workbook.sheets ?? [], sheetIndex);
  }

  /**
   * Whether the sheet at `sheetIndex` is hidden or veryHidden. Convenience over
   * {@link sheetVisibility}; mirrors {@link PptxPresentation.isHidden} (non-
   * clamped: out-of-range / non-integer ⇒ `false`).
   */
  isHidden(sheetIndex: number): boolean {
    return this.sheetVisibility(sheetIndex) !== 'visible';
  }

  async getWorksheet(sheetIndex: number): Promise<Worksheet> {
    this.assertResourceHealthy();
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) return cached;
    const active = this.sheetLoads.get(sheetIndex);
    if (active) return active;
    const load = this.loadWorksheet(sheetIndex);
    this.sheetLoads.set(sheetIndex, load);
    try {
      return await load;
    } finally {
      if (this.sheetLoads.get(sheetIndex) === load) this.sheetLoads.delete(sheetIndex);
    }
  }

  /** Detached comments for one worksheet, in authored order. Worksheet models
   * are materialized lazily, so this accessor is asynchronous. */
  async getComments(sheetIndex: number): Promise<readonly Readonly<XlsxComment>[]> {
    const worksheet = await this.getWorksheet(sheetIndex);
    return structuredClone(worksheet.comments ?? []);
  }

  /** Return a fresh content-free metrics snapshot, including lazy worksheet and
   * media work completed since load. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    const metrics = this.metrics;
    if (!metrics) throw new Error('Workbook not loaded');
    if (this.delimitedTextBacked) {
      const report = metrics.current();
      if (!report) throw new Error('OOXML resource metrics are not ready');
      return report;
    }
    return readLatestOoxmlResourceMetrics(metrics, async (timeoutMs) => {
      const response = await this.requireBridge().request(
        (id) => ({ type: 'resourceUsage', id }) satisfies WorkerRequest,
        undefined,
        { timeoutMs },
      );
      return (response as Extract<WorkerResponse, { type: 'resourceUsage' }>).usage;
    });
  }

  private async loadWorksheet(sheetIndex: number): Promise<Worksheet> {
    // The worker retained its transferred archive at parse time; loaded state
    // is represented by the workbook bootstrap, not a duplicate source buffer.
    if (!this.parsedWorkbook) {
      throw new Error('Workbook not loaded');
    }
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);

    return this.runArchiveOperation(() => this.loadWorksheetStream(sheetIndex, sheetMeta.name));
  }

  private async loadWorksheetStream(sheetIndex: number, sheetName: string): Promise<Worksheet> {
    const client = this.ensureWorksheetPullClient();
    const rows: Worksheet['rows'] = [];
    let modelUsage: WorksheetModelUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0 };
    let terminal: Worksheet | undefined;
    let nextCacheUsage: WorksheetCacheUsage | undefined;
    // The compatibility adapter knows the workbook sheet index/name but not
    // the resolved OPC relationship target. Omit `part` rather than fabricate
    // a package address; Rust-originated violations carry the real xl/... path.
    const part = undefined;
    try {
      for await (const unit of client.stream(sheetIndex, sheetName)) {
        if (unit.kind === 'rows') {
          const nextUsage = addWorksheetUsage(modelUsage, measureRows(unit.rows));
          assertWorksheetModelUsage(
            nextUsage,
            'get-worksheet',
            part,
            unit.usage,
          );
          rows.push(...unit.rows);
          modelUsage = nextUsage;
          continue;
        }
        const worksheet = unit.worksheet;
        worksheet.rows = worksheet.parseError ? [] : rows;
        // Terminal metadata contains no rows, but measure the final public model
        // to cover exact monolithic JSON before its cache admission.
        const measured = completeWorksheetUsage(
          worksheet,
          worksheet.parseError
            ? { rows: 0, cells: 0, ownedUtf8Bytes: 0 }
            : modelUsage,
        );
        assertWorksheetModelUsage(
          measured,
          'get-worksheet',
          part,
          unit.usage,
        );
        assertWorksheetJsonBytes(
          measured.jsonBytes,
          'get-worksheet',
          part,
          unit.usage,
        );
        const retainedUsage = this.retainedSheetUsage ?? {
          rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
        };
        const nextCache = addWorksheetCacheUsage(retainedUsage, measured);
        assertWorksheetCacheUsage(
          nextCache,
          'get-worksheet',
          part,
          unit.usage,
        );
        terminal = worksheet;
        nextCacheUsage = nextCache;
      }
      if (!terminal || !nextCacheUsage) {
        throw new Error(`XLSX worksheet ${sheetIndex} did not produce a terminal model`);
      }
      // The coordinator has ACKed the accepted terminal before it completes.
      // Only now commit Browser-retained cache ownership/accounting.
      this.retainedSheetUsage = nextCacheUsage;
      this.sheetCache.set(sheetIndex, terminal);
      return terminal;
    } catch (error) {
      if (error instanceof OoxmlResourceLimitError) this.resourceFailure ??= error;
      throw error;
    }
  }

  private ensureWorksheetPullClient(): XlsxWorksheetPullClient {
    if (this.worksheetPullClient) return this.worksheetPullClient;
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    this.worksheetPullClient = new XlsxWorksheetPullClient({
      generation: this.generation || 1,
      transport: this.requireBridge().transport(isXlsxWorksheetPullResponse),
      sharedStrings: this.parsedWorkbook.sharedStrings,
      timeoutMs: this.workerTimeoutMs,
      open: async (sheetIndex, sheetName, identity, timeoutMs) => {
        await this.requireBridge().request(
          (id) => ({ type: 'openSheetSession', id, sheetIndex, sheetName, ...identity }),
          undefined,
          { timeoutMs },
        );
      },
    });
    return this.worksheetPullClient;
  }

  private runArchiveOperation<T>(operation: () => Promise<T>): Promise<T> {
    const run = async (): Promise<T> => {
      this.assertResourceHealthy();
      try {
        return await operation();
      } catch (error) {
        if (error instanceof OoxmlResourceLimitError) this.resourceFailure ??= error;
        throw error;
      }
    };
    const result = (this.archiveOperationTail ?? Promise.resolve()).then(run, run);
    this.archiveOperationTail = result.then(() => undefined, () => undefined);
    return result;
  }

  /**
   * Fetch an embedded image's bytes by zip path (e.g. `xl/media/image1.png`),
   * wrapped in a Blob of the given MIME. The bytes are pulled through the
   * persistent worker via the `extractImage` message (twin of pptx/docx's
   * `getImage`/`getMedia`); results are cached by path for the lifetime of this
   * instance. The renderer's `fetchImage` option points here so image bytes are
   * extracted lazily rather than inlined as base64 at parse time.
   *
   * Routed through the persistent worker so all WASM `extract_image` decoding
   * stays with the archive owner.
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    this.assertResourceHealthy();
    this.requireArchiveBridge();
    const queued = this.queuedImageLoads?.get(imagePath);
    if (queued) return queued;
    const p = this.runArchiveOperation(() =>
      this.getImageWithinArchiveOperation(imagePath, mimeType));
    this.queuedImageLoads ??= new Map();
    this.queuedImageLoads.set(imagePath, p);
    void p.finally(() => {
      if (this.queuedImageLoads.get(imagePath) === p) this.queuedImageLoads.delete(imagePath);
    }).catch(() => undefined);
    return p;
  }

  private getImageWithinArchiveOperation(imagePath: string, mimeType: string): Promise<Blob> {
    return this.rawParts.get(
      imagePath,
      mimeType,
      () => this.requestImage(imagePath, mimeType),
    );
  }

  private requestImage(imagePath: string, mimeType: string): Promise<Blob> {
    return this.requireArchiveBridge()
      .request((id) => ({ type: 'extractImage', id, path: imagePath }) satisfies WorkerRequest)
      .then((res) => {
        const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      });
  }

  /**
   * Project the workbook to GitHub-flavoured markdown: each sheet becomes a
   * `## SheetName` section followed by a pipe table of its populated bounding
   * box (fully-empty middle rows trimmed, ULP noise masked). Styling, charts,
   * and drawings are discarded — the projection is meant for AI ingestion and
   * full-text search, not layout.
   *
   * Runs entirely in the worker off the archive opened at {@link load} (no
   * re-copy of the file, no re-parse of the model on the main thread), so it
   * works in BOTH `mode: 'main'` and `mode: 'worker'`.
   *
   * @example
   * const wb = await XlsxWorkbook.load(buffer);
   * const md = await wb.toMarkdown();
   */
  async toMarkdown(): Promise<string> {
    this.assertResourceHealthy();
    const res = await this.runArchiveOperation(() => this.requireArchiveBridge().request(
      (id) => ({ type: 'toMarkdown', id }) satisfies WorkerRequest,
    ));
    return (res as Extract<WorkerResponse, { type: 'markdownRendered' }>).markdown;
  }

  /**
   * Resolve a `list`-type data-validation `formula1` (ECMA-376 §18.3.1.32) into
   * the set of allowed values to display, evaluated relative to `sheetIndex`
   * (the sheet that owns the validation, used to resolve unqualified ranges):
   *
   * - Inline quoted list `"A,B,C"`        → the literal values.
   * - Range ref `$B$2:$B$5`               → each non-empty cell's *display
   *   string* (the same formatted text the grid shows, via {@link formatCellValue}),
   *   walked row-major. `Sheet2!$A$1:$A$9` resolves against the named sheet
   *   (lazily parsed via {@link getWorksheet}, hence async).
   * - Named range / complex formula       → `{ kind: 'formula' }` carrying the
   *   raw text so the caller can disclose it rather than blanking it.
   *
   * Read-only: this only reads cell values for display; it never writes.
   */
  async resolveValidationList(
    sheetIndex: number,
    formula1: string | undefined,
  ): Promise<ResolvedList> {
    this.assertResourceHealthy();
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    const parsed = parseListFormula(formula1);
    if (parsed.kind !== 'range') {
      // Inline / unresolved need no cell lookup.
      return resolveListValues(parsed, () => null);
    }

    // Pick the target sheet: the qualifier name (case-insensitive) or, when the
    // range is unqualified, the sheet that owns the validation.
    let targetIndex = sheetIndex;
    if (parsed.sheet) {
      const names = this.sheetNames;
      const found = names.findIndex(
        (n) => n.toLowerCase() === parsed.sheet?.toLowerCase(),
      );
      // Unknown sheet name (e.g. an external reference) → cannot expand;
      // surface the formula instead of silently dropping it.
      if (found < 0) return { kind: 'formula', formula: formula1 ?? '' };
      targetIndex = found;
    }

    const ws = await this.getWorksheet(targetIndex);
    const styles = this.parsedWorkbook.styles;
    // Index the target sheet's cells by "row:col" for O(1) lookup during the
    // row-major walk in resolveListValues.
    const byRC = new Map<string, Cell>();
    for (const r of ws.rows) {
      for (const c of r.cells) byRC.set(`${c.row}:${c.col}`, c);
    }

    return resolveListValues(parsed, (row, col) => {
      const cell = byRC.get(`${row}:${col}`);
      if (!cell) return null;
      return formatCellValue(cell, styles, null, ws.date1904);
    });
  }

  /**
   * IX2 — the display string a cell shows on the grid, i.e. exactly what
   * {@link renderViewport} would draw (number formats, dates, booleans, rich
   * text flattened). Used by {@link XlsxViewer.findText} to search the *rendered*
   * text rather than the raw stored value, so a search matches what the user
   * sees. Threads the workbook styles + the sheet's date system through the
   * shared {@link formatCellValue} (the same call the renderer and
   * validation-list expansion use). Returns `''` before the workbook is loaded.
   */
  cellText(ws: Worksheet, cell: Cell): string {
    if (!this.parsedWorkbook) return '';
    return formatCellValue(cell, this.parsedWorkbook.styles, null, ws.date1904);
  }

  /** Render a sheet viewport into `target`. Image bytes and decoded-image cache
   * ownership stay with this workbook instance. */
  async renderViewport(
    target: HTMLCanvasElement | OffscreenCanvas,
    sheetIndex: number,
    viewport: ViewportRange,
    opts?: XlsxRenderViewportOptions,
  ): Promise<void>;
  async renderViewport(
    target: HTMLCanvasElement | OffscreenCanvas,
    sheetIndex: number,
    viewport: ViewportRange,
    opts: RenderViewportOptions = {},
  ): Promise<void> {
    this.assertResourceHealthy();
    if (this._mode === 'worker') {
      throw new Error(
        "renderViewport(canvas) is unavailable in mode: 'worker'; use renderViewportToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    const styles = this.parsedWorkbook.styles;
    const extracted = extractViewerRenderContext(opts as WireRenderViewportOptions);
    const { sizeOverrides, ...renderOpts } = extracted.opts;
    return this.withWorksheetArchiveOperation(sheetIndex, (source) => {
      const ws = extracted.worksheet ?? createSizeOverriddenWorksheet(source, sizeOverrides);
      if (ws !== source) inheritSheetRenderCache(source, ws);
      if (extracted.layoutMetrics) {
        GridGeometry.forWorksheet(ws, extracted.layoutMetrics.maximumDigitWidth);
      }
      return renderWorksheetViewport(
        {
          ws,
          styles,
          math: this.math,
          threeD: this.threeD,
          regionMap: this.regionMap,
          chartEx: this.chartEx,
          tiff: this.tiff,
        },
        target,
        viewport,
        // The stable closure uses the archive operation already reserved by
        // withWorksheetArchiveOperation, avoiding a nested FIFO acquisition.
        { ...renderOpts, fetchImage: this._fetchImage },
      );
    });
  }

  /**
   * Render a sheet viewport and return it as an ImageBitmap (both modes; in
   * worker mode the render runs entirely off the main thread). `opts.width` /
   * `opts.height` are required: there is no DOM element to measure in a worker
   * or on an OffscreenCanvas. Paint with
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   */
  async renderViewportToBitmap(
    sheetIndex: number,
    viewport: ViewportRange,
    opts: RenderViewportToBitmapOptions,
  ): Promise<ImageBitmap>;
  async renderViewportToBitmap(
    sheetIndex: number,
    viewport: ViewportRange,
    opts: WireRenderViewportOptions & { width: number; height: number },
  ): Promise<ImageBitmap> {
    this.assertResourceHealthy();
    const extracted = extractViewerRenderContext(opts);
    const wireOpts = { ...extracted.opts, dpr: opts.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= this.sheetCount) {
        throw new Error(`Sheet index ${sheetIndex} out of range (count: ${this.sheetCount})`);
      }
      const res = await this.withWorksheetArchiveOperation(sheetIndex, () =>
        this.requireBridge().request(
          (id) => ({
            type: 'renderViewport',
            id,
            sheetIndex,
            viewport,
            opts: wireOpts,
            layoutMetrics: extracted.layoutMetrics,
            viewProjection: extracted.projection,
          }) satisfies RenderWorkerRequest,
        ));
      return (res as Extract<RenderWorkerResponse, { type: 'viewportRendered' }>).bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderViewport(off, sheetIndex, viewport, wireOpts);
    return off.transferToImageBitmap();
  }

  /** @internal Drop projections owned by a destroyed viewer. */
  [releaseXlsxViewerProjection](projectionId: number): void {
    if (this._mode !== 'worker') return;
    this.requireBridge().post(
      { type: 'releaseViewProjection', projectionId } satisfies RenderWorkerRequest,
    );
  }

  private withWorksheetArchiveOperation<T>(
    sheetIndex: number,
    operation: (worksheet: Worksheet) => Promise<T>,
  ): Promise<T> {
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) return this.runArchiveOperation(() => operation(cached));
    const active = this.sheetLoads.get(sheetIndex);
    if (active) {
      return this.runArchiveOperation(async () => operation(await active));
    }
    if (!this.parsedWorkbook) {
      return Promise.reject(new Error('Workbook not loaded'));
    }
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) return Promise.reject(new Error(`Sheet index ${sheetIndex} out of range`));

    let resolveLoad!: (worksheet: Worksheet) => void;
    let rejectLoad!: (error: unknown) => void;
    const load = new Promise<Worksheet>((resolve, reject) => {
      resolveLoad = resolve;
      rejectLoad = reject;
    });
    // The composite render promise is the primary caller. A concurrent
    // getWorksheet may also observe `load`, but a lone failed render must not
    // leave this coordination promise as an unhandled rejection.
    void load.catch(() => undefined);
    this.sheetLoads.set(sheetIndex, load);
    const combined = this.runArchiveOperation(async () => {
      try {
        const worksheet = await this.loadWorksheetStream(sheetIndex, sheetMeta.name);
        resolveLoad(worksheet);
        return await operation(worksheet);
      } catch (error) {
        rejectLoad(error);
        throw error;
      } finally {
        if (this.sheetLoads.get(sheetIndex) === load) this.sheetLoads.delete(sheetIndex);
      }
    });
    return combined;
  }

  destroy(): void {
    this.generation = (this.generation ?? 1) + 1;
    void this.worksheetPullClient?.cancelAll('closed').catch(() => undefined);
    this.worksheetPullClient = null;
    this.bridge?.terminate();
    this.bridge = null;
    this.parsedWorkbook = null;
    this.sheetCache.clear();
    this.sheetLoads.clear();
    this.fontsDestroyed = true;
    for (const retained of this.retainedFontSets.values()) {
      if (retained.faces) unloadGoogleFonts(retained.faces);
      // An in-flight registration observes fontsDestroyed in its own completion
      // callback and releases exactly once when the faces become available.
    }
    this.retainedFontSets.clear();
    this.googleFontNames = [];
    // Frame-local lookup maps never escape the renderer; drop the owning core
    // caches to release decoded surfaces and SVG references.
    dropDecodedBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
    this.rawParts.clear();
    this.queuedImageLoads?.clear();
  }

  private assertResourceHealthy(): void {
    if (this.resourceFailure) throw this.resourceFailure;
  }

  private requireBridge(): WorkbookBridge {
    if (!this.bridge) {
      throw new Error('This operation requires an active workbook worker');
    }
    return this.bridge;
  }

  private requireArchiveBridge(): WorkbookBridge {
    if (this.delimitedTextBacked || !this.bridge) {
      throw new Error('This operation requires an active archive-backed workbook');
    }
    return this.bridge;
  }
}
