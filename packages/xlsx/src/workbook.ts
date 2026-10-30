import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  normalizeXlsxWorksheetPolicy,
  type NormalizedXlsxWorksheetPolicy,
} from './internal/worksheet-policy.js';
import { bindCellFontLanguage, snapshotJapaneseCellFonts } from './internal/cell-font-language.js';
import { resolveCjkFallback, type CjkLang } from '@silurus/ooxml-core';
import { xlsxCjkFallback } from './font-requests.js';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/xlsx_parser_bg.wasm?url';
import {
  FontResolverBridgeHost,
  normalizeFontOptions,
  type FontSessionDemand,
  type FontResolverFailures,
  type FontSource,
  type FontSourceSession,
  type InstalledFamilyExclusion,
} from '@silurus/ooxml-core/internal/font-sources';
import {
  loadOfficeFontFallbacks,
  unloadOfficeFontFallbacks,
  type LoadedOfficeFontFallbacks,
  type OfficeFontFallbackRequest,
  WorkerBridge,
  defaultDpr,
  isHTMLCanvas,
  dropDecodedBitmapCache,
  dropSvgImageCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type ModelSource,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type TiffRenderer,
  OoxmlResourceLimitError,
  type OoxmlResourceMetrics,
  workerRendererDescriptors,
} from '@silurus/ooxml-core';
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
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, XlsxRenderViewportOptions, WorkerRequest, WorkerResponse, Cell, SheetVisibility, XlsxComment, Row, XlsxLoadedFontSources } from './types.js';
import { selectSheetVisibility } from './sheet-visibility.js';
import { renderWorksheetViewport, xlsxRenderCommitGuardOf } from './render-orchestrator.js';
import { xlsxOfficeFontRequests, xlsxWorksheetOfficeFontRequests } from './font-requests.js';
import { officeRequestKey } from './shape-office-line.js';
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
import { applyAutoRowHeights, bindXlsxOfficeFontRoutes, bindXlsxWorksheetOfficeFontRoutes, inheritSheetRenderCache, getGridGeometryForWorksheet } from './renderer.js';
import {
  assertDelimitedTextSourceBytes,
  resolveDelimitedTextOptions,
  type ResolvedDelimitedTextOptions,
  type XlsxSheetLoadOptions,
} from './delimited-text.js';
import { readDelimitedTextResponse } from './delimited-text-source.js';
import { WorksheetPreview } from './internal/worksheet-preview.js';
import { setWorksheetPreviewBounds } from './internal/worksheet-content-bounds.js';
import { bindWorksheetPolicy } from './worksheet-policy-context.js';
import {
  clampPartFailureMessage,
  prepareErrorPlaceholder,
  resolveErrorPlaceholderFor,
  wireErrorPlaceholder,
  type PartFailureLog,
  type PreparedErrorPlaceholder,
  type WireErrorPlaceholder,
} from '@silurus/ooxml-core/internal/part-failures';
import {
  xlsxPartFailureLog,
  type XlsxPartFailure,
  type XlsxSheetPartFailure,
} from './part-failures.js';
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
  loaded: { office: LoadedOfficeFontFallbacks } | null;
  readonly loading: Promise<{ office: LoadedOfficeFontFallbacks }>;
  /** This workbook's font-source session in this set (`LoadOptions.fonts`);
   * null without sources. */
  session: FontSourceSession | null;
  /** Serialises this set's demand deltas in workbook order. */
  fontQueue: Promise<unknown>;
}

/** Options for {@link XlsxWorkbook.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`fonts`, `resourceLimits`, the
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
  /**
   * Opt-in model source for workbook input that is not an OOXML package (a
   * `ModelSource<'xlsx'>`). When its `claim()` accepts the raw bytes, it
   * opens them in the parser realm through its own module; otherwise the
   * ordinary OOXML path runs unchanged. A source for another format is a
   * `TypeError`. Omitting this option imports and runs nothing extra.
   */
  legacyXls?: ModelSource<'xlsx'>;
}

type WorkbookBridge = WorkerBridge<
  WorkerResponse | RenderWorkerResponse | DelimitedTextParseResponse |
  PullSessionResponse<ArrayBuffer, number>
>;

export class XlsxWorkbook {
  private metrics: OoxmlResourceMetricsSession | null = null;
  private bridge: WorkbookBridge | null = null;
  private delimitedTextBacked = false;
  private readonly japaneseCellFonts = snapshotJapaneseCellFonts();
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  /** Cache insertion order is the workbook session's LRU order. A lease pins a
   * model while an operation or a viewer is using its content graph. */
  private sheetCacheUsage = new Map<number, WorksheetCacheUsage>();
  private sheetLeases = new Map<number, number>();
  private evictingSheets = new Map<number, { done: Promise<void>; complete: () => void }>();
  /** One materialization per sheet at a time. This becomes the ownership seam
   * for the bounded worksheet cursor: concurrent callers share one cursor and
   * one eventual mutable compatibility object instead of doubling peak work. */
  private sheetLoads = new Map<number, Promise<Worksheet>>();
  private sheetPreviews = new Map<number, WorksheetPreview>();
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
  /** Normalized per-session worksheet admission policy. */
  private worksheetPolicy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY;
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderViewport` call reuses it — equations in shapes render when present,
   *  and are skipped when omitted. */
  private cjkFallback: CjkLang = 'jp';
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
  /** Validated `LoadOptions.fonts` (normalised from the deprecated
   *  `useGoogleFonts`); undefined without sources. */
  private fonts: readonly FontSource[] | undefined;
  /** The workbook's incremental font-source demand (main realm), and the
   *  installed-family probe every delta passes before it registers. */
  private fontDemand: import('./font-source-load.js').XlsxFontDemand | null = null;
  private installedFamilies: InstalledFamilyExclusion | null = null;
  private fontSessionType: typeof FontSourceSession | null = null;
  private fontAcquirerFor: typeof import('./font-source-load.js').inRealmResolverFaceAcquirer | null = null;
  /** Every demand delta so far, replayed into each newly retained set. */
  private fontDeltas: Promise<FontSessionDemand>[] = [];
  private officeFontRequests: OfficeFontFallbackRequest[] = [];
  private readonly retainedFontSets = new Map<FontFaceSet, RetainedFontSet>();
  /** Worker mode: runs this workbook's `fontResolver` functions for the
   *  worker's font-source session (core `resolver-bridge-host.ts`); disposed
   *  in {@link destroy}, which aborts every in-flight resolver call. */
  private fontResolverBridge: FontResolverBridgeHost | null = null;
  /** Resolver failures of this workbook, shared by its viewer-realm sessions
   * and its worker bridge (core `resolver-resolution.ts`). */
  private fontResolverFailures: FontResolverFailures = new Map();
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
        if (this.fontResolverBridge?.accept(res)) return;
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
    opts = normalizeFontOptions(opts);
    const worksheetPolicy = normalizeXlsxWorksheetPolicy(opts);
    opts = { ...opts, xlsxWorksheetLimits: worksheetPolicy.worksheet };
    opts = { ...opts, cjkFallback: resolveCjkFallback(opts.cjkFallback) };
    const delimited = resolveDelimitedTextOptions(sourceOptions);
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode,
      policy: resourceOptions.policy,
      formatLimits: { xlsxWorksheetLimits: { ...worksheetPolicy.worksheet } },
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

  /** Parse an XLSX from a URL or ArrayBuffer. Cell-font theme selection uses a
   * snapshot of the viewer language; re-open to apply a language change. */
  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<XlsxWorkbook> {
    // Before anything else reads the options: `useGoogleFonts` becomes
    // `fonts: [googleFonts()]` (deprecated alias, warned once per realm) and
    // `fonts` is validated. Nothing below reads `useGoogleFonts`.
    opts = normalizeFontOptions(opts);
    const worksheetPolicy = normalizeXlsxWorksheetPolicy(opts);
    opts = { ...opts, xlsxWorksheetLimits: worksheetPolicy.worksheet };
    if (__OOXML_MODEL_SOURCES__ && opts.legacyXls !== undefined) {
      return (await import('./internal/workbook-model-source.js')).loadXlsxModelSource(source, opts, XlsxWorkbook);
    }
    opts = { ...opts, cjkFallback: resolveCjkFallback(opts.cjkFallback) };
    const resourceOptions = normalizeLoadResourceOptions(opts);
    const mode = opts.mode ?? 'main';
    const metrics = new OoxmlResourceMetricsSession({
      enabled: true,
      format: 'xlsx',
      mode,
      policy: resourceOptions.policy,
      formatLimits: { xlsxWorksheetLimits: { ...worksheetPolicy.worksheet } },
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
    buffer = toArrayBuffer(await resolveOoxmlContainer(buffer, opts.password));
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
    const worksheetPolicy = normalizeXlsxWorksheetPolicy(opts);
    const bridge = this.requireBridge();
    this.worksheetPolicy = worksheetPolicy;
    this.resourceFailure = null;
    this.retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
    this.sheetCache.clear();
    this.sheetCacheUsage.clear();
    this.sheetLeases.clear();
    this.completePendingEvictions();
    await this.worksheetPullClient?.cancelAll('closed');
    this.worksheetPullClient = null;
    this.generation = (this.generation ?? 0) + 1;
    this.workerTimeoutMs = opts.workerTimeoutMs;
    this.cjkFallback = resolveCjkFallback(opts.cjkFallback);
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
    // In worker mode the worker loads fonts before its first render
    // (rendering measures text), so the font sources are forwarded; the main
    // realm still loads them below for viewer geometry.
    // Preserve XLSX's historical caller-owned ArrayBuffer contract only when
    // the resolved ZIP is literally the caller's buffer. URL and decrypted
    // buffers are library-owned and can transfer directly without a peak copy.
    const workerData = preserveCallerBuffer ? data.slice(0) : data;
    this.attachFontResolverBridge(bridge, opts.fonts);
    const parsed = await bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({
              type: 'parse',
              id,
              data: workerData,
              resourcePolicy,
              worksheetPolicy,
              ...(opts.fonts ? { fonts: opts.fonts } : {}),
              cjkFallback: this.cjkFallback,
              renderers: rendererDescriptors,
              japaneseCellFonts: this.japaneseCellFonts,
            } satisfies RenderWorkerRequest)
          : ({
              type: 'parse',
              id,
              data: workerData,
              resourcePolicy,
              worksheetPolicy,
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
    this.cjkFallback = xlsxCjkFallback(parsedWorkbook, this.cjkFallback);
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
      this._partFailureLog.add({ kind: 'workbook', message: workbookError });
    }
    this.officeFontRequests = xlsxOfficeFontRequests(parsedWorkbook);
    // The composite viewer computes hit/scroll/overlay geometry on the main
    // realm even when paint runs in a worker. Register the same font-source
    // faces in both realms before any worksheet geometry snapshot is made so
    // ECMA-376 MDW is identical across paint and interaction.
    await this.prepareFontSources(parsedWorkbook, opts.fonts);
    if (typeof document !== 'undefined' && document.fonts) {
      await this.retainFontsInSet(document.fonts);
    }
  }

  private async _loadDelimitedText(
    data: ArrayBuffer,
    opts: LoadOptions,
    options: ResolvedDelimitedTextOptions,
  ): Promise<void> {
    const worksheetPolicy = normalizeXlsxWorksheetPolicy(opts);
    const bridge = this.requireBridge();
    this.worksheetPolicy = worksheetPolicy;
    this.delimitedTextBacked = true;
    this.workerTimeoutMs = opts.workerTimeoutMs;
    this.cjkFallback = resolveCjkFallback(opts.cjkFallback);
    this.generation++;
    this.math = this._mode === 'worker' ? undefined : opts.math;
    this.threeD = this._mode === 'worker' ? undefined : opts.threeD;
    this.regionMap = this._mode === 'worker' ? undefined : opts.regionMap;
    this.chartEx = this._mode === 'worker' ? undefined : opts.chartEx;
    this.tiff = this._mode === 'worker' ? undefined : opts.tiff;
    const rendererDescriptors = this._mode === 'worker'
      ? workerRendererDescriptors(opts)
      : undefined;
    this.attachFontResolverBridge(bridge, opts.fonts);
    const response = await bridge.request(
      (id) => ({
        type: 'parseDelimitedText',
        japaneseCellFonts: this.japaneseCellFonts,
        id,
        data,
        options,
        worksheetPolicy,
        // Only the render worker paints; the parse-only worker needs no fonts.
        ...(this._mode === 'worker' && opts.fonts ? { fonts: opts.fonts } : {}),
        cjkFallback: this.cjkFallback,
        renderers: rendererDescriptors,
      } satisfies DelimitedTextParseRequest),
      [data],
      { timeoutMs: opts.workerTimeoutMs },
    ) as Extract<DelimitedTextParseResponse, { type: 'delimitedTextParsed' }>;
    const worksheet = JSON.parse(
      new TextDecoder().decode(response.worksheetJson),
    ) as Worksheet;
    bindWorksheetPolicy(worksheet, worksheetPolicy);
    bindCellFontLanguage(worksheet, this.japaneseCellFonts);
    const sheets = response.workbook.workbook.sheets;
    if (sheets.length !== 1 || sheets[0]?.name !== worksheet.name) {
      throw new Error('Delimited text worker returned inconsistent worksheet metadata');
    }
    const measured = measureWorksheet(worksheet, worksheetPolicy);
    assertWorksheetModelUsage(measured, 'load-delimited-text', undefined, undefined, worksheetPolicy);
    assertWorksheetJsonBytes(measured.jsonBytes, 'load-delimited-text', undefined, undefined, worksheetPolicy);
    assertWorksheetCacheUsage(measured, 'load-delimited-text', undefined, undefined, worksheetPolicy);
    this.parsedWorkbook = response.workbook;
    this.cjkFallback = xlsxCjkFallback(response.workbook, this.cjkFallback);
    this.sheetCache.set(0, worksheet);
    this.sheetCacheUsage.set(0, measured);
    this.retainedSheetUsage = measured;

    this.officeFontRequests = xlsxOfficeFontRequests(response.workbook);
    await this.prepareFontSources(response.workbook, opts.fonts);
    if (typeof document !== 'undefined' && document.fonts) {
      await this.retainFontsInSet(document.fonts);
      await this.retainWorksheetOfficeFonts(worksheet);
      await this.demandWorksheetFonts(worksheet, worksheet.rows);
      const office = this.retainedFontSets.get(document.fonts)?.loaded?.office;
      bindXlsxWorksheetOfficeFontRoutes(worksheet, office?.routes, this.loadedFontSources(document.fonts));
    }
  }

  /** Worker mode: answer the render worker's `fontResolver` calls for the
   * document about to be parsed (core `resolver-bridge-host.ts`). */
  private attachFontResolverBridge(bridge: WorkbookBridge, fonts: readonly FontSource[] | undefined): void {
    this.fontResolverBridge?.dispose();
    this.fontResolverFailures = new Map();
    this.fontResolverBridge = this._mode === 'worker'
      ? new FontResolverBridgeHost({
        post: (message, transfer) => bridge.post(message, transfer),
        sources: fonts,
        // One outcome per family for the whole workbook: the viewer-realm
        // sessions and the worker share this failure record (and main's
        // resolution cache), so paint and hit geometry use the same faces.
        resolverFailures: this.fontResolverFailures,
      })
      : null;
  }

  /**
   * Font sources (`LoadOptions.fonts`): the workbook's demand, computed once
   * in this realm from the parsed model. The session, the CSS adapters and the
   * demand collector load only when sources are configured; the workbook
   * delta is replayed into every set this workbook retains.
   */
  private async prepareFontSources(workbook: ParsedWorkbook, fonts: readonly FontSource[] | undefined): Promise<void> {
    this.fonts = fonts;
    if (!fonts?.length) return;
    const fontSourceLoad = await import('./font-source-load.js');
    this.fontSessionType = fontSourceLoad.FontSourceSession;
    this.fontAcquirerFor = fontSourceLoad.inRealmResolverFaceAcquirer;
    this.fontDemand = new fontSourceLoad.XlsxFontDemand(workbook, this.cjkFallback, fontSourceLoad.declaredFontSubstitutions(fonts));
    // Installed authored faces are probed before any source face of a delta
    // registers, and never demanded (core installed-families.ts).
    this.installedFamilies = new fontSourceLoad.InstalledFamilyExclusion();
    await this.pushFontDelta(this.fontDemand.take());
  }

  /** Queue one (installed-excluded) delta on a retained set's session. */
  private enqueueFontDelta(retained: RetainedFontSet, delta: Promise<FontSessionDemand>): Promise<void> {
    const task = retained.fontQueue.then(async () => {
      const demand = await delta;
      if (retained.session && !this.fontsDestroyed) await retained.session.ensure(demand);
    });
    retained.fontQueue = task.catch(() => undefined);
    return task;
  }

  /** Record a demand delta and ensure it in every retained set. A ceiling
   * breach (`font-source-limit`) rejects; a failed face only warns. */
  private async pushFontDelta(delta: FontSessionDemand | null): Promise<void> {
    if (!delta || !this.installedFamilies || this.fontsDestroyed) return;
    const excluded = this.installedFamilies.exclude(delta);
    this.fontDeltas.push(excluded);
    await Promise.all([...this.retainedFontSets.values()].map((retained) => this.enqueueFontDelta(retained, excluded)));
  }

  /** The worksheet's font-source delta: its non-cell paint once per model,
   * plus `rows` (cells about to be measured or painted). */
  private async demandWorksheetFonts(worksheet: Worksheet | null, rows?: readonly Row[]): Promise<void> {
    const demand = this.fontDemand;
    if (!demand) return;
    if (worksheet) demand.worksheetParts(worksheet);
    if (rows?.length) demand.rows(rows, worksheet ?? undefined);
    await this.pushFontDelta(demand.take());
  }

  /** What this workbook's font sources loaded in `fontSet`, for the renderer. */
  private loadedFontSources(fontSet: FontFaceSet | null | undefined): XlsxLoadedFontSources | undefined {
    return fontSet ? this.retainedFontSets.get(fontSet)?.session?.loaded : undefined;
  }

  private async retainFontsInSet(fontSet: FontFaceSet): Promise<() => void> {
    if (this.fontsDestroyed) return () => undefined;
    let retained = this.retainedFontSets.get(fontSet);
    if (retained) {
      retained.refs++;
    } else {
      const session = this.fonts?.length && this.fontSessionType
        ? new this.fontSessionType({
          sources: this.fonts, fontSet, resolverFailures: this.fontResolverFailures,
          // Worker mode: the page acquires resolver URL faces for the worker
          // (core resolver-bridge-host.ts); this realm acquires through the
          // same cache and failure record, so both realms paint the same bytes.
          ...(this._mode === 'worker' && this.fontAcquirerFor
            ? { acquireResolverFaces: this.fontAcquirerFor(this.fontResolverFailures) } : {}),
        })
        : null;
      const record = {
        refs: 1, loaded: null, session, fontQueue: Promise.resolve() as Promise<unknown>,
      } as Omit<RetainedFontSet, 'loading'> as RetainedFontSet;
      // Replay every delta so far, in order, into this set.
      const sourcesReady = Promise.all(this.fontDeltas.map((delta) => this.enqueueFontDelta(record, delta)));
      // Settle both loads: a rejected source delta (font-source-limit) must
      // not drop the parallel Office-face result, whose faces would then never
      // be released. On any failure every face this set took is returned.
      const loading = Promise.allSettled([
        sourcesReady,
        loadOfficeFontFallbacks(this.officeFontRequests, fontSet),
      ]).then(([sources, office]) => {
        if (sources.status === 'fulfilled' && office.status === 'fulfilled') return { office: office.value };
        if (office.status === 'fulfilled') unloadOfficeFontFallbacks(office.value.faces);
        record.session?.release();
        throw sources.status === 'rejected' ? sources.reason : (office as PromiseRejectedResult).reason;
      });
      retained = Object.assign(record, { loading });
      this.retainedFontSets.set(fontSet, retained);
      loading.then((loaded) => {
        retained!.loaded = loaded;
        if (this.fontsDestroyed) {
          unloadOfficeFontFallbacks(loaded.office.faces);
          retained!.session?.release();
        }
      }, () => {
        // The awaiting caller observes the rejection (a font-source limit);
        // the session latched it and is released with the set.
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
      current.session?.release();
      if (current.loaded) {
        unloadOfficeFontFallbacks(current.loaded.office.faces);
      } else void current.loading.then((loaded) => {
        unloadOfficeFontFallbacks(loaded.office.faces);
      }, () => undefined);
    };
  }

  private async retainWorksheetOfficeFonts(worksheet: Worksheet): Promise<void> {
    const additional = xlsxWorksheetOfficeFontRequests(worksheet).filter((request) =>
      !this.officeFontRequests.some((known) => officeRequestKey(known) === officeRequestKey(request)));
    if (additional.length === 0) return;
    this.officeFontRequests.push(...additional);
    // Even an initially empty font registry is retained at bootstrap and by
    // each viewer. Worksheet-local shape tuples can then extend every live
    // document's registry without inventing a new lifetime or missing popups.
    await Promise.all([...this.retainedFontSets].map(async ([set, retained]) => {
      const current = await retained.loading;
      const office = await loadOfficeFontFallbacks(additional, set);
      // A popup can release its FontFaceSet while this extra sheet tuple is
      // loading. The retained record then leaves the map; attaching the late
      // face to that orphan would leak its registry reference forever.
      if (this.fontsDestroyed || this.retainedFontSets.get(set) !== retained || retained.refs <= 0) {
        unloadOfficeFontFallbacks(office.faces);
      } else {
        current.office.faces.push(...office.faces);
        const completed = new Set(current.office.checked);
        for (const key of office.checked) {
          if (completed.has(key)) continue;
          completed.add(key);
          current.office.checked.push(key);
        }
        Object.assign(current.office.routes, office.routes);
      }
    }));
  }

  /** @internal Retain required faces in the document that owns a viewer canvas. */
  async [retainXlsxViewerFonts](targetDocument: Document): Promise<() => void> {
    return await this.retainFontsInSet(targetDocument.fonts);
  }

  /** @internal Fonts are retained before this hook runs, so Canvas text
   * measurement observes the same faces that the subsequent paint uses. */
  [prepareXlsxViewerRowHeights](worksheet: Worksheet, ctx: CanvasRenderingContext2D): void {
    if (!this.parsedWorkbook) return;
    const set = isHTMLCanvas(ctx.canvas)
      ? ctx.canvas.ownerDocument.fonts : null;
    const routes = set ? this.retainedFontSets.get(set)?.loaded?.office.routes : undefined;
    bindXlsxOfficeFontRoutes(ctx, worksheet, routes, this.loadedFontSources(set));
    getGridGeometryForWorksheet(worksheet);
    applyAutoRowHeights(ctx, worksheet, this.parsedWorkbook.styles, this.cjkFallback);
  }

  get sheetNames(): string[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.name) ?? [];
  }

  /**
   * Package parts that failed to parse (#1974): a `workbook` entry for a
   * broken shared part (known at load) first, then one `sheet` entry per
   * worksheet/chartsheet/dialogsheet part, ascending by index. Sheets are
   * parsed lazily, so a sheet is listed once its model has been loaded (a
   * viewer shows it, or {@link getWorksheet} is called). Never routed to
   * `onError`.
   */
  get partFailures(): readonly XlsxPartFailure[] {
    return this._partFailureLog.list;
  }

  private get _partFailureLog(): PartFailureLog<XlsxPartFailure> {
    return xlsxPartFailureLog(this);
  }

  /** Main mode: resolve from the worksheet actually painted. */
  private sheetErrorPlaceholderText(
    prepared: PreparedErrorPlaceholder<XlsxSheetPartFailure>,
    sheetIndex: number,
    worksheet: Worksheet,
  ): WireErrorPlaceholder | undefined {
    if (prepared.kind === 'none' || !worksheet.parseError) return undefined;
    const recorded = this._partFailureLog.get(sheetIndex);
    const failure: XlsxSheetPartFailure = recorded?.kind === 'sheet' ? recorded : Object.freeze({
      kind: 'sheet',
      sheetIndex,
      sheetName: worksheet.name,
      message: clampPartFailureMessage(worksheet.parseError),
    });
    const resolved = resolveErrorPlaceholderFor(prepared, failure);
    return resolved === null ? undefined : wireErrorPlaceholder(resolved);
  }

  /**
   * Worker mode: main may not know yet whether the worker's model of this
   * sheet is degraded (a provisional preview here, a terminal there). Then a
   * function option is sent as `resolve-on-main`, and the worker asks back
   * once instead of painting. `final` (the retry) never asks again.
   */
  private workerErrorPlaceholderText(
    prepared: PreparedErrorPlaceholder<XlsxSheetPartFailure>,
    sheetIndex: number,
    final: boolean,
  ): WireErrorPlaceholder | undefined {
    if (prepared.kind === 'none') return undefined;
    if (prepared.kind === 'static') return prepared.text === null ? undefined : wireErrorPlaceholder(prepared.text);
    const recorded = this._partFailureLog.get(sheetIndex);
    if (recorded?.kind === 'sheet') {
      const resolved = resolveErrorPlaceholderFor(prepared, recorded);
      return wireErrorPlaceholder(resolved);
    }
    if (final || this.sheetCache.has(sheetIndex)) return undefined;
    return { state: 'resolve-on-main' };
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
    const pending = this.evictingSheets.get(sheetIndex);
    if (pending) {
      await pending.done;
      return this.getWorksheet(sheetIndex);
    }
    const release = this.pinWorksheet(sheetIndex);
    try {
      return await this.getOrLoadWorksheet(sheetIndex);
    } finally {
      release();
    }
  }

  private async getOrLoadWorksheet(sheetIndex: number): Promise<Worksheet> {
    this.assertResourceHealthy();
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) {
      this.touchWorksheet(sheetIndex, cached);
      return cached;
    }
    const active = this.sheetLoads.get(sheetIndex);
    if (active) return active;
    const progress = new WorksheetPreview([]);
    (this.sheetPreviews ??= new Map()).set(sheetIndex, progress);
    const load = this.loadWorksheet(sheetIndex, progress);
    void load.catch((error: unknown) => progress.fail(error));
    this.sheetLoads.set(sheetIndex, load);
    try {
      return await load;
    } finally {
      if (this.sheetLoads.get(sheetIndex) === load) this.sheetLoads.delete(sheetIndex);
      if (this.sheetPreviews?.get(sheetIndex) === progress) this.sheetPreviews.delete(sheetIndex);
    }
  }

  private touchWorksheet(sheetIndex: number, worksheet: Worksheet): void {
    this.sheetCache.delete(sheetIndex);
    this.sheetCache.set(sheetIndex, worksheet);
  }

  private pinWorksheet(sheetIndex: number): () => void {
    this.sheetLeases.set(sheetIndex, (this.sheetLeases.get(sheetIndex) ?? 0) + 1);
    let released = false;
    return () => {
      if (released) return;
      released = true;
      const count = this.sheetLeases.get(sheetIndex) ?? 0;
      if (count <= 1) this.sheetLeases.delete(sheetIndex);
      else this.sheetLeases.set(sheetIndex, count - 1);
    };
  }

  private completePendingEvictions(): void {
    for (const pending of this.evictingSheets.values()) pending.complete();
    this.evictingSheets.clear();
  }

  /** The lease begins before an asynchronous pull so another admission cannot
   * discard a sheet just as the waiting caller receives it. */
  private async acquireWorksheetLease(sheetIndex: number): Promise<{
    worksheet: Worksheet;
    release: () => void;
  }> {
    const pending = this.evictingSheets.get(sheetIndex);
    if (pending) {
      await pending.done;
      return this.acquireWorksheetLease(sheetIndex);
    }
    const release = this.pinWorksheet(sheetIndex);
    try {
      return { worksheet: await this.getWorksheet(sheetIndex), release };
    } catch (error) {
      release();
      throw error;
    }
  }

  /** Viewer lease may resolve at the first exact row-local viewport. Public
   * worksheet and find callers continue to wait for terminal cache admission.
   * Row pulling is independent of painting: a render may itself need a later
   * row, so waiting for the paint here would create a mutual wait. */
  private async acquireWorksheetPreviewLease(sheetIndex: number): Promise<{
    worksheet: Worksheet;
    release: () => void;
    partial: boolean;
    completion: Promise<Worksheet>;
    waitForRows: (row: number) => Promise<void>;
  }> {
    const pending = this.evictingSheets.get(sheetIndex);
    if (pending) {
      await pending.done;
      return this.acquireWorksheetPreviewLease(sheetIndex);
    }
    const releasePin = this.pinWorksheet(sheetIndex);
    const release = () => { releasePin(); };
    try {
      const completion = this.getOrLoadWorksheet(sheetIndex);
      // Preview readiness can reject before this lease exposes completion (for
      // example, on destroy). Observe both promises immediately; callers still
      // receive failures through ready or the original completion promise.
      void completion.catch(() => undefined);
      const progress = this.sheetPreviews?.get(sheetIndex);
      const worksheet = progress ? await progress.ready : await completion;
      return {
        worksheet, release, partial: !!progress?.worksheet && !progress.complete, completion,
        waitForRows: (row) => progress ? progress.waitFor(row) : Promise.resolve(),
      };
    } catch (error) {
      release();
      throw error;
    }
  }

  /** Detached comments for one worksheet, in authored order. Worksheet models
   * are materialized lazily, so this accessor is asynchronous. */
  async getComments(sheetIndex: number): Promise<readonly Readonly<XlsxComment>[]> {
    const lease = await this.acquireWorksheetLease(sheetIndex);
    try {
      return structuredClone(lease.worksheet.comments ?? []);
    } finally {
      lease.release();
    }
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

  private async loadWorksheet(sheetIndex: number, progress: WorksheetPreview): Promise<Worksheet> {
    // The worker retained its transferred archive at parse time; loaded state
    // is represented by the workbook bootstrap, not a duplicate source buffer.
    if (!this.parsedWorkbook) {
      throw new Error('Workbook not loaded');
    }
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);

    return this.runArchiveOperation(() => this.loadWorksheetStream(sheetIndex, sheetMeta.name, progress));
  }

  private async loadWorksheetStream(sheetIndex: number, sheetName: string, progress: WorksheetPreview): Promise<Worksheet> {
    const policy = this.worksheetPolicy;
    const client = this.ensureWorksheetPullClient();
    const rows = progress.rows;
    let modelUsage: WorksheetModelUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0 };
    let terminal: Worksheet | undefined;
    let nextCacheUsage: WorksheetCacheUsage | undefined;
    let terminalUsage: WorksheetCacheUsage | undefined;
    let terminalAcknowledged = false;
    let rowsWithoutColumnStyles = false;
    // The compatibility adapter knows the workbook sheet index/name but not
    // the resolved OPC relationship target. Omit `part` rather than fabricate
    // a package address; Rust-originated violations carry the real xl/... path.
    const part = undefined;
    try {
      for await (const unit of client.stream(sheetIndex, sheetName)) {
        if (unit.kind === 'preview') {
          if (unit.worksheet) {
            bindWorksheetPolicy(unit.worksheet, policy);
            bindCellFontLanguage(unit.worksheet, this.japaneseCellFonts);
          }
          progress.preview(unit.worksheet, unit.reason, unit.maxRow, unit.maxCol);
          if (unit.worksheet) {
            setWorksheetPreviewBounds(unit.worksheet, { maxRow: unit.maxRow, maxCol: unit.maxCol });
            await this.retainWorksheetOfficeFonts(unit.worksheet);
            await this.demandWorksheetFonts(unit.worksheet, unit.worksheet.rows);
            const mainFonts = typeof document !== 'undefined' ? document.fonts : undefined;
            const mainOffice = mainFonts ? this.retainedFontSets.get(mainFonts)?.loaded?.office : undefined;
            bindXlsxWorksheetOfficeFontRoutes(unit.worksheet, mainOffice?.routes, this.loadedFontSources(mainFonts));
          }
          continue;
        }
        if (unit.kind === 'rows') {
          const nextUsage = addWorksheetUsage(modelUsage, measureRows(unit.rows, policy), policy);
          assertWorksheetModelUsage(
            nextUsage,
            'get-worksheet',
            part,
            unit.usage,
            policy,
          );
          // Faces for these cells load before any of them is measured. Without
          // a preview model the column styles are unknown yet: the terminal
          // model charges these rows again with them.
          if (!progress.worksheet) rowsWithoutColumnStyles = true;
          await this.demandWorksheetFonts(progress.worksheet, unit.rows);
          progress.append(unit.rows);
          modelUsage = nextUsage;
          continue;
        }
        const worksheet = unit.worksheet;
        bindWorksheetPolicy(worksheet, policy);
        bindCellFontLanguage(worksheet, this.japaneseCellFonts);
        worksheet.rows = worksheet.parseError ? [] : rows;
        // Terminal metadata contains no rows, but measure the final public model
        // to cover exact monolithic JSON before its cache admission.
        const measured = completeWorksheetUsage(
          worksheet,
          worksheet.parseError
            ? { rows: 0, cells: 0, ownedUtf8Bytes: 0 }
            : modelUsage,
          policy,
        );
        assertWorksheetModelUsage(
          measured,
          'get-worksheet',
          part,
          unit.usage,
          policy,
        );
        assertWorksheetJsonBytes(
          measured.jsonBytes,
          'get-worksheet',
          part,
          unit.usage,
          policy,
        );
        // Resource governance is a library policy, not an OOXML rule. Plan the
        // complete eviction before changing either realm. The worker receives
        // the same victims before terminal ACK, so it cannot retain a stale
        // model or reject a sheet that fits after eviction.
        let remaining = this.retainedSheetUsage;
        let nextCache = addWorksheetCacheUsage(remaining, measured, {}, policy);
        const victims: number[] = [];
        for (const [candidate] of this.sheetCache) {
          try {
            assertWorksheetCacheUsage(nextCache, 'get-worksheet', part, unit.usage, policy);
            break;
          } catch (error) {
            if (!(error instanceof OoxmlResourceLimitError)) throw error;
          }
          if ((this.sheetLeases.get(candidate) ?? 0) > 0) continue;
          const previous = this.sheetCacheUsage.get(candidate);
          if (!previous) continue;
          // Subtract from the unsaturated retained total, then add the incoming
          // sheet again. addWorksheetCacheUsage caps an over-budget candidate
          // at limit + 1, so subtracting from that candidate loses exact usage.
          remaining = addWorksheetCacheUsage(remaining, {
            rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
          }, previous, policy);
          nextCache = addWorksheetCacheUsage(remaining, measured, {}, policy);
          victims.push(candidate);
        }
        assertWorksheetCacheUsage(nextCache, 'get-worksheet', part, unit.usage, policy);
        let pendingEviction: { done: Promise<void>; complete: () => void } | undefined;
        if (victims.length && this._mode === 'worker') {
          let complete!: () => void;
          const done = new Promise<void>((resolve) => { complete = resolve; });
          pendingEviction = { done, complete };
          for (const candidate of victims) this.evictingSheets.set(candidate, pendingEviction);
        }
        try {
          if (pendingEviction) {
            await this.requireBridge().request(
              (id) => ({ type: 'evictWorksheets', id, sheetIndices: victims }) satisfies RenderWorkerRequest,
            );
            if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
          }
          for (const candidate of victims) {
            if ((this.sheetLeases.get(candidate) ?? 0) > 0) {
              throw new Error(`Cannot evict leased worksheet ${candidate}`);
            }
          }
          for (const candidate of victims) {
            this.sheetCache.delete(candidate);
            this.sheetCacheUsage.delete(candidate);
          }
          // The terminal can still be canceled or fail its ACK. Eviction has
          // already happened in both realms, so its accounting commits now.
          this.retainedSheetUsage = remaining;
        } catch (error) {
          if (pendingEviction) {
            // A lost reply or divergent lease cannot prove the worker state.
            // Close both caches rather than let either realm serve a stale copy.
            this.destroy();
          }
          throw error;
        } finally {
          if (pendingEviction) {
            for (const candidate of victims) {
              if (this.evictingSheets.get(candidate) === pendingEviction) {
                this.evictingSheets.delete(candidate);
              }
            }
            pendingEviction.complete();
          }
        }
        terminal = worksheet;
        terminalUsage = measured;
        nextCacheUsage = nextCache;
      }
      terminalAcknowledged = true;
      if (!terminal || !nextCacheUsage || !terminalUsage) {
        throw new Error(`XLSX worksheet ${sheetIndex} did not produce a terminal model`);
      }
      // The coordinator has ACKed the accepted terminal before it completes.
      // Only now commit Browser-retained cache ownership/accounting.
      this.retainedSheetUsage = nextCacheUsage;
      this.sheetCache.set(sheetIndex, terminal);
      this.sheetCacheUsage.set(sheetIndex, terminalUsage);
      // #1974: a failure is known only once main has accepted the terminal.
      if (terminal.parseError) {
        this._partFailureLog.add({
          kind: 'sheet', sheetIndex, sheetName: terminal.name, message: terminal.parseError,
        });
      }
      // Both realms now retain the same model even if an optional font load
      // rejects; a later operation can still acquire the cached worksheet.
      await this.retainWorksheetOfficeFonts(terminal);
      // Rows were demanded as they streamed in; the terminal model adds its
      // drawings (shapes, charts, slicers).
      await this.demandWorksheetFonts(terminal, rowsWithoutColumnStyles ? terminal.rows : undefined);
      const mainFonts = typeof document !== 'undefined' ? document.fonts : undefined;
      const mainOffice = mainFonts ? this.retainedFontSets.get(mainFonts)?.loaded?.office : undefined;
      bindXlsxWorksheetOfficeFontRoutes(terminal, mainOffice?.routes, this.loadedFontSources(mainFonts));
      progress.finish(terminal);
      return terminal;
    } catch (error) {
      // Terminal ACK may have committed on the worker even when its reply was
      // lost. The main cache has not admitted that model; terminate the worker
      // and discard all local models before another operation can read it.
      if (this._mode === 'worker' && terminal && !terminalAcknowledged) this.destroy();
      this.latchFatalResourceFailure(error);
      progress.fail(error);
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
        this.latchFatalResourceFailure(error);
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
    const provisional = [...(this.sheetPreviews?.values() ?? [])]
      .some((preview) => preview.worksheet && !preview.complete);
    const p = provisional
      ? this.getImageWithinArchiveOperation(imagePath, mimeType)
      : this.runArchiveOperation(() => this.getImageWithinArchiveOperation(imagePath, mimeType));
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
   * Produce a best-effort, text-focused GitHub-flavoured markdown projection:
   * each sheet becomes a
   * `## SheetName` section followed by a pipe table of its populated bounding
   * box (fully-empty middle rows trimmed, ULP noise masked). Styling, charts,
   * and drawings are discarded, while review comments are kept in a final
   * quoted appendix. The projection is meant for AI ingestion and full-text
   * search, not an authoritative semantic or layout representation.
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

    const lease = await this.acquireWorksheetLease(targetIndex);
    try {
      const ws = lease.worksheet;
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
    } finally {
      lease.release();
    }
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
    // This invocation owns its viewport identity across resource awaits and
    // caller callbacks; a later edit to the caller's object is another frame.
    viewport = { row: viewport.row, col: viewport.col, rows: viewport.rows, cols: viewport.cols };
    if (this._mode === 'worker') {
      throw new Error(
        "renderViewport(canvas) is unavailable in mode: 'worker'; use renderViewportToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    const styles = this.parsedWorkbook.styles;
    const extracted = extractViewerRenderContext(opts as WireRenderViewportOptions);
    const { sizeOverrides, errorPlaceholder, ...renderOpts } = extracted.opts as WireRenderViewportOptions
      & Pick<XlsxRenderViewportOptions, 'errorPlaceholder'>;
    const preparedPlaceholder = prepareErrorPlaceholder<XlsxSheetPartFailure>(errorPlaceholder);
    const targetFontSet = isHTMLCanvas(target)
      ? target.ownerDocument.fonts
      : (typeof document !== 'undefined' ? document.fonts : null);
    const preview = this.sheetPreviews?.get(sheetIndex);
    if (preview?.worksheet && extracted.worksheet) {
      await preview.waitFor(Math.max(
        viewport.row + viewport.rows - 1,
        renderOpts.freezeRows ?? 0,
      ));
      this.assertResourceHealthy();
      const ws = extracted.worksheet;
      const errorPlaceholderText = this.sheetErrorPlaceholderText(preparedPlaceholder, sheetIndex, ws);
      // Visible images may be fetched during the provisional paint. Their
      // requests use the pull owner's image path while the cursor holds the
      // package operation, so the render itself must not queue behind that
      // operation.
      await renderWorksheetViewport(
        {
          ws, styles, cjkFallback: this.cjkFallback, math: this.math,
          threeD: this.threeD, regionMap: this.regionMap,
          chartEx: this.chartEx, tiff: this.tiff,
        }, target, viewport,
        {
          ...renderOpts,
          authoritativeMdw: extracted.layoutMetrics?.maximumDigitWidth,
          officeFontRoutes: targetFontSet
            ? this.retainedFontSets.get(targetFontSet)?.loaded?.office.routes
            : undefined,
          fontSources: this.loadedFontSources(targetFontSet),
          fetchImage: this._fetchImage,
          ...(errorPlaceholderText === undefined ? {} : { errorPlaceholderText }),
        },
      );
      return;
    }
    return this.withWorksheetArchiveOperation(sheetIndex, (source) => {
      const ws = extracted.worksheet ?? createSizeOverriddenWorksheet(source, sizeOverrides);
      if (ws !== source) inheritSheetRenderCache(source, ws);
      const errorPlaceholderText = this.sheetErrorPlaceholderText(preparedPlaceholder, sheetIndex, ws);
      // The render bind may invalidate a geometry snapshot made in another
      // font realm. Pin the viewer's MDW inside the renderer after binding.
      return renderWorksheetViewport(
        {
          ws,
          styles,
          cjkFallback: this.cjkFallback,
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
        {
          ...renderOpts,
          authoritativeMdw: extracted.layoutMetrics?.maximumDigitWidth,
          officeFontRoutes: targetFontSet
            ? this.retainedFontSets.get(targetFontSet)?.loaded?.office.routes
            : undefined,
          fontSources: this.loadedFontSources(targetFontSet),
          fetchImage: this._fetchImage,
          ...(errorPlaceholderText === undefined ? {} : { errorPlaceholderText }),
        },
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
    // This invocation owns its viewport identity across resource awaits and
    // caller callbacks; a later edit to the caller's object is another frame.
    viewport = { row: viewport.row, col: viewport.col, rows: viewport.rows, cols: viewport.cols };
    const extracted = extractViewerRenderContext(opts);
    // The option is extracted before any spread: a function never reaches
    // postMessage, not even for healthy sheets (#1974).
    const { errorPlaceholder, ...callerOpts } = extracted.opts as WireRenderViewportOptions
      & Pick<XlsxRenderViewportOptions, 'errorPlaceholder'>;
    const wireOpts = { ...callerOpts, dpr: opts.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      const preparedPlaceholder = prepareErrorPlaceholder<XlsxSheetPartFailure>(errorPlaceholder);
      // A viewer frame's freshness latch: a superseded frame drops the retry.
      const isCurrent = xlsxRenderCommitGuardOf(opts as RenderViewportOptions);
      if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= this.sheetCount) {
        throw new Error(`Sheet index ${sheetIndex} out of range (count: ${this.sheetCount})`);
      }
      // WHATWG HTML: an OffscreenCanvas constructed here snapshots this
      // document's language and direction, while one constructed in the Worker
      // has unknown language. Transferring it before any getContext() carries
      // those inherited values, matching the main-mode surface below. Each
      // send constructs its own surface because transfer detaches it.
      const request = (final = false) => {
        if (final && isCurrent?.() === false) return Promise.resolve(null);
        // Resolved at dispatch: a pull that ran just before may have admitted
        // this sheet's terminal on main.
        const errorPlaceholderText = this.workerErrorPlaceholderText(preparedPlaceholder, sheetIndex, final);
        const canvas = new OffscreenCanvas(1, 1);
        return this.requireBridge().request(
          (id) => ({
            type: 'renderViewport',
            id,
            sheetIndex,
            viewport,
            opts: errorPlaceholderText === undefined ? wireOpts : { ...wireOpts, errorPlaceholderText },
            layoutMetrics: extracted.layoutMetrics,
            viewProjection: extracted.projection,
            canvas,
          }) satisfies RenderWorkerRequest,
          [canvas],
        );
      };
      const preview = this.sheetPreviews?.get(sheetIndex);
      let res;
      if (preview?.worksheet && extracted.worksheet) {
        await preview.waitFor(Math.max(
          viewport.row + viewport.rows - 1,
          wireOpts.freezeRows ?? 0,
        ));
        this.assertResourceHealthy();
        res = await request();
      } else {
        res = await this.withWorksheetArchiveOperation(sheetIndex, () => request());
      }
      if ((res as RenderWorkerResponse).type === 'errorPlaceholderRequired') {
        // Tentative: publish nothing yet. Wait for main's own guarded
        // admission of this terminal (it records the failure), then retry
        // once. Admission failure, cancellation or destroy rejects here.
        await this.getWorksheet(sheetIndex);
        this.assertResourceHealthy();
        res = isCurrent?.() === false
          ? null
          : await this.withWorksheetArchiveOperation(sheetIndex, () => request(true));
        // Superseded between the reply and the retry: the frame is discarded
        // by its owner, so hand back an empty surface without resolving text.
        if (res === null) return new OffscreenCanvas(1, 1).transferToImageBitmap();
        if ((res as RenderWorkerResponse).type === 'errorPlaceholderRequired') {
          throw new Error('XLSX render worker requested placeholder text twice');
        }
      }
      return (res as Extract<RenderWorkerResponse, { type: 'viewportRendered' }>).bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderViewport(
      off,
      sheetIndex,
      viewport,
      errorPlaceholder === undefined ? wireOpts : { ...wireOpts, errorPlaceholder } as RenderViewportOptions,
    );
    return off.transferToImageBitmap();
  }

  /** @internal Drop projections owned by a destroyed viewer. */
  [releaseXlsxViewerProjection](projectionId: number): void {
    if (this._mode !== 'worker' || !this.bridge) return;
    this.bridge.post(
      { type: 'releaseViewProjection', projectionId } satisfies RenderWorkerRequest,
    );
  }

  /** The viewer may keep displaying a model after its cache entry is evicted.
   * Restore its active lease without materializing or changing that model. */
  private async retainWorksheetReference(sheetIndex: number): Promise<() => void> {
    const pending = this.evictingSheets.get(sheetIndex);
    if (pending) {
      await pending.done;
      return this.retainWorksheetReference(sheetIndex);
    }
    return this.parsedWorkbook ? this.pinWorksheet(sheetIndex) : () => undefined;
  }

  private withWorksheetArchiveOperation<T>(
    sheetIndex: number,
    operation: (worksheet: Worksheet) => Promise<T>,
  ): Promise<T> {
    const pending = this.evictingSheets.get(sheetIndex);
    if (pending) return pending.done.then(() => this.withWorksheetArchiveOperation(sheetIndex, operation));
    const release = this.pinWorksheet(sheetIndex);
    return this.withPinnedWorksheetArchiveOperation(sheetIndex, operation).finally(release);
  }

  private withPinnedWorksheetArchiveOperation<T>(
    sheetIndex: number,
    operation: (worksheet: Worksheet) => Promise<T>,
  ): Promise<T> {
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) {
      this.touchWorksheet(sheetIndex, cached);
      return this.runArchiveOperation(() => operation(cached));
    }
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
    const progress = new WorksheetPreview([]);
    (this.sheetPreviews ??= new Map()).set(sheetIndex, progress);
    const combined = this.runArchiveOperation(async () => {
      try {
        const worksheet = await this.loadWorksheetStream(sheetIndex, sheetMeta.name, progress);
        resolveLoad(worksheet);
        return await operation(worksheet);
      } catch (error) {
        rejectLoad(error);
        throw error;
      } finally {
        if (this.sheetLoads.get(sheetIndex) === load) this.sheetLoads.delete(sheetIndex);
        if (this.sheetPreviews?.get(sheetIndex) === progress) this.sheetPreviews.delete(sheetIndex);
      }
    });
    return combined;
  }

  destroy(): void {
    this.generation = (this.generation ?? 1) + 1;
    this._partFailureLog.close();
    void this.worksheetPullClient?.cancelAll('closed').catch(() => undefined);
    this.worksheetPullClient = null;
    this.bridge?.terminate();
    this.bridge = null;
    this.parsedWorkbook = null;
    this.sheetCache.clear();
    this.sheetCacheUsage.clear();
    this.sheetLeases.clear();
    this.completePendingEvictions();
    this.retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
    this.sheetLoads.clear();
    for (const progress of this.sheetPreviews?.values() ?? []) {
      progress.fail(new Error('XLSX workbook has been destroyed'));
    }
    this.sheetPreviews?.clear();
    this.fontsDestroyed = true;
    this.fontResolverBridge?.dispose();
    this.fontResolverBridge = null;
    this.fontResolverFailures = new Map();
    for (const retained of this.retainedFontSets.values()) {
      // Releasing settles an in-flight ensure (it resolves empty).
      retained.session?.release();
      if (retained.loaded) {
        unloadOfficeFontFallbacks(retained.loaded.office.faces);
      }
      // An in-flight registration observes fontsDestroyed in its own completion
      // callback and releases exactly once when the faces become available.
    }
    this.retainedFontSets.clear();
    // The descriptors may hold application font bytes (fontFiles data,
    // inline data: URLs): a destroyed workbook must not keep them reachable.
    this.fonts = undefined;
    this.fontDemand = null;
    this.installedFamilies = null;
    this.fontSessionType = null;
    this.fontAcquirerFor = null;
    this.fontDeltas = [];
    this.officeFontRequests = [];
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

  private latchFatalResourceFailure(error: unknown): void {
    // Aggregate cache capacity can change when a lease ends. Preserve the
    // document-level poison boundary for hard model/package limits only.
    if (
      error instanceof OoxmlResourceLimitError &&
      error.details.violation.resource !== 'worksheet-cache'
    ) this.resourceFailure ??= error;
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

/** @internal Viewer-owned operations retain a model until their content reads
 * finish. This adapter stays out of the package entry's public API. */
export function acquireXlsxWorksheet(workbook: XlsxWorkbook, sheetIndex: number): Promise<{
  worksheet: Worksheet;
  release: () => void;
}> {
  return workbook['acquireWorksheetLease'](sheetIndex);
}

/** @internal Fast viewer acquisition. Structural test doubles retain the
 * established full-model contract. */
export function acquireXlsxWorksheetPreview(workbook: XlsxWorkbook, sheetIndex: number): Promise<{
  worksheet: Worksheet;
  release: () => void;
  partial: boolean;
  completion: Promise<Worksheet>;
  waitForRows?: (row: number) => Promise<void>;
}> {
  const acquire = workbook['acquireWorksheetPreviewLease'];
  if (typeof acquire === 'function') return acquire.call(workbook, sheetIndex);
  return acquireXlsxWorksheet(workbook, sheetIndex).then((lease) => ({
    ...lease, partial: false, completion: Promise.resolve(lease.worksheet),
  }));
}

/** @internal Restore the active viewer lease for a displayed worksheet even
 * when its workbook cache entry has already been evicted. */
export function retainXlsxWorksheetReference(workbook: XlsxWorkbook, sheetIndex: number): Promise<() => void> {
  const retain = workbook['retainWorksheetReference'];
  // Structural viewer test doubles may not implement this cache-only hook.
  return typeof retain === 'function'
    ? retain.call(workbook, sheetIndex)
    : Promise.resolve(() => undefined);
}
declare const __OOXML_MODEL_SOURCES__: boolean;
