import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/xlsx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  dropBitmapCacheByPath,
  dropDuotoneBitmapCache,
  dropSvgImageCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, WorkerRequest, WorkerResponse, Cell, SheetVisibility } from './types.js';
import { selectSheetVisibility } from './sheet-visibility.js';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import { formatCellValue } from './number-format.js';
import { resolveSharedStrings } from './shared-strings.js';
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

/** Options for {@link XlsxWorkbook.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `maxZipEntryBytes`, `math`)
 *  with the worker-rendering mode. */
export interface LoadOptions extends CoreLoadOptions {
  /** Maximum decompressed bytes across distinct ZIP entries (default: 512 MiB). */
  maxZipTotalBytes?: number;
  /** Maximum ZIP central-directory entries (default: 20,000). */
  maxZipEntries?: number;
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link XlsxWorkbook.renderViewportToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   * The math engine is unavailable in this mode (equations are skipped).
   */
  mode?: 'main' | 'worker';
}

export class XlsxWorkbook {
  private worker: Worker;
  private bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  /** Cache of decoded image sources keyed by their zip `imagePath`. Shared
   *  across sheets. */
  private imageCache = new Map<string, CanvasImageSource | null>();
  /** Cache of fetched image *bytes* (as Blobs) keyed by zip path, populated by
   *  {@link XlsxWorkbook.getImage}. Twin of pptx/docx's per-instance
   *  `_imageCache`; kept separate from {@link XlsxWorkbook.imageCache} (decoded
   *  sources) so each layer dedupes independently. */
  private imageBlobCache = new Map<string, Promise<Blob>>();
  /** One stable closure per instance: core's path-keyed SVG cache namespaces on
   *  this identity, so two open workbooks never swap a shared zip path (e.g.
   *  xl/media/image1.svg). Reusing one reference also lets the SVG cache hit
   *  across viewport renders. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImage(path, mime);
  private rawData: ArrayBuffer | null = null;
  private maxZipEntryBytes: number | undefined;
  private maxZipTotalBytes: number | undefined;
  private maxZipEntries: number | undefined;
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderViewport` call reuses it — equations in shapes render when present,
   *  and are skipped (engine tree-shaken) when omitted. */
  private math: MathRenderer | undefined;
  /** Google-Fonts `FontFace` objects this workbook preloaded into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with its
   *  own FontFaceSet). Released in {@link destroy} so they do not leak into the
   *  shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in core,
   *  so a web font shared with another open workbook survives until both go). */
  private googleFontFaces: FontFace[] = [];
  private _mode: 'main' | 'worker' = 'main';

  private constructor(worker: Worker, mode: 'main' | 'worker', wasmUrlOverride?: string | URL) {
    this.worker = worker;
    this._mode = mode;
    this.bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this.worker, {
      correlate: (res) => res.id,
      toError: (res) => (res.type === 'error' ? res.message : undefined),
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
    this.bridge.post({ type: 'init', wasmUrl } satisfies WorkerRequest);
  }

  /** Parse an XLSX from a URL or ArrayBuffer. */
  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<XlsxWorkbook> {
    const mode = opts.mode ?? 'main';
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
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    buffer = toArrayBuffer(await resolveOoxmlContainer(buffer, opts.password));
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const wb = new XlsxWorkbook(worker, mode, opts.wasmUrl);
    await wb._load(buffer, opts);
    return wb;
  }

  // `load()` always resolves a URL/string source to an ArrayBuffer (via
  // resolveOoxmlContainer, so decryption sees the container before the render
  // worker is constructed) before calling `_load`, so this only ever receives
  // an ArrayBuffer — no separate string-source fetch branch is needed here.
  private async _load(data: ArrayBuffer, opts: LoadOptions = {}): Promise<void> {
    this.rawData = data;
    this.maxZipEntryBytes = opts.maxZipEntryBytes;
    this.maxZipTotalBytes = opts.maxZipTotalBytes;
    this.maxZipEntries = opts.maxZipEntries;
    this.math = opts.math;
    if (opts.math && this._mode === 'worker') {
      console.warn(
        "[ooxml] the math engine is unavailable in mode: 'worker'; equations will be skipped. Use mode: 'main' for workbooks with equations.",
      );
    }
    // In worker mode the worker preloads fonts before its first render
    // (rendering measures text), so the flag is forwarded; in main mode fonts
    // are loaded here after parse.
    const parsed = await this.bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({
              type: 'parse',
              id,
              data: data.slice(0),
              maxZipEntryBytes: this.maxZipEntryBytes,
              maxZipTotalBytes: this.maxZipTotalBytes,
              maxZipEntries: this.maxZipEntries,
              useGoogleFonts: !!opts.useGoogleFonts,
            } satisfies RenderWorkerRequest)
          : ({
              type: 'parse',
              id,
              data: data.slice(0),
              maxZipEntryBytes: this.maxZipEntryBytes,
              maxZipTotalBytes: this.maxZipTotalBytes,
              maxZipEntries: this.maxZipEntries,
            } satisfies WorkerRequest),
      undefined,
      { timeoutMs: opts.workerTimeoutMs },
    );
    // Both modes carry the light, workbook-level ParsedWorkbook back, so
    // sheetNames / tabColors / resolveValidationList keep working. In parse mode
    // it arrives as transferred UTF-8 JSON bytes — decode + parse once here.
    if (this._mode === 'worker') {
      this.parsedWorkbook = (parsed as Extract<RenderWorkerResponse, { type: 'parsed' }>).workbook;
    } else {
      const { workbookJson } = parsed as Extract<WorkerResponse, { type: 'parsed' }>;
      this.parsedWorkbook = JSON.parse(
        new TextDecoder().decode(new Uint8Array(workbookJson)),
      ) as ParsedWorkbook;
    }
    // #773: a workbook-level degradation (a present-but-corrupt shared part such
    // as `xl/sharedStrings.xml`, which blanks every string cell across all sheets)
    // still opens the workbook, but must not be SILENT. Surface it once at load —
    // the model also carries it on `workbook.parseError` for callers that inspect
    // it. Per-sheet placeholders (a broken worksheet) already surface via the
    // sheet-grid overlay, so they are not re-logged here.
    const workbookError = this.parsedWorkbook?.workbook.parseError;
    if (workbookError) {
      console.warn(`[ooxml] xlsx opened with a degraded part: ${workbookError}`);
    }
    if (this._mode === 'main' && opts.useGoogleFonts) {
      this.googleFontFaces = await preloadGoogleFonts(
        xlsxFontPreloadNames(this.parsedWorkbook),
        XLSX_GOOGLE_FONTS,
      );
    }
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
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) return cached;
    // `!this.rawData` guards that `parse` has run: the worker retained the
    // whole-workbook buffer at parse time, and `parseSheet` reuses it. We no
    // longer re-send the buffer here (it previously structured-cloned the entire
    // file per sheet switch); the retained `rawData` is still kept for
    // `getImage`'s route-through-worker path.
    if (!this.parsedWorkbook || !this.rawData) {
      throw new Error('Workbook not loaded');
    }
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);

    const res = await this.bridge.request((id) => ({
      type: 'parseSheet',
      id,
      sheetIndex,
      sheetName: sheetMeta.name,
      maxZipEntryBytes: this.maxZipEntryBytes,
      maxZipTotalBytes: this.maxZipTotalBytes,
      maxZipEntries: this.maxZipEntries,
    }));
    // Parse mode: the worker forwards the sheet as transferred UTF-8 JSON bytes
    // — decode + parse once here. Worker (render) mode: the worker already
    // decoded it and sends the object back as a structured clone.
    let ws: Worksheet;
    if (this._mode === 'worker') {
      ws = (res as Extract<RenderWorkerResponse, { type: 'parsedSheet' }>).worksheet;
    } else {
      const { worksheetJson } = res as Extract<WorkerResponse, { type: 'parsedSheet' }>;
      ws = JSON.parse(new TextDecoder().decode(new Uint8Array(worksheetJson))) as Worksheet;
    }
    // Resolve shared-string references (`{ type: 'shared', si }`) against the
    // workbook's dedup'd sharedStrings table so no consumer downstream ever
    // sees an unresolved cell. Covers BOTH decode paths above.
    resolveSharedStrings(ws, this.parsedWorkbook.sharedStrings);
    this.sheetCache.set(sheetIndex, ws);
    return ws;
  }

  /**
   * Fetch an embedded image's bytes by zip path (e.g. `xl/media/image1.png`),
   * wrapped in a Blob of the given MIME. The bytes are pulled through the
   * persistent worker via the `extractImage` message (twin of pptx/docx's
   * `getImage`/`getMedia`); results are cached by path for the lifetime of this
   * instance. The renderer's `fetchImage` option points here so image bytes are
   * extracted lazily rather than inlined as base64 at parse time.
   *
   * Routed through the worker even though the main thread also retains
   * `rawData`, to keep all WASM `extract_image` decoding on the worker (the
   * route-through-worker decision).
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    const hit = this.imageBlobCache.get(imagePath);
    if (hit) return hit;
    const p = this.bridge
      .request((id) => ({ type: 'extractImage', id, path: imagePath }) satisfies WorkerRequest)
      .then((res) => {
        const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      });
    this.imageBlobCache.set(imagePath, p);
    return p;
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
    const res = await this.bridge.request(
      (id) => ({ type: 'toMarkdown', id }) satisfies WorkerRequest,
    );
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

  /**
   * Render a sheet viewport into `target`. Note: `opts.fetchImage` is ignored
   * here — image bytes always come from this workbook's own archive through its
   * stable per-instance loader, whose closure identity keys the shared decoded
   * caches, the render-pass lease, and {@link destroy}'s cache drops. Callers
   * needing a custom byte source should use the standalone
   * `renderWorksheetViewport` orchestrator directly.
   */
  async renderViewport(
    target: HTMLCanvasElement | OffscreenCanvas,
    sheetIndex: number,
    viewport: ViewportRange,
    opts: RenderViewportOptions = {},
  ): Promise<void> {
    if (this._mode === 'worker') {
      throw new Error(
        "renderViewport(canvas) is unavailable in mode: 'worker'; use renderViewportToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    // Hot path: during scroll the worksheet is already cached. Skip the await
    // to keep the whole render in a single synchronous task so the browser
    // doesn't paint between the canvas clear and the draw.
    const ws = this.sheetCache.get(sheetIndex) ?? await this.getWorksheet(sheetIndex);
    return renderWorksheetViewport(
      { ws, styles: this.parsedWorkbook.styles, imageCache: this.imageCache, math: this.math },
      target,
      viewport,
      // Always render with THIS instance's stable `_fetchImage` closure (placed
      // after the spread so it wins over a caller-supplied one, matching
      // docx/pptx). The closure's identity is the namespace key for the shared
      // per-document caches AND the render-pass lease counter, and destroy()
      // drops exactly `_fetchImage`'s namespaces — a per-call closure would
      // split the cache/lease namespace every render and leave its bitmaps
      // undropped at destroy(). Callers who need a custom byte source use the
      // orchestrator's renderWorksheetViewport directly.
      { ...opts, fetchImage: this._fetchImage },
    );
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
    opts: WireRenderViewportOptions & { width: number; height: number },
  ): Promise<ImageBitmap> {
    const wireOpts = { ...opts, dpr: opts.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= this.sheetCount) {
        throw new Error(`Sheet index ${sheetIndex} out of range (count: ${this.sheetCount})`);
      }
      const res = await this.bridge.request(
        (id) => ({ type: 'renderViewport', id, sheetIndex, viewport, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'viewportRendered' }>).bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderViewport(off, sheetIndex, viewport, wireOpts);
    return off.transferToImageBitmap();
  }

  destroy(): void {
    this.bridge.terminate();
    this.parsedWorkbook = null;
    this.sheetCache.clear();
    // Release the Google-Fonts substitutes this workbook preloaded into the
    // shared FontFaceSet (main mode). Refcounted in core: a web font also used by
    // another open workbook stays until that one is destroyed too. Without this,
    // every opened workbook left its Google FontFace objects in `document.fonts`
    // forever (SPA memory leak).
    if (this.googleFontFaces.length > 0) {
      unloadGoogleFonts(this.googleFontFaces);
      this.googleFontFaces = [];
    }
    // The per-instance imageCache is a pure synchronous-lookup map into the
    // shared, per-`_fetchImage` core caches (base raster via getCachedBitmapByPath,
    // duotone recolour via getCachedDuotoneBitmapByPath, SVG via
    // getCachedSvgImageByPath) — the same ownership split docx/pptx use. Clearing
    // the map only drops lookup references; the GPU-backed ImageBitmaps and SVG
    // object URLs are released by dropping the three shared caches keyed by
    // `_fetchImage`, so a discarded workbook frees its GPU/URL handles promptly
    // rather than waiting for GC.
    this.imageCache.clear();
    dropBitmapCacheByPath(this._fetchImage);
    dropDuotoneBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
    this.imageBlobCache.clear();
    this.rawData = null;
  }
}
