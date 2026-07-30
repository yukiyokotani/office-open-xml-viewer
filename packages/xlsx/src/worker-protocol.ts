import type {
  ViewportRange,
  RenderViewportOptions,
  WorkerResponse,
  ParsedWorkbook,
  Worksheet,
} from './types.js';

/**
 * View-only per-band size overrides for one sheet, carried with every worker
 * `renderViewport` request. The render worker draws from its own worker-local
 * parsed-sheet cache, so main-thread Worksheet mutations (outline
 * collapse/expand via the size-0 hidden encoding, drag-to-resize #567) never
 * reach it on their own — without this channel the gutter/overlays update but
 * the grid bitmap stays stale.
 *
 * Semantics: keys are 1-based band indices; a number is the band's current
 * `rowHeights` / `colWidths` model value, `null` means "no entry — fall back
 * to the sheet default". The main thread accumulates every band the user has
 * touched this session (entries are updated in place, never removed), so
 * re-applying the full map is idempotent and converges the worker's cached
 * sheet to the main model even across worker-side re-parses.
 */
export interface WireSizeOverrides {
  rows?: Record<number, number | null>;
  cols?: Record<number, number | null>;
}

/**
 * Apply {@link WireSizeOverrides} to a worksheet's size maps (mutates `ws`).
 * Runs worker-side on every `renderViewport` before drawing. Touches only the
 * two size maps, so the worker's memoized render cache (cell map, merge sets —
 * keyed by the Worksheet object identity) stays valid.
 */
export function applySizeOverrides(ws: Worksheet, overrides: WireSizeOverrides | undefined): void {
  if (!overrides) return;
  if (overrides.rows) {
    for (const [k, v] of Object.entries(overrides.rows)) {
      const idx = Number(k);
      if (v === null) delete ws.rowHeights[idx];
      else ws.rowHeights[idx] = v;
    }
  }
  if (overrides.cols) {
    for (const [k, v] of Object.entries(overrides.cols)) {
      const idx = Number(k);
      if (v === null) delete ws.colWidths[idx];
      else ws.colWidths[idx] = v;
    }
  }
}

/** Serializable subset of RenderViewportOptions: drop the callback, the image
 *  cache, and the `fetchImage` loader (all non-cloneable; the worker owns its
 *  own cache and supplies its own in-worker fetchImage). Extended with the
 *  optional {@link WireSizeOverrides} so view-only size mutations reach the
 *  worker's local sheet copy; absent (the common case) when nothing has been
 *  resized or collapsed, keeping the wire payload unchanged. */
export type WireRenderViewportOptions = Omit<
  RenderViewportOptions,
  'onTextRun' | 'loadedImages' | 'fetchImage'
> & { sizeOverrides?: WireSizeOverrides };

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number; useGoogleFonts?: boolean }
  // `parseSheet` lets worker-mode XlsxWorkbook.getWorksheet (and the
  // resolveValidationList range path that awaits it) work, mirroring how the
  // pptx render worker handles `extractMedia` for getMedia. The render worker
  // already holds `rawData` + `workbook` from `parse`, so only `sheetIndex` is
  // load-bearing here; `sheetName` is carried for shape-compat with the
  // main-mode message getWorksheet posts but is ignored (the worker derives the
  // sheet name from its own `workbook`). No `data`: like main-mode `parseSheet`,
  // the buffer retained at `parse` is reused — never re-sent per sheet.
  | { type: 'parseSheet'; id: number; sheetIndex: number; sheetName?: string; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number }
  | { type: 'renderViewport'; id: number; sheetIndex: number; viewport: ViewportRange; opts: WireRenderViewportOptions }
  // Worker render mode decodes images in-worker via a getImage closure; this arm
  // exists only for protocol parity with worker.ts (so a stray extractImage
  // never hangs). The render worker reads bytes straight from its retained
  // rawData.
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerResponse =
  // `imageExtracted` / `error` are reused from WorkerResponse. `parsed` and
  // `parsedSheet` are NOT reused: the parse-only worker returns those models as
  // transferred JSON bytes, but the render worker has already decoded them
  // worker-side (it consumes them to render), so it sends the objects back to
  // the proxy as structured clones — no re-serialization. The light,
  // workbook-level ParsedWorkbook keeps synchronous getters (sheetNames,
  // tabColors, …) working; per-sheet data stays worker-side and is parsed on
  // demand.
  | Exclude<WorkerResponse, { type: 'parsed' } | { type: 'parsedSheet' }>
  | { type: 'parsed'; id: number; workbook: ParsedWorkbook }
  | { type: 'parsedSheet'; id: number; worksheet: Worksheet }
  | { type: 'viewportRendered'; id: number; bitmap: ImageBitmap };
