import type {
  ViewportRange,
  RenderViewportOptions,
  WorkerResponse,
  ParsedWorkbook,
  Worksheet,
} from './types.js';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import type { NormalizedOoxmlResourcePolicy } from '@silurus/ooxml-core/worker';
import type { PullSessionIdentity } from '@silurus/ooxml-core/worker';
import { GridGeometry } from './internal/grid-geometry.js';

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
 * touched this session (entries are updated in place, never removed). Each
 * render applies the full map to a render-local worksheet projection, leaving
 * the workbook/worker cache unchanged for other viewers.
 */
export interface WireSizeOverrides {
  rows?: Record<number, number | null>;
  cols?: Record<number, number | null>;
}

/**
 * Apply {@link WireSizeOverrides} to a worksheet's size maps (mutates `ws`).
 * Shared render paths call this only on a shallow render-local projection.
 */
export function applySizeOverrides(ws: Worksheet, overrides: WireSizeOverrides | undefined): void {
  if (!overrides) return;
  let changed = false;
  if (overrides.rows) {
    for (const [k, v] of Object.entries(overrides.rows)) {
      const idx = Number(k);
      if (v === null) {
        if (Object.hasOwn(ws.rowHeights, idx)) {
          delete ws.rowHeights[idx];
          changed = true;
        }
      } else if (ws.rowHeights[idx] !== v) {
        ws.rowHeights[idx] = v;
        changed = true;
      }
    }
  }
  if (overrides.cols) {
    for (const [k, v] of Object.entries(overrides.cols)) {
      const idx = Number(k);
      if (v === null) {
        if (Object.hasOwn(ws.colWidths, idx)) {
          delete ws.colWidths[idx];
          changed = true;
        }
      } else if (ws.colWidths[idx] !== v) {
        ws.colWidths[idx] = v;
        changed = true;
      }
    }
  }
  if (changed) GridGeometry.invalidate(ws);
}

/** Return a render-local worksheet when view overrides exist. The cached
 * worksheet remains immutable across independent viewers. */
export function createSizeOverriddenWorksheet(
  source: Worksheet,
  overrides: WireSizeOverrides | undefined,
): Worksheet {
  if (!overrides) return source;
  const view = {
    ...source,
    rowHeights: { ...source.rowHeights },
    colWidths: { ...source.colWidths },
  };
  applySizeOverrides(view, overrides);
  return view;
}

export interface WireViewProjection {
  readonly id: number;
  readonly revision: number;
  /** The owning viewer supplied every display-derived automatic row height in
   * `sizeOverrides.rows`; the worker must not rescan cells for this revision. */
  readonly autoRowHeightsPrepared?: boolean;
}

/** Worker-local cache for one shallow worksheet projection per viewer/sheet.
 * Repeated viewport paints at an unchanged revision reuse the same maps. */
export class WorksheetViewProjectionCache {
  private readonly entries = new Map<
    string,
    Readonly<{ revision: number; source: Worksheet; worksheet: Worksheet }>
  >();
  /** Viewer teardown can overtake a render already awaiting fonts/archive work.
   * A tombstone prevents that late render from resurrecting released entries. */
  private readonly releasedProjectionIds = new Set<number>();

  resolve(
    source: Worksheet,
    sheetIndex: number,
    projection: WireViewProjection | undefined,
    overrides: WireSizeOverrides | undefined,
  ): Readonly<{ worksheet: Worksheet; created: boolean }> {
    if (!projection) {
      const worksheet = createSizeOverriddenWorksheet(source, overrides);
      return { worksheet, created: worksheet !== source };
    }
    const cacheable = !this.releasedProjectionIds.has(projection.id);
    const key = `${projection.id}:${sheetIndex}`;
    const cached = cacheable ? this.entries.get(key) : undefined;
    if (
      cached &&
      cached.revision === projection.revision &&
      cached.source === source
    ) {
      return { worksheet: cached.worksheet, created: false };
    }
    const worksheet = createSizeOverriddenWorksheet(source, overrides);
    if (worksheet !== source && cacheable) {
      this.entries.set(key, { revision: projection.revision, source, worksheet });
      return { worksheet, created: true };
    }
    this.entries.delete(key);
    return { worksheet, created: worksheet !== source };
  }

  release(projectionId: number): void {
    this.releasedProjectionIds.add(projectionId);
    const prefix = `${projectionId}:`;
    for (const key of this.entries.keys()) {
      if (key.startsWith(prefix)) this.entries.delete(key);
    }
  }

  clear(): void {
    this.entries.clear();
    this.releasedProjectionIds.clear();
  }
}

/** Serializable subset of RenderViewportOptions: drop the callback, the image
 *  cache, and the `fetchImage` loader (all non-cloneable; the worker owns its
 *  own cache and supplies its own in-worker fetchImage). Extended with the
 *  optional {@link WireSizeOverrides} so view-only size mutations reach a
 *  render-local sheet projection; absent when nothing changed. */
export type WireRenderViewportOptions = Omit<
  RenderViewportOptions,
  'onTextRun' | 'loadedImages' | 'fetchImage'
> & {
  sizeOverrides?: WireSizeOverrides;
};

const viewerRenderContext = Symbol('xlsx-viewer-render-context');

type ViewerRenderViewportOptions = WireRenderViewportOptions & {
  [viewerRenderContext]?: Readonly<{
    maximumDigitWidth: number;
    worksheet?: Worksheet;
    projection?: WireViewProjection;
  }>;
};

/** @internal Attach the viewer-only render context. The symbol keeps the main-
 * thread worksheet projection out of the public OOXML options and out of
 * structured clone; workbook transport extracts only serializable fields for
 * worker rendering. */
export function withViewerRenderContext<T extends WireRenderViewportOptions>(
  opts: T,
  maximumDigitWidth: number,
  view?: Readonly<{
    worksheet: Worksheet;
    projection?: WireViewProjection;
  }>,
): T {
  if (!Number.isFinite(maximumDigitWidth) || maximumDigitWidth <= 0) {
    throw new Error('XLSX maximum digit width must be a finite positive number');
  }
  return {
    ...opts,
    [viewerRenderContext]: {
      maximumDigitWidth,
      worksheet: view?.worksheet,
      projection: view?.projection,
    },
  } as T & ViewerRenderViewportOptions;
}

/** @internal Split wire options from the viewer-only render context. */
export function extractViewerRenderContext(opts: WireRenderViewportOptions): {
  readonly opts: WireRenderViewportOptions;
  readonly layoutMetrics?: Readonly<{ maximumDigitWidth: number }>;
  readonly worksheet?: Worksheet;
  readonly projection?: WireViewProjection;
} {
  const internal = opts as ViewerRenderViewportOptions;
  const layoutMetrics = internal[viewerRenderContext];
  const wire = { ...opts } as ViewerRenderViewportOptions;
  delete wire[viewerRenderContext];
  return layoutMetrics
    ? {
        opts: wire,
        layoutMetrics: { maximumDigitWidth: layoutMetrics.maximumDigitWidth },
        worksheet: layoutMetrics.worksheet,
        projection: layoutMetrics.projection,
      }
    : { opts: wire };
}

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; resourcePolicy: NormalizedOoxmlResourcePolicy; useGoogleFonts?: boolean; useFontProvider?: boolean; renderers?: import('@silurus/ooxml-core/worker').WorkerRendererDescriptors }
  | ({ type: 'openSheetSession'; id: number; sheetIndex: number; sheetName: string } & PullSessionIdentity<number>)
  | {
      type: 'renderViewport';
      id: number;
      sheetIndex: number;
      viewport: ViewportRange;
      opts: WireRenderViewportOptions;
      /** Internal Window→Worker geometry authority. Not part of the public
       * render options because it exists only to align viewer interaction and
       * worker paint across font realms. */
      layoutMetrics?: Readonly<{ maximumDigitWidth: number }>;
      /** Viewer-local projection cache identity. The worker rebuilds its shallow
       * worksheet projection only when the revision changes. */
      viewProjection?: WireViewProjection;
    }
  | { type: 'releaseViewProjection'; projectionId: number }
  // Worker render mode decodes images in-worker via a getImage closure; this arm
  // exists only for protocol parity with worker.ts (so a stray extractImage
  // never hangs). The render worker reads bytes straight from its retained
  // archive.
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'resourceUsage'; id: number }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerResponse =
  // `imageExtracted` / `error` are reused from WorkerResponse. `parsed` is not:
  // the render worker already decoded the light workbook-level model
  // worker-side and sends it back as a structured clone. The light,
  // workbook-level ParsedWorkbook keeps synchronous getters (sheetNames,
  // tabColors, …) working; per-sheet data stays worker-side and is parsed on
  // demand.
  | Exclude<WorkerResponse, { type: 'parsed' }>
  | {
      type: 'parsed';
      id: number;
      workbook: ParsedWorkbook;
      usage?: OoxmlResourceUsageSnapshot;
    }
  | { type: 'viewportRendered'; id: number; bitmap: ImageBitmap }
  | ({ type: 'sheetSessionOpened'; id: number } & PullSessionIdentity<number>);
