import type { DocComment, DocNote, RenderPageOptions, WorkerResponse } from './types';
import type { DocxTextRunInfo } from './renderer';

/** Lightweight summary returned by the render worker's `parse` — everything
 *  the main-thread proxy needs for its synchronous getters. The full model
 *  stays in the worker. */
export interface DocumentMeta {
  pageCount: number;
  comments: DocComment[];
  footnotes: DocNote[];
  endnotes: DocNote[];
  /** ECMA-376 §17.6.13 / §17.6.11 — per-page page size (pt), one entry per page,
   *  index-aligned with `pageCount`. Built worker-side from canonical page geometry
   *  so the main thread can lay out (e.g. a scroll viewer's spacer)
   *  without the full model. Genuinely per-page for a mixed-geometry document. */
  pageSizes: { widthPt: number; heightPt: number }[];
  /** ECMA-376 §17.13.6.2 — `bookmarkName → 0-based page index` for internal
   *  hyperlink anchors (`<w:hyperlink w:anchor>`, §17.16.23). Built worker-side
   *  from the paginated pages (the same source `pageSizes` uses) so an internal
   *  link can resolve its destination page in worker mode without the full model.
   *  Serialized as `[name, pageIndex]` entries (a `Map` can't cross the wire). */
  bookmarkPages: [string, number][];
}

/** Serializable subset of RenderPageOptions (callbacks cannot cross the wire). */
export type WireRenderPageOptions = Omit<RenderPageOptions, 'onTextRun'>;

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number; useGoogleFonts?: boolean; defaultCurrentDateMs: number }
  | { type: 'renderPage'; id: number; pageIndex: number; opts: WireRenderPageOptions }
  // IX6 — collect a page's text-run geometry WITHOUT transferring a bitmap. The
  // find controller scans every page for its runs; a bitmap per page would be
  // wasted work + transfer for pages the user never looks at.
  | { type: 'collectRuns'; id: number; pageIndex: number; opts: WireRenderPageOptions }
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerResponse =
  | Exclude<WorkerResponse, { type: 'parsed' }>
  | { type: 'parsedMeta'; id: number; meta: DocumentMeta }
  // OffscreenCanvas cannot select/probe the OpenType `vert` feature. A render
  // worker that parses vertical East-Asian content returns the normalized model
  // so the proxy can continue through main-thread rendering instead of silently
  // painting horizontal glyph forms.
  | { type: 'mainThreadVerticalFallback'; id: number; documentJson: ArrayBuffer }
  // The worker projects structured-clone-safe run geometry from the same
  // retained layout variant it paints and ships it beside the bitmap.
  | { type: 'pageRendered'; id: number; bitmap: ImageBitmap; runs: DocxTextRunInfo[] }
  | { type: 'runsCollected'; id: number; runs: DocxTextRunInfo[] };
