import type { DocComment, DocNote, DocRevision, RenderPageOptions, WorkerResponse } from './types';
import type { CommentAnchorRange } from './comments';
import type { RevisionAnchorRange } from './revisions';
import type { DocxTextRunInfo } from './renderer';
import type {
  NormalizedOoxmlResourcePolicy,
  PullSessionCommand,
  PullSessionIdentity,
  PullSessionResponse,
  WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import type { DocxElementContextOptions } from './element-context';
import type { DocxElementContext, DocxPagePoint } from './selection-context';

/** Lightweight summary returned by the render worker's `parse` — everything
 *  the main-thread proxy needs for its synchronous getters. The full model
 *  stays in the worker. */
export interface DocumentMeta {
  pageCount: number;
  revisions: DocRevision[];
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
  /** ECMA-376 §17.13.4 comment-anchor ranges resolved from every retained
   *  story. Built worker-side with the same source identities used by rendered
   *  text runs, so consumers can join anchors to geometry without the full
   *  model. Absent for metadata produced by an older worker build. */
  commentAnchorRanges?: CommentAnchorRange[];
  /** ECMA-376 §17.13.5 revision ranges joined to retained source identities.
   * Deletions carry a deterministic final-state geometry fallback. */
  revisionAnchorRanges?: RevisionAnchorRange[];
}

/**
 * The layout-derived half of {@link DocumentMeta}: exactly what changes when the
 * viewed layout VARIANT changes.
 *
 * `revisions`, `comments`, `footnotes` and `endnotes` come from the parsed model
 * and are the same for every variant, so a variant switch keeps the ones the
 * load established instead of re-cloning them across the wire — the same
 * bandwidth rule {@link DocumentLayoutPartial.review} follows.
 */
export type DocumentLayoutMeta = Pick<
  DocumentMeta,
  'pageCount' | 'pageSizes' | 'bookmarkPages' | 'commentAnchorRanges' | 'revisionAnchorRanges'
>;

/**
 * Provisional layout geometry published by the render worker while it is still
 * paginating under `progressiveLayout`.
 *
 * Deliberately a strict SUBSET of {@link DocumentMeta}. The two anchor-range
 * fields are omitted because building them needs a whole-document run index
 * (`textRunSourceIndexForDocument`), which costs more per publication than the
 * publication saves — and a prefix's answers would be wrong for the pages that
 * do not exist yet. `DocumentMeta` already treats both as absent-tolerant, so
 * the host simply leaves them unset until the authoritative `parsedMeta`.
 */
export interface DocumentLayoutPartial {
  /** Pages published so far — NOT the document's total. */
  pageCount: number;
  /** Per-page size (pt) for the published pages, index-aligned with `pageCount`. */
  pageSizes: { widthPt: number; heightPt: number }[];
  /** Bookmark anchors resolvable within the published prefix. Cheap: the map is
   *  built from the prefix pages this publication just laid out. Anchors beyond
   *  the prefix are simply absent until layout completes. */
  bookmarkPages: [string, number][];
  /** Whether these pages are known to match the final layout. Always false
   *  today — see `layout/progressive.ts` on unbounded paginator lookahead. */
  exact: boolean;
  /** Model-derived review data, sent with the FIRST publication only.
   *
   *  It comes from the parsed model rather than the layout, so it costs nothing
   *  to produce — but re-cloning it on every publication would be real wire
   *  bandwidth on a heavily reviewed document. The host seeds its metadata from
   *  this and keeps it across later publications. */
  review?: Pick<DocumentMeta, 'revisions' | 'comments' | 'footnotes' | 'endnotes'>;
}

/** Serializable subset of RenderPageOptions (callbacks cannot cross the wire). */
export type WireRenderPageOptions = Omit<RenderPageOptions, 'onTextRun'>;

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  // `currentDateMs` / `showTrackedChanges` select the layout VARIANT the worker
  // paginates and reports metadata for. They are acquisition inputs, not paint
  // flags: hiding deletions changes line breaking, and DATE/TIME field text
  // changes measured widths, so each combination is a genuinely different
  // pagination with its own page count. Omitted means the document's default
  // view, which is what every load selected before these existed.
  | { type: 'parse'; id: number; data: ArrayBuffer; resourcePolicy: NormalizedOoxmlResourcePolicy; useGoogleFonts?: boolean; defaultCurrentDateMs: number; currentDateMs?: number; showTrackedChanges?: boolean; renderers?: WorkerRendererDescriptors; progressiveLayout?: boolean }
  // Repaginate the retained document as a DIFFERENT variant and report that
  // variant's metadata. `setLayoutView` in worker mode is this round-trip: the
  // model lives in the worker, so the host cannot compute the new page count,
  // page sizes, or anchor projections itself — and until it has them its
  // geometry accessors would describe a layout nobody is painting. Same field
  // semantics as `parse`: omitted means the document's default view.
  | { type: 'selectLayoutView'; id: number; currentDateMs?: number; showTrackedChanges?: boolean }
  | { type: 'renderPage'; id: number; pageIndex: number; opts: WireRenderPageOptions }
  // IX6 — collect a page's text-run geometry WITHOUT transferring a bitmap. The
  // find controller scans every page for its runs; a bitmap per page would be
  // wasted work + transfer for pages the user never looks at.
  | { type: 'collectRuns'; id: number; pageIndex: number; opts: WireRenderPageOptions }
  | {
      type: 'hitTestElement';
      id: number;
      pageIndex: number;
      point: DocxPagePoint;
      opts: DocxElementContextOptions;
    }
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'resourceUsage'; id: number }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerWireRequest =
  | RenderWorkerRequest
  | PullSessionCommand<number>;

export type RenderWorkerResponse =
  | Exclude<WorkerResponse, { type: 'documentSessionOpened' }>
  | PullSessionResponse<ArrayBuffer, number>
  | {
      type: 'parsedMeta';
      id: number;
      meta: DocumentMeta;
      usage?: OoxmlResourceUsageSnapshot;
    }
  // OffscreenCanvas cannot select/probe the OpenType `vert` feature. A render
  // worker that parses vertical East-Asian content opens a bounded model stream
  // so the proxy can continue through main-thread rendering instead of silently
  // painting horizontal glyph forms or receiving one monolithic JSON value.
  | ({
      type: 'mainThreadVerticalFallback';
      id: number;
      usage?: OoxmlResourceUsageSnapshot;
    } & PullSessionIdentity<number>)
  // The worker projects structured-clone-safe run geometry from the same
  // retained layout variant it paints and ships it beside the bitmap.
  | { type: 'pageRendered'; id: number; bitmap: ImageBitmap; runs: DocxTextRunInfo[] }
  | { type: 'runsCollected'; id: number; runs: DocxTextRunInfo[] }
  // The selected variant's geometry, for the host to install atomically with
  // the active-view record it belongs to.
  | { type: 'layoutViewSelected'; id: number; meta: DocumentLayoutMeta }
  // Progressive layout in worker mode: the worker publishes its provisional
  // prefixes as they are primed, then answers the original `parse` with the
  // authoritative `parsedMeta`. Keyed by `forId` rather than `id` ON PURPOSE —
  // `WorkerBridge` resolves a pending request on the FIRST response whose
  // `correlate()` returns its id, so an `id`-keyed partial would settle the
  // parse early and the authoritative metadata would arrive with nowhere to go.
  // An uncorrelated message is routed to the bridge's `onUnsolicited` hook
  // instead, which is exactly the push channel this needs.
  | { type: 'layoutPartial'; forId: number; partial: DocumentLayoutPartial }
  // Throttled liveness + progress heartbeat during a progressive parse. Also
  // load-bearing for safety: a document that never publishes a preview (one
  // short enough that previewing is pointless, or whose preview attempt threw)
  // would otherwise be silent between `parse` and `parsedMeta`, and the host
  // has given up its request timeout for the duration of a background layout.
  // This is what tells it the worker is alive rather than wedged.
  | { type: 'layoutProgress'; forId: number; committedPages: number }
  | { type: 'elementHit'; id: number; context: DocxElementContext | null };
