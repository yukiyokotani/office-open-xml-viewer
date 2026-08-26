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

/** Serializable subset of RenderPageOptions (callbacks cannot cross the wire). */
export type WireRenderPageOptions = Omit<RenderPageOptions, 'onTextRun'>;

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; resourcePolicy: NormalizedOoxmlResourcePolicy; useGoogleFonts?: boolean; defaultCurrentDateMs: number; renderers?: WorkerRendererDescriptors }
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
  | { type: 'elementHit'; id: number; context: DocxElementContext | null };
