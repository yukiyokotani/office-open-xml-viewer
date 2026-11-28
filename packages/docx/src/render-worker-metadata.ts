import { buildBookmarkPageMap } from './bookmark-nav.js';
import { collectLayoutSourceCommentRangesIfPresent, type CommentAnchorPlacement } from './comments.js';
import { textRunsForPage } from './text-run-projection.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';
import {
  reviewProjectionIndexForDocument,
} from './layout/text-index.js';
import { collectLayoutSourceRevisionRangesIfPresent } from './revisions.js';
import type { RetainedRenderWorkerDocumentLayout } from './render-worker-layout.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import type { DeepReadonly, DocumentLayout } from './layout/types.js';
import type { DocumentLayoutMeta, DocumentMeta } from './worker-protocol.js';

export type RenderWorkerReviewIndexInput = Pick<DocumentMeta, 'comments' | 'revisions'>;

/** The single worker-side projection from an immutable selected layout to the
 * synchronous geometry and indexes exposed by the host. Initial parse and
 * runtime view switching both delegate here. */
export function projectRenderWorkerLayoutMeta(
  layout: DeepReadonly<DocumentLayout>,
  source: LayoutSourceStore,
  review: RenderWorkerReviewIndexInput,
  options: Readonly<{ provisional?: boolean }> = {},
): DocumentLayoutMeta {
  const hasComments = review.comments.length > 0;
  const hasRevisions = review.revisions.length > 0;
  const reviewIndex = hasComments || hasRevisions
    ? reviewProjectionIndexForDocument(layout)
    : undefined;
  const reviewProjection = reviewIndex
    ? {
        ...(options.provisional ? { completedSourceKeys: reviewIndex.completedSourceKeys } : {}),
        // Each comment's card owner is decided once against this layout.
        cardOwners: {
          sourcePages: reviewIndex.sourcePages,
          runsForPage: (pageIndex: number) => textRunsForPage(layout as DocumentLayout, pageIndex, { scale: 1 }),
        },
        splitRuns: reviewIndex.splitRuns,
        textRunIndex: reviewIndex.textRunIndex,
        // A placement is kept beside its range, so it is sent beside the
        // ranges: the collector fills this as it returns them.
        commentAnchorPlacements: [] as (CommentAnchorPlacement | null)[],
      }
    : undefined;
  return {
    diagnostics: layout.diagnostics,
    pageCount: layout.pages.length,
    pageSizes: layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })),
    bookmarkPages: [...buildBookmarkPageMap(layout)],
    commentAnchorRanges: hasComments
      ? collectLayoutSourceCommentRangesIfPresent(
          review.comments,
          source,
          reviewIndex!.renderedRunIndex,
          reviewProjection,
        )
      : [],
    // Evaluated after the ranges above. Always present: a view switch
    // replaces the previous view's placements.
    commentAnchorPlacements: reviewProjection?.commentAnchorPlacements ?? [],
    revisionAnchorRanges: hasRevisions
      ? collectLayoutSourceRevisionRangesIfPresent(
          review.revisions,
          source,
          reviewIndex!.renderedRunIndex,
          reviewProjection,
        )
      : [],
  };
}

/** Select a runtime worker variant through the canonical store, then project
 * it through the same metadata authority used by initial parse. */
export function renderWorkerLayoutMeta(
  doc: RetainedRenderWorkerDocumentLayout,
  review: RenderWorkerReviewIndexInput,
  currentDateMs: number,
  showTrackedChanges: boolean,
): DocumentLayoutMeta {
  const source = layoutSourceStoreOf(doc.layoutServices);
  if (!source) throw new Error('Document layout source is not initialized');
  const layout = doc.layoutVariants.layoutFor(normalizeLayoutOptions(
    currentDateMs,
    doc.defaultCurrentDateMs,
    showTrackedChanges,
  ));
  return projectRenderWorkerLayoutMeta(layout, source, review);
}

const COMMENT_ANCHORS_BY_LAYOUT = new WeakMap<object, Pick<
  DocumentLayoutMeta, 'commentAnchorRanges' | 'commentAnchorPlacements'
>>();

/** The comment anchors of one retained worker variant, placed against it,
 * without selecting it: the host joins them to runs it collects for a view
 * other than the active one. Kept with the layout. */
export function renderWorkerCommentAnchors(
  doc: RetainedRenderWorkerDocumentLayout,
  review: RenderWorkerReviewIndexInput,
  currentDateMs: number,
  showTrackedChanges: boolean,
): Pick<DocumentLayoutMeta, 'commentAnchorRanges' | 'commentAnchorPlacements'> {
  const source = layoutSourceStoreOf(doc.layoutServices);
  if (!source) throw new Error('Document layout source is not initialized');
  const layout = doc.layoutVariants.layoutFor(normalizeLayoutOptions(
    currentDateMs,
    doc.defaultCurrentDateMs,
    showTrackedChanges,
  ));
  let anchors = COMMENT_ANCHORS_BY_LAYOUT.get(layout);
  if (!anchors) {
    const meta = projectRenderWorkerLayoutMeta(layout, source, { comments: review.comments, revisions: [] });
    anchors = {
      commentAnchorRanges: meta.commentAnchorRanges,
      commentAnchorPlacements: meta.commentAnchorPlacements,
    };
    COMMENT_ANCHORS_BY_LAYOUT.set(layout, anchors);
  }
  return anchors;
}
