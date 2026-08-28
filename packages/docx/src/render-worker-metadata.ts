import { buildBookmarkPageMap } from './bookmark-nav.js';
import { collectLayoutSourceCommentRangesIfPresent } from './comments.js';
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
  const reviewProjection = options.provisional && reviewIndex
    ? { completedSourceKeys: reviewIndex.completedSourceKeys }
    : undefined;
  return {
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
