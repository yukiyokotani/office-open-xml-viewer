import { buildBookmarkPageMap } from './bookmark-nav.js';
import { collectLayoutSourceCommentRangesIfPresent } from './comments.js';
import { collectLayoutSourceRevisionRangesIfPresent } from './revisions.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { layoutSourceStoreOf } from './layout/runtime-state.js';
import { textRunSourceIndexForDocument } from './layout/text-index.js';
import type { LayoutSourceStore } from './layout/layout-source-store.js';
import type { DeepReadonly, DocumentLayout } from './layout/types.js';
import type { RetainedRenderWorkerDocumentLayout } from './render-worker-layout.js';
import type { DocumentLayoutMeta, DocumentMeta } from './worker-protocol.js';

/**
 * Worker-side seam for the RUNTIME layout-view switch (`selectLayoutView`).
 *
 * ECMA-376 §17.13.5 makes the tracked-change view an acquisition input, not a
 * paint flag: hiding deletions changes line breaking, so the two views are
 * genuinely different paginations with different page counts and different
 * per-page geometry. In `mode: 'worker'` the model never crosses the wire, so
 * `DocxDocument`'s geometry accessors read worker metadata — which means a
 * switch that moved the active view without bringing the selected variant's
 * metadata back left the host measuring one pagination while painting another.
 *
 * Kept out of `render-worker.ts` so it can be driven directly by a test: that
 * module reaches WASM and `self` at import time. `parse` keeps building its own
 * metadata inline (its route is pinned by the layout-boundary gate); the two
 * share the anchor projection below, and `render-worker-layout-view.test.ts`
 * pins the remaining fields against each other so the routes cannot drift.
 */

/**
 * The §17.13.4 / §17.13.5 anchor projections for one built layout.
 *
 * Both are whole-document joins over the run index of the layout that is
 * actually being painted, so they belong to a variant exactly as much as the
 * page count does.
 */
export function renderWorkerReviewAnchors(
  layout: DeepReadonly<DocumentLayout>,
  source: LayoutSourceStore,
  review: Pick<DocumentMeta, 'comments' | 'revisions'>,
): Pick<DocumentMeta, 'commentAnchorRanges' | 'revisionAnchorRanges'> {
  const renderedRunIndex = textRunSourceIndexForDocument(layout);
  return {
    commentAnchorRanges: collectLayoutSourceCommentRangesIfPresent(
      review.comments,
      source,
      renderedRunIndex,
    ),
    revisionAnchorRanges: collectLayoutSourceRevisionRangesIfPresent(
      review.revisions,
      source,
      renderedRunIndex,
    ),
  };
}

/**
 * Build the selected variant and report the layout-derived metadata for it.
 *
 * The variant comes from the same retained store `parse` and every render
 * request select from, so a switched view is served the identical pagination
 * the next `renderPage` will paint. Only the layout-derived half is returned:
 * `comments`, `revisions`, `footnotes` and `endnotes` are model-derived and
 * identical across variants, so the host keeps the ones its load established
 * rather than paying to re-clone them across the wire on every toggle.
 */
export function renderWorkerLayoutViewMeta(
  retained: RetainedRenderWorkerDocumentLayout,
  review: Pick<DocumentMeta, 'comments' | 'revisions'>,
  view: Readonly<{ currentDateMs?: number; showTrackedChanges?: boolean }> = {},
): DocumentLayoutMeta {
  const source = layoutSourceStoreOf(retained.layoutServices);
  if (!source) throw new Error('Document layout source is not initialized');
  const selected = retained.layoutVariants.layoutFor(
    normalizeLayoutOptions(
      view.currentDateMs,
      retained.defaultCurrentDateMs,
      view.showTrackedChanges,
    ),
  );
  return {
    pageCount: selected.pages.length,
    pageSizes: selected.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })),
    bookmarkPages: [...buildBookmarkPageMap(selected)],
    ...renderWorkerReviewAnchors(selected, source, review),
  };
}
