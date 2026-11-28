/** Scope for projecting review anchors from a retained layout. */
export interface ReviewAnchorProjectionOptions {
  /** The paragraphs that have ended in a published prefix. When present the
   * projection is provisional. A revision may use fallback geometry only
   * for such paragraphs. A comment is published only when every paragraph
   * of its authored extent is one of them and its stand-in text is decided
   * (comments.ts): the prefix then places it as the complete layout will,
   * and a comment whose extent goes on beyond the prefix appears later, on
   * purpose, rather than somewhere it will not stay. Omit for authoritative
   * projection. */
  readonly completedSourceKeys?: ReadonlySet<string>;
  /** The laid-out document to decide comment card owners against
   * (comments.ts assignCommentCardOwners): the pages that hold runs of each
   * paragraph source, and the projected runs of a page. Comment anchors
   * only. */
  readonly cardOwners?: Readonly<{
    sourcePages: ReadonlyMap<string, readonly number[]>;
    runsForPage: (pageIndex: number) => readonly Readonly<import('./types.js').DocxTextRunInfo>[];
  }>;
  /** The runs of the laid-out document that cross a page break
   * (text-index.ts ReviewProjectionIndex.splitRuns): a comment shown at one
   * end of such a run stands on the page that holds that end. Comment
   * anchors only. */
  readonly splitRuns?: import('./layout/text-index.js').SplitRunIndex;
  /** The laid-out runs that hold text (text-index.ts
   * ReviewProjectionIndex.textRunIndex), where the rendered run index given
   * beside these options also lists tabs and inline objects. Text that
   * stands in for content without geometry (a revision's fallback, a
   * comment's adjacent-paragraph fallback) is chosen among these. */
  readonly textRunIndex?: ReadonlyMap<string, ReadonlySet<number>>;
  /** Filled with the placement of each collected comment range, in their
   * order (comments.ts CommentAnchorPlacement; null for a range without
   * one). The render worker sends them beside the ranges. Comment anchors
   * only. */
  readonly commentAnchorPlacements?: (import('./comments.js').CommentAnchorPlacement | null)[];
}

const PROVISIONAL_LAYOUTS = new WeakSet<object>();

/** Record that `layout` is a prefix published while its document is still
 * being paginated. The mark belongs to the layout: whichever route reads it
 * later (its publication, a view selection, a runs request, in the worker or
 * on the main thread) projects its review anchors as from a prefix, and a
 * complete layout of another view is projected whole while this one loads. */
export function markProvisionalLayout(layout: object): void {
  PROVISIONAL_LAYOUTS.add(layout);
}

export function isProvisionalLayout(layout: object): boolean {
  return PROVISIONAL_LAYOUTS.has(layout);
}
