/** Scope for projecting review anchors from a retained layout. */
export interface ReviewAnchorProjectionOptions {
  /** When present, provisional projection may use fallback geometry only for
   * paragraphs whose final fragment is already in the published prefix.
   * Omit for authoritative projection. */
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
  /** Filled with the placement of each collected comment range, in their
   * order (comments.ts CommentAnchorPlacement; null for a range without
   * one). The render worker sends them beside the ranges. Comment anchors
   * only. */
  readonly commentAnchorPlacements?: (import('./comments.js').CommentAnchorPlacement | null)[];
}
