/** Scope for projecting review anchors from a retained layout. */
export interface ReviewAnchorProjectionOptions {
  /** When present, provisional projection may use fallback geometry only for
   * paragraphs whose final fragment is already in the published prefix.
   * Omit for authoritative projection. */
  readonly completedSourceKeys?: ReadonlySet<string>;
}
