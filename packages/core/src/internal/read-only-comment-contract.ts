/** Static data contract shared by the viewer and the lazily loaded built-in UI. */
export const READ_ONLY_COMMENT_MARGIN_WIDTH_PX = 280;
export const READ_ONLY_COMMENT_MARKER_SIZE_PX = 24;

export interface ReadOnlyCommentMessage {
  readonly messageKey: string;
  readonly sourceId?: string;
  readonly author?: string;
  readonly date?: string;
  readonly text: string;
  readonly status?: 'active' | 'resolved' | 'closed';
}

export interface ReadOnlyCommentThread {
  readonly occurrenceKey: string;
  readonly root: ReadOnlyCommentMessage;
  readonly replies: readonly ReadOnlyCommentMessage[];
}
