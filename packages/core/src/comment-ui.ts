// Comments and change history are separate review features with different
// OOXML source models. The public API does not merge them into one
// cross-format review-item union, and their record and geometry types stay
// format-specific. Revision records are exposed only for DOCX
// (`DocxDocument.revisions`); XLSX and PPTX have no revisions member, not even
// a placeholder, because an empty array would read as "no changes" rather than
// "not implemented". A comparison of two presentations, if ever added, is a
// separate operation and not a revision record of one presentation.

/** Visibility policy shared by the built-in read-only comment UIs. */
export interface ViewerCommentsOptions {
  /** Include resolved or closed threads. DOCX/PPTX default false; XLSX default true. */
  readonly includeResolved?: boolean;
}

export type ViewerCommentConnectorRoute = 'bezier' | 'orthogonal';
export type ViewerCommentConnectorStroke = 'solid' | 'dashed';

/** Appearance of the optional DOCX/PPTX anchor-to-card connectors. */
export interface ViewerCommentConnectorOptions {
  /** Connector route. Default `bezier`. */
  readonly route?: ViewerCommentConnectorRoute;
  /** Connector stroke. Default `solid`. */
  readonly stroke?: ViewerCommentConnectorStroke;
  /** CSS color used for connectors. Also used when active unless `activeColor` is set. */
  readonly color?: string;
  /** Optional CSS color used for the selected connector. */
  readonly activeColor?: string;
}

/** Detached message data shared by format-specific comment contexts. */
export interface ViewerCommentMessageContext {
  readonly id?: string;
  readonly author?: string;
  readonly date?: string;
  readonly text: string;
  readonly status?: 'active' | 'resolved' | 'closed';
}

/** Detached root and replies for a selected comment thread. */
export interface ViewerCommentThreadContext {
  readonly root: ViewerCommentMessageContext;
  readonly replies: readonly ViewerCommentMessageContext[];
}
