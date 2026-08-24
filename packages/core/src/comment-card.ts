/** Format-neutral view model for one read-only OOXML comment thread. */
export interface ViewerCommentCard {
  readonly id: string;
  readonly author?: string;
  readonly date?: string;
  readonly text: string;
  readonly replies?: readonly ViewerCommentCard[];
}

/** Common context accepted by a card renderer shared between DOCX and PPTX. */
export interface ViewerCommentCardRenderContext {
  readonly view: ViewerCommentCard;
  readonly active: boolean;
  /** Absolute viewer zoom (`1` is the document's natural CSS size). */
  readonly zoom: number;
  readonly activate: () => void;
}

/** Framework-neutral mount hook. Return a cleanup callback when needed. */
export type ViewerCommentCardRenderer = (
  host: HTMLElement,
  context: ViewerCommentCardRenderContext,
) => void | (() => void);
