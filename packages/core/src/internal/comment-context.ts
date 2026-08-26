import type {
  ViewerCommentMessageContext,
  ViewerCommentThreadContext,
} from '../comment-ui.js';

export const MAX_COMMENT_CONTEXT_TEXT_CHARACTERS = 65_536;

export interface BoundedViewerCommentThreadContext {
  readonly thread: ViewerCommentThreadContext;
  readonly truncated: boolean;
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

function boundedMaximum(value: number | undefined): number {
  if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  return Math.min(
    MAX_COMMENT_CONTEXT_TEXT_CHARACTERS,
    Math.floor(value ?? MAX_COMMENT_CONTEXT_TEXT_CHARACTERS),
  );
}

/** Bound only authored message text; identifiers and source locations remain intact. */
export function boundedViewerCommentThreadContext(
  root: ViewerCommentMessageContext,
  replies: readonly ViewerCommentMessageContext[],
  requestedMaximum?: number,
): BoundedViewerCommentThreadContext {
  const maxTextCharacters = boundedMaximum(requestedMaximum);
  let textCharacters = 0;
  let truncated = false;
  const bound = (message: ViewerCommentMessageContext): ViewerCommentMessageContext => {
    const text = safeUtf16Prefix(message.text, Math.max(0, maxTextCharacters - textCharacters));
    textCharacters += text.length;
    if (text.length < message.text.length) truncated = true;
    return { ...message, text };
  };
  const thread = Object.freeze({
    root: Object.freeze(bound(root)),
    replies: Object.freeze(replies.map((reply) => Object.freeze(bound(reply)))),
  });
  return Object.freeze({ thread, truncated, textCharacters, maxTextCharacters });
}
