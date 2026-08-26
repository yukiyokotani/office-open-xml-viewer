import type { PptxComment } from './types.js';

export function pptxCommentOccurrenceKey(
  comment: Readonly<PptxComment>,
  index: number,
  slideIndex: number,
): string {
  const source = comment.id ?? `classic:${comment.authorId ?? 'unknown'}:${comment.index ?? index}`;
  return `slide:${slideIndex}:${source}`;
}
