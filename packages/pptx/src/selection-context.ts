import type {
  TextSelectionContextOptions,
  ViewerCommentThreadContext,
} from '@silurus/ooxml-core';
import { boundedViewerCommentThreadContext } from '@silurus/ooxml-core/internal/comment-context';
import type { PptxComment, SlideElementOrigin } from './types.js';
import { readBoundedNativeTextSelection } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';

/** Snapshot-local locator for a rendered run intersecting a PPTX text selection. */
export interface PptxSelectionRunLocator {
  readonly slideIndex: number;
  readonly runIndex: number;
  readonly shapeId?: string;
  readonly elementIndex?: number;
  readonly origin?: SlideElementOrigin;
}

export interface PptxTextSelectionContext {
  readonly format: 'pptx';
  readonly kind: 'text';
  readonly text: string;
  readonly slideIndexes: readonly number[];
  readonly shapeIds: readonly string[];
  readonly runs: readonly PptxSelectionRunLocator[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text' | 'runs')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
  readonly maxRunLocators: number;
}

/** Detached selected slide-comment thread for read-only AI/MCP handoff. */
export interface PptxCommentSelectionContext {
  readonly format: 'pptx';
  readonly kind: 'comment';
  readonly slideIndex: number;
  readonly commentIndex: number;
  readonly occurrenceId: string;
  readonly commentId?: string;
  readonly point?: Readonly<{ x: number; y: number }>;
  readonly thread: ViewerCommentThreadContext;
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

export function createPptxCommentSelectionContext(
  comment: Readonly<PptxComment>,
  slideIndex: number,
  commentIndex: number,
  occurrenceId: string,
  options: TextSelectionContextOptions = {},
): PptxCommentSelectionContext {
  const bounded = boundedViewerCommentThreadContext(
    {
      id: comment.id,
      author: comment.author,
      date: comment.date,
      text: comment.text,
      status: comment.status ?? 'active',
    },
    (comment.replies ?? []).map((reply) => ({
      id: reply.id,
      author: reply.author,
      date: reply.date,
      text: reply.text,
      status: reply.status ?? 'active',
    })),
    options.maxTextCharacters,
  );
  return Object.freeze({
    format: 'pptx',
    kind: 'comment',
    slideIndex,
    commentIndex,
    occurrenceId,
    ...(comment.id ? { commentId: comment.id } : {}),
    ...(Number.isFinite(comment.x) && Number.isFinite(comment.y)
      ? { point: Object.freeze({ x: comment.x as number, y: comment.y as number }) }
      : {}),
    thread: bounded.thread,
    truncated: bounded.truncated,
    truncationReasons: bounded.truncated ? ['text'] as const : [],
    textCharacters: bounded.textCharacters,
    maxTextCharacters: bounded.maxTextCharacters,
  });
}

function nonNegativeInteger(value: string | undefined): number | null {
  if (value === undefined || !/^\d+$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : null;
}

function slideIndexFor(run: HTMLElement): number | null {
  for (let element: HTMLElement | null = run; element; element = element.parentElement) {
    const index = nonNegativeInteger(element.dataset.slideIndex);
    if (index !== null) return index;
  }
  return null;
}

export function readPptxTextSelectionContext(
  root: HTMLElement,
  selection: Selection | null,
  options: TextSelectionContextOptions = {},
): PptxTextSelectionContext | null {
  const bounded = readBoundedNativeTextSelection(root, selection, (run) => {
    const slideIndex = slideIndexFor(run);
    const runIndex = nonNegativeInteger(run.dataset.runIndex);
    if (slideIndex === null || runIndex === null) return null;
    const elementIndex = nonNegativeInteger(run.dataset.elementIndex);
    const origin = run.dataset.elementOrigin;
    const hasElementLocator = elementIndex !== null &&
      (origin === 'master' || origin === 'layout' || origin === 'slide');
    return {
      slideIndex,
      runIndex,
      ...(run.dataset.shapeId === undefined ? {} : { shapeId: run.dataset.shapeId }),
      ...(hasElementLocator ? { elementIndex, origin } : {}),
    } satisfies PptxSelectionRunLocator;
  }, {
    maxChars: options.maxTextCharacters,
    maxLocators: options.maxRunLocators,
  });
  if (!bounded) return null;
  const runs = [...bounded.locators].sort(
    (left, right) => left.slideIndex - right.slideIndex || left.runIndex - right.runIndex,
  );
  return {
    format: 'pptx',
    kind: 'text',
    text: bounded.text,
    slideIndexes: [...new Set(runs.map((run) => run.slideIndex))],
    shapeIds: [...new Set(runs.flatMap((run) => run.shapeId ? [run.shapeId] : []))],
    runs,
    truncated: bounded.truncated,
    truncationReasons: bounded.truncationReasons,
    textCharacters: bounded.textCharacters,
    maxTextCharacters: bounded.maxTextCharacters,
    maxRunLocators: bounded.maxLocators,
  };
}
