import type {
  TextSelectionContextOptions,
  ViewerCommentThreadContext,
} from '@silurus/ooxml-core';
import { boundedViewerCommentThreadContext } from '@silurus/ooxml-core/internal/comment-context';
import { readBoundedNativeTextSelection } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { CommentAnchorRange } from './comments.js';
import type { DocComment } from './types.js';

/** Bounds for a DOCX selection-context snapshot. Extensible per format. */
export type DocxSelectionContextOptions = TextSelectionContextOptions;

export interface DocxSelectionSourceLocator {
  readonly story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';
  readonly storyInstance: string;
  readonly path: readonly number[];
}

/** Snapshot-local locator for a rendered run intersecting a DOCX text selection. */
export interface DocxSelectionRunLocator {
  readonly pageIndex: number;
  readonly runIndex: number;
  readonly paragraphId?: string;
  readonly source?: DocxSelectionSourceLocator;
}

/** Detached, bounded native-text context for a read-only AI/MCP handoff. */
export interface DocxTextSelectionContext {
  readonly format: 'docx';
  readonly kind: 'text';
  readonly text: string;
  readonly pageIndexes: readonly number[];
  readonly paragraphIds: readonly string[];
  readonly runs: readonly DocxSelectionRunLocator[];
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text' | 'runs')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
  readonly maxRunLocators: number;
}

export interface DocxPagePoint {
  readonly xPt: number;
  readonly yPt: number;
}

export interface DocxElementContext {
  readonly format: 'docx';
  readonly kind: 'element';
  readonly pageIndex: number;
  /** Index in the retained page paint snapshot, not a mutable document index. */
  readonly elementIndex: number;
  readonly elementType: 'chart' | 'image' | 'shape';
  readonly point: DocxPagePoint;
  readonly bounds: Readonly<DocxPagePoint & { widthPt: number; heightPt: number }>;
  readonly source: DocxSelectionSourceLocator;
  readonly text?: string;
  readonly mimeType?: string;
  readonly seriesCount?: number;
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

/** Detached selected comment thread for read-only AI/MCP handoff. */
export interface DocxCommentSelectionContext {
  readonly format: 'docx';
  readonly kind: 'comment';
  readonly pageIndex: number;
  readonly commentId: string;
  readonly source?: DocxSelectionSourceLocator;
  readonly thread: ViewerCommentThreadContext;
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

export type DocxSelectionContext =
  | DocxTextSelectionContext
  | DocxElementContext
  | DocxCommentSelectionContext;

export function createDocxCommentSelectionContext(
  comments: readonly Readonly<DocComment>[],
  anchors: readonly Readonly<CommentAnchorRange>[],
  commentId: string,
  pageIndex: number,
  options: DocxSelectionContextOptions = {},
): DocxCommentSelectionContext | null {
  const root = comments.find((comment) => comment.id === commentId && comment.parentId === undefined);
  if (!root) return null;
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const replies = comments.filter((comment) => {
    if (comment.parentId === undefined) return false;
    let current: Readonly<DocComment> | undefined = comment;
    const seen = new Set<string>();
    while (current.parentId !== undefined && !seen.has(current.id)) {
      seen.add(current.id);
      const parent = byId.get(current.parentId);
      if (!parent) return false;
      if (parent.id === root.id) return true;
      current = parent;
    }
    return false;
  });
  const bounded = boundedViewerCommentThreadContext(
    {
      id: root.id,
      author: root.author,
      date: root.date,
      text: root.paragraphs?.join('\n') ?? root.text,
      status: root.resolved ? 'resolved' : 'active',
    },
    replies.map((reply) => ({
      id: reply.id,
      author: reply.author,
      date: reply.date,
      text: reply.paragraphs?.join('\n') ?? reply.text,
      status: reply.resolved ? 'resolved' : 'active',
    })),
    options.maxTextCharacters,
  );
  const source = anchors.find((anchor) => anchor.commentId === commentId)?.source;
  return Object.freeze({
    format: 'docx',
    kind: 'comment',
    pageIndex,
    commentId,
    ...(source ? { source: Object.freeze({ ...source, path: Object.freeze([...source.path]) }) } : {}),
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

function pageIndexFor(run: HTMLElement): number | null {
  for (let element: HTMLElement | null = run; element; element = element.parentElement) {
    const index = nonNegativeInteger(element.dataset.pageIndex);
    if (index !== null) return index;
  }
  return null;
}

const SOURCE_STORIES = new Set<DocxSelectionSourceLocator['story']>([
  'body', 'header', 'footer', 'footnote', 'endnote', 'textbox',
]);

function sourceLocatorFor(run: HTMLElement): DocxSelectionSourceLocator | null {
  const story = run.dataset.sourceStory as DocxSelectionSourceLocator['story'] | undefined;
  const storyInstance = run.dataset.sourceStoryInstance;
  const encodedPath = run.dataset.sourcePath;
  if (!story || !SOURCE_STORIES.has(story) || !storyInstance || !encodedPath) return null;
  try {
    const path: unknown = JSON.parse(encodedPath);
    if (!Array.isArray(path) || path.length === 0 || path.length > 32 ||
      !path.every((index) => Number.isSafeInteger(index) && index >= 0)) return null;
    return { story, storyInstance, path: [...path] as number[] };
  } catch {
    return null;
  }
}

export function readDocxTextSelectionContext(
  root: HTMLElement,
  selection: Selection | null,
  options: DocxSelectionContextOptions = {},
): DocxTextSelectionContext | null {
  const bounded = readBoundedNativeTextSelection(root, selection, (run) => {
    const pageIndex = pageIndexFor(run);
    const runIndex = nonNegativeInteger(run.dataset.runIndex);
    if (pageIndex === null || runIndex === null) return null;
    return {
      pageIndex,
      runIndex,
      ...(run.dataset.paragraphId === undefined ? {} : { paragraphId: run.dataset.paragraphId }),
      ...(sourceLocatorFor(run) === null ? {} : { source: sourceLocatorFor(run)! }),
    } satisfies DocxSelectionRunLocator;
  }, {
    maxChars: options.maxTextCharacters,
    maxLocators: options.maxRunLocators,
  });
  if (!bounded) return null;
  const runs = [...bounded.locators].sort(
    (left, right) => left.pageIndex - right.pageIndex || left.runIndex - right.runIndex,
  );
  return {
    format: 'docx',
    kind: 'text',
    text: bounded.text,
    pageIndexes: [...new Set(runs.map((run) => run.pageIndex))],
    paragraphIds: [...new Set(runs.flatMap((run) => run.paragraphId ? [run.paragraphId] : []))],
    runs,
    truncated: bounded.truncated,
    truncationReasons: bounded.truncationReasons,
    textCharacters: bounded.textCharacters,
    maxTextCharacters: bounded.maxTextCharacters,
    maxRunLocators: bounded.maxLocators,
  };
}
