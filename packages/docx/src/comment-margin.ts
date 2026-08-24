import { overlayPercent, type ViewerCommentCardRenderContext } from '@silurus/ooxml-core';
import {
  buildReadOnlyCommentMargin,
  type ReadOnlyCommentCard,
  type ReadOnlyCommentCardRenderer,
} from '@silurus/ooxml-core/internal/read-only-comment-margin';
import { resolveCommentAnchorRuns, type CommentAnchorRange } from './comments.js';
import type { DocxTextRunInfo } from './renderer.js';
import type { DocComment } from './types.js';

const COMMENT_TINT = 'rgba(59, 130, 246, 0.18)';
const ACTIVE_COMMENT_TINT = 'rgba(37, 99, 235, 0.34)';

interface CommentThread {
  readonly root: DocComment;
  readonly replies: readonly DocComment[];
}

export interface DocxCommentMarginModel {
  readonly comments: readonly DocComment[];
  readonly anchors: readonly CommentAnchorRange[];
}

/** Context supplied when a consumer replaces the built-in comment card.
 * ScrollViewer still owns placement, virtualization, activation, and cleanup. */
export interface DocxCommentCardRenderContext extends ViewerCommentCardRenderContext {
  readonly comment: Readonly<DocComment>;
  readonly replies: readonly Readonly<DocComment>[];
}

export type DocxCommentCardRenderer = (
  host: HTMLElement,
  context: DocxCommentCardRenderContext,
) => void | (() => void);

function commentThreads(comments: readonly DocComment[]): CommentThread[] {
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const roots = comments.filter((comment) => comment.parentId === undefined);
  const replies = new Map<string, DocComment[]>();
  for (const comment of comments) {
    if (comment.parentId === undefined) continue;
    let current = comment;
    const seen = new Set<string>([current.id]);
    while (current.parentId !== undefined) {
      const parent = byId.get(current.parentId);
      if (!parent || seen.has(parent.id)) {
        current = comment;
        break;
      }
      seen.add(parent.id);
      current = parent;
    }
    if (current.parentId !== undefined || current === comment) continue;
    const list = replies.get(current.id) ?? [];
    if (!replies.has(current.id)) replies.set(current.id, list);
    list.push(comment);
  }
  return roots
    .filter((root) => root.resolved !== true)
    .map((root) => ({ root, replies: Object.freeze(replies.get(root.id) ?? []) }));
}

function toCard(comment: DocComment): ReadOnlyCommentCard {
  return {
    id: comment.id,
    author: comment.author,
    date: comment.date,
    text: comment.paragraphs?.join('\n') ?? comment.text,
  };
}

function createTint(
  layer: HTMLDivElement,
  run: Readonly<DocxTextRunInfo>,
  cssWidth: number,
  cssHeight: number,
  active: boolean,
): void {
  const tint = layer.ownerDocument.createElement('div');
  tint.style.cssText =
    'position:absolute;pointer-events:none;' +
    `left:${overlayPercent(run.x, cssWidth)};top:${overlayPercent(run.y, cssHeight)};` +
    `width:${overlayPercent(run.w, cssWidth)};height:${overlayPercent(run.h, cssHeight)};` +
    `background:${active ? ACTIVE_COMMENT_TINT : COMMENT_TINT};`;
  if (run.transform) {
    tint.style.transform = run.transform;
    tint.style.transformOrigin = 'top left';
  }
  layer.appendChild(tint);
}

/** Build one page's range tint and authored-order margin cards. A thread card is
 * emitted on the page containing its first structural anchor; later ranges are
 * still tinted but do not duplicate the card. */
export function buildDocxCommentMargin(
  tintLayer: HTMLDivElement,
  margin: HTMLDivElement,
  runs: readonly Readonly<DocxTextRunInfo>[],
  model: DocxCommentMarginModel,
  cssWidth: number,
  cssHeight: number,
  activeId: string | null,
  onActivate: (id: string | null) => void,
  zoom: number,
  renderCard?: DocxCommentCardRenderer,
): void {
  margin.dataset.ooxmlCommentZoom = String(zoom);
  tintLayer.innerHTML = '';
  const threads = commentThreads(model.comments);
  const firstAnchor = new Map<string, CommentAnchorRange>();
  for (const anchor of model.anchors) {
    if (!firstAnchor.has(anchor.commentId)) firstAnchor.set(anchor.commentId, anchor);
    const active = activeId === anchor.commentId;
    for (const run of resolveCommentAnchorRuns(anchor, runs)) {
      createTint(tintLayer, run, cssWidth, cssHeight, active);
    }
  }
  const visibleThreads = new Map<string, CommentThread>();
  const cards = threads.flatMap((thread): ReadOnlyCommentCard[] => {
    const anchor = firstAnchor.get(thread.root.id);
    if (!anchor || resolveCommentAnchorRuns(anchor, runs).length === 0) return [];
    visibleThreads.set(thread.root.id, thread);
    return [{
      ...toCard(thread.root),
      replies: thread.replies.map(toCard),
    }];
  });
  const sharedRenderer: ReadOnlyCommentCardRenderer | undefined = renderCard
    ? (host, context) => {
        const thread = visibleThreads.get(context.view.id);
        if (!thread) return;
        return renderCard(host, {
          view: context.view,
          comment: thread.root,
          replies: thread.replies,
          active: context.active,
          zoom: context.zoom,
          activate: context.activate,
        });
      }
    : undefined;
  buildReadOnlyCommentMargin(margin, cards, activeId, onActivate, sharedRenderer);
}
