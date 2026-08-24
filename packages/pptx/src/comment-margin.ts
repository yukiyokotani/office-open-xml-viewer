import { overlayPercent } from '@silurus/ooxml-core';
import {
  buildReadOnlyCommentMargin,
  type ReadOnlyCommentCard,
  type ReadOnlyCommentCardRenderer,
} from '@silurus/ooxml-core/internal/read-only-comment-margin';
import type { PptxComment, PptxCommentReply } from './types.js';

export interface PptxCommentCardRenderContext {
  readonly comment: Readonly<PptxComment>;
  readonly replies: readonly Readonly<PptxCommentReply>[];
  readonly active: boolean;
  readonly activate: () => void;
}

export type PptxCommentCardRenderer = (
  host: HTMLElement,
  context: PptxCommentCardRenderContext,
) => void | (() => void);

function commentId(comment: Readonly<PptxComment>, index: number): string {
  return comment.id ?? `classic:${comment.authorId ?? 'unknown'}:${comment.index ?? index}`;
}

function toReplyCard(reply: Readonly<PptxCommentReply>, index: number): ReadOnlyCommentCard {
  return {
    id: reply.id ?? `reply:${index}`,
    author: reply.author,
    date: reply.date,
    text: reply.text,
  };
}

export function buildPptxCommentMargin(
  markerLayer: HTMLDivElement,
  margin: HTMLDivElement,
  comments: readonly Readonly<PptxComment>[],
  slideWidthEmu: number,
  slideHeightEmu: number,
  activeId: string | null,
  onActivate: (id: string | null) => void,
  renderCard?: PptxCommentCardRenderer,
): void {
  markerLayer.replaceChildren();
  const visible = comments
    .map((comment, index) => ({ comment, index, id: commentId(comment, index) }))
    .filter(({ comment }) => comment.status !== 'resolved' && comment.status !== 'closed');

  for (const [visibleIndex, entry] of visible.entries()) {
    const { comment, id } = entry;
    if (!Number.isFinite(comment.x) || !Number.isFinite(comment.y)) continue;
    const marker = markerLayer.ownerDocument.createElement('button');
    marker.type = 'button';
    marker.dataset.ooxmlCommentId = id;
    marker.setAttribute('aria-label', `Comment ${visibleIndex + 1}`);
    marker.setAttribute('aria-pressed', String(activeId === id));
    const left = Math.max(0, Math.min(comment.x as number, slideWidthEmu));
    const top = Math.max(0, Math.min(comment.y as number, slideHeightEmu));
    marker.style.cssText =
      'position:absolute;transform:translate(-50%,-50%);width:22px;height:22px;' +
      'padding:0;border:2px solid #fff;border-radius:999px;cursor:pointer;pointer-events:auto;' +
      'font:600 11px/18px system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;' +
      'color:#fff;background:var(--ooxml-comment-marker-background,#2563eb);' +
      `left:${overlayPercent(left, slideWidthEmu)};top:${overlayPercent(top, slideHeightEmu)};` +
      `box-shadow:${activeId === id ? '0 0 0 3px rgba(37,99,235,.35)' : '0 1px 3px rgba(15,23,42,.28)'};`;
    marker.textContent = String(visibleIndex + 1);
    marker.addEventListener('click', () => onActivate(activeId === id ? null : id));
    markerLayer.appendChild(marker);
  }

  const cards: ReadOnlyCommentCard[] = visible.map(({ comment, id }) => ({
    id,
    author: comment.author,
    date: comment.date,
    text: comment.text,
    replies: comment.replies?.map(toReplyCard),
  }));
  const byId = new Map(visible.map((entry) => [entry.id, entry.comment]));
  const sharedRenderer: ReadOnlyCommentCardRenderer | undefined = renderCard
    ? (host, context) => {
        const comment = byId.get(context.comment.id);
        if (!comment) return;
        return renderCard(host, {
          comment,
          replies: comment.replies ?? [],
          active: context.active,
          activate: context.activate,
        });
      }
    : undefined;
  buildReadOnlyCommentMargin(margin, cards, activeId, onActivate, sharedRenderer);
}
