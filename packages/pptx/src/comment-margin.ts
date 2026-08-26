import {
  EMU_PER_PX,
  overlayPercent,
  type ViewerCommentsOptions,
  type ViewerCommentConnectorOptions,
} from '@silurus/ooxml-core';
import {
  buildReadOnlyCommentMargin,
  createReadOnlyCommentMarker,
  ensureReadOnlyCommentStyles,
  READ_ONLY_COMMENT_MARKER_SIZE_PX,
  type ReadOnlyCommentMessage,
  type ReadOnlyCommentThread,
} from '@silurus/ooxml-core/internal/read-only-comment-margin';
import type {
  ReadOnlyCommentMarginGeometry,
  ReadOnlyCommentRect,
  ReadOnlyCommentThreadGeometry,
} from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import { relativeElementRect } from '@silurus/ooxml-core/internal/dom-geometry';
import type { PptxElementBounds } from './element-selection.js';
import type { PptxComment, PptxCommentReply } from './types.js';
import { pptxCommentOccurrenceKey } from './comment-occurrence.js';

export interface PptxCommentsOptions extends ViewerCommentsOptions {
  /** Show the built-in margin cards. Set false for an application-owned list. Default true. */
  readonly cards?: boolean;
  /** Margin side. `auto` follows the Viewer container's CSS direction. Default `auto`. */
  readonly side?: 'auto' | 'left' | 'right';
  /** Show authored comment glyphs on the slide. Default true. */
  readonly markers?: boolean;
  /** Draw marker-to-card connectors with the requested geometry. Default none. */
  readonly connectors?: ViewerCommentConnectorOptions;
}

function toReplyCard(
  reply: Readonly<PptxCommentReply>,
  occurrenceKey: string,
  index: number,
): ReadOnlyCommentMessage {
  return {
    messageKey: `${occurrenceKey}:reply:${reply.id ?? index}`,
    sourceId: reply.id,
    author: reply.author,
    date: reply.date,
    text: reply.text,
    status: reply.status,
  };
}

export function buildPptxCommentMargin(
  markerLayer: HTMLDivElement,
  margin: HTMLDivElement | null,
  comments: readonly Readonly<PptxComment>[],
  elementBounds: readonly Readonly<PptxElementBounds>[],
  slideIndex: number,
  slideWidthEmu: number,
  slideHeightEmu: number,
  activeId: string | null,
  onSetActive: (id: string, active: boolean) => void,
  zoom: number,
  logicalMarginWidth: number,
  showMarkers: boolean,
  includeResolved = false,
  onGeometryChange?: () => void,
  onScrollGeometryChange?: () => void,
): ReadOnlyCommentMarginGeometry {
  ensureReadOnlyCommentStyles(markerLayer.ownerDocument);
  if (margin) margin.dataset.ooxmlCommentZoom = String(zoom);
  markerLayer.replaceChildren();
  const visible = comments
    .map((comment, index) => ({
      comment,
      index,
      id: pptxCommentOccurrenceKey(comment, index, slideIndex),
    }))
    .filter(({ comment }) => includeResolved ||
      (comment.status !== 'resolved' && comment.status !== 'closed'));
  const anchorRects = new Map<string, ReadOnlyCommentRect>();
  const markersById = new Map<string, HTMLButtonElement>();
  const boundsByElementId = new Map(elementBounds.map((entry) => [entry.elementId, entry]));
  const surface = markerLayer.parentElement;

  for (const [visibleIndex, entry] of visible.entries()) {
    const { comment, id } = entry;
    const targetBounds = (comment.anchors ?? []).flatMap((anchor) => {
      if ((anchor.type !== 'drawingElement' && anchor.type !== 'textRange') || !anchor.elementId) {
        return [];
      }
      const target = boundsByElementId.get(anchor.elementId);
      return target ? [target] : [];
    });
    if (activeId === id) {
      for (const target of targetBounds) {
        const frame = markerLayer.ownerDocument.createElement('div');
        frame.dataset.ooxmlCommentTarget = id;
        frame.style.cssText =
          `left:${overlayPercent(target.bounds.x, slideWidthEmu)};` +
          `top:${overlayPercent(target.bounds.y, slideHeightEmu)};` +
          `width:${overlayPercent(target.bounds.width, slideWidthEmu)};` +
          `height:${overlayPercent(target.bounds.height, slideHeightEmu)};` +
          `border-width:${2 * zoom}px;transform:rotate(${target.bounds.rotation}deg);`;
        markerLayer.appendChild(frame);
      }
    }
    const firstTarget = targetBounds[0]?.bounds;
    const hasPosition = Number.isFinite(comment.x) && Number.isFinite(comment.y);
    if (!hasPosition && !firstTarget) continue;
    const left = Math.max(0, Math.min(
      firstTarget
        ? firstTarget.x + (hasPosition ? comment.x as number : firstTarget.width)
        : comment.x as number,
      slideWidthEmu,
    ));
    const top = Math.max(0, Math.min(
      firstTarget
        ? firstTarget.y + (hasPosition ? comment.y as number : 0)
        : comment.y as number,
      slideHeightEmu,
    ));
    anchorRects.set(id, Object.freeze({
      x: left / EMU_PER_PX * zoom - READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom / 2,
      y: top / EMU_PER_PX * zoom - READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom / 2,
      width: READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom,
      height: READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom,
    }));
    if (activeId === id && targetBounds.length === 0 && !showMarkers) {
      const frame = markerLayer.ownerDocument.createElement('div');
      frame.dataset.ooxmlCommentTarget = id;
      frame.style.cssText =
        `left:${overlayPercent(left, slideWidthEmu)};` +
        `top:${overlayPercent(top, slideHeightEmu)};` +
        `width:${READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom}px;` +
        `height:${READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom}px;` +
        `border-width:${2 * zoom}px;border-radius:50%;transform:translate(-50%,-50%);`;
      markerLayer.appendChild(frame);
    }
    if (showMarkers) {
      const marker = createReadOnlyCommentMarker(markerLayer.ownerDocument, {
        occurrenceKey: id,
        visibleIndex,
        author: comment.author,
        active: activeId === id,
        zoom,
        onSetActive,
      });
      marker.style.left = overlayPercent(left, slideWidthEmu);
      marker.style.top = overlayPercent(top, slideHeightEmu);
      markerLayer.appendChild(marker);
      markersById.set(id, marker);
    }
  }

  const cardThreads: ReadOnlyCommentThread[] = visible.map(({ comment, id }) => ({
    occurrenceKey: id,
    root: {
      messageKey: `${id}:root`,
      sourceId: comment.id,
      author: comment.author,
      date: comment.date,
      text: comment.text,
      status: comment.status,
    },
    replies: comment.replies?.map((reply, index) => toReplyCard(reply, id, index)) ?? [],
  }));
  const cardHosts = margin
    ? buildReadOnlyCommentMargin(margin, cardThreads, {
        activeId,
        zoom,
        logicalWidth: logicalMarginWidth,
        onSetActive,
        onGeometryChange,
        onScrollGeometryChange,
        preferredTopById: new Map(cardThreads.map((thread) => {
          const anchor = anchorRects.get(thread.occurrenceKey);
          return [thread.occurrenceKey, anchor?.y ?? 0] as const;
        })),
      })
    : new Map<string, HTMLButtonElement>();
  for (const entry of visible) {
    const card = cardHosts.get(entry.id);
    const marker = markersById.get(entry.id);
    if (card?.id && marker) marker.setAttribute('aria-controls', card.id);
  }
  if (!onGeometryChange && !onScrollGeometryChange) {
    return Object.freeze({
      threads: Object.freeze(cardThreads.map((thread): ReadOnlyCommentThreadGeometry => {
        const anchorRect = anchorRects.get(thread.occurrenceKey);
        return Object.freeze({
          occurrenceKey: thread.occurrenceKey,
          active: activeId === thread.occurrenceKey,
          anchorRects: Object.freeze(anchorRect ? [anchorRect] : []),
        });
      })),
      scrollTop: margin?.scrollTop ?? 0,
    });
  }
  const marginRect = margin && surface ? relativeElementRect(margin, surface) : undefined;
  const geometry = Object.freeze(cardThreads.map((thread): ReadOnlyCommentThreadGeometry => {
    const cardHost = cardHosts.get(thread.occurrenceKey);
    const cardRect = cardHost && surface
      ? relativeElementRect(cardHost, surface)
      : undefined;
    const anchorRect = anchorRects.get(thread.occurrenceKey);
    return Object.freeze({
      occurrenceKey: thread.occurrenceKey,
      active: activeId === thread.occurrenceKey,
      anchorRects: Object.freeze(anchorRect ? [anchorRect] : []),
      ...(cardRect ? { cardRect } : {}),
    });
  }));
  return Object.freeze({
    threads: geometry,
    ...(marginRect ? { cardClipBounds: marginRect } : {}),
    scrollTop: margin?.scrollTop ?? 0,
  });
}
