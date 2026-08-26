import {
  overlayPercent,
  type ViewerCommentsOptions,
  type ViewerCommentConnectorOptions,
} from '@silurus/ooxml-core';
import {
  buildReadOnlyCommentMargin,
  createReadOnlyCommentMarker,
  ensureReadOnlyCommentStyles,
  READ_ONLY_COMMENT_MARKER_SIZE_PX,
  readOnlyCommentAuthorAccent,
  type ReadOnlyCommentMessage,
  type ReadOnlyCommentThread,
} from '@silurus/ooxml-core/internal/read-only-comment-margin';
import type {
  ReadOnlyCommentMarginGeometry,
  ReadOnlyCommentRect,
  ReadOnlyCommentThreadGeometry,
} from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import { relativeElementRect } from '@silurus/ooxml-core/internal/dom-geometry';
import { resolveCommentAnchorRuns, type CommentAnchorRange } from './comments.js';
import { sourceKey } from './layout/source-key.js';
import type { DocxTextRunInfo } from './renderer.js';
import type { DocComment } from './types.js';

interface CommentThread {
  readonly root: DocComment;
  readonly replies: readonly DocComment[];
}

export interface DocxCommentMarginModel {
  readonly comments: readonly DocComment[];
  readonly anchors: readonly CommentAnchorRange[];
}

interface ResolvedPageCommentAnchor {
  readonly anchor: Readonly<CommentAnchorRange>;
  readonly runs: readonly Readonly<DocxTextRunInfo>[];
}

/** Resolve only anchors whose authored, reference, or final-state fallback
 * source occurs in this page's projected runs. A mounted page must not pay the
 * `anchors × runs` join cost for comments belonging to every other page. */
export function resolvePageCommentAnchors(
  anchors: readonly Readonly<CommentAnchorRange>[],
  runs: readonly Readonly<DocxTextRunInfo>[],
): readonly ResolvedPageCommentAnchor[] {
  const pageSources = new Set<string>();
  for (const run of runs) {
    if (run.source) pageSources.add(sourceKey(run.source));
  }
  if (pageSources.size === 0) return [];
  const pageAnchors: ResolvedPageCommentAnchor[] = [];
  for (const anchor of anchors) {
    const mayResolve = pageSources.has(sourceKey(anchor.source)) ||
      pageSources.has(sourceKey(anchor.reference.source)) ||
      (anchor.geometryFallback !== undefined &&
        pageSources.has(sourceKey(anchor.geometryFallback.source)));
    if (!mayResolve) continue;
    const resolved = resolveCommentAnchorRuns(anchor, runs);
    if (resolved.length > 0) pageAnchors.push({ anchor, runs: resolved });
  }
  return pageAnchors;
}

export interface DocxCommentsOptions extends ViewerCommentsOptions {
  /** Show the built-in margin cards. Set false for an application-owned list. Default true. */
  readonly cards?: boolean;
  /** Margin side. `auto` follows the Viewer container's CSS direction. Default `auto`. */
  readonly side?: 'auto' | 'left' | 'right';
  /** Show authored comment glyphs beside anchored text. Default true. */
  readonly markers?: boolean;
  /** Draw anchor-to-card connectors with the requested geometry. Default none. */
  readonly connectors?: ViewerCommentConnectorOptions;
}

function commentThreads(comments: readonly DocComment[], includeResolved: boolean): CommentThread[] {
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
    .filter((root) => includeResolved || root.resolved !== true)
    .map((root) => ({ root, replies: Object.freeze(replies.get(root.id) ?? []) }));
}

function toMessage(comment: DocComment, occurrenceKey: string, index: number): ReadOnlyCommentMessage {
  return {
    messageKey: index === 0
      ? `${occurrenceKey}:root`
      : `${occurrenceKey}:reply:${comment.id || index - 1}`,
    sourceId: comment.id,
    author: comment.author,
    date: comment.date,
    text: comment.paragraphs?.join('\n') ?? comment.text,
    status: comment.resolved ? 'resolved' : 'active',
  };
}

function createTint(
  layer: HTMLDivElement,
  run: Readonly<DocxTextRunInfo>,
  cssWidth: number,
  cssHeight: number,
  active: boolean,
  accent: string,
): HTMLDivElement {
  const tint = layer.ownerDocument.createElement('div');
  tint.style.cssText =
    `--ooxml-comment-author-accent:${accent};` +
    `left:${overlayPercent(run.x, cssWidth)};top:${overlayPercent(run.y, cssHeight)};` +
    `width:${overlayPercent(run.w, cssWidth)};height:${overlayPercent(run.h, cssHeight)};`;
  tint.dataset.ooxmlCommentHighlight = '';
  tint.dataset.active = String(active);
  if (run.transform) {
    tint.style.transform = run.transform;
  }
  layer.appendChild(tint);
  return tint;
}

function wordHighlightRun(run: Readonly<DocxTextRunInfo>): DocxTextRunInfo {
  const bounds = run.highlightBounds;
  return {
    ...run,
    x: bounds?.x ?? run.x,
    y: bounds?.y ?? run.y,
    w: bounds?.width ?? run.w,
    h: bounds?.height ?? run.h,
  };
}

/** Fill the whitespace between consecutive anchor runs on one rendered line.
 * Exact line geometry and transform equality are the boundary: this never uses
 * a proximity threshold and never joins different baselines or transforms. */
function mergeSameLineRuns(
  runs: readonly Readonly<DocxTextRunInfo>[],
): Readonly<DocxTextRunInfo>[] {
  const merged: DocxTextRunInfo[] = [];
  for (const run of runs) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.y === run.y &&
      previous.h === run.h &&
      previous.transform === run.transform
    ) {
      const left = Math.min(previous.x, run.x);
      const right = Math.max(previous.x + previous.w, run.x + run.w);
      merged[merged.length - 1] = { ...previous, x: left, w: right - left };
      continue;
    }
    merged.push({ ...run });
  }
  return merged;
}

/** Build one page's range tint and authored-order margin cards. A thread card is
 * emitted on the page containing its first structural anchor; later ranges are
 * still tinted but do not duplicate the card. */
export function buildDocxCommentMargin(
  tintLayer: HTMLDivElement,
  margin: HTMLDivElement | null,
  runs: readonly Readonly<DocxTextRunInfo>[],
  model: DocxCommentMarginModel,
  cssWidth: number,
  cssHeight: number,
  activeId: string | null,
  onSetActive: (id: string, active: boolean) => void,
  zoom: number,
  logicalMarginWidth: number,
  showMarkers: boolean,
  includeResolved = false,
  onGeometryChange?: () => void,
  onScrollGeometryChange?: () => void,
): ReadOnlyCommentMarginGeometry {
  ensureReadOnlyCommentStyles(tintLayer.ownerDocument);
  if (margin) margin.dataset.ooxmlCommentZoom = String(zoom);
  tintLayer.innerHTML = '';
  const threads = commentThreads(model.comments, includeResolved);
  const accentById = new Map(threads.map((thread) => [
    thread.root.id,
    readOnlyCommentAuthorAccent(thread.root.author),
  ]));
  const visibleThreadIds = new Set(threads.map((thread) => thread.root.id));
  const firstAnchor = new Map<string, CommentAnchorRange>();
  for (const anchor of model.anchors) {
    if (visibleThreadIds.has(anchor.commentId) && !firstAnchor.has(anchor.commentId)) {
      firstAnchor.set(anchor.commentId, anchor);
    }
  }
  const resolvedPageAnchors = resolvePageCommentAnchors(model.anchors, runs);
  const resolvedAnchorSet = new Set(resolvedPageAnchors.map(({ anchor }) => anchor));
  const anchorRects = new Map<string, ReadOnlyCommentRect[]>();
  const markerAnchorById = new Map<string, Readonly<{
    rect: ReadOnlyCommentRect;
    direction?: 'ltr' | 'rtl';
  }>>();
  const surface = tintLayer.parentElement;
  for (const { anchor, runs: anchorRuns } of resolvedPageAnchors) {
    if (!visibleThreadIds.has(anchor.commentId)) continue;
    const active = activeId === anchor.commentId;
    for (const run of mergeSameLineRuns(anchorRuns.map(wordHighlightRun))) {
      const tint = createTint(
        tintLayer,
        run,
        cssWidth,
        cssHeight,
        active,
        accentById.get(anchor.commentId) ?? readOnlyCommentAuthorAccent(undefined),
      );
      const rect = run.transform && surface
        ? relativeElementRect(tint, surface)
        : Object.freeze({ x: run.x, y: run.y, width: run.w, height: run.h });
      const list = anchorRects.get(anchor.commentId) ?? [];
      if (!anchorRects.has(anchor.commentId)) anchorRects.set(anchor.commentId, list);
      const resolvedRect = rect ?? Object.freeze({ x: run.x, y: run.y, width: run.w, height: run.h });
      list.push(resolvedRect);
      if (!markerAnchorById.has(anchor.commentId)) {
        markerAnchorById.set(anchor.commentId, Object.freeze({
          rect: resolvedRect,
          ...(run.direction ? { direction: run.direction } : {}),
        }));
      }
    }
  }
  const cardThreads = threads.flatMap((thread): ReadOnlyCommentThread[] => {
    const anchor = firstAnchor.get(thread.root.id);
    if (!anchor || !resolvedAnchorSet.has(anchor)) return [];
    return [{
      occurrenceKey: thread.root.id,
      root: toMessage(thread.root, thread.root.id, 0),
      replies: thread.replies.map((reply, index) => toMessage(reply, thread.root.id, index + 1)),
    }];
  });
  const cardHosts = margin
    ? buildReadOnlyCommentMargin(margin, cardThreads, {
        activeId,
        zoom,
        logicalWidth: logicalMarginWidth,
        onSetActive,
        onGeometryChange,
        onScrollGeometryChange,
        preferredTopById: new Map(cardThreads.map((thread) => {
          const first = anchorRects.get(thread.occurrenceKey)?.[0];
          return [thread.occurrenceKey, first?.y ?? 0] as const;
        })),
      })
    : new Map<string, HTMLButtonElement>();
  if (showMarkers) {
    for (const [visibleIndex, thread] of cardThreads.entries()) {
      const anchor = markerAnchorById.get(thread.occurrenceKey);
      if (!anchor) continue;
      const marker = createReadOnlyCommentMarker(tintLayer.ownerDocument, {
        occurrenceKey: thread.occurrenceKey,
        visibleIndex,
        author: thread.root.author,
        active: activeId === thread.occurrenceKey,
        zoom,
        onSetActive,
      });
      const half = READ_ONLY_COMMENT_MARKER_SIZE_PX * zoom / 2;
      const gap = 4 * zoom;
      const inheritedDirection = tintLayer.ownerDocument.defaultView?.getComputedStyle?.(
        tintLayer,
      ).direction;
      const rtl = (anchor.direction ?? inheritedDirection) === 'rtl';
      const preferredLeft = rtl
        ? anchor.rect.x - gap - half
        : anchor.rect.x + anchor.rect.width + gap + half;
      const left = Math.max(half, Math.min(preferredLeft, cssWidth - half));
      const top = Math.max(
        half,
        Math.min(anchor.rect.y + anchor.rect.height / 2, cssHeight - half),
      );
      marker.style.left = `${left / cssWidth * 100}%`;
      marker.style.top = `${top / cssHeight * 100}%`;
      const card = cardHosts.get(thread.occurrenceKey);
      if (card?.id) marker.setAttribute('aria-controls', card.id);
      tintLayer.appendChild(marker);
    }
  }
  if (!onGeometryChange && !onScrollGeometryChange) {
    return Object.freeze({
      threads: Object.freeze(cardThreads.map((thread): ReadOnlyCommentThreadGeometry => Object.freeze({
        occurrenceKey: thread.occurrenceKey,
        active: activeId === thread.occurrenceKey,
        anchorRects: Object.freeze(anchorRects.get(thread.occurrenceKey) ?? []),
      }))),
      scrollTop: margin?.scrollTop ?? 0,
    });
  }
  const marginRect = margin && surface ? relativeElementRect(margin, surface) : undefined;
  const geometry = Object.freeze(cardThreads.map((thread): ReadOnlyCommentThreadGeometry => {
    const cardHost = cardHosts.get(thread.occurrenceKey);
    const cardRect = cardHost && surface
      ? relativeElementRect(cardHost, surface)
      : undefined;
    return Object.freeze({
      occurrenceKey: thread.occurrenceKey,
      active: activeId === thread.occurrenceKey,
      anchorRects: Object.freeze(anchorRects.get(thread.occurrenceKey) ?? []),
      ...(cardRect ? { cardRect } : {}),
    });
  }));
  return Object.freeze({
    threads: geometry,
    ...(marginRect ? { cardClipBounds: marginRect } : {}),
    scrollTop: margin?.scrollTop ?? 0,
  });
}
