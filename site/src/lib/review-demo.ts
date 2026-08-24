import {
  DocxDocument,
  resolveCommentAnchorRuns,
  resolveRevisionAnchorRuns,
  type DocComment,
  type DocRevision,
  type DocxTextRunInfo,
} from '@silurus/ooxml-docx';

const SVG_NS = 'http://www.w3.org/2000/svg';
const PAGE_WIDTH = 760;
const MOBILE_MARGIN_QUERY = '(max-width: 520px)';

interface ReviewHighlightRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface ReviewHighlightBand extends ReviewHighlightRect {
  paragraphKey: string;
}

interface AnchoredReviewEntry {
  readonly id: string;
  readonly kind: 'comment' | 'revision';
  readonly bands: readonly ReviewHighlightRect[];
  readonly comment?: Readonly<DocComment>;
  readonly revision?: Readonly<DocRevision>;
}

function reviewParagraphKey(run: Readonly<DocxTextRunInfo>): string {
  if (run.source) {
    return `${run.source.story}:${run.source.storyInstance}:${run.source.path.join('.')}`;
  }
  return `paragraph:${run.paragraphId ?? ''}`;
}

/** Presentation-only coalescing: rendered run fragments stay exact in the
 * package API, while fragments from one paragraph and visual line become one
 * selection band. Whitespace between words is part of the selected range;
 * separate paragraphs and lines stay split. */
export function mergeReviewHighlightRuns(
  runs: readonly Readonly<DocxTextRunInfo>[],
): ReviewHighlightRect[] {
  const lineTolerancePx = 1;
  const sorted = [...runs]
    .filter((run) => run.w > 0 && run.h > 0)
    .sort((left, right) => left.y - right.y || left.x - right.x);
  const merged: ReviewHighlightBand[] = [];
  for (const run of sorted) {
    const previous = merged.at(-1);
    const paragraphKey = reviewParagraphKey(run);
    const sameLine = previous !== undefined
      && previous.paragraphKey === paragraphKey
      && Math.abs(previous.y - run.y) <= lineTolerancePx
      && Math.abs(previous.h - run.h) <= lineTolerancePx;
    if (previous && sameLine) {
      previous.w = Math.max(previous.x + previous.w, run.x + run.w) - previous.x;
      previous.h = Math.max(previous.h, run.h);
    } else {
      merged.push({ x: run.x, y: run.y, w: run.w, h: run.h, paragraphKey });
    }
  }
  return merged.map(({ paragraphKey: _, ...rect }) => rect);
}

function initials(author: string | undefined): string {
  return (author ?? '?')
    .split(/\s+/u)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? '')
    .join('');
}

function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function commentCard(
  comment: Readonly<DocComment>,
  replies: readonly Readonly<DocComment>[],
): HTMLElement {
  const card = element('article', 'review-comment-card');
  card.dataset.commentId = comment.id;
  const heading = element('div', 'review-comment-heading');
  const avatar = element('span', 'review-avatar', initials(comment.author));
  const identity = element('div', 'review-comment-identity');
  identity.append(
    element('strong', undefined, comment.author || 'Unknown reviewer'),
    element('span', undefined, comment.resolved ? 'Resolved thread' : 'Open thread'),
  );
  heading.append(avatar, identity);
  if (comment.resolved) heading.append(element('span', 'review-resolved', 'Resolved'));
  card.append(heading, element('p', 'review-comment-body', comment.text));

  for (const reply of replies) {
    const row = element('div', 'review-reply');
    row.append(
      element('span', 'review-avatar review-avatar-small', initials(reply.author)),
      element('p', undefined, reply.text),
    );
    card.append(row);
  }
  return card;
}

function revisionCard(revision: Readonly<DocRevision>): HTMLElement {
  const card = element('article', 'review-comment-card review-change-card');
  const heading = element('div', 'review-comment-heading');
  const avatar = element('span', 'review-avatar', initials(revision.author));
  const identity = element('div', 'review-comment-identity');
  const label = revision.kind === 'moveFrom' ? 'Moved from' : 'Deleted';
  identity.append(
    element('strong', undefined, revision.author || 'Unknown reviewer'),
    element('span', undefined, `${label} text`),
  );
  heading.append(avatar, identity);
  const body = element('p', 'review-comment-body review-change-body');
  body.append(element('span', 'review-change-label', label), element('del', undefined, revision.text));
  card.append(heading, body);
  return card;
}

export function mountReviewDemo(root: HTMLElement, url: string): void {
  if (root.dataset.mounted === '1') return;
  root.dataset.mounted = '1';
  const canvas = root.querySelector<HTMLCanvasElement>('[data-review-canvas]');
  const highlightSvg = root.querySelector<SVGSVGElement>('[data-review-highlights]');
  const connectorSvg = root.querySelector<SVGSVGElement>('[data-review-connectors]');
  const layoutRoot = root.querySelector<HTMLElement>('[data-review-layout]');
  const threadList = root.querySelector<HTMLElement>('[data-review-threads]');
  const status = root.querySelector<HTMLElement>('[data-review-status]');
  const count = root.querySelector<HTMLElement>('[data-review-count]');
  const changeCount = root.querySelector<HTMLElement>('[data-review-change-count]');
  const pageLabel = root.querySelector<HTMLElement>('[data-review-page-label]');
  if (!canvas || !highlightSvg || !connectorSvg || !layoutRoot || !threadList || !status) return;

  let loaded: DocxDocument | undefined;
  let resizeObserver: ResizeObserver | undefined;
  const marginQuery = matchMedia(MOBILE_MARGIN_QUERY);
  let onMarginQueryChange: (() => void) | undefined;
  const cleanup = () => {
    resizeObserver?.disconnect();
    if (onMarginQueryChange) marginQuery.removeEventListener('change', onMarginQueryChange);
    loaded?.destroy();
    loaded = undefined;
  };
  window.addEventListener('pagehide', cleanup, { once: true });

  void DocxDocument.load(url, { useGoogleFonts: true }).then(async (documentModel) => {
    loaded = documentModel;
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const comments = documentModel.comments;
    const roots = comments.filter((comment) => !comment.parentId);
    const anchors = documentModel.commentAnchorRanges();
    const revisionAnchors = documentModel.revisionAnchorRanges();
    const marginRevision = (revision: Readonly<DocRevision> | undefined) =>
      revision?.kind === 'deletion' || revision?.kind === 'moveFrom';
    let pageIndex = 0;
    let runsByComment = new Map<string, readonly Readonly<DocxTextRunInfo>[]>();
    let runsByRevision = new Map<number, readonly Readonly<DocxTextRunInfo>[]>();

    for (let candidate = 0; candidate < documentModel.pageCount; candidate += 1) {
      const candidateRuns = await documentModel.collectPageRuns(candidate, { width: PAGE_WIDTH });
      const candidateMap = new Map<string, readonly Readonly<DocxTextRunInfo>[]>();
      const candidateRevisionMap = new Map<number, readonly Readonly<DocxTextRunInfo>[]>();
      for (const anchor of anchors) {
        const resolved = resolveCommentAnchorRuns(anchor, candidateRuns);
        if (resolved.length === 0) continue;
        const previous = candidateMap.get(anchor.commentId) ?? [];
        candidateMap.set(anchor.commentId, [...previous, ...resolved]);
      }
      for (const anchor of revisionAnchors) {
        if (!marginRevision(documentModel.revisions[anchor.revisionIndex])) continue;
        const resolved = resolveRevisionAnchorRuns(anchor, candidateRuns);
        if (resolved.length === 0) continue;
        const previous = candidateRevisionMap.get(anchor.revisionIndex) ?? [];
        candidateRevisionMap.set(anchor.revisionIndex, [...previous, ...resolved]);
      }
      if (roots.some((comment) => (candidateMap.get(comment.id)?.length ?? 0) > 0)
        || candidateRevisionMap.size > 0) {
        pageIndex = candidate;
        runsByComment = candidateMap;
        runsByRevision = candidateRevisionMap;
        break;
      }
    }

    await documentModel.renderPage(canvas, pageIndex, { width: PAGE_WIDTH, dpr });
    // renderPage sets a fixed CSS width for standalone use. This responsive
    // composition scales the canvas, highlights, and review margin together.
    canvas.style.width = '100%';
    canvas.style.height = 'auto';
    const size = documentModel.pageSize(pageIndex);
    const pageHeight = PAGE_WIDTH * size.heightPt / size.widthPt;
    highlightSvg.setAttribute('viewBox', `0 0 ${PAGE_WIDTH} ${pageHeight}`);
    if (pageLabel) pageLabel.textContent = `Page ${pageIndex + 1} · Accepted view`;

    const repliesFor = (id: string) => comments.filter((comment) => comment.parentId === id);
    const commentEntries: AnchoredReviewEntry[] = roots.flatMap((comment) => {
      const geometry = runsByComment.get(comment.id) ?? [];
      const bands = mergeReviewHighlightRuns(geometry);
      return bands.length > 0 ? [{
        id: `comment:${comment.id}`,
        kind: 'comment' as const,
        comment,
        bands,
      }] : [];
    });
    const revisionEntries: AnchoredReviewEntry[] = [...runsByRevision].flatMap(
      ([revisionIndex, geometry]) => {
        const revision = documentModel.revisions[revisionIndex];
        const bands = mergeReviewHighlightRuns(geometry);
        return revision && marginRevision(revision) && bands.length > 0 ? [{
          id: `revision:${revisionIndex}`,
          kind: 'revision' as const,
          revision,
          bands,
        }] : [];
      },
    );
    const anchored = [...commentEntries, ...revisionEntries].sort((left, right) =>
      left.bands[0]!.y - right.bands[0]!.y
      || left.bands[0]!.x - right.bands[0]!.x
      || left.id.localeCompare(right.id));

    if (count) count.textContent = `${roots.length} threads`;
    if (changeCount) changeCount.textContent = `${documentModel.revisions.length} changes`;

    for (const entry of anchored) {
      const { bands } = entry;
      for (const band of bands) {
        const rect = document.createElementNS(SVG_NS, 'rect');
        rect.setAttribute('x', String(band.x));
        rect.setAttribute('y', String(band.y));
        rect.setAttribute('width', String(Math.max(band.w, 2)));
        rect.setAttribute('height', String(band.h));
        rect.setAttribute('rx', '3');
        rect.dataset.reviewId = entry.id;
        if (entry.kind === 'revision') rect.classList.add('is-revision');
        highlightSvg.append(rect);
      }

      const card = entry.kind === 'comment'
        ? commentCard(entry.comment!, repliesFor(entry.comment!.id))
        : revisionCard(entry.revision!);
      card.dataset.reviewId = entry.id;
      card.tabIndex = 0;
      card.setAttribute('role', 'button');
      card.setAttribute('aria-pressed', 'false');
      threadList.append(card);
    }

    let pinnedReviewId: string | undefined;
    const select = (id: string) => {
      root.dataset.selectedReview = id;
      for (const node of root.querySelectorAll<HTMLElement>('[data-review-id]')) {
        const selected = node.dataset.reviewId === id;
        node.classList.toggle('is-selected', selected);
        if (node.classList.contains('review-comment-card')) {
          node.setAttribute('aria-pressed', String(selected));
        }
      }
    };
    const clearSelection = () => {
      delete root.dataset.selectedReview;
      for (const node of root.querySelectorAll<HTMLElement>('[data-review-id]')) {
        node.classList.remove('is-selected');
        if (node.classList.contains('review-comment-card')) {
          node.setAttribute('aria-pressed', 'false');
        }
      }
    };
    const restorePinnedSelection = () => {
      if (pinnedReviewId) select(pinnedReviewId);
      else clearSelection();
    };

    for (const card of threadList.querySelectorAll<HTMLElement>('.review-comment-card')) {
      const id = card.dataset.reviewId;
      if (!id) continue;
      card.addEventListener('mouseenter', () => select(id));
      card.addEventListener('mouseleave', restorePinnedSelection);
      card.addEventListener('focus', () => select(id));
      card.addEventListener('blur', restorePinnedSelection);
      const togglePinned = () => {
        if (pinnedReviewId === id) {
          pinnedReviewId = undefined;
          clearSelection();
        } else {
          pinnedReviewId = id;
          select(id);
        }
      };
      card.addEventListener('click', (event) => {
        event.stopPropagation();
        togglePinned();
      });
      card.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter' && event.key !== ' ') return;
        event.preventDefault();
        togglePinned();
      });
    }
    root.addEventListener('click', () => {
      pinnedReviewId = undefined;
      clearSelection();
    });

    const layoutMargin = () => {
      connectorSvg.replaceChildren();
      const cards = [...threadList.querySelectorAll<HTMLElement>('.review-comment-card')];
      if (marginQuery.matches) {
        for (const card of cards) card.style.removeProperty('top');
        return;
      }

      const layoutRect = layoutRoot.getBoundingClientRect();
      const pageRect = canvas.getBoundingClientRect();
      const threadRect = threadList.getBoundingClientRect();
      if (layoutRect.width <= 0 || pageRect.width <= 0 || threadRect.width <= 0) return;
      connectorSvg.setAttribute('viewBox', `0 0 ${layoutRect.width} ${layoutRect.height}`);

      let nextTop = 6;
      const gap = 12;
      for (const [index, entry] of anchored.entries()) {
        const card = cards[index];
        const band = entry.bands[0];
        if (!card || !band) continue;
        const anchorY = pageRect.top - layoutRect.top
          + (band.y + band.h / 2) / pageHeight * pageRect.height;
        const desiredTop = anchorY - (threadRect.top - layoutRect.top) - 18;
        const top = Math.max(nextTop, desiredTop);
        card.style.top = `${top}px`;
        nextTop = top + card.offsetHeight + gap;

        const startX = pageRect.left - layoutRect.left
          + (band.x + band.w) / PAGE_WIDTH * pageRect.width;
        const startY = anchorY;
        const endX = threadRect.left - layoutRect.left + 1;
        const endY = threadRect.top - layoutRect.top + top + Math.min(25, card.offsetHeight / 2);
        const horizontal = Math.max(14, (endX - startX) * .5);
        const path = document.createElementNS(SVG_NS, 'path');
        path.setAttribute(
          'd',
          `M ${startX} ${startY} C ${startX + horizontal} ${startY}, ${endX - horizontal} ${endY}, ${endX} ${endY}`,
        );
        path.dataset.reviewId = entry.id;
        if (root.dataset.selectedReview === entry.id) path.classList.add('is-selected');
        connectorSvg.append(path);
      }
    };

    resizeObserver = new ResizeObserver(layoutMargin);
    resizeObserver.observe(layoutRoot);
    resizeObserver.observe(canvas);
    onMarginQueryChange = layoutMargin;
    marginQuery.addEventListener('change', onMarginQueryChange);
    clearSelection();
    status.hidden = true;
    requestAnimationFrame(() => requestAnimationFrame(layoutMargin));
    void document.fonts?.ready.then(layoutMargin);
  }).catch((error: unknown) => {
    status.textContent = error instanceof Error ? error.message : 'Unable to load the review sample.';
    status.dataset.error = '1';
    cleanup();
  });
}
