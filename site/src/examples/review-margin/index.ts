import {
  DocxDocument,
  resolveCommentAnchorRuns,
  resolveRevisionAnchorRuns,
  type DocComment,
  type DocRevision,
  type DocxTextRunInfo,
} from '@silurus/ooxml/docx';

const PAGE_WIDTH = 760;
const SVG_NS = 'http://www.w3.org/2000/svg';

interface ReviewItem {
  id: string;
  label: string;
  text: string;
  meta: string;
  runs: readonly Readonly<DocxTextRunInfo>[];
  marker: 'range' | 'point' | 'fallback';
  deleted: boolean;
  replies?: readonly DocComment[];
}

interface ReviewController {
  updatePage(pageIndex: number): Promise<void>;
  destroy(): void;
}

function sameSource(a: DocxTextRunInfo['source'], b: DocxTextRunInfo['source']): boolean {
  return a !== undefined && b !== undefined && a.story === b.story
    && a.storyInstance === b.storyInstance
    && a.path.length === b.path.length
    && a.path.every((part, index) => part === b.path[index]);
}

function lineBands(runs: readonly Readonly<DocxTextRunInfo>[]) {
  const result: Array<Readonly<DocxTextRunInfo> & { sourceKey: string }> = [];
  for (const run of [...runs].sort((a, b) => a.y - b.y || a.x - b.x)) {
    if (run.w <= 0 || run.h <= 0) continue;
    const sourceKey = run.source
      ? `${run.source.story}:${run.source.storyInstance}:${run.source.path.join('.')}`
      : `paragraph:${run.paragraphId ?? ''}`;
    const previous = result.at(-1);
    const gap = previous ? run.x - (previous.x + previous.w) : Number.POSITIVE_INFINITY;
    if (previous && previous.sourceKey === sourceKey && previous.transform === run.transform
      && Math.abs(previous.y - run.y) <= 1 && Math.abs(previous.h - run.h) <= 1
      && gap >= -1 && gap <= Math.max(2, run.h * .4)) {
      result[result.length - 1] = { ...previous, w: run.x + run.w - previous.x };
    } else {
      result.push({ ...run, sourceKey });
    }
  }
  return result;
}

function commentThreads(comments: readonly DocComment[]) {
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const rootOf = (comment: DocComment): DocComment => {
    const seen = new Set<string>();
    let current = comment;
    while (current.parentId && byId.has(current.parentId) && !seen.has(current.id)) {
      seen.add(current.id);
      current = byId.get(current.parentId) as DocComment;
    }
    return current;
  };
  const replies = new Map<string, DocComment[]>();
  for (const comment of comments) {
    const root = rootOf(comment);
    if (root === comment) continue;
    const group = replies.get(root.id) ?? [];
    group.push(comment);
    replies.set(root.id, group);
  }
  return { byId, replies, rootOf };
}

function revisionLabel(revision: DocRevision): string {
  const labels: Record<string, string> = {
    insertion: 'Inserted', deletion: 'Deleted', moveFrom: 'Moved from', moveTo: 'Moved to',
  };
  const action = labels[revision.kind] ?? `Changed (${revision.kind})`;
  return `${action} by ${revision.author || 'Unknown reviewer'}`;
}

function reviewItems(doc: DocxDocument, runs: readonly Readonly<DocxTextRunInfo>[]): ReviewItem[] {
  const threads = commentThreads(doc.comments);
  const comments = new Map<string, ReviewItem>();
  for (const anchor of doc.commentAnchorRanges()) {
    const comment = threads.byId.get(anchor.commentId);
    if (!comment) continue;
    const resolved = resolveCommentAnchorRuns(anchor, runs);
    if (resolved.length === 0) continue;
    const root = threads.rootOf(comment);
    const direct = resolved.some((run) => run.sourceRunIndex !== undefined
      && sameSource(run.source, anchor.source)
      && run.sourceRunIndex >= anchor.startRunIndex && run.sourceRunIndex < anchor.endRunIndex);
    const point = !direct && anchor.startRunIndex === anchor.endRunIndex
      && resolved.some((run) => run.sourceRunIndex !== undefined && sameSource(run.source, anchor.source));
    const reference = !direct && !point
      && resolved.some((run) => run.sourceRunIndex !== undefined
        && sameSource(run.source, anchor.reference.source));
    const marker = direct ? 'range' : point || reference ? 'point' : 'fallback';
    const existing = comments.get(root.id);
    const orphan = root.parentId !== undefined && !threads.byId.has(root.parentId);
    comments.set(root.id, {
      id: `comment-${root.id}`,
      label: root.author || 'Unknown reviewer',
      text: root.text,
      meta: `${root.resolved ? 'Resolved thread' : 'Open thread'}${orphan ? ' · orphaned reply' : ''}${marker === 'point' ? ' · authored boundary' : marker === 'fallback' ? ' · approximate final-state position' : ''}`,
      runs: [...(existing?.runs ?? []), ...resolved],
      marker: existing?.marker === 'range' || marker === 'range' ? 'range'
        : existing?.marker === 'point' || marker === 'point' ? 'point' : 'fallback',
      deleted: false,
      replies: threads.replies.get(root.id) ?? [],
    });
  }

  const revisions = doc.revisionAnchorRanges().flatMap((anchor): ReviewItem[] => {
    const revision = doc.revisions[anchor.revisionIndex];
    if (!revision) return [];
    const resolved = resolveRevisionAnchorRuns(anchor, runs);
    if (resolved.length === 0) return [];
    const direct = resolved.some((run) => run.sourceRunIndex !== undefined
      && sameSource(run.source, anchor.source)
      && run.sourceRunIndex >= anchor.startRunIndex && run.sourceRunIndex < anchor.endRunIndex);
    return [{
      id: `revision-${anchor.revisionIndex}`,
      label: revisionLabel(revision),
      text: revision.text,
      meta: direct ? revision.kind : `${revision.kind} · approximate final-state position`,
      runs: resolved,
      marker: direct ? 'range' : 'fallback',
      deleted: revision.kind === 'deletion' || revision.kind === 'moveFrom',
    }];
  });
  return [...comments.values(), ...revisions];
}

export async function mountReviewExample(root: HTMLElement, signal?: AbortSignal): Promise<ReviewController> {
  const canvas = root.querySelector('canvas') as HTMLCanvasElement;
  const highlights = root.querySelector('[data-review-highlights]') as SVGSVGElement;
  const list = root.querySelector('[data-review-items]') as HTMLOListElement;
  const status = root.querySelector('[data-review-status]') as HTMLElement;
  const content = root.querySelector('[data-review-content]') as HTMLElement;
  const empty = root.querySelector('[data-review-empty]') as HTMLElement;
  const pageStatus = root.querySelector('[data-review-page-status]') as HTMLElement;
  const previous = root.querySelector('[data-review-previous]') as HTMLButtonElement;
  const next = root.querySelector('[data-review-next]') as HTMLButtonElement;
  const events = new AbortController();
  let doc: DocxDocument | undefined;
  let generation = 0;
  let destroyed = false;
  let selectedId: string | undefined;
  let previewId: string | undefined;

  const paintSelection = (id?: string): void => {
    root.querySelectorAll<HTMLElement>('[data-review-id]').forEach((element) => {
      const selected = id !== undefined && element.dataset.reviewId === id;
      if (element.classList.contains('review-example__item')) {
        element.setAttribute('aria-pressed', String(selected));
      } else {
        element.dataset.selected = String(selected);
      }
    });
  };
  root.addEventListener('click', (event) => {
    const target = (event.target as Element).closest<HTMLElement>('[data-review-id]');
    if (target) {
      selectedId = selectedId === target.dataset.reviewId ? undefined : target.dataset.reviewId;
      paintSelection(selectedId);
    }
  }, { signal: events.signal });
  root.addEventListener('mouseover', (event) => {
    const target = (event.target as Element).closest<HTMLElement>('[data-review-id]');
    if (target) { previewId = target.dataset.reviewId; paintSelection(previewId); }
  }, { signal: events.signal });
  root.addEventListener('mouseout', (event) => {
    const target = (event.target as Element).closest<HTMLElement>('[data-review-id]');
    if (target) { previewId = undefined; paintSelection(selectedId); }
  }, { signal: events.signal });
  root.addEventListener('focusin', (event) => {
    const target = (event.target as Element).closest<HTMLElement>('[data-review-id]');
    if (target) { previewId = target.dataset.reviewId; paintSelection(previewId); }
  }, { signal: events.signal });
  root.addEventListener('focusout', () => {
    previewId = undefined;
    paintSelection(selectedId);
  }, { signal: events.signal });

  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    generation += 1;
    events.abort();
    doc?.destroy();
    doc = undefined;
  };
  signal?.addEventListener('abort', destroy, { once: true });
  if (signal?.aborted) destroy();

  const updatePage = async (requestedPage: number): Promise<void> => {
    if (!doc || destroyed) return;
    const pageIndex = Math.max(0, Math.min(requestedPage, doc.pageCount - 1));
    const request = ++generation;
    const stage = document.createElement('canvas');
    const runs: DocxTextRunInfo[] = [];
    await doc.renderPage(stage, pageIndex, { width: PAGE_WIDTH, onTextRun: (run) => runs.push(run) });
    if (destroyed || request !== generation) return;

    canvas.width = stage.width;
    canvas.height = stage.height;
    canvas.getContext('2d')?.drawImage(stage, 0, 0);
    canvas.style.width = '100%';
    canvas.style.height = 'auto';
    const page = doc.pageSize(pageIndex);
    const pageHeight = PAGE_WIDTH * page.heightPt / page.widthPt;
    highlights.setAttribute('viewBox', `0 0 ${PAGE_WIDTH} ${pageHeight}`);
    highlights.replaceChildren();
    list.replaceChildren();

    const items = reviewItems(doc, runs);
    for (const item of items) {
      const row = document.createElement('li');
      row.id = `review-row-${item.id}`;
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'review-example__item';
      button.dataset.reviewId = item.id;
      button.setAttribute('aria-pressed', 'false');
      const quote = item.runs.map((run) => run.text).join('').trim();
      button.setAttribute('aria-label', `${item.label}. Page ${pageIndex + 1}. ${item.text}. Anchored to ${quote || 'a nearby final-state position'}.`);
      const label = document.createElement('strong');
      const body = item.deleted ? document.createElement('del') : document.createElement('span');
      const meta = document.createElement('span');
      label.textContent = item.label;
      body.textContent = item.text;
      meta.className = 'review-example__meta';
      meta.textContent = `Page ${pageIndex + 1} · ${item.meta}`;
      button.append(label, body, meta);
      if (item.replies?.length) {
      const replies = document.createElement('span');
        replies.className = 'review-example__replies';
        for (const reply of item.replies) {
          const paragraph = document.createElement('span');
          paragraph.textContent = `${reply.author || 'Unknown reviewer'}: ${reply.text}`;
          replies.append(paragraph);
        }
        button.append(replies);
      }
      row.append(button);
      list.append(row);

      const bands = item.marker === 'range' ? lineBands(item.runs) : lineBands(item.runs).slice(0, 1);
      const anchorIds: string[] = [];
      for (const [index, band] of bands.entries()) {
        const rect = document.createElementNS(SVG_NS, 'rect');
        rect.id = `review-anchor-${item.id}-${index}`;
        anchorIds.push(rect.id);
        rect.classList.add('review-example__anchor');
        if (item.marker !== 'range') rect.classList.add('review-example__anchor--point');
        if (item.marker === 'fallback') rect.classList.add('review-example__anchor--fallback');
        rect.dataset.reviewId = item.id;
        rect.setAttribute('aria-hidden', 'true');
        rect.setAttribute('focusable', 'false');
        rect.setAttribute('x', String(band.x));
        rect.setAttribute('y', String(band.y));
        rect.setAttribute('width', String(item.marker === 'range' ? band.w : 3));
        rect.setAttribute('height', String(band.h));
        if (band.transform) {
          rect.style.transform = band.transform;
          rect.style.transformOrigin = `${band.x}px ${band.y}px`;
        }
        highlights.append(rect);
      }
      if (anchorIds.length > 0) button.setAttribute('aria-controls', anchorIds.join(' '));
    }

    empty.hidden = items.length > 0;
    previous.disabled = pageIndex === 0;
    next.disabled = pageIndex === doc.pageCount - 1;
    previous.dataset.page = String(pageIndex - 1);
    next.dataset.page = String(pageIndex + 1);
    pageStatus.textContent = `Page ${pageIndex + 1} of ${doc.pageCount}`;
    status.textContent = `Page ${pageIndex + 1} loaded with ${items.length} review item${items.length === 1 ? '' : 's'}.`;
    status.classList.add('review-example__status--ready');
    content.hidden = false;
    selectedId = undefined;
    previewId = undefined;
    paintSelection(undefined);
  };

  const navigate = (event: Event): void => {
    const page = Number((event.currentTarget as HTMLButtonElement).dataset.page);
    if (Number.isInteger(page)) {
      void updatePage(page).catch((error: unknown) => {
        status.classList.remove('review-example__status--ready');
        status.textContent = error instanceof Error ? error.message : 'Unable to render this page.';
      });
    }
  };
  previous.addEventListener('click', navigate, { signal: events.signal });
  next.addEventListener('click', navigate, { signal: events.signal });

  try {
    const source = root.dataset.src;
    if (!source) throw new Error('Set data-src to a same-origin or CORS-enabled DOCX URL.');
    doc = await DocxDocument.load(source);
    if (destroyed) { doc.destroy(); return { updatePage, destroy }; }
    if (doc.pageCount === 0) throw new Error('The DOCX has no renderable pages.');
    await updatePage(Number(root.dataset.page ?? 0));
  } catch (error) {
    doc?.destroy();
    doc = undefined;
    status.textContent = error instanceof Error ? error.message : 'Unable to load review data.';
    throw error;
  }

  return {
    updatePage,
    destroy,
  };
}

const lifetime = new AbortController();
const mounted = await Promise.all(
  [...document.querySelectorAll<HTMLElement>('[data-review-example]')].map(async (root) => {
    try { return await mountReviewExample(root, lifetime.signal); }
    catch { return undefined; } // mountReviewExample has already announced the error.
  }),
);
const controllers = mounted.filter((item): item is ReviewController => item !== undefined);
// This site uses Astro navigation. In a framework component, call controller.destroy()
// from that component's unmount/effect cleanup instead.
document.addEventListener('astro:before-swap', () => {
  lifetime.abort();
  controllers.forEach((controller) => controller.destroy());
}, { once: true });
