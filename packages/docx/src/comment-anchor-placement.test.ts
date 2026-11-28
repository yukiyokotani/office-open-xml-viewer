import { describe, expect, it, vi } from 'vitest';
import {
  adoptCommentAnchorPlacements,
  collectLayoutSourceCommentRangesIfPresent,
  commentAnchorPlacement,
  resolveCommentAnchorRuns,
  resolveDocxCommentThreads,
  resolveDocxCommentThreadsOnPage,
  type CommentAnchorRange,
} from './comments.js';
import { buildDocxCommentMargin } from './comment-margin.js';
import { DocxDocument } from './document.js';
import { DocxFindController } from './find.js';
import { layoutDocumentInput } from './layout/document.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import type { LayoutOptions } from './layout/options.js';
import { normalizeLayoutOptions } from './layout/options.js';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { reviewProjectionIndexForDocument } from './layout/text-index.js';
import type { DocumentLayout } from './layout/types.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import {
  projectRenderWorkerLayoutMeta, renderWorkerCommentAnchors, renderWorkerLayoutMeta,
} from './render-worker-metadata.js';
import { paginateRenderWorkerDocumentProgressively } from './render-worker-progressive.js';
import { buildDocxTextLayer } from './text-layer.js';
import { DocxScrollCommentNavigation } from './scroll-comment-navigation.js';
import { installDom, makeEl } from './scroll-viewer-test-dom.js';
import {
  fakeElement, inlinePictureRun, oracleMeasureContext, paintedDocument, PICTURE_PART, withFakeDocument,
  type FakeElement,
} from './test-support/painted-glyphs.test-support.js';
import { resolveRevisionAnchorRuns } from './revisions.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxClipRegion, DocxDocumentModel, DocxTextRunInfo } from './types.js';

const COMMENT_PART = {
  parts: {
    'word/comments.xml': '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:comment w:id="1" w:author="A"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>',
  },
  relationships: '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    + 'Target="comments.xml"/>',
  contentTypes: '<Override PartName="/word/comments.xml" '
    + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
} as const;
const sized = (value: string) => `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const EXACT_12PT = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>';
const REFERENCE = '<w:r><w:commentReference w:id="1"/></w:r>';

/** The anchors of a laid-out document as the document and the render worker
 * collect them. */
function placedAnchors(layout: DocumentLayout, model: DocxDocumentModel): CommentAnchorRange[] {
  const index = reviewProjectionIndexForDocument(layout);
  return collectLayoutSourceCommentRangesIfPresent(model.comments, layoutSourceStore(model), index.renderedRunIndex, {
    cardOwners: { sourcePages: index.sourcePages, runsForPage: (page) => textRunsForPage(layout, page, { scale: 1 }) },
    splitRuns: index.splitRuns,
    textRunIndex: index.textRunIndex,
  });
}

function cardsAndTints(layout: DocumentLayout, model: DocxDocumentModel, anchors: readonly CommentAnchorRange[]) {
  installDom();
  return layout.pages.map((_, pageIndex) => {
    const tint = makeEl('div');
    const margin = makeEl('div');
    margin.clientHeight = 800;
    buildDocxCommentMargin(tint as never, margin as never, textRunsForPage(layout, pageIndex, { scale: 1 }),
      { comments: model.comments ?? [], anchors, pageIndex }, 612, 792, null, () => undefined, 1, 260, false);
    return [margin.children.length, tint.children.length > 0];
  });
}

describe('a reference-only comment beside a run that crosses a page break', () => {
  const wrapped = sized(`${'word '.repeat(910)}end`);

  it('after the run: stands with the run\'s last fragment, on the second page', async () => {
    const { layout, model } = await paintedDocument(`<w:p>${EXACT_12PT}${wrapped}${REFERENCE}</w:p>`, COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    const anchors = placedAnchors(layout, model);
    expect(anchors).toHaveLength(1);
    // Nothing of the placement is a field of the public range.
    expect(Object.keys(anchors[0]!).sort()).toEqual(
      ['cardOwner', 'commentId', 'endRunIndex', 'reference', 'source', 'startRunIndex']);
    expect(commentAnchorPlacement(anchors[0]!)?.cardPageIndex).toBe(1);
    const threads = (page: number) => resolveDocxCommentThreadsOnPage(
      model.comments ?? [], anchors, textRunsForPage(layout, page, { scale: 1 }), page);
    expect(threads(0)).toEqual([]);
    expect(threads(1)).toHaveLength(1);
    expect(cardsAndTints(layout, model, anchors)).toEqual([[0, false], [1, true]]);
  });

  it('before the run: stands with the run\'s first fragment, on the first page', async () => {
    const { layout, model } = await paintedDocument(`<w:p>${EXACT_12PT}${REFERENCE}${wrapped}</w:p>`, COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    const anchors = placedAnchors(layout, model);
    expect(commentAnchorPlacement(anchors[0]!)?.cardPageIndex).toBe(0);
    expect(cardsAndTints(layout, model, anchors)).toEqual([[1, true], [0, false]]);
  });

  it('keeps its page when the anchors reach the host from the render worker', async () => {
    const { layout, model } = await paintedDocument(`<w:p>${EXACT_12PT}${wrapped}${REFERENCE}</w:p>`, COMMENT_PART);
    const meta = structuredClone(projectRenderWorkerLayoutMeta(layout, layoutSourceStore(model), {
      comments: model.comments ?? [], revisions: [],
    }));
    const anchors = meta.commentAnchorRanges!;
    adoptCommentAnchorPlacements(anchors, meta.commentAnchorPlacements);
    expect(commentAnchorPlacement(anchors[0]!)?.cardPageIndex).toBe(1);
    expect(cardsAndTints(layout, model, anchors)).toEqual([[0, false], [1, true]]);
  });

  it('is placed by the page its runs were collected for', async () => {
    const { layout, model } = await paintedDocument(`<w:p>${EXACT_12PT}${wrapped}${REFERENCE}</w:p>`, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    const anchors = doc.commentAnchorRanges();
    const threads = async (page: number) =>
      resolveDocxCommentThreads(doc.comments, anchors, await doc.collectPageRuns(page));
    expect(await threads(0)).toEqual([]);
    expect(await threads(1)).toHaveLength(1);
    expect(await doc.getCommentThreads(0)).toEqual([]);
    expect(await doc.getCommentThreads(1)).toHaveLength(1);
  });
});

describe('a comment range that starts in a paragraph of deleted text only', () => {
  it('marks its laid-out part alone, which carries the card', async () => {
    // The range starts before a tracked deletion, the only run of its
    // paragraph, and ends after "Covered" in the next; "Outside" follows.
    const { layout, model, runs } = await paintedDocument(
      '<w:p><w:commentRangeStart w:id="1"/><w:del w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r>'
      + '<w:rPr><w:sz w:val="24"/></w:rPr><w:delText>Deleted</w:delText></w:r></w:del></w:p>'
      + `<w:p>${sized('Covered')}<w:commentRangeEnd w:id="1"/>${REFERENCE}${sized(' Outside')}</w:p>`, COMMENT_PART);
    const covered = runs.find((run) => run.text === 'Covered')!;
    const anchors = placedAnchors(layout, model);
    expect(anchors.map((anchor) => [anchor.source.path, anchor.cardOwner])).toEqual([[covered.source!.path, true]]);
    const [thread] = resolveDocxCommentThreadsOnPage(model.comments ?? [], anchors, runs, 0);
    expect(thread!.anchors).toHaveLength(1);
    expect(thread!.anchors[0]!.rects).toHaveLength(1);
    expect(thread!.anchors[0]!.rects[0]!.x).toBeCloseTo(covered.x, 6);
    expect(thread!.anchors[0]!.rects[0]!.width).toBeCloseTo(covered.w, 6);
  });
});

const DELETION = '<w:del w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
  + '<w:delText>Deleted</w:delText></w:r></w:del>';

describe('a comment range none of whose text is laid out', () => {
  it('is shown once, where one run stands in for the whole of it', async () => {
    // The range covers deleted text in two paragraphs; its reference stands
    // before "After".
    const { layout, model, runs } = await paintedDocument(
      `<w:p><w:commentRangeStart w:id="1"/>${DELETION}</w:p>`
      + `<w:p>${DELETION.replace('w:id="9"', 'w:id="10"')}<w:commentRangeEnd w:id="1"/>${REFERENCE}${sized('After')}</w:p>`,
      COMMENT_PART);
    const anchors = placedAnchors(layout, model);
    expect(anchors).toHaveLength(2);
    const rects = resolveDocxCommentThreadsOnPage(model.comments ?? [], anchors, runs, 0)
      .flatMap((thread) => thread.anchors.flatMap((anchor) => anchor.rects));
    expect(rects).toHaveLength(1);
    expect(rects[0]!.x).toBeCloseTo(runs.find((run) => run.text === 'After')!.x, 6);
    installDom();
    const tint = makeEl('div');
    const margin = makeEl('div');
    margin.clientHeight = 800;
    buildDocxCommentMargin(tint as never, margin as never, runs, { comments: model.comments ?? [], anchors, pageIndex: 0 },
      612, 792, null, () => undefined, 1, 260, false);
    expect(tint.children).toHaveLength(1);
    expect(margin.children).toHaveLength(1);
  });
});

describe('text that stands in for content that is not laid out', () => {
  const body = `<w:p><w:commentRangeStart w:id="1"/>${DELETION}<w:commentRangeEnd w:id="1"/>${REFERENCE}</w:p>`
    + `<w:p>${inlinePictureRun(40, 30)}</w:p><w:p>${sized('After')}</w:p>`;
  const parts = {
    parts: { ...COMMENT_PART.parts, ...PICTURE_PART.parts },
    relationships: COMMENT_PART.relationships + PICTURE_PART.relationships,
    contentTypes: COMMENT_PART.contentTypes + PICTURE_PART.contentTypes,
  };

  it('is text for a deletion: an inline picture between is passed over', async () => {
    const { layout, model, runs } = await paintedDocument(body, parts);
    const meta = projectRenderWorkerLayoutMeta(layout, layoutSourceStore(model), {
      comments: [], revisions: model.revisions ?? [],
    });
    expect(meta.revisionAnchorRanges).toHaveLength(1);
    expect(resolveRevisionAnchorRuns(meta.revisionAnchorRanges![0]!, runs).map((run) => run.text)).toEqual(['After']);
  });

  it('is text for a deletion in a main-mode document as well', async () => {
    const { layout, model } = await paintedDocument(body, parts);
    const { doc } = documentOf(model, layout);
    Object.assign(doc, { _revisionAnchorRanges: null });
    const ranges = doc.revisionAnchorRanges();
    expect(ranges).toHaveLength(1);
    expect(resolveRevisionAnchorRuns(ranges[0]!, await doc.collectPageRuns(0)).map((run) => run.text)).toEqual(['After']);
  });

  it('is text for a comment on that deletion too', async () => {
    const { layout, model, runs } = await paintedDocument(body, parts);
    const anchors = placedAnchors(layout, model);
    expect(anchors.flatMap((anchor) => resolveCommentAnchorRuns(anchor, runs)).map((run) => run.text)).toEqual(['After']);
  });
});

describe('a comment range an application makes itself', () => {
  it('is joined as given to the runs of a document that has no such comment', async () => {
    const { layout, model } = await paintedDocument(`<w:p>${sized('Word')}</w:p>`);
    const { doc } = documentOf(model, layout);
    const runs = await doc.collectPageRuns(0);
    const word = runs.find((run) => run.text === 'Word')!;
    const anchor: CommentAnchorRange = {
      commentId: 'app-1', source: word.source!, startRunIndex: word.sourceRunIndex!, endRunIndex: word.sourceRunIndex! + 1,
      reference: { source: word.source!, runIndex: word.sourceRunIndex! + 1, affinity: 'preceding' },
    };
    expect(resolveCommentAnchorRuns(anchor, runs)).toEqual([word]);
    expect(resolveDocxCommentThreads([{ id: 'app-1', text: 'mine' }], [anchor], runs)).toHaveLength(1);
  });
});

describe('a copy of a document\'s comment range', () => {
  it('is the application\'s own: it is joined as given, without the document\'s placement', async () => {
    // A reference-only comment after a run that crosses onto page two: the
    // document's range stands on page two; a copy of it carries no
    // placement and marks the fragment of the run on each page.
    const { layout, model } = await paintedDocument(
      `<w:p>${EXACT_12PT}${sized(`${'word '.repeat(910)}end`)}${REFERENCE}</w:p>`, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    const [anchor] = doc.commentAnchorRanges();
    const first = await doc.collectPageRuns(0);
    expect(resolveCommentAnchorRuns(anchor!, first)).toEqual([]);
    const copy = { ...anchor! };
    expect(resolveCommentAnchorRuns(copy, first).length).toBeGreaterThan(0);
    expect(resolveDocxCommentThreads(doc.comments, [copy], first)).toHaveLength(1);
  });
});

describe('a comment in a paragraph that every page repeats', () => {
  it('is shown on each of them: a repeated paragraph is not one crossing pages', async () => {
    const { layout, model } = await paintedDocument(
      `<w:p>${sized('One')}</w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr>${sized('Two')}</w:p>`, {
        parts: {
          ...COMMENT_PART.parts,
          'word/header1.xml': '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + `<w:p>${sized('Head')}${REFERENCE}</w:p></w:hdr>`,
        },
        relationships: `${COMMENT_PART.relationships}<Relationship Id="rIdH" `
          + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        contentTypes: `${COMMENT_PART.contentTypes}<Override PartName="/word/header1.xml" `
          + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
        section: '<w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          + 'w:type="default" r:id="rIdH"/>',
      });
    expect(layout.pages).toHaveLength(2);
    const index = reviewProjectionIndexForDocument(layout);
    expect(index.splitRuns.size).toBe(0);
    const anchors = placedAnchors(layout, model);
    expect(anchors).toHaveLength(1);
    // One card, on the first page; the anchor tinted on both.
    expect(cardsAndTints(layout, model, anchors)).toEqual([[1, true], [0, true]]);
  });
});

/** A main-mode document over `model`, its final view (`layout`) active. */
function documentOf(model: DocxDocumentModel, layout: DocumentLayout) {
  const source = layoutSourceStore(model);
  const services = createLayoutServices(source, { measureContext: oracleMeasureContext() });
  const builds: LayoutOptions[] = [];
  const variants = attachDocumentLayoutVariants({
    source,
    services,
    defaultCurrentDateMs: 0,
    buildLayout: (options) => {
      builds.push(options);
      return options.showTrackedChanges === true || options.currentDateMs !== 0
        ? layoutDocumentInput(source.bodyLayoutInput, services, options)
        : layout;
    },
  });
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, {
    _mode: 'main', _document: model, _source: source, _meta: null, _review: null,
    _commentAnchorRanges: null, _reviewProjectionIndex: null,
  });
  attachDocumentLayoutRuntime(doc, 0);
  documentLayoutRuntimeOf(doc).services = services;
  return { doc, services, variants, builds };
}

describe('comment threads asked for a view other than the active one', () => {
  const body = `<w:p>${sized('Before ')}${REFERENCE}`
    + '<w:del w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
    + `<w:delText xml:space="preserve">Deleted </w:delText></w:r></w:del>${sized('After')}</w:p>`;

  it('join that view\'s runs to that view\'s anchors', async () => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    const final = await doc.collectPageRuns(0);
    const markup = await doc.collectPageRuns(0, { showTrackedChanges: true });
    const deleted = markup.find((run) => run.text.startsWith('Deleted'))!;
    expect(deleted).toBeDefined();
    expect(final.some((run) => run.text.startsWith('Deleted'))).toBe(false);

    // The final view shows the point at the run after the deletion.
    const [shown] = await doc.getCommentThreads(0);
    expect(shown!.anchors[0]!.rects[0]!.x).toBeCloseTo(final.find((run) => run.text === 'After')!.x, 6);
    // The markup view lays the deleted run out: the point stands at it.
    const [marked] = await doc.getCommentThreads(0, { showTrackedChanges: true });
    expect(marked!.anchors[0]!.rects[0]!.x).toBeCloseTo(deleted.x, 6);
    // The public resolvers, given the markup view's runs and the document's
    // own anchors, place the point as the markup layout does.
    const anchors = doc.commentAnchorRanges();
    const [joined] = resolveDocxCommentThreads(doc.comments, anchors, markup);
    expect(joined!.anchors[0]!.anchor).toBe(anchors[0]);
    expect(joined!.anchors[0]!.rects[0]!.x).toBeCloseTo(deleted.x, 6);
    expect(resolveCommentAnchorRuns(anchors[0]!, markup).map((run) => run.text)).toEqual([deleted.text]);
    expect(resolveCommentAnchorRuns(anchors[0]!, final).map((run) => run.text)).toEqual(['After']);
    // The active view's anchors are untouched.
    const [again] = await doc.getCommentThreads(0);
    expect(again!.anchors[0]!.rects[0]!.x).toBeCloseTo(shown!.anchors[0]!.rects[0]!.x, 6);
  });

  it('a range segment that view leaves out is absent from its runs', async () => {
    // The range starts before a tracked deletion and ends after a line
    // break in the next paragraph. The final view lays out no covered run
    // and keeps both segments for the fallback; the markup view lays the
    // deletion out and leaves the break-only segment out.
    const { layout, model } = await paintedDocument(
      '<w:p><w:commentRangeStart w:id="1"/><w:del w:id="9" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r>'
      + '<w:rPr><w:sz w:val="24"/></w:rPr><w:delText>Deleted</w:delText></w:r></w:del></w:p>'
      + `<w:p><w:r><w:br/></w:r><w:commentRangeEnd w:id="1"/>${REFERENCE}${sized('Outside')}</w:p>`, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    const anchors = doc.commentAnchorRanges();
    expect(anchors).toHaveLength(2);
    const markup = await doc.collectPageRuns(0, { showTrackedChanges: true });
    const expected = (await doc.getCommentThreads(0, { showTrackedChanges: true }))
      .flatMap((thread) => thread.anchors.flatMap((anchor) => anchor.rects));
    expect(expected).toHaveLength(1);
    expect(expected[0]!.x).toBeCloseTo(markup.find((run) => run.text === 'Deleted')!.x, 6);
    expect(resolveDocxCommentThreads(doc.comments, anchors, markup)
      .flatMap((thread) => thread.anchors.flatMap((anchor) => anchor.rects))).toEqual(expected);
    expect(anchors.flatMap((anchor) => resolveCommentAnchorRuns(anchor, markup)).map((run) => run.text))
      .toEqual(['Deleted']);
  });

  it('runs keep the layout that produced them when the Date they were asked with changes', async () => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    const { doc, builds } = documentOf(model, layout);
    const date = new Date(86_400_000);
    const runs = await doc.collectPageRuns(0, { currentDate: date, showTrackedChanges: true });
    const anchors = doc.commentAnchorRanges();
    expect(resolveCommentAnchorRuns(anchors[0]!, runs)[0]!.text.startsWith('Deleted')).toBe(true);
    const built = builds.length;
    date.setTime(2 * 86_400_000);
    expect(resolveCommentAnchorRuns(anchors[0]!, runs)[0]!.text.startsWith('Deleted')).toBe(true);
    // No other layout is selected, or built, for the runs already collected.
    expect(builds).toHaveLength(built);
  });

  it('runs keep what their layout showed after the document is destroyed', async () => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    Object.assign(doc, { _bridge: { terminate: () => undefined }, _rawParts: new Map(), _embeddedFontFaces: [], _officeFontFaces: [] });
    const comments = doc.comments;
    const anchors = doc.commentAnchorRanges();
    const runs = await doc.collectPageRuns(0);
    const before = {
      runs: resolveCommentAnchorRuns(anchors[0]!, runs),
      threads: resolveDocxCommentThreads(comments, anchors, runs),
    };
    expect(before.runs.map((run) => run.text)).toEqual(['After']);
    expect(before.threads).toHaveLength(1);
    doc.destroy();
    expect(resolveCommentAnchorRuns(anchors[0]!, runs)).toEqual(before.runs);
    expect(resolveDocxCommentThreads(comments, anchors, runs)).toEqual(before.threads);
  });

  it('a worker document reads a Date when the runs are asked for, not when they arrive', async () => {
    const doc = Object.create(DocxDocument.prototype) as DocxDocument;
    const posted: { type: string; opts?: { currentDate?: unknown } }[] = [];
    Object.assign(doc, {
      _mode: 'worker', _review: null,
      _meta: { comments: [{ id: '1', text: 'c' }], revisions: [], commentAnchorRanges: [] },
      _bridge: {
        request: async (build: (id: number) => { type: string; opts?: { currentDate?: unknown } }) => {
          const request = build(1);
          posted.push(request);
          await Promise.resolve();
          return { type: 'runsCollected', id: 1, runs: [], commentAnchorRanges: [] };
        },
      },
    });
    attachDocumentLayoutRuntime(doc, 0);
    const date = new Date(86_400_000);
    const pending = doc.collectPageRuns(0, { currentDate: date });
    date.setTime(Number.NaN);
    await expect(pending).resolves.toEqual([]);
    expect(posted.map((request) => [request.type, request.opts?.currentDate])).toEqual([['collectRuns', 86_400_000]]);
  });

  it('a worker document takes them from the answer that brought the runs', async () => {
    // One answer, read from one retained layout: a publication between two
    // requests cannot give the runs and the anchors different layouts.
    const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const run: DocxTextRunInfo = {
      source, sourceRunIndex: 0, text: 'Marked', x: 10, y: 20, w: 30, h: 12, fontSize: 12, font: '12px sans-serif',
    };
    const range: CommentAnchorRange = {
      commentId: '1', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    };
    const doc = Object.create(DocxDocument.prototype) as DocxDocument;
    const posted: { type: string; commentAnchors?: true }[] = [];
    Object.assign(doc, {
      _mode: 'worker', _review: null,
      _meta: { comments: [{ id: '1', text: 'c' }], revisions: [], commentAnchorRanges: [] },
      _bridge: {
        request: async (build: (id: number) => { type: string; commentAnchors?: true }) => {
          const request = build(1);
          posted.push(request);
          // Only the answer to a request for both carries the range.
          return {
            type: 'runsCollected', id: 1, runs: [run],
            ...(request.commentAnchors ? { commentAnchorRanges: [range] } : {}),
          };
        },
      },
    });
    attachDocumentLayoutRuntime(doc, 0);
    const threads = await doc.getCommentThreads(0, { showTrackedChanges: true });
    expect(posted).toEqual([expect.objectContaining({ type: 'collectRuns', commentAnchors: true })]);
    expect(threads).toHaveLength(1);
    // The active view asks for the runs alone and reads the document's anchors.
    posted.length = 0;
    expect(await doc.getCommentThreads(0)).toEqual([]);
    expect(posted).toEqual([expect.not.objectContaining({ commentAnchors: true })]);
  });

  it('runs keep the layout that produced them when the document changes its view', async () => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    const { doc } = documentOf(model, layout);
    const final = await doc.collectPageRuns(0);
    // The document now shows the markup view; its anchors are that view's.
    documentLayoutRuntimeOf(doc).activeLayoutOptions = normalizeLayoutOptions(undefined, 0, true);
    Object.assign(doc, { _commentAnchorRanges: null, _reviewProjectionIndex: null });
    const markup = await doc.collectPageRuns(0);
    expect(markup.some((run) => run.text.startsWith('Deleted'))).toBe(true);
    const anchors = doc.commentAnchorRanges();
    const after = final.find((run) => run.text === 'After')!;
    // The runs collected in the final view still show the point where the
    // final layout places it.
    expect(resolveCommentAnchorRuns(anchors[0]!, final).map((run) => run.text)).toEqual(['After']);
    const [thread] = resolveDocxCommentThreads(doc.comments, anchors, final);
    expect(thread!.anchors[0]!.rects[0]!.x).toBeCloseTo(after.x, 6);
    expect(resolveCommentAnchorRuns(anchors[0]!, markup)[0]!.text.startsWith('Deleted')).toBe(true);
  });

  it('asked of a worker document leave no anchors behind on the host', async () => {
    const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const anchor = (): CommentAnchorRange => ({
      commentId: '1', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    });
    const doc = Object.create(DocxDocument.prototype) as DocxDocument;
    const requests: string[] = [];
    Object.assign(doc, {
      _mode: 'worker', _review: null,
      _meta: { comments: [{ id: '1', text: 'c' }], revisions: [], commentAnchorRanges: [anchor()] },
      _bridge: {
        request: async (build: (id: number) => { type: string }) => {
          const request = build(1);
          requests.push(request.type);
          return { type: 'runsCollected', id: 1, runs: [], commentAnchorRanges: [anchor()] };
        },
      },
    });
    attachDocumentLayoutRuntime(doc, 0);
    for (let day = 1; day <= 20; day += 1) await doc.getCommentThreads(0, { currentDate: day * 86_400_000 });
    // One request per call: the worker keeps the anchors with its layouts.
    expect(requests).toEqual(Array.from({ length: 20 }, () => 'collectRuns'));
    const retained = Object.values(doc as unknown as Record<string, unknown>)
      .filter((value) => value instanceof Map || value instanceof Set || Array.isArray(value))
      .map((value) => (value instanceof Map || value instanceof Set ? value.size : (value as unknown[]).length));
    expect(retained.every((size) => size === 0)).toBe(true);
  });

  it('are collected by the render worker from the layout of that view', async () => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    const { services, variants } = documentOf(model, layout);
    const retained = { layoutServices: services, layoutVariants: variants.store, defaultCurrentDateMs: 0 };
    const review = { comments: model.comments ?? [], revisions: [] };
    const collected = structuredClone(renderWorkerCommentAnchors(retained, review, 0, true));
    adoptCommentAnchorPlacements(collected.commentAnchorRanges!, collected.commentAnchorPlacements);
    const markupLayout = variants.store.layoutFor({ currentDateMs: 0, showTrackedChanges: true });
    const markup = textRunsForPage(markupLayout as DocumentLayout, 0, { scale: 1 });
    const [thread] = resolveDocxCommentThreadsOnPage(model.comments ?? [], collected.commentAnchorRanges!, markup, 0);
    expect(thread!.anchors[0]!.rects[0]!.x).toBeCloseTo(markup.find((run) => run.text.startsWith('Deleted'))!.x, 6);
    // Kept with the layout: a second request does not collect again.
    expect(renderWorkerCommentAnchors(retained, review, 0, true).commentAnchorRanges)
      .toBe(renderWorkerCommentAnchors(retained, review, 0, true).commentAnchorRanges);
  });
});

describe('navigation to a comment that starts in clipped text and ends on a later page', () => {
  const exactRowCell = (rowTwips: number, paragraph: string): string => (
    '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    + `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${rowTwips}" w:hRule="exact"/></w:trPr>`
    + `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl>`);
  const body = exactRowCell(240, '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="800" w:lineRule="exact"/></w:pPr>'
    + `<w:commentRangeStart w:id="1"/>${sized('Hidden')}</w:p>`)
    + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${sized('Shown')}<w:commentRangeEnd w:id="1"/>${REFERENCE}</w:p>`;

  const navigate = async (anchors: (layout: DocumentLayout, model: DocxDocumentModel) => readonly CommentAnchorRange[],
    pageIndex?: number, document = body, parts: Parameters<typeof paintedDocument>[1] = COMMENT_PART, scale = 1) => {
    const { layout, model } = await paintedDocument(document, parts);
    const ranges = anchors(layout, model);
    const doc = {
      comments: model.comments ?? [],
      pageCount: layout.pages.length,
      layoutComplete: true,
      commentAnchorRanges: () => ranges,
      collectPageRuns: async (page: number) => textRunsForPage(layout, page, { scale }),
    } as unknown as DocxDocument;
    const select = vi.fn<(commentId: string, page: number, point: Readonly<{ x: number; y: number }>) => void>();
    const navigation = new DocxScrollCommentNavigation({
      document: () => doc,
      destroyed: () => false,
      scale: () => scale,
      pageWidth: () => 612,
      currentDate: () => undefined,
      showTrackedChanges: () => false,
      waitForLayout: async () => undefined,
      select,
    });
    const reached = await navigation.goToComment('1', pageIndex === undefined ? undefined : { pageIndex });
    return { reached, select, layout };
  };
  /** Whether `point` lies in the box of the run `text` on `page`. */
  const inRun = (layout: DocumentLayout, page: number, text: string, point: Readonly<{ x: number; y: number }>) => {
    const run = textRunsForPage(layout, page, { scale: 1 }).find((candidate) => candidate.text === text)!;
    return point.x > run.x && point.x < run.x + run.w && point.y > run.y && point.y < run.y + run.h;
  };

  it('goes to the page of the card, at the text that is seen', async () => {
    const { reached, select, layout } = await navigate(placedAnchors);
    expect(layout.pages).toHaveLength(2);
    expect(reached).toBe(true);
    expect(select).toHaveBeenCalledTimes(1);
    expect(select.mock.calls[0]![1]).toBe(1);
    expect(inRun(layout, 1, 'Shown', select.mock.calls[0]![2])).toBe(true);
  });

  it('finds that page for anchors that carry no card placement', async () => {
    const { reached, select, layout } = await navigate((layout, model) => {
      const index = reviewProjectionIndexForDocument(layout);
      return collectLayoutSourceCommentRangesIfPresent(model.comments, layoutSourceStore(model), index.renderedRunIndex);
    });
    expect(reached).toBe(true);
    expect(select.mock.calls[0]![1]).toBe(1);
    expect(inRun(layout, 1, 'Shown', select.mock.calls[0]![2])).toBe(true);
  });

  it('goes to the part of a picture that its fixed cell shows, not to the middle of the picture', async () => {
    // A 450pt picture in a fixed 100pt cell, at 2 CSS px per point: the cell
    // shows its first 100pt (page x 144..344 px); the middle of the picture's
    // own box (x 594 px) lies where the cell hides it.
    const picture = exactRowCell(1200, `<w:p><w:commentRangeStart w:id="1"/>${inlinePictureRun(450, 40)}`
      + `<w:commentRangeEnd w:id="1"/>${REFERENCE}</w:p>`).replace('<w:tblCellMar>', '<w:tblCellMar>'
      + '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>') + '<w:p/>';
    const { reached, select } = await navigate(placedAnchors, undefined, picture, {
      parts: { ...COMMENT_PART.parts, ...PICTURE_PART.parts },
      relationships: COMMENT_PART.relationships + PICTURE_PART.relationships,
      contentTypes: COMMENT_PART.contentTypes + PICTURE_PART.contentTypes,
    }, 2);
    expect(reached).toBe(true);
    const point = select.mock.calls[0]![2];
    expect(Object.keys(point).sort()).toEqual(['x', 'y']);
    expect(point.x).toBeGreaterThan(144);
    expect(point.x).toBeLessThan(344);
    expect(point.y).toBeGreaterThan(144);
    expect(point.y).toBeLessThan(144 + 80);
  });

  it('goes to a point that the text box rectangles of the picture show, not to the gap between them', async () => {
    // A 400px inline object under 257 disjoint rectangles in two groups,
    // x 0..129 and 272..400: more than are tracked as separate pieces, so
    // their extent stands for them, and its middle (x 200) lies in the gap.
    const strips = [...Array.from({ length: 129 }, (_, index) => index),
      ...Array.from({ length: 128 }, (_, index) => 272 + index)];
    const clip: DocxClipRegion = {
      x: 50, y: 100, width: 400, height: 20,
      rects: strips.map((x) => ({ x, y: 0, width: 0.5, height: 20 })),
    };
    const source = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const run: DocxTextRunInfo = {
      source, sourceRunIndex: 0, text: '', inlineContent: 'object', x: 50, y: 100, w: 400, h: 20,
      fontSize: 0, font: '', clips: [clip],
    };
    const doc = {
      comments: [{ id: '1', text: 'c' }],
      pageCount: 1,
      layoutComplete: true,
      commentAnchorRanges: () => [{
        commentId: '1', source, startRunIndex: 0, endRunIndex: 1,
        reference: { source, runIndex: 1, affinity: 'preceding' as const },
      }],
      collectPageRuns: async () => [run],
    } as unknown as DocxDocument;
    const select = vi.fn<(commentId: string, page: number, point: Readonly<{ x: number; y: number }>) => void>();
    const navigation = new DocxScrollCommentNavigation({
      document: () => doc, destroyed: () => false, scale: () => 1, pageWidth: () => 612,
      currentDate: () => undefined, showTrackedChanges: () => false, waitForLayout: async () => undefined, select,
    });
    expect(await navigation.goToComment('1')).toBe(true);
    const point = select.mock.calls[0]![2];
    expect(strips.some((x) => point.x > 50 + x && point.x < 50 + x + 0.5)).toBe(true);
    expect(point.y).toBeGreaterThan(100);
    expect(point.y).toBeLessThan(120);
  });

  it('goes to the tint of a comment that starts with a tab, not to the tab\'s line box', async () => {
    // A tab and a 12pt word are commented; 80pt text shares their line, so
    // the line box is far taller than the tint.
    const { reached, select, layout } = await navigate(placedAnchors, undefined,
      `<w:p><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:tab/></w:r>${sized('word')}`
      + `<w:commentRangeEnd w:id="1"/>${REFERENCE}<w:r><w:rPr><w:sz w:val="160"/></w:rPr><w:t>Tall</w:t></w:r></w:p>`);
    expect(reached).toBe(true);
    const point = select.mock.calls[0]![2];
    const { model } = await paintedDocument(
      `<w:p><w:commentRangeStart w:id="1"/><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:tab/></w:r>${sized('word')}`
      + `<w:commentRangeEnd w:id="1"/>${REFERENCE}<w:r><w:rPr><w:sz w:val="160"/></w:rPr><w:t>Tall</w:t></w:r></w:p>`,
      COMMENT_PART);
    const rects = resolveDocxCommentThreadsOnPage(model.comments ?? [], placedAnchors(layout, model),
      textRunsForPage(layout, 0, { scale: 1 }), 0).flatMap((thread) => thread.anchors.flatMap((anchor) => anchor.rects));
    expect(rects.length).toBeGreaterThan(0);
    expect(rects.some((rect) => point.x >= rect.x && point.x <= rect.x + rect.width
      && point.y >= rect.y && point.y <= rect.y + rect.height)).toBe(true);
  });

  it('does not go to a page that shows nothing of the comment', async () => {
    const { reached, select } = await navigate(placedAnchors, 0);
    expect(reached).toBe(false);
    expect(select).not.toHaveBeenCalled();
  });
});

describe('find beside text that a clip hides', () => {
  it('matches the letters that are seen, not a hidden letter and one of them', async () => {
    // "A" below a 12pt exact-height row is hidden; "AA" follows the table.
    const { runs } = await paintedDocument(
      '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
      + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="240" w:hRule="exact"/></w:trPr>'
      + '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" '
      + `w:line="800" w:lineRule="exact"/></w:pPr>${sized('A')}</w:p></w:tc></w:tr></w:tbl><w:p>${sized('AA')}</w:p>`);
    expect(runs.map((run) => run.text).join('')).toBe('AAA');
    const visible = runs.findIndex((run) => run.text === 'AA');
    const find = new DocxFindController(() => 1, async () => runs);
    expect(await find.find('AA')).toHaveLength(1);
    expect(find.pageHighlights(0).map((highlight) => highlight.slices)).toEqual(
      [[{ runIndex: visible, start: 0, end: 2 }]]);
    // The hidden letter alone is no match; each visible one is.
    expect(await find.find('A')).toHaveLength(2);
  });
});

describe('comment anchors read from a prefix the render worker has published', () => {
  it('are projected as from a prefix: no fallback stands in for a paragraph still to come', async () => {
    // "Unrelated" on page one; "Before" and an empty paragraph holding a
    // reference-only comment on page two; more pages follow.
    const filler = Array.from({ length: 400 }, (_, index) => `<w:p>${sized(`filler ${index}`)}</w:p>`).join('');
    const { model } = await paintedDocument(`<w:p>${sized('Unrelated')}</w:p>`
      + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${sized('Before')}</w:p><w:p>${REFERENCE}</w:p>${filler}`, COMMENT_PART);
    const source = layoutSourceStore(model);
    const retained = retainRenderWorkerDocumentLayout(
      source, createLayoutServices(source, { measureContext: oracleMeasureContext() }), 0);
    const review = { comments: model.comments ?? [], revisions: [] };
    const seen: { pages: number; published: number; read: number; selected: number }[] = [];
    await paginateRenderWorkerDocumentProgressively(retained, source, {
      publish: (publication) => {
        // What a request for the runs and anchors of this view reads now.
        const read = renderWorkerCommentAnchors(retained, review, 0, false);
        seen.push({
          pages: publication.pageCount,
          published: publication.commentAnchorRanges?.length ?? 0,
          read: read.commentAnchorRanges?.length ?? 0,
          // What selecting this view again (tracked changes on, then off)
          // gives the host as its metadata.
          selected: renderWorkerLayoutMeta(retained, review, 0, false).commentAnchorRanges?.length ?? 0,
        });
      },
      progress: () => undefined,
    }, { currentDateMs: 0 }, undefined, review);
    const first = seen.find((publication) => publication.pages === 1);
    expect(first).toEqual({ pages: 1, published: 0, read: 0, selected: 0 });
    // The complete layout places the comment.
    expect(renderWorkerCommentAnchors(retained, review, 0, false).commentAnchorRanges).toHaveLength(1);
    expect(renderWorkerLayoutMeta(retained, review, 0, false).commentAnchorRanges).toHaveLength(1);
  });
});

describe('text an exact-height row hides, in the text layer', () => {
  it('takes no pointer: its box does not cover what the row shows', async () => {
    const { runs } = await paintedDocument(
      '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
      + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="240" w:hRule="exact"/></w:trPr>'
      + `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${sized('Visible')}<w:r><w:br/></w:r>${sized('Secret')}</w:p>`
      + '</w:tc></w:tr></w:tbl><w:p/>');
    const spans: FakeElement[] = [];
    const collect = (element: FakeElement): void => {
      if (element.children.length === 0 && (element as { textContent?: string }).textContent) spans.push(element);
      element.children.forEach(collect);
    };
    withFakeDocument(() => {
      const layer = fakeElement();
      buildDocxTextLayer(layer as never, runs, 612, 792);
      collect(layer);
    });
    const of = (value: string) => spans.filter((span) => (
      span.dataset.runIndex === String(runs.findIndex((run) => run.text === value))));
    expect(of('Secret').length).toBeGreaterThan(0);
    for (const span of of('Secret')) {
      expect(span.style.cssText).toContain('pointer-events:none');
      expect(span.style.cssText).not.toContain('pointer-events:all');
    }
    for (const span of of('Visible')) expect(span.style.cssText).toContain('pointer-events:all');
  });
});
