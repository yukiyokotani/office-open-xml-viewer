import { describe, expect, it, vi } from 'vitest';
import {
  adoptCommentAnchorPlacements,
  collectLayoutSourceCommentRangesIfPresent,
  commentAnchorPlacement,
  resolveDocxCommentThreads,
  resolveDocxCommentThreadsOnPage,
  type CommentAnchorRange,
} from './comments.js';
import { buildDocxCommentMargin } from './comment-margin.js';
import { DocxDocument } from './document.js';
import { layoutDocumentInput } from './layout/document.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { reviewProjectionIndexForDocument } from './layout/text-index.js';
import type { DocumentLayout } from './layout/types.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { projectRenderWorkerLayoutMeta, renderWorkerCommentAnchors } from './render-worker-metadata.js';
import { DocxScrollCommentNavigation } from './scroll-comment-navigation.js';
import { installDom, makeEl } from './scroll-viewer-test-dom.js';
import { oracleMeasureContext, paintedDocument } from './test-support/painted-glyphs.test-support.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxDocumentModel, DocxTextRunInfo } from './types.js';

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
  const variants = attachDocumentLayoutVariants({
    source,
    services,
    defaultCurrentDateMs: 0,
    buildLayout: (options) => (options.showTrackedChanges === true
      ? layoutDocumentInput(source.bodyLayoutInput, services, options)
      : layout),
  });
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, {
    _mode: 'main', _document: model, _source: source, _meta: null, _review: null,
    _commentAnchorRanges: null, _reviewProjectionIndex: null,
  });
  attachDocumentLayoutRuntime(doc, 0);
  documentLayoutRuntimeOf(doc).services = services;
  return { doc, services, variants };
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
    // The active view's anchors are untouched.
    const [again] = await doc.getCommentThreads(0);
    expect(again!.anchors[0]!.rects[0]!.x).toBeCloseTo(shown!.anchors[0]!.rects[0]!.x, 6);
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
    pageIndex?: number) => {
    const { layout, model } = await paintedDocument(body, COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    const ranges = anchors(layout, model);
    const doc = {
      comments: model.comments ?? [],
      pageCount: layout.pages.length,
      layoutComplete: true,
      commentAnchorRanges: () => ranges,
      collectPageRuns: async (page: number) => textRunsForPage(layout, page, { scale: 1 }),
    } as unknown as DocxDocument;
    const select = vi.fn<(commentId: string, page: number, run: Readonly<DocxTextRunInfo>) => void>();
    const navigation = new DocxScrollCommentNavigation({
      document: () => doc,
      destroyed: () => false,
      scale: () => 1,
      pageWidth: () => 612,
      currentDate: () => undefined,
      showTrackedChanges: () => false,
      waitForLayout: async () => undefined,
      select,
    });
    const reached = await navigation.goToComment('1', pageIndex === undefined ? undefined : { pageIndex });
    return { reached, select };
  };

  it('goes to the page of the card, at the text that is seen', async () => {
    const { reached, select } = await navigate(placedAnchors);
    expect(reached).toBe(true);
    expect(select).toHaveBeenCalledTimes(1);
    expect(select.mock.calls[0]![1]).toBe(1);
    expect(select.mock.calls[0]![2].text).toBe('Shown');
  });

  it('finds that page for anchors that carry no card placement', async () => {
    const { reached, select } = await navigate((layout, model) => {
      const index = reviewProjectionIndexForDocument(layout);
      return collectLayoutSourceCommentRangesIfPresent(model.comments, layoutSourceStore(model), index.renderedRunIndex);
    });
    expect(reached).toBe(true);
    expect(select.mock.calls[0]![1]).toBe(1);
    expect(select.mock.calls[0]![2].text).toBe('Shown');
  });

  it('does not go to a page that shows nothing of the comment', async () => {
    const { reached, select } = await navigate(placedAnchors, 0);
    expect(reached).toBe(false);
    expect(select).not.toHaveBeenCalled();
  });
});
