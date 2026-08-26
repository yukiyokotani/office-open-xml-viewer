import {
  resolveDocxCommentThreads,
  type DocxDocument,
  type DocxTextRunInfo,
} from '@silurus/ooxml/docx';

export async function renderReviewPage(
  doc: DocxDocument,
  canvas: HTMLCanvasElement,
  requestedPage: number,
) {
  if (doc.pageCount === 0) throw new Error('The DOCX has no renderable pages.');
  const pageIndex = Math.max(0, Math.min(requestedPage, doc.pageCount - 1));
  const runs: DocxTextRunInfo[] = [];

  await doc.renderPage(canvas, pageIndex, {
    width: 760,
    onTextRun: (run) => runs.push(run),
  });

  const comments = resolveDocxCommentThreads(
    doc.comments,
    doc.commentAnchorRanges(),
    runs,
  );

  return { pageIndex, comments };
}
