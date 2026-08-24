import {
  resolveCommentAnchorRuns,
  resolveRevisionAnchorRuns,
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

  const comments = doc.commentAnchorRanges().flatMap((anchor) => {
    const comment = doc.comments.find((item) => item.id === anchor.commentId);
    const geometry = resolveCommentAnchorRuns(anchor, runs);
    return comment && geometry.length > 0 ? [{ comment, anchor, geometry }] : [];
  });

  const revisions = doc.revisionAnchorRanges().flatMap((anchor) => {
    const revision = doc.revisions[anchor.revisionIndex];
    const geometry = resolveRevisionAnchorRuns(anchor, runs);
    if (!revision || geometry.length === 0) return [];
    return [{
      revision,
      geometry,
      // deletion/moveFrom geometry is a nearby final-state position, not its old range.
      isPositionHint: revision.kind === 'deletion' || revision.kind === 'moveFrom',
    }];
  });

  return { pageIndex, comments, revisions };
}
