interface CommentSurfaceScaleInput {
  readonly currentScale: number;
  readonly viewportWidth: number;
  readonly contentWidth: number;
  readonly leadingExtent: number;
  readonly trailingExtent: number;
}

interface CommentDemoZoomViewer {
  getScale(): number;
  setScale(scale: number): void;
}

/**
 * Fit a centered document surface plus its one-sided review margin.
 *
 * ScrollViewer centers the authored page/slide independently of comment cards,
 * so the larger side extent must also be reserved on the opposite side. This
 * keeps the authored surface centered while making the adjacent margin visible.
 */
export function commentSurfaceFitScale(
  input: Readonly<CommentSurfaceScaleInput>,
): number | null {
  const values = [
    input.currentScale,
    input.viewportWidth,
    input.contentWidth,
    input.leadingExtent,
    input.trailingExtent,
  ];
  if (!values.every(Number.isFinite)) return null;
  if (input.currentScale <= 0 || input.viewportWidth <= 0 || input.contentWidth <= 0) {
    return null;
  }
  const sideExtent = Math.max(0, input.leadingExtent, input.trailingExtent);
  const centeredSurfaceWidth = input.contentWidth + sideExtent * 2;
  if (centeredSurfaceWidth <= input.viewportWidth) return null;
  const scale = input.currentScale * input.viewportWidth / centeredSurfaceWidth;
  return Number.isFinite(scale) && scale > 0 ? scale : null;
}

/** Shrink only the official comment demo's initial DOCX/PPTX view. */
export function fitVisibleCommentSurface(
  viewer: CommentDemoZoomViewer,
  host: HTMLElement,
): boolean {
  const margin = [...host.querySelectorAll<HTMLElement>(
    '[data-ooxml-comment-ui="margin"]',
  )].find((candidate) => candidate.querySelector('[data-ooxml-comment-card]') !== null);
  const wrapper = margin?.parentElement;
  const canvas = wrapper?.querySelector<HTMLCanvasElement>(':scope > canvas');
  const viewport = wrapper?.parentElement;
  if (!margin || !canvas || !viewport) return false;

  const content = canvas.getBoundingClientRect();
  const review = margin.getBoundingClientRect();
  const scale = commentSurfaceFitScale({
    currentScale: viewer.getScale(),
    viewportWidth: viewport.clientWidth,
    contentWidth: content.width,
    leadingExtent: content.left - review.left,
    trailingExtent: review.right - content.right,
  });
  if (scale === null) return false;
  viewer.setScale(scale);
  return true;
}
