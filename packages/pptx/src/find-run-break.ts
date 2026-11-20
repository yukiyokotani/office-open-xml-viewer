import type { PptxTextRunInfo } from './renderer';

/**
 * Whether layout separates two runs that are adjacent in a slide's run list, so
 * a whole-word match may end at the first and start at the second although the
 * joined search text has no character between them.
 *
 * Runs of different shapes or table cells are separate. Within a text body the
 * renderer reports the span of authored text each run belongs to (`textSpan`),
 * and that alone decides: a span ends at a paragraph end, a manual line break,
 * a tab or an equation, so runs of different spans are separate, and runs of
 * one span join, wherever they are drawn. A long word that wraps onto a second
 * line stays one word, and so do the glyph cells of stacked vertical text,
 * where every glyph is a run.
 *
 * Runs without `textSpan` are compared by position: another drawn line, or a
 * gap of more than a quarter of the font size on one line, separates.
 */
export function pptxRunsBreak(previous: PptxTextRunInfo, next: PptxTextRunInfo): boolean {
  if (
    (previous.renderElementIndex ?? previous.elementIndex) !== (next.renderElementIndex ?? next.elementIndex) ||
    previous.origin !== next.origin || previous.shapeId !== next.shapeId ||
    previous.shapeX !== next.shapeX || previous.shapeY !== next.shapeY ||
    previous.tableCell?.row !== next.tableCell?.row || previous.tableCell?.column !== next.tableCell?.column ||
    previous.textSpan !== next.textSpan
  ) return true;
  if (previous.textSpan !== undefined) return false;
  if (Math.abs(previous.inShapeY - next.inShapeY) > 0.5) return true;
  const gap = Math.max(
    next.inShapeX - (previous.inShapeX + previous.w), previous.inShapeX - (next.inShapeX + next.w));
  return gap > 0.25 * Math.min(previous.fontSize, next.fontSize);
}

/** The module's default export is what the find controller loads on demand. */
export default pptxRunsBreak;
