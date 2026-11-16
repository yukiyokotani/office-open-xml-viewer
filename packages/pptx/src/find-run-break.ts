import type { PptxTextRunInfo } from './renderer';

/**
 * Whether layout separates two runs that are adjacent in a slide's run list, so
 * a whole-word match may end at the first and start at the second although the
 * joined search text has no character between them.
 *
 * Runs of different shapes or table cells are separate. Within a text body the
 * renderer reports the authored line of each run (`logicalLine`: a paragraph,
 * or the part of one between manual line breaks): runs of different authored
 * lines are separate, and runs of one authored line join, also across an
 * automatic wrap (a long word broken over two lines stays one word) and down
 * a column of stacked vertical text, where every glyph is a run. On one drawn
 * line, runs apart by more than a quarter of the font size are separate (a
 * tab stop). A run without `logicalLine` is compared by position: another
 * drawn line separates.
 */
export function pptxRunsBreak(previous: PptxTextRunInfo, next: PptxTextRunInfo): boolean {
  if (
    (previous.renderElementIndex ?? previous.elementIndex) !== (next.renderElementIndex ?? next.elementIndex) ||
    previous.origin !== next.origin || previous.shapeId !== next.shapeId ||
    previous.shapeX !== next.shapeX || previous.shapeY !== next.shapeY ||
    previous.tableCell?.row !== next.tableCell?.row || previous.tableCell?.column !== next.tableCell?.column ||
    previous.logicalLine !== next.logicalLine
  ) return true;
  if (Math.abs(previous.inShapeY - next.inShapeY) > 0.5) return previous.logicalLine === undefined;
  const gap = Math.max(
    next.inShapeX - (previous.inShapeX + previous.w), previous.inShapeX - (next.inShapeX + next.w));
  return gap > 0.25 * Math.min(previous.fontSize, next.fontSize);
}
