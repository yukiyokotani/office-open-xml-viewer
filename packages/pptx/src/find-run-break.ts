import type { PptxTextRunInfo } from './renderer';

/**
 * Whether layout separates two runs that are adjacent in a slide's run list, so
 * a whole-word match may end at the first and start at the second although the
 * joined search text has no character between them: the runs belong to
 * different shapes or table cells, lie on different lines of a text body, or
 * are apart by more than a quarter of the font size on the same line (a tab
 * stop). Runs that merely continue a line (a formatting change inside a word)
 * join. In stacked vertical text every glyph is a run of its own and a line is
 * a column, so the glyph cells of one column join and a change of column
 * separates.
 */
export function pptxRunsBreak(previous: PptxTextRunInfo, next: PptxTextRunInfo): boolean {
  if (
    (previous.renderElementIndex ?? previous.elementIndex) !== (next.renderElementIndex ?? next.elementIndex) ||
    previous.origin !== next.origin || previous.shapeId !== next.shapeId ||
    previous.shapeX !== next.shapeX || previous.shapeY !== next.shapeY ||
    previous.tableCell?.row !== next.tableCell?.row || previous.tableCell?.column !== next.tableCell?.column
  ) return true;
  if (previous.stackedColumn !== undefined || next.stackedColumn !== undefined) {
    return previous.stackedColumn !== next.stackedColumn;
  }
  if (Math.abs(previous.inShapeY - next.inShapeY) > 0.5) return true;
  const gap = Math.max(
    next.inShapeX - (previous.inShapeX + previous.w), previous.inShapeX - (next.inShapeX + next.w));
  return gap > 0.25 * Math.min(previous.fontSize, next.fontSize);
}
