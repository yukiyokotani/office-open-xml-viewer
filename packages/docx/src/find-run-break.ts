import type { DocxTextRunInfo } from './renderer';

/**
 * Whether layout separates two runs that are adjacent in a page's run list, so
 * a whole-word match may end at the first and start at the second although the
 * joined search text has no character between them: the runs belong to
 * different paragraphs, lie on different lines, or are apart by more than a
 * quarter of the font size on the same line (a tab stop or a column gap).
 * Runs that merely continue a line (a formatting change inside a word) join.
 */
export function docxRunsBreak(previous: DocxTextRunInfo, next: DocxTextRunInfo): boolean {
  const a = previous.source;
  const b = next.source;
  if (a && b && (
    a.story !== b.story || a.storyInstance !== b.storyInstance ||
    a.path.length !== b.path.length || a.path.some((step, index) => step !== b.path[index])
  )) return true;
  // A vertical page reports physical positions: a line is a column, so its
  // runs share x and advance along y.
  if (previous.transform || next.transform) return Math.abs(previous.x - next.x) > 0.5;
  if (Math.abs(previous.y - next.y) > 0.5) return true;
  const gap = Math.max(next.x - (previous.x + previous.w), previous.x - (next.x + next.w));
  return gap > 0.25 * Math.min(previous.fontSize, next.fontSize);
}
