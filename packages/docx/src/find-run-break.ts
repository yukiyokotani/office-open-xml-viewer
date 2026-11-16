import type { DocxTextRunInfo } from './renderer';

/**
 * The unit vectors of a run's inline and block axes on the page. A run reports
 * its physical top-left and the CSS transform its overlay box applies about
 * that point: none for horizontal text, a quarter turn for a vertical page or
 * cell, or a `matrix(a, b, c, d, 0, 0)` for a rotated text box.
 */
function runAxes(transform: string | undefined): readonly [number, number, number, number] {
  if (transform === 'rotate(90deg)') return [0, 1, -1, 0];
  if (transform === 'rotate(-90deg)') return [0, -1, 1, 0];
  const matrix = transform && /^matrix\(([^)]*)\)$/.exec(transform);
  if (matrix) {
    const [a, b, c, d] = matrix[1].split(',').map(Number);
    if ([a, b, c, d].every(Number.isFinite)) return [a, b, c, d];
  }
  return [1, 0, 0, 1];
}

/**
 * Whether layout separates two runs that are adjacent in a page's run list, so
 * a whole-word match may end at the first and start at the second although the
 * joined search text has no character between them.
 *
 * Runs of different paragraphs are separate. Within a paragraph the layout
 * reports each run's range in the paragraph text (`textRange`), and that alone
 * decides: runs adjacent in the text join, wherever they are placed, so a long
 * word that wraps onto a second line stays one word; between runs that are not
 * adjacent lies a tab, a manual line break or an inline object, and they are
 * separate.
 *
 * Runs without `textRange` are compared by position: another orientation,
 * another line, or a gap of more than a quarter of the font size on one line
 * separates. Positions are taken in the runs' own writing frame, so this is
 * the same for horizontal, vertical and rotated text.
 */
export function docxRunsBreak(previous: DocxTextRunInfo, next: DocxTextRunInfo): boolean {
  const a = previous.source;
  const b = next.source;
  if (a && b && (
    a.story !== b.story || a.storyInstance !== b.storyInstance ||
    a.path.length !== b.path.length || a.path.some((step, index) => step !== b.path[index])
  )) return true;
  const from = previous.textRange;
  const to = next.textRange;
  // Adjacent in either order: a right-to-left line lists its runs against the text.
  if (from && to) return from.end !== to.start && to.end !== from.start;
  if ((previous.transform ?? '') !== (next.transform ?? '')) return true;
  const [inlineX, inlineY, blockX, blockY] = runAxes(previous.transform);
  const inline = (run: DocxTextRunInfo): number => run.x * inlineX + run.y * inlineY;
  const block = (run: DocxTextRunInfo): number => run.x * blockX + run.y * blockY;
  if (Math.abs(block(previous) - block(next)) > 0.5) return true;
  const gap = Math.max(
    inline(next) - (inline(previous) + previous.w), inline(previous) - (inline(next) + next.w));
  return gap > 0.25 * Math.min(previous.fontSize, next.fontSize);
}
