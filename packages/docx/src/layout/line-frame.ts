import type { LineLayout, Matrix2DData, ParagraphLayout } from './types.js';

/**
 * The line-local paint frame of a retained line, composed inside its
 * paragraph's frame, or null when the line is painted as laid out.
 *
 * A tbLrV section is painted through the vertical-lr frame (physical x =
 * logical y), a reflection. Under `word-section-tblrv-line-strip` a tbLrV
 * line strip has the content arrangement of a tbRl strip and only the strip
 * order differs (left to right), so a `block-reflected` line is reflected
 * across its own block band, y ↦ 2·top + height − y. Composed with the
 * vertical-lr frame this is the clockwise quarter turn of a tbRl page
 * translated onto the strip, which keeps glyphs, inline objects and their
 * decorations unmirrored. Paint applies the same retained contract
 * (paint/canvas-text.ts lineFrameMatrix) without importing layout.
 */
export function lineFrameTransform(
  paragraph: Pick<ParagraphLayout, 'lineFrame'>,
  line: Pick<LineLayout, 'bounds'>,
): Matrix2DData | null {
  if (paragraph.lineFrame !== 'block-reflected') return null;
  return {
    a: 1,
    b: 0,
    c: 0,
    d: -1,
    e: 0,
    f: 2 * line.bounds.yPt + line.bounds.heightPt,
  };
}
