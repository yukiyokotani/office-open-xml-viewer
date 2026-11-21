import type { PhysicalAnchorFrame } from './acquisition-context.js';
import {
  physicalToLogicalMatrix,
  transformRect,
  transformRectEdges,
} from './coordinate-space.js';
import type { LayoutRect, WritingMode } from './types.js';

/**
 * Projections between a vertical section's upright physical page and its
 * section-logical frame for DrawingML anchors and upright tables (§17.6.20).
 *
 * Each projection is the section frame's own canonical matrix
 * (coordinate-space.ts logicalToPhysicalMatrix/physicalToLogicalMatrix),
 * written per frame so that extents are carried as extents rather than
 * recovered by differencing transformed corners. The vertical-rl
 * (clockwise) and sideways-lr (counter-clockwise) formulas keep the
 * arithmetic order their callers have always used, so their results are
 * unchanged to the bit; vertical-lr (tbLrV) is the reflection
 * physical (x, y) = logical (y, x).
 */

type Box = Readonly<{ x: number; y: number; w: number; h: number }>;

function unsupportedFrame(writingMode: WritingMode): never {
  throw new RangeError(`An upright physical frame requires a vertical section, got ${writingMode}`);
}

/** Physical y of the top of a logical inline band [inlineStart, inlineStart
 * + inlineExtent]: the band start where logical x maps to physical y
 * (vertical-rl, vertical-lr), and the band end in the counter-clockwise frame
 * (physical y = page height - logical x). */
export function physicalInlineBandTopPt(
  frame: PhysicalAnchorFrame,
  inlineStartPt: number,
  inlineExtentPt: number,
): number {
  switch (frame.writingMode) {
    case 'vertical-rl':
    case 'vertical-lr':
      return inlineStartPt;
    case 'sideways-lr':
      return frame.pageHeight - (inlineStartPt + inlineExtentPt);
    default:
      return unsupportedFrame(frame.writingMode);
  }
}

/** A box resolved on the upright physical page, in the section-logical frame
 * (width and height swap). */
export function physicalAnchorBoxToLogical(frame: PhysicalAnchorFrame, box: Box): Box {
  switch (frame.writingMode) {
    case 'vertical-rl':
      return { x: box.y, y: frame.pageWidth - (box.x + box.w), w: box.h, h: box.w };
    case 'vertical-lr':
      return { x: box.y, y: box.x, w: box.h, h: box.w };
    case 'sideways-lr':
      return { x: frame.pageHeight - (box.y + box.h), y: box.x, w: box.h, h: box.w };
    default:
      return unsupportedFrame(frame.writingMode);
  }
}

/** Physical dist* padding relabelled with the logical edges it lands on. */
export function physicalDistToLogical(
  frame: PhysicalAnchorFrame,
  dist: Readonly<{ dl: number; dr: number; dt: number; db: number }>,
): { dl: number; dr: number; dt: number; db: number } {
  const edges = transformRectEdges(
    physicalToLogicalMatrix(frame.writingMode, { widthPt: frame.pageWidth, heightPt: frame.pageHeight }),
    { top: dist.dt, right: dist.dr, bottom: dist.db, left: dist.dl },
  );
  return { dl: edges.left, dr: edges.right, dt: edges.top, db: edges.bottom };
}

/** Physical origin of an upright block, such as an upright table or owner
 * segment, whose physical width is its logical block extent and whose
 * physical height is its logical inline extent, starting at the logical
 * `cursorPt` (column top, current block position). The clockwise frame places
 * it at physical x = page width − logical y − width, y = logical x; the
 * vertical-lr frame of tbLrV at x = logical y, y = logical x, following its
 * left-to-right line progression; a counter-clockwise frame at x = logical y
 * with the block ending at the column bottom. The clockwise and vertical-lr
 * placements are `word-vertical-section-upright-block-table`; the
 * counter-clockwise placement is generic library policy transformed through
 * the section's own frame, which no comparison with a Word-produced reference
 * has established. */
export function uprightBlockPhysicalOrigin(
  frame: PhysicalAnchorFrame,
  cursorPt: Readonly<{ xPt: number; yPt: number }>,
  physicalWidthPt: number,
  physicalHeightPt: number,
): Readonly<{ xPt: number; yPt: number }> {
  switch (frame.writingMode) {
    case 'vertical-rl':
      return { xPt: frame.pageWidth - cursorPt.yPt - physicalWidthPt, yPt: cursorPt.xPt };
    case 'vertical-lr':
      return { xPt: cursorPt.yPt, yPt: cursorPt.xPt };
    case 'sideways-lr':
      return { xPt: cursorPt.yPt, yPt: frame.pageHeight - cursorPt.xPt - physicalHeightPt };
    default:
      return unsupportedFrame(frame.writingMode);
  }
}

/** A rectangle of the upright physical page in the section-logical frame, for
 * upright table exclusions. */
export function uprightPhysicalRectToLogical(
  frame: PhysicalAnchorFrame,
  rect: LayoutRect,
): LayoutRect {
  switch (frame.writingMode) {
    case 'vertical-rl':
      return Object.freeze({
        xPt: rect.yPt,
        yPt: frame.pageWidth - rect.xPt - rect.widthPt,
        widthPt: rect.heightPt,
        heightPt: rect.widthPt,
      });
    case 'vertical-lr':
      return Object.freeze({
        xPt: rect.yPt,
        yPt: rect.xPt,
        widthPt: rect.heightPt,
        heightPt: rect.widthPt,
      });
    case 'sideways-lr':
      return Object.freeze(transformRect(
        physicalToLogicalMatrix('sideways-lr', { widthPt: frame.pageWidth, heightPt: frame.pageHeight }),
        rect,
      ));
    default:
      return unsupportedFrame(frame.writingMode);
  }
}
