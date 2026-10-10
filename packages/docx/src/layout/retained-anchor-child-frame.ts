import type { AnchorAcquisitionInput } from './anchor-input.js';
import type { LayoutRect, Matrix2DData } from './types.js';
import { transformRect } from './coordinate-space.js';

/** Shared retained-group frame projection: normal anchors and an explicit
 * relocated reading scene consume the same acquired child geometry. */
export function retainedAnchorChildFrame(
  acquisition: AnchorAcquisitionInput,
  outerFrame: LayoutRect,
  coordinateSpace?: Readonly<{
    physicalToLogical: Matrix2DData;
    logicalToPhysical: Matrix2DData;
  }>,
): LayoutRect {
  const child = acquisition.group?.resolvedChildFrame;
  if (!child) return outerFrame;
  const authoredWidthPt = acquisition.extent.widthPt;
  const authoredHeightPt = acquisition.extent.heightPt;
  if (
    acquisition.extent.widthStatus !== 'valid'
    || acquisition.extent.heightStatus !== 'valid'
    || authoredWidthPt === null
    || authoredHeightPt === null
    || authoredWidthPt <= 0
    || authoredHeightPt <= 0
  ) {
    throw new Error('resolved grouped anchor requires its authored wp:extent');
  }
  const physicalOuter = coordinateSpace === undefined
    ? outerFrame
    : transformRect(coordinateSpace.logicalToPhysical, outerFrame);
  const scaleX = physicalOuter.widthPt / authoredWidthPt;
  const scaleY = physicalOuter.heightPt / authoredHeightPt;
  const physicalChild = {
    xPt: physicalOuter.xPt + child.offsetXPt * scaleX,
    yPt: physicalOuter.yPt + child.offsetYPt * scaleY,
    widthPt: child.widthPt * scaleX,
    heightPt: child.heightPt * scaleY,
  };
  return coordinateSpace === undefined
    ? physicalChild
    : transformRect(coordinateSpace.physicalToLogical, physicalChild);
}
