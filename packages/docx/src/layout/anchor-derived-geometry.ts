import type { AnchorFrameResult } from './anchor-frame.js';
import type { LayoutRect } from './types.js';

/** Rebuild effect/wrap geometry when an authored anchor object frame receives
 * an effective extent (for example shape text auto-fit). Edge effects remain
 * point distances; authored wrap polygons scale with the object coordinate
 * space, matching the anchor solver's §20.4 transaction. */
export function resizeDerivedAnchorRect(
  derived: LayoutRect,
  authored: LayoutRect,
  effective: LayoutRect,
): LayoutRect {
  const leftPt = authored.xPt - derived.xPt;
  const topPt = authored.yPt - derived.yPt;
  const rightPt = derived.xPt + derived.widthPt - authored.xPt - authored.widthPt;
  const bottomPt = derived.yPt + derived.heightPt - authored.yPt - authored.heightPt;
  return {
    xPt: effective.xPt - leftPt,
    yPt: effective.yPt - topPt,
    widthPt: Math.max(0, effective.widthPt + leftPt + rightPt),
    heightPt: Math.max(0, effective.heightPt + topPt + bottomPt),
  };
}

export function resizeResolvedAnchorGeometry(
  result: Extract<AnchorFrameResult, { status: 'resolved' }>,
  effectiveObjectFrame: LayoutRect,
): Extract<AnchorFrameResult, { status: 'resolved' }> {
  const authored = result.geometry.objectFrame;
  if (
    authored.xPt === effectiveObjectFrame.xPt
    && authored.yPt === effectiveObjectFrame.yPt
    && authored.widthPt === effectiveObjectFrame.widthPt
    && authored.heightPt === effectiveObjectFrame.heightPt
  ) return result;
  const scaleX = authored.widthPt === 0 ? 1 : effectiveObjectFrame.widthPt / authored.widthPt;
  const scaleY = authored.heightPt === 0 ? 1 : effectiveObjectFrame.heightPt / authored.heightPt;
  const polygon = result.geometry.wrap.polygon;
  return {
    ...result,
    geometry: {
      ...result.geometry,
      objectFrame: effectiveObjectFrame,
      inkBounds: resizeDerivedAnchorRect(result.geometry.inkBounds, authored, effectiveObjectFrame),
      wrapBounds: result.geometry.wrapBounds
        ? resizeDerivedAnchorRect(result.geometry.wrapBounds, authored, effectiveObjectFrame)
        : null,
      wrap: {
        ...result.geometry.wrap,
        polygon: polygon ? {
          ...polygon,
          points: polygon.points.map((point) => ({
            xPt: effectiveObjectFrame.xPt + (point.xPt - authored.xPt) * scaleX,
            yPt: effectiveObjectFrame.yPt + (point.yPt - authored.yPt) * scaleY,
          })),
        } : null,
      },
    },
  };
}
