import type { LayoutRect, NativeReadingImagePlan, PaintResourceKind, UprightResourceOrientation } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

/** Paint one retained non-text resource using the orientation selected during
 * layout acquisition. A vertical section rotates its section-logical frame a
 * quarter turn (clockwise, or counter-clockwise for a native BtoT section);
 * physical graphics apply the retained inverse turn locally so their authored
 * DrawingML transform is subsequently composed in an upright frame. */
export function paintRetainedResource(
  resourceKey: string,
  resourceKind: PaintResourceKind,
  bounds: LayoutRect,
  orientation: UprightResourceOrientation | undefined,
  context: CanvasPaintContext,
  nativeImagePlan?: NativeReadingImagePlan,
): void {
  // Counter-turning the command itself exchanges its destination axes. A
  // retained reading plan must be acquired in that frame, not silently rebuilt
  // here. Current reading producers use the drawing-local transform instead.
  if (nativeImagePlan && orientation !== undefined) throw new Error('Reading image command orientation is not acquired');
  if (orientation === undefined) {
    context.resources.paint(resourceKey, resourceKind, bounds, context.ctx, nativeImagePlan);
    return;
  }
  const { ctx } = context;
  ctx.save();
  ctx.translate(
    bounds.xPt + bounds.widthPt / 2,
    bounds.yPt + bounds.heightPt / 2,
  );
  ctx.rotate(orientation === 'upright-physical' ? -Math.PI / 2 : Math.PI / 2);
  context.resources.paint(resourceKey, resourceKind, {
    xPt: -bounds.heightPt / 2,
    yPt: -bounds.widthPt / 2,
    widthPt: bounds.heightPt,
    heightPt: bounds.widthPt,
  }, ctx);
  ctx.restore();
}
