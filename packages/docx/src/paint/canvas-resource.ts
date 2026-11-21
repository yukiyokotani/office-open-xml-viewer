import type { LayoutRect, PaintResourceKind, UprightResourceOrientation } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';
import { applyCanvasTransform } from './canvas-transform.js';

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
): void {
  if (orientation === undefined) {
    context.resources.paint(resourceKey, resourceKind, bounds, context.ctx);
    return;
  }
  const { ctx } = context;
  ctx.save();
  ctx.translate(
    bounds.xPt + bounds.widthPt / 2,
    bounds.yPt + bounds.heightPt / 2,
  );
  if (orientation === 'upright-physical-transposed') {
    applyCanvasTransform(ctx, { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 });
  } else {
    ctx.rotate(orientation === 'upright-physical' ? -Math.PI / 2 : Math.PI / 2);
  }
  context.resources.paint(resourceKey, resourceKind, {
    xPt: -bounds.heightPt / 2,
    yPt: -bounds.widthPt / 2,
    widthPt: bounds.heightPt,
    heightPt: bounds.widthPt,
  }, ctx);
  ctx.restore();
}
