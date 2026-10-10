import type { Fill, PathCmd, Stroke } from '../types/common';
import { resolveArrowPaint, paintResolvedArrowHead } from './arrow';
import { trackPaintPath, currentStrokeBounds } from './paint-bounds';
import { appendGeometryPath } from './path-data';
import { applyResolvedStroke, hexToRgba, resolveFill, usesPathShade } from './paint';
import type { ResolvedStrokeGeometry } from './stroke-geometry';
import { resolveDrawingMLGeometry, requireResolvedDrawingMLGeometry, type ResolvedDrawingMLGeometry } from './drawingml-geometry';
import { pathFillModeOverlay } from './preset-geometry';

// Retained DOCX DrawingML plans use points. The shared pattern bitmap's one
// point cells therefore need no CSS-pixel conversion in this painter.
const PATTERN_PT_TO_SHAPE_UNITS = 1;

type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
  : T extends readonly (infer U)[] ? readonly DeepReadonly<U>[]
  : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;

export type DrawingMLShapeFill = DeepReadonly<Exclude<Fill, { fillType: 'image' }>>;

export type DrawingMLShapeGeometry =
  | Readonly<{
      kind: 'preset';
      name: string;
      adjustments: readonly (number | null)[];
    }>
  | Readonly<{
      kind: 'custom';
      subpaths: readonly (readonly PathCmd[])[];
      /** ECMA-376 §20.1.9.15 per-path `fill` mode and `stroke` flag,
       *  parallel to `subpaths`; absent when every path uses the defaults. */
      paint?: readonly DrawingMLPathPaint[];
    }>;

export type DrawingMLPathPaint = Readonly<{
  /** ST_PathFillMode (§20.1.10.37) other than `norm`. */
  fill?: 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  /** Present only when the path is not stroked. */
  stroke?: false;
}>;

export interface DrawingMLShapePaintPlan {
  /** DOCX acquisition always supplies this; ordinary core callers may resolve through the compatibility adapter. */
  readonly resolvedGeometry?: ResolvedDrawingMLGeometry;
  readonly rect: Readonly<{ x: number; y: number; w: number; h: number }>;
  readonly geometry: DrawingMLShapeGeometry;
  readonly fill: DrawingMLShapeFill | null;
  readonly stroke: DeepReadonly<Stroke> | null;
  readonly transform: Readonly<{
    rotationDeg: number;
    flipH: boolean;
    flipV: boolean;
  }>;
}

/** Execute paint in the authored DrawingML shape transform. Resource-backed
 * fills use this same frame as ordinary solid/gradient shape paint. */
export function withDrawingMLShapeTransform(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  paint: () => void,
): void {
  const { x, y, w, h } = plan.rect;
  const { rotationDeg, flipH, flipV } = plan.transform;
  ctx.save();
  try {
    if (rotationDeg !== 0 || flipH || flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      if (rotationDeg !== 0) ctx.rotate(rotationDeg * Math.PI / 180);
      ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    paint();
  } finally {
    ctx.restore();
  }
}

/** Append the same fill-bearing silhouette for clipping and path shading. */
function appendDrawingMLShapeOutline(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  x: number, y: number, w: number, h: number,
): void {
  const geometry = plan.resolvedGeometry
    ? requireResolvedDrawingMLGeometry(plan, plan.resolvedGeometry.scale)
    : resolveDrawingMLGeometry(plan, 1);
  if (w !== geometry.width || h !== geometry.height) throw new RangeError('Retained DrawingML silhouette frame mismatch');
  appendGeometryPath(ctx, geometry.fillSilhouette, x - geometry.originX, y - geometry.originY);
}

/** Clip the current state; the caller owns save/restore and the shape transform. */
export function clipDrawingMLShape(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
): void {
  const { x, y, w, h } = plan.rect;
  ctx.beginPath();
  appendDrawingMLShapeOutline(ctx, plan, x, y, w, h);
  // Use Canvas's nonzero rule, matching normal DrawingML shape fill and PPTX
  // picture clipping. `evenodd` would turn overlapping silhouette subpaths into
  // XOR holes that do not exist in the authored geometry.
  ctx.clip();
}

function applyDrawingMLStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke,
  geometry: ResolvedStrokeGeometry,
  rect: DrawingMLShapePaintPlan['rect'],
  rotationDeg: number,
): void {
  ctx.strokeStyle = hexToRgba(stroke.color);
  applyResolvedStroke(ctx, geometry);
  if (stroke.fill) {
    const paint = resolveFill(
      stroke.fill as Fill,
      ctx,
      rect.x,
      rect.y,
      rect.w,
      rect.h,
      rotationDeg,
      PATTERN_PT_TO_SHAPE_UNITS,
      undefined, undefined, currentStrokeBounds(ctx),
    );
    if (paint) ctx.strokeStyle = paint;
  }
}

/** Replay the shared immutable geometry using the existing brush frames. */
export function paintDrawingMLShape(
  ctx: CanvasRenderingContext2D, plan: DrawingMLShapePaintPlan, unitToDevice: number,
): void {
  const geometry = plan.resolvedGeometry
    ? requireResolvedDrawingMLGeometry(plan, unitToDevice)
    : resolveDrawingMLGeometry(plan, unitToDevice);
  const retainedPlan = plan.resolvedGeometry ? plan : { ...plan, resolvedGeometry: geometry };
  if (usesPathShade(plan.stroke?.fill)) ctx = trackPaintPath(ctx);
  const { x, y, w, h } = plan.rect;
  withDrawingMLShapeTransform(ctx, plan, () => {
    const fillStyle = resolveFill(plan.fill as Fill | null, ctx, x, y, w, h,
      plan.transform.rotationDeg, PATTERN_PT_TO_SHAPE_UNITS, undefined,
      (target, bx, by, bw, bh) => appendDrawingMLShapeOutline(target, retainedPlan, bx, by, bw, bh));
    const stroke = plan.stroke as Stroke | null;
    for (const path of geometry.paths) {
      ctx.beginPath();
      appendGeometryPath(ctx, path.path, x - geometry.originX, y - geometry.originY);
      if (fillStyle && path.fill !== 'none') {
        ctx.save();
        try {
          ctx.fillStyle = fillStyle;
          if (path.fillRule === 'evenodd') ctx.fill('evenodd'); else ctx.fill();
          const overlay = pathFillModeOverlay(path.fill);
          if (overlay) { ctx.fillStyle = overlay; if (path.fillRule === 'evenodd') ctx.fill('evenodd'); else ctx.fill(); }
        } finally { ctx.restore(); }
      }
      if (stroke && path.stroke) {
        applyDrawingMLStroke(ctx, stroke, path.stroke, plan.rect, plan.transform.rotationDeg);
        ctx.stroke();
      }
    }
    if (stroke) for (const arrow of geometry.arrows) {
      const tipX = x - geometry.originX + arrow.tipX, tipY = y - geometry.originY + arrow.tipY;
      paintResolvedArrowHead(ctx, tipX, tipY, arrow.angle, arrow.geometry, stroke,
        resolveArrowPaint(ctx, stroke, unitToDevice, tipX, tipY, arrow.angle, arrow.end,
          plan.rect, plan.transform.rotationDeg, PATTERN_PT_TO_SHAPE_UNITS));
    }
  });
}
