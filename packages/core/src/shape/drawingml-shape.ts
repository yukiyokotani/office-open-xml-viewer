import type { Fill, PathCmd, Stroke } from '../types/common';
import { drawArrowHead, lineEndRetract, retractLineEndpoint } from './arrow';
import { buildCustomPath } from './custGeom';
import { getCustGeomEndpoints } from './custgeom-endpoints';
import { applyStroke, resolveFill } from './paint';
import { buildShapePath } from './preset';
import {
  buildPresetGeometryFillPath,
  getConnectorAnchors,
  hasPreset,
  pathFillModeOverlay,
  renderPresetShape,
} from './preset-geometry';

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

/** The per-path paint of a custom geometry, or null when it has none or
 *  it does not line up with the subpaths. */
function customPathPaint(
  geometry: Extract<DrawingMLShapeGeometry, { kind: 'custom' }>,
): readonly DrawingMLPathPaint[] | null {
  return geometry.paint && geometry.paint.length === geometry.subpaths.length
    ? geometry.paint
    : null;
}

export interface DrawingMLShapePaintPlan {
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

/** Clip the current canvas state to the exact DrawingML shape silhouette.
 * The caller owns save/restore and should enter the shape transform first. */
export function clipDrawingMLShape(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
): void {
  const { x, y, w, h } = plan.rect;
  ctx.beginPath();
  if (plan.geometry.kind === 'preset') {
    const adjustments = [...plan.geometry.adjustments];
    if (!buildPresetGeometryFillPath(
      ctx, plan.geometry.name, x, y, w, h, adjustments,
    )) {
      buildShapePath(
        ctx, plan.geometry.name, x, y, w, h,
        adjustments[0], adjustments[1], adjustments[2], adjustments[3],
      );
    }
  } else {
    // The silhouette is the fill-bearing paths (an unfilled path is only a
    // line), as for PPTX custom geometry.
    const paint = customPathPaint(plan.geometry);
    const subpaths = paint
      ? plan.geometry.subpaths.filter((_, index) => paint[index].fill !== 'none')
      : plan.geometry.subpaths;
    buildCustomPath(ctx, subpaths as PathCmd[][], x, y, w, h);
  }
  // Use Canvas's nonzero rule, matching normal DrawingML shape fill and PPTX
  // picture clipping. `evenodd` would turn overlapping silhouette subpaths into
  // XOR holes that do not exist in the authored geometry.
  ctx.clip();
}

const CONNECTOR_GEOMETRIES = new Set([
  'line', 'straightconnector1',
  'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
  'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
]);

const CALLOUT_GEOMETRIES = new Set([
  'callout1', 'callout2', 'callout3',
  'bordercallout1', 'bordercallout2', 'bordercallout3',
  'accentcallout1', 'accentcallout2', 'accentcallout3',
  'accentbordercallout1', 'accentbordercallout2', 'accentbordercallout3',
]);

function retractableLeader(geometry: string): boolean {
  return CALLOUT_GEOMETRIES.has(geometry)
    || geometry === 'line'
    || geometry === 'straightconnector1'
    || geometry.startsWith('bentconnector');
}

function applyDrawingMLStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke,
  unitToDevice: number,
  rect: DrawingMLShapePaintPlan['rect'],
  rotationDeg: number,
): void {
  applyStroke(ctx, stroke, unitToDevice);
  if (stroke.fill) {
    const paint = resolveFill(
      stroke.fill as Fill,
      ctx,
      rect.x,
      rect.y,
      rect.w,
      rect.h,
      rotationDeg,
    );
    if (paint) ctx.strokeStyle = paint;
  }
}

function paintConnectorEnds(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  geometry: string,
  unitToDevice: number,
): void {
  const stroke = plan.stroke as Stroke | null;
  if (!stroke || (!CONNECTOR_GEOMETRIES.has(geometry) && !CALLOUT_GEOMETRIES.has(geometry))) {
    return;
  }
  const { x, y, w, h } = plan.rect;
  const effectivePaint = stroke.fill
    ? resolveFill(
        stroke.fill as Fill,
        ctx,
        x,
        y,
        w,
        h,
        plan.transform.rotationDeg,
      ) ?? undefined
    : undefined;
  const adjustments = plan.geometry.kind === 'preset' ? plan.geometry.adjustments : [];
  const anchors = getConnectorAnchors(geometry, x, y, w, h, [...adjustments]);
  if (!anchors) return;
  if (retractableLeader(geometry)
    && anchors.vertices.length >= 2
    && (stroke.headEnd || stroke.tailEnd)) {
    const points = anchors.vertices.map((vertex) => ({ x: vertex.x, y: vertex.y }));
    if (stroke.tailEnd) {
      points[points.length - 1] = retractLineEndpoint(
        points[points.length - 1],
        points[points.length - 2],
        lineEndRetract(stroke.tailEnd, stroke, unitToDevice),
      );
    }
    if (stroke.headEnd) {
      points[0] = retractLineEndpoint(
        points[0],
        points[1],
        lineEndRetract(stroke.headEnd, stroke, unitToDevice),
      );
    }
    applyDrawingMLStroke(ctx, stroke, unitToDevice, plan.rect, plan.transform.rotationDeg);
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    for (let index = 1; index < points.length; index++) {
      ctx.lineTo(points[index].x, points[index].y);
    }
    ctx.stroke();
  }
  if (stroke.tailEnd) {
    drawArrowHead(
      ctx, anchors.end.x, anchors.end.y, anchors.end.angle,
      stroke.tailEnd, stroke, unitToDevice,
      effectivePaint,
    );
  }
  if (stroke.headEnd) {
    drawArrowHead(
      ctx, anchors.start.x, anchors.start.y, anchors.start.angle,
      stroke.headEnd, stroke, unitToDevice,
      effectivePaint,
    );
  }
}

function paintCustomEnds(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  unitToDevice: number,
): void {
  if (plan.geometry.kind !== 'custom') return;
  const stroke = plan.stroke as Stroke | null;
  if (!stroke || (!stroke.headEnd && !stroke.tailEnd)) return;
  const endpoints = getCustGeomEndpoints(plan.geometry.subpaths as PathCmd[][]);
  const { x, y, w, h } = plan.rect;
  const effectivePaint = stroke.fill
    ? resolveFill(
        stroke.fill as Fill,
        ctx,
        x,
        y,
        w,
        h,
        plan.transform.rotationDeg,
      ) ?? undefined
    : undefined;
  if (endpoints.start && stroke.headEnd) {
    drawArrowHead(
      ctx,
      x + endpoints.start.x * w,
      y + endpoints.start.y * h,
      Math.atan2(endpoints.start.dy * h, endpoints.start.dx * w),
      stroke.headEnd,
      stroke,
      unitToDevice,
      effectivePaint,
    );
  }
  if (endpoints.end && stroke.tailEnd) {
    drawArrowHead(
      ctx,
      x + endpoints.end.x * w,
      y + endpoints.end.y * h,
      Math.atan2(endpoints.end.dy * h, endpoints.end.dx * w),
      stroke.tailEnd,
      stroke,
      unitToDevice,
      effectivePaint,
    );
  }
}

export function paintDrawingMLShape(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  unitToDevice: number,
): void {
  const { x, y, w, h } = plan.rect;
  withDrawingMLShapeTransform(ctx, plan, () => {
    // Shared fill resolution is observational; retained plans keep gradient
    // stops readonly so layout snapshots cannot be mutated by a painter.
    const fillStyle = resolveFill(
      plan.fill as Fill | null,
      ctx,
      x,
      y,
      w,
      h,
      plan.transform.rotationDeg,
    );
    const stroke = plan.stroke as Stroke | null;
    const applyAndStroke = stroke
      ? () => {
          applyDrawingMLStroke(
            ctx,
            stroke,
            unitToDevice,
            plan.rect,
            plan.transform.rotationDeg,
          );
          ctx.stroke();
        }
      : null;
    if (plan.geometry.kind === 'preset') {
      const geometry = plan.geometry.name.toLowerCase();
      const adjustments = [...plan.geometry.adjustments];
      const hasDecoratedRetractableLeader = retractableLeader(geometry)
        && !!(stroke?.headEnd || stroke?.tailEnd);
      const painted = hasPreset(geometry) && renderPresetShape(
        ctx,
        geometry,
        x,
        y,
        w,
        h,
        adjustments,
        fillStyle,
        applyAndStroke,
        () => {},
        hasDecoratedRetractableLeader ? { skipTrailingStroke: true } : undefined,
      );
      if (!painted) {
        ctx.beginPath();
        buildShapePath(
          ctx, geometry, x, y, w, h,
          adjustments[0], adjustments[1], adjustments[2], adjustments[3],
        );
        if (fillStyle && geometry !== 'arc') {
          ctx.fillStyle = fillStyle;
          if (geometry === 'donut' || geometry === 'smileyface' || geometry === 'frame') {
            ctx.fill('evenodd');
          } else {
            ctx.fill();
          }
        }
        if (applyAndStroke) applyAndStroke();
      }
      paintConnectorEnds(ctx, plan, geometry, unitToDevice);
    } else {
      const paint = customPathPaint(plan.geometry);
      if (paint) {
        // ECMA-376 §20.1.9.15: each path is filled and stroked on its own.
        plan.geometry.subpaths.forEach((subpath, index) => {
          ctx.beginPath();
          buildCustomPath(ctx, [subpath as PathCmd[]], x, y, w, h);
          const mode = paint[index].fill;
          if (fillStyle && mode !== 'none') {
            ctx.fillStyle = fillStyle;
            ctx.fill();
            const overlay = pathFillModeOverlay(mode);
            if (overlay) {
              ctx.save();
              ctx.fillStyle = overlay;
              ctx.fill();
              ctx.restore();
            }
          }
          if (applyAndStroke && paint[index].stroke !== false) applyAndStroke();
        });
      } else {
        ctx.beginPath();
        buildCustomPath(ctx, plan.geometry.subpaths as PathCmd[][], x, y, w, h);
        if (fillStyle) {
          ctx.fillStyle = fillStyle;
          ctx.fill();
        }
        if (applyAndStroke) applyAndStroke();
      }
      paintCustomEnds(ctx, plan, unitToDevice);
    }
  });
}
