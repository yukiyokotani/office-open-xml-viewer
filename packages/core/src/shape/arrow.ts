import type { ArrowEnd, Stroke } from '../types/common';
import { hexToRgba, resolveFill, usesPathShade } from './paint';

import { lineEndLocalBounds, resolveArrowGeometry, trimGeometryPathEnds, type ResolvedArrowGeometry } from './arrow-geometry';
import { appendGeometryPath, createGeometryPath, type GeometryPathSink } from './path-data';
import { resolvePresetGeometryPaths } from './preset-geometry/geometry';
export { lineEndRetract, lineEndPaintExtent, retractLineEndpoint, trimGeometryPathEnds, type Point } from './arrow-geometry';

/** Trace the path `build` emits into `sink`, shortened `head` / `tail` px at
 *  its open ends (see {@link trimGeometryPathEnds}). */
export function traceTrimmedPath(
  sink: GeometryPathSink, build: (sink: GeometryPathSink) => void, head: number, tail: number,
): void {
  const path = createGeometryPath();
  build(path.sink);
  appendGeometryPath(sink, trimGeometryPathEnds(path.commands, head, tail));
}

/** Trace a connector/callout preset's leader (its trailing unfilled stroke
 *  path, the one `renderPresetShape` skips with `skipTrailingStroke`) trimmed
 *  at its ends. Returns false when the preset has no such leader. */
export function tracePresetLeader(
  sink: GeometryPathSink, geom: string, x: number, y: number, w: number, h: number,
  adj: (number | null | undefined)[], head: number, tail: number,
): boolean {
  const paths = resolvePresetGeometryPaths(geom, w, h, adj, undefined, x, y);
  const leader = paths?.[paths.length - 1];
  if (!leader || !leader.stroke || !(leader.fill === 'none' || leader.fill == null)) return false;
  appendGeometryPath(sink, trimGeometryPathEnds(leader.path, head, tail));
  return true;
}

/**
 * Draw a DrawingML line-end decoration (arrow head) at `(tipX, tipY)`,
 * oriented along `angle` radians (0 = pointing right, +x axis).
 *
 * Sizes and outlines follow the observed Office rule documented in
 * `arrow-geometry.ts` (ECMA-376 §20.1.8.38 / §20.1.8.57, ST_LineEndType /
 * ST_LineEndWidth / ST_LineEndLength), shared by the PPTX and DOCX renderers.
 *
 * `scale` converts stroke-width units to device px (stroke width in px is
 * `stroke.width * scale`, as in core's `applyStroke`); `unitsPerPoint` is the
 * number of stroke-width units per point (12700 for EMU strokes).
 */
export function drawArrowHead(
  ctx: CanvasRenderingContext2D,
  tipX: number,
  tipY: number,
  angle: number,
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
  unitsPerPoint: number,
  effectivePaint?: string | CanvasGradient | CanvasPattern,
  paintInHostFrame = false,
): void {
  const geometry = resolveArrowGeometry(arrowEnd, stroke, scale, unitsPerPoint);
  if (geometry) paintResolvedArrowHead(ctx, tipX, tipY, angle, geometry, stroke, effectivePaint, paintInHostFrame);
}

/** Execute a resolved arrow; no path or endpoint geometry is acquired here. */
export function paintResolvedArrowHead(
  ctx: CanvasRenderingContext2D, tipX: number, tipY: number, angle: number,
  geometry: ResolvedArrowGeometry, stroke: Stroke,
  effectivePaint?: string | CanvasGradient | CanvasPattern,
  /** Paint every gradient / pattern brush in the host (line) frame, so the
   *  decoration continues the brush of the line it ends. Hosts without prior
   *  decoration behaviour opt in; the default keeps the frames described below. */
  paintInHostFrame = false,
): void {

  const paint = effectivePaint ?? hexToRgba(stroke.color);

  // Only the new untiled rect/shape raster is authored in the host frame by
  // default. Ordinary patterns and tiled gradients retain main's
  // decoration-local CTM; native solid/linear/circle paints also keep their
  // existing frames. A CanvasPattern alone cannot distinguish those brushes
  // from a path raster. `paintInHostFrame` paints every non-solid brush at
  // the host CTM instead.
  const hostTransform = (usesPathShade(stroke.fill) || (paintInHostFrame && typeof paint === 'object'))
    && typeof ctx.getTransform === 'function' ? ctx.getTransform() : undefined;
  const anchorPaint = () => {
    // Canvas patterns follow the paint-time CTM; native gradients retain their
    // creation frame. The path has already captured the decoration transform.
    // Paint a host-box pattern at the host CTM so rotating/translating the
    // arrow geometry does not translate/rotate its gradient a second time.
    if (hostTransform && typeof paint === 'object' && (paintInHostFrame || 'setTransform' in paint)) ctx.setTransform(hostTransform);
  };
  ctx.save();
  ctx.translate(tipX, tipY);
  ctx.rotate(angle);
  ctx.fillStyle = paint;
  ctx.strokeStyle = paint;
  ctx.lineWidth = geometry.stroke.lineWidth;
  ctx.setLineDash([]);
  ctx.beginPath();
  appendGeometryPath(ctx, geometry.path);
  if (geometry.paint === 'stroke') {
    ctx.lineCap = geometry.stroke.lineCap;
    ctx.lineJoin = geometry.stroke.lineJoin;
    ctx.miterLimit = geometry.stroke.miterLimit;
  }
  anchorPaint();
  if (geometry.paint === 'fill') ctx.fill(); else ctx.stroke();
  ctx.restore();
}

/** Resolve decoration paint from the very geometry drawn by drawArrowHead.
 * Coverage is the decoration's local box (`lineEndLocalBounds`, which already
 * includes the open arrow's pen); the oval uses its exact ellipse support.
 * Keep the gradient's host frame separate from this device-space coverage. */
export function resolveArrowPaint(
  ctx: CanvasRenderingContext2D, stroke: Stroke, scale: number, unitsPerPoint: number,
  tipX: number, tipY: number, angle: number, end: ArrowEnd,
  frame: { x: number; y: number; w: number; h: number }, rotation: number, ptUnits: number,
  patternCoordinateTransform?: DOMMatrix2DInit,
): string | CanvasGradient | CanvasPattern | undefined {
  if (!stroke.fill || end.type === 'none') return undefined;
  if (!usesPathShade(stroke.fill)
    || typeof ctx.getTransform !== 'function') {
    return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits, patternCoordinateTransform) ?? undefined;
  }
  const box = lineEndLocalBounds(end, stroke, scale, unitsPerPoint);
  if (!box) return undefined;
  const m = ctx.getTransform(); const c = Math.cos(angle); const s = Math.sin(angle);
  const a = m.a * c + m.c * s; const b = m.b * c + m.d * s;
  const cc = -m.a * s + m.c * c; const d = -m.b * s + m.d * c;
  const mid = (box.x0 + box.x1) / 2, half = (box.x1 - box.x0) / 2;
  const cx = m.a * tipX + m.c * tipY + m.e + a * mid;
  const cy = m.b * tipX + m.d * tipY + m.f + b * mid;
  const oval = end.type === 'oval';
  const hx = oval ? Math.hypot(a * half, cc * box.halfY) : Math.abs(a) * half + Math.abs(cc) * box.halfY;
  const hy = oval ? Math.hypot(b * half, d * box.halfY) : Math.abs(b) * half + Math.abs(d) * box.halfY;
  return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits,
    undefined, undefined, { x: cx - hx, y: cy - hy, w: 2 * hx, h: 2 * hy }) ?? undefined;
}
