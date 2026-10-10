import type { ArrowEnd, Stroke } from '../types/common';
import { hexToRgba, resolveFill, usesPathShade } from './paint';

import { arrowGeom, resolveArrowGeometry, type ResolvedArrowGeometry } from './arrow-geometry';
import { appendGeometryPath } from './path-data';
export { lineEndRetract, lineEndPaintExtent, retractLineEndpoint, type Point } from './arrow-geometry';

/**
 * Draw a DrawingML line-end decoration (arrow head) at `(tipX, tipY)`,
 * oriented along `angle` radians (0 = pointing right, +x axis).
 *
 * ECMA-376 §20.1.8.3 (CT_LineEndProperties) / §20.1.10.33 (ST_LineEndType:
 * none / triangle / stealth / diamond / oval / arrow) / §20.1.10.31–.32
 * (ST_LineEndWidth / ST_LineEndLength: sm / med / lg). The spec only names
 * the w/len steps as *relative* sizes, not exact ratios — the multiples of
 * line width below are calibrated against PowerPoint's rendering and shared
 * between the pptx and docx renderers so connector arrows look identical.
 *
 * `scale` is the EMU → device-px factor (same convention as core's
 * `applyStroke`, where stroke width in px is `stroke.width * scale`).
 */
export function drawArrowHead(
  ctx: CanvasRenderingContext2D,
  tipX: number,
  tipY: number,
  angle: number,
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
  effectivePaint?: string | CanvasGradient | CanvasPattern,
): void {
  const geometry = resolveArrowGeometry(arrowEnd, stroke, scale);
  if (geometry) paintResolvedArrowHead(ctx, tipX, tipY, angle, geometry, stroke, effectivePaint);
}

/** Execute a resolved arrow; no path or endpoint geometry is acquired here. */
export function paintResolvedArrowHead(
  ctx: CanvasRenderingContext2D, tipX: number, tipY: number, angle: number,
  geometry: ResolvedArrowGeometry, stroke: Stroke,
  effectivePaint?: string | CanvasGradient | CanvasPattern,
): void {

  const paint = effectivePaint ?? hexToRgba(stroke.color);

  // Only the new untiled rect/shape raster is authored in the host frame.
  // Ordinary patterns and tiled gradients retain main's decoration-local CTM;
  // native solid/linear/circle paints also keep their existing frames. A
  // CanvasPattern alone cannot distinguish those brushes from a path raster.
  const hostTransform = usesPathShade(stroke.fill)
    && typeof ctx.getTransform === 'function' ? ctx.getTransform() : undefined;
  const anchorPaint = () => {
    // Canvas patterns follow the paint-time CTM; native gradients retain their
    // creation frame. The path has already captured the decoration transform.
    // Paint a host-box pattern at the host CTM so rotating/translating the
    // arrow geometry does not translate/rotate its gradient a second time.
    if (hostTransform && typeof paint === 'object' && 'setTransform' in paint) ctx.setTransform(hostTransform);
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
  }
  anchorPaint();
  if (geometry.paint === 'fill') ctx.fill(); else ctx.stroke();
  ctx.restore();
}

/** Resolve decoration paint from the very geometry drawn by drawArrowHead.
 * The filled end is contained by [-len,0] × [-halfW,halfW]; oval uses its
 * exact ellipse support. Open arrows add the transformed round stroke pen.
 * Keep the gradient's host frame separate from this device-space coverage. */
export function resolveArrowPaint(
  ctx: CanvasRenderingContext2D, stroke: Stroke, scale: number,
  tipX: number, tipY: number, angle: number, end: ArrowEnd,
  frame: { x: number; y: number; w: number; h: number }, rotation: number, ptUnits: number,
): string | CanvasGradient | CanvasPattern | undefined {
  if (!stroke.fill || end.type === 'none') return undefined;
  if (!usesPathShade(stroke.fill)
    || typeof ctx.getTransform !== 'function') {
    return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits) ?? undefined;
  }
  const { lw, halfW, len } = arrowGeom(end, stroke, scale);
  const m = ctx.getTransform(); const c = Math.cos(angle); const s = Math.sin(angle);
  const a = m.a * c + m.c * s; const b = m.b * c + m.d * s;
  const cc = -m.a * s + m.c * c; const d = -m.b * s + m.d * c;
  const cx = m.a * tipX + m.c * tipY + m.e - a * len / 2;
  const cy = m.b * tipX + m.d * tipY + m.f - b * len / 2;
  const hx = end.type === 'oval' ? Math.hypot(a * len / 2, cc * halfW) : Math.abs(a) * len / 2 + Math.abs(cc) * halfW;
  const hy = end.type === 'oval' ? Math.hypot(b * len / 2, d * halfW) : Math.abs(b) * len / 2 + Math.abs(d) * halfW;
  const px = end.type === 'arrow' ? lw / 2 * Math.hypot(a, cc) : 0;
  const py = end.type === 'arrow' ? lw / 2 * Math.hypot(b, d) : 0;
  return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits,
    undefined, undefined, { x: cx - hx - px, y: cy - hy - py, w: 2 * (hx + px), h: 2 * (hy + py) }) ?? undefined;
}
