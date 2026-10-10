import type { ArrowEnd, Stroke } from '../types/common';
import { createGeometryPath, type GeometryPath } from './path-data';
import type { ResolvedStrokeGeometry } from './stroke-geometry';

/** A 2D point in canvas pixels. */
export interface Point {
  x: number;
  y: number;
}

/**
 * Resolve a line-end decoration's pixel geometry. The width/length steps
 * (sm/med/lg) are *relative* in the spec (§20.1.10.31–.32); the multipliers of
 * line width below are calibrated against PowerPoint. `lw` is the line width in
 * device px, `halfW` the half-span across the line, `len` the extent along it.
 */
export function arrowGeom(
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
): { lw: number; halfW: number; len: number } {
  const lw = Math.max(0.5, stroke.width * scale);
  const wMul = arrowEnd.w === 'sm' ? 4 : arrowEnd.w === 'lg' ? 8 : 6;
  const lMul = arrowEnd.len === 'sm' ? 4 : arrowEnd.len === 'lg' ? 8 : 6;
  return { lw, halfW: (lw * wMul) / 2, len: lw * lMul };
}

/** Decorations whose filled body covers the tip→`-len` span, so the leader line
 *  must stop at `-len` (retract) for its cap to hide inside the shape. The open
 *  `arrow` (a stroked V) and `none` keep the line running all the way to the tip. */
const RETRACTING_ENDS = new Set(['triangle', 'stealth', 'diamond', 'oval']);

/**
 * How far (device px) the leader line should be pulled back from the tip so a
 * line-end decoration's filled body hides the line's end cap. Zero for `arrow`
 * and `none`. Matches the `len` used by {@link drawArrowHead} so the line stops
 * exactly at the decoration's base.
 */
export function lineEndRetract(arrowEnd: ArrowEnd, stroke: Stroke, scale: number): number {
  if (!RETRACTING_ENDS.has(arrowEnd.type)) return 0;
  return arrowGeom(arrowEnd, stroke, scale).len;
}

/** Conservative radius around a line-end tip occupied by its painted pixels. */
export function lineEndPaintExtent(arrowEnd: ArrowEnd, stroke: Stroke, scale: number): number {
  if (arrowEnd.type === 'none') return 0;
  const { lw, halfW, len } = arrowGeom(arrowEnd, stroke, scale);
  return Math.max(len, halfW) + lw / 2;
}

/**
 * Pull `p` toward its neighbour `toward` by `amount` px, clamped so it never
 * passes the neighbour. Used to retract a polyline's terminal vertex before
 * stroking, so a decorated end stops at the decoration's base.
 */
export function retractLineEndpoint(p: Point, toward: Point, amount: number): Point {
  if (amount <= 0) return { x: p.x, y: p.y };
  const dx = toward.x - p.x;
  const dy = toward.y - p.y;
  const d = Math.hypot(dx, dy);
  if (d < 1e-9) return { x: p.x, y: p.y };
  const t = Math.min(amount, d) / d;
  return { x: p.x + dx * t, y: p.y + dy * t };
}


export interface ResolvedArrowGeometry {
  readonly path: GeometryPath;
  readonly paint: 'fill' | 'stroke';
  readonly stroke: ResolvedStrokeGeometry;
}
export function resolveArrowGeometry(end: ArrowEnd, stroke: Stroke, scale: number,
  charge: (units: number) => void = () => {}): ResolvedArrowGeometry | null {
  if (end.type === 'none') return null;
  const { lw, halfW, len } = arrowGeom(end, stroke, scale);
  const path = createGeometryPath(charge), ctx = path.sink;
  switch (end.type) {
    case 'triangle': case 'stealth':
      ctx.moveTo(0, 0); ctx.lineTo(-len, -halfW); ctx.lineTo(-len, halfW); ctx.closePath(); break;
    case 'arrow':
      ctx.moveTo(-len, -halfW); ctx.lineTo(0, 0); ctx.lineTo(-len, halfW); break;
    case 'diamond':
      ctx.moveTo(0, 0); ctx.lineTo(-len / 2, -halfW); ctx.lineTo(-len, 0); ctx.lineTo(-len / 2, halfW); ctx.closePath(); break;
    case 'oval': ctx.ellipse(-len / 2, 0, len / 2, halfW, 0, 0, Math.PI * 2); break;
  }
  return Object.freeze({ path: Object.freeze(path.commands), paint: end.type === 'arrow' ? 'stroke' : 'fill',
    stroke: Object.freeze({ lineWidth: lw, lineCap: end.type === 'arrow' ? 'round' : 'butt',
      lineJoin: end.type === 'arrow' ? 'round' : 'miter', miterLimit: 10, dash: Object.freeze([]) }) });
}
