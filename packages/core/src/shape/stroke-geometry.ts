import type { Stroke } from '../types/common';
import { drawingmlLineDashArray, shapeStrokeDashArray } from '../draw/dash';
export interface ResolvedStrokeGeometry {
  readonly lineWidth: number;
  readonly lineCap: 'butt' | 'round' | 'square';
  readonly lineJoin: 'round' | 'bevel' | 'miter';
  readonly miterLimit: number;
  readonly dash: readonly number[];
}
/** One shared owner for stroke-width/dash/cap normalization. */
export function resolveStrokeGeometry(stroke: Stroke | null, scale: number): ResolvedStrokeGeometry {
  if (!Number.isFinite(scale) || scale < 0) throw new RangeError('Invalid DrawingML stroke scale');
  if (!stroke) return Object.freeze({ lineWidth: 0, lineCap: 'butt', lineJoin: 'miter', miterLimit: 10, dash: Object.freeze([]) });
  if (!Number.isFinite(stroke.width) || stroke.width < 0) throw new RangeError('Invalid DrawingML stroke width');
  const lw = Math.max(0.5, stroke.width * scale);
  const dash = stroke.customDash != null
    ? drawingmlLineDashArray(stroke.customDash, null, lw)
    : stroke.dashStyle ? shapeStrokeDashArray(stroke.dashStyle, lw) : [];
  const requestedCap = stroke.lineCap ?? 'butt';
  const hasZeroDash = dash.some((length, index) => index % 2 === 0 && length === 0);
  const value = { lineWidth: lw, lineCap: requestedCap === 'butt' && hasZeroDash ? 'square' as const : requestedCap,
    lineJoin: stroke.lineJoin ?? 'miter', miterLimit: stroke.miterLimit ?? 10, dash: Object.freeze(dash) };
  if (!Number.isFinite(value.lineWidth) || !Number.isFinite(value.miterLimit) || value.miterLimit <= 0
    || !value.dash.every((v) => Number.isFinite(v) && v >= 0)) throw new RangeError('Invalid resolved DrawingML stroke');
  return Object.freeze(value);
}
