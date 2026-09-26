// Classic chart geometry helpers.
import { drawingmlLineDashArray, pptxPresetDashArray } from '../../draw/dash.js';
import type { ChartModel } from '../../types/chart';


export function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}


/** Append `pts` to the CURRENT path starting from `pts[0]` (which the caller has
 *  already `moveTo`'d, or the first point is the current pen position). When
 *  `smooth` and there are ≥3 points, draw a Catmull-Rom → cubic-Bézier curve
 *  through the points (tangents from neighbours, the same formula scatter uses,
 *  ECMA-376 §21.2.2.194); otherwise straight `lineTo` segments. The caller owns
 *  `beginPath`/`moveTo`/`stroke`/`fill` so this composes into both the line
 *  stroke and the area fill's top edge. */
export function appendCurve(
  ctx: CanvasRenderingContext2D,
  pts: Array<{ x: number; y: number }>,
  smooth: boolean,
): void {
  if (pts.length === 0) return;
  if (smooth && pts.length >= 3) {
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] ?? pts[i];
      const p1 = pts[i];
      const p2 = pts[i + 1];
      const p3 = pts[i + 2] ?? p2;
      const cp1x = p1.x + (p2.x - p0.x) / 6;
      const cp1y = p1.y + (p2.y - p0.y) / 6;
      const cp2x = p2.x - (p3.x - p1.x) / 6;
      const cp2y = p2.y - (p3.y - p1.y) / 6;
      ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
    }
  } else {
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
  }
}


export function dashPatternForPreset(preset: string | undefined, lineWidth = 1): number[] {
  const scale = Number.isFinite(lineWidth) && lineWidth > 0 ? lineWidth : 1;
  return pptxPresetDashArray(preset ?? 'solid', scale);
}


export function dashPatternForLine(
  customDash: ChartModel['chartBorderCustomDash'],
  preset: string | null | undefined,
  lineWidth = 1,
): number[] {
  const scale = Number.isFinite(lineWidth) && lineWidth > 0 ? lineWidth : 1;
  return drawingmlLineDashArray(customDash, preset, scale);
}
