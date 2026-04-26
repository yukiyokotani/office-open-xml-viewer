/**
 * Executes a single ECMA-376 `<path>` element against a Canvas 2D context.
 *
 * All `<arcTo>` angles are interpreted per spec as *visual* angles (not
 * parametric), matching the `ooxmlArcTo` helper in renderer.ts. The ellipse
 * center is back-solved from the pen position so multiple arcs chain
 * continuously.
 */

import type { Evaluator } from './evaluator';

// 60 000-ths of a degree per full revolution (2π radians).
const DEG60K_TO_RAD = (Math.PI * 2) / 21600000;

export interface PresetPath {
  /** Path-local coord-space width. `null` → use shape width as-is. */
  w: number | null;
  h: number | null;
  /** Fill modifier: null | "norm" | "none" | "lighten" | "lightenLess" | "darken" | "darkenLess". */
  fill: string | null;
  stroke: boolean;
  extrusionOk: boolean;
  cmds: Array<[string, ...string[]]>;
}

/**
 * Apply the path's commands to the current ctx sub-path. Caller owns
 * beginPath / fill / stroke. The pen position is tracked in absolute canvas
 * pixels because `<arcTo>` back-computes the ellipse centre from it.
 */
export function applyPresetPath(
  ctx: CanvasRenderingContext2D,
  path: PresetPath,
  evaluator: Evaluator,
  shapeX: number,
  shapeY: number,
  shapeW: number,
  shapeH: number,
): void {
  // Path-local coords map to [0, shapeW]×[0, shapeH] on canvas.
  const sx = path.w != null ? shapeW / path.w : 1;
  const sy = path.h != null ? shapeH / path.h : 1;
  const toAbsX = (v: number) => shapeX + v * sx;
  const toAbsY = (v: number) => shapeY + v * sy;

  let penX = 0;  // absolute canvas px
  let penY = 0;

  for (const cmd of path.cmds) {
    switch (cmd[0]) {
      case 'm': {
        const ax = toAbsX(evaluator.resolve(cmd[1]));
        const ay = toAbsY(evaluator.resolve(cmd[2]));
        ctx.moveTo(ax, ay);
        penX = ax; penY = ay;
        break;
      }
      case 'l': {
        const ax = toAbsX(evaluator.resolve(cmd[1]));
        const ay = toAbsY(evaluator.resolve(cmd[2]));
        ctx.lineTo(ax, ay);
        penX = ax; penY = ay;
        break;
      }
      case 'C': {
        const ax1 = toAbsX(evaluator.resolve(cmd[1]));
        const ay1 = toAbsY(evaluator.resolve(cmd[2]));
        const ax2 = toAbsX(evaluator.resolve(cmd[3]));
        const ay2 = toAbsY(evaluator.resolve(cmd[4]));
        const ax  = toAbsX(evaluator.resolve(cmd[5]));
        const ay  = toAbsY(evaluator.resolve(cmd[6]));
        ctx.bezierCurveTo(ax1, ay1, ax2, ay2, ax, ay);
        penX = ax; penY = ay;
        break;
      }
      case 'Q': {
        const ax1 = toAbsX(evaluator.resolve(cmd[1]));
        const ay1 = toAbsY(evaluator.resolve(cmd[2]));
        const ax  = toAbsX(evaluator.resolve(cmd[3]));
        const ay  = toAbsY(evaluator.resolve(cmd[4]));
        ctx.quadraticCurveTo(ax1, ay1, ax, ay);
        penX = ax; penY = ay;
        break;
      }
      case 'a': {
        // Path-local radii: visual-to-parametric conversion is computed in
        // path-local space (where the OOXML visual angle is defined). The
        // canvas-scaled radii are used only for the actual pixel placement.
        // When sx ≠ sy (e.g. cloudCallout in a non-square box) using the
        // scaled ratio here would skew the start/sweep angles.
        const wR_local = evaluator.resolve(cmd[1]);
        const hR_local = evaluator.resolve(cmd[2]);
        const wR = wR_local * sx;
        const hR = hR_local * sy;
        const stDeg = evaluator.resolve(cmd[3]) * DEG60K_TO_RAD;
        const swDeg = evaluator.resolve(cmd[4]) * DEG60K_TO_RAD;
        // OOXML arc uses visual angles; Canvas's ellipse() takes parametric.
        // atan2 collapses every full revolution, so split the sweep into
        // whole turns + remainder and compute a monotonic parametric sweep
        // whose sign matches swDeg. Without this, sun/etc. with swDeg=360°
        // produce start==end and nothing is drawn.
        const TWO_PI = Math.PI * 2;
        const visualToParam = (v: number) =>
          Math.atan2(wR_local * Math.sin(v), hR_local * Math.cos(v));
        const stP = visualToParam(stDeg);
        const fullRevs = Math.trunc(swDeg / TWO_PI);
        const remainder = swDeg - fullRevs * TWO_PI;
        let delta = visualToParam(stDeg + remainder) - stP;
        if (remainder > 0 && delta < 0) delta += TWO_PI;
        else if (remainder < 0 && delta > 0) delta -= TWO_PI;
        const endP = stP + delta + fullRevs * TWO_PI;
        // Back-solve center from pen & start angle (canvas-scaled radii).
        const cx = penX - wR * Math.cos(stP);
        const cy = penY - hR * Math.sin(stP);
        if (Math.abs(wR) > 1e-6 && Math.abs(hR) > 1e-6) {
          ctx.ellipse(
            cx, cy, Math.abs(wR), Math.abs(hR),
            0, stP, endP, swDeg < 0,
          );
          penX = cx + wR * Math.cos(endP);
          penY = cy + hR * Math.sin(endP);
        }
        break;
      }
      case 'c':
        ctx.closePath();
        break;
    }
  }
}
