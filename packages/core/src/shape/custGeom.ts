import type { PathCmd } from '../types/common';

/**
 * Build a canvas path from normalised custGeom sub-paths.
 * Coordinates in each PathCmd are in [0,1] relative to the shape bounding box;
 * this function maps them to the canvas pixel rectangle (x, y, w, h).
 *
 * Pen position is tracked so `arcTo` can back-calculate the ellipse centre
 * from the current pen point and `stAng`.
 */
export function buildCustomPath(
  ctx: CanvasRenderingContext2D,
  subpaths: PathCmd[][],
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  for (const cmds of subpaths) {
    let penX = 0;
    let penY = 0;
    for (const cmd of cmds) {
      switch (cmd.cmd) {
        case 'moveTo':
          ctx.moveTo(x + cmd.x * w, y + cmd.y * h);
          penX = cmd.x; penY = cmd.y;
          break;
        case 'lineTo':
          ctx.lineTo(x + cmd.x * w, y + cmd.y * h);
          penX = cmd.x; penY = cmd.y;
          break;
        case 'cubicBezTo':
          ctx.bezierCurveTo(
            x + cmd.x1 * w, y + cmd.y1 * h,
            x + cmd.x2 * w, y + cmd.y2 * h,
            x + cmd.x  * w, y + cmd.y  * h,
          );
          penX = cmd.x; penY = cmd.y;
          break;
        case 'arcTo': {
          const rw = cmd.wr * w;
          const rh = cmd.hr * h;
          if (rw <= 0 || rh <= 0) break;
          const stRad = (cmd.stAng * Math.PI) / 180;
          const swRad = (cmd.swAng * Math.PI) / 180;
          const penAbsX = x + penX * w;
          const penAbsY = y + penY * h;
          const cx = penAbsX - rw * Math.cos(stRad);
          const cy = penAbsY - rh * Math.sin(stRad);
          const endRad = stRad + swRad;
          ctx.ellipse(cx, cy, rw, rh, 0, stRad, endRad, swRad < 0);
          penX = (cx + rw * Math.cos(endRad) - x) / w;
          penY = (cy + rh * Math.sin(endRad) - y) / h;
          break;
        }
        case 'close':
          ctx.closePath();
          break;
      }
    }
  }
}
