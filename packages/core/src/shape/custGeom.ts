import type { GeometryPathSink } from './path-data';
import type { PathCmd } from '../types/common';

export interface CustomGeometryBounds {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Conservative canvas-space bounds of normalized DrawingML custom geometry.
 *
 * CT_Path2D coordinates may be negative or exceed the path coordinate system,
 * so an `<a:xfrm>` rectangle is not a safe effect crop. Control points and the
 * complete ellipse behind an arc are included deliberately: that can
 * overestimate exact curve extrema, but never clips valid paint.
 */
export function getCustomGeometryBounds(
  subpaths: readonly (readonly PathCmd[])[],
  x: number,
  y: number,
  w: number,
  h: number,
): CustomGeometryBounds | null {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  const include = (px: number, py: number) => {
    if (!Number.isFinite(px) || !Number.isFinite(py)) return;
    minX = Math.min(minX, px);
    minY = Math.min(minY, py);
    maxX = Math.max(maxX, px);
    maxY = Math.max(maxY, py);
  };

  for (const commands of subpaths) {
    let penX = 0;
    let penY = 0;
    for (const command of commands) {
      switch (command.cmd) {
        case 'moveTo':
        case 'lineTo':
          penX = command.x;
          penY = command.y;
          include(x + penX * w, y + penY * h);
          break;
        case 'cubicBezTo':
          include(x + command.x1 * w, y + command.y1 * h);
          include(x + command.x2 * w, y + command.y2 * h);
          penX = command.x;
          penY = command.y;
          include(x + penX * w, y + penY * h);
          break;
        case 'quadBezTo':
          include(x + command.x1 * w, y + command.y1 * h);
          penX = command.x;
          penY = command.y;
          include(x + penX * w, y + penY * h);
          break;
        case 'arcTo': {
          const radiusX = Math.abs(command.wr * w);
          const radiusY = Math.abs(command.hr * h);
          if (radiusX <= 0 || radiusY <= 0) break;
          const start = command.stAng * Math.PI / 180;
          const sweep = command.swAng * Math.PI / 180;
          const penAbsX = x + penX * w;
          const penAbsY = y + penY * h;
          const centerX = penAbsX - radiusX * Math.cos(start);
          const centerY = penAbsY - radiusY * Math.sin(start);
          include(centerX - radiusX, centerY - radiusY);
          include(centerX + radiusX, centerY + radiusY);
          const end = start + sweep;
          penX = (centerX + radiusX * Math.cos(end) - x) / w;
          penY = (centerY + radiusY * Math.sin(end) - y) / h;
          break;
        }
        case 'close':
          break;
      }
    }
  }

  if (!Number.isFinite(minX)) return null;
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

/**
 * Build a canvas path from normalised custGeom sub-paths.
 * Coordinates in each PathCmd are in [0,1] relative to the shape bounding box;
 * this function maps them to the canvas pixel rectangle (x, y, w, h).
 *
 * Pen position is tracked so `arcTo` can back-calculate the ellipse centre
 * from the current pen point and `stAng`.
 */
export function buildCustomPath(
  ctx: GeometryPathSink,
  subpaths: readonly (readonly PathCmd[])[],
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
        case 'quadBezTo':
          ctx.quadraticCurveTo(
            x + cmd.x1 * w,
            y + cmd.y1 * h,
            x + cmd.x * w,
            y + cmd.y * h,
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
