/** Canonical preset painter; geometry-only consumers import ./geometry. */
import { appendGeometryPath } from '../path-data';
import { resolvePresetGeometryPaths, pathFillModeOverlay } from './geometry';
export * from './geometry';

/**
 * Render a preset shape onto the canvas. Handles all paths (including
 * secondary outline-only / highlight paths) with per-path fill/stroke
 * semantics. The caller provides the base fillStyle and an `applyStroke`
 * closure that configures stroke properties (dash, width, colour, …)
 * immediately before each `ctx.stroke()` call.
 *
 * Returns false if the preset is unknown — fall back to legacy rendering.
 */
export function renderPresetShape(
  ctx: CanvasRenderingContext2D,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: (number | null | undefined)[],
  baseFill: string | CanvasGradient | CanvasPattern | null,
  applyAndStroke: (() => void) | null,
  clearShadow: () => void,
  opts?: {
    skipTrailingStroke?: boolean;
    /**
     * Paint the already-traced fill-bearing path. The painter may clip and draw,
     * but must leave the current path intact for tint overlays and stroking.
     * Returns true only when it painted the path.
     */
    paintFill?: (ctx: CanvasRenderingContext2D) => boolean;
  },
): boolean {
  const paths = resolvePresetGeometryPaths(geom, w, h, adj, undefined, x, y);
  if (!paths) return false;

  let shadowCleared = false;

  const lastIdx = paths.length - 1;
  for (let i = 0; i < paths.length; i++) {
    const path = paths[i];
    ctx.beginPath();
    appendGeometryPath(ctx, path.path);

    const fillMode = path.fill;
    const wantFill = fillMode !== 'none' && baseFill != null;

    if (fillMode !== 'none' && opts?.paintFill) {
      let painted = false;
      ctx.save();
      try {
        painted = opts.paintFill(ctx);
      } finally {
        ctx.restore();
      }
      if (painted) {
        const overlay = pathFillModeOverlay(fillMode);
        if (overlay) {
          ctx.save();
          ctx.fillStyle = overlay;
          ctx.fill();
          ctx.restore();
        }
        if (!shadowCleared) {
          clearShadow();
          shadowCleared = true;
        }
      }
    } else if (wantFill) {
      ctx.save();
      ctx.fillStyle = baseFill!;
      ctx.fill();
      // For "lighten" / "darken" modifiers, overlay a translucent tint so
      // multi-path 3D shapes (can, cube, pentagon) get highlights/shadows
      // without re-parsing the base fill.
      const overlay = pathFillModeOverlay(fillMode);
      if (overlay) {
        ctx.fillStyle = overlay;
        ctx.fill();
      }
      ctx.restore();
      if (!shadowCleared) {
        clearShadow();
        shadowCleared = true;
      }
    }

    if (path.stroke && applyAndStroke) {
      // A connector/callout's leader line is the geometry's trailing stroke
      // path with no fill — usually fill="none", but the `line` preset's sole
      // path uses fill:null, so treat a missing fill the same (else `line`
      // double-strokes and its cap pokes through the arrow tip). When the
      // caller re-strokes the leader retracted from its decorated ends (so the
      // line stops at the arrow base), skip it here. The accent bar is also
      // fill="none" but is spared because it is NOT the last path; rect borders
      // (fill≠none) likewise always stroke.
      const isTrailingLeader = i === lastIdx && (path.fill === 'none' || path.fill == null);
      if (!(opts?.skipTrailingStroke && isTrailingLeader)) {
        applyAndStroke();
      }
    }
  }

  return true;
}
