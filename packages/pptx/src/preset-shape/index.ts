/**
 * Preset-shape rendering engine. Drives ECMA-376 §20.1.9 preset geometries
 * off the authoritative `presetShapeDefinitions.xml` data (shipped as
 * `presets.json`), instead of hand-rolled per-shape switch cases.
 *
 * Public entry point: `renderPresetShape`. Returns true when the shape was
 * found and drawn; callers fall back to their legacy codepath on false.
 */

import presetsJson from './presets.json';
import { createEvaluator } from './evaluator';
import { applyPresetPath, type PresetPath } from './path-executor';

interface PresetDef {
  adj: [string, string][];
  gd: [string, string][];
  paths: PresetPath[];
}

const PRESETS = presetsJson as unknown as Record<string, PresetDef>;

export function hasPreset(geom: string): boolean {
  return geom.toLowerCase() in PRESETS;
}

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
): boolean {
  const key = geom.toLowerCase();
  const def = PRESETS[key];
  if (!def) return false;

  const evaluator = createEvaluator({ w, h, adj }, def.adj, def.gd);

  let shadowCleared = false;

  for (const path of def.paths) {
    ctx.beginPath();
    applyPresetPath(ctx, path, evaluator, x, y, w, h);

    const fillMode = path.fill;
    const wantFill = fillMode !== 'none' && baseFill != null;

    if (wantFill) {
      ctx.save();
      ctx.fillStyle = baseFill!;
      ctx.fill();
      // For "lighten" / "darken" modifiers, overlay a translucent tint so
      // multi-path 3D shapes (can, cube, pentagon) get highlights/shadows
      // without re-parsing the base fill.
      const overlay = tintOverlay(fillMode);
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
      applyAndStroke();
    }
  }

  return true;
}

/**
 * For connector presets (straight / bent / curved), return the canvas-space
 * tip points and outgoing tangent angles at the start and end of the path.
 * Used by the renderer to place arrow heads with the correct orientation.
 *
 * `start.angle` is the direction **from** the path **toward** the start tip
 * (so arrow heads "headEnd" point outward); `end.angle` is the direction
 * the pen was moving as it reached the end tip.
 */
export function getConnectorAnchors(
  geom: string,
  x: number, y: number, w: number, h: number,
  adj: (number | null | undefined)[],
): {
  start: { x: number; y: number; angle: number };
  end:   { x: number; y: number; angle: number };
} | null {
  const def = PRESETS[geom.toLowerCase()];
  if (!def || def.paths.length === 0) return null;
  const path = def.paths[0];
  const evaluator = createEvaluator({ w, h, adj }, def.adj, def.gd);
  const sx = path.w != null ? w / path.w : 1;
  const sy = path.h != null ? h / path.h : 1;
  const toAbsX = (v: number) => x + v * sx;
  const toAbsY = (v: number) => y + v * sy;

  let startX = 0, startY = 0;
  let penX = 0, penY = 0;
  let startTanX = 0, startTanY = 0;
  let startTanSet = false;
  let endTanX = 0, endTanY = 0;

  for (const cmd of path.cmds) {
    switch (cmd[0]) {
      case 'm': {
        penX = toAbsX(evaluator.resolve(cmd[1]));
        penY = toAbsY(evaluator.resolve(cmd[2]));
        startX = penX; startY = penY;
        break;
      }
      case 'l': {
        const nx = toAbsX(evaluator.resolve(cmd[1]));
        const ny = toAbsY(evaluator.resolve(cmd[2]));
        if (!startTanSet) { startTanX = nx - penX; startTanY = ny - penY; startTanSet = true; }
        endTanX = nx - penX; endTanY = ny - penY;
        penX = nx; penY = ny;
        break;
      }
      case 'C': {
        const c1x = toAbsX(evaluator.resolve(cmd[1]));
        const c1y = toAbsY(evaluator.resolve(cmd[2]));
        const c2x = toAbsX(evaluator.resolve(cmd[3]));
        const c2y = toAbsY(evaluator.resolve(cmd[4]));
        const nx  = toAbsX(evaluator.resolve(cmd[5]));
        const ny  = toAbsY(evaluator.resolve(cmd[6]));
        if (!startTanSet) { startTanX = c1x - penX; startTanY = c1y - penY; startTanSet = true; }
        endTanX = nx - c2x; endTanY = ny - c2y;
        penX = nx; penY = ny;
        break;
      }
    }
  }

  // Start arrow points opposite the outgoing tangent (away from the path).
  const startAngle = Math.atan2(startTanY, startTanX) + Math.PI;
  const endAngle   = Math.atan2(endTanY,   endTanX);
  return {
    start: { x: startX, y: startY, angle: startAngle },
    end:   { x: penX,   y: penY,   angle: endAngle   },
  };
}

function tintOverlay(mode: string | null): string | null {
  switch (mode) {
    case 'lighten':     return 'rgba(255,255,255,0.30)';
    case 'lightenLess': return 'rgba(255,255,255,0.15)';
    case 'darken':      return 'rgba(0,0,0,0.30)';
    case 'darkenLess':  return 'rgba(0,0,0,0.15)';
    default: return null;
  }
}
