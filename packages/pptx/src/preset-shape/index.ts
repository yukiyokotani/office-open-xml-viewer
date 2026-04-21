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

function tintOverlay(mode: string | null): string | null {
  switch (mode) {
    case 'lighten':     return 'rgba(255,255,255,0.30)';
    case 'lightenLess': return 'rgba(255,255,255,0.15)';
    case 'darken':      return 'rgba(0,0,0,0.30)';
    case 'darkenLess':  return 'rgba(0,0,0,0.15)';
    default: return null;
  }
}
