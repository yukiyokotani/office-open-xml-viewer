/**
 * Preset-shape rendering engine. Drives ECMA-376 §20.1.9 preset geometries
 * off the authoritative `presetShapeDefinitions.xml` data (shipped as
 * `presets.json`), instead of hand-rolled per-shape switch cases.
 *
 * Public entry point: `renderPresetShape`. Returns true when the shape was
 * found and drawn; callers fall back to their legacy codepath on false.
 */

import presetsJson from './presets.json';
import { createEvaluator, compileFormula, type CompiledFormula } from './evaluator';
import { applyPresetPath, type PresetPath } from './path-executor';

interface PresetDef {
  adj: [string, string][];
  gd: [string, string][];
  paths: PresetPath[];
}

const PRESETS = presetsJson as unknown as Record<string, PresetDef>;

/**
 * A preset definition's adjust/guide formulas, tokenised once. The raw formula
 * strings in a {@link PresetDef} are immutable (presets.json ships ~200 fixed
 * definitions), so splitting each postfix expression into `{op, argTokens}` need
 * only happen once per preset — not once per shape per render, which is what the
 * old inline `expr.trim().split(/\s+/)` inside the evaluator did. The `paths` are
 * NOT part of this: their per-command tokens are resolved directly by the path
 * executor and vary only by the (already-cheap) `evaluator.resolve` lookup.
 */
interface CompiledDef {
  adj: [string, CompiledFormula][];
  gd: [string, CompiledFormula][];
}

/**
 * Lazily-populated cache of compiled adjust/guide formulas, keyed by the
 * PresetDef OBJECT IDENTITY. Every def we evaluate is a stable singleton — an
 * entry of the module-level `PRESETS` map or the module-level `RECT_DEF` — so
 * identity keying is sound and needs no name plumbing. A def is compiled the
 * first time it is evaluated and reused thereafter (a used preset is compiled
 * once for the life of the module; unused presets are never compiled).
 */
const COMPILED = new WeakMap<PresetDef, CompiledDef>();

/** Get (compiling on first use) the tokenised adjust/guide formulas for a def. */
function compiledDefFor(def: PresetDef): CompiledDef {
  let c = COMPILED.get(def);
  if (!c) {
    c = {
      adj: def.adj.map(([name, fmla]) => [name, compileFormula(fmla)]),
      gd: def.gd.map(([name, fmla]) => [name, compileFormula(fmla)]),
    };
    COMPILED.set(def, c);
  }
  return c;
}

/**
 * Build an evaluator for a def, drawing its adjust/guide formulas from the
 * compiled cache. The single choke point every `createEvaluator` call goes
 * through so the tokenisation is shared across the render / silhouette / anchor
 * entry points.
 */
function evaluatorForDef(
  def: PresetDef,
  w: number,
  h: number,
  adj: (number | null | undefined)[],
) {
  const c = compiledDefFor(def);
  return createEvaluator({ w, h, adj }, c.adj, c.gd);
}

export function hasPreset(geom: string): boolean {
  return geom.toLowerCase() in PRESETS;
}

/**
 * Append a preset geometry's outline to the CURRENT path of `ctx` (no
 * `beginPath` / `fill` / `stroke` — the caller owns those). Every `<path>` of the
 * preset is emitted as its own subpath, so the result is suitable for use as a
 * clip region, a silhouette to stroke, or an even-odd cut-out.
 *
 * This is the geometry-only counterpart of {@link renderPresetShape}: it shares
 * the same evaluator + path executor, so a picture clipped by `<a:prstGeom>`
 * (ECMA-376 §20.1.9.18 — a picture's preset geometry is its clip silhouette) is
 * traced by exactly the same code that draws the equivalent `<p:sp>`. Returns
 * false when the preset name is unknown so the caller can fall back to a rect.
 *
 * Note: fill modifiers (`lighten` / `darken`) and multi-path highlight overlays
 * are irrelevant to a silhouette, so every path is emitted regardless of its
 * `fill` mode — the union of all subpaths is the shape's outer outline.
 */
export function buildPresetGeometryPath(
  ctx: CanvasRenderingContext2D,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: (number | null | undefined)[] = [],
): boolean {
  const def = PRESETS[geom.toLowerCase()];
  if (!def) return false;
  const evaluator = evaluatorForDef(def, w, h, adj);
  for (const path of def.paths) {
    applyPresetPath(ctx, path, evaluator, x, y, w, h);
  }
  return true;
}

/** Append only the fill-bearing paths of a preset geometry.
 *
 * CT_Path2D paths whose `fill` is `none` are decorative stroke overlays, not
 * part of the object's painted silhouette. Picture and image-fill clipping must
 * therefore exclude them; otherwise an open accent/leader path is implicitly
 * closed by Canvas and cuts a false region into the image. Fill modifiers such
 * as `lighten` and `darken` remain included because those paths do carry fill.
 */
export function buildPresetGeometryFillPath(
  ctx: CanvasRenderingContext2D,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: (number | null | undefined)[] = [],
): boolean {
  const def = PRESETS[geom.toLowerCase()];
  if (!def) return false;
  const evaluator = evaluatorForDef(def, w, h, adj);
  for (const path of def.paths) {
    if (path.fill === 'none') continue;
    applyPresetPath(ctx, path, evaluator, x, y, w, h);
  }
  return true;
}

// A wedge callout's tail only protrudes when its tip is OUTSIDE the body. When
// the author drags the tip into the body (e.g. to hide the tail), PowerPoint
// renders just the base shape — no inward dent. The literal ECMA-376 preset
// path always draws the tail vertices, which on a thin near-horizontal edge
// shows as a visible notch. So when the tip is inside the body bounds, fall
// back to the base geometry. With the preset DEFAULT adjusts the tip is
// outside, so ordinary callouts keep their tails.
const WEDGE_CALLOUT_BASE: Record<string, string | null> = {
  wedgeroundrectcallout: 'roundrect', // corner radius = adj3
  wedgeellipsecallout: 'ellipse',
  wedgerectcallout: null, // → inline RECT_DEF
};
const RECT_DEF: PresetDef = {
  adj: [],
  gd: [],
  paths: [{
    w: null, h: null, fill: null, stroke: true, extrusionOk: false,
    cmds: [['m', 'l', 't'], ['l', 'r', 't'], ['l', 'r', 'b'], ['l', 'l', 'b'], ['c']],
  }],
};

/** Effective adjust value i, falling back to the preset's declared default. */
function effAdj(adj: (number | null | undefined)[], def: PresetDef, i: number): number {
  const s = adj[i];
  if (typeof s === 'number') return s;
  const d = def.adj[i];
  return d ? Number(d[1].replace(/^val\s+/, '')) || 0 : 0;
}

function effectivePreset(
  key: string,
  w: number,
  h: number,
  adj: (number | null | undefined)[],
): { def: PresetDef; adj: (number | null | undefined)[] } | null {
  let def = PRESETS[key];
  if (!def) return null;

  // Suppress a wedge-callout tail whose tip sits inside the body (see note above).
  if (key in WEDGE_CALLOUT_BASE) {
    const a1 = effAdj(adj, def, 0); // dxPos fraction (×1e5) of width from centre
    const a2 = effAdj(adj, def, 1); // dyPos fraction (×1e5) of height from centre
    const xPos = w / 2 + (w * a1) / 100000;
    const yPos = h / 2 + (h * a2) / 100000;
    const tipInside = xPos >= 0 && xPos <= w && yPos >= 0 && yPos <= h;
    if (tipInside) {
      const baseKey = WEDGE_CALLOUT_BASE[key];
      if (baseKey === 'roundrect') {
        def = PRESETS.roundrect;
        adj = [effAdj(adj, PRESETS[key], 2)]; // adj3 → corner radius
      } else if (baseKey && PRESETS[baseKey]) {
        def = PRESETS[baseKey];
        adj = [];
      } else {
        def = RECT_DEF;
        adj = [];
      }
    }
  }

  return { def, adj };
}

/**
 * Conservative painted-path bounds for a preset geometry in canvas space.
 *
 * DrawingML adjustments can place vertices outside the shape's `<a:xfrm>`
 * rectangle (wedge callout tips are the common example). Effect rasters must
 * include those vertices or their shadow/reflection/soft edge is clipped.
 * Curve control points and complete ellipse boxes are included deliberately:
 * they may overestimate a curve's exact extrema, but never crop valid paint.
 */
export function getPresetGeometryBounds(
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: (number | null | undefined)[] = [],
): { x: number; y: number; w: number; h: number } | null {
  const selected = effectivePreset(geom.toLowerCase(), w, h, adj);
  if (!selected) return null;

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
  const boundsCtx = {
    moveTo: include,
    lineTo: include,
    bezierCurveTo(x1: number, y1: number, x2: number, y2: number, ex: number, ey: number) {
      include(x1, y1);
      include(x2, y2);
      include(ex, ey);
    },
    quadraticCurveTo(x1: number, y1: number, ex: number, ey: number) {
      include(x1, y1);
      include(ex, ey);
    },
    ellipse(cx: number, cy: number, rx: number, ry: number) {
      include(cx - Math.abs(rx), cy - Math.abs(ry));
      include(cx + Math.abs(rx), cy + Math.abs(ry));
    },
    closePath() {},
  } as unknown as CanvasRenderingContext2D;
  const evaluator = evaluatorForDef(selected.def, w, h, selected.adj);
  for (const path of selected.def.paths) {
    applyPresetPath(boundsCtx, path, evaluator, x, y, w, h);
  }

  if (!Number.isFinite(minX)) return { x, y, w, h };
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
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
  const key = geom.toLowerCase();
  const selected = effectivePreset(key, w, h, adj);
  if (!selected) return false;
  const { def } = selected;
  adj = selected.adj;

  const evaluator = evaluatorForDef(def, w, h, adj);

  let shadowCleared = false;

  const lastIdx = def.paths.length - 1;
  for (let i = 0; i < def.paths.length; i++) {
    const path = def.paths[i];
    ctx.beginPath();
    applyPresetPath(ctx, path, evaluator, x, y, w, h);

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
        const overlay = tintOverlay(fillMode);
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

/**
 * For connector AND callout presets, return the canvas-space tip points and
 * outgoing tangent angles at the two ends of the geometry's *leader line*
 * (its last `<path>`). Used by the renderer to place line-end decorations
 * (headEnd / tailEnd) with the correct orientation.
 *
 * - Connectors (straight / bent / curved) have a single path, so the leader is
 *   that path.
 * - Callouts (callout1/2/3 and their border / accent variants) emit the leader
 *   as the LAST path — after the text rectangle and any accent bar. callout1's
 *   leader is a 2-point line; callout2/3 are 3-/4-point polylines whose tip is
 *   the final vertex.
 *
 * `start.angle` is the direction **from** the line **toward** the attach (start)
 * tip, so a "headEnd" decoration points outward; `end.angle` is the pen's travel
 * direction as it reaches the tip, so a "tailEnd" decoration points outward.
 */
export function getConnectorAnchors(
  geom: string,
  x: number, y: number, w: number, h: number,
  adj: (number | null | undefined)[],
): {
  start: { x: number; y: number; angle: number };
  end:   { x: number; y: number; angle: number };
  /** Every vertex of the leader polyline (m/l/C endpoints), in draw order, so
   *  callers can retract a decorated end before re-stroking the line. */
  vertices: Array<{ x: number; y: number }>;
} | null {
  const def = PRESETS[geom.toLowerCase()];
  if (!def || def.paths.length === 0) return null;
  // The decoratable line is the geometry's LAST <path>. For a connector that is
  // its only path; for a callout it is the leader line, which presets.json
  // emits after the text rectangle (and, for accent variants, the accent bar).
  // So paths[last] picks the leader for callout1/2/3 and is a no-op for the
  // single-path connectors.
  const path = def.paths[def.paths.length - 1];
  const evaluator = evaluatorForDef(def, w, h, adj);
  const sx = path.w != null ? w / path.w : 1;
  const sy = path.h != null ? h / path.h : 1;
  const toAbsX = (v: number) => x + v * sx;
  const toAbsY = (v: number) => y + v * sy;

  let startX = 0, startY = 0;
  let penX = 0, penY = 0;
  let startTanX = 0, startTanY = 0;
  let startTanSet = false;
  let endTanX = 0, endTanY = 0;
  const vertices: Array<{ x: number; y: number }> = [];

  for (const cmd of path.cmds) {
    switch (cmd[0]) {
      case 'm': {
        penX = toAbsX(evaluator.resolve(cmd[1]));
        penY = toAbsY(evaluator.resolve(cmd[2]));
        startX = penX; startY = penY;
        vertices.push({ x: penX, y: penY });
        break;
      }
      case 'l': {
        const nx = toAbsX(evaluator.resolve(cmd[1]));
        const ny = toAbsY(evaluator.resolve(cmd[2]));
        if (!startTanSet) { startTanX = nx - penX; startTanY = ny - penY; startTanSet = true; }
        endTanX = nx - penX; endTanY = ny - penY;
        penX = nx; penY = ny;
        vertices.push({ x: penX, y: penY });
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
        vertices.push({ x: penX, y: penY });
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
    vertices,
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
