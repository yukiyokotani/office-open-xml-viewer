import type { GeometryPathSink } from './path-data';
// Pre-set shape paths and helpers for OOXML <a:prstGeom>.
//
// Originally lived inline in the pptx renderer (see commit history of
// packages/pptx/src/renderer.ts). Moved here so the docx renderer can share
// the same prstGeom coverage for `wps:wsp` shapes that don't carry a
// `<a:custGeom>` payload — see ECMA-376 §20.1.9.18 for the exhaustive
// preset list.
//
// Functions are pure path-builders: they assume the caller has already
// called `ctx.beginPath()` and will call `ctx.fill()` / `ctx.stroke()`
// after this returns.
//
// PARTIALLY UNIFIED with the spec engine (Phase 4 A2): body rendering is
// already driven by the spec engine (shape/preset-geometry, presets.json)
// everywhere; this legacy switch remains reachable only as (a) the silhouette
// tracer for pptx effect masks and (b) the fallback for names the spec engine
// doesn't carry (`rect`). Presets whose hand-written case was proven
// coordinate-identical to the spec engine (see preset-parity.test.ts) have
// had their case DELETED and are routed through the engine via
// SPEC_MIGRATED_PRESETS below. Presets whose outline STRUCTURE differs but
// whose FILLED region is identical (see the fill-region scan in
// preset-parity.test.ts) are likewise deleted and routed through the engine
// via SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS: their only live legacy call path
// (the pptx effect-mask silhouette) fills path 0, so an equal fill region
// means an equal mask. Presets whose outline TOPOLOGY matches but whose
// vertices/arcs differ by a small, spec-CORRECT amount (legacy is an
// approximation, the spec engine is ECMA-faithful) are deleted and routed via
// SPEC_MIGRATED_SPEC_CORRECT_PRESETS: safe for the same path-0 reason, and here
// the visible body already renders through the spec engine so the mask simply
// matches the body it feathers.
//
// Every case still in the switch is a KNOWN DIFFERENCE from the spec engine,
// kept verbatim so no rendered pixel moves. Nature of the differences
// (details per preset in the parity harness's PRESET_PARITY_REPORT table):
//  - geometry approximations — the legacy math deviates from the ECMA-376
//    guide formulas (regular-polygon rings for pentagon…decagon and
//    star5/6/7; fixed-ratio instead of ss-based arrows/chevron/homeplate;
//    quadratic instead of circular corner arcs for round1rect / snip2* /
//    snipRoundRect / plaque; simplified
//    moon/teardrop/pie/chord/blockarc/wave/… constructions). The spec engine is
//    the ECMA-faithful side; converging means adopting its geometry and
//    re-approving VRT references. (The small-Δ, same-topology members of this
//    family — round2SameRect, round2DiagRect, dodecagon, star10 — have been
//    migrated; see the batch-6 bullet below.)
//  - structural, fill-DIFFERENT — the legacy body emits a different FILLED
//    silhouette (single-outline cloud/cloudCallout, no 3D highlight overlays
//    for can/cube/bevel, merged contours for
//    plus/mathMultiply/mathNotEqual/arrow-callouts, donut hole winding). These
//    stay legacy because their mask silhouette would move.
//  - structural, fill-EQUIVALENT — same filled silhouette, difference confined
//    to decorative interior stroke lines the fill scan ignores. The first such
//    batch (A2 batch 5) — accentCallout1, accentBorderCallout1,
//    flowchartPredefinedProcess, flowchartSort, flowchartInternalStorage,
//    flowchartSummingJunction — was migrated via
//    SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS (case DELETED, mask unchanged).
//  - spec-correct micro-geometry — same outline TOPOLOGY, vertices/arcs off by a
//    small amount where the spec engine is the corrected side. A2 batch 6 —
//    round2SameRect, round2DiagRect (quad vs true arc corners, Δ ≤ 1.01px),
//    dodecagon (Δ ≤ 5.97px) and star10 (Δ ≤ 6.60px, both ring vs ECMA vertex
//    guides) — was migrated via SPEC_MIGRATED_SPEC_CORRECT_PRESETS (case DELETED;
//    the silhouette now matches the spec-engine body it feathers).
//  - intentional — `arc` (the engine's spec shape is pie-wedge fill + open
//    arc stroke; the legacy open arc is kept for effect-mask semantics, see
//    the pptx renderer's paintShapeBody) and the bent/curved connectors
//    (legacy draws a straight diagonal, engine the spec elbow/curve).
//  - `flowchartAlternateProcess` matches at default adjusts but honours an
//    adj1 the spec declares no avLst for; kept so rogue avLst payloads don't
//    change effect-mask silhouettes.

/* eslint-disable */

import { buildPresetGeometryPath } from './preset-geometry/geometry';

// ── Star helper ──────────────────────────────────────────────────────────────
export function drawStar(
  ctx: GeometryPathSink,
  cx: number, cy: number,
  rx: number, ry: number,
  points: number,
  innerRatio: number,
  startAngle = -Math.PI / 2
) {
  const total = points * 2;
  for (let i = 0; i < total; i++) {
    const angle = startAngle + (i * Math.PI) / points;
    const r = i % 2 === 0 ? 1.0 : innerRatio;
    const px = cx + rx * r * Math.cos(angle);
    const py = cy + ry * r * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

// ── Regular polygon helper ───────────────────────────────────────────────────
export function drawPolygon(
  ctx: GeometryPathSink,
  cx: number, cy: number,
  rx: number, ry: number,
  sides: number,
  startAngle = -Math.PI / 2
) {
  for (let i = 0; i < sides; i++) {
    const angle = startAngle + (i * 2 * Math.PI) / sides;
    const px = cx + rx * Math.cos(angle);
    const py = cy + ry * Math.sin(angle);
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  }
  ctx.closePath();
}

/**
 * Emulate OOXML `<arcTo>` on a Canvas path.
 *
 * OOXML arc semantics (DrawingML §20.1.9.4): `stAng`/`swAng` are *visual*
 * angles — the angle of the radius ray from the ellipse center, not the
 * parametric ellipse angle. That is why the canonical preset-shape `gdLst`
 * formulas compute angles with plain `at2` on raw dimensions and still land
 * on the correct point on an elongated ellipse.
 *
 * Canvas's `ellipse()` takes parametric angles, so we convert:
 *   θ_parametric = atan2(wR * sin θ_visual, hR * cos θ_visual)
 *
 * The center is placed so the pen sits on the ellipse at the parametric
 * equivalent of `stAng` (guaranteed non-degenerate when `wR, hR > 0`).
 *
 * Returns the arc's geometric end point, so the caller can chain.
 */
export function ooxmlArcTo(
  ctx: GeometryPathSink,
  curX: number, curY: number,
  wR: number, hR: number,
  stAng: number, swAng: number,
): { x: number; y: number } {
  const visualToParam = (v: number) => Math.atan2(wR * Math.sin(v), hR * Math.cos(v));
  const stP  = visualToParam(stAng);
  const endP = visualToParam(stAng + swAng);
  const cx   = curX - wR * Math.cos(stP);
  const cy   = curY - hR * Math.sin(stP);
  // Canvas draws from stP to endP in the direction set by `counterclockwise`.
  // OOXML positive swAng = clockwise in screen coords = parametric angle
  // increasing, so pass `counterclockwise = swAng < 0`.
  ctx.ellipse(cx, cy, Math.abs(wR), Math.abs(hR), 0, stP, endP, swAng < 0);
  return { x: cx + wR * Math.cos(endP), y: cy + hR * Math.sin(endP) };
}
/**
 * Non-spec alias labels the legacy switch historically accepted, mapped to
 * the canonical ECMA-376 preset whose case body they shared (or duplicated).
 * Only aliases whose canonical target has MIGRATED to the spec engine live
 * here — aliases of still-legacy presets keep their `case` labels below.
 */
const LEGACY_SPEC_ALIASES: Record<string, string> = {
  oval: 'ellipse',
  rtriangle: 'rttriangle',
  roundrectangle: 'roundrect',
  // Historic misspelling the legacy switch accepted as a synonym for
  // flowchartSummingJunction (presets.json only carries the correct spelling).
  flowchartsumingjunction: 'flowchartsummingjunction',
};

/**
 * Presets whose hand-written case was verified coordinate-identical to the
 * spec-driven engine (preset-parity.test.ts: subpath structure, closed flags,
 * winding, and ≤5e-3 px symmetric deviation on square/wide/tall boxes at
 * default AND perturbed adjust values) and therefore deleted from the switch.
 * `buildShapePath` routes them through `buildPresetGeometryPath`, so the
 * silhouette/fallback callers keep emitting bit-identical geometry while the
 * spec engine becomes the single source of truth for these shapes.
 */
export const SPEC_MIGRATED_PRESETS: ReadonlySet<string> = new Set([
  // batch 1 — basic solids & rectangles
  'ellipse',
  'rttriangle',
  'triangle',
  'diamond',
  'trapezoid',
  'roundrect',
  'snip1rect',
  'frame',
  'irregularseal1',
  'irregularseal2',
  // batch 2 — stars (star5/6/7/10 stay legacy: their inner-vertex rings
  // deviate from the spec guide formulas — see the parity table)
  'star4',
  'star8',
  'star12',
  'star16',
  'star24',
  'star32',
  // batch 3 — straight connectors, plain callout1 pair, matched arrows
  // (bent/curved connectors stay legacy: their case draws a straight
  // diagonal, not the spec elbow/curve. The accent callout1 pair is
  // fill-equivalent, not exact, and lives in
  // SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS below.)
  'line',
  'straightconnector1',
  'callout1',
  'bordercallout1',
  'leftuparrow',
  'quadarrowcallout',
  // batch 4 — math operators & flowchart shapes
  // (flowchartAlternateProcess is NOT migrated despite matching at default
  // adjusts: its legacy body honours an adj1 the spec declares no avLst for,
  // so rogue <a:avLst> payloads would change silhouettes)
  'mathequal',
  'mathplus',
  'mathminus',
  'flowchartdecision',
  'flowchartmanualinput',
  'flowchartconnector',
  'flowchartinputoutput',
  'flowchartmerge',
  'flowchartextract',
  'flowchartpreparation',
  'flowchartcollate',
]);

/**
 * Presets whose legacy case emits a DIFFERENT outline structure from the spec
 * engine (extra/fewer subpaths, shifted decorative lines) yet whose FILLED
 * region is identical — proven by the dense winding-rule fill scan in
 * preset-parity.test.ts (fillEquivalent), across the same square/wide/tall
 * boxes. These are safe to route through the spec engine because their ONLY
 * live legacy call path is the pptx effect-mask silhouette (paintShapeBody),
 * which fills path 0: an equal fill region ⇒ an equal mask. The decorative
 * interior strokes that differ (predefinedProcess side bars, sort/summing
 * divider lines, internalStorage inner rules, accent callout bar + leader) are
 * emitted by the spec engine's own direct-draw path (renderPresetShape), not
 * by this function, so their appearance is unchanged.
 *
 * Distinct from SPEC_MIGRATED_PRESETS (coordinate-identical) because the pin
 * test asserts fill-region equivalence for these, not exact geometry.
 *
 * A2 batch 5 (fill-equivalent).
 */
export const SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS: ReadonlySet<string> = new Set([
  'accentcallout1',
  'accentbordercallout1',
  'flowchartpredefinedprocess',
  'flowchartsort',
  'flowchartinternalstorage',
  'flowchartsummingjunction',
]);

/**
 * Presets whose legacy case and the spec engine share the SAME outline topology
 * (identical subpath count, closed flags, winding, and — for the round2* rects —
 * the same corner-arc segment structure) but whose vertex/arc coordinates differ
 * by a small, spec-CORRECT amount: the legacy body is an approximation and the
 * spec engine is the ECMA-376-faithful side (presets.json IS
 * presetShapeDefinitions.xml §20.1.9). Migrating adopts the spec geometry, so
 * these are NOT coordinate-identical — the pin below records the design-time Δ
 * instead of asserting exact parity (which is impossible by construction).
 *
 * The difference is bounded and its nature is per-preset:
 *  - round2samerect / round2diagrect (Δ ≤ 1.01px, max over the square/wide/tall
 *    audit boxes): the legacy corners are QUADRATIC Béziers; the spec draws true
 *    circular quarter-arcs (`a` commands). Same 8-point outline, arc vs quad only.
 *  - dodecagon (Δ ≤ 5.97px): the legacy body is a regular 12-gon inscribed in the
 *    bbox via `drawPolygon`; the spec uses the ECMA-376 gdLst vertex guides
 *    (`x1..x4`, `y1..y4`), which place the vertices on the spec's own construction
 *    rather than a uniform ring. Same 12-vertex closed polygon.
 *  - star10 (Δ ≤ 6.60px): the legacy body is a uniform 10-point star (single
 *    inner-radius ratio via `drawStar`); the spec uses the per-vertex
 *    `sx1..sx6`/`sy1..sy4` guides derived from the ECMA `hf` horizontal factor,
 *    so inner and outer vertices are not on concentric uniform rings. Same
 *    20-vertex closed star.
 *
 * Safe to route through the spec engine for the same reason as the sets above:
 * the ONLY live legacy call path for an engine-known preset is the pptx
 * effect-mask silhouette (paintShapeBody), which fills path 0 — and here the
 * whole visible BODY already renders through the spec engine (renderPresetShape),
 * so after this migration the silhouette simply matches the body it masks.
 *
 * Distinct from SPEC_MIGRATED_PRESETS (coordinate-identical, exact-parity pin)
 * and from SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS (different topology / decorative
 * strokes, fill-region-equivalence pin): here the topology is identical and the
 * geometry is close-but-not-equal because the spec side is the corrected one.
 *
 * A2 batch 6 (spec-correct micro-geometry).
 */
export const SPEC_MIGRATED_SPEC_CORRECT_PRESETS: ReadonlySet<string> = new Set([
  'round2samerect',
  'round2diagrect',
  'dodecagon',
  'star10',
]);

/** Build the canvas path for a given OOXML preset geometry (`<a:prstGeom>`).
 *
 * The caller is responsible for `ctx.beginPath()` / `ctx.fill()` /
 * `ctx.stroke()` — this function only emits path commands.
 *
 * @param geom OOXML preset name (e.g. "ellipse", "rtTriangle", "roundRect").
 *             Unknown values fall through to a plain rect.
 * @param adj  First adjustment value from avLst (0–100000 range), used by
 *             shapes like trapezoid.
 */
export function buildShapePath(
  ctx: GeometryPathSink,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: number | null = null,
  adj2: number | null = null,
  adj3: number | null = null,
  adj4: number | null = null,
) {
  const cx = x + w / 2;
  const cy = y + h / 2;

  // Migrated presets: the spec-driven engine is the single implementation.
  // Three tiers, all routed here: SPEC_MIGRATED_PRESETS (coordinate-identical to
  // the legacy body), SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS (same filled
  // silhouette, decorative strokes aside), and SPEC_MIGRATED_SPEC_CORRECT_PRESETS
  // (same topology, small spec-CORRECT vertex/arc drift). All three are safe
  // because the sole live legacy caller is the pptx effect-mask that fills
  // path 0. Everything not in any set keeps its legacy output — including the
  // plain-rect default — untouched.
  {
    const raw = geom.toLowerCase();
    const key = LEGACY_SPEC_ALIASES[raw] ?? raw;
    if (
      SPEC_MIGRATED_PRESETS.has(key) ||
      SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS.has(key) ||
      SPEC_MIGRATED_SPEC_CORRECT_PRESETS.has(key)
    ) {
      if (buildPresetGeometryPath(ctx, key, x, y, w, h, [adj, adj2, adj3, adj4])) return;
    }
  }

  // OOXML preset names are camelCase (e.g. `straightConnector1`, `roundRect`,
  // `rtTriangle`), but every `case` below is lowercase. Normalize here so the
  // catalog is matched case-insensitively — otherwise camelCase presets fall
  // through to the `default` plain-rect branch. (The pptx renderer already
  // lower-cases at its call site; this makes the function correct for every
  // caller, e.g. the docx renderer which passes the raw `prst` value.)
  switch (geom.toLowerCase()) {
    // ── Quadrilaterals ────────────────────────────────────────────────────────
    case 'parallelogram': {
      // adj controls horizontal slant; default 25000 = 25% of width
      const offset = w * Math.min(0.5, (adj ?? 25000) / 100000);
      ctx.moveTo(x + offset, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - offset, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Regular polygons ──────────────────────────────────────────────────────
    case 'pentagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 5);
      break;
    case 'hexagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 6, 0);
      break;
    case 'heptagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 7);
      break;
    case 'octagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 8, -Math.PI / 8);
      break;
    case 'decagon':
      drawPolygon(ctx, cx, cy, w / 2, h / 2, 10);
      break;
    // dodecagon migrated to the spec engine (A2 batch 6, spec-correct
    // micro-geometry): the legacy regular 12-gon ring drifts ≤5.97px from the
    // ECMA-376 gdLst vertex guides — see SPEC_MIGRATED_SPEC_CORRECT_PRESETS.

    // ── Stars ─────────────────────────────────────────────────────────────────
    // Inner-radius defaults from ECMA-376 prstGeom avLst: adj / 50000 = innerR / outerR.
    case 'star5':
    case 'star':
      drawStar(ctx, cx, cy, w / 2, h / 2, 5, (adj ?? 19098) / 50000);
      break;
    case 'star6':
      drawStar(ctx, cx, cy, w / 2, h / 2, 6, (adj ?? 28868) / 50000, 0);
      break;
    case 'star7':
      drawStar(ctx, cx, cy, w / 2, h / 2, 7, (adj ?? 34142) / 50000);
      break;
    // star10 migrated to the spec engine (A2 batch 6, spec-correct
    // micro-geometry): the legacy uniform-ring star drifts ≤6.60px from the
    // ECMA-376 per-vertex guides — see SPEC_MIGRATED_SPEC_CORRECT_PRESETS.

    // ── Arrows ────────────────────────────────────────────────────────────────
    case 'rightarrow': {
      // adj1=shaft height (% of h, default 50000), adj2=arrowhead from right (% of w, default 50000)
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 50000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x, sy + sh);
      ctx.closePath();
      break;
    }
    case 'leftarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 50000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x + w, sy);
      ctx.lineTo(x + ahw, sy);
      ctx.lineTo(x + ahw, y);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + ahw, y + h);
      ctx.lineTo(x + ahw, sy + sh);
      ctx.lineTo(x + w, sy + sh);
      ctx.closePath();
      break;
    }
    case 'uparrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(1, (adj2 ?? 50000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h);
      ctx.lineTo(sx, y + h);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }
    case 'downarrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(1, (adj2 ?? 50000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y + h);
      ctx.lineTo(x + w, y + h - ahh);
      ctx.lineTo(sx + sw, y + h - ahh);
      ctx.lineTo(sx + sw, y);
      ctx.lineTo(sx, y);
      ctx.lineTo(sx, y + h - ahh);
      ctx.lineTo(x, y + h - ahh);
      ctx.closePath();
      break;
    }
    case 'leftrightarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(0.5, (adj2 ?? 25000) / 100000);
      const sy = y + (h - sh) / 2;
      ctx.moveTo(x, cy);
      ctx.lineTo(x + ahw, y);
      ctx.lineTo(x + ahw, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x + ahw, sy + sh);
      ctx.lineTo(x + ahw, y + h);
      ctx.closePath();
      break;
    }
    case 'updownarrow': {
      const sw = w * Math.min(1, (adj  ?? 50000) / 100000);
      const ahh = h * Math.min(0.5, (adj2 ?? 25000) / 100000);
      const sx = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h - ahh);
      ctx.lineTo(x + w, y + h - ahh);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - ahh);
      ctx.lineTo(sx, y + h - ahh);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }
    case 'notchedrightarrow': {
      const sh = h * Math.min(1, (adj  ?? 50000) / 100000);
      const ahw = w * Math.min(1, (adj2 ?? 35000) / 100000);
      const sy = y + (h - sh) / 2;
      const notch = ahw * 0.43; // notch depth relative to arrowhead width
      ctx.moveTo(x, sy);
      ctx.lineTo(x + w - ahw, sy);
      ctx.lineTo(x + w - ahw, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h);
      ctx.lineTo(x + w - ahw, sy + sh);
      ctx.lineTo(x, sy + sh);
      ctx.lineTo(x + notch, cy);
      ctx.closePath();
      break;
    }

    // ── Process flow shapes ───────────────────────────────────────────────────
    case 'chevron': {
      // adj = kink position from left as fraction of width; default 50000 (50%)
      // Kink at x=kink: right arrow-tip spans from kink to w; left V-notch at kink
      const kink = w * Math.min(1, Math.max(0, (adj ?? 50000) / 100000));
      ctx.moveTo(x, y);
      ctx.lineTo(x + kink, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + kink, y + h);
      ctx.lineTo(x, y + h);
      if (kink > 0) ctx.lineTo(x + kink, cy);
      ctx.closePath();
      break;
    }
    case 'homeplate': {
      const tip = h * 0.4;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - tip);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - tip);
      ctx.closePath();
      break;
    }

    // ── Brackets / braces ─────────────────────────────────────────────────────
    case 'leftbracket': {
      // Square bracket [ shape. adj (default 8333) controls corner arc height
      // as fraction of h; clamp to [0, 50000] per OOXML spec.
      const a = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH2 = Math.min(h * a / 100000, h / 2); // never let arcs overlap
      // Top arc: (w, 0) → quadratic via (0, 0) → (0, arcH)
      ctx.moveTo(x + w, y);
      ctx.quadraticCurveTo(x, y, x, y + arcH2);
      // Straight left side — omit when arcs just meet (path continues from arc end)
      if (h - 2 * arcH2 > 0.5) ctx.lineTo(x, y + h - arcH2);
      // Bottom arc: (0, h-arcH) → quadratic via (0, h) → (w, h)
      ctx.quadraticCurveTo(x, y + h, x + w, y + h);
      break;
    }
    case 'rightbracket': {
      // Square bracket ] shape — mirror of leftBracket.
      const a = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH2 = Math.min(h * a / 100000, h / 2);
      ctx.moveTo(x, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + arcH2);
      if (h - 2 * arcH2 > 0.5) ctx.lineTo(x + w, y + h - arcH2);
      ctx.quadraticCurveTo(x + w, y + h, x, y + h);
      break;
    }
    case 'leftbrace': {
      // { shape
      const mid = cy;
      const nb = w * 0.45;
      ctx.moveTo(x + w, y);
      ctx.bezierCurveTo(x + w - nb, y, x + w - nb, mid - h * 0.08, x, mid);
      ctx.bezierCurveTo(x + w - nb, mid + h * 0.08, x + w - nb, y + h, x + w, y + h);
      break;
    }
    case 'rightbrace': {
      const mid = cy;
      const nb = w * 0.45;
      ctx.moveTo(x, y);
      ctx.bezierCurveTo(x + nb, y, x + nb, mid - h * 0.08, x + w, mid);
      ctx.bezierCurveTo(x + nb, mid + h * 0.08, x + nb, y + h, x, y + h);
      break;
    }

    // ── Callouts ──────────────────────────────────────────────────────────────
    // The callout1 family (2-point: attach + tip) is fully spec-engine-driven:
    // callout1/bordercallout1 via SPEC_MIGRATED_PRESETS and accentCallout1/
    // accentBorderCallout1 via SPEC_MIGRATED_FILL_EQUIVALENT_PRESETS (A2 batch 5)
    // — the accent bar + leader are decorative strokes that leave the filled
    // silhouette unchanged. callout2 / callout3 require adj5..adj8 which exceed
    // this function's 4-adjustment signature, so they also go through the
    // generic preset engine (`renderPresetShape`) reading the full presets.json.
    case 'wedgerectcallout': {
      // Wedge (triangle-tail) callout: rect + filled triangle pointer.
      ctx.rect(x, y, w, h * 0.8);
      const tipX = x + w * 0.2;
      const tipY = y + h;
      ctx.moveTo(x + w * 0.1, y + h * 0.8);
      ctx.lineTo(tipX, tipY);
      ctx.lineTo(x + w * 0.3, y + h * 0.8);
      ctx.closePath();
      break;
    }

    case 'wedgeellipsecallout': {
      // Ellipse body + triangular pointer to tip defined by adj/adj2
      // adj/adj2 are offsets from center in 1/100000 of shape dimensions
      const tipDx = (adj ?? -20000) / 100000 * w;
      const tipDy = (adj2 ?? 120000) / 100000 * h;
      const tipX = cx + tipDx;
      const tipY = cy + tipDy;
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      // Triangular pointer
      const angle = Math.atan2(tipDy, tipDx);
      const perp = Math.PI / 10;
      const rx = w / 2, ry = h / 2;
      const p1x = cx + rx * Math.cos(angle - perp);
      const p1y = cy + ry * Math.sin(angle - perp);
      const p2x = cx + rx * Math.cos(angle + perp);
      const p2y = cy + ry * Math.sin(angle + perp);
      ctx.moveTo(p1x, p1y);
      ctx.lineTo(tipX, tipY);
      ctx.lineTo(p2x, p2y);
      ctx.closePath();
      break;
    }
    case 'cloudcallout': {
      // Simplified cloud (series of arcs) + small circular tail
      const bumpR = Math.min(w, h) * 0.22;
      const bumps = [
        [cx - w * 0.25, y + h * 0.35],
        [cx - w * 0.10, y + h * 0.15],
        [cx + w * 0.10, y + h * 0.10],
        [cx + w * 0.28, y + h * 0.20],
        [cx + w * 0.35, y + h * 0.40],
      ] as [number, number][];
      // Draw cloud outline
      ctx.moveTo(bumps[0][0] - bumpR, bumps[0][1]);
      for (const [bx2, by2] of bumps) {
        ctx.arc(bx2, by2, bumpR, Math.PI, 0);
      }
      ctx.arc(cx, y + h * 0.65, w * 0.45, 0, Math.PI);
      ctx.closePath();
      // Tail: small circle leading to tip
      const tipX2 = cx + (adj ?? -20000) / 100000 * w;
      const tipY2 = cy + (adj2 ?? 120000) / 100000 * h;
      ctx.moveTo(cx + w * 0.05, y + h * 0.8);
      ctx.arc(tipX2, tipY2, Math.min(w, h) * 0.07, 0, Math.PI * 2);
      break;
    }

    // ── Connectors ────────────────────────────────────────────────────────────
    case 'bentconnector2':
    case 'bentconnector3':
    case 'bentconnector4':
    case 'bentconnector5':
    case 'curvedconnector2':
    case 'curvedconnector3':
    case 'curvedconnector4':
    case 'curvedconnector5':
      // Connectors run diagonally from top-left to bottom-right of their bounding box.
      // Flip transforms (already applied to ctx) handle other orientations.
      // (Straight `line`/`straightConnector1` migrated to the spec engine; the
      // bent/curved ones stay because this legacy body draws a straight
      // diagonal, not the spec elbow/curve geometry.)
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y + h);
      break;

    // ── Heart ─────────────────────────────────────────────────────────────────
    case 'heart': {
      ctx.moveTo(cx, y + h * 0.32);
      ctx.bezierCurveTo(cx, y, x + w * 0.05, y, x, y + h * 0.3);
      ctx.bezierCurveTo(x, y + h * 0.68, cx - w * 0.05, y + h * 0.78, cx, y + h);
      ctx.bezierCurveTo(cx + w * 0.05, y + h * 0.78, x + w, y + h * 0.68, x + w, y + h * 0.3);
      ctx.bezierCurveTo(x + w - w * 0.05, y, cx, y, cx, y + h * 0.32);
      break;
    }

    // ── Donut / ring ──────────────────────────────────────────────────────────
    case 'donut': {
      // OOXML: dr = min(wd2, hd2) * adj / 100000; iRx = wd2 - dr; iRy = hd2 - dr
      const rx = w / 2, ry = h / 2;
      const dr  = Math.min(rx, ry) * (adj ?? 25000) / 100000;
      const irx = rx - dr, iry = ry - dr;
      ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2, false);
      ctx.moveTo(cx + irx, cy);
      ctx.ellipse(cx, cy, irx, iry, 0, 0, Math.PI * 2, true);
      break;
    }

    // ── No smoking / prohibition sign ─────────────────────────────────────────
    // Ring = outer CW + inner CCW (nonzero creates donut hole).
    // Bar (UL→LR backslash): single CW path: LR arc (0°→90°) + diagonal line
    // + UL arc (180°→270°) + diagonal close — fills the bar strip in the inner hole.
    case 'nosmoking':
    case 'nosmokingsign': {
      const adjFrac = (adj ?? 18750) / 100000;
      const rx  = w / 2;
      const ry  = h / 2;
      const rix = rx * (1 - 2 * adjFrac);
      const riy = ry * (1 - 2 * adjFrac);
      // Outer ring: outer CW + inner CCW
      ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2, false);
      ctx.moveTo(cx + rix, cy);
      ctx.ellipse(cx, cy, rix, riy, 0, 0, Math.PI * 2, true);
      // Bar fill: LR quad (0°→90°) → diagonal → UL quad (180°→270°) → diagonal close
      ctx.moveTo(cx + rix, cy);
      ctx.ellipse(cx, cy, rix, riy, 0, 0, Math.PI / 2, false);
      ctx.lineTo(cx - rix, cy);
      ctx.ellipse(cx, cy, rix, riy, 0, Math.PI, 3 * Math.PI / 2, false);
      ctx.closePath();
      break;
    }

    // ── Wedge / pie slice ─────────────────────────────────────────────────────
    case 'pie': {
      const stAng = (adj  ?? 0)        / 21600000 * Math.PI * 2;
      const enAng = (adj2 ?? 16200000) / 21600000 * Math.PI * 2;
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, Math.min(w, h) / 2, stAng, enAng);
      ctx.closePath();
      break;
    }

    // ── Cloud ─────────────────────────────────────────────────────────────────
    case 'cloud': {
      // Simplified cloud using arcs
      const r = h * 0.28;
      ctx.arc(x + w * 0.25, y + h * 0.55, r, Math.PI, Math.PI * 1.8);
      ctx.arc(x + w * 0.45, y + h * 0.35, r * 1.1, Math.PI * 1.3, Math.PI * 1.9);
      ctx.arc(x + w * 0.65, y + h * 0.4, r, Math.PI * 1.5, Math.PI * 2);
      ctx.arc(x + w * 0.8, y + h * 0.6, r * 0.9, Math.PI * 1.6, Math.PI * 0.1);
      ctx.arc(x + w * 0.55, y + h * 0.75, r, 0, Math.PI * 0.7);
      ctx.arc(x + w * 0.25, y + h * 0.7, r * 0.9, 0, Math.PI);
      ctx.closePath();
      break;
    }

    // ── Parallelogram / funnel ────────────────────────────────────────────────
    case 'funnel': {
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(cx + w * 0.15, y + h);
      ctx.lineTo(cx - w * 0.15, y + h);
      ctx.closePath();
      break;
    }

    // ── Smiley face ───────────────────────────────────────────────────────────
    // Spec: filled circle body + two filled eye circles + smile quadratic arc
    case 'smileyface': {
      // Body circle
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
      ctx.closePath();
      // Left eye (filled sub-path, evenodd makes it a hole in fill)
      const eyeRx = w * 0.05;
      const eyeRy = h * 0.05;
      const eyeY  = cy - h * 0.12;
      ctx.moveTo(cx - w * 0.2 + eyeRx, eyeY);
      ctx.ellipse(cx - w * 0.2, eyeY, eyeRx, eyeRy, 0, 0, Math.PI * 2);
      // Right eye
      ctx.moveTo(cx + w * 0.2 + eyeRx, eyeY);
      ctx.ellipse(cx + w * 0.2, eyeY, eyeRx, eyeRy, 0, 0, Math.PI * 2);
      // Smile: open arc rendered as stroke sub-path
      ctx.moveTo(cx - w * 0.25, cy + h * 0.05);
      ctx.quadraticCurveTo(cx, cy + h * 0.3, cx + w * 0.25, cy + h * 0.05);
      break;
    }

    // ── Fold / document ───────────────────────────────────────────────────────
    case 'document':
    case 'foldedcorner': {
      const fold = Math.min(w, h) * 0.15;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - fold, y);
      ctx.lineTo(x + w, y + fold);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      ctx.moveTo(x + w - fold, y);
      ctx.lineTo(x + w - fold, y + fold);
      ctx.lineTo(x + w, y + fold);
      break;
    }

    // ── Snipped-corner rectangles ─────────────────────────────────────────────
    case 'snip2samerect': {
      // Two snipped corners (top-right + bottom-left); adj = snip size
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x + s, y + h);
      ctx.lineTo(x, y + h - s);
      ctx.closePath();
      break;
    }
    case 'snip2diagrect': {
      // Two snipped diagonal corners (top-right + bottom-left)
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + s, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h - s);
      ctx.lineTo(x + w - s, y + h);
      ctx.lineTo(x + s, y + h);
      ctx.lineTo(x, y + h - s);
      ctx.lineTo(x, y + s);
      ctx.closePath();
      break;
    }
    case 'sniproundrect': {
      // One snipped + one rounded corner
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const s = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + s, y);
      ctx.lineTo(x + w - s, y);
      ctx.lineTo(x + w, y + s);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.quadraticCurveTo(x, y, x + s, y);
      ctx.closePath();
      break;
    }
    case 'round1rect': {
      // One rounded corner (top-left); adj = corner size
      const a = Math.min(50000, Math.max(0, adj ?? 16667));
      const r = Math.min(w, h) * a / 100000;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
      break;
    }
    // round2samerect and round2diagrect migrated to the spec engine (A2 batch 6,
    // spec-correct micro-geometry): the legacy corners are QUADRATIC Béziers,
    // the spec draws true circular quarter-arcs (Δ ≤ 1.01px, same 8-point
    // outline) — see SPEC_MIGRATED_SPEC_CORRECT_PRESETS.

    // ── Misc shapes ───────────────────────────────────────────────────────────
    case 'plaque': {
      // Rectangle with concave quarter-circle corners
      const r = Math.min(w, h) * 0.25;
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w - r, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + r);
      ctx.lineTo(x + w, y + h - r);
      ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
      ctx.lineTo(x + r, y + h);
      ctx.quadraticCurveTo(x, y + h, x, y + h - r);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
      break;
    }
    case 'can': {
      const ry = h * 0.1;
      // Top face (full ellipse, filled + stroked as outline)
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, 0, Math.PI * 2);
      // Body (open path; fill() implicitly closes with top chord, stroke() draws open)
      ctx.moveTo(x, y + ry);
      ctx.lineTo(x, y + h - ry);
      ctx.ellipse(cx, y + h - ry, w / 2, ry, 0, Math.PI, 2 * Math.PI);
      ctx.lineTo(x + w, y + ry);
      break;
    }
    case 'cube': {
      const off = Math.min(w, h) * 0.2;
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - off);
      ctx.lineTo(x + w - off, y + h);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x, y + off);
      ctx.closePath();
      ctx.moveTo(x + off, y);
      ctx.lineTo(x + off, y + off);
      ctx.lineTo(x + w - off, y + off);
      ctx.moveTo(x + off, y + off);
      ctx.lineTo(x, y + off);
      break;
    }
    case 'bevel': {
      // Beveled rectangle (inset rectangle + corner lines)
      const bev = Math.min(w, h) * 0.1;
      ctx.rect(x, y, w, h);
      ctx.moveTo(x, y);
      ctx.lineTo(x + bev, y + bev);
      ctx.lineTo(x + w - bev, y + bev);
      ctx.lineTo(x + w, y);
      ctx.moveTo(x + w - bev, y + bev);
      ctx.lineTo(x + w - bev, y + h - bev);
      ctx.lineTo(x + w, y + h);
      ctx.moveTo(x + w - bev, y + h - bev);
      ctx.lineTo(x + bev, y + h - bev);
      ctx.lineTo(x, y + h);
      ctx.moveTo(x + bev, y + h - bev);
      ctx.lineTo(x + bev, y + bev);
      break;
    }
    case 'halfframe': {
      // L-shaped half-frame
      const th = Math.min(w, h) * 0.25;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + th);
      ctx.lineTo(x + th, y + th);
      ctx.lineTo(x + th, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'corner': {
      // L-shaped corner bracket
      const th = Math.min(w, h) * 0.25;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + th);
      ctx.lineTo(x + th, y + th);
      ctx.lineTo(x + th, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }
    case 'flowchartalternateprocess':
    case 'flowchartprocess': {
      const a2 = Math.min(50000, Math.max(0, adj ?? 16667));
      const r2 = Math.min(w, h) * a2 / 100000;
      ctx.roundRect(x, y, w, h, [{ x: r2, y: r2 }]);
      break;
    }
    case 'flowchartterminator': {
      // Stadium / pill shape
      const sr = Math.min(w, h) / 2;
      ctx.roundRect(x, y, w, h, [{ x: sr, y: sr }]);
      break;
    }
    case 'flowchartdocument': {
      // Rectangle with wavy bottom
      const waveH = h * 0.1;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      break;
    }
    // flowchartPredefinedProcess and flowchartSort migrated to the spec engine
    // (A2 batch 5, fill-equivalent): the side bars / centre divider are
    // decorative strokes; the filled body (rect / diamond) is identical.
    case 'moon': {
      // Crescent moon
      ctx.arc(cx, cy, Math.min(w, h) / 2, -Math.PI / 2, Math.PI / 2);
      ctx.arc(cx - w * 0.2, cy, Math.min(w, h) / 2, Math.PI / 2, -Math.PI / 2, true);
      ctx.closePath();
      break;
    }
    case 'arc': {
      // OOXML arc: adj = stAng (default 270°=top), adj2 = swAng (default 90°)
      const FULL = 21600000;
      const startA = (adj  ?? 16200000) / FULL * Math.PI * 2;
      const swingA = (adj2 ?? 5400000)  / FULL * Math.PI * 2;
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, startA, startA + swingA, swingA < 0);
      break;
    }

    // ── Math operator shapes (ECMA-376 presets) ───────────────────────────────
    case 'mathmultiply': {
      // ECMA-376 preset: "×" aligned to bbox diagonals, thickness = ss * a1 / 100000
      const a1 = Math.min(51965, Math.max(0, adj ?? 23520));
      const th = Math.min(w, h) * a1 / 100000;
      const ang = Math.atan2(h, w);
      const sa = Math.sin(ang), ca = Math.cos(ang);
      const halfTX = th / 2 * sa;
      const halfTY = th / 2 * ca;
      // Bar 1: corner (x,y) → (x+w, y+h)
      ctx.moveTo(x + halfTX,     y - halfTY);
      ctx.lineTo(x - halfTX,     y + halfTY);
      ctx.lineTo(x + w - halfTX, y + h + halfTY);
      ctx.lineTo(x + w + halfTX, y + h - halfTY);
      ctx.closePath();
      // Bar 2: corner (x+w, y) → (x, y+h)
      ctx.moveTo(x + w - halfTX, y - halfTY);
      ctx.lineTo(x + w + halfTX, y + halfTY);
      ctx.lineTo(x + halfTX,     y + h + halfTY);
      ctx.lineTo(x - halfTX,     y + h - halfTY);
      ctx.closePath();
      break;
    }


    case 'mathdivide': {
      const a1 = Math.min(36745, Math.max(1000, adj ?? 23520));
      const ma1 = -a1;
      const ma3h = (73490 + ma1) / 4;
      const ma3w = 36745 * w / h;
      const maxAdj3 = Math.min(ma3h, ma3w);
      const a3 = Math.min(maxAdj3, Math.max(1000, adj3 ?? 11760));
      const maxAdj2 = 73490 + (4 * a3) - a1;
      const a2 = Math.min(maxAdj2, Math.max(0, adj2 ?? 5880));
      const dy1 = h * a1 / 200000;
      const yg  = h * a2 / 100000;
      const rad = h * a3 / 100000;
      const dx1 = w * 73490 / 200000;
      const y3 = cy - dy1;
      const y4 = cy + dy1;
      const y2 = y3 - (yg + rad);
      const y1 = y2 - rad;
      const y5 = (y + h) - y1;
      const x1 = cx - dx1;
      const x2 = cx + dx1;
      ctx.rect(x1, y3, x2 - x1, y4 - y3);
      ctx.moveTo(cx + rad, y1 + rad);
      ctx.arc(cx, y1 + rad, rad, 0, Math.PI * 2);
      ctx.moveTo(cx + rad, y5 - rad);
      ctx.arc(cx, y5 - rad, rad, 0, Math.PI * 2);
      break;
    }

    // ── 4-direction arrow ────────────────────────────────────────────────────
    case 'quadarrow': {
      const sw  = w * (adj  ?? 23000) / 100000;
      const ahw = w * (adj2 ?? 30000) / 100000;
      const sx  = x + (w - sw) / 2;
      const sy2 = y + (h - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w - ahw, y + ahw);
      ctx.lineTo(x + w - ahw, sy2);
      ctx.lineTo(sx + sw, sy2);
      ctx.lineTo(sx + sw, y + ahw);
      ctx.lineTo(x + ahw, y + ahw);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - ahw, y + h - ahw);
      ctx.lineTo(sx + sw, y + h - ahw);
      ctx.lineTo(sx + sw, sy2 + sw);
      ctx.lineTo(x + w - ahw, sy2 + sw);
      ctx.lineTo(x + w - ahw, y + h - ahw);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x + ahw, y + h - ahw);
      ctx.lineTo(x + ahw, sy2 + sw);
      ctx.lineTo(sx, sy2 + sw);
      ctx.lineTo(sx, y + h - ahw);
      ctx.lineTo(x, cy);
      ctx.lineTo(x + ahw, y + ahw);
      ctx.lineTo(sx, y + ahw);
      ctx.lineTo(sx, sy2);
      ctx.lineTo(x + ahw, sy2);
      ctx.closePath();
      break;
    }
    // ── Wave ──────────────────────────────────────────────────────────────────
    // OOXML: wavy top and bottom filling the bounding box. adj=12500 (12.5% amplitude).
    case 'wave': {
      const wAmp = h * (adj ?? 12500) / 100000;
      const yw1 = y + wAmp;        // top wave baseline (wAmp below top)
      const yw2 = y + h - wAmp;    // bottom wave baseline (wAmp above bottom)
      // Top wave (L→R): peaks at y, troughs at y + 2*wAmp
      ctx.moveTo(x, yw1);
      ctx.bezierCurveTo(x + w * 0.25, y,             x + w * 0.25, y + wAmp * 2, x + w * 0.5, yw1);
      ctx.bezierCurveTo(x + w * 0.75, y + wAmp * 2,  x + w * 0.75, y,             x + w, yw1);
      // Right side
      ctx.lineTo(x + w, yw2);
      // Bottom wave (R→L, half-period shift): peaks toward y+h, troughs toward y+h-2*wAmp
      ctx.bezierCurveTo(x + w * 0.75, y + h,              x + w * 0.75, y + h - wAmp * 2, x + w * 0.5, yw2);
      ctx.bezierCurveTo(x + w * 0.25, y + h - wAmp * 2,   x + w * 0.25, y + h,             x, yw2);
      ctx.closePath();
      break;
    }

    // ── Double wave (wavy top AND bottom edges) ───────────────────────────────
    // OOXML default adj=6250 (6.25% amplitude). Bezier CPs stay inside bounding box.
    case 'doublewave': {
      const wAmp = h * (adj ?? 6250) / 100000;
      const y1 = y + wAmp;       // top wave baseline
      const y2 = y + h - wAmp;   // bottom wave baseline
      // Top wave (L→R): peaks at y (top), troughs at y + 2*wAmp
      ctx.moveTo(x, y1);
      ctx.bezierCurveTo(x + w * 0.25, y,            x + w * 0.25, y + wAmp * 2, x + w * 0.5, y1);
      ctx.bezierCurveTo(x + w * 0.75, y + wAmp * 2, x + w * 0.75, y,            x + w, y1);
      // Right side
      ctx.lineTo(x + w, y2);
      // Bottom wave (R→L): peaks at y+h (bottom), troughs at y+h - 2*wAmp
      ctx.bezierCurveTo(x + w * 0.75, y + h,              x + w * 0.75, y + h - wAmp * 2, x + w * 0.5, y2);
      ctx.bezierCurveTo(x + w * 0.25, y + h - wAmp * 2,   x + w * 0.25, y + h,             x, y2);
      // Left side (closePath draws left edge)
      ctx.closePath();
      break;
    }

    // ── Sun (8 triangular rays + central disc) ────────────────────────────────
    case 'sun': {
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR * ((adj ?? 25000) / 100000 + 0.5);
      const clampedInner = Math.min(innerR, outerR * 0.9);
      const halfRayAng = Math.PI / 16;
      for (let i = 0; i < 8; i++) {
        const a = (i / 8) * Math.PI * 2;
        ctx.moveTo(cx + clampedInner * Math.cos(a - halfRayAng), cy + clampedInner * Math.sin(a - halfRayAng));
        ctx.lineTo(cx + outerR       * Math.cos(a),             cy + outerR       * Math.sin(a));
        ctx.lineTo(cx + clampedInner * Math.cos(a + halfRayAng), cy + clampedInner * Math.sin(a + halfRayAng));
        ctx.closePath();
      }
      ctx.moveTo(cx + clampedInner, cy);
      ctx.arc(cx, cy, clampedInner, 0, Math.PI * 2);
      break;
    }

    // ── Lightning bolt ────────────────────────────────────────────────────────
    case 'lightningbolt': {
      ctx.moveTo(cx + w * 0.1, y);
      ctx.lineTo(x, cy - h * 0.05);
      ctx.lineTo(cx + w * 0.05, cy - h * 0.05);
      ctx.lineTo(cx - w * 0.1, y + h);
      ctx.lineTo(x + w, cy + h * 0.05);
      ctx.lineTo(cx - w * 0.05, cy + h * 0.05);
      ctx.closePath();
      break;
    }

    // ── Bracket pair [] ───────────────────────────────────────────────────────
    case 'bracketpair': {
      const a   = Math.min(50000, Math.max(0, adj ?? 8333));
      const arcH = h * a / 100000;
      ctx.moveTo(x + w * 0.4, y);
      ctx.quadraticCurveTo(x, y, x, y + arcH);
      if (h - 2 * arcH > 0) ctx.lineTo(x, y + h - arcH);
      ctx.quadraticCurveTo(x, y + h, x + w * 0.4, y + h);
      ctx.moveTo(x + w * 0.6, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + arcH);
      if (h - 2 * arcH > 0) ctx.lineTo(x + w, y + h - arcH);
      ctx.quadraticCurveTo(x + w, y + h, x + w * 0.6, y + h);
      break;
    }

    // ── Brace pair {} ─────────────────────────────────────────────────────────
    case 'bracepair': {
      const nb = w * 0.2;
      ctx.moveTo(x + w * 0.4, y);
      ctx.bezierCurveTo(x + w * 0.4 - nb, y, x + w * 0.4 - nb, cy - h * 0.08, x, cy);
      ctx.bezierCurveTo(x + w * 0.4 - nb, cy + h * 0.08, x + w * 0.4 - nb, y + h, x + w * 0.4, y + h);
      ctx.moveTo(x + w * 0.6, y);
      ctx.bezierCurveTo(x + w * 0.6 + nb, y, x + w * 0.6 + nb, cy - h * 0.08, x + w, cy);
      ctx.bezierCurveTo(x + w * 0.6 + nb, cy + h * 0.08, x + w * 0.6 + nb, y + h, x + w * 0.6, y + h);
      break;
    }

    // ── Chord (arc + closing line) ────────────────────────────────────────────
    case 'chord': {
      const startA = (adj  ?? 2700000)  / 21600000 * Math.PI * 2;
      const endA   = (adj2 ?? 16200000) / 21600000 * Math.PI * 2;
      ctx.ellipse(cx, cy, w / 2, h / 2, 0, startA, endA);
      ctx.closePath();
      break;
    }

    // ── Block arc ─────────────────────────────────────────────────────────────
    case 'blockarc': {
      const outerR    = Math.min(w, h) / 2;
      const stAngRaw  = adj  ?? 10800000;  // default 180° (left)
      const enAngRaw  = adj2 ?? 0;          // default 0° (right)
      const innerFrac = (adj3 ?? 25000) / 100000;
      const innerR    = outerR * (1 - innerFrac);
      const startA    = stAngRaw / 21600000 * Math.PI * 2;
      const endA      = enAngRaw / 21600000 * Math.PI * 2;
      ctx.arc(cx, cy, outerR, startA, endA, false);
      ctx.arc(cx, cy, innerR, endA, startA, true);
      ctx.closePath();
      break;
    }

    // ── Teardrop ──────────────────────────────────────────────────────────────
    case 'teardrop': {
      const r   = Math.min(w, h) * 0.4;
      const bCx = x + r;
      const bCy = y + h - r;
      ctx.arc(bCx, bCy, r, 0, Math.PI * 2 * 0.75);
      ctx.bezierCurveTo(bCx - r * 0.1, bCy - r, x + w - r, y + r, x + w, y);
      ctx.bezierCurveTo(x + w - r * 0.2, y + r * 0.5, bCx + r, bCy - r * 1.1, bCx + r, bCy);
      ctx.closePath();
      break;
    }

    // ── Diagonal stripe ───────────────────────────────────────────────────────
    case 'diagstripe': {
      const thH = h * (adj ?? 50000) / 100000;
      const x1  = thH * w / h;
      ctx.moveTo(x + x1, y);
      ctx.lineTo(x + w,      y);
      ctx.lineTo(x + w - x1, y + h);
      ctx.lineTo(x,          y + h);
      ctx.closePath();
      break;
    }

    // ── Wedge round-rect callout ──────────────────────────────────────────────
    case 'wedgeroundrectcallout': {
      const r2 = Math.min(w, h) * 0.1;
      ctx.roundRect(x, y, w, h * 0.85, r2);
      ctx.moveTo(x + w * 0.1, y + h * 0.85);
      ctx.lineTo(x + w * 0.2, y + h);
      ctx.lineTo(x + w * 0.3, y + h * 0.85);
      ctx.closePath();
      break;
    }

    // ── Arrow callouts ────────────────────────────────────────────────────────
    case 'rightarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 50000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x, sy, shW, shH);
      ctx.moveTo(x + shW, y); ctx.lineTo(x + w, cy); ctx.lineTo(x + shW, y + h); ctx.closePath();
      break;
    }
    case 'leftarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 50000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x + w - shW, sy, shW, shH);
      ctx.moveTo(x + w - shW, y); ctx.lineTo(x, cy); ctx.lineTo(x + w - shW, y + h); ctx.closePath();
      break;
    }
    case 'uparrowcallout': {
      const shW = w * (adj  ?? 50000) / 100000;
      const shH = h * (adj2 ?? 50000) / 100000;
      const sx  = x + (w - shW) / 2;
      ctx.rect(sx, y + shH, shW, h - shH);
      ctx.moveTo(x, y + shH); ctx.lineTo(cx, y); ctx.lineTo(x + w, y + shH); ctx.closePath();
      break;
    }
    case 'downarrowcallout': {
      const shW = w * (adj  ?? 50000) / 100000;
      const shH = h * (adj2 ?? 50000) / 100000;
      const sx  = x + (w - shW) / 2;
      ctx.rect(sx, y, shW, h - shH);
      ctx.moveTo(x, y + h - shH); ctx.lineTo(cx, y + h); ctx.lineTo(x + w, y + h - shH); ctx.closePath();
      break;
    }
    case 'leftrightarrowcallout': {
      const shH = h * (adj  ?? 50000) / 100000;
      const shW = w * (adj2 ?? 25000) / 100000;
      const sy  = y + (h - shH) / 2;
      ctx.rect(x + shW, sy, w - 2 * shW, shH);
      ctx.moveTo(x + shW, y); ctx.lineTo(x, cy); ctx.lineTo(x + shW, y + h); ctx.closePath();
      ctx.moveTo(x + w - shW, y); ctx.lineTo(x + w, cy); ctx.lineTo(x + w - shW, y + h); ctx.closePath();
      break;
    }

    // ── Left-right-up arrow ───────────────────────────────────────────────────
    case 'leftrightuparrow': {
      const sw  = w * (adj  ?? 25000) / 100000;
      const ahh = h * (adj2 ?? 30000) / 100000;
      const sx  = x + (w - sw) / 2;
      ctx.moveTo(cx, y);
      ctx.lineTo(x + w, y + ahh);
      ctx.lineTo(sx + sw, y + ahh);
      ctx.lineTo(sx + sw, y + h);
      ctx.lineTo(sx, y + h);
      ctx.lineTo(sx, y + ahh);
      ctx.lineTo(x, y + ahh);
      ctx.closePath();
      break;
    }

    // ── Left-up arrow ─────────────────────────────────────────────────────────
    // ECMA-376 leftUpArrow preset: L-shape with arrowheads on the up (vertical)
    // and left (horizontal) arms, meeting at the bottom-right outer corner.
    //   adj1 (default 25000): arrow-head overhang (shaft width control; dx3 = ss*a1/200000)
    //   adj2 (default 25000): shaft offset from bbox edge (= head half-width; dx4 = ss*a2/100000)
    //   adj3 (default 25000): arrow-head length along arm (x1 = ss*a3/100000)
    // ── U-turn arrow ──────────────────────────────────────────────────────────
    // Spec (ECMA-376): outer half-arc on top, arrowhead on right side pointing down
    case 'uturnarrow': {
      const sw     = w * (adj ?? 25000) / 100000;   // shaft width
      const outerR = (w - sw) / 2;                   // outer bend radius
      const innerR = Math.max(0, outerR - sw);        // inner bend radius
      const arcCX  = x + sw + outerR;                // arc center X
      const arcCY  = y + sw + outerR;                // arc center Y
      const ahW    = sw * 2;                          // arrowhead full width
      const ahBase = y + h - sw * 2.5;               // where arrowhead base starts
      // shaft: left side down, U-arc across top, right side down to arrowhead
      ctx.moveTo(x, y + h);
      ctx.lineTo(x, arcCY);
      ctx.arc(arcCX, arcCY, outerR, Math.PI, 0);
      ctx.lineTo(x + w, ahBase);
      // arrowhead (pointing downward on right side)
      ctx.lineTo(x + w + (ahW - sw) / 2, ahBase);
      ctx.lineTo(arcCX + sw / 2, y + h);  // tip
      ctx.lineTo(x + w - (ahW - sw) / 2 - sw, ahBase);
      ctx.lineTo(x + w - sw, ahBase);
      ctx.lineTo(x + w - sw, arcCY);
      ctx.arc(arcCX, arcCY, innerR, 0, Math.PI, true);
      ctx.lineTo(x + sw, y + h);
      ctx.closePath();
      break;
    }

    // ── Bent arrow / bent-up arrow ────────────────────────────────────────────
    case 'bentarrow':
    case 'bentuparrow': {
      const t = Math.min(w, h) * 0.25;
      ctx.moveTo(x, cy - t / 2);
      ctx.lineTo(x + w - t * 2, cy - t / 2);
      ctx.lineTo(x + w - t * 2, y + t);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x + w - t * 2, y + h - t);
      ctx.lineTo(x + w - t * 2, cy + t / 2);
      ctx.lineTo(x, cy + t / 2);
      ctx.closePath();
      break;
    }

    // ── Plus shape (non-math) ─────────────────────────────────────────────────
    case 'plus': {
      const t = Math.min(w, h) * (adj ?? 25000) / 100000;
      ctx.rect(cx - t, y, 2 * t, h);
      ctx.rect(x, cy - t, w, 2 * t);
      break;
    }

    // ── Math not-equal ────────────────────────────────────────────────────────
    // ECMA-376 prstGeom mathNotEqual:
    //   adj1 = bar thickness (default 23520, pin 0..50000)
    //   adj2 = cross angle in 60000ths of degrees (default 6600000 = 110°, pin 4200000..6600000)
    //   adj3 = gap between bars (default 11760, pin 0..100000-2*adj1)
    case 'mathnotequal': {
      const a1 = Math.min(50000, Math.max(0, adj  ?? 23520));
      const crAngRaw = Math.min(6600000, Math.max(4200000, adj2 ?? 6600000));
      const a3 = Math.min(100000 - 2 * a1, Math.max(0, adj3 ?? 11760));
      const dy1 = h * a1 / 100000;
      const dy2 = h * a3 / 200000;
      const dx1 = w * 73490 / 200000;
      const hd2 = h / 2;
      const cadj2 = (crAngRaw / 60000 - 90) * Math.PI / 180;
      const xadj2 = hd2 * Math.tan(cadj2);
      const len = Math.hypot(xadj2, hd2);
      const bhw = len * dy1 / hd2;
      // Bars centered on cx with width 2*dx1 ≈ 0.7349w
      ctx.rect(cx - dx1, cy - dy2 - dy1, 2 * dx1, dy1);
      ctx.rect(cx - dx1, cy + dy2,       2 * dx1, dy1);
      // Diagonal slash as a parallelogram: top at (cx+xadj2), bottom at (cx-xadj2).
      // bhw is horizontal thickness (so perpendicular thickness matches dy1).
      ctx.moveTo(cx + xadj2 - bhw / 2, y);
      ctx.lineTo(cx + xadj2 + bhw / 2, y);
      ctx.lineTo(cx - xadj2 + bhw / 2, y + h);
      ctx.lineTo(cx - xadj2 - bhw / 2, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: connector (circle) ─────────────────────────────────────────
    // ── Flowchart: delay (D-shape) ────────────────────────────────────────────
    case 'flowchartdelay': {
      const r = h / 2;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w - r, y);
      ctx.arc(x + w - r, cy, r, -Math.PI / 2, Math.PI / 2);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: display (pentagon-like) ────────────────────────────────────
    case 'flowchartdisplay': {
      const lx = w * 0.2;
      const rx = w * 0.15;
      ctx.moveTo(x + lx, y);
      ctx.lineTo(x + w - rx, y);
      ctx.arc(x + w - rx, cy, h / 2, -Math.PI / 2, Math.PI / 2);
      ctx.lineTo(x + lx, y + h);
      ctx.lineTo(x, cy);
      ctx.closePath();
      break;
    }

    // ── Flowchart: punched card (parallelogram approximation) ────────────────
    // (flowchartInputOutput left this shared body for the spec engine — the
    // w/5 slant IS its spec geometry; punchedCard's spec shape is a rectangle
    // with a snipped top-left corner, so this body remains a known diff.)
    case 'flowchartpunchedcard': {
      const sl = w * 0.2;
      ctx.moveTo(x + sl, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w - sl, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: off-page connector (pentagon pointing down) ────────────────
    case 'flowchartoffpageconnector': {
      const tipH = h * 0.3;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - tipH);
      ctx.lineTo(cx, y + h);
      ctx.lineTo(x, y + h - tipH);
      ctx.closePath();
      break;
    }

    // ── Flowchart: online storage / manual label (rect fallback) ─────────────
    case 'flowchartonlinestorage':
    case 'flowchartmanuallabel':
    case 'flowchartpuncheddisk': {
      ctx.rect(x, y, w, h);
      break;
    }

    // ── Horizontal scroll ─────────────────────────────────────────────────────
    case 'horizontalscroll': {
      const r = Math.min(w, h) * 0.15;
      ctx.roundRect(x + r, y, w - r, h, r);
      ctx.moveTo(x + r, y + r * 2);
      ctx.arc(x + r, y + r, r, Math.PI / 2, Math.PI * 2.5);
      break;
    }

    // ── Vertical scroll ───────────────────────────────────────────────────────
    case 'verticalscroll': {
      const r = Math.min(w, h) * 0.15;
      ctx.roundRect(x, y + r, w, h - r, r);
      ctx.moveTo(x + r * 2, y + r);
      ctx.arc(x + r, y + r, r, 0, Math.PI * 2);
      break;
    }

    // ── Ribbon ───────────────────────────────────────────────────────────────
    // ECMA-376 prstGeom ribbon: tails at top, main body extends downward with
    // two side fold tabs. adj1 = tail depth (default 16667, pin 0..33333),
    // adj2 = body width percent (default 50000, pin 25000..75000).
    case 'ribbon': {
      const a1 = Math.min(33333, Math.max(0, adj  ?? 16667));
      const a2 = Math.min(75000, Math.max(25000, adj2 ?? 50000));
      const dx2 = w * a2 / 200000;
      const wd8 = w / 8, wd32 = w / 32;
      const x2r = w / 2 - dx2, x9r = w / 2 + dx2;
      const x3r = x2r + wd32, x8r = x9r - wd32;
      const x5r = x2r + wd8,  x6r = x9r - wd8;
      const x4r = x5r - wd32, x7r = x6r + wd32;
      const x10r = w - wd8;
      const y1r = h * a1 / 200000;
      const y2r = h * a1 / 100000;
      const y4r = h - y2r;
      const y3r = y4r / 2;
      // Outer outline (straight-line approximation of wd32 arcs)
      ctx.moveTo(x,          y);
      ctx.lineTo(x + x4r,    y);
      ctx.lineTo(x + x3r,    y + y1r);
      ctx.lineTo(x + x8r,    y + y2r);
      ctx.lineTo(x + x7r,    y + y1r);
      ctx.lineTo(x + w,      y);
      ctx.lineTo(x + x10r,   y + y3r);
      ctx.lineTo(x + w,      y + y4r);
      ctx.lineTo(x + x9r,    y + y4r);
      ctx.lineTo(x + x9r,    y + h);
      ctx.lineTo(x + x3r,    y + h);
      ctx.lineTo(x + x2r,    y + y4r);
      ctx.lineTo(x,          y + y4r);
      ctx.lineTo(x + wd8,    y + y3r);
      ctx.closePath();
      break;
    }

    // ── Ribbon2 (mirrored vertically: tails at bottom, body above) ───────────
    case 'ribbon2': {
      const a1 = Math.min(33333, Math.max(0, adj  ?? 16667));
      const a2 = Math.min(75000, Math.max(25000, adj2 ?? 50000));
      const dx2 = w * a2 / 200000;
      const wd8 = w / 8, wd32 = w / 32;
      const x2r = w / 2 - dx2, x9r = w / 2 + dx2;
      const x3r = x2r + wd32, x8r = x9r - wd32;
      const x5r = x2r + wd8,  x6r = x9r - wd8;
      const x4r = x5r - wd32, x7r = x6r + wd32;
      const x10r = w - wd8;
      const dy1 = h * a1 / 200000;
      const dy2 = h * a1 / 100000;
      const y1r = h - dy1;          // tail upper ridge
      const y2r = h - dy2;          // tail bottom ridge
      const y4r = dy2;              // top of body bottom
      const y3r = (y4r + h) / 2;    // bottom indent of tails
      // Mirror of ribbon around horizontal center
      ctx.moveTo(x,          y + h);
      ctx.lineTo(x + x4r,    y + h);
      ctx.lineTo(x + x3r,    y + y1r);
      ctx.lineTo(x + x8r,    y + y2r);
      ctx.lineTo(x + x7r,    y + y1r);
      ctx.lineTo(x + w,      y + h);
      ctx.lineTo(x + x10r,   y + y3r);
      ctx.lineTo(x + w,      y + y4r);
      ctx.lineTo(x + x9r,    y + y4r);
      ctx.lineTo(x + x9r,    y);
      ctx.lineTo(x + x3r,    y);
      ctx.lineTo(x + x2r,    y + y4r);
      ctx.lineTo(x,          y + y4r);
      ctx.lineTo(x + wd8,    y + y3r);
      ctx.closePath();
      break;
    }

    // ── Ellipse ribbon (ECMA-376 prstGeom ellipseRibbon) ─────────────────────
    // Arched ribbon: top edge is a downward parabola, bottom has center fold.
    // adj1 = overall band depth (default 25000, pin 0..100000)
    // adj2 = body width % (default 50000, pin 25000..75000)
    // adj3 = arch depth (default 12500, pin minAdj3..adj1)
    case 'ellipseribbon': {
      const a1 = Math.min(100000, Math.max(0, adj  ?? 25000));
      const a2 = Math.min(75000,  Math.max(25000, adj2 ?? 50000));
      const minAdj3 = Math.max(0, a1 - (100000 - a1) / 2);
      const a3 = Math.min(a1, Math.max(minAdj3, adj3 ?? 12500));
      const wd8 = w / 8;
      const dx2 = w * a2 / 200000;
      const x2e = w / 2 - dx2;
      const x3e = x2e + wd8;
      const x4e = w - x3e;
      const x5e = w - x2e;
      const x6e = w - wd8;
      const dy1 = h * a3 / 100000;
      const f1 = 4 * dy1 / w;
      // top outer arch
      const q2a = x3e - x3e * x3e / w;
      const y1e = f1 * q2a;
      const cx1 = x3e / 2, cy1 = f1 * cx1;
      const cx2 = w - cx1;
      // top inner fold
      const q1b = h * a1 / 100000;
      const dy3 = q1b - dy1;
      const q4b = x2e - x2e * x2e / w;
      const q5  = f1 * q4b;
      const y3e = q5 + dy3;
      const q7  = (dy1 + dy3 - y3e) + dy1;
      const cy3 = q7 + dy3;
      const rh  = h - q1b;
      const y2e = (dy1 * 14 / 16 + rh) / 2;
      const y5e = q5 + rh;
      const y6e = y3e + rh;
      const cx4 = x2e / 2, cy4 = f1 * cx4 + rh;
      const cx5 = w - cx4;
      const cy6 = cy3 + rh;
      ctx.moveTo(x,            y);
      ctx.quadraticCurveTo(x + cx1, y + cy1, x + x3e, y + y1e);
      ctx.lineTo(x + x2e,       y + y3e);
      ctx.quadraticCurveTo(x + w / 2, y + cy3, x + x5e, y + y3e);
      ctx.lineTo(x + x4e,       y + y1e);
      ctx.quadraticCurveTo(x + cx2, y + cy1, x + w, y);
      ctx.lineTo(x + x6e,       y + y2e);
      ctx.lineTo(x + w,         y + rh);
      ctx.quadraticCurveTo(x + cx5, y + cy4, x + x5e, y + y5e);
      ctx.lineTo(x + x5e,       y + y6e);
      ctx.quadraticCurveTo(x + w / 2, y + cy6, x + x2e, y + y6e);
      ctx.lineTo(x + x2e,       y + y5e);
      ctx.quadraticCurveTo(x + cx4, y + cy4, x, y + rh);
      ctx.lineTo(x + wd8,       y + y2e);
      ctx.closePath();
      break;
    }

    // ── Ellipse ribbon 2 (ECMA-376 prstGeom ellipseRibbon2: mirrored) ────────
    case 'ellipseribbon2': {
      const a1 = Math.min(100000, Math.max(0, adj  ?? 25000));
      const a2 = Math.min(75000,  Math.max(25000, adj2 ?? 50000));
      const minAdj3 = Math.max(0, a1 - (100000 - a1) / 2);
      const a3 = Math.min(a1, Math.max(minAdj3, adj3 ?? 12500));
      const wd8 = w / 8;
      const dx2 = w * a2 / 200000;
      const x2e = w / 2 - dx2;
      const x3e = x2e + wd8;
      const x4e = w - x3e;
      const x5e = w - x2e;
      const x6e = w - wd8;
      const dy1 = h * a3 / 100000;
      const f1 = 4 * dy1 / w;
      const q2a = x3e - x3e * x3e / w;
      const u1  = f1 * q2a;
      const y1e = h - u1;
      const cx1 = x3e / 2;
      const cu1 = f1 * cx1;
      const cy1 = h - cu1;
      const cx2 = w - cx1;
      const q1b = h * a1 / 100000;
      const dy3 = q1b - dy1;
      const q4b = x2e - x2e * x2e / w;
      const q5  = f1 * q4b;
      const u3  = q5 + dy3;
      const y3e = h - u3;
      const q7  = (dy1 + dy3 - u3) + dy1;
      const cu3 = q7 + dy3;
      const cy3 = h - cu3;
      const rh  = h - q1b;
      const u2  = (dy1 * 14 / 16 + rh) / 2;
      const y2e = h - u2;
      const u5  = q5 + rh;
      const y5e = h - u5;
      const u6  = u3 + rh;
      const y6e = h - u6;
      const cx4 = x2e / 2;
      const q9  = f1 * cx4;
      const cu4 = q9 + rh;
      const cy4 = h - cu4;
      const cx5 = w - cx4;
      const cu6 = cu3 + rh;
      const cy6 = h - cu6;
      ctx.moveTo(x,            y + h);
      ctx.quadraticCurveTo(x + cx1, y + cy1, x + x3e, y + y1e);
      ctx.lineTo(x + x2e,       y + y3e);
      ctx.quadraticCurveTo(x + w / 2, y + cy3, x + x5e, y + y3e);
      ctx.lineTo(x + x4e,       y + y1e);
      ctx.quadraticCurveTo(x + cx2, y + cy1, x + w, y + h);
      ctx.lineTo(x + x6e,       y + y2e);
      ctx.lineTo(x + w,         y + q1b);
      ctx.quadraticCurveTo(x + cx5, y + cy4, x + x5e, y + y5e);
      ctx.lineTo(x + x5e,       y + y6e);
      ctx.quadraticCurveTo(x + w / 2, y + cy6, x + x2e, y + y6e);
      ctx.lineTo(x + x2e,       y + y5e);
      ctx.quadraticCurveTo(x + cx4, y + cy4, x, y + q1b);
      ctx.lineTo(x + wd8,       y + y2e);
      ctx.closePath();
      break;
    }

    // ── Circular arrow (donut sector + arrowhead) ─────────────────────────────
    // OOXML defaults: stAng=0 (east), swAng=270°, thickW=50% of radius
    case 'circulararrow': {
      const stAng  = ((adj2 ?? 0)        / 60000) * Math.PI / 180;
      const swAng  = ((adj  ?? 16200000) / 60000) * Math.PI / 180;  // default 270°
      const thkPct = (adj3 ?? 50000) / 100000;
      const outerR = Math.min(w, h) / 2;
      const innerR = outerR * (1 - thkPct);
      const midR   = (outerR + innerR) / 2;
      const arcW   = outerR - innerR;
      const endAng = stAng + swAng;

      // Arc body (donut sector): outer CW then inner CCW
      ctx.arc(cx, cy, outerR, stAng, endAng, false);
      ctx.arc(cx, cy, innerR, endAng, stAng, true);
      ctx.closePath();

      // Arrowhead at endAng: filled triangle pointing in clockwise tangent direction
      // Tangent (CW): (sin(endAng), -cos(endAng))
      const tx = Math.sin(endAng), ty = -Math.cos(endAng);
      // Tip: extend midR point by ahLen in tangent direction
      const ahLen = arcW * 1.5;
      const tipX = cx + midR * Math.cos(endAng) + ahLen * tx;
      const tipY = cy + midR * Math.sin(endAng) + ahLen * ty;
      ctx.moveTo(tipX, tipY);
      ctx.lineTo(cx + outerR * Math.cos(endAng), cy + outerR * Math.sin(endAng));  // outer base
      ctx.lineTo(cx + innerR * Math.cos(endAng), cy + innerR * Math.sin(endAng));  // inner base
      ctx.closePath();
      break;
    }

    // ── Curved directional arrows (ECMA-376 §20.1.9.11–14) ────────────────────
    // adj1 = shaft thickness (pin 0..a2), adj2 = arrowhead half-width (pin 0..maxAdj2),
    // adj3 = arrowhead length along main axis (pin 0..maxAdj3).
    case 'curvedrightarrow': {
      const ss  = Math.min(w, h);
      const hd2 = h / 2;
      const maxAdj2 = 50000 * h / ss;
      const a2  = Math.min(maxAdj2, Math.max(0, adj2 ?? 50000));
      const a1  = Math.min(a2,      Math.max(0, adj  ?? 25000));
      const th  = ss * a1 / 100000;
      const aw  = ss * a2 / 100000;
      const hR  = hd2 - (th + aw) / 4;
      const q10 = (2 * hR) ** 2 - th ** 2;
      const idx = Math.sqrt(Math.max(0, q10)) * w / (2 * hR);
      const maxAdj3 = 100000 * idx / ss;
      const a3  = Math.min(maxAdj3, Math.max(0, adj3 ?? 25000));
      const ah  = ss * a3 / 100000;
      const dy  = Math.sqrt(Math.max(0, w * w - ah * ah)) * hR / w;
      const y3  = hR + th;
      const y5  = hR + dy;    // +- hR dy 0 = hR + dy - 0
      const y7  = y3 + dy;
      const dh  = (aw - th) / 2;
      const y4  = y5 - dh;
      const y8  = y7 + dh;
      const y6  = h - aw / 2;
      const x1  = w - ah;
      const swAng  = Math.atan2(ah, dy);    // at2 returns angle of (dy, ah): but OOXML at2 a b = atan2(b, a), i.e. atan2(dy, ah)? Check spec — at2 x y returns the angle whose tan = y/x, so at2 ah dy = atan2(dy, ah).
      const mswAng = -swAng;
      const stAng  = Math.PI - swAng;       // cd2 - swAng
      // Outer path: start at (l, hR), outer upper arc, arrowhead, inner lower arc, close
      ctx.moveTo(x, y + hR);
      ooxmlArcTo(ctx, x, y + hR, w, hR, Math.PI, mswAng);
      ctx.lineTo(x + x1, y + y4);
      ctx.lineTo(x + w,  y + y6);
      ctx.lineTo(x + x1, y + y8);
      ctx.lineTo(x + x1, y + y7);
      ooxmlArcTo(ctx, x + x1, y + y7, w, hR, stAng, swAng);
      ctx.closePath();
      break;
    }
    case 'curvedleftarrow': {
      const ss  = Math.min(w, h);
      const hd2 = h / 2;
      const maxAdj2 = 50000 * h / ss;
      const a2  = Math.min(maxAdj2, Math.max(0, adj2 ?? 50000));
      const a1  = Math.min(a2,      Math.max(0, adj  ?? 25000));
      const th  = ss * a1 / 100000;
      const aw  = ss * a2 / 100000;
      const hR  = hd2 - (th + aw) / 4;
      const q10 = (2 * hR) ** 2 - th ** 2;
      const idx = Math.sqrt(Math.max(0, q10)) * w / (2 * hR);
      const maxAdj3 = 100000 * idx / ss;
      const a3  = Math.min(maxAdj3, Math.max(0, adj3 ?? 25000));
      const ah  = ss * a3 / 100000;
      const dy  = Math.sqrt(Math.max(0, w * w - ah * ah)) * hR / w;
      const y3  = hR + th;
      const y5  = hR + dy;
      const y7  = y3 + dy;
      const dh  = (aw - th) / 2;
      const y4  = y5 - dh;
      const y8  = y7 + dh;
      const y6  = h - aw / 2;
      const x1  = ah;
      const swAng  = Math.atan2(ah, dy);
      const q12    = th / 2;
      const dang2  = Math.atan2(q12, idx);
      const swAng2 = dang2 - swAng;
      const swAng3 = swAng - dang2;
      const stAng3 = -dang2;
      // moveTo (l, y6); lnTo (x1, y4); lnTo (x1, y5); arcTo wR=w hR=hR stAng=swAng swAng=swAng2;
      //   arcTo wR=w hR=hR stAng=stAng3 swAng=swAng3; lnTo (x1, y8); close
      ctx.moveTo(x,      y + y6);
      ctx.lineTo(x + x1, y + y4);
      ctx.lineTo(x + x1, y + y5);
      const p1 = ooxmlArcTo(ctx, x + x1, y + y5, w, hR, swAng, swAng2);
      ooxmlArcTo(ctx, p1.x, p1.y, w, hR, stAng3, swAng3);
      ctx.lineTo(x + x1, y + y8);
      ctx.closePath();
      break;
    }
    case 'curveduparrow': {
      const ss  = Math.min(w, h);
      const wd2 = w / 2;
      const maxAdj2 = 50000 * w / ss;
      const a2  = Math.min(maxAdj2, Math.max(0, adj2 ?? 50000));
      const a1  = Math.min(100000,  Math.max(0, adj  ?? 25000));
      const th  = ss * a1 / 100000;
      const aw  = ss * a2 / 100000;
      const wR  = wd2 - (th + aw) / 4;
      const q10 = (2 * wR) ** 2 - th ** 2;
      const idy = Math.sqrt(Math.max(0, q10)) * h / (2 * wR);
      const maxAdj3 = 100000 * idy / ss;
      const a3  = Math.min(maxAdj3, Math.max(0, adj3 ?? 25000));
      const ah  = ss * a3 / 100000;
      const dx  = Math.sqrt(Math.max(0, h * h - ah * ah)) * wR / h;
      const x3  = wR + th;
      const x5  = wR + dx;
      const x7  = x3 + dx;
      const dh  = (aw - th) / 2;
      const x4  = x5 - dh;
      const x8  = x7 + dh;
      const x6  = w - aw / 2;
      const y1  = ah;
      const swAng   = Math.atan2(ah, dx);
      const q12     = th / 2;
      const dang2   = Math.atan2(q12, idy);
      const swAng2  = dang2 - swAng;
      const swAng3  = swAng - dang2;   // +- swAng dang2 0 = swAng - dang2
      const stAng3  = Math.PI / 2 - swAng;  // cd4 - swAng
      const stAng2  = Math.PI / 2 - dang2;  // cd4 - dang2
      // moveTo (x6, t); lnTo (x8, y1); lnTo (x7, y1); arcTo wR=wR hR=h stAng=stAng3 swAng=swAng3;
      //   arcTo wR=wR hR=h stAng=stAng2 swAng=swAng2; lnTo (x4, y1); close
      ctx.moveTo(x + x6, y);
      ctx.lineTo(x + x8, y + y1);
      ctx.lineTo(x + x7, y + y1);
      const p1 = ooxmlArcTo(ctx, x + x7, y + y1, wR, h, stAng3, swAng3);
      ooxmlArcTo(ctx, p1.x, p1.y, wR, h, stAng2, swAng2);
      ctx.lineTo(x + x4, y + y1);
      ctx.closePath();
      break;
    }
    case 'curveddownarrow': {
      const ss  = Math.min(w, h);
      const wd2 = w / 2;
      const maxAdj2 = 50000 * w / ss;
      const a2  = Math.min(maxAdj2, Math.max(0, adj2 ?? 50000));
      const a1  = Math.min(100000,  Math.max(0, adj  ?? 25000));
      const th  = ss * a1 / 100000;
      const aw  = ss * a2 / 100000;
      const wR  = wd2 - (th + aw) / 4;
      const q10 = (2 * wR) ** 2 - th ** 2;
      const idy = Math.sqrt(Math.max(0, q10)) * h / (2 * wR);
      const maxAdj3 = 100000 * idy / ss;
      const a3  = Math.min(maxAdj3, Math.max(0, adj3 ?? 25000));
      const ah  = ss * a3 / 100000;
      const dx  = Math.sqrt(Math.max(0, h * h - ah * ah)) * wR / h;
      const x3  = wR + th;
      const x5  = wR + dx;
      const x7  = x3 + dx;
      const dh  = (aw - th) / 2;
      const x4  = x5 - dh;
      const x8  = x7 + dh;
      const x6  = w - aw / 2;
      const y1  = h - ah;
      const swAng   = Math.atan2(ah, dx);
      const q12     = th / 2;
      const dang2   = Math.atan2(q12, idy);
      const stAng   = 3 * Math.PI / 2 + swAng;   // 3cd4 + swAng
      const stAng2  = 3 * Math.PI / 2 - dang2;   // 3cd4 - dang2
      const swAng2  = dang2 - Math.PI / 2;       // dang2 - cd4
      const swAng3  = Math.PI / 2 - dang2;       // cd4 - dang2
      // ECMA: moveTo (x6, b); lnTo (x4, y1); lnTo (x5, y1); arcTo stAng=stAng swAng=mswAng;
      //   lnTo (x3, t); arcTo stAng=3cd4 swAng=swAng; lnTo (x8, y1); close
      ctx.moveTo(x + x6, y + h);
      ctx.lineTo(x + x4, y + y1);
      ctx.lineTo(x + x5, y + y1);
      ooxmlArcTo(ctx, x + x5, y + y1, wR, h, stAng, -swAng);
      ctx.lineTo(x + x3, y);
      ooxmlArcTo(ctx, x + x3, y, wR, h, 3 * Math.PI / 2, swAng);
      ctx.lineTo(x + x8, y + y1);
      void stAng2; void swAng2; void swAng3; void x7;
      ctx.closePath();
      break;
    }

    // ── Striped right arrow (3 stripes + arrowhead) ───────────────────────────
    // Spec: ssd = min(w,h), ssd32=ssd/32, ssd8=ssd/8 etc. adj=arrowhead length
    case 'stripedrightarrow': {
      const ssd   = Math.min(w, h);
      const ssd32 = ssd / 32;
      const ssd16 = ssd / 16;
      const ssd8  = ssd / 8;
      const shH   = ssd * (adj ?? 50000) / 100000;  // shaft height
      const ahW   = w * (adj2 ?? 50000) / 100000;   // arrowhead width
      const y1    = cy - shH / 2;
      const y2    = cy + shH / 2;
      const x4    = x + w - ahW;
      // stripe 1
      ctx.rect(x, y1, ssd32, shH);
      // stripe 2
      ctx.rect(x + ssd16, y1, ssd16, shH);
      // stripe 3 (narrower, bridging to arrowhead)
      ctx.rect(x + ssd8, y1, ssd8, shH);
      // arrow body + head
      ctx.moveTo(x4, y1);
      ctx.lineTo(x4, y);
      ctx.lineTo(x + w, cy);
      ctx.lineTo(x4, y + h);
      ctx.lineTo(x4, y2);
      ctx.lineTo(x + ssd8 * 2, y2);
      ctx.lineTo(x + ssd8 * 2, y1);
      ctx.closePath();
      break;
    }

    // ── Flowchart: magnetic disk (vertical cylinder) ──────────────────────────
    case 'flowchartmagneticdisk': {
      const ry = h * 0.15;
      ctx.moveTo(x, y + ry);
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, Math.PI, 0);
      ctx.lineTo(x + w, y + h - ry);
      ctx.ellipse(cx, y + h - ry, w / 2, ry, 0, 0, Math.PI);
      ctx.lineTo(x, y + ry);
      ctx.closePath();
      // top cap stroke line
      ctx.moveTo(x + w, y + ry);
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, 0, Math.PI);
      break;
    }

    // flowchartInternalStorage migrated to the spec engine (A2 batch 5,
    // fill-equivalent): the two inner rules are decorative strokes; the filled
    // body is the plain rectangle, identical to the engine's.

    // ── Flowchart: magnetic drum (cylinder on its side with left cap) ─────────
    case 'flowchartmagneticdrum': {
      const rx = w * 0.15;
      ctx.moveTo(x + rx, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x + rx, y + h);
      ctx.ellipse(x + rx, cy, rx, h / 2, 0, Math.PI / 2, -Math.PI / 2, true);
      ctx.closePath();
      // right cap open arc
      ctx.moveTo(x + w, y);
      ctx.ellipse(x + w, cy, rx, h / 2, 0, -Math.PI / 2, Math.PI / 2);
      break;
    }

    // flowchartSummingJunction (and its historic misspelling alias) migrated to
    // the spec engine (A2 batch 5, fill-equivalent): the inscribed X is a
    // decorative stroke; the filled body is the circle, identical to the
    // engine's. The alias is mapped in LEGACY_SPEC_ALIASES.

    // ── Flowchart: magnetic tape (circle with tail) ───────────────────────────
    case 'flowchartmagnetictape': {
      // circle from bottom going around, with a small tail at bottom-right
      const r = Math.min(w, h) / 2;
      const tailX = cx + r * 0.5;
      ctx.moveTo(cx, y + h);
      ctx.arc(cx, cy, r, Math.PI / 2, Math.PI / 2 + Math.PI * 2 * 0.875);
      ctx.lineTo(tailX, cy + r * 0.5);
      ctx.lineTo(tailX, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: punched tape (wave bottom) ─────────────────────────────────
    case 'flowchartpunchedtape': {
      const waveH = h * 0.12;
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      // second wave on top for symmetry
      ctx.moveTo(x, y + waveH);
      ctx.bezierCurveTo(x + w * 0.25, y, x + w * 0.75, y + waveH * 2, x + w, y + waveH);
      break;
    }

    // ── Flowchart: manual operation (inverted trapezoid) ─────────────────────
    case 'flowchartmanualoperation': {
      const sl = w * 0.15;
      ctx.moveTo(x + sl, y);
      ctx.lineTo(x + w - sl, y);
      ctx.lineTo(x + w, y + h);
      ctx.lineTo(x, y + h);
      ctx.closePath();
      break;
    }

    // ── Flowchart: multidocument (stacked wave documents) ────────────────────
    case 'flowchartmultidocument': {
      const waveH = h * 0.1;
      const shiftX = w * 0.04;
      // back documents (offset rects)
      ctx.rect(x + shiftX * 2, y - h * 0.08, w - shiftX * 2, h * 0.1);
      ctx.rect(x + shiftX, y - h * 0.04, w - shiftX, h * 0.06);
      // main document with wave bottom
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + h - waveH);
      ctx.bezierCurveTo(x + w * 0.75, y + h, x + w * 0.25, y + h - waveH * 2, x, y + h - waveH);
      ctx.closePath();
      break;
    }

    case 'rttriangle': {
      // Right triangle — right angle at bottom-left corner
      ctx.moveTo(x, y);
      ctx.lineTo(x, y + h);
      ctx.lineTo(x + w, y + h);
      ctx.closePath();
      break;
    }

    default:
      // rect and everything else
      ctx.rect(x, y, w, h);
      break;
  }
}
