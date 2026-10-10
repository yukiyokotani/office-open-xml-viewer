import type { PathCmd } from '../types/common';

/**
 * An endpoint of a custom-geometry path, used to place a line-end decoration
 * (arrow head) on a freeform / curved line.
 *
 * Coordinates `(x, y)` are in the same **normalised** [0,1] space as the
 * incoming {@link PathCmd}s (relative to the shape bounding box). `(dx, dy)` is
 * the **outward** tangent direction at that endpoint, also expressed in
 * normalised space — i.e. it points *away* from the line so an arrow head drawn
 * along it faces outward (head at the start, tail at the end).
 *
 * The tangent is returned as a raw direction vector rather than a baked angle
 * because the shape may be scaled anisotropically (w ≠ h). The caller scales the
 * vector by the box dimensions — `atan2(dy·h, dx·w)` — to obtain the device-space
 * orientation. {@link CustGeomEndpoint.angle} provides the isotropic angle
 * (`atan2(dy, dx)`) for convenience / testing.
 */
export interface CustGeomEndpoint {
  /** Normalised x in [0,1]. */
  x: number;
  /** Normalised y in [0,1]. */
  y: number;
  /** Outward tangent x-component (normalised space). */
  dx: number;
  /** Outward tangent y-component (normalised space). */
  dy: number;
  /** Isotropic outward angle in radians (`atan2(dy, dx)`). */
  angle: number;
}

/** Two endpoints may be `null` independently (closed end ⇒ no arrow head). */
export interface CustGeomEndpoints {
  start: CustGeomEndpoint | null;
  end: CustGeomEndpoint | null;
}

const EPS = 1e-9;

function isDrawCmd(cmd: PathCmd): boolean {
  return cmd.cmd === 'lineTo' || cmd.cmd === 'cubicBezTo' ||
    cmd.cmd === 'quadBezTo' || cmd.cmd === 'arcTo';
}

function makeEndpoint(x: number, y: number, dx: number, dy: number): CustGeomEndpoint {
  // Normalise signed zero so the tangent (and any atan2 derived from it) is
  // canonical: `atan2(-0, -k)` is -π while `atan2(0, -k)` is +π — same
  // direction, but the negative zero leaks an arbitrary sign into the angle.
  const ndx = dx === 0 ? 0 : dx;
  const ndy = dy === 0 ? 0 : dy;
  return { x, y, dx: ndx, dy: ndy, angle: Math.atan2(ndy, ndx) };
}

/**
 * Forward tangent (direction of travel) at the **start** of a sub-path, i.e.
 * leaving `(penX, penY)` along the first drawn command `cmd`.
 */
function startForwardTangent(
  penX: number,
  penY: number,
  cmd: PathCmd,
): { dx: number; dy: number } {
  switch (cmd.cmd) {
    case 'lineTo':
      return { dx: cmd.x - penX, dy: cmd.y - penY };
    case 'cubicBezTo': {
      // Tangent at t=0 of a cubic Bézier points toward the first control point.
      let dx = cmd.x1 - penX;
      let dy = cmd.y1 - penY;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Degenerate first control point: fall back to the second, then the end.
        dx = cmd.x2 - penX;
        dy = cmd.y2 - penY;
      }
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        dx = cmd.x - penX;
        dy = cmd.y - penY;
      }
      return { dx, dy };
    }
    case 'quadBezTo': {
      let dx = cmd.x1 - penX;
      let dy = cmd.y1 - penY;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        dx = cmd.x - penX;
        dy = cmd.y - penY;
      }
      return { dx, dy };
    }
    case 'arcTo': {
      // Ellipse: P(t) = (cx + wr cos t, cy + hr sin t), with the pen at t=stAng
      // (matches buildCustomPath). Tangent dP/dt = (-wr sin t, hr cos t); for
      // sweep>0 travel increases t, for sweep<0 it decreases t.
      const stRad = (cmd.stAng * Math.PI) / 180;
      const dir = cmd.swAng < 0 ? -1 : 1;
      const dx = -cmd.wr * Math.sin(stRad) * dir;
      const dy = cmd.hr * Math.cos(stRad) * dir;
      return { dx, dy };
    }
    default:
      return { dx: 0, dy: 0 };
  }
}

/**
 * Advance the pen across one command, returning the new pen position. This is
 * the single source of truth for pen traversal and mirrors `buildCustomPath`
 * **exactly** — in particular a degenerate arc (`wr<=0 || hr<=0`) is a no-op
 * that leaves the pen unchanged (buildCustomPath does `break` before advancing
 * the pen), so the next drawn command starts from the pre-arc point.
 *
 * `close` and any non-drawing command leave the pen where it is.
 */
function advancePen(px: number, py: number, cmd: PathCmd): { x: number; y: number } {
  switch (cmd.cmd) {
    case 'moveTo':
    case 'lineTo':
    case 'cubicBezTo':
    case 'quadBezTo':
      return { x: cmd.x, y: cmd.y };
    case 'arcTo': {
      // Same degenerate guard as buildCustomPath (`rw<=0 || rh<=0`, with w,h>0
      // this reduces to wr<=0 || hr<=0): skip the arc and keep the pen put.
      if (cmd.wr <= 0 || cmd.hr <= 0) return { x: px, y: py };
      const stRad = (cmd.stAng * Math.PI) / 180;
      const endRad = stRad + (cmd.swAng * Math.PI) / 180;
      // Centre back-calculated from the pen at stAng (matches buildCustomPath):
      // cx = penX - wr cos(stAng); end = (cx + wr cos(endRad), cy + hr sin(endRad)).
      const cx = px - cmd.wr * Math.cos(stRad);
      const cy = py - cmd.hr * Math.sin(stRad);
      return { x: cx + cmd.wr * Math.cos(endRad), y: cy + cmd.hr * Math.sin(endRad) };
    }
    default:
      return { x: px, y: py };
  }
}

/**
 * Forward tangent (direction of travel) **arriving** at the end of a drawn
 * command, given the point `(prevX, prevY)` it started from. The end point is
 * computed via {@link advancePen} so it always agrees with buildCustomPath's
 * traversal (a degenerate arc contributes no movement and no tangent).
 */
function endForwardTangent(
  prevX: number,
  prevY: number,
  cmd: PathCmd,
): { dx: number; dy: number; x: number; y: number } {
  const { x, y } = advancePen(prevX, prevY, cmd);
  switch (cmd.cmd) {
    case 'lineTo':
      return { dx: cmd.x - prevX, dy: cmd.y - prevY, x, y };
    case 'cubicBezTo': {
      // Tangent at t=1 points from the second control point to the end.
      let dx = cmd.x - cmd.x2;
      let dy = cmd.y - cmd.y2;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Degenerate: second control coincides with the end → use first control.
        dx = cmd.x - cmd.x1;
        dy = cmd.y - cmd.y1;
      }
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        // Fully degenerate → chord from the previous point to the end.
        dx = cmd.x - prevX;
        dy = cmd.y - prevY;
      }
      return { dx, dy, x, y };
    }
    case 'quadBezTo': {
      let dx = cmd.x - cmd.x1;
      let dy = cmd.y - cmd.y1;
      if (Math.abs(dx) < EPS && Math.abs(dy) < EPS) {
        dx = cmd.x - prevX;
        dy = cmd.y - prevY;
      }
      return { dx, dy, x, y };
    }
    case 'arcTo': {
      // A degenerate arc draws nothing (advancePen kept the pen put), so it has
      // no outward tangent — matching buildCustomPath, which skips it entirely.
      if (cmd.wr <= 0 || cmd.hr <= 0) return { dx: 0, dy: 0, x, y };
      const stRad = (cmd.stAng * Math.PI) / 180;
      const endRad = stRad + (cmd.swAng * Math.PI) / 180;
      const dir = cmd.swAng < 0 ? -1 : 1;
      const dx = -cmd.wr * Math.sin(endRad) * dir;
      const dy = cmd.hr * Math.cos(endRad) * dir;
      return { dx, dy, x, y };
    }
    default:
      return { dx: 0, dy: 0, x, y };
  }
}

/** Track the pen position across one sub-path to find its final terminal point. */
function terminalPoint(cmds: readonly PathCmd[]): { x: number; y: number } | null {
  let px = 0;
  let py = 0;
  let started = false;
  for (const cmd of cmds) {
    if (cmd.cmd === 'moveTo') started = true;
    ({ x: px, y: py } = advancePen(px, py, cmd));
  }
  return started ? { x: px, y: py } : null;
}

/** Does this sub-path form a closed loop (explicit `close` or terminal ≈ start)? */
function isClosed(cmds: readonly PathCmd[]): boolean {
  if (cmds.some((c) => c.cmd === 'close')) return true;
  const first = cmds.find((c) => c.cmd === 'moveTo') as
    | Extract<PathCmd, { cmd: 'moveTo' }>
    | undefined;
  if (!first) return false;
  const term = terminalPoint(cmds);
  if (!term) return false;
  // Only a path that actually draws something can be an implicit loop.
  const hasDraw = cmds.some(isDrawCmd);
  if (!hasDraw) return false;
  return Math.abs(term.x - first.x) < EPS && Math.abs(term.y - first.y) < EPS;
}

/**
 * Extract the start and end decoration points of a custom-geometry path.
 *
 * - `start` = the very first `moveTo` of the first sub-path. Its tangent is the
 *   **reverse** of the direction leaving that point (so a head faces outward).
 * - `end`   = the terminal point of the last drawn command of the last sub-path.
 *   Its tangent is the direction of travel arriving there (already outward).
 *
 * An end belonging to a **closed** sub-path returns `null` (no arrow head), which
 * matches PowerPoint: line-end decorations only apply to open paths.
 *
 * Relationship to {@link getConnectorAnchors} (preset-geometry/index.ts): that
 * helper serves the *preset* connector/callout geometries (`straightConnector1`,
 * `bentConnector*`, `curvedConnector*`, …). It evaluates a preset's guide
 * formulas and returns tip points + tangent **angles already in device space**,
 * and it only walks `m`/`l`/`C` commands (presets never emit `arcTo`). This
 * helper is its custom-geometry counterpart: it consumes the *raw* normalised
 * `<a:custGeom>` sub-paths (free-form / curve shapes), handles `arcTo`, and
 * returns the tangent as a **normalised direction vector** so the caller can
 * scale it for an anisotropic box before taking `atan2`. The two are kept
 * separate on purpose — they take different inputs and different coordinate
 * spaces, so merging them would only add branching.
 *
 * Currently the only consumer is the pptx renderer (free-form / curve lines with
 * `<a:ln><a:headEnd|tailEnd>`). When docx / xlsx grow free-form line-end arrow
 * support they are expected to share this helper rather than re-deriving the
 * endpoints, which is why it lives in `@silurus/ooxml-core`.
 *
 * @param subpaths Normalised (`[0,1]`) custGeom sub-paths.
 */
export function getCustGeomEndpoints(subpaths: readonly (readonly PathCmd[])[]): CustGeomEndpoints {
  const result: CustGeomEndpoints = { start: null, end: null };
  if (!subpaths || subpaths.length === 0) return result;

  // ── start: first sub-path's opening moveTo + first drawn command ──────────
  const firstSub = subpaths[0];
  if (firstSub && firstSub.length > 0 && !isClosed(firstSub)) {
    const mv = firstSub.find((c) => c.cmd === 'moveTo') as
      | Extract<PathCmd, { cmd: 'moveTo' }>
      | undefined;
    const firstDraw = firstSub.find(isDrawCmd);
    if (mv && firstDraw) {
      const fwd = startForwardTangent(mv.x, mv.y, firstDraw);
      if (Math.abs(fwd.dx) > EPS || Math.abs(fwd.dy) > EPS) {
        // Outward (head) tangent is the reverse of the travel direction.
        result.start = makeEndpoint(mv.x, mv.y, -fwd.dx, -fwd.dy);
      }
    }
  }

  // ── end: last sub-path's final drawn command ─────────────────────────────
  const lastSub = subpaths[subpaths.length - 1];
  if (lastSub && lastSub.length > 0 && !isClosed(lastSub)) {
    // Replay the pen to the point just before the last drawn command.
    let px = 0;
    let py = 0;
    let lastDrawIdx = -1;
    for (let i = 0; i < lastSub.length; i++) {
      if (isDrawCmd(lastSub[i])) lastDrawIdx = i;
    }
    if (lastDrawIdx >= 0) {
      for (let i = 0; i < lastDrawIdx; i++) {
        ({ x: px, y: py } = advancePen(px, py, lastSub[i]));
      }
      const t = endForwardTangent(px, py, lastSub[lastDrawIdx]);
      if (Math.abs(t.dx) > EPS || Math.abs(t.dy) > EPS) {
        result.end = makeEndpoint(t.x, t.y, t.dx, t.dy);
      }
    }
  }

  return result;
}
