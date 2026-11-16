import type { ArrowEnd, Stroke } from '../types/common';
import { drawArrowHead, resolveArrowPaint, tracePresetLeader, traceTrimmedPath } from './arrow';
import { lineEndRetract } from './arrow-geometry';
import type { GeometryPathSink } from './path-data';
import type { getConnectorAnchors } from './preset-geometry/geometry';

/**
 * Line-end decoration policy for preset geometries whose leader line carries
 * `<a:ln><a:headEnd|tailEnd>` (ECMA-376 §20.1.8.38 / §20.1.8.57), shared by
 * every DrawingML host (PPTX, XLSX, and the retained DOCX plan in
 * `drawingml-geometry.ts`).
 *
 * The decorated line is the geometry's LAST `<path>` (its leader): the only
 * path of a connector, the leader of a callout (after the text rectangle and
 * any accent bar). `getConnectorAnchors` resolves its two ends for all of them.
 */
export const CONNECTOR_PRESETS: ReadonlySet<string> = new Set([
  'line', 'straightconnector1',
  'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
  'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
]);

/** Callout family: callout1 = 2-point leader, callout2/3 = 3-/4-point polylines. */
export const CALLOUT_PRESETS: ReadonlySet<string> = new Set([
  'callout1', 'callout2', 'callout3',
  'bordercallout1', 'bordercallout2', 'bordercallout3',
  'accentcallout1', 'accentcallout2', 'accentcallout3',
  'accentbordercallout1', 'accentbordercallout2', 'accentbordercallout3',
]);

/** A preset (lower-case name) whose leader takes head/tail decorations. */
export function hasPresetLineEnds(name: string): boolean {
  return CONNECTOR_PRESETS.has(name) || CALLOUT_PRESETS.has(name);
}

/**
 * A leader that is re-stroked from its vertices, trimmed at its decorated
 * ends: every callout plus the straight / bent (polyline) connectors. Curved
 * connectors are excluded: their leader is a Bézier and the anchor vertices
 * only capture segment endpoints, so re-stroking from vertices would
 * straighten the curve. Their preset leader is trimmed along the curve instead.
 */
export function isRetractableLeaderPreset(name: string): boolean {
  return CALLOUT_PRESETS.has(name) || name === 'line' || name === 'straightconnector1' || name.startsWith('bentconnector');
}

const decorated = (end: ArrowEnd | undefined): boolean => !!end && end.type !== 'none';

/** How a preset's leader is painted once `renderPresetShape` skipped it. */
export interface PresetLeaderPlan {
  /** Device px each decorated end pulls the line back (observed Office rule in arrow-geometry.ts). */
  readonly headTrim: number;
  readonly tailTrim: number;
  /** Either end carries a decoration other than `none`. */
  readonly decorated: boolean;
  /**
   * `vertices`: re-stroke the trimmed vertex polyline; `curve`: re-stroke the
   * preset's curved leader trimmed along the curve; null: the preset engine
   * strokes the leader itself. Non-null means the host must pass
   * `skipTrailingStroke` to `renderPresetShape`, so the full-length leader
   * is not stroked under the trimmed one (a double line / a cap poking
   * through the arrow tip).
   */
  readonly restroke: 'vertices' | 'curve' | null;
}

/**
 * Plan the leader paint of a line-end preset (`name` lower-case), or null when
 * the preset takes no line ends or there is no stroke. `restrokeUndecorated`
 * keeps re-stroking retractable leaders from their vertices even without
 * decorations (the PPTX host composes the leader with its decorations).
 */
export function planPresetLeader(
  name: string, stroke: Stroke | null | undefined, scale: number, unitsPerPoint: number,
  restrokeUndecorated = false,
): PresetLeaderPlan | null {
  if (!stroke || !hasPresetLineEnds(name)) return null;
  const headTrim = stroke.headEnd ? lineEndRetract(stroke.headEnd, stroke, scale, unitsPerPoint) : 0;
  const tailTrim = stroke.tailEnd ? lineEndRetract(stroke.tailEnd, stroke, scale, unitsPerPoint) : 0;
  const isDecorated = decorated(stroke.headEnd) || decorated(stroke.tailEnd);
  const restroke = isRetractableLeaderPreset(name) && (isDecorated || restrokeUndecorated)
    ? 'vertices'
    : name.startsWith('curvedconnector') && (headTrim > 0 || tailTrim > 0) ? 'curve' : null;
  return { headTrim, tailTrim, decorated: isDecorated, restroke };
}

type ConnectorAnchors = NonNullable<ReturnType<typeof getConnectorAnchors>>;

/**
 * Trace the planned re-stroked leader into `sink` (the caller begins the path
 * and strokes it). Returns false when nothing was traced. The vertex leader is
 * trimmed across its whole length, so a short terminal leg cannot stop the
 * retraction early; a leader consumed by both ends traces nothing visible.
 */
export function tracePlannedPresetLeader(
  sink: GeometryPathSink, plan: PresetLeaderPlan, name: string, anchors: ConnectorAnchors,
  x: number, y: number, w: number, h: number, adj: (number | null | undefined)[],
): boolean {
  if (plan.restroke === 'vertices') {
    const points = anchors.vertices;
    if (points.length < 2) return false;
    traceTrimmedPath(sink, (path) => {
      path.moveTo(points[0].x, points[0].y);
      for (let i = 1; i < points.length; i++) path.lineTo(points[i].x, points[i].y);
    }, plan.headTrim, plan.tailTrim);
    return true;
  }
  if (plan.restroke === 'curve') return tracePresetLeader(sink, name, x, y, w, h, adj, plan.headTrim, plan.tailTrim);
  return false;
}

/** The two ends of a decorated line: tip points and outward angles (radians);
 *  null for an end that takes no decoration (a closed subpath). */
export interface LineEndAnchors {
  readonly start: { readonly x: number; readonly y: number; readonly angle: number } | null;
  readonly end: { readonly x: number; readonly y: number; readonly angle: number } | null;
}

/**
 * Paint the tail decoration at the line's end, then the head decoration at
 * its start, each in the stroke's resolved paint (`frame` / `rotation` /
 * `ptUnits` are the host's shade frame, as for its body stroke). `brush`
 * optionally gives the host's pattern coordinate transform and paints
 * gradient / pattern brushes in the host frame, so a decoration continues
 * the brush of the line it ends.
 */
export function drawPresetLineEndArrows(
  ctx: CanvasRenderingContext2D, anchors: LineEndAnchors, stroke: Stroke,
  scale: number, unitsPerPoint: number,
  frame: { x: number; y: number; w: number; h: number }, rotation: number, ptUnits: number,
  brush?: { readonly hostFrame: true; readonly patternCoordinateTransform?: DOMMatrix2DInit },
): void {
  const paint = (end: ArrowEnd, tip: { x: number; y: number; angle: number }) => drawArrowHead(
    ctx, tip.x, tip.y, tip.angle, end, stroke, scale, unitsPerPoint,
    resolveArrowPaint(ctx, stroke, scale, unitsPerPoint, tip.x, tip.y, tip.angle, end, frame, rotation, ptUnits,
      brush?.patternCoordinateTransform),
    brush?.hostFrame ?? false,
  );
  if (stroke.tailEnd && anchors.end) paint(stroke.tailEnd, anchors.end);
  if (stroke.headEnd && anchors.start) paint(stroke.headEnd, anchors.start);
}
