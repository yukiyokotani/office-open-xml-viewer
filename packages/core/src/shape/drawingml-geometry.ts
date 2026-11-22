import type { Stroke, ArrowEnd } from '../types/common';
import type { DrawingMLShapePaintPlan } from './drawingml-shape';
import { createGeometryPath, assertGeometryPath, GeometryWorkBudgetError, type GeometryPath } from './path-data';
import { resolveStrokeGeometry, type ResolvedStrokeGeometry } from './stroke-geometry';
import { resolveArrowGeometry, lineEndRetract, trimGeometryPathEnds, type ResolvedArrowGeometry } from './arrow-geometry';
import { buildShapePath } from './preset';
import { buildCustomPath } from './custGeom';
import { getCustGeomEndpoints } from './custgeom-endpoints';
import { resolvePresetGeometryPaths, buildPresetGeometryFillPath, getConnectorAnchors } from './preset-geometry/geometry';
import { geometryPathBounds, type GeometryMatrix, type GeometryBox } from './path-enclosure';
import { hasPresetLineEnds, isRetractableLeaderPreset } from './preset-line-ends';

export interface ResolvedDrawingMLPath {
  readonly path: GeometryPath;
  readonly fill: string | null;
  readonly fillRule: 'nonzero' | 'evenodd';
  readonly stroke: ResolvedStrokeGeometry | null;
}
export interface ResolvedDrawingMLArrow {
  readonly tipX: number; readonly tipY: number; readonly angle: number;
  readonly end: Readonly<ArrowEnd>;
  readonly geometry: ResolvedArrowGeometry;
}
export interface ResolvedDrawingMLGeometry {
  readonly width: number; readonly height: number; readonly scale: number;
  readonly originX: number; readonly originY: number;
  readonly sourceKey: string;
  readonly paths: readonly ResolvedDrawingMLPath[];
  readonly fillSilhouette: GeometryPath;
  readonly arrows: readonly ResolvedDrawingMLArrow[];
  readonly workUnits: number;
}
function sourceKey(plan: DrawingMLShapePaintPlan): string {
  return JSON.stringify({ geometry: plan.geometry, stroke: plan.stroke, fill: plan.fill });
}

// Retained DrawingML shape plans (DOCX) are authored in points, so their stroke
// widths are points too; line-end decorations need this for their 2 pt floor.
const PLAN_UNITS_PER_POINT = 1;

/** Acquire numeric paths and paint scheduling once. No Canvas or painter is called. */
export function resolveDrawingMLGeometry(plan: DrawingMLShapePaintPlan, scale: number, maximumWork = Infinity): ResolvedDrawingMLGeometry {
  let workUnits = 0;
  if (!(maximumWork === Infinity || (Number.isSafeInteger(maximumWork) && maximumWork > 0))) throw new GeometryWorkBudgetError('Invalid DrawingML geometry budget');
  const charge = (units: number) => {
    workUnits += units;
    if (!Number.isSafeInteger(workUnits) || workUnits > maximumWork) throw new GeometryWorkBudgetError('DrawingML geometry work budget exceeded');
  };
  const { x, y, w, h } = plan.rect;
  if (![scale, plan.rect.x, plan.rect.y, w, h, plan.transform.rotationDeg].every(Number.isFinite) || scale < 0 || w < 0 || h < 0) throw new RangeError('Invalid DrawingML geometry frame');
  charge(1);
  const strokeSource = plan.stroke as Stroke | null;
  charge(strokeSource?.customDash?.length ?? 0);
  if (plan.geometry.kind === 'preset' && plan.geometry.adjustments.some(a => a !== null && !Number.isFinite(a))) throw new RangeError('Non-finite DrawingML adjustment');
  const stroke = strokeSource ? resolveStrokeGeometry(strokeSource, scale) : null;
  const paths: ResolvedDrawingMLPath[] = [], arrows: ResolvedDrawingMLArrow[] = [];
  const add = (path: GeometryPath, fill: string | null, strokeGeometry: ResolvedStrokeGeometry | null, fillRule: 'nonzero' | 'evenodd' = 'nonzero') => {
    paths.push(Object.freeze({ path, fill, stroke: strokeGeometry, fillRule }));
  };
  const arrow = (tipX: number, tipY: number, angle: number, end: Readonly<ArrowEnd> | undefined) => {
    if (!end || !strokeSource) return;
    const geometry = resolveArrowGeometry(end, strokeSource, scale, PLAN_UNITS_PER_POINT, charge);
    if (!geometry) return;
    if (![tipX, tipY, angle].every(Number.isFinite)) throw new RangeError('Invalid resolved DrawingML arrow frame');
    arrows.push(Object.freeze({ tipX, tipY, angle, end: Object.freeze({ ...end }), geometry }));
  };
  // How far each decorated end pulls its line back (observed Office rule in
  // arrow-geometry.ts). Leaders that cannot be re-stroked from straight
  // vertices (curved connectors, custom paths) are trimmed along the path.
  const headTrim = strokeSource?.headEnd ? lineEndRetract(strokeSource.headEnd, strokeSource, scale, PLAN_UNITS_PER_POINT) : 0;
  const tailTrim = strokeSource?.tailEnd ? lineEndRetract(strokeSource.tailEnd, strokeSource, scale, PLAN_UNITS_PER_POINT) : 0;
  /** Add a stroked path trimmed at its ends; a filled path keeps its untrimmed fill. */
  const addTrimmed = (path: GeometryPath, fill: string | null, fillRule: 'nonzero' | 'evenodd', head: number, tail: number) => {
    if (!(head > 0 || tail > 0) || !stroke) { add(path, fill, stroke, fillRule); return; }
    const trimmed = trimGeometryPathEnds(path, head, tail);
    // The retained validator charges 1 + arguments per command of every path.
    charge(trimmed.reduce((units, command) => units + 1 + command.args.length, 0));
    if (fill !== 'none') add(path, fill, null, fillRule);
    add(trimmed, 'none', stroke, fillRule);
  };
  const silhouette = createGeometryPath(charge);
  if (plan.geometry.kind === 'preset') {
    charge(plan.geometry.adjustments.length);
    const name = plan.geometry.name.toLowerCase(), adjustments = [...plan.geometry.adjustments];
    const decorated = isRetractableLeaderPreset(name) && !!(strokeSource?.headEnd || strokeSource?.tailEnd);
    const trimsCurve = name.startsWith('curvedconnector');
    const presetPaths = resolvePresetGeometryPaths(name, w, h, adjustments, charge, x, y);
    if (presetPaths) {
      presetPaths.forEach((path, index) => {
        const trailing = index === presetPaths.length - 1 && (path.fill === 'none' || path.fill == null);
        if (trimsCurve && trailing && path.stroke) addTrimmed(path.path, path.fill, 'nonzero', headTrim, tailTrim);
        else add(path.path, path.fill, path.stroke && !(decorated && trailing) ? stroke : null);
      });
    } else {
      // Only names the spec engine does not define reach here: `rect` (implicit
      // box geometry) or a schema-invalid value (see buildShapePath).
      const path = createGeometryPath(charge);
      buildShapePath(path.sink, name, x, y, w, h, adjustments[0], adjustments[1], adjustments[2], adjustments[3]);
      add(Object.freeze(path.commands), null, stroke, 'nonzero');
    }
    // The clip/shade silhouette excludes the decorative `fill="none"` paths.
    if (!buildPresetGeometryFillPath(silhouette.sink, plan.geometry.name, x, y, w, h, adjustments)) {
      buildShapePath(silhouette.sink, plan.geometry.name, x, y, w, h, adjustments[0], adjustments[1], adjustments[2], adjustments[3]);
    }
    if (strokeSource && hasPresetLineEnds(name)) {
      const anchors = getConnectorAnchors(name, x, y, w, h, adjustments);
      if (anchors) {
        if (decorated && anchors.vertices.length >= 2) {
          // Trim across the whole leader, so a short terminal leg cannot stop
          // the retraction early; a leader consumed by both ends paints nothing.
          const path = createGeometryPath(charge), points = anchors.vertices;
          path.sink.moveTo(points[0].x, points[0].y);
          for (let i = 1; i < points.length; i++) path.sink.lineTo(points[i].x, points[i].y);
          const leader = trimGeometryPathEnds(Object.freeze(path.commands), headTrim, tailTrim);
          if (leader.length >= 2) add(leader, 'none', stroke);
        }
        arrow(anchors.end.x, anchors.end.y, anchors.end.angle, strokeSource.tailEnd);
        arrow(anchors.start.x, anchors.start.y, anchors.start.angle, strokeSource.headEnd);
      }
    }
  } else {
    const { subpaths } = plan.geometry;
    for (const path of subpaths) charge(path.length + 1);
    const paint = plan.geometry.paint?.length === subpaths.length ? plan.geometry.paint : null;
    const ends = strokeSource && (headTrim > 0 || tailTrim > 0) ? getCustGeomEndpoints(subpaths) : null;
    const customHead = ends?.start ? headTrim : 0, customTail = ends?.end ? tailTrim : 0;
    if (paint) {
      subpaths.forEach((subpath, index) => {
        const path = createGeometryPath(charge);
        buildCustomPath(path.sink, [subpath], x, y, w, h);
        const commands = Object.freeze(path.commands), fill = paint[index].fill ?? null;
        if (paint[index].stroke === false) add(commands, fill, null, paint[index].fillRule);
        else addTrimmed(commands, fill, paint[index].fillRule ?? 'nonzero',
          index === 0 ? customHead : 0, index === subpaths.length - 1 ? customTail : 0);
      });
    } else {
      const path = createGeometryPath(charge);
      buildCustomPath(path.sink, subpaths, x, y, w, h);
      addTrimmed(Object.freeze(path.commands), null, 'nonzero', customHead, customTail);
    }
    buildCustomPath(silhouette.sink, paint ? subpaths.filter((_, i) => paint[i].fill !== 'none') : subpaths, x, y, w, h);
    if (strokeSource) {
      const endpoints = getCustGeomEndpoints(subpaths);
      if (endpoints.start) arrow(x + endpoints.start.x * w, y + endpoints.start.y * h, Math.atan2(endpoints.start.dy * h, endpoints.start.dx * w), strokeSource.headEnd);
      if (endpoints.end) arrow(x + endpoints.end.x * w, y + endpoints.end.y * h, Math.atan2(endpoints.end.dy * h, endpoints.end.dx * w), strokeSource.tailEnd);
    }
  }
  return Object.freeze({ width: w, height: h, scale, originX: x, originY: y, sourceKey: sourceKey(plan), paths: Object.freeze(paths), arrows: Object.freeze(arrows), fillSilhouette: Object.freeze(silhouette.commands), workUnits });
}

/** Validate frame/source binding without reacquiring geometry. */
export function requireResolvedDrawingMLGeometry(plan: DrawingMLShapePaintPlan, scale: number): ResolvedDrawingMLGeometry {
  const geometry = plan.resolvedGeometry;
  if (![scale, plan.rect.x, plan.rect.y, plan.rect.w, plan.rect.h, plan.transform.rotationDeg].every(Number.isFinite)
    || scale < 0 || plan.rect.w < 0 || plan.rect.h < 0 || typeof plan.transform.flipH !== 'boolean' || typeof plan.transform.flipV !== 'boolean') throw new RangeError('Invalid retained DrawingML frame');
  if (!geometry || geometry.width !== plan.rect.w || geometry.height !== plan.rect.h || geometry.scale !== scale || geometry.sourceKey !== sourceKey(plan)) throw new RangeError('Missing or stale retained DrawingML geometry');
  if (![geometry.originX, geometry.originY].every(Number.isFinite)) throw new RangeError('Invalid retained DrawingML origin');
  if (!Number.isSafeInteger(geometry.workUnits) || geometry.workUnits < 1 || !Array.isArray(geometry.paths) || !Array.isArray(geometry.arrows)) throw new RangeError('Invalid retained DrawingML work/scheduling');
  if (geometry.paths.length + geometry.arrows.length > geometry.workUnits) throw new GeometryWorkBudgetError('Retained geometry scheduling exceeds its acquired work');
  let pathWork = 0;
  const charge = (units: number) => {
    pathWork += units;
    if (pathWork > geometry.workUnits) throw new GeometryWorkBudgetError('Retained geometry paths exceed their acquired work');
  };
  const stroke = (value: ResolvedStrokeGeometry) => {
    if (!Number.isFinite(value.lineWidth) || value.lineWidth < 0 || !Number.isFinite(value.miterLimit) || value.miterLimit <= 0
      || !['butt', 'round', 'square'].includes(value.lineCap) || !['round', 'bevel', 'miter'].includes(value.lineJoin)
      || !Array.isArray(value.dash) || value.dash.length > geometry.workUnits || !value.dash.every(v => Number.isFinite(v) && v >= 0)) throw new RangeError('Invalid retained stroke geometry');
  };
  for (const path of geometry.paths) {
    assertGeometryPath(path.path, charge);
    if (path.fill !== null && !['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess'].includes(path.fill)) throw new RangeError('Invalid retained path fill');
    if (path.fillRule !== 'nonzero' && path.fillRule !== 'evenodd') throw new RangeError('Invalid retained fill rule');
    if (path.stroke) stroke(path.stroke);
  }
  assertGeometryPath(geometry.fillSilhouette, charge);
  for (const arrow of geometry.arrows) {
    if (![arrow.tipX, arrow.tipY, arrow.angle].every(Number.isFinite) || !['fill', 'stroke'].includes(arrow.geometry.paint)) throw new RangeError('Invalid retained arrow frame');
    assertGeometryPath(arrow.geometry.path, charge); stroke(arrow.geometry.stroke);
  }
  return geometry;
}
function composeGeometryMatrix(a: GeometryMatrix, b: GeometryMatrix): GeometryMatrix {
  const m = { a: a.a * b.a + a.c * b.b, b: a.b * b.a + a.d * b.b, c: a.a * b.c + a.c * b.d,
    d: a.b * b.c + a.d * b.d, e: a.a * b.e + a.c * b.f + a.e, f: a.b * b.e + a.d * b.f + a.f };
  if (!Object.values(m).every(Number.isFinite) || !Number.isFinite(m.a * m.d - m.b * m.c) || m.a * m.d - m.b * m.c === 0) throw new RangeError('Invalid DrawingML geometry transform');
  return m;
}
const identity: GeometryMatrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
const translation = (e: number, f: number): GeometryMatrix => ({ ...identity, e, f });
const rotation = (angle: number): GeometryMatrix => ({ a: Math.cos(angle), b: Math.sin(angle), c: -Math.sin(angle), d: Math.cos(angle), e: 0, f: 0 });
/** Read the shared numeric geometry under an explicit owner transform. */
export function drawingMLGeometryBounds(plan: DrawingMLShapePaintPlan, owner: GeometryMatrix = identity): GeometryBox | null {
  const geometry = requireResolvedDrawingMLGeometry(plan, 1);
  const { x, y, w, h } = plan.rect, transform = plan.transform;
  let matrix = owner;
  if (transform.rotationDeg !== 0 || transform.flipH || transform.flipV) {
    matrix = composeGeometryMatrix(matrix, translation(x + w / 2, y + h / 2));
    if (transform.rotationDeg !== 0) matrix = composeGeometryMatrix(matrix, rotation(transform.rotationDeg * Math.PI / 180));
    matrix = composeGeometryMatrix(matrix, { ...identity, a: transform.flipH ? -1 : 1, d: transform.flipV ? -1 : 1 });
    matrix = composeGeometryMatrix(matrix, translation(-(x + w / 2), -(y + h / 2)));
  }
  matrix = composeGeometryMatrix(matrix, translation(x - geometry.originX, y - geometry.originY));
  const boxes: GeometryBox[] = [];
  const fillStroke = resolveStrokeGeometry(null, 1);
  const include = (path: GeometryPath, m: GeometryMatrix, stroke: ResolvedStrokeGeometry) => {
    const box = geometryPathBounds(path, m, stroke); if (box) boxes.push(box);
  };
  for (const path of geometry.paths) {
    if (plan.fill && plan.fill.fillType !== 'none' && path.fill !== 'none') include(path.path, matrix, fillStroke);
    if (path.stroke) include(path.path, matrix, path.stroke);
  }
  for (const arrow of geometry.arrows) {
    const m = composeGeometryMatrix(composeGeometryMatrix(matrix, translation(arrow.tipX, arrow.tipY)), rotation(arrow.angle));
    include(arrow.geometry.path, m, arrow.geometry.paint === 'stroke' ? arrow.geometry.stroke : fillStroke);
  }
  if (boxes.length === 0) return null;
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  for (const b of boxes) { left = Math.min(left, b.x); top = Math.min(top, b.y); right = Math.max(right, b.x + b.w); bottom = Math.max(bottom, b.y + b.h); }
  const result = { x: left, y: top, w: right - left, h: bottom - top };
  if (!Object.values(result).every(Number.isFinite)) throw new RangeError('Non-finite DrawingML enclosure');
  return result;
}
