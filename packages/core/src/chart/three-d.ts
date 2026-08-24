import type { ChartThreeD } from '../types/chart.js';

export interface ThreeDRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ThreeDPoint {
  x: number;
  y: number;
}

export interface ThreeDScenePoint extends ThreeDPoint {
  depth: number;
}

export interface ThreeDCameraNormal {
  x: number;
  y: number;
  z: number;
}

export interface ThreeDBarClusterSlot {
  /** Offset from the start of one category interval. */
  offset: number;
  /** Marker width on the category axis. */
  size: number;
}

/**
 * Resolve one bar/column marker inside a 3-D category cluster.
 *
 * ECMA-376 §21.2.2.74/.75 define gapDepth and gapWidth around bar or
 * column *clusters*. Excel therefore places ordinary series beside each other
 * on the category axis and extrudes the complete cluster through one shared
 * depth interval. Stacked series reuse the complete category-axis footprint.
 */
export function planThreeDBarClusterSlot(
  categoryInterval: number,
  gapWidthPercent: number,
  seriesIndex: number,
  seriesCount: number,
  stacked: boolean,
): ThreeDBarClusterSlot {
  const interval = Number.isFinite(categoryInterval) && categoryInterval > 0
    ? categoryInterval : 0;
  const gap = Number.isFinite(gapWidthPercent)
    ? clamp(gapWidthPercent, 0, 500) : 150;
  const count = Math.max(1, Math.trunc(seriesCount));
  const index = clamp(Math.trunc(seriesIndex), 0, count - 1);
  // gapWidth is expressed as a percentage of one marker width, not of the
  // complete multi-series group. If marker width is M, one category interval
  // is `seriesCount * M + gapWidth * M` for clustered data, or
  // `M + gapWidth * M` for a stack.
  const markerCount = stacked ? 1 : count;
  const size = interval / (markerCount + gap / 100);
  const group = size * markerCount;
  return {
    offset: (interval - group) / 2 + (stacked ? 0 : index * size),
    size,
  };
}

/** Ring scale for `coneToMax` / `pyramidToMax`. Both ends of a stacked
 * segment use the same axis-coordinate function, so neighbouring segments
 * share one ring instead of restarting from a full-width base. */
export function threeDToMaxScale(
  coordinate: number,
  axisMin: number,
  axisMax: number,
): number {
  if (![coordinate, axisMin, axisMax].every(Number.isFinite)) return 1;
  const bound = coordinate >= 0 ? axisMax : axisMin;
  return 1 - Math.min(
    1,
    Math.abs(coordinate) / Math.max(Number.MIN_VALUE, Math.abs(bound)),
  );
}

/** Resolve pie3D `<c:view3D><c:hPercent>` as the Office-defined multiplier
 * of the radial family's default wall thickness (MS-OE376 §2.1.1501(b)).
 * Omission keeps automatic/default thickness. Public models created before
 * authored-provenance was exposed continue to treat a supplied value as
 * authored. */
export function pieThreeDThicknessMultiplier(view: ChartThreeD): number {
  const authored = view.heightPercentAuthored ?? view.heightPercent != null;
  if (!authored || view.heightPercent == null || !Number.isFinite(view.heightPercent)) return 1;
  if (view.heightPercent < 5 || view.heightPercent > 500) return 1;
  return view.heightPercent / 100;
}

export interface ChartThreeDSceneTopology {
  farX: 'min' | 'max';
  farY: 'min' | 'max';
  axisX: 'min' | 'max';
  axisY: 'min' | 'max';
  nearDepth: 0 | 1;
  farDepth: 0 | 1;
}

export interface ChartThreeDProjection {
  /** Authored 3-D scene box fitted into the available plot rectangle. */
  scene: ThreeDRect;
  /** Front plotting plane after reserving the projected depth vector. */
  front: ThreeDRect;
  /** Full front-to-back offset in CSS pixels. */
  depthX: number;
  depthY: number;
  /** Model-space scene depth before camera projection/refit. */
  modelDepth: number;
  /** Excel-compatible compact vertical radius for 3-D pie tops. */
  pieScaleY: number;
  /** Bounded visible pie-wall thickness as a fraction of the horizontal radius. */
  pieThicknessFraction: number;
  project: (x: number, y: number, depth: number) => ThreeDPoint;
  /** Project model geometry outside the base 0..1 depth interval. Used only
   * for authored CT_Surface thickness; ordinary data geometry remains clipped
   * to the chart volume through {@link project}. */
  projectUnbounded: (x: number, y: number, depth: number) => ThreeDPoint;
  /** Camera-space depth used by the shared painter's far-to-near scene sort. */
  cameraDepth: (x: number, y: number, depth: number) => number;
  /** Reciprocal clip-space W used for perspective-correct interpolation of
   * camera depth across one projected face. Equals 1 in the affine limit. */
  cameraProjectionWeight: (x: number, y: number, depth: number) => number;
  /** True when an outward-wound scene face is visible to the camera. */
  cameraFacing: (points: readonly ThreeDScenePoint[]) => boolean;
  /** Unit outward normal transformed into camera space. Material lighting is
   * deliberately a later renderer stage and never changes this geometry. */
  cameraNormal: (points: readonly ThreeDScenePoint[]) => ThreeDCameraNormal | null;
  /** Far wall planes and near axis plane selected from the authored camera. */
  topology: ChartThreeDSceneTopology;
  seriesDepth: (seriesIndex: number, seriesCount: number, stacked?: boolean) => number;
  prismDepth: (seriesCount: number) => number;
  prismInterval: (
    seriesIndex: number,
    seriesCount: number,
    stacked?: boolean,
  ) => { near: number; far: number };
}

export interface ChartThreeDProjectionOptions {
  /**
   * Model-space chart depth as a fraction of chart width at depthPercent=100.
   * The camera remains shared; this only describes how much of the model's Z
   * axis the chart family occupies before projection.
   */
  sceneDepthScale?: number;
  /**
   * Compatibility multiplier applied to tan(FOV / 2). The default preserves
   * the Office line/area boundary corpus; callers with an Office view whose
   * authored FOV is observed directly can opt into the normative gain of 1.
   */
  perspectiveTangentGain?: number;
  /** Model-space scene height as a fraction of scene width when the chart has
   * no authored hPercent. Callers use an Office-observed family boundary
   * instead of fitting an otherwise empty full-height cartesian cuboid. */
  sceneHeightScale?: number;
}

function fitProjectedGeometry(
  projection: ChartThreeDProjection,
  projected: readonly ThreeDPoint[],
  target: ThreeDRect,
  paddingFraction: number,
): ChartThreeDProjection {
  if (!projected.length || projected.length > 100_000
    || ![target.x, target.y, target.w, target.h].every(Number.isFinite)
    || target.w <= 0 || target.h <= 0) return projection;
  if (!projected.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) {
    return projection;
  }
  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (const point of projected) {
    minX = Math.min(minX, point.x);
    maxX = Math.max(maxX, point.x);
    minY = Math.min(minY, point.y);
    maxY = Math.max(maxY, point.y);
  }
  const width = maxX - minX;
  const height = maxY - minY;
  if (!(width > Number.EPSILON) || !(height > Number.EPSILON)) return projection;
  const padding = clamp(finiteOr(paddingFraction, 0.06), 0, 0.45);
  const availableW = target.w * (1 - 2 * padding);
  const availableH = target.h * (1 - 2 * padding);
  const scale = Math.min(availableW / width, availableH / height);
  if (!(scale > 0) || !Number.isFinite(scale)) return projection;
  const sourceCenter = { x: (minX + maxX) / 2, y: (minY + maxY) / 2 };
  const targetCenter = { x: target.x + target.w / 2, y: target.y + target.h / 2 };
  const transformProject = (
    baseProject: ChartThreeDProjection['project'],
  ): ChartThreeDProjection['project'] => (x, y, depth) => {
    const point = baseProject(x, y, depth);
    return {
      x: targetCenter.x + (point.x - sourceCenter.x) * scale,
      y: targetCenter.y + (point.y - sourceCenter.y) * scale,
    };
  };
  return {
    ...projection,
    project: transformProject(projection.project),
    projectUnbounded: transformProject(projection.projectUnbounded),
    depthX: projection.depthX * scale,
    depthY: projection.depthY * scale,
  };
}

/** Reframe an existing homogeneous camera around the geometry actually used
 * by one chart. This is a final uniform viewport transform only: camera-space
 * depth, culling, straight lines and vanishing points remain unchanged. */
export function fitChartThreeDProjectionToPoints(
  projection: ChartThreeDProjection,
  points: readonly ThreeDScenePoint[],
  target: ThreeDRect,
  paddingFraction = 0.06,
): ChartThreeDProjection {
  if (!points.length || points.length > 100_000) return projection;
  return fitProjectedGeometry(
    projection,
    points.map(point => projection.project(point.x, point.y, point.depth)),
    target,
    paddingFraction,
  );
}

export type ChartThreeDSurfaceKind = 'floor' | 'sideWall' | 'backWall';

export interface ChartThreeDSurfaceGeometry {
  /** Authored thickness in model-space chart units. */
  thickness: number;
  /** Plot-volume boundary face. */
  inner: ThreeDScenePoint[];
  /** Parallel exterior face. Equals `inner` when thickness is zero/invalid. */
  outer: ThreeDScenePoint[];
  /** One planar face at zero thickness, otherwise the closed six-face slab. */
  faces: ThreeDScenePoint[][];
  /** Projected plot-face width/height used by Office plain picture stacking. */
  pictureStackAspect: number | null;
  /** CSS-pixel metric represented by one normalized scene-depth unit. */
  modelDepth: number;
}

export interface ChartThreeDSurfaceGridSegment {
  /** Index into {@link ChartThreeDSurfaceGeometry.faces}. */
  faceIndex: number;
  scenePoints: [ThreeDScenePoint, ThreeDScenePoint];
}

const MAX_UNSIGNED_INT = 4_294_967_295;

/** Resolve one CT_Surface as a bounded slab outside the plot volume.
 * ECMA-376 §21.2.2.206 defines thickness as a percentage of the largest
 * dimension of the plot volume. No screen-space offsets or family constants
 * are involved: all six faces pass through the shared camera. */
export function planChartThreeDSurfaceGeometry(
  projection: ChartThreeDProjection,
  kind: ChartThreeDSurfaceKind,
  thicknessPercent: number | null | undefined,
): ChartThreeDSurfaceGeometry {
  const { front } = projection;
  const xMin = front.x;
  const xMax = front.x + front.w;
  const sideX = projection.topology.farX === 'min' ? xMin : xMax;
  const floorY = projection.topology.axisY === 'min' ? front.y : front.y + front.h;
  const topY = floorY === front.y ? front.y + front.h : front.y;
  const { nearDepth, farDepth } = projection.topology;
  const stackReference = [
    projection.projectUnbounded(xMin, topY, farDepth),
    projection.projectUnbounded(xMax, topY, farDepth),
    projection.projectUnbounded(xMin, floorY, farDepth),
  ];
  // Office sizes one plain-stack repetition from the complete projected plot
  // volume width, then anchors that repetition on the selected floor/wall
  // face. The near depth edge can extend beyond the back-wall edge under
  // perspective, so measuring only the far face makes the repeated picture
  // too short and exposes source rows that Office clips above the wall.
  const projectedVolumeX = [xMin, xMax].flatMap(x =>
    [topY, floorY].flatMap(y =>
      [nearDepth, farDepth].map(depth => projection.projectUnbounded(x, y, depth).x)
    )
  );
  const stackReferenceWidth = Math.max(...projectedVolumeX) - Math.min(...projectedVolumeX);
  const stackReferenceHeight = Math.hypot(
    stackReference[0].x - stackReference[2].x,
    stackReference[0].y - stackReference[2].y,
  );
  const pictureStackAspect = stackReferenceWidth > 0 && stackReferenceHeight > 0
    ? stackReferenceWidth / stackReferenceHeight
    : null;
  const rawPercent = thicknessPercent == null ? 0 : thicknessPercent;
  const validPercent = Number.isFinite(rawPercent)
    && rawPercent >= 0 && rawPercent <= MAX_UNSIGNED_INT ? rawPercent : 0;
  const thickness = Math.max(front.w, front.h, projection.modelDepth) * validPercent / 100;
  let inner: ThreeDScenePoint[];
  let outer: ThreeDScenePoint[];
  if (kind === 'floor') {
    inner = [
      { x: xMin, y: floorY, depth: nearDepth },
      { x: xMax, y: floorY, depth: nearDepth },
      { x: xMax, y: floorY, depth: farDepth },
      { x: xMin, y: floorY, depth: farDepth },
    ];
    const outerY = floorY + (floorY === front.y ? -thickness : thickness);
    outer = inner.map(point => ({ ...point, y: outerY }));
  } else if (kind === 'sideWall') {
    inner = [
      { x: sideX, y: floorY, depth: nearDepth },
      { x: sideX, y: floorY, depth: farDepth },
      { x: sideX, y: topY, depth: farDepth },
      { x: sideX, y: topY, depth: nearDepth },
    ];
    const outerX = sideX + (sideX === xMin ? -thickness : thickness);
    outer = inner.map(point => ({ ...point, x: outerX }));
  } else {
    inner = [
      { x: xMin, y: floorY, depth: farDepth },
      { x: xMax, y: floorY, depth: farDepth },
      { x: xMax, y: topY, depth: farDepth },
      { x: xMin, y: topY, depth: farDepth },
    ];
    const depthOffset = projection.modelDepth > 0 ? thickness / projection.modelDepth : 0;
    const outerDepth = farDepth === 0 ? -depthOffset : 1 + depthOffset;
    outer = inner.map(point => ({ ...point, depth: outerDepth }));
  }
  if (!(thickness > 0)) {
    return {
      thickness: 0,
      inner,
      outer: [...inner],
      faces: [inner],
      pictureStackAspect,
      modelDepth: projection.modelDepth,
    };
  }
  const sides = inner.map((point, index) => [
    point,
    inner[(index + 1) % inner.length],
    outer[(index + 1) % outer.length],
    outer[index],
  ]);
  const rawFaces = [inner, outer, ...sides];
  const solidCenter = [...inner, ...outer].reduce(
    (sum, point) => ({
      x: sum.x + point.x / 8,
      y: sum.y + point.y / 8,
      depth: sum.depth + point.depth / 8,
    }),
    { x: 0, y: 0, depth: 0 },
  );
  const faces = rawFaces.map(face => {
    const [a, b, c] = face;
    const ab = { x: b.x - a.x, y: b.y - a.y, depth: b.depth - a.depth };
    const ac = { x: c.x - a.x, y: c.y - a.y, depth: c.depth - a.depth };
    const normal = {
      x: ab.y * ac.depth - ab.depth * ac.y,
      y: ab.depth * ac.x - ab.x * ac.depth,
      depth: ab.x * ac.y - ab.y * ac.x,
    };
    const centroid = face.reduce(
      (sum, point) => ({
        x: sum.x + point.x / face.length,
        y: sum.y + point.y / face.length,
        depth: sum.depth + point.depth / face.length,
      }),
      { x: 0, y: 0, depth: 0 },
    );
    const outward = {
      x: centroid.x - solidCenter.x,
      y: centroid.y - solidCenter.y,
      depth: centroid.depth - solidCenter.depth,
    };
    const dot = normal.x * outward.x + normal.y * outward.y + normal.depth * outward.depth;
    return dot < 0 ? [...face].reverse() : face;
  });
  return { thickness, inner, outer, faces, pictureStackAspect, modelDepth: projection.modelDepth };
}

/** Continue one grid rule over the planar face or every corresponding face of
 * a positive-thickness CT_Surface slab. The input fraction is expressed in
 * the owning plot-face x or y direction, so no screen-space fitting or camera
 * heuristic is involved. */
export function planChartThreeDSurfaceGridSegments(
  geometry: ChartThreeDSurfaceGeometry,
  kind: ChartThreeDSurfaceKind,
  coordinate: 'x' | 'y',
  fraction: number,
): ChartThreeDSurfaceGridSegment[] {
  if (!Number.isFinite(fraction) || fraction < 0 || fraction > 1) return [];
  if ((coordinate === 'x' && kind === 'sideWall')
    || (coordinate === 'y' && kind === 'floor')) return [];
  const interpolate = (start: ThreeDScenePoint, end: ThreeDScenePoint): ThreeDScenePoint => ({
    x: start.x + (end.x - start.x) * fraction,
    y: start.y + (end.y - start.y) * fraction,
    depth: start.depth + (end.depth - start.depth) * fraction,
  });
  let innerStart: ThreeDScenePoint;
  let innerEnd: ThreeDScenePoint;
  let outerStart: ThreeDScenePoint;
  let outerEnd: ThreeDScenePoint;
  let startFaceIndex: number;
  let endFaceIndex: number;
  if (coordinate === 'x') {
    innerStart = interpolate(geometry.inner[0], geometry.inner[1]);
    innerEnd = interpolate(geometry.inner[3], geometry.inner[2]);
    outerStart = interpolate(geometry.outer[0], geometry.outer[1]);
    outerEnd = interpolate(geometry.outer[3], geometry.outer[2]);
    startFaceIndex = 2;
    endFaceIndex = 4;
  } else {
    innerStart = interpolate(geometry.inner[0], geometry.inner[3]);
    innerEnd = interpolate(geometry.inner[1], geometry.inner[2]);
    outerStart = interpolate(geometry.outer[0], geometry.outer[3]);
    outerEnd = interpolate(geometry.outer[1], geometry.outer[2]);
    startFaceIndex = 5;
    endFaceIndex = 3;
  }
  const segments: ChartThreeDSurfaceGridSegment[] = [{
    faceIndex: 0,
    scenePoints: [innerStart, innerEnd],
  }];
  if (geometry.thickness > 0) {
    segments.push(
      { faceIndex: 1, scenePoints: [outerStart, outerEnd] },
      { faceIndex: startFaceIndex, scenePoints: [innerStart, outerStart] },
      { faceIndex: endFaceIndex, scenePoints: [innerEnd, outerEnd] },
    );
  }
  return segments;
}

/** Refit the complete base cuboid and the three authored surface slabs into
 * the existing plot rectangle with the same 3% scene margin used by the base
 * camera. This is one uniform viewport transform, so data, axes and surfaces
 * keep a coherent projection. */
export function fitChartThreeDProjectionToWallThickness(
  projection: ChartThreeDProjection,
  view: Pick<ChartThreeD, 'floor' | 'sideWall' | 'backWall'>,
  target: ThreeDRect,
): ChartThreeDProjection {
  const specs: Array<[ChartThreeDSurfaceKind, number | null | undefined]> = [
    ['floor', view.floor?.thicknessPercent],
    ['sideWall', view.sideWall?.thicknessPercent],
    ['backWall', view.backWall?.thicknessPercent],
  ];
  const geometries = specs.map(([kind, thickness]) =>
    planChartThreeDSurfaceGeometry(projection, kind, thickness)
  );
  if (!geometries.some(geometry => geometry.thickness > 0)) return projection;
  const points = geometries.flatMap(geometry => geometry.faces.flat());
  return fitProjectedGeometry(
    projection,
    points.map(point => projection.projectUnbounded(point.x, point.y, point.depth)),
    target,
    0.03,
  );
}

const finiteOr = (value: number | null | undefined, fallback: number): number =>
  typeof value === 'number' && Number.isFinite(value) ? value : fallback;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

/**
 * Resolve the application-generated 3-D view into one deterministic, bounded
 * homogeneous camera plan shared by bar/column, line and area painters.
 *
 * ECMA-376 defines the authored `view3D` fields and their schema bounds but not
 * Excel's raster projection.  The omitted baseline (rotY=20, rotX=15,
 * depth=100, perspective=30, gapDepth=150) and the depth/rotation responses are
 * repeated observations from the local boundary corpus. The implementation is
 * intentionally a small scene transform, but every cartesian primitive passes
 * through that one transform so straight lines, common planes and vanishing
 * points remain geometrically coherent.
 */
export function planChartThreeDProjection(
  view: ChartThreeD,
  plot: ThreeDRect,
  options: ChartThreeDProjectionOptions = {},
): ChartThreeDProjection | null {
  if (![plot.x, plot.y, plot.w, plot.h].every(Number.isFinite) || plot.w <= 0 || plot.h <= 0) {
    return null;
  }
  const rotationX = clamp(finiteOr(view.rotationX, 15), -90, 90);
  const rotationYRaw = clamp(finiteOr(view.rotationY, 20), 0, 360);
  const rotationY = ((rotationYRaw + 180) % 360) - 180;
  const depthPercent = clamp(finiteOr(view.depthPercent, 100), 20, 2000);
  const perspective = clamp(finiteOr(view.perspective, 30), 0, 240);
  const gapDepth = clamp(finiteOr(view.gapDepthPercent, 150), 0, 500);
  const heightPercentAuthored = view.heightPercentAuthored ?? view.heightPercent != null;
  const authoredHeightPercent = heightPercentAuthored && view.heightPercent != null
    && Number.isFinite(view.heightPercent)
    ? clamp(view.heightPercent, 5, 500)
    : null;
  const inferredHeightPercent = options.sceneHeightScale != null
    && Number.isFinite(options.sceneHeightScale)
    ? clamp(options.sceneHeightScale * 100, 5, 500)
    : null;
  const sceneHeightPercent = authoredHeightPercent ?? inferredHeightPercent;
  let scene = plot;
  if (sceneHeightPercent != null) {
    // ECMA-376 §21.2.2.83 defines hPercent as the 3-D chart height relative to
    // its width. Fit that authored scene box inside the renderer's available
    // plot without changing the ratio.
    const ratio = sceneHeightPercent / 100;
    const sceneWidth = Math.min(plot.w, plot.h / ratio);
    const sceneHeight = sceneWidth * ratio;
    scene = {
      x: plot.x + (plot.w - sceneWidth) / 2,
      y: plot.y + (plot.h - sceneHeight) / 2,
      w: sceneWidth,
      h: sceneHeight,
    };
  }
  const radians = Math.PI / 180;
  // ECMA-376 §21.2.2.41 defines depthPercent relative to chart width. Preserve
  // its linear response in model space; the complete projected box is fitted
  // afterward, so even the 2000% schema boundary remains finite and visible.
  // Office uses one camera but different Z occupancy for clustered prisms and
  // line/area series planes. Vector boundary observations give approximately
  // 10% of chart width for bar/column and 40% for line/area at depth=100.
  // Keeping that distinction in model space avoids chart-family angle hacks:
  // every wall, axis and data primitive still passes through this one camera.
  const sceneDepthScale = clamp(finiteOr(options.sceneDepthScale, 0.10), 0.01, 2);
  const depthMagnitude = scene.w * sceneDepthScale * (depthPercent / 100);
  const centreX = scene.x + scene.w / 2;
  const centreY = scene.y + scene.h / 2;
  const yaw = -rotationY * radians;
  const pitch = rotationX * radians;
  const cosYaw = Math.cos(yaw);
  const sinYaw = Math.sin(yaw);
  const cosPitch = Math.cos(pitch);
  const sinPitch = Math.sin(pitch);
  const perspectiveEnabled = view.rightAngleAxes !== true && perspective > 0;
  // ECMA-376 §21.2.2.136 stores the full field-of-view in half-degree units,
  // hence the normative pinhole half-angle is value * 0.25°. Existing Office
  // line/area vectors require a stronger fitted response, while measured
  // standard Bar views use the normative response. Keep this compatibility
  // choice explicit and family-scoped instead of changing the authored angle.
  // atan() keeps the complete 0..240 schema range below 90°.
  const normativeHalfAngle = clamp(perspective * 0.25, 0.25, 60) * radians;
  const perspectiveTangentGain = clamp(
    finiteOr(options.perspectiveTangentGain, 2), 0.25, 4,
  );
  const perspectiveHalfAngle = Math.atan(
    perspectiveTangentGain * Math.tan(normativeHalfAngle),
  );
  const sceneDiagonal = Math.hypot(scene.w, scene.h, depthMagnitude);
  const requestedCameraDistance = perspectiveEnabled
    ? sceneDiagonal * 0.5 / Math.tan(perspectiveHalfAngle)
    : Number.POSITIVE_INFINITY;

  const cameraPoint = (x: number, y: number, depth: number, clampDepth = true) => {
    const worldX = x - centreX;
    const worldY = centreY - y;
    // Increasing chart depth moves away from the viewer.
    const finiteDepth = Number.isFinite(depth) ? depth : 0;
    const worldZ = (0.5 - (clampDepth ? clamp(finiteDepth, 0, 1) : finiteDepth))
      * depthMagnitude;
    const yawX = cosYaw * worldX + sinYaw * worldZ;
    const yawZ = -sinYaw * worldX + cosYaw * worldZ;
    return {
      x: yawX,
      y: cosPitch * worldY - sinPitch * yawZ,
      z: sinPitch * worldY + cosPitch * yawZ,
    };
  };
  const cameraFaceNormal = (points: readonly ThreeDScenePoint[]) => {
    if (points.length < 3) return null;
    const cameraPoints = points.map(point => cameraPoint(point.x, point.y, point.depth, false));
    const a = cameraPoints[0];
    let unitNormal: ThreeDCameraNormal | null = null;
    // A clipped/sign-crossing solid can collapse one end of a quad to a
    // triangle while retaining its four-index topology. Do not classify that
    // valid face as edge-on merely because the first three stored vertices
    // contain the duplicated crossing point; find the first non-collinear fan
    // triangle, matching mesh winding normalization.
    for (let first = 1; first + 1 < cameraPoints.length && !unitNormal; first++) {
      for (let second = first + 1; second < cameraPoints.length; second++) {
        const b = cameraPoints[first];
        const c = cameraPoints[second];
        const ab = { x: b.x - a.x, y: b.y - a.y, z: b.z - a.z };
        const ac = { x: c.x - a.x, y: c.y - a.y, z: c.z - a.z };
        const normal = {
          x: ab.y * ac.z - ab.z * ac.y,
          y: ab.z * ac.x - ab.x * ac.z,
          z: ab.x * ac.y - ab.y * ac.x,
        };
        const length = Math.hypot(normal.x, normal.y, normal.z);
        if (length > Number.EPSILON) {
          unitNormal = {
            x: normal.x / length,
            y: normal.y / length,
            z: normal.z / length,
          };
          break;
        }
      }
    }
    if (!unitNormal) return null;
    return {
      normal: unitNormal,
      centroid: cameraPoints.reduce(
        (sum, point) => ({
          x: sum.x + point.x / cameraPoints.length,
          y: sum.y + point.y / cameraPoints.length,
          z: sum.z + point.z / cameraPoints.length,
        }),
        { x: 0, y: 0, z: 0 },
      ),
    };
  };
  let maxCameraZ = Number.NEGATIVE_INFINITY;
  for (const x of [scene.x, scene.x + scene.w]) {
    for (const y of [scene.y, scene.y + scene.h]) {
      for (const depth of [0, 1]) {
        maxCameraZ = Math.max(maxCameraZ, cameraPoint(x, y, depth).z);
      }
    }
  }
  const cameraDistance = perspectiveEnabled
    ? Math.max(requestedCameraDistance, maxCameraZ + sceneDiagonal * 0.01)
    : Number.POSITIVE_INFINITY;
  const rawProject = (
    x: number,
    y: number,
    depth: number,
    clampDepth = true,
  ): ThreeDPoint => {
    const camera = cameraPoint(x, y, depth, clampDepth);
    if (!perspectiveEnabled) return { x: camera.x, y: -camera.y };
    const denominator = Math.max(cameraDistance * 1e-9, cameraDistance - camera.z);
    const scale = cameraDistance / denominator;
    return { x: camera.x * scale, y: -camera.y * scale };
  };

  const rawCorners: ThreeDPoint[] = [];
  for (const x of [scene.x, scene.x + scene.w]) {
    for (const y of [scene.y, scene.y + scene.h]) {
      for (const depth of [0, 1]) rawCorners.push(rawProject(x, y, depth));
    }
  }
  const rawMinX = Math.min(...rawCorners.map(point => point.x));
  const rawMaxX = Math.max(...rawCorners.map(point => point.x));
  const rawMinY = Math.min(...rawCorners.map(point => point.y));
  const rawMaxY = Math.max(...rawCorners.map(point => point.y));
  const rawWidth = Math.max(Number.MIN_VALUE, rawMaxX - rawMinX);
  const rawHeight = Math.max(Number.MIN_VALUE, rawMaxY - rawMinY);
  const fitScale = Math.min(plot.w / rawWidth, plot.h / rawHeight) * 0.94;
  const fitOffsetX = plot.x + (plot.w - rawWidth * fitScale) / 2 - rawMinX * fitScale;
  const fitOffsetY = plot.y + (plot.h - rawHeight * fitScale) / 2 - rawMinY * fitScale;
  const project = (x: number, y: number, depth: number): ThreeDPoint => {
    const raw = rawProject(x, y, depth);
    return {
      x: fitOffsetX + raw.x * fitScale,
      y: fitOffsetY + raw.y * fitScale,
    };
  };
  const projectUnbounded = (x: number, y: number, depth: number): ThreeDPoint => {
    const raw = rawProject(x, y, depth, false);
    return {
      x: fitOffsetX + raw.x * fitScale,
      y: fitOffsetY + raw.y * fitScale,
    };
  };
  // `front` is the logical z=0 data plane. Its visual position is obtained
  // only through project(); no Canvas-horizontal surrogate plane exists.
  const front: ThreeDRect = { ...scene };
  const depthNear = project(centreX, centreY, 0);
  const depthFar = project(centreX, centreY, 1);
  const depthX = depthFar.x - depthNear.x;
  const depthY = depthFar.y - depthNear.y;
  const planeDepth = (
    axis: 'x' | 'y' | 'depth',
    endpoint: 'min' | 'max',
  ): number => {
    const x = axis === 'x' ? (endpoint === 'min' ? scene.x : scene.x + scene.w) : centreX;
    const y = axis === 'y' ? (endpoint === 'min' ? scene.y : scene.y + scene.h) : centreY;
    const depth = axis === 'depth' ? (endpoint === 'min' ? 0 : 1) : 0.5;
    return cameraPoint(x, y, depth).z;
  };
  const farX: 'min' | 'max' = planeDepth('x', 'min') <= planeDepth('x', 'max') ? 'min' : 'max';
  const farY: 'min' | 'max' = planeDepth('y', 'min') <= planeDepth('y', 'max') ? 'min' : 'max';
  const nearDepth: 0 | 1 = planeDepth('depth', 'min') >= planeDepth('depth', 'max') ? 0 : 1;
  const farDepth: 0 | 1 = nearDepth === 0 ? 1 : 0;
  const verticalEdgeMeanX = (endpoint: 'min' | 'max'): number => {
    const x = endpoint === 'min' ? scene.x : scene.x + scene.w;
    const top = project(x, scene.y, nearDepth);
    const bottom = project(x, scene.y + scene.h, nearDepth);
    return (top.x + bottom.x) / 2;
  };
  const horizontalEdgeMeanY = (endpoint: 'min' | 'max'): number => {
    const y = endpoint === 'min' ? scene.y : scene.y + scene.h;
    const left = project(scene.x, y, nearDepth);
    const right = project(scene.x + scene.w, y, nearDepth);
    return (left.y + right.y) / 2;
  };
  const axisX: 'min' | 'max' = verticalEdgeMeanX('min') <= verticalEdgeMeanX('max') ? 'min' : 'max';
  const axisY: 'min' | 'max' = horizontalEdgeMeanY('min') >= horizontalEdgeMeanY('max') ? 'min' : 'max';
  const prismDepth = (seriesCount: number): number => {
    const count = Math.max(1, Math.trunc(seriesCount));
    // ECMA-376 §21.2.2.74 gapDepth is the gap between adjacent 3-D series as
    // a percentage of marker depth. Each series owns one equal depth slot;
    // dividing the slot by (1 + gap ratio) preserves that authored ratio and
    // leaves the remainder as the inter-series/scene-edge spacing.
    return 1 / count / (1 + gapDepth / 100);
  };
  const seriesDepth = (seriesIndex: number, seriesCount: number, stacked = false): number => {
    if (stacked || seriesCount <= 1) return 0.5;
    const index = clamp(Math.trunc(seriesIndex), 0, Math.max(0, seriesCount - 1));
    return (index + 0.5) / seriesCount;
  };
  return {
    scene,
    front,
    depthX,
    depthY,
    modelDepth: depthMagnitude,
    // Existing Office vectors give a top-ellipse ratio of about .21 at 15°
    // and approximately 1 at 89°. A slight power curve captures both observed
    // boundaries without coupling radial charts to cartesian scene depth.
    pieScaleY: clamp(
      Math.pow(Math.sin(Math.max(1, Math.abs(rotationX)) * radians), 1.15),
      0.20,
      1,
    ),
    // The 3-D pie depth=100/2000 references are pixel-identical: cartesian
    // depthPercent/gapDepth do not control the pie wall. At the default 15°
    // elevation the wall is about .29r, then diminishes to zero as the top
    // becomes face-on near 90°. Keep this observed radial-family rule separate
    // from the cartesian scene depth.
    pieThicknessFraction: 0.30 * Math.max(0, Math.cos(Math.abs(rotationX) * radians)),
    project,
    projectUnbounded,
    cameraDepth(x, y, depth) {
      return cameraPoint(x, y, depth, false).z;
    },
    cameraProjectionWeight(x, y, depth) {
      if (!perspectiveEnabled) return 1;
      const z = cameraPoint(x, y, depth, false).z;
      return 1 / Math.max(cameraDistance * 1e-9, cameraDistance - z);
    },
    cameraFacing(points) {
      const face = cameraFaceNormal(points);
      if (!face) return false;
      const { normal, centroid } = face;
      const viewVector = perspectiveEnabled
        ? { x: -centroid.x, y: -centroid.y, z: cameraDistance - centroid.z }
        : { x: 0, y: 0, z: 1 };
      const dot = normal.x * viewVector.x + normal.y * viewVector.y + normal.z * viewVector.z;
      const magnitude = Math.hypot(viewVector.x, viewVector.y, viewVector.z);
      return magnitude > 0 && dot > magnitude * 1e-10;
    },
    cameraNormal(points) {
      const face = cameraFaceNormal(points);
      return face?.normal ?? null;
    },
    topology: { farX, farY, axisX, axisY, nearDepth, farDepth },
    seriesDepth,
    prismDepth,
    prismInterval(seriesIndex, seriesCount, stacked = false) {
      const centre = seriesDepth(seriesIndex, seriesCount, stacked);
      const half = prismDepth(stacked ? 1 : seriesCount) / 2;
      // The slot centre is authored in [0,1]. Intersecting with the scene is a
      // defensive fallback only; ordinary n>=1/gap>=0 intervals remain equal.
      const near = clamp(centre - half, 0, 1);
      const far = clamp(centre + half, 0, 1);
      return { near, far };
    },
  };
}
