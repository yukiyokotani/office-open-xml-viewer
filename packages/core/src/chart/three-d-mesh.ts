import type { ThreeDScenePoint } from './three-d.js';

/** Shapes supported by ECMA-376 CT_BarShape/ST_Shape. */
export type ThreeDMeshShape =
  | 'box'
  | 'cylinder'
  | 'cone'
  | 'coneToMax'
  | 'pyramid'
  | 'pyramidToMax';

export type ThreeDMeshKind = ThreeDMeshShape | 'areaStrip' | 'lineRibbon' | 'pieSector';

export type ThreeDMeshFaceRole = 'baseCap' | 'endCap' | 'side';

/** One outward-wound face in model space. No projection or paint is stored. */
export interface ThreeDMeshFace {
  readonly indices: readonly number[];
  readonly role: ThreeDMeshFaceRole;
  /** Curved sides are tessellated geometry whose internal facet edges are not outlines. */
  readonly smoothSurface: boolean;
  readonly segmentIndex?: number;
  /** Ring edges bounding a side facet. Used only to derive an authored outer
   * rim when the adjacent cap is back-facing; they are not facet outlines. */
  readonly baseRimEdge?: readonly [number, number];
  readonly endRimEdge?: readonly [number, number];
}

/** A complete bounded solid before camera projection. */
export interface ThreeDMesh {
  readonly shape: ThreeDMeshKind;
  readonly vertices: readonly ThreeDScenePoint[];
  readonly faces: readonly ThreeDMeshFace[];
  /** Longitudinal candidate edges used to derive a curved silhouette after culling. */
  readonly silhouetteEdges: ReadonlyArray<readonly [number, number]>;
}

export interface ThreeDAreaStripMeshOptions {
  readonly x0: number;
  readonly x1: number;
  readonly lower0: number;
  readonly lower1: number;
  readonly upper0: number;
  readonly upper1: number;
  readonly nearDepth: number;
  readonly farDepth: number;
  readonly capStart: boolean;
  readonly capEnd: boolean;
}

export interface ThreeDLineRibbonMeshOptions {
  /** One bounded stroke polygon in the chart's model-space X/Y plane. */
  readonly outline: readonly { readonly x: number; readonly y: number }[];
  readonly nearDepth: number;
  readonly farDepth: number;
}

export interface ThreeDPieSectorMeshOptions {
  readonly centerX: number;
  readonly centerY: number;
  readonly centerDepth: number;
  readonly radius: number;
  /** Physical model-space depth represented by normalized depth 0..1. */
  readonly modelDepth: number;
  readonly thickness: number;
  readonly startAngle: number;
  readonly endAngle: number;
  readonly segments?: number;
}

export interface ThreeDShapeMeshOptions {
  readonly shape: string;
  readonly horizontal: boolean;
  readonly crossStart: number;
  readonly crossSize: number;
  readonly baseCoord: number;
  readonly endCoord: number;
  readonly nearDepth: number;
  readonly farDepth: number;
  readonly toMaxBaseScale?: number;
  readonly toMaxEndScale?: number;
  /** Explicit scales preserve the cross-sections of a solid clipped by an
   * authored axis domain. When omitted, the shape's full 1→0 contract applies. */
  readonly baseScale?: number;
  readonly endScale?: number;
  readonly omitBaseCap?: boolean;
  readonly omitEndCap?: boolean;
  readonly roundSegments?: number;
}

export const DEFAULT_THREE_D_ROUND_SEGMENTS = 32;

export function normalizeThreeDMeshShape(shape: string): ThreeDMeshShape {
  switch (shape) {
    case 'cylinder':
    case 'cone':
    case 'coneToMax':
    case 'pyramid':
    case 'pyramidToMax':
      return shape;
    default:
      return 'box';
  }
}

const clamp01 = (value: number): number => Math.max(0, Math.min(1, value));

function faceNormal(
  vertices: readonly ThreeDScenePoint[],
  indices: readonly number[],
): ThreeDScenePoint | null {
  if (indices.length < 3) return null;
  const a = vertices[indices[0]];
  // A sign-crossing area strip collapses one end of a quad into a triangle.
  // Find the first non-collinear fan triangle rather than assuming the first
  // three stored vertices are distinct.
  for (let first = 1; first + 1 < indices.length; first++) {
    for (let second = first + 1; second < indices.length; second++) {
      const b = vertices[indices[first]];
      const c = vertices[indices[second]];
      const ab = { x: b.x - a.x, y: b.y - a.y, depth: b.depth - a.depth };
      const ac = { x: c.x - a.x, y: c.y - a.y, depth: c.depth - a.depth };
      const normal = {
        x: ab.y * ac.depth - ab.depth * ac.y,
        y: ab.depth * ac.x - ab.x * ac.depth,
        depth: ab.x * ac.y - ab.y * ac.x,
      };
      const length = Math.hypot(normal.x, normal.y, normal.depth);
      if (length > Number.EPSILON) {
        return { x: normal.x / length, y: normal.y / length, depth: normal.depth / length };
      }
    }
  }
  return null;
}

function outwardWoundFaces(
  vertices: readonly ThreeDScenePoint[],
  faces: readonly ThreeDMeshFace[],
): ThreeDMeshFace[] {
  const solidCenter = vertices.reduce(
    (sum, point) => ({
      x: sum.x + point.x / vertices.length,
      y: sum.y + point.y / vertices.length,
      depth: sum.depth + point.depth / vertices.length,
    }),
    { x: 0, y: 0, depth: 0 },
  );
  return faces.map(face => {
    const normal = faceNormal(vertices, face.indices);
    if (!normal) return face;
    const center = face.indices.reduce(
      (sum, index) => ({
        x: sum.x + vertices[index].x / face.indices.length,
        y: sum.y + vertices[index].y / face.indices.length,
        depth: sum.depth + vertices[index].depth / face.indices.length,
      }),
      { x: 0, y: 0, depth: 0 },
    );
    const outward = {
      x: center.x - solidCenter.x,
      y: center.y - solidCenter.y,
      depth: center.depth - solidCenter.depth,
    };
    const dot = normal.x * outward.x + normal.y * outward.y + normal.depth * outward.depth;
    return dot >= 0 ? face : { ...face, indices: [...face.indices].reverse() };
  });
}

/**
 * Build a real model-space mesh for a 3-D bar datum.
 *
 * Geometry is intentionally independent from Canvas paint and the camera:
 * - box/pyramid use a four-vertex rectangular cross-section,
 * - cylinder/cone use bounded circular rings,
 * - ToMax variants use two rings scaled from the same virtual axis apex.
 *
 * Every face is outward-wound after construction. Projection, back-face
 * culling, depth ordering and material lighting happen in later stages.
 */
export function buildThreeDShapeMesh(options: ThreeDShapeMeshOptions): ThreeDMesh | null {
  const {
    horizontal,
    crossStart,
    crossSize,
    baseCoord,
    endCoord,
    nearDepth,
    farDepth,
  } = options;
  if (![crossStart, crossSize, baseCoord, endCoord, nearDepth, farDepth]
    .every(Number.isFinite)
    || crossSize <= 0
    || baseCoord === endCoord
    || nearDepth === farDepth) return null;

  const shape = normalizeThreeDMeshShape(options.shape);
  const round = shape === 'cylinder' || shape === 'cone' || shape === 'coneToMax';
  const tapered = shape !== 'box' && shape !== 'cylinder';
  const toMax = shape === 'coneToMax' || shape === 'pyramidToMax';
  const segments = round
    ? Math.max(8, Math.min(64, Math.trunc(options.roundSegments ?? DEFAULT_THREE_D_ROUND_SEGMENTS)))
    : 4;
  const baseScale = clamp01(options.baseScale
    ?? (toMax ? options.toMaxBaseScale ?? 1 : 1));
  const endScale = clamp01(options.endScale
    ?? (tapered ? (toMax ? options.toMaxEndScale ?? 0 : 0) : 1));
  if (baseScale === 0 && endScale === 0) return null;

  const centerCross = crossStart + crossSize / 2;
  const crossRadius = crossSize / 2;
  const centerDepth = (nearDepth + farDepth) / 2;
  const depthRadius = Math.abs(farDepth - nearDepth) / 2;
  const vertices: ThreeDScenePoint[] = [];
  const ring = (coord: number, scale: number): number[] => {
    if (scale === 0) {
      const index = vertices.length;
      vertices.push(horizontal
        ? { x: coord, y: centerCross, depth: centerDepth }
        : { x: centerCross, y: coord, depth: centerDepth });
      return [index];
    }
    const indices: number[] = [];
    for (let index = 0; index < segments; index++) {
      // Four points describe the actual rectangular cross-section corners;
      // round shapes use the inscribed circular ring in the same slot.
      const angle = round
        ? index / segments * Math.PI * 2
        : Math.PI / 4 + index / segments * Math.PI * 2;
      const radiusFactor = round ? 1 : Math.SQRT2;
      const cross = centerCross + Math.cos(angle) * crossRadius * scale * radiusFactor;
      const depth = centerDepth + Math.sin(angle) * depthRadius * scale * radiusFactor;
      indices.push(vertices.length);
      vertices.push(horizontal ? { x: coord, y: cross, depth } : { x: cross, y: coord, depth });
    }
    return indices;
  };

  const base = ring(baseCoord, baseScale);
  const end = ring(endCoord, endScale);
  const smoothSurface = round;
  const faces: ThreeDMeshFace[] = [];
  if (base.length > 1 && options.omitBaseCap !== true) {
    faces.push({ indices: base, role: 'baseCap', smoothSurface });
  }
  if (end.length > 1 && options.omitEndCap !== true) {
    faces.push({ indices: end, role: 'endCap', smoothSurface });
  }
  const silhouetteEdges: Array<readonly [number, number]> = [];
  const sideCount = Math.max(base.length, end.length);
  for (let index = 0; index < sideCount; index++) {
    const next = (index + 1) % sideCount;
    const baseIndex = base.length === 1 ? base[0] : base[index];
    const baseNext = base.length === 1 ? base[0] : base[next];
    const endIndex = end.length === 1 ? end[0] : end[index];
    const endNext = end.length === 1 ? end[0] : end[next];
    const indices = base.length === 1
      ? [baseIndex, endNext, endIndex]
      : end.length === 1
        ? [baseIndex, baseNext, endIndex]
        : [baseIndex, baseNext, endNext, endIndex];
    faces.push({
      indices,
      role: 'side',
      smoothSurface,
      segmentIndex: index,
      baseRimEdge: base.length > 1 ? [baseIndex, baseNext] : undefined,
      endRimEdge: end.length > 1 ? [endIndex, endNext] : undefined,
    });
    silhouetteEdges.push([baseIndex, endIndex]);
  }

  return {
    shape,
    vertices,
    faces: outwardWoundFaces(vertices, faces),
    silhouetteEdges,
  };
}

function buildAreaStripMeshSingle(options: ThreeDAreaStripMeshOptions): ThreeDMesh | null {
  let {
    x0, x1, lower0, lower1, upper0, upper1,
    nearDepth, farDepth, capStart, capEnd,
  } = options;
  if (![x0, x1, lower0, lower1, upper0, upper1, nearDepth, farDepth]
    .every(Number.isFinite) || x0 === x1 || nearDepth === farDepth) return null;
  if (x1 < x0) {
    [x0, x1] = [x1, x0];
    [lower0, lower1] = [lower1, lower0];
    [upper0, upper1] = [upper1, upper0];
    [capStart, capEnd] = [capEnd, capStart];
  }
  const z0 = Math.min(nearDepth, farDepth);
  const z1 = Math.max(nearDepth, farDepth);
  const top0 = Math.min(lower0, upper0);
  const top1 = Math.min(lower1, upper1);
  const bottom0 = Math.max(lower0, upper0);
  const bottom1 = Math.max(lower1, upper1);
  if (Math.max(bottom0 - top0, bottom1 - top1) < 1e-9) return null;
  const vertices: ThreeDScenePoint[] = [
    { x: x0, y: top0, depth: z0 },
    { x: x1, y: top1, depth: z0 },
    { x: x1, y: bottom1, depth: z0 },
    { x: x0, y: bottom0, depth: z0 },
    { x: x0, y: top0, depth: z1 },
    { x: x1, y: top1, depth: z1 },
    { x: x1, y: bottom1, depth: z1 },
    { x: x0, y: bottom0, depth: z1 },
  ];
  const rawFaces: ThreeDMeshFace[] = [
    { indices: [0, 3, 2, 1], role: 'side', smoothSurface: false },
    { indices: [4, 5, 6, 7], role: 'side', smoothSurface: false },
    { indices: [0, 4, 7, 3], role: 'baseCap', smoothSurface: false },
    { indices: [1, 2, 6, 5], role: 'endCap', smoothSurface: false },
    { indices: [0, 1, 5, 4], role: 'side', smoothSurface: false },
    { indices: [3, 7, 6, 2], role: 'side', smoothSurface: false },
  ];
  const faces = rawFaces.filter(face =>
    (face.role !== 'baseCap' || capStart) && (face.role !== 'endCap' || capEnd));
  return {
    shape: 'areaStrip',
    vertices,
    faces: outwardWoundFaces(vertices, faces),
    silhouetteEdges: [],
  };
}

/** Build one or two closed model-space solids for an extruded area interval.
 * A sign change is split at the shared datum crossing so neither broad face is
 * a self-intersecting bow-tie. */
export function buildThreeDAreaStripMeshes(
  options: ThreeDAreaStripMeshOptions,
): ThreeDMesh[] {
  const delta0 = options.upper0 - options.lower0;
  const delta1 = options.upper1 - options.lower1;
  if (Number.isFinite(delta0) && Number.isFinite(delta1) && delta0 * delta1 < 0) {
    const t = delta0 / (delta0 - delta1);
    const crossingX = options.x0 + (options.x1 - options.x0) * t;
    const crossingValue = options.lower0 + (options.lower1 - options.lower0) * t;
    return [
      buildAreaStripMeshSingle({
        ...options,
        x1: crossingX,
        lower1: crossingValue,
        upper1: crossingValue,
        capEnd: false,
      }),
      buildAreaStripMeshSingle({
        ...options,
        x0: crossingX,
        lower0: crossingValue,
        upper0: crossingValue,
        capStart: false,
      }),
    ].filter((mesh): mesh is ThreeDMesh => mesh != null);
  }
  const mesh = buildAreaStripMeshSingle(options);
  return mesh ? [mesh] : [];
}

/** Extrude one already-tessellated line-stroke polygon through its authored
 * series-depth interval. Line3D is a solid ribbon in Excel, not a flat Canvas
 * stroke; reusing the bounded stroke polygon preserves dash/cap/join geometry
 * while the ordinary mesh projector supplies culling, depth order and material
 * shading exactly like bars and area strips. */
export function buildThreeDLineRibbonMesh(
  options: ThreeDLineRibbonMeshOptions,
): ThreeDMesh | null {
  const outline = options.outline;
  if (outline.length < 3
    || outline.length > 64
    || ![options.nearDepth, options.farDepth].every(Number.isFinite)
    || Math.abs(options.nearDepth - options.farDepth) < 1e-9
    || !outline.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) {
    return null;
  }
  const z0 = Math.min(options.nearDepth, options.farDepth);
  const z1 = Math.max(options.nearDepth, options.farDepth);
  const count = outline.length;
  const vertices: ThreeDScenePoint[] = [
    ...outline.map(point => ({ ...point, depth: z0 })),
    ...outline.map(point => ({ ...point, depth: z1 })),
  ];
  const near = Array.from({ length: count }, (_, index) => index);
  const far = Array.from({ length: count }, (_, index) => count + index).reverse();
  const faces: ThreeDMeshFace[] = [
    { indices: near, role: 'side', smoothSurface: false },
    { indices: far, role: 'side', smoothSurface: false },
    ...Array.from({ length: count }, (_, index): ThreeDMeshFace => {
      const next = (index + 1) % count;
      return {
        indices: [index, next, count + next, count + index],
        role: 'side',
        smoothSurface: false,
      };
    }),
  ];
  return {
    shape: 'lineRibbon',
    vertices,
    faces: outwardWoundFaces(vertices, faces),
    silhouetteEdges: [],
  };
}

/** Build a closed cylindrical sector in the shared cartesian camera space.
 * The pie top lies in the X/Z plane and thickness follows model Y, so rotX,
 * rotY, perspective, culling and occlusion are the same operations as every
 * other 3-D mesh rather than an independently squashed screen ellipse. */
export function buildThreeDPieSectorMesh(
  options: ThreeDPieSectorMeshOptions,
): ThreeDMesh | null {
  const {
    centerX, centerY, centerDepth, radius, modelDepth, thickness,
    startAngle, endAngle,
  } = options;
  if (![centerX, centerY, centerDepth, radius, modelDepth, thickness, startAngle, endAngle]
    .every(Number.isFinite)
    || !(radius > 0) || !(modelDepth > 0) || !(thickness > 0)
    || !(endAngle > startAngle)) return null;
  const sweep = Math.min(Math.PI * 2, endAngle - startAngle);
  const fullSweep = sweep >= Math.PI * 2 - 1e-9;
  const segments = Math.max(2, Math.min(128, Math.trunc(options.segments
    ?? Math.ceil(DEFAULT_THREE_D_ROUND_SEGMENTS * sweep / (Math.PI * 2)))));
  const topY = centerY - thickness / 2;
  const bottomY = centerY + thickness / 2;
  const vertices: ThreeDScenePoint[] = [
    { x: centerX, y: topY, depth: centerDepth },
    { x: centerX, y: bottomY, depth: centerDepth },
  ];
  const topArc: number[] = [];
  const bottomArc: number[] = [];
  const arcPointCount = fullSweep ? segments : segments + 1;
  for (let index = 0; index < arcPointCount; index++) {
    const angle = startAngle + sweep * index / segments;
    const x = centerX + Math.cos(angle) * radius;
    const depth = centerDepth + Math.sin(angle) * radius / modelDepth;
    topArc.push(vertices.length);
    vertices.push({ x, y: topY, depth });
    bottomArc.push(vertices.length);
    vertices.push({ x, y: bottomY, depth });
  }
  const faces: ThreeDMeshFace[] = fullSweep ? [
    // A complete pie is a cylinder, not a sector with a coincident radial
    // cut. Ring caps and wrapped side facets leave no internal seam to paint.
    { indices: [...topArc], role: 'baseCap', smoothSurface: true },
    { indices: [...bottomArc], role: 'endCap', smoothSurface: true },
  ] : [
    { indices: [0, ...topArc], role: 'baseCap', smoothSurface: true },
    { indices: [1, ...bottomArc], role: 'endCap', smoothSurface: true },
    {
      indices: [0, 1, bottomArc[0], topArc[0]],
      role: 'baseCap', smoothSurface: false,
    },
    {
      indices: [0, topArc.at(-1) as number, bottomArc.at(-1) as number, 1],
      role: 'endCap', smoothSurface: false,
    },
  ];
  for (let index = 0; index < segments; index++) {
    const next = fullSweep ? (index + 1) % segments : index + 1;
    faces.push({
      indices: [topArc[index], bottomArc[index], bottomArc[next], topArc[next]],
      role: 'side',
      smoothSurface: true,
      segmentIndex: index,
      baseRimEdge: [topArc[index], topArc[next]],
      endRimEdge: [bottomArc[index], bottomArc[next]],
    });
  }
  return {
    shape: 'pieSector',
    vertices,
    faces: outwardWoundFaces(vertices, faces),
    silhouetteEdges: topArc.slice(0, fullSweep ? undefined : -1).map((index, segment) =>
      [index, bottomArc[segment]] as const),
  };
}
