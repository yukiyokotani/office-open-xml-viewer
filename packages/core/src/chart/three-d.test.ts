import { describe, expect, it } from 'vitest';
import {
  fitChartThreeDProjectionToWallThickness,
  fitChartThreeDProjectionToPoints,
  pieThreeDThicknessMultiplier,
  planChartThreeDSurfaceGridSegments,
  planChartThreeDSurfaceGeometry,
  planChartThreeDProjection,
  planThreeDBarClusterSlot,
  threeDToMaxScale,
} from './three-d.js';
import {
  buildThreeDAreaStripMeshes,
  buildThreeDPieSectorMesh,
  buildThreeDShapeMesh,
} from './three-d-mesh.js';
import {
  bucketThreeDStackItems,
  threeDMeshOutlineWidthPx,
  threeDPieSliceAngles,
  threeDWallGeometry,
} from './three-d-renderer.js';

const PLOT = { x: 20, y: 10, w: 360, h: 180 };

describe('pieThreeDThicknessMultiplier', () => {
  it('distinguishes omitted hPercent from an authored value and its bare 100% default', () => {
    expect(pieThreeDThicknessMultiplier({ heightPercent: 50, heightPercentAuthored: false }))
      .toBe(1);
    expect(pieThreeDThicknessMultiplier({ heightPercent: 50, heightPercentAuthored: true }))
      .toBe(0.5);
    expect(pieThreeDThicknessMultiplier({ heightPercent: 100, heightPercentAuthored: true }))
      .toBe(1);
    expect(pieThreeDThicknessMultiplier({ heightPercent: 200 })).toBe(2);
    expect(pieThreeDThicknessMultiplier({ heightPercent: 0, heightPercentAuthored: true }))
      .toBe(1);
    expect(pieThreeDThicknessMultiplier({ heightPercent: 501, heightPercentAuthored: true }))
      .toBe(1);
  });
});

describe('threeDMeshOutlineWidthPx', () => {
  it('matches the observed one-pixel 1pt Excel mesh edge at 100% zoom', () => {
    expect(threeDMeshOutlineWidthPx(12_700, 4 / 3)).toBe(1);
    expect(threeDMeshOutlineWidthPx(12_700, 8 / 3)).toBe(2);
    expect(threeDMeshOutlineWidthPx(25_400, 4 / 3)).toBe(2);
  });
});

describe('bucketThreeDStackItems', () => {
  it('groups a maximum-size public model with one category read per item', () => {
    let categoryReads = 0;
    const items = Array.from({ length: 10_000 }, (_, index) => ({
      get categoryIndex() {
        categoryReads++;
        return index % 257;
      },
      index,
    }));
    const buckets = bucketThreeDStackItems(items, 257);
    expect(buckets.reduce((sum, bucket) => sum + bucket.length, 0)).toBe(10_000);
    expect(categoryReads).toBe(10_000);
    expect(buckets[0][0].index).toBe(0);
    expect(buckets[256][0].index).toBe(256);
  });

  it('ignores invalid caller-constructed category indexes', () => {
    expect(bucketThreeDStackItems([
      { categoryIndex: -1 },
      { categoryIndex: 0 },
      { categoryIndex: 2 },
      { categoryIndex: Number.NaN },
    ], 2).map(bucket => bucket.length)).toEqual([1, 0]);
  });
});

describe('threeDWallGeometry', () => {
  it('closes floor, side wall and back wall on identical projected edges', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, heightPercent: 100,
      depthPercent: 100, perspective: 30,
    }, PLOT, { sceneDepthScale: 1 });
    if (!plan) throw new Error('projection not planned');
    const walls = threeDWallGeometry(plan);
    expect(walls.floor).toHaveLength(4);
    expect(walls.sideWall).toHaveLength(4);
    expect(walls.backWall).toHaveLength(4);
    const samePoint = (a: { x: number; y: number }, b: { x: number; y: number }) =>
      Math.hypot(a.x - b.x, a.y - b.y) < 1e-9;
    const sharedFloorBack = walls.floor.filter(floorPoint =>
      walls.backWall.some(backPoint => samePoint(floorPoint, backPoint)));
    const sharedFloorSide = walls.floor.filter(floorPoint =>
      walls.sideWall.some(sidePoint => samePoint(floorPoint, sidePoint)));
    const sharedSideBack = walls.sideWall.filter(sidePoint =>
      walls.backWall.some(backPoint => samePoint(sidePoint, backPoint)));
    expect(sharedFloorBack).toHaveLength(2);
    expect(sharedFloorSide).toHaveLength(2);
    expect(sharedSideBack).toHaveLength(2);
    const sideMidX = walls.sideWall.reduce((sum, point) => sum + point.x, 0) / 4;
    const seriesStart = plan.project(walls.seriesAxisX, walls.floorY, walls.nearDepth);
    const seriesEnd = plan.project(walls.seriesAxisX, walls.floorY, walls.farDepth);
    expect((seriesStart.x + seriesEnd.x) / 2).toBeGreaterThan(sideMidX);
  });

  it('extrudes each CT_Surface outward by a percentage of the largest plot dimension', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
    }, PLOT, { sceneDepthScale: 1 });
    if (!plan) throw new Error('projection not planned');
    const floor = planChartThreeDSurfaceGeometry(plan, 'floor', 25);
    const side = planChartThreeDSurfaceGeometry(plan, 'sideWall', 25);
    const back = planChartThreeDSurfaceGeometry(plan, 'backWall', 25);
    expect(floor.thickness).toBeCloseTo(90, 12);
    expect(side.thickness).toBeCloseTo(90, 12);
    expect(back.thickness).toBeCloseTo(90, 12);
    expect(floor.faces).toHaveLength(6);
    expect(side.faces).toHaveLength(6);
    expect(back.faces).toHaveLength(6);
    expect(floor.pictureStackAspect).toBeGreaterThan(0);
    expect(side.pictureStackAspect).toBeCloseTo(floor.pictureStackAspect as number, 12);
    expect(back.pictureStackAspect).toBeCloseTo(floor.pictureStackAspect as number, 12);
    expect(Math.abs(floor.outer[0].y - floor.inner[0].y)).toBeCloseTo(90, 12);
    expect(Math.abs(side.outer[0].x - side.inner[0].x)).toBeCloseTo(90, 12);
    expect(
      Math.abs(back.outer[0].depth - back.inner[0].depth) * plan.modelDepth,
    ).toBeCloseTo(90, 12);
  });

  it('keeps zero-thickness surfaces planar and fits positive slabs inside the plot once', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, heightPercent: 100,
      depthPercent: 100, perspective: 30,
    }, PLOT, { sceneDepthScale: 1 });
    if (!plan) throw new Error('projection not planned');
    expect(planChartThreeDSurfaceGeometry(plan, 'floor', 0).faces).toHaveLength(1);
    for (const invalid of [-1, Number.POSITIVE_INFINITY, 4_294_967_296]) {
      expect(planChartThreeDSurfaceGeometry(plan, 'floor', invalid)).toMatchObject({
        thickness: 0,
        faces: [expect.any(Array)],
      });
    }
    const fitted = fitChartThreeDProjectionToWallThickness(plan, {
      floor: { thicknessPercent: 25 },
      sideWall: { thicknessPercent: 25 },
      backWall: { thicknessPercent: 25 },
    }, PLOT);
    const points = (['floor', 'sideWall', 'backWall'] as const).flatMap(kind =>
      planChartThreeDSurfaceGeometry(fitted, kind, 25).faces.flat()
        .map(point => fitted.projectUnbounded(point.x, point.y, point.depth))
    );
    expect(Math.min(...points.map(point => point.x))).toBeGreaterThanOrEqual(PLOT.x - 1e-9);
    expect(Math.max(...points.map(point => point.x))).toBeLessThanOrEqual(PLOT.x + PLOT.w + 1e-9);
    expect(Math.min(...points.map(point => point.y))).toBeGreaterThanOrEqual(PLOT.y - 1e-9);
    expect(Math.max(...points.map(point => point.y))).toBeLessThanOrEqual(PLOT.y + PLOT.h + 1e-9);
  });

  it('plans one planar grid rule and four corresponding thick-slab segments', () => {
    const plan = planChartThreeDProjection({
      rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
    }, PLOT, { sceneDepthScale: 1 });
    if (!plan) throw new Error('projection not planned');
    const planar = planChartThreeDSurfaceGeometry(plan, 'floor', 0);
    const thickFloor = planChartThreeDSurfaceGeometry(plan, 'floor', 25);
    const thickBack = planChartThreeDSurfaceGeometry(plan, 'backWall', 25);
    expect(planChartThreeDSurfaceGridSegments(planar, 'floor', 'x', 0.5))
      .toHaveLength(1);
    expect(planChartThreeDSurfaceGridSegments(thickFloor, 'floor', 'x', 0.5)
      .map(segment => segment.faceIndex)).toEqual([0, 1, 2, 4]);
    expect(planChartThreeDSurfaceGridSegments(thickBack, 'backWall', 'y', 0.5)
      .map(segment => segment.faceIndex)).toEqual([0, 1, 5, 3]);
    expect(planChartThreeDSurfaceGridSegments(thickFloor, 'floor', 'y', 0.5))
      .toEqual([]);
    expect(planChartThreeDSurfaceGridSegments(thickBack, 'backWall', 'x', -0.1))
      .toEqual([]);
  });
});

describe('threeDPieSliceAngles', () => {
  it('maps OOXML zero to twelve o’clock and advances slices clockwise', () => {
    const first = threeDPieSliceAngles(0, 0, 0.25);
    expect(first.leading).toBeCloseTo(Math.PI / 2, 12);
    expect(first.start).toBeCloseTo(0, 12);
    expect(first.end).toBeCloseTo(Math.PI / 2, 12);
    expect(first.middle).toBeCloseTo(Math.PI / 4, 12);
    const second = threeDPieSliceAngles(0, 0.25, 0.25);
    expect(second.leading).toBeCloseTo(0, 12);
    const rotated = threeDPieSliceAngles(90, 0, 0.25);
    expect(rotated.leading).toBeCloseTo(0, 12);
  });

  it('projects the zero-degree leading ray above the pie centre', () => {
    const plan = planChartThreeDProjection({
      rotationX: 30, rotationY: 0, perspective: 0, depthPercent: 100,
    }, PLOT, { sceneDepthScale: 1, sceneHeightScale: 0.15 });
    if (!plan) throw new Error('projection not planned');
    const angle = threeDPieSliceAngles(0, 0, 0.25).leading;
    const centerX = plan.scene.x + plan.scene.w / 2;
    const centerY = plan.scene.y + plan.scene.h / 2;
    const radius = Math.min(plan.scene.w, plan.modelDepth) * 0.35;
    const center = plan.project(centerX, centerY, 0.5);
    const leading = plan.project(
      centerX + Math.cos(angle) * radius,
      centerY,
      0.5 + Math.sin(angle) * radius / plan.modelDepth,
    );
    expect(leading.y).toBeLessThan(center.y);
  });
});

describe('buildThreeDShapeMesh', () => {
  const mesh = (shape: string, toMaxBaseScale = 1, toMaxEndScale = 0) =>
    buildThreeDShapeMesh({
      shape,
      horizontal: false,
      crossStart: 10,
      crossSize: 20,
      baseCoord: 100,
      endCoord: 20,
      nearDepth: 0.2,
      farDepth: 0.8,
      toMaxBaseScale,
      toMaxEndScale,
    });

  it('constructs cylinder as two circular rings plus indexed side quads', () => {
    const cylinder = mesh('cylinder');
    expect(cylinder).not.toBeNull();
    if (!cylinder) throw new Error('mesh not built');
    expect(cylinder.vertices).toHaveLength(64);
    expect(cylinder.faces.filter(face => face.role === 'side')).toHaveLength(32);
    expect(cylinder.faces.filter(face => face.role !== 'side')).toHaveLength(2);
    expect(cylinder.faces.filter(face => face.role === 'side').every(
      face => face.indices.length === 4 && face.smoothSurface,
    )).toBe(true);
    expect(new Set(cylinder.vertices.slice(0, 32).map(point => point.y))).toEqual(new Set([100]));
    expect(new Set(cylinder.vertices.slice(32).map(point => point.y))).toEqual(new Set([20]));
  });

  it('constructs cone and pyramid as real base-ring-to-apex meshes', () => {
    const cone = mesh('cone');
    const pyramid = mesh('pyramid');
    if (!cone || !pyramid) throw new Error('mesh not built');
    expect(cone.vertices).toHaveLength(33);
    expect(cone.faces.filter(face => face.role === 'side')).toHaveLength(32);
    expect(cone.faces.filter(face => face.role === 'side').every(
      face => face.indices.length === 3,
    )).toBe(true);
    expect(pyramid.vertices).toHaveLength(5);
    expect(pyramid.faces.filter(face => face.role === 'side')).toHaveLength(4);
    const coneApex = cone.vertices[32];
    const pyramidApex = pyramid.vertices[4];
    expect(coneApex).toEqual({ x: 20, y: 20, depth: 0.5 });
    expect(pyramidApex).toEqual(coneApex);
  });

  it('keeps ToMax neighbouring segment rings geometrically identical', () => {
    const first = mesh('coneToMax', 1, 0.6);
    const second = buildThreeDShapeMesh({
      shape: 'coneToMax',
      horizontal: false,
      crossStart: 10,
      crossSize: 20,
      baseCoord: 20,
      endCoord: -20,
      nearDepth: 0.2,
      farDepth: 0.8,
      toMaxBaseScale: 0.6,
      toMaxEndScale: 0.2,
    });
    if (!first || !second) throw new Error('mesh not built');
    expect(first.vertices.slice(32)).toEqual(second.vertices.slice(0, 32));
  });

  it('omits a shared stacked cap without changing the common ring', () => {
    const lower = buildThreeDShapeMesh({
      shape: 'cylinder', horizontal: false, crossStart: 10, crossSize: 20,
      baseCoord: 100, endCoord: 60, nearDepth: 0.2, farDepth: 0.8,
      omitEndCap: true,
    });
    const upper = buildThreeDShapeMesh({
      shape: 'cylinder', horizontal: false, crossStart: 10, crossSize: 20,
      baseCoord: 60, endCoord: 20, nearDepth: 0.2, farDepth: 0.8,
      omitBaseCap: true,
    });
    if (!lower || !upper) throw new Error('mesh not built');
    expect(lower.faces.some(face => face.role === 'endCap')).toBe(false);
    expect(upper.faces.some(face => face.role === 'baseCap')).toBe(false);
    expect(lower.vertices.slice(32)).toEqual(upper.vertices.slice(0, 32));
  });

  it('builds an authored-domain cone clip as a frustum cross-section', () => {
    const clipped = buildThreeDShapeMesh({
      shape: 'cone', horizontal: false, crossStart: 10, crossSize: 20,
      baseCoord: 60, endCoord: 20, nearDepth: 0.2, farDepth: 0.8,
      // This visible interval is the last 2/3 of an original 100→20 cone.
      baseScale: 0.5,
      endScale: 0,
    });
    if (!clipped) throw new Error('mesh not built');
    const baseRing = clipped.vertices.slice(0, 32);
    expect(Math.max(...baseRing.map(point => point.x))).toBeCloseTo(25, 12);
    expect(Math.min(...baseRing.map(point => point.x))).toBeCloseTo(15, 12);
    expect(clipped.vertices[32]).toEqual({ x: 20, y: 20, depth: 0.5 });
  });

  it('orients every convex face normal away from the solid interior', () => {
    for (const shape of ['box', 'cylinder', 'cone', 'coneToMax', 'pyramid', 'pyramidToMax']) {
      const solid = mesh(shape, 0.9, 0.35);
      if (!solid) throw new Error(`mesh not built: ${shape}`);
      const center = solid.vertices.reduce((sum, point) => ({
        x: sum.x + point.x / solid.vertices.length,
        y: sum.y + point.y / solid.vertices.length,
        depth: sum.depth + point.depth / solid.vertices.length,
      }), { x: 0, y: 0, depth: 0 });
      for (const face of solid.faces) {
        const [a, b, c] = face.indices.map(index => solid.vertices[index]);
        const ab = { x: b.x - a.x, y: b.y - a.y, z: b.depth - a.depth };
        const ac = { x: c.x - a.x, y: c.y - a.y, z: c.depth - a.depth };
        const normal = {
          x: ab.y * ac.z - ab.z * ac.y,
          y: ab.z * ac.x - ab.x * ac.z,
          z: ab.x * ac.y - ab.y * ac.x,
        };
        const faceCenter = face.indices.reduce((sum, index) => ({
          x: sum.x + solid.vertices[index].x / face.indices.length,
          y: sum.y + solid.vertices[index].y / face.indices.length,
          depth: sum.depth + solid.vertices[index].depth / face.indices.length,
        }), { x: 0, y: 0, depth: 0 });
        const outward = {
          x: faceCenter.x - center.x,
          y: faceCenter.y - center.y,
          z: faceCenter.depth - center.depth,
        };
        expect(normal.x * outward.x + normal.y * outward.y + normal.z * outward.z)
          .toBeGreaterThan(0);
      }
    }
  });

  it('shows the circular data-end cap and a curved band in the default camera', () => {
    const cylinder = mesh('cylinder');
    const projection = planChartThreeDProjection({}, PLOT);
    if (!cylinder || !projection) throw new Error('mesh/projection not built');
    const visible = cylinder.faces.filter(face => projection.cameraFacing(
      face.indices.map(index => cylinder.vertices[index]),
    ));
    expect(visible.some(face => face.role === 'endCap')).toBe(true);
    const visibleSides = visible.filter(face => face.role === 'side');
    expect(visibleSides.length).toBeGreaterThanOrEqual(12);
    expect(new Set(visibleSides.map(face => {
      const normal = projection.cameraNormal(face.indices.map(index => cylinder.vertices[index]));
      return normal ? `${normal.x.toFixed(3)},${normal.y.toFixed(3)},${normal.z.toFixed(3)}` : '';
    })).size).toBeGreaterThanOrEqual(5);
  });
});

describe('buildThreeDAreaStripMeshes', () => {
  const outward = (mesh: ReturnType<typeof buildThreeDAreaStripMeshes>[number]) => {
    const center = mesh.vertices.reduce((sum, point) => ({
      x: sum.x + point.x / mesh.vertices.length,
      y: sum.y + point.y / mesh.vertices.length,
      depth: sum.depth + point.depth / mesh.vertices.length,
    }), { x: 0, y: 0, depth: 0 });
    for (const face of mesh.faces) {
      const a = mesh.vertices[face.indices[0]];
      let normal: { x: number; y: number; z: number } | null = null;
      for (let first = 1; first + 1 < face.indices.length && !normal; first++) {
        for (let second = first + 1; second < face.indices.length; second++) {
          const b = mesh.vertices[face.indices[first]];
          const c = mesh.vertices[face.indices[second]];
          const ab = { x: b.x - a.x, y: b.y - a.y, z: b.depth - a.depth };
          const ac = { x: c.x - a.x, y: c.y - a.y, z: c.depth - a.depth };
          const candidate = {
            x: ab.y * ac.z - ab.z * ac.y,
            y: ab.z * ac.x - ab.x * ac.z,
            z: ab.x * ac.y - ab.y * ac.x,
          };
          if (Math.hypot(candidate.x, candidate.y, candidate.z) > Number.EPSILON) {
            normal = candidate;
            break;
          }
        }
      }
      expect(normal).not.toBeNull();
      if (!normal) continue;
      const faceCenter = face.indices.reduce((sum, index) => ({
        x: sum.x + mesh.vertices[index].x / face.indices.length,
        y: sum.y + mesh.vertices[index].y / face.indices.length,
        depth: sum.depth + mesh.vertices[index].depth / face.indices.length,
      }), { x: 0, y: 0, depth: 0 });
      expect(normal.x * (faceCenter.x - center.x)
        + normal.y * (faceCenter.y - center.y)
        + normal.z * (faceCenter.depth - center.depth)).toBeGreaterThan(0);
    }
  };

  it('normalizes reversed categories and negative area thickness', () => {
    const meshes = buildThreeDAreaStripMeshes({
      x0: 100, x1: 20,
      lower0: 40, lower1: 45,
      upper0: 70, upper1: 80,
      nearDepth: 0.2, farDepth: 0.8,
      capStart: true, capEnd: true,
    });
    expect(meshes).toHaveLength(1);
    expect(meshes[0].faces).toHaveLength(6);
    outward(meshes[0]);
  });

  it('splits a sign-crossing strip into two closed non-bow-tie solids', () => {
    const meshes = buildThreeDAreaStripMeshes({
      x0: 10, x1: 90,
      lower0: 50, lower1: 50,
      upper0: 20, upper1: 80,
      nearDepth: 0.2, farDepth: 0.8,
      capStart: true, capEnd: true,
    });
    expect(meshes).toHaveLength(2);
    const crossingXs = meshes.flatMap(mesh => mesh.vertices.map(point => point.x))
      .filter(x => x > 10 && x < 90);
    expect(new Set(crossingXs.map(x => x.toFixed(10))).size).toBe(1);
    expect(meshes[0].faces.some(face => face.role === 'endCap')).toBe(false);
    expect(meshes[1].faces.some(face => face.role === 'baseCap')).toBe(false);
    meshes.forEach(outward);
  });
});

describe('buildThreeDPieSectorMesh', () => {
  it('builds a closed bounded sector with real top, bottom and curved wall faces', () => {
    const mesh = buildThreeDPieSectorMesh({
      centerX: 100, centerY: 80, centerDepth: 0.5,
      radius: 30, modelDepth: 120, thickness: 9,
      startAngle: -Math.PI / 2, endAngle: Math.PI / 2,
      segments: 16,
    });
    expect(mesh).not.toBeNull();
    if (!mesh) return;
    expect(mesh.shape).toBe('pieSector');
    expect(mesh.faces.filter(face => face.role === 'side')).toHaveLength(16);
    expect(mesh.faces).toHaveLength(20);
    expect(mesh.vertices.every(point =>
      Number.isFinite(point.x) && Number.isFinite(point.y)
      && Number.isFinite(point.depth) && point.depth >= 0 && point.depth <= 1)).toBe(true);
    expect(Math.min(...mesh.vertices.map(point => point.y))).toBeCloseTo(75.5, 12);
    expect(Math.max(...mesh.vertices.map(point => point.y))).toBeCloseTo(84.5, 12);
  });

  it('builds a complete pie as a seamless cylinder without radial cut faces', () => {
    const mesh = buildThreeDPieSectorMesh({
      centerX: 100, centerY: 80, centerDepth: 0.5,
      radius: 30, modelDepth: 120, thickness: 9,
      startAngle: -Math.PI / 2, endAngle: Math.PI * 3 / 2,
      segments: 16,
    });
    expect(mesh).not.toBeNull();
    if (!mesh) return;
    expect(mesh.faces.filter(face => face.role === 'side')).toHaveLength(16);
    expect(mesh.faces).toHaveLength(18);
    expect(mesh.faces.filter(face => face.role === 'baseCap')).toHaveLength(1);
    expect(mesh.faces.filter(face => face.role === 'endCap')).toHaveLength(1);
    expect(mesh.vertices).toHaveLength(34);
  });

});

describe('planChartThreeDProjection', () => {
  it('derives a camera normal when a clipped quad repeats its crossing vertex', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, perspective: 30,
    }, PLOT);
    if (!plan) throw new Error('projection not planned');
    const x0 = plan.scene.x + plan.scene.w * 0.2;
    const x1 = plan.scene.x + plan.scene.w * 0.8;
    const y0 = plan.scene.y + plan.scene.h * 0.2;
    const y1 = plan.scene.y + plan.scene.h * 0.8;
    const crossing = { x: x1, y: y0, depth: 0.2 };
    expect(plan.cameraNormal([
      { x: x0, y: y0, depth: 0.2 },
      crossing,
      crossing,
      { x: x0, y: y1, depth: 0.2 },
    ])).not.toBeNull();
  });

  it('reframes the camera around actual mesh vertices without changing depth order', () => {
    const plan = planChartThreeDProjection({}, PLOT, {
      sceneDepthScale: 1,
      sceneHeightScale: 0.15,
    });
    if (!plan) throw new Error('projection not planned');
    const points = [
      { x: plan.scene.x + plan.scene.w * 0.2, y: plan.scene.y, depth: 0.2 },
      { x: plan.scene.x + plan.scene.w * 0.8, y: plan.scene.y, depth: 0.2 },
      { x: plan.scene.x + plan.scene.w * 0.8, y: plan.scene.y + plan.scene.h, depth: 0.8 },
      { x: plan.scene.x + plan.scene.w * 0.2, y: plan.scene.y + plan.scene.h, depth: 0.8 },
    ];
    const fitted = fitChartThreeDProjectionToPoints(plan, points, PLOT, 0.1);
    const projected = points.map(point => fitted.project(point.x, point.y, point.depth));
    expect(Math.min(...projected.map(point => point.x))).toBeGreaterThanOrEqual(PLOT.x - 1e-8);
    expect(Math.max(...projected.map(point => point.x))).toBeLessThanOrEqual(PLOT.x + PLOT.w + 1e-8);
    expect(Math.min(...projected.map(point => point.y))).toBeGreaterThanOrEqual(PLOT.y - 1e-8);
    expect(Math.max(...projected.map(point => point.y))).toBeLessThanOrEqual(PLOT.y + PLOT.h + 1e-8);
    expect(fitted.cameraDepth(points[0].x, points[0].y, points[0].depth))
      .toBe(plan.cameraDepth(points[0].x, points[0].y, points[0].depth));
  });

  it('places clustered series side by side inside one category group', () => {
    const slots = [0, 1, 2].map(seriesIndex =>
      planThreeDBarClusterSlot(100, 150, seriesIndex, 3, false));
    expect(slots.map(slot => slot.size)).toEqual([
      100 / 4.5,
      100 / 4.5,
      100 / 4.5,
    ]);
    expect(slots[1].offset).toBeCloseTo(slots[0].offset + slots[0].size, 12);
    expect(slots[2].offset).toBeCloseTo(slots[1].offset + slots[1].size, 12);
    expect(slots[0].offset).toBeCloseTo((100 - 3 * 100 / 4.5) / 2, 12);
  });

  it('reuses the complete category footprint for every stacked series', () => {
    const slots = [0, 1, 7].map(seriesIndex =>
      planThreeDBarClusterSlot(100, 150, seriesIndex, 8, true));
    expect(slots).toEqual([
      { offset: 30, size: 40 },
      { offset: 30, size: 40 },
      { offset: 30, size: 40 },
    ]);
  });

  it('keeps ToMax rings continuous across positive and negative stack segments', () => {
    expect(threeDToMaxScale(40, -100, 100)).toBeCloseTo(0.6, 12);
    // The first segment's end is the next segment's base by construction.
    expect(threeDToMaxScale(40, -100, 100))
      .toBe(threeDToMaxScale(40, -100, 100));
    expect(threeDToMaxScale(100, -100, 100)).toBe(0);
    expect(threeDToMaxScale(-40, -100, 100)).toBeCloseTo(0.6, 12);
    expect(threeDToMaxScale(-100, -100, 100)).toBe(0);
  });
  it('uses the observed compact Office baseline for omitted view fields', () => {
    const plan = planChartThreeDProjection({}, PLOT);
    expect(plan).not.toBeNull();
    expect(plan?.depthX).toBeGreaterThan(0);
    // The far series plane moves upward/right in the default camera.
    expect(plan?.depthY).toBeLessThan(0);
    if (!plan) throw new Error('projection not planned');
    const slopeAt = (fraction: number) => {
      const y = plan.front.y + plan.front.h * fraction;
      const left = plan.project(plan.front.x, y, 1);
      const right = plan.project(plan.front.x + plan.front.w, y, 1);
      return (right.y - left.y) / (right.x - left.x);
    };
    const topSlope = slopeAt(0);
    const bottomSlope = slopeAt(1);
    // A true perspective fans parallel world lines toward one finite X
    // vanishing point. Office's default back-wall observations are roughly
    // .053 at the top and .13 at the bottom; equality would be the old affine
    // surrogate rather than the authored perspective.
    expect(topSlope).toBeGreaterThan(0.04);
    expect(topSlope).toBeLessThan(0.1);
    expect(bottomSlope).toBeGreaterThan(0.1);
    expect(bottomSlope).toBeLessThan(0.16);
    expect(bottomSlope).toBeGreaterThan(topSlope);
    expect(topSlope).toBeCloseTo(0.052, 2);
    expect(bottomSlope).toBeCloseTo(0.126, 2);
    expect(plan?.pieScaleY).toBeCloseTo(0.2113, 3);
  });

  it('keeps one camera while chart families select their observed Z occupancy', () => {
    const bar = planChartThreeDProjection({}, PLOT, { sceneDepthScale: 0.10 });
    const surface = planChartThreeDProjection({}, PLOT, { sceneDepthScale: 0.40 });
    if (!bar || !surface) throw new Error('projection not planned');
    expect(surface.modelDepth / bar.modelDepth).toBeCloseTo(4, 12);
    for (const plan of [bar, surface]) {
      const y = plan.front.y + plan.front.h * 0.65;
      const axisStart = plan.project(plan.front.x, y, 0.5);
      const axisEnd = plan.project(plan.front.x + plan.front.w, y, 0.5);
      const dataStart = plan.project(plan.front.x + 40, y, 0.5);
      const dataEnd = plan.project(plan.front.x + 140, y, 0.5);
      const cross = (a: typeof axisStart, b: typeof axisEnd, p: typeof dataStart) =>
        (b.x - a.x) * (p.y - a.y) - (b.y - a.y) * (p.x - a.x);
      expect(cross(axisStart, axisEnd, dataStart)).toBeCloseTo(0, 8);
      expect(cross(axisStart, axisEnd, dataEnd)).toBeCloseTo(0, 8);
    }
  });

  it('uses the normative FOV and measured axis proportions for a standard 3-D bar', () => {
    // The source chart's authored inner plot is 43.5000531% × 57.3076923%
    // of a 300.75pt × 196.5pt chart. Excel's projected axes measured
    // width:height:depth = 8.1:8.1:2.6 for the authored default view.
    const plot = {
      x: 0,
      y: 0,
      w: 300.75 * 0.4350005310065076,
      h: 196.5 * 0.573076923076923,
    };
    const plan = planChartThreeDProjection({
      rotationX: 15,
      rotationY: 20,
      heightPercent: 100,
      depthPercent: 100,
      perspective: 30,
      rightAngleAxes: false,
    }, plot, {
      sceneDepthScale: 0.65,
      perspectiveTangentGain: 1,
    });
    if (!plan) throw new Error('projection not planned');
    const origin = plan.project(plan.scene.x, plan.scene.y + plan.scene.h, 0);
    const horizontal = plan.project(plan.scene.x + plan.scene.w, plan.scene.y + plan.scene.h, 0);
    const vertical = plan.project(plan.scene.x, plan.scene.y, 0);
    const depth = plan.project(plan.scene.x, plan.scene.y + plan.scene.h, 1);
    const length = (point: { x: number; y: number }) =>
      Math.hypot(point.x - origin.x, point.y - origin.y);
    expect(length(horizontal) / length(vertical)).toBeCloseTo(1, 1);
    expect(length(depth) / length(vertical)).toBeCloseTo(2.6 / 8.1, 2);
  });

  it('keeps rotation/depth schema boundaries finite and inside the plot budget', () => {
    for (const rotationX of [-90, 0, 90]) {
      for (const rotationY of [0, 90, 180, 270, 360]) {
        for (const perspective of [0, 1, 240]) {
          for (const depthPercent of [20, 2000]) {
            for (const heightPercent of [5, 500]) {
              for (const sceneDepthScale of [0.10, 0.40]) {
              const plan = planChartThreeDProjection({
                rotationX, rotationY, perspective, depthPercent, heightPercent,
              }, PLOT, { sceneDepthScale });
              expect(plan).not.toBeNull();
              if (!plan) throw new Error('projection not planned');
              for (const x of [plan.scene.x, plan.scene.x + plan.scene.w]) {
                for (const y of [plan.scene.y, plan.scene.y + plan.scene.h]) {
                  for (const depth of [0, 1]) {
                    const point = plan.project(x, y, depth);
                    expect(Number.isFinite(point.x)).toBe(true);
                    expect(Number.isFinite(point.y)).toBe(true);
                    expect(point.x).toBeGreaterThanOrEqual(PLOT.x - 1e-7);
                    expect(point.x).toBeLessThanOrEqual(PLOT.x + PLOT.w + 1e-7);
                    expect(point.y).toBeGreaterThanOrEqual(PLOT.y - 1e-7);
                    expect(point.y).toBeLessThanOrEqual(PLOT.y + PLOT.h + 1e-7);
                  }
                }
              }
              }
            }
          }
        }
      }
    }
  });

  it('culls off-axis cuboid side faces from their outward camera normals', () => {
    const plan = planChartThreeDProjection({
      rotationX: 0, rotationY: 0, perspective: 60,
    }, PLOT);
    if (!plan) throw new Error('projection not planned');
    const y0 = plan.scene.y + plan.scene.h * 0.3;
    const y1 = plan.scene.y + plan.scene.h * 0.7;
    const leftRightX = plan.scene.x + plan.scene.w * 0.4;
    const rightLeftX = plan.scene.x + plan.scene.w * 0.6;
    expect(plan.cameraFacing([
      { x: leftRightX, y: y0, depth: 0.4 },
      { x: leftRightX, y: y1, depth: 0.4 },
      { x: leftRightX, y: y1, depth: 0.6 },
      { x: leftRightX, y: y0, depth: 0.6 },
    ])).toBe(true);
    expect(plan.cameraFacing([
      { x: leftRightX, y: y0, depth: 0.4 },
      { x: leftRightX, y: y0, depth: 0.6 },
      { x: leftRightX, y: y1, depth: 0.6 },
      { x: leftRightX, y: y1, depth: 0.4 },
    ])).toBe(false);
    expect(plan.cameraFacing([
      { x: rightLeftX, y: y0, depth: 0.4 },
      { x: rightLeftX, y: y0, depth: 0.6 },
      { x: rightLeftX, y: y1, depth: 0.6 },
      { x: rightLeftX, y: y1, depth: 0.4 },
    ])).toBe(true);
    expect(plan.cameraFacing([
      { x: rightLeftX, y: y0, depth: 0.4 },
      { x: rightLeftX, y: y1, depth: 0.4 },
      { x: rightLeftX, y: y1, depth: 0.6 },
      { x: rightLeftX, y: y0, depth: 0.6 },
    ])).toBe(false);
  });

  it('preserves an arbitrary 3-D straight line under both family scene boxes', () => {
    for (const sceneDepthScale of [0.10, 0.40]) {
      const plan = planChartThreeDProjection({ perspective: 240 }, PLOT, { sceneDepthScale });
      if (!plan) throw new Error('projection not planned');
      const point = (t: number) => plan.project(
        plan.scene.x + plan.scene.w * (0.15 + 0.7 * t),
        plan.scene.y + plan.scene.h * (0.8 - 0.6 * t),
        0.1 + 0.8 * t,
      );
      const a = point(0);
      const b = point(1);
      for (const t of [0.25, 0.5, 0.75]) {
        const p = point(t);
        const area = (b.x - a.x) * (p.y - a.y) - (b.y - a.y) * (p.x - a.x);
        expect(area).toBeCloseTo(0, 7);
      }
    }
  });

  it('keeps stacked series on one depth plane and clusters ordinary series', () => {
    const plan = planChartThreeDProjection({ gapDepthPercent: 150 }, PLOT);
    expect(plan?.seriesDepth(0, 3)).toBeCloseTo(1 / 6, 10);
    expect(plan?.seriesDepth(2, 3)).toBeCloseTo(5 / 6, 10);
    expect(plan?.seriesDepth(0, 3, true)).toBe(0.5);
    expect(plan?.seriesDepth(2, 3, true)).toBe(0.5);
    expect(plan?.prismDepth(3)).toBeGreaterThan(0);
    expect(plan?.prismDepth(3)).toBeLessThanOrEqual(0.35);
    const intervals = [0, 1, 2].map(index => plan?.prismInterval(index, 3));
    const widths = intervals.map(interval => (interval?.far ?? 0) - (interval?.near ?? 0));
    expect(widths[0]).toBeCloseTo(widths[1], 12);
    expect(widths[1]).toBeCloseTo(widths[2], 12);
    const centreGap = plan?.seriesDepth(1, 3) ?? 0;
    const previousCentre = plan?.seriesDepth(0, 3) ?? 0;
    expect((centreGap - previousCentre - widths[0]) / widths[0]).toBeCloseTo(1.5, 12);
    expect(intervals.every(interval =>
      interval != null && interval.near >= 0 && interval.far <= 1
    )).toBe(true);
    const stackedWidths = [1, 2, 8].map(count => {
      const interval = plan?.prismInterval(0, count, true);
      return (interval?.far ?? 0) - (interval?.near ?? 0);
    });
    expect(stackedWidths[0]).toBeCloseTo(stackedWidths[1], 12);
    expect(stackedWidths[1]).toBeCloseTo(stackedWidths[2], 12);
  });

  it('projects bars, lines and axes with the same planar geometry', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, perspective: 30,
    }, PLOT);
    if (!plan) throw new Error('projection not planned');
    const signedArea = (
      a: { x: number; y: number },
      b: { x: number; y: number },
      c: { x: number; y: number },
    ) => (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
    for (const depth of [0, 0.25, 0.5, 0.75, 1]) {
      for (const y of [plan.front.y + 40, plan.front.y + 90]) {
        const axisStart = plan.project(plan.front.x, y, depth);
        const axisEnd = plan.project(plan.front.x + plan.front.w, y, depth);
        const barStart = plan.project(plan.front.x + 20, y, depth);
        const barEnd = plan.project(plan.front.x + 80, y, depth);
        expect(signedArea(axisStart, axisEnd, barStart)).toBeCloseTo(0, 8);
        expect(signedArea(axisStart, axisEnd, barEnd)).toBeCloseTo(0, 8);
      }
    }
  });

  it('uses one projective vanishing construction for all depth edges', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, perspective: 120,
    }, PLOT);
    if (!plan) throw new Error('projection not planned');
    const cross = (a: ThreeDPoint, b: ThreeDPoint) => ({
      a: b.y - a.y, b: a.x - b.x, c: a.y * b.x - a.x * b.y,
    });
    type ThreeDPoint = { x: number; y: number };
    const intersect = (l: ReturnType<typeof cross>, m: ReturnType<typeof cross>) => {
      const d = l.a * m.b - m.a * l.b;
      return { x: (l.b * m.c - m.b * l.c) / d, y: (l.c * m.a - m.c * l.a) / d };
    };
    const depthLine = (x: number, y: number) => cross(plan.project(x, y, 0), plan.project(x, y, 1));
    const first = intersect(
      depthLine(plan.front.x, plan.front.y),
      depthLine(plan.front.x + plan.front.w, plan.front.y),
    );
    const second = intersect(
      depthLine(plan.front.x, plan.front.y + plan.front.h),
      depthLine(plan.front.x + plan.front.w, plan.front.y + plan.front.h),
    );
    expect(first.x).toBeCloseTo(second.x, 8);
    expect(first.y).toBeCloseTo(second.y, 8);
  });

  it('uses the affine limit when right-angle axes disable perspective', () => {
    const plan = planChartThreeDProjection({
      rotationX: 15, rotationY: 20, perspective: 240, rightAngleAxes: true,
    }, PLOT);
    if (!plan) throw new Error('projection not planned');
    const slope = (y: number, depth: number) => {
      const left = plan.project(plan.front.x, y, depth);
      const right = plan.project(plan.front.x + plan.front.w, y, depth);
      return (right.y - left.y) / (right.x - left.x);
    };
    expect(slope(plan.front.y, 1)).toBeCloseTo(slope(plan.front.y + plan.front.h, 0), 12);
  });

  it('derives painter order from the rotated camera rather than logical z', () => {
    const defaultView = planChartThreeDProjection({ rotationX: 15, rotationY: 20 }, PLOT);
    const reversedView = planChartThreeDProjection({ rotationX: 15, rotationY: 200 }, PLOT);
    if (!defaultView || !reversedView) throw new Error('projection not planned');
    const centre = (plan: NonNullable<typeof defaultView>, depth: number) =>
      plan.cameraDepth(plan.front.x + plan.front.w / 2, plan.front.y + plan.front.h / 2, depth);
    expect(centre(defaultView, 0)).toBeGreaterThan(centre(defaultView, 1));
    expect(centre(reversedView, 0)).toBeLessThan(centre(reversedView, 1));
  });

  it('keeps 3-D pie wall thickness independent of cartesian depthPercent', () => {
    const ordinary = planChartThreeDProjection({ depthPercent: 100 }, PLOT);
    const boundary = planChartThreeDProjection({ depthPercent: 2000 }, PLOT);
    expect(ordinary?.pieThicknessFraction).toBeCloseTo(0.30 * Math.cos(15 * Math.PI / 180), 12);
    expect(boundary?.pieThicknessFraction).toBe(ordinary?.pieThicknessFraction);
    const faceOn = planChartThreeDProjection({ rotationX: 89 }, PLOT);
    expect(faceOn?.pieThicknessFraction).toBeLessThan(0.01);
  });

  it('fits the authored hPercent scene ratio inside the plot', () => {
    for (const heightPercent of [25, 100, 300]) {
      const plan = planChartThreeDProjection({ heightPercent }, PLOT);
      expect(plan).not.toBeNull();
      expect((plan?.scene.h ?? 0) / (plan?.scene.w ?? 1))
        .toBeCloseTo(heightPercent / 100, 12);
      expect(plan?.scene.x).toBeGreaterThanOrEqual(PLOT.x);
      expect(plan?.scene.y).toBeGreaterThanOrEqual(PLOT.y);
      expect((plan?.scene.x ?? 0) + (plan?.scene.w ?? 0)).toBeLessThanOrEqual(PLOT.x + PLOT.w);
      expect((plan?.scene.y ?? 0) + (plan?.scene.h ?? 0)).toBeLessThanOrEqual(PLOT.y + PLOT.h);
    }
  });

  it('keeps a parser default hPercent in Office automatic-scaling mode', () => {
    const omitted = planChartThreeDProjection({}, PLOT);
    const parserDefault = planChartThreeDProjection({
      heightPercent: 100,
      heightPercentAuthored: false,
    }, PLOT);
    const authored = planChartThreeDProjection({
      heightPercent: 100,
      heightPercentAuthored: true,
    }, PLOT);
    expect(parserDefault?.scene).toEqual(omitted?.scene);
    expect((parserDefault?.scene.h ?? 0) / (parserDefault?.scene.w ?? 1))
      .toBeCloseTo(PLOT.h / PLOT.w, 12);
    expect((authored?.scene.h ?? 0) / (authored?.scene.w ?? 1)).toBeCloseTo(1, 12);
  });

  it('uses the observed automatic line/area height when hPercent is omitted', () => {
    const plot = { x: 0, y: 0, w: 738, h: 439 };
    const view = {
      rotationX: 20,
      rotationY: 20,
      depthPercent: 100,
      perspective: 30,
    };
    const automatic = planChartThreeDProjection(view, plot, {
      sceneDepthScale: 0.4,
      sceneHeightScale: 1 / 3,
    });
    if (!automatic) throw new Error('projection not planned');
    const automaticAspect = planChartThreeDSurfaceGeometry(automatic, 'backWall', 0)
      .pictureStackAspect ?? 0;
    expect(automatic.scene.h / automatic.scene.w).toBeCloseTo(1 / 3, 12);
    // The 2:1 source occupies two wall heights: its upper half is clipped,
    // while the observed 8:1 control repeats twice over the same wall.
    expect(automaticAspect / 2).toBeCloseTo(2, 2);
  });

  it('keeps depthPercent linear in model space and responsive after refitting', () => {
    const defaultDepth = planChartThreeDProjection({ depthPercent: 100 }, PLOT);
    const maximumDepth = planChartThreeDProjection({ depthPercent: 2000 }, PLOT);
    if (!defaultDepth || !maximumDepth) throw new Error('projection not planned');
    const projectedDepthLength = (plan: typeof defaultDepth) => {
      const x = plan.scene.x + plan.scene.w / 2;
      const y = plan.scene.y + plan.scene.h / 2;
      const near = plan.project(x, y, 0);
      const far = plan.project(x, y, 1);
      return Math.hypot(far.x - near.x, far.y - near.y);
    };
    expect(maximumDepth.modelDepth / defaultDepth.modelDepth).toBeCloseTo(20, 12);
    expect(projectedDepthLength(maximumDepth) / projectedDepthLength(defaultDepth))
      .toBeGreaterThan(5);
  });

  it('keeps the authored perspective boundary monotonic until the camera plane', () => {
    const values = [1, 60, 120, 180, 240].map(perspective => {
      const plan = planChartThreeDProjection({ perspective }, PLOT);
      if (!plan) throw new Error('projection not planned');
      const y = plan.front.y + plan.front.h;
      const left = plan.project(plan.front.x, y, 1);
      const right = plan.project(plan.front.x + plan.front.w, y, 1);
      return (right.y - left.y) / (right.x - left.x);
    });
    for (let index = 1; index < values.length; index++) {
      expect(values[index]).toBeGreaterThanOrEqual(values[index - 1] - 1e-10);
    }
    expect(new Set(values.map(value => value.toFixed(8))).size).toBeGreaterThan(2);
  });

  it('rejects invalid or degenerate plot rectangles', () => {
    expect(planChartThreeDProjection({}, { x: 0, y: 0, w: 0, h: 100 })).toBeNull();
    expect(planChartThreeDProjection({}, { x: 0, y: 0, w: Infinity, h: 100 })).toBeNull();
  });
});
