import { describe, expect, it } from 'vitest';
import {
  isObservedAutomaticSurfaceCamera,
  surfaceMaterialFactor,
  surfacePerspectiveTangentGain,
} from './material-color.js';
import { automaticSurfaceMajorUnit } from './axis-scale.js';

describe('surface automatic material', () => {
  it('is winding-invariant and bounded to the surface compatibility range', () => {
    const lit = surfaceMaterialFactor({ x: 0.2, y: 0.4, z: 0.9 });
    expect(surfaceMaterialFactor({ x: -0.2, y: -0.4, z: -0.9 })).toBeCloseTo(lit, 12);
    expect(lit).toBeGreaterThan(1);
    expect(surfaceMaterialFactor({ x: 0.2, y: -0.4, z: 0.1 })).toBeGreaterThanOrEqual(0.48);
  });

  it('lights camera-space upper-right facets more strongly than their mirrors', () => {
    expect(surfaceMaterialFactor({ x: 0.4, y: 0, z: 0.92 }))
      .toBeGreaterThan(surfaceMaterialFactor({ x: -0.4, y: 0, z: 0.92 }));
    expect(surfaceMaterialFactor({ x: 0, y: 0.4, z: 0.92 }))
      .toBeGreaterThan(surfaceMaterialFactor({ x: 0, y: -0.4, z: 0.92 }));
  });
});

describe('surface automatic camera compatibility', () => {
  const observedCamera = {
    rotationX: 15,
    rotationY: 20,
    rightAngleAxes: false,
    perspective: 30,
  };

  it('uses the observed perspective gain only for the effective omitted-view camera', () => {
    expect(isObservedAutomaticSurfaceCamera(observedCamera)).toBe(true);
    expect(surfacePerspectiveTangentGain(observedCamera)).toBe(2);
  });

  it.each([
    { ...observedCamera, rotationX: 30 },
    { ...observedCamera, rotationY: 45 },
    { ...observedCamera, rightAngleAxes: true },
    { ...observedCamera, perspective: 20 },
  ])('keeps authored camera projections outside the observed boundary', camera => {
    expect(isObservedAutomaticSurfaceCamera(camera)).toBe(false);
    expect(surfacePerspectiveTangentGain(camera)).toBe(1);
  });
});

describe('surface automatic major unit', () => {
  it('uses projected axis length and the compact edge-on floor', () => {
    expect(automaticSurfaceMajorUnit(0, 30, 220)).toBe(5);
    expect(automaticSurfaceMajorUnit(0, 40, 136)).toBe(10);
    expect(automaticSurfaceMajorUnit(10, 90, 0)).toBe(20);
    expect(automaticSurfaceMajorUnit(1, 32, 220)).toBe(5);
  });
});
