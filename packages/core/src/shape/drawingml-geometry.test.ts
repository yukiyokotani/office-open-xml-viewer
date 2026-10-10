import { describe, expect, it } from 'vitest';
import type { DrawingMLShapePaintPlan } from './drawingml-shape';
import { drawingMLGeometryBounds, resolveDrawingMLGeometry, requireResolvedDrawingMLGeometry } from './drawingml-geometry';
import { GeometryWorkBudgetError } from './path-data';

const rectangle: DrawingMLShapePaintPlan = {
  rect: { x: 10, y: 20, w: 30, h: 40 },
  geometry: { kind: 'preset', name: 'rect', adjustments: [] },
  fill: { fillType: 'solid', color: 'FF0000' }, stroke: { color: '000000', width: 2 },
  transform: { rotationDeg: 0, flipH: false, flipV: false },
};
function retained(plan: DrawingMLShapePaintPlan) {
  return { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) };
}
describe('shared retained DrawingML geometry', () => {
  it('excludes unfilled/unstroked paths from enclosure without discarding their source order', () => {
    const plan = retained({ ...rectangle, stroke: null,
      geometry: { kind: 'custom', subpaths: [
        [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1, y: 0 }, { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' }],
        [{ cmd: 'moveTo', x: -100, y: -100 }, { cmd: 'lineTo', x: 100, y: 100 }],
      ], paint: [{}, { fill: 'none', stroke: false }] },
    });
    expect(plan.resolvedGeometry.paths).toHaveLength(2);
    expect(drawingMLGeometryBounds(plan)).toEqual({ x: 10, y: 20, w: 30, h: 40 });
    expect(plan.resolvedGeometry.fillSilhouette).toHaveLength(4);
  });
  it('retains the retracted leader and full arrow body rather than the anchor rectangle', () => {
    const plan = retained({ ...rectangle, rect: { x: 10, y: 20, w: 30, h: 0 }, fill: null,
      geometry: { kind: 'preset', name: 'line', adjustments: [] },
      stroke: { color: '000000', width: 2, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
    });
    const leader = plan.resolvedGeometry.paths.find(p => p.stroke !== null)!;
    expect(leader.path).toEqual([{ op: 'moveTo', args: [10, 20] }, { op: 'lineTo', args: [24, 20] }]);
    expect(drawingMLGeometryBounds(plan)).toEqual({ x: 10, y: 12, w: 30, h: 16 });
  });
  it('keeps origin translation compatible with a cloned immutable local plan', () => {
    const original = retained(rectangle), copy = structuredClone(original);
    const moved = { ...copy, rect: { ...copy.rect, x: 110, y: 220 } };
    expect(requireResolvedDrawingMLGeometry(moved, 1)).toEqual(original.resolvedGeometry);
    expect(drawingMLGeometryBounds(moved)).toEqual({ x: 109, y: 219, w: 32, h: 42 });
    expect(drawingMLGeometryBounds(original, { a: 0, b: 1, c: -1, d: 0, e: 100, f: 200 })).toEqual({ x: 39, y: 209, w: 42, h: 32 });
    expect(() => requireResolvedDrawingMLGeometry({ ...copy, rect: { ...copy.rect, w: 31 } }, 1)).toThrow(/stale/);
  });
  it('fails geometry work atomically and refuses nonfinite resolved commands', () => {
    const before = JSON.stringify(rectangle);
    expect(() => resolveDrawingMLGeometry(rectangle, 1, 1)).toThrow(GeometryWorkBudgetError);
    expect(JSON.stringify(rectangle)).toBe(before);
    const acquired = retained(rectangle);
    expect(() => requireResolvedDrawingMLGeometry({ ...acquired, resolvedGeometry: { ...acquired.resolvedGeometry, workUnits: 1 } }, 1)).toThrow(GeometryWorkBudgetError);
    expect(() => requireResolvedDrawingMLGeometry({ ...acquired, rect: { ...acquired.rect, x: NaN } }, 1)).toThrow(/frame/);
    expect(() => resolveDrawingMLGeometry({ ...rectangle, geometry: { kind: 'custom', subpaths: [[{ cmd: 'moveTo', x: NaN, y: 0 }]] } }, 1)).toThrow(/Non-finite/);
  });
});
