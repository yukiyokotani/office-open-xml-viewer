import { describe, expect, it } from 'vitest';
import { resolveDrawingMLGeometry, type DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
import type { DrawingLayout } from './types.js';
import { snapshotPlainData } from './plain-data.js';
import { deriveNativeWrapPaintExtent } from './native-wrap-paint-bounds.js';

const rectangle: DrawingMLShapePaintPlan = {
  rect: { x: 10, y: 20, w: 30, h: 40 },
  geometry: { kind: 'preset', name: 'rect', adjustments: [] },
  fill: { fillType: 'solid', color: 'FF0000' },
  stroke: { color: '000000', width: 2, lineJoin: 'miter', lineCap: 'butt' },
  transform: { rotationDeg: 0, flipH: false, flipV: false },
};

function drawing(plan: DrawingMLShapePaintPlan, extra: Partial<DrawingLayout> = {}): DrawingLayout {
  return snapshotPlainData({
    kind: 'drawing', id: 'invented-drawing', flowDomainId: 'invented-flow',
    flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 }, advancePt: 0, ordinaryFlow: false,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    commands: [{ kind: 'drawingml-shape', plan: { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) } }], ...extra,
  }, 'invented paint proof');
}

describe('retained native reading-policy paint extent', () => {
  it('includes the actual rectangle miter stroke and drawing owner transform', () => {
    const result = deriveNativeWrapPaintExtent(drawing(rectangle, {
      transform: { a: 0, b: 1, c: -1, d: 0, e: 100, f: 200 },
    }), 1000);
    expect(result.coordinateSpace).toBe('section-logical-pt');
    expect(result.bounds).toEqual({ xPt: 39, yPt: 209, widthPt: 42, heightPt: 32 });
  });

  it('includes arrow/cap protrusions from the retained rotated connector geometry', () => {
    const result = deriveNativeWrapPaintExtent(drawing({
      ...rectangle, rect: { x: 10, y: 20, w: 30, h: 0 },
      geometry: { kind: 'preset', name: 'line', adjustments: [] }, fill: null,
      stroke: { color: '000000', width: 2, lineCap: 'square', tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
      transform: { rotationDeg: 90, flipH: false, flipV: false },
    }), 1000);
    expect(result.bounds?.xPt).toBeLessThan(25);
    expect((result.bounds?.xPt ?? 0) + (result.bounds?.widthPt ?? 0)).toBeGreaterThan(25);
    expect(result.bounds?.yPt).toBeLessThan(5);
    expect((result.bounds?.yPt ?? 0) + (result.bounds?.heightPt ?? 0)).toBeGreaterThanOrEqual(35);
  });

  it('does not substitute the anchor frame for custom geometry outside it', () => {
    const result = deriveNativeWrapPaintExtent(drawing({
      ...rectangle, stroke: null,
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: -1, y: 0 }, { cmd: 'lineTo', x: 2, y: 0 },
        { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' },
      ]] },
    }), 1000);
    expect(result.bounds?.xPt).toBe(-20);
    expect((result.bounds?.xPt ?? 0) + (result.bounds?.widthPt ?? 0)).toBe(70);
  });

  it('rejects unproved textbox/resource/brush classes and work exhaustion atomically', () => {
    expect(() => deriveNativeWrapPaintExtent(drawing(rectangle, { textBoxIds: ['invented-textbox'] }), 1000)).toThrow(/owned textboxes/);
    expect(() => deriveNativeWrapPaintExtent(drawing(rectangle, {
      commands: [
        { kind: 'drawingml-shape', plan: { ...rectangle, resolvedGeometry: resolveDrawingMLGeometry(rectangle, 1) } },
        { kind: 'resource', resourceKey: 'invented-image', resourceKind: 'image', rect: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10 } },
      ],
    }), 1000)).toThrow(/cannot prove command/);
    expect(() => deriveNativeWrapPaintExtent(drawing({ ...rectangle,
      fill: { fillType: 'pattern', fg: '000000', bg: 'FFFFFF', preset: 'pct25' },
    }), 1000)).toThrow(/solid/);
    const source = drawing(rectangle);
    const before = JSON.stringify(source);
    expect(() => deriveNativeWrapPaintExtent(source, 1)).toThrow(/budget/);
    expect(JSON.stringify(source)).toBe(before);
    expect(() => deriveNativeWrapPaintExtent({ ...source }, 1000)).toThrow(/sealed/);
  });
});
