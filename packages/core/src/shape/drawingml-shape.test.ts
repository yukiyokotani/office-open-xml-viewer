import { describe, expect, it } from 'vitest';
import {
  clipDrawingMLShape,
  paintDrawingMLShape,
  type DrawingMLShapePaintPlan,
} from './drawingml-shape.js';

function recordingContext() {
  const operations: Array<{ name: string; args: unknown[] }> = [];
  let lineWidth = 1;
  let lineCap: CanvasLineCap = 'butt';
  let fillStyle: string | CanvasGradient | CanvasPattern = '';
  let strokeStyle: string | CanvasGradient | CanvasPattern = '';
  const gradient = {
    addColorStop(...args: unknown[]) { operations.push({ name: 'addColorStop', args }); },
  } as unknown as CanvasGradient;
  const ctx = {
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      fillStyle = value;
      operations.push({ name: 'fillStyle', args: [value] });
    },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string | CanvasGradient | CanvasPattern) {
      strokeStyle = value;
      operations.push({ name: 'strokeStyle', args: [value] });
    },
    get lineWidth() { return lineWidth; },
    set lineWidth(value: number) { lineWidth = value; operations.push({ name: 'lineWidth', args: [value] }); },
    get lineCap() { return lineCap; },
    set lineCap(value: CanvasLineCap) { lineCap = value; operations.push({ name: 'lineCap', args: [value] }); },
    save() { operations.push({ name: 'save', args: [] }); },
    restore() { operations.push({ name: 'restore', args: [] }); },
    translate(...args: unknown[]) { operations.push({ name: 'translate', args }); },
    rotate(...args: unknown[]) { operations.push({ name: 'rotate', args }); },
    scale(...args: unknown[]) { operations.push({ name: 'scale', args }); },
    beginPath() { operations.push({ name: 'beginPath', args: [] }); },
    rect(...args: unknown[]) { operations.push({ name: 'rect', args }); },
    moveTo(...args: unknown[]) { operations.push({ name: 'moveTo', args }); },
    lineTo(...args: unknown[]) { operations.push({ name: 'lineTo', args }); },
    bezierCurveTo(...args: unknown[]) { operations.push({ name: 'bezierCurveTo', args }); },
    ellipse(...args: unknown[]) { operations.push({ name: 'ellipse', args }); },
    closePath() { operations.push({ name: 'closePath', args: [] }); },
    clip(...args: unknown[]) { operations.push({ name: 'clip', args }); },
    fill(...args: unknown[]) { operations.push({ name: 'fill', args }); },
    stroke() { operations.push({ name: 'stroke', args: [] }); },
    setLineDash(...args: unknown[]) { operations.push({ name: 'setLineDash', args }); },
    createLinearGradient(...args: unknown[]) {
      operations.push({ name: 'createLinearGradient', args }); return gradient;
    },
    createRadialGradient(...args: unknown[]) {
      operations.push({ name: 'createRadialGradient', args }); return gradient;
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, operations, gradient };
}

describe('shared DrawingML shape painter', () => {
  it('renders preset geometry, gradients, transforms, and point-width strokes once', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'preset', name: 'rect', adjustments: [] },
      fill: {
        fillType: 'gradient', angle: 0, gradType: 'linear',
        stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '0000FF' }],
      },
      stroke: { color: '000000', width: 2, dashStyle: 'dash', lineCap: 'round' },
      transform: { rotationDeg: 90, flipH: true, flipV: false },
    };
    const { ctx, operations } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'translate', args: [60, 45] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'scale', args: [-1, 1] },
      { name: 'createLinearGradient', args: [10, 45, 110, 45] },
      { name: 'addColorStop', args: [0, 'rgba(255,0,0,1)'] },
      { name: 'addColorStop', args: [1, 'rgba(0,0,255,1)'] },
      { name: 'lineWidth', args: [2] },
      { name: 'lineCap', args: ['round'] },
    ]));
  });

  it('renders custom cubic geometry and its terminal arrows', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: .25, y1: 0, x2: .75, y2: 1, x: 1, y: 1 },
      ]] },
      fill: null,
      stroke: {
        color: '123456', width: 1, lineCap: 'square',
        headEnd: { type: 'arrow', w: 'sm', len: 'sm' },
        tailEnd: { type: 'triangle', w: 'med', len: 'lg' },
      },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx, operations } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'moveTo', args: [10, 20] },
      { name: 'bezierCurveTo', args: [35, 20, 85, 70, 110, 70] },
      { name: 'translate', args: [10, 20] },
      { name: 'translate', args: [110, 70] },
    ]));
  });

  it('uses an authored DrawingML gradient as the line paint', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: null,
      stroke: {
        color: 'AABBCC',
        width: 2,
        lineCap: 'round',
        fill: {
          fillType: 'gradient',
          angle: 90,
          gradType: 'linear',
          stops: [
            { position: 0.19, color: '11223300' },
            { position: 1, color: 'AABBCC' },
          ],
        },
      },
      transform: { rotationDeg: 0, flipH: true, flipV: false },
    };
    const { ctx, operations } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    const gradientOp = operations.find(({ name }) => name === 'createLinearGradient');
    expect(gradientOp?.args).toHaveLength(4);
    expect(gradientOp?.args.map(Number)).toEqual([
      expect.closeTo(60),
      expect.closeTo(20),
      expect.closeTo(60),
      expect.closeTo(70),
    ]);
    expect(operations).toEqual(expect.arrayContaining([
      { name: 'addColorStop', args: [0.19, 'rgba(17,34,51,0)'] },
      { name: 'addColorStop', args: [1, 'rgba(170,187,204,1)'] },
      { name: 'lineCap', args: ['round'] },
      { name: 'stroke', args: [] },
    ]));
  });

  it('uses the same resolved gradient for a connector shaft and arrow head', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'preset', name: 'line', adjustments: [] },
      fill: null,
      stroke: {
        color: 'AABBCC',
        width: 2,
        tailEnd: { type: 'triangle', w: 'med', len: 'med' },
        fill: {
          fillType: 'gradient',
          angle: 0,
          gradType: 'linear',
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'AABBCC' },
          ],
        },
      },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx, operations, gradient } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    expect(operations).toContainEqual({ name: 'strokeStyle', args: [gradient] });
    expect(operations).toContainEqual({ name: 'fillStyle', args: [gradient] });
  });

  it('preserves arbitrary adjustment counts for preset geometry', () => {
    const adjustments = Array.from({ length: 12 }, (_, index) => index * 1000);
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 0, y: 0, w: 100, h: 100 },
      geometry: { kind: 'preset', name: 'rect', adjustments },
      fill: { fillType: 'solid', color: 'FFFFFF' }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx } = recordingContext();

    expect(() => paintDrawingMLShape(ctx, plan, 1)).not.toThrow();
    expect(plan.geometry.kind === 'preset' && plan.geometry.adjustments).toHaveLength(12);
  });

  it('clips overlapping custom subpaths with the normal nonzero rule', () => {
    const { ctx, operations } = recordingContext();
    clipDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 100, h: 100 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'lineTo', x: 1, y: 0 },
        { cmd: 'lineTo', x: 1, y: 1 },
        { cmd: 'close' },
      ], [
        { cmd: 'moveTo', x: .25, y: .25 },
        { cmd: 'lineTo', x: .75, y: .25 },
        { cmd: 'lineTo', x: .75, y: .75 },
        { cmd: 'close' },
      ]] },
      fill: null,
      stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    });

    expect(operations.filter(({ name }) => name === 'clip')).toEqual([
      { name: 'clip', args: [] },
    ]);
  });

  // ECMA-376 §20.1.9.15: custom paths are filled and stroked on their own.
  it('honours per-path fill modes and stroke flags of custom geometry', () => {
    const square = (at: number) => [
      { cmd: 'moveTo' as const, x: at, y: at },
      { cmd: 'lineTo' as const, x: at + .2, y: at },
      { cmd: 'lineTo' as const, x: at + .2, y: at + .2 },
      { cmd: 'close' as const },
    ];
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 0, y: 0, w: 100, h: 100 },
      geometry: {
        kind: 'custom',
        subpaths: [square(0), square(.3), square(.6)],
        paint: [{ fill: 'none' }, { stroke: false }, { fill: 'darken' }],
      },
      fill: { fillType: 'solid', color: 'FF0000' },
      stroke: { color: '0000FF', width: 1 },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx, operations } = recordingContext();
    paintDrawingMLShape(ctx, plan, 1);
    const paints = operations
      .filter(({ name }) => name === 'fill' || name === 'stroke' || name === 'beginPath')
      .map(({ name }) => name);
    expect(paints).toEqual([
      'beginPath', 'stroke',
      'beginPath', 'fill',
      'beginPath', 'fill', 'fill', 'stroke',
    ]);
    expect(operations.filter(({ name }) => name === 'fillStyle').map(({ args }) => args[0]))
      .toContain('rgba(0,0,0,0.4)');

    const clip = recordingContext();
    clipDrawingMLShape(clip.ctx, plan);
    expect(clip.operations.filter(({ name }) => name === 'moveTo')).toHaveLength(2);
  });
});
