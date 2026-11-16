import { describe, expect, it } from 'vitest';
import { lineEndRetract } from '@silurus/ooxml-core';
import { renderViewport } from './renderer.js';
import type { ShapeInfo, Styles, Worksheet } from './types.js';

// ECMA-376 §20.5.2.13 xdr:cxnSp + §20.1.8.38 / §20.1.8.57 a:headEnd / a:tailEnd:
// a decorated leader is stroked once, stopping where its arrow head starts.

/** A 2D affine matrix with the DOMMatrix members the painter reads. */
function matrix(a: number, b: number, c: number, d: number, e: number, f: number) {
  return {
    a, b, c, d, e, f,
    multiply(m: { a: number; b: number; c: number; d: number; e: number; f: number }) {
      return matrix(
        a * m.a + c * m.b, b * m.a + d * m.b,
        a * m.c + c * m.d, b * m.c + d * m.d,
        a * m.e + c * m.f + e, b * m.e + d * m.f + f,
      );
    },
  };
}
type Matrix = ReturnType<typeof matrix>;
type Point = { x: number; y: number };

const RED = 'rgba(255,0,0,1)';
const TRIANGLE = { type: 'triangle', w: 'lg', len: 'lg' };
const STROKE_EMU = 19050;

/** Paint one shape over a 3-column × 3-row anchor; record its device-space paint. */
function record(shape: Partial<ShapeInfo>) {
  /** Stroked subpaths in the connector's paint, in device space. */
  const strokes: { points: Point[]; closed: boolean; style: unknown }[] = [];
  /** Arrow heads: the device origin of the decoration frame (its tip) when filled. */
  const tips: Point[] = [];
  /** The transform each connector-paint fill / stroke was painted under. */
  const fillFrames: Matrix[] = [];
  const strokeFrames: Matrix[] = [];
  let subpaths: { points: Point[]; closed: boolean }[] = [];
  const pen = (p: Point) => {
    if (subpaths.length === 0) subpaths.push({ points: [], closed: false });
    subpaths[subpaths.length - 1].points.push(p);
  };
  const stack: Matrix[] = [];
  let current = matrix(1, 0, 0, 1, 0, 0);
  const at = (x: number, y: number): Point => ({ x: current.a * x + current.c * y + current.e, y: current.b * x + current.d * y + current.f });
  const isShapePaint = (style: unknown) => style === RED || (typeof style === 'object' && style !== null && 'gradient' in style);
  const state: Record<string, unknown> = { globalAlpha: 1, lineWidth: 1, lineCap: 'butt', lineJoin: 'miter', miterLimit: 10 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (typeof prop === 'string' && prop in target) return target[prop];
      switch (prop) {
        case 'canvas': return { width: 400, height: 300 };
        case 'measureText': return () => ({ width: 7 });
        case 'getTransform': return () => current;
        case 'getLineDash': return () => [];
        case 'save': return () => { stack.push(current); };
        case 'restore': return () => { current = stack.pop() ?? current; };
        case 'translate': return (x: number, y: number) => { current = current.multiply(matrix(1, 0, 0, 1, x, y)); };
        case 'scale': return (x: number, y: number) => { current = current.multiply(matrix(x, 0, 0, y, 0, 0)); };
        case 'rotate': return (r: number) => {
          current = current.multiply(matrix(Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0));
        };
        case 'setTransform': return (a: number | Matrix = 1, b = 0, c = 0, d = 1, e = 0, f = 0) => {
          current = typeof a === 'object' ? matrix(a.a, a.b, a.c, a.d, a.e, a.f) : matrix(a, b, c, d, e, f);
        };
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ gradient: true, addColorStop() {} });
        case 'createPattern': return () => ({ setTransform() {} });
        case 'beginPath': return () => { subpaths = []; };
        case 'moveTo': return (x: number, y: number) => { subpaths.push({ points: [at(x, y)], closed: false }); };
        case 'lineTo': return (x: number, y: number) => { pen(at(x, y)); };
        case 'bezierCurveTo': return (...args: number[]) => { pen(at(args[4], args[5])); };
        case 'quadraticCurveTo': return (...args: number[]) => { pen(at(args[2], args[3])); };
        case 'closePath': return () => { if (subpaths.length) subpaths[subpaths.length - 1].closed = true; };
        case 'stroke': return () => {
          if (isShapePaint(target.strokeStyle)) {
            for (const subpath of subpaths) strokes.push({ ...subpath, points: [...subpath.points], style: target.strokeStyle });
            strokeFrames.push(current);
          }
        };
        case 'fill': return () => {
          if (isShapePaint(target.fillStyle)) {
            tips.push({ x: current.e, y: current.f });
            fillFrames.push(current);
          }
        };
        default: return typeof prop === 'string' ? () => undefined : undefined;
      }
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const ws = {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], images: [],
    shapeGroups: [{
      fromCol: 1, fromColOff: 0, fromRow: 2, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 5, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      shapes: [{
        x: 0, y: 0, w: 1, h: 1, rot: 0, flipH: false, flipV: false,
        strokeColor: '#FF0000', strokeWidth: STROKE_EMU,
        geom: { type: 'preset', name: 'straightConnector1', adj: [] },
        ...shape,
      }],
    }],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
    { row: 1, col: 1, rows: 10, cols: 10 });
  // A callout's text box is a closed subpath; the leader is the open one.
  return { leaders: strokes.filter((s) => !s.closed && s.points.length >= 2), tips, fillFrames, strokeFrames };
}

/** Device px a tail triangle retracts the line (core's observed Office rule). */
const retract = lineEndRetract(TRIANGLE, { color: '#FF0000', width: STROKE_EMU }, 1 / 9525, 12700);
const distance = (a: Point, b: Point) => Math.hypot(a.x - b.x, a.y - b.y);
const preset = (name: string): Partial<ShapeInfo> => ({ geom: { type: 'preset', name, adj: [] } as ShapeInfo['geom'] });

describe('XLSX connector line ends', () => {
  it('strokes each decorated leader once, stopping at its arrow head', () => {
    const cases: [string, Partial<ShapeInfo>][] = [
      ['straight connector', preset('straightConnector1')],
      ['bent connector', preset('bentConnector3')],
      // Curved leaders are trimmed along the curve; the preset engine must not
      // also stroke the full-length curve under it.
      ['curved connector', preset('curvedConnector3')],
      // Callout leaders take line ends like connectors (shared core policy).
      ['callout leader', preset('borderCallout1')],
      // A custom-geometry connector ends on its last path's open end.
      ['custom-geometry connector', { geom: { type: 'custom', paths: [{ w: 100, h: 100, commands: [
        { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 50, y: 0 },
        { op: 'cubicBezTo', x1: 80, y1: 0, x2: 100, y2: 20, x3: 100, y3: 100 },
      ] }] } } as Partial<ShapeInfo>],
      // A closed subpath ahead of the open line does not hide the line's end.
      ['custom path with a closed subpath first', { geom: { type: 'custom', paths: [{ w: 100, h: 100, commands: [
        { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 10, y: 0 }, { op: 'lineTo', x: 10, y: 10 }, { op: 'close' },
        { op: 'moveTo', x: 0, y: 50 }, { op: 'lineTo', x: 100, y: 50 },
      ] }] } } as Partial<ShapeInfo>],
      // A flat custom-geometry connector (zero height) is still a stroked line.
      ['zero-height custom connector', { h: 0, geom: { type: 'custom', paths: [{ w: 100, h: 100, commands: [
        { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 100, y: 0 },
      ] }] } } as Partial<ShapeInfo>],
      // A painted (hatch) body fill must not cancel the leader trim.
      ['hatch-filled line', { ...preset('line'), fill: { fillType: 'hatch', color: '#00FF00', angle: 0, distancePt: 3, lines: 1 } } as Partial<ShapeInfo>],
    ];
    for (const [name, shape] of cases) {
      const { leaders, tips } = record({ ...shape, strokeTailEnd: TRIANGLE });
      expect(leaders, name).toHaveLength(1);
      expect(tips, name).toHaveLength(1);
      const gap = distance(leaders[0].points[leaders[0].points.length - 1], tips[0]);
      // Along a curve the chord to the tip is at most the trimmed arc length.
      expect(gap, name).toBeGreaterThan(retract - 0.25);
      expect(gap, name).toBeLessThanOrEqual(retract + 1e-9);
    }

    // flipH mirrors the leader and its decoration: the tail moves to the left.
    const plain = record({ strokeTailEnd: TRIANGLE });
    const flipped = record({ strokeTailEnd: TRIANGLE, flipH: true });
    const [start, end] = [plain.leaders[0].points[0], plain.leaders[0].points.at(-1)!];
    const [fStart, fEnd] = [flipped.leaders[0].points[0], flipped.leaders[0].points.at(-1)!];
    expect(plain.tips[0].x).toBeGreaterThan(end.x);
    expect(end.x).toBeGreaterThan(start.x);
    expect(flipped.tips[0].x).toBeLessThan(fEnd.x);
    expect(fEnd.x).toBeLessThan(fStart.x);
    expect(flipped.tips[0].x + plain.tips[0].x).toBeCloseTo(start.x + fStart.x, 6);
  });

  it('paints a gradient line and its ends in the gradient of the line frame', () => {
    const { leaders, fillFrames, strokeFrames } = record({
      strokeTailEnd: TRIANGLE,
      strokeFill: {
        fillType: 'gradient', angle: 0, gradType: 'linear',
        stops: [{ position: 0, color: '00FF00' }, { position: 1, color: '0000FF' }],
      } as ShapeInfo['strokeFill'],
    });
    expect(leaders).toHaveLength(1);
    expect(leaders[0].style).toMatchObject({ gradient: true });
    // The arrow head continues the line's brush: it is filled under the frame
    // the line was stroked in, not the translated / rotated decoration frame.
    expect(fillFrames).toHaveLength(1);
    const coefficients = ({ a, b, c, d, e, f }: Matrix) => [a, b, c, d, e, f];
    expect(coefficients(fillFrames[0])).toEqual(coefficients(strokeFrames[0]));
  });
});
