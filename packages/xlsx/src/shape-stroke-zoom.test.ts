import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { ShapeInfo, Styles, Worksheet } from './types.js';

// ECMA-376 §20.1.2.2.24 `a:ln@w` is an EMU length and the §20.1.8.38 /
// §20.1.8.57 line-end sizes are multiples of it: both scale with the sheet
// zoom (`cellScale`) like the shape geometry, as in Excel (issue #2251).

type Point = { x: number; y: number };
type Matrix = { a: number; b: number; c: number; d: number; e: number; f: number };
const multiply = (m: Matrix, n: Matrix): Matrix => ({
  a: m.a * n.a + m.c * n.b, b: m.b * n.a + m.d * n.b,
  c: m.a * n.c + m.c * n.d, d: m.b * n.c + m.d * n.d,
  e: m.a * n.e + m.c * n.f + m.e, f: m.b * n.e + m.d * n.f + m.f,
});

const RED = 'rgba(255,0,0,1)';
/** 3 pt: above the 2 pt line-end base floor, so the decoration unit is the line width. */
const STROKE_EMU = 38100;
const TRIANGLE = { type: 'triangle', w: 'lg', len: 'lg' };

/** Paint one shape over a 3-column × 3-row anchor at `cellScale`; record its outline paint. */
function record(shape: Partial<ShapeInfo>, cellScale?: number) {
  const strokes: { lineWidth: number; dash: number[]; points: Point[] }[] = [];
  const fills: Point[][] = [];
  let subpaths: Point[][] = [];
  let dash: number[] = [];
  const stack: Matrix[] = [];
  let current: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const at = (x: number, y: number): Point => ({ x: current.a * x + current.c * y + current.e, y: current.b * x + current.d * y + current.f });
  const pen = (p: Point) => { if (subpaths.length === 0) subpaths.push([]); subpaths[subpaths.length - 1].push(p); };
  const state: Record<string, unknown> = { globalAlpha: 1, lineWidth: 1, lineCap: 'butt', lineJoin: 'miter', miterLimit: 10 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (typeof prop === 'string' && prop in target) return target[prop];
      switch (prop) {
        case 'canvas': return { width: 800, height: 600 };
        case 'measureText': return () => ({ width: 7 });
        case 'getTransform': return () => current;
        case 'getLineDash': return () => dash;
        case 'setLineDash': return (d: number[]) => { dash = [...d]; };
        case 'save': return () => { stack.push(current); };
        case 'restore': return () => { current = stack.pop() ?? current; };
        case 'translate': return (x: number, y: number) => { current = multiply(current, { a: 1, b: 0, c: 0, d: 1, e: x, f: y }); };
        case 'scale': return (x: number, y: number) => { current = multiply(current, { a: x, b: 0, c: 0, d: y, e: 0, f: 0 }); };
        case 'rotate': return (r: number) => {
          current = multiply(current, { a: Math.cos(r), b: Math.sin(r), c: -Math.sin(r), d: Math.cos(r), e: 0, f: 0 });
        };
        case 'setTransform': return (a: number | Matrix = 1, b = 0, c = 0, d = 1, e = 0, f = 0) => {
          current = typeof a === 'object' ? { ...a } : { a, b, c, d, e, f };
        };
        case 'createPattern': return () => ({ setTransform() {} });
        case 'beginPath': return () => { subpaths = []; };
        case 'moveTo': return (x: number, y: number) => { subpaths.push([at(x, y)]); };
        case 'lineTo': return (x: number, y: number) => { pen(at(x, y)); };
        case 'rect': return (x: number, y: number, w: number, h: number) => {
          subpaths.push([at(x, y), at(x + w, y), at(x + w, y + h), at(x, y + h)]);
        };
        case 'stroke': return () => {
          if (target.strokeStyle !== RED) return;
          // `lineWidth` is in the current user space; report it in device px.
          const unit = Math.hypot(current.a, current.b);
          for (const points of subpaths) {
            strokes.push({ lineWidth: (target.lineWidth as number) * unit, dash: dash.map(v => v * unit), points });
          }
        };
        case 'fill': return () => { if (target.fillStyle === RED) fills.push(...subpaths.map(p => [...p])); };
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
    { row: 1, col: 1, rows: 10, cols: 10 }, cellScale === undefined ? {} : { cellScale });
  return { strokes, fills };
}

const distance = (a: Point, b: Point) => Math.hypot(a.x - b.x, a.y - b.y);

/** The tail triangle's length along the line and full width across it. */
function triangleSize(points: Point[]) {
  expect(points).toHaveLength(3);
  const [tip, b1, b2] = points;
  const mid = { x: (b1.x + b2.x) / 2, y: (b1.y + b2.y) / 2 };
  return { length: distance(tip, mid), width: distance(b1, b2) };
}

const connectors: [string, Partial<ShapeInfo>][] = [
  ['preset connector', { strokeTailEnd: TRIANGLE } as Partial<ShapeInfo>],
  ['custom-geometry connector', {
    strokeTailEnd: TRIANGLE,
    geom: { type: 'custom', paths: [{ w: 100, h: 100, commands: [
      { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 100, y: 100 },
    ] }] },
  } as Partial<ShapeInfo>],
];

describe('XLSX shape outline zoom (issue #2251)', () => {
  it('scales the line width and its line-end size with the sheet zoom', () => {
    for (const [name, shape] of connectors) {
      const base = record(shape, 1);
      // 3 pt = 4 px at 100%; lg × lg triangle = 5 × line width each way.
      expect(base.strokes.map(s => s.lineWidth), name).toEqual([4]);
      expect(triangleSize(base.fills[0]).length, name).toBeCloseTo(20, 9);
      expect(triangleSize(base.fills[0]).width, name).toBeCloseTo(20, 9);
      for (const zoom of [0.5, 2]) {
        const zoomed = record(shape, zoom);
        expect(zoomed.strokes.map(s => s.lineWidth), `${name} @${zoom}`).toEqual([4 * zoom]);
        expect(zoomed.fills, `${name} @${zoom}`).toHaveLength(1);
        const size = triangleSize(zoomed.fills[0]);
        expect(size.length, `${name} @${zoom}`).toBeCloseTo(20 * zoom, 9);
        expect(size.width, `${name} @${zoom}`).toBeCloseTo(20 * zoom, 9);
        // The leader still stops where the (scaled) decoration starts: the gap
        // to the tip keeps its proportion to the line.
        const leader = zoomed.strokes[0].points;
        const baseLeader = base.strokes[0].points;
        expect(distance(leader.at(-1)!, zoomed.fills[0][0]), `${name} @${zoom}`)
          .toBeCloseTo(distance(baseLeader.at(-1)!, base.fills[0][0]) * zoom, 6);
      }
    }
  });

  it('scales a body outline and its dash pattern with the zoom', () => {
    const dashed: Partial<ShapeInfo> = { geom: { type: 'preset', name: 'rect', adj: [] }, strokeDashStyle: 'dash' };
    const base = record(dashed, 1);
    expect(base.strokes.length).toBeGreaterThan(0);
    expect(base.strokes[0].lineWidth).toBe(4);
    for (const zoom of [0.5, 2]) {
      const zoomed = record(dashed, zoom);
      expect(zoomed.strokes[0].lineWidth).toBe(4 * zoom);
      expect(zoomed.strokes[0].dash).toEqual(base.strokes[0].dash.map(v => v * zoom));
    }
  });

  it('keeps the 100% paint unchanged and the hairline floor zoom-independent', () => {
    for (const [name, shape] of connectors) {
      // The default (no cellScale) and an explicit 100% paint the same outline.
      expect(record(shape, 1), name).toEqual(record(shape));
    }
    // A 1 px (9525 EMU) line at 25% zoom would be 0.25 px: it keeps core's
    // 0.5 px minimum painted width (the same floor PPTX applies).
    const thin = record({ strokeWidth: 9525 }, 0.25);
    expect(thin.strokes.map(s => s.lineWidth)).toEqual([0.5]);
  });
});
