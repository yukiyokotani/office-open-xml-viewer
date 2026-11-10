import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { ShapeHatchFill, Styles, Worksheet } from './types.js';

// An OpenDocument hatch (ODF 1.3 Part 3 §16.42) is stroked directly inside
// the shape's clip: no pattern bitmap is built, and the number of lines is
// bounded by the visible canvas, whatever the authored spacing and zoom.

/** A 2D affine matrix with the DOMMatrix members the painter reads. */
function matrix(a: number, b: number, c: number, d: number, e: number, f: number) {
  return {
    a, b, c, d, e, f,
    inverse() {
      const det = a * d - b * c;
      return matrix(d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det);
    },
    transformPoint(p: { x: number; y: number }) {
      return { x: a * p.x + c * p.y + e, y: b * p.x + d * p.y + f };
    },
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

/** The hatch colour, distinct from every grid and header colour. */
const HATCH = '#123456';

/** Paint one `size` × `size` px rectangle with `fill` on a 400 × 300 canvas. */
function paint(fill: ShapeHatchFill, size = 100) {
  const calls: string[] = [];
  let lines = 0;
  const stack: Matrix[] = [];
  let current = matrix(1, 0, 0, 1, 0, 0);
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop === 'canvas') return { width: 400, height: 300 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'getTransform') return () => current;
      if (prop === 'save') return () => { stack.push(current); };
      if (prop === 'restore') return () => { current = stack.pop() ?? current; };
      if (prop === 'translate') return (x: number, y: number) => { current = current.multiply(matrix(1, 0, 0, 1, x, y)); };
      if (prop === 'scale') return (x: number, y: number) => { current = current.multiply(matrix(x, 0, 0, y, 0, 0)); };
      if (prop === 'setTransform') {
        return (a = 1, b = 0, c = 0, d = 1, e = 0, f = 0) => { current = matrix(a, b, c, d, e, f); };
      }
      if (prop === 'resetTransform') return () => { current = matrix(1, 0, 0, 1, 0, 0); };
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
      if (prop === 'rotate') return (r: number) => { current = current.multiply(matrix(Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0)); };
      if (prop === 'fillRect') {
        return () => {
          if (String(target.fillStyle).toLowerCase() === HATCH) lines++;
          calls.push('fillRect');
        };
      }
      if (typeof prop === 'string') return (...args: unknown[]) => { calls.push(prop); void args; };
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const ws = {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], images: [],
    shapeGroups: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 0, toColOff: size * 9525, toRow: 0, toRowOff: size * 9525,
      nativeExtCx: 0, nativeExtCy: 0,
      shapes: [{
        x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
        fill,
        geom: { type: 'preset', name: 'rect' },
      }],
    }],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
    { row: 1, col: 1, rows: 10, cols: 10 });
  return { calls, lines };
}

const hatch = (distancePt: number, lines = 1, angle = 0): ShapeHatchFill => ({
  fillType: 'hatch', color: HATCH, angle, distancePt, lines,
});

describe('OpenDocument hatch fills', () => {
  it('draws lines at the authored spacing, one family per hatch style', () => {
    // 7.5 pt = 10 px: lines at 0, 10, …, 100 across the 100 px box.
    expect(paint(hatch(7.5)).lines).toBe(11);
    expect(paint(hatch(7.5, 2)).lines).toBe(22);
    expect(paint(hatch(7.5, 3)).lines).toBeGreaterThan(22);
  });

  it('never builds a bitmap, however wide the authored spacing', () => {
    // 10000 pt is a 13 333 px period: one line, no pattern, no canvas.
    const { calls, lines } = paint(hatch(10_000, 3, 30));
    expect(lines).toBeLessThanOrEqual(3);
    expect(calls).not.toContain('createPattern');
    expect(calls).not.toContain('drawImage');
  });

  it('bounds the lines by the visible canvas at any spacing and zoom', () => {
    // A 0.01 pt spacing would be 10 000 lines over the box; lines closer
    // than two device pixels merge, so the 100 px box holds at most ~51.
    expect(paint(hatch(0.01)).lines).toBeLessThanOrEqual(52);
    // A shape far larger than the canvas (as when zoomed far in): only the
    // visible part of the 400 × 300 canvas is hatched, the lines again two
    // device pixels apart at least. Covering the whole box would take
    // millions of lines.
    const bound = 2 * (Math.ceil(Math.hypot(400, 300) / 2) + 2);
    const huge = paint(hatch(0.01, 2), 10_000_000).lines;
    expect(huge).toBeGreaterThan(100);
    expect(huge).toBeLessThanOrEqual(bound);
  });
});
