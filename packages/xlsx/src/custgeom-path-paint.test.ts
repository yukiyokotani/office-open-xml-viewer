import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { PathInfo, Styles, Worksheet } from './types.js';

// ECMA-376 §20.1.9.15 a:path@fill / @stroke: each custom-geometry path is
// filled and stroked on its own. `fill="none"` leaves the path unfilled and
// `stroke="0"` unstroked; the shading modes add the shared overlay.
function recordAll(paths: PathInfo[]) {
  const ops: string[] = [];
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop === 'canvas') return { width: 400, height: 300 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      if (prop === 'fill' || prop === 'stroke' || prop === 'beginPath') {
        return () => ops.push(`${prop}:${String(prop === 'stroke' ? target.strokeStyle : target.fillStyle)}`);
      }
      if (typeof prop === 'string') return () => undefined;
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const ws = {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], images: [],
    shapeGroups: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      shapes: [{
        x: 0, y: 0, w: 1, h: 1, rot: 0,
        fillColor: '#FF0000', strokeColor: '#0000FF', strokeWidth: 12700,
        geom: { type: 'custom', paths },
      }],
    }],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
    { row: 1, col: 1, rows: 10, cols: 10 });
  return ops;
}

function recordCustomShape(paths: PathInfo[]) {
  // Grid lines and cell paint use other colours; keep the shape's own paint.
  return recordAll(paths).filter((op) => /^(fill|stroke):rgba\((255,0,0|0,0,255),1\)$/.test(op));
}

const square = (extra: Partial<PathInfo> = {}): PathInfo => ({
  w: 10, h: 10,
  commands: [
    { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 10, y: 0 },
    { op: 'lineTo', x: 10, y: 10 }, { op: 'close' },
  ],
  ...extra,
});

describe('XLSX custom geometry per-path paint', () => {
  it('fills and strokes every path by default', () => {
    const ops = recordCustomShape([square(), square()]);
    expect(ops.filter((op) => op.startsWith('fill:'))).toHaveLength(2);
    expect(ops.filter((op) => op.startsWith('stroke:'))).toHaveLength(2);
  });

  it('honours fill="none" and stroke="0" per path', () => {
    const ops = recordCustomShape([square({ fill: 'none' }), square({ stroke: false })]);
    expect(ops).toEqual(['stroke:rgba(0,0,255,1)', 'fill:rgba(255,0,0,1)']);
  });

  it('shades a lighten/darken path over the ordinary fill', () => {
    const ops = recordCustomShape([square({ fill: 'darken', stroke: false })]);
    expect(ops).toEqual(['fill:rgba(255,0,0,1)']);
    const all = recordAll([square({ fill: 'darken', stroke: false })]);
    const at = all.indexOf('fill:rgba(255,0,0,1)');
    expect(all[at + 1]).toBe('fill:rgba(0,0,0,0.4)');
  });
});
