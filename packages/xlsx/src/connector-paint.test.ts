import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { ShapeInfo, Styles, Worksheet } from './types.js';

// ECMA-376 §20.5.2 xdr:cxnSp + §20.1.8.38/§20.1.8.57 a:headEnd/a:tailEnd:
// a connector is a stroked preset line whose ends may carry arrow decorations.
function record(shape: Partial<ShapeInfo>, toRow = 2) {
  const ops: string[] = [];
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop === 'canvas') return { width: 400, height: 300 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      if (prop === 'fill' || prop === 'stroke') {
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
      fromCol: 0, fromColOff: 0, fromRow: 2, fromRowOff: 0,
      toCol: 3, toColOff: 0, toRow: toRow, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      shapes: [{
        x: 0, y: 0, w: 1, h: 1, rot: 0, flipH: false, flipV: false,
        strokeColor: '#FF0000', strokeWidth: 19050,
        geom: { type: 'preset', name: 'straightConnector1', adj: [] },
        ...shape,
      }],
    }],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
    { row: 1, col: 1, rows: 10, cols: 10 });
  // Grid and cell paint use other colours; keep the connector's own red paint.
  return ops.filter((op) => op === 'stroke:rgba(255,0,0,1)' || op === 'fill:rgba(255,0,0,1)');
}


describe('XLSX connector paint', () => {
  it('paints a filled arrow head at a decorated end', () => {
    const ops = record({ strokeTailEnd: { type: 'triangle', w: 'med', len: 'med' } }, 2);
    expect(ops).toContain('fill:rgba(255,0,0,1)');
    expect(ops).toContain('stroke:rgba(255,0,0,1)');
  });

  it('paints one decoration per decorated end', () => {
    const one = record({ strokeTailEnd: { type: 'triangle', w: 'med', len: 'med' } }, 2);
    const both = record({
      strokeHeadEnd: { type: 'triangle', w: 'med', len: 'med' },
      strokeTailEnd: { type: 'triangle', w: 'med', len: 'med' },
    }, 2);
    expect(both.filter((op) => op.startsWith('fill:'))).toHaveLength(2 * one.filter((op) => op.startsWith('fill:')).length);
  });

  it('does not decorate an end of type none', () => {
    expect(record({ strokeTailEnd: { type: 'none', w: 'med', len: 'med' } }, 2)).toEqual(['stroke:rgba(255,0,0,1)']);
  });
});
