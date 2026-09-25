import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

describe('worksheet picture transforms', () => {
  it('wraps crop and alpha paint in the authored centre rotation and flips', () => {
    const ops: Array<{ name: string; args: unknown[]; alpha?: number }> = [];
    const state: Record<string, unknown> = { globalAlpha: 1 };
    const ctx = new Proxy(state, {
      get(target, prop) {
        if (prop === 'canvas') return { width: 400, height: 300 };
        if (prop === 'measureText') return () => ({ width: 7 });
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
          return () => ({ addColorStop() {} });
        }
        if (prop === 'drawImage') return (...args: unknown[]) => {
          ops.push({ name: 'drawImage', args, alpha: target.globalAlpha as number });
          if (target.throwDraw) throw new Error('synthetic draw failure');
        };
        if (typeof prop === 'string') return (...args: unknown[]) => ops.push({ name: prop, args });
      },
      set(target, prop, value) { target[String(prop)] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    const ws = {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
      freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], shapeGroups: [],
      images: [{
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/asymmetric.png', mimeType: 'image/png',
        rotation: 90, flipH: true, flipV: true,
        srcRect: { l: .1, t: .2, r: .3, b: .1 }, alpha: .4,
      }],
    } as Worksheet;

    renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      { loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]) });

    expect(ops.find((op) => op.name === 'rotate')?.args[0]).toBeCloseTo(Math.PI / 2, 12);
    expect(ops).toContainEqual({ name: 'scale', args: [-1, -1] });
    const draw = ops.find((op) => op.name === 'drawImage');
    expect(draw?.args).toHaveLength(9);
    expect(draw?.args.slice(1, 5)).toEqual([20, 20, 120, 70]);
    expect(draw?.args.slice(5)).toEqual([50, 22, 134, 40]);
    expect(draw?.alpha).toBe(.4);
    const rotatesAt = ops.findIndex((op) => op.name === 'rotate');
    expect(ops.slice(rotatesAt - 1, rotatesAt + 3)).toEqual([
      { name: 'translate', args: [117, 42] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'scale', args: [-1, -1] },
      { name: 'translate', args: [-117, -42] },
    ]);

    ops.length = 0;
    state.throwDraw = true;
    expect(() => renderViewport(
      ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      { loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]) },
    )).toThrow('synthetic draw failure');
    const failedDraw = ops.findIndex((op) => op.name === 'drawImage');
    expect(ops.slice(failedDraw + 1)).toEqual([
      { name: 'restore', args: [] }, // alpha frame
      { name: 'restore', args: [] }, // DrawingML transform frame
    ]);

    // Exercise the other reflection combination in a non-default viewport.
    // RTL changes placement only; the authored local horizontal reflection and
    // negative clockwise angle retain their DrawingML order around the new box.
    ops.length = 0;
    state.throwDraw = false;
    ws.rightToLeft = true;
    ws.images[0].rotation = -30;
    ws.images[0].flipH = true;
    ws.images[0].flipV = false;
    renderViewport(
      ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      {
        cellScale: 1.5, scrollOffsetX: 2, scrollOffsetY: 3,
        loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]),
      },
    );
    const secondDraw = ops.find((op) => op.name === 'drawImage');
    expect(secondDraw?.args.slice(1, 5)).toEqual([20, 20, 120, 70]);
    expect(secondDraw?.args.slice(5)).toEqual([126, 28.5, 202, 60]);
    const secondRotate = ops.findIndex((op) => op.name === 'rotate');
    expect(ops.slice(secondRotate - 1, secondRotate + 3)).toEqual([
      { name: 'translate', args: [227, 58.5] },
      { name: 'rotate', args: [-Math.PI / 6] },
      { name: 'scale', args: [-1, 1] },
      { name: 'translate', args: [-227, -58.5] },
    ]);
  });
});
