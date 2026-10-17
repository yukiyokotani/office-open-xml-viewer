import { describe, expect, it } from 'vitest';
import { renderPreparedWorksheetViewport } from './index.js';
import type { Styles, Worksheet } from './types.js';

const styles: Styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] };

function worksheet(overrides: Partial<Worksheet> = {}): Worksheet {
  return {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 1, freezeCols: 1, conditionalFormats: [],
    images: [], charts: [], shapeGroups: [], ...overrides,
  };
}

function canvas(): HTMLCanvasElement {
  const target = { width: 1, height: 1 };
  const state: Record<string, unknown> = {
    canvas: target,
    measureText: (text: string) => ({ width: text.length * 7 }),
    createLinearGradient: () => ({ addColorStop() {} }),
  };
  const context = new Proxy(state, {
    get: (object, key: string) => object[key] ?? (() => undefined),
    set: (object, key: string, value) => { object[key] = value; return true; },
  });
  return Object.assign(target, { getContext: () => context }) as unknown as HTMLCanvasElement;
}

describe('prepared worksheet viewport', () => {
  it('counts frozen bands in the cell budget and rejects before changing the canvas', async () => {
    const target = canvas();
    const ws = worksheet();
    const options = { width: 800, height: 600, scrollX: 200, scrollY: 400 };
    await expect(renderPreparedWorksheetViewport(target, ws, styles, { ...options, maxVisibleCells: 1 }))
      .rejects.toMatchObject({ code: 'ooxml-resource-limit', details: {
        violation: { resource: 'viewport-cells', limit: 1 },
      } });
    expect([target.width, target.height]).toEqual([1, 1]);
    const frame = await renderPreparedWorksheetViewport(target, ws, styles, options);
    const count = (frame.range.rows + 1) * (frame.range.cols + 1);
    await expect(renderPreparedWorksheetViewport(target, ws, styles, { ...options, maxVisibleCells: count }))
      .resolves.toBeDefined();
    await expect(renderPreparedWorksheetViewport(target, ws, styles, { ...options, maxVisibleCells: count - 1 }))
      .rejects.toMatchObject({ details: { violation: { observed: count } } });
  });

  it('uses painted geometry for merged cells, hidden rows and frozen-pane hit testing', async () => {
    const ws = worksheet({
      mergeCells: [{ top: 3, bottom: 4, left: 3, right: 4 }],
      rowHeights: { 2: 0, 3: 45 },
    });
    const frame = await renderPreparedWorksheetViewport(canvas(), ws, styles,
      { width: 800, height: 600, scrollX: 50, scrollY: 10 });
    const rect = frame.cellRect(4, 4)!;
    expect(frame.cellAt(rect.x + rect.w / 2, rect.y + rect.h / 2)).toEqual({ row: 3, col: 3 });
    expect(frame.cellRect(2, 1)).toBeNull();
    const frozen = frame.cellRect(1, 1)!;
    expect(frame.cellAt(frozen.x + 1, frozen.y + 1)).toEqual({ row: 1, col: 1 });
    expect(frame.cellAt(0, 0)).toBeNull();
  });
});
