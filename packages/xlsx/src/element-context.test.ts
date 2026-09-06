import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import type { Worksheet } from './types.js';
import { hitTestXlsxElementContext, projectXlsxElementContext } from './element-context.js';

function worksheet(): Worksheet {
  const chart = {
    chartType: 'bar',
    title: 'Revenue',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'Actual', values: [10, 20] }],
  } as ChartModel;
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      imagePath: 'xl/media/image1.png', mimeType: 'image/png',
    }],
    shapeGroups: [],
    charts: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
      chart,
    }],
  } as Worksheet;
}

const viewport = {
  width: 800,
  height: 600,
  cellScale: 1,
  viewport: { row: 1, col: 1, rows: 30, cols: 12 },
  scrollOffsetX: 0,
  scrollOffsetY: 0,
  freezeRows: 0,
  freezeCols: 0,
};

describe('hitTestXlsxElementContext', () => {
  it('returns the topmost chart with bounded AI-readable data', () => {
    expect(hitTestXlsxElementContext(worksheet(), 0, { x: 100, y: 80 }, viewport)).toMatchObject({
      format: 'xlsx',
      kind: 'element',
      sheetIndex: 0,
      sheetName: 'Sheet1',
      elementType: 'chart',
      elementIndex: 0,
      seriesCount: 1,
      text: 'Chart type: bar\nTitle: Revenue\nCategories: Q1, Q2\nSeries Actual: 10, 20',
      anchor: {
        from: { row: 1, col: 1 },
        to: { row: 9, col: 5 },
      },
    });
  });

  it('falls through to an image and ignores the clipped/header area', () => {
    const ws = worksheet();
    ws.charts = [];
    expect(hitTestXlsxElementContext(ws, 0, { x: 100, y: 80 }, viewport)).toMatchObject({
      elementType: 'image',
      mimeType: 'image/png',
    });
    expect(hitTestXlsxElementContext(ws, 0, { x: 20, y: 10 }, viewport)).toBeNull();
  });

  it('hit-tests the rotated picture footprint rather than its unrotated box', () => {
    const ws = worksheet();
    ws.charts = [];
    ws.images[0].rotation = 90;
    expect(hitTestXlsxElementContext(ws, 0, { x: 180, y: 180 }, viewport))
      .toMatchObject({ elementType: 'image' });
    expect(hitTestXlsxElementContext(ws, 0, { x: 60, y: 80 }, viewport)).toBeNull();
  });

  it('does no worksheet-cell scan during an object hit', () => {
    const ws = worksheet();
    Object.defineProperty(ws, 'rows', {
      get(): never { throw new Error('cell rows must not be read'); },
    });
    expect(hitTestXlsxElementContext(ws, 0, { x: 100, y: 80 }, viewport)?.elementType)
      .toBe('chart');
  });

  it('projects retained chart focus into the clipped viewport without scanning cells', () => {
    const ws = worksheet();
    const context = hitTestXlsxElementContext(ws, 0, { x: 100, y: 80 }, viewport)!;
    const projection = projectXlsxElementContext(ws, context, viewport);

    expect(projection).toMatchObject({
      clip: { x: 50, y: 22, width: 750, height: 578 },
      rect: { x: 50, y: 22 },
      rotation: 0,
    });
    expect(projection!.rect.width).toBeGreaterThan(0);
    expect(projection!.rect.height).toBeGreaterThan(0);
  });
});
