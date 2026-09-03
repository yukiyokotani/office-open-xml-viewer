import { describe, expect, it, vi } from 'vitest';
import type { ChartExRenderer, ChartModel } from '@silurus/ooxml-core';
import {
  createPaintResourceRegistry,
} from '../layout/paint-resources.js';
import type { PaintResourceDescriptor } from '../layout/types.js';
import {
  createPaintResourceSession,
  unavailablePaintResourceHandle,
} from './resource-session.js';
import { createCanvasPaintResourcePainter } from './canvas-page.js';
import {
  canonicalCanvasPaintResourceHandlers,
  createCanonicalCanvasPaintResourceHandlers,
} from './canonical-resource-handlers.js';

function chartModel(): ChartModel {
  return {
    chartType: 'line', title: null, categories: [], series: [], varyColors: false,
    showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null,
    valAxisTitle: null, catAxisHidden: false, valAxisHidden: false,
    catAxisLineHidden: false, valAxisLineHidden: false, plotAreaBg: null,
    chartBg: 'ABCDEF', chartBorderColor: '123456', chartBorderWidthEmu: 12_700,
    showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'cross', catAxisMajorTickMark: 'cross',
    titleFontSizeHpt: null, titleFontColor: null, titleFontFace: null,
    catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
  } as ChartModel;
}

function recordingContext() {
  const operations: Array<{ name: string; args: unknown[]; alpha?: number }> = [];
  const alphaStack: number[] = [];
  let alpha = .8;
  const ctx = {
    get globalAlpha() { return alpha; },
    set globalAlpha(value: number) { alpha = value; },
    fillStyle: '', strokeStyle: '', lineWidth: 1, font: '',
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
    letterSpacing: '0px', fontKerning: 'auto',
    save() { alphaStack.push(alpha); operations.push({ name: 'save', args: [] }); },
    restore() { alpha = alphaStack.pop() ?? alpha; operations.push({ name: 'restore', args: [] }); },
    translate(...args: unknown[]) { operations.push({ name: 'translate', args }); },
    rotate(...args: unknown[]) { operations.push({ name: 'rotate', args }); },
    scale(...args: unknown[]) { operations.push({ name: 'scale', args }); },
    drawImage(...args: unknown[]) {
      operations.push({ name: 'drawImage', args, alpha });
    },
    fillRect(...args: unknown[]) { operations.push({ name: 'fillRect', args }); },
    strokeRect(...args: unknown[]) { operations.push({ name: 'strokeRect', args }); },
    fillText(...args: unknown[]) { operations.push({ name: 'fillText', args }); },
    setLineDash() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
    measureText() { throw new Error('resource adapter must not measure'); },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, operations };
}

function painter(
  descriptors: PaintResourceDescriptor[],
  handles: Array<{ resourceKey: string; kind: PaintResourceDescriptor['kind']; handle: unknown }>,
) {
  const registry = createPaintResourceRegistry(descriptors);
  return createCanvasPaintResourcePainter(
    createPaintResourceSession(registry, handles),
    canonicalCanvasPaintResourceHandlers,
  );
}

describe('canonical Canvas paint resource handlers', () => {
  it('paints cropped images in point-space with DrawingML alpha', () => {
    const resourceKey = 'image:body:0';
    const image = { width: 100, height: 80 } as CanvasImageSource;
    const paint = painter([{
      kind: 'image', resourceKey, partPath: 'word/media/image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 40, heightPt: 30 },
      srcRect: { l: .1, t: .25, r: .2, b: 0 }, alpha: .5,
    }], [{ kind: 'image', resourceKey, handle: image }]);
    const { ctx, operations } = recordingContext();

    paint.paint(resourceKey, 'image', { xPt: 10, yPt: 20, widthPt: 40, heightPt: 30 }, ctx);

    expect(operations.find((operation) => operation.name === 'drawImage')).toEqual({
      name: 'drawImage', alpha: .4,
      args: [image, 10, 20, 70, 60, 10, 20, 40, 30],
    });
  });

  it('composes image rotation and reflection around the retained bounds center', () => {
    const resourceKey = 'image:rotated';
    const image = { width: 20, height: 10 } as CanvasImageSource;
    const paint = painter([{
      kind: 'image', resourceKey, partPath: 'word/media/image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 40, heightPt: 30 }, rotation: 90, flipH: true, flipV: false,
    }], [{ kind: 'image', resourceKey, handle: image }]);
    const { ctx, operations } = recordingContext();

    paint.paint(resourceKey, 'image', { xPt: 10, yPt: 20, widthPt: 40, heightPt: 30 }, ctx);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'translate', args: [30, 35] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'scale', args: [-1, 1] },
      { name: 'drawImage', alpha: .8, args: [image, -20, -15, 40, 30] },
    ]));
  });

  it.each(['math', 'picture-bullet'] as const)(
    'draws %s handles into the retained point-space bounds',
    (kind) => {
      const resourceKey = `${kind}:body:0`;
      const image = { width: 20, height: 10 } as CanvasImageSource;
      const descriptor: PaintResourceDescriptor = kind === 'math'
        ? { kind, resourceKey }
        : {
            kind, resourceKey, partPath: 'word/media/bullet.gif', mimeType: 'image/gif',
            intrinsicSize: { widthPt: 6, heightPt: 7 },
          };
      const paint = painter([descriptor], [{ kind, resourceKey, handle: image }]);
      const { ctx, operations } = recordingContext();

      paint.paint(resourceKey, kind, { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 }, ctx);

      expect(operations.find((operation) => operation.name === 'drawImage')).toEqual({
        name: 'drawImage', alpha: .8, args: [image, 1, 2, 3, 4],
      });
    },
  );

  it('skips an explicitly unavailable drawable without treating it as a missing handle', () => {
    const resourceKey = 'math:unavailable';
    const paint = painter(
      [{ kind: 'math', resourceKey }],
      [{
        kind: 'math', resourceKey,
        handle: unavailablePaintResourceHandle('optional math renderer unavailable'),
      }],
    );
    const { ctx, operations } = recordingContext();

    paint.paint(resourceKey, 'math', { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 }, ctx);

    expect(operations).toEqual([]);
  });

  it('paints a placeholder for an image whose optional TIFF codec is unavailable', () => {
    const resourceKey = 'image:tiff-unavailable';
    const paint = painter(
      [{
        kind: 'image', resourceKey, partPath: 'word/media/image.tiff', mimeType: 'image/tiff',
        intrinsicSize: { widthPt: 40, heightPt: 30 }, rotation: 90,
      }],
      [{
        kind: 'image', resourceKey,
        handle: unavailablePaintResourceHandle(
          'optional TIFF codec unavailable',
          { placeholder: 'tiff' },
        ),
      }],
    );
    const { ctx, operations } = recordingContext();

    paint.paint(resourceKey, 'image', { xPt: 10, yPt: 20, widthPt: 40, heightPt: 30 }, ctx);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'translate', args: [30, 35] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'fillText', args: ['TIFF image unavailable', 0, 0, 40] },
    ]));
    expect(operations.some(operation => operation.name === 'drawImage')).toBe(false);
  });

  it('passes chart bounds and point scaling once to the shared core renderer', () => {
    const resourceKey = 'chart:body:0';
    const paint = painter([{
      kind: 'chart', resourceKey, intrinsicSize: { widthPt: 120, heightPt: 80 },
      model: chartModel(),
    }], [{ kind: 'chart', resourceKey, handle: null }]);
    const { ctx, operations } = recordingContext();

    paint.paint(resourceKey, 'chart', { xPt: 10, yPt: 20, widthPt: 120, heightPt: 80 }, ctx);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'fillRect', args: [10, 20, 120, 80] },
      { name: 'strokeRect', args: [10.5, 20.5, 119, 79] },
      { name: 'fillText', args: ['(no data)', 70, 60] },
    ]));
  });

  it('forwards the opt-in ChartEx module through the DOCX chart paint boundary', () => {
    const resourceKey = 'chart:chartex';
    const render = vi.fn(() => true);
    const chartEx = { render } as ChartExRenderer;
    const registry = createPaintResourceRegistry([{
      kind: 'chart', resourceKey, intrinsicSize: { widthPt: 120, heightPt: 80 },
      model: {
        ...chartModel(),
        chartType: 'boxWhisker',
        chartexBox: { categories: [], series: [] },
      },
    }]);
    const paint = createCanvasPaintResourcePainter(
      createPaintResourceSession(registry, [{ kind: 'chart', resourceKey, handle: null }]),
      createCanonicalCanvasPaintResourceHandlers(undefined, undefined, undefined, chartEx),
    );
    const { ctx } = recordingContext();

    paint.paint(resourceKey, 'chart', { xPt: 10, yPt: 20, widthPt: 120, heightPt: 80 }, ctx);

    expect(render).toHaveBeenCalledOnce();
    expect(render).toHaveBeenCalledWith(
      ctx,
      expect.objectContaining({ chartType: 'boxWhisker' }),
      { x: 10, y: 20, w: 120, h: 80 },
      1,
      0,
    );
  });
});
