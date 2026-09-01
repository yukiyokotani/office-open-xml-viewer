import { describe, expect, it, vi } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import type { Slide } from './types.js';

const coreMocks = vi.hoisted(() => {
  const bitmap = { width: 12, height: 10 } as unknown as ImageBitmap;
  let resolved: CanvasImageSource | null | undefined;
  let lookupFill: Record<string, unknown> = {
    fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png', mimeType: 'image/png',
  };
  return {
    bitmap,
    decode: vi.fn(async (..._args: unknown[]) => bitmap),
    decodeSvg: vi.fn(async (..._args: unknown[]) => bitmap),
    setLookupFill(fill: Record<string, unknown>) { lookupFill = fill; },
    renderChart: vi.fn((...args: unknown[]) => {
      const lookup = args[7] as ((fill: unknown) => CanvasImageSource | null | undefined);
      resolved = lookup(lookupFill);
    }),
    resolved: () => resolved,
  };
});

vi.mock('@silurus/ooxml-core', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  getCachedDuotoneBitmapByPath: coreMocks.decode,
  getCachedSvgImageByPath: coreMocks.decodeSvg,
  renderChart: coreMocks.renderChart,
}));

import { renderSlide } from './renderer.js';

function canvas(): HTMLCanvasElement {
  const state: Record<string, unknown> = { fillStyle: '', globalAlpha: 1 };
  const context = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      return () => undefined;
    },
    set(target, property: string, value) {
      target[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return {
    width: 0, height: 0, style: {}, offsetWidth: 960,
    getContext: () => context,
  } as unknown as HTMLCanvasElement;
}

describe('PPTX chart picture-marker preload', () => {
  it('warms the shared bitmap cache before the synchronous chart paint', async () => {
    const chart = {
      chartType: 'line', categories: ['A'],
      series: [{
        name: 'Picture', values: [1], markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png', mimeType: 'image/png',
        },
      }],
    } as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [{
        type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000,
        rotation: 0, flipH: false, flipV: false, chart,
      }],
    } as Slide;
    const fetchImage = vi.fn(async () => new Blob(['png'], { type: 'image/png' }));
    const tiff = { render: vi.fn() };

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage, tiff,
    });

    expect(coreMocks.decode).toHaveBeenCalledTimes(1);
    expect(coreMocks.decode.mock.calls[0]?.slice(0, 4)).toEqual([
      'ppt/media/chart-marker.png', 'image/png', undefined, fetchImage,
    ]);
    expect(coreMocks.decode.mock.calls[0]?.[4]).toMatchObject({ tiff });
    expect(coreMocks.renderChart).toHaveBeenCalledTimes(1);
    expect(coreMocks.resolved()).toBe(coreMocks.bitmap);
  });

  it('prefers an SVG twin, falls back to raster, and uses raster for duotone', async () => {
    coreMocks.decode.mockClear();
    coreMocks.decodeSvg.mockClear();
    coreMocks.renderChart.mockClear();
    coreMocks.decodeSvg.mockRejectedValueOnce(new Error('SVG decode failed'));
    coreMocks.setLookupFill({
      fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png',
      svgImagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/png',
    });
    const chart = {
      chartType: 'line', categories: ['A'],
      series: [{
        name: 'Picture', values: [1], markerSymbol: 'picture', markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png',
          svgImagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/png',
        },
      }],
    } as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [{
        type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000,
        rotation: 0, flipH: false, flipV: false, chart,
      }],
    } as Slide;
    const fetchImage = vi.fn(async () => new Blob(['image'], { type: 'image/png' }));

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    expect(coreMocks.decodeSvg).toHaveBeenCalledWith('ppt/media/chart-marker.svg', fetchImage);
    expect(coreMocks.decode).toHaveBeenCalled();
    expect(coreMocks.decode.mock.calls.at(-1)?.[4]).toMatchObject({
      failClosedOnDuotoneFailure: true,
    });
    expect(coreMocks.resolved()).toBe(coreMocks.bitmap);

    coreMocks.decode.mockClear();
    coreMocks.decodeSvg.mockClear();
    coreMocks.setLookupFill({
      fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png',
      svgImagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/png',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    });
    (chart.series[0] as { markerFillPaint: unknown }).markerFillPaint = {
      fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.png',
      svgImagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/png',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    };
    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    expect(coreMocks.decodeSvg).not.toHaveBeenCalled();
    expect(coreMocks.decode).toHaveBeenCalled();

    coreMocks.decode.mockClear();
    coreMocks.decodeSvg.mockClear();
    coreMocks.setLookupFill({
      fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/svg+xml',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    });
    (chart.series[0] as { markerFillPaint: unknown }).markerFillPaint = {
      fillType: 'image', stretch: true, imagePath: 'ppt/media/chart-marker.svg', mimeType: 'image/svg+xml',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    };
    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    expect(coreMocks.decodeSvg).not.toHaveBeenCalled();
    expect(coreMocks.decode).not.toHaveBeenCalled();
    expect(coreMocks.resolved()).toBeNull();
  });
});
