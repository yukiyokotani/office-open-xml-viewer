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

function pngBlob(width = 12_090, height = 9_063): Blob {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return new Blob([bytes as BlobPart], { type: 'image/png' });
}

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
    const fetchImage = vi.fn(async () => pngBlob());
    const tiff = { render: vi.fn() };

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage, tiff,
    });

    expect(coreMocks.decode).toHaveBeenCalledTimes(1);
    expect(coreMocks.decode.mock.calls[0]?.slice(0, 4)).toEqual([
      'ppt/media/chart-marker.png', 'image/png', undefined, fetchImage,
    ]);
    expect(coreMocks.decode.mock.calls[0]?.[4]).toMatchObject({
      tiff,
      targetWidthPx: 420,
      targetHeightPx: 315,
    });
    expect(coreMocks.renderChart).toHaveBeenCalledTimes(1);
    expect(coreMocks.resolved()).toBe(coreMocks.bitmap);
  });

  it.each([
    ['zero-width', 0, 3_000_000, undefined],
    ['zero-height', 4_000_000, 0, undefined],
    ['negative-width', -1, 3_000_000, undefined],
    ['negative-height', 4_000_000, -1, undefined],
    ['non-finite-width', Number.NaN, 3_000_000, undefined],
    ['non-finite-height', 4_000_000, Number.POSITIVE_INFINITY, undefined],
    [
      'overflowing-derived-width',
      Number.MAX_VALUE,
      3_000_000,
      { l: 0.5, t: 0, r: 0.4999999999999999, b: 0 },
    ],
  ] as const)(
    'excludes %s chart frames before aggregate source gating and preload',
    async (_case, hiddenWidth, hiddenHeight, overflowCrop) => {
      coreMocks.decode.mockClear();
      coreMocks.decodeSvg.mockClear();
      coreMocks.renderChart.mockClear();
      const visibleFill = {
        fillType: 'image' as const,
        stretch: true,
        imagePath: 'ppt/media/visible-chart-fill.png',
        mimeType: 'image/png',
      };
      coreMocks.setLookupFill(visibleFill);
      const visibleChart = {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Visible', values: [1], markerFillPaint: visibleFill }],
      } as ChartModel;
      const sourceCount = overflowCrop ? 257 : 256;
      const hiddenChart = {
        chartType: 'line',
        categories: Array.from({ length: sourceCount }, (_, index) => String(index)),
        series: [{
          name: 'Invisible',
          values: Array.from({ length: sourceCount }, () => 1),
          showMarker: false,
          dataPointOverrides: Array.from({ length: sourceCount }, (_, idx) => ({
            idx,
            markerSymbol: 'picture' as const,
            markerFillPaint: {
              fillType: 'image' as const,
              ...(overflowCrop
                ? { stretch: true, srcRect: overflowCrop }
                : {
                    tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
                    dpi: 96,
                  }),
              imagePath: `ppt/media/invisible-${idx}.png`,
              mimeType: 'image/png',
            },
          })),
        }],
      } as ChartModel;
      const slide = {
        index: 0, slideNumber: 1, background: null,
        elements: [
          {
            type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000,
            rotation: 0, flipH: false, flipV: false, chart: visibleChart,
          },
          {
            type: 'chart', x: 0, y: 0, width: hiddenWidth, height: hiddenHeight,
            rotation: 0, flipH: false, flipV: false, chart: hiddenChart,
          },
        ],
      } as Slide;
      const fetchImage = vi.fn(async () => pngBlob());

      await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
        width: 960, dpr: 1, fetchImage,
      });

      expect(coreMocks.decode).toHaveBeenCalledTimes(1);
      expect(coreMocks.decode.mock.calls[0]?.slice(0, 4)).toEqual([
        'ppt/media/visible-chart-fill.png', 'image/png', undefined, fetchImage,
      ]);
      expect(coreMocks.decode.mock.calls[0]?.[4]).toHaveProperty('targetWidthPx');
      expect(coreMocks.decode.mock.calls[0]?.[4]).toHaveProperty('targetHeightPx');
    },
  );

  it('deduplicates a shared fill at the largest crop- and DPR-aware chart-frame target', async () => {
    coreMocks.decode.mockClear();
    coreMocks.decodeSvg.mockClear();
    coreMocks.renderChart.mockClear();
    const fill = {
      fillType: 'image' as const,
      stretch: true,
      imagePath: 'ppt/media/shared-chart-fill.png',
      mimeType: 'image/png',
      srcRect: { l: 0.25, t: 0.2, r: 0.25, b: 0.2 },
    };
    coreMocks.setLookupFill(fill);
    const chart = {
      chartType: 'line', categories: ['A'],
      series: [{ name: 'Series', values: [1] }],
      chartFill: fill,
    } as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [
        {
          type: 'chart', x: 0, y: 0, width: 2_000_000, height: 1_500_000,
          rotation: 0, flipH: false, flipV: false, chart,
        },
        {
          type: 'chart', x: 2_000_000, y: 0, width: 4_000_000, height: 3_000_000,
          rotation: 0, flipH: false, flipV: false, chart,
        },
      ],
    } as Slide;
    const fetchImage = vi.fn(async () => pngBlob(16_000, 10_000));

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 2, fetchImage,
    });

    expect(coreMocks.decode).toHaveBeenCalledTimes(1);
    expect(coreMocks.decode.mock.calls[0]?.[4]).toMatchObject({
      // Larger frame: ~420 × 315 CSS px × DPR 2, then magnified by the
      // 50%-wide / 60%-high visible source rectangle.
      targetWidthPx: 1680,
      targetHeightPx: 1050,
    });
    expect(coreMocks.renderChart).toHaveBeenCalledTimes(2);
    expect(coreMocks.resolved()).toBe(coreMocks.bitmap);
  });

  it.each([false, true])(
    'uses every same-chart crop/fillRect occurrence and forces raster SVG fallback (reversed=%s)',
    async reversed => {
      coreMocks.decode.mockClear();
      coreMocks.decodeSvg.mockClear();
      coreMocks.renderChart.mockClear();
      const shared = {
        fillType: 'image' as const,
        stretch: true,
        imagePath: 'ppt/media/shared-crop.png',
        svgImagePath: 'ppt/media/shared-crop.svg',
        mimeType: 'image/png',
      };
      const plain = { ...shared };
      const cropped = {
        ...shared,
        srcRect: { l: 0.25, t: 0.25, r: 0.25, b: 0.25 },
        fillRect: { l: -0.25, t: -0.25, r: -0.25, b: -0.25 },
      };
      const [chartFill, plotAreaFill] = reversed ? [cropped, plain] : [plain, cropped];
      coreMocks.setLookupFill(cropped);
      const chart = {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Series', values: [1] }],
        chartFill,
        plotAreaFill,
      } as ChartModel;
      const slide = {
        index: 0, slideNumber: 1, background: null,
        elements: [{
          type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000,
          rotation: 0, flipH: false, flipV: false, chart,
        }],
      } as Slide;
      const fetchImage = vi.fn(async () => pngBlob());

      await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
        width: 960, dpr: 2, fetchImage,
      });

      expect(coreMocks.decodeSvg).not.toHaveBeenCalled();
      expect(coreMocks.decode).toHaveBeenCalledTimes(1);
      const options = coreMocks.decode.mock.calls[0]?.[4] as Record<string, number>;
      expect(options).toMatchObject({ targetWidthPx: 2520, targetHeightPx: 1890 });
      expect(options.widthPt).toBeCloseTo(4_000_000 / 12_700 * 3);
      expect(options.heightPt).toBeCloseTo(3_000_000 / 12_700 * 3);
      expect(coreMocks.resolved()).toBe(coreMocks.bitmap);
    },
  );

  it.each([false, true])(
    'keeps a same-chart shared source native-sized when one use is tiled (reversed=%s)',
    async reversed => {
    coreMocks.decode.mockClear();
    coreMocks.decodeSvg.mockClear();
    coreMocks.renderChart.mockClear();
    const shared = {
      fillType: 'image' as const,
      imagePath: 'ppt/media/shared-natural-size.png',
      mimeType: 'image/png',
    };
    const stretchFill = { ...shared, stretch: true };
    const tileFill = {
      ...shared,
      tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
      dpi: 96,
    };
    coreMocks.setLookupFill(tileFill);
    const [chartFill, plotAreaFill] = reversed
      ? [tileFill, stretchFill]
      : [stretchFill, tileFill];
    const chart = {
      chartType: 'line', categories: ['A'],
      series: [{ name: 'Series', values: [1] }],
      chartFill,
      plotAreaFill,
    } as ChartModel;
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [{
        type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000,
        rotation: 0, flipH: false, flipV: false, chart,
      }],
    } as Slide;
    const fetchImage = vi.fn(async () => pngBlob());

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 2, fetchImage,
    });

    expect(coreMocks.decode).toHaveBeenCalledTimes(1);
    expect(coreMocks.decode.mock.calls[0]?.[4]).not.toHaveProperty('targetWidthPx');
    expect(coreMocks.decode.mock.calls[0]?.[4]).not.toHaveProperty('targetHeightPx');
    expect(coreMocks.renderChart).toHaveBeenCalledTimes(1);
    expect(coreMocks.resolved()).toBe(coreMocks.bitmap);
    },
  );

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
    const fetchImage = vi.fn(async () => pngBlob());

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    expect(coreMocks.decodeSvg).toHaveBeenCalledWith(
      'ppt/media/chart-marker.svg',
      fetchImage,
      {
        targetWidthPx: 420,
        targetHeightPx: 315,
        workerDecoder: undefined,
      },
    );
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
