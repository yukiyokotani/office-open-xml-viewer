import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildPageLayers } from '../layout/page-graph.js';
import { rasterPaintOccurrencesForPage } from '../layout/text-index.js';
import type { DocumentLayout, LayoutPage, PaintResourceRegistry } from '../layout/types.js';
import { renderSelectedDocumentPage } from './canvas-document.js';
import { TiffDecodeError, type ChartModel, type TiffRenderOptions } from '@silurus/ooxml-core';

class RecordingContext {
  readonly operations: string[] = [];
  fillStyle = '';
  globalAlpha = 1;

  scale(): void { this.operations.push('scale'); }
  fillRect(): void { this.operations.push('fillRect'); }
  save(): void { this.operations.push('save'); }
  restore(): void { this.operations.push('restore'); }
  drawImage(): void { this.operations.push('drawImage'); }
  fillText(): void { this.operations.push('fillText'); }
  translate(): void { this.operations.push('translate'); }
  rotate(): void { this.operations.push('rotate'); }
  setTransform(): void { this.operations.push('setTransform'); }
  clearRect(): void { this.operations.push('clearRect'); }
}

class ElementCanvas {
  width = 1;
  height = 1;
  isConnected = false;
  removeCalls = 0;
  ownerDocument?: {
    defaultView?: { HTMLCanvasElement: typeof ElementCanvas };
  };
  readonly attributes: Record<string, string> = {};
  readonly style: Record<string, string> = {};
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
  setAttribute(name: string, value: string): void {
    this.attributes[name] = value;
  }
  remove(): void {
    this.isConnected = false;
    this.removeCalls += 1;
  }
}

class WorkerCanvas {
  width = 1;
  height = 1;
  readonly context = new RecordingContext();

  getContext(): RecordingContext { return this.context; }
}

const section: LayoutPage['section'] = {
  geometry: {
    pageWidth: 200, pageHeight: 100,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0,
  },
  columns: [{ xPt: 0, wPt: 200 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'tbRl', verticalAlignment: 'top',
};

const page: LayoutPage = {
  pageIndex: 0,
  geometry: {
    xPt: 0, yPt: 0, widthPt: 200, heightPt: 100,
    contentTopPt: 0, contentBottomPt: 100,
  },
  flowDomains: [], section, sectionOccurrenceId: 'section:0', parityBlank: false,
  bookmarkStarts: [],
  pageNumber: { displayNumber: 1, format: 'decimal', sectionOccurrenceId: 'section:0' },
  sectionRegions: [], columnSeparators: [], pageBorder: null,
  layers: {
    ...buildPageLayers([]),
    capabilities: { requiresElementBackedVerticalGlyphPaint: true },
  },
  readingOrder: [],
};

const layout: DocumentLayout = { pages: [page], diagnostics: [] };
const registry: PaintResourceRegistry = {
  keys: [], descriptors: [],
  resolve() { throw new Error('empty registry'); },
};

function tiffDimensions(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(38);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49], 0);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, 2, true);
  view.setUint16(10, 256, true);
  view.setUint16(12, 4, true);
  view.setUint32(14, 1, true);
  view.setUint32(18, width, true);
  view.setUint16(22, 257, true);
  view.setUint16(24, 4, true);
  view.setUint32(26, 1, true);
  view.setUint32(30, height, true);
  return bytes;
}

afterEach(() => vi.unstubAllGlobals());

describe('vertical OpenType paint target projection', () => {
  it('sizes a grouped image decode from its final retained paint frame under an adaptive budget', async () => {
    const resourceKey = 'image:body:grouped';
    const finalFrame = { xPt: 20, yPt: 10, widthPt: 8, heightPt: 9 };
    const drawing = {
      kind: 'drawing' as const,
      id: 'grouped-image-drawing',
      source: { story: 'body' as const, storyInstance: 'body', path: [0, 1] },
      flowDomainId: 'body:domain',
      flowBounds: finalFrame,
      inkBounds: finalFrame,
      advancePt: 0,
      ordinaryFlow: false,
      commands: [{
        kind: 'resource' as const,
        resourceKind: 'image' as const,
        resourceKey,
        rect: finalFrame,
      }],
    };
    const imageDescriptor = {
      kind: 'image' as const,
      resourceKey,
      partPath: 'word/media/grouped.tiff',
      mimeType: 'image/tiff',
      // The authored child extent precedes grouped/relative anchor scaling.
      intrinsicSize: { widthPt: 4, heightPt: 3 },
    };
    const imageRegistry: PaintResourceRegistry = {
      keys: [resourceKey],
      descriptors: [imageDescriptor],
      resolve() { return imageDescriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: buildPageLayers([{ layer: 'body', node: drawing }]),
    };
    const directLayout: DocumentLayout = { pages: [directPage], diagnostics: [] };
    const rasterPaintOccurrences = rasterPaintOccurrencesForPage(directLayout, 0);
    const renderTiff = vi.fn(async () => ({
      width: 11, height: 12, close() {},
    }) as unknown as ImageBitmap);
    const bytes = tiffDimensions(4_249, 6_137);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));

    await renderSelectedDocumentPage(
      directLayout,
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      {
        dpr: 1,
        parseError: false,
        registry: imageRegistry,
        rasterPaintOccurrences,
        textRuns: [],
        fetchImage,
        tiff: { render: renderTiff },
        imageResources: { decodedByteBudget: 1_024 },
      },
    );

    expect(renderTiff).toHaveBeenCalledWith(expect.any(Uint8Array), expect.objectContaining({
      // 8x9 pt at the 96-dpi default paint scale, rounded outward.
      targetWidthPx: 11,
      targetHeightPx: 12,
    }));
    expect(rasterPaintOccurrences).toEqual([{
      resourceKey,
      resourceKind: 'image',
      widthPt: 8,
      heightPt: 9,
    }]);
  });

  it('contains a missing optional TIFF codec at the retained image node', async () => {
    const resourceKey = 'image:body:optional-tiff';
    const frame = { xPt: 20, yPt: 10, widthPt: 80, heightPt: 40 };
    const drawing = {
      kind: 'drawing' as const,
      id: 'optional-tiff-drawing',
      source: { story: 'body' as const, storyInstance: 'body', path: [0, 2] },
      flowDomainId: 'body:domain',
      flowBounds: frame,
      inkBounds: frame,
      advancePt: 0,
      ordinaryFlow: false,
      commands: [{
        kind: 'resource' as const,
        resourceKind: 'image' as const,
        resourceKey,
        rect: frame,
      }],
    };
    const descriptor = {
      kind: 'image' as const,
      resourceKey,
      partPath: 'word/media/optional.tiff',
      mimeType: 'image/tiff',
      intrinsicSize: { widthPt: 80, heightPt: 40 },
    };
    const imageRegistry: PaintResourceRegistry = {
      keys: [resourceKey],
      descriptors: [descriptor],
      resolve() { return descriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: buildPageLayers([{ layer: 'body', node: drawing }]),
    };
    const directLayout: DocumentLayout = { pages: [directPage], diagnostics: [] };
    const canvas = new WorkerCanvas();
    const fetchImage = vi.fn(async () => new Blob([
      tiffDimensions(80, 40) as BlobPart,
    ], { type: 'image/tiff' }));

    await expect(renderSelectedDocumentPage(
      directLayout,
      directPage,
      canvas as unknown as OffscreenCanvas,
      {
        dpr: 1,
        parseError: false,
        registry: imageRegistry,
        rasterPaintOccurrences: rasterPaintOccurrencesForPage(directLayout, 0),
        textRuns: [],
        fetchImage,
      },
    )).resolves.toBeUndefined();

    expect(canvas.context.operations).toContain('fillText');
    expect(canvas.context.operations).not.toContain('drawImage');
  });

  it('prefetches one relationship-backed chart picture marker before synchronous paint', async () => {
    const fetchImage = vi.fn(async () => new Blob(['png'], { type: 'image/png' }));
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 8, height: 8 })));
    const model = {
      chartType: 'line', categories: ['A'],
      series: [{
        name: 'Picture', values: [1], markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'word/media/chart-marker-prefetch.png',
          svgImagePath: 'word/media/chart-marker-prefetch.svg',
          mimeType: 'image/png',
          srcRect: { l: 0.1, t: 0, r: 0, b: 0 },
        },
      }],
    } as ChartModel;
    const chartDescriptor = {
      kind: 'chart' as const,
      resourceKey: 'chart:body:picture',
      intrinsicSize: { widthPt: 100, heightPt: 60 },
      model,
    };
    const hiddenChartDescriptor = {
      ...chartDescriptor,
      resourceKey: 'chart:other-page:picture',
      model: {
        ...model,
        series: [{
          ...model.series[0],
          markerFillPaint: {
            ...model.series[0]!.markerFillPaint,
            imagePath: 'word/media/other-page.png',
          },
        }],
      } as ChartModel,
    };
    const chartRegistry: PaintResourceRegistry = {
      keys: [chartDescriptor.resourceKey, hiddenChartDescriptor.resourceKey],
      descriptors: [chartDescriptor, hiddenChartDescriptor],
      resolve(key) {
        return (key === chartDescriptor.resourceKey ? chartDescriptor : hiddenChartDescriptor) as never;
      },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: {
        ...page.layers,
        capabilities: {
          requiresElementBackedVerticalGlyphPaint: false,
          resourceKeys: [chartDescriptor.resourceKey],
        },
      },
    };

    await renderSelectedDocumentPage(
      { pages: [directPage], diagnostics: [] },
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      {
        dpr: 1,
        parseError: false,
        registry: chartRegistry,
        rasterPaintOccurrences: [{
          resourceKey: chartDescriptor.resourceKey,
          resourceKind: 'chart',
          widthPt: 100,
          heightPt: 60,
        }],
        textRuns: [],
        fetchImage,
      },
    );

    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(
      'word/media/chart-marker-prefetch.png',
      'image/png',
    );
    expect(fetchImage).not.toHaveBeenCalledWith(
      'word/media/chart-marker-prefetch.svg',
      expect.anything(),
    );
    expect(fetchImage).not.toHaveBeenCalledWith('word/media/other-page.png', expect.anything());
  });

  it('keeps a valid chart occurrence when another retained use has an invalid frame', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 8, height: 8 })));
    const imagePath = 'word/media/chart-valid-occurrence.png';
    const fetchImage = vi.fn(async () => new Blob(['png'], { type: 'image/png' }));
    const chartDescriptor = {
      kind: 'chart' as const,
      resourceKey: 'chart:body:mixed-occurrences',
      intrinsicSize: { widthPt: 1, heightPt: 1 },
      model: {
        chartType: 'line', categories: ['A'],
        series: [{
          name: 'Picture', values: [1], markerFillPaint: {
            fillType: 'image', stretch: true, imagePath, mimeType: 'image/png',
          },
        }],
      } as ChartModel,
    };
    const chartRegistry: PaintResourceRegistry = {
      keys: [chartDescriptor.resourceKey],
      descriptors: [chartDescriptor],
      resolve() { return chartDescriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: {
        ...page.layers,
        capabilities: { requiresElementBackedVerticalGlyphPaint: false },
      },
    };

    await renderSelectedDocumentPage(
      { pages: [directPage], diagnostics: [] },
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      {
        dpr: 1,
        parseError: false,
        registry: chartRegistry,
        rasterPaintOccurrences: [{
          resourceKey: chartDescriptor.resourceKey,
          resourceKind: 'chart',
          widthPt: 100,
          heightPt: 60,
        }, {
          resourceKey: chartDescriptor.resourceKey,
          resourceKind: 'chart',
          widthPt: Number.NaN,
          heightPt: 60,
        }],
        textRuns: [],
        fetchImage,
      },
    );

    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(imagePath, 'image/png');
  });

  it.each([
    ['zero-width', 0, 60, undefined],
    ['zero-height', 100, 0, undefined],
    ['negative-width', -1, 60, undefined],
    ['negative-height', 100, -1, undefined],
    ['non-finite-width', Number.NaN, 60, undefined],
    ['non-finite-height', 100, Number.POSITIVE_INFINITY, undefined],
    [
      'overflowing-derived-width',
      Number.MAX_VALUE / 4,
      60,
      { l: 0.5, t: 0, r: 0.4999999999999999, b: 0 },
    ],
  ] as const)(
    'excludes %s retained chart occurrences before aggregate source gating and preload',
    async (_case, invalidWidthPt, invalidHeightPt, overflowCrop) => {
      vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 8, height: 8 })));
      const fetchImage = vi.fn(async () => new Blob(['png'], { type: 'image/png' }));
      const visiblePath = 'word/media/visible-chart-fill.png';
      const visibleFill = {
        fillType: 'image' as const,
        stretch: true,
        imagePath: visiblePath,
        mimeType: 'image/png',
      };
      const visibleChart = {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Visible', values: [1], markerFillPaint: visibleFill }],
      } as ChartModel;
      const sourceCount = overflowCrop ? 257 : 256;
      const invalidChart = {
        chartType: 'line',
        categories: Array.from({ length: sourceCount }, (_, index) => String(index)),
        series: [{
          name: 'Invisible',
          values: Array.from({ length: sourceCount }, () => 1),
          showMarker: false,
          dataPointOverrides: Array.from({ length: sourceCount }, (_, index) => ({
            idx: index,
            markerSymbol: 'picture' as const,
            markerFillPaint: {
              fillType: 'image' as const,
              ...(overflowCrop
                ? { stretch: true, srcRect: overflowCrop }
                : {
                    tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
                    dpi: 96,
                  }),
              imagePath: `word/media/invisible-${index}.png`,
              mimeType: 'image/png',
            },
          })),
        }],
      } as ChartModel;
      const visibleDescriptor = {
        kind: 'chart' as const,
        resourceKey: 'chart:body:visible',
        intrinsicSize: { widthPt: 100, heightPt: 60 },
        model: visibleChart,
      };
      const invalidDescriptor = {
        kind: 'chart' as const,
        resourceKey: `chart:body:invalid-${_case}`,
        intrinsicSize: { widthPt: 100, heightPt: 60 },
        model: invalidChart,
      };
      const chartRegistry: PaintResourceRegistry = {
        keys: [visibleDescriptor.resourceKey, invalidDescriptor.resourceKey],
        descriptors: [visibleDescriptor, invalidDescriptor],
        resolve(resourceKey) {
          return (resourceKey === visibleDescriptor.resourceKey
            ? visibleDescriptor
            : invalidDescriptor) as never;
        },
      };
      const directPage: LayoutPage = {
        ...page,
        layers: {
          ...page.layers,
          capabilities: { requiresElementBackedVerticalGlyphPaint: false },
        },
      };

      await renderSelectedDocumentPage(
        { pages: [directPage], diagnostics: [] },
        directPage,
        new WorkerCanvas() as unknown as OffscreenCanvas,
        {
          dpr: 1,
          parseError: false,
          registry: chartRegistry,
          rasterPaintOccurrences: [{
            resourceKey: visibleDescriptor.resourceKey,
            resourceKind: 'chart',
            widthPt: 100,
            heightPt: 60,
          }, {
            resourceKey: invalidDescriptor.resourceKey,
            resourceKind: 'chart',
            widthPt: invalidWidthPt,
            heightPt: invalidHeightPt,
          }],
          textRuns: [],
          fetchImage,
        },
      );

      expect(fetchImage).toHaveBeenCalledTimes(1);
      expect(fetchImage).toHaveBeenCalledWith(visiblePath, 'image/png');
    },
  );

  it('surfaces a TIFF diagnostic after sizing the chart picture from its retained occurrence', async () => {
    const failure = new TiffDecodeError('Unsupported TIFF compression');
    const renderTiff = vi.fn(async () => { throw failure; });
    const bytes = tiffDimensions(4_249, 6_137);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const model = {
      chartType: 'line', categories: ['A'],
      series: [{
        name: 'TIFF picture', values: [1], markerFillPaint: {
          fillType: 'image', stretch: true,
          imagePath: 'word/media/chart-marker.tiff',
          mimeType: 'image/tiff',
        },
      }],
    } as ChartModel;
    const chartDescriptor = {
      kind: 'chart' as const,
      resourceKey: 'chart:body:tiff',
      intrinsicSize: { widthPt: 1, heightPt: 1 },
      model,
    };
    const chartRegistry: PaintResourceRegistry = {
      keys: [chartDescriptor.resourceKey],
      descriptors: [chartDescriptor],
      resolve() { return chartDescriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: {
        ...page.layers,
        capabilities: { requiresElementBackedVerticalGlyphPaint: false },
      },
    };

    await expect(renderSelectedDocumentPage(
      { pages: [directPage], diagnostics: [] },
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      {
        dpr: 1,
        parseError: false,
        registry: chartRegistry,
        rasterPaintOccurrences: [{
          resourceKey: chartDescriptor.resourceKey,
          resourceKind: 'chart',
          widthPt: 100,
          heightPt: 60,
        }],
        textRuns: [],
        fetchImage,
        tiff: { render: renderTiff },
        imageResources: { decodedByteBudget: 200_000 },
      },
    )).rejects.toBe(failure);
    expect(renderTiff).toHaveBeenCalledWith(expect.any(Uint8Array), expect.objectContaining({
      targetWidthPx: 134,
      targetHeightPx: 80,
    }));
  });

  it.each([false, true])(
    'keeps a same-chart shared source native-sized when one fill is tiled (reversed=%s)',
    async reversed => {
    const renderTiff = vi.fn(async (
      _bytes: Uint8Array,
      _options?: Readonly<TiffRenderOptions>,
    ) => ({
      width: 1200, height: 900, close() {},
    }) as unknown as ImageBitmap);
    const fetchImage = vi.fn(async () => new Blob([
      new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]) as BlobPart,
    ], { type: 'image/tiff' }));
    const shared = {
      fillType: 'image' as const,
      imagePath: 'word/media/shared-chart-fill.tiff',
      mimeType: 'image/tiff',
    };
    const stretchFill = { ...shared, stretch: true };
    const tileFill = {
      ...shared,
      tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
      dpi: 96,
    };
    const [chartFill, plotAreaFill] = reversed
      ? [tileFill, stretchFill]
      : [stretchFill, tileFill];
    const chartDescriptor = {
      kind: 'chart' as const,
      resourceKey: 'chart:body:shared-tile',
      intrinsicSize: { widthPt: 200, heightPt: 120 },
      model: {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Series', values: [1] }],
        chartFill,
        plotAreaFill,
      } as ChartModel,
    };
    const chartRegistry: PaintResourceRegistry = {
      keys: [chartDescriptor.resourceKey],
      descriptors: [chartDescriptor],
      resolve() { return chartDescriptor as never; },
    };
    const directPage: LayoutPage = {
      ...page,
      layers: {
        ...page.layers,
        capabilities: { requiresElementBackedVerticalGlyphPaint: false },
      },
    };

    await renderSelectedDocumentPage(
      { pages: [directPage], diagnostics: [] },
      directPage,
      new WorkerCanvas() as unknown as OffscreenCanvas,
      {
        dpr: 2,
        parseError: false,
        registry: chartRegistry,
        rasterPaintOccurrences: [{
          resourceKey: chartDescriptor.resourceKey,
          resourceKind: 'chart',
          widthPt: 200,
          heightPt: 120,
        }],
        textRuns: [],
        fetchImage,
        tiff: { render: renderTiff },
      },
    );

    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(renderTiff).toHaveBeenCalledTimes(1);
    expect(renderTiff.mock.calls[0]?.[1]?.targetWidthPx).toBeUndefined();
    expect(renderTiff.mock.calls[0]?.[1]?.targetHeightPx).toBeUndefined();
    },
  );

  it('paints into an element-backed surface before copying to an OffscreenCanvas target', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement(tag: string) {
        if (tag !== 'canvas') throw new Error(`unexpected element ${tag}`);
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new WorkerCanvas();

    await renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      {
        dpr: 1, parseError: false, registry, rasterPaintOccurrences: [], textRuns: [],
      },
    );

    expect(created).toHaveLength(1);
    expect(created[0]!.context.operations).toContain('fillRect');
    expect(target.context.operations).toEqual(['drawImage']);
    expect({ width: target.width, height: target.height }).toEqual({
      width: created[0]!.width,
      height: created[0]!.height,
    });
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
    expect(created[0]!.attributes['aria-hidden']).toBe('true');
  });

  it('leaves a detached caller canvas detached and projects from a hidden attached surface', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new ElementCanvas();

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, rasterPaintOccurrences: [], textRuns: [],
    });

    expect(target.isConnected).toBe(false);
    expect(target.context.operations).toEqual(['drawImage']);
    expect(created).toHaveLength(1);
    expect(created[0]!.context.operations).toContain('fillRect');
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('paints directly into an already attached caller canvas without reparenting it', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = new ElementCanvas();
    target.isConnected = true;

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, rasterPaintOccurrences: [], textRuns: [],
    });

    expect(created).toEqual([]);
    expect(target.isConnected).toBe(true);
    expect(target.removeCalls).toBe(0);
    expect(target.context.operations).toContain('fillRect');
    expect(target.context.operations).not.toContain('drawImage');
  });

  it('removes the hidden attached surface when target projection fails', async () => {
    const created: ElementCanvas[] = [];
    vi.stubGlobal('HTMLCanvasElement', ElementCanvas);
    vi.stubGlobal('document', {
      createElement() {
        const canvas = new ElementCanvas();
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ElementCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    });
    const target = {
      width: 1,
      height: 1,
      getContext() { return null; },
    };

    await expect(renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      { dpr: 1, parseError: false, registry, rasterPaintOccurrences: [], textRuns: [] },
    )).rejects.toThrow('2D canvas is unavailable for DOCX paint projection');

    expect(created).toHaveLength(1);
    expect(created[0]!.isConnected).toBe(false);
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('creates the hidden paint surface in a detached caller canvas owner document', async () => {
    class MainCanvas {}
    class ForeignCanvas extends ElementCanvas {}
    const created: ForeignCanvas[] = [];
    const foreignDocument = {
      defaultView: { HTMLCanvasElement: ForeignCanvas },
      createElement() {
        const canvas = new ForeignCanvas();
        canvas.ownerDocument = foreignDocument;
        created.push(canvas);
        return canvas;
      },
      body: {
        appendChild(canvas: ForeignCanvas) {
          canvas.isConnected = true;
          return canvas;
        },
      },
    };
    vi.stubGlobal('HTMLCanvasElement', MainCanvas);
    vi.stubGlobal('document', {
      createElement() {
        throw new Error('global document must not own a foreign canvas staging surface');
      },
      body: {
        appendChild() {
          throw new Error('global document must not receive a foreign canvas staging surface');
        },
      },
    });
    const target = new ForeignCanvas();
    target.ownerDocument = foreignDocument;

    await renderSelectedDocumentPage(layout, page, target as unknown as HTMLCanvasElement, {
      dpr: 1, parseError: false, registry, rasterPaintOccurrences: [], textRuns: [],
    });

    expect(created).toHaveLength(1);
    expect(target.isConnected).toBe(false);
    expect(target.context.operations).toEqual(['drawImage']);
    expect(created[0]!.ownerDocument).toBe(foreignDocument);
    expect(created[0]!.attributes['aria-hidden']).toBe('true');
    expect(created[0]!.removeCalls).toBe(1);
  });

  it('does not require a document staging surface for parse-error fallback paint', async () => {
    vi.stubGlobal('HTMLCanvasElement', undefined);
    vi.stubGlobal('document', undefined);
    const target = new WorkerCanvas();

    await renderSelectedDocumentPage(
      layout,
      page,
      target as unknown as OffscreenCanvas,
      { dpr: 1, parseError: true, registry, rasterPaintOccurrences: [], textRuns: [] },
    );

    expect(target.context.operations).toContain('clearRect');
    expect(target.context.operations).not.toContain('drawImage');
  });
});
