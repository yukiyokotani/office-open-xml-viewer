import { describe, it, expect, vi, afterEach } from 'vitest';
import { prefetchImages, decodeImageSource, renderWorksheetViewport } from './render-orchestrator';
import type { Worksheet, ParsedWorkbook, ImageAnchor } from './types';
import {
  dropBitmapCacheByPath,
  dropDuotoneBitmapCache,
  dropSvgImageCache,
  getCachedBitmapByPath,
  chartImageFillKey,
  OoxmlDecodedImageLimitError,
  TiffDecodeError,
  type OffscreenFactory,
  type TiffRenderOptions,
} from '@silurus/ooxml-core';
import { isOptionalImageUnavailable } from './internal/optional-image-fallback.js';

/**
 * The render orchestrator decodes embedded images lazily by zip path:
 * `decodeImageSource(imagePath, mimeType, svgImagePath?, fetchImage)` returns a
 * `CanvasImageSource` (an ImageBitmap for raster via `createImageBitmap`, an
 * HTMLImageElement for SVG via core's `getCachedSvgImageByPath`).
 * `prefetchImages` collects every image path from BOTH `ws.images` (top-level
 * `twoCellAnchor` pictures) AND the image leaves inside `ws.shapeGroups`,
 * resolves each against the SHARED, per-`fetchImage` core caches
 * (`getCachedBitmapByPath` for raster/metafile, `getCachedDuotoneBitmapByPath`
 * for a `<a:duotone>` recolour, `getCachedSvgImageByPath` for an SVG vector
 * original), and records the drawable in the passed lookup map keyed by
 * `imageCacheKey`. The map is a pure synchronous-lookup layer; ownership of the
 * decoded bitmaps lives in the shared caches (dropped per document on
 * destroy / re-parse), the same split docx/pptx use.
 */

// A minimal stand-in for a decoded raster bitmap.
class FakeBitmap {
  readonly width = 1;
  readonly height = 1;
  close() {}
  constructor(public readonly tag: string) {}
}

/** Flush pending microtasks so a drop's close-through-promise (`promise.then(b =>
 *  b.close())`) has run before the assertion — mirrors core's own cache tests. */
const flush = () => new Promise((r) => setTimeout(r, 0));

/** Build a minimal standard (non-placeable) WMF that draws one polyline, so the
 *  shared core player produces non-empty geometry (→ a non-null bitmap). */
function buildMinimalWmf(): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff);
  const i16 = (v: number) => u16(v & 0xffff);
  const u32 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  u16(1); u16(9); u16(0x0300); u32(0); u16(8); u32(0); u16(0); // 18-byte header
  const rec = (fn: number, params: number[]) => { u32(3 + params.length); u16(fn); for (const p of params) i16(p); };
  rec(0x020b, [0, 0]);             // SETWINDOWORG
  rec(0x020c, [100, 100]);         // SETWINDOWEXT
  rec(0x02fa, [0, 1, 0, 0, 0]);    // CREATEPENINDIRECT (color as low/high words)
  rec(0x012d, [0]);                // SELECTOBJECT
  rec(0x0325, [2, 0, 0, 50, 50]);  // POLYLINE
  u32(3); u16(0x0000);             // EOF
  return new Uint8Array(b);
}

/** Build a true EMF (ENHMETAHEADER) header so isEmf detects it. */
function buildEmfHeader(): Uint8Array {
  const buf = new Uint8Array(48);
  const dv = new DataView(buf.buffer);
  dv.setUint32(0, 1, true); // EMR_HEADER iType
  dv.setUint32(40, 0x464d4520, true); // " EMF"
  return buf;
}

function tiffHeader(width: number, height: number): Uint8Array {
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

/** Stub OffscreenCanvas (the WMF player's target) for the node test env. */
function stubOffscreenCanvas(): void {
  vi.stubGlobal(
    'OffscreenCanvas',
    class {
      width: number;
      height: number;
      constructor(w: number, h: number) { this.width = w; this.height = h; }
      getContext() {
        return {
          fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
          lineJoin: 'miter', lineCap: 'butt',
          save() {}, restore() {}, beginPath() {}, closePath() {},
          moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
        };
      }
    },
  );
}

/** An offscreen surface whose getImageData returns a fixed near-white pixel grid
 *  and whose putImageData records the mutated buffer, so a duotone recolour can
 *  run (and be observed) without a real canvas. Cast to the core
 *  `OffscreenFactory` at the boundary (a partial mock). Shared by the duotone
 *  ownership and recolour tests. */
function recordingFactory(record: { out?: Uint8ClampedArray }): OffscreenFactory {
  return ((w: number, h: number) => ({
    width: w,
    height: h,
    getContext() {
      return {
        drawImage() {},
        getImageData(_sx: number, _sy: number, sw: number, sh: number) {
          // All near-white opaque pixels (t≈0.96) → should map toward clr2.
          const data = new Uint8ClampedArray(sw * sh * 4).fill(246);
          for (let i = 3; i < data.length; i += 4) data[i] = 255; // alpha
          return { data, width: sw, height: sh } as unknown as ImageData;
        },
        putImageData(img: ImageData) {
          record.out = img.data;
        },
      };
    },
  })) as unknown as OffscreenFactory;
}

/** Build a Worksheet with one top-level image and one group-leaf image, each at
 *  a distinct zip path, plus enough required fields to satisfy the type. */
function worksheetWithImages(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.png',
        mimeType: 'image/png',
      },
    ],
    shapeGroups: [
      {
        fromCol: 3, fromColOff: 0, fromRow: 3, fromRowOff: 0,
        toCol: 5, toColOff: 0, toRow: 5, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        shapes: [
          {
            x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
            geom: {
              type: 'image',
              imagePath: 'xl/media/image2.png',
              mimeType: 'image/png',
            },
          },
        ],
      },
    ],
  } as Worksheet;
}

describe('render-orchestrator image decode (lazy bytes)', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('prefetchImages collects BOTH ws.images and group-leaf images, keyed by imagePath, decoded once each', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>();

    await prefetchImages(ws, cache, fetchImage);

    // Both paths decoded and cached under their zip path (not a data URL).
    expect(cache.has('xl/media/image1.png')).toBe(true);
    expect(cache.has('xl/media/image2.png')).toBe(true);
    expect(cache.size).toBe(2);
    // Each path fetched exactly once.
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image2.png', 'image/png');
  });

  it('threads anchor geometry and effective DPR into oversized raster decoding', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 12_090);
    new DataView(png.buffer).setUint32(20, 9_063);
    const decode = vi.fn(async (_source: unknown, _options?: ImageBitmapOptions) => new FakeBitmap('poster'));
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));
    const ws = worksheetWithImages();
    ws.shapeGroups = [];

    await prefetchImages(ws, new Map(), fetchImage, { effectiveDpr: 2 });

    expect(decode).toHaveBeenCalledOnce();
    const options = decode.mock.calls[0]?.[1] as ImageBitmapOptions | undefined;
    expect(options).toMatchObject({ resizeQuality: 'high' });
    expect(options?.resizeWidth).toBeGreaterThan(0);
    expect(options?.resizeWidth).toBeLessThan(12_090);
  });

  it('merges duplicate same-path placements to the larger required raster target in either order', async () => {
    const tiffBytes = tiffHeader(12_090, 9_063);
    const placement = (toCol: number, toRow: number): ImageAnchor => ({
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol, toColOff: 0, toRow, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      imagePath: 'xl/media/shared-target.tiff',
      mimeType: 'image/tiff',
    } as ImageAnchor);
    const small = placement(1, 1);
    const large = placement(4, 6);

    const run = async (images: ImageAnchor[]) => {
      const render = vi.fn(async (
        _bytes: Uint8Array,
        options?: Readonly<TiffRenderOptions>,
      ) => new FakeBitmap('shared-target') as unknown as ImageBitmap);
      const fetchImage = vi.fn(async () =>
        new Blob([tiffBytes as BlobPart], { type: 'image/tiff' }));
      const cache = new Map<string, CanvasImageSource | null>();
      const ws = worksheetWithImages();
      ws.images = images;
      ws.shapeGroups = [];

      await prefetchImages(ws, cache, fetchImage, {
        effectiveDpr: 2,
        tiff: { render },
      });

      expect(fetchImage).toHaveBeenCalledOnce();
      expect(render).toHaveBeenCalledOnce();
      expect(cache.size).toBe(1);
      expect(cache.has('xl/media/shared-target.tiff')).toBe(true);
      const target = render.mock.calls[0]?.[1];
      dropBitmapCacheByPath(fetchImage);
      return target;
    };

    const smallOnly = await run([small]);
    const largeOnly = await run([large]);
    expect(largeOnly?.targetWidthPx).toBeGreaterThan(smallOnly?.targetWidthPx ?? 0);
    expect(largeOnly?.targetHeightPx).toBeGreaterThan(smallOnly?.targetHeightPx ?? 0);
    await expect(run([large, small])).resolves.toEqual(largeOnly);
    await expect(run([small, large])).resolves.toEqual(largeOnly);
  });

  it('preserves native decoding for an ordinary visible raster by default', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 800);
    new DataView(png.buffer).setUint32(20, 600);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const decode = vi.fn(async () => new FakeBitmap('display'));
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async () => blob);
    const ws = worksheetWithImages();
    ws.shapeGroups = [];

    await prefetchImages(ws, new Map(), fetchImage, { effectiveDpr: 1 });

    expect(decode).toHaveBeenCalledWith(blob);
  });

  it('decodes an ordinary visible raster at its bounded display grid when requested', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 800);
    new DataView(png.buffer).setUint32(20, 600);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const decode = vi.fn(async () => new FakeBitmap('display'));
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async () => blob);
    const ws = worksheetWithImages();
    ws.shapeGroups = [];

    await prefetchImages(ws, new Map(), fetchImage, {
      effectiveDpr: 1,
      imageResources: { resolution: 'display' },
    });

    expect(decode).toHaveBeenCalledWith(blob, {
      resizeWidth: 800,
      resizeHeight: 54,
      resizeQuality: 'high',
    });
  });

  it('keeps DrawingML pixel effects on the authored source grid during prefetch', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 800);
    new DataView(png.buffer).setUint32(20, 600);
    const decode = vi.fn(async (source: Blob | { width?: number; height?: number }) => ({
      width: source instanceof Blob ? 800 : source.width ?? 800,
      height: source instanceof Blob ? 600 : source.height ?? 600,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async () =>
      new Blob([png as BlobPart], { type: 'image/png' }));
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    (ws.images as ImageAnchor[])[0] = {
      ...(ws.images as ImageAnchor[])[0],
      imagePath: 'xl/media/duotone-native-grid.png',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    };

    await prefetchImages(ws, new Map(), fetchImage, {
      effectiveDpr: 2,
      offscreenFactory: recordingFactory({}),
    });

    expect(decode).toHaveBeenCalledTimes(2);
    expect(decode.mock.calls[0]).toHaveLength(1);
    expect(decode.mock.calls[1]).toHaveLength(1);
  });

  it('retains the authored display target for worker-decoded SVG images', async () => {
    const svgDecoder = vi.fn(async (_blob: Blob, target?: {
      targetWidthPx?: number;
      targetHeightPx?: number;
    }) => ({
      width: target?.targetWidthPx ?? 1,
      height: target?.targetHeightPx ?? 1,
      close() {},
    }) as unknown as ImageBitmap);
    const fetchImage = vi.fn(async () => new Blob([
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"/>',
    ], { type: 'image/svg+xml' }));
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    (ws.images as ImageAnchor[])[0] = {
      ...(ws.images as ImageAnchor[])[0],
      imagePath: 'xl/media/worker-icon.svg',
      mimeType: 'image/svg+xml',
    };

    await prefetchImages(ws, new Map(), fetchImage, {
      effectiveDpr: 2,
      svgDecoder,
    });

    expect(svgDecoder).toHaveBeenCalledWith(expect.any(Blob), {
      targetWidthPx: expect.any(Number),
      targetHeightPx: expect.any(Number),
    });
    const target = svgDecoder.mock.calls[0]?.[1];
    expect(target?.targetWidthPx).toBeGreaterThan(0);
    expect(target?.targetHeightPx).toBeGreaterThan(0);
  });

  it('does not apply the anchor-sized pixel cap to a TIFF codec result', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const bitmap = { width: 1000, height: 1000, close() {} } as unknown as ImageBitmap;
    const render = vi.fn(async () => bitmap);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    (ws.images as ImageAnchor[])[0] = {
      ...(ws.images as ImageAnchor[])[0],
      imagePath: 'xl/media/photo.tiff',
      mimeType: 'image/tiff',
    };

    await expect(prefetchImages(ws, new Map(), fetchImage, {
      effectiveDpr: 1,
      tiff: { render },
    })).resolves.toBeUndefined();
    expect(render).toHaveBeenCalledOnce();
  });

  it('applies one adaptive quality scale to all visible worksheet images', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    const view = new DataView(png.buffer);
    view.setUint32(16, 3200);
    view.setUint32(20, 1000);
    const decode = vi.fn(async (_source: unknown, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 3200,
      height: options?.resizeHeight ?? 1000,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async () => new Blob([png as BlobPart], { type: 'image/png' }));

    await prefetchImages(worksheetWithImages(), new Map(), fetchImage, {
      effectiveDpr: 10,
      imageResources: { decodedByteBudget: 1_024_000, strategy: 'adaptive' },
    });

    expect(decode).toHaveBeenCalledTimes(2);
    const targets = decode.mock.calls.map((call) => ({
      width: call[1]?.resizeWidth as number,
      height: call[1]?.resizeHeight as number,
    }));
    expect(targets[0]).toEqual(targets[1]);
    expect(targets[0].width * targets[0].height * 4 * targets.length)
      .toBeLessThanOrEqual(1_024_000);
    expect(decode.mock.calls.every((call) => call[1]?.resizeQuality === 'high')).toBe(true);
  });

  it('prefetchImages sizes oversized chart picture fills from the chart anchor', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 12_090);
    new DataView(png.buffer).setUint32(20, 9_063);
    const decode = vi.fn(async (_source: unknown, _options?: ImageBitmapOptions) =>
      new FakeBitmap('chart-picture'));
    vi.stubGlobal('createImageBitmap', decode);
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([png as BlobPart], { type: mime }));
    const ws = worksheetWithImages();
    ws.images = [];
    ws.shapeGroups = [];
    ws.charts = [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
      chart: {
        chartType: 'line', categories: ['A'], showDataLabels: false,
        title: null, valMin: null, valMax: null, catAxisTitle: null, valAxisTitle: null,
        catAxisHidden: false, valAxisHidden: false, catAxisLineHidden: false,
        valAxisLineHidden: false, plotAreaBg: null, chartBg: null, showLegend: false,
        legendPos: null, catAxisCrossBetween: 'between', valAxisMajorTickMark: 'out',
        catAxisMajorTickMark: 'out', titleFontSizeHpt: null, titleFontColor: null,
        titleFontFace: null, catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
        dataLabelFontSizeHpt: null, subtotalIndices: [],
        series: [{
          name: 'Series', color: null, values: [1], showMarker: true,
          markerSymbol: 'picture', markerFillPaint: {
            fillType: 'image', stretch: true, imagePath: 'xl/media/chart-marker.png', mimeType: 'image/png',
          },
        }],
      },
    } as Worksheet['charts'][number]];
    const cache = new Map<string, CanvasImageSource | null>();
    await prefetchImages(ws, cache, fetchImage, { effectiveDpr: 2 });
    expect(fetchImage).toHaveBeenCalledOnce();
    expect(fetchImage).toHaveBeenCalledWith('xl/media/chart-marker.png', 'image/png');
    expect(decode).toHaveBeenCalledOnce();
    expect(decode.mock.calls[0]?.[1]).toMatchObject({
      // Four authored 64-character columns resolve to 2048 CSS px in this
      // fixture; DPR 2 needs 4096 px, and the available geometry share retains
      // up to 2× that display grid.
      resizeWidth: 8_192,
      resizeQuality: 'high',
    });
    expect(cache.has(chartImageFillKey({
      fillType: 'image', stretch: true, imagePath: 'xl/media/chart-marker.png', mimeType: 'image/png',
    }))).toBe(true);
  });

  it.each([
    ['zero-width', 0, 0, 8, 0],
    ['zero-height', 4, 0, 0, 0],
    ['negative-width', 0, -1, 8, 0],
    ['negative-height', 4, 0, 0, -1],
    ['non-finite-width', 0, Number.POSITIVE_INFINITY, 8, 0],
    ['non-finite-height', 4, 0, 0, Number.NaN],
  ] as const)(
    'excludes %s chart anchors before aggregate source gating without a viewport',
    async (_case, toCol, toColOff, toRow, toRowOff) => {
      vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
      const visibleFill = {
        fillType: 'image' as const,
        stretch: true,
        imagePath: 'xl/media/visible-chart-fill.png',
        mimeType: 'image/png',
      };
      const visibleChart = {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Visible', values: [1], markerFillPaint: visibleFill }],
      };
      const sourceCount = 256;
      const hiddenChart = {
        chartType: 'line',
        categories: Array.from({ length: sourceCount }, (_, index) => String(index)),
        series: [{
          name: 'Invalid anchor',
          values: Array.from({ length: sourceCount }, () => 1),
          showMarker: false,
          dataPointOverrides: Array.from({ length: sourceCount }, (_, idx) => ({
            idx,
            markerSymbol: 'picture' as const,
            markerFillPaint: {
              fillType: 'image' as const,
              tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
              dpi: 96,
              imagePath: `xl/media/invalid-chart-${idx}.png`,
              mimeType: 'image/png',
            },
          })),
        }],
      };
      const ws = worksheetWithImages();
      ws.images = [];
      ws.shapeGroups = [];
      ws.charts = [
        {
          fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
          toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
          chart: visibleChart,
        },
        {
          fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
          toCol, toColOff, toRow, toRowOff,
          chart: hiddenChart,
        },
      ] as Worksheet['charts'];
      const fetchImage = vi.fn(async (path: string, mime: string) =>
        new Blob([new TextEncoder().encode(path)], { type: mime }));
      const cache = new Map<string, CanvasImageSource | null>();

      await prefetchImages(ws, cache, fetchImage, { effectiveDpr: 1 });

      expect(fetchImage).toHaveBeenCalledOnce();
      expect(fetchImage).toHaveBeenCalledWith('xl/media/visible-chart-fill.png', 'image/png');
      expect(cache.size).toBe(1);
      expect(cache.has(chartImageFillKey(visibleFill))).toBe(true);
    },
  );

  it('excludes a chart whose finite frame overflows one usage before no-viewport aggregate gating', async () => {
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    png.set([0x49, 0x48, 0x44, 0x52], 12);
    new DataView(png.buffer).setUint32(16, 12_090);
    new DataView(png.buffer).setUint32(20, 9_063);
    const decode = vi.fn(async (_source: unknown, _options?: ImageBitmapOptions) =>
      new FakeBitmap('visible-chart'));
    vi.stubGlobal('createImageBitmap', decode);
    const visibleFill = {
      fillType: 'image' as const,
      stretch: true,
      imagePath: 'xl/media/visible-overflow-gate.png',
      mimeType: 'image/png',
    };
    const sourceCount = 257;
    const overflowChart = {
      chartType: 'line',
      categories: Array.from({ length: sourceCount }, (_, index) => String(index)),
      series: [{
        name: 'Overflow',
        values: Array.from({ length: sourceCount }, () => 1),
        showMarker: false,
        dataPointOverrides: Array.from({ length: sourceCount }, (_, idx) => ({
          idx,
          markerSymbol: 'picture' as const,
          markerFillPaint: {
            fillType: 'image' as const,
            stretch: true,
            ...(idx === sourceCount - 1
              ? { fillRect: { l: -Number.MAX_VALUE / 2, t: 0, r: 0, b: 0 } }
              : {}),
            imagePath: `xl/media/overflow-chart-${idx}.png`,
            mimeType: 'image/png',
          },
        })),
      }],
    };
    const ws = worksheetWithImages();
    ws.images = [];
    ws.shapeGroups = [];
    ws.charts = [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
        chart: {
          chartType: 'line', categories: ['A'],
          series: [{ name: 'Visible', values: [1], markerFillPaint: visibleFill }],
        },
      },
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
        chart: overflowChart,
      },
    ] as Worksheet['charts'];
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([png as BlobPart], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage, { effectiveDpr: 1 });

    expect(fetchImage).toHaveBeenCalledOnce();
    expect(fetchImage).toHaveBeenCalledWith('xl/media/visible-overflow-gate.png', 'image/png');
    expect(decode).toHaveBeenCalledOnce();
    expect(decode.mock.calls[0]?.[1]).toMatchObject({
      resizeWidth: 4_096,
      resizeQuality: 'high',
    });
    expect(cache.size).toBe(1);
    expect(cache.has(chartImageFillKey(visibleFill))).toBe(true);
  });

  it.each([false, true])(
    'keeps a same-chart shared source native-sized when one fill is tiled (reversed=%s)',
    async reversed => {
    const tiffBytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const render = vi.fn(async (
      _bytes: Uint8Array,
      _options?: Readonly<TiffRenderOptions>,
    ) => new FakeBitmap('chart-tile') as unknown as ImageBitmap);
    const fetchImage = vi.fn(async () =>
      new Blob([tiffBytes as BlobPart], { type: 'image/tiff' }));
    const shared = {
      fillType: 'image' as const,
      imagePath: 'xl/media/shared-chart-fill.tiff',
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
    const chart = {
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0,
      chart: {
        chartType: 'line', categories: ['A'],
        series: [{ name: 'Series', values: [1] }],
        chartFill,
        plotAreaFill,
      },
    } as Worksheet['charts'][number];
    const ws = worksheetWithImages();
    ws.images = [];
    ws.shapeGroups = [];
    ws.charts = [chart];
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage, {
      effectiveDpr: 2,
      tiff: { render },
    });

    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(render).toHaveBeenCalledTimes(1);
    expect(render.mock.calls[0]?.[1]?.targetWidthPx).toBeUndefined();
    expect(render.mock.calls[0]?.[1]?.targetHeightPx).toBeUndefined();
    expect(cache.has(chartImageFillKey(stretchFill))).toBe(true);
    dropBitmapCacheByPath(fetchImage);
    },
  );

  it('prefetchImages skips anchors wholly outside the current viewport', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(worksheetWithImages(), cache, fetchImage, {
      viewport: { row: 1, col: 1, rows: 2, cols: 2 },
    });

    expect(cache.has('xl/media/image1.png')).toBe(true);
    expect(cache.has('xl/media/image2.png')).toBe(false);
    expect(fetchImage).toHaveBeenCalledTimes(1);
  });

  it('prefetchImages retains an off-cell anchor whose signed offset reaches the viewport', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    const group = ws.shapeGroups?.[0];
    if (!group) throw new Error('fixture group missing');
    // CT_Marker offsets are signed ST_Coordinate values. Although the from-cell
    // is outside cols 1..2, this authored negative offset moves the frame back
    // into the visible sheet range and must not be culled before decode.
    group.fromColOff = -20_000_000;
    group.fromRowOff = -20_000_000;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 1, col: 1, rows: 2, cols: 2 },
    });

    expect(cache.has('xl/media/image2.png')).toBe(true);
  });

  it('does not fetch small one-cell images above a distant viewport', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    const image = ws.images?.[0];
    if (!image) throw new Error('fixture image missing');
    image.editAs = 'oneCell';
    image.nativeExtCx = 914_400;
    image.nativeExtCy = 914_400;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 10_000, col: 1, rows: 20, cols: 20 },
    });

    expect(fetchImage).not.toHaveBeenCalled();
    expect(cache.size).toBe(0);
  });

  it('falls back to both to-markers when a one-cell native extent is incomplete', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    const image = ws.images?.[0];
    if (!image) throw new Error('fixture image missing');
    image.editAs = 'oneCell';
    image.fromCol = 4;
    image.toCol = 8;
    image.nativeExtCx = 1;
    image.nativeExtCy = 0;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 1, col: 8, rows: 2, cols: 2 },
      width: 640,
      height: 480,
    });

    expect(fetchImage).toHaveBeenCalledOnce();
  });

  it('does not decode reversed two-cell anchors that paint no rectangle', async () => {
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    const image = ws.images?.[0];
    if (!image) throw new Error('fixture image missing');
    image.fromCol = 5;
    image.toCol = 2;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 1, col: 1, rows: 20, cols: 20 },
      width: 800,
      height: 600,
    });

    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('limits an authored whole-sheet freeze to frozen bands visible on canvas', async () => {
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    ws.freezeRows = 1_048_576;
    ws.freezeCols = 16_384;
    const image = ws.images?.[0];
    if (!image) throw new Error('fixture image missing');
    image.fromRow = 10_000;
    image.toRow = 10_001;
    image.fromCol = 100;
    image.toCol = 101;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 1, col: 1, rows: 20, cols: 20 },
      width: 640,
      height: 480,
      freezeRows: ws.freezeRows,
      freezeCols: ws.freezeCols,
    });

    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('uses scaled sparse axes for distant-anchor visibility', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }));
    const cache = new Map<string, CanvasImageSource | null>();
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    ws.colWidths = { 2_000: 17.25 };
    ws.rowHeights = { 2_000: 19.5 };
    const image = ws.images?.[0];
    if (!image) throw new Error('fixture image missing');
    image.fromCol = 1_999;
    image.toCol = 2_000;
    image.fromRow = 1_999;
    image.toRow = 2_000;

    await prefetchImages(ws, cache, fetchImage, {
      viewport: { row: 2_000, col: 2_000, rows: 1, cols: 1 },
      width: 320,
      height: 240,
      cellScale: 1.25,
    });

    expect(fetchImage).toHaveBeenCalledOnce();
  });

  it('prefetchImages does not re-fetch an already-decoded path (shared cache hit, not a stale map entry)', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>();

    // First pass warms the shared path-keyed cache (both images fetched once).
    await prefetchImages(ws, cache, fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(2);

    // Second pass re-resolves every referenced image (the way docx/pptx do, so an
    // LRU eviction of a still-referenced path is re-decoded rather than served
    // stale/closed) — but each hits the shared cache, so NO byte is re-fetched.
    await prefetchImages(ws, cache, fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(2);

    dropBitmapCacheByPath(fetchImage);
  });

  it('prefetchImages is a no-op when fetchImage is absent (cache stays empty)', async () => {
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>([
      ['stale-from-previous-sheet', {} as CanvasImageSource],
    ]);
    await prefetchImages(ws, cache, undefined);
    expect(cache.size).toBe(0);
  });

  it('decodeImageSource decodes raster via createImageBitmap from fetched bytes', async () => {
    const bmp = new FakeBitmap('image/png');
    const createImageBitmap = vi.fn(async () => bmp);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob(['X'], { type: mime }));

    const src = await decodeImageSource('xl/media/image1.png', 'image/png', undefined, fetchImage);

    expect(src).toBe(bmp);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
  });

  it('threads the opt-in TIFF codec through the XLSX image path', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const bitmap = new FakeBitmap('image/tiff') as unknown as ImageBitmap;
    const render = vi.fn(async () => bitmap);
    const browserDecode = vi.fn();
    vi.stubGlobal('createImageBitmap', browserDecode);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));

    await expect(decodeImageSource(
      'xl/media/image1.tiff',
      'image/tiff',
      undefined,
      fetchImage,
      0,
      0,
      null,
      null,
      undefined,
      false,
      { render },
    )).resolves.toBe(bitmap);

    expect(render).toHaveBeenCalledTimes(1);
    expect(browserDecode).not.toHaveBeenCalled();
  });

  it('contains a TIFF codec failure at the affected picture', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const failure = new TiffDecodeError('synthetic TIFF decode failure');
    const render = vi.fn(async () => {
      throw failure;
    });
    const fetchImage = vi.fn(async () =>
      new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const ws = worksheetWithImages();
    ws.images = [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      imagePath: 'xl/media/broken.tiff',
      mimeType: 'image/tiff',
    } as ImageAnchor];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    await expect(prefetchImages(ws, cache, fetchImage, {
      tiff: { render },
    })).resolves.toBeUndefined();
    expect(cache.get('xl/media/broken.tiff')).toBeNull();
    expect(isOptionalImageUnavailable(cache, 'xl/media/broken.tiff', 'tiff')).toBe(true);
  });

  it('keeps decoded-image limit failures fail-closed for TIFF pictures', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const failure = new OoxmlDecodedImageLimitError('image-pixels', 10, 11);
    const render = vi.fn(async () => {
      throw failure;
    });
    const fetchImage = vi.fn(async () =>
      new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const ws = worksheetWithImages();
    ws.images = [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      imagePath: 'xl/media/over-budget.tiff',
      mimeType: 'image/tiff',
    } as ImageAnchor];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    await expect(prefetchImages(ws, cache, fetchImage, {
      tiff: { render },
    })).rejects.toBe(failure);
    expect(cache.size).toBe(0);
  });

  it('marks a TIFF image unavailable instead of failing when the optional codec is absent', async () => {
    const fetchImage = vi.fn(async () =>
      new Blob([tiffHeader(32, 24) as BlobPart], { type: 'image/tiff' }));
    const ws = worksheetWithImages();
    ws.images = [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
      nativeExtCx: 0, nativeExtCy: 0,
      imagePath: 'xl/media/optional.tiff',
      mimeType: 'image/tiff',
    } as ImageAnchor];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    await expect(prefetchImages(ws, cache, fetchImage)).resolves.toBeUndefined();

    expect(cache.get('xl/media/optional.tiff')).toBeNull();
    expect(isOptionalImageUnavailable(cache, 'xl/media/optional.tiff', 'tiff')).toBe(true);
  });

  it('decodeImageSource forces the raster (not the SVG vector) when the picture is cropped', async () => {
    // A cropped picture (a non-null `srcRect`) with an svgBlip vector original
    // must decode the RASTER fallback: the renderer's `<a:srcRect>` crop math
    // needs the bitmap's native pixel grid, which an SVG element lacks. So even
    // with svgImagePath present, createImageBitmap (raster) is the path taken.
    const bmp = new FakeBitmap('image/png');
    const createImageBitmap = vi.fn(async () => bmp);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob(['X'], { type: mime }));

    const src = await decodeImageSource(
      'xl/media/image1.png',
      'image/png',
      'xl/media/image1.svg', // svgBlip present …
      fetchImage,
      0,
      0,
      { l: 0.1, t: 0, r: 0.1, b: 0 }, // … but the picture is cropped → raster wins
    );

    expect(src).toBe(bmp);
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    // The SVG part is never fetched when a crop forces the raster path.
    expect(fetchImage).not.toHaveBeenCalledWith('xl/media/image1.svg', expect.anything());
  });

  it('falls back from an SVG original to its raster twin when SVG decode fails', async () => {
    const bitmap = new FakeBitmap('image/png');
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => {
      if (blob.type === 'image/svg+xml') throw new Error('SVG decode failed');
      return bitmap;
    }));
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob(['X'], { type: mime }));

    const source = await decodeImageSource(
      'xl/media/image1.png', 'image/png', 'xl/media/image1.svg', fetchImage,
    );

    expect(source).toBe(bitmap);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.svg', 'image/svg+xml');
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
  });

  it('uses the raster twin for duotone and fails closed for SVG-only duotone', async () => {
    const bitmap = new FakeBitmap('image/png');
    vi.stubGlobal('createImageBitmap', vi.fn(async () => bitmap));
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob(['X'], { type: mime }));
    const duo = { clr1: '000000', clr2: 'FFFFFF' };

    await decodeImageSource(
      'xl/media/image1.png', 'image/png', 'xl/media/image1.svg', fetchImage,
      0, 0, null, duo, recordingFactory({}),
    );
    expect(fetchImage).not.toHaveBeenCalledWith('xl/media/image1.svg', expect.anything());
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');

    fetchImage.mockClear();
    await expect(decodeImageSource(
      'xl/media/image-only.svg', 'image/svg+xml', undefined, fetchImage,
      0, 0, null, duo,
    )).resolves.toBeNull();
    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('fails closed for a chart duotone when pixel readback is unavailable', async () => {
    const bitmap = new FakeBitmap('image/png');
    vi.stubGlobal('createImageBitmap', vi.fn(async () => bitmap));
    const fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob(['X'], { type: mime }));
    const duo = { clr1: '000000', clr2: 'FFFFFF' };

    await expect(decodeImageSource(
      'xl/media/image1.png', 'image/png', undefined, fetchImage,
      0, 0, null, duo, () => null, true,
    )).resolves.toBeNull();
    // The established non-chart compatibility path still returns the source.
    await expect(decodeImageSource(
      'xl/media/image1.png', 'image/png', undefined, fetchImage,
      0, 0, null, duo, () => null, false,
    )).resolves.toBe(bitmap);
  });

  it('decodeImageSource rasterizes a WMF blip (no throw) instead of vanishing', async () => {
    // A WMF blob used to throw in createImageBitmap; now decodeImageSource routes
    // through core's decodeRasterOrMetafile, which sniffs + rasterizes it.
    stubOffscreenCanvas();
    vi.stubGlobal('createImageBitmap', vi.fn(async (s: { width: number; height: number }) =>
      ({ width: s.width, height: s.height, close() {} }) as unknown as ImageBitmap));
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([buildMinimalWmf() as BlobPart], { type: 'image/wmf' }));

    const src = await decodeImageSource('xl/media/chart1.wmf', 'image/wmf', undefined, fetchImage, 100, 100);

    expect(src).not.toBeNull();
    expect((src as ImageBitmap).width).toBe(200); // wmfRasterTarget(100,100) → 200×200
  });

  it('decodeImageSource returns null for an unsupported metafile (true EMF), not a throw', async () => {
    const cib = vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([buildEmfHeader() as BlobPart], { type: 'image/emf' }));

    const src = await decodeImageSource('xl/media/diagram.emf', 'image/emf', undefined, fetchImage, 100, 100);

    expect(src).toBeNull();
    expect(cib).not.toHaveBeenCalled(); // EMF branch never touches createImageBitmap
  });

  it('prefetchImages caches an EMF decode as null (sniffed once) — renderer skips a null source', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap));
    const ws = worksheetWithImages();
    // Point the top-level image at an EMF; the group leaf stays a PNG.
    ws.images[0].imagePath = 'xl/media/image1.emf';
    ws.images[0].mimeType = 'image/emf';
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      path.endsWith('.emf')
        ? new Blob([buildEmfHeader() as BlobPart], { type: mime })
        : new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage);

    // EMF decodes to null but is CACHED as null (matching pptx's getCachedBitmap):
    // has() short-circuits the per-render prefetch so it isn't re-fetched every
    // frame, and the renderer skips a null (falsy) source.
    expect(cache.has('xl/media/image1.emf')).toBe(true);
    expect(cache.get('xl/media/image1.emf')).toBeNull();
    expect(cache.has('xl/media/image2.png')).toBe(true);
    expect(cache.size).toBe(2);

    // A second prefetch must NOT re-fetch the now-cached (null) EMF — the whole
    // point of caching the null: the unsupported blip is sniffed exactly once.
    const callsAfterFirst = fetchImage.mock.calls.length;
    await prefetchImages(ws, cache, fetchImage);
    expect(fetchImage.mock.calls.length).toBe(callsAfterFirst);
  });
});

// ── Teardown ownership: the shared caches own the bitmaps, not the map ────────
// After #781, xlsx decodes through the SAME per-`fetchImage` core caches
// docx/pptx use, so the passed lookup map only ever holds references; the
// GPU-backed ImageBitmaps are released by dropping the shared caches keyed by
// `fetchImage`. These tests pin that ownership (dropping the shared cache closes
// the decoded bitmap — i.e. #779's teardown leak is not reintroduced) and that
// the lookup map is a pure, non-owning layer (clearing it never closes).
describe('shared-cache consolidation (teardown ownership / no leak)', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('prefetchImages decodes the base raster into the SHARED path-keyed cache; dropBitmapCacheByPath closes it', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages(); // two distinct PNG paths
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage);

    // The lookup map exposes the drawable for the synchronous grid draw …
    expect(cache.get('xl/media/image1.png')).toBeTruthy();
    expect(cache.get('xl/media/image2.png')).toBeTruthy();
    // … but ownership is the shared cache: no bitmap was closed yet.
    expect(close).not.toHaveBeenCalled();

    // Dropping the shared cache (destroy / re-parse) closes every GPU bitmap —
    // proving the decode landed in getCachedBitmapByPath, not an orphaned map.
    dropBitmapCacheByPath(fetchImage);
    await flush();
    expect(close).toHaveBeenCalledTimes(2);
  });

  it('the lookup map is non-owning: clearing it never closes a bitmap (no double close)', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage);
    cache.clear(); // dropping lookup references must NOT close the shared bitmap
    await flush();
    expect(close).not.toHaveBeenCalled();

    // The shared cache still owns and closes them exactly once.
    dropBitmapCacheByPath(fetchImage);
    await flush();
    expect(close).toHaveBeenCalledTimes(2);
  });

  it('a <a:duotone> recolour is owned by the shared duotone cache; dropDuotoneBitmapCache closes it', async () => {
    const baseClose = vi.fn();
    const duoClose = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (src: unknown) =>
        (src instanceof Blob
          ? { width: 4, height: 4, close: baseClose }
          : { width: 4, height: 4, close: duoClose }) as unknown as ImageBitmap,
      ),
    );
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const record: { out?: Uint8ClampedArray } = {};
    const ws = worksheetWithImages();
    ws.images = [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.png',
        mimeType: 'image/png',
        duotone: { clr1: '000000', clr2: 'FFF3F4' },
      },
    ];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage, { offscreenFactory: recordingFactory(record) });

    // The recoloured variant is what the draw site looks up (colour-suffixed key).
    expect(cache.get('xl/media/image1.png|duo:000000:FFF3F4')).toBeTruthy();

    // The recolour raster is owned by the duotone cache …
    dropDuotoneBitmapCache(fetchImage);
    await flush();
    expect(duoClose).toHaveBeenCalledTimes(1);
    // … and the colour-free base by the base cache.
    dropBitmapCacheByPath(fetchImage);
    await flush();
    expect(baseClose).toHaveBeenCalledTimes(1);
  });

  it('an SVG vector original is owned by the shared SVG cache (dropSvgImageCache clears it)', async () => {
    // Route the picture through the SVG decoder: an svg mime with no separate
    // raster blip. getCachedSvgImageByPath decodes via <img>, which node lacks,
    // so the decode rejects and prefetchImages swallows it — the point here is
    // that the SVG object-URL cache is keyed by `fetchImage` and released by
    // dropSvgImageCache, not by the lookup map.
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    ws.images = [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.svg',
        mimeType: 'image/svg+xml',
      },
    ];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    await expect(prefetchImages(ws, cache, fetchImage)).resolves.toBeUndefined();
    // Releasing the SVG cache for this fetchImage must not throw (no leak of the
    // per-document object URLs).
    expect(() => dropSvgImageCache(fetchImage)).not.toThrow();
  });
});

// ── <a:duotone> recolour at decode time (§20.1.8.23) ─────────────────────────
// A picture carrying a duotone effect is decoded once, recoloured along the
// clr1→clr2 luminance ramp, and cached under a colour-suffixed key so the raw
// blip and its recoloured variant never collide. `applyDuotone` reads the base
// bitmap's pixels via getImageData → transform → putImageData → a NEW bitmap, so
// we inject an offscreen factory + stub createImageBitmap to exercise the path
// without a real canvas.
describe('render-orchestrator duotone (§20.1.8.23)', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('recolours a duotone picture and caches it under a colour-suffixed key', async () => {
    // A fake bitmap exposes width/height so imageNaturalSize sizes the surface.
    const baseBitmap = { width: 4, height: 4, tag: 'base' } as unknown as ImageBitmap;
    const recoloured = { width: 4, height: 4, tag: 'duo' } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: unknown) => {
      // The base decode passes a Blob; applyDuotone passes the offscreen surface.
      return src instanceof Blob ? baseBitmap : recoloured;
    }));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    ws.images = [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.png',
        mimeType: 'image/png',
        alpha: 0.7,
        duotone: { clr1: '000000', clr2: 'FFF3F4' },
      },
    ];
    ws.shapeGroups = [];
    const record: { out?: Uint8ClampedArray } = {};
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage, {
      offscreenFactory: recordingFactory(record),
    });

    // Cached under path + duotone colours (NOT the bare path).
    const key = 'xl/media/image1.png|duo:000000:FFF3F4';
    expect(cache.has(key)).toBe(true);
    expect(cache.has('xl/media/image1.png')).toBe(false);
    // The cached source is the recoloured bitmap, not the base.
    expect(cache.get(key)).toBe(recoloured);
    // putImageData saw the recoloured buffer: near-white (246) → toward FFF3F4
    // (R=0xFF=255, G=0xF3=243, B=0xF4=244), so R>G and R>B and all high.
    expect(record.out).toBeDefined();
    const out = record.out as Uint8ClampedArray;
    expect(out[0]).toBeGreaterThan(240); // R
    expect(out[0]).toBeGreaterThanOrEqual(out[1]); // R>=G
    expect(out[0]).toBeGreaterThanOrEqual(out[2]); // R>=B
  });

  it('keeps a duotone variant separate from the same path without duotone', async () => {
    const baseBitmap = { width: 2, height: 2 } as unknown as ImageBitmap;
    const recoloured = { width: 2, height: 2, tag: 'duo' } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: unknown) =>
      src instanceof Blob ? baseBitmap : recoloured,
    ));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    ws.images = [
      { fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0, nativeExtCx: 0, nativeExtCy: 0, imagePath: 'xl/media/image1.png', mimeType: 'image/png' },
      { fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0, nativeExtCx: 0, nativeExtCy: 0, imagePath: 'xl/media/image1.png', mimeType: 'image/png', duotone: { clr1: '000000', clr2: 'FFF3F4' } },
    ];
    ws.shapeGroups = [];
    const record: { out?: Uint8ClampedArray } = {};
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage, { offscreenFactory: recordingFactory(record) });

    expect(cache.has('xl/media/image1.png')).toBe(true); // plain
    expect(cache.has('xl/media/image1.png|duo:000000:FFF3F4')).toBe(true); // recoloured
    expect(cache.size).toBe(2);
  });
});

// ── Render-pass liveness: LRU eviction must never hand the draw a closed bitmap ─
// The shared base cache is LRU-bounded (256): a single prefetch pass resolving
// MORE images than the cap evicts — and GPU-closes — bitmaps decoded earlier in
// the SAME pass, while the lookup map still references them for the synchronous
// draw. renderWorksheetViewport therefore holds a core render-pass lease
// (acquireBitmapCacheLease) across prefetch→draw: evictions still remove cache
// entries (bounded size), but their closes are deferred until the pass ends. The
// failure path is pinned too: a re-resolve that fails must DELETE the stale
// lookup entry (the prior bitmap may have been evicted+closed), because the
// renderer skips only a missing/falsy source, not a closed one.
describe('render-pass lease: >cap prefetch never draws a closed bitmap', () => {
  afterEach(() => vi.unstubAllGlobals());

  /** A fake HTMLCanvas + proxy 2D context whose drawImage records the closed
   *  state of every image it is handed AT DRAW TIME. All other context members
   *  no-op (the resize-test pattern). */
  function makeRecordingCanvas(drawn: {
    closedAtDraw: boolean[];
    texts?: string[];
    strokes?: number;
  }) {
    const target: Record<string, unknown> = {
      drawImage: (img: unknown) => {
        drawn.closedAtDraw.push(Boolean((img as { closed?: boolean }).closed));
      },
      fillText: (text: string) => drawn.texts?.push(text),
      stroke: () => {
        if (drawn.strokes !== undefined) drawn.strokes++;
      },
      measureText: (s: string) => ({ width: [...String(s)].length * 7 }),
      createLinearGradient: () => ({ addColorStop() {} }),
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(4) }),
      setTransform: () => undefined,
    };
    const ctx = new Proxy(target, {
      get(t, prop: string) {
        if (prop in t) return t[prop];
        return () => undefined;
      },
      set(t, prop: string, value: unknown) {
        t[prop] = value;
        return true;
      },
    });
    const canvas = {
      width: 0,
      height: 0,
      clientWidth: 800,
      clientHeight: 600,
      style: {} as Record<string, string>,
      getContext: () => ctx as unknown as CanvasRenderingContext2D,
    };
    return canvas as unknown as HTMLCanvasElement;
  }

  const STYLES = {
    fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [],
  } as unknown as ParsedWorkbook['styles'];

  it('renders the missing-codec TIFF placeholder without blanking the worksheet', async () => {
    const fetchImage = vi.fn(async () =>
      new Blob([tiffHeader(32, 24) as BlobPart], { type: 'image/tiff' }));
    const ws = {
      name: 'S', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 64, defaultRowHeight: 20,
      mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], charts: [], shapeGroups: [],
      images: [{
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/optional.tiff',
        mimeType: 'image/tiff',
      } as ImageAnchor],
    } as unknown as Worksheet;
    const drawn = { closedAtDraw: [] as boolean[], texts: [] as string[] };

    await expect(renderWorksheetViewport(
      { ws, styles: STYLES },
      makeRecordingCanvas(drawn),
      { row: 1, col: 1, rows: 10, cols: 10 },
      { fetchImage, width: 800, height: 600, dpr: 1 },
    )).resolves.toBeUndefined();

    expect(drawn.texts).toContain('TIFF image unavailable');
    expect(drawn.closedAtDraw).toEqual([]);
  });

  it('keeps cells, gridlines, and healthy pictures when a configured TIFF codec rejects one picture', async () => {
    const healthy = { width: 32, height: 24, closed: false, close() {} } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async () => healthy));
    const failure = new TiffDecodeError('Unsupported TIFF Predictor: 2');
    const render = vi.fn(async () => {
      throw failure;
    });
    const fetchImage = vi.fn(async (path: string, mime: string) => new Blob(
      [path.endsWith('.tiff') ? tiffHeader(32, 24) as BlobPart : 'healthy-png'],
      { type: mime },
    ));
    const ws = {
      name: 'S',
      rows: [{
        index: 1,
        height: null,
        cells: [{ row: 1, col: 1, styleIndex: 0, value: { type: 'text', text: 'cell survives' } }],
      }],
      colWidths: {}, rowHeights: {},
      defaultColWidth: 64, defaultRowHeight: 20,
      mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], charts: [], shapeGroups: [],
      images: [{
        fromCol: 2, fromColOff: 0, fromRow: 2, fromRowOff: 0,
        toCol: 4, toColOff: 0, toRow: 4, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/healthy.png',
        mimeType: 'image/png',
      }, {
        fromCol: 5, fromColOff: 0, fromRow: 5, fromRowOff: 0,
        toCol: 7, toColOff: 0, toRow: 7, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/unsupported.tiff',
        mimeType: 'image/tiff',
      }] as ImageAnchor[],
    } as unknown as Worksheet;
    const drawn = { closedAtDraw: [] as boolean[], texts: [] as string[], strokes: 0 };

    await expect(renderWorksheetViewport(
      { ws, styles: STYLES, tiff: { render } },
      makeRecordingCanvas(drawn),
      { row: 1, col: 1, rows: 10, cols: 10 },
      { fetchImage, width: 800, height: 600, dpr: 1 },
    )).resolves.toBeUndefined();

    expect(drawn.texts).toContain('cell survives');
    expect(drawn.texts).toContain('TIFF image unavailable');
    expect(drawn.closedAtDraw).toEqual([false]);
    expect(drawn.strokes).toBeGreaterThan(0);
  });

  it('draws 300 images (cap 256) in one pass with every bitmap still open; evicted ones close after the pass', async () => {
    // Each decode yields a bitmap with a live `closed` flag the recording
    // drawImage reads at draw time.
    const bitmaps: Array<{ closed: boolean }> = [];
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      const bmp = {
        width: 4,
        height: 4,
        closed: false,
        close() { this.closed = true; },
      };
      bitmaps.push(bmp);
      return bmp as unknown as ImageBitmap;
    }));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );

    const N = 300; // > IMAGE_BITMAP_CACHE_MAX (256) → forces mid-pass evictions
    const images: ImageAnchor[] = [];
    for (let i = 0; i < N; i++) {
      images.push({
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: `xl/media/lease-${i}.png`,
        mimeType: 'image/png',
      } as ImageAnchor);
    }
    const ws = {
      name: 'S', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 64, defaultRowHeight: 20,
      mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], charts: [], images, shapeGroups: [],
    } as unknown as Worksheet;

    const drawn = { closedAtDraw: [] as boolean[] };
    const canvas = makeRecordingCanvas(drawn);
    await renderWorksheetViewport(
      { ws, styles: STYLES },
      canvas,
      { row: 1, col: 1, rows: 10, cols: 10 },
      { fetchImage, width: 800, height: 600, dpr: 1 },
    );

    // Sanity: the pass really decoded past the cap and really drew the anchors.
    expect(bitmaps.length).toBe(N);
    expect(drawn.closedAtDraw.length).toBe(N);
    // The pinned property: NO bitmap handed to drawImage was closed at draw time.
    expect(drawn.closedAtDraw.every((c) => c === false)).toBe(true);

    // After the pass (lease released), the mid-pass evictions' deferred closes
    // run: exactly N − cap bitmaps close, proving eviction did happen and was
    // deferred (not suppressed).
    await flush();
    const closedAfter = bitmaps.filter((b) => b.closed).length;
    expect(closedAfter).toBe(N - 256);

    dropBitmapCacheByPath(fetchImage);
  });

  it('prefetchImages deletes the stale lookup entry when a re-resolve fails', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({
      width: 4, height: 4, close() {},
    }) as unknown as ImageBitmap));
    let fail = false;
    const fetchImage = vi.fn(async (path: string, mime: string) => {
      if (fail) throw new Error('byte source unavailable');
      return new Blob([new TextEncoder().encode(path)], { type: mime });
    });
    const ws = worksheetWithImages();
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();

    // Pass 1: healthy decode lands in the lookup map.
    await prefetchImages(ws, cache, fetchImage);
    expect(cache.has('xl/media/image1.png')).toBe(true);

    // The shared entry is evicted+closed (LRU pressure / drop) …
    dropBitmapCacheByPath(fetchImage);
    await flush();
    // … and the re-resolve on the next pass fails. The stale lookup entry MUST
    // be deleted — its bitmap may be closed, and the renderer only skips a
    // missing/falsy source.
    fail = true;
    await prefetchImages(ws, cache, fetchImage);
    expect(cache.has('xl/media/image1.png')).toBe(false);
  });

  it('replaces a duotone pass-through bitmap after base-cache eviction between passes', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      const bmp = {
        width: 4,
        height: 4,
        closed: false,
        close(): void { this.closed = true; },
      };
      return bmp as unknown as ImageBitmap;
    }));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    ws.images = [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.png',
        mimeType: 'image/png',
        duotone: { clr1: '000000', clr2: 'FFF3F4' },
      },
    ];
    ws.shapeGroups = [];
    const cache = new Map<string, CanvasImageSource | null>();
    const key = 'xl/media/image1.png|duo:000000:FFF3F4';

    await prefetchImages(ws, cache, fetchImage);
    const pass1 = cache.get(key) as ImageBitmap & { closed: boolean };
    expect(pass1).toBeDefined();
    expect(pass1.closed).toBe(false);

    for (let i = 0; i < 256; i++) {
      await getCachedBitmapByPath(`xl/media/pressure-${i}.png`, 'image/png', fetchImage);
    }
    await flush();
    expect(pass1.closed).toBe(true);

    await prefetchImages(ws, cache, fetchImage);
    const pass2 = cache.get(key) as ImageBitmap & { closed: boolean };
    expect(pass2).toBeDefined();
    expect(pass2.closed).toBe(false);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });
});
