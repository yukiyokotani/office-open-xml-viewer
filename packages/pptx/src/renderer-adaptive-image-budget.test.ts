import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderSlide, renderSlideWithEmbeddedFonts } from './renderer.js';
import type { PictureElement, Slide } from './types.js';

function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function canvas(): HTMLCanvasElement {
  const state: Record<PropertyKey, unknown> = { fillStyle: '', globalAlpha: 1 };
  const context = new Proxy(state, {
    get(target, property) {
      if (property in target) return target[property];
      if (property === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      return () => undefined;
    },
    set(target, property, value) {
      target[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  const result = {
    width: 0,
    height: 0,
    style: {},
    offsetWidth: 960,
    getContext: () => context,
  } as unknown as HTMLCanvasElement;
  state.canvas = result;
  return result;
}

function picture(path: string, x: number): PictureElement {
  return {
    type: 'picture',
    x,
    y: 0,
    width: 4_572_000,
    height: 6_858_000,
    rotation: 0,
    flipH: false,
    flipV: false,
    imagePath: path,
    mimeType: 'image/png',
    stroke: null,
  };
}

function twoPictureSlide(): Slide {
  return {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [
      picture('ppt/media/left.png', 0),
      picture('ppt/media/right.png', 4_572_000),
    ],
  } as Slide;
}

describe('PPTX adaptive decoded-image budget', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('allocates one bounded quality level across every raster on the slide', async () => {
    const decode = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 4_000,
      height: options?.resizeWidth ?? 4_000,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', decode);
    const bytes = pngHeader(4_000, 4_000);
    const fetchImage = vi.fn(async (_path: string, _mime: string) =>
      new Blob([bytes as BlobPart], { type: 'image/png' }));
    const budget = 1_048_576;

    await renderSlide(canvas(), twoPictureSlide(), 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
      imageResources: { decodedByteBudget: budget, strategy: 'adaptive' },
    });

    expect(decode).toHaveBeenCalledTimes(2);
    const targets = decode.mock.calls.map((call) => ({
      width: call[1]?.resizeWidth as number,
      height: call[1]?.resizeWidth as number,
    }));
    expect(targets[0]).toEqual(targets[1]);
    expect(targets[0].width * targets[0].height * 4 * targets.length)
      .toBeLessThanOrEqual(budget);
  });

  it('preserves native decoding for an ordinary slide raster that fits', async () => {
    const decode = vi.fn(async () => ({
      width: 800,
      height: 600,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', decode);
    const blob = new Blob([pngHeader(800, 600) as BlobPart], { type: 'image/png' });
    const fetchImage = vi.fn(async () => blob);
    const slide = { ...twoPictureSlide(), elements: [picture('ppt/media/photo.png', 0)] };

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage,
    });

    expect(decode).toHaveBeenCalledWith(blob);
  });

  it('renders the reported 109,571,670-pixel picture at its visible target', async () => {
    const decode = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 12_090,
      height: options?.resizeWidth
        ? Math.ceil(9_063 * options.resizeWidth / 12_090)
        : 9_063,
      close() {},
    }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', decode);
    const bytes = pngHeader(12_090, 9_063);
    expect(12_090 * 9_063).toBe(109_571_670);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/png' }));
    const slide = { ...twoPictureSlide(), elements: [picture('ppt/media/large-photo.png', 0)] };

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
    });

    expect(decode).toHaveBeenCalledWith(expect.any(Blob), {
      resizeWidth: 1_921,
      resizeQuality: 'high',
    });
  });

  it('does not apply the display-sized pixel cap to a TIFF codec result', async () => {
    const bitmap = { width: 1000, height: 1000, close() {} } as unknown as ImageBitmap;
    const render = vi.fn(async () => bitmap);
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
    const fetchImage = vi.fn(async () => new Blob([bytes as BlobPart], { type: 'image/tiff' }));
    const element = { ...picture('ppt/media/photo.tiff', 0), mimeType: 'image/tiff' };
    const slide = { ...twoPictureSlide(), elements: [element] };

    await expect(renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage,
      tiff: { render },
    })).resolves.toBeDefined();
    expect(render).toHaveBeenCalledOnce();
  });

  it('keeps the authored display target for a worker-decoded SVG outside the raster plan', async () => {
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
    const element = { ...picture('ppt/media/icon.svg', 0), mimeType: 'image/svg+xml' };
    const slide = { ...twoPictureSlide(), elements: [element] };

    await renderSlideWithEmbeddedFonts(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
      svgDecoder,
    });

    expect(svgDecoder).toHaveBeenCalledWith(expect.any(Blob), {
      targetWidthPx: 960,
      targetHeightPx: 1440,
    });
  });

  it('keeps a strict mode that surfaces the typed aggregate quota crossing', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
      width: options?.resizeWidth ?? 4_000,
      height: options?.resizeHeight ?? 4_000,
      close() {},
    }) as unknown as ImageBitmap));
    const bytes = pngHeader(4_000, 4_000);
    const fetchImage = vi.fn(async (_path: string, _mime: string) =>
      new Blob([bytes as BlobPart], { type: 'image/png' }));

    await expect(renderSlide(canvas(), twoPictureSlide(), 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
      imageResources: { decodedByteBudget: 1_048_576, strategy: 'strict' },
    })).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'active-decoded-bytes',
    });
  });

  it('drops a superseded canvas render while it is still waiting for image admission', async () => {
    let releaseFirst!: () => void;
    const firstDecode = new Promise<void>((resolve) => { releaseFirst = resolve; });
    let decodeIndex = 0;
    vi.stubGlobal('createImageBitmap', vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      if (decodeIndex++ === 0) await firstDecode;
      return {
        width: options?.resizeWidth ?? 100,
        height: options?.resizeHeight ?? 100,
        close() {},
      } as unknown as ImageBitmap;
    }));
    const bytes = pngHeader(100, 100);
    const fetchImage = vi.fn(async (_path: string, _mime: string) =>
      new Blob([bytes as BlobPart], { type: 'image/png' }));
    const single = (path: string): Slide => ({
      ...twoPictureSlide(),
      elements: [picture(path, 0)],
    });
    const blockerCanvas = canvas();
    const reusedCanvas = canvas();

    const blocker = renderSlide(blockerCanvas, single('ppt/media/blocker.png'), 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    await vi.waitFor(() => expect(globalThis.createImageBitmap).toHaveBeenCalledTimes(1));
    const stale = renderSlide(reusedCanvas, single('ppt/media/stale.png'), 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });
    const newest = renderSlide(reusedCanvas, single('ppt/media/newest.png'), 9_144_000, 6_858_000, {
      width: 960, dpr: 1, fetchImage,
    });

    releaseFirst();
    await Promise.all([blocker, stale, newest]);
    expect(fetchImage.mock.calls.map(([path]) => path)).not.toContain('ppt/media/stale.png');
    expect(fetchImage.mock.calls.map(([path]) => path)).toContain('ppt/media/newest.png');
  });
});
