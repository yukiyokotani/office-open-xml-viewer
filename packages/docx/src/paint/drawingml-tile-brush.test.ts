import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  OoxmlDecodedImageLimitError,
} from '@silurus/ooxml-core';
import type { ImagePaintResourceDescriptor } from '../layout/types.js';
import { createPaintResourceRegistry } from '../layout/paint-resources.js';
import { acquireDrawingMLTileBrush } from './browser-vml-image-brush.js';
import { drawingMLTileCellKey, paintImageDecodeKey, preloadPaintImages } from './browser-images.js';
import { paintImageResource } from './image-resource.js';
import { createProductionPaintResourceSession } from './resource-session.js';
import { drawingMLTileOrigin, type DrawingMLTileBrush } from './vml-image-brush.js';

// ECMA-376 §20.1.8.58 CT_TileInfoProperties on a shape picture fill.

const fixtures = vi.hoisted(() => ({
  source: { width: 40, height: 20 } as object,
  derived: new Map<string, unknown>(),
  effectFrames: new Map<object, { x: number; y: number; width: number; height: number }>(),
  outsets: new Map<object, { x: number; y: number }>(),
  resolution: null as { dpiX: number; dpiY: number } | null,
}));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return {
    ...actual,
    getCachedBitmapByPath: vi.fn(async () => fixtures.source),
    getCachedDuotoneBitmapByPath: vi.fn(async () => fixtures.source),
    inspectCachedRasterSource: vi.fn(async () => ({
      format: 'png', dimensions: null, resolution: fixtures.resolution,
    })),
    // The shared derived-surface cache, reduced to its key contract.
    getCachedDerivedBitmap: vi.fn(async (_namespace: string, key: string, _owner: unknown,
      create: () => Promise<{ bitmap: unknown }>) => {
      if (!fixtures.derived.has(key)) fixtures.derived.set(key, (await create()).bitmap);
      return fixtures.derived.get(key);
    }),
    // Decoded effect geometry is supplied by the decoder; tests state it.
    imageEffectFrame: vi.fn((image: object) => fixtures.effectFrames.get(image)),
    imageEffectOutsetPeriods: vi.fn((image: object, cropped: boolean) => (
      cropped ? { x: 0, y: 0 } : fixtures.outsets.get(image) ?? { x: 0, y: 0 })),
  };
});

const source = () => fixtures.source as unknown as ImageBitmap;
const BOX = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 40 };

function descriptor(
  tile: NonNullable<ImagePaintResourceDescriptor['tile']>,
  extra: Partial<ImagePaintResourceDescriptor> = {},
): ImagePaintResourceDescriptor {
  return {
    kind: 'image', resourceKey: 'image:tile', partPath: 'word/media/tile.png', mimeType: 'image/png',
    intrinsicSize: { widthPt: 100, heightPt: 40 }, tile, ...extra,
  };
}

function brushOf(
  tileDescriptor: ImagePaintResourceDescriptor,
  image: CanvasImageSource = source(),
  cell?: CanvasImageSource | null,
): Readonly<DrawingMLTileBrush> {
  const brush = acquireDrawingMLTileBrush(tileDescriptor, image, cell);
  if (!brush) throw new Error('expected a tile brush');
  return brush;
}

interface DrawCall { args: unknown[]; transform: string }

/** Scratch surface stub recording mirrored draws into the tile cell. */
function stubScratchCanvas() {
  const sizes: number[][] = [];
  const draws: DrawCall[] = [];
  const surfaces: Array<{ width: number; height: number }> = [];
  class Scratch {
    constructor(public width: number, public height: number) {
      sizes.push([width, height]);
      surfaces.push(this);
    }
    getContext() {
      const stack: string[] = [];
      let transform = '';
      return {
        save: () => { stack.push(transform); },
        restore: () => { transform = stack.pop() ?? ''; },
        translate: (x: number, y: number) => { transform += `t(${x},${y})`; },
        scale: (x: number, y: number) => { transform += `s(${x},${y})`; },
        drawImage: (...args: unknown[]) => { draws.push({ args, transform }); },
      };
    }
  }
  vi.stubGlobal('OffscreenCanvas', Scratch);
  vi.stubGlobal('createImageBitmap', vi.fn(async (canvas: Scratch) => (
    { width: canvas.width, height: canvas.height, close() {} })));
  return { sizes, draws, finalSizes: () => surfaces.map((surface) => [surface.width, surface.height]) };
}

async function prepare(tileDescriptor: ImagePaintResourceDescriptor) {
  return preloadPaintImages([tileDescriptor],
    [{ resourceKind: 'image', resourceKey: tileDescriptor.resourceKey, widthPt: 100, heightPt: 40 }],
    async () => new Blob());
}

beforeEach(() => {
  fixtures.source = { width: 40, height: 20 };
  fixtures.derived.clear();
  fixtures.effectFrames.clear();
  fixtures.outsets.clear();
  fixtures.resolution = null;
});
afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

describe('DrawingML tile brush', () => {
  it('scales the native tile, aligns it in the painted box, then offsets it', () => {
    // 40x20 px at 96 DPI is 30x15 pt; sx 50% gives 15x15 pt. ctr registers it
    // at ((100-15)/2, (40-15)/2) and tx/ty add 1pt / -2pt.
    const brush = brushOf(descriptor({ tx: 12_700, ty: -25_400, sx: 0.5, sy: 1, flip: 'none', algn: 'ctr' }));
    expect(brush).toMatchObject({
      kind: 'drawingml-tile-brush', image: source(), pixelWidth: 40, pixelHeight: 20,
      tileWidthPt: 15, tileHeightPt: 15,
    });
    const origin = drawingMLTileOrigin(brush, BOX);
    expect(origin.xPt).toBeCloseTo(43.5);
    expect(origin.yPt).toBeCloseTo(10.5);
    expect(drawingMLTileOrigin(brush, { xPt: 5, yPt: 0, widthPt: 612, heightPt: 792 }).xPt)
      .toBeCloseTo(5 + (612 - 15) / 2 + 1);
  });

  it('reads absent placement attributes as the identity and right/bottom alignments from the far edge', () => {
    const identity = brushOf(descriptor({ flip: 'none' }));
    expect([identity.tileWidthPt, identity.tileHeightPt]).toEqual([30, 15]);
    expect(drawingMLTileOrigin(identity, BOX)).toEqual({ xPt: 0, yPt: 0 });
    expect(drawingMLTileOrigin(brushOf(descriptor({ flip: 'none', algn: 'br' })), BOX))
      .toEqual({ xPt: 70, yPt: 25 });
  });

  it('sizes the tile at the authored §20.1.8.14 blipFill DPI', () => {
    // 192 DPI halves the physical size of 40x20 source pixels.
    const brush = brushOf(descriptor({ flip: 'none' }, { blipDpi: 192 }));
    expect([brush.tileWidthPt, brush.tileHeightPt]).toEqual([15, 7.5]);
  });

  it('uses the DPI declared by the blip when blipFill@dpi is zero or absent', async () => {
    fixtures.resolution = { dpiX: 192, dpiY: 144 };
    const tiled = descriptor({ flip: 'none' });
    const decoded = await prepare(tiled);
    const brush = brushOf(tiled, decoded.get(paintImageDecodeKey(tiled)) as CanvasImageSource);
    expect([brush.tileWidthPt, brush.tileHeightPt]).toEqual([15, 10]);
    // A blip without resolution metadata keeps the 96-DPI convention.
    fixtures.resolution = null;
    fixtures.source = { width: 40, height: 20 };
    const plain = await prepare(tiled);
    expect(brushOf(tiled, plain.get(paintImageDecodeKey(tiled)) as CanvasImageSource).tileWidthPt).toBe(30);
  });

  it('samples physical blip effects of a tile at the effective DPI', async () => {
    fixtures.resolution = { dpiX: 192, dpiY: 192 };
    await prepare(descriptor({ flip: 'none' }, {
      blipEffects: [{ type: 'grayscale' }],
      effectHost: { rotation: 0, flipH: false, flipV: false, tileDpi: 96 },
    }));
    expect(vi.mocked(getCachedDuotoneBitmapByPath).mock.calls[0]![2]).toMatchObject({
      emuPerPixelX: 914400 / 192, emuPerPixelY: 914400 / 192,
    });
  });

  it('repeats the authored frame of a grown effect surface and wraps its outset', async () => {
    // A 40x20 frame inside a 60x40 effect surface repeats at 30x15pt, not 45x30pt.
    fixtures.source = { width: 60, height: 40 };
    fixtures.effectFrames.set(fixtures.source, { x: 10, y: 10, width: 40, height: 20 });
    fixtures.outsets.set(fixtures.source, { x: 1, y: 1 });
    const scratch = stubScratchCanvas();
    const tiled = descriptor({ flip: 'none' });
    const decoded = await prepare(tiled);
    const brush = brushOf(tiled, source(), decoded.get(drawingMLTileCellKey(tiled)) as CanvasImageSource);
    expect([brush.tileWidthPt, brush.tileHeightPt]).toEqual([30, 15]);
    expect(scratch.sizes).toEqual([[40, 20]]);
    // The period and its eight neighbours, each placing the frame on its period.
    expect(scratch.draws).toHaveLength(9);
    expect(scratch.draws.map((draw) => draw.transform)).toContain('t(-40,-20)s(1,1)');
    expect(scratch.draws[0]!.args.slice(1)).toEqual([-10, -10, 60, 40]);
  });

  it('composes mirrored neighbours into one accounted cell during preparation', async () => {
    const scratch = stubScratchCanvas();
    const tiled = descriptor({ flip: 'xy', algn: 'tl' });
    const decoded = await prepare(tiled);
    expect(scratch.sizes).toEqual([[80, 40]]);
    expect(scratch.draws.map((draw) => draw.transform)).toEqual([
      't(0,0)s(1,1)', 't(80,0)s(-1,1)', 't(0,40)s(1,-1)', 't(80,40)s(-1,-1)',
    ]);
    // The scratch surface is baked into a cached bitmap, then released.
    expect(scratch.finalSizes()).toEqual([[1, 1]]);
    const cell = decoded.get(drawingMLTileCellKey(tiled)) as CanvasImageSource;
    const brush = brushOf(tiled, source(), cell);
    expect(brush.image).toBe(cell);
    expect([brush.pixelWidth, brush.pixelHeight]).toEqual([40, 20]);
  });

  it('keys a derived cell by the decoded source variant', async () => {
    const scratch = stubScratchCanvas();
    const tiled = descriptor({ flip: 'x' });
    await prepare(tiled);
    // The same source decoded at another grid (another page's target).
    fixtures.source = { width: 80, height: 40 };
    const decoded = await prepare(tiled);
    expect(scratch.sizes).toEqual([[80, 20], [160, 40]]);
    expect(decoded.get(drawingMLTileCellKey(tiled))).toMatchObject({ width: 160, height: 40 });
  });

  it('repeats the §20.1.8.55 cropped source', async () => {
    const scratch = stubScratchCanvas();
    const cropped = descriptor({ flip: 'none' }, { srcRect: { l: 0.25, t: 0, r: 0, b: 0.5 } });
    await prepare(cropped);
    expect(scratch.sizes).toEqual([[30, 10]]);
    expect(scratch.draws[0]!.args.slice(1)).toEqual([10, 0, 30, 10, 0, 0, 30, 10]);
    const brush = brushOf(cropped, source(), {} as CanvasImageSource);
    expect(brush.tileWidthPt).toBeCloseTo(22.5);
    expect(brush.tileHeightPt).toBeCloseTo(7.5);
  });

  it('keeps a fractional crop window as the physical period with one backing pixel', async () => {
    fixtures.source = { width: 1, height: 1 };
    const scratch = stubScratchCanvas();
    const cropped = descriptor({ flip: 'none' }, { srcRect: { l: 0.6, t: 0, r: 0, b: 0 } });
    const decoded = await prepare(cropped);
    expect(scratch.sizes).toEqual([[1, 1]]);
    const brush = brushOf(cropped, source(), decoded.get(drawingMLTileCellKey(cropped)) as CanvasImageSource);
    expect(brush.tileWidthPt).toBeCloseTo(0.4 * 0.75);
    expect(brush.tileHeightPt).toBeCloseTo(0.75);
    expect([brush.pixelWidth, brush.pixelHeight]).toEqual([1, 1]);
    // A window with no positive extent selects nothing.
    expect(acquireDrawingMLTileBrush(descriptor({ flip: 'none' }, { srcRect: { l: 0.6, t: 0, r: 0.4, b: 0 } }),
      source())).toBeNull();
  });

  it('refuses a cell over the shared surface ceilings before allocating it', async () => {
    const scratch = stubScratchCanvas();
    // A negative crop inset amplifies the 40x20 source to 200040 px wide.
    await expect(prepare(descriptor({ flip: 'xy' }, { srcRect: { l: -2500, t: 0, r: -2500, b: 0 } })))
      .rejects.toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(scratch.sizes).toEqual([]);
  });

  it('keeps the native source grid under display-resolution decoding', async () => {
    await preloadPaintImages([descriptor({ flip: 'none' })],
      [{ resourceKind: 'image', resourceKey: 'image:tile', widthPt: 612, heightPt: 792 }],
      async () => new Blob(), undefined, 0.1, undefined, { resolution: 'display' });
    const options = vi.mocked(getCachedBitmapByPath).mock.calls.at(-1)![3] as Record<string, unknown>;
    expect(options.targetWidthPx).toBeUndefined();
    expect(options.targetHeightPx).toBeUndefined();
  });

  it('paints a production-resolved brush as a pattern anchored at the tile origin', () => {
    const tiled = descriptor({ tx: 12_700, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' });
    const session = createProductionPaintResourceSession(
      createPaintResourceRegistry([tiled]),
      (candidate) => brushOf(candidate as ImagePaintResourceDescriptor),
    );
    const calls: unknown[][] = [];
    const ctx = {
      globalAlpha: 1,
      save: () => {}, restore: () => {},
      translate: (...args: number[]) => { calls.push(['translate', ...args]); },
      scale: (...args: number[]) => { calls.push(['scale', ...args]); },
      createPattern: (image: unknown, repetition: string) => {
        calls.push(['pattern', image === source(), repetition]);
        return { pattern: true };
      },
      fillRect: (...args: number[]) => { calls.push(['fillRect', ...args]); },
      set fillStyle(_value: unknown) { calls.push(['fillStyle']); },
    } as unknown as CanvasRenderingContext2D;
    paintImageResource(session.resolve('image:tile', 'image'),
      { xPt: 10, yPt: 20, widthPt: 100, heightPt: 40 }, ctx);
    expect(calls[0]).toEqual(['pattern', true, 'repeat']);
    expect(calls[1]![0]).toBe('translate');
    expect(calls[1]![1]).toBeCloseTo(11);
    expect(calls[1]![2]).toBeCloseTo(20);
    expect(calls[2]![0]).toBe('scale');
    expect(calls[2]![1]).toBeCloseTo(0.75);
    expect(calls.at(-1)![0]).toBe('fillRect');
  });
});
