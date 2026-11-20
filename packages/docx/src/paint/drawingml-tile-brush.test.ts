import { afterEach, describe, expect, it, vi } from 'vitest';
import { getCachedBitmapByPath, OoxmlDecodedImageLimitError } from '@silurus/ooxml-core';
import type { ImagePaintResourceDescriptor } from '../layout/types.js';
import { createPaintResourceRegistry } from '../layout/paint-resources.js';
import { acquireDrawingMLTileBrush } from './browser-vml-image-brush.js';
import { drawingMLTileCellKey, imageKey, preloadPaintImages } from './browser-images.js';
import { paintImageResource } from './image-resource.js';
import { createProductionPaintResourceSession } from './resource-session.js';
import { drawingMLTileOrigin } from './vml-image-brush.js';

// ECMA-376 §20.1.8.58 CT_TileInfoProperties on a shape picture fill.

const source = vi.hoisted(() => ({ width: 40, height: 20 }) as unknown as ImageBitmap);

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return {
    ...actual,
    getCachedBitmapByPath: vi.fn(async () => source),
    getCachedDerivedBitmap: vi.fn(async (_namespace: string, _key: string, _owner: unknown,
      create: () => Promise<{ bitmap: unknown }>) => (await create()).bitmap),
  };
});

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

interface DrawCall { args: unknown[]; transform: string }

/** Scratch surface stub recording mirrored draws into the tile cell. */
function stubScratchCanvas(): { sizes: number[][]; draws: DrawCall[]; baked: number[][]; finalSizes: () => number[][] } {
  const sizes: number[][] = [];
  const draws: DrawCall[] = [];
  const baked: number[][] = [];
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
  vi.stubGlobal('createImageBitmap', vi.fn(async (canvas: Scratch) => {
    baked.push([canvas.width, canvas.height]);
    return { width: canvas.width, height: canvas.height, close() {} };
  }));
  return { sizes, draws, baked, finalSizes: () => surfaces.map((surface) => [surface.width, surface.height]) };
}

async function prepare(tileDescriptor: ImagePaintResourceDescriptor) {
  return preloadPaintImages([tileDescriptor],
    [{ resourceKind: 'image', resourceKey: tileDescriptor.resourceKey, widthPt: 100, heightPt: 40 }],
    async () => new Blob());
}

afterEach(() => { vi.unstubAllGlobals(); });

describe('DrawingML tile brush', () => {
  it('scales the native 96-DPI tile, aligns it in the painted box, then offsets it', () => {
    // 40x20 px at 96 DPI is 30x15 pt; sx 50% gives 15x15 pt. ctr registers it
    // at ((100-15)/2, (40-15)/2) and tx/ty add 1pt / -2pt.
    const brush = acquireDrawingMLTileBrush(descriptor({
      tx: 12_700, ty: -25_400, sx: 0.5, sy: 1, flip: 'none', algn: 'ctr',
    }), source);
    expect(brush).toMatchObject({
      kind: 'drawingml-tile-brush', image: source, pixelWidth: 40, pixelHeight: 20,
      tileWidthPt: 15, tileHeightPt: 15,
    });
    const origin = drawingMLTileOrigin(brush, BOX);
    expect(origin.xPt).toBeCloseTo(43.5);
    expect(origin.yPt).toBeCloseTo(10.5);
    // The same resource re-registers against any painted extent.
    expect(drawingMLTileOrigin(brush, { xPt: 5, yPt: 0, widthPt: 612, heightPt: 792 }).xPt)
      .toBeCloseTo(5 + (612 - 15) / 2 + 1);
  });

  it('reads absent placement attributes as the identity and right/bottom alignments from the far edge', () => {
    const identity = acquireDrawingMLTileBrush(descriptor({ flip: 'none' }), source);
    expect([identity.tileWidthPt, identity.tileHeightPt]).toEqual([30, 15]);
    expect(drawingMLTileOrigin(identity, BOX)).toEqual({ xPt: 0, yPt: 0 });
    const bottomRight = acquireDrawingMLTileBrush(descriptor({ flip: 'none', algn: 'br' }), source);
    expect(drawingMLTileOrigin(bottomRight, BOX)).toEqual({ xPt: 70, yPt: 25 });
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
    expect(scratch.baked).toEqual([[80, 40]]);
    expect(scratch.finalSizes()).toEqual([[1, 1]]);
    const cell = decoded.get(drawingMLTileCellKey(tiled)) as unknown as CanvasImageSource;
    const brush = acquireDrawingMLTileBrush(tiled, decoded.get(imageKey(tiled.partPath)) as ImageBitmap, cell);
    expect(brush.image).toBe(cell);
    expect([brush.pixelWidth, brush.pixelHeight]).toEqual([40, 20]);
  });

  it('repeats the §20.1.8.55 cropped source', async () => {
    const scratch = stubScratchCanvas();
    const cropped = descriptor({ flip: 'none' }, { srcRect: { l: 0.25, t: 0, r: 0, b: 0.5 } });
    await prepare(cropped);
    expect(scratch.sizes).toEqual([[30, 10]]);
    expect(scratch.draws[0]!.args.slice(1)).toEqual([10, 0, 30, 10, 0, 0, 30, 10]);
    const brush = acquireDrawingMLTileBrush(cropped, source, {} as CanvasImageSource);
    expect(brush.tileWidthPt).toBeCloseTo(22.5);
    expect(brush.tileHeightPt).toBeCloseTo(7.5);
  });

  it('refuses a cell over the shared surface ceilings before allocating it', async () => {
    const scratch = stubScratchCanvas();
    // A negative crop inset amplifies the 40x20 source to 200040 px wide.
    await expect(prepare(descriptor({ flip: 'xy' }, { srcRect: { l: -2500, t: 0, r: -2500, b: 0 } })))
      .rejects.toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(scratch.sizes).toEqual([]);
  });

  it('paints a production-resolved brush as a pattern anchored at the tile origin', () => {
    const tiled = descriptor({ tx: 12_700, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' });
    const session = createProductionPaintResourceSession(
      createPaintResourceRegistry([tiled]),
      (candidate) => acquireDrawingMLTileBrush(candidate as ImagePaintResourceDescriptor, source),
    );
    const calls: unknown[][] = [];
    const ctx = {
      globalAlpha: 1,
      save: () => {}, restore: () => {},
      translate: (...args: number[]) => { calls.push(['translate', ...args]); },
      scale: (...args: number[]) => { calls.push(['scale', ...args]); },
      createPattern: (image: unknown, repetition: string) => {
        calls.push(['pattern', image === source, repetition]);
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

  it('keeps the native source grid under display-resolution decoding', async () => {
    // A display target would shrink the bitmap and with it the tile period.
    await preloadPaintImages([descriptor({ flip: 'none' })],
      [{ resourceKind: 'image', resourceKey: 'image:tile', widthPt: 612, heightPt: 792 }],
      async () => new Blob(), undefined, 0.1, undefined, { resolution: 'display' });
    const options = vi.mocked(getCachedBitmapByPath).mock.calls.at(-1)![3] as Record<string, unknown>;
    expect(options.targetWidthPx).toBeUndefined();
    expect(options.targetHeightPx).toBeUndefined();
  });
});
