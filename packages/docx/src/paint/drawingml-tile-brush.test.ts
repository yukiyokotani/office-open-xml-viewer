import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ImagePaintResourceDescriptor } from '../layout/types.js';
import { acquireDrawingMLTileBrush } from './browser-vml-image-brush.js';
import { paintImageResource } from './image-resource.js';
import { drawingMLTileOrigin } from './vml-image-brush.js';

const BOX = { xPt: 0, yPt: 0, widthPt: 100, heightPt: 40 };

// ECMA-376 §20.1.8.58 CT_TileInfoProperties on a shape picture fill.

const image = { width: 40, height: 20 } as unknown as CanvasImageSource;

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

/** Scratch canvas stub recording mirrored draws into the tile cell. */
function stubScratchCanvas(): { sizes: number[][]; draws: DrawCall[] } {
  const sizes: number[][] = [];
  const draws: DrawCall[] = [];
  class Scratch {
    constructor(public width: number, public height: number) { sizes.push([width, height]); }
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
  return { sizes, draws };
}

afterEach(() => { vi.unstubAllGlobals(); });

describe('DrawingML tile brush', () => {
  it('scales the native 96-DPI tile, aligns it in the box, then offsets it', () => {
    // 40x20 px at 96 DPI is 30x15 pt; sx 50% gives 15x15 pt. ctr registers it
    // at ((100-15)/2, (40-15)/2) and tx/ty add 1pt / -2pt.
    const brush = acquireDrawingMLTileBrush(descriptor({
      tx: 12_700, ty: -25_400, sx: 0.5, sy: 1, flip: 'none', algn: 'ctr',
    }), image);
    expect(brush).toMatchObject({
      kind: 'drawingml-tile-brush', image, pixelWidth: 40, pixelHeight: 20,
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
    const identity = acquireDrawingMLTileBrush(descriptor({ flip: 'none' }), image);
    expect([identity.tileWidthPt, identity.tileHeightPt]).toEqual([30, 15]);
    expect(drawingMLTileOrigin(identity, BOX)).toEqual({ xPt: 0, yPt: 0 });
    const bottomRight = acquireDrawingMLTileBrush(descriptor({ flip: 'none', algn: 'br' }), image);
    expect(drawingMLTileOrigin(bottomRight, BOX)).toEqual({ xPt: 70, yPt: 25 });
  });

  it('composes mirrored neighbours into one repeating cell with an unmirrored first tile', () => {
    const scratch = stubScratchCanvas();
    const brush = acquireDrawingMLTileBrush(descriptor({ flip: 'xy', algn: 'tl' }), image);
    expect(scratch.sizes).toEqual([[80, 40]]);
    expect(brush.image).not.toBe(image);
    expect([brush.pixelWidth, brush.pixelHeight]).toEqual([40, 20]);
    expect(scratch.draws.map((draw) => draw.transform)).toEqual([
      't(0,0)s(1,1)', 't(80,0)s(-1,1)', 't(0,40)s(1,-1)', 't(80,40)s(-1,-1)',
    ]);
  });

  it('repeats the §20.1.8.55 cropped source', () => {
    const scratch = stubScratchCanvas();
    const brush = acquireDrawingMLTileBrush(descriptor({ flip: 'none' }, {
      srcRect: { l: 0.25, t: 0, r: 0, b: 0.5 },
    }), image);
    expect(brush.tileWidthPt).toBeCloseTo(22.5);
    expect(brush.tileHeightPt).toBeCloseTo(7.5);
    expect(scratch.sizes).toEqual([[30, 10]]);
    expect(scratch.draws[0]!.args.slice(1)).toEqual([10, 0, 30, 10, 0, 0, 30, 10]);
  });

  it('paints the brush as a pattern anchored at the tile origin over the shape bounds', () => {
    const brush = acquireDrawingMLTileBrush(descriptor({ tx: 12_700, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' }), image);
    const calls: unknown[][] = [];
    const ctx = {
      globalAlpha: 1,
      save: () => {}, restore: () => {},
      translate: (...args: number[]) => { calls.push(['translate', ...args]); },
      scale: (...args: number[]) => { calls.push(['scale', ...args]); },
      createPattern: (source: unknown, repetition: string) => {
        calls.push(['pattern', source === image, repetition]);
        return { pattern: true };
      },
      fillRect: (...args: number[]) => { calls.push(['fillRect', ...args]); },
      set fillStyle(_value: unknown) { calls.push(['fillStyle']); },
    } as unknown as CanvasRenderingContext2D;
    paintImageResource({ descriptor: descriptor({ flip: 'none' }), handle: brush },
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
