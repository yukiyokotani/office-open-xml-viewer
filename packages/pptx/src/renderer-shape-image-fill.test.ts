import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Slide } from './types.js';

const mocks = vi.hoisted(() => ({
  image: { width: 2, height: 1, close() {} } as unknown as ImageBitmap,
  decode: vi.fn(),
}));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return { ...actual, getCachedDuotoneBitmapByPath: mocks.decode };
});

import { paintPreparedShapeImageFill, renderSlide } from './renderer.js';

function recordingCanvas() {
  const calls: Array<[string, ...unknown[]]> = [];
  const state: Record<string, unknown> = { globalAlpha: 1, fillStyle: '', strokeStyle: '' };
  const stack: Array<Record<string, unknown>> = [];
  const ctx = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      if (property === 'save') return () => stack.push({ ...target });
      if (property === 'restore') return () => Object.assign(target, stack.pop());
      if (property === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (property === 'measureText') return (text: string) => ({ width: text.length * 8 });
      return (...args: unknown[]) => { calls.push([property, ...args]); };
    },
    set(target, property: string, value) { target[property] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const canvas = {
    width: 0, height: 0, style: {}, offsetWidth: 960,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
  return { canvas, ctx, calls, state };
}

function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function shape(imagePath: string, x = 0, withText = false) {
  return {
    type: 'shape', x, y: 0, width: 914_400, height: 914_400,
    rotation: 0, flipH: false, flipV: false, geometry: 'ellipse',
    fill: {
      fillType: 'image', imagePath, mimeType: 'image/png', stretch: true,
      rotWithShape: true, alpha: 0.5,
    },
    stroke: { color: '000000', width: 12_700 },
    textBody: withText ? {
      verticalAnchor: 't', paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null,
        spaceAfter: null, spaceLine: null, lvl: 0, bullet: { type: 'inherit' },
        defFontSize: null, defColor: null, defBold: null, defItalic: null,
        defFontFamily: null, tabStops: [], eaLnBrk: true,
        runs: [{
          type: 'text', text: 'Picture fill text', bold: null, italic: null,
          underline: false, strikethrough: false, fontSize: 18,
          color: null, fontFamily: null,
        }],
      }], defaultFontSize: 18, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
    } : null,
    defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null,
    adj5: null, adj6: null, adj7: null, adj8: null, shadow: null,
  } as const;
}

describe('PPTX ordinary-shape image fills', () => {
  beforeEach(() => {
    mocks.decode.mockReset();
    mocks.decode.mockResolvedValue(mocks.image);
  });

  it('clips the asymmetric source to the current path and isolates alpha/crop state', () => {
    const { ctx, calls, state } = recordingCanvas();
    const painted = paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/asymmetric.png', mimeType: 'image/png',
      stretch: true, alpha: 0.5,
      srcRect: { l: 0.5, t: 0, r: 0, b: 0 },
      fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 },
    }, { image: mocks.image }, { x: 10, y: 20, w: 100, h: 50 }, 1, true);

    expect(painted).toBe(true);
    expect(calls).toContainEqual(['clip', 'evenodd']);
    const draw = calls.find(([name]) => name === 'drawImage');
    expect(draw?.slice(0, 8)).toEqual(['drawImage', mocks.image, 1, 0, 1, 1, 20, 30]);
    expect(draw?.[8]).toBeCloseTo(60);
    expect(draw?.[9]).toBeCloseTo(35);
    expect(state.globalAlpha).toBe(1);
  });

  it('preserves the established full-box stretch when fill mode is absent', () => {
    const { ctx, calls } = recordingCanvas();
    expect(paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/mode-absent.png', mimeType: 'image/png',
    }, { image: mocks.image }, { x: 10, y: 20, w: 100, h: 50 }, 1)).toBe(true);
    expect(calls).toContainEqual(['drawImage', mocks.image, 10, 20, 100, 50]);
  });

  it('prepares a shared resource once, then paints each owning shape before its stroke', async () => {
    const { canvas, calls } = recordingCanvas();
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [shape('ppt/media/shared.png', 0, true), shape('ppt/media/shared.png', 1_000_000)],
    } as unknown as Slide;
    const textRuns = vi.fn();
    await renderSlide(canvas, slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['png'], { type: 'image/png' })),
    }, textRuns);

    expect(mocks.decode).toHaveBeenCalledTimes(1);
    expect(calls.filter(([name]) => name === 'drawImage')).toHaveLength(2);
    const firstImage = calls.findIndex(([name]) => name === 'drawImage');
    const firstStroke = calls.findIndex(([name]) => name === 'stroke');
    expect(firstImage).toBeGreaterThan(-1);
    expect(firstImage).toBeLessThan(firstStroke);
    expect(textRuns.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
  });

  it('contains an ordinary decode failure while retaining shape stroke and text', async () => {
    mocks.decode.mockRejectedValue(new Error('bad image'));
    const { canvas, calls } = recordingCanvas();
    const textRuns = vi.fn();
    await expect(renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null,
      elements: [shape('ppt/media/bad.png', 0, true)],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['bad'], { type: 'image/png' })),
    }, textRuns)).resolves.toBeDefined();

    expect(calls.some(([name]) => name === 'drawImage')).toBe(false);
    expect(calls.some(([name]) => name === 'stroke')).toBe(true);
    expect(textRuns.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
  });

  it.each([false, true])('keeps a shared tile native-sized regardless of consumer order (%s)', async (reversed) => {
    const shared = 'ppt/media/tile-and-stretch.png';
    const tiled = {
      ...shape(shared),
      fill: {
        fillType: 'image', imagePath: shared, mimeType: 'image/png', dpi: 300,
        tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
      },
    };
    const stretched = {
      ...shape(shared, 1_000_000),
      fill: {
        fillType: 'image', imagePath: shared, mimeType: 'image/png', stretch: true,
        srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
        fillRect: { l: 0.1, t: 0, r: 0.2, b: 0 },
      },
    };
    mocks.decode.mockClear();
    await renderSlide(recordingCanvas().canvas, {
      index: 0, slideNumber: 1, background: null,
      elements: reversed ? [stretched, tiled] : [tiled, stretched],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(300, 150) as BlobPart], { type: 'image/png' })),
    });

    expect(mocks.decode).toHaveBeenCalledTimes(1);
    expect(mocks.decode.mock.calls[0][4]).not.toHaveProperty('targetWidthPx');
    expect(mocks.decode.mock.calls[0][4]).not.toHaveProperty('targetHeightPx');
  });

  it.each([
    { sx: Number.POSITIVE_INFINITY, sy: 1, flip: 'none' },
    { sx: 1_000_000, sy: 1_000_000, flip: 'xy' },
  ])('rejects an unbounded tile auxiliary and restores caller state: %o', (tile) => {
    const { ctx, calls, state } = recordingCanvas();
    expect(paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/huge.png', mimeType: 'image/png',
      dpi: 96, tile: { algn: 'tl', tx: 0, ty: 0, ...tile },
    }, { image: mocks.image }, { x: 0, y: 0, w: 100, h: 100 }, 1)).toBe(false);
    expect(calls.some(([name]) => name === 'drawImage')).toBe(false);
    expect(state.globalAlpha).toBe(1);
  });

  it('checks the rounded tile scratch allocation at the area boundary', () => {
    const allocations: Array<[number, number]> = [];
    class ScratchCanvas {
      constructor(width: number, height: number) { allocations.push([width, height]); }
      getContext() { return recordingCanvas().ctx; }
    }
    vi.stubGlobal('OffscreenCanvas', ScratchCanvas);
    const paint = (sx: number) => {
      const { ctx } = recordingCanvas();
      ctx.createPattern = () => ({ setTransform() {} }) as unknown as CanvasPattern;
      return paintPreparedShapeImageFill(ctx, {
        fillType: 'image', imagePath: 'ppt/media/boundary.png', mimeType: 'image/png',
        dpi: 96,
        tile: { algn: 'tl', tx: 0, ty: 0, sx, sy: 4096.1, flip: 'none' },
      }, { image: mocks.image }, { x: 0, y: 0, w: 100, h: 100 }, 1 / 9525);
    };

    expect(paint(2047.05)).toBe(true); // cell 4094.1×4096.1 → alloc 4095×4097
    expect(allocations).toEqual([[4095, 4097]]);
    expect(paint(2047.55)).toBe(false); // cell 4095.1×4096.1 → alloc 4096×4097
    expect(allocations).toEqual([[4095, 4097]]);
    vi.unstubAllGlobals();
  });

  it.todo('counter-positions rotWithShape=false image frames under rotated/flipped shapes');
});
