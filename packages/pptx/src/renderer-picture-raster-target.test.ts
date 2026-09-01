import { describe, expect, it, vi } from 'vitest';
import type { Slide } from './types.js';

const coreMocks = vi.hoisted(() => ({
  bitmap: { width: 1920, height: 720, close() {} } as unknown as ImageBitmap,
  decode: vi.fn(),
}));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  coreMocks.decode.mockResolvedValue(coreMocks.bitmap);
  return {
    ...actual,
    getCachedDuotoneBitmapByPath: coreMocks.decode,
  };
});

import { renderSlide } from './renderer.js';

function canvas(): HTMLCanvasElement {
  const state: Record<string, unknown> = {
    fillStyle: '', strokeStyle: '', globalAlpha: 1, lineWidth: 1,
  };
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

describe('PPTX display-sized picture decode', () => {
  it('derives full-source device pixels from the frame, effective DPR, and srcRect crop', async () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'picture',
        x: 0,
        y: 0,
        width: 4_572_000, // half of the 10-inch slide => 480 CSS px
        height: 3_429_000, // half slide height => 360 CSS px
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: 'ppt/media/108mp-poster.jpg',
        mimeType: 'image/jpeg',
        srcRect: { l: 0.25, t: 0, r: 0.25, b: 0 },
      }],
    } as Slide;
    const fetchImage = vi.fn(async () => new Blob(['jpeg'], { type: 'image/jpeg' }));

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
    });

    expect(coreMocks.decode).toHaveBeenCalled();
    for (const call of coreMocks.decode.mock.calls) {
      expect(call[4]).toMatchObject({
        // 480 CSS px × DPR 2, then / 50% visible source width.
        targetWidthPx: 1920,
        targetHeightPx: 720,
      });
    }
  });
});
