import { describe, expect, it, vi } from 'vitest';
import type { Slide } from './types.js';

const bitmapDecode = vi.hoisted(() => {
  let resolve!: (bitmap: ImageBitmap) => void;
  const promise = new Promise<ImageBitmap>((accept) => { resolve = accept; });
  return { promise, resolve, decode: vi.fn(() => promise) };
});

vi.mock('@silurus/ooxml-core', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  getCachedDuotoneBitmapByPath: bitmapDecode.decode,
}));

import { invalidatePptxRenderTarget, renderSlide } from './renderer.js';

describe('PPTX main-render target invalidation', () => {
  it('does not draw a decoded background after its caller canvas was restored', async () => {
    const calls: string[] = [];
    const context = new Proxy({
      fillStyle: '',
      globalAlpha: 1,
    } as Record<string, unknown>, {
      get(target, property: string) {
        if (property in target) return target[property];
        return () => { calls.push(property); };
      },
      set(target, property: string, value) {
        target[property] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
    const canvas = {
      width: 0,
      height: 0,
      style: {} as CSSStyleDeclaration,
      offsetWidth: 960,
      getContext: () => context,
    } as unknown as HTMLCanvasElement;
    const slide: Slide = {
      index: 0,
      slideNumber: 1,
      background: {
        fillType: 'image',
        imagePath: 'ppt/media/background.png',
        mimeType: 'image/png',
      },
      elements: [],
    };

    const rendering = renderSlide(canvas, slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage: vi.fn(async () => new Blob()),
    });
    await vi.waitFor(() => expect(bitmapDecode.decode).toHaveBeenCalledOnce());

    invalidatePptxRenderTarget(canvas);
    calls.length = 0; // models CallerCanvasMount restoring the original bitmap
    bitmapDecode.resolve({ width: 1, height: 1, close: vi.fn() } as unknown as ImageBitmap);
    await rendering;

    expect(calls).not.toContain('drawImage');
    expect(calls).toHaveLength(0);
  });
});
