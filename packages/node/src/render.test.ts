import { describe, it, expect, afterEach, vi } from 'vitest';
import { installOffscreenCanvasShim, installImageBitmapShim } from './render';
import type { NodeCanvasFactory } from './render';
import { loadSkiaForTests } from './test-imports';

// skia-canvas ships a native binding via a devDependency, so `pnpm install`
// provides it in CI as well as locally — these canvas-backed checks run
// everywhere. Load it through the shared test helper: absent → skip cleanly
// (local), but under OOXML_REQUIRE_SKIA=1 (CI) a load failure becomes a hard
// error instead of a silent skip. `describe.skipIf(!skia)` gates every block.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (buf) => loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
};

const g = globalThis as unknown as { OffscreenCanvas?: unknown };

describe('installImageBitmapShim display-sized decode', () => {
  it('forwards the target to the codec and resizes a native-sized fallback surface', async () => {
    const drawImage = vi.fn();
    const resizedSurface = {
      width: 1280,
      height: 960,
      getContext: () => ({ measureText: () => ({ width: 0 }), drawImage }),
    };
    const sourceImage = { width: 12_090, height: 9_063 };
    const loadImage = vi.fn(async () => sourceImage);
    const fallbackFactory: NodeCanvasFactory = {
      createCanvas: vi.fn(() => resizedSurface),
      loadImage,
    };
    const restore = installImageBitmapShim(fallbackFactory);
    try {
      const bitmap = await createImageBitmap(new Blob(['encoded']), {
        resizeWidth: 1280,
        resizeHeight: 960,
        resizeQuality: 'high',
      });
      expect(loadImage).toHaveBeenCalledWith(expect.any(ArrayBuffer), { width: 1280 });
      expect(drawImage).toHaveBeenCalledWith(sourceImage, 0, 0, 1280, 960);
      expect(bitmap).toBe(resizedSurface);
    } finally {
      restore();
    }
  });
});

describe.skipIf(!skia)('installOffscreenCanvasShim', () => {
  afterEach(() => {
    // Make sure each test cleans up; restore is the responsibility of the test
    // body, but guard against leakage if an assertion threw.
    delete g.OffscreenCanvas;
  });

  it('injects a global OffscreenCanvas backed by the factory', () => {
    expect(typeof g.OffscreenCanvas).toBe('undefined');
    const restore = installOffscreenCanvasShim(factory);
    expect(typeof g.OffscreenCanvas).toBe('function');

    const oc = new (g.OffscreenCanvas as new (w: number, h: number) => unknown)(64, 48) as {
      width: number;
      height: number;
      getContext(k: '2d'): CanvasRenderingContext2D;
    };
    expect(oc.width).toBe(64);
    expect(oc.height).toBe(48);
    const ctx = oc.getContext('2d');
    expect(typeof ctx.getImageData).toBe('function');
    expect(typeof ctx.putImageData).toBe('function');
    expect(typeof ctx.drawImage).toBe('function');

    restore();
    expect(typeof g.OffscreenCanvas).toBe('undefined');
  });

  it('does not overwrite an already-defined global OffscreenCanvas', () => {
    const sentinel = function PreExisting() {} as unknown;
    g.OffscreenCanvas = sentinel;
    const restore = installOffscreenCanvasShim(factory);
    expect(g.OffscreenCanvas).toBe(sentinel);
    restore();
    // restore returns the global to its pre-call value (the sentinel).
    expect(g.OffscreenCanvas).toBe(sentinel);
    delete g.OffscreenCanvas;
  });

  it('shimmed OffscreenCanvas supports a ctx.filter blur pass (feathered edge)', () => {
    const restore = installOffscreenCanvasShim(factory);
    const oc = new (g.OffscreenCanvas as new (w: number, h: number) => unknown)(100, 100) as {
      getContext(k: '2d'): CanvasRenderingContext2D;
    };
    const ctx = oc.getContext('2d');
    ctx.filter = 'blur(6px)';
    ctx.fillStyle = '#000';
    ctx.fillRect(40, 40, 20, 20);
    ctx.filter = 'none';
    // Just outside the original square edge, blur must leak partial alpha.
    const outside = ctx.getImageData(33, 50, 1, 1).data;
    expect(outside[3]).toBeGreaterThan(0);
    expect(outside[3]).toBeLessThan(255);
    restore();
    delete g.OffscreenCanvas;
  });

  it('shimmed OffscreenCanvas works as a drawImage source onto another canvas', () => {
    const restore = installOffscreenCanvasShim(factory);
    const oc = new (g.OffscreenCanvas as new (w: number, h: number) => unknown)(50, 50) as {
      getContext(k: '2d'): CanvasRenderingContext2D;
    };
    const octx = oc.getContext('2d');
    octx.fillStyle = '#ff0000';
    octx.fillRect(0, 0, 50, 50);

    const dest = factory.createCanvas(50, 50);
    const dctx = dest.getContext('2d') as unknown as CanvasRenderingContext2D;
    dctx.drawImage(oc as unknown as CanvasImageSource, 0, 0);
    const px = dctx.getImageData(25, 25, 1, 1).data;
    expect(Array.from(px)).toEqual([255, 0, 0, 255]);
    restore();
    delete g.OffscreenCanvas;
  });
});

describe.skipIf(!skia)('installImageBitmapShim', () => {
  it('restores the previous global', () => {
    const gg = globalThis as unknown as { createImageBitmap?: unknown };
    const before = gg.createImageBitmap;
    const restore = installImageBitmapShim(factory);
    expect(typeof gg.createImageBitmap).toBe('function');
    restore();
    expect(gg.createImageBitmap).toBe(before);
  });
});
