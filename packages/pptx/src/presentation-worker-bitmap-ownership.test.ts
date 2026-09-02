import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import type { PptxTextRunInfo } from './renderer.js';

function ownedBitmap(close = vi.fn()) {
  const bitmap = { width: 4, height: 3, close } as unknown as ImageBitmap;
  return { bitmap, close };
}

function workerPresentation(bitmap: ImageBitmap, runs: PptxTextRunInfo[]): PptxPresentation {
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  const slideWidth = 9_144_000;
  const slideHeight = 6_858_000;
  Object.assign(instance, {
    _mode: 'worker',
    _resourceFailure: null,
    _bootstrap: { slideCount: 1, slideWidth, slideHeight },
    _preflight: {
      slideCount: 1,
      slideWidth,
      slideHeight,
      slides: [{ mediaElements: [] }],
    },
    _availableSlideCount: 1,
    _destroyed: false,
    _fetchMedia: vi.fn(async () => new Blob()),
    _fetchImage: vi.fn(async () => new Blob()),
    _bridge: {
      request: vi.fn(async () => ({ kind: 'slideRendered', id: 1, bitmap, runs })),
    },
  });
  return instance as unknown as PptxPresentation;
}

function canvasWithContexts(
  ...contexts: Array<CanvasRenderingContext2D | null>
): HTMLCanvasElement {
  const getContext = vi.fn();
  for (const context of contexts) getContext.mockReturnValueOnce(context);
  return {
    width: 0,
    height: 0,
    offsetWidth: 320,
    style: {},
    getContext,
  } as unknown as HTMLCanvasElement;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('PptxPresentation worker bitmap ownership', () => {
  it('releases the received bitmap once and preserves the callback failure', async () => {
    const { bitmap, close } = ownedBitmap();
    const runs = [
      { text: 'first' },
      { text: 'second' },
    ] as PptxTextRunInfo[];
    const failure = new Error('text-run callback failed');
    const onTextRun = vi.fn((run: PptxTextRunInfo) => {
      if (run === runs[1]) throw failure;
    });
    const presentation = workerPresentation(bitmap, runs);

    await expect(presentation.renderSlideToBitmap(0, {
      width: 320,
      dpr: 1,
      onTextRun,
    })).rejects.toBe(failure);

    expect(onTextRun).toHaveBeenCalledTimes(2);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('hands the still-open bitmap to the caller after successful callback replay', async () => {
    const { bitmap, close } = ownedBitmap();
    const runs = [{ text: 'success' }] as PptxTextRunInfo[];
    const onTextRun = vi.fn();
    const presentation = workerPresentation(bitmap, runs);

    await expect(presentation.renderSlideToBitmap(0, {
      width: 320,
      dpr: 1,
      onTextRun,
    })).resolves.toBe(bitmap);

    expect(onTextRun).toHaveBeenCalledWith(runs[0]);
    expect(close).not.toHaveBeenCalled();
  });

  it('releases the bitmap when presentSlide cannot acquire its draw context', async () => {
    const { bitmap, close } = ownedBitmap();
    const presentation = workerPresentation(bitmap, []);
    const canvas = canvasWithContexts(
      {} as CanvasRenderingContext2D,
      null,
    );

    await expect(presentation.presentSlide(canvas, 0, {
      width: 320,
      dpr: 1,
    })).rejects.toThrow('2D context not available');

    expect(close).toHaveBeenCalledTimes(1);
  });

  it('releases the bitmap and preserves a presentSlide draw failure', async () => {
    const { bitmap, close } = ownedBitmap();
    const failure = new Error('draw failed');
    const drawImage = vi.fn(() => { throw failure; });
    const presentation = workerPresentation(bitmap, []);
    const canvas = canvasWithContexts(
      {} as CanvasRenderingContext2D,
      { drawImage } as unknown as CanvasRenderingContext2D,
    );

    await expect(presentation.presentSlide(canvas, 0, {
      width: 320,
      dpr: 1,
    })).rejects.toBe(failure);

    expect(drawImage).toHaveBeenCalledWith(bitmap, 0, 0);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('preserves a presentSlide draw failure when bitmap cleanup also throws', async () => {
    const cleanupFailure = new Error('bitmap cleanup failed');
    const close = vi.fn(() => { throw cleanupFailure; });
    const { bitmap } = ownedBitmap(close);
    const drawFailure = new Error('draw failed');
    const presentation = workerPresentation(bitmap, []);
    const canvas = canvasWithContexts(
      {} as CanvasRenderingContext2D,
      { drawImage: () => { throw drawFailure; } } as unknown as CanvasRenderingContext2D,
    );

    await expect(presentation.presentSlide(canvas, 0, {
      width: 320,
      dpr: 1,
    })).rejects.toBe(drawFailure);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('releases the bitmap exactly once after a successful presentSlide draw', async () => {
    const { bitmap, close } = ownedBitmap();
    const drawImage = vi.fn();
    const baseDrawImage = vi.fn();
    const presentation = workerPresentation(bitmap, []);
    const canvas = canvasWithContexts(
      {} as CanvasRenderingContext2D,
      { drawImage } as unknown as CanvasRenderingContext2D,
    );
    const baseCanvas = {
      width: 0,
      height: 0,
      getContext: vi.fn(() => ({ drawImage: baseDrawImage })),
    } as unknown as HTMLCanvasElement;
    vi.stubGlobal('document', {
      createElement: vi.fn(() => baseCanvas),
    });

    await expect(presentation.presentSlide(canvas, 0, {
      width: 320,
      dpr: 1,
    })).resolves.toBeDefined();

    expect(drawImage).toHaveBeenCalledWith(bitmap, 0, 0);
    expect(baseDrawImage).toHaveBeenCalledWith(canvas, 0, 0);
    expect(close).toHaveBeenCalledTimes(1);
  });
});
