import { describe, expect, it, vi } from 'vitest';
import {
  WorkerSvgDecodeClient,
  boundedSvgRasterSize,
  decodeSvgBlobOnMainThread,
  type WorkerSvgDecodeRequest,
} from './svg-decode-bridge';

describe('WorkerSvgDecodeClient', () => {
  it('correlates a transferred SVG decode without losing the display target', async () => {
    const sent: WorkerSvgDecodeRequest[] = [];
    const client = new WorkerSvgDecodeClient((message) => sent.push(message as WorkerSvgDecodeRequest));
    const pending = client.decode(new Blob(['<svg/>']), {
      targetWidthPx: 512,
      targetHeightPx: 256,
    });
    await vi.waitFor(() => expect(sent).toHaveLength(1));
    expect(sent[0]).toMatchObject({
      kind: 'ooxmlDecodeSvg',
      decodeId: 1,
      targetWidthPx: 512,
      targetHeightPx: 256,
    });
    const bitmap = { width: 512, height: 256, close: vi.fn() } as unknown as ImageBitmap;
    expect(client.accept({ kind: 'ooxmlSvgDecoded', decodeId: 1, bitmap })).toBe(true);
    await expect(pending).resolves.toBe(bitmap);
  });

  it('bounds extreme vector targets by axis and aggregate pixel budgets', () => {
    const size = boundedSvgRasterSize(500_000, 500_000);
    expect(size.width).toBe(size.height);
    expect(size.width * size.height).toBeLessThanOrEqual(1 << 25);
    expect(size.width).toBeLessThanOrEqual(32_767);
  });

  it('keeps bounded dimensions finite when a finite target product overflows', () => {
    const size = boundedSvgRasterSize(Number.MAX_VALUE, Number.MAX_VALUE);
    expect(Number.isSafeInteger(size.width)).toBe(true);
    expect(Number.isSafeInteger(size.height)).toBe(true);
    expect(size.width * size.height).toBeLessThanOrEqual(1 << 25);
  });
});

describe('decodeSvgBlobOnMainThread', () => {
  it('rasterizes a viewBox-only SVG into the requested bounded surface before transfer', async () => {
    const drawImage = vi.fn();
    const bitmap = { width: 400, height: 240, close: vi.fn() } as unknown as ImageBitmap;
    const revokeObjectURL = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:test-svg');
    vi.stubGlobal('Image', class {
      width = 0;
      height = 0;
      // Chromium assigns a viewBox-only SVG a 150px-tall natural box whose
      // width preserves the viewBox ratio (10:6 => 250:150).
      naturalWidth = 250;
      naturalHeight = 150;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      decode = vi.fn(async () => undefined);
      set src(_value: string) { queueMicrotask(() => this.onload?.()); }
    });
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return { drawImage }; }
      transferToImageBitmap() { return bitmap; }
    });

    try {
      await expect(decodeSvgBlobOnMainThread(
        new Blob(['<svg viewBox="0 0 10 6"/>'], { type: 'image/svg+xml' }),
        { targetWidthPx: 400, targetHeightPx: 240 },
      )).resolves.toBe(bitmap);
      expect(drawImage).toHaveBeenCalledWith(expect.anything(), 0, 0, 400, 240);
      expect(revokeObjectURL).toHaveBeenCalledWith('blob:test-svg');
    } finally {
      vi.restoreAllMocks();
      vi.unstubAllGlobals();
    }
  });

  it('covers a one-axis target while preserving the SVG intrinsic aspect ratio', async () => {
    const drawImage = vi.fn();
    const bitmap = { width: 80, height: 40, close: vi.fn() } as unknown as ImageBitmap;
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:one-axis-svg');
    vi.stubGlobal('Image', class {
      width = 0;
      height = 0;
      naturalWidth = 300;
      naturalHeight = 150;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      decode = vi.fn(async () => undefined);
      set src(_value: string) { queueMicrotask(() => this.onload?.()); }
    });
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return { drawImage }; }
      transferToImageBitmap() { return bitmap; }
    });

    try {
      await expect(decodeSvgBlobOnMainThread(
        new Blob(['<svg viewBox="0 0 16 8"/>'], { type: 'image/svg+xml' }),
        { targetHeightPx: 40 },
      )).resolves.toBe(bitmap);
      expect(drawImage).toHaveBeenCalledWith(expect.anything(), 0, 0, 80, 40);
    } finally {
      vi.restoreAllMocks();
      vi.unstubAllGlobals();
    }
  });

  it('covers mismatched target axes without distorting the SVG source grid', async () => {
    const drawImage = vi.fn();
    const bitmap = { width: 480, height: 240, close: vi.fn() } as unknown as ImageBitmap;
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:coverage-svg');
    vi.stubGlobal('Image', class {
      width = 0;
      height = 0;
      naturalWidth = 300;
      naturalHeight = 150;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      decode = vi.fn(async () => undefined);
      set src(_value: string) { queueMicrotask(() => this.onload?.()); }
    });
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return { drawImage }; }
      transferToImageBitmap() { return bitmap; }
    });

    try {
      await expect(decodeSvgBlobOnMainThread(
        new Blob(['<svg viewBox="0 0 16 8"/>'], { type: 'image/svg+xml' }),
        { targetWidthPx: 400, targetHeightPx: 240 },
      )).resolves.toBe(bitmap);
      expect(drawImage).toHaveBeenCalledWith(expect.anything(), 0, 0, 480, 240);
    } finally {
      vi.restoreAllMocks();
      vi.unstubAllGlobals();
    }
  });
});
