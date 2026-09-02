import { describe, it, expect, afterEach, vi } from 'vitest';
import { getCachedSvgImageByPath, dropSvgImageCache } from './svg-image-by-path';
import { getCachedBitmapByPath } from './bitmap-image-by-path';

describe('getCachedSvgImageByPath', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('fetches bytes, makes an object URL, loads an <img>, dedupes by path', async () => {
    let created = 0;
    let revoked = 0;
    vi.stubGlobal('URL', { createObjectURL: () => { created++; return `blob:${created}`; },
                           revokeObjectURL: () => { revoked++; } });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    const b = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    expect(a).toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(created).toBe(1);
    expect(revoked).toBe(1); // raw Blob handle is not retained with the image
  });

  it('self-evicts on failure (no poisoned cache) and revokes the failed object URL', async () => {
    let created = 0;
    let revoked = 0;
    vi.stubGlobal('URL', {
      createObjectURL: () => { created++; return `blob:${created}`; },
      revokeObjectURL: () => { revoked++; },
    });
    // <img> that always fails to load.
    class FailImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onerror && this.onerror()); } }
    vi.stubGlobal('Image', FailImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(revoked).toBe(1); // failed URL revoked, not leaked
    // Second call must RETRY (cache self-evicted), not return a cached rejection.
    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(fetchImage).toHaveBeenCalledTimes(2);
  });

  it('namespaces the cache by fetchImage — same zip path from two documents does not cross-contaminate', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    // Two different documents reference the SAME internal zip path. Their byte
    // sources (fetchImage closures) differ — the cache must not serve doc A's
    // decoded SVG for doc B's request.
    const fetchA = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const fetchB = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    const b = await getCachedSvgImageByPath('word/media/image1.svg', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1); // B must consult its OWN source, not hit A's cache
    expect(a).not.toBe(b); // distinct decoded images
    // Within one document, the path still dedupes.
    const a2 = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    expect(a2).toBe(a);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('dropSvgImageCache forgets decoded images and lets them re-decode', async () => {
    let revoked = 0;
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => { revoked++; } });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('a.svg', fetchImage);
    await getCachedSvgImageByPath('b.svg', fetchImage);
    expect(revoked).toBe(2); // each URL was already released after decode
    dropSvgImageCache(fetchImage); // e.g. on Document.destroy()
    expect(revoked).toBe(2);
    await getCachedSvgImageByPath('a.svg', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });

  it('awaits img.decode() before resolving when available', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    let decoded = false;
    class DecodeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      decode() { return Promise.resolve().then(() => { decoded = true; }); }
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', DecodeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('d.svg', fetchImage);
    expect(decoded).toBe(true); // resolved only after decode() completed
  });

  it('uses the shared decoded owner when HTMLImageElement is unavailable', async () => {
    vi.stubGlobal('Image', undefined);
    const bitmap = { width: 20, height: 10, close() {} } as unknown as ImageBitmap;
    const createImageBitmap = vi.fn(async () => bitmap);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    const result = await getCachedSvgImageByPath('worker.svg', fetchImage);
    expect(result).toBe(bitmap);
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('worker.svg', 'image/svg+xml');
  });

  it('uses the Window decode bridge at the requested display size in a worker', async () => {
    vi.stubGlobal('Image', undefined);
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      throw new Error('Chromium workers cannot decode this SVG Blob');
    }));
    const bitmap = { width: 640, height: 360, close() {} } as unknown as ImageBitmap;
    const workerDecoder = vi.fn(async () => bitmap);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    const result = await getCachedSvgImageByPath('worker-bridged.svg', fetchImage, {
      targetWidthPx: 640,
      targetHeightPx: 360,
      workerDecoder,
    });

    expect(result).toBe(bitmap);
    expect(workerDecoder).toHaveBeenCalledWith(expect.any(Blob), {
      targetWidthPx: 640,
      targetHeightPx: 360,
    });
  });

  it('admits at most two concurrent SVG fetch/decode operations per document', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload?.()); } }
    vi.stubGlobal('Image', FakeImg);
    let active = 0;
    let maximum = 0;
    const releases: Array<() => void> = [];
    const fetchImage = vi.fn(async () => {
      active++;
      maximum = Math.max(maximum, active);
      await new Promise<void>((resolve) => releases.push(resolve));
      active--;
      return new Blob(['<svg/>'], { type: 'image/svg+xml' });
    });

    const pending = [0, 1, 2].map((i) => getCachedSvgImageByPath(`${i}.svg`, fetchImage));
    await Promise.resolve();
    await Promise.resolve();
    expect(fetchImage).toHaveBeenCalledTimes(2);
    releases.shift()?.();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(fetchImage).toHaveBeenCalledTimes(3);
    while (releases.length > 0) releases.shift()?.();
    await Promise.all(pending);
    expect(maximum).toBe(2);
  });

  it('shares the same two decode slots across SVG and raster work', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload?.()); } }
    vi.stubGlobal('Image', FakeImg);
    vi.stubGlobal('createImageBitmap', vi.fn(async () => (
      { width: 1, height: 1, close() {} } as unknown as ImageBitmap
    )));
    let active = 0;
    let maximum = 0;
    const releases: Array<() => void> = [];
    const fetchImage = vi.fn(async (_path: string, mime: string) => {
      active++;
      maximum = Math.max(maximum, active);
      await new Promise<void>((resolve) => releases.push(resolve));
      active--;
      return new Blob([mime === 'image/svg+xml' ? '<svg/>' : new Uint8Array([0])], { type: mime });
    });

    const pending = [
      getCachedSvgImageByPath('a.svg', fetchImage),
      getCachedBitmapByPath('a.png', 'image/png', fetchImage),
      getCachedSvgImageByPath('b.svg', fetchImage),
      getCachedBitmapByPath('b.png', 'image/png', fetchImage),
    ];
    await vi.waitFor(() => expect(active).toBe(2));
    expect(fetchImage).toHaveBeenCalledTimes(2);
    while (releases.length > 0) {
      releases.shift()?.();
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await Promise.all(pending);
    expect(maximum).toBe(2);
  });
});
