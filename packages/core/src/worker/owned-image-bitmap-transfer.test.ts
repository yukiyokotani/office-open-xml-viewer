import { readFile } from 'node:fs/promises';
import { describe, expect, it, vi } from 'vitest';
import { postOwnedImageBitmap } from './owned-image-bitmap-transfer.js';

function bitmapWithClose(close = vi.fn()): ImageBitmap {
  return { width: 4, height: 3, close } as unknown as ImageBitmap;
}

describe('postOwnedImageBitmap', () => {
  it('transfers the bitmap without closing ownership after a successful post', () => {
    const close = vi.fn();
    const bitmap = bitmapWithClose(close);
    const message = { kind: 'rendered', bitmap } as const;
    const post = vi.fn();

    postOwnedImageBitmap(post, message);

    expect(post).toHaveBeenCalledWith(message, [bitmap]);
    expect(close).not.toHaveBeenCalled();
  });

  it('closes the still-owned bitmap and rethrows the original post failure', () => {
    const close = vi.fn();
    const bitmap = bitmapWithClose(close);
    const failure = new DOMException('worker already terminated', 'InvalidStateError');
    const post = vi.fn(() => { throw failure; });

    expect(() => postOwnedImageBitmap(post, { bitmap })).toThrow(failure);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('preserves the post failure when bitmap cleanup also throws', () => {
    const cleanupFailure = new Error('bitmap cleanup failed');
    const bitmap = bitmapWithClose(vi.fn(() => { throw cleanupFailure; }));
    const postFailure = new DOMException('worker already terminated', 'InvalidStateError');
    const post = vi.fn(() => { throw postFailure; });

    expect(() => postOwnedImageBitmap(post, { bitmap })).toThrow(postFailure);
  });
});

describe('render-worker bitmap transfer ownership', () => {
  it.each([
    [
      'DOCX page',
      new URL('../../../docx/src/render-worker.ts', import.meta.url),
      /postOwnedImageBitmap\(\s*post,\s*\{\s*type:\s*'pageRendered'/s,
    ],
    [
      'PPTX slide',
      new URL('../../../pptx/src/render-worker.ts', import.meta.url),
      /postOwnedImageBitmap\(\s*post,\s*\{\s*kind:\s*'slideRendered'/s,
    ],
    [
      'XLSX viewport',
      new URL('../../../xlsx/src/render-worker.ts', import.meta.url),
      /postOwnedImageBitmap\(\s*post,\s*\{\s*type:\s*'viewportRendered'/s,
    ],
  ] as const)('routes the %s response through the transactional helper', async (
    _label,
    sourceUrl,
    expectedCall,
  ) => {
    const source = await readFile(sourceUrl, 'utf8');
    expect(source).toMatch(expectedCall);
  });
});
