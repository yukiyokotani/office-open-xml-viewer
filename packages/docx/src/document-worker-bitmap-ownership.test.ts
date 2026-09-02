import { describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';
import type { DocxTextRunInfo } from './renderer.js';

function ownedBitmap(close = vi.fn()) {
  const bitmap = { width: 4, height: 3, close } as unknown as ImageBitmap;
  return { bitmap, close };
}

function workerDocument(bitmap: ImageBitmap, runs: DocxTextRunInfo[]): DocxDocument {
  const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
  Object.assign(instance, {
    _mode: 'worker',
    _bridge: {
      request: vi.fn(async () => ({ type: 'pageRendered', id: 1, bitmap, runs })),
    },
  });
  const document = instance as unknown as DocxDocument;
  attachDocumentLayoutRuntime(document, 0);
  return document;
}

describe('DocxDocument worker bitmap callback ownership', () => {
  it('releases the received bitmap once and preserves the callback failure', async () => {
    const { bitmap, close } = ownedBitmap();
    const runs = [
      { text: 'first' },
      { text: 'second' },
    ] as DocxTextRunInfo[];
    const failure = new Error('text-run callback failed');
    const onTextRun = vi.fn((run: DocxTextRunInfo) => {
      if (run === runs[1]) throw failure;
    });
    const document = workerDocument(bitmap, runs);

    await expect(document.renderPageToBitmap(0, { dpr: 1, onTextRun })).rejects.toBe(failure);

    expect(onTextRun).toHaveBeenCalledTimes(2);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('hands the still-open bitmap to the caller after successful callback replay', async () => {
    const { bitmap, close } = ownedBitmap();
    const runs = [{ text: 'success' }] as DocxTextRunInfo[];
    const onTextRun = vi.fn();
    const document = workerDocument(bitmap, runs);

    await expect(document.renderPageToBitmap(0, { dpr: 1, onTextRun })).resolves.toBe(bitmap);

    expect(onTextRun).toHaveBeenCalledWith(runs[0]);
    expect(close).not.toHaveBeenCalled();
  });

  it('preserves the callback failure when bitmap cleanup also throws', async () => {
    const cleanupFailure = new Error('bitmap cleanup failed');
    const close = vi.fn(() => { throw cleanupFailure; });
    const { bitmap } = ownedBitmap(close);
    const callbackFailure = new Error('text-run callback failed');
    const document = workerDocument(bitmap, [{ text: 'run' }] as DocxTextRunInfo[]);

    await expect(document.renderPageToBitmap(0, {
      dpr: 1,
      onTextRun: () => { throw callbackFailure; },
    })).rejects.toBe(callbackFailure);
    expect(close).toHaveBeenCalledTimes(1);
  });
});
