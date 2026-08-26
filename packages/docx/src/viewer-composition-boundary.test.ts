import { describe, expect, it, vi } from 'vitest';
import viewerSource from './viewer.ts?raw';
import scrollViewerSource from './scroll-viewer.ts?raw';
import { renderDocxFocusedPage } from './focused-view-runtime.js';
import type { DocxDocument } from './document.js';

describe('DOCX focused-view composition boundary', () => {
  it('keeps engine paint selection out of both public Viewer implementations', () => {
    for (const source of [viewerSource, scrollViewerSource]) {
      expect(source).toContain('renderDocxFocusedPage');
      expect(source).not.toContain('.renderPage(');
      expect(source).not.toContain('.renderPageToBitmap(');
    }
  });

  it('routes main and worker paint through the canonical document operations', async () => {
    const bitmap = {} as ImageBitmap;
    const document = {
      renderPage: vi.fn(() => Promise.resolve()),
      renderPageToBitmap: vi.fn(() => Promise.resolve(bitmap)),
    } as unknown as DocxDocument;
    const canvas = {} as HTMLCanvasElement;
    const options = { width: 640, dpr: 2 };

    await expect(renderDocxFocusedPage(document, canvas, 3, 'main', options)).resolves.toBeUndefined();
    await expect(renderDocxFocusedPage(document, canvas, 4, 'worker', options)).resolves.toBe(bitmap);
    expect(document.renderPage).toHaveBeenCalledWith(canvas, 3, options);
    expect(document.renderPageToBitmap).toHaveBeenCalledWith(4, options);
  });
});
