import { describe, expect, it, vi } from 'vitest';
import viewerSource from './viewer.ts?raw';
import scrollViewerSource from './scroll-viewer.ts?raw';
import { renderPptxFocusedSlide } from './focused-view-runtime.js';
import type { PptxPresentation } from './presentation.js';

describe('PPTX focused-view composition boundary', () => {
  it('keeps static engine paint selection out of both public Viewer implementations', () => {
    for (const source of [viewerSource, scrollViewerSource]) {
      expect(source).toContain('renderPptxFocusedSlide');
      expect(source).not.toContain('.renderSlide(');
      expect(source).not.toContain('.renderSlideToBitmap(');
    }
  });

  it('routes main and worker paint through the canonical presentation operations', async () => {
    const bitmap = {} as ImageBitmap;
    const presentation = {
      renderSlide: vi.fn(() => Promise.resolve()),
      renderSlideToBitmap: vi.fn(() => Promise.resolve(bitmap)),
    } as unknown as PptxPresentation;
    const canvas = {} as HTMLCanvasElement;
    const options = { width: 960, dpr: 2 };

    await expect(renderPptxFocusedSlide(presentation, canvas, 2, 'main', options)).resolves.toBeUndefined();
    await expect(renderPptxFocusedSlide(presentation, canvas, 3, 'worker', options)).resolves.toBe(bitmap);
    expect(presentation.renderSlide).toHaveBeenCalledWith(canvas, 2, options);
    expect(presentation.renderSlideToBitmap).toHaveBeenCalledWith(3, options);
  });
});
