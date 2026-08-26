import type {
  PptxPresentation,
  RenderSlideToBitmapOptions,
} from './presentation.js';

/**
 * Format-local static paint boundary shared by the focused Canvas Viewer and
 * every virtualized ScrollViewer slot. Presentation handles remain caller-owned
 * because their media lifetime differs between a focused view and a slot pool.
 */
export function renderPptxFocusedSlide(
  presentation: PptxPresentation,
  canvas: HTMLCanvasElement,
  slideIndex: number,
  mode: 'main',
  options: RenderSlideToBitmapOptions,
): Promise<void>;
export function renderPptxFocusedSlide(
  presentation: PptxPresentation,
  canvas: HTMLCanvasElement,
  slideIndex: number,
  mode: 'worker',
  options: RenderSlideToBitmapOptions,
): Promise<ImageBitmap>;
export function renderPptxFocusedSlide(
  presentation: PptxPresentation,
  canvas: HTMLCanvasElement,
  slideIndex: number,
  mode: 'main' | 'worker',
  options: RenderSlideToBitmapOptions,
): Promise<ImageBitmap | void> {
  if (mode === 'worker') return presentation.renderSlideToBitmap(slideIndex, options);
  return presentation.renderSlide(canvas, slideIndex, options);
}
