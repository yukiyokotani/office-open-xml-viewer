import type { RenderPageToBitmapOptions } from './document.js';
import type { DocxDocument } from './document.js';

/**
 * Format-local paint boundary shared by the focused Canvas Viewer and every
 * virtualized ScrollViewer slot. The caller owns render generations, bitmap
 * commit, overlays, and slot lifetime; this boundary only selects the engine's
 * canonical main/worker paint operation.
 */
export function renderDocxFocusedPage(
  document: DocxDocument,
  canvas: HTMLCanvasElement,
  pageIndex: number,
  mode: 'main',
  options: RenderPageToBitmapOptions,
): Promise<void>;
export function renderDocxFocusedPage(
  document: DocxDocument,
  canvas: HTMLCanvasElement,
  pageIndex: number,
  mode: 'worker',
  options: RenderPageToBitmapOptions,
): Promise<ImageBitmap>;
export function renderDocxFocusedPage(
  document: DocxDocument,
  canvas: HTMLCanvasElement,
  pageIndex: number,
  mode: 'main' | 'worker',
  options: RenderPageToBitmapOptions,
): Promise<ImageBitmap | void> {
  if (mode === 'worker') return document.renderPageToBitmap(pageIndex, options);
  return document.renderPage(canvas, pageIndex, options);
}
