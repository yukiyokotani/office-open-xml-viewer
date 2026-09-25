import type { ImageFill } from '../types/common.js';

export type ActiveChartImageFillPainter = (
  ctx: CanvasRenderingContext2D,
  fill: ImageFill,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx: number,
  shapeRotationDeg: number,
) => boolean;

let activePainter: ActiveChartImageFillPainter | undefined;

/** Keep leaf frame painters independent from image collection and decoding.
 * The host-owned chart scope installs the already-bounded synchronous painter. */
export function withActiveChartImageFillPainter<T>(
  painter: ActiveChartImageFillPainter | undefined,
  paint: () => T,
): T {
  const previous = activePainter;
  activePainter = painter;
  try {
    return paint();
  } finally {
    activePainter = previous;
  }
}

export function paintActiveChartImageFill(
  ctx: CanvasRenderingContext2D,
  fill: ImageFill,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  return activePainter?.(ctx, fill, x, y, w, h, ptToPx, shapeRotationDeg) ?? false;
}
