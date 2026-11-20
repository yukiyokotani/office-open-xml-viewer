import { fillTextAtReference } from '@silurus/ooxml-core';
import type { CanvasPaintContext } from './types.js';

/** `fillText` with the em layout measured `font` at: a shorthand layout
 * planned at the reference em (core reference-advances.ts) is drawn there
 * under a linear scale; any other shorthand is drawn as set. The caller has
 * already assigned `font` and the run's letter spacing to `ctx`. */
export function fillRetainedText(
  context: CanvasPaintContext,
  font: string,
  text: string,
  x: number,
  y: number,
): void {
  const reference = context.textAdvances?.planned(font);
  if (reference) fillTextAtReference(context.ctx, reference, text, x, y);
  else context.ctx.fillText(text, x, y);
}
