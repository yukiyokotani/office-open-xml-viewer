import { fillTextAtReference } from '@silurus/ooxml-core';
import { restoreCanvasLanguage, selectCanvasLanguage } from '@silurus/ooxml-core/internal/canvas-language';
import type { TextPlacement } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

/** `fillText` with the em layout measured `font` at: a shorthand layout
 * planned at the reference em (core reference-advances.ts) is drawn there
 * under a linear scale; any other shorthand is drawn as set. The caller has
 * already assigned `font` and the run's letter spacing to `ctx`. `lang` is
 * the route language the caller selected (none for untagged text), kept
 * across the reference-em `save()`. */
export function fillRetainedText(
  context: CanvasPaintContext,
  font: string,
  text: string,
  x: number,
  y: number,
  lang?: string,
): void {
  const reference = context.textAdvances?.planned(font);
  if (reference) fillTextAtReference(context.ctx, reference, text, x, y, lang);
  else context.ctx.fillText(text, x, y);
}

/** {@link fillRetainedText} of one string painted through `route`, in the
 * route's language (the state layout measured it in), restoring the
 * context's own. `font` is the route's shorthand, already assigned. */
export function fillRetainedRouteText(
  context: CanvasPaintContext,
  route: TextPlacement['fontRoute'],
  font: string,
  text: string,
  x: number,
  y: number,
): void {
  const previousLang = selectCanvasLanguage(context.ctx, route.lang);
  fillRetainedText(context, font, text, x, y, route.lang);
  restoreCanvasLanguage(context.ctx, previousLang);
}
