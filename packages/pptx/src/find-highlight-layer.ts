/**
 * IX2 pptx find-highlight overlay.
 *
 * The highlight twin of {@link buildPptxTextLayer}: it draws a visible box per
 * matched run-slice, grouped into one positioned + rotated `<div>` per shape
 * frame exactly as the selection overlay groups its transparent spans, so a box
 * tracks the drawn (rotated) text. Riding the same shape-grouped DOM overlay
 * (rather than a canvas draw pass) means highlights rotate with their shape and
 * compose with the selection / hyperlink layer.
 *
 * A slice's horizontal extent within its run is the shared core
 * `sliceHorizontalExtent`, measured against the run's font; the box's vertical
 * extent is the run's line box (`h`). Boxes are placed in the shape's own
 * coordinate frame (`inShapeX`/`inShapeY`), so the shape div's `rotate()` lays
 * them along the glyphs. The active match uses a distinct emphasis colour.
 */
import {
  sliceHorizontalExtent,
  overlayPercent,
  type FindHighlightColors,
  type MatchRunSlice,
} from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';
import { pptxRunFrameKey, pptxRunFrameTransform } from './run-frame-transform.js';

export interface PptxHighlightMatch {
  slices: MatchRunSlice[];
  active: boolean;
}

/** Browser find-bar palette (translucent so glyphs stay legible). */
export const DEFAULT_FIND_HIGHLIGHT = 'rgba(255, 214, 0, 0.42)';
export const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'rgba(255, 140, 0, 0.55)';

/** Format-specific compatibility alias for the shared colour contract. */
export type PptxHighlightColors = FindHighlightColors;

/**
 * Populate a highlight overlay layer with a box per matched run-slice, grouped
 * by shape frame (with the shape's rotation) so each box lands on the drawn
 * glyphs.
 *
 * All coordinates are PERCENTAGES of `cssWidth`/`cssHeight`, and the container's
 * own size is left untouched (`width:100%;height:100%` from the caller), so the
 * highlights track the canvas's ACTUAL rendered box even when a consumer scales
 * the canvas down with external CSS — mirroring {@link buildPptxTextLayer}.
 *
 * @param layer     the overlay div (cleared here; sized `100%` by the caller).
 * @param runs      the slide's runs (same array the slide was rendered from).
 * @param matches   the slide's matches (run-slices + active flag).
 * @param cssWidth  the slide's intended CSS width (px, number) — the % denominator.
 * @param cssHeight the slide's intended CSS height (px, number) — the % denominator.
 * @param measureForFont returns a width-measurer primed with a run's font.
 * @param colors    optional colour overrides.
 */
export function buildPptxHighlightLayer(
  layer: HTMLDivElement,
  runs: PptxTextRunInfo[],
  matches: PptxHighlightMatch[],
  cssWidth: number,
  cssHeight: number,
  measureForFont: (font: string) => (s: string) => number,
  colors: PptxHighlightColors = {},
): void {
  layer.innerHTML = '';

  const matchColor = colors.match ?? DEFAULT_FIND_HIGHLIGHT;
  const activeColor = colors.active ?? DEFAULT_FIND_ACTIVE_HIGHLIGHT;

  // One positioned + rotated div per shape frame (keyed like the text layer), so
  // boxes inside it inherit the shape's rotation. Reused across matches/slices.
  // The frame is placed as a % of the slide box (so it tracks the scaled canvas)
  // with % width/height too, so the boxes inside (positioned as % of THIS frame)
  // scale with it under the shape's rotate().
  const shapeMap = new Map<string, { div: HTMLDivElement; w: number; h: number }>();
  const shapeDiv = (run: PptxTextRunInfo): { div: HTMLDivElement; w: number; h: number } => {
    const transform = pptxRunFrameTransform(run);
    const key = pptxRunFrameKey(run, transform);
    let entry = shapeMap.get(key);
    if (!entry) {
      const div = document.createElement('div');
      div.style.cssText =
        `position:absolute;` +
        `left:${overlayPercent(run.shapeX, cssWidth)};top:${overlayPercent(run.shapeY, cssHeight)};` +
        `width:${overlayPercent(run.shapeW, cssWidth)};height:${overlayPercent(run.shapeH, cssHeight)};` +
        `pointer-events:none;overflow:hidden;`;
      if (transform) {
        div.style.transformOrigin = 'center center';
        div.style.transform = transform;
      }
      entry = { div, w: run.shapeW, h: run.shapeH };
      shapeMap.set(key, entry);
      layer.appendChild(div);
    }
    return entry;
  };

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const measure = measureForFont(run.font);
      const { x, width } = sliceHorizontalExtent(run.text, slice.start, slice.end, measure);
      if (width <= 0) continue;
      const shape = shapeDiv(run);
      const box = document.createElement('div');
      // Placed as a % of the shape frame (shapeW/shapeH), so the box scales with
      // the frame when the whole overlay is scaled down by external CSS.
      box.style.cssText =
        `position:absolute;` +
        `left:${overlayPercent(run.inShapeX + x, shape.w)};top:${overlayPercent(run.inShapeY, shape.h)};` +
        `width:${overlayPercent(width, shape.w)};height:${overlayPercent(run.h, shape.h)};` +
        `background:${fill};pointer-events:none;`;
      shape.div.appendChild(box);
    }
  }
}
