import { countSpacingCharacters, resolvedBidiLevels } from './character-spacing.js';
import { bidiPlacedOffsets } from './layout/visual-cluster-geometry.js';
/**
 * IX2 docx find-highlight overlay.
 *
 * The mirror image of {@link buildDocxTextLayer}: instead of a transparent
 * selection span per run, it draws a visible highlight box per matched
 * run-slice, positioned with the SAME per-run geometry (`DocxTextRunInfo.x/y/h`
 * + the run's `transform` for vertical pages) so a box lands exactly on the
 * drawn glyphs. This rides the existing DOM-overlay mechanism (a positioned
 * layer over the canvas) rather than adding a canvas draw pass — the same
 * approach the selection / hyperlink overlay uses, so highlights compose with
 * them and survive re-render the same way.
 *
 * The horizontal extent of a slice (its x offset + width within the run) is the
 * shared core `sliceHorizontalExtent`, measured against the run's font via a
 * caller-supplied `measure` primed with `ctx.font = run.font` — the same
 * measureText↔fillText correspondence the selection overlay relies on. The
 * vertical extent is the run's line box (`y` … `y + h`), owned here.
 *
 * The active match (the one `findNext`/`findPrev` last landed on) is drawn in a
 * distinct emphasis colour so the user can tell it apart from the other hits,
 * matching a browser find bar's current-vs-other highlighting.
 */
import {
  hasStrongRtl,
  sliceHorizontalExtent,
  overlayPercent,
  type FindHighlightColors,
  type MatchRunSlice,
} from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer';
import { tateChuYokoOverlayScale } from './tate-chu-yoko-overlay';
import { runOverlayMeasure } from './run-overlay-measure';

/** One page's highlight input: the run-slices a match covers, and whether that
 *  match is the active one (emphasis colour). */
export interface DocxHighlightMatch {
  slices: MatchRunSlice[];
  active: boolean;
}

/** Default highlight colours (browser find-bar palette): a soft yellow for
 *  every match and a stronger orange for the active one. Both translucent so the
 *  drawn glyphs beneath stay legible. */
export const DEFAULT_FIND_HIGHLIGHT = 'rgba(255, 214, 0, 0.42)';
export const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'rgba(255, 140, 0, 0.55)';

/** Format-specific compatibility alias for the shared colour contract. */
export type DocxHighlightColors = FindHighlightColors;

/**
 * Populate a highlight overlay layer with one box per matched run-slice.
 *
 * Every box is positioned as a PERCENTAGE of `cssWidth`/`cssHeight`, and the
 * container's own size is left untouched (`width:100%;height:100%` from the
 * caller), so the highlights track the canvas's ACTUAL rendered box even when a
 * consumer scales the canvas down with external CSS — mirroring
 * {@link buildDocxTextLayer}.
 *
 * @param layer    the overlay div (cleared here; sized `100%` by the caller).
 * @param runs     the page's runs (same array the page was rendered/text-layered from).
 * @param matches  the page's matches (run-slices + active flag).
 * @param cssWidth  the page's intended CSS width (px, number) — the x-axis % denominator.
 * @param cssHeight the page's intended CSS height (px, number) — the y-axis % denominator.
 * @param measureForFont  returns a width-measurer primed with a run's `font`
 *                        (the viewer closes over a canvas 2d context). Kept as a
 *                        factory so the font is set once per run, not per glyph.
 * @param colors   optional colour overrides.
 */
export function buildDocxHighlightLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  matches: DocxHighlightMatch[],
  cssWidth: number,
  cssHeight: number,
  measureForFont: (font: string) => (s: string) => number,
  colors: DocxHighlightColors = {},
): void {
  layer.innerHTML = '';

  const matchColor = colors.match ?? DEFAULT_FIND_HIGHLIGHT;
  const activeColor = colors.active ?? DEFAULT_FIND_ACTIVE_HIGHLIGHT;

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      // Measure the way the run was laid out and drawn (its reference-em font
      // scaled onto the run when projected with one).
      const measure = runOverlayMeasure(run, measureForFont);
      const extent = sliceHorizontalExtent(run.text, slice.start, slice.end, measure);
      // Match offsets are UTF-16; count tracking with the same rendered-
      // character rule as layout. Marks and unselected SHY add no run pitch.
      // Exclude the pitch following the slice's final rendered character.
      const pitch = run.letterSpacingPx ?? 0;
      const start = Math.max(0, Math.min(slice.start, run.text.length));
      const end = Math.max(start, Math.min(slice.end, run.text.length));
      const sliceCount = countSpacingCharacters(run.text.slice(start, end)).pitchBearing;
      const prefixCount = countSpacingCharacters(run.text.slice(0, start)).pitchBearing;
      const pitchedX = extent.x + prefixCount * pitch;
      // Keep main's multiply-before-add order for ordinary slices as well.
      const pitchedWidth = extent.width + Math.max(0, sliceCount - 1) * pitch;
      // §17.3.2.10 縦中横 alone compresses natural extents into the retained
      // one-em cell. Ordinary runs preserve the established natural measurement.
      const k = tateChuYokoOverlayScale(run, measure);
      // A run with right-to-left text is painted in the visual order of its
      // bidi levels: the slice covers one box per painted interval.
      const intervals = run.direction === 'rtl' || hasStrongRtl(run.text)
        ? visualSliceIntervals(run.text, run.direction ?? 'ltr', start, end, measure, pitch)
        : [{ x: pitchedX, width: pitchedWidth }];
      for (const interval of intervals) {
      const x = interval.x * k;
      const width = interval.width * k;
      if (width <= 0) continue;
      const box = document.createElement('div');
      // A vertical (tbRl) page reports the run at its physical top-left and
      // carries a rotate transform; apply the same transform (about the box's
      // top-left) so the highlight lies along the rotated glyph run, exactly as
      // the selection span does.
      const transform = run.transform
        ? `transform:${run.transform};transform-origin:top left;`
        : '';
      // Positioned as a % of the page's intended CSS box so the box scales with
      // the container under external CSS scaling. A vertical (tbRl) run's
      // transform is about `top left`, so the %-placed origin rotates into the
      // column exactly as before at 1× scale.
      box.style.cssText =
        `position:absolute;` +
        `left:${overlayPercent(run.x + x, cssWidth)};top:${overlayPercent(run.y, cssHeight)};` +
        `width:${overlayPercent(width, cssWidth)};height:${overlayPercent(run.h, cssHeight)};` +
        transform +
        `background:${fill};pointer-events:none;`;
      layer.appendChild(box);
      }
    }
  }
}

/** Where the characters `[start, end)` of a run are painted, from the run's
 * left edge: each character's measured advance placed in the visual order
 * of the run's UAX#9 levels (the order the run was painted in), adjacent
 * characters merged into one interval. */
function visualSliceIntervals(
  text: string,
  direction: 'ltr' | 'rtl',
  start: number,
  end: number,
  measure: (value: string) => number,
  pitch: number,
): { x: number; width: number }[] {
  const levels = resolvedBidiLevels(text, direction);
  const items: { xPt: number; advancePt: number; level: number; from: number }[] = [];
  let offset = 0;
  let pitched = 0;
  for (const character of text) {
    const extent = sliceHorizontalExtent(text, offset, offset + character.length, measure);
    const bearing = countSpacingCharacters(character).pitchBearing;
    items.push({ xPt: extent.x + pitched * pitch, advancePt: extent.width + bearing * pitch,
      level: levels[offset] ?? (direction === 'rtl' ? 1 : 0), from: offset });
    pitched += bearing;
    offset += character.length;
  }
  const placed = bidiPlacedOffsets(items);
  const boxes = items
    .map((item, index) => ({ x: placed[index]!, width: item.advancePt, from: item.from }))
    .filter((box) => box.from >= start && box.from < end && box.width > 0)
    .sort((left, right) => left.x - right.x);
  const merged: { x: number; width: number }[] = [];
  for (const box of boxes) {
    const last = merged.at(-1);
    if (last && Math.abs(last.x + last.width - box.x) < 1e-6) last.width += box.width;
    else merged.push({ x: box.x, width: box.width });
  }
  return merged;
}
