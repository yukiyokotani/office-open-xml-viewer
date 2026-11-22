import { countSpacingCharacters } from './character-spacing.js';
import { OverlayClipFrames } from './run-clip.js';
import { clipRunRect, runFontBox, runOverlayRects, type RunRect } from './overlay-geometry.js';
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
 * A slice's boxes come from the overlay geometry pipeline
 * (overlay-geometry.ts): the glyph rects of the slice's characters, cut to
 * the run's clips and joined where they touch. Only a run without cluster
 * boxes (one not projected from a layout) is measured, with the shared core
 * `sliceHorizontalExtent` against the run's font via a caller-supplied
 * `measure` primed with `ctx.font = run.font`.
 *
 * The active match (the one `findNext`/`findPrev` last landed on) is drawn in a
 * distinct emphasis colour so the user can tell it apart from the other hits,
 * matching a browser find bar's current-vs-other highlighting.
 */
import {
  sliceHorizontalExtent,
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
  const clipFrames = new OverlayClipFrames(layer, cssWidth, cssHeight);

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const start = Math.max(0, Math.min(slice.start, run.text.length));
      const end = Math.max(start, Math.min(slice.end, run.text.length));
      // The slice's glyphs where they are painted inside the run's clips
      // (overlay-geometry.ts). A run without cluster boxes (not projected
      // from a layout) is measured along the run and cut the same way.
      const rects = run.clusterBoxes
        ? runOverlayRects(run, start, end)
        : measuredSliceRects(run, start, end, measureForFont);
      for (const rect of rects) {
        // A clipped run's box also lies in containers sized to its clips,
        // which cut it like the glyphs at any angle.
        const frame = clipFrames.place(run);
        const box = document.createElement('div');
        // A vertical (tbRl) page reports the run at its physical top-left and
        // carries a rotate transform; apply the same transform (about the box's
        // top-left) so the highlight lies along the rotated glyph run, exactly as
        // the selection span does.
        const transform = frame.orientation
          ? `transform:${frame.orientation};transform-origin:top left;`
          : '';
        // Positioned as a % of the page's intended CSS box so the box scales with
        // the container under external CSS scaling. The rect's offsets run
        // along the run's own axes, turned with the run.
        box.style.cssText =
          `position:absolute;` +
          frame.position(rect.x, rect.y) +
          `width:${frame.width(rect.width)};height:${frame.height(rect.height)};` +
          transform +
          `background:${fill};pointer-events:none;`;
        frame.parent.appendChild(box);
      }
    }
  }
}

/** The rects of a slice of a run that carries no cluster boxes: its measured
 * extent along the run over the run's font box, cut to the run's clips. */
function measuredSliceRects(
  run: DocxTextRunInfo,
  start: number,
  end: number,
  measureForFont: (font: string) => (s: string) => number,
): RunRect[] {
  const { x, width } = measuredSliceInterval(run, start, end, run.letterSpacingPx ?? 0, measureForFont);
  const across = runFontBox(run);
  return clipRunRect(run, { x, y: across.y, width, height: across.height });
}

/** The extent of a slice of a run that carries no cluster boxes (a run not
 * projected from a layout): measured as a prefix of the run's text, from
 * its left edge. */
function measuredSliceInterval(
  run: DocxTextRunInfo,
  start: number,
  end: number,
  pitch: number,
  measureForFont: (font: string) => (s: string) => number,
): { x: number; width: number } {
  // Measure the way the run was laid out and drawn (its reference-em font
  // scaled onto the run when projected with one).
  const measure = runOverlayMeasure(run, measureForFont);
  const extent = sliceHorizontalExtent(run.text, start, end, measure);
  // Match offsets are UTF-16; count tracking with the same rendered-
  // character rule as layout. Marks and unselected SHY add no run pitch.
  // Exclude the pitch following the slice's final rendered character.
  const sliceCount = countSpacingCharacters(run.text.slice(start, end)).pitchBearing;
  const prefixCount = countSpacingCharacters(run.text.slice(0, start)).pitchBearing;
  const pitchedX = extent.x + prefixCount * pitch;
  // Keep main's multiply-before-add order for ordinary slices as well.
  const pitchedWidth = extent.width + Math.max(0, sliceCount - 1) * pitch;
  // §17.3.2.10 縦中横 alone compresses natural extents into the retained
  // one-em cell. Ordinary runs preserve the established natural measurement.
  const k = tateChuYokoOverlayScale(run, measure);
  return { x: pitchedX * k, width: pitchedWidth * k };
}
