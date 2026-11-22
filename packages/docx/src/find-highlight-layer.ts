import { countSpacingCharacters } from './character-spacing.js';
import { OverlayClipFrames } from './run-clip.js';
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
  const clipFrames = new OverlayClipFrames(layer, cssWidth, cssHeight);

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const start = Math.max(0, Math.min(slice.start, run.text.length));
      const end = Math.max(start, Math.min(slice.end, run.text.length));
      const pitch = run.letterSpacingPx ?? 0;
      // A projected run carries where its clusters are painted: the slice
      // covers one box per painted interval, and nothing is measured.
      const intervals = run.clusterBoxes
        ? clusterSliceIntervals(run.clusterBoxes, start, end, pitch)
        : [measuredSliceInterval(run, start, end, pitch, measureForFont)];
      for (const { x, width } of intervals) {
        if (width <= 0) continue;
        // WORD_CELL_CONTENT_CLIP: a clipped run's box lies in a container sized
        // to its clip, which cuts it like the glyphs.
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
        // the container under external CSS scaling. A vertical (tbRl) run's
        // transform is about `top left`, so the %-placed origin rotates into the
        // column exactly as before at 1× scale.
        box.style.cssText =
          `position:absolute;` +
          // The slice offset runs along the run's own axis.
          frame.position(x, 0) +
          `width:${frame.width(width)};height:${frame.height(run.h)};` +
          transform +
          `background:${fill};pointer-events:none;`;
        frame.parent.appendChild(box);
      }
    }
  }
}

type ClusterBoxes = NonNullable<DocxTextRunInfo['clusterBoxes']>;

/** Where the characters `[start, end)` of a run are painted, from the run's
 * retained cluster boxes: adjacent boxes merged into one interval. A slice
 * that covers part of a cluster takes that share of its box. The tracking
 * that follows the slice's final rendered cluster is left out, on the side
 * that cluster ends on. */
function clusterSliceIntervals(
  clusters: ClusterBoxes,
  start: number,
  end: number,
  pitch: number,
): { x: number; width: number }[] {
  // Clusters are listed in source order: the first one that ends after `start`.
  let low = 0;
  let high = clusters.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (clusters[middle]!.end <= start) low = middle + 1;
    else high = middle;
  }
  const boxes: { x: number; width: number; rtl: boolean }[] = [];
  for (let index = low; index < clusters.length && clusters[index]!.start < end; index += 1) {
    const cluster = clusters[index]!;
    const length = Math.max(1, cluster.end - cluster.start);
    const from = (Math.max(start, cluster.start) - cluster.start) / length;
    const to = (Math.min(end, cluster.end) - cluster.start) / length;
    boxes.push({
      x: cluster.x + cluster.width * (cluster.rtl ? 1 - to : from),
      width: cluster.width * (to - from),
      rtl: cluster.rtl,
    });
  }
  let final = boxes.length - 1;
  while (final >= 0 && !(boxes[final]!.width > 0)) final -= 1;
  if (final >= 0 && pitch !== 0) {
    const box = boxes[final]!;
    const trimmed = Math.min(pitch, box.width);
    box.width -= trimmed;
    if (box.rtl) box.x += trimmed;
  }
  const merged: { x: number; width: number }[] = [];
  for (const box of boxes.filter((candidate) => candidate.width > 0).sort((left, right) => left.x - right.x)) {
    const last = merged.at(-1);
    if (last && Math.abs(last.x + last.width - box.x) < 1e-6) last.width += box.width;
    else merged.push({ x: box.x, width: box.width });
  }
  return merged;
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
