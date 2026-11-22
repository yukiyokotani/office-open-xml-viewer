import { overlayPercent } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './types.js';

type CssRect = Readonly<{ x: number; y: number; width: number; height: number }>;

/** True when DocxTextRunInfo.clipBounds reports nothing of the run visible. */
export function runIsClippedAway(run: Readonly<DocxTextRunInfo>): boolean {
  const clip = run.clipBounds;
  return clip !== undefined && !(clip.width > 0 && clip.height > 0);
}

/** A highlight rectangle of an unturned `run`, cut to its clip; null when
 * nothing of it is visible. A turned run keeps its rectangle (its clip is
 * enforced by the overlay container, see overlayClipFrame). */
export function clipRunRect(run: Readonly<DocxTextRunInfo>, rect: CssRect): CssRect | null {
  const clip = run.clipBounds;
  if (!clip) return rect;
  if (runIsClippedAway(run)) return null;
  if (run.transform) return rect;
  const left = Math.max(rect.x, clip.x);
  const top = Math.max(rect.y, clip.y);
  const right = Math.min(rect.x + rect.width, clip.x + clip.width);
  const bottom = Math.min(rect.y + rect.height, clip.y + clip.height);
  if (right <= left || bottom <= top) return null;
  return { x: left, y: top, width: right - left, height: bottom - top };
}

/** Where an overlay element for a run is placed, and how its page CSS px
 * geometry is written there. */
export interface OverlayClipFrame {
  readonly parent: HTMLElement;
  readonly left: (xPx: number) => string;
  readonly top: (yPx: number) => string;
  readonly width: (widthPx: number) => string;
  readonly height: (heightPx: number) => string;
}

/** WORD_CELL_CONTENT_CLIP: the overlay frame for `run`. An unclipped run is
 * placed on `layer` in percentages of the page. A clipped run is placed in a
 * container positioned and sized to its clip (in percentages of the page, so
 * it scales with the page) with `overflow:hidden`, so the browser clips the
 * run's hit box and highlights for any transform or responsive resize.
 * Runs with the same clip share one container, kept in `containers`. */
export function overlayClipFrame(
  layer: HTMLElement,
  run: Readonly<DocxTextRunInfo>,
  cssWidth: number,
  cssHeight: number,
  containers: Map<string, HTMLElement>,
): OverlayClipFrame {
  const clip = run.clipBounds;
  if (!clip) {
    return {
      parent: layer,
      left: (x) => overlayPercent(x, cssWidth),
      top: (y) => overlayPercent(y, cssHeight),
      width: (w) => overlayPercent(w, cssWidth),
      height: (h) => overlayPercent(h, cssHeight),
    };
  }
  const width = Math.max(0, clip.width);
  const height = Math.max(0, clip.height);
  const key = `${clip.x},${clip.y},${width},${height}`;
  let parent = containers.get(key);
  if (!parent) {
    parent = (layer.ownerDocument ?? document).createElement('div');
    parent.style.cssText =
      'position:absolute;' +
      `left:${overlayPercent(clip.x, cssWidth)};top:${overlayPercent(clip.y, cssHeight)};` +
      `width:${overlayPercent(width, cssWidth)};height:${overlayPercent(height, cssHeight)};` +
      'overflow:hidden;pointer-events:none;';
    layer.appendChild(parent);
    containers.set(key, parent);
  }
  return {
    parent,
    left: (x) => overlayPercent(x - clip.x, width),
    top: (y) => overlayPercent(y - clip.y, height),
    width: (w) => overlayPercent(w, width),
    height: (h) => overlayPercent(h, height),
  };
}
