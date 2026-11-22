import { overlayPercent } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './types.js';

type CssRect = Readonly<{ x: number; y: number; width: number; height: number }>;
type RunClip = NonNullable<DocxTextRunInfo['clips']>[number];
/** [a, b, c, d, e, f]: a CSS-px affine map (x' = a x + c y + e). */
type Affine = readonly [number, number, number, number, number, number];

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0];

/** True when DocxTextRunInfo.clipBounds reports nothing of the run visible. */
export function runIsClippedAway(run: Readonly<DocxTextRunInfo>): boolean {
  const clip = run.clipBounds;
  return clip !== undefined && !(clip.width > 0 && clip.height > 0);
}

/** A highlight rectangle of an unturned `run`, cut to its clip; null when
 * nothing of it is visible. A turned run keeps its rectangle (its clips are
 * enforced by the overlay containers, see OverlayClipFrames). */
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

/** The linear part of an orientation transform as text-run-projection emits
 * it (`rotate(±90deg)` or `matrix(a, b, c, d, 0, 0)`); identity when absent. */
function orientation(transform: string | undefined): readonly [number, number, number, number] {
  if (!transform) return [1, 0, 0, 1];
  if (transform === 'rotate(90deg)') return [0, 1, -1, 0];
  if (transform === 'rotate(-90deg)') return [0, -1, 1, 0];
  const match = /^matrix\(([^,]+),([^,]+),([^,]+),([^,]+),\s*0,\s*0\)$/u.exec(transform);
  const values = match ? match.slice(1, 5).map(Number) : [];
  return values.length === 4 && values.every(Number.isFinite)
    ? [values[0]!, values[1]!, values[2]!, values[3]!] : [1, 0, 0, 1];
}

function placed(x: number, y: number, transform: string | undefined): Affine {
  const [a, b, c, d] = orientation(transform);
  return [a, b, c, d, x, y];
}

function compose(outer: Affine, inner: Affine): Affine {
  const [a, b, c, d, e, f] = outer;
  const [p, q, r, s, t, u] = inner;
  return [a * p + c * q, b * p + d * q, a * r + c * s, b * r + d * s, a * t + c * u + e, b * t + d * u + f];
}

function invert(m: Affine): Affine {
  const [a, b, c, d, e, f] = m;
  const det = a * d - b * c;
  if (!(Math.abs(det) > 0)) return IDENTITY;
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det];
}

/** The CSS orientation of a linear map, or '' for none. */
function orientationCss(m: Affine): string {
  const [a, b, c, d] = m;
  const near = (value: number, target: number) => Math.abs(value - target) < 1e-9;
  if (near(a, 1) && near(b, 0) && near(c, 0) && near(d, 1)) return '';
  return `matrix(${a}, ${b}, ${c}, ${d}, 0, 0)`;
}

interface OpenContainer {
  readonly key: string;
  readonly element: HTMLElement;
  /** Page CSS px to the container's local px. */
  readonly fromPage: Affine;
  readonly width: number;
  readonly height: number;
}

/** Where an overlay element of a run goes, and its CSS placement there. */
export interface OverlayPlacement {
  readonly parent: HTMLElement;
  /** `left`/`top` in percentages of the parent, for a box at offsets
   * (dx, dy) along the run's own axes from its origin. */
  readonly position: (dx: number, dy: number) => string;
  /** A width or height in percentages of the parent's width or height. */
  readonly width: (lengthPx: number) => string;
  readonly height: (lengthPx: number) => string;
  /** The element's orientation against the parent (about its top-left), ''
   * for none; further transforms compose after it. */
  readonly orientation: string;
}

/** WORD_CELL_CONTENT_CLIP: overlay elements of clipped runs are placed in
 * nested `overflow:hidden` containers, one per clip in DocxTextRunInfo.clips,
 * each positioned, sized and oriented to its clip in percentages of its
 * parent, so the browser clips hit testing, selection and highlights for any
 * transform or responsive resize. Containers are shared only by runs placed
 * one after another with the same clips, so the overlay keeps document order:
 * an unclipped or differently clipped run in between opens new containers. */
export class OverlayClipFrames {
  private readonly open: OpenContainer[] = [];

  constructor(
    private readonly layer: HTMLElement,
    private readonly cssWidth: number,
    private readonly cssHeight: number,
  ) {}

  /** The placement of an element of `run`, added after every element placed
   * before it. */
  place(run: Readonly<DocxTextRunInfo>): OverlayPlacement {
    const clips = run.clips ?? [];
    let shared = 0;
    while (shared < clips.length && shared < this.open.length
      && this.open[shared]!.key === clipKey(clips[shared]!)) shared++;
    this.open.length = shared;
    for (const clip of clips.slice(shared)) this.openContainer(clip);
    const frame = this.open.at(-1);
    const width = frame?.width ?? this.cssWidth;
    const height = frame?.height ?? this.cssHeight;
    const runToParent = compose(frame?.fromPage ?? IDENTITY, placed(run.x, run.y, run.transform));
    return {
      parent: frame?.element ?? this.layer,
      // On the layer itself a run keeps its own transform text.
      orientation: frame ? orientationCss(runToParent) : run.transform ?? '',
      position: (dx, dy) => {
        const corner = compose(runToParent, [1, 0, 0, 1, dx, dy]);
        return `left:${overlayPercent(corner[4], width)};top:${overlayPercent(corner[5], height)};`;
      },
      width: (length) => overlayPercent(length, width),
      height: (length) => overlayPercent(length, height),
    };
  }

  private openContainer(clip: RunClip): void {
    const outer = this.open.at(-1);
    const parent = outer?.element ?? this.layer;
    const parentWidth = outer?.width ?? this.cssWidth;
    const parentHeight = outer?.height ?? this.cssHeight;
    const clipToPage = placed(clip.x, clip.y, clip.transform);
    const clipToParent = compose(outer?.fromPage ?? IDENTITY, clipToPage);
    const width = Math.max(0, clip.width);
    const height = Math.max(0, clip.height);
    const turn = orientationCss(clipToParent);
    const element = (parent.ownerDocument ?? document).createElement('div');
    element.style.cssText =
      'position:absolute;' +
      `left:${overlayPercent(clipToParent[4], parentWidth)};top:${overlayPercent(clipToParent[5], parentHeight)};` +
      `width:${overlayPercent(width, parentWidth)};height:${overlayPercent(height, parentHeight)};` +
      (turn ? `transform:${turn};transform-origin:top left;` : '') +
      'overflow:hidden;pointer-events:none;';
    parent.appendChild(element);
    this.open.push({ key: clipKey(clip), element, fromPage: invert(clipToPage), width, height });
  }
}

function clipKey(clip: RunClip): string {
  return `${clip.x},${clip.y},${clip.width},${clip.height},${clip.transform ?? ''}`;
}
