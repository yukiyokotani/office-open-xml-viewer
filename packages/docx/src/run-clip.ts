import { overlayPercent } from '@silurus/ooxml-core';
import type { DocxClipRegion, DocxTextRunInfo } from './types.js';

type RunClip = NonNullable<DocxTextRunInfo['clips']>[number];
/** [a, b, c, d, e, f]: a CSS-px affine map (x' = a x + c y + e). */
type Affine = readonly [number, number, number, number, number, number];

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0];

type Point = Readonly<{ x: number; y: number }>;

/** The page polygon of a box placed like DocxTextRunInfo.highlightBounds. */
function placedPolygon(x: number, y: number, width: number, height: number, transform: string | undefined): Point[] {
  const [a, b, c, d] = orientation(transform);
  return [[0, 0], [width, 0], [width, height], [0, height]].map(([u, v]) => ({
    x: x + a * u! + c * v!, y: y + b * u! + d * v!,
  }));
}

function area(polygon: readonly Point[]): number {
  return polygon.reduce((sum, point, index) => {
    const next = polygon[(index + 1) % polygon.length]!;
    return sum + point.x * next.y - next.x * point.y;
  }, 0) / 2;
}

/** `subject` cut by the convex polygon `clip` (Sutherland-Hodgman). */
function cutPolygon(subject: readonly Point[], clip: readonly Point[]): Point[] {
  const sign = area(clip) >= 0 ? 1 : -1;
  let points = [...subject];
  for (let index = 0; index < clip.length && points.length > 0; index++) {
    const p = clip[index]!;
    const q = clip[(index + 1) % clip.length]!;
    const side = (point: Point) => sign * ((q.x - p.x) * (point.y - p.y) - (q.y - p.y) * (point.x - p.x));
    const next: Point[] = [];
    for (let i = 0; i < points.length; i++) {
      const current = points[i]!;
      const following = points[(i + 1) % points.length]!;
      const sc = side(current);
      const sf = side(following);
      if (sc >= 0) next.push(current);
      if ((sc >= 0) !== (sf >= 0)) {
        const t = sc / (sc - sf);
        next.push({ x: current.x + t * (following.x - current.x), y: current.y + t * (following.y - current.y) });
      }
    }
    points = next;
  }
  return points;
}

/** WORD_CELL_CONTENT_CLIP: whether any of a box placed like
 * DocxTextRunInfo.highlightBounds survives `clips` (each a union of its
 * rectangles); empty overlaps (zero area) count as hidden. */
export function clippedBoxVisible(
  box: Readonly<{ x: number; y: number; width: number; height: number; transform?: string }>,
  clips: readonly Readonly<DocxClipRegion>[] | undefined,
): boolean {
  let pieces: Point[][] = [placedPolygon(box.x, box.y, box.width, box.height, box.transform)];
  for (const clip of clips ?? []) {
    const regions = (clip.rects ?? [{ x: 0, y: 0, width: clip.width, height: clip.height }]).map((rect) => {
      const [a, b, c, d] = orientation(clip.transform);
      return placedPolygon(clip.x + a * rect.x + c * rect.y, clip.y + b * rect.x + d * rect.y, rect.width, rect.height, clip.transform);
    });
    pieces = pieces.flatMap((piece) => regions.map((region) => cutPolygon(piece, region)))
      .filter((piece) => piece.length >= 3 && Math.abs(area(piece)) > 1e-9);
    if (pieces.length === 0) return false;
  }
  return pieces.length > 0;
}

/** The comment highlight of a run, shared by the comment API and the
 * viewer's tint: the run's font box (highlightBounds, else its box) with its
 * orientation and clips; null when the clips leave none of it visible. */
export function runHighlightRect(run: Readonly<DocxTextRunInfo>): Readonly<{
  x: number; y: number; width: number; height: number;
  transform?: string; clips?: readonly Readonly<DocxClipRegion>[];
}> | null {
  const bounds = run.highlightBounds;
  const rect = {
    x: bounds?.x ?? run.x,
    y: bounds?.y ?? run.y,
    width: bounds?.width ?? run.w,
    height: bounds?.height ?? run.h,
    ...(run.transform ? { transform: run.transform } : {}),
  };
  if (!(rect.width > 0 && rect.height > 0)) return null;
  if (run.clips && !clippedBoxVisible(rect, run.clips)) return null;
  return run.clips ? { ...rect, clips: run.clips } : rect;
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
      unionClipPath(clip, width, height) +
      'overflow:hidden;pointer-events:none;';
    parent.appendChild(element);
    this.open.push({ key: clipKey(clip), element, fromPage: invert(clipToPage), width, height });
  }
}

function clipKey(clip: RunClip): string {
  return JSON.stringify(clip);
}

/** A union clip's rectangles as one clip-path polygon in percentages of the
 * container: each rectangle is traced from a common start point and joined
 * by zero-width bridges, so only the rectangles are inside. */
function unionClipPath(clip: RunClip, width: number, height: number): string {
  if (!clip.rects || clip.rects.length === 0) return '';
  const point = (x: number, y: number) => `${overlayPercent(x, width)} ${overlayPercent(y, height)}`;
  const parts: string[] = [point(0, 0)];
  for (const rect of clip.rects) {
    const right = rect.x + rect.width;
    const bottom = rect.y + rect.height;
    parts.push(point(rect.x, rect.y), point(right, rect.y), point(right, bottom), point(rect.x, bottom),
      point(rect.x, rect.y), point(0, 0));
  }
  return `clip-path:polygon(${parts.join(',')});`;
}
