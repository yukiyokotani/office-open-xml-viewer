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

/** At most this many polygon pieces are tracked while cutting a box by
 * nested union clips; past it a union is replaced by its bounding box, a
 * superset that keeps the work linear in the number of clips. */
const MAX_CLIP_PIECES = 64;

type Box = Readonly<{ x: number; y: number; width: number; height: number; transform?: string }>;

type LocalRect = Readonly<{ x: number; y: number; width: number; height: number }>;

/** Unions with more rectangles than this skip the pairwise containment pass. */
const MAX_CONTAINMENT_RECTS = 512;

/** A clip's rectangles reduced to those that matter for `near` (page
 * polygons being cut): rectangles without area, outside the polygons'
 * extent, duplicated or inside another rectangle are dropped. An empty
 * result clips everything. Returned as page polygons. */
function clipRegions(clip: Readonly<DocxClipRegion>, near: readonly Point[][]): Point[][] {
  const [a, b, c, d] = orientation(clip.transform);
  const det = a * d - b * c;
  // The extent of `near` in the clip's own frame.
  let minU = Infinity;
  let minV = Infinity;
  let maxU = -Infinity;
  let maxV = -Infinity;
  for (const polygon of near) {
    for (const point of polygon) {
      const dx = point.x - clip.x;
      const dy = point.y - clip.y;
      const u = (d * dx - c * dy) / det;
      const v = (-b * dx + a * dy) / det;
      minU = Math.min(minU, u);
      minV = Math.min(minV, v);
      maxU = Math.max(maxU, u);
      maxV = Math.max(maxV, v);
    }
  }
  const seen = new Set<string>();
  let rects: LocalRect[] = [];
  for (const rect of clip.rects ?? [{ x: 0, y: 0, width: clip.width, height: clip.height }]) {
    if (!(rect.width > 0 && rect.height > 0)) continue;
    if (rect.x >= maxU || rect.y >= maxV || rect.x + rect.width <= minU || rect.y + rect.height <= minV) continue;
    const key = `${rect.x},${rect.y},${rect.width},${rect.height}`;
    if (seen.has(key)) continue;
    seen.add(key);
    rects.push(rect);
  }
  if (rects.length > 1 && rects.length <= MAX_CONTAINMENT_RECTS) {
    const kept: LocalRect[] = [];
    for (const rect of [...rects].sort((left, right) => right.width * right.height - left.width * left.height)) {
      const inside = kept.some((outer) => rect.x >= outer.x && rect.y >= outer.y
        && rect.x + rect.width <= outer.x + outer.width && rect.y + rect.height <= outer.y + outer.height);
      if (!inside) kept.push(rect);
    }
    rects = kept;
  }
  return rects.map((rect) => placedPolygon(
    clip.x + a * rect.x + c * rect.y, clip.y + b * rect.x + d * rect.y, rect.width, rect.height, clip.transform,
  ));
}

function boundingPolygon(polygons: readonly Point[][]): Point[] {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const polygon of polygons) {
    for (const point of polygon) {
      left = Math.min(left, point.x);
      top = Math.min(top, point.y);
      right = Math.max(right, point.x);
      bottom = Math.max(bottom, point.y);
    }
  }
  return [{ x: left, y: top }, { x: right, y: top }, { x: right, y: bottom }, { x: left, y: bottom }];
}

/** The visible pieces (convex page polygons) of `box` under `clips`, each a
 * union of its rectangles; identical pieces are kept once. The cap applies
 * after each union is reduced (clipRegions), so duplicate or nested
 * rectangles never trigger the bounding-box superset. */
function visiblePieces(box: Box, clips: readonly Readonly<DocxClipRegion>[] | undefined): Point[][] {
  let pieces: Point[][] = [placedPolygon(box.x, box.y, box.width, box.height, box.transform)];
  for (const clip of clips ?? []) {
    let regions = clipRegions(clip, pieces);
    if (regions.length === 0) return [];
    // Past the cap a union is cut as its bounding box, a superset.
    if (pieces.length * regions.length > MAX_CLIP_PIECES) regions = [boundingPolygon(regions)];
    const seen = new Set<string>();
    pieces = pieces.flatMap((piece) => regions.map((region) => cutPolygon(piece, region)))
      .filter((piece) => {
        if (piece.length < 3 || !(Math.abs(area(piece)) > 1e-9)) return false;
        const key = piece.map((point) => `${point.x.toFixed(6)},${point.y.toFixed(6)}`).join(';');
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    if (pieces.length === 0) return [];
  }
  return pieces;
}

/** WORD_CELL_CONTENT_CLIP: whether any of a box placed like
 * DocxTextRunInfo.highlightBounds survives `clips` (each a union of its
 * rectangles with area; an empty union hides everything). */
export function clippedBoxVisible(box: Box, clips: readonly Readonly<DocxClipRegion>[] | undefined): boolean {
  return visiblePieces(box, clips).length > 0;
}

/** The page bounding box of the visible part of `box` under `clips`, in the
 * box's CSS px space; null when nothing is visible. */
export function clippedBoxBounds(
  box: Box,
  clips: readonly Readonly<DocxClipRegion>[] | undefined,
): Readonly<{ x: number; y: number; width: number; height: number }> | null {
  const pieces = visiblePieces(box, clips);
  if (pieces.length === 0) return null;
  const [corner, , opposite] = boundingPolygon(pieces);
  return { x: corner!.x, y: corner!.y, width: opposite!.x - corner!.x, height: opposite!.y - corner!.y };
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
  if (run.clips && !runGlyphsVisible(run)) return null;
  return run.clips ? { ...rect, clips: run.clips } : rect;
}

/** WORD_CELL_CONTENT_CLIP: whether any glyph of `run` is painted inside its
 * clips: a glyph box (DocxTextRunInfo.clusterBoxes) along the run, the font
 * box across it; tracking after a glyph paints nothing. A run without
 * cluster boxes is judged by its font box. */
export function runGlyphsVisible(run: Readonly<DocxTextRunInfo>): boolean {
  const boxes = run.clusterBoxes;
  if (!boxes || boxes.length === 0) return runPartVisible(run, 0, run.highlightBounds?.width ?? run.w);
  return boxes.some((box) => {
    const width = box.glyphWidth ?? box.width;
    return width > 0 && runPartVisible(run, box.glyphX ?? box.x, width);
  });
}

/** The painted box of the part of `run` from `offset` along its own axis
 * (from the run's origin), `length` long: across the run it spans the run's
 * font box (highlightBounds), where its glyphs are painted, not the line
 * height that additional line spacing adds. */
export function runPaintedPart(
  run: Readonly<DocxTextRunInfo>,
  offset: number,
  length: number,
): Readonly<{ x: number; y: number; width: number; height: number; transform?: string }> {
  const [a, b, c, d] = orientation(run.transform);
  const font = run.highlightBounds;
  // The font box's position across the run, from the run's origin.
  const across = font ? c * (font.x - run.x) + d * (font.y - run.y) : 0;
  return {
    x: run.x + a * offset + c * across, y: run.y + b * offset + d * across, width: length,
    height: font?.height ?? run.h,
    ...(run.transform ? { transform: run.transform } : {}),
  };
}

/** Whether anything painted in the part of `run` from `offset` along its own
 * axis, `length` long, survives its clips (runPaintedPart). */
export function runPartVisible(run: Readonly<DocxTextRunInfo>, offset: number, length: number): boolean {
  return clippedBoxVisible(runPaintedPart(run, offset, length), run.clips);
}

/** Two highlight boxes (placed like DocxTextRunInfo.highlightBounds) of one
 * rendered line, joined across the gap between them: they must share their
 * orientation, height and position across the line, measured in their own
 * frame, so a turned line joins along its own axis and two turned columns
 * never join. Returns the joined box's origin and length, or null. */
export function joinLineBoxes(
  previous: Readonly<{ x: number; y: number; width: number; height: number; transform?: string }>,
  next: Readonly<{ x: number; y: number; width: number; height: number; transform?: string }>,
): Readonly<{ x: number; y: number; width: number }> | null {
  if (previous.transform !== next.transform || previous.height !== next.height) return null;
  const [a, b, c, d] = orientation(previous.transform);
  // Orientations are rotations or reflections: the inverse is the transpose.
  const along = (x: number, y: number) => a * x + b * y;
  const across = (x: number, y: number) => c * x + d * y;
  const line = across(previous.x, previous.y);
  if (Math.abs(line - across(next.x, next.y)) > 1e-9) return null;
  const start = Math.min(along(previous.x, previous.y), along(next.x, next.y));
  const end = Math.max(along(previous.x, previous.y) + previous.width, along(next.x, next.y) + next.width);
  return { x: a * start + c * line, y: b * start + d * line, width: end - start };
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
 * nested clipping (`overflow:clip`) containers, one per clip in DocxTextRunInfo.clips,
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
      // overflow:clip, unlike hidden, makes no scroll container, so browser
      // find cannot scroll the frame off the painted text; browsers without
      // it keep hidden.
      'overflow:hidden;overflow:clip;overflow-clip-margin:0;pointer-events:none;';
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
  if (!clip.rects) return '';
  const rects = clip.rects.filter((rect) => rect.width > 0 && rect.height > 0);
  // An empty union clips everything, as the Canvas clip does.
  if (rects.length === 0) return 'clip-path:polygon(0% 0%,0% 0%,0% 0%);';
  const point = (x: number, y: number) => `${overlayPercent(x, width)} ${overlayPercent(y, height)}`;
  const parts: string[] = [point(0, 0)];
  for (const rect of rects) {
    const right = rect.x + rect.width;
    const bottom = rect.y + rect.height;
    parts.push(point(rect.x, rect.y), point(right, rect.y), point(right, bottom), point(rect.x, bottom),
      point(rect.x, rect.y), point(0, 0));
  }
  return `clip-path:polygon(${parts.join(',')});`;
}
