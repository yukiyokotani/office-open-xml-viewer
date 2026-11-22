import type { DocxTextRunInfo } from './types.js';

type CssRect = Readonly<{ x: number; y: number; width: number; height: number }>;
type Orientation = readonly [number, number, number, number];

/** The linear part of a run's CSS orientation transform (text-run-projection
 * emits `rotate(±90deg)` or `matrix(a, b, c, d, 0, 0)`); null when unknown. */
function orientation(transform: string | undefined): Orientation | null {
  if (!transform) return [1, 0, 0, 1];
  if (transform === 'rotate(90deg)') return [0, 1, -1, 0];
  if (transform === 'rotate(-90deg)') return [0, -1, 1, 0];
  const match = /^matrix\(([^,]+),([^,]+),([^,]+),([^,]+),\s*0,\s*0\)$/u.exec(transform);
  if (!match) return null;
  const values = match.slice(1, 5).map(Number);
  return values.every(Number.isFinite) ? values as unknown as Orientation : null;
}

/** A highlight or hit rectangle of `run`, cut to its paint clip
 * (DocxTextRunInfo.clipBounds); null when nothing of it is visible. Both are
 * placed like highlightBounds: x/y is the page position of the rectangle's
 * top-left corner, width/height lie along the run's own (transformed) axes. */
export function clipRunRect(run: Readonly<DocxTextRunInfo>, rect: CssRect): CssRect | null {
  const clip = run.clipBounds;
  if (!clip) return rect;
  const m = orientation(run.transform);
  if (!m) return rect;
  const [a, b, c, d] = m;
  const det = a * d - b * c;
  // Page positions to offsets along the run's axes from its origin, and back.
  const local = (x: number, y: number) => {
    const dx = x - run.x;
    const dy = y - run.y;
    return { x: (d * dx - c * dy) / det, y: (-b * dx + a * dy) / det };
  };
  const r = local(rect.x, rect.y);
  const k = local(clip.x, clip.y);
  const left = Math.max(r.x, k.x);
  const top = Math.max(r.y, k.y);
  const right = Math.min(r.x + rect.width, k.x + clip.width);
  const bottom = Math.min(r.y + rect.height, k.y + clip.height);
  if (right <= left || bottom <= top) return null;
  return {
    x: run.x + a * left + c * top,
    y: run.y + b * left + d * top,
    width: right - left,
    height: bottom - top,
  };
}

/** A rectangle at offsets (dx, dy) along `run`'s own axes from its origin,
 * placed like clipRunRect's rectangles. */
export function placeRunRect(
  run: Readonly<DocxTextRunInfo>,
  dx: number,
  dy: number,
  width: number,
  height: number,
): CssRect {
  const [a, b, c, d] = orientation(run.transform) ?? [1, 0, 0, 1];
  return { x: run.x + a * dx + c * dy, y: run.y + b * dx + d * dy, width, height };
}

/** The visible part of `run`'s own box ([0, w] x [0, h] along its axes from
 * its origin) under its clip; undefined when the run is unclipped, null when
 * nothing is visible. */
export function visibleRunExtent(
  run: Readonly<DocxTextRunInfo>,
): Readonly<{ left: number; top: number; right: number; bottom: number }> | null | undefined {
  if (!run.clipBounds) return undefined;
  const visible = clipRunRect(run, placeRunRect(run, 0, 0, run.w, run.h));
  if (!visible) return null;
  const [a, b, c, d] = orientation(run.transform) ?? [1, 0, 0, 1];
  const det = a * d - b * c;
  const dx = visible.x - run.x;
  const dy = visible.y - run.y;
  const left = (d * dx - c * dy) / det;
  const top = (-b * dx + a * dy) / det;
  return { left, top, right: left + visible.width, bottom: top + visible.height };
}
