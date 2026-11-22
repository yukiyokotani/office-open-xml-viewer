import type { DocxTextRunInfo } from './types.js';

type CssRect = Readonly<{ x: number; y: number; width: number; height: number }>;

/** A highlight or hit rectangle of `run`, cut to the run's paint clip
 * (DocxTextRunInfo.clipBounds); null when nothing of it is visible. A
 * transformed run keeps its rectangle: its clip is not axis-aligned with it. */
export function clipRunRect(run: Readonly<DocxTextRunInfo>, rect: CssRect): CssRect | null {
  const clip = run.clipBounds;
  if (!clip || run.transform) return rect;
  const left = Math.max(rect.x, clip.x);
  const top = Math.max(rect.y, clip.y);
  const right = Math.min(rect.x + rect.width, clip.x + clip.width);
  const bottom = Math.min(rect.y + rect.height, clip.y + clip.height);
  if (right <= left || bottom <= top) return null;
  return { x: left, y: top, width: right - left, height: bottom - top };
}
