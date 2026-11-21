import type { PointPt, ResourcePlacement } from './types.js';

/** Visible outline of an inline resource placement, in the placement's own
 * coordinates: its authored paint box turned by `rotationDeg` about the box
 * centre (clockwise, as painted), cut to the `lineClip` band when the paint is
 * clipped (WORD_INLINE_ROTATED_PICTURE_LINE_CLIP). Hit testing and element
 * bounds use this instead of the authored `bounds`. Empty when the band
 * removes everything. */
export function resourcePlacementOutline(
  placement: Pick<ResourcePlacement, 'bounds' | 'rotationDeg' | 'lineClip'>,
): readonly PointPt[] {
  const { xPt, yPt, widthPt, heightPt } = placement.bounds;
  const cx = xPt + widthPt / 2;
  const cy = yPt + heightPt / 2;
  const angle = ((placement.rotationDeg ?? 0) * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  let outline: PointPt[] = [
    [-widthPt / 2, -heightPt / 2], [widthPt / 2, -heightPt / 2],
    [widthPt / 2, heightPt / 2], [-widthPt / 2, heightPt / 2],
  ].map(([dx, dy]) => ({ xPt: cx + dx! * cos - dy! * sin, yPt: cy + dx! * sin + dy! * cos }));
  const clip = placement.lineClip;
  if (!clip) return outline;
  outline = clipHorizontal(outline, yPt + clip.topPt, 1);
  return clipHorizontal(outline, yPt + clip.bottomPt, -1);
}

/** One Sutherland-Hodgman pass against the horizontal line y = edge, keeping
 * the side where `side * (y - edge) >= 0`. */
function clipHorizontal(polygon: readonly PointPt[], edge: number, side: 1 | -1): PointPt[] {
  const kept: PointPt[] = [];
  for (let index = 0; index < polygon.length; index++) {
    const current = polygon[index]!;
    const previous = polygon[(index + polygon.length - 1) % polygon.length]!;
    const currentIn = side * (current.yPt - edge) >= 0;
    const previousIn = side * (previous.yPt - edge) >= 0;
    if (currentIn !== previousIn) {
      const t = (edge - previous.yPt) / (current.yPt - previous.yPt);
      kept.push({ xPt: previous.xPt + t * (current.xPt - previous.xPt), yPt: edge });
    }
    if (currentIn) kept.push(current);
  }
  return kept;
}
