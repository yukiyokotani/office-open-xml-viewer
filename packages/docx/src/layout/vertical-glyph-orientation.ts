// Shared East Asian upright-glyph projection for vertical text frames.
//
// A vertical frame (a DrawingML text box with vert/eaVert/mongolianVert, or a
// table cell with ECMA-376 §17.4.72 textDirection) lays its content out
// horizontally in a local frame that paint rotates by a quarter turn. In the
// "upright" variants (eaVert, mongolianVert, and the §17.18.93 `V` cell
// directions) each East Asian cluster is painted upright again while other
// text keeps the rotation. This module owns that per-cluster paint-operation
// split so both frames use one projection.

import type { TextPlacement } from './types.js';
import { EAST_ASIAN_RE } from './text.js';

/** Split a text placement's paint operations per cluster, marking East Asian
 * clusters upright (centred on their advance) and others sideways. */
export function eastAsianUprightPaintOps(
  placement: TextPlacement,
): TextPlacement['paintOps'] {
  return placement.clusters.map((cluster) => {
    const text = placement.text.slice(
      cluster.range.start - placement.range.start,
      cluster.range.end - placement.range.start,
    );
    const template = placement.paintOps.find((operation) =>
      operation.range.start <= cluster.range.start && operation.range.end >= cluster.range.end)
      ?? placement.paintOps[0]!;
    const upright = EAST_ASIAN_RE.test(text);
    return {
      ...template,
      text,
      range: cluster.range,
      offset: upright
        ? { xPt: cluster.offset.xPt + cluster.advancePt / 2, yPt: cluster.offset.yPt }
        : cluster.offset,
      glyphOrientation: upright ? 'upright' as const : 'sideways' as const,
    };
  });
}
