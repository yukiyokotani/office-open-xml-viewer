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
  wordArtEmoji = false,
  /** WORD_VERTICAL_LINE_CENTRED_BOX: the cross-line offset of sideways
   * (turned) clusters from the line axis the upright glyphs are centred on. */
  sidewaysOffsetPt = 0,
): TextPlacement['paintOps'] {
  return placement.clusters.map((cluster) => {
    const text = placement.optionalHyphenGlyph ? placement.text : placement.text.slice(
      cluster.range.start - placement.range.start,
      cluster.range.end - placement.range.start,
    );
    const template = placement.paintOps.find((operation) =>
      operation.range.start <= cluster.range.start && operation.range.end >= cluster.range.end)
      ?? placement.paintOps[0]!;
    // Word's stacked WordArt also uprights emoji presentation clusters.
    // Issue #1668 controls: split ZWJ, modifier and regional-indicator runs;
    // text-presentation symbols (copyright, registered) remain sideways.
    const upright = EAST_ASIAN_RE.test(text)
      || (wordArtEmoji && /\p{Emoji_Presentation}/u.test(text));
    return {
      ...template,
      text,
      range: cluster.range,
      offset: upright
        ? { xPt: cluster.offset.xPt + cluster.advancePt / 2, yPt: cluster.offset.yPt }
        : { xPt: cluster.offset.xPt, yPt: cluster.offset.yPt + sidewaysOffsetPt },
      glyphOrientation: upright ? 'upright' as const : 'sideways' as const,
    };
  });
}

/** WORD_VERTICAL_LINE_CENTRED_BOX: in an upright frame (eaVert,
 * mongolianVert, a tbRlV cell) a turned run keeps its baseline half its font
 * box's (ascent - descent) tailward of the line axis, as in a vertical
 * section; the run's font box gives that offset. */
export function uprightFrameSidewaysOffsetPt(placement: TextPlacement): number {
  const box = placement.highlightBounds;
  return box ? placement.origin.yPt - (box.yPt + box.heightPt / 2) : 0;
}

/** WORD_VERTICAL_RUBY_UPRIGHT_GUIDE in an upright frame: the guide's East
 * Asian clusters stand upright centred on their advance at the guide axis;
 * other clusters stay turned with their baseline offset scaled from the
 * base run's. */
export function uprightFrameRuby(placement: TextPlacement): TextPlacement['ruby'] {
  const ruby = placement.ruby;
  if (!ruby || ruby.paintOps.length === 0) return ruby;
  const ops = ruby.paintOps;
  const start = ops[0]!.origin.xPt;
  const end = start + ruby.advancePt;
  const baseOffset = uprightFrameSidewaysOffsetPt(placement);
  return {
    ...ruby,
    paintOps: ops.map((op, index) => {
      if (op.glyphOrientation) return op;
      const advance = (ops[index + 1]?.origin.xPt ?? end) - op.origin.xPt;
      if (EAST_ASIAN_RE.test(op.text)) {
        return { ...op, origin: { xPt: op.origin.xPt + advance / 2, yPt: op.origin.yPt }, glyphOrientation: 'upright' as const };
      }
      const offset = placement.fontSizePt > 0 ? baseOffset * op.fontSizePt / placement.fontSizePt : 0;
      return { ...op, glyphOrientation: 'sideways' as const, ...(offset !== 0 ? { glyphOffsetPt: { xPt: 0, yPt: offset } } : {}) };
    }),
  };
}
