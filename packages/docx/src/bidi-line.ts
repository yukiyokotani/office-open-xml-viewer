// WordprocessingML segment ordering uses the shared UAX#9 kernel with
// Word's HL1 run overrides and content-level Canvas direction.
import {
  computeSegmentLineVisualOrder, segmentLineHasRtl,
  type SegmentLineVisualOrder,
} from '@silurus/ooxml-core/internal/bidi-line';
import { wordRtlAmbiguousCharacter, wordRunDirectionSpan } from './layout/script-compatibility.js';

const options = {
  isTab: (segment: unknown) => 'isTab' in (segment as object),
  isRtlMarked: (segment: unknown) => (segment as { rtl?: unknown }).rtl === true,
  digitsAsAN: (segment: unknown) => (segment as { digitsAsAN?: unknown }).digitsAsAN === true,
  isRtlAmbiguous: wordRtlAmbiguousCharacter,
  directionFromAnyOdd: true,
  isolateDirection: wordRunDirectionSpan,
  paintDirectionFromTrailingSpace: true,
};
export type LineVisualOrder = SegmentLineVisualOrder;
export const segmentsHaveRtl = (segments: readonly unknown[]): boolean =>
  segmentLineHasRtl(segments, options);
const unauthoredDirectionOptions = { ...options, isolateDirection: undefined };
const lineLevelOptions = {
  ...options, isolateDirection: undefined, paintDirectionFromTrailingSpace: undefined,
};
/** How a line's runs are ordered: 'word' resolves run-direction isolates
 *  (WORD_RUN_DIRECTION_ISOLATE) and keeps line-end spaces in their runs
 *  (WORD_LINE_END_SPACE_STAYS_IN_RUN); 'unauthored' (runs without w:rtl) skips
 *  the isolates; 'line-level' keeps plain line-level UAX #9 for lines whose
 *  character-spread gap model has not adopted those rules. */
export type LineOrderModel = 'word' | 'unauthored' | 'line-level';
export const computeLineVisualOrder = (
  segments: readonly unknown[], baseRtl: boolean, model: LineOrderModel = 'word',
): LineVisualOrder => model === 'line-level'
  ? computeSegmentLineVisualOrder(segments, baseRtl, lineLevelOptions)
  : computeSegmentLineVisualOrder(
    segments, baseRtl, model === 'word' ? options : unauthoredDirectionOptions, true,
  );

/** Physical edge a line aligns to, resolving logical start/end against base direction. */
export type AlignEdge = 'left' | 'right' | 'center' | 'justify';

/**
 * Resolve a paragraph's `w:jc` value (and base direction) to a physical edge.
 * ALL edge values are logical in WordprocessingML: `start`/`end` by definition
 * (§17.18.44), and the transitional `left`/`right` are defined as
 * "semantically equivalent to start/end" (ECMA-376 Part 4 §14.11.2) — so every
 * edge flips under an RTL base. An unset alignment defaults to the leading
 * (logical-start) edge.
 */
export function resolveAlignEdge(alignment: string | undefined, baseRtl: boolean): AlignEdge {
  switch (alignment) {
    case 'center':
      return 'center';
    case 'both':
    case 'justify':
    case 'distribute':
    // ECMA-376 §17.18.44: the three kashida settings (lowKashida / mediumKashida /
    // highKashida) and thaiDistribute are all forms of full justification between
    // both text margins — they differ only in HOW the extra space is distributed
    // (Arabic kashida elongation, Thai per-character spacing). The physical edge is
    // "justify" for all of them. NOTE: this is a MAPPING to the existing
    // inter-word/inter-character justification.
    // TODO(§17.18.44): true kashida (U+0640 tatweel) elongation — tracked in
    // https://github.com/yukiyokotani/office-open-xml-viewer/issues/724
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return 'justify';
    case 'end':
    case 'right':
      return baseRtl ? 'left' : 'right';
    case 'start':
    case 'left':
    case undefined:
    default:
      return baseRtl ? 'right' : 'left';
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value fully justifies each line by
 *  expanding inter-word (and, for distribute/thaiDistribute, inter-character)
 *  spacing. Covers `both` / `justify` / `distribute` plus the kashida and Thai
 *  variants, which this renderer maps onto the same slack-distribution kernel
 *  (see {@link resolveAlignEdge}; true kashida elongation is a follow-up). */
export function jcIsFullyJustified(alignment: string | undefined): boolean {
  switch (alignment) {
    case 'both':
    case 'justify':
    case 'distribute':
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return true;
    default:
      return false;
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value also stretches the paragraph's
 *  last line. `word-thai-distribute-cluster-policy` records that
 *  `thaiDistribute`, like `both`, leaves its final line ragged; only
 *  `distribute` stretches the last line. */
export function jcStretchesLastLine(alignment: string | undefined): boolean {
  return alignment === 'distribute';
}
