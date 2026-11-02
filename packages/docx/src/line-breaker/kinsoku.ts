import { popMixedLineItem, replaceLastMixedLineItem } from './mixed-space-fit.js';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { crossRunKinsokuRetract, graphemeClusterOffsets } from '@silurus/ooxml-core';
import { type LayoutImageSeg, type LayoutMathSeg, type LayoutTabSeg, type LayoutTextSeg } from './model.js';
import { createLineBreakerState } from './break-queue.js';
import { RESET_SLICED_TEXT_MEASUREMENT, protectedNoBreakOffsets, slicedTextMetadata } from './advance.js';
import { rebaseSeaBreaks } from './text-runs.js';


export type CrossRunKinsokuRetraction =
  | { readonly kind: 'none' }
  | { readonly kind: 'blocked' }
  | { readonly kind: 'retracted'; readonly tail: LayoutTextSeg };


/** Retract one legal suffix without cutting a protected source seam. */
export function retractLeadingKinsoku(
  breakerState: ReturnType<typeof createLineBreakerState>,
  kinsoku: KinsokuRules,
  materializeLatinSpaceCompression: () => void,
  strAdvance: (
    segment: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number,
  next: LayoutTextSeg,
  /** Layout scale for the line summary (WORD_COMPRESSED_SPACE_LINE_FIT). */
  scale: number,
): CrossRunKinsokuRetraction {
  const firstCp = next.text.codePointAt(0);
  const lastSeg = breakerState.currentLine[breakerState.currentLine.length - 1];
  if (
    firstCp === undefined ||
    !kinsoku.lineStartForbidden.has(firstCp) ||
    lastSeg === undefined ||
    !('text' in lastSeg)
  ) {
    return { kind: 'none' };
  }

  const lastText = lastSeg as LayoutTextSeg;
  const chars = [...lastText.text];
  const minKeep = breakerState.currentLine.length > 1 ? 0 : 1;
  const retractCount = crossRunKinsokuRetract(chars, kinsoku, minKeep);
  if (retractCount <= 0) return { kind: 'none' };

  const headText = chars.slice(0, chars.length - retractCount).join('');
  const split = headText.length;
  // Moving the whole segment would cut the hard source seam immediately
  // before it. A protected no-break edge is equally indivisible.
  if (
    (split === 0 && lastText.hardJoinPrev === true) ||
    protectedNoBreakOffsets(lastText).has(split)
  ) {
    return { kind: 'blocked' };
  }

  materializeLatinSpaceCompression();
  breakerState.latinLineHomogeneous = false;

  const tailText = lastText.text.slice(split);
  const tail: LayoutTextSeg = {
    ...lastText,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: tailText,
    ...slicedTextMetadata(lastText, split, lastText.text.length),
    measuredWidth: strAdvance(lastText, tailText, true),
    // The retraction creates a real line boundary. The old source seam was
    // either retained in `headText` or was soft; it must not be projected
    // onto this newly-created suffix.
    joinPrev: undefined,
    hardJoinPrev: undefined,
    src: {
      segIndex: lastText.src!.segIndex,
      charOffset: lastText.src!.charOffset + split,
    },
    seaBreaks: rebaseSeaBreaks(lastText.seaBreaks, split),
  };

  if (headText) {
    // The head keeps its first glyph and predecessor, so it keeps the
    // incoming WORD_KERNED_SPACE_PAIRS seam pair (bounded by its own advance).
    const headNaturalW = strAdvance(lastText, headText);
    const seamPair = Math.max(lastText.leadingWordBoundaryPx ?? 0, -headNaturalW);
    const headW = headNaturalW + seamPair;
    breakerState.currentWidth -= lastText.measuredWidth - headW;
    replaceLastMixedLineItem(breakerState, {
      ...lastText,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: headText,
      measuredWidth: headW,
      ...(seamPair === 0 ? {} : { leadingWordBoundaryPx: seamPair }),
      ...slicedTextMetadata(lastText, 0, split),
    }, scale);
  } else {
    breakerState.currentWidth -= lastText.measuredWidth;
    popMixedLineItem(breakerState);
  }
  return { kind: 'retracted', tail };
}


/** Keep a forbidden leader with its owner when retraction is blocked. */
export function keepLeadingKinsoku(
  breakerState: ReturnType<typeof createLineBreakerState>,
  strNaturalAdvance: (segment: LayoutTextSeg, text: string) => number,
  addToLine: (
    segment: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    width: number,
    height: number,
    ascent: number,
    descent: number,
  ) => void,
  queueEmergencyTail: (segment: LayoutTextSeg, split: number) => void,
  segment: LayoutTextSeg,
  h: number,
  asc: number,
  desc: number,
): boolean {
  const firstEnd = graphemeClusterOffsets(segment.text)[0] ?? segment.text.length;
  if (firstEnd <= 0) return false;
  const prefix = segment.text.slice(0, firstEnd);
  const prefixWidth = strNaturalAdvance(segment, prefix);
  addToLine(
    {
      ...segment,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: prefix,
      measuredWidth: prefixWidth,
      ...slicedTextMetadata(segment, 0, firstEnd),
    },
    prefixWidth,
    h,
    asc,
    desc,
  );
  if (firstEnd < segment.text.length) queueEmergencyTail(segment, firstEnd);
  return true;
}
