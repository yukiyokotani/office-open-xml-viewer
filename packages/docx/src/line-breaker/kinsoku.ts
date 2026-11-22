import { advanceLineBoundary } from './source-cursor.js';
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
  | {
    readonly kind: 'retracted';
    readonly tail: LayoutTextSeg;
    /** A retracted ruby cell's later pieces, in order after `tail`. */
    readonly followers?: readonly LayoutTextSeg[];
  };

/** Queue a retraction ahead of the segment that caused it. */
export function queueRetraction(
  queue: { unshift(segment: LayoutTextSeg): void },
  retraction: Extract<CrossRunKinsokuRetraction, { kind: 'retracted' }>,
): void {
  for (const follower of [...(retraction.followers ?? [])].reverse()) queue.unshift(follower);
  queue.unshift(retraction.tail);
}


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
  if (lastText.rubyGroup !== undefined) {
    // An opener left at the line end by the retracted cell moves with it;
    // when that opener would empty the line, the cell stays (blocked) and
    // the forbidden leader is kept with it instead.
    const line = breakerState.currentLine;
    let start = line.length - 1;
    while (start > 0 && 'text' in line[start - 1]! && (line[start - 1] as LayoutTextSeg).rubyGroup === lastText.rubyGroup) {
      start -= 1;
    }
    const before = start > 0 ? line[start - 1] : undefined;
    if (before && 'text' in before && before.text.length > 0
      && [...before.text].every((char) => kinsoku.lineEndForbidden.has(char.codePointAt(0)!))
      && start - 1 === 0) {
      return { kind: 'blocked' };
    }
    const cell = retractRubyCell(breakerState, materializeLatinSpaceCompression);
    if (cell.kind !== 'retracted') return cell;
    const opener = retractLineEndForbidden(breakerState, kinsoku, materializeLatinSpaceCompression, strAdvance, scale);
    return opener.kind === 'retracted'
      ? { kind: 'retracted', tail: opener.tail,
        followers: [...(opener.followers ?? []), cell.tail, ...(cell.followers ?? [])] }
      : cell;
  }
  const chars = [...lastText.text];
  const minKeep = breakerState.currentLine.length > 1 ? 0 : 1;
  const retractCount = crossRunKinsokuRetract(chars, kinsoku, minKeep);
  if (retractCount <= 0) return { kind: 'none' };
  return retractSuffix(breakerState, materializeLatinSpaceCompression, strAdvance, lastText,
    chars.slice(0, chars.length - retractCount).join('').length, scale);
}

/** §17.3.1.16: a line must not end with a line-end-forbidden character
 * (an opening bracket). When the unit after it moves to the next line, the
 * forbidden suffix of the line's last segment, or that whole ruby cell,
 * moves with it. */
export function retractLineEndForbidden(
  breakerState: ReturnType<typeof createLineBreakerState>,
  kinsoku: KinsokuRules,
  materializeLatinSpaceCompression: () => void,
  strAdvance: (segment: LayoutTextSeg, text: string, retainTrailingPunctuationCompression?: boolean) => number,
  scale: number,
): CrossRunKinsokuRetraction {
  const lastSeg = breakerState.currentLine.at(-1);
  if (!kinsoku.enabled || lastSeg === undefined || !('text' in lastSeg)) return { kind: 'none' };
  const lastText = lastSeg as LayoutTextSeg;
  const chars = [...lastText.text];
  let count = 0;
  while (count < chars.length && kinsoku.lineEndForbidden.has(chars[chars.length - 1 - count]!.codePointAt(0)!)) {
    count += 1;
  }
  if (count === 0) return { kind: 'none' };
  if (lastText.rubyGroup !== undefined) return retractRubyCell(breakerState, materializeLatinSpaceCompression);
  if (count === chars.length && breakerState.currentLine.length === 1) return { kind: 'blocked' };
  return retractSuffix(breakerState, materializeLatinSpaceCompression, strAdvance, lastText,
    chars.slice(0, chars.length - count).join('').length, scale);
}

/** WORD_RUBY_ALIGN_PLACEMENT: a ruby cell is indivisible. It moves whole,
 * every piece of it, unless it alone holds the line or is glued to what
 * precedes it; the line returns to its state before the cell was admitted
 * (its extents included), which placeRubyCell recorded. */
function retractRubyCell(
  breakerState: ReturnType<typeof createLineBreakerState>,
  materializeLatinSpaceCompression: () => void,
): CrossRunKinsokuRetraction {
  const line = breakerState.currentLine;
  const group = (line.at(-1) as LayoutTextSeg).rubyGroup;
  let start = line.length - 1;
  while (start > 0) {
    const before = line[start - 1]!;
    if (!('text' in before) || before.rubyGroup !== group) break;
    start -= 1;
  }
  const pieces = line.slice(start) as LayoutTextSeg[];
  if (start === 0 || pieces[0]!.hardJoinPrev === true) return { kind: 'blocked' };
  materializeLatinSpaceCompression();
  for (let index = pieces.length - 1; index >= 0; index -= 1) {
    breakerState.currentWidth -= pieces[index]!.measuredWidth;
    popMixedLineItem(breakerState);
  }
  const admitted = breakerState.rubyCellStart;
  breakerState.latinLineHomogeneous = false;
  if (!admitted || admitted.group !== group || admitted.line !== line || admitted.lineLength !== start) {
    return { kind: 'retracted', tail: pieces[0]!, followers: pieces.slice(1) };
  }
  const mixedSpace = breakerState.mixedSpace;
  Object.assign(breakerState, admitted.scalars);
  breakerState.mixedSpace = mixedSpace;
  breakerState.snapBlock = admitted.snapBlock ? { ...admitted.snapBlock } : null;
  breakerState.latinLineGaps.length = admitted.latinLineGapCount;
  breakerState.latinLineHomogeneous = false;
  breakerState.rubyCellStart = null;
  return { kind: 'retracted', tail: pieces[0]!, followers: pieces.slice(1) };
}

/** Retract `lastText` from `split` on: the head stays, the suffix leads the
 * next line. */
function retractSuffix(
  breakerState: ReturnType<typeof createLineBreakerState>,
  materializeLatinSpaceCompression: () => void,
  strAdvance: (segment: LayoutTextSeg, text: string, retainTrailingPunctuationCompression?: boolean) => number,
  lastText: LayoutTextSeg,
  split: number,
  scale: number,
): CrossRunKinsokuRetraction {
  const headText = lastText.text.slice(0, split);
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
  const tailSlice: LayoutTextSeg = {
    ...lastText,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: tailText,
    ...slicedTextMetadata(lastText, split, lastText.text.length),
  };
  const tail: LayoutTextSeg = {
    ...tailSlice,
    measuredWidth: strAdvance(tailSlice, tailText, true),
    // The retraction creates a real line boundary. The old source seam was
    // either retained in `headText` or was soft; it must not be projected
    // onto this newly-created suffix.
    joinPrev: undefined,
    hardJoinPrev: undefined,
    src: advanceLineBoundary(lastText.src!, split),
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
