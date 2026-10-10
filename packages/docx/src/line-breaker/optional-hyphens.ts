import type { BreakOpportunityIteratorContext } from './break-opportunities.js';
import type { LayoutTextSeg } from './model.js';
import { RESET_SLICED_TEXT_MEASUREMENT, slicedTextMetadata } from './advance.js';
import type { KinsokuRules } from '@silurus/ooxml-core';

/** Intrinsic minima and actual discretionary placement consume the same
 * eligibility. A marker cannot break an authored hard seam, fixed cell,
 * empty prefix, end of word, or forbidden continuation line start. */
export function optionalHyphenBreakAllowed(
  member: LayoutTextSeg, offset: number, next: LayoutTextSeg | undefined,
  visiblePrefix: boolean, kinsoku: KinsokuRules,
): boolean {
  const hasContinuation = offset < member.text.length;
  const leading = hasContinuation ? member.text.codePointAt(offset) : next?.text.codePointAt(0);
  return visiblePrefix && leading !== undefined
    && !/^\s/u.test(String.fromCodePoint(leading))
    && !member.hardJoinPrev && (hasContinuation || !next?.hardJoinPrev)
    && member.fitTextRegionIndex === undefined
    && (!kinsoku.enabled || !kinsoku.lineStartForbidden.has(leading));
}

/** ECMA-376 §17.3.3.29: authored discretionary opportunities are independent
 * of dictionary hyphenation. Fit the word without any conditional glyph first;
 * on overflow choose the last fitting marker, including that marker's own
 * glyph advance. Source/font seams remain joined, and the exact measured glyph
 * becomes retained paint. All lookahead is charged to the pass's work budget. */
export function placeOptionalHyphenPrefix(
  context: BreakOpportunityIteratorContext,
  first: LayoutTextSeg,
): boolean {
  if (!first.optionalHyphenWord || first.joinPrev && context.breakerState.currentLine.length > 0) return false;
  context.reservePrefixWork(first.text.length + 1);
  const members = [first];
  for (const next of context.breakerState.queue) {
    if (!('text' in next) || !next.joinPrev) break;
    context.reservePrefixWork(next.text.length + 1);
    members.push(next);
  }
  let width = context.breakerState.currentWidth;
  let candidate: { index: number; offset: number; glyph: LayoutTextSeg } | undefined;
  let visiblePrefix = false;
  for (const [index, member] of members.entries()) {
    const markers = member.optionalHyphen
      ? [{ offset: 0, glyph: member.optionalHyphen }]
      : member.optionalHyphenBreaks ?? [];
    for (const marker of markers) {
      // A marker at a word's end has no continuation to break. Fixed fitText
      // cells and authored no-break ownership keep their ordinary atomicity.
      // Prefixes can be revisited quadratically in a dense authored word.
      // Reserve before allocating or shaping, including cache hits and
      // zero-offset marker probes. Inspect the continuation without copying it.
      context.reservePrefixWork(marker.offset + 1);
      const prefix = member.text.slice(0, marker.offset);
      const next = members[index + 1];
      if (!optionalHyphenBreakAllowed(member, marker.offset, next,
        visiblePrefix || prefix.trim().length > 0, context.kinsoku)) continue;
      const glyph: LayoutTextSeg = { ...marker.glyph, ...RESET_SLICED_TEXT_MEASUREMENT,
        src: member.src ? { ...member.src, charOffset: member.src.charOffset + marker.offset } : undefined,
        optionalHyphenGlyph: true, joinPrev: true };
      const prefixWidth = marker.offset > 0 ? context.strAdvance(member, prefix) : 0;
      if (context.fitsMeasuredWidth(width + prefixWidth + context.segAdvance(glyph), context.availW())) {
        candidate = { index, offset: marker.offset, glyph };
      }
    }
    if (!member.optionalHyphen) {
      width += context.segAdvance(member);
      visiblePrefix ||= member.text.trim().length > 0;
    }
  }
  const last = members.at(-1)!;
  // A collapsible trailing space does not force an otherwise fitting word
  // to select an optional hyphen. RTL keeps its ordinary containment policy.
  if (!context.baseRtl && last.text.endsWith(' ')) {
    context.reservePrefixWork(last.text.length);
    width -= context.segAdvance(last) - context.strAdvance(last, last.text.replace(/ +$/u, ''));
  }
  if (context.fitsMeasuredWidth(width, context.availW())) return false;
  // An unselected owner cannot bypass the existing registered-face space
  // fit. Its conditional glyph is absent from this complete-word request.
  // Real visible font seams retain that policy's mixed-face exclusion.
  if (members.every(member => member.optionalHyphen || context.sameLatinSpaceFace(member, first))
    && context.fitHomogeneousLatinSpaces(first, width - context.breakerState.currentWidth)) return false;
  if (!candidate) {
    // Move the complete word to the next band before resorting to the
    // ordinary overlong-word policy. A too-wide conditional glyph must not
    // overflow the current line or create an empty hyphen-only line.
    if (context.breakerState.currentLine.length > 0 && !first.joinPrev) {
      context.flush(undefined, false, first.src);
      context.breakerState.queue.unshift(first);
      return true;
    }
    return false;
  }
  for (let index = 0; index <= candidate.index; index += 1) {
    if (index > 0) context.breakerState.queue.shift();
    const source = members[index]!;
    let member = source;
    if (index === candidate.index) {
      context.reservePrefixWork(source.text.length);
      const tail = source.text.slice(candidate.offset);
      if (tail) context.breakerState.queue.unshift({ ...source, ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail, measuredWidth: 0, joinPrev: undefined, hardJoinPrev: undefined,
        ...slicedTextMetadata(source, candidate.offset, source.text.length),
        src: source.src ? { ...source.src, charOffset: source.src.charOffset + candidate.offset } : undefined });
      if (candidate.offset === 0) continue;
      member = { ...source, ...RESET_SLICED_TEXT_MEASUREMENT,
        text: source.text.slice(0, candidate.offset), measuredWidth: 0,
        ...slicedTextMetadata(source, 0, candidate.offset) };
    }
    if (member.optionalHyphen) {
      context.addToLine(member, 0, 0, 0, 0);
      continue;
    }
    const box = context.textSegmentBox(member);
    member.measuredWidth = box.width;
    context.addToLine(member, box.width, box.height, box.ascent, box.descent);
  }
  // The chosen boundary consumed the seam before the retained continuation,
  // including standalone and sequence-end markers with no sliced text tail.
  // Its next line starts with its own ordinary source/atomicity ownership.
  const continuation = context.breakerState.queue.peek();
  if (continuation && 'text' in continuation && (continuation.joinPrev || continuation.hardJoinPrev)) {
    context.breakerState.queue.shift();
    context.breakerState.queue.unshift({ ...continuation, joinPrev: undefined, hardJoinPrev: undefined });
  }
  const glyphBox = context.textSegmentBox(candidate.glyph);
  candidate.glyph.measuredWidth = glyphBox.width;
  context.addToLine(candidate.glyph, glyphBox.width, glyphBox.height, glyphBox.ascent, glyphBox.descent);
  context.flush(undefined, false, context.breakerState.queue.peek()?.src);
  return true;
}
