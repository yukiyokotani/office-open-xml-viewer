// Justified-text slack distribution for the docx renderer.
//
// ECMA-376 / ISO-29500 §17.18.44 ST_Jc defines the "fill the line" alignments:
//
//   both (Justified)  "This type of justification shall only affect the inter-
//                      word spacing on each line, and not the inter-character
//                      spacing within each word when justifying its contents."
//   distribute        "This type of justification shall equally affect the
//                      inter-word spacing on each line as well as the inter-
//                      character spacing between each word ... an equal amount of
//                      additional character pitch shall be added to all
//                      characters on the line."
//
// The `both` clause "inter-word, not inter-character within each word" is written
// for LATIN script, where a word is a run of letters delimited by spaces. CJK has
// no inter-word spaces and each ideograph behaves as its own word, so an
// inter-CJK boundary is an inter-word boundary, not an intra-word one.
// `word-cjk-both-inter-character-expansion` records the resulting compatibility
// behavior: `both` widens ASCII word gaps and inter-CJK boundaries, but not the
// boundary between two Latin letters in one word.
//
// `both` and `distribute` select the same gap opportunities and differ only in
// the last line: `both` leaves the paragraph's final line natural, `distribute`
// fills it too — a policy the RENDERER owns (it decides whether to call this for
// a given line). NOTE: per spec `distribute` should additionally add pitch
// between Latin letters within a word ("all characters on the line"). That
// pure-Latin refinement is explicitly unsupported pending independent evidence;
// CJK content, this kernel's target, is unaffected.
//
// The gap geometry itself lives in the shared kernel
// `@silurus/ooxml-core` → `distributeLineSlack` (packages/core/src/text/
// line-distribute.ts), which the pptx justifier shares too. This module is a thin
// docx adapter: it keeps the renderer's positional call shape and injects docx's
// whitespace predicate (U+0020 / U+3000, the core default) and the
// expansion-vs-compression gap rule (compression touches spaces only, never
// overlaps ideographs).
//
// Why a per-segment / per-character model: docx layout splits a paragraph into
// segments at spaces and style boundaries (splitTextForLayout), so a CJK phrase
// like "観察することで" is ONE segment with no internal spaces. Its inter-CJK
// gaps fall INSIDE the segment, between code points. The kernel walks the line's
// whole code-point stream and reports, per segment, where each gap falls:
// internal split points (to slice the glyph drawing) and a trailing-edge flag
// (the inter-segment boundary). The renderer applies `perGap` at each.

import {
  distributeLineSlack as distributeLineSlackCore,
  type DistributeResult,
} from '@silurus/ooxml-core';

// Re-export the shared shapes so the renderer keeps importing them from here.
export type {
  DistributeSeg,
  DistributeResult,
  SegStretch,
} from '@silurus/ooxml-core';

/**
 * Distribute `slack` px across a justified line's gap opportunities (inter-word
 * spaces AND, for expansion, inter-CJK boundaries), per ECMA-376 §17.18.44.
 *
 * Thin positional adapter over the shared `distributeLineSlack` kernel; see
 * {@link DistributeResult} and the core docs for the gap model. docx's
 * whitespace set is U+0020 / U+3000 — exactly the kernel default — so no
 * whitespace predicate is injected.
 *
 * @param segments       The line's segments in LOGICAL (reading) order.
 * @param slack          availWidth - naturalWidth, px. >0 stretches; <0 compresses.
 * @param firstContentSi Index of the first segment holding non-whitespace content;
 *                       earlier (leading-indent / 字下げ) whitespace segments are
 *                       fixed. Pass 0 to disable the skip (e.g. under bidi).
 * @param lastDrawnSi    Index of the visually-final segment; it and the boundary
 *                       into it get no gap (EXACT match, not `>=`, for bidi).
 * @param minPerGap      Lower bound on a (negative) `perGap` when compressing.
 *                       Ignored when slack >= 0.
 * @param includeCJK     When true (default), inter-CJK boundaries open gaps too
 *                       (EXPANSION). When false, only inter-word spaces open gaps
 *                       — used for COMPRESSION (negative slack), where shrinking a
 *                       space is legitimate but overlapping two ideographs is not,
 *                       so the renderer passes includeCJK = slack > 0.
 * @param seaClusterGaps When true (jc=thaiDistribute, §17.18.44 "Thai Language
 *                       Justification"), a gap ALSO opens at every UAX#29
 *                       grapheme-cluster boundary interior to a Southeast-Asian
 *                       span (Thai/Lao/Khmer) whose both sides are SEA, so a
 *                       space-free Thai line is justified by widening inter-cluster
 *                       gaps (a combining mark stays glued to its base). Off for
 *                       `both`/`distribute` per
 *                       `word-thai-distribute-cluster-policy`.
 * @returns The distribution, or `null` when nothing stretches.
 */
export function distributeLineSlack(
  segments: readonly { text?: string }[],
  slack: number,
  firstContentSi: number,
  lastDrawnSi: number,
  minPerGap = -Infinity,
  includeCJK = true,
  seaClusterGaps = false,
): DistributeResult | null {
  return distributeLineSlackCore(segments, slack, {
    firstContentSi,
    lastDrawnSi,
    minPerGap,
    seaClusterGaps,
    // Compression (includeCJK=false) opens spaces only; never widen an inter-CJK
    // boundary, which would overlap ideographs. The kernel default isGapChar
    // (core.isCjkBreakChar) handles expansion.
    ...(includeCJK ? {} : { isGapChar: () => false }),
  });
}

/** The px total actually applied by a {@link DistributeResult} = perGap summed
 *  over every gap it opened (interior splits + inter-segment trailing gaps).
 *  Negative for compression. The drawn line width is `Σ measuredWidth + this`. */
export function distributedDelta(dist: DistributeResult | null): number {
  if (!dist) return 0;
  let gaps = 0;
  for (const s of dist.perSeg.values()) gaps += s.splitBefore.length + (s.trailingGap ? 1 : 0);
  return dist.perGap * gaps;
}
