import { measureFitTextUnit, measureJoinedTextUnit } from './atomic-units.js';
import { placeOptionalHyphenPrefix } from './optional-hyphens.js';
import {
  kinsokuAdjustedSplit,
  isGraphemeFillText,
  isDictionarySeaText,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
} from '@silurus/ooxml-core';
import { EAST_ASIAN_RE } from '../layout/text.js';
import {
  wordIsOverflowPunctuation,
  wordIdeographicSpaceLineEndAllowanceCount,
  wordVisiblePrefixFitWidthPx,
} from '../layout/line-compatibility.js';
import {
  type LayoutImageSeg,
  type LayoutMathSeg,
  type LayoutSeg,
  type LayoutTabSeg,
  type LayoutTextSeg,
  type LineBoundary,
} from './model.js';
import {
  RESET_SLICED_TEXT_MEASUREMENT,
  charScaleFactor,
  charSpacingDeltaPx,
  legalTextSplitAtOrBefore,
  segAdvanceWidth,
  segmentCharacterGridDeltaPx,
  slicedTextMetadata,
  snapToCharsClass,
} from './advance.js';
import { createBidiTabCellResolver, bidiTabFrame, nextLineTabStop, positionalTabTarget, tabAlignmentRole } from './tabs.js';
import { wordPositionalTabReferenceBox } from '../layout/line-compatibility.js';
import { buildFont } from './font-routes.js';
import {
  mixedCandidateMayShrink,
  mixedLineHasVisibleText,
  noteMixedSpaceWork,
  type MixedSpaceCandidate,
} from './mixed-space-fit.js';
import {
  extendThroughTrailingIdeographicSpaces,
  hasCJKBreakOpportunity,
  rebaseSeaBreaks,
} from './text-runs.js';
import { fitCJKPrefix, hasEastAsianVisiblePredecessor } from './fit-search.js';
import { LineGapRejection, fitJustifiedCompression, justifiedCompressionApplies, type PassOperationState } from './pass-operations.js';

/** The iterator reads only its declared slice of the explicit pass state. */
export type BreakOpportunityIteratorContext = Pick<
  PassOperationState,
  | 'reservePrefixWork'
  | 'breakerState'
  | 'flush'
  | 'baseRtl'
  | 'addToLine'
  | 'scale'
  | 'firstIndent'
  | 'tabOriginPx'
  | 'maxWidth'
  | 'marginRightPx'
  | 'tabFollowWidth'
  | 'measureText'
  | 'verticalInkExtra'
  | 'characterGrid'
  | 'tabStops'
  | 'defaultTabPt'
  | 'tabFollowingMetrics'
  | 'bidiCustomStopsPx'
  | 'bidiIntervalPx'
  | 'decimalAlignmentPoint'
  | 'availW'
  | 'setMeasureFont'
  | 'fontFamilyClasses'
  | 'measurement'
  | 'textSegmentBox'
  | 'prospectiveSnapAdvance'
  | 'segAdvance'
  | 'strAdvance'
  | 'justifiedCompression'
  | 'isJustified'
  | 'stretchLastLine'
  | 'widthPolicy'
  | 'overflowPunct'
  | 'sameLatinSpaceFace'
  | 'fitsMeasuredWidth'
  | 'fitHomogeneousLatinSpaces'
  | 'mixedSpaceRequirement'
  | 'markMixedSpacesCompressed'
  | 'appendQueuedIdeographicSpaceSegment'
  | 'emergencyTextSplit'
  | 'effectiveFontPx'
  | 'ctx'
  | 'verticalGlyphMeasurement'
  | 'kinsoku'
  | 'strNaturalAdvance'
  | 'retractCurrentLineForLeadingKinsoku'
  | 'keepLeadingKinsokuWithCurrentLine'
  | 'explicitTextSplit'
  | 'queueEmergencyTail'
  | 'forcedPlacement'
  | 'minimalLegalTextWidth'
>;

/** Transaction hooks for WORD_FLOAT_GAP_FLOW admission (#1670). */
export interface GapTransactionHooks {
  readonly captureGapSnapshot: () => void;
  readonly rejectGap: (rejection: LineGapRejection) => void;
}

function sameBoundary(left: LineBoundary | undefined, right: LineBoundary): boolean {
  return left !== undefined && left.segIndex === right.segIndex && left.charOffset === right.charOffset;
}

/** Index in the current line where the joined unit ending at its last item
 * starts: source seams marked `joinPrev` permit no line boundary. */
function joinedUnitStart(line: readonly LayoutSeg[]): number {
  let index = line.length - 1;
  while (index > 0) {
    const item = line[index]!;
    if (!('text' in item) || !item.joinPrev) break;
    index -= 1;
  }
  return Math.max(0, index);
}

/** The single rollback predicate of a narrowed gap: the head unit's placed
 * advance (from the line start, joined followers included, collapsible
 * U+0020 suffix excluded outside RTL) exceeds the fragment. */
function placedAdvanceExceeds(
  context: BreakOpportunityIteratorContext,
  segment: LayoutTextSeg,
  placedText: string,
): boolean {
  const { breakerState, strAdvance, availW, fitsMeasuredWidth, baseRtl } = context;
  // A line-end U+3000 run hangs as whitespace (word-ideographic-space hang).
  const visible = (baseRtl ? placedText : placedText.replace(/ +$/u, '')).replace(/\u3000+$/u, '');
  return !fitsMeasuredWidth(breakerState.currentWidth + strAdvance(segment, visible), availW());
}

function lineAdvanceFrom(line: readonly LayoutSeg[], start: number): number {
  let width = 0;
  for (let index = start; index < line.length; index += 1) {
    width += (line[index] as { measuredWidth?: number }).measuredWidth ?? 0;
  }
  return width;
}

/** Consume one prepared queue in source order, applying all legal break paths.
 * A narrowed float-gap fragment is a transaction: a forced placement inside it
 * rolls the fragment back (`rejectGap`) and processing resumes from the
 * restored queue in the next admissible window. */
export function iterateBreakOpportunities(
  context: BreakOpportunityIteratorContext,
  hooks?: GapTransactionHooks,
): void {
  const { breakerState, flush } = context;
  const resolveCompletedTabCells = createBidiTabCellResolver();
  while (breakerState.queue.length > 0) {
    const transaction = breakerState.gapTransaction;
    if (transaction?.narrowed) {
      const head = breakerState.queue.peek()!;
      if (transaction.stopBefore && breakerState.currentLine.length > 0
        && sameBoundary(head.src, transaction.stopBefore)) {
        // Replay of a rejected fragment ends before its forced unit.
        flush(undefined, false, head.src);
        continue;
      }
      if (breakerState.currentLine.length === 0) hooks?.captureGapSnapshot();
    }
    const seg = breakerState.queue.shift()!;
    breakerState.inHand = seg;
    try {
      processQueuedSegment(context, seg, resolveCompletedTabCells);
    } catch (error) {
      if (!(error instanceof LineGapRejection) || !hooks) throw error;
      hooks.rejectGap(error);
    }
    breakerState.inHand = undefined;
  }
}

function processQueuedSegment(
  context: BreakOpportunityIteratorContext,
  seg: LayoutSeg,
  resolveCompletedTabCells: ReturnType<typeof createBidiTabCellResolver>,
): void {
  const { breakerState, flush } = context;
  // ── Line-break sentinel ──────────────────────────────
  if ('lineBreak' in seg) {
    // The line being flushed ends at a MANUAL break (§17.3.3.1) — mark it so a
    // justified paragraph left-aligns it like its final line (§17.18.44).
    flush(seg.fontSize, true);
    breakerState.trailingBreakFontSize = seg.fontSize;
    return;
  }
  breakerState.trailingBreakFontSize = null;

  // ── Tab segment ──────────────────────────────────────
  if ('isTab' in seg) {
    processTabSegment(context, seg, resolveCompletedTabCells);
    return;
  }

  // ── Image segment ────────────────────────────────────
  if ('imagePath' in seg) {
    processImageSegment(context, seg);
    return;
  }

  // ── Math segment ─────────────────────────────────────
  if ('math' in seg) {
    processMathSegment(context, seg);
    return;
  }

  // ── Text segment ─────────────────────────────────────
  if ('text' in seg) {
    // A fixed cell returns before ordinary text admission. An authored
    // marker after that complete cell still owns its external word seam;
    // the selector excludes every opportunity inside the cell itself.
    if (seg.fitTextRegionIndex !== undefined && placeOptionalHyphenPrefix(context, seg)) return;
    if (seg.optionalHyphen) {
      context.addToLine(seg, 0, 0, 0, 0);
      return;
    }
  }
  processTextSegment(context, seg as LayoutTextSeg);
}

function processTextSegment(context: BreakOpportunityIteratorContext, seg: LayoutTextSeg): void {
  const {
    breakerState,
    flush,
    addToLine,
    maxWidth,
    characterGrid,
    availW,
    textSegmentBox,
    prospectiveSnapAdvance,
    segAdvance,
    strAdvance,
    isJustified,
    stretchLastLine,
    overflowPunct,
    sameLatinSpaceFace,
  } = context;
  const s = seg as LayoutTextSeg;
  const segmentBox = textSegmentBox(s);
  const w = segmentBox.width;
  const prospectiveWidth = prospectiveSnapAdvance(s, w);
  const h = segmentBox.height;
  const asc = segmentBox.ascent;
  const desc = segmentBox.descent;
  const paragraphFinalIdeographicSpaceTail = s.paragraphFinalIdeographicSpaceTail === true;
  const paragraphFinalIdeographicSpaceCount = s.paragraphFinalIdeographicSpaceCount ?? 0;
  const paragraphFinalIdeographicSpaceLocalCount = s.paragraphFinalIdeographicSpaceLocalCount ?? 0;
  const visibleBeforeParagraphFinalTail = paragraphFinalIdeographicSpaceTail
    ? s.text.slice(0, Math.max(0, s.text.length - paragraphFinalIdeographicSpaceLocalCount))
    : s.text;
  if (
    paragraphFinalIdeographicSpaceTail &&
    paragraphFinalIdeographicSpaceCount > 1 &&
    visibleBeforeParagraphFinalTail.length > 0
  ) {
    const visibleSegment: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: visibleBeforeParagraphFinalTail,
      paragraphFinalIdeographicSpaceTail: undefined,
      paragraphFinalIdeographicSpaceLocalCount: undefined,
      paragraphFinalIdeographicSpaceCount: undefined,
      paragraphFinalIdeographicSpaceTailStart: undefined,
      measuredWidth: 0,
      ...slicedTextMetadata(s, 0, visibleBeforeParagraphFinalTail.length),
    };
    const trailingSegment: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: s.text.slice(visibleBeforeParagraphFinalTail.length),
      paragraphFinalIdeographicSpaceLocalCount,
      joinPrev: undefined,
      hardJoinPrev: undefined,
      paragraphFinalIdeographicSpaceTailStart: true,
      measuredWidth: 0,
      ...slicedTextMetadata(s, visibleBeforeParagraphFinalTail.length, s.text.length),
      src: s.src
        ? {
            segIndex: s.src.segIndex,
            charOffset: s.src.charOffset + visibleBeforeParagraphFinalTail.length,
          }
        : undefined,
    };
    breakerState.queue.unshift(trailingSegment);
    breakerState.queue.unshift(visibleSegment);
    return;
  }
  if (
    paragraphFinalIdeographicSpaceTail &&
    /^\u3000+$/u.test(s.text) &&
    s.paragraphFinalIdeographicSpaceTailStart === true
  ) {
    const currentLineHasVisibleText = breakerState.currentLine.some(
      (candidate) => 'text' in candidate && /[^\u3000]/u.test(candidate.text),
    );
    if (currentLineHasVisibleText) {
      let trailingTailWidth = w;
      for (const candidate of breakerState.queue) {
        if (!('text' in candidate) || candidate.paragraphFinalIdeographicSpaceTail !== true) break;
        trailingTailWidth += segAdvance(candidate);
      }
      if (breakerState.currentWidth + trailingTailWidth > availW()) {
        flush(undefined, false, s.src);
        breakerState.queue.unshift(s);
        return;
      }
    }
  }

  // ECMA-376 §17.3.2.14: a fit region is an atomic fixed-width cell. The
  // first segment judges the WHOLE resolved region; after an optional flush,
  // every member is added without entering the CJK/overlong-word split paths.
  // This also handles a target wider than the line: it overflows as one unit
  // instead of violating the required internal non-wrap boundary.
  if (s.fitTextRegionIndex !== undefined) {
    if (s.fitTextRegionStart) {
      const regionWidth = measureFitTextUnit(s, breakerState.queue, context);
      if (
        breakerState.currentLine.length > 0 &&
        breakerState.currentWidth + regionWidth > availW()
      ) {
        flush(undefined, false, s.src);
      }
      // The region is one fixed cell: it cannot be forced into a narrower gap.
      if (breakerState.currentLine.length === 0 && !context.fitsMeasuredWidth(regionWidth, availW())) {
        context.forcedPlacement(regionWidth);
      }
    }
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    return;
  }
  // A terminal separator may collapse when this word becomes line-final;
  // visible glyphs still need to fit at their natural measured advance.
  const trimmed = s.text.replace(/ +$/, '');
  // Subtract the full-model advance of the trimmed text (not the natural width)
  // so the grid delta, w:w scale and w:spacing pitch on the retained glyphs all
  // cancel and trailingSpaceW is the bare trailing-space advance — keeping `w`
  // and `wForFit` on the one advance model (`strAdvance` == the model behind `w`).
  const trailingSpaceW = snapToCharsClass(s, characterGrid)
    ? 0
    : s.text.endsWith(' ')
      ? w - strAdvance(s, trimmed)
      : 0;
  s.latinNaturalTrailingSpacePx =
    trailingSpaceW > 0 &&
    // One U+0020 after a word. A separator emitted as its own run is outside
    // both observed classes and makes the line inhomogeneous.
    s.latinSpaceCompressionEligible === true &&
    /^[^ ]+ $/u.test(s.text)
      ? trailingSpaceW
      : undefined;
  s.latinSpaceCompressionPx = undefined;
  // WORD_COMPRESSED_SPACE_LINE_FIT: the trailing U+0020 of a word, or a
  // space-only run after visible content, are shrinkable spaces of a mixed line.
  const mixedSpaces =
    s.mixedSpaceAverageWidthRatio !== undefined &&
    trailingSpaceW > 0 &&
    !trimmed.includes(' ') &&
    (trimmed.length > 0 ||
      // O(1) from the reversible line summary (no per-segment line scan).
      mixedLineHasVisibleText(breakerState));
  s.mixedNaturalTrailingSpacePx = mixedSpaces ? trailingSpaceW : undefined;
  s.mixedNaturalTrailingSpaceCount = mixedSpaces ? s.text.length - trimmed.length : undefined;
  const interwordCompression = justifiedCompressionApplies(context);
  // Library containment policy: an RTL line is anchored at its right edge,
  // so even an invisible trailing-space advance shifts visible LTR cells left.
  // Count that advance during fitting instead of admitting glyphs past the band.
  const fitWidthFor = (
    widthPx: number,
    trailingSpacePx: number,
  ): number => wordVisiblePrefixFitWidthPx(widthPx, trailingSpacePx, context.baseRtl);
  const wForFit = fitWidthFor(prospectiveWidth, trailingSpaceW);
  // ECMA-376 §17.3.1.33 does not prescribe a line-breaking tolerance.
  // Word-for-Mac controls with Calibri and Arial, left/center/right aligned
  // 10pt table cells, wrap a trailing Latin word below its natural advance
  // boundary (including <1pt overflow). Times New Roman differs by a
  // sub-point at that boundary, so this is a conservative library fit policy,
  // not a claim that every Office face and script has identical break points.
  // An earlier global 25%-of-spaces allowance pulled words up even when Word
  // did not; it also had no proven bound matching paint compression.
  // Dictionary-SEA candidate (Thai/Lao/Khmer; grapheme-fill Myanmar/Tibetan
  // stays on its per-cluster greedy path). Per-codepoint scan: a rare segment
  // mixing both SEA families is not dictionary-SEA, so
  // it keeps the pre-#991 greedy path instead of moving a grapheme-fill span
  // inside an atomic chunk.
  const sDictSea = s.seaBreaks !== undefined && isDictionarySeaText(s.text);

  // Atomic glued group: when THIS segment starts a glued group (its followers
  // in the queue are `joinPrev` pieces — small-caps case-pieces of the SAME
  // word like "I" then "NTRODUCTION", or a UAX#14 LB13 non-starter authored in
  // its own run like a trailing "," / "。"), the per-segment wrap below would
  // let the group split across lines. Pre-measure it and, if it does not fit on
  // the current (non-empty) line, flush so it starts fresh.
  //
  // ONLY when the lead segment is NOT itself CJK-breakable. A glued group whose
  // lead is a CJK run (e.g. "…通過する" + "。") is NOT atomic: the run splits at
  // an inter-CJK boundary and the trailing non-starter stays on its LAST piece
  // (§17.3.1.16 kinsoku keeps it off the next line's head when enabled — the
  // default; with kinsoku off it may lead the line, as it did before PR #602).
  // Pre-flushing the whole run instead leaves the prior line far short, which a
  // `both` line then stretches wide in a justified paragraph. `joinPrev`
  // stays a pure "this is a non-starter" marker; the atomic-vs-breakable decision lives here. A
  // non-breakable Latin / small-caps lead is genuinely atomic, so the pre-flush
  // (and the over-long-word char-break path below) still applies there.
  // Complete-unit justified admission precedes the ordinary natural-only
  // joined-unit preflight, so a style/source seam cannot bypass compression.
  s.measuredWidth = w;
  if (!prefersWholeWordAtScriptBoundary(context, s) && fitJustifiedCompression(context, s)) return;
  // The authored opportunity follows the existing whole-unit justified fit;
  // it must not introduce a wrap before a word admitted by that policy.
  if (placeOptionalHyphenPrefix(context, s)) return;
  if (prepareAtomicTextFit(context, { s, w, trailingSpaceW, sDictSea, fitWidthFor })) return;

  // §17.3.1.21 permits one eligible punctuation character past the text
  // extent. The isolated compatibility predicate owns both the CJK-language
  // sets and the bounded parent-run extensions owned by
  // `wordIsOverflowPunctuation`. CJK
  // segments that need an internal split retain their separate
  // overflowPunct-vs-kinsoku rule.
  const visibleSegmentScalars = [...trimmed];
  const trailingOverflowCharacter = visibleSegmentScalars.at(-1);
  const textBeforeTrailingOverflow = visibleSegmentScalars.slice(0, -1).join('');
  // §17.3.1.21 permits paragraph-edge hanging, not intrusion into a
  // DrawingML exclusion (§20.4.2.17). A narrowed float gap owns its full ink.
  const admitsTrailingOverflowPunctuation =
    overflowPunct &&
    breakerState.gapTransaction?.endsAtExclusion !== true &&
    trailingOverflowCharacter !== undefined &&
    (breakerState.currentLine.length > 0 || textBeforeTrailingOverflow.length > 0) &&
    wordIsOverflowPunctuation(
      trailingOverflowCharacter,
      s.eastAsiaLanguage,
      s.overflowPunctuationEastAsianRun === true,
      s.script === 'ascii' || s.script === 'highAnsi',
      s.script === 'complexScript',
      s.overflowPunctuationBidiLanguage,
    ) &&
    breakerState.currentWidth + strAdvance(s, textBeforeTrailingOverflow) <= availW();

  // WORD_COMPRESSED_SPACE_LINE_FIT: once a mixed line's spaces are shrunk, a
  // following unit (such as a closing mark split into its own source run) is
  // judged by the same reduction rule; joined and split runs must agree.
  if (breakerState.mixedSpace.compressed && mixedCandidateMayShrink(breakerState, s.text)) {
    const required = context.mixedSpaceRequirement(mixedSeamCandidate(context, s, s.text, wForFit));
    if (required !== undefined) {
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
      context.appendQueuedIdeographicSpaceSegment(s);
      return;
    }
  }

  // A line already admitted using this homogeneous-face rule cannot lend
  // that prior compression to a later mixed-face candidate. Its allocation
  // is finalized here, and the new route starts a fresh line.
  if (
    breakerState.latinAppliedPerGap > 0 &&
    (!breakerState.latinLineHomogeneous ||
      !breakerState.latinLineFace ||
      !sameLatinSpaceFace(s, breakerState.latinLineFace))
  ) {
    flush(undefined, false, s.src);
    breakerState.queue.unshift(s);
    return;
  }

  placeOrSplitText(context, {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  });
}

function processMathSegment(context: BreakOpportunityIteratorContext, seg: LayoutMathSeg): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    availW,
    setMeasureFont,
    fontFamilyClasses,
    measurement,
  } = context;

  const render = seg.mathMetadata;
  if (!render || render.available === false) {
    const emPx = seg.fontSize * scale;
    setMeasureFont(buildFont(false, false, emPx, null, fontFamilyClasses));
    const m = measurement.measureCurrentText(seg.fallbackText);
    const w = m.width;
    const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? emPx * 0.8;
    const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? emPx * 0.2;
    seg.measuredWidth = w;
    seg.mathAscent = asc;
    seg.mathDescent = desc;
    if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
      flush(undefined, false, seg.src);
    }
    if (breakerState.currentLine.length === 0 && w > availW()) context.forcedPlacement(w);
    addToLine(seg, w, seg.fontSize, Math.max(asc, emPx * 0.8), Math.max(desc, emPx * 0.2));
    return;
  }
  const emPx = seg.fontSize * scale;
  const w = render.widthEm * emPx;
  const asc = render.ascentEm * emPx;
  const desc = render.descentEm * emPx;
  seg.measuredWidth = w;
  // Ink extents (from the MathJax SVG viewBox) position the rasterized
  // glyph relative to the baseline when drawing.
  seg.mathAscent = asc;
  seg.mathDescent = desc;
  // …but the LINE BOX must reserve at least a normal single line for the
  // run's font size. A short equation — e.g. a lone "−" — has near-zero ink
  // height; using that as the line height would collapse the line (and the
  // table row) and pin the glyph to the very top of the cell. Floor to the
  // font's natural ascent/descent so math occupies a full line like text
  // does (tall math — fractions, big operators — keeps its larger ink box).
  const lineAsc = Math.max(asc, emPx * 0.8);
  const lineDesc = Math.max(desc, emPx * 0.2);
  if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
    flush(undefined, false, seg.src);
  }
  if (breakerState.currentLine.length === 0 && w > availW()) context.forcedPlacement(w);
  addToLine(seg, w, seg.fontSize, lineAsc, lineDesc);
  return;
}

function processImageSegment(context: BreakOpportunityIteratorContext, seg: LayoutImageSeg): void {
  const { breakerState, flush, addToLine, scale, availW } = context;

  if (seg.anchor) {
    seg.measuredWidth = 0;
    return;
  }
  const w = seg.widthPt * scale;
  const h = seg.heightPt;
  const asc = seg.heightPt * scale;
  seg.measuredWidth = w;
  if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
    flush(undefined, false, seg.src);
  }
  if (breakerState.currentLine.length === 0 && w > availW()) context.forcedPlacement(w);
  addToLine(seg, w, h, asc, 0);
  return;
}

/** Commit a cell already proven to fit its band. Keeping the aligned cell's
 * measurement path preserves ordinary TOC/field allocation, while callers leave
 * oversized cells in the queue for normal legal-break and emergency fitting. */
function commitAlignedTabCell(context: BreakOpportunityIteratorContext): void {
  const { breakerState, scale, addToLine, measureText, verticalInkExtra, characterGrid } = context;
  while (breakerState.queue.length > 0) {
    const q = breakerState.queue.peek()!;
    if ('isTab' in q || 'lineBreak' in q) break;
    breakerState.queue.shift();
    if ('imagePath' in q) {
      const w = q.widthPt * scale;
      q.measuredWidth = w;
      addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
    } else if ('math' in q) {
      addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
    } else {
      const m = measureText(q);
      // #1014 — fold the vo=Tr ink deficit into the committed advance too.
      const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), characterGrid, scale);
      q.measuredWidth = w;
      const asc =
        m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
      const desc =
        m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
      addToLine(q, w, q.fontSize, asc, desc);
    }
  }
}

/**
 * Allocation available to an aligned (right/center/decimal) ordinary tab cell.
 *
 * ECMA-376 §17.3.1.37 positions custom stops relative to the page margins and
 * does not bound them by the paragraph's trailing indent. So on a line that
 * no float narrows, an aligned ordinary cell may extend to the text margin
 * (or the indent edge when a negative indent lies beyond it) wherever no
 * exclusion intersects that extension on the line; the pass records it as
 * `lineMarginExtension`. A line that uses it carries the extension as part of
 * its band, so alignment and justification slack are measured against the
 * same allocation. Positional tabs, ordinary text and left tabs keep the
 * actual band; a float-narrowed window keeps it too, since the #1672 Word
 * controls overlap floats and that overlap is not emulated.
 */
function alignedTabCellAvailW(context: BreakOpportunityIteratorContext): number {
  return context.availW() + context.breakerState.lineMarginExtension;
}

function processTabSegment(
  context: BreakOpportunityIteratorContext,
  seg: LayoutTabSeg,
  resolveCompletedTabCells: ReturnType<typeof createBidiTabCellResolver>,
): void {
  const {
    breakerState,
    flush,
    baseRtl,
    addToLine,
    scale,
    firstIndent,
    tabOriginPx,
    maxWidth,
    marginRightPx,
    tabFollowWidth,
    tabStops,
    defaultTabPt,
    tabFollowingMetrics,
    availW,
  } = context;
  seg.marginAllocation = false;

  // Ordinary RTL stops still require the complete cell's visual order. A
  // positional tab additionally owns a normative next-line decision, which
  // must happen here, while the queue can still move (§17.3.3.23).
  if (baseRtl) {
    let oversizedMarginLeading = false;
    if (seg.ptab) {
      const input = { ...context, ...breakerState };
      breakerState.currentWidth += resolveCompletedTabCells(input);
      const { startPen, leftLimit, frame } = bidiTabFrame(input);
      let followingWidth = 0;
      for (const q of breakerState.queue) {
        if ('isTab' in q || 'lineBreak' in q) break;
        followingWidth += tabFollowWidth(q);
      }
      const box = wordPositionalTabReferenceBox(
        seg.ptab.relativeTo === 'indent' ? frame.indentStart : 0,
        seg.ptab.relativeTo === 'indent' ? frame.indentEnd : leftLimit,
        frame.bandStart, frame.bandEnd, frame.narrowed);
      // Library scope policy preserves the established no-float normal-fitting
      // contract for an oversized leading margin cell. Its complete contents
      // cannot align within the reference, so retain the ordinary automatic
      // gap used by that contract, without claiming an additional observed rule.
      oversizedMarginLeading = !frame.narrowed && seg.ptab.alignment === 'left'
        && seg.ptab.relativeTo === 'margin' && followingWidth > box.end - box.start;
      const target = positionalTabTarget(seg.ptab.alignment, box.start, box.end, followingWidth);
      if (target < startPen + breakerState.currentWidth && breakerState.currentLine.length > 0) {
        flush(undefined, false, seg.src);
        breakerState.queue.unshift(seg);
        return;
      }
    }
    const input = { ...context, ...breakerState };
    const { startPen, frame } = bidiTabFrame(input);
    if (frame.narrowed || oversizedMarginLeading || seg.ptab?.relativeTo === 'indent') {
      // Resolve leading gaps before fitting, in the same reading frame as
      // paint. A provisional zero gap can fit text that the final tab would
      // push into the exclusion. Without float exclusion, the final ordinary
      // tab gap stays provisional, even after multiple tabs: the bidi walk can
      // reduce the last gap as its cell fills the band. Freezing that gap here
      // prematurely wraps otherwise in-band text. An indent-relative leading
      // ptab instead has an authored fixed gap.
      // Aligned cells retain post-pass alignment,
      // bounded by the actual band after normal text fitting.
      breakerState.currentWidth += resolveCompletedTabCells(input);
      const pen = startPen + breakerState.currentWidth;
      const stop = seg.ptab && !oversizedMarginLeading ? undefined : nextLineTabStop(pen,
        context.bidiCustomStopsPx, context.bidiIntervalPx, frame.leadingShift);
      const leading = seg.ptab ? seg.ptab.alignment === 'left'
        : !stop || tabAlignmentRole(stop.alignment) === 'leading';
      if (leading) {
        const target = seg.ptab && !oversizedMarginLeading
          ? wordPositionalTabReferenceBox(
            seg.ptab.relativeTo === 'indent' ? frame.indentStart : 0,
            seg.ptab.relativeTo === 'indent' ? frame.indentEnd : input.marginRightPx + tabOriginPx,
            frame.bandStart, frame.bandEnd, frame.narrowed).start
          : stop?.pos ?? pen;
        if (target > frame.bandEnd && breakerState.currentLine.length > 0) {
          flush(undefined, false, seg.src);
          breakerState.queue.unshift(seg);
          return;
        }
        const gap = target > frame.bandEnd ? 0 : Math.max(0, target - pen);
        seg.readingGap = gap;
        seg.measuredWidth = gap;
        seg.leader = stop?.leader;
        addToLine(seg, gap, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
        return;
      }
    } else if (!seg.ptab) {
      // Earlier cells are complete when another ordinary tab is reached. Charge
      // their resolved gaps so the next cell cannot overflow the line, but leave
      // this final gap provisional: the bidi walk may shrink it as text fits.
      breakerState.currentWidth += resolveCompletedTabCells(input);
    }
    seg.measuredWidth = 0;
    addToLine(seg, 0, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    return;
  }

  // LTR pen coordinates include the available line window's displacement.
  // RTL tabs are resolved in reading order by the post-pass above.
  const lineOrigin = breakerState.lineXOffset;
  const absFromParaX = lineOrigin + breakerState.currentWidth
    + (breakerState.isFirst ? firstIndent : 0);

  // ── ECMA-376 §17.3.3.23 absolute-position tab (<w:ptab>) ──────────────
  // A ptab ignores the paragraph's custom tab stops and the default-tab
  // interval; it advances to a fixed position on the line derived from its
  // `alignment` (§17.18.71) and `relativeTo` (§17.18.73). The `alignment`
  // ALSO governs how the text after the ptab aligns to that position (left /
  // centered / right). All coordinates below are paraX-relative px.
  //
  if (seg.ptab) {
    seg.resolvedAlignment = seg.ptab.alignment;
    const bandLeft = breakerState.lineXOffset + (breakerState.isFirst ? Math.min(0, firstIndent) : 0);
    const bandRight = breakerState.lineXOffset + breakerState.lineMaxWidth;
    const box = wordPositionalTabReferenceBox(
      seg.ptab.relativeTo === 'indent' ? 0 : -tabOriginPx,
      seg.ptab.relativeTo === 'indent' ? maxWidth : marginRightPx,
      bandLeft, bandRight, breakerState.lineXOffset !== 0 || breakerState.lineMaxWidth !== maxWidth);
    // Width of the content that trails the ptab up to the next tab / line end
    // — needed to right-/center-align it against `target` (the trailing text
    // is what aligns to the stop, §17.18.71).
    let followW = 0;
    for (const q of breakerState.queue) {
      if ('isTab' in q || 'lineBreak' in q) break;
      followW += tabFollowWidth(q);
    }
    const target = positionalTabTarget(seg.ptab.alignment, box.start, box.end, followW);
    let tabW = target - absFromParaX;
    // §17.3.3.23: "If the alignment location … cannot be found on the current
    // line, because the starting location is past that point, then the tab …
    // shall advance to that location on the next available line." So when the
    // pen already sits past the target, wrap the ptab (and its trailing
    // content) to a fresh line — unless the line is empty (nowhere to wrap).
    if (tabW < 0) {
      if (breakerState.currentLine.length > 0) {
        flush(undefined, false, seg.src);
        breakerState.queue.unshift(seg);
        return;
      }
      // Empty line: cannot advance backwards; contribute no width but keep the
      // segment so the line-height reflects the ptab's font.
      tabW = 0;
    }
    if (breakerState.currentWidth + tabW > availW()) {
      if (breakerState.currentLine.length > 0) {
        flush(undefined, false, seg.src);
        breakerState.queue.unshift(seg);
        return;
      }
      tabW = 0;
    }
    seg.measuredWidth = tabW;
    addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    // §17.3.3.23 selects the reference target independently of paragraph
    // indents. Library policy limits allocation to the actual paragraph/float
    // band: a fitting cell stays aligned, and all other cells use normal breaks.
    if (seg.ptab.alignment !== 'left' && breakerState.currentWidth + followW <= availW()) {
      commitAlignedTabCell(context);
    }
    return;
  }
  // ECMA-376 §17.3.1.37 / §17.15.1.25 — resolve the next stop in TEXT-MARGIN
  // coordinates (the same origin as custom stops): the current pen position is
  // `absFromParaX + tabOriginPx`, custom stops are `pos * scale`, and the
  // automatic grid interval is `defaultTabPt * scale`. Mixing paraX and margin
  // coordinates is what diverged leading-tab rows from labeled ones; computing
  // both in margin space and converting back keeps them aligned.
  const curMarginPx = absFromParaX + tabOriginPx;
  const customStopsPx = tabStops.map((t) => ({
    pos: t.pos * scale,
    alignment: t.alignment,
    leader: t.leader,
  }));
  const stop = nextLineTabStop(curMarginPx, customStopsPx, defaultTabPt * scale,
    lineOrigin);
  seg.resolvedAlignment = stop?.alignment ?? 'left';
  // Convert the chosen margin-space stop back to paraX-relative px.
  const stopParaX = stop ? stop.pos - tabOriginPx : absFromParaX;
  // Right/center/decimal tab: place the tab + its trailing content (up to the next
  // tab / line end) so the content ends at / centers on the stop. An oversized
  // cell still uses ordinary break opportunities
  // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
  // Automatic stops returned by nextTabStop are left-aligned, so they fall
  // through to the left-tab path below.
  const alignmentRole = stop ? tabAlignmentRole(stop.alignment) : 'leading';
  if (stop && alignmentRole !== 'leading') {
    const stopX = stopParaX;
    seg.leader = stop.leader;
    const following = tabFollowingMetrics();
    const alignmentWidth =
      alignmentRole === 'center'
        ? following.totalWidth / 2
        : alignmentRole === 'decimal'
          ? (following.decimalPrefixWidth ?? following.totalWidth)
          : following.totalWidth;
    let tabW = stopX - absFromParaX - alignmentWidth;
    if (tabW <= 0) tabW = 0;
    // Stop coordinates are margin-relative (§17.3.1.37): a fitting cell keeps
    // that allocation on an unnarrowed line, past an authored right indent.
    // Otherwise the gap and cell are limited to the actual paragraph/float band.
    const cellAvail = alignedTabCellAvailW(context);
    const cellLimit = breakerState.currentWidth + tabW + following.totalWidth <= cellAvail
      ? cellAvail : availW();
    seg.marginAllocation = cellLimit > availW()
      && breakerState.currentWidth + tabW + following.totalWidth > availW();
    if (breakerState.currentWidth + tabW > cellLimit) {
      if (breakerState.currentLine.length > 0) {
        flush(undefined, false, seg.src);
        breakerState.queue.unshift(seg);
        return;
      }
      tabW = 0;
    }
    seg.measuredWidth = tabW;
    addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    // Keep an aligned cell atomic only when it fits its band. Oversized cells
    // return to the iterator at their ordinary break sites.
    if (breakerState.currentWidth + following.totalWidth <= cellLimit) {
      commitAlignedTabCell(context);
    }
    return;
  }

  // Left-aligned tab (custom 'left'/'bar'/'clear' or an automatic stop): the
  // pen moves to the stop's paraX. nextTabStop already applied the §17.15.1.25
  // "after all custom stops" automatic grid, so there is no separate fallback.
  let tabWidth = stopParaX - absFromParaX;
  if (stop) seg.leader = stop.leader;
  // ECMA-376 §§17.3.1.37–38: an unavailable stop breaks the line. Library
  // containment policy also applies to float-displaced targets. On an empty
  // line an out-of-band tab contributes no gap, so its cell can use the normal
  // fitting path rather than retrying an unreachable target forever.
  if (breakerState.currentWidth + tabWidth > availW()) {
    if (breakerState.currentLine.length > 0) {
      flush(undefined, false, seg.src);
      breakerState.queue.unshift(seg);
      return;
    }
    tabWidth = 0;
  }
  tabWidth = Math.max(0, tabWidth);
  seg.measuredWidth = tabWidth;
  addToLine(seg, tabWidth, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
  return;
}

interface TextFitFrame {
  readonly s: LayoutTextSeg;
  readonly w: number;
  readonly h: number;
  readonly asc: number;
  readonly desc: number;
  readonly wForFit: number;
  readonly paragraphFinalIdeographicSpaceTail: boolean;
  readonly admitsTrailingOverflowPunctuation: boolean | undefined;
}

/**
 * WORD_COMPRESSED_SPACE_LINE_FIT decides on the paragraph's joined character
 * sequence, not on source runs. When admitted text ends exactly at a run seam
 * that the joined text does not break at — a kinsoku line-end/line-start pair
 * across the seam (§17.15.1.58-.60), a non-starter or word continuation that
 * segmentation marks `joinPrev`/`hardJoinPrev`, or a word without break
 * opportunities continuing in the next run — the following text up to the
 * joined sequence's next legal break joins the candidate. Bounded by the
 * appended characters.
 */
function mixedSeamCandidate(
  context: BreakOpportunityIteratorContext,
  segment: LayoutTextSeg,
  text: string,
  fitWidth: number,
): MixedSpaceCandidate {
  const pieces: { segment: LayoutTextSeg; text: string }[] = [{ segment, text }];
  if (text !== segment.text || text.length === 0 || text.endsWith(' ')) return { pieces, fitWidth };
  const { kinsoku, strAdvance } = context;
  const pairIllegal = (previous: string, next: string): boolean => kinsoku.enabled && (
    kinsoku.lineEndForbidden.has(previous.codePointAt(0)!) ||
    kinsoku.lineStartForbidden.has(next.codePointAt(0)!)
  );
  let previous = [...text].at(-1)!;
  let width = fitWidth;
  for (const follower of context.breakerState.queue) {
    if (!('text' in follower) || follower.text.length === 0) break;
    const characters = [...follower.text];
    const breakable = hasCJKBreakOpportunity(follower.text);
    let take = 0;
    let complete = false;
    while (take < characters.length) {
      const character = characters[take]!;
      if (character === ' ') {
        while (take < characters.length && characters[take] === ' ') take += 1;
        complete = true;
        break;
      }
      const illegal = take === 0
        ? follower.joinPrev === true || follower.hardJoinPrev === true || pairIllegal(previous, character)
        : !breakable || pairIllegal(previous, character);
      if (!illegal) {
        complete = true;
        break;
      }
      previous = character;
      take += 1;
    }
    if (take > 0) {
      const piece = characters.slice(0, take).join('');
      noteMixedSpaceWork(piece.length);
      pieces.push({ segment: follower, text: piece });
      const visible = piece.replace(/ +$/u, '');
      if (visible.length > 0) width += strAdvance(follower, visible);
    }
    if (complete || take < characters.length) break;
  }
  return { pieces, fitWidth: width };
}

/** WORD_COMPRESSED_SPACE_LINE_FIT admission of a whole segment. */
function fitMixedSpaces(
  context: BreakOpportunityIteratorContext,
  segment: LayoutTextSeg,
  fitWidth: number,
): boolean {
  if (!mixedCandidateMayShrink(context.breakerState, segment.text)) return false;
  const required = context.mixedSpaceRequirement(
    mixedSeamCandidate(context, segment, segment.text, fitWidth),
  );
  if (required === undefined) return false;
  if (required > 0) context.markMixedSpacesCompressed();
  return true;
}

/** Whole-word admission at a CJK boundary precedes any ordinary hyphen
 * fallback. Evaluate at the word's origin, before a formatting prefix is placed. */
function prefersWholeWordAtScriptBoundary(context: BreakOpportunityIteratorContext, s: LayoutTextSeg): boolean {
  const previous = context.breakerState.currentLine.at(-1);
  const previousChar = previous && 'text' in previous ? [...previous.text].at(-1) : undefined;
  return s.hyperlink?.kind !== 'external' && previousChar !== undefined
    && hasCJKBreakOpportunity(previousChar) && !hasCJKBreakOpportunity(s.text);
}

function placeOrSplitText(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    availW,
    fitsMeasuredWidth,
    fitHomogeneousLatinSpaces,
    appendQueuedIdeographicSpaceSegment,
    emergencyTextSplit,
    strNaturalAdvance,
    explicitTextSplit,
    queueEmergencyTail,
  } = context;
  const {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;
  if (
    (fitsMeasuredWidth(breakerState.currentWidth + wForFit, availW()) &&
      breakerState.latinAppliedPerGap === 0) ||
    fitHomogeneousLatinSpaces(s, wForFit) ||
    fitMixedSpaces(context, s, wForFit) ||
    fitsMeasuredWidth(breakerState.currentWidth + wForFit, availW())
  ) {
    // Fits on current line as-is
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    appendQueuedIdeographicSpaceSegment(s);
  } else if (admitsTrailingOverflowPunctuation) {
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    appendQueuedIdeographicSpaceSegment(s);
  } else if (
    hasCJKBreakOpportunity(s.text) &&
    s.seaBreaks === undefined &&
    s.hardJoinPrev !== true
  ) {
    splitCjkOverflow(context, frame);
  } else if (s.seaBreaks !== undefined && s.hardJoinPrev !== true) {
    splitSeaOverflow(context, frame);
  } else if (breakerState.currentLine.length === 0) {
    // `word-overlong-token-emergency-break`: for a single non-CJK token wider
    // than a full line, fit the widest character prefix (at least one
    // character), draw it, and re-queue the remainder. Segments are already
    // space-delimited, so this cannot bypass an ordinary space opportunity.
    const semanticSplit = explicitTextSplit(s, availW());
    const split = semanticSplit || emergencyTextSplit(s, availW());
    // An explicit hyphen/URL opportunity is a legal break; any other split is forced,
    // and so is a retained prefix (even one whole grapheme) wider than the gap.
    if ((semanticSplit === 0 && split < s.text.length)
      || placedAdvanceExceeds(context, s, s.text.slice(0, split))) {
      context.forcedPlacement(context.minimalLegalTextWidth(s));
    }
    if (split >= s.text.length) {
      // The visible glyphs actually fit (only a trailing space pushed it over the
      // fit test) — place the word whole.
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
    } else {
      const prefix = s.text.slice(0, split);
      const pw = strNaturalAdvance(s, prefix);
      addToLine(
        {
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        },
        pw,
        h,
        asc,
        desc,
      );
      queueEmergencyTail(s, split);
    }
  } else {
    const semanticSplit = prefersWholeWordAtScriptBoundary(context, s)
      ? 0 : explicitTextSplit(s, availW() - breakerState.currentWidth);
    if (semanticSplit > 0 && semanticSplit < s.text.length) {
      const prefix = s.text.slice(0, semanticSplit);
      const pw = strNaturalAdvance(s, prefix);
      addToLine(
        {
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        },
        pw,
        h,
        asc,
        desc,
      );
      queueEmergencyTail(s, semanticSplit);
      return;
    }
    if (s.joinPrev) {
      // LB14 and the other UAX glue rules prohibit a line boundary at this
      // source seam. If the complete glued group is wider than the fresh
      // line, split this member at the widest legal grapheme boundary that
      // fits the actual remaining band. This bases the decision on the group
      // advance, not on the follower's standalone width.
      const remaining = availW() - breakerState.currentWidth;
      const split = emergencyTextSplit(s, remaining, true);
      // Splitting a glued member, or letting it overflow, forces its unit.
      const retained = (remaining > 0 || s.hardJoinPrev === true) && split > 0 && split < s.text.length
        ? s.text.slice(0, split) : s.text;
      if (split < s.text.length || placedAdvanceExceeds(context, s, retained)) {
        const unitStart = joinedUnitStart(breakerState.currentLine);
        context.forcedPlacement(
          lineAdvanceFrom(breakerState.currentLine, unitStart) + context.minimalLegalTextWidth(s),
          unitStart,
        );
      }
      if ((remaining > 0 || s.hardJoinPrev === true) && split > 0 && split < s.text.length) {
        const prefix = s.text.slice(0, split);
        const pw = strNaturalAdvance(s, prefix);
        addToLine(
          {
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: prefix,
            measuredWidth: pw,
            ...slicedTextMetadata(s, 0, prefix.length),
          },
          pw,
          h,
          asc,
          desc,
        );
        queueEmergencyTail(s, split);
        return;
      }
      // A scalar span that continues the preceding grapheme (or another
      // explicitly glued piece) may overflow a pathological narrow line, but
      // it must never become a new line head and tear the cluster.
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
      return;
    }
    // Latin token does not fit on the current (non-empty) line: move it to a fresh
    // line and re-process. There it either fits, or — when it is wider than the
    // whole column — the empty-line branch above breaks it at the character level
    // (overflow-wrap). Re-queueing rather than force-adding is what lets that
    // over-long-word path run instead of letting the word spill the column.
    flush(undefined, false, s.src);
    breakerState.queue.unshift(s);
  }
}

function prepareAtomicTextFit(
  context: BreakOpportunityIteratorContext,
  frame: {
    s: LayoutTextSeg;
    w: number;
    trailingSpaceW: number;
    sDictSea: boolean;
    fitWidthFor: (widthPx: number, trailingSpacePx: number, next: LayoutSeg | undefined) => number;
  },
): boolean {
  const { breakerState, flush, availW, segAdvance, strAdvance } = context;
  const { s, w, trailingSpaceW, sDictSea, fitWidthFor } = frame;
  const preferWholeWord = prefersWholeWordAtScriptBoundary(context, s);
  const follower = breakerState.queue.peek() as LayoutTextSeg | undefined;
  // A leader's own internal opportunity ends the admission unit. If its
  // full segment fits but the next legal follower prefix does not, select
  // that opportunity before placing the leader; otherwise the tail forces overflow.
  if (!preferWholeWord && s.explicitBreaks && follower?.joinPrev) {
    const whole = measureJoinedTextUnit(s, breakerState.queue, context, w, trailingSpaceW, 0, false, 'whole-leader');
    if (breakerState.currentWidth + fitWidthFor(whole.width, whole.trailingSpace, whole.next) > availW()) {
      const split = context.explicitTextSplit(s, availW() - breakerState.currentWidth);
      if (split > 0 && split < s.text.length) {
        context.queueEmergencyTail(s, split);
        breakerState.queue.unshift({ ...s, ...RESET_SLICED_TEXT_MEASUREMENT,
          text: s.text.slice(0, split), ...slicedTextMetadata(s, 0, split) });
        return true;
      }
    }
  }
  if (
    !s.joinPrev &&
    breakerState.currentLine.length > 0 &&
    (follower?.joinPrev || preferWholeWord && follower?.explicitBreakBefore) &&
    ((breakerState.queue.peek() as LayoutTextSeg | undefined)?.hardJoinPrev === true ||
      !hasCJKBreakOpportunity(s.text)) &&
    // A SEA (Thai/Lao/Khmer) lead with usable word breaks is NOT atomic — the
    // run splits at a dictionary boundary (issue #797), mirroring the CJK gate.
    ((breakerState.queue.peek() as LayoutTextSeg | undefined)?.hardJoinPrev === true ||
      !(s.seaBreaks && s.seaBreaks.length > 0))
  ) {
    const group = measureJoinedTextUnit(s, breakerState.queue, context, w, trailingSpaceW, 0, false, preferWholeWord ? 'whole-word' : 'prefix');
    const groupFitWidth = fitWidthFor(group.width, group.trailingSpace, group.next);
    if (
      breakerState.currentWidth + groupFitWidth > availW() &&
      // WORD_COMPRESSED_SPACE_LINE_FIT judges a joined unit as one candidate,
      // so a source-run seam inside it cannot change the decision.
      !(mixedCandidateMayShrink(breakerState, group.pieces.map((piece) => piece.text).join(''))
        && context.mixedSpaceRequirement({ pieces: group.pieces, fitWidth: groupFitWidth }) !== undefined)
    ) {
      flush(undefined, false, s.src);
    }
  }

  // `word-dictionary-sea-atomic-chunk`: ECMA-376 prescribes no SEA
  // line-breaking algorithm. Treat dictionary boundaries inside a no-space
  // Thai/Lao/Khmer chunk as secondary opportunities: move a chunk that fits a
  // full line as a unit; only a full-line-overlong chunk breaks at dictionary
  // boundaries through the greedy SEA branch below.
  //
  // Judged only at chunk START: if the previously committed token is a text
  // segment glued to `s` (no trailing space), the whole chunk already passed
  // this judgment when its head was placed, so a mid-chunk segment never
  // needs it. The chunk spans `s` plus following queue segments while they
  // stay dictionary-SEA text glued without intervening spaces. Grapheme-fill
  // scripts (Myanmar/Tibetan) are excluded because their per-cluster path
  // fills the remaining width.
  if (
    sDictSea &&
    breakerState.currentLine.length > 0 &&
    (() => {
      const last = breakerState.currentLine[breakerState.currentLine.length - 1];
      return !('text' in last) || (last as LayoutTextSeg).text.endsWith(' ');
    })()
  ) {
    let chunkW = w;
    let chunkTrail = trailingSpaceW;
    let next = breakerState.queue.peek();
    if (!s.text.endsWith(' ')) {
      const following = breakerState.queue[Symbol.iterator]();
      for (let step = following.next(); !step.done; step = following.next()) {
        const f = step.value;
        next = f;
        if (!('text' in f) || (f as LayoutTextSeg).seaBreaks === undefined) break;
        if (!isDictionarySeaText((f as LayoutTextSeg).text)) break;
        const ft = f as LayoutTextSeg;
        const fw = segAdvance(ft);
        const fTrim = ft.text.replace(/ +$/, '');
        chunkW += fw;
        chunkTrail = ft.text.endsWith(' ') ? fw - strAdvance(ft, fTrim) : 0;
        if (ft.text.endsWith(' ')) {
          next = following.next().value;
          break;
        } // a space ends the chunk
        next = undefined;
      }
    }
    const chunkWForFit = fitWidthFor(chunkW, chunkTrail, next);
    if (
      breakerState.currentWidth + chunkWForFit > availW() &&
      chunkWForFit <= breakerState.lineMaxWidth
    ) {
      flush(undefined, false, s.src);
    }
  }
  return false;
}

function splitCjkOverflow(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    characterGrid,
    availW,
    setMeasureFont,
    fontFamilyClasses,
    measurement,
    strAdvance,
    overflowPunct,
    appendQueuedIdeographicSpaceSegment,
    emergencyTextSplit,
    effectiveFontPx,
    ctx,
    verticalGlyphMeasurement,
    kinsoku,
    strNaturalAdvance,
    retractCurrentLineForLeadingKinsoku,
    keepLeadingKinsokuWithCurrentLine,
  } = context;
  const {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;

  // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
  // A segment that ALSO contains SEA (a mixed CJK+SEA `<w:cs/>` run) is routed
  // to the SEA branch below instead — its `seaBreaks` already merges the CJK
  // per-character opportunities with the SEA dictionary/transition ones
  // (issue #960), so both scripts break by their own rule from one offset set.
  // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
  //  separate: it sums per-char advances, whereas this path uses substring
  //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
  const available = availW() - breakerState.currentWidth;
  let rawPrefix = '';
  const maximumIdeographicSpaceHang = paragraphFinalIdeographicSpaceTail
    ? wordIdeographicSpaceLineEndAllowanceCount(
        hasEastAsianVisiblePredecessor(s.text),
        s.paragraphFinalIdeographicSpaceCount ?? 0,
      )
    : Number.POSITIVE_INFINITY;
  if (available > 0) {
    const nonMonotoneAllocation =
      charSpacingDeltaPx(s, scale) < 0 || snapToCharsClass(s, characterGrid) === 'latin';
    if (nonMonotoneAllocation) {
      rawPrefix = s.text.slice(0, emergencyTextSplit(s, available, false));
    } else {
      setMeasureFont(
        buildFont(
          s.bold,
          s.italic,
          effectiveFontPx(s),
          s.fontFamily,
          fontFamilyClasses,
          s.fontRoute,
        ),
      );
      measurement.withSegmentKerning(s, () => {
        rawPrefix = fitCJKPrefix(
          ctx,
          s.text,
          available,
          segmentCharacterGridDeltaPx(s, characterGrid, scale),
          charScaleFactor(s),
          charSpacingDeltaPx(s, scale),
          s.verticalRun === true,
          verticalGlyphMeasurement,
          (prefix) => strAdvance(s, prefix),
          maximumIdeographicSpaceHang,
        );
      });
    }
  }
  // WORD_COMPRESSED_SPACE_LINE_FIT: a mixed line's U+0020 may shrink to admit
  // further characters of this run. Extend the natural prefix while the space
  // floors and the East Asian overflow limit admit it; kinsoku below may still
  // retract the break.
  //
  // Admission is monotone in the prefix length: a longer prefix needs at least
  // as much reduction, and its core (closing marks excluded) overflows at
  // least as far. The longest admitted prefix is therefore found by binary
  // search, measuring O(log n) prefixes; lines outside the rule's scope stop
  // at the O(1) gate before any candidate text is built.
  const allChars = [...s.text];
  const minSplit = breakerState.currentLine.length > 0 ? 0 : 1;
  const proposedPrefixFor = (split: number, extensions: boolean): string => {
    const hangingSplit =
      extensions &&
      overflowPunct &&
      breakerState.gapTransaction?.endsAtExclusion !== true &&
      split < allChars.length &&
      (breakerState.currentLine.length > 0 || split > 0) &&
      wordIsOverflowPunctuation(
        allChars[split],
        s.eastAsiaLanguage,
        s.overflowPunctuationEastAsianRun === true,
        s.script === 'ascii' || s.script === 'highAnsi',
        s.script === 'complexScript',
        s.overflowPunctuationBidiLanguage,
      )
        ? split + 1
        : null;
    const adjusted = hangingSplit ?? kinsokuAdjustedSplit(allChars, split, kinsoku, minSplit);
    const proposedSplit = extensions
      ? extendThroughTrailingIdeographicSpaces(
          allChars,
          adjusted,
          paragraphFinalIdeographicSpaceTail && maximumIdeographicSpaceHang === 0
            ? 0
            : maximumIdeographicSpaceHang,
        )
      : adjusted;
    const proposedPrefix = allChars.slice(0, proposedSplit).join('').length;
    return s.text.slice(0, legalTextSplitAtOrBefore(s, proposedPrefix, minSplit > 0 ? 1 : 0));
  };
  let mixedSpaceExtended = false;
  const naturalRawSplit = [...rawPrefix].length;
  if (breakerState.currentLine.length > 0 && mixedCandidateMayShrink(breakerState, s.text)) {
    const characters = [...s.text];
    const admitted = (count: number): boolean => {
      const candidate = characters.slice(0, count).join('');
      noteMixedSpaceWork(candidate.length);
      return !candidate.endsWith(' ') && context.mixedSpaceRequirement(
        mixedSeamCandidate(context, s, candidate, strAdvance(s, candidate)),
      ) !== undefined;
    };
    // Galloping bound, then binary search: every measured prefix is at most
    // twice the admitted extension past the natural break.
    // Start from the natural break as the ordinary pipeline resolves it
    // (including its hanging allowances), so an allowance that needs no
    // reduction is not turned into one by a source-run seam.
    const naturalHead = [...proposedPrefixFor(naturalRawSplit, true)].length;
    let low = Math.max(naturalRawSplit, naturalHead);
    let high = characters.length;
    for (let step = 1; low + step <= characters.length; step *= 2) {
      if (!admitted(low + step)) {
        high = low + step - 1;
        break;
      }
      low += step;
    }
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (admitted(middle)) low = middle;
      else high = middle - 1;
    }
    // Every shorter prefix is admitted too; keep the longest one that is also
    // a kinsoku-legal break inside this run, so the extension never relies on
    // kinsoku's unrestricted fallback (a seam before closing marks must
    // resolve exactly like the joined text, by cross-run retraction).
    const natural = Math.max(naturalRawSplit, naturalHead);
    const legal = (count: number): boolean => !kinsoku.enabled || count >= characters.length || (
      !kinsoku.lineStartForbidden.has(characters[count]!.codePointAt(0)!) &&
      !kinsoku.lineEndForbidden.has(characters[count - 1]!.codePointAt(0)!)
    );
    while (low > natural && !legal(low)) low -= 1;
    if (low > natural) {
      rawPrefix = characters.slice(0, low).join('');
      mixedSpaceExtended = true;
    }
  }
  // Apply kinsoku to the break position: retract leftwards so the tail
  // never begins with a 行頭禁則 char and the head never ends with a
  // 行末禁則 char (ECMA-376 §17.15.1.58–.60). When the current line
  // already has content, retracting to an empty prefix is allowed — the
  // whole run moves to the next (fresh) line, which is Word's 追い出し.
  // When the line is empty we keep at least one char (minSplit=1) so we
  // never lose forward progress.
  const rawSplit = [...rawPrefix].length;
  // ECMA-376 §17.3.1.21 permits one punctuation character beyond the
  // paragraph extents. The isolated compatibility projection resolves the
  // language-specific set and its precedence over kinsoku at this internal
  // CJK split.
  let prefix = proposedPrefixFor(rawSplit, true);
  if (mixedSpaceExtended) {
    // Every fit decision on a shrinking line goes through the one predicate:
    // the final head (after punctuation hanging and source
    // protection) must itself be admitted. Otherwise keep the checked
    // kinsoku-legal head, and failing that the natural break.
    const needsShrink = (head: string): boolean =>
      head.length > 0 && breakerState.currentWidth + strAdvance(s, head) > availW();
    // The break after the head must be legal in the joined character sequence
    // (kinsoku pair with the next character, here or in a following run).
    const nextCharacter = (head: string): string | undefined => {
      if (head.length < s.text.length) return [...s.text.slice(head.length)][0];
      for (const follower of breakerState.queue) {
        if (!('text' in follower)) return undefined;
        if (follower.text.length > 0) return [...follower.text][0];
      }
      return undefined;
    };
    const legalBreakAfter = (head: string): boolean => {
      if (!kinsoku.enabled || head.length === 0) return true;
      const next = nextCharacter(head);
      return next === undefined || (
        !kinsoku.lineStartForbidden.has(next.codePointAt(0)!) &&
        !kinsoku.lineEndForbidden.has([...head].at(-1)!.codePointAt(0)!)
      );
    };
    const admittedHead = (head: string): boolean => !needsShrink(head) || (
      legalBreakAfter(head) &&
      context.mixedSpaceRequirement(mixedSeamCandidate(context, s, head, strAdvance(s, head))) !== undefined
    );
    if (!admittedHead(prefix)) {
      const checked = proposedPrefixFor(rawSplit, false);
      prefix = admittedHead(checked) ? checked : proposedPrefixFor(naturalRawSplit, true);
    }
    mixedSpaceExtended = needsShrink(prefix);
  }
  if (breakerState.currentLine.length === 0) {
    // minSplit keeps progress on an empty line even when no legal prefix
    // fits; that retained prefix is forced unless a legal one exists.
    const legalSplit = kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, 0);
    const legalUtf16 = allChars.slice(0, legalSplit).join('').length;
    if (legalTextSplitAtOrBefore(s, legalUtf16, 1) === 0 || placedAdvanceExceeds(context, s, prefix)) {
      context.forcedPlacement(context.minimalLegalTextWidth(s));
    }
  }
  if (prefix.length > 0) {
    if (mixedSpaceExtended) {
      // The retained head overflows naturally and was admitted above.
      context.markMixedSpacesCompressed();
    }
    // Grid advance for the head piece — the same model as the line box / draw.
    const pw = strNaturalAdvance(s, prefix);
    const headSeg: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: prefix,
      measuredWidth: pw,
      ...slicedTextMetadata(s, 0, prefix.length),
    };
    addToLine(headSeg, pw, h, asc, desc);
    const tail = s.text.slice(prefix.length);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, prefix.length, s.text.length),
        measuredWidth: 0,
        src: {
          segIndex: s.src!.segIndex,
          charOffset: s.src!.charOffset + prefix.length,
        },
      });
    } else {
      appendQueuedIdeographicSpaceSegment(s);
    }
  } else if (breakerState.currentLine.length > 0) {
    // No prefix of `s` fits. If `s` would lead the next line with a 行頭禁則
    // char, kinsokuAdjustedSplit can't fix it from within `s` (the offending
    // char is its first); pull trailing graphemes of the current line's last
    // text segment down so they lead the next line ahead of `s` — cross-run
    // 追い出し (§17.3.1.16). See crossRunKinsokuRetract for the bounded,
    // re-validating, whitespace-guarded retraction count.
    const retraction = retractCurrentLineForLeadingKinsoku(s);
    if (retraction.kind === 'blocked') {
      // Keeping the forbidden leader overflows the line: force its unit.
      const unitStart = joinedUnitStart(breakerState.currentLine);
      context.forcedPlacement(
        lineAdvanceFrom(breakerState.currentLine, unitStart)
          + strAdvance(s, s.text.slice(0, graphemeClusterOffsets(s.text).find((offset) => offset > 0) ?? s.text.length)),
        unitStart,
      );
      keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
      return;
    }
    flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
    breakerState.queue.unshift(s);
    if (retraction.kind === 'retracted') breakerState.queue.unshift(retraction.tail);
  } else {
    // Empty line and not even one char fits — force-fit one char to guarantee progress
    const forcedChars = [...s.text];
    const forcedSplit =
      forcedChars.length > 0
        ? extendThroughTrailingIdeographicSpaces(
            forcedChars,
            1,
            s.paragraphFinalIdeographicSpaceTail === true
              ? wordIdeographicSpaceLineEndAllowanceCount(
                  EAST_ASIAN_RE.test(forcedChars[0] ?? ''),
                  s.paragraphFinalIdeographicSpaceCount ?? 0,
                )
              : Number.POSITIVE_INFINITY,
          )
        : 0;
    const forcedUtf16 = forcedChars.slice(0, forcedSplit).join('').length;
    const legalForcedUtf16 =
      legalTextSplitAtOrBefore(s, forcedUtf16) || emergencyTextSplit(s, availW(), true);
    const firstChar = s.text.slice(0, legalForcedUtf16);
    if (firstChar) {
      const fw = strNaturalAdvance(s, firstChar);
      const headSeg: LayoutTextSeg = {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: firstChar,
        measuredWidth: fw,
        ...slicedTextMetadata(s, 0, firstChar.length),
      };
      addToLine(headSeg, fw, h, asc, desc);
      const tail = s.text.slice(firstChar.length);
      if (tail) {
        breakerState.queue.unshift({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: tail,
          ...slicedTextMetadata(s, firstChar.length, s.text.length),
          measuredWidth: 0,
          src: {
            segIndex: s.src!.segIndex,
            charOffset: s.src!.charOffset + firstChar.length,
          },
        });
      } else {
        appendQueuedIdeographicSpaceSegment(s);
      }
    }
  }
}

function splitSeaOverflow(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    characterGrid,
    availW,
    strAdvance,
    emergencyTextSplit,
    strNaturalAdvance,
    retractCurrentLineForLeadingKinsoku,
    keepLeadingKinsokuWithCurrentLine,
  } = context;
  const {
    s: segment,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;
  const s = segment as LayoutTextSeg & { seaBreaks: readonly number[] };

  // No-inter-word-space line wrap: Thai/Lao/Khmer dictionary words (#797) or
  // Myanmar/Tibetan grapheme clusters (#961). This ONE segment is a whole run;
  // break it only at a member of `s.seaBreaks` — the UNION (#960) of the
  // dictionary word (or grapheme-cluster) boundaries, the no-space SEA↔non-SEA
  // script transitions, and (for a mixed CJK+SEA `<w:cs/>` run) the CJK
  // per-character opportunities, already kinsoku-filtered by
  // `seaMixedBreakOffsets`. Entered for ANY such segment (even one with no
  // interior boundary — a single word/cluster wider than the column, or
  // Segmenter unavailable) so the emergency split below stays GRAPHEME-safe
  // instead of falling to the code-point path. Kinsoku 行頭/行末禁則 was applied
  // when the offsets were built (so a forbidden CJK char never heads/tails a
  // line); choosing an earlier legal offset is the only remaining adjustment,
  // which fitSeaWordPrefix already does. The run stays one contiguous draw per
  // line (measure==paint); the tail re-queues with its offsets rebased.
  const available = availW() - breakerState.currentWidth;
  const measureSub = (sub: string): number => strAdvance(s, sub);
  // Grapheme-fill runs (Myanmar/Tibetan) have DENSE offsets (one per cluster),
  // so use the monotone binary-search fit — a per-line full scan would be O(n²)
  // down a long run. Dictionary runs keep the negative-spacing-safe full scan.
  const monotone =
    isGraphemeFillText(s.text) &&
    charSpacingDeltaPx(s, scale) >= 0 &&
    snapToCharsClass(s, characterGrid) !== 'latin';
  const split = fitSeaWordPrefix(s.text, s.seaBreaks, 0, available, measureSub, monotone);
  if (split > 0 && breakerState.currentLine.length === 0
    && placedAdvanceExceeds(context, s, s.text.slice(0, split))) {
    context.forcedPlacement(context.minimalLegalTextWidth(s));
  }
  if (split > 0) {
    const prefix = s.text.slice(0, split);
    const pw = strNaturalAdvance(s, prefix);
    addToLine(
      {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: prefix,
        measuredWidth: pw,
        ...slicedTextMetadata(s, 0, prefix.length),
      },
      pw,
      h,
      asc,
      desc,
    );
    const tail = s.text.slice(split);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, split, s.text.length),
        measuredWidth: 0,
        src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + split },
        seaBreaks: rebaseSeaBreaks(s.seaBreaks, split),
      });
    }
  } else if (breakerState.currentLine.length > 0) {
    // No whole word fits the remaining band — move the run to a fresh line and
    // re-process (Latin-word style). If `s` would then LEAD the next line with
    // a 行頭禁則 char (a mixed CJK+SEA run whose first glyph is a forbidden
    // leader — #960 routes it here, where the offset set cannot fix a
    // segment-initial char), pull trailing graphemes of the current line's
    // last text segment down so they lead ahead of `s` — the same cross-run
    // 追い出し (§17.3.1.16) the CJK branch does.
    const retraction = retractCurrentLineForLeadingKinsoku(s);
    if (retraction.kind === 'blocked') {
      // Keeping the forbidden leader overflows the line: force its unit.
      const unitStart = joinedUnitStart(breakerState.currentLine);
      context.forcedPlacement(
        lineAdvanceFrom(breakerState.currentLine, unitStart)
          + strAdvance(s, s.text.slice(0, graphemeClusterOffsets(s.text).find((offset) => offset > 0) ?? s.text.length)),
        unitStart,
      );
      keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
      return;
    }
    flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
    breakerState.queue.unshift(s);
    if (retraction.kind === 'retracted') breakerState.queue.unshift(retraction.tail);
  } else {
    // Empty line and the first dictionary word is wider than the whole
    // column: emergency GRAPHEME-safe split (a code-point split would tear a
    // base + tone/combining mark, both BMP). Guarantee ≥1 cluster of progress.
    context.forcedPlacement(context.minimalLegalTextWidth(s));
    const firstWordEnd = s.seaBreaks[0] ?? s.text.length;
    const firstWord = s.text.slice(0, firstWordEnd);
    const graphemes = graphemeClusterOffsets(firstWord);
    let gsplit = fitSeaWordPrefix(firstWord, graphemes, 0, available, measureSub, monotone);
    if (gsplit <= 0) gsplit = graphemes.length > 0 ? graphemes[0] : firstWord.length;
    gsplit = legalTextSplitAtOrBefore(s, gsplit) || emergencyTextSplit(s, available, true);
    const prefix = s.text.slice(0, gsplit);
    const pw = strNaturalAdvance(s, prefix);
    addToLine(
      {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: prefix,
        measuredWidth: pw,
        ...slicedTextMetadata(s, 0, prefix.length),
      },
      pw,
      h,
      asc,
      desc,
    );
    const tail = s.text.slice(gsplit);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, gsplit, s.text.length),
        measuredWidth: 0,
        src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + gsplit },
        seaBreaks: rebaseSeaBreaks(s.seaBreaks, gsplit),
      });
    }
  }
}
