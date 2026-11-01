import { createPrefixWorkBudget } from './prefix-work-budget.js';
import { LineMeasurementAdapter } from './measurement-adapter.js';
import type { TabStop } from '../types';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { type PreparedFloatWrap } from '../float-layout.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from '../layout/measurement-capabilities.js';
import { wordUniformRunPositionPaintPt } from '../layout/line-compatibility.js';
import { calcEffectiveFontPx } from '../layout/text.js';
import {
  type DocGridCtx,
  type LayoutImageSeg,
  type LayoutLine,
  type LayoutMathSeg,
  type LayoutSeg,
  type LayoutTabSeg,
  type LayoutTextSeg,
  type LineBoundary,
  type WrapLayoutCtx,
} from './model.js';
import { createLineBreakerState, prepareBreakQueue } from './break-queue.js';
import { SegmentQueue } from './segment-queue.js';
import { buildFont } from './font-routes.js';
import { type CrossRunKinsokuRetraction } from './kinsoku.js';
import { iterateBreakOpportunities } from './break-opportunities.js';
import { finalizeRetainedLineShapes } from './line-finalize.js';
import {
  performMarkMixedSpacesCompressed,
  performMixedSpaceRequirement,
  type MixedSpaceCandidate,
} from './mixed-space-fit.js';
import {
  performLineHeadRequirement,
  performForcedPlacement,
  performMinimalLegalTextWidth,
  performRejectGap,
  performCaptureGapSnapshot,
  performSameLatinSpaceFace,
  performMaterializeLatinSpaceCompression,
  performStartLine,
  performAvailW,
  performFitsMeasuredWidth,
  performFlush,
  performProspectiveSnapAdvance,
  performAddToLine,
  performAdmitLineMetric,
  performSegNaturalAdvance,
  performStandaloneSnapAdvance,
  performSegAdvance,
  performStrNaturalAdvance,
  performEastAsianSnapCellCount,
  performStrAdvance,
  performFitHomogeneousLatinSpaces,
  performTextSegmentBox,
  performAppendQueuedIdeographicSpaceSegment,
  performTabFollowWidth,
  performDecimalAlignmentPoint,
  performDecimalAlignmentPrefixWidth,
  performTabFollowingMetrics,
  performEmergencyTextSplit,
  performExplicitTextSplit,
  performQueueEmergencyTail,
  performRetractCurrentLineForLeadingKinsoku,
  performKeepLeadingKinsokuWithCurrentLine,
} from './pass-operations.js';
import type { PassOperationState } from './pass-operations.js';

/** Paragraph line-fit policy: the WORD_KASHIDA_LINE_FIT reserve and the
 * WORD_RTL_JUSTIFIED_VISIBLE_FIT visible-prefix fit of right-to-left lines. */
export interface LineFitWidth {
  readonly fraction: number;
  /** Share of the line's start-to-far-edge width reserved (before mode 15). */
  readonly reserveShare: number;
  /** A right-to-left line fits its visible prefix, without its line-end U+0020. */
  readonly rtlVisiblePrefix: boolean;
}

export interface LineBreakerPassInput {
  readonly ctx: MeasurementTextContext;
  readonly segs: LayoutSeg[];
  readonly maxWidth: number;
  readonly firstIndent: number;
  readonly scale: number;
  readonly tabStops: TabStop[];
  readonly wrapCtx?: WrapLayoutCtx;
  readonly fontFamilyClasses: Record<string, string>;
  readonly tabOriginPx: number;
  readonly kinsoku: KinsokuRules;
  readonly characterGrid?: DocGridCtx;
  readonly defaultTabPt: number;
  readonly marginRightPx: number;
  readonly baseRtl: boolean;
  readonly isJustified: boolean;
  readonly stretchLastLine: boolean;
  readonly startBoundary?: LineBoundary;
  readonly widthPolicy: 'bounded' | 'intrinsic' | 'unwrapped';
  readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService;
  readonly overflowPunct: boolean;
  readonly justifiedCompression?: boolean;
  /** WORD_KASHIDA_LINE_FIT: a bounded line's text may fill `fraction` of its
   * available width less its reserve (absent: the whole width; paint still
   * justifies to the full width), and WORD_RTL_JUSTIFIED_VISIBLE_FIT. */
  readonly lineFit?: LineFitWidth;
  readonly passContext: Readonly<{
    probeHeights: readonly number[] | null;
    /** Monotone per-physical-line exclusion probe heights (≥ probeHeights). */
    probeFloors?: readonly number[] | null;
    preparedFloatWrap?: PreparedFloatWrap;
  }>;
}

/**
 * One immutable-input line-break pass; all mutable line state is pass-local.
 * This driver is deliberately a longer, flat wiring function: it binds the
 * explicit pass state to the named measurement, fit, tab, and break operations.
 * The decisions and mutations live in those operations and the iterator.
 */
export function runLineBreakerPass(input: LineBreakerPassInput): LayoutLine[] {
  const {
    ctx,
    segs,
    maxWidth,
    scale,
    tabStops,
    wrapCtx,
    fontFamilyClasses,
    kinsoku,
    characterGrid,
    defaultTabPt,
    baseRtl,
    startBoundary,
    widthPolicy,
    verticalGlyphMeasurement,
    passContext,
  } = input;

  const { probeHeights, preparedFloatWrap } = passContext;
  const probeFloors = passContext.probeFloors ?? probeHeights;
  const breakerState = createLineBreakerState(maxWidth, wrapCtx);
  // WORD_COMPRESSED_SPACE_LINE_FIT scope, fixed for the paragraph: only
  // segments acquired under its document gate carry the eligibility. No
  // Word control measured U+3000, whose hanging and paragraph-final rules
  // (WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE) the observed rule does not
  // define; a paragraph holding U+3000 keeps this compression rule disabled.
  breakerState.mixedSpaceEnabled = segs.some(
    (segment) => 'text' in segment
      && (segment.mixedSpaceAverageWidthRatio !== undefined || segment.autoSpaceCompressible === true),
  ) && !segs.some((segment) => 'text' in segment && segment.text.includes('\u3000'));

  let operationState: PassOperationState;
  const sameLatinSpaceFace = performSameLatinSpaceFace;

  const materializeLatinSpaceCompression = (): void =>
    performMaterializeLatinSpaceCompression(operationState);
  const snapPitchPx =
    characterGrid?.type === 'snapToChars' &&
    characterGrid.characterPitchPt != null &&
    characterGrid.characterPitchPt > 0
      ? characterGrid.characterPitchPt * scale
      : null;

  const lineHeadRequirement = (boundary?: LineBoundary): number =>
    performLineHeadRequirement(operationState, boundary);
  const forcedPlacement = (requiredWidth: number, unitStart?: number): void =>
    performForcedPlacement(operationState, requiredWidth, unitStart);
  const minimalLegalTextWidth = (segment: LayoutTextSeg): number =>
    performMinimalLegalTextWidth(operationState, segment);

  // Compute wrap constraints for a new line fragment about to start. Mutates
  // lineXOffset/lineMaxWidth/currentLineTopY. `requirement` seeds the gap
  // search; placement itself admits or rejects a narrowed gap (#1670).
  const startLine = (requirement: number = 0): void => performStartLine(operationState, requirement);

  // Intrinsic acquisition and DrawingML wrap=none disable automatic wrapping
  // while retaining the real paragraph/anchor width for tabs and alignment.
  // Unwrapped paint still finalizes retained shapes; intrinsic measurement does
  // not. Neither mode invents an oversized page.
  const availW = () => performAvailW(operationState);

  // AutoFit can set a table column to the measured text advance plus its
  // first-line indent and grouped cell insets. Cell acquisition subtracts those
  // same grouped insets. At exact equality, subtracting the indent from the
  // line width may round down while adding it to the advance does not. Both
  // inequalities are equivalent in real arithmetic; check both operation
  // orders for a finite bounded line. This changes arithmetic order only;
  // it adds no fixed overflow allowance.
  const fitsMeasuredWidth = (used: number, available: number): boolean =>
    performFitsMeasuredWidth(operationState, used, available);

  // ECMA-376 §17.3.1.37 tab stops in leading-edge px, for the bidi post-pass.
  const bidiCustomStopsPx = baseRtl
    ? tabStops.map((t) => ({ pos: t.pos * scale, alignment: t.alignment, leader: t.leader }))
    : [];
  const bidiIntervalPx = defaultTabPt * scale;

  // Rewrite a finalized bidi line's tab widths (+ leaders) in the VISUAL frame
  // (§17.3.1.6 base RTL). The line's tabs were laid out with provisional width 0
  // by the tab block below (the LTR pen math does not apply under an RTL base);
  // here we place each tab-delimited cell at its mirrored stop. No-op for a line
  // without tabs (LTR paragraphs skip this entirely — `baseRtl` is false).

  const flush = (forceHeight?: number, brTerminated = false, nextStart?: LineBoundary) =>
    performFlush(operationState, forceHeight, brTerminated, nextStart);

  const prospectiveSnapAdvance = (s: LayoutTextSeg, naturalWidth: number): number =>
    performProspectiveSnapAdvance(operationState, s, naturalWidth);

  const addToLine = (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
  ) => performAddToLine(operationState, s, w, h, asc, desc);
  const admitLineMetric = (s: LayoutImageSeg, h: number, asc: number) =>
    performAdmitLineMetric(operationState, s, h, asc);

  const measurement = new LineMeasurementAdapter(
    ctx,
    scale,
    (segment) =>
      buildFont(
        segment.bold,
        segment.italic,
        calcEffectiveFontPx(segment, scale),
        segment.fontFamily,
        fontFamilyClasses,
        segment.fontRoute,
      ),
    verticalGlyphMeasurement,
  );
  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);
  const measureText = (s: LayoutTextSeg, clusterGeometry = false): TextMetrics =>
    measurement.measureSegment(s, clusterGeometry);
  const verticalInkExtra = (s: LayoutTextSeg, text: string): number =>
    measurement.verticalInkExtra(s, text);
  const setMeasureFont = (font: string): void => measurement.setFont(font);

  const endBoundary: LineBoundary = { segIndex: segs.length, charOffset: 0 };
  breakerState.queue = new SegmentQueue(prepareBreakQueue(segs, startBoundary, kinsoku, scale, measurement));

  // The segment's laid-out ADVANCE (= its measuredWidth): natural width plus the
  // character-grid delta, the §17.3.2.43 horizontal glyph scale (w:w) and the
  // §17.3.2.35 character-spacing pitch (w:spacing). This is the SINGLE source of
  // truth shared with the draw paths (segAdvanceWidth) — every line-break / fit /
  // tab measurement uses it so line wrapping packs the grid's char count and the
  // box matches what is drawn (measure==paint). `kerning` (§17.3.2.19) is applied
  // via `ctx.fontKerning` inside `withSegKerning`, wrapping the measureText call.
  // The #1014 vo=Tr ink deficit (`verticalInkExtra`, defined above) is folded into
  // the natural width so measure == paint on an under-reporting vertical run.
  const segNaturalAdvance = (s: LayoutTextSeg): number =>
    performSegNaturalAdvance(operationState, s);
  const standaloneSnapAdvance = (s: LayoutTextSeg, naturalWidth: number): number =>
    performStandaloneSnapAdvance(operationState, s, naturalWidth);
  const segAdvance = (s: LayoutTextSeg): number => performSegAdvance(operationState, s);
  // Grid advance of an arbitrary substring under a segment's font (for split
  // prefixes/tails). Selects the font (and the run's kerning state), then applies
  // the same width model as a whole segment BUT with the substring's own
  // text/length so char-spacing scales with the piece — the split-prefix vs
  // whole-segment advances must agree.
  const strNaturalAdvance = (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression = false,
  ): number =>
    performStrNaturalAdvance(operationState, s, text, retainTrailingPunctuationCompression);
  const eastAsianSnapCellCount = (s: LayoutTextSeg): number =>
    performEastAsianSnapCellCount(operationState, s);
  const strAdvance = (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression = false,
  ): number => performStrAdvance(operationState, s, text, retainTrailingPunctuationCompression);

  /** Compatibility projection governed by WORD_LATIN_INTERWORD_XAVG_FLOOR.
   * The selected face's OS/2 xAvgCharWidth / 2 is the minimum inter-word
   * advance. Equal-face gaps share the required deficit uniformly; a natural
   * space narrower than that minimum retains its natural advance. This is not
   * an ECMA-376 definition or a replacement for glyph advance.
   * Mixed faces, authored spacing, and snap-to-character cells are outside
   * the measured scope and retain their natural widths. The §17.6.5 grid
   * pitch is additive to the selected-face floor. The Word 6 compatibility
   * setting excludes the fit projection at segment acquisition. */
  const fitHomogeneousLatinSpaces = (next: LayoutTextSeg, nextFitWidth: number): boolean =>
    performFitHomogeneousLatinSpaces(operationState, next, nextFitWidth);
  const mixedSpaceRequirement = (candidate: MixedSpaceCandidate): number | undefined =>
    performMixedSpaceRequirement(operationState, candidate);
  const markMixedSpacesCompressed = (): void => performMarkMixedSpacesCompressed(operationState);

  /** Measure one text segment's canonical advance and vertical contribution.
   * Every path that commits a complete text segment to a line must use this
   * authority so font fallback, small-caps, position, ruby and grid metrics do
   * not diverge at internal segment seams. */
  const textSegmentBox = (
    s: LayoutTextSeg,
  ): Readonly<{
    width: number;
    height: number;
    ascent: number;
    descent: number;
  }> => performTextSegmentBox(operationState, s);

  /** Continue the existing U+3000 line-end hanging rule across an internal
   * width-balance segment seam. The split space keeps its own font/grid advance
   * for measure == paint. UAX #14 classifies U+3000 as BA, so an internal seam
   * before it must not become an authored break opportunity. Restrict
   * consumption to the same authored run; an actual source boundary remains
   * independently modeled. */
  const appendQueuedIdeographicSpaceSegment = (source: LayoutTextSeg): void =>
    performAppendQueuedIdeographicSpaceSegment(operationState, source);

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => performTabFollowWidth(operationState, q);

  /** Resolve the registered decimal alignment point independently of run/style
   * seams so both LTR and mirrored bidi tab paths consume one source boundary. */
  const decimalAlignmentPoint = performDecimalAlignmentPoint;

  const decimalAlignmentPrefixWidth = (segments: readonly LayoutSeg[]): number | undefined =>
    performDecimalAlignmentPrefixWidth(operationState, segments);

  const tabFollowingMetrics = (): Readonly<{
    totalWidth: number;
    decimalPrefixWidth?: number;
  }> => performTabFollowingMetrics(operationState);

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph, the new line is empty but still occupies one line
  // height. Track the trailing break so it can be flushed after the loop.

  // Establish the first line's wrap window now that the content queue exists.

  /**
   * Return the widest grapheme-safe UTF-16 prefix that fits an emergency
   * break band. This is the single authority used whether the overlong token
   * starts on an empty line or consumes the useful remainder of the current
   * line. Keeping both cases here prevents measurement and retained paint
   * partitions from drifting apart.
   */
  const emergencyTextSplit = (
    segment: LayoutTextSeg,
    available: number,
    forceAtLeastOne = true,
  ): number => performEmergencyTextSplit(operationState, segment, available, forceAtLeastOne);

  /** Select the last semantic URL candidate that fits the prospective band.
   * Every candidate is evaluated independently: signed character spacing can
   * make prefix advances non-monotone, and snapToChars must include the current
   * line's active script block rather than treating the prefix in isolation. */
  const explicitTextSplit = (segment: LayoutTextSeg, available: number): number =>
    performExplicitTextSplit(operationState, segment, available);

  const queueEmergencyTail = (segment: LayoutTextSeg, split: number): void =>
    performQueueEmergencyTail(operationState, segment, split);

  /**
   * Move a legal suffix of the current line ahead of a segment whose first
   * glyph is forbidden at line start. This is the single cross-run 追い出し
   * authority for both the CJK and SEA overflow paths.
   *
   * A source seam marked by `hardJoinPrev` is indivisible: moving the complete
   * following segment would strand its owner on the previous line. Likewise,
   * a split at either edge of an authored no-break range is not legal. When
   * either constraint blocks retraction, the caller must keep the forbidden
   * leader on the current line instead of weakening the authored constraint.
   */
  const retractCurrentLineForLeadingKinsoku = (next: LayoutTextSeg): CrossRunKinsokuRetraction =>
    performRetractCurrentLineForLeadingKinsoku(operationState, next);

  /** Keep one otherwise-forbidden line-start grapheme with the current line.
   * This is the only legal fallback when cross-run retraction would split an
   * authored hard/no-break group. Reprocessing the tail repeats the rule for a
   * sequence of forbidden leaders while guaranteeing grapheme-safe progress. */
  const keepLeadingKinsokuWithCurrentLine = (
    segment: LayoutTextSeg,
    h: number,
    asc: number,
    desc: number,
  ): boolean => performKeepLeadingKinsokuWithCurrentLine(operationState, segment, h, asc, desc);

  operationState = {
    ...input,
    reservePrefixWork: createPrefixWorkBudget(),
    probeHeights,
    probeFloors,
    preparedFloatWrap,
    breakerState,
    sameLatinSpaceFace,
    materializeLatinSpaceCompression,
    snapPitchPx,
    lineHeadRequirement,
    startLine,
    forcedPlacement,
    minimalLegalTextWidth,
    availW,
    fitsMeasuredWidth,
    bidiCustomStopsPx,
    bidiIntervalPx,
    flush,
    prospectiveSnapAdvance,
    addToLine,
    admitLineMetric,
    measurement,
    effectiveFontPx,
    measureText,
    verticalInkExtra,
    setMeasureFont,
    endBoundary,
    segNaturalAdvance,
    standaloneSnapAdvance,
    segAdvance,
    strNaturalAdvance,
    eastAsianSnapCellCount,
    strAdvance,
    fitHomogeneousLatinSpaces,
    mixedSpaceRequirement,
    markMixedSpacesCompressed,
    textSegmentBox,
    appendQueuedIdeographicSpaceSegment,
    tabFollowWidth,
    decimalAlignmentPoint,
    decimalAlignmentPrefixWidth,
    tabFollowingMetrics,
    emergencyTextSplit,
    explicitTextSplit,
    queueEmergencyTail,
    retractCurrentLineForLeadingKinsoku,
    keepLeadingKinsokuWithCurrentLine,
  };
  // The first numbered line's body indent may be resolved against its float
  // window (placeNumberingFirstLine); every first-line reader sees it.
  Object.defineProperty(operationState, 'firstIndent', {
    get: () => breakerState.firstIndentOverride ?? input.firstIndent,
    enumerable: true,
  });

  startLine(lineHeadRequirement());

  iterateBreakOpportunities(operationState, {
    captureGapSnapshot: () => performCaptureGapSnapshot(operationState, false),
    rejectGap: (rejection) => performRejectGap(operationState, rejection),
  });

  if (breakerState.currentLine.length > 0) flush();
  // Trailing <w:br/>: emit the empty line it opened (§17.3.3.1).
  else if (breakerState.trailingBreakFontSize !== null) flush(breakerState.trailingBreakFontSize);

  // Gap fragments on a physical baseline share the tallest line metrics. The
  // next convergence pass probes every fragment with that same band, so a
  // taller later gap cannot silently collide with a polygon above/below it.
  for (let start = 0; start < breakerState.lines.length;) {
    let end = start + 1;
    const first = breakerState.lines[start];
    while (end < breakerState.lines.length
      && breakerState.lines[end].physicalLineIndex === first.physicalLineIndex) end += 1;
    const last = breakerState.lines[end - 1];
    // Metrics accumulate until horizontal continuation ends. Publish the final
    // physical union, including object leading and position ownership, to all
    // fragments; each segment is visited once, independent of gap count.
    const { segments: _segments, xOffset: _x, availWidth: _width,
      marginExtension: _extension, consumedEnd: _end, endsWithBreak: _break, ...metrics } = last;
    let positionReference: number | undefined;
    for (let index = start; index < end; index += 1) {
      for (const segment of breakerState.lines[index].segments) {
        if ('isTab' in segment) continue;
        const position = 'text' in segment && segment.positionExtendsLineBox !== false
          ? segment.position ?? 0 : 0;
        positionReference = positionReference === undefined ? position
          : positionReference === position ? positionReference : 0;
      }
    }
    for (let index = start; index < end; index += 1) {
      const line = breakerState.lines[index];
      Object.assign(line, metrics);
      line.endsWithBreak = last.endsWithBreak;
      for (const segment of line.segments) {
        if ('text' in segment) segment.lineRelativePosition =
          wordUniformRunPositionPaintPt(segment.position ?? 0, positionReference ?? 0);
      }
    }
    start = end;
  }
  finalizeRetainedLineShapes(breakerState.lines, widthPolicy, measureText);

  return breakerState.lines;
}
