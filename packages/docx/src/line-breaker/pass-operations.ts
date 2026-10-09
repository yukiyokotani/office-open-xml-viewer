import { textBreakOffsetAt, textBreakOffsets } from './text-break-window.js';
import { candidateUnit, lineGapModel, type GapSegment } from './line-gaps.js';
import { LineMeasurementAdapter } from './measurement-adapter.js';
import { enumerateGaps, graphemeClusterOffsets, kinsokuAdjustedSplit } from '@silurus/ooxml-core';
import {
  prepareFloatWrap,
  computePreparedLineFloatWindow,
  type PreparedFloatWrap,
} from '../float-layout.js';
import { calcEffectiveFontPx, EAST_ASIAN_RE, independentTextShapeRequest, sliceTextShapeRequest } from '../layout/text.js';
import {
  wordSnapToCharsEastAsianCellCount,
  wordIdeographicSpaceLineEndAllowanceCount,
  wordJustifiedInterwordCompressionFactor,
} from '../layout/line-compatibility.js';
import {
  type LayoutImageSeg,
  type LayoutMathSeg,
  type LayoutSeg,
  type LayoutTabSeg,
  type LayoutTextSeg,
  type LineBoundary,
} from './model.js';
import { createLineBreakerState, type GapTransaction, type GapWindow } from './break-queue.js';
import {
  commitMixedLineItem,
  createMixedSpaceState,
  performSettleMixedSpaces,
  type MixedSpaceCandidate,
} from './mixed-space-fit.js';
import { applyBidiTabPostPass } from './tabs.js';
import {
  eastAsianGridCountSinglePx,
  measuredLineMetrics,
  nativeCanvasLineRatio,
} from './line-metrics.js';
import {
  RESET_SLICED_TEXT_MEASUREMENT,
  charScaleFactor,
  segLetterSpacingPx,
  charSpacingDeltaPx,
  protectedNoBreakOffsets,
  hardJoinPrefixEnd,
  legalTextSplitAtOrBefore,
  segAdvanceWidth,
  segmentCharacterGridDeltaPx,
  slicedPunctuationCompressions,
  slicedTextMetadata,
  snapToCharsAllocatedWidthPx,
  snapToCharsClass,
} from './advance.js';
import {
  buildFont,
  segmentEastAsiaFloorSingleLinePx,
  segmentIntendedSingleLinePx,
} from './font-routes.js';
import { rubyAscentReservePx } from './ruby-metrics.js';
import { fitCJKPrefix, hasEastAsianVisiblePredecessor } from './fit-search.js';
import { rebaseSeaBreaks, hasCJKBreakOpportunity } from './text-runs.js';
import {
  keepLeadingKinsoku,
  retractLeadingKinsoku,
  type CrossRunKinsokuRetraction,
} from './kinsoku.js';

import type { LineBreakerPassInput } from './pass-driver.js';

export interface PassOperationState extends LineBreakerPassInput {
  readonly reservePrefixWork: (utf16Units: number) => void;
  readonly breakerState: ReturnType<typeof createLineBreakerState>;
  readonly sameLatinSpaceFace: (candidate: LayoutTextSeg, reference: LayoutTextSeg) => boolean;
  readonly materializeLatinSpaceCompression: () => void;
  readonly snapPitchPx: number | null;
  readonly lineHeadRequirement: (boundary?: LineBoundary) => number;
  readonly startLine: (requirement?: number) => void;
  /** Report a forced placement of the unit starting at `unitStart` in the
   * current line; throws LineGapRejection inside a narrowed float gap. */
  readonly forcedPlacement: (requiredWidth: number, unitStart?: number) => void;
  /** Smallest legal line-head advance of a text segment under placement rules. */
  readonly minimalLegalTextWidth: (segment: LayoutTextSeg) => number;
  readonly availW: () => number;
  readonly fitsMeasuredWidth: (used: number, available: number) => boolean;
  readonly bidiCustomStopsPx: {
    pos: number;
    alignment: 'left' | 'start' | 'center' | 'right' | 'end' | 'decimal' | 'bar' | 'clear' | 'num';
    leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  }[];
  readonly bidiIntervalPx: number;
  readonly flush: (forceHeight?: number, brTerminated?: boolean, nextStart?: LineBoundary) => void;
  readonly prospectiveSnapAdvance: (s: LayoutTextSeg, naturalWidth: number) => number;
  readonly addToLine: (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
  ) => void;
  readonly measurement: LineMeasurementAdapter;
  readonly effectiveFontPx: (s: LayoutTextSeg) => number;
  readonly measureText: (s: LayoutTextSeg, clusterGeometry?: boolean) => TextMetrics;
  readonly verticalInkExtra: (s: LayoutTextSeg, text: string) => number;
  readonly setMeasureFont: (font: string) => void;
  readonly endBoundary: LineBoundary;
  readonly segNaturalAdvance: (s: LayoutTextSeg) => number;
  readonly standaloneSnapAdvance: (s: LayoutTextSeg, naturalWidth: number) => number;
  readonly segAdvance: (s: LayoutTextSeg) => number;
  readonly strNaturalAdvance: (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number;
  readonly eastAsianSnapCellCount: (s: LayoutTextSeg) => number;
  readonly strAdvance: (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number;
  readonly fitHomogeneousLatinSpaces: (next: LayoutTextSeg, nextFitWidth: number) => boolean;
  /** WORD_COMPRESSED_SPACE_LINE_FIT (mixed-script lines); see mixed-space-fit.ts. */
  readonly mixedSpaceRequirement: (candidate: MixedSpaceCandidate) => number | undefined;
  readonly markMixedSpacesCompressed: () => void;
  readonly textSegmentBox: (
    s: LayoutTextSeg,
  ) => Readonly<{ width: number; height: number; ascent: number; descent: number }>;
  readonly appendQueuedIdeographicSpaceSegment: (source: LayoutTextSeg) => void;
  readonly tabFollowWidth: (q: LayoutSeg) => number;
  readonly decimalAlignmentPoint: (
    segments: readonly LayoutSeg[],
  ) => Readonly<{ segmentIndex: number; charOffset: number }> | null;
  readonly decimalAlignmentPrefixWidth: (segments: readonly LayoutSeg[]) => number | undefined;
  readonly tabFollowingMetrics: () => Readonly<{ totalWidth: number; decimalPrefixWidth?: number }>;
  readonly emergencyTextSplit: (
    segment: LayoutTextSeg,
    available: number,
    forceAtLeastOne?: boolean,
  ) => number;
  readonly explicitTextSplit: (segment: LayoutTextSeg, available: number) => number;
  readonly queueEmergencyTail: (segment: LayoutTextSeg, split: number) => void;
  readonly retractCurrentLineForLeadingKinsoku: (next: LayoutTextSeg) => CrossRunKinsokuRetraction;
  readonly keepLeadingKinsokuWithCurrentLine: (
    segment: LayoutTextSeg,
    h: number,
    asc: number,
    desc: number,
  ) => boolean;
  readonly probeHeights: readonly number[] | null;
  readonly probeFloors: readonly number[] | null;
  readonly preparedFloatWrap?: PreparedFloatWrap;
}

export function performSameLatinSpaceFace(
  candidate: LayoutTextSeg,
  reference: LayoutTextSeg,
): boolean {
  return (
    candidate.latinSpaceCompressionEligible === true &&
    !candidate.verticalRun &&
    !reference.verticalRun &&
    !candidate.tateChuYoko &&
    !reference.tateChuYoko &&
    candidate.latinSpaceAverageWidthRatio === reference.latinSpaceAverageWidthRatio &&
    candidate.fontRoute?.fingerprint === reference.fontRoute?.fingerprint &&
    candidate.fontFamily === reference.fontFamily &&
    candidate.fontSize === reference.fontSize &&
    candidate.bold === reference.bold &&
    candidate.italic === reference.italic &&
    (candidate.charScale == null || candidate.charScale === 1) &&
    (reference.charScale == null || reference.charScale === 1) &&
    candidate.kerning === reference.kerning &&
    candidate.widthBalanceGridDeltaFactor === reference.widthBalanceGridDeltaFactor &&
    !candidate.rtl &&
    candidate.fitTextRegionIndex === undefined
  );
}

export function performMaterializeLatinSpaceCompression(operationState: PassOperationState): void {
  const { breakerState } = operationState;

  for (let index = 0; index < breakerState.latinAppliedGapCount; index += 1) {
    const gap = breakerState.latinLineGaps[index];
    gap.measuredWidth -= breakerState.latinAppliedPerGap;
    gap.latinSpaceCompressionPx = breakerState.latinAppliedPerGap;
  }
  breakerState.latinAppliedGapCount = 0;
  breakerState.latinAppliedPerGap = 0;
}

/**
 * WORD_FLOAT_GAP_FLOW (#1670): one physical line may be split into fragments,
 * one per free gap, filled in reading order. ECMA-376 §20.4.2.17–.19 permits
 * text on both sides of a square/tight/through object but does not say how a
 * line is partitioned. The registered rule selects the first gap in reading
 * order whose width admits the content placed there, then continues into
 * later gaps on the same baseline.
 *
 * Admission is a placement transaction, not a separate width predictor. A
 * fragment narrowed by an exclusion is filled by ordinary placement. If any
 * placement path would force content into it (an emergency split, an illegal
 * kinsoku split, or an advance past the fragment edge), the fragment rolls
 * back to its first placement and the search continues with a requirement
 * strictly larger than the rejected gap. A full paragraph band (no exclusion
 * narrows it) keeps the established forced-placement behavior, so progress is
 * guaranteed once a unit fits no narrowed gap.
 */
export class LineGapRejection extends Error {
  constructor(
    /** Lower bound of the forced unit's advance, from placement's own measurement. */
    readonly requiredWidth: number,
    /** Complete units precede the forced unit: end the fragment before it. */
    readonly stopBefore?: LineBoundary,
  ) {
    super('A narrowed float gap cannot admit the line-head unit');
    this.name = 'LineGapRejection';
  }
}

/** Seed requirement for a line head. Content needs no more than the solver's
 * numeric floor: placement decides admission. A mark-only head (no inline
 * content remains) keeps `word-empty-mark-float-side-gap`. */
export function performLineHeadRequirement(
  operationState: PassOperationState,
  boundary?: LineBoundary,
): number {
  const { wrapCtx, segs, breakerState, scale } = operationState;
  if (!wrapCtx) return 0;
  const start = boundary?.segIndex ?? 0;
  // Source boundaries must start directly at their index: walking the already
  // consumed prefix at every fragment would make long paragraphs quadratic.
  function* sourceCandidates(): IterableIterator<LayoutSeg> {
    for (let index = start; index < segs.length; index += 1) yield segs[index];
  }
  const candidates = boundary ? sourceCandidates() : breakerState.queue;
  let mark: LayoutTextSeg | undefined;
  for (const candidate of candidates) {
    if (('text' in candidate && !candidate.metricOnly && candidate.text.length > 0)
      || ('imagePath' in candidate && !candidate.anchor) || 'math' in candidate
      || 'isTab' in candidate || 'lineBreak' in candidate) return 0;
    if ('text' in candidate && candidate.metricOnly) mark ??= candidate;
  }
  if (boundary ? start >= segs.length : breakerState.queue.length === 0) return 0;
  return wrapCtx.paragraphMarkLineStartWidth ?? (mark ? mark.fontSize * scale : 0);
}

/** Smallest prefix placement may legally put at a line head: the first SEA
 * boundary, the first CJK split legal under kinsoku and protected ranges, the
 * first URL syntax opportunity, a hard seam's protected prefix, else the whole
 * segment (segments are already delimited at ordinary opportunities). Used
 * only as the lower bound reported when a narrowed gap rejects its head. */
export function performMinimalLegalTextWidth(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
): number {
  const { kinsoku, strAdvance, baseRtl } = operationState;
  const text = segment.text;
  let end = text.length;
  if (segment.fitTextRegionIndex === undefined && !segment.ruby && !segment.tateChuYoko) {
    const protectedOffsets = protectedNoBreakOffsets(segment);
    if (segment.hardJoinPrev === true) {
      end = hardJoinPrefixEnd(segment) ?? text.length;
    } else if (segment.seaBreaks !== undefined) {
      end = segment.seaBreaks.find((offset) => offset > 0 && offset < text.length) ?? text.length;
    } else if (hasCJKBreakOpportunity(text)) {
      const characters = [...text];
      for (let count = 1; count < characters.length; count += 1) {
        if (kinsokuAdjustedSplit(characters, count, kinsoku, 0) !== count) continue;
        const utf16 = characters.slice(0, count).join('').length;
        if (legalTextSplitAtOrBefore(segment, utf16, 1) === utf16) {
          end = utf16;
          break;
        }
      }
    } else {
      end = [...textBreakOffsets(segment.explicitBreaks)].find((offset) =>
        offset > 0 && offset < text.length && !protectedOffsets.has(offset)) ?? text.length;
    }
  }
  const prefix = text.slice(0, end);
  // The fit tests exclude a collapsible U+0020 suffix except in an RTL line.
  return strAdvance(segment, baseRtl ? prefix : prefix.replace(/ +$/u, ''));
}

/** Vertical allocation belongs to the physical line; only a failed gap
 * continuation opens a new line. Horizontal queue/fit state resets per gap. */
function resetPhysicalLineMetrics({ breakerState }: PassOperationState): void {
  breakerState.lineHeight = 0;
  breakerState.lineAscent = 0;
  breakerState.lineDescent = 0;
  breakerState.lineIntendedSingle = 0;
  breakerState.lineHasInlinePicture = false;
  breakerState.linePictureMarkSingle = 0;
  breakerState.lineGridCountSingle = 0;
  breakerState.lineLatinGridCountSingle = 0;
  breakerState.lineVisibleAscent = 0;
  breakerState.lineVisibleDescent = 0;
  breakerState.lineVisibleIntendedSingle = 0;
  breakerState.lineHasVisibleMetrics = false;
  breakerState.lineHasRuby = false;
  breakerState.lineEastAsian = false;
  breakerState.positionReferencePt = undefined;
  breakerState.firstPositioned = undefined;
  breakerState.uniformPositionEligible = true;
}

/** Exclusion probe band of a physical line. The initial pass measures without
 * exclusions. Afterwards every physical line is probed in the pass that
 * reaches it: a line the previous pass did not produce uses the last resolved
 * physical allocation, and the next pass replaces it with its own. With fixed
 * metrics, the result converges in three passes (measure, resolve, confirm)
 * regardless of how many physical lines or gaps exclusions create; varying
 * metrics retain the fail-closed pass guard. */
function physicalProbeHeight(probeHeights: readonly number[] | null, index: number): number | undefined {
  if (!probeHeights || probeHeights.length === 0) return undefined;
  return probeHeights[index] ?? probeHeights[probeHeights.length - 1];
}

function openPhysicalLine(operationState: PassOperationState): void {
  const { breakerState } = operationState;
  resetPhysicalLineMetrics(operationState);
  if (breakerState.lines.length > 0) breakerState.physicalLineIndex += 1;
}

/** Width between the paragraph's trailing indent and the text margin (zero
 * when a negative indent already reaches past the margin). */
function marginExtensionWidth({ marginRightPx, maxWidth }: PassOperationState): number {
  return Math.max(0, marginRightPx - maxWidth);
}

export function performStartLine(operationState: PassOperationState, requirement: number = 0): void {
  const { breakerState, maxWidth, wrapCtx, firstIndent, probeFloors } = operationState;

  breakerState.snapBlock = null;
  breakerState.lineXOffset = 0;
  breakerState.lineMaxWidth = maxWidth;
  breakerState.lineMarginExtension = marginExtensionWidth(operationState);
  breakerState.gapTransaction = null;
  const cursor = breakerState.fragmentCursor;
  breakerState.fragmentCursor = null;
  // Every gap of this physical line uses the same observed band. New gaps
  // are horizontal placements, so they need no additional convergence pass.
  // Without an observed band (first pass), the line is measured unconstrained.
  if (!wrapCtx || physicalProbeHeight(probeFloors, breakerState.physicalLineIndex) === undefined) {
    openPhysicalLine(operationState);
    return;
  }
  const transaction: GapTransaction = {
    cursor,
    requestTopY: breakerState.currentLineTopY,
    requirement: requirement + (breakerState.isFirst ? Math.max(0, firstIndent) : 0),
    window: null,
    narrowed: false,
    endsAtExclusion: false,
    snapshot: null,
    stopBefore: null,
  };
  breakerState.gapTransaction = transaction;
  if (!cursor) openPhysicalLine(operationState);
  placeLineWindow(operationState, transaction, null);
}

/** Search the first admissible window for the transaction's requirement:
 * a later gap on the continuation baseline, else a new physical line. */
function placeLineWindow(
  operationState: PassOperationState,
  transaction: GapTransaction,
  rejected: GapWindow | null,
): void {
  const { breakerState, maxWidth, wrapCtx, baseRtl, firstIndent, probeFloors, preparedFloatWrap } =
    operationState;
  if (!wrapCtx) return;
  // §17.3.1.12 removes a hanging indent from the paragraph's first-line
  // indentation, not from an object's exclusion (§20.4.2.17–.19). Query the
  // expanded first-line band before subtracting floats; applying the hanging
  // offset after that subtraction would move text back inside an object.
  // Positive first-line indents remain an inset within the selected window.
  // The measured-line contract still carries the authored firstIndent: restore
  // it in the returned width so fit/tab arithmetic and planLine consume the
  // same safe window, including the mirrored logical start of RTL paragraphs.
  const hangingOffset = breakerState.isFirst ? Math.min(0, firstIndent) : 0;
  const lineBandX = wrapCtx.paraX + (baseRtl ? 0 : hangingOffset);
  const lineBandWidth = maxWidth - hangingOffset;
  const reference = {
    xLeftPt: wrapCtx.referenceXPt ?? wrapCtx.paraX,
    xRightPt: (wrapCtx.referenceXPt ?? wrapCtx.paraX) + (wrapCtx.referenceWidthPt ?? maxWidth),
    readingDirection: wrapCtx.readingDirection ?? (baseRtl ? 'rtl' : 'ltr'),
  } as const;
  const query = (topY: number, x: number, width: number, height: number,
    requiredWidth = transaction.requirement) => {
    if (wrapCtx.lineWindow) {
      const win = wrapCtx.lineWindow({
        topYPt: topY, minimumStartWidthPt: requiredWidth,
        squareMinimumStartWidthPt: requiredWidth, probeHeightPt: height,
        paragraphXPt: x, maximumWidthPt: width,
        columnXPt: wrapCtx.columnXPt, columnWidthPt: wrapCtx.columnWidthPt,
      });
      return { topY: win.topYPt, xOffset: win.xOffsetPt, maxWidth: win.maximumWidthPt };
    }
    return computePreparedLineFloatWindow(
      topY, requiredWidth, height, x, width,
      preparedFloatWrap ?? prepareFloatWrap(wrapCtx.floats),
      wrapCtx.columnXPt, wrapCtx.columnXPt + wrapCtx.columnWidthPt,
      reference, requiredWidth,
    );
  };
  let accepted: GapWindow & { narrowed: boolean } | null = null;
  const cursor = transaction.cursor;
  if (cursor) {
    const probeH = physicalProbeHeight(probeFloors, breakerState.physicalLineIndex);
    const left = baseRtl ? lineBandX : cursor.right;
    const right = baseRtl ? cursor.left : lineBandX + lineBandWidth;
    if (probeH !== undefined && right > left) {
      const next = query(cursor.topY, left, right - left, probeH);
      if (next.topY === cursor.topY && next.maxWidth >= transaction.requirement) {
        // A strictly smaller unvisited band gives monotonic horizontal
        // progress; the column and largest-side reference never shrink.
        accepted = {
          topY: next.topY,
          xOffset: left - wrapCtx.paraX + next.xOffset,
          maxWidth: next.maxWidth,
          solverWidth: next.maxWidth,
          narrowed: true,
        };
      }
    }
    if (!accepted) {
      // No later gap on this baseline admits the head: the physical line ends.
      transaction.cursor = null;
      openPhysicalLine(operationState);
      breakerState.currentLineTopY = transaction.requestTopY;
    }
  }
  if (!accepted) {
    const probeH = physicalProbeHeight(probeFloors, breakerState.physicalLineIndex);
    if (probeH === undefined) {
      breakerState.lineXOffset = 0;
      breakerState.lineMaxWidth = maxWidth;
      breakerState.lineMarginExtension = marginExtensionWidth(operationState);
      transaction.window = null;
      transaction.narrowed = false;
      transaction.endsAtExclusion = false;
      transaction.snapshot = null;
      return;
    }
    const win = query(transaction.requestTopY, lineBandX, lineBandWidth, probeH);
    accepted = {
      topY: win.topY,
      xOffset: win.xOffset,
      maxWidth: win.maxWidth + hangingOffset,
      solverWidth: win.maxWidth,
      narrowed: win.xOffset !== 0 || win.maxWidth !== lineBandWidth,
    };
  }
  // After a rejection the requirement exceeds the rejected width. A window
  // still narrower than it is not admitted by the solver's exact comparison
  // (it differs only by representable rounding) or comes from a boundary
  // that does not honor the requirement. Accept it without a transaction:
  // forced placement is then the established behavior, and the search ends.
  if (rejected && accepted.narrowed && accepted.solverWidth < transaction.requirement) {
    accepted = { ...accepted, narrowed: false };
  }
  breakerState.currentLineTopY = accepted.topY;
  breakerState.lineXOffset = accepted.xOffset;
  breakerState.lineMaxWidth = accepted.maxWidth;
  // A trailing-indent extension exists only beside an unnarrowed band, and
  // only where no exclusion intersects it on this line (§20.4.2.17–.19).
  const extension = marginExtensionWidth(operationState);
  const extensionProbeH = physicalProbeHeight(probeFloors, breakerState.physicalLineIndex);
  if (accepted.narrowed || extension <= 0 || baseRtl) {
    breakerState.lineMarginExtension = 0;
  } else if (extensionProbeH === undefined) {
    breakerState.lineMarginExtension = extension;
  } else {
    const free = query(accepted.topY, wrapCtx.paraX + maxWidth, extension, extensionProbeH, extension);
    breakerState.lineMarginExtension = free.topY === accepted.topY && free.xOffset === 0
      && free.maxWidth >= extension ? extension : 0;
  }
  transaction.window = accepted;
  transaction.narrowed = accepted.narrowed;
  // Absolute line-end edge versus the paragraph band edge, in reading order.
  const windowStart = wrapCtx.paraX + accepted.xOffset;
  const windowEnd = windowStart + accepted.maxWidth;
  transaction.endsAtExclusion = baseRtl
    ? windowStart > lineBandX + 1e-9
    : windowEnd < lineBandX + lineBandWidth - 1e-9;
  transaction.snapshot = null;
  if (accepted.narrowed) performCaptureGapSnapshot(operationState, true);
}

/** Record the rollback image of a narrowed fragment before its first unit.
 * Called when the window is accepted (the iterator may hold the segment being
 * processed) and again at each iterator step while the fragment is empty, so
 * content re-queued by the preceding flush (e.g. kinsoku retraction) is kept. */
export function performCaptureGapSnapshot(
  operationState: PassOperationState,
  includeInHand: boolean,
): void {
  const { breakerState } = operationState;
  const transaction = breakerState.gapTransaction;
  if (!transaction?.narrowed) return;
  const {
    lines, queue, currentLine: _line, latinLineGaps: _gaps, gapTransaction: _transaction,
    inHand, snapBlock, ...scalars
  } = breakerState;
  transaction.snapshot = {
    scalars: { ...scalars },
    snapBlock: snapBlock ? { ...snapBlock } : null,
    linesLength: lines.length,
    queue: queue.snapshot(includeInHand ? inHand : undefined),
  };
}

/** Called by every placement path that would force its unit into the
 * current fragment. On a full band this is the established behavior; in a
 * narrowed gap the fragment is rejected (or ended before the unit). */
export function performForcedPlacement(
  operationState: PassOperationState,
  requiredWidth: number,
  unitStart = 0,
): void {
  const { breakerState } = operationState;
  const transaction = breakerState.gapTransaction;
  if (!transaction?.narrowed || !transaction.snapshot) return;
  // Inkless items before the unit (anchor characters, mark metrics) travel
  // with it; only inked content can end the fragment before the unit.
  let committed = false;
  for (let index = 0; index < unitStart && index < breakerState.currentLine.length; index += 1) {
    if (!isInklessLineItem(breakerState.currentLine[index])) {
      committed = true;
      break;
    }
  }
  const lead = committed ? breakerState.currentLine[unitStart] : undefined;
  throw new LineGapRejection(requiredWidth, lead?.src ? { ...lead.src } : undefined);
}

function isInklessLineItem(item: LayoutSeg): boolean {
  return ('text' in item && item.metricOnly === true) || ('imagePath' in item && Boolean(item.anchor));
}

/** Lower bound of a queued unit's line-head advance, from placement's measures. */
function headUnitLowerBound(operationState: PassOperationState, segment: LayoutSeg | undefined): number {
  if (!segment) return 0;
  if ('text' in segment) return operationState.minimalLegalTextWidth(segment);
  if ('imagePath' in segment) return segment.anchor ? 0 : segment.widthPt * operationState.scale;
  if ('math' in segment) return segment.measuredWidth;
  return 0;
}

/** Roll a rejected fragment back and continue the search. */
export function performRejectGap(
  operationState: PassOperationState,
  rejection: LineGapRejection,
): void {
  const { breakerState, firstIndent, maxWidth } = operationState;
  const transaction = breakerState.gapTransaction;
  const snapshot = transaction?.snapshot;
  if (!transaction || !snapshot || !transaction.window) throw rejection;
  Object.assign(breakerState, snapshot.scalars);
  breakerState.snapBlock = snapshot.snapBlock
    ? { ...(snapshot.snapBlock as NonNullable<typeof breakerState.snapBlock>) }
    : null;
  breakerState.lines.length = snapshot.linesLength;
  breakerState.queue.restore(snapshot.queue);
  breakerState.currentLine = [];
  breakerState.latinLineGaps = [];
  breakerState.mixedSpace = createMixedSpaceState();
  breakerState.inHand = undefined;
  // A replay ends the fragment before the forced unit. If the replay cannot
  // reach that source boundary, the whole fragment is rejected instead.
  if (rejection.stopBefore && !transaction.stopBefore) {
    transaction.stopBefore = rejection.stopBefore;
    return;
  }
  transaction.stopBefore = null;
  const rejected = transaction.window;
  const required = rejection.requiredWidth
    + (breakerState.isFirst ? Math.max(0, firstIndent) : 0);
  const hangingOffset = breakerState.isFirst ? Math.min(0, firstIndent) : 0;
  // The forced unit did not fit this gap, so its advance exceeds the width
  // the solver admitted. A lower bound that does not (non-monotone advances)
  // sends the unit to the full band instead of creeping through gaps.
  transaction.requirement = required > rejected.solverWidth
    ? Math.max(transaction.requirement, required)
    : maxWidth - hangingOffset;
  placeLineWindow(operationState, transaction, rejected);
}

export function performAvailW(operationState: PassOperationState) {
  const { breakerState, firstIndent, widthPolicy } = operationState;
  return widthPolicy !== 'bounded'
    ? Number.POSITIVE_INFINITY
    : breakerState.lineMaxWidth - (breakerState.isFirst ? firstIndent : 0);
}

export function performFitsMeasuredWidth(
  operationState: PassOperationState,
  used: number,
  available: number,
): boolean {
  const { breakerState, firstIndent, widthPolicy } = operationState;

  // Intrinsic AutoFit widths can become the exact final line width. The
  // margin subtraction and point/pixel round trip may differ by an ulp from
  // the same shaped advance. Use the existing grapheme-fit numerical epsilon
  // (below), not an Office width allowance, before forcing an emergency split.
  if (used <= available + 1e-9) return true;
  if (
    widthPolicy !== 'bounded' ||
    !Number.isFinite(used) ||
    !Number.isFinite(breakerState.lineMaxWidth)
  ) {
    return false;
  }
  return used + (breakerState.isFirst ? firstIndent : 0) <= breakerState.lineMaxWidth;
}

export function performFlush(
  operationState: PassOperationState,
  forceHeight?: number,
  brTerminated = false,
  nextStart?: LineBoundary,
) {
  const {
    breakerState,
    materializeLatinSpaceCompression,
    lineHeadRequirement,
    startLine,
    bidiCustomStopsPx,
    bidiIntervalPx,
    endBoundary,
    strAdvance,
    decimalAlignmentPoint,
    firstIndent,
    scale,
    wrapCtx,
    tabOriginPx,
    marginRightPx,
    baseRtl,
  } = operationState;

  // An anchor character or paragraph-mark metric has no advance; it moves
  // with the unit that follows it. A narrowed fragment holding only such
  // items cannot end before that unit: the unit is its head (forced here).
  // Tabs and breaks are pen/line controls, not such units.
  const followingUnit = headUnitLowerBound(operationState, breakerState.inHand);
  if (!brTerminated && nextStart !== undefined && followingUnit > 0
    && breakerState.currentLine.length > 0 && breakerState.currentLine.every(isInklessLineItem)) {
    operationState.forcedPlacement(followingUnit);
  }
  materializeLatinSpaceCompression();
  performSettleMixedSpaces(operationState);
  breakerState.currentWidth += applyBidiTabPostPass({
    baseRtl,
    currentLine: breakerState.currentLine,
    marginRightPx,
    maxWidth: operationState.maxWidth,
    lineXOffset: breakerState.lineXOffset,
    lineMaxWidth: breakerState.lineMaxWidth,
    isFirst: breakerState.isFirst,
    firstIndent,
    tabOriginPx,
    bidiCustomStopsPx,
    bidiIntervalPx,
    decimalAlignmentPoint,
    strAdvance,
  });
  // §17.3.2.24 defines `position` relative to surrounding non-positioned
  // text. A line whose every metric-bearing item shares the same inherited
  // position has no differently-positioned peer to pin the resulting line
  // box to one side. `word-uniform-run-position-leading` owns the compatibility
  // placement of that box around the glyphs. Keep mixed
  // lines relative to zero so their authored displacement and ink union
  // remain unchanged. Images/math
  // provide a zero-position reference; tabs do not contribute vertical
  // metrics. The fixed drop-cap path intentionally keeps its paint-only
  // lowering and therefore opts out of this normalization.
  for (const segment of breakerState.currentLine) {
    if ('isTab' in segment) continue;
    const position = 'text' in segment && segment.positionExtendsLineBox !== false
      ? (segment.position ?? 0) : 0;
    if (breakerState.positionReferencePt === undefined) breakerState.positionReferencePt = position;
    else if (breakerState.positionReferencePt !== position) breakerState.positionReferencePt = null;
    if (!('text' in segment)) {
      breakerState.uniformPositionEligible = false;
      continue;
    }
    breakerState.firstPositioned ??= segment;
    const first = breakerState.firstPositioned;
    breakerState.uniformPositionEligible &&= segment.text.length > 0
      && !segment.metricOnly && !segment.ruby && !segment.vertAlign
      && segment.positionExtendsLineBox !== false
      && segment.position === first.position
      && segment.fontFamily === first.fontFamily
      && segment.fontRoute?.fingerprint === first.fontRoute?.fingerprint
      && segment.bold === first.bold && segment.italic === first.italic
      && segment.fontSize === first.fontSize
      && segment.resolvedDesignDescentRatio === first.resolvedDesignDescentRatio
      && segment.referenceFontVerticalMetric === first.referenceFontVerticalMetric
      && segment.resolvedResourceVerticalMetric === first.resolvedResourceVerticalMetric;
  }
  const linePositionReferencePt = breakerState.positionReferencePt ?? 0;
  // §17.3.3.1 — the break is one run among the line's runs: its own size
  // participates in the line height but must not override a taller peer.
  const h =
    forceHeight !== undefined
      ? Math.max(breakerState.lineHeight, forceHeight)
      : breakerState.lineHeight || 10;
  // If the line has no measured content (empty/line-break line), synthesize
  // stable ascent/descent from the effective font size so wrap/baseline math
  // stays consistent with non-empty lines.
  const hasContent = breakerState.lineAscent > 0 || breakerState.lineDescent > 0;
  const asc = hasContent ? breakerState.lineAscent : h * scale * 0.8;
  const desc = hasContent ? breakerState.lineDescent : h * scale * 0.2;
  const visibleAscent = breakerState.lineHasVisibleMetrics ? breakerState.lineVisibleAscent : asc;
  const visibleDescent = breakerState.lineHasVisibleMetrics
    ? breakerState.lineVisibleDescent
    : desc;
  const visibleIntendedSingle = breakerState.lineHasVisibleMetrics
    ? breakerState.lineVisibleIntendedSingle
    : breakerState.lineIntendedSingle;
  const gridCountSingle =
    breakerState.lineGridCountSingle ||
    (breakerState.lineEastAsian
      ? eastAsianGridCountSinglePx(breakerState.lineIntendedSingle, h * scale)
      : asc + desc);
  const inlinePictureTextSingle = breakerState.lineHasInlinePicture
    ? Math.max(breakerState.lineIntendedSingle, breakerState.linePictureMarkSingle)
    : 0;
  // Only project that registered rule when every metric-bearing item is
  // visible text in one admitted face tuple. Canvas fallback geometry does
  // not reveal hhea descent, and mixed styles cannot share one descent
  // reserve. The rule uses face data, never a family-specific correction.
  const firstPositioned = breakerState.firstPositioned;
  const uniformPositionAuto =
    linePositionReferencePt !== 0 && breakerState.uniformPositionEligible
    && firstPositioned?.resolvedDesignDescentRatio != null
    && (firstPositioned.referenceFontVerticalMetric || firstPositioned.resolvedResourceVerticalMetric)
      ? {
          normalSinglePx: Math.max(
            asc + desc - Math.abs(linePositionReferencePt * scale),
            breakerState.lineIntendedSingle,
          ),
          positionPx: linePositionReferencePt * scale,
          designDescentPx:
            firstPositioned.resolvedDesignDescentRatio * firstPositioned.fontSize * scale,
        }
      : undefined;
  breakerState.lines.push({
    physicalLineIndex: breakerState.physicalLineIndex,
    ...(justifiedCompressionApplies(operationState)
      ? {
          justifiedCompressionPx: breakerState.justifiedCompressionPx,
          gapPlan: {
            ...lineGapModel(breakerState.currentLine.map(s => gapSegment(operationState, s))),
            // CJK/ideographic/SEA expansion has no proportional measurement.
            // Retain its previous opportunities; family selection lives in core.
            expansionGaps: enumerateGaps(breakerState.currentLine.map(s => ({
              text: 'text' in s && !s.metricOnly && s.fitTextRegionIndex === undefined && s.snapGridClass === undefined
                ? s.text : undefined,
            })), { lastDrawnSi: Infinity }).gaps,
          },
        } : {}),
    segments: breakerState.currentLine,
    height: h,
    ascent: asc,
    descent: desc,
    visibleAscent,
    visibleDescent,
    visibleIntendedSingle,
    intendedSingle: breakerState.lineIntendedSingle,
    latinGridCountSingle: breakerState.lineLatinGridCountSingle,
    ...(inlinePictureTextSingle > 0 ? { inlinePictureTextSingle } : {}),
    uniformPositionAuto,
    // Empty/synthetic East Asian lines use the same design-height rule as a
    // text run; their synthesized Canvas box must not reintroduce a
    // scale-dependent cell count.
    gridCountSingle,
    xOffset: breakerState.lineXOffset,
    availWidth: breakerState.lineMaxWidth,
    ...(breakerState.currentLine.some((segment) => 'isTab' in segment && segment.marginAllocation)
      ? { marginExtension: breakerState.lineMarginExtension } : {}),
    topY: wrapCtx ? breakerState.currentLineTopY : undefined,
    hasRuby: breakerState.lineHasRuby,
    eastAsian: breakerState.lineEastAsian,
    endsWithBreak: brTerminated,
    consumedEnd: nextStart ?? breakerState.queue.peek()?.src ?? endBoundary,
  });
  if (wrapCtx) {
    if (!brTerminated && nextStart !== undefined) {
      breakerState.fragmentCursor = {
        topY: breakerState.currentLineTopY,
        left: wrapCtx.paraX + breakerState.lineXOffset
          + (breakerState.isFirst ? Math.min(0, firstIndent) : 0),
        right: wrapCtx.paraX + breakerState.lineXOffset + breakerState.lineMaxWidth,
      };
    }
    // Use the physical allocation that supplied the exclusion probe. Local
    // fragment metrics can be smaller than a later gap or paragraph-wide ruby
    // reserve (§17.3.3.25); they cannot independently advance the next origin.
    breakerState.currentLineTopY += operationState.probeHeights?.[breakerState.physicalLineIndex]
      ?? wrapCtx.lineBoxH(
      asc,
      desc,
      breakerState.lineHasRuby,
      breakerState.lineIntendedSingle,
      breakerState.lineEastAsian,
      gridCountSingle,
      uniformPositionAuto,
      inlinePictureTextSingle,
      breakerState.lineLatinGridCountSingle,
    );
  }
  breakerState.currentLine = [];
  breakerState.currentWidth = 0;
  breakerState.justifiedGapModel = undefined;
  breakerState.justifiedCompressionPx = 0;
  breakerState.justifiedUnitEnd = undefined;
  breakerState.latinLineFace = undefined;
  breakerState.latinLineHomogeneous = true;
  breakerState.latinLineGaps = [];
  breakerState.latinUniformGapCapacity = undefined;
  breakerState.isFirst = false;
  startLine(lineHeadRequirement(nextStart));
}

export function performProspectiveSnapAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  naturalWidth: number,
): number {
  const { breakerState, snapPitchPx, eastAsianSnapCellCount, characterGrid } = operationState;

  const kind = snapToCharsClass(s, characterGrid);
  if (!kind || snapPitchPx == null) return naturalWidth;
  if (kind === 'eastAsia') {
    const cells = eastAsianSnapCellCount(s);
    return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx, cells);
  }
  if (breakerState.snapBlock?.kind === kind) {
    return (
      snapToCharsAllocatedWidthPx(
        breakerState.snapBlock.naturalWidthPx + naturalWidth,
        kind,
        snapPitchPx,
      ) - breakerState.snapBlock.allocatedWidthPx
    );
  }
  return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
}

export function performAddToLine(
  operationState: PassOperationState,
  s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
  w: number,
  h: number,
  asc: number,
  desc: number,
) {
  const {
    breakerState,
    sameLatinSpaceFace,
    materializeLatinSpaceCompression,
    snapPitchPx,
    effectiveFontPx,
    eastAsianSnapCellCount,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
  } = operationState;

  let committedWidth = w;
  if ('text' in s) {
    const kind = snapToCharsClass(s, characterGrid);
    const naturalWidth = s.snapGridNaturalWidthPx ?? w;
    if (kind && snapPitchPx != null) {
      s.snapGridClass = kind;
      s.snapGridNaturalWidthPx = naturalWidth;
      s.snapGridCellPitchPx = snapPitchPx;
      if (kind === 'eastAsia') {
        const cellCount = eastAsianSnapCellCount(s);
        committedWidth = snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx, cellCount);
        s.snapGridLeadingPadPx = 0;
        s.snapGridTrailingPadPx = committedWidth - naturalWidth;
        s.measuredWidth = committedWidth;
        breakerState.snapBlock = null;
      } else if (breakerState.snapBlock?.kind === kind) {
        const previousLeading = breakerState.snapBlock.first.snapGridLeadingPadPx ?? 0;
        const previousTrailing = breakerState.snapBlock.last.snapGridTrailingPadPx ?? 0;
        const combinedNatural = breakerState.snapBlock.naturalWidthPx + naturalWidth;
        const combinedAllocated = snapToCharsAllocatedWidthPx(combinedNatural, kind, snapPitchPx);
        const slack = combinedAllocated - combinedNatural;
        const leading = kind === 'latin' ? slack / 2 : 0;
        const trailing = slack - leading;
        breakerState.snapBlock.first.measuredWidth -= previousLeading;
        breakerState.snapBlock.first.snapGridLeadingPadPx = leading;
        breakerState.snapBlock.first.measuredWidth += leading;
        breakerState.snapBlock.last.measuredWidth -= previousTrailing;
        s.snapGridLeadingPadPx = 0;
        s.snapGridTrailingPadPx = trailing;
        s.measuredWidth = naturalWidth + trailing;
        committedWidth = combinedAllocated - breakerState.snapBlock.allocatedWidthPx;
        breakerState.snapBlock = {
          kind,
          first: breakerState.snapBlock.first,
          last: s,
          naturalWidthPx: combinedNatural,
          allocatedWidthPx: combinedAllocated,
        };
      } else {
        const allocated = snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
        const slack = allocated - naturalWidth;
        const leading = kind === 'latin' ? slack / 2 : 0;
        const trailing = slack - leading;
        s.snapGridLeadingPadPx = leading;
        s.snapGridTrailingPadPx = trailing;
        s.measuredWidth = allocated;
        committedWidth = allocated;
        breakerState.snapBlock = {
          kind,
          first: s,
          last: s,
          naturalWidthPx: naturalWidth,
          allocatedWidthPx: allocated,
        };
      }
    } else {
      s.snapGridClass = undefined;
      s.snapGridLeadingPadPx = undefined;
      s.snapGridTrailingPadPx = undefined;
      s.snapGridCellPitchPx = undefined;
      s.measuredWidth = w;
      breakerState.snapBlock = null;
    }
  } else {
    breakerState.snapBlock = null;
  }
  commitMixedLineItem(breakerState, s, scale);
  breakerState.currentWidth += committedWidth;
  if (justifiedCompressionApplies(operationState)) {
    const previous = breakerState.justifiedGapModel;
    breakerState.justifiedGapModel = previous?.segmentCount === breakerState.currentLine.length - 1
      ? lineGapModel([gapSegment(operationState, s)], previous, false)
      : lineGapModel(breakerState.currentLine.map(item => gapSegment(operationState, item)), undefined, false);
  }
  if (
    'text' in s &&
    s.latinSpaceCompressionEligible === true && s.latinSpaceAverageWidthRatio != null &&
    s.fontRoute
  ) {
    if (breakerState.latinLineFace && !sameLatinSpaceFace(s, breakerState.latinLineFace)) {
      materializeLatinSpaceCompression();
      breakerState.latinLineHomogeneous = false;
    }
    breakerState.latinLineFace ??= s;
    if (s.latinNaturalTrailingSpacePx !== undefined) {
      const capacity = Math.max(
        0,
        s.latinNaturalTrailingSpacePx - (
          ((calcEffectiveFontPx(s, scale) * s.latinSpaceAverageWidthRatio!) / 2) * charScaleFactor(s) +
          segmentCharacterGridDeltaPx(s, characterGrid, scale)
        ),
      );
      if (
        breakerState.latinUniformGapCapacity !== undefined &&
        Math.abs(capacity - breakerState.latinUniformGapCapacity) > 1e-6
      ) {
        materializeLatinSpaceCompression();
        breakerState.latinLineHomogeneous = false;
      }
      breakerState.latinUniformGapCapacity ??= capacity;
      breakerState.latinLineGaps.push(s);
    }
  } else if (!(operationState.isJustified && isInklessLineItem(s))) {
    // In a justified line an anchor character or paragraph-mark metric has no
    // advance or separator, so it neither forms nor interrupts a gap.
    materializeLatinSpaceCompression();
    breakerState.latinLineHomogeneous = false;
  }
  if (h > breakerState.lineHeight) breakerState.lineHeight = h;
  if ('imagePath' in s && s.inlinePicture === true) {
    breakerState.lineHasInlinePicture = true;
    breakerState.linePictureMarkSingle = Math.max(
      breakerState.linePictureMarkSingle,
      (s.paragraphMarkSinglePx ?? 0) * scale,
    );
  }
  if (asc > breakerState.lineAscent) breakerState.lineAscent = asc;
  if (desc > breakerState.lineDescent) breakerState.lineDescent = desc;
  const paintsInlineInk = !('text' in s) || s.metricOnly !== true;
  if (paintsInlineInk) {
    breakerState.lineHasVisibleMetrics = true;
    if (asc > breakerState.lineVisibleAscent) breakerState.lineVisibleAscent = asc;
    if (desc > breakerState.lineVisibleDescent) breakerState.lineVisibleDescent = desc;
  }
  // Grid-count height for docGrid cell allocation (§17.6.5). Only East Asian
  // TEXT (and tall inline objects) drives the count — a Latin run keeps its
  // natural height and is NOT cell-rounded, so it must not contribute (its
  // substituted Canvas box would otherwise inflate the count). An EA text run
  // counts from its DESIGN height when tabled, else the deterministic Word FE
  // 1.3em fallback; an image/math object counts its measured box. The line's
  // value is the max.
  let segGridCount = 0;
  if (!('isTab' in s) && !('imagePath' in s) && !('math' in s)) {
    const ts = s as LayoutTextSeg;
    if (ts.ruby) breakerState.lineHasRuby = true;
    const metricEastAsian = ts.metricEastAsian === true || EAST_ASIAN_RE.test(ts.text);
    if (!breakerState.lineEastAsian && metricEastAsian) breakerState.lineEastAsian = true;
    // Prefer the selected resource's single-line height. Without admitted
    // geometry, the generic East Asian grid fallback remains authoritative.
    // Small caps (non-super/sub) keep the FULL run size here so the line box
    // follows the run size, not the 2pt-reduced glyphs (§17.3.2.33).
    const intendedEm = ts.smallCaps && !ts.vertAlign ? ts.fontSize * scale : effectiveFontPx(ts);
    // The OpenType code-page class selects the general line ratio even for
    // Latin text. This script hint selects an optional East-Asian-specific
    // floor and grid-cell counting; ruby keeps its measured annotation box.
    const segScriptHint = metricEastAsian && !ts.ruby;
    const nativeRatio =
      ts.resolvedLineHeightRatio == null
        ? nativeCanvasLineRatio(
            ctx,
            fontFamilyClasses,
            ts.fontRoute,
            ts.fontFamily,
            ts.bold ? 700 : 400,
            ts.italic ? 'italic' : 'normal',
            ts.text,
          )
        : null;
    const designIntended =
      ts.textBoxLineFloor && ts.ruby
        ? 0
        : Math.max(
            segmentIntendedSingleLinePx(ts, intendedEm, segScriptHint),
            ts.textBoxLineFloor || ts.metricEastAsian === true
              ? segmentEastAsiaFloorSingleLinePx(ts, intendedEm, segScriptHint)
              : 0,
          );
    if (paintsInlineInk && !metricEastAsian && !ts.ruby
      && ts.resolvedLatinGridCellAllocation === true) {
      breakerState.lineLatinGridCountSingle = Math.max(breakerState.lineLatinGridCountSingle, designIntended);
    }
    const intended = Math.max(designIntended, (nativeRatio ?? 0) * intendedEm);
    if (intended > breakerState.lineIntendedSingle) breakerState.lineIntendedSingle = intended;
    if (paintsInlineInk && intended > breakerState.lineVisibleIntendedSingle) {
      breakerState.lineVisibleIntendedSingle = intended;
    }
    // Only East Asian text is cell-rounded. The native Canvas probe can
    // establish a browser-selected font box for ordinary auto lines, but it
    // cannot establish Word's Far-East design height or OS/2 code-page class.
    // For an untabled tuple retain the documented 1.3em grid fallback;
    // parsed resource/reference design metrics still take precedence.
    if (segScriptHint) segGridCount = eastAsianGridCountSinglePx(designIntended, intendedEm);
  } else if (!('isTab' in s)) {
    // Image/math object: a tall inline object sizes the line's cells too.
    segGridCount = asc + desc;
  }
  if (segGridCount > breakerState.lineGridCountSingle)
    breakerState.lineGridCountSingle = segGridCount;
}

export function performSegNaturalAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): number {
  const { measureText, verticalInkExtra, scale, characterGrid } = operationState;
  return segAdvanceWidth(
    s,
    measureText(s).width + verticalInkExtra(s, s.text),
    characterGrid,
    scale,
  );
}

export function performStandaloneSnapAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  naturalWidth: number,
): number {
  const { snapPitchPx, eastAsianSnapCellCount, characterGrid } = operationState;

  const kind = snapToCharsClass(s, characterGrid);
  if (!kind || snapPitchPx == null || s.text.length === 0) return naturalWidth;
  return snapToCharsAllocatedWidthPx(
    naturalWidth,
    kind,
    snapPitchPx,
    kind === 'eastAsia' ? eastAsianSnapCellCount(s) : 1,
  );
}

export function performSegAdvance(operationState: PassOperationState, s: LayoutTextSeg): number {
  const { segNaturalAdvance, standaloneSnapAdvance } = operationState;
  return standaloneSnapAdvance(s, segNaturalAdvance(s));
}

export function performStrNaturalAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  text: string,
  retainTrailingPunctuationCompression = false,
): number {
  const { measurement, effectiveFontPx, verticalInkExtra, scale, characterGrid } = operationState;

  const start = retainTrailingPunctuationCompression ? s.text.length - text.length : 0;
  const measuredSegment = {
    ...s,
    text,
    punctuationCompressions: slicedPunctuationCompressions(
      s,
      Math.max(0, start),
      Math.max(0, start) + text.length,
    ),
  };
  if (s.textLayoutService && s.textShapeRequest) {
    const shaped = s.textLayoutService.shape({
      ...sliceTextShapeRequest(s.textShapeRequest, Math.max(0, start), Math.max(0, start) + text.length),
      fontSizePt: effectiveFontPx(s),
      measure: true,
      clusterGeometry: false,
    });
    return segAdvanceWidth(
      measuredSegment,
      shaped.advancePt + verticalInkExtra(s, text),
      characterGrid,
      scale,
    );
  }
  const natural = measurement.measureRunText(s, text).width;
  return segAdvanceWidth(
    measuredSegment,
    natural + verticalInkExtra(s, text),
    characterGrid,
    scale,
  );
}

export function performEastAsianSnapCellCount(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): number {
  const { snapPitchPx, measurement, measureText, verticalInkExtra, scale, characterGrid } =
    operationState;

  if (snapPitchPx == null) return 1;
  if (s.textLayoutService && s.textShapeRequest && !s.shapedClusters) {
    measureText(s, true);
  }
  const shapedClusters = s.shapedClusters?.length ? s.shapedClusters : null;
  const boundaries =
    shapedClusters == null
      ? [...new Set([0, ...graphemeClusterOffsets(s.text), s.text.length])].sort((a, b) => a - b)
      : null;
  const ranges =
    shapedClusters?.map((cluster) => ({
      start: cluster.range.start,
      end: cluster.range.end,
      advancePx: cluster.advancePt,
    })) ??
    boundaries!.slice(0, -1).map((start, index) => ({
      start,
      end: boundaries![index + 1]!,
      advancePx: undefined,
    }));
  let cells = 0;
  for (const range of ranges) {
    const { start, end } = range;
    if (end <= start) continue;
    const text = s.text.slice(start, end);
    const measuredSegment = {
      ...s,
      text,
      ...slicedTextMetadata(s, start, end),
    };
    let naturalAdvancePx: number;
    if (range.advancePx != null) {
      naturalAdvancePx = segAdvanceWidth(
        measuredSegment,
        range.advancePx + verticalInkExtra(s, text),
        characterGrid,
        scale,
      );
    } else {
      const naturalWidthPx = measurement.measureRunText(s, text).width;
      naturalAdvancePx = segAdvanceWidth(
        measuredSegment,
        naturalWidthPx + verticalInkExtra(s, text),
        characterGrid,
        scale,
      );
    }
    cells += wordSnapToCharsEastAsianCellCount(naturalAdvancePx, snapPitchPx);
  }
  return Math.max(1, cells);
}

export function performStrAdvance(
  operationState: PassOperationState,
  s: LayoutTextSeg,
  text: string,
  retainTrailingPunctuationCompression = false,
): number {
  const { standaloneSnapAdvance, strNaturalAdvance } = operationState;

  const start = retainTrailingPunctuationCompression ? s.text.length - text.length : 0;
  const candidate = {
    ...s,
    text,
    // Snap-cell acquisition can remeasure this candidate, so it needs the
    // same retained range as strNaturalAdvance, not the parent word request.
    ...slicedTextMetadata(s, Math.max(0, start), Math.max(0, start) + text.length),
    shapedClusters: text === s.text ? s.shapedClusters : undefined,
  };
  return standaloneSnapAdvance(
    candidate,
    strNaturalAdvance(s, text, retainTrailingPunctuationCompression),
  );
}

export function performFitHomogeneousLatinSpaces(
  operationState: PassOperationState,
  next: LayoutTextSeg,
  nextFitWidth: number,
): boolean {
  const {
    breakerState,
    sameLatinSpaceFace,
    availW,
    fitsMeasuredWidth,
    characterGrid,
    baseRtl,
    isJustified,
    widthPolicy,
  } = operationState;

  if (
    isJustified ||
    baseRtl ||
    widthPolicy !== 'bounded' ||
    characterGrid?.type === 'snapToChars' ||
    (characterGrid?.type === 'linesAndChars' && next.widthBalanceGridDeltaFactor !== 0.5) ||
    next.latinSpaceCompressionEligible !== true ||
    next.latinSpaceAverageWidthRatio == null ||
    !next.fontRoute ||
    next.rtl ||
    next.verticalRun ||
    next.tateChuYoko ||
    next.fitTextRegionIndex !== undefined ||
    !breakerState.latinLineHomogeneous ||
    !breakerState.latinLineFace ||
    !sameLatinSpaceFace(next, breakerState.latinLineFace) ||
    breakerState.latinLineGaps.length === 0
  )
    return false;
  const totalCapacity =
    (breakerState.latinUniformGapCapacity ?? 0) * breakerState.latinLineGaps.length;
  if (totalCapacity <= 0) return false;
  const restored = breakerState.latinAppliedPerGap * breakerState.latinAppliedGapCount;
  const required = Math.max(0, breakerState.currentWidth + restored + nextFitWidth - availW());
  if (
    required > totalCapacity ||
    !fitsMeasuredWidth(breakerState.currentWidth + restored + nextFitWidth - required, availW())
  ) {
    return false;
  }
  // Keep aggregate fit width current. Write each retained gap exactly once
  // when the line is finalized, avoiding quadratic work on long lines.
  breakerState.currentWidth += restored - required;
  breakerState.latinAppliedGapCount = breakerState.latinLineGaps.length;
  breakerState.latinAppliedPerGap = required / breakerState.latinAppliedGapCount;
  return true;
}

/** Closed paragraph/layout gates only; no source/content/face taxonomy. */
export function justifiedCompressionApplies(state: Pick<PassOperationState,
  'justifiedCompression' | 'baseRtl' | 'widthPolicy' | 'characterGrid'>): boolean {
  return state.justifiedCompression === true && !state.baseRtl
    && state.widthPolicy === 'bounded'
    && state.characterGrid?.type !== 'snapToChars'
    && state.characterGrid?.type !== 'linesAndChars';
}

function gapSegment(state: Pick<PassOperationState, 'strAdvance' | 'scale' | 'characterGrid'>, segment: LayoutSeg): GapSegment {
  if ('text' in segment) {
    // §17.3.2.14 owns fixed-width fitText geometry and atomic wrapping.
    // Its interaction with proportional justification is unmeasured: keep
    // the cell and adjacent spaces fixed, as for the existing pitch policy.
    if (segment.fitTextRegionIndex !== undefined) {
      return { widthPx: segment.measuredWidth, spacePx: 0 };
    }
    const spaceWidths = new Map<number, number>();
    const spaceClusters = segment.shapedSpaceClusters ?? segment.shapedClusters;
    if (spaceClusters && !segment.ruby) {
      const advances = new Map(spaceClusters.map(cluster => [cluster.range.start, cluster.advancePt]));
      let utf16 = 0;
      let cpOffset = 0;
      for (const character of segment.text) {
        if (character === ' ') {
          const advance = advances.get(utf16);
          if (advance === undefined) throw new Error('A text space lacks authoritative cluster geometry');
          spaceWidths.set(cpOffset, advance * charScaleFactor(segment)
            + segLetterSpacingPx(segment, state.characterGrid, state.scale));
        }
        utf16 += character.length;
        cpOffset += 1;
      }
    }
    const start = segment.text.indexOf(' ');
    const spacePx = spaceWidths.values().next().value ?? (start >= 0 && !segment.ruby
      ? state.strAdvance({ ...segment, text: ' ', ...slicedTextMetadata(segment, start, start + 1) }, ' ')
      : 0);
    return {
      text: segment.metricOnly ? '' : segment.ruby ? undefined : segment.text,
      widthPx: segment.measuredWidth, spacePx, spaceWidths,
    };
  }
  return { widthPx: segment.measuredWidth, spacePx: 0 };
}

/** Admit the whole unit against the same natural gap model that layout uses
 * for paint. Advances are never shortened here; compression is line-owned. */
export function fitJustifiedCompression(state: Pick<PassOperationState,
  'justifiedCompression' | 'baseRtl' | 'widthPolicy' | 'characterGrid' | 'breakerState'
  | 'segAdvance' | 'strAdvance' | 'availW' | 'textSegmentBox' | 'addToLine' | 'scale' | 'measurement' | 'flush'>, first: LayoutTextSeg): boolean {
  if (!justifiedCompressionApplies(state) || first.fitTextRegionIndex !== undefined) return false;
  const { breakerState: breaker } = state;
  if (breaker.justifiedUnitEnd) {
    if (breaker.justifiedUnitEnd === first) breaker.justifiedUnitEnd = undefined;
    return false;
  }
  const members = candidateUnit(first, breaker.queue);
  // Opaque successors stay with their own placement path. Ruby is itself an
  // opaque, measured text unit and participates without opening adjacent gaps.
  if (members.some(member => !('text' in member))) return false;
  let previousText = breaker.currentLine.at(-1);
  const boundaries: number[] = [];
  const boxes = members.map(member => {
    if (!('text' in member)) throw new Error('A text candidate lost its placement unit');
    const box = state.textSegmentBox(member);
    const boundary = state.measurement.wordBoundaryAdvance(
      previousText && 'text' in previousText ? previousText : undefined, member);
    boundaries.push(boundary);
    member.measuredWidth = box.width + boundary;
    previousText = member;
    return box;
  });
  if (!breaker.justifiedGapModel || breaker.justifiedGapModel.segmentCount !== breaker.currentLine.length) {
    breaker.justifiedGapModel = lineGapModel(breaker.currentLine.map(s => gapSegment(state, s)), undefined, false);
  }
  const previous = breaker.justifiedGapModel;
  const model = lineGapModel(members.map(s => gapSegment(state, s)), previous, false);
  const overflow = model.visibleWidthPx - state.availW();
  const factor = overflow > 0 ? wordJustifiedInterwordCompressionFactor({
    overflow, naturalGapSum: model.S,
    candidateLineEndSeparator: model.lineEndSeparatorPx,
    previousOpportunitySum: previous.S + previous.lineEndSeparatorPx,
    expansionWithoutCandidate: state.availW() - previous.visibleWidthPx,
  }) : 0;
  if (factor === undefined) {
    for (const [index, member] of members.entries()) member.measuredWidth = boxes[index].width;
    // An adjusted candidate that was refused must be reconsidered at a fresh
    // line origin; the ordinary isolated-width path cannot override this fit.
    if (boundaries.some(delta => delta !== 0) && breaker.currentLine.length > 0) {
      state.flush(undefined, false, first.src);
      breaker.queue.unshift(first);
      return true;
    }
    if (members.length > 1) breaker.justifiedUnitEnd = members.at(-1);
    return false;
  }
  for (const [index, member] of members.entries()) {
    if (member !== first) breaker.queue.shift();
    if (!('text' in member)) throw new Error('A text candidate lost its placement unit');
    const box = boxes[index];
    member.leadingWordBoundaryPx = boundaries[index];
    state.addToLine(member, member.measuredWidth, box.height, box.ascent, box.descent);
  }
  breaker.justifiedCompressionPx = Math.max(0, overflow);
  return true;
}

export function performTextSegmentBox(
  operationState: PassOperationState,
  s: LayoutTextSeg,
): Readonly<{
  width: number;
  height: number;
  ascent: number;
  descent: number;
}> {
  const {
    measurement,
    effectiveFontPx,
    measureText,
    verticalInkExtra,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
  } = operationState;

  s.leadingWordBoundaryPx = undefined;
  // Fitting needs the spaces' contextual advances, not every prefix of a
  // possibly overlong word. Full cluster acquisition belongs to final slices.
  const measured = justifiedCompressionApplies(operationState) && s.text.includes(' ') && !s.ruby
    ? measurement.measureSegment(s, 'spaces')
    : measureText(s, snapToCharsClass(s, characterGrid) === 'eastAsia');
  let width = segAdvanceWidth(
    s,
    measured.width + verticalInkExtra(s, s.text),
    characterGrid,
    scale,
  );
  s.snapGridNaturalWidthPx = width;
  // Max-content uses the unbroken native context, as its ordinary intrinsic
  // text merge already does. Registered physical units must stay intact;
  // acquire only their separator pair through the shared boundary oracle.
  // This is not the bounded-line Word justification/compression policy.
  if (operationState.widthPolicy === 'intrinsic') {
    const previous = operationState.breakerState.currentLine.at(-1);
    if (previous && 'text' in previous && (previous.semanticSlotSpans || s.semanticSlotSpans)) {
      const boundary = measurement.wordBoundaryAdvance(previous, s);
      s.leadingWordBoundaryPx = boundary;
      width += boundary;
    }
  }

  const fullPx = s.fontSize * scale;
  let metricMeasurement = measured;
  let metricEmPx = effectiveFontPx(s);
  if (s.metricOnly && s.metricProbeText) {
    // Native reserved-separator participant (MS-DOC 2.3.3): its own text is
    // empty, so its vertical metrics come from its bounded probe through the
    // same selected face, at the existing effective metric size in this
    // caller's scale (super/sub scaling; small caps keep the full-size policy
    // of the branch below). Width stays the empty measurement above; probe
    // clusters, ink and advance are never retained. Position applies once below.
    const probeEmPx = s.smallCaps && !s.vertAlign ? fullPx : metricEmPx;
    const probe = s.textLayoutService && s.textShapeRequest
      ? s.textLayoutService.shape({
          ...independentTextShapeRequest(s.textShapeRequest, s.metricProbeText),
          fontSizePt: probeEmPx,
          measure: true,
          clusterGeometry: false,
        })
      : undefined;
    const fallback = probe ? undefined : measurement.measureWithFont(
      buildFont(s.bold, s.italic, probeEmPx, s.fontFamily, fontFamilyClasses, s.fontRoute),
      s.metricProbeText,
    );
    metricMeasurement = {
      width: measured.width,
      actualBoundingBoxAscent: probe ? probe.ascentPt : fallback!.actualBoundingBoxAscent,
      actualBoundingBoxDescent: probe ? probe.descentPt : fallback!.actualBoundingBoxDescent,
      fontBoundingBoxAscent: probe ? probe.ascentPt : fallback!.fontBoundingBoxAscent,
      fontBoundingBoxDescent: probe ? probe.descentPt : fallback!.fontBoundingBoxDescent,
    } as TextMetrics;
    metricEmPx = probeEmPx;
  } else if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
    if (s.textLayoutService && s.textShapeRequest) {
      const shaped = s.textLayoutService.shape({
        ...(s.text ? s.textShapeRequest : independentTextShapeRequest(s.textShapeRequest, 'X')),
        fontSizePt: fullPx,
        measure: true,
        clusterGeometry: false,
      });
      metricMeasurement = {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    } else {
      metricMeasurement = measurement.measureWithFont(
        buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses, s.fontRoute),
        s.text || 'X',
      );
    }
    metricEmPx = fullPx;
  }

  const corrected = measuredLineMetrics(metricMeasurement, fullPx);
  // Selected resource sides come from the same admitted face as the line
  // height; native reference sides are policy geometry only. Ruby and an
  // authored baseline position compose additional boxes outside either
  // simple OpenType projection, so retain measured sides for those inputs.
  const designOwnsSides =
    (s.resolvedResourceVerticalMetric || s.referenceFontVerticalMetric) &&
    !s.ruby &&
    (s.position ?? 0) === 0 &&
    s.resolvedDesignAscentRatio != null &&
    s.resolvedDesignDescentRatio != null;
  let ascent = designOwnsSides ? s.resolvedDesignAscentRatio! * metricEmPx : corrected.ascent;
  let descent = designOwnsSides ? s.resolvedDesignDescentRatio! * metricEmPx : corrected.descent;
  if (s.positionExtendsLineBox !== false) {
    const positionPx = (s.position ?? 0) * scale;
    if (positionPx > 0) ascent += positionPx;
    else if (positionPx < 0) descent -= positionPx;
  }
  if (s.ruby && (!s.textBoxLineFloor || s.textBoxVertical)) {
    ascent += rubyAscentReservePx(
      s.ruby.fontSizePt,
      s.ruby.hpsRaisePt,
      scale,
      s,
      ctx,
      fontFamilyClasses,
    );
  }
  return { width, height: s.fontSize, ascent, descent };
}

export function performAppendQueuedIdeographicSpaceSegment(
  operationState: PassOperationState,
  source: LayoutTextSeg,
): void {
  const { breakerState, addToLine, textSegmentBox } = operationState;

  if (
    /\s$/u.test(source.text) ||
    source.ruby !== undefined ||
    source.tateChuYoko === true ||
    source.fitTextRegionIndex !== undefined
  )
    return;
  const follower = breakerState.queue.peek();
  if (
    !follower ||
    !('text' in follower) ||
    follower.joinPrev !== true ||
    follower.text.length === 0 ||
    [...follower.text].some((character) => character !== '\u3000')
  )
    return;
  breakerState.queue.shift();
  const hangingCount = wordIdeographicSpaceLineEndAllowanceCount(
    hasEastAsianVisiblePredecessor(source.text),
    follower.paragraphFinalIdeographicSpaceCount ?? [...follower.text].length,
  );
  if (hangingCount === 0) {
    breakerState.queue.unshift(follower);
    return;
  }
  const hangingText = follower.text.slice(0, hangingCount);
  const hangingSegment: LayoutTextSeg = {
    ...follower,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: hangingText,
    measuredWidth: 0,
    ...slicedTextMetadata(follower, 0, hangingText.length),
  };
  const followerBox = textSegmentBox(hangingSegment);
  hangingSegment.measuredWidth = followerBox.width;
  addToLine(
    hangingSegment,
    followerBox.width,
    followerBox.height,
    followerBox.ascent,
    followerBox.descent,
  );
  const remainder = follower.text.slice(hangingText.length);
  if (remainder.length > 0) {
    breakerState.queue.unshift({
      ...follower,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: remainder,
      measuredWidth: 0,
      joinPrev: undefined,
      hardJoinPrev: undefined,
      ...slicedTextMetadata(follower, hangingText.length, follower.text.length),
      src: follower.src
        ? {
            segIndex: follower.src.segIndex,
            charOffset: follower.src.charOffset + hangingText.length,
          }
        : undefined,
    });
  }
}

export function performTabFollowWidth(operationState: PassOperationState, q: LayoutSeg): number {
  const { segAdvance, scale } = operationState;

  if ('isTab' in q) return q.measuredWidth || 0;
  if ('imagePath' in q) return q.widthPt * scale;
  if ('math' in q) return q.measuredWidth || 0;
  if ('lineBreak' in q) return 0;
  return segAdvance(q);
}

export function performDecimalAlignmentPoint(
  segments: readonly LayoutSeg[],
): Readonly<{ segmentIndex: number; charOffset: number }> | null {
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex]!;
    if (!('text' in segment)) continue;
    const separator = segment.text.indexOf('.');
    if (separator >= 0) return { segmentIndex, charOffset: separator };
  }

  let lastDigit: Readonly<{ segmentIndex: number; charOffset: number }> | null = null;
  let inFirstNumber = false;
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex]!;
    if (!('text' in segment)) {
      if (inFirstNumber) return lastDigit;
      continue;
    }
    let charOffset = 0;
    for (const scalar of segment.text) {
      charOffset += scalar.length;
      if (/\p{Decimal_Number}/u.test(scalar)) {
        inFirstNumber = true;
        lastDigit = { segmentIndex, charOffset };
      } else if (inFirstNumber) {
        return lastDigit;
      }
    }
  }
  return lastDigit;
}

export function performDecimalAlignmentPrefixWidth(
  operationState: PassOperationState,
  segments: readonly LayoutSeg[],
): number | undefined {
  const { strAdvance, tabFollowWidth, decimalAlignmentPoint } = operationState;

  const point = decimalAlignmentPoint(segments);
  if (!point) return undefined;
  let width = 0;
  for (let index = 0; index < point.segmentIndex; index += 1) {
    width += tabFollowWidth(segments[index]!);
  }
  const segment = segments[point.segmentIndex]!;
  if (!('text' in segment)) return width;
  return width + strAdvance(segment, segment.text.slice(0, point.charOffset));
}

export function performTabFollowingMetrics(operationState: PassOperationState): Readonly<{
  totalWidth: number;
  decimalPrefixWidth?: number;
}> {
  const { breakerState, tabFollowWidth, decimalAlignmentPrefixWidth } = operationState;

  const following: LayoutSeg[] = [];
  let totalWidth = 0;
  for (const q of breakerState.queue) {
    if ('isTab' in q || 'lineBreak' in q) break;
    following.push(q);
    totalWidth += tabFollowWidth(q);
  }
  const decimalPrefixWidth = decimalAlignmentPrefixWidth(following);
  return decimalPrefixWidth === undefined ? { totalWidth } : { totalWidth, decimalPrefixWidth };
}

/** Bracket close to the line band; probing a whole suffix midpoint on
 * every queued line repeats long measurements for dense sparse breaks. */
function lastFittingMonotoneIndex(count: number, fits: (index: number) => boolean): number {
  if (!count || !fits(0)) return -1;
  let lower = 0, upper = 1;
  while (upper < count && fits(upper)) {
    lower = upper;
    upper = upper * 2 + 1;
  }
  upper = Math.min(upper, count);
  while (lower + 1 < upper) {
    const middle = Math.floor((lower + upper) / 2);
    if (fits(middle)) lower = middle;
    else upper = middle;
  }
  return lower;
}

export function performEmergencyTextSplit(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  available: number,
  forceAtLeastOne = true,
): number {
  const {
    prospectiveSnapAdvance,
    measurement,
    effectiveFontPx,
    setMeasureFont,
    strNaturalAdvance,
    strAdvance,
    ctx,
    scale,
    fontFamilyClasses,
    characterGrid,
    verticalGlyphMeasurement,
  } = operationState;

  operationState.reservePrefixWork(segment.text.length);
  const protectedOffsets = protectedNoBreakOffsets(segment);
  const graphemeOffsets = [...new Set([0, ...graphemeClusterOffsets(segment.text), segment.text.length])];
  let split = 0;
  if (available > 0) {
    const monotoneAllocation =
      charSpacingDeltaPx(segment, scale) >= 0 &&
      snapToCharsClass(segment, characterGrid) !== 'latin';
    if (monotoneAllocation) {
      setMeasureFont(
        buildFont(
          segment.bold,
          segment.italic,
          effectiveFontPx(segment),
          segment.fontFamily,
          fontFamilyClasses,
          segment.fontRoute,
        ),
      );
      measurement.withSegmentKerning(segment, () => {
        const fitted = fitCJKPrefix(
          ctx,
          segment.text,
          available,
          segmentCharacterGridDeltaPx(segment, characterGrid, scale),
          charScaleFactor(segment),
          charSpacingDeltaPx(segment, scale),
          segment.verticalRun === true,
          verticalGlyphMeasurement,
          (prefix) => {
            operationState.reservePrefixWork(prefix.length);
            return strAdvance(segment, prefix);
          },
        ).length;
        split =
          graphemeOffsets
            .filter((offset) => offset <= fitted && !protectedOffsets.has(offset))
            .at(-1) ?? 0;
      });
    } else if (charSpacingDeltaPx(segment, scale) >= 0) {
      // A Latin block's ceil((activeNatural + prefixNatural) / pitch) * pitch
      // minus its fixed allocation is monotone in the retained natural width.
      // Use that exact active-block allocator, rather than standalone rounding.
      const legal = graphemeOffsets.filter(offset => offset > 0 && !protectedOffsets.has(offset));
      const index = lastFittingMonotoneIndex(legal.length, i => {
        operationState.reservePrefixWork(legal[i]!);
        const natural = strNaturalAdvance(segment, segment.text.slice(0, legal[i]!));
        return prospectiveSnapAdvance(segment, natural) <= available + 1e-9;
      });
      if (index >= 0) split = legal[index]!;
    } else {
      // Signed spacing can make prefix advances
      // non-monotone. Evaluate every legal retained candidate against the
      // exact prospective line block rather than binary-searching a
      // standalone approximation.
      for (const offset of graphemeOffsets) {
        if (offset <= 0 || protectedOffsets.has(offset)) continue;
        operationState.reservePrefixWork(offset);
        const natural = strNaturalAdvance(segment, segment.text.slice(0, offset));
        if (prospectiveSnapAdvance(segment, natural) <= available + 1e-9) split = offset;
      }
    }
  }
  if (split <= 0 && forceAtLeastOne) {
    split =
      graphemeOffsets.find((offset) => offset > 0 && !protectedOffsets.has(offset)) ??
      segment.text.length;
  }
  // Preserve the existing JLReq/Word line-end hanging rule after switching
  // the emergency splitter from code-point indexes to UTF-16 grapheme offsets.
  while (segment.text.startsWith('\u3000', split)) split += 1;
  return split;
}

export function performExplicitTextSplit(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  available: number,
): number {
  const { prospectiveSnapAdvance, strNaturalAdvance, scale } = operationState;
  const window = segment.explicitBreaks;
  if (!(available > 0) || !window) return 0;
  const count = window.end - window.start;
  const fits = (index: number): boolean => {
    const offset = textBreakOffsetAt(window, index);
    operationState.reservePrefixWork(offset);
    const natural = strNaturalAdvance(segment, segment.text.slice(0, offset));
    return prospectiveSnapAdvance(segment, natural) <= available + 1e-9;
  };
  // Use the emergency splitter's existing monotone-allocation contract. Signed
  // spacing retains exact candidate evaluation. Positive Latin snap allocation
  // is a monotone ceil map of the same natural-prefix authority, even with an
  // existing block; fits() uses its actual prospective allocation.
  if (charSpacingDeltaPx(segment, scale) < 0) {
    let selected = 0;
    for (let i = 0; i < count; i++) if (fits(i)) selected = textBreakOffsetAt(window, i);
    return selected;
  }
  const index = lastFittingMonotoneIndex(count, fits);
  return index < 0 ? 0 : textBreakOffsetAt(window, index);
}

export function performQueueEmergencyTail(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  split: number,
): void {
  const { breakerState } = operationState;

  breakerState.queue.unshift({
    ...segment,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: segment.text.slice(split),
    ...slicedTextMetadata(segment, split, segment.text.length),
    seaBreaks: rebaseSeaBreaks(segment.seaBreaks, split),
    measuredWidth: 0,
    // The emergency split itself is now the legal line boundary. A source-
    // boundary glue marker protects only the first retained prefix; carrying
    // it onto the tail would make the next line overflow again.
    joinPrev: undefined,
    hardJoinPrev: undefined,
    src: {
      segIndex: segment.src!.segIndex,
      charOffset: segment.src!.charOffset + split,
    },
  });
}

export function performRetractCurrentLineForLeadingKinsoku(
  operationState: PassOperationState,
  next: LayoutTextSeg,
): CrossRunKinsokuRetraction {
  const { breakerState, materializeLatinSpaceCompression, strAdvance, kinsoku } = operationState;
  return retractLeadingKinsoku(
    breakerState,
    kinsoku,
    materializeLatinSpaceCompression,
    strAdvance,
    next,
    operationState.scale,
  );
}

export function performKeepLeadingKinsokuWithCurrentLine(
  operationState: PassOperationState,
  segment: LayoutTextSeg,
  h: number,
  asc: number,
  desc: number,
): boolean {
  const { breakerState, addToLine, strNaturalAdvance, queueEmergencyTail } = operationState;
  return keepLeadingKinsoku(
    breakerState,
    strNaturalAdvance,
    addToLine,
    queueEmergencyTail,
    segment,
    h,
    asc,
    desc,
  );
}
