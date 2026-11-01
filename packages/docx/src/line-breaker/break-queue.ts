import { createMixedSpaceState } from './mixed-space-fit.js';
import { LineMeasurementAdapter } from './measurement-adapter.js';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { containsSeaScript, seaMixedBreakOffsets } from '@silurus/ooxml-core';
import { type LayoutImageSeg, type LayoutLine, type LayoutMathSeg, type LayoutSeg, type LayoutTabSeg, type LayoutTextSeg, type LineBoundary, type WrapLayoutCtx } from './model.js';
import { protectedNoBreakOffsets, slicedTextMetadata } from './advance.js';
import { rebaseSeaBreaks } from './text-runs.js';
import { resolveFitTextSegments } from './segment-builder.js';
import type { LineGapModel } from './line-gaps.js';
import { SegmentQueue, type SegmentQueueCursor } from './segment-queue.js';

/** Prepare source-anchored break opportunities and the resumable queue.
 * SEA dictionary boundaries, protected ranges, paragraph-final hanging spaces,
 * and fitText allocation are resolved once per pass at the requested scale. */
export function prepareBreakQueue(
  segs: LayoutSeg[],
  startBoundary: LineBoundary | undefined,
  kinsoku: KinsokuRules,
  scale: number,
  measurement: LineMeasurementAdapter,
): LayoutSeg[] {
  const sourcedSegs = segs.map((seg, segIndex) => {
    seg.src = { segIndex, charOffset: 0 };
    // Issue #797 / #960 — attach the SEA (Thai/Lao/Khmer) break offsets ONCE per
    // segment (perf: never per line/char). Only for SEA text; non-SEA segments
    // keep `seaBreaks` absent so their wrap path is byte-identical. The set now
    // UNIONS the dictionary word boundaries (#797) with the no-space SEA↔non-SEA
    // script transitions and, for a mixed CJK+SEA run (a `<w:cs/>` run keeps CJK
    // in the same cs segment), the CJK per-character opportunities — so each
    // script keeps its own break rule inside one contiguous segment (#960). The
    // layout kinsoku set (§17.15.1.58–.60) drops positions that would orphan a
    // forbidden char at a line head/tail, replacing the CJK path's retract.
    if ('text' in seg && containsSeaScript(seg.text)) {
      const protectedOffsets = protectedNoBreakOffsets(seg);
      seg.seaBreaks = seaMixedBreakOffsets(seg.text, { cjk: true, kinsoku }).filter(
        (offset) => !protectedOffsets.has(offset),
      );
    }
    return seg;
  });
  let queue: LayoutSeg[];
  if (!startBoundary) {
    queue = sourcedSegs;
  } else if (startBoundary.segIndex >= sourcedSegs.length) {
    queue = [];
  } else {
    const first = sourcedSegs[startBoundary.segIndex];
    if (startBoundary.charOffset > 0) {
      if (!('text' in first) || startBoundary.charOffset > first.text.length) {
        queue = [];
      } else {
        const text = first.text.slice(startBoundary.charOffset);
        queue = text
          ? [
              {
                ...first,
                text,
                measuredWidth: 0,
                src: { ...startBoundary },
                // A retained resume boundary has already consumed the source
                // seam. Carrying either marker would invent new ownership at
                // the start of this suffix.
                joinPrev: undefined,
                hardJoinPrev: undefined,
                ...slicedTextMetadata(first, startBoundary.charOffset, first.text.length),
                // Rebase the SEA break offsets onto the resumed (sliced) text so
                // a paginated Thai paragraph still breaks at word boundaries.
                seaBreaks: rebaseSeaBreaks(first.seaBreaks, startBoundary.charOffset),
              },
              ...sourcedSegs.slice(startBoundary.segIndex + 1),
            ]
          : sourcedSegs.slice(startBoundary.segIndex + 1);
      }
    } else {
      queue = sourcedSegs.slice(startBoundary.segIndex);
    }
  }

  // Mark the paragraph-final U+3000 suffix once. A backwards pass avoids the
  // O(N²) suffix rescans that would result from queue.every/reduce per segment.
  let paragraphFinalIdeographicSpaceCount = 0;
  let paragraphFinalIdeographicSpaceTailStartIndex = -1;
  const markedParagraphFinalTail: Array<
    Readonly<{
      index: number;
      segment: LayoutTextSeg;
    }>
  > = [];
  for (let index = queue.length - 1; index >= 0; index -= 1) {
    const candidate = queue[index];
    if (!candidate || !('text' in candidate) || candidate.text.length === 0) break;
    // `fitText` (§17.3.2.14) and tate-chu-yoko (§17.3.2.10) are indivisible
    // layout cells. Ruby owns one base/guide pair. Paragraph-final whitespace
    // may affect how those cells measure, but must never split or clone them.
    if (
      candidate.fitTextRegionIndex !== undefined ||
      candidate.tateChuYoko === true ||
      candidate.ruby !== undefined
    ) {
      // buildSegments can split an atomic source run before its U+3000 tail
      // (ruby is retained only on the first emitted segment). Undo any markers
      // already assigned to trailing pieces from that same authored run.
      if (candidate.sourceRunIndex !== undefined) {
        for (
          let markedIndex = markedParagraphFinalTail.length - 1;
          markedIndex >= 0;
          markedIndex -= 1
        ) {
          const marked = markedParagraphFinalTail[markedIndex];
          if (marked.segment.sourceRunIndex !== candidate.sourceRunIndex) continue;
          marked.segment.paragraphFinalIdeographicSpaceTail = undefined;
          marked.segment.paragraphFinalIdeographicSpaceLocalCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceTailStart = undefined;
          markedParagraphFinalTail.splice(markedIndex, 1);
        }
        paragraphFinalIdeographicSpaceTailStartIndex = markedParagraphFinalTail.at(-1)?.index ?? -1;
      }
      break;
    }
    const trailingSpaces = /^\u3000+$/u.test(candidate.text);
    const visibleWithTrailingSpaces = /[^\u3000]\u3000+$/u.test(candidate.text);
    if (!trailingSpaces && !visibleWithTrailingSpaces) break;
    const localTrailingCount = trailingSpaces
      ? [...candidate.text].length
      : [...candidate.text].reverse().findIndex((character) => character !== '\u3000');
    paragraphFinalIdeographicSpaceCount += localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceTail = true;
    candidate.paragraphFinalIdeographicSpaceLocalCount = localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceCount = paragraphFinalIdeographicSpaceCount;
    paragraphFinalIdeographicSpaceTailStartIndex = index;
    markedParagraphFinalTail.push({ index, segment: candidate });
    if (visibleWithTrailingSpaces) break;
  }
  if (paragraphFinalIdeographicSpaceTailStartIndex >= 0) {
    const start = queue[paragraphFinalIdeographicSpaceTailStartIndex];
    if (start && 'text' in start) start.paragraphFinalIdeographicSpaceTailStart = true;
  }

  // Resolve §17.3.2.14 from RAW natural advances at this exact layout scale.
  // The resulting per-gap is folded into segAdvanceWidth below, so the line
  // breaker and paint pen use one width authority. Cached w:spacing is ignored.
  // #1014 — the natural width includes the vo=Tr ink deficit so the resolved gap
  // (target − natural)/n, plus the ink-grown cell the paint draws, still sums to
  // the fitText target (measure == paint); 0 for non-under-reporting runs.
  resolveFitTextSegments(
    queue.filter((seg): seg is LayoutTextSeg => 'text' in seg),
    scale,
    (segment) =>
      measurement.measureSegment(segment).width +
      measurement.verticalInkExtra(segment, segment.text),
  );

  return queue;
}


export type SnapBlockState = {
    kind: 'latin' | 'complexScript';
    first: LayoutTextSeg;
    last: LayoutTextSeg;
    naturalWidthPx: number;
    allocatedWidthPx: number;
  };


/** A window accepted for one line fragment. `solverWidth` is the width the
 * float solver compared with the requirement (before hanging-indent restore). */
export interface GapWindow {
  readonly topY: number;
  readonly xOffset: number;
  readonly maxWidth: number;
  readonly solverWidth: number;
  /** The solver's own window (absolute left, width), as an exclusion key. */
  readonly solverWindow?: Readonly<{ leftPt: number; widthPt: number }>;
}

/** WORD_FLOAT_GAP_FLOW admission as a placement transaction (#1670).
 * A fragment narrowed by an exclusion is filled by ordinary placement; if
 * placement would need a forced break or overflow at the fragment head, the
 * fragment rolls back and the next gap is searched with a strictly larger
 * requirement. Admission and placement therefore share one implementation. */
export interface GapTransaction {
  /** Continuation origin while this physical line may still take a later gap. */
  cursor: { topY: number; left: number; right: number } | null;
  /** Top requested for a new physical line (solver searches downward from it). */
  requestTopY: number;
  /** Width the next gap must offer, in solver units (incl. positive first-line indent). */
  requirement: number;
  window: GapWindow | null;
  narrowed: boolean;
  /** The fragment's line-end edge (reading order) is an exclusion, not the
   * paragraph edge: §17.3.1.21 hanging punctuation may not cross it. */
  endsAtExclusion: boolean;
  snapshot: Readonly<{
    scalars: Readonly<Record<string, unknown>>;
    snapBlock: unknown;
    linesLength: number;
    queue: SegmentQueueCursor;
  }> | null;
  /** Complete units precede a forced unit: end the fragment before this source. */
  stopBefore: LineBoundary | null;
  /** Windows a WORD_KASHIDA_LINE_FIT unit did not fit, skipped by the search. */
  excluded?: Array<{ topY: number; leftPt: number; widthPt: number }>;
  /** Excluded-window retries taken by this line (bounded). */
  excludedRetries?: number;
}

export function createLineBreakerState(maxWidth: number, wrapCtx?: WrapLayoutCtx) {
  return {
    lines: [] as LayoutLine[],
    /** First-line body indent once a float moved a numbered first line's
     * leading edge (WrapLayoutCtx.numberingFirstLine); else the pass input. */
    firstIndentOverride: undefined as number | undefined,
    numberingFirstLine: undefined as LayoutLine['numberingFirstLine'],
    currentLine: [] as (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[],
    currentWidth: 0,
    justifiedGapModel: undefined as LineGapModel | undefined,
    justifiedCompressionPx: 0,
    justifiedUnitEnd: undefined as LayoutSeg | undefined,
    latinLineFace: undefined as LayoutTextSeg | undefined,
    latinLineHomogeneous: true,
    latinLineGaps: [] as LayoutTextSeg[],
    /** WORD_COMPRESSED_SPACE_LINE_FIT state of the current line. */
    mixedSpace: createMixedSpaceState(),
    /** Some segment of the paragraph carries the rule's eligibility; when
     * false the projection does no work at all. */
    mixedSpaceEnabled: false,
    latinUniformGapCapacity: undefined as number | undefined,
    latinAppliedGapCount: 0,
    latinAppliedPerGap: 0,
    snapBlock: null as SnapBlockState | null,
    physicalLineIndex: 0,
    positionReferencePt: undefined as number | null | undefined,
    firstPositioned: undefined as LayoutTextSeg | undefined,
    uniformPositionEligible: true,
    fragmentCursor: null as { topY: number; left: number; right: number } | null,
    /** Pass-local admission transaction of the line fragment being filled. */
    gapTransaction: null as GapTransaction | null,
    /** Queue item taken by the iterator and not yet committed or re-queued. */
    inHand: undefined as LayoutSeg | undefined,
    lineHeight: 0,
    lineAscent: 0,
    lineDescent: 0,
    lineIntendedSingle: 0,
    lineHasInlinePicture: false,
    linePictureMarkSingle: 0,
    lineGridCountSingle: 0,
    lineLatinGridCountSingle: 0,
    lineVisibleAscent: 0,
    lineVisibleDescent: 0,
    lineVisibleIntendedSingle: 0,
    lineHasVisibleMetrics: false,
    /** Picture-bullet marker metric admitted to the current line (pt, px). */
    lineMarkerHeight: 0,
    lineMarkerAscent: 0,
    isFirst: true,
    lineMaxWidth: maxWidth,
    lineXOffset: 0,
    /** Exclusion-free width past the paragraph's trailing indent, up to the
     *  text margin, that an aligned tab cell on this line may allocate. */
    lineMarginExtension: 0,
    currentLineTopY: wrapCtx?.startPageY ?? 0,
    lineHasRuby: false,
    lineEastAsian: false,
    queue: new SegmentQueue(),
    trailingBreakFontSize: null as number | null,
  };
}
