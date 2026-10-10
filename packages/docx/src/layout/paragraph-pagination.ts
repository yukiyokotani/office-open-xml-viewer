import { nativeReadingRequiresFreshColumn } from './native-reading-admission.js';
import { adjustForWidowOrphan, selectLargestFittingEnd } from '../line-fit-policy.js';
import type { LineBoundary } from '../line-layout.js';
import {
  wordFinalParagraphAdmissionExtentPt,
  wordVerticalRlFinalLineAdmissionExtentPt,
} from './body-pagination-compatibility.js';
import { sliceParagraphLayout } from './paragraph.js';
import type { ParagraphLayout, WritingMode } from './types.js';

export interface ParagraphFragmentCursor {
  readonly boundary: LineBoundary | null;
  readonly sourceRangeStart?: number;
  readonly uniformRubyAdvancePt?: number;
}

export interface ParagraphFragmentSelection {
  readonly fragment: ParagraphLayout | null;
  readonly nextCursor: ParagraphFragmentCursor | null;
  readonly requiresFreshFlowRegion: boolean;
  readonly additionalReservePt: number;
  /** Flow charge admitted to this region; retained line/ink geometry can be taller. */
  readonly admittedBlockExtentPt: number;
  /** The fit extent this selection compared for `fragment`, excluding its
   * reserve: retained advance for a continuing fragment, else
   * {@link completingParagraphAdmissionExtentPt}. Fit checks made after
   * admission (the fresh-page footnote guard) use this same extent, never the
   * retained advance, so they cannot reject a fragment this selection
   * admitted. 0 without a fragment. */
  readonly admissionExtentPt: number;
}

export type ParagraphFragmentation =
  | Readonly<{ kind: 'indivisible' }>
  | Readonly<{
      kind: 'splittable';
      /** Source cursor after every retained visual line, including the final line. */
      lineEndBoundaries: readonly LineBoundary[];
    }>;

/**
 * The selector's fit extent of a fragment that completes its paragraph.
 * Retained advance stays authoritative for placement and paint; only the
 * page/column-edge fit comparison ignores authored trailing whitespace
 * (WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION), and any retained trailing
 * extent beyond it (for example a bottom border) remains. Vertical final-line
 * admission follows WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION. That edge
 * policy is not evidence for a body/note-band boundary: footnote plans charge
 * the full retained advance there (§17.3.1.33).
 */
function completingParagraphAdmissionExtentPt(
  fragment: ParagraphLayout,
  authoredSpaceAfterPt: number,
  writingMode: WritingMode,
  availableBlockExtentPt: number,
): number {
  const logicalLineBoxExtentPt = wordFinalParagraphAdmissionExtentPt({
    advancePt: fragment.advancePt,
    retainedSpaceAfterPt: fragment.spacing.afterPt,
    authoredSpaceAfterPt,
  });
  return wordVerticalRlFinalLineAdmissionExtentPt({
    paragraph: fragment,
    writingMode,
    logicalLineBoxExtentPt,
    availableBlockExtentPt,
  });
}

function compareLineBoundaries(left: LineBoundary, right: LineBoundary): number {
  return left.segIndex - right.segIndex || left.charOffset - right.charOffset;
}

export function selectParagraphFragment(
  acquired: ParagraphLayout,
  cursor: ParagraphFragmentCursor,
  fragmentation: ParagraphFragmentation,
  availableBlockExtentPt: number,
  freshFlowRegionBlockExtentPt: number,
  canRelocate: boolean,
  policy: Readonly<{
    keepLines: boolean;
    widowControl: boolean;
    /** Authored §17.3.1.33 trailing whitespace; final-fragment fit is governed
     * by WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION. */
    authoredSpaceAfterPt?: number;
    /** Owning section-region flow axis. Vertical final-line admission is
     * governed by WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION. */
    writingMode?: WritingMode;
    /** Lines at or after this index may not be admitted in this flow region:
     * a proven anchor-line deferral (anchor-line-deferral.ts) ends the page
     * just above that line. Keep and widow rules then act as for overflow. */
    lineEndLimit?: number;
  }>,
  additionalReserveFor?: (fragment: ParagraphLayout) => number,
  uniformRubyAdvancePt?: number,
  additionalReserveFits?: (reservePt: number) => boolean,
): ParagraphFragmentSelection {
  if (![availableBlockExtentPt, freshFlowRegionBlockExtentPt].every(
    (value) => Number.isFinite(value) && value >= 0,
  )) throw new RangeError('Paragraph fragment extents must be finite and non-negative');
  if (
    fragmentation.kind === 'splittable'
    && fragmentation.lineEndBoundaries.length !== acquired.lines.length
  ) {
    throw new RangeError('Splittable paragraph source boundaries must align with retained lines');
  }
  if (fragmentation.kind === 'indivisible' && cursor.boundary !== null) {
    throw new RangeError('Indivisible paragraph cannot carry a continuation boundary');
  }
  const authoredSpaceAfterPt = policy.authoredSpaceAfterPt ?? 0;
  if (!Number.isFinite(authoredSpaceAfterPt) || authoredSpaceAfterPt < 0) {
    throw new RangeError('Authored paragraph spaceAfter must be finite and non-negative');
  }
  const total = acquired.lines.length;
  const lineEndLimit = policy.lineEndLimit;
  if (lineEndLimit !== undefined
    && (!Number.isInteger(lineEndLimit) || lineEndLimit < 0 || lineEndLimit >= total)) {
    throw new RangeError('Paragraph line end limit must name a retained line');
  }
  const slice = (end: number) => sliceParagraphLayout(acquired, {
    lineStart: 0,
    lineEnd: end,
    continuesFromPrevious: cursor.boundary !== null,
    continuesOnNext: end < total,
  });
  const reserveFor = (fragment: ParagraphLayout): number => {
    const reserve = additionalReserveFor?.(fragment) ?? 0;
    if (!Number.isFinite(reserve) || reserve < 0) {
      throw new RangeError('Paragraph page-local reserve must be finite and non-negative');
    }
    return reserve;
  };
  const reserveFits = (reservePt: number): boolean => additionalReserveFits?.(reservePt) ?? true;
  const admissionExtent = (fragment: ParagraphLayout, completesParagraph: boolean): number => (
    completesParagraph
      ? completingParagraphAdmissionExtentPt(
        fragment, authoredSpaceAfterPt, policy.writingMode ?? 'horizontal-tb', availableBlockExtentPt,
      )
      : fragment.advancePt
  );
  if (fragmentation.kind === 'indivisible') {
    const completeReserve = reserveFor(acquired);
    const completeExtentPt = admissionExtent(acquired, true);
    if (acquired.nativeReadingRelocations?.length && nativeReadingRequiresFreshColumn(
      completeExtentPt + completeReserve, availableBlockExtentPt, freshFlowRegionBlockExtentPt,
      canRelocate, reserveFits(completeReserve),
    )) return { fragment: null, nextCursor: cursor, requiresFreshFlowRegion: true,
      additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0 };
    if (canRelocate && lineEndLimit !== undefined) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
      };
    }
    if (canRelocate && (
      completeExtentPt + completeReserve > availableBlockExtentPt
      || !reserveFits(completeReserve)
    ) && completeExtentPt + completeReserve <= freshFlowRegionBlockExtentPt) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
      };
    }
    return {
      fragment: acquired,
      nextCursor: null,
      requiresFreshFlowRegion: false,
      additionalReservePt: completeReserve,
      admittedBlockExtentPt: Math.min(acquired.advancePt, availableBlockExtentPt),
      admissionExtentPt: completeExtentPt,
    };
  }
  if (total === 0) {
    const reserve = reserveFor(acquired);
    const completeExtentPt = admissionExtent(acquired, true);
    if (canRelocate && (
      completeExtentPt + reserve > availableBlockExtentPt
      || !reserveFits(reserve)
    )
      && completeExtentPt + reserve <= freshFlowRegionBlockExtentPt) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
      };
    }
    return {
      fragment: acquired, nextCursor: null,
      requiresFreshFlowRegion: false, additionalReservePt: reserve,
      admittedBlockExtentPt: Math.min(acquired.advancePt, availableBlockExtentPt),
      admissionExtentPt: completeExtentPt,
    };
  }
  const completeReserve = reserveFor(acquired);
  const completeExtentPt = admissionExtent(acquired, true);
  if (cursor.boundary === null && policy.keepLines && canRelocate
    && (
      lineEndLimit !== undefined
      || completeExtentPt + completeReserve > availableBlockExtentPt
      || !reserveFits(completeReserve)
    )
    && completeExtentPt + completeReserve <= freshFlowRegionBlockExtentPt) {
    return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
    };
  }
  // Acquisition retains one entry per physical line; the same index owns
  // its full source boundary, including all horizontal gap placements.
  const allGroupEnds = acquired.lines.map((_, index) => index + 1);
  const groupEnds = allGroupEnds.filter(end => end <= (lineEndLimit ?? total));
  if (groupEnds.length === 0) return {
    fragment: null, nextCursor: cursor, requiresFreshFlowRegion: true,
    additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
  };
  // Generic selector contract: the caller's page-local reserve need not grow
  // with the candidate, so a completed fragment can fit although a shorter
  // continuing prefix does not. Test completion independently before the
  // greedy prefix walk; never infer it from prefix fits. Footnote reserves
  // are now committed per reference line and are nondecreasing (see
  // word-reference-line-footnote-allocation), so this branch must not encode
  // any footnote priority. The usual slice/cursor/ownership path below
  // finalizes either case.
  const completeFits = lineEndLimit === undefined
    && completeExtentPt + completeReserve <= availableBlockExtentPt
    && reserveFits(completeReserve);
  let groupEnd = completeFits ? groupEnds.length : selectLargestFittingEnd(
    0,
    groupEnds.length,
    availableBlockExtentPt,
    (groupIndex) => (() => {
      const lineEnd = groupEnds[groupIndex - 1]!;
      const candidate = slice(lineEnd);
      const reserve = reserveFor(candidate);
      return reserveFits(reserve)
        ? admissionExtent(candidate, lineEnd === total) + reserve
        : availableBlockExtentPt + 1;
    })(),
  ).end;
  if (groupEnd === 0) {
    if (canRelocate) return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
    };
    groupEnd = 1;
  }
  for (;;) {
    const widow = adjustForWidowOrphan({
      widowControl: policy.widowControl,
      start: 0,
      end: groupEnd,
      totalLines: allGroupEnds.length,
      canRelocate,
    });
    if (widow.kind === 'relocate') {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0, admissionExtentPt: 0,
      };
    }
    if (widow.kind !== 'dropLastLine') break;
    groupEnd -= 1;
  }
  const end = groupEnds[groupEnd - 1]!;
  const fragment = slice(end);
  const nextBoundary = end < total ? fragmentation.lineEndBoundaries[end - 1]! : null;
  if (
    nextBoundary !== null
    && cursor.boundary !== null
    && compareLineBoundaries(nextBoundary, cursor.boundary) <= 0
  ) {
    throw new Error('Paragraph continuation source boundary did not advance');
  }
  return {
    fragment,
    nextCursor: nextBoundary === null ? null : Object.freeze({
      boundary: nextBoundary,
      sourceRangeStart: fragment.lines.at(-1)!.range.end,
      ...(uniformRubyAdvancePt === undefined ? {} : { uniformRubyAdvancePt }),
    }),
    requiresFreshFlowRegion: false,
    additionalReservePt: reserveFor(fragment),
    admittedBlockExtentPt: Math.min(fragment.advancePt, availableBlockExtentPt),
    admissionExtentPt: admissionExtent(fragment, end === total),
  };
}
