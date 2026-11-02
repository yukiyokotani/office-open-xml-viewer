import { enumerateGaps, type GapScanState, type LineGap, type DistributeOptions } from '@silurus/ooxml-core';
import type { LayoutSeg } from './model.js';

/** Width authority is the breaker's advance or retained layout geometry, never
 * a font-name approximation. Opaque cells include tabs, inline objects and ruby;
 * zero-advance metric hosts have empty text and do not interrupt adjacency. */
export interface GapSegment {
  text?: string;
  widthPx: number;
  spacePx: number;
  /** Contextual per-space advances, indexed in code points. */
  spaceWidths?: ReadonlyMap<number, number>;
  /** Advances of the U+3000 in the segment's trailing U+0020/U+3000 run,
   * indexed in code points, present only where a line-end ideographic space
   * hangs: it then stands outside the visible width like U+0020
   * (WORD_LINE_END_SPACE_HANGS). */
  ideographicSpaceWidths?: ReadonlyMap<number, number>;
  /** Right-to-left text: the line keeps its U+3000 inside the visible width. */
  rtl?: boolean;
  /** WORD_EAST_ASIAN_AUTO_SPACE gap included in `widthPx` before the first
   * character; a gap of the same family as U+0020. */
  leadingGapPx?: number;
}

export const textCellGapOptions: DistributeOptions = {
  textCellSpaceGaps: true,
  atomicSpaceGaps: true,
  isWhitespace: cp => cp === 0x20,
  isGapChar: () => false,
  lastDrawnSi: Infinity,
};

/** Complete opportunities retained once when the accepted line closes. The
 * streaming prefix remains constant-size during candidate comparisons. */
export interface LineGapPlan extends LineGapModel {
  expansionGaps: LineGap[];
}

export interface LineGapModel {
  gaps: LineGap[];
  /** Natural opportunity sum, without the gaps before a hanging U+3000 tail. */
  S: number;
  visibleWidthPx: number;
  lineEndSeparatorPx: number;
  naturalWidthPx: number;
  segmentCount: number;
  /** Raw enumerator state for prefix accumulation (tail gaps included). */
  scan: GapScanState;
  /** Incremental tail state: U+0020 tail, U+0020/U+3000 tail, the natural
   * width of gaps opening into a hanging U+3000 tail, RTL text seen. */
  tail: Readonly<{ latinPx: number; fullPx: number; gapPx: number; rtl: boolean }>;
}

/** WORD_JUSTIFIED_INTERWORD_COMPRESSION: the same enumerator owns fit and
 * positive/negative justification. Consecutive U+0020 all participate only
 * between text cells; NBSP, tab/object/ruby adjacency and line edges stay fixed.
 * Prefix accumulation retains constant-size state, so candidate lookahead and
 * committed appends are linear even with long fragmented words/space runs. */
export function lineGapModel(
  segments: readonly GapSegment[],
  prefix?: Readonly<LineGapModel>,
  collect = true,
): LineGapModel {
  const offset = prefix?.segmentCount ?? 0;
  const enumerate = (input: readonly GapSegment[], collectGaps: boolean) => enumerateGaps(input, {
    ...textCellGapOptions,
    segmentOffset: offset,
    prefix: prefix?.scan,
    collect: collectGaps,
    gapWidth: (si, cpOffset) => input[si - offset].spaceWidths?.get(cpOffset) ?? input[si - offset].spacePx,
    leadingGapWidth: (si) => input[si - offset]?.leadingGapPx ?? 0,
  });
  const enumeration = enumerate(segments, collect);
  let naturalWidthPx = prefix?.naturalWidthPx ?? 0;
  let { latinPx, fullPx, gapPx, rtl } = prefix?.tail ?? { latinPx: 0, fullPx: 0, gapPx: 0, rtl: false };
  // The last unit of this call before the U+0020/U+3000 tail, as [local
  // segment index, code point].
  let lastVisible: readonly [number, number] | undefined;
  segments.forEach((segment, index) => {
    naturalWidthPx += segment.widthPx;
    if (segment.rtl) rtl = true;
    if (segment.text === '') return;
    if (segment.text === undefined) {
      latinPx = 0;
      fullPx = 0;
      lastVisible = [index, 0];
      return;
    }
    if (/^ +$/u.test(segment.text)) {
      latinPx += segment.widthPx;
      fullPx += segment.widthPx;
      return;
    }
    const characters = [...segment.text];
    const spaceWidth = (off: number): number => segment.spaceWidths?.get(off) ?? segment.spacePx;
    let latin = characters.length;
    while (latin > 0 && characters[latin - 1] === ' ') latin -= 1;
    latinPx = 0;
    for (let off = latin; off < characters.length; off++) latinPx += spaceWidth(off);
    let full = characters.length;
    while (full > 0 && (characters[full - 1] === ' '
      || segment.ideographicSpaceWidths?.has(full - 1) === true)) full -= 1;
    let trailingPx = 0;
    for (let off = full; off < characters.length; off++) {
      trailingPx += characters[off] === ' ' ? spaceWidth(off) : segment.ideographicSpaceWidths!.get(off)!;
    }
    if (full === 0) {
      fullPx += trailingPx;
    } else {
      fullPx = trailingPx;
      lastVisible = [index, full - 1];
    }
  });
  // Gaps that open into a hanging U+3000 tail (before or inside it) take no
  // slack. The enumerator state keeps them for later prefixes; their sum is
  // the growth of the opportunity sum after the last visible unit.
  if (fullPx === latinPx) {
    gapPx = 0;
  } else if (lastVisible === undefined) {
    gapPx += enumeration.state.naturalGapSum - (prefix?.scan.naturalGapSum ?? 0);
  } else {
    const [index, cp] = lastVisible;
    const visible = segments.slice(0, index + 1);
    const last = visible[index]!;
    if (last.text !== undefined) visible[index] = { ...last, text: [...last.text].slice(0, cp + 1).join('') };
    gapPx = enumeration.state.naturalGapSum - enumerate(visible, false).state.naturalGapSum;
  }
  const inTail = (gap: LineGap): boolean => lastVisible === undefined
    || gap.segIndex - offset > lastVisible[0]
    || (gap.segIndex - offset === lastVisible[0] && gap.cpOffset >= lastVisible[1]);
  const ideographicTail = !rtl;
  const lineEndSeparatorPx = ideographicTail ? fullPx : latinPx;
  return {
    gaps: ideographicTail && gapPx > 0 ? enumeration.gaps.filter((gap) => !inTail(gap)) : enumeration.gaps,
    S: enumeration.state.naturalGapSum - (ideographicTail ? gapPx : 0),
    naturalWidthPx,
    visibleWidthPx: naturalWidthPx - lineEndSeparatorPx,
    lineEndSeparatorPx,
    segmentCount: offset + segments.length,
    scan: enumeration.state,
    tail: { latinPx, fullPx, gapPx, rtl },
  };
}

/** A complete next text unit, including its separator runs. Opaque cells are
 * indivisible units of their own. No following-word eligibility lookahead. */
export function candidateUnit(first: LayoutSeg, queue: Iterable<LayoutSeg>): LayoutSeg[] {
  const segments = [first];
  if (!('text' in first) || first.ruby || first.fitTextRegionIndex !== undefined) return segments;
  let hasSeparator = first.text.endsWith(' ');
  for (const next of queue) {
    if (!('text' in next) || next.ruby || next.fitTextRegionIndex !== undefined
      || (hasSeparator && /[^ ]/u.test(next.text) && !next.joinPrev && !next.hardJoinPrev)
      // A WORD_EAST_ASIAN_AUTO_SPACE seam is a legal East Asian/Latin break
      // (UAX #14) unless segmentation glued it.
      || (next.autoSpaceBeforePt !== undefined && !next.joinPrev && !next.hardJoinPrev)) break;
    segments.push(next);
    if (next.text.length > 0) hasSeparator = next.text.endsWith(' ');
  }
  return segments;
}
