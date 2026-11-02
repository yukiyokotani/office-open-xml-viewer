import { isCjkBreakChar } from './cjk-ranges.js';
import { graphemeClusterOffsets, isSeaScriptCodePoint, isSeaGraphemeExtend } from './sea-break.js';
import type { DistributeSeg, DistributeOptions } from './line-distribute.js';

/** Default inter-word whitespace (WordprocessingML): ASCII space + ideographic
 *  space. */
const defaultIsWhitespace = (cp: number): boolean => cp === 0x20 || cp === 0x3000;

/** One opportunity after a code point; widths are supplied by the caller.
 * `leading` is a caller-measured gap that a segment carries before its first
 * code point (see `leadingGapWidth`), owned by the preceding code point. */
export interface LineGap {
  kind: 'space' | 'boundary' | 'leading';
  codePoint?: number;
  segIndex: number;
  cpOffset: number;
  naturalPx: number;
  trailing: boolean;
}

type GapUnit = { si: number; off: number; cp?: number; ws: boolean; len: number };
/** Constant-size prefix state for fit. Pending spaces become opportunities
 * only when their right neighbour arrives; cloning this state is O(1). */
export interface GapScanState {
  last?: GapUnit;
  hasContent: boolean;
  pendingWidth: number;
  pendingCount: number;
  pendingLastWidth: number;
  naturalGapSum: number;
  gapCount: number;
}

/** Enumerate once for paint, or accumulate a prefix without retaining gaps for
 * fit. Both modes use exactly the same adjacency decisions. Work is linear in
 * code points; neither a space run nor a fragmented word rescans its prefix. */
export function enumerateGaps<T extends DistributeSeg>(
  segments: readonly T[],
  opts: DistributeOptions & {
    prefix?: Readonly<GapScanState>;
    segmentOffset?: number;
    collect?: boolean;
  } = {},
): { gaps: LineGap[]; state: GapScanState } {
  const state: GapScanState = opts.prefix ? { ...opts.prefix } : {
    hasContent: false, pendingWidth: 0, pendingCount: 0, pendingLastWidth: 0, naturalGapSum: 0, gapCount: 0,
  };
  const gaps: LineGap[] = [];
  let pending: LineGap[] = [];
  const collect = opts.collect !== false;
  const offset = opts.segmentOffset ?? 0;
  const final = opts.lastDrawnSi ?? segments.length - 1;
  const isWhitespace = opts.isWhitespace ?? defaultIsWhitespace;
  const isGapChar = opts.isGapChar ?? isCjkBreakChar;
  const width = (u: GapUnit): number => opts.gapWidth?.(u.si, u.off, u.cp ?? 0) ?? 1;
  const descriptor = (u: GapUnit, kind: LineGap['kind'] = 'space'): LineGap => ({
    kind, codePoint: u.cp,
    segIndex: u.si, cpOffset: u.off, naturalPx: width(u), trailing: u.off === u.len - 1,
  });
  const visit = (u: GapUnit): void => {
    const previous = state.last;
    const textCell = (v: GapUnit | undefined): boolean => v?.cp !== undefined && v.cp !== 0x09;
    // A caller-measured gap before a segment's first code point opens only
    // between two text cells on the same line, with no space run between.
    const leadingPx = u.off === 0 && !u.ws && opts.leadingGapWidth ? opts.leadingGapWidth(u.si) : 0;
    if (leadingPx > 0 && state.hasContent && state.pendingCount === 0
      && textCell(previous) && textCell(u) && previous!.si !== final) {
      state.naturalGapSum += leadingPx;
      state.gapCount += 1;
      if (collect) gaps.push({ ...descriptor(previous!, 'leading'), naturalPx: leadingPx });
    }
    if (u.ws) {
      if (state.hasContent && u.si !== final) {
        if (opts.seaClusterSpaceGaps && opts.seaClusterGaps && state.pendingCount === 0
          && previous?.cp !== undefined && isSeaScriptCodePoint(previous.cp) && previous.si !== final) {
          const before = width(previous);
          state.pendingWidth += before;
          state.pendingCount += 1;
          if (collect) pending.push(descriptor(previous, 'boundary'));
        }
        state.pendingLastWidth = width(u);
        state.pendingWidth += state.pendingLastWidth;
        state.pendingCount += 1;
        if (collect) pending.push(descriptor(u));
      }
      // Preserve the nearest non-space cell for the whole space run.
      return;
    }
    if (state.pendingCount > 0) {
      // UAX #29: SPACE followed by an Extend/SpacingMark/ZWJ can be one
      // grapheme even across source seams. Keep the final space's cluster
      // atomic; earlier standalone spaces in the run remain opportunities.
      // The two-scalar query is bounded, independent of line/prefix length.
      if (opts.atomicSpaceGaps && u.cp !== undefined
        && !graphemeClusterOffsets(` ${String.fromCodePoint(u.cp)}`).includes(1)) {
        state.pendingWidth -= state.pendingLastWidth;
        state.pendingCount -= 1;
        if (collect) pending.pop();
      }
      if (!opts.textCellSpaceGaps || (textCell(previous) && textCell(u))) {
        state.naturalGapSum += state.pendingWidth;
        state.gapCount += state.pendingCount;
        if (collect) for (const gap of pending) gaps.push(gap);
      }
      state.pendingWidth = 0;
      state.pendingCount = 0;
      state.pendingLastWidth = 0;
      pending = [];
    } else if (previous && previous.si !== final && state.hasContent) {
      // A whitespace boundary was counted above; never count it twice.
      // lastIsSpace is represented separately below even when a fixed space
      // run (e.g. in the visually-final segment) owns no opportunities.
      if (!lastIsSpace && (
        (previous.cp !== undefined && isGapChar(previous.cp)) ||
        (u.cp !== undefined && isGapChar(u.cp)) ||
        (opts.seaClusterGaps && previous.cp !== undefined && u.cp !== undefined
          && isSeaScriptCodePoint(previous.cp) && isSeaScriptCodePoint(u.cp)
          && !isSeaGraphemeExtend(u.cp))
      )) {
        state.naturalGapSum += width(previous);
        state.gapCount += 1;
        if (collect) gaps.push(descriptor(previous, 'boundary'));
      }
    }
    state.hasContent = true;
    state.last = u;
  };
  let lastIsSpace = (opts.prefix?.pendingCount ?? 0) > 0;
  for (let index = opts.firstContentSi ?? 0; index < segments.length; index++) {
    const seg = segments[index];
    const si = offset + index;
    if (seg.text === undefined) {
      visit({ si, off: 0, ws: false, len: 1 });
      lastIsSpace = false;
      continue;
    }
    const chars = [...seg.text];
    for (let off = 0; off < chars.length; off++) {
      const cp = chars[off].codePointAt(0)!;
      const ws = isWhitespace(cp);
      visit({ si, off, cp, ws, len: chars.length });
      lastIsSpace = ws;
    }
  }
  return { gaps, state };
}
