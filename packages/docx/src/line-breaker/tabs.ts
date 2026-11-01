import type { TabStop, DocSettings } from '../types';
import { nextTabStop, nextTabStopRtl, type ResolvedTabStop } from '../layout/text.js';
import { wordFloatTabStopPosition, wordPositionalTabReferenceBox } from '../layout/line-compatibility.js';
import { type LayoutImageSeg, type LayoutMathSeg, type LayoutSeg, type LayoutTabSeg, type LayoutTextSeg } from './model.js';

/** ECMA-376 §17.15.1.25 — the ABSENT default for `<w:defaultTabStop>`: "If this
 *  element is omitted, then automatic tab stops should be generated at 720
 *  twentieths of a point (0.5")", i.e. 36 pt. Used ONLY as the fallback when a
 *  document carries no `<w:defaultTabStop>`; a document that sets one overrides
 *  this via {@link resolveDefaultTabPt}. Shared by line layout and the
 *  numbered-list marker's retained trailing-tab advance. */
export const DEFAULT_TAB_PT = 36;


/** ECMA-376 §17.15.1.25 — resolve the document's automatic tab-stop interval
 *  (pt): the explicit `<w:defaultTabStop>` value when present, else the spec
 *  absent default of 720 twips (36pt). Mirrors {@link resolveKinsokuRules}: the
 *  resolved value is threaded into both the measure and draw passes so they
 *  agree. */
export function resolveDefaultTabPt(settings: DocSettings | undefined): number {
  const v = settings?.defaultTabStop;
  // §17.15.1.25 defines automatic stops as multiples of the interval, which is
  // undefined for a non-positive interval; fall back to the documented absent
  // default (720 twips = 36pt) so the automatic grid always advances.
  return v != null && v > 0 ? v : DEFAULT_TAB_PT;
}


/** Resolve normative stop eligibility, then project the displaced line frame
 * through WORD_TAB_DISPLACED_READING_FRAME in layout/line-compatibility.ts. */
export function nextLineTabStop(
  pen: number,
  stops: readonly ResolvedTabStop[],
  interval: number,
  leadingShift: number,
): ResolvedTabStop | null {
  const stop = nextTabStop(pen, stops, interval);
  if (!stop || leadingShift <= 0) return stop;
  const custom = stops.some((entry) => entry.alignment !== 'bar' && entry.pos === stop.pos);
  return { ...stop, pos: wordFloatTabStopPosition(stop.pos, custom, leadingShift) };
}

/** §17.3.3.23 / §17.18.71: align a positional tab's following cell within its
 * reference box, independently of ordinary stops, in reading coordinates. */
export function positionalTabTarget(
  alignment: NonNullable<LayoutTabSeg['ptab']>['alignment'],
  left: number,
  right: number,
  followingWidth: number,
): number {
  return alignment === 'left' ? left
    : alignment === 'right' ? right - followingWidth
      : (left + right - followingWidth) / 2;
}

/** One entry in a bidi line's LOGICAL-order sequence, for {@link layoutBidiTabStops}. */
export interface BidiTabItem {
  /** True for a tab segment (its width is (re)computed); false for content. */
  isTab: boolean;
  /** Content width in px (ignored for tabs). Set by the LTR layout pass. */
  width: number;
  /** Logical advance from this item's start to the resolved decimal alignment
   * point. The bidi resolver converts the complete cell's logical prefix to a
   * physical reading-frame offset. */
  decimalOffset?: number;
  ptab?: LayoutTabSeg['ptab'];
  readingGap?: number;
  /** Authored leader retained with a gap already allocated during fitting. */
  leader?: TabStop['leader'];
}


/** Per-segment result of {@link layoutBidiTabStops}. */
export interface BidiTabResult {
  /** New measuredWidth for the tab at this LOGICAL index (non-tabs: unchanged). */
  width: number;
  /** Leader to paint across this tab's span (`'none'`/undefined ⇒ blank). */
  leader?: TabStop['leader'];
}


/** Resolve ST_TabJc aliases to the physical role used by the paragraph's
 * reading frame. `start`/`end` are the strict logical aliases of
 * `left`/`right`; `num` is the leading tab between a list marker and its text.
 * Both the LTR and mirrored bidi algorithms operate in reading-frame
 * coordinates, so the role mapping itself is direction-independent. */
export function tabAlignmentRole(
  alignment: TabStop['alignment'],
): 'leading' | 'center' | 'trailing' | 'decimal' {
  if (alignment === 'center') return 'center';
  if (alignment === 'decimal') return 'decimal';
  if (alignment === 'right' || alignment === 'end') {
    return 'trailing';
  }
  return 'leading';
}


/**
 * ECMA-376 §17.3.1.37 / §17.15.1.25 / §17.18.84 — lay out ONE line of a BIDI
 * (RTL-base) paragraph's tab-aligned cells, returning each tab's width and
 * leader BY LOGICAL INDEX.
 *
 * The LTR layout pass ({@link layoutLines}) resolves tab widths against the pen
 * in LOGICAL order but in a LEFT-to-right frame, which is wrong for a bidi
 * paragraph: a tab advances the pen in READING order, which under an RTL base
 * runs RIGHT-to-LEFT, and the tab-delimited cells then reorder visually. The LTR
 * result lands the trailing content (a TOC page number, a footer field) on the
 * wrong visual side — often overflowing and wrapping to a new line (the "leaders
 * appear/disappear" and "page number on its own row" symptoms of issue #820).
 *
 * This lays the line out in the RTL READING frame: the pen starts at the right
 * TEXT MARGIN (pen 0) and moves LEFT (increasing pen). A tab stop's `pos` is its
 * distance from that margin, not from the paragraph's indented edge. Content
 * begins at `startPenPx` (the paragraph's leading-indent + first-line
 * indent); the Nth tab in reading order advances to the next stop further left
 * (larger `pos`), exactly like the LTR pen advances rightward through stops.
 * Alignment is logical (Part 4 §14.11.2): physical `left` = `start` (leading ⇒
 * following content's leading/RIGHT edge on the stop), physical `right` = `end`
 * (trailing ⇒ its trailing/LEFT edge on the stop); `center` is unchanged;
 * `bar`/`clear` advance like `start`. Automatic stops fall on the §17.15.1.25
 * grid from the margin, after all custom stops. A stop past the LEFT text
 * margin (`leftLimitPx`) invokes `word-tab-stop-page-edge-clamp`.
 *
 * The widths returned here reproduce the intended layout through the draw
 * loop's visual walk because {@link computeLineVisualOrder} classifies tabs as UAX#9 S:
 * rule L2 then reverses cells AND tabs together, so the logical tab between
 * cells k−1 and k sits visually between the mirrored cells k and k−1 — its
 * reading-frame gap IS its visual gap. (This is why results map back by logical
 * index; resolving stops against the visual sequence would reverse the tab→stop
 * assignment and paint the leader in the wrong cell gap — the #830 follow-up
 * bug where the TOC leader appeared between the title and the chapter number
 * instead of between the page number and the title.)
 *
 * @param items line segments in LOGICAL order (as `line.segments`).
 * @param customStopsPx custom tab stops in margin px (`pos * scale`).
 * @param startPenPx reading-frame pen at the line's content start = the leading
 *   (logical-left ⇒ physical-right) indent, plus any first-line indent.
 * @param leftLimitPx reading-frame position of the LEFT text margin (= the
 *   margin-to-margin text width).
 * @param intervalPx automatic-stop interval = `defaultTabPt * scale`.
 * @returns one {@link BidiTabResult} per LOGICAL index (1:1 with `items`).
 */
export function layoutBidiTabStops(
  items: BidiTabItem[],
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  startPenPx: number,
  leftLimitPx: number,
  intervalPx: number,
  frame?: Readonly<{ leadingShift: number; indentStart: number; indentEnd: number; bandStart: number; bandEnd: number; fitEnd?: number; narrowed: boolean }>,
): BidiTabResult[] {
  return resolveBidiTabStops(items, customStopsPx, startPenPx, leftLimitPx, intervalPx, frame).results;
}

function resolveBidiTabStops(
  items: BidiTabItem[],
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  startPenPx: number,
  leftLimitPx: number,
  intervalPx: number,
  frame?: Readonly<{ leadingShift: number; indentStart: number; indentEnd: number; bandStart: number; bandEnd: number; fitEnd?: number; narrowed: boolean }>,
): { results: BidiTabResult[]; endPen: number } {
  const n = items.length;
  const width = items.map((it) => it.width);
  const leader: (TabStop['leader'] | undefined)[] = new Array(n).fill(undefined);

  // Width of the content run immediately FOLLOWING index `i` in reading order
  // (up to the next tab / line end) — the trailing/centered stop needs it.
  // `word-rtl-decimal-tab-physical-alignment`
  const followAlignmentWidth = (
    from: number,
    role: ReturnType<typeof tabAlignmentRole>,
  ): Readonly<{ total: number; alignment: number }> => {
    let total = 0;
    let decimal: number | undefined;
    for (let j = from; j < n; j++) {
      if (items[j].isTab) break;
      if (decimal === undefined && items[j].decimalOffset !== undefined) {
        decimal = total + items[j].decimalOffset!;
      }
      total += width[j];
    }
    return {
      total,
      alignment: role === 'center'
        ? total / 2
        : role === 'trailing'
          ? total
          : role === 'decimal'
            ? decimal === undefined ? 0 : total - decimal
            : 0,
    };
  };

  // Reading-frame walk. `pen` = distance from the right TEXT MARGIN; content and
  // tabs push it further LEFT (increasing).
  let pen = startPenPx;
  for (let i = 0; i < n; i++) {
    const it = items[i];
    if (!it.isTab) {
      pen += width[i];
      continue;
    }
    if (it.readingGap !== undefined) {
      width[i] = it.readingGap;
      leader[i] = it.leader;
      pen += width[i];
      continue;
    }
    if (it.ptab && frame) {
      const following = followAlignmentWidth(i + 1, 'leading').total;
      const referenceStart = it.ptab.relativeTo === 'indent' ? frame.indentStart : 0;
      const referenceEnd = it.ptab.relativeTo === 'indent' ? frame.indentEnd : leftLimitPx;
      const box = wordPositionalTabReferenceBox(referenceStart, referenceEnd,
        frame.bandStart, frame.bandEnd, frame.narrowed);
      const target = positionalTabTarget(it.ptab.alignment, box.start, box.end, following);
      // §17.3.3.23 reachability was checked by the queue-owning iterator.
      // An unreachable target on an empty line can only contribute zero gap.
      // The fitted prefix, rather than the full queued cell, owns alignment.
      // The final projection may reduce a gap but cannot enlarge its allocation
      // past the fitting band (library policy, not an Office overflow fallback).
      width[i] = Math.max(0, Math.min(target,
        (it.ptab.alignment === 'left' ? frame.fitEnd ?? frame.bandEnd : frame.bandEnd) - following) - pen);
      pen += width[i];
      continue;
    }
    const stop = frame
      ? nextLineTabStop(pen, customStopsPx, intervalPx, frame.leadingShift)
      : nextTabStopRtl(pen, customStopsPx, intervalPx);
    if (!stop) {
      // No stop further left: the tab collapses (following content continues).
      width[i] = 0;
      continue;
    }
    // The tab's leading (right) edge sits at the pen; its trailing (left) edge
    // is the stop-aligned target, giving the gap it fills.
    const role = tabAlignmentRole(stop.alignment);
    const following = followAlignmentWidth(i + 1, role);
    const fw = following.total;
    let target: number; // pen value after the tab (its trailing/left edge)
    if (role !== 'leading') {
      // end aligns the full following cell, center half of it, and decimal the
      // reading-frame distance through the first halfwidth period. The
      // registered no-separator fallback uses the numeric cell's physical
      // right edge, which is the reading-leading edge in this mirrored frame.
      target = stop.pos - following.alignment;
    } else {
      // start/leading (or bar/clear/left): following content's LEADING (right)
      // edge on the stop.
      target = stop.pos;
    }
    // Preserve the established text-margin clamp when no float narrows the
    // line. §17.3.1.12 allows a negative trailing indent to extend the authored
    // band past that margin; expanding the clamp to that edge would move cells
    // that already fit (and charge larger earlier gaps before later-cell fit).
    // A positive trailing indent still limits allocation to its authored band.
    // Only a narrowed float window replaces the margin clamp with its available
    // edge. This is library containment policy, not an Office overflow rule.
    const bandLimit = frame?.narrowed
      ? frame.bandEnd
      : Math.min(leftLimitPx, frame?.bandEnd ?? leftLimitPx);
    // WORD_KASHIDA_LINE_FIT: content after a leading tab ends within the fit
    // width; an aligned cell may reach its stop anywhere in the band.
    const trailingLimit = frame?.fitEnd === undefined || role !== 'leading'
      ? bandLimit : Math.min(bandLimit, frame.fitEnd);
    if (target + fw > trailingLimit) target = trailingLimit - fw;
    // Never let a tab move the pen backwards (right).
    if (target < pen) target = pen;
    width[i] = target - pen;
    leader[i] = stop.leader;
    pen = target;
  }

  return { results: items.map((_, i) => ({ width: width[i], leader: leader[i] })), endPen: pen };
}


export interface BidiTabPostPassInput {
  readonly baseRtl: boolean;
  /** WORD_KASHIDA_LINE_FIT: the line's fit width when it is narrower than its
   * band; resolved cells end within it, while stops keep the band's frame. */
  readonly fitWidth?: number;
  readonly currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[];
  readonly marginRightPx: number;
  readonly maxWidth: number;
  readonly lineXOffset: number;
  readonly lineMaxWidth: number;
  readonly isFirst: boolean;
  readonly firstIndent: number;
  readonly tabOriginPx: number;
  readonly bidiCustomStopsPx: {
    pos: number;
    alignment: NonNullable<TabStop['alignment']>;
    leader?: TabStop['leader'];
  }[];
  readonly bidiIntervalPx: number;
  readonly decimalAlignmentPoint: (segments: readonly LayoutSeg[]) => Readonly<{
    segmentIndex: number;
    charOffset: number;
  }> | null;
  readonly strAdvance: (segment: LayoutTextSeg, text: string) => number;
}


/** One reading-frame projection for positional-tab reachability and the
 * final bidi walk. Both must compare the pen with the same float/indent band. */
export function bidiTabFrame(input: Pick<BidiTabPostPassInput,
  'marginRightPx' | 'maxWidth' | 'lineXOffset' | 'lineMaxWidth' |
  'isFirst' | 'firstIndent' | 'tabOriginPx' | 'fitWidth'>) {
  const { marginRightPx, maxWidth, lineXOffset, lineMaxWidth, isFirst, firstIndent, tabOriginPx } = input;
  const startPen = marginRightPx - (lineXOffset + lineMaxWidth) + (isFirst ? firstIndent : 0);
  const bandEnd = marginRightPx - lineXOffset;
  return {
    startPen,
    leftLimit: marginRightPx + tabOriginPx,
    frame: {
      leadingShift: maxWidth - (lineXOffset + lineMaxWidth),
      indentStart: marginRightPx - maxWidth,
      indentEnd: marginRightPx,
      bandStart: marginRightPx - (lineXOffset + lineMaxWidth) + (isFirst ? Math.min(0, firstIndent) : 0),
      bandEnd,
      // Where fitted content must end: the band end, or the narrower fit width.
      fitEnd: input.fitWidth === undefined ? bandEnd : Math.min(bandEnd, startPen + input.fitWidth),
      narrowed: lineXOffset !== 0 || lineMaxWidth !== maxWidth,
    },
  };
}

/** Resolve bidi tab positions after line content is known, in the visual frame. */
export function applyBidiTabPostPass(input: BidiTabPostPassInput): number {
  if (!input.baseRtl || !input.currentLine.some((segment) => 'isTab' in segment)) return 0;
  return resolveBidiTabRange(input, 0, bidiTabFrame(input).startPen).delta;
}

/** Called only at a tab boundary: all preceding cells are complete. Tabs reset
 * snap/space fitting blocks, so appending the next cell cannot change that
 * prefix. Carry its exact reading pen instead of revisiting it at every tab.
 * The new tab remains provisional; flush still projects the final whole line
 * after any trailing-cell fitting/retraction. Thus each completed cell is
 * visited once here and once at flush, including collapsed zero-width tabs.
 * This is a traversal optimization of §17.3.1.37, not a new tab policy. */
export function createBidiTabCellResolver(): (input: BidiTabPostPassInput) => number {
  let line: BidiTabPostPassInput['currentLine'] | undefined;
  let end = 0;
  let pen = 0;
  return (input) => {
    if (line !== input.currentLine) {
      line = input.currentLine;
      end = 0;
      pen = bidiTabFrame(input).startPen;
    }
    const resolved = resolveBidiTabRange(input, end, pen);
    end = input.currentLine.length;
    pen = resolved.endPen;
    return resolved.delta;
  };
}

function resolveBidiTabRange(
  input: BidiTabPostPassInput,
  from: number,
  startPen: number,
): { delta: number; endPen: number } {
  const {
    baseRtl,
    bidiCustomStopsPx,
    bidiIntervalPx,
    decimalAlignmentPoint,
    strAdvance,
  } = input;
  if (!baseRtl) return { delta: 0, endPen: startPen };
  const currentLine = input.currentLine.slice(from);
  // LOGICAL order — the reading-frame walk resolves the Nth tab against the
  // Nth-reachable stop in the logical reading frame. Do not feed the visual
  // sequence here: UAX#9 L2 reverses cells AND tabs together, so a
  // visual-order walk assigns the stops in reverse and paints the leader in
  // the wrong cell gap (the #830 follow-up bug — the TOC underscore leader
  // appeared between the title and the chapter number instead of between the
  // page number and the title). Because the reversal is symmetric, each
  // logical tab's reading-frame gap IS its visual gap, so widths mapped back
  // by logical index tile correctly under the draw loop's visual walk.
  const items: BidiTabItem[] = currentLine.map((s) => ({
    isTab: 'isTab' in s,
    width: s.measuredWidth,
    ptab: 'isTab' in s ? s.ptab : undefined,
    readingGap: 'isTab' in s ? s.readingGap : undefined,
    leader: 'isTab' in s ? s.leader : undefined,
  }));
  for (let tabIndex = 0; tabIndex < currentLine.length; tabIndex += 1) {
    if (!('isTab' in currentLine[tabIndex]!)) continue;
    let cellEnd = tabIndex + 1;
    while (cellEnd < currentLine.length && !('isTab' in currentLine[cellEnd]!)) {
      cellEnd += 1;
    }
    const cell = currentLine.slice(tabIndex + 1, cellEnd);
    const point = decimalAlignmentPoint(cell);
    if (!point) continue;
    const itemIndex = tabIndex + 1 + point.segmentIndex;
    const segment = currentLine[itemIndex]!;
    if ('text' in segment) {
      items[itemIndex]!.decimalOffset = strAdvance(
        segment,
        segment.text.slice(0, point.charOffset),
      );
    }
  }
  // Margin-anchored frame (§17.3.1.37 — stops measure from the TEXT MARGIN):
  // pen 0 = right text margin. Content starts after the leading indent — the
  // line window's RIGHT edge is paraX-relative `lineXOffset + lineMaxWidth`
  // (= maxWidth when no float narrows it), so its margin distance is
  // marginRightPx minus that — plus the first line's first-line indent
  // (which narrows the leading edge under an RTL base, mirroring the draw
  // loop's `effAvailW`). The left text margin sits tabOriginPx past the
  // paragraph box (its trailing indent).
  const { leftLimit, frame } = bidiTabFrame(input);
  const { results: res, endPen } = resolveBidiTabStops(items, bidiCustomStopsPx, startPen, leftLimit, bidiIntervalPx, frame);
  let delta = 0;
  for (let i = 0; i < currentLine.length; i++) {
    const s = currentLine[i];
    if (!('isTab' in s)) continue;
    delta += res[i].width - s.measuredWidth;
    s.measuredWidth = res[i].width;
    (s as LayoutTabSeg).leader = res[i].leader;
  }
  return { delta, endPen };
}
