import { advanceLineBoundary } from './source-cursor.js';
import { retractLineItems } from './line-admission.js';
import { placedUnitOf } from './ruby-cell.js';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import { type LayoutImageSeg, type LayoutMathSeg, type LayoutTabSeg, type LayoutTextSeg } from './model.js';
import { isInklessLineItem, type createLineBreakerState } from './break-queue.js';
import { RESET_SLICED_TEXT_MEASUREMENT, protectedNoBreakOffsets, slicedTextMetadata } from './advance.js';
import { rebaseSeaBreaks } from './text-runs.js';


export type CrossRunKinsokuRetraction =
  | { readonly kind: 'none' }
  | { readonly kind: 'blocked' }
  | {
    readonly kind: 'retracted';
    readonly tail: LayoutTextSeg;
    /** The items retracted after `tail`, in line order. */
    readonly followers?: readonly LayoutTextSeg[];
  };

/** Queue a retraction ahead of the segment that caused it. */
export function queueRetraction(
  queue: { unshift(segment: LayoutTextSeg): void },
  retraction: Extract<CrossRunKinsokuRetraction, { kind: 'retracted' }>,
): void {
  for (const follower of [...(retraction.followers ?? [])].reverse()) queue.unshift(follower);
  queue.unshift(retraction.tail);
}

type BreakerState = ReturnType<typeof createLineBreakerState>;
type LineItem = BreakerState['currentLine'][number];

/** What a retraction needs of its pass. */
export interface KinsokuLine {
  readonly breakerState: BreakerState;
  readonly kinsoku: KinsokuRules;
  readonly materializeLatinSpaceCompression: () => void;
  readonly strAdvance: (
    segment: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression?: boolean,
  ) => number;
  /** The advance of `text` before any character-grid allocation. */
  readonly strNaturalAdvance: (segment: LayoutTextSeg, text: string) => number;
  /** The pass's one admission (performAddToLine): the head of a split item
   * is admitted again through it. */
  readonly addToLine: (segment: LineItem, width: number, height: number, ascent: number, descent: number) => void;
}

/** A boundary of the current line: before item `index`, or `offset` UTF-16
 * units into that text item. The line end is `{ index: line.length, offset: 0 }`. */
interface LinePoint {
  readonly index: number;
  readonly offset: number;
}

/** Where the current line can end so that what moves to the next line, and
 * what stays, both keep §17.3.1.16. */
type RetractionPoint = LinePoint | 'none' | 'blocked';

const isText = (item: LineItem | undefined): item is LayoutTextSeg => item !== undefined && 'text' in item;

/** The grapheme cluster boundaries of a line item's text (0 and its length
 * included), computed once per item. */
const clusterBoundaries = new WeakMap<LayoutTextSeg, readonly number[]>();

function boundariesOf(item: LayoutTextSeg): readonly number[] {
  let bounds = clusterBoundaries.get(item);
  if (!bounds) {
    bounds = [...new Set([0, ...graphemeClusterOffsets(item.text), item.text.length])].sort((left, right) => left - right);
    clusterBoundaries.set(item, bounds);
  }
  return bounds;
}

/** The start of the grapheme cluster of `item` that ends at or spans the
 * position before `offset`: a retraction never cuts a cluster (a base with
 * its combining marks or variation selector). */
function clusterStartBefore(item: LayoutTextSeg, offset: number): number {
  const bounds = boundariesOf(item);
  let low = 0;
  let high = bounds.length - 1;
  while (low < high) {
    const middle = (low + high + 1) >>> 1;
    if (bounds[middle]! < offset) low = middle;
    else high = middle - 1;
  }
  return bounds[low]!;
}

/** Items the breaker never divides are one unit of the search each: a ruby
 * cell with the fitText regions it shares a segment with (placedUnitOf), a
 * fitText region without ruby (§17.3.2.14) and a tate-chu-yoko cell
 * (§17.3.2.10). */
function sameUnit(left: LayoutTextSeg, right: LayoutTextSeg): boolean {
  const unit = placedUnitOf(left);
  if (unit || placedUnitOf(right)) return unit === placedUnitOf(right);
  return left.fitTextRegionIndex !== undefined && left.fitTextRegionIndex === right.fitTextRegionIndex;
}

function indivisible(item: LayoutTextSeg): boolean {
  return placedUnitOf(item) !== undefined || item.fitTextRegionIndex !== undefined || item.tateChuYoko === true;
}

/** The boundary one unit before `point`: a grapheme cluster of ordinary
 * text, a whole indivisible unit, or an empty text item. Undefined at the
 * line start and before an item that is not text, which no retraction
 * crosses. */
function previousPoint(line: readonly LineItem[], point: LinePoint): LinePoint | undefined {
  if (point.offset > 0) {
    return { index: point.index, offset: clusterStartBefore(line[point.index] as LayoutTextSeg, point.offset) };
  }
  const item = line[point.index - 1];
  if (!isText(item)) return undefined;
  if (indivisible(item)) {
    let start = point.index - 1;
    while (start > 0 && isText(line[start - 1]) && sameUnit(line[start - 1] as LayoutTextSeg, item)) start -= 1;
    return { index: start, offset: 0 };
  }
  return { index: point.index - 1, offset: item.text.length === 0 ? 0 : clusterStartBefore(item, item.text.length) };
}

/** The character that would end the line broken at `point`: the base code
 * point of the last cluster before it. */
function codePointEndingAt(line: readonly LineItem[], point: LinePoint): number | undefined {
  if (point.offset > 0) {
    const item = line[point.index] as LayoutTextSeg;
    return item.text.codePointAt(clusterStartBefore(item, point.offset));
  }
  for (let index = point.index - 1; index >= 0; index -= 1) {
    const item = line[index];
    if (!isText(item)) return undefined;
    if (item.text.length > 0) return item.text.codePointAt(clusterStartBefore(item, item.text.length));
  }
  return undefined;
}

/** The grapheme cluster that would lead the next line broken at `point`,
 * `next` being the text after the line. */
function clusterLeadingAt(line: readonly LineItem[], point: LinePoint, next: string): string | undefined {
  for (let index = point.index, offset = point.offset; index < line.length; index += 1, offset = 0) {
    const item = line[index];
    if (!isText(item)) return undefined;
    if (offset < item.text.length) {
      const end = boundariesOf(item).find((bound) => bound > offset) ?? item.text.length;
      return item.text.slice(offset, end);
    }
  }
  const end = graphemeClusterOffsets(next).find((bound) => bound > 0) ?? next.length;
  return next.slice(0, end);
}

/** The end of the line's text before item `index`, through as many text
 * items as it takes: enough of it to hold the cluster that ends there, which
 * may itself lie in several items (an emoji sequence whose joiner is a
 * segment of its own). */
function textBefore(line: readonly LineItem[], index: number): string {
  let text = '';
  for (let at = index - 1; at >= 0 && text.length < CLUSTER_CONTEXT_UNITS; at -= 1) {
    const item = line[at];
    if (!isText(item)) break;
    text = item.text.slice(-CLUSTER_CONTEXT_UNITS) + text;
  }
  return text;
}

/** UTF-16 units of context that decide whether text continues a cluster:
 * longer than any cluster of ordinary text (an emoji joiner sequence, a base
 * with its marks). */
const CLUSTER_CONTEXT_UNITS = 64;

/** How far into `after` the grapheme cluster that ends `before` reaches
 * (UAX #29): 0 when `after` opens a cluster of its own. */
function clusterContinuation(before: () => string, after: string): number {
  // Nothing below U+0300 continues a cluster (CR LF apart, which no segment
  // seam divides), so ordinary Latin text never asks for the line's text.
  if (after === '' || after.charCodeAt(0) < 0x300) return 0;
  return continuationOf(before(), after);
}

function continuationOf(before: string, after: string): number {
  if (before === '') return 0;
  const head = after.slice(0, CLUSTER_CONTEXT_UNITS);
  const end = graphemeClusterOffsets(before + head).find((bound) => bound >= before.length)
    ?? before.length + head.length;
  return end - before.length;
}

/** The length of the prefix of `segment` that belongs to the grapheme
 * cluster ending the line: combining marks, a variation selector or the
 * rest of an emoji sequence that open a run of other formatting. That
 * prefix belongs to the line's last character, not to the text after it. */
export function lineClusterContinuation(line: readonly LineItem[], segment: LayoutTextSeg): number {
  return clusterContinuation(() => textBefore(line, line.length), segment.text);
}

/** A boundary between two items lies inside a grapheme cluster when the
 * text after it continues the cluster before it: a variation selector or a
 * combining mark in a run of its own stays with its base. Boundaries inside
 * an item are that item's cluster boundaries already. */
function splitsCluster(line: readonly LineItem[], point: LinePoint, lead: string | undefined): boolean {
  if (point.offset > 0 || lead === undefined) return false;
  return clusterContinuation(() => textBefore(line, point.index), lead) > 0;
}

const WHITESPACE_RE = /\s/u;

/** The line keeps content of its own before item `index`: a list label or
 * an inkless item (an anchor character, a mark metric) is not content. */
function contentBefore(line: readonly LineItem[], index: number): boolean {
  for (let at = index - 1; at >= 0; at -= 1) {
    const item = line[at]!;
    if (!isInklessLineItem(item) && !('listLabel' in item && item.listLabel !== undefined)) return true;
  }
  return false;
}

/** §17.3.1.16 over the whole line: the last boundary of the line, at or
 * before its end (`mustMove`: strictly before it, the unit after the line
 * being line-start-forbidden), that leaves no line-end-forbidden character
 * ending the line and no line-start-forbidden character or space leading
 * the next. Units are crossed whole, through any number of segments: the
 * rule is about characters, so a run seam does not end the search,
 * consecutive openers of different runs move together, and a ruby cell is
 * never cut. 'blocked': no such boundary leaves content on the line.
 *
 * A boundary inside the line is legal or not whatever follows the line, so
 * a search that found none need not be repeated: the line's length at that
 * search is kept (`kinsokuBlockedLength`), and a later search over the
 * longer line stops there. A line of forbidden characters that grows one
 * overflowing character at a time is therefore searched in linear work. */
function lineRetractionPoint(
  breakerState: BreakerState,
  kinsoku: KinsokuRules,
  next: string,
  mustMove: boolean,
): RetractionPoint {
  const line = breakerState.currentLine;
  let point: LinePoint | undefined = { index: line.length, offset: 0 };
  if (mustMove) point = previousPoint(line, point);
  else {
    // The line may end where it does unless that leaves a line-end-forbidden
    // character there or cuts a grapheme cluster (the text after the line
    // opens with a mark of the line's last character).
    const end = codePointEndingAt(line, point);
    if ((end === undefined || !kinsoku.lineEndForbidden.has(end))
      && !splitsCluster(line, point, clusterLeadingAt(line, point, next))) return 'none';
  }
  for (; point && point.index >= breakerState.kinsokuBlockedLength; point = previousPoint(line, point)) {
    const leadCluster = clusterLeadingAt(line, point, next);
    const lead = leadCluster?.codePointAt(0);
    const end = codePointEndingAt(line, point);
    if (splitsCluster(line, point, leadCluster)) continue;
    if (lead !== undefined
      && (WHITESPACE_RE.test(String.fromCodePoint(lead)) || kinsoku.lineStartForbidden.has(lead))) continue;
    if (end !== undefined && kinsoku.lineEndForbidden.has(end)) continue;
    if (point.offset === 0 && !contentBefore(line, point.index)) break;
    return point;
  }
  breakerState.kinsokuBlockedLength = line.length;
  return 'blocked';
}

/** A segment as it leads, or follows on, the next line after a retraction. */
function requeued(
  context: KinsokuLine,
  item: LayoutTextSeg,
  start: number,
  leadsLine: boolean,
): LayoutTextSeg {
  // A placed unit's members keep their identity: they are their unit's
  // members and their cell's pieces.
  if (placedUnitOf(item)) return item;
  const text = item.text.slice(start);
  const slice: LayoutTextSeg = {
    ...item,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text,
    ...slicedTextMetadata(item, start, item.text.length),
  };
  return {
    ...slice,
    measuredWidth: context.strAdvance(slice, text, true),
    // The retraction creates a real line boundary. The old source seam was
    // either retained in the head or was soft; it must not be projected
    // onto this newly-created suffix.
    ...(leadsLine ? { joinPrev: undefined, hardJoinPrev: undefined } : {}),
    src: advanceLineBoundary(item.src!, start),
    seaBreaks: rebaseSeaBreaks(item.seaBreaks, start),
  };
}

/** The advance of the line from `point` on, as it would lead the next line. */
function retractedAdvance(context: KinsokuLine, point: LinePoint): number {
  const line = context.breakerState.currentLine;
  let total = 0;
  for (let index = point.index; index < line.length; index += 1) {
    const item = line[index] as LayoutTextSeg;
    total += index === point.index && point.offset > 0
      ? requeued(context, item, point.offset, true).measuredWidth
      : item.measuredWidth;
  }
  return total;
}

/** End the current line at `point`: everything after it is taken back, in
 * any number of items, and returned to be queued. The line returns to its
 * state before the first retracted item was admitted (line-admission.ts);
 * an item cut by the point has its head admitted again, so the line's
 * extents, grid allocation and gaps are those of the content it keeps.
 * Blocked without cutting when the point is a protected source seam. */
function retractLineFrom(context: KinsokuLine, point: LinePoint): CrossRunKinsokuRetraction {
  const { breakerState } = context;
  const line = breakerState.currentLine;
  const cut = line[point.index] as LayoutTextSeg;
  // Moving a whole segment would cut the hard source seam immediately
  // before it. A protected no-break edge is equally indivisible.
  if (point.offset === 0 ? cut.hardJoinPrev === true : protectedNoBreakOffsets(cut).has(point.offset)) {
    return { kind: 'blocked' };
  }
  context.materializeLatinSpaceCompression();
  const admission = breakerState.admissions[point.index]!;
  const taken = retractLineItems(breakerState, point.index) as LayoutTextSeg[];
  const moved = taken.map((item, index) => requeued(context, item, index === 0 ? point.offset : 0, index === 0));
  if (point.offset > 0) {
    // The head keeps its first glyph and predecessor, so it keeps the
    // incoming WORD_KERNED_SPACE_PAIRS seam pair (bounded by its own advance).
    // It is admitted by its natural width: the admission allocates it on
    // the character grid, in the script block it joins.
    const headText = cut.text.slice(0, point.offset);
    const headNaturalW = context.strNaturalAdvance(cut, headText);
    const seamPair = Math.max(cut.leadingWordBoundaryPx ?? 0, -headNaturalW);
    const headW = headNaturalW + seamPair;
    context.addToLine({
      ...cut,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: headText,
      measuredWidth: headW,
      ...(seamPair === 0 ? {} : { leadingWordBoundaryPx: seamPair }),
      ...slicedTextMetadata(cut, 0, point.offset),
    }, headW, admission.h, admission.asc, admission.desc);
    breakerState.latinLineHomogeneous = false;
  }
  return { kind: 'retracted', tail: moved[0]!, followers: moved.slice(1) };
}

function retractAt(context: KinsokuLine, point: RetractionPoint): CrossRunKinsokuRetraction {
  return typeof point === 'string' ? { kind: point } : retractLineFrom(context, point);
}

/** §17.3.1.16: `next`, which does not fit the current line, would lead the
 * next line with a line-start-forbidden character. The end of the current
 * line moves down to lead it instead (cross-run 追い出し). 'blocked' when
 * nothing can move: the caller keeps the leader on the current line. */
export function retractLeadingKinsoku(context: KinsokuLine, next: LayoutTextSeg): CrossRunKinsokuRetraction {
  const { kinsoku } = context;
  const line = context.breakerState.currentLine;
  const firstCp = next.text.codePointAt(0);
  const last = line.at(-1);
  if (firstCp === undefined || !kinsoku.lineStartForbidden.has(firstCp) || !isText(last)) {
    return { kind: 'none' };
  }
  const point = kinsoku.enabled ? lineRetractionPoint(context.breakerState, kinsoku, next.text, true) : 'none';
  return retractAt(context, point);
}

/** §17.3.1.16: a line must not end with a line-end-forbidden character (an
 * opening bracket), and UAX #29: nor inside a grapheme cluster. When `next`
 * moves to the next line, the line's forbidden end (every consecutive
 * opener, whichever runs they are of, a ruby cell ending in one, or the
 * character whose mark opens `next`) moves with it, provided `admits` the
 * moved advance beside the unit `next` leads on a line; otherwise it cannot
 * be kept from a line edge and stays. */
export function retractLineEndForbidden(
  context: KinsokuLine,
  next: LayoutTextSeg,
  admits: (retractedAdvancePx: number) => boolean,
): CrossRunKinsokuRetraction {
  const { kinsoku } = context;
  const line = context.breakerState.currentLine;
  const end = { index: line.length, offset: 0 };
  // A mark is never parted from its base, whatever fits and whether or not
  // kinsoku applies; an opener stays when the move does not fit.
  const cutsCluster = splitsCluster(line, end, clusterLeadingAt(line, end, next.text));
  if (!kinsoku.enabled && !cutsCluster) return { kind: 'none' };
  const point = lineRetractionPoint(context.breakerState, kinsoku, next.text, false);
  if (typeof point === 'string') return { kind: 'none' };
  if (!cutsCluster && !admits(retractedAdvance(context, point))) return { kind: 'none' };
  const retraction = retractLineFrom(context, point);
  return retraction.kind === 'retracted' ? retraction : { kind: 'none' };
}


/** Keep a forbidden leader with its owner when retraction is blocked. */
export function keepLeadingKinsoku(
  breakerState: ReturnType<typeof createLineBreakerState>,
  strNaturalAdvance: (segment: LayoutTextSeg, text: string) => number,
  addToLine: (
    segment: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    width: number,
    height: number,
    ascent: number,
    descent: number,
  ) => void,
  queueEmergencyTail: (segment: LayoutTextSeg, split: number) => void,
  segment: LayoutTextSeg,
  h: number,
  asc: number,
  desc: number,
): boolean {
  const firstEnd = graphemeClusterOffsets(segment.text)[0] ?? segment.text.length;
  if (firstEnd <= 0) return false;
  const prefix = segment.text.slice(0, firstEnd);
  const prefixWidth = strNaturalAdvance(segment, prefix);
  addToLine(
    {
      ...segment,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: prefix,
      measuredWidth: prefixWidth,
      ...slicedTextMetadata(segment, 0, firstEnd),
    },
    prefixWidth,
    h,
    asc,
    desc,
  );
  if (firstEnd < segment.text.length) queueEmergencyTail(segment, firstEnd);
  return true;
}
