import type { createLineBreakerState, SnapBlockState } from './break-queue.js';
import { popMixedLineItem } from './mixed-space-fit.js';
import type { LayoutTextSeg } from './model.js';

type BreakerState = ReturnType<typeof createLineBreakerState>;

/** The line state one admission (performAddToLine) changes besides the pen:
 * the line's extents and the per-line facts derived from its items. Every
 * other field of the breaker state belongs to the line as a whole (its band,
 * its float gap, its place in the paragraph) and outlives any one item. */
const LINE_ADMISSION_KEYS = [
  'justifiedUnitEnd',
  'justifiedGapModel',
  'latinLineFace',
  'latinLineHomogeneous',
  'latinUniformGapCapacity',
  'lineLabelShiftPx',
  'lineHeight',
  'lineAscent',
  'lineDescent',
  'lineSignedDescent',
  'lineSignedAscent',
  'lineIntendedSingle',
  'lineHasInlinePicture',
  'linePictureMarkSingle',
  'lineObjectBaselineMinimum',
  'lineObjectAdvanceMinimum',
  'lineGridCountSingle',
  'lineLatinGridCountSingle',
  'lineVisibleAscent',
  'lineVisibleDescent',
  'lineVisibleIntendedSingle',
  'lineHasVisibleMetrics',
  'lineHasRuby',
  'lineEastAsian',
] as const satisfies readonly (keyof BreakerState)[];

type LineAdmissionKey = typeof LINE_ADMISSION_KEYS[number];

/** The line as it was before one item was admitted. */
export interface LineAdmissionState {
  readonly values: { readonly [Key in LineAdmissionKey]: BreakerState[Key] };
  /** The open snapToChars script block and the pads of its two edge members,
   * which a joining member re-allocates. */
  readonly snapBlock: SnapBlockState | null;
  readonly blockLeadingPadPx: number | undefined;
  readonly blockTrailingPadPx: number | undefined;
  readonly latinLineGapCount: number;
}

/** One record per item of the current line, in line order: the line before
 * the item, what the item added to the pen, and the box it was admitted
 * with. Any suffix of the line can therefore be taken back
 * (retractLineItems), and the head of a split item admitted again, with the
 * line's extents recomputed rather than kept. */
export interface LineAdmission {
  readonly before: LineAdmissionState;
  /** The pen advance the admission added to the line. */
  readonly advancePx: number;
  /** The item's own measured width as admitted (a later space compression
   * narrows the item and the line together). */
  readonly measuredPx: number;
  readonly h: number;
  readonly asc: number;
  readonly desc: number;
}

export function captureLineAdmission(breakerState: BreakerState): LineAdmissionState {
  const values = {} as { [Key in LineAdmissionKey]: BreakerState[Key] };
  for (const key of LINE_ADMISSION_KEYS) (values as Record<string, unknown>)[key] = breakerState[key];
  const block = breakerState.snapBlock;
  return {
    values,
    snapBlock: block,
    blockLeadingPadPx: block?.first.snapGridLeadingPadPx,
    blockTrailingPadPx: block?.last.snapGridTrailingPadPx,
    latinLineGapCount: breakerState.latinLineGaps.length,
  };
}

function measuredWidthOf(item: BreakerState['currentLine'][number]): number {
  return 'measuredWidth' in item && typeof item.measuredWidth === 'number' ? item.measuredWidth : 0;
}

/** Record the admission of the line's last item. */
export function recordLineAdmission(
  breakerState: BreakerState,
  before: LineAdmissionState,
  penBeforePx: number,
  h: number,
  asc: number,
  desc: number,
): void {
  breakerState.admissions.push({
    before,
    advancePx: breakerState.currentWidth - penBeforePx,
    measuredPx: measuredWidthOf(breakerState.currentLine.at(-1)!),
    h,
    asc,
    desc,
  });
}

/** Take the line's items from `start` on back. The line returns to its
 * state before item `start` was admitted: its pen, its extents, its open
 * snapToChars block (the block's edge members take their pads back) and its
 * Latin gaps. The caller has materialized any Latin space compression, which
 * stays applied to the items that remain. Returns the items taken. */
export function retractLineItems(
  breakerState: BreakerState,
  start: number,
): BreakerState['currentLine'] {
  const line = breakerState.currentLine;
  const admissions = breakerState.admissions;
  if (admissions.length !== line.length) {
    throw new Error('Line admissions do not describe the current line');
  }
  if (start >= line.length) return [];
  const taken = line.slice(start);
  for (let index = line.length - 1; index >= start; index -= 1) {
    const admission = admissions[index]!;
    breakerState.currentWidth -= admission.advancePx + (measuredWidthOf(line[index]!) - admission.measuredPx);
    popMixedLineItem(breakerState);
  }
  const { before } = admissions[start]!;
  admissions.length = start;
  breakerState.kinsokuBlockedLength = Math.min(breakerState.kinsokuBlockedLength, start);
  Object.assign(breakerState, before.values);
  const block = before.snapBlock;
  if (block) {
    const restorePad = (member: LayoutTextSeg, key: 'snapGridLeadingPadPx' | 'snapGridTrailingPadPx', padPx: number | undefined) => {
      member.measuredWidth += (padPx ?? 0) - (member[key] ?? 0);
      member[key] = padPx;
    };
    restorePad(block.first, 'snapGridLeadingPadPx', before.blockLeadingPadPx);
    restorePad(block.last, 'snapGridTrailingPadPx', before.blockTrailingPadPx);
  }
  breakerState.snapBlock = block;
  breakerState.latinLineGaps.length = before.latinLineGapCount;
  // Materialized compression is no longer a uniform per-gap state.
  breakerState.latinLineHomogeneous = false;
  return taken;
}
