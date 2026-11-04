import type { LineVisualOrder } from '../bidi-line.js';

export interface BidiSourceSpaceSlot {
  readonly sourceIndex: number;
  readonly slotIndex: number;
  readonly side: 'left' | 'right';
}

/** WORD_LINE_END_SPACE_STAYS_IN_RUN: a closing space at the other level
 * moves outside its content-level block under UAX #9 L2. Painting all word
 * slices in that space's direction preserves glyph order, but also circulates
 * the source whitespace between slots. Equal spaces hide the difference;
 * three closing spaces or a formatting seam expose it. Retain the exact
 * source-to-physical-slot permutation rather than guessing an edge width.
 * Maps are acquisition-local and never enter DocumentLayout. O(segments). */
export function bidiSourceSpaceSlots(
  visual: LineVisualOrder,
  segmentCount: number,
): Readonly<{
  bySlot: ReadonlyMap<number, BidiSourceSpaceSlot>;
  bySource: ReadonlyMap<number, BidiSourceSpaceSlot>;
}> | undefined {
  const groups = visual.spaceDirectionOverrides;
  if (!groups?.length) return undefined;
  const groupAt = new Int32Array(segmentCount).fill(-1);
  const orders: number[][] = groups.map(() => []);
  groups.forEach((group, groupIndex) => {
    for (let index = group.startIndex; index < group.endIndex; index++) groupAt[index] = groupIndex;
  });
  for (const index of visual.order) {
    const group = groupAt[index]!;
    if (group >= 0) orders[group]!.push(index);
  }
  const bySlot = new Map<number, BidiSourceSpaceSlot>();
  const bySource = new Map<number, BidiSourceSpaceSlot>();
  groups.forEach((group, groupIndex) => {
    const order = orders[groupIndex]!;
    order.forEach((slotIndex, at) => {
      const sourceIndex = order[(at + (group.paintRtl ? order.length - 1 : 1)) % order.length]!;
      const slot = { sourceIndex, slotIndex, side: group.paintRtl ? 'left' as const : 'right' as const };
      bySlot.set(slotIndex, slot);
      bySource.set(sourceIndex, slot);
    });
  });
  return { bySlot, bySource };
}
