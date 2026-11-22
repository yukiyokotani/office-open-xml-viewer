import type { LayoutRect } from './types.js';

/**
 * The upright block tables of one page under
 * `word-vertical-section-upright-block-table`, in section-logical page points
 * (x inline, y block).
 *
 * Placement is one forward pass: a block is placed once, when the flow reaches
 * it or the paragraph fragment before it (body-paginator.ts), against the
 * boxes already on the page, and never moves afterwards. This index is the
 * only thing a placement reads:
 * - `boxes`, every placed box of the page, whichever column placed it. A box
 *   is registered once; the tables that run on from one another share one
 *   box, their bounding rectangle;
 * - `slots`, the tables placed ahead of the flow (the tables of a run after
 *   its first, and tables placed with the paragraph before them), by
 *   top-level body index. A table with a slot takes it.
 *
 * `boxes` is bucketed along the block axis, so a registration touches a
 * bounded number of buckets and a query reads only the boxes near its block
 * range. The bucket size is a performance parameter; results do not depend
 * on it.
 */
export interface UprightBlockPageIndex {
  readonly buckets: Map<number, LayoutRect[]>;
  /** Boxes spanning more buckets than MAX_BUCKET_SPAN; every query reads them. */
  readonly wide: LayoutRect[];
  readonly slots: Map<number, UprightBlockSlot>;
  /** Text exclusions registered so far; numbers their float identities. */
  exclusionCount: number;
}

/** A table placed ahead of the flow reaching it. */
export interface UprightBlockSlot {
  readonly sourceIndex: number;
  /** The column the table was placed for; a table reached elsewhere has no slot. */
  readonly flowDomainId: string;
  /** Its box between the outer edges of its border rules. */
  readonly bounds: LayoutRect;
  /** Placed with the paragraph before it, which already wraps around it: the
   * text exclusion is registered and the placement is final. */
  readonly reserved: boolean;
}

/** What accepting upright blocks adds to the page: at most one box for the
 * index, the slots of tables placed ahead of the flow, and the text
 * exclusions of the blocks placed now. */
export interface UprightBlockRegistration {
  readonly box: LayoutRect | null;
  readonly slots: readonly UprightBlockSlot[];
  readonly exclusions: readonly Readonly<{ id: string; bounds: LayoutRect }>[];
}

const BUCKET_PT = 72;
const MAX_BUCKET_SPAN = 32;

export function emptyUprightBlockPageIndex(): UprightBlockPageIndex {
  return { buckets: new Map(), wide: [], slots: new Map(), exclusionCount: 0 };
}

/** Registers `box`, met along the block axis up to `reachPt` past its far edge. */
export function addUprightBlockBox(index: UprightBlockPageIndex, box: LayoutRect, reachPt: number): void {
  const first = Math.floor(box.yPt / BUCKET_PT);
  const last = Math.floor((box.yPt + box.heightPt + reachPt) / BUCKET_PT);
  if (!(last - first < MAX_BUCKET_SPAN)) {
    index.wide.push(box);
    return;
  }
  for (let bucket = first; bucket <= last; bucket += 1) {
    const boxes = index.buckets.get(bucket);
    if (boxes) boxes.push(box);
    else index.buckets.set(bucket, [box]);
  }
}

/**
 * The registered boxes a block spanning [`blockStartPt`, `blockEndPt`] meets:
 * those starting at or before its end whose far edge, extended by `reachPt`,
 * lies past its start. Ordered by inline start.
 */
export function uprightBlockBoxesMet(
  index: UprightBlockPageIndex,
  blockStartPt: number,
  blockEndPt: number,
  reachPt: number,
): LayoutRect[] {
  const met: LayoutRect[] = [];
  const meets = (box: LayoutRect) => box.yPt <= blockEndPt && blockStartPt < box.yPt + box.heightPt + reachPt;
  for (const box of index.wide) if (meets(box)) met.push(box);
  const first = Math.floor(blockStartPt / BUCKET_PT);
  const last = Math.floor(blockEndPt / BUCKET_PT);
  if (last - first < MAX_BUCKET_SPAN) {
    for (let bucket = first; bucket <= last; bucket += 1) {
      const boxes = index.buckets.get(bucket);
      if (!boxes) continue;
      for (const box of boxes) {
        // A box registered in several buckets is taken from the first of
        // them that this query reads.
        const firstShared = Math.max(first, Math.floor(box.yPt / BUCKET_PT));
        if (bucket === firstShared && meets(box)) met.push(box);
      }
    }
  } else {
    // A block wider than the bucketed range reads every bucket once.
    const seen = new Set<LayoutRect>();
    for (const boxes of index.buckets.values()) {
      for (const box of boxes) {
        if (seen.has(box)) continue;
        seen.add(box);
        if (meets(box)) met.push(box);
      }
    }
  }
  return met.sort((left, right) => left.xPt - right.xPt);
}
