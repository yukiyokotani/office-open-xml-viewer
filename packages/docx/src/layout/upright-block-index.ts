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
 *   top-level body index. A table with a slot takes it;
 * - `exclusions`, the text exclusions of the placed blocks. A paragraph is
 *   laid out against those that can meet its lines only
 *   (uprightBlockExclusionsAhead), not against every block of the page.
 *
 * `boxes` is bucketed along the block axis and `exclusions` along both axes,
 * so a registration touches a bounded number of buckets and a query reads
 * only the rectangles near its range. The bucket size is a performance
 * parameter; results do not depend on it.
 */
export interface UprightBlockPageIndex {
  readonly buckets: Map<number, LayoutRect[]>;
  /** Boxes spanning more buckets than MAX_BUCKET_SPAN; every query reads them. */
  readonly wide: LayoutRect[];
  readonly slots: Map<number, UprightBlockSlot>;
  /** Text exclusions by `${block bucket}:${inline bucket}`. */
  readonly exclusionCells: Map<string, UprightBlockExclusionEntry[]>;
  /** Exclusions covering more cells than MAX_EXCLUSION_CELLS; every query reads them. */
  readonly wideExclusions: UprightBlockExclusionEntry[];
  /** Highest block bucket holding an exclusion. */
  lastExclusionBucket: number;
  /** Text exclusions registered so far; numbers their float identities. */
  exclusionCount: number;
}

/** A text exclusion of an upright block: its section-logical rectangle (the
 * block's box extended by the text clearance past its far edge) under a
 * page-unique identity. */
export type UprightBlockExclusion = Readonly<{ id: string; bounds: LayoutRect }>;

/** A registered exclusion with its registration ordinal on the page. */
export type UprightBlockExclusionEntry = Readonly<{ exclusion: UprightBlockExclusion; ordinal: number }>;

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
  readonly exclusions: readonly UprightBlockExclusion[];
}

const BUCKET_PT = 72;
const MAX_BUCKET_SPAN = 32;
const MAX_EXCLUSION_CELLS = 256;

export function emptyUprightBlockPageIndex(): UprightBlockPageIndex {
  return {
    buckets: new Map(),
    wide: [],
    slots: new Map(),
    exclusionCells: new Map(),
    wideExclusions: [],
    lastExclusionBucket: Number.NEGATIVE_INFINITY,
    exclusionCount: 0,
  };
}

const bucketOf = (pt: number) => Math.floor(pt / BUCKET_PT);

/** Registers a text exclusion and returns its entry. */
export function addUprightBlockExclusion(
  index: UprightBlockPageIndex,
  exclusion: UprightBlockExclusion,
): UprightBlockExclusionEntry {
  const entry = Object.freeze({ exclusion, ordinal: index.exclusionCount });
  index.exclusionCount += 1;
  const { xPt, yPt, widthPt, heightPt } = exclusion.bounds;
  const firstBlock = bucketOf(yPt);
  const lastBlock = bucketOf(yPt + heightPt);
  const firstInline = bucketOf(xPt);
  const lastInline = bucketOf(xPt + widthPt);
  if (!((lastBlock - firstBlock + 1) * (lastInline - firstInline + 1) <= MAX_EXCLUSION_CELLS)) {
    index.wideExclusions.push(entry);
    return entry;
  }
  for (let block = firstBlock; block <= lastBlock; block += 1) {
    for (let inline = firstInline; inline <= lastInline; inline += 1) {
      const key = `${block}:${inline}`;
      const cell = index.exclusionCells.get(key);
      if (cell) cell.push(entry);
      else index.exclusionCells.set(key, [entry]);
    }
  }
  index.lastExclusionBucket = Math.max(index.lastExclusionBucket, lastBlock);
  return entry;
}

/**
 * The registered exclusions that lines laid out from block position
 * `blockFromPt` on, in the inline band [`inlineStartPt`, `inlineEndPt`], can
 * meet: those ending past `blockFromPt` whose inline extent touches the band.
 * A line never lies before the block position its paragraph starts at, so an
 * exclusion that ends at or before it is out of reach, as is one beside the
 * band (a table of another column that does not reach into this one). In
 * registration order.
 */
export function uprightBlockExclusionsAhead(
  index: UprightBlockPageIndex,
  blockFromPt: number,
  inlineStartPt: number,
  inlineEndPt: number,
): UprightBlockExclusionEntry[] {
  const ahead: UprightBlockExclusionEntry[] = [];
  const reaches = ({ exclusion: { bounds } }: UprightBlockExclusionEntry) =>
    bounds.yPt + bounds.heightPt > blockFromPt
    && bounds.xPt <= inlineEndPt && bounds.xPt + bounds.widthPt >= inlineStartPt;
  for (const entry of index.wideExclusions) if (reaches(entry)) ahead.push(entry);
  const firstBlock = bucketOf(blockFromPt);
  const firstInline = bucketOf(inlineStartPt);
  const lastInline = bucketOf(inlineEndPt);
  for (let block = firstBlock; block <= index.lastExclusionBucket; block += 1) {
    for (let inline = firstInline; inline <= lastInline; inline += 1) {
      const cell = index.exclusionCells.get(`${block}:${inline}`);
      if (!cell) continue;
      for (const entry of cell) {
        // An exclusion registered in several cells is taken from the first
        // of them that this query reads.
        const { bounds } = entry.exclusion;
        if (block === Math.max(firstBlock, bucketOf(bounds.yPt))
          && inline === Math.max(firstInline, bucketOf(bounds.xPt))
          && reaches(entry)) ahead.push(entry);
      }
    }
  }
  return ahead.sort((left, right) => left.ordinal - right.ordinal);
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
