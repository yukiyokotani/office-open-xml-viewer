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
  /** The table placed last (UprightBlockTail). */
  tail: UprightBlockTail | null;
  /** The boxes of the placed top-level tables with the body index of the
   * last table in each, in registration order. */
  readonly tables: Array<Readonly<{ sourceIndex: number; box: LayoutRect }>>;
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
  /** The column the table was placed for. A table the flow reaches in
   * another column has no slot, unless the slot is `pinned`. */
  readonly flowDomainId: string;
  /** Its box between the outer edges of its border rules. */
  readonly bounds: LayoutRect;
  /** Placed with the paragraph before it, which already wraps around it, or
   * pinned: the text exclusion is registered and the placement is final. */
  readonly reserved: boolean;
  /** Registered when the page's layout started (a page pin): the table takes
   * the slot in whichever column of the page the flow reaches it. */
  readonly pinned: boolean;
  /** The table does not fit the band from `bounds` and is cut between rows
   * there; the rest continues on the next page. */
  readonly splitRows: boolean;
}

/** The last upright block table placed in a column and where it ends: the
 * table that directly follows it in the body continues from there. */
export type UprightBlockTail = Readonly<{ sourceIndex: number; flowDomainId: string; bounds: LayoutRect }>;

/** What accepting upright blocks adds to the page: at most one box for the
 * index, the slots of tables placed ahead of the flow, the text exclusions
 * of the blocks placed now, and the new tail. */
export interface UprightBlockRegistration {
  readonly box: LayoutRect | null;
  readonly slots: readonly UprightBlockSlot[];
  readonly exclusions: readonly UprightBlockExclusion[];
  readonly tail?: UprightBlockTail;
}

const BUCKET_PT = 72;
const MAX_BUCKET_SPAN = 32;
const MAX_EXCLUSION_CELLS = 256;

export function emptyUprightBlockPageIndex(): UprightBlockPageIndex {
  return {
    buckets: new Map(),
    wide: [],
    slots: new Map(),
    tail: null,
    tables: [],
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

/**
 * `word-vertical-section-upright-block-table`: the boxes of the upright block
 * tables placed so far for tables before body index `beforeSourceIndex`: the
 * ones a wrapped drawing anchored by the paragraph at that index is moved
 * below.
 */
export function uprightTableBoxesBefore(
  index: UprightBlockPageIndex | undefined,
  beforeSourceIndex: number,
): readonly LayoutRect[] {
  if (!index || index.tables.length === 0) return NO_BOXES;
  const boxes: LayoutRect[] = [];
  for (const table of index.tables) if (table.sourceIndex < beforeSourceIndex) boxes.push(table.box);
  return boxes.length === 0 ? NO_BOXES : boxes;
}

const NO_BOXES: readonly LayoutRect[] = Object.freeze([]);

/**
 * How far along the inline axis (x) a drawing with frame `rect` moves to lie
 * below every one of `boxes` it would cover: past the inline end of each box
 * it overlaps, taken in inline order so that a box reached by the move is
 * cleared too. Zero when it overlaps none.
 */
export function inlineShiftPastBoxesPt(rect: LayoutRect, boxes: readonly LayoutRect[]): number {
  if (boxes.length === 0) return 0;
  let startPt = rect.xPt;
  for (const box of [...boxes].sort((left, right) => left.xPt - right.xPt)) {
    if (box.yPt < rect.yPt + rect.heightPt && rect.yPt < box.yPt + box.heightPt
      && box.xPt < startPt + rect.widthPt && startPt < box.xPt + box.widthPt) {
      startPt = box.xPt + box.widthPt;
    }
  }
  return startPt - rect.xPt;
}
