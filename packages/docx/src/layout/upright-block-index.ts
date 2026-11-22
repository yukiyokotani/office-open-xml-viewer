import type { FloatRegistryEntryPt, LayoutRect, ParagraphLayout } from './types.js';

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
 * Invariants that every placement relies on:
 * - Drawings before tables. A wrapped drawing anchored by a paragraph is
 *   in the way of every table after that paragraph where the drawing lies:
 *   registered on the page (committed with its paragraph, or registered
 *   when the page's layout started), or, while the paragraph fragment that
 *   anchors it is planned with the table after it and not committed yet,
 *   handed to that plan (BodyTableAcquisitionInput.precedingDrawings). A
 *   drawing anchored after a table lies below it: when a table box is
 *   placed, the drawings it covers move below it here before any later
 *   table is placed against them.
 * - Drawings move past tables. A wrapped drawing laid out while a table
 *   box is in place never covers it: it moves along the inline axis to lie
 *   below it, or past it where the page ends (uprightDrawingMovePt), whichever of the two comes
 *   first in the body. That is a drawing anchored after the table; one
 *   anchored before a table placed ahead of the flow (a pin, or the tables
 *   planned with the paragraph fragment before them, whose plan is accepted
 *   only when the drawings of the fragment that is committed leave its box
 *   clear, body-paginator.ts); and one carried from an earlier layout of
 *   its page, from the frame it keeps. Tables never move for a drawing laid
 *   out after them and a drawing a table has moved never moves back, so the
 *   two cannot chase each other; how that holds for a drawing that follows
 *   its own text is stated at AnchorReflowMoves (paragraph.ts).
 * - What a box obstructs. A table box obstructs only the fragments after
 *   it in body order (`boxOrders`). A box of a later fragment is on the
 *   page early only as a pin, and the layout that pinned it placed the
 *   earlier fragments without it. A pin's text exclusion holds for all
 *   text of the page: that is what it is for.
 * - Termination. Each layout of a page either completes or is laid out
 *   again for a bounded reason (body-paginator.ts): on each page a fragment
 *   is pinned at most once and a table released at most once, and a layout
 *   that holds pins and fails is laid out once more without any pin, after
 *   which the page asks for none.
 *
 * `boxes` is bucketed along the block axis, `exclusions` and the page's
 * drawings along both axes, so a registration touches a bounded number of
 * buckets and a query reads only the rectangles near its range. The bucket size is a performance
 * parameter; results do not depend on it.
 */
export interface UprightBlockPageIndex {
  readonly buckets: Map<number, LayoutRect[]>;
  /** Boxes spanning more buckets than MAX_BUCKET_SPAN; every query reads them. */
  readonly wide: LayoutRect[];
  /** The body order of each box's first fragment (uprightFragmentOrder): a
   * box obstructs only the fragments after it in body order. */
  readonly boxOrders: Map<LayoutRect, number>;
  readonly slots: Map<number, UprightBlockSlot>;
  /** The owner segments pinned when the page's layout started, by segment
   * key (uprightSegmentKey): a pinned segment takes these bounds. */
  readonly segmentPins: Map<string, LayoutRect>;
  /** The table placed last (UprightBlockTail). */
  tail: UprightBlockTail | null;
  /** The boxes of the placed top-level tables with the body index of the
   * last table in each, bucketed along the block axis (addUprightTableBox). */
  readonly tableBuckets: Map<number, UprightTableBox[]>;
  /** Number of boxes in `tableBuckets`. */
  tableCount: number;
  /** The page's wrapped anchored drawings by `${block bucket}:${inline
   * bucket}`, each cell a max-heap on inline end (UprightDrawingCell); those
   * covering more than MAX_EXCLUSION_CELLS cells; and how many of the
   * page's wrap rectangles have been looked at for them
   * (syncUprightDrawings). */
  readonly drawingCells: Map<string, UprightDrawingCell>;
  readonly wideDrawings: UprightPageDrawing[];
  drawingsSeen: number;
  /** Identifies the boxes of `tableBuckets` in registration order: equal for
   * two pages (or two layouts of one page) exactly when they hold the same
   * boxes for the same tables, so a result that depends on them can be keyed
   * by it across layouts (UprightTableBoxVersions). */
  tableBoxesVersion: number;
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
  /** Body index of the table, or of the first table of a logical table. */
  readonly sourceIndex: number;
  /** Body index of its last table. */
  readonly lastSourceIndex: number;
  /** The column the table was placed for. A table the flow reaches in
   * another column has no slot, unless the slot is `anyColumn`. */
  readonly flowDomainId: string;
  /** Its box between the outer edges of its border rules. */
  readonly bounds: LayoutRect;
  /** Placed with the paragraph before it, which already wraps around it, or
   * pinned: the text exclusion is registered and the placement is final. */
  readonly reserved: boolean;
  /** The table takes the slot in whichever column of the page the flow
   * reaches it: the text before it was moved to another column around it. */
  readonly anyColumn: boolean;
  /** Registered when the page's layout started (a page pin): all text of the
   * page already flows around the table, so it asks for no further pin. */
  readonly pinned: boolean;
  /** The first table of the plan that placed it. It speaks for the tables
   * placed with it: only it asks for a page pin, which holds them all. */
  readonly first: boolean;
  /** The table does not fit the band from `bounds` and is cut between rows
   * there; the rest continues on the next page. */
  readonly splitRows: boolean;
}

/** The last upright block table placed in a column and where it ends: the
 * table that directly follows it in the body continues from there. */
export type UprightBlockTail = Readonly<{ sourceIndex: number; flowDomainId: string; bounds: LayoutRect }>;

/** A wrapped anchored drawing of the page: the body index of the paragraph
 * that anchors it, its text exclusion and its frame, in section-logical page
 * points. */
export type UprightPageDrawing = {
  readonly anchorIndex: number;
  /** Where the drawing lies, below every table placed for a table before its
   * paragraph that it would cover (moveUprightPageDrawingsBelowTableBox).
   * Replaced, never mutated, when it moves. */
  exclusion: LayoutRect;
  frame: LayoutRect;
  /** The frame as the page's wrap rectangles hold it, which the text laid
   * out so far wraps around. */
  readonly registeredFrame: LayoutRect;
};

/** What accepting upright blocks adds to the page: at most one box for the
 * index, the slots of tables placed ahead of the flow, the text exclusions
 * of the blocks placed now, and the new tail. */
export interface UprightBlockRegistration {
  readonly box: LayoutRect | null;
  /** Body order of the first fragment `box` holds (uprightFragmentOrder). */
  readonly boxOrder?: number;
  readonly slots: readonly UprightBlockSlot[];
  readonly exclusions: readonly UprightBlockExclusion[];
  readonly tail?: UprightBlockTail;
  /** A box for the history of placed tables only (addUprightTableBox), where
   * `box` and `tail` give none: an owner segment's, which a drawing anchored
   * after its table lies below. */
  readonly tableBox?: UprightTableBox;
  /** An owner segment placed for good at `bounds` (a page pin). */
  readonly segmentPin?: Readonly<{ key: string; bounds: LayoutRect }>;
}

/** Identity of one owner segment of the table at body index `sourceIndex`,
 * by the fragment-start key of its first row. */
export function uprightSegmentKey(sourceIndex: number, fragmentStartKey: string): string {
  return `${sourceIndex}|${fragmentStartKey}`;
}

const BUCKET_PT = 72;
const MAX_BUCKET_SPAN = 32;
const MAX_EXCLUSION_CELLS = 256;

export function emptyUprightBlockPageIndex(): UprightBlockPageIndex {
  return {
    buckets: new Map(),
    wide: [],
    boxOrders: new Map(),
    slots: new Map(),
    segmentPins: new Map(),
    tail: null,
    tableBuckets: new Map(),
    tableCount: 0,
    tableBoxesVersion: 0,
    drawingCells: new Map(),
    wideDrawings: [],
    drawingsSeen: 0,
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

/** Registers `box`, whose first fragment has body order `order`, met along
 * the block axis up to `reachPt` past its far edge. */
export function addUprightBlockBox(index: UprightBlockPageIndex, box: LayoutRect, reachPt: number, order: number): void {
  index.boxOrders.set(box, order);
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
 * The registered boxes a block spanning [`blockStartPt`, `blockEndPt`] whose
 * fragment has body order `order` meets: those of fragments before it in
 * body order starting at or before its end whose far edge, extended by
 * `reachPt`, lies past its start. Ordered by inline start.
 *
 * A box of a later fragment is on the page before this one only as a pin,
 * registered when the page's layout started; the layout that pinned it
 * placed this fragment before it, so it does not obstruct this one.
 */
export function uprightBlockBoxesMet(
  index: UprightBlockPageIndex,
  blockStartPt: number,
  blockEndPt: number,
  reachPt: number,
  order: number,
): LayoutRect[] {
  const met: LayoutRect[] = [];
  const meets = (box: LayoutRect) => box.yPt <= blockEndPt && blockStartPt < box.yPt + box.heightPt + reachPt
    && index.boxOrders.get(box)! < order;
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
 * Numbers the histories of placed table boxes of one layout execution. A
 * history is its predecessor extended by one box, so each gets one number
 * from one short key, and two histories have the same number exactly when
 * they hold the same boxes in the same order. Zero is the empty history.
 */
export type UprightTableBoxVersions = Map<string, number>;

/** The box of placed tables that end at body index `sourceIndex`. */
export type UprightTableBox = Readonly<{ sourceIndex: number; box: LayoutRect }>;

/** Records the box of placed tables ending at body index `sourceIndex`. */
export function addUprightTableBox(
  index: UprightBlockPageIndex,
  sourceIndex: number,
  box: LayoutRect,
  versions: UprightTableBoxVersions,
): void {
  const entry: UprightTableBox = Object.freeze({ sourceIndex, box });
  // Registered in at most MAX_BUCKET_SPAN + 1 buckets from its block start:
  // no page is wide enough for a query past that range to meet the box.
  const first = bucketOf(box.yPt);
  const last = Math.min(bucketOf(box.yPt + box.heightPt), first + MAX_BUCKET_SPAN);
  for (let bucket = first; bucket <= last; bucket += 1) {
    const entries = index.tableBuckets.get(bucket);
    if (entries) entries.push(entry);
    else index.tableBuckets.set(bucket, [entry]);
  }
  index.tableCount += 1;
  // The version of this history: the one before it extended by this box.
  const step = `${index.tableBoxesVersion}|${sourceIndex}:${box.xPt},${box.yPt},${box.widthPt},${box.heightPt}`;
  let version = versions.get(step);
  if (version === undefined) versions.set(step, version = versions.size + 1);
  index.tableBoxesVersion = version;
}

/** The page's ends along the inline and the block axis, in section-logical
 * page points: a drawing moved for a table stays before them. */
export type UprightDrawingLimits = Readonly<{ inlineEndPt: number; blockEndPt: number }>;

/**
 * `word-vertical-section-upright-block-table`: how far a wrapped drawing
 * with frame `rect` moves so that it covers no upright block table. Zero
 * when it covers none. Null when no place tried lies on the page: the
 * drawing then keeps its place, over the table, and stays visible. Library
 * policy where no control exists (the measured case is a drawing with room
 * below its table).
 *
 * Two routes are tried from the drawing's own position, in this order, and
 * the first that ends clear is taken:
 * 1. below: along the inline axis (x) to lie directly below what it covers,
 *    and, where that would pass the page's inline end, along the block axis
 *    (y) past the far edge instead;
 * 2. past: along the block axis past the far edge of what it covers, and,
 *    where that would pass the page's block end, along the inline axis
 *    instead.
 * A route ends without a place when neither move stays before the page's
 * ends.
 *
 * The tables are those placed for tables before body index
 * `beforeSourceIndex`, and `planned`. A paragraph's own acquisition passes
 * no limit (Number.POSITIVE_INFINITY): every box on the page when the
 * paragraph is laid out is in place for good, also one of a later table
 * placed ahead of the flow, and the drawing moves past it
 * (UprightBlockPageIndex, "Drawings move past tables").
 *
 * `blockers` are rectangles the drawing may not cover either: the drawings
 * it collides with (ECMA-376 §20.4.2.3, allowOverlap false). They are
 * cleared in the same search, so the place found is clear of the tables and
 * of the blockers at once.
 *
 * Each route ends: a step moves the drawing past every rectangle it
 * covers, along one axis, and it never moves back along either, so a
 * rectangle left behind is not met again and a route takes at most as many
 * steps as there are rectangles. Only the table boxes in the drawing's
 * current block range are read.
 */
export function uprightDrawingMovePt(
  index: UprightBlockPageIndex | undefined,
  beforeSourceIndex: number,
  rect: LayoutRect,
  limits: UprightDrawingLimits,
  planned?: readonly LayoutRect[],
  blockers?: readonly LayoutRect[],
): Readonly<{ xPt: number; yPt: number }> | null {
  const route = (inlineFirst: boolean): Readonly<{ xPt: number; yPt: number }> | null => {
    let xPt = rect.xPt;
    let yPt = rect.yPt;
    for (;;) {
      const endXPt = xPt + rect.widthPt;
      const endYPt = yPt + rect.heightPt;
      let nextXPt = Number.NEGATIVE_INFINITY;
      let nextYPt = Number.NEGATIVE_INFINITY;
      const meet = (box: LayoutRect) => {
        if (box.xPt < endXPt && xPt < box.xPt + box.widthPt && box.yPt < endYPt && yPt < box.yPt + box.heightPt) {
          nextXPt = Math.max(nextXPt, box.xPt + box.widthPt);
          nextYPt = Math.max(nextYPt, box.yPt + box.heightPt);
        }
      };
      if (planned) for (const box of planned) meet(box);
      if (blockers) for (const box of blockers) meet(box);
      if (index && index.tableCount > 0) {
        const first = bucketOf(yPt);
        const last = bucketOf(endYPt);
        for (let bucket = first; bucket <= last; bucket += 1) {
          const entries = index.tableBuckets.get(bucket);
          if (!entries) continue;
          for (const entry of entries) if (entry.sourceIndex < beforeSourceIndex) meet(entry.box);
        }
      }
      if (nextXPt === Number.NEGATIVE_INFINITY) return { xPt: xPt - rect.xPt, yPt: yPt - rect.yPt };
      const inlineFits = nextXPt + rect.widthPt <= limits.inlineEndPt;
      const blockFits = nextYPt + rect.heightPt <= limits.blockEndPt;
      if (inlineFits && (inlineFirst || !blockFits)) xPt = nextXPt;
      else if (blockFits) yPt = nextYPt;
      else return null;
    }
  };
  const below = route(true);
  // A drawing that covers nothing, or has a place below, takes it.
  return below ?? route(false);
}

/** The box a registration adds to the history of placed tables
 * (addUprightTableBox), or null when it adds none. */
export function uprightRegistrationTableBox(registration: UprightBlockRegistration): UprightTableBox | null {
  return registration.tableBox
    ?? (registration.tail && registration.box
      ? { sourceIndex: registration.tail.sourceIndex, box: registration.box } : null);
}

/** Whether two rectangles share an area: more than a common edge, within
 * the tolerance line bands use (1e-6 pt). */
export function uprightRectsShareArea(first: LayoutRect, second: LayoutRect): boolean {
  return Math.min(first.xPt + first.widthPt, second.xPt + second.widthPt) - Math.max(first.xPt, second.xPt) > 1e-6
    && Math.min(first.yPt + first.heightPt, second.yPt + second.heightPt) - Math.max(first.yPt, second.yPt) > 1e-6;
}

/** The extent of what a recorded drawing places or was registered with: its
 * exclusion, its frame and its registered frame. It is kept in the cells
 * this extent covers. */
function drawingExtent(drawing: UprightPageDrawing): Readonly<{
  blockStartPt: number; blockEndPt: number; inlineStartPt: number; inlineEndPt: number;
}> {
  const rects = [drawing.exclusion, drawing.frame, drawing.registeredFrame];
  let blockStartPt = Number.POSITIVE_INFINITY;
  let blockEndPt = Number.NEGATIVE_INFINITY;
  let inlineStartPt = Number.POSITIVE_INFINITY;
  let inlineEndPt = Number.NEGATIVE_INFINITY;
  for (let at = 0; at < rects.length; at += 1) {
    const rect = rects[at]!;
    blockStartPt = Math.min(blockStartPt, rect.yPt);
    blockEndPt = Math.max(blockEndPt, rect.yPt + rect.heightPt);
    inlineStartPt = Math.min(inlineStartPt, rect.xPt);
    inlineEndPt = Math.max(inlineEndPt, rect.xPt + rect.widthPt);
  }
  return { blockStartPt, blockEndPt, inlineStartPt, inlineEndPt };
}

/** Inline end of a recorded drawing's extent: each cell is a max-heap on it. */
const drawingInlineEndPt = (drawing: UprightPageDrawing) => drawingExtent(drawing).inlineEndPt;

/** The cells a recorded drawing is kept in, by block and inline bucket, or
 * null when it covers more than MAX_EXCLUSION_CELLS (kept in `wideDrawings`).
 * Along the block axis at most MAX_BUCKET_SPAN + 1 buckets from its start:
 * no page is wide enough for a query past that range to meet it. */
function drawingCells(drawing: UprightPageDrawing): Readonly<{
  firstBlock: number; lastBlock: number; firstInline: number; lastInline: number;
}> | null {
  const extent = drawingExtent(drawing);
  const firstBlock = bucketOf(extent.blockStartPt);
  const lastBlock = Math.min(bucketOf(extent.blockEndPt), firstBlock + MAX_BUCKET_SPAN);
  const firstInline = bucketOf(extent.inlineStartPt);
  const lastInline = bucketOf(extent.inlineEndPt);
  if (!((lastBlock - firstBlock + 1) * (lastInline - firstInline + 1) <= MAX_EXCLUSION_CELLS)) return null;
  return { firstBlock, lastBlock, firstInline, lastInline };
}

/** The drawings of one cell: a binary max-heap on inline end, with each
 * drawing's place in it. Insertion and removal take O(log n) swaps; the
 * drawings ending past a position are enumerated by descending from the
 * root and stopping at every one that does not, so a query reads at most
 * twice as many drawings as it returns, plus one. */
export type UprightDrawingCell = {
  readonly heap: UprightPageDrawing[];
  readonly at: Map<UprightPageDrawing, number>;
};

function placeInHeap(cell: UprightDrawingCell, drawing: UprightPageDrawing, at: number): void {
  cell.heap[at] = drawing;
  cell.at.set(drawing, at);
}

/** Moves the drawing at `from` up to its place; returns that place. */
function siftUp(cell: UprightDrawingCell, from: number): number {
  const drawing = cell.heap[from]!;
  const endPt = drawingInlineEndPt(drawing);
  let at = from;
  while (at > 0) {
    const parent = (at - 1) >>> 1;
    const above = cell.heap[parent]!;
    if (!(drawingInlineEndPt(above) < endPt)) break;
    placeInHeap(cell, above, at);
    at = parent;
  }
  placeInHeap(cell, drawing, at);
  return at;
}

function siftDown(cell: UprightDrawingCell, from: number): void {
  const drawing = cell.heap[from]!;
  const endPt = drawingInlineEndPt(drawing);
  let at = from;
  for (;;) {
    const left = 2 * at + 1;
    if (left >= cell.heap.length) break;
    const right = left + 1;
    const child = right < cell.heap.length
      && drawingInlineEndPt(cell.heap[right]!) > drawingInlineEndPt(cell.heap[left]!) ? right : left;
    const below = cell.heap[child]!;
    if (!(drawingInlineEndPt(below) > endPt)) break;
    placeInHeap(cell, below, at);
    at = child;
  }
  placeInHeap(cell, drawing, at);
}

function insertDrawing(index: UprightBlockPageIndex, drawing: UprightPageDrawing): void {
  const cells = drawingCells(drawing);
  if (!cells) {
    index.wideDrawings.push(drawing);
    return;
  }
  for (let block = cells.firstBlock; block <= cells.lastBlock; block += 1) {
    for (let inline = cells.firstInline; inline <= cells.lastInline; inline += 1) {
      const key = `${block}:${inline}`;
      let cell = index.drawingCells.get(key);
      if (!cell) index.drawingCells.set(key, cell = { heap: [], at: new Map() });
      cell.heap.push(drawing);
      siftUp(cell, cell.heap.length - 1);
    }
  }
}

function removeDrawing(index: UprightBlockPageIndex, drawing: UprightPageDrawing): void {
  const cells = drawingCells(drawing);
  if (!cells) {
    index.wideDrawings.splice(index.wideDrawings.indexOf(drawing), 1);
    return;
  }
  for (let block = cells.firstBlock; block <= cells.lastBlock; block += 1) {
    for (let inline = cells.firstInline; inline <= cells.lastInline; inline += 1) {
      const cell = index.drawingCells.get(`${block}:${inline}`)!;
      const at = cell.at.get(drawing)!;
      cell.at.delete(drawing);
      const last = cell.heap.pop()!;
      if (last === drawing) continue;
      placeInHeap(cell, last, at);
      if (siftUp(cell, at) === at) siftDown(cell, at);
    }
  }
}

/** Records a wrapped anchored drawing of the page, moved first so that it
 * covers none of the upright tables already placed for tables before its
 * paragraph (uprightDrawingMovePt). */
export function addUprightPageDrawing(
  index: UprightBlockPageIndex,
  drawing: UprightPageDrawing,
  limits: UprightDrawingLimits,
): void {
  shiftDrawing(drawing, uprightDrawingMovePt(index, drawing.anchorIndex, drawing.frame, limits));
  insertDrawing(index, drawing);
}

/** Moves a drawing's exclusion and frame by `move`. */
function shiftDrawing(drawing: UprightPageDrawing, move: Readonly<{ xPt: number; yPt: number }> | null): void {
  if (!move || (move.xPt === 0 && move.yPt === 0)) return;
  const moved = (rect: LayoutRect) => Object.freeze({ ...rect, xPt: rect.xPt + move.xPt, yPt: rect.yPt + move.yPt });
  drawing.exclusion = moved(drawing.exclusion);
  drawing.frame = moved(drawing.frame);
}

/**
 * `word-vertical-section-upright-block-table`: a wrapped drawing anchored
 * after a table does not cover it. When the box of tables ending at body
 * index `sourceIndex` is placed, the recorded drawings of paragraphs after
 * them that it covers move clear of it here (uprightDrawingMovePt), before
 * any later table is placed against them, as their paragraphs' own
 * acquisition places them. Each move takes the drawing out of its cells and
 * puts it into its new ones, in O(log n) per cell.
 */
export function moveUprightPageDrawingsBelowTableBox(
  index: UprightBlockPageIndex,
  table: UprightTableBox,
  limits: UprightDrawingLimits,
): void {
  const { box } = table;
  const moves: Array<Readonly<{ drawing: UprightPageDrawing; move: Readonly<{ xPt: number; yPt: number }> }>> = [];
  for (const drawing of uprightPageDrawingsNear(index, box.yPt, box.yPt + box.heightPt, box.xPt, box.xPt + box.widthPt)) {
    const { frame } = drawing;
    if (drawing.anchorIndex > table.sourceIndex
      && frame.xPt < box.xPt + box.widthPt && box.xPt < frame.xPt + frame.widthPt
      && frame.yPt < box.yPt + box.heightPt && box.yPt < frame.yPt + frame.heightPt) {
      const move = uprightDrawingMovePt(index, drawing.anchorIndex, frame, limits);
      if (move && (move.xPt !== 0 || move.yPt !== 0)) moves.push({ drawing, move });
    }
  }
  for (const { drawing, move } of moves) {
    removeDrawing(index, drawing);
    shiftDrawing(drawing, move);
    insertDrawing(index, drawing);
  }
}

/** The recorded drawings whose extent (exclusion, frame or registered
 * frame) can lie in the block range [`blockStartPt`, `blockEndPt`] and the
 * inline range (`inlineFromPt`, `inlineToPt`), each once. Only the cells of
 * that range are read, and in each only the drawings reaching past
 * `inlineFromPt` (UprightDrawingCell): a drawing beside the range along
 * either axis is looked at only when it shares a cell with its edge. */
export function uprightPageDrawingsNear(
  index: UprightBlockPageIndex,
  blockStartPt: number,
  blockEndPt: number,
  inlineFromPt: number,
  inlineToPt: number,
): UprightPageDrawing[] {
  const near: UprightPageDrawing[] = [];
  for (let at = 0; at < index.wideDrawings.length; at += 1) {
    const drawing = index.wideDrawings[at]!;
    const extent = drawingExtent(drawing);
    if (extent.blockStartPt <= blockEndPt && blockStartPt <= extent.blockEndPt
      && extent.inlineEndPt > inlineFromPt && extent.inlineStartPt < inlineToPt) near.push(drawing);
  }
  const firstBlock = bucketOf(blockStartPt);
  const lastBlock = bucketOf(blockEndPt);
  const firstInline = bucketOf(inlineFromPt);
  const lastInline = bucketOf(inlineToPt);
  const pending: number[] = [];
  for (let block = firstBlock; block <= lastBlock; block += 1) {
    for (let inline = firstInline; inline <= lastInline; inline += 1) {
      const cell = index.drawingCells.get(`${block}:${inline}`);
      if (!cell || cell.heap.length === 0) continue;
      pending.length = 0;
      pending.push(0);
      while (pending.length > 0) {
        const at = pending.pop()!;
        const drawing = cell.heap[at]!;
        if (!(drawingInlineEndPt(drawing) > inlineFromPt)) continue;
        if (2 * at + 1 < cell.heap.length) pending.push(2 * at + 1);
        if (2 * at + 2 < cell.heap.length) pending.push(2 * at + 2);
        // A drawing kept in several cells is taken from the first of them
        // that this query reads.
        const cells = drawingCells(drawing)!;
        if (block === Math.max(firstBlock, cells.firstBlock) && inline === Math.max(firstInline, cells.firstInline)
          && drawingExtent(drawing).inlineStartPt < inlineToPt) near.push(drawing);
      }
    }
  }
  return near;
}

/**
 * `word-vertical-section-upright-block-table`: the text exclusions of the
 * wrapped drawings that `fragment` anchors, among the float registry
 * entries its acquisition would commit. A table placed after the fragment
 * before those entries are committed (planned with the fragment, or looked
 * ahead to from a keep set that holds it) meets these instead
 * (BodyTableAcquisitionInput.precedingDrawings). Placement and look-ahead
 * both take them from here.
 */
export function uprightPrecedingDrawings(
  entries: readonly FloatRegistryEntryPt[] | undefined,
  fragment: ParagraphLayout,
): LayoutRect[] {
  return fragmentDrawingEntries(entries, fragment).map((entry) => entry.exclusionBounds);
}

/** The frames of those drawings: what has to leave the box of a plan clear
 * for the plan to be accepted (body-paginator.ts). */
export function uprightDrawingFrames(
  entries: readonly FloatRegistryEntryPt[] | undefined,
  fragment: ParagraphLayout,
): LayoutRect[] {
  return fragmentDrawingEntries(entries, fragment).map((entry) => entry.bounds);
}

function fragmentDrawingEntries(
  entries: readonly FloatRegistryEntryPt[] | undefined,
  fragment: ParagraphLayout,
): FloatRegistryEntryPt[] {
  if (!entries || entries.length === 0) return [];
  const anchored = new Set<string>();
  for (const drawing of fragment.drawings) {
    const occurrenceId = drawing.anchorLayer?.acquisitionOccurrenceId ?? drawing.anchorLayer?.occurrenceId;
    if (occurrenceId !== undefined) anchored.add(occurrenceId);
  }
  return entries.filter((entry) =>
    entry.kind === 'shape' && entry.wrap !== undefined && anchored.has(entry.occurrenceId));
}
