import { hasStrongRtl } from '@silurus/ooxml-core';
import { resolvedBidiLevels } from '../character-spacing.js';
import type { TextPlacement } from './types.js';

/** WORD_CONDENSED_SPACING_ADVANCE_FLOOR splits a segment's paint at floored
 * clusters, and Canvas resolves bidirectional order only inside one
 * `fillText`. Place the split operations as the single operation they replace
 * would have been drawn: UAX#9 levels of the painted text in the paint
 * direction (Canvas applies no higher-level overrides), then each L2 reversal
 * (highest level first, down to the lowest odd level) mirrors its contiguous
 * run of operations within the run's own span. Operations are level-uniform
 * (split at level changes). A floored cluster drawn alone at an odd level
 * keeps its glyph's trailing edge on the right-to-left pen, its drawn
 * advance behind it. Work is linear per level. */
export function bidiPlacedOffsets(
  items: readonly Readonly<{ xPt: number; advancePt: number; drawnAdvancePt?: number; level: number }>[],
): number[] {
  const xs = items.map((item) => item.xPt);
  let order = items.map((_, index) => index);
  let highest = 0;
  let lowestOdd = Number.POSITIVE_INFINITY;
  for (const item of items) {
    highest = Math.max(highest, item.level);
    if (item.level % 2 === 1) lowestOdd = Math.min(lowestOdd, item.level);
  }
  for (let level = highest; level >= lowestOdd; level -= 1) {
    const next: number[] = [];
    let cursor = 0;
    while (cursor < order.length) {
      if (items[order[cursor]!]!.level < level) {
        next.push(order[cursor]!);
        cursor += 1;
        continue;
      }
      let end = cursor;
      let startPt = Number.POSITIVE_INFINITY;
      let endPt = Number.NEGATIVE_INFINITY;
      while (end < order.length && items[order[end]!]!.level >= level) {
        const index = order[end]!;
        startPt = Math.min(startPt, xs[index]!);
        endPt = Math.max(endPt, xs[index]! + items[index]!.advancePt);
        end += 1;
      }
      for (let member = end - 1; member >= cursor; member -= 1) {
        const index = order[member]!;
        xs[index] = startPt + endPt - xs[index]! - items[index]!.advancePt;
        next.push(index);
      }
      cursor = end;
    }
    order = next;
  }
  return items.map((item, index) => item.drawnAdvancePt !== undefined && item.level % 2 === 1
    ? xs[index]! + item.advancePt - item.drawnAdvancePt
    : xs[index]!);
}

type ClusteredText = Pick<TextPlacement, 'text' | 'range' | 'clusters' | 'direction' | 'origin' | 'bounds'>
  & Partial<Pick<TextPlacement, 'advancePt' | 'tateChuYoko' | 'paintOps' | 'highlightBounds' | 'fontSizePt'
    | 'ownedTrailingSlackPt'>>
  & Readonly<{ kashidaInsertions?: readonly Readonly<{ beforeCp: number; count: number }>[] }>;
type VisualClusters = Readonly<{
  /** Visual start of each cluster from the left edge of the piece's box. */
  startsPt: readonly number[];
  /** Advance of each cluster from its pen to the next pen (kashida
   * elongation and justification included). */
  advancesPt: readonly number[];
  odd: readonly boolean[];
  /** The last cluster's advance was set to end at the piece's box: cut when
   * the tracking after the final character of a fitted run lies outside it,
   * extended when a compressed fitted run paints its final glyph past its
   * pitch. */
  lastCut: boolean;
}>;
const visualClustersByPlacement = new WeakMap<object, VisualClusters | null>();
/** Extra advance each cluster is painted with under kashida justification.
 * The piece paints one operation whose text holds the inserted tatweels
 * while the retained clusters keep their natural advances; the insertions
 * come from the plan (`kashidaInsertions`), never from the painted text,
 * where an authored tatweel is indistinguishable. An equal share of the
 * added width per inserted tatweel extends the cluster before each one. */
function kashidaExtensionsPt(placement: ClusteredText): number[] | null {
  const insertions = placement.kashidaInsertions;
  if (!insertions || placement.advancePt === undefined) return null;
  const total = insertions.reduce((sum, insertion) => sum + insertion.count, 0);
  const addedPt = placement.advancePt - placement.clusters.reduce((sum, cluster) => sum + cluster.advancePt, 0);
  if (total === 0 || !(addedPt > 0)) return null;
  // UTF-16 offset of each code point, in order.
  const starts: number[] = [];
  let offset = placement.range.start;
  for (const character of placement.text) {
    starts.push(offset);
    offset += character.length;
  }
  const extensions = placement.clusters.map(() => 0);
  let cluster = 0;
  for (const insertion of [...insertions].sort((left, right) => left.beforeCp - right.beforeCp)) {
    const at = starts[insertion.beforeCp] ?? offset;
    while (cluster < placement.clusters.length - 1 && placement.clusters[cluster]!.range.end <= at) cluster += 1;
    // Before this character: after the cluster that precedes it.
    const owner = placement.clusters[cluster]!.range.start >= at && cluster > 0 ? cluster - 1 : cluster;
    extensions[owner]! += insertion.count;
  }
  return extensions.map((count) => count * addedPt / total);
}

/** Where each retained cluster of a piece is painted, or null when the piece
 * paints in logical order from its origin at its retained offsets. Retained
 * cluster offsets are logical pen positions; a piece with right-to-left text
 * paints in the visual order of its UAX#9 levels within the same extent, its
 * paint origin may stand past the trailing whitespace that visual order
 * moves to the left, and kashida elongation widens clusters in paint only.
 * Everything derived from a piece's geometry (anchor positions, source
 * ownership) reads positions through this one table. Computed once per
 * placement. */
export function visualClusters(placement: ClusteredText): VisualClusters | null {
  const cached = visualClustersByPlacement.get(placement);
  if (cached !== undefined) return cached;
  let result: VisualClusters | null = null;
  const direction = placement.direction;
  if ((direction === 'ltr' || direction === 'rtl') && placement.clusters.length > 0
    && (direction === 'rtl' || hasStrongRtl(placement.text))) {
    const levels = resolvedBidiLevels(placement.text, direction);
    const level = (index: number): number =>
      levels[placement.clusters[index]!.range.start - placement.range.start] ?? (direction === 'rtl' ? 1 : 0);
    const extensions = kashidaExtensionsPt(placement);
    let shiftPt = 0;
    const logical = placement.clusters.map((cluster, index) => {
      const item = { xPt: cluster.offset.xPt + shiftPt, advancePt: tiledAdvancePt(placement, index) + (extensions?.[index] ?? 0),
        level: level(index) };
      shiftPt += extensions?.[index] ?? 0;
      return item;
    });
    // The piece's box ends its last cluster: tracking after the final
    // character of a widened fitted run lies outside the fitted width, and
    // the final glyph of a compressed one paints past its pitch to the box
    // edge.
    const last = logical.at(-1)!;
    const overflowPt = last.xPt + last.advancePt - logical[0]!.xPt - placement.bounds.widthPt;
    const fitted = Math.abs(overflowPt) > 1e-9 && last.advancePt - overflowPt >= 0;
    if (fitted) last.advancePt -= overflowPt;
    else if (overflowPt > 1e-9) last.advancePt = 0;
    result = {
      lastCut: Math.abs(overflowPt) > 1e-9,
      startsPt: bidiPlacedOffsets(logical),
      advancesPt: logical.map((item) => item.advancePt),
      odd: logical.map((item) => item.level % 2 === 1),
    };
  }
  visualClustersByPlacement.set(placement, result);
  return result;
}

/** Visual x of the logical boundary before cluster `index` (after the last
 * cluster when `index` is the cluster count): the leading edge of that
 * cluster, which is its right edge in a right-to-left run. */
export function visualBoundaryXPt(placement: ClusteredText, index: number): number {
  const clusters = placement.clusters;
  if (clusters.length === 0) return placement.bounds.xPt;
  const visual = visualClusters(placement);
  const at = Math.min(index, clusters.length - 1);
  const startPt = visual ? placement.bounds.xPt + visual.startsPt[at]! : placement.origin.xPt + clusters[at]!.offset.xPt;
  const leadingIsRight = visual?.odd[at] === true;
  const trailing = index >= clusters.length;
  if (leadingIsRight === trailing) return startPt;
  const endPt = startPt + (visual?.advancesPt[at] ?? clusters[at]!.advancePt);
  // A piece in logical order ends at its box as well, as its final cluster
  // box does (clusterBoxes): cut where its trailing tracking or compressed
  // space lies outside, extended where a compressed fitted run's final glyph
  // paints past its pitch.
  return visual || !trailing ? endPt : Math.max(startPt, placement.bounds.xPt + placement.bounds.widthPt);
}

/** Visual x of the start (left edge) of cluster `index`. */
export function visualClusterStartXPt(placement: ClusteredText, index: number): number {
  const visual = visualClusters(placement);
  return visual
    ? placement.bounds.xPt + visual.startsPt[index]!
    : placement.origin.xPt + placement.clusters[index]!.offset.xPt;
}

/** `clusters[from, to)` (contiguous in source order) as maximal groups that
 * are also contiguous where they are painted: `[first, end)` index pairs. A
 * mixed-direction span paints as several disjoint intervals. */
export function visuallyContiguousGroups(placement: ClusteredText, from: number, to: number): [number, number][] {
  const visual = visualClusters(placement);
  if (!visual || to - from <= 1) return to > from ? [[from, to]] : [];
  const groups: [number, number][] = [];
  let first = from;
  for (let index = from + 1; index <= to; index += 1) {
    const adjacent = index < to && (
      Math.abs(visual.startsPt[index - 1]! + visual.advancesPt[index - 1]! - visual.startsPt[index]!) < 1e-6
      || Math.abs(visual.startsPt[index]! + visual.advancesPt[index]! - visual.startsPt[index - 1]!) < 1e-6);
    if (!adjacent) {
      groups.push([first, index]);
      first = index;
    }
  }
  return groups;
}

/** Where one cluster is painted, for readers that address text by offset.
 * A cluster has two intervals along the run, both from the listing's
 * origin: its advance box, from its pen to the next pen (its tracking
 * included), and its glyph box, where its paint operation draws its glyph
 * (paintedGlyphPt). Readers that ask what is painted (find highlights,
 * comment tints, clip visibility) use the glyph box; readers that place the
 * caret or size a selection use the advance box. */
export interface ClusterBox {
  /** UTF-16 offsets of the cluster in the text the boxes are listed for. */
  readonly start: number;
  readonly end: number;
  /** Left edge and width of the advance box. */
  readonly xPt: number;
  readonly widthPt: number;
  /** Left edge and width of the glyph box. */
  readonly glyphXPt: number;
  readonly glyphWidthPt: number;
  /** The cluster reads right to left: it starts at its right edge, and its
   * tracking lies on its left. */
  readonly rtl: boolean;
  /** The run's final cluster, set to end at the run's box: the tracking
   * after it is already outside its width (a fitted run). */
  readonly trackingExcluded?: true;
  /** A tate-chu-yoko cluster stands upright across the run: `xPt`/`widthPt`
   * (and its glyph box along the run) span the run's cell along it; these
   * place its advance box across the run, from the top of the listing's
   * box. */
  readonly crossPt?: number;
  readonly crossWidthPt?: number;
  /** Where the glyph is painted across the run, from the top of the
   * listing's box, when that is not the run's font box: a tate-chu-yoko
   * character's own interval, or the font box of a vertical-text cell where
   * its operation draws it (centred on the column, or on its shifted
   * baseline). */
  readonly glyphCrossPt?: number;
  readonly glyphCrossWidthPt?: number;
}

/** The painted boxes of `clusters[first, last)`, with text offsets from
 * `textStart` (clipped to `textEnd`) and positions from `originXPt`. */
export function clusterBoxes(
  placement: ClusteredText,
  first: number,
  last: number,
  textStart: number,
  textEnd: number,
  originXPt: number,
): ClusterBox[] {
  const visual = visualClusters(placement);
  const boxes: ClusterBox[] = [];
  // §17.3.2.10 tate-chu-yoko: the characters stand upright side by side
  // across the run, centred on the paint origin, each over the run's cell.
  const upright = tateChuYokoRow(placement);
  const slackOwner = (placement.ownedTrailingSlackPt ?? 0) > 0 ? rightmostCluster(placement, visual) : -1;
  for (let index = first; index < last; index += 1) {
    const cluster = placement.clusters[index]!;
    const operation = operationAt(placement, cluster.range.start);
    const alone = operation && operation.range.start === cluster.range.start
      && operation.range.end === cluster.range.end ? operation : undefined;
    if (upright) {
      const fromPt = cluster.offset.xPt - upright.originPt;
      // The row reads toward the run's top edge (it is turned back upright):
      // a character's glyph lies at the far end of its advance.
      const crossPt = upright.endPt - fromPt - cluster.advancePt - placement.bounds.yPt;
      const naturalPt = Math.max(0, alone?.naturalAdvancePt ?? cluster.glyphAdvancePt ?? cluster.advancePt);
      boxes.push({
        start: Math.max(cluster.range.start, textStart) - textStart,
        end: Math.min(cluster.range.end, textEnd) - textStart,
        xPt: placement.bounds.xPt - originXPt,
        widthPt: placement.bounds.widthPt,
        glyphXPt: (upright.glyph?.startPt ?? placement.bounds.xPt) - originXPt,
        glyphWidthPt: upright.glyph?.widthPt ?? placement.bounds.widthPt,
        rtl: false,
        crossPt,
        crossWidthPt: cluster.advancePt,
        glyphCrossPt: crossPt + cluster.advancePt - naturalPt,
        glyphCrossWidthPt: naturalPt,
      });
      continue;
    }
    const startPt = visualClusterStartXPt(placement, index);
    const final = index === placement.clusters.length - 1;
    const ownAdvancePt = visual ? visual.advancesPt[index]!
      : final
        ? Math.max(0, placement.bounds.xPt + placement.bounds.widthPt - startPt)
        : tiledAdvancePt(placement, index);
    // The justification the piece owns after itself lies between its last
    // painted cluster and the next piece's pen: that cluster's advance.
    const advancePt = index === slackOwner ? ownAdvancePt + (placement.ownedTrailingSlackPt ?? 0) : ownAdvancePt;
    const odd = visual?.odd[index] === true;
    const cut = final && (visual ? visual.lastCut : Math.abs(ownAdvancePt - cluster.advancePt) > 1e-9);
    const xPt = startPt - originXPt;
    const glyph = paintedGlyphPt(placement, cluster, alone, startPt, advancePt, odd);
    const across = operation ? paintedGlyphCrossPt(placement, operation) : null;
    boxes.push({
      start: Math.max(cluster.range.start, textStart) - textStart,
      end: Math.min(cluster.range.end, textEnd) - textStart,
      xPt,
      widthPt: advancePt,
      glyphXPt: glyph.startPt - originXPt,
      glyphWidthPt: glyph.widthPt,
      rtl: odd,
      ...(cut ? { trackingExcluded: true as const } : {}),
      ...(across ? { glyphCrossPt: across.startPt, glyphCrossWidthPt: across.widthPt } : {}),
    });
  }
  return boxes;
}

/** Where paint draws a vertical-text cell's glyph across the run, from the
 * top of the placement's box, when that is not its font box: the font box's
 * extent where the operation centres the glyph (an upright or rotated cell,
 * drawn on its middle baseline at its origin and glyph offset), or moved
 * with the baseline by the glyph offset of a sideways cell (the font box
 * already follows the operation's own baseline offset). Null for an
 * operation of horizontal text, which is drawn in the font box. */
function paintedGlyphCrossPt(
  placement: ClusteredText,
  operation: NonNullable<ClusteredText['paintOps']>[number],
): Readonly<{ startPt: number; widthPt: number }> | null {
  if (operation.glyphOrientation === undefined) return null;
  const font = placement.highlightBounds ?? placement.bounds;
  let startPt: number;
  if (operation.glyphOrientation === 'sideways') {
    startPt = font.yPt + (operation.glyphOffsetPt?.yPt ?? 0);
  } else {
    // Turned back by a quarter, an upright glyph's inline offset points
    // against the block axis.
    const centrePt = placement.origin.yPt + operation.offset.yPt + (operation.glyphOrientation === 'upright'
      ? -(operation.glyphOffsetPt?.xPt ?? 0) : operation.glyphOffsetPt?.yPt ?? 0);
    startPt = centrePt - font.heightPt / 2;
  }
  if (Math.abs(startPt - font.yPt) < 1e-9) return null;
  return { startPt: startPt - placement.bounds.yPt, widthPt: font.heightPt };
}

/** The cluster whose advance ends at the piece's right edge: the last one
 * in source order, or in a piece painted in visual order the one placed
 * furthest right. */
function rightmostCluster(placement: ClusteredText, visual: VisualClusters | null): number {
  if (!visual) return placement.clusters.length - 1;
  let found = -1;
  let endPt = Number.NEGATIVE_INFINITY;
  visual.startsPt.forEach((startPt, index) => {
    if (startPt + visual.advancesPt[index]! > endPt + 1e-9) {
      endPt = startPt + visual.advancesPt[index]!;
      found = index;
    }
  });
  return found;
}

/** Where paint draws one cluster's glyph along the run, in the placement's
 * coordinates, read from the operation that draws it:
 *
 * - an operation that draws this cluster alone, centred on its origin (an
 *   upright or rotated cell of vertical text): its drawn advance, centred on
 *   the operation's origin and glyph offset, wherever the cell's pen and
 *   tracking lie;
 * - an operation that draws this cluster alone from its own origin with a
 *   retained drawn advance (a floored condensed cluster): that advance, the
 *   width of the glyph shaped in isolation, from the operation's origin;
 * - any other operation (several clusters shaped together, or one cluster on
 *   its pen): the cluster's retained glyph advance (its contextual advance
 *   in that operation) from its pen, the left edge, or the right edge of a
 *   right-to-left cluster, whose tracking lies on its left. Without a
 *   retained glyph advance the glyph fills its advance. */
function paintedGlyphPt(
  placement: ClusteredText,
  cluster: ClusteredText['clusters'][number],
  alone: NonNullable<ClusteredText['paintOps']>[number] | undefined,
  startPt: number,
  advancePt: number,
  odd: boolean,
): Readonly<{ startPt: number; widthPt: number }> {
  const originPt = placement.origin.xPt;
  if (alone && (alone.glyphOrientation === 'upright' || alone.glyphOrientation === 'rotate')) {
    const widthPt = Math.max(0, alone.naturalAdvancePt ?? cluster.glyphAdvancePt ?? advancePt);
    // An upright glyph is drawn turned back by a quarter: its glyph-local
    // block offset lies along the run.
    const offsetPt = alone.glyphOrientation === 'upright' ? alone.glyphOffsetPt?.yPt : alone.glyphOffsetPt?.xPt;
    return { startPt: originPt + alone.offset.xPt + (offsetPt ?? 0) - widthPt / 2, widthPt };
  }
  if (alone?.naturalAdvancePt !== undefined) {
    return {
      startPt: originPt + alone.offset.xPt + (alone.glyphOffsetPt?.xPt ?? 0),
      widthPt: Math.max(0, alone.naturalAdvancePt),
    };
  }
  const widthPt = Math.max(0, cluster.glyphAdvancePt ?? advancePt);
  return { startPt: odd ? startPt + advancePt - widthPt : startPt, widthPt };
}

/** Cluster `index`'s advance from its pen to the next cluster's pen, so
 * what is added between glyphs (justification, distribution) belongs to the
 * advance before it; the last cluster keeps its own advance. */
function tiledAdvancePt(placement: ClusteredText, index: number): number {
  const clusters = placement.clusters;
  const cluster = clusters[index]!;
  const next = clusters[index + 1];
  return next ? Math.max(cluster.advancePt, next.offset.xPt - cluster.offset.xPt, 0) : cluster.advancePt;
}

/** A tate-chu-yoko run paints its characters upright in one row across the
 * run, at their natural advances and centred on its paint origin (the
 * operation's text is centred there): the row's far end across the run, the
 * natural origin its advances count from, and where the row's glyphs stand
 * along the run (the em of the run's font as the operation scales it,
 * centred on the operation's origin on the middle baseline; absent without
 * the font size, when the row spans the run's cell). Null for any other
 * run. */
function tateChuYokoRow(placement: ClusteredText): Readonly<{
  endPt: number;
  originPt: number;
  glyph?: Readonly<{ startPt: number; widthPt: number }>;
}> | null {
  if (!placement.tateChuYoko || placement.clusters.length === 0) return null;
  let start = Number.POSITIVE_INFINITY;
  let end = Number.NEGATIVE_INFINITY;
  for (const cluster of placement.clusters) {
    start = Math.min(start, cluster.offset.xPt);
    end = Math.max(end, cluster.offset.xPt + cluster.advancePt);
  }
  const operation = placement.paintOps?.[0];
  const centrePt = placement.origin.yPt + (operation?.offset.yPt ?? 0);
  const row = { endPt: centrePt + (end - start) / 2, originPt: start };
  if (!operation || placement.fontSizePt === undefined) return row;
  // Turned back by a quarter, the glyph-local block axis lies along the run,
  // scaled as paint scales it.
  const scale = operation.scaleX !== 1 || operation.scaleY !== undefined
    ? operation.writingMode === 'vertical-rl' ? operation.scaleX : operation.scaleY ?? 1 : 1;
  const widthPt = placement.fontSizePt * scale;
  return {
    ...row,
    glyph: {
      startPt: placement.origin.xPt + operation.offset.xPt + (operation.glyphOffsetPt?.yPt ?? 0) * scale - widthPt / 2,
      widthPt,
    },
  };
}

/** The paint operation that draws the character at source offset `offset`
 * of a piece; undefined without operations. The operations are ordered once
 * per piece, however many owners list its boxes. */
const orderedOperations = new WeakMap<object, readonly NonNullable<ClusteredText['paintOps']>[number][]>();
function operationAt(
  placement: ClusteredText,
  offset: number,
): NonNullable<ClusteredText['paintOps']>[number] | undefined {
  const operations = placement.paintOps;
  if (!operations) return undefined;
  let ordered = orderedOperations.get(operations);
  if (!ordered) {
    ordered = [...operations].sort((left, right) => left.range.start - right.range.start);
    orderedOperations.set(operations, ordered);
  }
  let low = 0;
  let high = ordered.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (ordered[middle]!.range.end <= offset) low = middle + 1;
    else high = middle;
  }
  const found = ordered[low];
  return found && found.range.start <= offset ? found : undefined;
}
