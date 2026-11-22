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

type ClusteredText = Pick<TextPlacement, 'text' | 'range' | 'clusters' | 'direction' | 'origin' | 'bounds'>;
type VisualClusters = Readonly<{ startsPt: readonly number[]; odd: readonly boolean[] }>;
const visualClustersByTable = new WeakMap<TextPlacement['clusters'], VisualClusters | null>();

/** The visual start of each retained cluster, relative to the left edge of
 * the placement's box (`bounds.xPt`), or null when the piece paints in
 * logical order from its origin. Retained cluster offsets are logical pen
 * positions; a piece with right-to-left text paints in the visual order of
 * its UAX#9 levels within the same extent, and its paint origin may stand
 * past the trailing whitespace that visual order moves to the left. Everything
 * derived from a piece's geometry (anchor positions, source ownership) reads
 * positions through this one table. Computed once per cluster table. */
export function visualClusters(placement: ClusteredText): VisualClusters | null {
  const cached = visualClustersByTable.get(placement.clusters);
  if (cached !== undefined) return cached;
  let result: VisualClusters | null = null;
  const direction = placement.direction;
  if ((direction === 'ltr' || direction === 'rtl') && placement.clusters.length > 0
    && (direction === 'rtl' || hasStrongRtl(placement.text))) {
    const levels = resolvedBidiLevels(placement.text, direction);
    const level = (index: number): number =>
      levels[placement.clusters[index]!.range.start - placement.range.start] ?? (direction === 'rtl' ? 1 : 0);
    result = {
      startsPt: bidiPlacedOffsets(placement.clusters.map((cluster, index) => ({
        xPt: cluster.offset.xPt, advancePt: cluster.advancePt, level: level(index),
      }))),
      odd: placement.clusters.map((_, index) => level(index) % 2 === 1),
    };
  }
  visualClustersByTable.set(placement.clusters, result);
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
  return leadingIsRight !== trailing ? startPt + clusters[at]!.advancePt : startPt;
}

/** Visual x of the start (left edge) of cluster `index`. */
export function visualClusterStartXPt(placement: ClusteredText, index: number): number {
  const visual = visualClusters(placement);
  return visual
    ? placement.bounds.xPt + visual.startsPt[index]!
    : placement.origin.xPt + placement.clusters[index]!.offset.xPt;
}
