import { sliceSemanticSlotSpans } from './text.js';
import type { TextPlacement } from './types.js';
import { visualClusters, visuallyContiguousGroups } from './visual-cluster-geometry.js';

/** Project original ownership from retained clusters; never reshape or repaint
 * source fragments. The full sequence remains the only paint authority. This
 * projection serves source-addressed overlays and destination-field convergence.
 * A seam inside a grapheme shares its retained cluster rectangle between owners;
 * detached marks must never acquire independent measured geometry. */
export type SourceOwnedTextGeometry = Omit<TextPlacement, 'paintOps'> & Readonly<{ letterSpacingPt: number }>;

export function sourceOwnedTextPlacements(placement: TextPlacement): readonly SourceOwnedTextGeometry[] {
  const { paintOps, trailingSpaceCompressionPt, ...geometry } = placement;
  // A floored single-cluster operation (WORD_CONDENSED_SPACING_ADVANCE_FLOOR)
  // paints without pitch; the run's tracking is that of its other operations.
  const tracked = paintOps.find((operation) => operation.naturalAdvancePt === undefined) ?? paintOps[0];
  const unreduced = { ...geometry, letterSpacingPt: tracked?.letterSpacingPt ?? 0 };
  const projected = trailingSpaceCompressionPt === undefined
    ? unreduced : { ...unreduced, trailingSpaceCompressionPt };
  if (!placement.sourceRuns) return [projected];
  // The retained reduction is an equal share of each trailing U+0020 (see
  // textPlanSegment); an owner reports only the share inside its own advance.
  const spacesStart = trailingSpaceCompressionPt === undefined
    ? placement.range.end
    : placement.range.start + placement.text.replace(/ +$/u, '').length;
  const perSpacePt = (trailingSpaceCompressionPt ?? 0) / Math.max(1, placement.range.end - spacesStart);
  const visual = visualClusters(placement);
  let cursor = 0;
  return placement.sourceRuns.flatMap(owner => {
    while (cursor < placement.clusters.length && placement.clusters[cursor]!.range.end <= owner.range.start) cursor++;
    let end = cursor;
    while (end < placement.clusters.length && placement.clusters[end]!.range.start < owner.range.end) end++;
    // One geometry per interval the owner's clusters are painted in: a
    // mixed-direction owner paints as several disjoint intervals, and one
    // box over them would cover other owners' text.
    const groups = visuallyContiguousGroups(placement, cursor, end);
    return (groups.length > 0 ? groups : [[cursor, end] as [number, number]]).map(([first, last]) => {
      const clusters = placement.clusters.slice(first, last);
      const range = groups.length <= 1 ? owner.range : {
        start: Math.max(owner.range.start, clusters[0]!.range.start),
        end: Math.min(owner.range.end, clusters.at(-1)!.range.end),
      };
      let from = Number.POSITIVE_INFINITY;
      let to = Number.NEGATIVE_INFINITY;
      // The rectangle is where the clusters are painted (the visual table);
      // the part's own cluster table stays in logical pen order from zero.
      let logicalFrom = Number.POSITIVE_INFINITY;
      clusters.forEach((cluster, member) => {
        const startPt = visual ? visual.startsPt[first + member]! : cluster.offset.xPt;
        from = Math.min(from, startPt);
        to = Math.max(to, startPt + (visual ? visual.advancesPt[first + member]! : cluster.advancePt));
        logicalFrom = Math.min(logicalFrom, cluster.offset.xPt);
      });
      if (clusters.length === 0) from = to = logicalFrom = 0;
      // A whole glyph/vertical cell keeps its existing physical rectangle.
      const whole = range.start === placement.range.start && range.end === placement.range.end;
      const offset = range.start - placement.range.start;
      const ownedSpaces = Math.max(0, range.end - Math.max(range.start, spacesStart));
      return {
        ...unreduced,
        ...(ownedSpaces > 0 && perSpacePt > 0 ? { trailingSpaceCompressionPt: ownedSpaces * perSpacePt } : {}),
        sourceRuns: undefined,
        // A part's text and clusters are its own; the piece's plan is not.
        kashidaInsertions: undefined,
        sourceRunIndex: owner.sourceRunIndex,
        role: owner.role ?? 'content',
        dependency: owner.dependency,
        range,
        text: placement.text.slice(offset, offset + range.end - range.start),
        ...(placement.semanticSlotSpans ? { semanticSlotSpans: sliceSemanticSlotSpans(
          placement.semanticSlotSpans, offset, offset + range.end - range.start,
        ) } : {}),
        origin: whole ? placement.origin : { ...placement.origin, xPt: (visual ? placement.bounds.xPt : placement.origin.xPt) + from },
        bounds: whole ? placement.bounds : { ...placement.bounds, xPt: placement.bounds.xPt + from, widthPt: to - from },
        ...(placement.highlightBounds ? { highlightBounds: whole ? placement.highlightBounds : {
          ...placement.highlightBounds, xPt: placement.highlightBounds.xPt + from, widthPt: to - from,
        } } : {}),
        advancePt: whole ? placement.advancePt : to - from,
        clusters: whole ? placement.clusters : clusters.map(cluster => ({ ...cluster,
          range: { start: Math.max(range.start, cluster.range.start), end: Math.min(range.end, cluster.range.end) },
          offset: { ...cluster.offset, xPt: cluster.offset.xPt - logicalFrom },
        })),
      };
    });
  });
}
