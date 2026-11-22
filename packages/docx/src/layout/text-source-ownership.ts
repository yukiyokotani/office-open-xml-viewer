import { sliceSemanticSlotSpans } from './text.js';
import type { TextPlacement } from './types.js';
import { visualClusters } from './visual-cluster-geometry.js';

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
  return placement.sourceRuns.map(owner => {
    while (cursor < placement.clusters.length && placement.clusters[cursor]!.range.end <= owner.range.start) cursor++;
    let end = cursor;
    while (end < placement.clusters.length && placement.clusters[end]!.range.start < owner.range.end) end++;
    const clusters = placement.clusters.slice(cursor, end);
    let from = Number.POSITIVE_INFINITY;
    let to = Number.NEGATIVE_INFINITY;
    // The owner's rectangle is where its clusters are painted (the visual
    // table); its own cluster table stays in logical pen order from zero.
    let logicalFrom = Number.POSITIVE_INFINITY;
    clusters.forEach((cluster, member) => {
      const startPt = visual ? visual.startsPt[cursor + member]! : cluster.offset.xPt;
      from = Math.min(from, startPt);
      to = Math.max(to, startPt + cluster.advancePt);
      logicalFrom = Math.min(logicalFrom, cluster.offset.xPt);
    });
    if (clusters.length === 0) from = to = logicalFrom = 0;
    // A whole glyph/vertical cell keeps its existing physical rectangle.
    const whole = owner.range.start === placement.range.start && owner.range.end === placement.range.end;
    const offset = owner.range.start - placement.range.start;
    const ownedSpaces = Math.max(0, owner.range.end - Math.max(owner.range.start, spacesStart));
    return {
      ...unreduced,
      ...(ownedSpaces > 0 && perSpacePt > 0 ? { trailingSpaceCompressionPt: ownedSpaces * perSpacePt } : {}),
      sourceRuns: undefined,
      sourceRunIndex: owner.sourceRunIndex,
      role: owner.role ?? 'content',
      dependency: owner.dependency,
      range: owner.range,
      text: placement.text.slice(offset, offset + owner.range.end - owner.range.start),
      ...(placement.semanticSlotSpans ? { semanticSlotSpans: sliceSemanticSlotSpans(
        placement.semanticSlotSpans, offset, offset + owner.range.end - owner.range.start,
      ) } : {}),
      origin: whole ? placement.origin : { ...placement.origin, xPt: (visual ? placement.bounds.xPt : placement.origin.xPt) + from },
      bounds: whole ? placement.bounds : { ...placement.bounds, xPt: placement.bounds.xPt + from, widthPt: to - from },
      ...(placement.highlightBounds ? { highlightBounds: whole ? placement.highlightBounds : {
        ...placement.highlightBounds, xPt: placement.highlightBounds.xPt + from, widthPt: to - from,
      } } : {}),
      advancePt: whole ? placement.advancePt : to - from,
      clusters: whole ? placement.clusters : clusters.map(cluster => ({ ...cluster,
        range: { start: Math.max(owner.range.start, cluster.range.start), end: Math.min(owner.range.end, cluster.range.end) },
        offset: { ...cluster.offset, xPt: cluster.offset.xPt - logicalFrom },
      })),
    };
  });
}
