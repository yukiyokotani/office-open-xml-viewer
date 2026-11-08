import type { SegmentLineSourceCell } from '@silurus/ooxml-core/internal/bidi-line';
import type { MeasuredLinePlanSegment, MeasuredTextPlanSegment } from './paragraph.js';
import type { TextPaintOp } from './types.js';

export interface SourceSpaceCells {
  readonly segments: readonly MeasuredLinePlanSegment[];
  readonly parentCellStarts: readonly number[];
  readonly spaceOnly: readonly boolean[];
  readonly projection: readonly SegmentLineSourceCell[];
}

interface PreparedCell extends SegmentLineSourceCell {
  readonly firstCluster: number;
  readonly endCluster: number;
  readonly fromPt: number;
  readonly toPt: number;
  readonly firstOperation: number;
  readonly endOperation: number;
}

interface PreparedSourceSpaceCells {
  readonly projection: readonly SegmentLineSourceCell[];
  readonly materialize: () => SourceSpaceCells;
}

/** A final-layout view, not new acquired text. Original indices, source ranges,
 * widths, contextual operations and authored bidi scopes remain parent-owned.
 * Atomic U+0020 clusters become independent cells BEFORE the whole-host UAX
 * pass and gap enumeration (§17.18.44). A visible joining island stays whole.
 * No substring measurement, glyph reversal or inferred neighbor direction.
 * Validate cuts before resolving the host, but allocate acquired cell geometry
 * only after that SAME resolver selects cells. Mixed-parent fallback therefore
 * pays only for the local primitive cut plan, not discarded derived geometry.
 * The plan captures immutable parents for one synchronous final-line call; it
 * is neither an acquired cache entry nor a retained/worker payload.
 *
 * Stage 1 leaves indivisible spaces, fixed/native-special geometry and an op
 * spanning multiple visible words on the old whole-line path. Their acquired
 * ink/context cannot be split safely by this representation. */
export function prepareSourceSpaceCells(parents: readonly MeasuredLinePlanSegment[]): PreparedSourceSpaceCells | undefined {
  if (!parents.some(segment => segment.kind === 'text' && segment.text.includes(' '))) return undefined;
  const projection: PreparedCell[] = [];
  const parentCellStarts: number[] = [];
  let hasSpaceCell = false;
  for (let parentIndex = 0; parentIndex < parents.length; parentIndex += 1) {
    const parent = parents[parentIndex]!;
    parentCellStarts.push(projection.length);
    if (parent.kind === 'text' && parent.nativeVisiblePrefix) {
      const prefix = parent.nativeVisiblePrefix, op = parent.basePaintOps[0];
      if (parent.nativeBaseDirection !== 'rtl' || !parent.text.endsWith(' ')
        || prefix.endUTF16 !== parent.text.replace(/ +$/u, '').length
        || prefix.request.text !== parent.text.slice(0, prefix.endUTF16)
        || prefix.request.baseDirection !== 'rtl' || prefix.measurement.nativeBaseDirection !== 'rtl'
        || prefix.request.fontRoute.fingerprint !== parent.fontRoute.fingerprint
        || prefix.request.fontSizePt !== parent.fontSizePt || prefix.request.weight !== parent.fontWeight
        || prefix.request.style !== parent.fontStyle || parent.basePaintOps.length !== 1 || !op
        || op.text !== parent.text || op.range.start !== parent.range.start || op.range.end !== parent.range.end
        || op.letterSpacingPt !== prefix.request.letterSpacingPt || op.scaleX !== 1
        || (op.scaleY !== undefined && op.scaleY !== 1) || op.glyphOrientation !== undefined
        || op.kerning !== (prefix.request.kerning === true ? 'normal' : 'none')) return undefined;
    }
    if (parent.kind === 'text' && (parent.fixedPitch || parent.fitText || parent.ruby || parent.tateChuYoko
      || parent.optionalHyphenGlyph || parent.retainedGeometry?.emphasis
      || (parent.decorations.length > 0 && !parent.retainedGeometry)
      || parent.basePaintOps.some(op => op.writingMode !== 'horizontal-tb'
        || op.sourceMapping === 'kashida' || op.naturalAdvancePt !== undefined))) return undefined;
    if (parent.kind !== 'text' || !parent.text.includes(' ')) {
      projection.push({ parentIndex, startUTF16: 0,
        endUTF16: parent.kind === 'text' ? parent.text.length : 0, spaceOnly: false,
        firstCluster: -1, endCluster: -1, fromPt: 0, toPt: 0, firstOperation: 0, endOperation: 0 });
      continue;
    }
    const ranges: Array<{ first: number; end: number; space: boolean }> = [];
    let first = 0;
    for (let index = 0; index < parent.clusters.length; index += 1) {
      const cluster = parent.clusters[index]!;
      const value = parent.text.slice(cluster.range.start - parent.range.start, cluster.range.end - parent.range.start);
      if (!value.includes(' ')) continue;
      // Do not split SPACE+mark/ZWJ or an indivisible multi-space native atom.
      if (value !== ' ') return undefined;
      hasSpaceCell = true;
      if (first < index) ranges.push({ first, end: index, space: false });
      ranges.push({ first: index, end: index + 1, space: true });
      first = index + 1;
    }
    if (first < parent.clusters.length) ranges.push({ first, end: parent.clusters.length, space: false });
    if (ranges.length === 0) return undefined;
    // Only an ASCII SPACE suffix can be removed while keeping the complete
    // visible operation. Contextual multiword/leading-space operations and
    // other whitespace (including NBSP) cannot be cut by this representation.
    if (parent.basePaintOps.some(op => op.text.replace(/ +$/u, '').includes(' '))) return undefined;
    let operationCursor = 0;
    for (let index = 0; index < ranges.length; index += 1) {
      const part = ranges[index]!;
      const start = parent.clusters[part.first]!.range.start;
      const end = parent.clusters[part.end - 1]!.range.end;
      // Acquired pen boundaries include contextual seams exactly once.
      const fromPt = index === 0 ? 0 : parent.clusters[part.first]!.offset.xPt;
      const toPt = index + 1 === ranges.length ? parent.measuredWidthPt
        : parent.clusters[ranges[index + 1]!.first]!.offset.xPt;
      if (!Number.isFinite(fromPt) || !Number.isFinite(toPt) || toPt < fromPt) return undefined;
      while (parent.basePaintOps[operationCursor]
        && parent.basePaintOps[operationCursor]!.range.end <= start) operationCursor += 1;
      let endOperation = operationCursor;
      let firstOpStart: number | undefined;
      let lastOpEnd: number | undefined;
      let cellClusterCursor = part.first;
      // All capability/coverage checks must precede the host-model decision.
      // Materializing an accepted plan is total: no later fallback can mix
      // the fine author policy with an unvalidated whole-parent paint view.
      for (; endOperation < parent.basePaintOps.length; endOperation += 1) {
        const op = parent.basePaintOps[endOperation]!;
        if (op.range.start >= end) break;
        const a = Math.max(start, op.range.start), b = Math.min(end, op.range.end);
        while (cellClusterCursor < part.end && parent.clusters[cellClusterCursor]!.range.start < a) cellClusterCursor += 1;
        const cluster = cellClusterCursor < part.end ? parent.clusters[cellClusterCursor] : undefined;
        if (!cluster || cluster.range.start !== a || b <= a) return undefined;
        firstOpStart ??= a;
        lastOpEnd = b;
      }
      if (firstOpStart !== start || lastOpEnd !== end) return undefined;
      projection.push({ parentIndex, startUTF16: start - parent.range.start,
        endUTF16: end - parent.range.start, spaceOnly: part.space,
        firstCluster: part.first, endCluster: part.end, fromPt, toPt,
        firstOperation: operationCursor, endOperation });
    }
  }
  parentCellStarts.push(projection.length);
  if (!hasSpaceCell) return undefined;
  return { projection, materialize: () => materializeSourceSpaceCells(parents, projection, parentCellStarts) };
}

function materializeSourceSpaceCells(
  parents: readonly MeasuredLinePlanSegment[], projection: readonly PreparedCell[],
  parentCellStarts: readonly number[],
): SourceSpaceCells {
  const segments: MeasuredLinePlanSegment[] = [];
  const spaceOnly: boolean[] = [];
  for (let parentIndex = 0; parentIndex < parents.length; parentIndex += 1) {
    const parent = parents[parentIndex]!;
    const firstIndex = parentCellStarts[parentIndex]!;
    const endIndex = parentCellStarts[parentIndex + 1]!;
    if (parent.kind !== 'text' || projection[firstIndex]!.firstCluster < 0) {
      segments.push(parent); spaceOnly.push(false); continue;
    }
    const { textShapeRequest: _request, textLayoutService: _service,
      nativeVisiblePrefix: prefixAuthority, ...retainedParent } = parent;
    let ownerCursor = 0;
    let slotCursor = 0;
    for (let index = firstIndex; index < endIndex; index += 1) {
      const part = projection[index]!;
      const firstCell = index === firstIndex;
      const start = parent.range.start + part.startUTF16;
      const end = parent.range.start + part.endUTF16;
      const { fromPt, toPt } = part;
      // The operation we actually paint is the trimmed prefix. Its acquired
      // native ink can differ from the full RTL Word+SPACE origin; neither
      // advance subtraction nor the parent's old ink certifies that prefix.
      const prefix = prefixAuthority && !part.spaceOnly && part.startUTF16 === 0
        && part.endUTF16 === prefixAuthority.endUTF16 ? prefixAuthority : undefined;
      const clusters: MeasuredTextPlanSegment['clusters'][number][] = [];
      for (let ci = part.firstCluster; ci < part.endCluster; ci += 1) {
        const cluster = parent.clusters[ci]!;
        // The acquired objects are readonly. A zero rebase preserves their
        // exact values and ownership; signed -0 still follows the original
        // subtraction, as do translated offsets.
        clusters.push(Object.is(fromPt, 0) ? cluster
          : { ...cluster, offset: { ...cluster.offset, xPt: cluster.offset.xPt - fromPt } });
      }
      const ops: TextPaintOp[] = [];
      let cellClusterCursor = part.firstCluster;
      for (let oi = part.firstOperation; oi < part.endOperation; oi += 1) {
        const op = parent.basePaintOps[oi]!;
        const a = Math.max(start, op.range.start), b = Math.min(end, op.range.end);
        while (parent.clusters[cellClusterCursor]!.range.start < a) cellClusterCursor += 1;
        const cluster = parent.clusters[cellClusterCursor]!;
        const { inkBounds, blockAxisInkBounds, ...facts } = op;
        ops.push({
          ...facts,
          text: op.text.slice(a - op.range.start, b - op.range.start),
          range: { start: a, end: b },
          offset: { ...op.offset, xPt: (a === op.range.start ? op.offset.xPt : cluster.offset.xPt) - fromPt },
          ...(!part.spaceOnly && (prefix?.measurement.inkBounds ?? inkBounds)
            ? { inkBounds: prefix?.measurement.inkBounds ?? inkBounds } : {}),
          ...(!part.spaceOnly && blockAxisInkBounds ? { blockAxisInkBounds: prefix ? {
            startPt: (op.glyphOffsetPt?.yPt ?? 0) - prefix.measurement.inkBounds.ascentPt,
            endPt: (op.glyphOffsetPt?.yPt ?? 0) + prefix.measurement.inkBounds.descentPt,
          } : blockAxisInkBounds } : {}),
        });
      }
      const relativeStart = start - parent.range.start, relativeEnd = end - parent.range.start;
      let sourceRuns: MeasuredTextPlanSegment['sourceRuns'];
      if (parent.sourceRuns) {
        while (parent.sourceRuns[ownerCursor] && parent.sourceRuns[ownerCursor]!.range.end <= start) ownerCursor += 1;
        const owners: NonNullable<MeasuredTextPlanSegment['sourceRuns']>[number][] = [];
        for (let oi = ownerCursor; oi < parent.sourceRuns.length; oi += 1) {
          const owner = parent.sourceRuns[oi]!;
          if (owner.range.start >= end) break;
          owners.push({ ...owner, range: { start: Math.max(start, owner.range.start), end: Math.min(end, owner.range.end) } });
        }
        sourceRuns = owners;
      }
      let semanticSlotSpans: MeasuredTextPlanSegment['semanticSlotSpans'];
      if (parent.semanticSlotSpans) {
        while (parent.semanticSlotSpans[slotCursor] && parent.semanticSlotSpans[slotCursor]!.end <= relativeStart) slotCursor += 1;
        const slots: NonNullable<MeasuredTextPlanSegment['semanticSlotSpans']>[number][] = [];
        for (let si = slotCursor; si < parent.semanticSlotSpans.length; si += 1) {
          const slot = parent.semanticSlotSpans[si]!;
          if (slot.start >= relativeEnd) break;
          slots.push({ ...slot, start: Math.max(relativeStart, slot.start) - relativeStart,
            end: Math.min(relativeEnd, slot.end) - relativeStart });
        }
        semanticSlotSpans = slots;
      }
      const cell: MeasuredTextPlanSegment = {
        ...retainedParent,
        breakBefore: firstCell ? parent.breakBefore : true,
        text: parent.text.slice(relativeStart, relativeEnd), range: { start, end },
        measuredWidthPt: toPt - fromPt,
        clusters,
        basePaintOps: ops,
        ...(parent.retainedGeometry && prefixAuthority ? {
          retainedGeometry: { ...parent.retainedGeometry, base: prefix ? {
            ascentPt: prefix.measurement.ascentPt, descentPt: prefix.measurement.descentPt,
            inkBounds: prefix.measurement.inkBounds,
          } : part.spaceOnly ? {
            ...parent.retainedGeometry.base,
            // The charged source cell paints no visible glyph. Keep its font
            // box for rules, but do not borrow the neighboring word's ink.
            inkBounds: { xMinPt: 0, xMaxPt: 0, ascentPt: 0, descentPt: 0 },
          } : parent.retainedGeometry.base },
        } : {}),
        ...(sourceRuns ? { sourceRuns } : {}),
        ...(semanticSlotSpans ? { semanticSlotSpans } : {}),
        autoSpace: firstCell ? parent.autoSpace : undefined,
        leadingSeamPairPt: firstCell ? parent.leadingSeamPairPt : undefined,
        decorationTerminalAdvancePt: parent.decorationTerminalAdvancePt === undefined ? undefined
          : Math.max(0, Math.min(toPt, parent.decorationTerminalAdvancePt) - fromPt),
      };
      segments.push(cell); spaceOnly.push(part.spaceOnly);
    }
  }
  return { segments, parentCellStarts, spaceOnly, projection };
}
