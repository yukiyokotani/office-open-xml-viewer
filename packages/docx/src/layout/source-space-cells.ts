import type { SegmentLineSourceCell } from '@silurus/ooxml-core/internal/bidi-line';
import type { MeasuredLinePlanSegment, MeasuredTextPlanSegment } from './paragraph.js';
import type { TextPaintOp } from './types.js';

export interface SourceSpaceCells {
  readonly segments: readonly MeasuredLinePlanSegment[];
  readonly parentCellStarts: readonly number[];
  readonly spaceOnly: readonly boolean[];
  readonly projection: readonly SegmentLineSourceCell[];
}

/** A final-layout view, not new acquired text. Original indices, source ranges,
 * widths, contextual operations and authored bidi scopes remain parent-owned.
 * Atomic U+0020 clusters become independent cells BEFORE the whole-host UAX
 * pass and gap enumeration (§17.18.44). A visible joining island stays whole.
 * No substring measurement, glyph reversal or inferred neighbor direction.
 *
 * Stage 1 leaves indivisible spaces, fixed/native-special geometry and an op
 * spanning multiple visible words on the old whole-line path. Their acquired
 * ink/context cannot be split safely by this representation. */
export function sourceSpaceCells(parents: readonly MeasuredLinePlanSegment[]): SourceSpaceCells | undefined {
  if (!parents.some(segment => segment.kind === 'text' && segment.text.includes(' '))) return undefined;
  const segments: MeasuredLinePlanSegment[] = [];
  const parentCellStarts: number[] = [];
  const spaceOnly: boolean[] = [];
  let hasSpaceCell = false;
  for (const parent of parents) {
    parentCellStarts.push(segments.length);
    if (parent.kind === 'text' && (parent.fixedPitch || parent.fitText || parent.ruby || parent.tateChuYoko
      || parent.optionalHyphenGlyph || parent.retainedGeometry?.emphasis
      || (parent.decorations.length > 0 && !parent.retainedGeometry)
      || parent.basePaintOps.some(op => op.writingMode !== 'horizontal-tb'
        || op.sourceMapping === 'kashida' || op.naturalAdvancePt !== undefined))) return undefined;
    if (parent.kind !== 'text' || !parent.text.includes(' ')) {
      segments.push(parent); spaceOnly.push(false); continue;
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
    // The existing string painter can trim a zero-ink suffix without cutting
    // a visible shaping island. It has no independently acquired per-word ink
    // for a contextual operation spanning several words or leading whitespace.
    // Only an ASCII SPACE suffix can be removed while keeping the complete
    // visible operation. Other whitespace (including NBSP) is not that suffix.
    if (parent.basePaintOps.some(op => op.text.replace(/ +$/u, '').includes(' '))) return undefined;
    const { textShapeRequest: _request, textLayoutService: _service, ...retainedParent } = parent;
    let operationCursor = 0;
    let ownerCursor = 0;
    let slotCursor = 0;
    for (let index = 0; index < ranges.length; index += 1) {
      const part = ranges[index]!;
      const clusters = parent.clusters.slice(part.first, part.end);
      const start = clusters[0]!.range.start;
      const end = clusters.at(-1)!.range.end;
      // Acquired pen boundaries include contextual seams exactly once. They
      // are not reconstructed by summing nominal glyph widths.
      const fromPt = index === 0 ? 0 : clusters[0]!.offset.xPt;
      const toPt = index + 1 === ranges.length ? parent.measuredWidthPt
        : parent.clusters[ranges[index + 1]!.first]!.offset.xPt;
      if (!Number.isFinite(fromPt) || !Number.isFinite(toPt) || toPt < fromPt) return undefined;
      while (parent.basePaintOps[operationCursor]
        && parent.basePaintOps[operationCursor]!.range.end <= start) operationCursor += 1;
      const ops: TextPaintOp[] = [];
      let cellClusterCursor = 0;
      for (let oi = operationCursor; oi < parent.basePaintOps.length; oi += 1) {
        const op = parent.basePaintOps[oi]!;
        if (op.range.start >= end) break;
        const a = Math.max(start, op.range.start), b = Math.min(end, op.range.end);
        while (clusters[cellClusterCursor] && clusters[cellClusterCursor]!.range.start < a) cellClusterCursor += 1;
        const cluster = clusters[cellClusterCursor];
        if (!cluster || cluster.range.start !== a || b <= a) return undefined;
        const { inkBounds, blockAxisInkBounds, ...facts } = op;
        ops.push({
          ...facts,
          text: op.text.slice(a - op.range.start, b - op.range.start),
          range: { start: a, end: b },
          offset: { ...op.offset, xPt: (a === op.range.start ? op.offset.xPt : cluster.offset.xPt) - fromPt },
          ...(!part.space && inkBounds ? { inkBounds } : {}),
          ...(!part.space && blockAxisInkBounds ? { blockAxisInkBounds } : {}),
        });
      }
      if (!ops.length || ops[0]!.range.start !== start || ops.at(-1)!.range.end !== end) return undefined;
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
        breakBefore: index === 0 ? parent.breakBefore : true,
        text: parent.text.slice(relativeStart, relativeEnd), range: { start, end },
        measuredWidthPt: toPt - fromPt,
        clusters: clusters.map(cluster => ({ ...cluster, offset: { ...cluster.offset, xPt: cluster.offset.xPt - fromPt } })),
        basePaintOps: ops,
        ...(sourceRuns ? { sourceRuns } : {}),
        ...(semanticSlotSpans ? { semanticSlotSpans } : {}),
        autoSpace: index === 0 ? parent.autoSpace : undefined,
        leadingSeamPairPt: index === 0 ? parent.leadingSeamPairPt : undefined,
        decorationTerminalAdvancePt: parent.decorationTerminalAdvancePt === undefined ? undefined
          : Math.max(0, Math.min(toPt, parent.decorationTerminalAdvancePt) - fromPt),
      };
      segments.push(cell); spaceOnly.push(part.space);
    }
  }
  parentCellStarts.push(segments.length);
  if (!hasSpaceCell) return undefined;
  const projection: SegmentLineSourceCell[] = [];
  for (let parentIndex = 0; parentIndex < parents.length; parentIndex++) {
    const parent = parents[parentIndex]!;
    for (let ci = parentCellStarts[parentIndex]!; ci < parentCellStarts[parentIndex + 1]!; ci++) {
      const cell = segments[ci]!;
      projection.push({ parentIndex,
        startUTF16: parent.kind === 'text' ? cell.range.start - parent.range.start : 0,
        endUTF16: parent.kind === 'text' ? cell.range.end - parent.range.start : 0,
        spaceOnly: spaceOnly[ci]!,
      });
    }
  }
  return { segments, parentCellStarts, spaceOnly, projection };
}
