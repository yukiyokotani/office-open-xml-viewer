import type { ParagraphLayoutSource, ParagraphTextBearingRun, TextShapeResult, TextShapeSpan, NativeBaseDirection } from './text.js';
import { sliceTextShapeRequest } from './text.js';
import type { MeasuredLinePlanSegment, MeasuredTextPlanSegment } from './paragraph.js';
import type { SourceSpaceCells } from './source-space-cells.js';
import type { SegStretch } from '../text-distribute.js';
import { indexedFontMetrics, selectResourceAverageWidthRatio, selectResourceMetric,
  mayUseExactLocalReferenceWidthMetric } from '../line-breaker/font-metrics.js';
import { referenceFontAverageWidthRatio } from '../reference-font-line-metrics.js';
import { checkedMetadataUtf16Work, checkedPrefixWorkTotal } from '../line-breaker/prefix-work-budget.js';

/** word-mixed-rtl-distribute-pitch: source opportunities, selected resources and
 * intact joining operations belong to ONE optional final-line view. Admission
 * widths stay native; this cannot reflow or repair an overflowing line. The
 * registered observation states the measured profiles and unresolved precision.
 * Unsupported lexical, acquired, resource or paint authority declines the WHOLE
 * view before the canonical host resolver is asked to choose it. */
export interface MixedRtlPitchSource {
  readonly paragraph: ParagraphLayoutSource;
  readonly end: number;
  readonly spaces: number;
  readonly followingCuts: number;
}

export function mixedRtlPitchSource(paragraph: ParagraphLayoutSource): MixedRtlPitchSource | undefined {
  if (paragraph.alignment !== 'distribute' || paragraph.bidi !== true) return undefined;
  let end = 0, spaces = 0, followingCuts = 0, previousAscii = false;
  let arabic = false, ascii = false, previousSpace = true;
  let tuple: readonly unknown[] | undefined;
  let previousKind: 'arabic' | 'ascii' | undefined;
  let pendingSpace: { ordinary: boolean; before: typeof previousKind } | undefined;
  for (const run of paragraph.runs) {
    if (run.type !== 'text') return undefined;
    const sourceRun: ParagraphTextBearingRun = run;
    // Same authored metric tuple across script slots and authored runs. This
    // is equality, not a face-name/size selector; a different next-run face or
    // CS size is outside the measured allocation even if Canvas can draw it.
    const currentTuple = [run.fontFamily, sourceRun.fontFamilyHighAnsi ?? run.fontFamily, run.fontFamilyEastAsia ?? run.fontFamily,
      run.fontFamilyCs ?? run.fontFamily, run.fontSize, run.fontSizeCs ?? run.fontSize,
      run.bold, run.boldCs ?? run.bold, run.italic, run.italicCs ?? run.italic];
    if (tuple && tuple.some((value, index) => value !== currentTuple[index])) return undefined;
    tuple ??= currentTuple;
    if (run.ruby || sourceRun.optionalHyphen || run.noteRef || run.fitTextVal != null
      || sourceRun.noBreakRanges?.length || sourceRun.noBreakBefore || sourceRun.noBreakAfter || run.eastAsianVert
      || run.allCaps || run.smallCaps || run.charSpacing || (run.charScale != null && run.charScale !== 1)
      || run.vertAlign || run.position || run.kerning || run.revision
      || run.underline || run.strikethrough || run.doubleStrikethrough || run.border
      || run.emphasisMark || run.background || run.highlight) return undefined;
    for (const character of run.text) {
      if (previousAscii) followingCuts++;
      previousAscii = false;
      if (character === ' ') {
        if (previousSpace) return undefined;
        spaces++; previousSpace = true;
        pendingSpace = { ordinary: !run.rtl && !run.cs, before: previousKind };
      } else if (/^[\x21-\x7e]$/u.test(character)) {
        if (run.rtl || run.cs) return undefined;
        // The registered view represents script-homogeneous U20-delimited
        // words only. A direct cross-script seam has no established separate
        // following-cut authority: keep the WHOLE old budget/order/paint,
        // rather than deleting one cut or inventing a replacement allocation.
        if (previousKind === 'arabic' && !previousSpace) return undefined;
        // Seam spaces adjacent to ordinary ASCII were ordinary-authored in
        // the registered evidence. An RTL-owned seam is a different UAX scope,
        // even for identical visible text, and retains its COMPLETE old view.
        if (pendingSpace && !pendingSpace.ordinary) return undefined;
        pendingSpace = undefined; previousKind = 'ascii';
        ascii = true; previousAscii = true; previousSpace = false;
      } else if (/^\p{Script=Arabic}$/u.test(character) && /^\p{Letter}$/u.test(character)) {
        if (!run.rtl) return undefined;
        if (previousKind === 'ascii' && !previousSpace) return undefined;
        if (pendingSpace && (pendingSpace.before === 'ascii' ? !pendingSpace.ordinary : pendingSpace.ordinary)) return undefined;
        pendingSpace = undefined; previousKind = 'arabic';
        arabic = true; previousSpace = false;
      } else return undefined;
      end += character.length;
    }
  }
  return arabic && ascii && spaces > 0 && !previousSpace
    ? { paragraph, end, spaces, followingCuts } : undefined;
}

type Parent = MeasuredTextPlanSegment;
export interface MixedRtlPitchPlan {
  readonly cells: SourceSpaceCells;
  readonly stretches: ReadonlyMap<number, SegStretch>;
  readonly deltaPt: number;
}

export function prepareMixedRtlPitch(parents: readonly MeasuredLinePlanSegment[], cells: SourceSpaceCells,
  source: MixedRtlPitchSource, widthPt: number): MixedRtlPitchPlan | undefined {
  const service = parents[0]?.kind === 'text' ? parents[0].textLayoutService : undefined;
  if (!service?.fontMetrics || !service.exactLocalFaces || !service.ownsGlyphs || !service.sameGlyphRaster) return;
  let cursor = 0, work = 0, sourceRunIndex = 0, sourceRunOffset = 0;
  const charge = (amount: number | undefined) => {
    const next = amount === undefined ? undefined : checkedPrefixWorkTotal(work, amount);
    if (next === undefined) return false;
    work = next; return true;
  };
  for (const parent of parents) {
    if (parent.kind !== 'text' || parent.range.start !== cursor || parent.range.end !== cursor + parent.text.length
      || parent.fontSizePt !== (parents[0] as Parent).fontSizePt
      || parent.fontWeight !== (parents[0] as Parent).fontWeight || parent.fontStyle !== (parents[0] as Parent).fontStyle
      || parent.textLayoutService !== service || !parent.textShapeRequest || parent.autoSpace || parent.leadingSeamPairPt
      || parent.characterGridDeltaPt || parent.nativeBaseDirection || parent.decorations.length
      || parent.basePaintOps.some(op => op.letterSpacingPt !== 0 || op.scaleX !== 1
        || (op.scaleY !== undefined && op.scaleY !== 1) || op.glyphOrientation !== undefined)
      || parent.textShapeRequest.text !== parent.text) return;
    const firstRequest = (parents[0] as Parent).textShapeRequest!;
    for (const slot of ['ascii', 'highAnsi', 'eastAsia', 'complexScript'] as const) {
      if (parent.textShapeRequest.fonts[slot] !== firstRequest.fonts[slot]
        || parent.textShapeRequest.themeFonts?.[slot] !== firstRequest.themeFonts?.[slot]
        || parent.textShapeRequest.themeFontPresence?.[slot] !== firstRequest.themeFontPresence?.[slot]) return;
    }
    const { text: _text, substituteContext: _context, ...metadata } = parent.textShapeRequest;
    if (!charge(checkedMetadataUtf16Work(metadata)) || !charge(parent.text.length)
      || !charge(parent.textShapeRequest.substituteContext?.text.length ?? 0)) return;
    for (let j = 0; j < parent.text.length; j++) {
      while (sourceRunIndex < source.paragraph.runs.length) {
        const run = source.paragraph.runs[sourceRunIndex]!;
        if (run.type !== 'text') return;
        if (sourceRunOffset < run.text.length) break;
        sourceRunIndex++; sourceRunOffset = 0;
      }
      const run = source.paragraph.runs[sourceRunIndex];
      if (run?.type !== 'text' || run.text.charCodeAt(sourceRunOffset++) !== parent.text.charCodeAt(j)) return;
    }
    cursor = parent.range.end;
  }
  if (cursor !== source.end) return;
  // Metadata-only plans cannot invoke the native measurer. Charge their actual
  // route/proof keys and every aggregate word request BEFORE optional native or
  // raster work. One line is bounded by the existing 16Mi UTF-16 work policy.
  const plans: TextShapeResult[] = [];
  // Temporary facts for this synchronous plan, not another document/font cache.
  // Capture each family once, so long metadata is not rescanned per SPACE.
  const familyFaces = new Map<string, ReturnType<NonNullable<typeof service.exactLocalFaces>>>();
  const facesFor = (family: string) => {
    let faces = familyFaces.get(family);
    if (!faces) { faces = service.exactLocalFaces!(family); familyFaces.set(family, faces); }
    return faces;
  };
  const metricIndex = indexedFontMetrics(service.fontMetrics);
  for (const item of parents) {
    const parent = item as Parent, request = parent.textShapeRequest!;
    const plan = service.shape({ ...request, measure: false, clusterGeometry: false });
    if (!plan.spans.length || !charge(checkedMetadataUtf16Work(plan.spans.map(span => ({
      font: span.font, route: span.fontRoute, faces: facesFor(span.font.requestedFamily),
    }))))) return;
    plans.push(plan);
  }
  const ratios = new Map<TextShapeSpan['font'], number | undefined>();
  const q = new Map<number, number>();
  const cellSpans = new Map<number, TextShapeSpan>();
  const spanCursors = parents.map(() => 0);
  const tupleKey = (span: TextShapeSpan, baseDirection?: NativeBaseDirection) => JSON.stringify([span.font.resourceIdentity,
    span.font.resolvedFamily, span.font.weight, span.font.style, span.fontRoute.fingerprint, baseDirection]);
  const witnesses = new Map<string, TextShapeSpan>();
  for (const [index, plan] of plans.entries()) for (const span of plan.spans) {
    const baseDirection = (parents[index] as Parent).textShapeRequest!.baseDirection;
    if (/[^ ]/u.test(span.text)) witnesses.set(tupleKey(span, baseDirection), span);
  }
  let sumQ = 0, visible = 0, foundSpaces = 0, foundCuts = 0;
  for (let i = 0; i < cells.segments.length; i++) {
    const cell = cells.segments[i];
    if (cell?.kind !== 'text') return;
    const projection = cells.projection[i]!, parent = parents[projection.parentIndex] as Parent;
    const shape = plans[projection.parentIndex]!;
    const offset = cell.range.start - parent.range.start;
    let spanIndex = spanCursors[projection.parentIndex]!;
    while (spanIndex < shape.spans.length && shape.spans[spanIndex]!.end <= offset) spanIndex++;
    spanCursors[projection.parentIndex] = spanIndex;
    const span = shape.spans[spanIndex];
    if (!span || span.start > offset || span.end < offset + cell.text.length) return;
    cellSpans.set(i, span);
    if (cells.spaceOnly[i]) {
      if (cell.text !== ' ' || cell.clusters.length !== 1) return;
      let ratio = ratios.get(span.font);
      if (!ratios.has(span.font)) {
        ratio = selectResourceAverageWidthRatio(metricIndex, span.font, ' ');
        if (ratio === undefined && mayUseExactLocalReferenceWidthMetric(span.font)) {
          ratio = referenceFontAverageWidthRatio(span.font.requestedFamily, span.font.weight, span.font.style);
        }
        ratios.set(span.font, ratio);
      }
      const value = ratio === undefined ? NaN : ratio * parent.fontSizePt;
      if (!Number.isFinite(value) || value <= 0) return;
      q.set(i, value); sumQ += value; foundSpaces++;
    } else {
      // Signed contextual cluster deltas are authoritative; never replace them
      // with isolated glyph widths or clamp a negative joining contribution.
      for (const cluster of cell.clusters) {
        if (!Number.isFinite(cluster.advancePt)) return;
        visible += cluster.advancePt;
      }
      if (/^[\x21-\x7e]+$/u.test(cell.text)) {
        if (cell.rtl || cell.basePaintOps.length !== 1 || cell.clusters.length !== cell.text.length
          || cell.clusters.some((c, j) => c.range.start !== cell.range.start + j || c.range.end !== c.range.start + 1)) return;
        foundCuts += cell.text.length - (cell.range.end === source.end ? 1 : 0);
        if (cell.text.length > 1) {
          const request = sliceTextShapeRequest(parent.textShapeRequest!, projection.startUTF16, projection.endUTF16);
          const { text: _text, substituteContext: _context, ...metadata } = request;
          if (!charge(checkedMetadataUtf16Work(metadata)) || !charge(request.text.length)
            || !charge(request.substituteContext?.text.length ?? 0)) return;
        }
      } else if (!cell.rtl || !/^[\p{Script=Arabic}]+$/u.test(cell.text)) return;
    }
  }
  const deltaPt = (widthPt - visible - sumQ) / (source.spaces + source.followingCuts);
  if (foundSpaces !== source.spaces || foundCuts !== source.followingCuts
    || !Number.isFinite(visible) || !Number.isFinite(sumQ) || !Number.isFinite(deltaPt) || deltaPt < 0) return;
  // Every proof key and the actual short native aggregate request is charged
  // atomically. The two raster reads belong to ONE cached service proof; this
  // line-local plan adds no cache, font selector or per-glyph native requests.
  const proofs: Array<{ span: TextShapeSpan; witness: TextShapeSpan; family: string; probe: string; baseDirection?: NativeBaseDirection }> = [];
  for (const [index, shape] of plans.entries()) for (const span of shape.spans) {
    const baseDirection = (parents[index] as Parent).textShapeRequest!.baseDirection;
    const font = span.font;
    if (!mayUseExactLocalReferenceWidthMetric(font)) return;
    const face = facesFor(font.requestedFamily).find(f => f.resolvedFamily === font.resolvedFamily
      && f.weight === font.weight && f.style === font.style);
    const metric = selectResourceMetric(metricIndex, font, span.text);
    const probe = metric?.glyphCoverageProbeFamily ?? face?.coverageProbeFamily;
    // The actual selected resource metric owns its registered probe. A
    // tuple-only fallback cannot choose among conflicting same-family probes.
    if (!metric?.glyphCoverageProbeFamily && face && facesFor(font.requestedFamily).some(f =>
      f.resolvedFamily === font.resolvedFamily && f.weight === font.weight && f.style === font.style
      && f.coverageProbeFamily !== face.coverageProbeFamily)) return;
    const witness = /[^ ]/u.test(span.text) ? span : witnesses.get(tupleKey(span, baseDirection));
    if (!face || !probe || !witness) return;
    const ownership = { ...(baseDirection ? { baseDirection } : {}), text: span.text, coverageProbeFamily: probe, weight: font.weight, style: font.style };
    const raster = { ...(baseDirection ? { baseDirection } : {}), text: witness.text, route: witness.fontRoute, faceFamily: face.resolvedFamily,
      weight: font.weight, style: font.style };
    if (!charge(checkedMetadataUtf16Work(ownership)) || !charge(span.text.length)
      || !charge(checkedMetadataUtf16Work(raster)) || !charge(witness.text.length * 2)) return;
    proofs.push({ span, witness, family: face.resolvedFamily, probe, ...(baseDirection ? { baseDirection } : {}) });
  }
  for (let i = 0; i < cells.segments.length; i++) {
    const cell = cells.segments[i] as Parent;
    if (cells.spaceOnly[i] || cell.text.length === 1 || !/^[\x21-\x7e]+$/u.test(cell.text)) continue;
    const projection = cells.projection[i]!, parent = parents[projection.parentIndex] as Parent;
    const request = sliceTextShapeRequest(parent.textShapeRequest!, projection.startUTF16, projection.endUTF16);
    const { text: _text, substituteContext: _context, ...metadata } = request;
    if (!charge(checkedMetadataUtf16Work({ ...metadata, letterSpacingPt: deltaPt,
      measure: true, clusterGeometry: false })) || !charge(request.text.length)) return;
  }
  for (const { span, witness, family, probe, baseDirection } of proofs) {
    if (!service.ownsGlyphs({ ...(baseDirection ? { baseDirection } : {}), text: span.text, coverageProbeFamily: probe,
      weight: span.font.weight, style: span.font.style })
      || !service.sameGlyphRaster({ ...(baseDirection ? { baseDirection } : {}), text: witness.text, route: witness.fontRoute, faceFamily: family,
        weight: span.font.weight, style: span.font.style })) return;
  }
  const stretches = new Map<number, SegStretch>();
  const selected = [...cells.segments];
  for (let i = 0; i < selected.length; i++) {
    const cell = selected[i] as Parent;
    if (cells.spaceOnly[i]) {
      const adjustment = q.get(i)! + deltaPt - cell.measuredWidthPt;
      // The existing source-cell geometry owns expansion only, not a new
      // separator contraction or terminal hanging policy.
      if (!Number.isFinite(adjustment) || adjustment < 0) return;
      stretches.set(i, { splitBefore: [], internalStretch: 0, trailingGap: true, trailingDelta: adjustment });
    } else if (/^[\x21-\x7e]+$/u.test(cell.text)) {
      if (cell.text.length === 1) {
        // No internal cut means the existing painter keeps its ORIGINAL native
        // operation. A following source cut changes only the outside advance.
        stretches.set(i, { splitBefore: [], internalStretch: 0,
          trailingGap: cell.range.end < source.end, trailingDelta: deltaPt });
        continue;
      }
      const projection = cells.projection[i]!, parent = parents[projection.parentIndex] as Parent;
      const request = sliceTextShapeRequest(parent.textShapeRequest!, projection.startUTF16, projection.endUTF16);
      const tracked = service.shape({ ...request, letterSpacingPt: deltaPt, measure: true, clusterGeometry: false });
      const span = tracked.spans[0], original = cellSpans.get(i);
      if (!original || tracked.spans.length !== 1 || !span || span.start !== 0 || span.end !== request.text.length
        || span.text !== request.text || span.fontRoute.fingerprint !== original.fontRoute.fingerprint
        || span.font.resourceIdentity !== original.font.resourceIdentity || span.font.resolvedFamily !== original.font.resolvedFamily
        || span.font.weight !== original.font.weight || span.font.style !== original.font.style
        || !tracked.inkBounds || !Object.values(tracked.inkBounds).every(Number.isFinite)
        || tracked.inkBounds.xMaxPt < tracked.inkBounds.xMinPt
        || tracked.inkBounds.ascentPt < 0 || tracked.inkBounds.descentPt < 0
        || !Number.isFinite(tracked.advancePt) || tracked.ascentPt !== cell.selectedFaceFontBox?.ascentPt || tracked.descentPt !== cell.selectedFaceFontBox?.descentPt) return;
      // Honoring the SAME numeric request is the existing provider contract,
      // not an exact binary64 N*delta identity. Native spacing storage may be
      // quantized; never turn this aggregate advance into a second fit oracle,
      // compare it with a new allowance, or retune the source-owned budget.
      // Native aggregate advance may include the trailing tracking charge. It
      // is NEVER used for fitting or V: source internal cuts own (N-1)*delta,
      // and the final following cut exists only before actual source EOF.
      const operation = cell.basePaintOps[0]!;
      selected[i] = { ...cell, basePaintOps: [{ ...operation, inkBounds: tracked.inkBounds,
        // Both published ink views describe this SAME native receipt. Keep the
        // existing glyph-offset convention; acquisition/font-box metrics and
        // horizontal line allocation remain untouched.
        blockAxisInkBounds: {
          startPt: (operation.glyphOffsetPt?.yPt ?? 0) - tracked.inkBounds.ascentPt,
          endPt: (operation.glyphOffsetPt?.yPt ?? 0) + tracked.inkBounds.descentPt,
        },
      }] };
      const splitBefore = Array.from({ length: cell.text.length - 1 }, (_, j) => j + 1);
      stretches.set(i, { splitBefore, internalStretch: splitBefore.length * deltaPt,
        trailingGap: cell.range.end < source.end, trailingDelta: deltaPt });
    }
  }
  return { cells: { ...cells, segments: selected }, stretches, deltaPt };
}
