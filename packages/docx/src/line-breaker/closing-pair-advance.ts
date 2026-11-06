import { checkedMetadataUtf16Work, checkedPrefixWorkTotal } from './prefix-work-budget.js';
import { segmentsHaveRtl } from '../bidi-line.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { independentTextShapeRequest, sliceTextShapeRequest, type ParagraphLayoutSource, type ParagraphTextBearingRun, type TextShapeResult, type TextLayoutService } from '../layout/text.js';
import type { ParagraphMeasurementEnvironment } from '../paragraph-measure.js';
import { fitsLineWidth } from './fit-search.js';
import type { LayoutSeg, LayoutTextSeg, LineLayoutEnvironment } from './model.js';

/** word-first-owner-closing-pair-advance (WORD_FIRST_OWNER_CLOSING_PAIR_ADVANCE):
 * the first scalar's complete request owns the selected cell and unchanged
 * native/contextual advance. Following typography supplies adjacency only.
 * Unknown glyph/route authority, prior reductions and unsupported source
 * boundaries preserve their old geometry. Evidence and the explicit absent-EA
 * Japanese library default are registered in layout/line-compatibility.ts;
 * document CSC eligibility is independent. */
const CLOSING_PAIRS = new Set(['。」', '、」', '」」', '」。', '」、', '」「', '）（']);

/** Dedicated root-story provenance, resolved identically for intrinsic and
 * bounded acquisition. Compatibility shape contexts do not retain WML authority. */
export function closingPairScopeApplies(paragraph: ParagraphLayoutSource, context: ParagraphLayoutContext,
  environment: ParagraphMeasurementEnvironment): boolean {
  // The registered left/both allocation excludes Kashida and Thai. A final
  // Thai carry does not establish the soft-line tracking needed after an
  // internal reduction; retain the whole paragraph's existing geometry until
  // that spacing authority is established. Other alignments remain unclaimed.
  // Read the parser-resolved value (both -> justify); unknown or missing
  // values do not establish eligibility.
  return (paragraph.alignment === 'left' || paragraph.alignment === 'justify')
    && environment.closingPairRootSource === true && context.hasEastAsianText && !paragraph.framePr
    && !context.runDirectionUnauthored && !context.baseRtl && !context.hasRuby
    && !context.lineGrid.active && !context.characterGrid.active
    && environment.pageWritingMode === 'horizontal-tb' && !environment.verticalPageFrame;
}

function ordinary(segment: LayoutSeg | undefined, runs: readonly ParagraphLayoutSource['runs'][number][]): segment is LayoutTextSeg {
  if (!segment || !('text' in segment) || !segment.text || !segment.textShapeRequest || !segment.textLayoutService
    || segment.metricOnly || segment.verticalRun || segment.ruby || segment.rubyPiece || segment.tateChuYoko
    || segment.optionalHyphen || segment.optionalHyphenGlyph || segment.optionalHyphenWord
    || segment.optionalHyphenBreaks?.length || segment.noBreakRanges?.length
    || segment.fitTextRegionIndex !== undefined || segment.fitTextPerGapPx !== undefined
    || segment.charSpacing || (segment.charScale !== undefined && segment.charScale !== 1)
    || segment.vertAlign || segment.position) return false;
  const run = segment.sourceRunIndex === undefined ? undefined : runs[segment.sourceRunIndex];
  return run?.type === 'text' && !run.rtl && !run.cs && !run.allCaps && !run.smallCaps;
}
const equal = (a: number, b: number) => Number.isFinite(a) && Number.isFinite(b)
  && fitsLineWidth(a, b) && fitsLineWidth(b, a);

export function retainClosingPairAdvances(runs: readonly ParagraphLayoutSource['runs'][number][],
  environment: LineLayoutEnvironment, segments: LayoutSeg[]): void {
  if (!environment.closingPairAdvanceAdmitted
    || ![undefined, 11, 12, 14, 15].includes(environment.compatibilityMode)) return;
  const service = environment.layoutServices?.text;
  if (!service?.acquireContextualAdvances || !service.exactLocalFaces || !service.ownsGlyphs || !service.sameGlyphRaster) return;
  // A complete ordinary source class, including constituents hidden by an
  // identical-format text sequence. Fields/protected/discretionary units and
  // controls are not established pair boundaries, even if they emit no glyph.
  if (runs.some(run => {
    if (run.type !== 'text') return true;
    const source: ParagraphTextBearingRun = run;
    return source.ruby || source.optionalHyphen || source.noteRef || source.fitTextVal != null
      || source.noBreakRanges?.length || source.noBreakBefore || source.noBreakAfter
      || source.eastAsianVert || source.allCaps || source.smallCaps || /[\p{Cc}\p{Cf}]/u.test(run.text);
  })) return;
  const groups: Array<{ segment: LayoutTextSeg; ranges: Array<{ start: number; end: number }>; pairs: string[] }> = [];
  for (let index = 0; index < segments.length; index++) {
    const segment = segments[index];
    if (!ordinary(segment, runs) || segment.textLayoutService !== service) continue;
    const language = segment.textShapeRequest!.eastAsiaLanguage?.toLowerCase();
    if (language != null && language !== 'ja-jp') continue;
    const ranges: Array<{ start: number; end: number }> = [], pairs: string[] = [];
    let compressionIndex = 0;
    for (let start = 0; start < segment.text.length; start++) {
      const first = segment.text[start]!;
      let second = segment.text[start + 1];
      if (second === undefined) {
        const next = segments[index + 1];
        if (ordinary(next, runs) && next.textLayoutService === service) {
          const context = segment.textShapeRequest!.substituteContext;
          const nextContext = next.textShapeRequest!.substituteContext;
          const contiguous = context && nextContext && (
            context.text === nextContext.text && context.offset + segment.text.length === nextContext.offset
            || next.sourceRunIndex === segment.sourceRunIndex! + 1
              && context.offset + segment.text.length === context.text.length && nextContext.offset === 0);
          if (contiguous) second = next.text[0];
        }
      }
      const pair = first + (second ?? '');
      while (compressionIndex < (segment.punctuationCompressions?.length ?? 0)
        && segment.punctuationCompressions![compressionIndex]!.end < start + 1) compressionIndex++;
      if (!CLOSING_PAIRS.has(pair) || segment.punctuationCompressions?.[compressionIndex]?.end === start + 1) continue;
      ranges.push({ start, end: start + 1 }); pairs.push(pair);
    }
    if (ranges.length) groups.push({ segment, ranges, pairs });
  }
  if (!groups.length || segmentsHaveRtl(segments)) return;
  // One atomic optional batch preflights ALL prefix/key work before native or
  // glyph/raster proofs. No full-source all-grapheme acquisition and no new
  // prefix cache. Refusal keeps every source advance unchanged.
  // Metadata planning itself is bounded before its shape keys are built.
  // One possible resolver/key operation per source unit is a conservative
  // bound; cache hits count. No full context is serialized per short request.
  const plans: Array<{ additionalWorkUtf16: number; shape?: TextShapeResult; faces?: readonly ReturnType<NonNullable<TextLayoutService['exactLocalFaces']>>[] }> = [];
  let planningWork = 0;
  for (const group of groups) {
    const { text: _text, substituteContext: _context, ...metadata } = group.segment.textShapeRequest!;
    const metadataWork = checkedMetadataUtf16Work({ ...metadata, measure: false, clusterGeometry: false });
    const work = metadataWork === undefined ? undefined : checkedPrefixWorkTotal(group.segment.text.length,
      metadataWork * Math.max(1, group.segment.text.length));
    const contextWork = work === undefined ? undefined : checkedPrefixWorkTotal(work, group.segment.textShapeRequest!.substituteContext?.text.length ?? group.segment.text.length);
    const next = contextWork === undefined ? undefined : checkedPrefixWorkTotal(planningWork, contextWork);
    if (next === undefined) return;
    planningWork = next;
    plans.push({ additionalWorkUtf16: contextWork! });
  }
  for (let i = 0; i < groups.length; i++) {
    const group = groups[i]!, request = group.segment.textShapeRequest!, plan = plans[i]!;
    plan.shape = service.shape({ ...request, measure: false, clusterGeometry: false });
    const { text: _text, substituteContext: _context, ...metadata } = request;
    const shortWork = checkedMetadataUtf16Work({ ...metadata, measure: true, clusterGeometry: false, fontHint: 'eastAsia' });
    const originalWork = checkedMetadataUtf16Work(metadata);
    if (shortWork === undefined || originalWork === undefined) return;
    let routeWork = 0;
    const ownedFaces: Array<ReturnType<NonNullable<TextLayoutService['exactLocalFaces']>>> = [];
    for (const span of plan.shape.spans) {
      const faces = Object.freeze(service.exactLocalFaces(span.font.requestedFamily).map(face => Object.freeze({
        resolvedFamily: face.resolvedFamily, weight: face.weight, style: face.style,
        ...(face.coverageProbeFamily ? { coverageProbeFamily: face.coverageProbeFamily } : {}),
      })));
      ownedFaces.push(faces);
      // Include all exact faces in the planning bound, before any glyph/raster
      // proof. The actual source/tuple is still validated from the receipt.
      const dataWork = checkedMetadataUtf16Work({ route: span.fontRoute, faces,
        requestedFamily: span.font.requestedFamily, resourceIdentity: span.font.resourceIdentity,
        weight: span.font.weight, style: span.font.style });
      const next = dataWork === undefined ? undefined : checkedPrefixWorkTotal(routeWork, dataWork);
      if (next === undefined) return;
      routeWork = next;
    }
    plan.faces = Object.freeze(ownedFaces);
    // Four short A/C/P/B shape metadata keys; three ownership scalar keys,
    // one service raster key and four native raster font selections. Charge
    // the whole route/tuple/coverage descriptor for each, even on cache hits.
    const proofWork = checkedPrefixWorkTotal(4 * (shortWork + originalWork),
      8 * routeWork + 5 * 6 + 5); // five short source units + escaped key text
    const total = proofWork === undefined ? undefined : checkedPrefixWorkTotal(plan.additionalWorkUtf16,
      proofWork * group.ranges.length);
    if (total === undefined) return;
    plan.additionalWorkUtf16 = total;
  }
  // A known missing exact tuple cannot become a positive resource receipt.
  // Reject that whole optional acquisition before spending native prefix work;
  // glyph ownership/raster remain mandatory after the atomic budget below.
  if (!groups.some((group, i) => {
    const plan = plans[i]!;
    let spanIndex = 0;
    return group.ranges.some(range => {
      while (spanIndex < plan.shape!.spans.length && plan.shape!.spans[spanIndex]!.end <= range.start) spanIndex++;
      const span = plan.shape!.spans[spanIndex];
      return span?.font.source === 'local' && span.font.resourceIdentity?.startsWith('office-local:')
        && plan.faces![spanIndex]!.some(face => face.resolvedFamily === span.font.resolvedFamily
          && face.weight === span.font.weight && face.style === span.font.style && face.coverageProbeFamily);
    });
  })) return;
  const batch = groups.map((g, i) => ({ request: g.segment.textShapeRequest!, ranges: g.ranges,
    additionalWorkUtf16: plans[i]!.additionalWorkUtf16 }));
  const receipts = service.acquireContextualAdvances(batch);
  if (!receipts || receipts.length !== groups.length || receipts.some((r, i) => r.clusters.length !== groups[i]!.ranges.length
    || !Number.isFinite(r.shape.advancePt)
    || r.shape.spans.length !== plans[i]!.shape!.spans.length
    || r.shape.spans.some((span, j) => {
      const planned = plans[i]!.shape!.spans[j]!;
      return span.start !== planned.start || span.end !== planned.end || span.text !== planned.text
        || span.fontRoute.fingerprint !== planned.fontRoute.fingerprint
        || span.fontRoute.familyList !== planned.fontRoute.familyList || span.fontRoute.scope !== planned.fontRoute.scope
        || span.font.resourceIdentity !== planned.font.resourceIdentity
        || span.font.resolvedFamily !== planned.font.resolvedFamily || span.font.requestedFamily !== planned.font.requestedFamily
        || span.font.weight !== planned.font.weight || span.font.style !== planned.font.style;
    })
    || r.clusters.some((c, j) => !Number.isFinite(c.offsetPt) || !Number.isFinite(c.advancePt) || c.range.start !== groups[i]!.ranges[j]!.start || c.range.end !== groups[i]!.ranges[j]!.end))) return;
  for (let i = 0; i < groups.length; i++) {
    const { segment, ranges, pairs } = groups[i]!, receipt = receipts[i]!;
    const request = segment.textShapeRequest!;
    const adjustments: Array<{ end: number; adjustmentPt: number }> = [];
    let spanIndex = 0;
    for (let j = 0; j < ranges.length; j++) {
      const range = ranges[j]!, pair = pairs[j]!;
      while (spanIndex < receipt.shape.spans.length && receipt.shape.spans[spanIndex]!.end <= range.start) spanIndex++;
      const span = receipt.shape.spans[spanIndex];
      if (!span || span.start > range.start || span.end < range.end || span.font.source !== 'local'
        || !span.font.resourceIdentity?.startsWith('office-local:')) continue;
      const face = plans[i]!.faces![spanIndex]!.find(f => f.resolvedFamily === span.font.resolvedFamily
        && f.weight === span.font.weight && f.style === span.font.style);
      if (!face?.coverageProbeFamily || !service.ownsGlyphs({ text: pair + '一', weight: span.font.weight,
        style: span.font.style, coverageProbeFamily: face.coverageProbeFamily })
        || !service.sameGlyphRaster({ text: pair + '一', route: span.fontRoute, faceFamily: face.resolvedFamily,
          weight: span.font.weight, style: span.font.style })) continue;
      const acquire = (r: typeof request) => service.shape({ ...r, measure: true, clusterGeometry: false });
      const a = acquire(sliceTextShapeRequest(request, range.start, range.end));
      const c = acquire({ ...independentTextShapeRequest(request, '一'), fontHint: 'eastAsia' });
      const p = acquire(range.end < request.text.length ? sliceTextShapeRequest(request, range.start, range.end + 1)
        : independentTextShapeRequest(request, pair));
      const b = acquire(independentTextShapeRequest(request, pair[1]!));
      const same = (shape: TextShapeResult) => shape.spans.length === 1
        && shape.spans[0]!.fontRoute.fingerprint === span.fontRoute.fingerprint
        && shape.spans[0]!.font.resourceIdentity === span.font.resourceIdentity
        && shape.spans[0]!.font.weight === span.font.weight && shape.spans[0]!.font.style === span.font.style;
      if (![a, c, p, b].every(same) || !(a.advancePt > 0 && c.advancePt > 0 && b.advancePt > 0)
        || !equal(a.advancePt, c.advancePt) || !equal(p.advancePt - b.advancePt, a.advancePt)
        || !equal(receipt.clusters[j]!.advancePt, a.advancePt)) continue;
      adjustments.push({ end: range.end, adjustmentPt: c.advancePt / 2 - a.advancePt });
    }
    if (adjustments.length) segment.punctuationCompressions = Object.freeze([
      ...(segment.punctuationCompressions ?? []), ...adjustments,
    ].sort((a, b) => a.end - b.end).map(a => Object.freeze(a)));
  }
}
