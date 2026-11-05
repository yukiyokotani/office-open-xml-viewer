import { bidiClass } from '@silurus/ooxml-core/internal/bidi-line';
import { lineOrderEngineIdentity, sameEngineState, resolveNativeSourceLevelUnits } from '../bidi-line.js';
import { nativeVisiblePrefixMatches, sliceTextShapeRequest } from '../layout/text.js';
import type { ParagraphLayoutRun } from '../layout/text.js';
import { fitsLineWidth } from './fit-search.js';
import { RESET_SLICED_TEXT_MEASUREMENT, slicedTextMetadata } from './advance.js';
import type { LayoutSeg, LayoutTextSeg, LineLayoutEnvironment } from './model.js';

/** Normative shaping boundary: Unicode 17 / UAX #9 rev.51 §3.5, through the
 * SAME whole authored host resolver. The initial library capability is narrower:
 * one ordinary strong-letter source run, ASCII SPACE, and a true nonnumeric
 * scope-final period. Digits, marks, controls, weak punctuation, transformations
 * and unknown author scopes remain ordinary BEFORE directed acquisition. */
export function acquireNativeSourceUnits(
  original: LayoutSeg[], runs: readonly ParagraphLayoutRun[], environment: LineLayoutEnvironment,
): LayoutSeg[] {
  const band = environment.nativeUnitMinimumBandPt;
  const service = environment.layoutServices?.text;
  if (band === undefined || !service?.acquireProjectableNativeUnits || !service.shapeDirected || runs.length !== 1) return original;
  // The band owner supplies the complete canonical paragraph, not a source
  // window. With exactly one ordinary run, paragraph end is also the authored
  // LRI scope end; a font/run boundary followed by another unmarked run would
  // NOT establish this EOS period authority and is outside this capability.
  const run = runs[0]!;
  if (run.type !== 'text' || run.rtl === true || run.cs === true || !run.text
    || run.smallCaps || run.allCaps || run.vertAlign != null
    || run.text.startsWith(' ') || !/\p{Letter}/u.test(run.text)) return original;
  let sawL = false, sawR = false;
  for (let offset = 0; offset < run.text.length;) {
    const code = run.text.codePointAt(offset)!;
    const character = String.fromCodePoint(code);
    const type = bidiClass(code);
    if (character !== ' ' && !(character === '.' && offset + 1 === run.text.length)) {
      if (!/\p{Letter}/u.test(character) || (type !== 'L' && type !== 'R' && type !== 'AL')) return original;
      sawL ||= type === 'L'; sawR ||= type === 'R' || type === 'AL';
    }
    offset += character.length;
  }
  if (!sawR || (!sawL && !run.text.endsWith('.'))) return original;
  const engine = lineOrderEngineIdentity();
  const resolved = resolveNativeSourceLevelUnits([{ text: run.text, rtl: false, digitsAsAN: false }], true);
  // Only the existing branded, unmodified Unicode-17 engine certifies this
  // closed strong-source theorem. A forwarding/custom engine remains ordinary.
  if (!resolved) return original;
  const units: Array<{ start: number; end: number; baseDirection: 'ltr' | 'rtl' }> = [];
  for (const part of resolved) {
    const text = run.text.slice(part.startUTF16, part.endUTF16);
    // L1 SPACE reset is not a visible native cut. Its independently acquired
    // zero-ink charged source cell is projected by the final whole-host view.
    if (/^ +$/u.test(text) && units.length > 0) { units[units.length - 1]!.end = part.endUTF16; continue; }
    units.push({ start: part.startUTF16, end: part.endUTF16, baseDirection: (part.level & 1) ? 'rtl' : 'ltr' });
  }
  if (units.length < 2) return original;
  const candidate: LayoutTextSeg[] = [];
  let cursor = 0;
  let hasVisibleCut = false;
  let unitIndex = 0;
  for (const segment of original) {
    if (!('text' in segment) || !segment.textShapeRequest || segment.metricOnly || segment.ruby
      || segment.tateChuYoko || segment.vertAlign != null || segment.script === 'eastAsia'
      || segment.latinSpaceCompressionEligible || segment.mixedSpaceAverageWidthRatio !== undefined
      || segment.widthBalanceSpaceSequence || segment.fitTextRegionIndex !== undefined || segment.verticalRun
      || segment.optionalHyphen || segment.optionalHyphenBreaks || segment.emphasisMark || segment.semanticSlotSpans
      || (segment.charSpacing ?? 0) !== 0 || (segment.charScale ?? 1) !== 1
      || segment.autoSpaceBeforePt !== undefined || segment.punctuationCompressions?.length) return original;
    const request = segment.textShapeRequest;
    if (request.substituteContext?.text !== run.text || request.substituteContext.offset !== cursor
      || run.text.slice(cursor, cursor + segment.text.length) !== segment.text) return original;
    const end = cursor + segment.text.length;
    while (units[unitIndex] && units[unitIndex]!.end <= cursor) unitIndex++;
    let currentUnit = unitIndex;
    let firstVisibleBase: 'ltr' | 'rtl' | undefined;
    // Both source partitions are ordered. Walk their intersections once;
    // scanning every unit for every word would be quadratic on long runs.
    for (; currentUnit < units.length && units[currentUnit]!.start < end; currentUnit++) {
      const unit = units[currentUnit]!;
      const first = Math.max(cursor, unit.start), last = Math.min(end, unit.end);
      if (first >= last) continue;
      if (/[^ ]/u.test(run.text.slice(first, last))) {
        if (firstVisibleBase === undefined) firstVisibleBase = unit.baseDirection;
        else hasVisibleCut ||= firstVisibleBase !== unit.baseDirection;
      }
      const start = first - cursor, finish = last - cursor;
      const text = segment.text.slice(start, finish);
      if (!text.trimEnd() || text.replace(/ +$/u, '').includes(' ')) return original;
      candidate.push({ ...segment, ...RESET_SLICED_TEXT_MEASUREMENT,
        ...slicedTextMetadata(segment, start, finish), text, measuredWidth: 0,
        // Native cuts inside a physical word are NOT legal source word breaks.
        joinPrev: start > 0 ? true : segment.joinPrev,
        textShapeRequest: { ...sliceTextShapeRequest(request, start, finish), baseDirection: unit.baseDirection },
        nativeLevelUnit: Object.freeze({ sourceText: run.text, ...unit }),
      });
    }
    unitIndex = Math.max(unitIndex, currentUnit - 1);
    cursor = end;
  }
  // Different directions in separate native words are already representable;
  // SPACE-only L1 cuts do not activate this capability. A visible boundary must
  // lie INSIDE an existing ordinary native operation before pair partners may
  // join the same selected source bundle.
  if (!hasVisibleCut || cursor !== run.text.length) return original;
  // A level seam is not itself proof of an original grapheme boundary:
  // Letter/strong filtering still admits GCB Prepend letters. Only actual
  // whole-source segmentation certifies these new cuts; a code-point fallback
  // or segmentation of already-cut children would manufacture that proof.
  if (typeof Intl === 'undefined' || typeof Intl.Segmenter !== 'function') return original;
  const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    .segment(run.text)[Symbol.iterator]();
  let boundary = 0;
  const isOriginalBoundary = (offset: number): boolean => {
    while (boundary < offset) {
      const next = graphemes.next();
      boundary = next.done ? run.text.length : next.value.index;
      if (next.done) break;
    }
    return boundary === offset;
  };
  let candidateEnd = 0;
  for (const piece of candidate) {
    const start = piece.textShapeRequest!.substituteContext!.offset;
    const end = start + piece.text.length;
    if (start !== candidateEnd || run.text.slice(start, end) !== piece.text
      || !isOriginalBoundary(start) || !isOriginalBoundary(end)) return original;
    candidateEnd = end;
  }
  if (candidateEnd !== run.text.length) return original;
  const requests = candidate.map(piece => ({ ...piece.textShapeRequest!,
    baseDirection: piece.nativeLevelUnit!.baseDirection, measure: true, clusterGeometry: 'spaces' as const }));
  const results = service.acquireProjectableNativeUnits(requests);
  if (!results) return original;
  if (results.length !== candidate.length) throw new TypeError('Native unit source bundle cardinality mismatch');
  let wordAdvance = 0;
  for (let index = 0; index < candidate.length; index++) {
    const piece = candidate[index]!, entry = results[index]!, shape = entry.shape;
    if (piece.nativeLevelUnit!.baseDirection === 'rtl' && piece.text.endsWith(' ')) {
      if (!nativeVisiblePrefixMatches(requests[index]!, shape, entry.visiblePrefix)) return original;
      piece.nativeVisiblePrefix = entry.visiblePrefix;
    }
    if (shape.nativeBaseDirection !== piece.nativeLevelUnit!.baseDirection
      || shape.spans.map(span => span.text).join('') !== piece.text) throw new TypeError('Native unit source/receipt mismatch');
    if (!Number.isFinite(shape.advancePt) || shape.advancePt < 0
      || shape.spans.some(span => span.fontRoute.fingerprint !== piece.fontRoute?.fingerprint)) return original;
    if (!piece.joinPrev) wordAdvance = 0;
    wordAdvance += shape.advancePt;
    if (!fitsLineWidth(wordAdvance, band)) return original;
  }
  // The supported producer is allowed to call the existing mutable engine
  // setter. Geometry receipts do not certify that later host policy; decline
  // this temporary bundle if its exact engine/lifecycle no longer owns it.
  return sameEngineState(engine, lineOrderEngineIdentity()) ? candidate : original;
}
