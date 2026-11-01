import { wordArabicTextKerned, wordKerningApplies } from '../layout/line-compatibility.js';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import type { LayoutTextSeg } from '../line-layout.js';
import type { MeasurementTextContext, VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx, semanticSlotStartIndex } from '../layout/text.js';
import {
  charScaleFactor,
  CONDENSED_CONTEXT_GRAPHEMES,
  condensedCharacterAdjustments,
  condensedSpacingMayFloor,
  contextualGraphemeAdvances,
  segLetterSpacingPx,
  segmentCharacterGridDeltaPx,
  type CondensedAdvanceUnits,
  type NaturalClusterAdvance,
} from './advance.js';
import type { DocGridCtx } from './model.js';
import { verticalRunInkExtra } from './vertical-text.js';

/** Measured UTF-16 units per text unit charged for one vertical cell plan:
 * routing probes (natural, vertical-form and ink probes) plus one
 * {@link CONDENSED_CONTEXT_GRAPHEMES}-bounded context pair per sideways
 * grapheme, the whole sideways piece and a floored cluster's drawn advance,
 * plus the two bidi level passes of a split piece. */
const VERTICAL_PLAN_WORK_PER_UNIT = 2 * (CONDENSED_CONTEXT_GRAPHEMES + 1) + 9;

/** Ownership retains the source's slots even after emergency slicing. Only
 * overlapping slots describe the current physical seam; an all-ASCII suffix
 * must keep the ordinary boundary contract despite its original mixed owner. */
function hasLiveSlotSeam(segment: LayoutTextSeg | undefined): boolean {
  const slots = segment?.semanticSlotSpans;
  if (!segment || !slots) return false;
  const start = segment.semanticSlotRange?.start ?? 0;
  const end = segment.semanticSlotRange?.end ?? segment.text.length;
  const first = slots[semanticSlotStartIndex(slots, start)];
  return first !== undefined && first.end < end;
}

/** Owns the Canvas state used by one line-breaking pass. The state recorded here
 * describes only assignments made by this adapter; a future advance cache must
 * also account for the document-scoped shaping service and its feature state. */
export class LineMeasurementAdapter {
  private selectedFont: string | null = null;
  private readonly initialFont: string;
  private selectedKerning: CanvasFontKerning | null;
  readonly letterSpacing: string;

  constructor(
    private readonly context: MeasurementTextContext,
    private readonly scale: number,
    private readonly fontForSegment: (segment: LayoutTextSeg) => string,
    private readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  ) {
    this.initialFont = context.font;
    this.selectedKerning = context.fontKerning;
    this.letterSpacing = context.letterSpacing;
  }

  setFont(font: string): void {
    if (font !== this.selectedFont) {
      this.context.font = font;
      this.selectedFont = font;
    }
  }

  selectSegmentFont(segment: LayoutTextSeg): void {
    this.setFont(this.fontForSegment(segment));
  }

  private selectSegmentKerning(segment: LayoutTextSeg): CanvasFontKerning | null {
    // ECMA-376 §17.3.2.19: an absent w:kern disables pair kerning even when
    // Canvas would otherwise choose its automatic kerning behavior.
    const previous = this.selectedKerning;
    const selected = this.segmentKerning(segment);
    this.context.fontKerning = selected;
    this.selectedKerning = selected;
    return previous;
  }

  private segmentKerning(segment: LayoutTextSeg): CanvasFontKerning {
    // The acquisition request owns the decision; retained paint consumes that
    // same value. Standalone segments apply the same content-text rule.
    return (segment.textShapeRequest?.kerning ?? wordKerningApplies(segment.fontSize, segment.kerning))
      || wordArabicTextKerned(segment.text)
      ? 'normal'
      : 'none';
  }

  private restoreKerning(previous: CanvasFontKerning | null): void {
    if (previous != null) {
      this.context.fontKerning = previous;
      this.selectedKerning = previous;
    }
  }

  withSegmentKerning<T>(segment: LayoutTextSeg, measure: () => T): T {
    const previous = this.selectSegmentKerning(segment);
    try {
      return measure();
    } finally {
      this.restoreKerning(previous);
    }
  }

  measureSegment(segment: LayoutTextSeg, clusterGeometry: boolean | 'spaces' = false): TextMetrics {
    if (segment.textLayoutService && segment.textShapeRequest) {
      if (segment.textShapeRequest.text !== segment.text) {
        throw new Error('Segment measurement does not match its retained text range context');
      }
      const shaped = segment.textLayoutService.shape({
        ...segment.textShapeRequest,
        fontSizePt: calcEffectiveFontPx(segment, this.scale),
        measure: true,
        clusterGeometry,
      });
      if (clusterGeometry) {
        if (clusterGeometry === 'spaces') segment.shapedSpaceClusters = shaped.clusters;
        else segment.shapedClusters = shaped.clusters;
        segment.selectedFaceFontBox = {
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
        segment.selectedFaceInkBounds = shaped.inkBounds ?? {
          xMinPt: 0,
          xMaxPt: shaped.advancePt,
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
      }
      return {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    }
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return this.context.measureText(segment.text);
    } finally {
      this.restoreKerning(previous);
    }
  }

  /** Browser pair-context repair at an actual ordinary word boundary. This
   * is native geometry, not an Office fitting allowance (§17.3.2.19).
   * Same-source, same-face horizontal non-complex text is measured as two
   * adjacent tokens and their concatenation; the difference belongs to the
   * next token's origin/advance. Each token is visited at most twice, without
   * a line-prefix cache. This does not promise arbitrary multi-token contextual
   * GSUB equivalence; RTL/complex and authored atomic units retain their own
   * shaping/placement contracts. Callers commit the result only on that line.
   * Registered physical units retain their ordinary slot metadata. A joined
   * boundary probe may carry the same uniform-allocation proof, but the service
   * must again select ONE exact registered face with all-or-none peer cmap.
   * Otherwise decline rather than charging a detached mark or estimating a
   * pair. Latin probes re-prove one ordinary WORD slot seam plus an optional
   * pure trailing U+0020 span; two individually admitted tokens can still
   * have too many joined word seams. All spaces retain full face/cmap proof.
   * This bounded two-token probe does not merge retained paint units.
   */
  wordBoundaryAdvance(left: LayoutTextSeg | undefined, right: LayoutTextSeg): number {
    const l = left?.textShapeRequest;
    const r = right.textShapeRequest;
    const service = right.textLayoutService;
    const physicalUnit = hasLiveSlotSeam(left) || hasLiveSlotSeam(right);
    if (!left || !l || !r || !service || service !== left.textLayoutService
      || !left.text.endsWith(' ') || right.text.startsWith(' ') || !right.text
      || left.metricOnly || right.metricOnly || left.ruby || right.ruby
      || left.fitTextRegionIndex !== undefined || right.fitTextRegionIndex !== undefined
      || left.verticalRun || right.verticalRun || left.rtl || right.rtl
      || l.complexScript || r.complexScript || l.kerning !== true
      || (left.script !== 'ascii' && left.script !== 'highAnsi')
      || (right.script !== 'ascii' && right.script !== 'highAnsi')
      || (!physicalUnit && left.script !== right.script)
      || left.fontRoute?.fingerprint !== right.fontRoute?.fingerprint
      || calcEffectiveFontPx(left, this.scale) !== calcEffectiveFontPx(right, this.scale)
      || l.weight !== r.weight || l.style !== r.style || l.kerning !== r.kerning
      || charScaleFactor(left) !== charScaleFactor(right)
      || left.sourceRunIndex === undefined || left.sourceRunIndex !== right.sourceRunIndex
      || left.sourceTextSequence !== right.sourceTextSequence) return 0;
    const lc = l.substituteContext;
    const rc = r.substituteContext;
    if (!lc || !rc || lc.text !== rc.text || lc.offset + left.text.length !== rc.offset) return 0;
    const measure = (request: typeof r) => service.shape({ ...request,
      fontSizePt: calcEffectiveFontPx(right, this.scale), measure: true, clusterGeometry: false });
    const joined = measure({ ...l, text: left.text + right.text,
      registeredLatinPieceText: undefined,
      ...(physicalUnit ? { joinRegisteredLatinSlots: true } : {}),
    });
    if (physicalUnit && !(joined.spans.length === 1 && joined.spans[0]?.semanticSlotSpans)) return 0;
    return (joined.advancePt - measure(l).advancePt - measure(r).advancePt) * charScaleFactor(right);
  }

  /** WORD_KERNED_SPACE_PAIRS: the font pair adjustment across a seam between
   * two adjacent text pieces of one line when exactly one side of the seam is
   * U+0020. A word-piece seam (`left` ends with its separator) of one slot
   * script uses {@link wordBoundaryAdvance}. A §17.3.2.26 slot seam inside a
   * piece (a highAnsi glyph before an ascii space, or a space before a glyph
   * of the other Latin slot) cannot be joined by the text service, which
   * shapes each slot alone; both sides are then proved to be one face (same
   * route fingerprint, size, weight, style, scale and source sequence) and the
   * pair is measured from that face as two glyphs against their isolated
   * advances, with kerning on. Any other seam has no pair here. */
  kernedSpaceSeamAdvance(left: LayoutTextSeg | undefined, right: LayoutTextSeg): number {
    if (!left || !left.text || !right.text) return 0;
    const leftSpace = left.text.endsWith(' ');
    if (leftSpace === right.text.startsWith(' ')) return 0;
    if (leftSpace && (left.script === right.script || hasLiveSlotSeam(left) || hasLiveSlotSeam(right))) {
      return this.wordBoundaryAdvance(left, right);
    }
    const l = left.textShapeRequest;
    const r = right.textShapeRequest;
    if (!l || !r || l.kerning !== true || r.kerning !== true
      || !right.textLayoutService || right.textLayoutService !== left.textLayoutService
      || left.metricOnly || right.metricOnly || left.ruby || right.ruby
      || left.fitTextRegionIndex !== undefined || right.fitTextRegionIndex !== undefined
      || left.verticalRun || right.verticalRun || left.rtl || right.rtl
      || l.complexScript || r.complexScript
      || (left.script !== 'ascii' && left.script !== 'highAnsi')
      || (right.script !== 'ascii' && right.script !== 'highAnsi')
      || !left.fontRoute || left.fontRoute.fingerprint !== right.fontRoute?.fingerprint
      || calcEffectiveFontPx(left, this.scale) !== calcEffectiveFontPx(right, this.scale)
      || left.bold !== right.bold || left.italic !== right.italic
      || l.weight !== r.weight || l.style !== r.style
      || charScaleFactor(left) !== charScaleFactor(right)
      || left.sourceRunIndex === undefined || left.sourceRunIndex !== right.sourceRunIndex
      || left.sourceTextSequence !== right.sourceTextSequence) return 0;
    const lc = l.substituteContext;
    const rc = r.substituteContext;
    if (!lc || !rc || lc.text !== rc.text || lc.offset + left.text.length !== rc.offset) return 0;
    const before = left.text.slice(graphemeClusterOffsets(left.text).at(-1) ?? 0);
    const after = right.text.slice(0, graphemeClusterOffsets(right.text)[0] ?? right.text.length);
    const font = this.fontForSegment(left);
    const previousFont = this.selectedFont ?? this.initialFont;
    const previousKerning = this.context.fontKerning;
    this.setFont(font);
    this.context.fontKerning = 'normal';
    try {
      const width = (text: string) => this.context.measureText(text).width;
      return (width(before + after) - width(before) - width(after)) * charScaleFactor(right);
    } finally {
      this.context.font = previousFont;
      this.selectedFont = previousFont;
      this.context.fontKerning = previousKerning;
    }
  }

  /** WORD_CONDENSED_SPACING_ADVANCE_FLOOR units of `segment.text[start, start +
   * text.length)`, or `undefined` when no character of the segment can reach a
   * negative advance (the signed segment arithmetic is then exact).
   *
   * Vertical runs use the planned glyph cells paint draws. Horizontal text uses
   * each grapheme cluster's contextual advance from {@link contextualGraphemeAdvances},
   * the authority retained geometry shares, so a probe is linear in its length. Whole-segment
   * units are computed once per segment text and reused by its prefixes (a
   * cluster's context never looks forward). Work is charged to `reserveWork`
   * in the UTF-16 units measured. */
  condensedAdvanceUnits(
    segment: LayoutTextSeg,
    text: string,
    start: number,
    grid: DocGridCtx | undefined,
    reserveWork: (utf16Units: number) => void,
  ): CondensedAdvanceUnits | undefined {
    if (!condensedSpacingMayFloor(segment, grid, this.scale)) return undefined;
    if (segment.text.slice(start, start + text.length) !== text) {
      throw new Error('Condensed advance units do not match their segment text range');
    }
    // The segment's adjustment table (built and charged once, then reused by
    // every probe and by segAdvanceWidth).
    const adjustmentPx = condensedCharacterAdjustments(segment, grid, this.scale, reserveWork);
    // Measurement without the vertical capability (intrinsic table widths)
    // already measures vertical text horizontally; keep that authority.
    if (segment.verticalRun && this.verticalGlyphMeasurement) {
      // Upper bound of the planner's measured UTF-16 units: per-character
      // routing probes plus bounded-context sideways graphemes.
      reserveWork(VERTICAL_PLAN_WORK_PER_UNIT * text.length);
      const gridPx = segmentCharacterGridDeltaPx(segment, grid, this.scale);
      const cells = this.verticalGlyphMeasurement.planRun({
        text,
        font: this.fontForSegment(segment),
        fontKerning: this.segmentKerning(segment),
        fontSizePt: calcEffectiveFontPx(segment, this.scale),
        letterSpacingPt: segLetterSpacingPx(segment, grid, this.scale),
        characterGridDeltaPt: gridPx,
        charScale: charScaleFactor(segment),
        growTrRotateInk: true,
        writingMode: 'vertical-rl',
        characterAdjustmentPt: (from: number, to: number) => adjustmentPx(start + from, start + to),
      });
      return {
        kind: 'vertical-cells',
        cells: cells.map((cell) => ({
          text: text.slice(cell.range.start, cell.range.end),
          advancePx: cell.advancePt,
          ...(cell.floored ? { floored: true } : {}),
        })),
      };
    }
    if (start === 0) {
      let whole = this.condensedClusterCache.get(segment);
      if (!whole || whole.text !== segment.text) {
        whole = { text: segment.text, clusters: this.contextualClusters(segment, 0, segment.text.length, reserveWork) };
        this.condensedClusterCache.set(segment, whole);
      }
      let end = 0;
      const clusters: NaturalClusterAdvance[] = [];
      for (const cluster of whole.clusters) {
        if (end + cluster.text.length > text.length) break;
        clusters.push(cluster);
        end += cluster.text.length;
      }
      if (end === text.length) {
        reserveWork(clusters.length);
        return { kind: 'natural-clusters', clusters };
      }
    }
    return {
      kind: 'natural-clusters',
      clusters: this.contextualClusters(segment, start, start + text.length, reserveWork),
    };
  }

  private readonly condensedClusterCache = new WeakMap<
    LayoutTextSeg,
    Readonly<{ text: string; clusters: readonly NaturalClusterAdvance[] }>
  >();

  private contextualClusters(
    segment: LayoutTextSeg,
    start: number,
    end: number,
    reserveWork: (utf16Units: number) => void,
  ): readonly NaturalClusterAdvance[] {
    return contextualGraphemeAdvances(
      segment, start, end, calcEffectiveFontPx(segment, this.scale),
      (text) => this.measureRunText(segment, text).width,
      reserveWork,
    );
  }

  measureRunText(segment: LayoutTextSeg, text: string): TextMetrics {
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return this.context.measureText(text);
    } finally {
      this.restoreKerning(previous);
    }
  }

  measureCurrentText(text: string): TextMetrics {
    return this.context.measureText(text);
  }

  measureWithFont(font: string, text: string): TextMetrics {
    const previousFont = this.selectedFont ?? this.initialFont;
    this.setFont(font);
    try {
      return this.context.measureText(text);
    } finally {
      this.context.font = previousFont;
      this.selectedFont = previousFont;
    }
  }

  verticalInkExtra(segment: LayoutTextSeg, text: string): number {
    if (!segment.verticalRun) return 0;
    if (!this.verticalGlyphMeasurement) {
      throw new Error('Vertical glyph measurement capability is required for vertical text');
    }
    this.selectSegmentFont(segment);
    const previous = this.selectSegmentKerning(segment);
    try {
      return verticalRunInkExtra(text, true, this.verticalGlyphMeasurement);
    } finally {
      this.restoreKerning(previous);
    }
  }
}
