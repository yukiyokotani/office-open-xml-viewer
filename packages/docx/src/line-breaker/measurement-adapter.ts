import { wordKerningApplies } from '../layout/line-compatibility.js';
import type { LayoutTextSeg } from '../line-layout.js';
import type { MeasurementTextContext, VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx, semanticSlotStartIndex } from '../layout/text.js';
import { charScaleFactor } from './advance.js';
import { verticalRunInkExtra } from './vertical-text.js';

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
    // The acquisition request owns the mode-scoped decision; retained paint
    // consumes that same value. Standalone segments have no compatibility mode.
    const selected = (segment.textShapeRequest?.kerning ?? wordKerningApplies(segment.fontSize, segment.kerning))
      ? 'normal'
      : 'none';
    this.context.fontKerning = selected;
    this.selectedKerning = selected;
    return previous;
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
      ...(physicalUnit ? { joinRegisteredLatinSlots: true } : {}),
    });
    if (physicalUnit && !(joined.spans.length === 1 && joined.spans[0]?.semanticSlotSpans)) return 0;
    return (joined.advancePt - measure(l).advancePt - measure(r).advancePt) * charScaleFactor(right);
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
