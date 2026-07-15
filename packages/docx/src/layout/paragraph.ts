import { autoContrastColor, createCanvasFontRoute } from '@silurus/ooxml-core';
import type { ParagraphLayoutContext } from '../layout-context.js';
import {
  createFloatWrapOracle,
  measureParagraph,
  type MeasuredParagraph,
  type ParagraphMeasurementEnvironment,
  type ParagraphPlacement as MeasurementPlacement,
  type TextMeasurer,
} from '../paragraph-measure.js';
import type {
  LayoutImageSeg,
  LayoutMathSeg,
  LayoutTabSeg,
  LayoutTextSeg,
} from '../line-layout.js';
import { calcEffectiveFontPx, EAST_ASIAN_RE, shapeRunToDocRun } from './text.js';
import type { DocParagraph, DocRun, ShapeRun } from '../types.js';
import {
  computeLineVisualOrder,
  jcIsFullyJustified,
  jcStretchesLastLine,
  resolveAlignEdge,
  segmentsHaveRtl,
} from '../bidi-line.js';
import {
  distributeLineSlack,
  distributedDelta,
  shrinkFitCompression,
  type DistributeResult,
  type SegStretch,
} from '../text-distribute.js';
import { computeKashidaDistribution, type KashidaLevel } from '../kashida-justify.js';
import { imageResourceKey } from './source-key.js';
import { stableFingerprint } from './fingerprint.js';
import { planShapeDrawing } from './shape-drawing-plan.js';
import {
  normalizeTextBoxInput,
  type NormalizedTextBoxParagraphInput,
} from './textbox-input.js';
import {
  numberingMarkerPhysicalLeft,
  resolveNumberingMarkerGeometry,
  shapeNumberingMarkerText,
} from './numbering-marker.js';
import { deepFreezePlainData, snapshotPlainData } from './plain-data.js';
import { retainedBorderTreatment } from './border-treatment.js';
import {
  translateDrawing,
  translateLine,
  translateParagraphLayout,
  translatePlacement,
  translatePoint,
  translateRect,
  translateTextBox,
} from './retained-geometry-translation.js';
export { translateParagraphLayout } from './retained-geometry-translation.js';
import type { ParagraphBorderEdges } from './paragraph-border-adjacency.js';
import {
  centeredLeaderGlyphOrigins,
  groupedRunBorderFragments,
  retainedEmphasisGlyphs,
  retainedTextDecorations,
  rubyPaintOperations,
  type RetainedEmphasisClusterInk,
  type RetainedEmphasisMarkInput,
} from './retained-typography.js';
import type { RunTypographyAcquisitionInput } from './typography-input.js';
import { resolveAnchorFrame, type AnchorReferenceFramesInput, type AnchorFrameResult } from './anchor-frame.js';
import type { AnchorAcquisitionInput, AnchorAxisChoiceInput, AnchorEdgesInput } from './anchor-input.js';
import { resizeResolvedAnchorGeometry } from './anchor-derived-geometry.js';
import { paragraphGapPt } from './paragraph-spacing.js';
import { paginationFieldDependency } from './pagination-fields.js';
import {
  measureParagraphIntrinsicWidth,
  type BodyFrameGroup,
} from './frame.js';
export {
  bodyFrameGroupFor,
  bodyParagraphBorderEdgesFor,
  collectBodyFrameGroups,
  prepareBodyFrameMetadata,
} from './frame.js';
export type { BodyFrameGroup } from './frame.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type {
  DrawingLayout,
  DrawingPaintCommand,
  AcquiredParagraphLayoutInput,
  InlineResourceLayout,
  LineLayout,
  LayoutRect,
  ParagraphLayout,
  ParagraphPlacement,
  PointPt,
  SourceRef,
  TextBoxLayout,
  TextPlacement,
  WrapExclusion,
} from './types.js';

function finiteNonNegative(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be finite and non-negative`);
  }
  return value;
}

export type MeasuredTextPlanSegment = Readonly<
  Omit<TextPlacement, 'origin' | 'bounds' | 'advancePt' | 'paintOps'> & {
    measuredWidthPt: number;
    basePaintOps: readonly import('./types.js').TextPaintOp[];
    /** False when this segment continues the preceding shaped grapheme. */
    breakBefore?: boolean;
    /** WordprocessingML bidi classification facts consumed by the shared UAX#9 seam. */
    rtl?: boolean;
    digitsAsAN?: boolean;
    /** A fixed-pitch fitText region is an atom for paragraph justification. */
    fixedPitch?: boolean;
    /** Acquisition-only authority used to shape the final contextual kashida string. */
    textLayoutService?: import('./text.js').TextLayoutService;
    textShapeRequest?: import('./text.js').TextShapeRequest;
    retainedGeometry?: RetainedTextGeometryPlan;
  }
>;

type RetainedInkMetric = Pick<
  import('./text.js').GlyphMeasurement,
  'ascentPt' | 'descentPt' | 'inkBounds'
>;

interface RetainedTextGeometryPlan {
  readonly base: RetainedInkMetric;
  readonly underline?: Readonly<{
    authoredStyle?: string;
    color: string;
    probe: RetainedInkMetric;
  }>;
  readonly strike?: Readonly<{
    double: boolean;
    probe: RetainedInkMetric;
    doubleProbe?: RetainedInkMetric;
  }>;
  readonly emphasis?: Readonly<{
    authored: string;
    glyph: string;
    mark: RetainedEmphasisMarkInput;
    clusterInk: readonly RetainedEmphasisClusterInk[];
  }>;
}

function retainedTypographyInput(run: DocRun | undefined): RunTypographyAcquisitionInput | undefined {
  if (!run || (run.type !== 'text' && run.type !== 'field')) return undefined;
  return (run as typeof run & Readonly<{
    typographyInput?: RunTypographyAcquisitionInput;
  }>).typographyInput;
}

export interface MeasuredTabPlanSegment {
  readonly kind: 'tab';
  readonly range: import('./types.js').TextRange;
  readonly measuredWidthPt: number;
  readonly leader: import('./types.js').TabPlacement['leader'];
  readonly fontSizePt: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly leaderShape?: Readonly<{
    glyph: string;
    advancePt: number;
    fontRoute: import('@silurus/ooxml-core').CanvasFontRoute;
    fontSizePt: number;
    fontWeight: number;
    fontStyle: 'normal' | 'italic';
    color: import('./types.js').TextColorPolicy;
  }>;
}

export interface MeasuredResourcePlanSegment {
  readonly kind: 'resource';
  readonly range: import('./types.js').TextRange;
  readonly measuredWidthPt: number;
  readonly resourceKey: string;
  readonly resourceKind: import('./types.js').InlineResourceKind;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly topOffsetPt: number;
}

export interface MeasuredAnchorHostPlanSegment {
  readonly kind: 'anchor-host';
  readonly measuredWidthPt: 0;
  readonly range: import('./types.js').TextRange;
  readonly sourceMetrics?: Readonly<{ ascentPt: number; descentPt: number }>;
  readonly anchorOccurrenceId?: string;
}

export type MeasuredLinePlanSegment =
  | MeasuredTextPlanSegment
  | MeasuredTabPlanSegment
  | MeasuredResourcePlanSegment
  | MeasuredAnchorHostPlanSegment;

export interface MeasuredLinePlanInput {
  readonly range: import('./types.js').TextRange;
  readonly topPt: number;
  readonly baselinePt: number;
  readonly advancePt: number;
  readonly xOffsetPt: number;
  readonly availableWidthPt: number;
  readonly endsWithBreak: boolean;
  readonly segments: readonly MeasuredLinePlanSegment[];
}

export interface PlanLineInput {
  readonly paragraphXPt: number;
  readonly availableWidthPt: number;
  readonly alignment?: string;
  readonly baseRtl: boolean;
  readonly isFirstLine: boolean;
  readonly isLastLine: boolean;
  readonly stretchLastLine: boolean;
  readonly firstLineIndentPt?: number;
  readonly numbering?: Readonly<{
    /** Resolved logical-start offset of the first-line body after the marker. */
    bodyOffsetPt: number;
  }>;
  /** Decimal stop relative to paragraphXPt for Word's numeric no-tab alignment. */
  readonly decimalAutoTabPt?: number;
  /** Effective m:jc for a one-display-math line. Absolute, never bidi-flipped. */
  readonly displayMathJustification?: string;
  readonly line: MeasuredLinePlanInput;
}

function displayMathEdge(justification: string): 'left' | 'right' | 'center' {
  switch (justification) {
    case 'left': return 'left';
    case 'right': return 'right';
    case 'center':
    case 'centerGroup':
    default: return 'center';
  }
}

function segmentWidth(segment: MeasuredLinePlanSegment): number {
  return finiteNonNegative(segment.measuredWidthPt, 'segment.measuredWidthPt');
}

function distributionSegments(segments: readonly MeasuredLinePlanSegment[]): readonly { text?: string }[] {
  return segments.map((segment) => segment.kind === 'text' && !segment.fixedPitch
    ? { text: segment.text }
    : {});
}

function kashidaLevel(alignment: string | undefined): KashidaLevel | null {
  if (alignment === 'lowKashida') return 'low';
  if (alignment === 'mediumKashida') return 'medium';
  if (alignment === 'highKashida') return 'high';
  return null;
}

function contextualAdvance(segment: MeasuredTextPlanSegment, text: string): number {
  if (!segment.textLayoutService || !segment.textShapeRequest) {
    throw new Error('Kashida acquisition requires the retained TextLayoutService authority');
  }
  const shaped = segment.textLayoutService.shape({
    ...segment.textShapeRequest,
    text,
    measure: true,
  });
  const scaleX = segment.basePaintOps[0]?.scaleX ?? 1;
  const pitchPt = segment.basePaintOps[0]?.letterSpacingPt ?? 0;
  return shaped.advancePt * scaleX + [...text].length * pitchPt;
}

function keepGraphemeSafeCuts(
  distribution: DistributeResult | null,
  segments: readonly MeasuredLinePlanSegment[],
): DistributeResult | null {
  if (!distribution) return null;
  const totalDeltaPt = distributedDelta(distribution);
  const retained = new Map<number, SegStretch>();
  let gapCount = 0;
  for (const [segmentIndex, stretch] of distribution.perSeg) {
    const segment = segments[segmentIndex];
    let splitBefore = stretch.splitBefore;
    if (segment?.kind === 'text') {
      const allowed = new Set(segment.clusters.slice(1).map((cluster) =>
        cluster.range.start - segment.range.start));
      const codePoints = [...segment.text];
      const utf16Offsets = [0];
      for (const codePoint of codePoints) {
        utf16Offsets.push((utf16Offsets.at(-1) ?? 0) + codePoint.length);
      }
      splitBefore = splitBefore.filter((cut) => allowed.has(utf16Offsets[cut] ?? -1));
    }
    const next = segments[segmentIndex + 1];
    const trailingGap = stretch.trailingGap
      && !(next?.kind === 'text' && next.breakBefore === false);
    gapCount += splitBefore.length + (trailingGap ? 1 : 0);
    retained.set(segmentIndex, {
      splitBefore: [...splitBefore],
      trailingGap,
      internalStretch: 0,
    });
  }
  if (gapCount === 0) return null;
  const perGap = totalDeltaPt / gapCount;
  for (const stretch of retained.values()) {
    stretch.internalStretch = stretch.splitBefore.length * perGap;
  }
  return { perGap, perSeg: retained };
}

function retainedTextGeometry(
  segment: MeasuredTextPlanSegment,
  stretch: SegStretch | undefined,
  perGapPt: number,
): Readonly<{
  clusters: readonly import('./types.js').TextClusterLayout[];
  paintOps: readonly import('./types.js').TextPaintOp[];
}> {
  if (!stretch || stretch.splitBefore.length === 0) {
    return { clusters: segment.clusters, paintOps: segment.basePaintOps };
  }
  const codePoints = [...segment.text];
  const cuts = [...stretch.splitBefore];
  if (cuts.some((cut, index) => cut <= 0 || cut >= codePoints.length || (index > 0 && cut <= (cuts[index - 1] ?? 0)))) {
    throw new Error('Internal paragraph justification contains an invalid code-point cut');
  }
  const utf16Offsets = [0];
  for (const codePoint of codePoints) {
    utf16Offsets.push((utf16Offsets.at(-1) ?? 0) + codePoint.length);
  }
  const cutUtf16 = cuts.map((cut) => utf16Offsets[cut] ?? -1);
  const clusterStarts = new Set(segment.clusters.map((cluster) =>
    cluster.range.start - segment.range.start));
  if (cutUtf16.some((cut) => !clusterStarts.has(cut))) {
    throw new Error('Internal paragraph justification must split at shaped cluster boundaries');
  }
  const boundaries = [0, ...cuts, codePoints.length];
  const paintSlices: Array<Readonly<{
    range: import('./types.js').TextRange;
    offset: import('./types.js').PointPt;
  }>> = [];
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const from = boundaries[index] ?? 0;
    const to = boundaries[index + 1] ?? from;
    const start = segment.range.start + (utf16Offsets[from] ?? 0);
    const firstCluster = segment.clusters.find((cluster) => cluster.range.start === start);
    if (!firstCluster) throw new Error('Internal paragraph justification is missing shaped cluster geometry');
    paintSlices.push({
      range: { start, end: segment.range.start + (utf16Offsets[to] ?? 0) },
      offset: { xPt: firstCluster.offset.xPt + index * perGapPt, yPt: firstCluster.offset.yPt },
    });
  }
  const clusters = segment.clusters.map((cluster) => {
    const relativeStart = cluster.range.start - segment.range.start;
    const precedingGaps = cutUtf16.filter((cut) => cut <= relativeStart).length;
    return {
      ...cluster,
      offset: { ...cluster.offset, xPt: cluster.offset.xPt + precedingGaps * perGapPt },
    };
  });
  const baseOp = segment.basePaintOps.length === 1 ? segment.basePaintOps[0] : undefined;
  if (!baseOp) throw new Error('Internal paragraph justification requires one contextual paint op');
  const fullyDistributed = cuts.length === codePoints.length - 1
    && cuts.every((cut, index) => cut === index + 1);
  if (fullyDistributed) {
    // Canvas applies uniform letter spacing without breaking the contextual
    // shaping unit. Keeping one op is essential for Japanese punctuation whose
    // isolated advance/ink differs from its `…：［…` context.
    return {
      clusters,
      paintOps: [{
        ...baseOp,
        letterSpacingPt: baseOp.letterSpacingPt + perGapPt,
      }],
    };
  }
  const paintOps: import('./types.js').TextPaintOp[] = paintSlices.map((slice) => ({
    ...baseOp,
    text: segment.text.slice(
      slice.range.start - segment.range.start,
      slice.range.end - segment.range.start,
    ),
    range: slice.range,
    offset: slice.offset,
  }));
  return { clusters, paintOps };
}

/** Converts a measured line snapshot into final point-space visual geometry.
 * Bidi order, alignment, compression, justification, and tab advances are
 * resolved here; paint consumes the resulting placements without source access. */
export function planLine(input: PlanLineInput): LineLayout {
  const { line } = input;
  let segments = line.segments;
  const bidi = input.baseRtl || segmentsHaveRtl(segments);
  const visual = computeLineVisualOrder(
    segments.map((segment) => segment.kind === 'tab'
      ? { isTab: true }
      : segment.kind === 'text'
        ? { text: segment.text, rtl: segment.rtl, digitsAsAN: segment.digitsAsAN }
        : {}),
    input.baseRtl,
  );
  let naturalWidthPt = segments.reduce((sum, segment) => sum + segmentWidth(segment), 0);
  const lineLeftPt = input.paragraphXPt + line.xOffsetPt;
  const availableWidthPt = Math.min(input.availableWidthPt, line.availableWidthPt);
  const logicalStartOffsetPt = !input.isFirstLine
    ? 0
    : input.numbering
      ? finiteNonNegative(input.numbering.bodyOffsetPt, 'numbering.bodyOffsetPt')
      : input.firstLineIndentPt ?? 0;
  const physicalStartOffsetPt = input.baseRtl ? 0 : logicalStartOffsetPt;
  const effectiveAvailableWidthPt = input.baseRtl
    ? availableWidthPt - logicalStartOffsetPt
    : availableWidthPt;
  let lineSlackPt = effectiveAvailableWidthPt - physicalStartOffsetPt - naturalWidthPt;
  const endsLogicalLine = input.isLastLine || line.endsWithBreak;
  const edge = input.displayMathJustification === undefined
    ? resolveAlignEdge(input.alignment, input.baseRtl)
    : displayMathEdge(input.displayMathJustification);
  const applyJustify = edge === 'justify' && (!endsLogicalLine || input.stretchLastLine);
  const kashida = applyJustify ? kashidaLevel(input.alignment) : null;
  if (kashida && lineSlackPt > 0) {
    const distribution = computeKashidaDistribution(
      segments.map((segment) => segment.kind === 'text' ? { text: segment.text } : {}),
      lineSlackPt,
      kashida,
      (segmentIndex, text) => {
        const segment = segments[segmentIndex];
        if (segment?.kind !== 'text') return 0;
        return contextualAdvance(segment, text);
      },
    );
    if (distribution) {
      segments = segments.map((segment, segmentIndex): MeasuredLinePlanSegment => {
        if (segment.kind !== 'text') return segment;
        const plan = distribution.perSeg.get(segmentIndex);
        if (!plan) return segment;
        const base = segment.basePaintOps[0];
        if (!base) throw new Error('Kashida acquisition requires a contextual text paint operation');
        return {
          ...segment,
          measuredWidthPt: segment.measuredWidthPt + plan.advanceDeltaPx,
          basePaintOps: [{ ...base, text: plan.text, sourceMapping: 'kashida' }],
        };
      });
      naturalWidthPt += distribution.appliedPx;
      lineSlackPt = distribution.residualPx;
    }
  }
  const lastDrawnIndex = visual.order.at(-1) ?? -1;
  let firstContentIndex = 0;
  if (!bidi) {
    const found = segments.findIndex((segment) => segment.kind !== 'text' || /\S/.test(segment.text));
    firstContentIndex = found < 0 ? 0 : found;
  }

  let stretchByIndex: ReadonlyMap<number, SegStretch> | null = null;
  let perGapPt = 0;
  let distributedWidthPt = 0;
  const distSegments = distributionSegments(segments);
  if (applyJustify) {
    const distribution = keepGraphemeSafeCuts(distributeLineSlack(
      distSegments,
      lineSlackPt,
      firstContentIndex,
      bidi ? lastDrawnIndex : segments.length,
      -(line.baselinePt - line.topPt) * .25,
      lineSlackPt > 0,
      input.alignment === 'thaiDistribute' && lineSlackPt > 0,
    ), segments);
    stretchByIndex = distribution?.perSeg ?? null;
    perGapPt = distribution?.perGap ?? 0;
    distributedWidthPt = distributedDelta(distribution);
  } else if (lineSlackPt < 0) {
    const compression = keepGraphemeSafeCuts(shrinkFitCompression(
      distSegments,
      lineSlackPt,
      firstContentIndex,
      bidi ? lastDrawnIndex : segments.length,
      line.baselinePt - line.topPt,
    ), segments);
    stretchByIndex = compression?.perSeg ?? null;
    perGapPt = compression?.perGap ?? 0;
    distributedWidthPt = distributedDelta(compression);
  }

  const drawnWidthPt = naturalWidthPt + distributedWidthPt;
  const alignmentSlackPt = lineSlackPt - distributedWidthPt;
  const naturalAlignmentOffsetPt = edge === 'right'
    ? alignmentSlackPt
    : edge === 'center'
      ? alignmentSlackPt / 2
      : edge === 'justify' && input.baseRtl && !applyJustify
        ? alignmentSlackPt
        : 0;
  const lineStartPt = lineLeftPt + physicalStartOffsetPt;
  const alignmentOffsetPt = input.decimalAutoTabPt === undefined
    ? naturalAlignmentOffsetPt
    : Math.max(0, input.paragraphXPt + input.decimalAutoTabPt - drawnWidthPt - lineStartPt);
  let xPt = lineStartPt + alignmentOffsetPt;
  const placements: ParagraphPlacement[] = [];
  for (const segmentIndex of visual.order) {
    const segment = segments[segmentIndex];
    if (!segment) continue;
    const stretch = stretchByIndex?.get(segmentIndex);
    const internalStretchPt = stretch?.internalStretch ?? 0;
    const widthPt = segmentWidth(segment) + internalStretchPt;
    if (segment.kind === 'tab') {
      const bounds = { xPt, yPt: line.topPt, widthPt: segment.measuredWidthPt, heightPt: line.advancePt };
      placements.push({
        kind: 'tab', range: segment.range,
        bounds,
        advancePt: segment.measuredWidthPt,
        leader: segment.leader,
        ...(segment.leader === 'none' ? {} : segment.leaderShape ? {
          leaderGlyphs: centeredLeaderGlyphOrigins({
            interval: bounds,
            baselinePt: line.baselinePt,
            ...segment.leaderShape,
          }),
        } : {}),
      });
    } else if (segment.kind === 'resource') {
      placements.push({
        kind: 'resource', range: segment.range,
        resourceKey: segment.resourceKey, resourceKind: segment.resourceKind,
        bounds: {
          xPt, yPt: line.baselinePt + segment.topOffsetPt,
          widthPt: segment.widthPt, heightPt: segment.heightPt,
        },
        advancePt: segment.measuredWidthPt,
      });
    } else if (segment.kind === 'anchor-host') {
      placements.push({
        kind: 'anchor-host', range: segment.range,
        bounds: { xPt, yPt: line.topPt, widthPt: 0, heightPt: line.advancePt },
        baselinePt: line.baselinePt,
        ...(segment.sourceMetrics ? { sourceMetrics: segment.sourceMetrics } : {}),
        ...(segment.anchorOccurrenceId ? { anchorOccurrenceId: segment.anchorOccurrenceId } : {}),
      });
    } else {
      const {
        measuredWidthPt: _measuredWidthPt,
        breakBefore: _breakBefore,
        rtl: _rtl,
        digitsAsAN: _digitsAsAN,
        fixedPitch: _fixedPitch,
        textLayoutService: _textLayoutService,
        textShapeRequest: _textShapeRequest,
        retainedGeometry,
        direction: _direction,
        ...style
      } = segment;
      const textGeometry = retainedTextGeometry(segment, stretch, perGapPt);
      const direction = visual.rtl[segmentIndex] ? 'rtl' : 'ltr';
      const paintOps = direction === 'rtl'
        ? textGeometry.paintOps.map((operation) => {
            const text = operation.text.trimEnd();
            return text === '' ? operation : {
              ...operation,
              text,
              ...(operation.sourceMapping === 'kashida' ? {} : {
                range: { ...operation.range, end: operation.range.start + text.length },
              }),
            };
          })
        : textGeometry.paintOps;
      const trailingWhitespaceStart = segment.text.trimEnd().length;
      const rtlLeadingGapPt = direction === 'rtl'
        ? (style.fitText?.trailingPadPt ?? 0) + segment.clusters
            .filter((cluster) => cluster.range.start >= segment.range.start + trailingWhitespaceStart)
            .reduce((sum, cluster) => sum + cluster.advancePt, 0)
        : 0;
      const ownedTrailingSlackPt = stretch?.trailingGap ? perGapPt : 0;
      const origin = { xPt: xPt + rtlLeadingGapPt, yPt: line.baselinePt };
      const geometryOrigin = {
        xPt,
        yPt: line.baselinePt - (style.positionPt ?? 0),
      };
      const decorations = retainedGeometry
        ? retainedTextDecorations({
            origin: geometryOrigin,
            advancePt: widthPt + ownedTrailingSlackPt,
            base: retainedGeometry.base,
            color: retainedColorString(style.color),
            ...(retainedGeometry.underline ? { underline: retainedGeometry.underline } : {}),
            ...(retainedGeometry.strike ? { strike: retainedGeometry.strike } : {}),
          })
        : style.decorations;
      const emphasis = retainedGeometry?.emphasis ? {
        authored: retainedGeometry.emphasis.authored,
        glyphs: retainedEmphasisGlyphs({
          authored: retainedGeometry.emphasis.authored,
          glyph: retainedGeometry.emphasis.glyph,
          origin: {
            xPt: origin.xPt,
            yPt: line.baselinePt - (style.positionPt ?? 0),
          },
          clusters: textGeometry.clusters,
          clusterInk: retainedGeometry.emphasis.clusterInk,
          mark: retainedGeometry.emphasis.mark,
          scaleX: segment.basePaintOps[0]?.scaleX ?? 1,
        }),
      } : undefined;
      const placed: TextPlacement = {
        ...style,
        kind: 'text',
        origin,
        bounds: { xPt, yPt: line.topPt, widthPt, heightPt: line.advancePt },
        advancePt: widthPt,
        clusters: textGeometry.clusters,
        paintOps: paintOps.map((operation) => ({ ...operation, direction })),
        decorations,
        ...(emphasis ? { emphasis } : {}),
        direction,
        ...(ownedTrailingSlackPt !== 0 ? { ownedTrailingSlackPt } : {}),
        ...((style.highlight || style.background) ? {
          highlightFragments: [{
            rect: {
              xPt, yPt: line.topPt,
              widthPt: widthPt + ownedTrailingSlackPt,
              heightPt: line.advancePt,
            },
            color: style.highlight ?? style.background!,
          }],
        } : {}),
        ...(style.ruby ? {
          ruby: {
            ...style.ruby,
            paintOps: style.ruby.paintOps.map((operation) => ({
              ...operation,
              origin: {
                xPt: operation.origin.xPt + xPt
                  + (widthPt - segment.measuredWidthPt) / 2,
                yPt: operation.origin.yPt + line.baselinePt,
              },
            })),
          },
        } : {}),
      };
      placements.push(placed);
    }
    xPt += widthPt;
    if (stretch?.trailingGap) xPt += perGapPt;
  }
  for (let start = 0; start < placements.length;) {
    const first = placements[start];
    if (first?.kind !== 'text' || !first.runBorder) {
      start += 1;
      continue;
    }
    let end = start + 1;
    while (end < placements.length) {
      const candidate = placements[end];
      if (candidate?.kind !== 'text' || !candidate.runBorder) break;
      end += 1;
    }
    const group = placements.slice(start, end) as TextPlacement[];
    const fragments = groupedRunBorderFragments(group.map((placement) => ({
      bounds: placement.bounds,
      trailingSlackPt: placement.ownedTrailingSlackPt ?? 0,
      border: placement.runBorder!,
    })));
    placements[start] = { ...first, runBorderFragments: fragments };
    start = end;
  }
  return deepFreezePlainData({
    range: line.range,
    bounds: {
      xPt: lineStartPt + alignmentOffsetPt,
      yPt: line.topPt,
      widthPt: drawnWidthPt,
      heightPt: line.advancePt,
    },
    baselinePt: line.baselinePt,
    advancePt: line.advancePt,
    placements,
  });
}

function sliceAdvance(input: AcquiredParagraphLayoutInput): number {
  const continuation = input.continuation;
  const start = continuation?.lineStart ?? 0;
  const end = continuation?.lineEnd ?? input.lines.length;
  if (start < 0 || end < start || end > input.lines.length) {
    throw new RangeError('Paragraph continuation line range is outside the retained lines');
  }
  let advancePt = continuation?.continuesFromPrevious ? 0 : input.spacing.beforePt;
  for (let index = start; index < end; index += 1) {
    const line = input.lines[index];
    if (!line) continue;
    if (index === 0 && !continuation?.continuesFromPrevious) {
      advancePt += Math.max(0,
        line.bounds.yPt - (input.flowBounds.yPt + input.spacing.beforePt));
    } else if (index > start) {
      const previous = input.lines[index - 1];
      advancePt += Math.max(0,
        line.bounds.yPt - ((previous?.bounds.yPt ?? line.bounds.yPt) + (previous?.advancePt ?? 0)));
    }
    advancePt += finiteNonNegative(line.advancePt, 'line.advancePt');
  }
  if (input.lines.length === 0 && input.paragraphMark) {
    advancePt += finiteNonNegative(input.paragraphMark.bounds.heightPt, 'paragraphMark.heightPt');
  }
  if (!continuation?.continuesOnNext) advancePt += input.spacing.afterPt;
  return advancePt;
}

/**
 * Finalizes the parser-independent paragraph acquisition snapshot. All coordinates
 * are scale-1 points; subsequent Canvas paint is a pure viewport transform.
 */
export function layoutParagraph(input: AcquiredParagraphLayoutInput): ParagraphLayout {
  const lineStart = input.continuation?.lineStart ?? 0;
  const lineEnd = input.continuation?.lineEnd ?? input.lines.length;
  const lines = input.lines.slice(lineStart, lineEnd);
  const advancePt = input.continuation
    ? sliceAdvance(input)
    : finiteNonNegative(input.flowBounds.heightPt, 'flowBounds.heightPt');
  const node: ParagraphLayout = {
    kind: 'paragraph',
    id: input.id,
    source: input.source,
    flowDomainId: input.flowDomainId,
    ordinaryFlow: input.ordinaryFlow,
    ...(input.styleId !== undefined ? { styleId: input.styleId } : {}),
    flowBounds: { ...input.flowBounds, heightPt: advancePt },
    inkBounds: input.inkBounds,
    ...(input.clipBounds ? { clipBounds: input.clipBounds } : {}),
    advancePt,
    spacing: input.spacing,
    contextualSpacing: input.contextualSpacing ?? false,
    lines,
    borders: input.borders,
    ...(input.shading ? { shading: input.shading } : {}),
    resources: input.resources,
    drawings: input.drawings,
    textBoxes: input.textBoxes,
    events: input.events,
    exclusions: input.exclusions,
    ...(input.anchorFrames ? { anchorFrames: input.anchorFrames } : {}),
    ...(input.paragraphMark ? { paragraphMark: input.paragraphMark } : {}),
    ...(input.continuation ? { continuation: input.continuation } : {}),
  };
  return deepFreezePlainData(node);
}

export interface ParagraphAcquisitionOptions {
  readonly id: string;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly ordinaryFlow: boolean;
  readonly context: ParagraphLayoutContext;
  readonly placement: MeasurementPlacement;
  readonly measurer: TextMeasurer;
  readonly environment: ParagraphMeasurementEnvironment;
  readonly exclusions: readonly WrapExclusion[];
  /** Effective enclosing fill, retained only for automatic text-color resolution. */
  readonly containerShading?: string | null;
  /** Layout-owned §17.3.1.7 edge selection for adjacent/sliced border boxes. */
  readonly paragraphBorderEdges?: ParagraphBorderEdges;
  /** Final flow reservation; may exceed w:after when a bottom border owns more space. */
  readonly trailingExtentPt?: number;
  /** The measurement starts after a consumed line boundary on another flow slice. */
  readonly continuesFromPrevious?: boolean;
  readonly anchorFrames?: Readonly<Pick<
    AnchorReferenceFramesInput,
    'page' | 'margin' | 'column' | 'pageParity'
  >>;
}

function runSource(source: SourceRef, runIndex: number): SourceRef {
  return { ...source, path: [...source.path, runIndex] };
}

function chartResourceKey(source: SourceRef): string {
  return stableFingerprint('chart-resource', source);
}

function fieldDependency(run: Extract<DocRun, { type: 'field' }>): TextPlacement['dependency'] {
  const paginationDependency = paginationFieldDependency(run);
  if (paginationDependency) return paginationDependency;
  if (/^date$/i.test(run.fieldType)) return 'date';
  if (/^time$/i.test(run.fieldType)) return 'time';
  return 'document';
}

function sourceRunIndex(segment: { sourceRunIndex?: number }): number | undefined {
  return segment.sourceRunIndex;
}

function selectedFaceSourceMetrics(
  segment: LayoutTextSeg,
): Readonly<{ ascentPt: number; descentPt: number }> | undefined {
  if (!segment.textLayoutService || !segment.textShapeRequest) return undefined;
  const shape = segment.textLayoutService.shape({
    ...segment.textShapeRequest,
    text: segment.text,
    measure: true,
  });
  return { ascentPt: shape.ascentPt, descentPt: shape.descentPt };
}

const HIGHLIGHT_COLOR_HEX: Readonly<Record<string, string>> = Object.freeze({
  yellow: '#FFFF00', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
});

function retainedHighlightColor(value: string): string {
  if (value.startsWith('#')) return value;
  return HIGHLIGHT_COLOR_HEX[value] ?? '#FFFF00';
}

function textPlacement(
  segment: LayoutTextSeg,
  paragraph: DocParagraph,
  sourceOffset: number,
  xPt: number,
  baselinePt: number,
  topPt: number,
  heightPt: number,
): TextPlacement | import('./types.js').AnchorHostPlacement {
  const runIndex = sourceRunIndex(segment);
  const run = runIndex === undefined ? undefined : paragraph.runs[runIndex];
  const typography = retainedTypographyInput(run);
  if (segment.metricOnly) {
    const sourceMetrics = selectedFaceSourceMetrics(segment);
    return {
      kind: 'anchor-host',
      range: { start: sourceOffset, end: sourceOffset },
      bounds: { xPt, yPt: topPt, widthPt: 0, heightPt },
      baselinePt,
      ...(sourceMetrics ? { sourceMetrics } : {}),
    };
  }
  const color: TextPlacement['color'] = segment.color
    ? { kind: 'explicit', color: `#${segment.color}` }
    : segment.colorAuto
      ? { kind: 'auto', ...(segment.background ? { background: `#${segment.background}` } : {}) }
      : { kind: 'default' };
  const fontRoute = segment.fontRoute ?? createCanvasFontRoute(
    segment.fontFamily ? `"${segment.fontFamily.replaceAll('"', '\\"')}"` : 'sans-serif',
    segment.fontFamily ? 'native' : 'generic',
  );
  const baseShape = segment.ruby && segment.textLayoutService && segment.textShapeRequest
    ? segment.textLayoutService.shape({
        ...segment.textShapeRequest,
        text: segment.text,
        measure: true,
      })
    : undefined;
  const rubyShape = segment.ruby && segment.textLayoutService && segment.textShapeRequest
    ? segment.textLayoutService.shape({
        ...segment.textShapeRequest,
        text: segment.ruby.text,
        fontSizePt: segment.ruby.fontSizePt,
        measure: true,
      })
    : undefined;
  const rubySpans = segment.ruby && rubyShape
    ? (rubyShape.clusters ?? []).map((cluster) => {
        const span = rubyShape.spans.find((candidate) =>
          candidate.start <= cluster.range.start && candidate.end >= cluster.range.end)
          ?? rubyShape.spans[0];
        if (!span) throw new Error('Ruby shaping produced no selected-face span');
        return {
          text: segment.ruby!.text.slice(cluster.range.start, cluster.range.end),
          offsetPt: cluster.offsetPt,
          fontRoute: span.fontRoute,
          fontSizePt: segment.ruby!.fontSizePt,
          fontWeight: span.font.weight,
          fontStyle: span.font.style,
          color,
        };
      })
    : [];
  const rubyRaisePt = typography?.ruby?.raisePt.status === 'valid'
    ? typography.ruby.raisePt.value ?? undefined
    : segment.ruby?.hpsRaisePt;
  const rubyPaintOps = segment.ruby && rubyShape
    ? rubyPaintOperations({
        baseOrigin: { xPt: 0, yPt: 0 },
        baseAdvancePt: segment.measuredWidth,
        guideAdvancePt: rubyShape.advancePt,
        ...(rubyRaisePt === undefined ? {} : { raisePt: rubyRaisePt }),
        ...(baseShape?.inkBounds && rubyShape.inkBounds ? {
          baseInkTopPt: -baseShape.inkBounds.ascentPt,
          guideInkBottomFromBaselinePt: rubyShape.inkBounds.descentPt,
        } : {}),
        spans: rubySpans,
      })
    : [];
  return {
    kind: 'text',
    text: segment.text,
    ...(runIndex === undefined ? {} : { sourceRunIndex: runIndex }),
    ...(run?.type === 'field' ? { role: 'field-result' as const, dependency: fieldDependency(run) } : {}),
    range: { start: sourceOffset, end: sourceOffset + segment.text.length },
    origin: { xPt, yPt: baselinePt + (segment.position ? -segment.position : 0) },
    bounds: { xPt, yPt: topPt, widthPt: segment.measuredWidth, heightPt },
    advancePt: segment.measuredWidth,
    clusters: [{
      range: { start: sourceOffset, end: sourceOffset + segment.text.length },
      offset: { xPt: 0, yPt: 0 },
      advancePt: segment.measuredWidth,
    }],
    color,
    fontRoute,
    fontSizePt: calcEffectiveFontPx(segment, 1),
    fontWeight: segment.bold ? 700 : 400,
    fontStyle: segment.italic ? 'italic' : 'normal',
    direction: segment.rtl ? 'rtl' : 'ltr',
    ...(segment.verticalRun ? { writingMode: 'vertical-rl' as const } : {}),
    ...(segment.charSpacing !== undefined ? { characterSpacingPt: segment.charSpacing } : {}),
    ...(segment.charScale !== undefined ? { characterScale: segment.charScale } : {}),
    ...(segment.fitTextRegionIndex !== undefined ? { fitText: {
      regionIndex: segment.fitTextRegionIndex,
      perGapPt: segment.fitTextPerGapPx ?? 0,
      trailingPadPt: segment.fitTextTrailingPadPx ?? 0,
    } } : {}),
    ...(segment.kerning !== undefined ? { kerning: segment.fontSize >= segment.kerning } : {}),
    ...(segment.position !== undefined ? { positionPt: segment.position } : {}),
    ...(segment.vertAlign ? { verticalAlign: segment.vertAlign } : {}),
    ...(segment.tateChuYoko ? { tateChuYoko: true } : {}),
    ...(segment.tateChuYokoCompress ? { tateChuYokoCompress: true } : {}),
    ...(segment.ruby && rubyShape ? { ruby: {
      text: segment.ruby.text,
      advancePt: rubyShape.advancePt,
      authored: {
        ...(typography?.ruby?.align.status === 'valid' && typography.ruby.align.value
          ? { align: typography.ruby.align.value } : {}),
        ...(typography?.ruby?.baseFontSizePt.status === 'valid'
          && typography.ruby.baseFontSizePt.value !== null
          ? { baseFontSizePt: typography.ruby.baseFontSizePt.value } : {}),
        ...(rubyRaisePt === undefined ? {} : { raisePt: rubyRaisePt }),
        ...(typography?.ruby?.language.status === 'valid' && typography.ruby.language.value
          ? { language: typography.ruby.language.value } : {}),
      },
      paintOps: rubyPaintOps,
    } } : {}),
    ...(segment.emphasisMark ? { emphasisMark: segment.emphasisMark } : {}),
    ...(segment.highlight ? {
      highlight: retainedHighlightColor(segment.highlight),
    } : {}),
    ...(segment.background ? { background: `#${segment.background}` } : {}),
    ...(segment.border ? { runBorder: {
      val: typography?.border?.val.value ?? segment.border.style,
      color: segment.border.color ? `#${segment.border.color}` : '#000000',
      widthPt: segment.border.width,
      spacePt: segment.border.space ?? 0,
      ...(typography?.border?.themeColor.value
        ? { themeColor: typography.border.themeColor.value } : {}),
      ...(typography?.border?.themeTint.value
        ? { themeTint: typography.border.themeTint.value } : {}),
      ...(typography?.border?.themeShade.value
        ? { themeShade: typography.border.themeShade.value } : {}),
      ...(typography?.border?.shadow.status === 'valid'
        && typography.border.shadow.value !== null
        ? { shadow: typography.border.shadow.value } : {}),
      ...(typography?.border?.frame.status === 'valid'
        && typography.border.frame.value !== null
        ? { frame: typography.border.frame.value } : {}),
    } } : {}),
    ...(segment.revision ? { revision: segment.revision } : {}),
    typography: {
      caps: typography?.caps ?? false,
      smallCaps: typography?.smallCaps ?? segment.smallCaps === true,
      strike: typography?.strike ?? segment.strikethrough,
      doubleStrike: typography?.doubleStrike ?? segment.doubleStrikethrough === true,
      verticalAlign: typography?.verticalAlign ?? {
        status: segment.vertAlign ? 'valid' : 'missing',
        raw: segment.vertAlign ?? null,
        value: segment.vertAlign ?? null,
      },
      positionPt: typography?.positionPt ?? {
        status: segment.position === undefined ? 'missing' : 'valid',
        raw: segment.position === undefined ? null : String(segment.position * 2),
        value: segment.position ?? null,
      },
      emphasis: typography?.emphasis ?? {
        status: segment.emphasisMark ? 'valid' : 'missing',
        raw: segment.emphasisMark ?? null,
        value: segment.emphasisMark ?? null,
      },
      ...(typography?.underline ? { underline: typography.underline } : {}),
    },
    decorations: [],
    paintOps: [{
      text: segment.text,
      range: { start: sourceOffset, end: sourceOffset + segment.text.length },
      offset: { xPt: 0, yPt: segment.position ? -segment.position : 0 },
      letterSpacingPt: segment.charSpacing ?? 0,
      scaleX: segment.charScale ?? 1,
      direction: segment.rtl ? 'rtl' : 'ltr',
      kerning: segment.kerning === undefined
        ? 'auto'
        : segment.fontSize >= segment.kerning ? 'normal' : 'none',
      writingMode: segment.verticalRun ? 'vertical-rl' : 'horizontal-tb',
    }],
    ...(segment.hyperlink?.kind === 'external' ? { hyperlink: segment.hyperlink.url } : {}),
  };
}

function plannedBaselinePt(
  measuredLine: MeasuredParagraph['lines'][number],
  context: ParagraphLayoutContext,
): number {
  const raw = measuredLine.layout;
  const visibleAscentPt = raw.visibleAscent ?? raw.ascent;
  const visibleDescentPt = raw.visibleDescent ?? raw.descent;
  const visibleNaturalPt = visibleAscentPt + visibleDescentPt;
  const autoMultiple = context.lineSpacing?.rule === 'auto'
    && !context.hasRuby
    && !context.lineGrid.active;
  const compressedAuto = autoMultiple && (context.lineSpacing?.value ?? 1) < 1;
  const centerBoxPt = autoMultiple && !compressedAuto
    ? Math.max(visibleNaturalPt, raw.visibleIntendedSingle ?? raw.intendedSingle)
    : measuredLine.advancePt;
  return measuredLine.topYPt + (centerBoxPt - visibleNaturalPt) / 2 + visibleAscentPt;
}

interface RetainedNumberingPlan {
  readonly bodyOffsetPt: number;
  readonly markerText: string;
  readonly markerWidthPt: number;
  readonly markerShiftPt: number;
  readonly shape: NonNullable<ReturnType<typeof shapeNumberingMarkerText>>['shape'] | null;
}

function retainedNumberingPlan(
  paragraph: ParagraphAcquisitionInput,
  context: ParagraphLayoutContext,
  options: Pick<ParagraphAcquisitionOptions, 'environment'>,
): RetainedNumberingPlan | undefined {
  const numbering = paragraph.numbering;
  if (!numbering) return undefined;
  if (context.numberingMarkerGeometry) return context.numberingMarkerGeometry;
  const markerInput = paragraph.numberingMarkerShapeInput;
  const service = options.environment.layoutServices?.text;
  if (!markerInput || !service) return undefined;
  return resolveNumberingMarkerGeometry(numbering, markerInput, {
    // Marker alignment is authored at the hanging-indent reference. The
    // context's firstIndentPt is already the resolved BODY offset.
    authoredFirstIndentPt: paragraph.indentFirst,
    physicalIndentLeftPt: context.physicalIndentLeftPt,
    tabStops: paragraph.tabStops,
    defaultTabPt: context.defaultTabPt,
  }, service);
}

function numberingMarkerPlacements(
  plan: RetainedNumberingPlan,
  paragraph: ParagraphAcquisitionInput,
  context: ParagraphLayoutContext,
  paragraphXPt: number,
  availableWidthPt: number,
  line: LineLayout,
): readonly TextPlacement[] {
  if (!plan.shape || plan.markerText === '') return [];
  const shape = plan.shape;
  const markerLeftPt = numberingMarkerPhysicalLeft({
    baseRtl: context.baseRtl,
    paragraphXPt,
    availableWidthPt,
    authoredFirstIndentPt: paragraph.indentFirst,
    markerShiftPt: plan.markerShiftPt,
    markerWidthPt: plan.markerWidthPt,
  });
  const rangeBase = -plan.markerText.length;
  const color: TextPlacement['color'] = paragraph.numbering?.color
    ? { kind: 'explicit', color: `#${paragraph.numbering.color}` }
    : paragraph.numbering?.colorAuto
      ? { kind: 'auto' }
      : paragraph.paragraphMarkColor
        ? { kind: 'explicit', color: `#${paragraph.paragraphMarkColor}` }
        : { kind: 'default' };
  let spanOffsetPt = 0;
  return shape.spans.map((span) => {
    const offsetPt = spanOffsetPt;
    spanOffsetPt += span.advancePt;
    const clusters = shape.clusters
      ? shape.clusters
          .filter((cluster) => cluster.range.start >= span.start && cluster.range.end <= span.end)
          .map((cluster) => ({
            range: { start: rangeBase + cluster.range.start, end: rangeBase + cluster.range.end },
            offset: { xPt: cluster.offsetPt - offsetPt, yPt: 0 },
            advancePt: cluster.advancePt,
          }))
      : [{
          range: { start: rangeBase + span.start, end: rangeBase + span.end },
          offset: { xPt: 0, yPt: 0 }, advancePt: span.advancePt,
        }];
    const xPt = markerLeftPt + offsetPt;
    return {
      kind: 'text', role: 'numbering-marker', text: span.text,
      range: { start: rangeBase + span.start, end: rangeBase + span.end },
      origin: { xPt, yPt: line.baselinePt },
      bounds: {
        xPt, yPt: line.baselinePt - span.ascentPt,
        widthPt: span.advancePt, heightPt: span.ascentPt + span.descentPt,
      },
      advancePt: span.advancePt, clusters,
      paintOps: [{
        text: span.text,
        range: { start: rangeBase + span.start, end: rangeBase + span.end },
        offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 0, scaleX: 1,
        direction: context.baseRtl ? 'rtl' : 'ltr',
        kerning: 'auto', writingMode: 'horizontal-tb',
      }],
      color, fontRoute: span.fontRoute,
      fontSizePt: paragraph.numberingMarkerShapeInput?.fontSizePt ?? span.ascentPt + span.descentPt,
      fontWeight: span.font.weight, fontStyle: span.font.style,
      direction: context.baseRtl ? 'rtl' : 'ltr', decorations: [],
    } satisfies TextPlacement;
  });
}

function retainedHexColor(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  return value.startsWith('#') ? value : `#${value}`;
}

function retainEffectiveTextBackground(
  lines: readonly LineLayout[],
  paragraphShading: string | null | undefined,
  containerShading: string | null | undefined,
): readonly LineLayout[] {
  const paragraphBackground = retainedHexColor(paragraphShading);
  const containerBackground = retainedHexColor(containerShading);
  return lines.map((line) => ({
    ...line,
    placements: line.placements.map((placement) => {
      if (placement.kind !== 'text') return placement;
      const effectiveBackground = placement.background
        ?? paragraphBackground
        ?? containerBackground;
      if (!effectiveBackground || placement.color.kind === 'explicit') return placement;
      return {
        ...placement,
        color: { kind: 'auto', background: effectiveBackground },
      } satisfies TextPlacement;
    }),
  }));
}

function visibleParagraphBorder(
  edge: NonNullable<ParagraphAcquisitionInput['borders']>['top'],
): edge is NonNullable<typeof edge> {
  return edge != null && edge.style !== 'none';
}

function paragraphDecorationBox(
  paragraph: ParagraphAcquisitionInput,
  lines: readonly LineLayout[],
  paragraphXPt: number,
  availableWidthPt: number,
  contentTopPt: number,
  contentHeightPt: number,
  borderEdges: NonNullable<ParagraphAcquisitionOptions['paragraphBorderEdges']>,
): LayoutRect {
  let leftPt = paragraphXPt;
  let rightPt = paragraphXPt + availableWidthPt;
  if (paragraph.indentFirst < 0) {
    if (paragraph.bidi) rightPt -= paragraph.indentFirst;
    else leftPt += paragraph.indentFirst;
  }
  for (const placement of lines.flatMap((line) => line.placements)) {
    const marker = placement.kind === 'text' && placement.role === 'numbering-marker'
      || placement.kind === 'resource' && placement.resourceKind === 'picture-bullet';
    if (!marker || !placement.bounds) continue;
    leftPt = Math.min(leftPt, placement.bounds.xPt);
    rightPt = Math.max(rightPt, placement.bounds.xPt + placement.bounds.widthPt);
  }
  const borders = paragraph.borders;
  const topEdge = borderEdges.top === 'none' ? null : borders?.[borderEdges.top] ?? null;
  const bottomEdge = borderEdges.bottom === 'none' ? null : borders?.bottom ?? null;
  const leftSpacePt = visibleParagraphBorder(borders?.left ?? null) ? borders!.left!.space ?? 0 : 0;
  const rightSpacePt = visibleParagraphBorder(borders?.right ?? null) ? borders!.right!.space ?? 0 : 0;
  const topSpacePt = visibleParagraphBorder(topEdge) ? topEdge.space ?? 0 : 0;
  const bottomSpacePt = visibleParagraphBorder(bottomEdge) ? bottomEdge.space ?? 0 : 0;
  return {
    xPt: leftPt - leftSpacePt,
    yPt: contentTopPt - topSpacePt,
    widthPt: rightPt - leftPt + leftSpacePt + rightSpacePt,
    heightPt: contentHeightPt + topSpacePt + bottomSpacePt,
  };
}

function retainedColorString(color: TextPlacement['color']): string {
  if (color.kind === 'explicit') return color.color;
  if (color.kind === 'auto') return autoContrastColor(color.background ?? '#FFFFFF');
  return '#000000';
}

function completeInkBounds(
  shape: import('./text.js').TextShapeResult,
): import('./text.js').GlyphInkBounds {
  return shape.inkBounds ?? {
    xMinPt: 0,
    xMaxPt: shape.advancePt,
    ascentPt: shape.ascentPt,
    descentPt: shape.descentPt,
  };
}

function emphasisGlyph(mark: string): string {
  if (mark === 'circle') return '○';
  if (mark === 'comma') return '﹅';
  return '•';
}

function retainedGeometryPlan(
  segment: LayoutTextSeg,
  sourceOffset: number,
  color: TextPlacement['color'],
): RetainedTextGeometryPlan | undefined {
  if (!(segment.underline || segment.strikethrough
    || segment.doubleStrikethrough || segment.emphasisMark)) return undefined;
  const service = segment.textLayoutService;
  const request = segment.textShapeRequest;
  if (!service || !request) {
    throw new Error('Retained typography geometry requires TextLayoutService');
  }
  const shape = (text: string) => service.shape({ ...request, text, measure: true });
  const base = shape(segment.text);
  const textColor = retainedColorString(color);
  const underline = segment.underline ? {
    ...(segment.underlineStyle ? { authoredStyle: segment.underlineStyle } : {}),
    color: segment.underlineColor && segment.underlineColor !== 'auto'
      ? `#${segment.underlineColor}` : textColor,
    probe: shape('_'),
  } : undefined;
  const strike = segment.strikethrough || segment.doubleStrikethrough ? {
    double: segment.doubleStrikethrough === true,
    probe: shape('-'),
    ...(segment.doubleStrikethrough ? { doubleProbe: shape('=') } : {}),
  } : undefined;
  const emphasis = segment.emphasisMark ? (() => {
    const glyph = emphasisGlyph(segment.emphasisMark);
    const markShape = shape(glyph);
    const markSpan = markShape.spans[0];
    if (!markSpan) throw new Error('Emphasis shaping produced no selected-face span');
    const clusterInk = (segment.shapedClusters ?? []).map((cluster): RetainedEmphasisClusterInk => {
      const text = segment.text.slice(cluster.range.start, cluster.range.end);
      return {
        text,
        range: {
          start: sourceOffset + cluster.range.start,
          end: sourceOffset + cluster.range.end,
        },
        ink: completeInkBounds(shape(text)),
      };
    });
    return {
      authored: segment.emphasisMark,
      glyph,
      mark: {
        inkBounds: completeInkBounds(markShape),
        fontRoute: markSpan.fontRoute,
        fontSizePt: request.fontSizePt,
        fontWeight: markSpan.font.weight,
        fontStyle: markSpan.font.style,
        color,
      },
      clusterInk,
    };
  })() : undefined;
  return {
    base,
    ...(underline ? { underline } : {}),
    ...(strike ? { strike } : {}),
    ...(emphasis ? { emphasis } : {}),
  };
}

function textPlanSegment(
  segment: LayoutTextSeg,
  paragraph: DocParagraph,
  sourceOffset: number,
  characterGridDeltaPt: number,
  sourceRun?: DocRun & Readonly<{ anchorOccurrenceId?: string }>,
): MeasuredTextPlanSegment | MeasuredAnchorHostPlanSegment {
  if (segment.metricOnly) {
    const sourceMetrics = selectedFaceSourceMetrics(segment);
    return {
      kind: 'anchor-host', measuredWidthPt: 0,
      range: { start: sourceOffset, end: sourceOffset },
      ...(sourceMetrics ? { sourceMetrics } : {}),
      ...(sourceRun?.type === 'anchorHost' && sourceRun.anchorOccurrenceId
        ? { anchorOccurrenceId: sourceRun.anchorOccurrenceId }
        : {}),
    };
  }
  const projected = textPlacement(segment, paragraph, sourceOffset, 0, 0, 0, 0);
  if (projected.kind !== 'text') throw new Error('Visible text segment projected as anchor host');
  const pitchPt = segment.fitTextPerGapPx
    ?? (segment.charSpacing ?? 0) + (segment.snapToCharacterGrid === false ? 0 : characterGridDeltaPt);
  const scaleX = segment.charScale ?? 1;
  const retainedGeometry = retainedGeometryPlan(segment, sourceOffset, projected.color);
  const candidateClusters = segment.shapedClusters;
  const shapedClusters = candidateClusters?.length
    && candidateClusters[0]?.range.start === 0
    && candidateClusters.at(-1)?.range.end === segment.text.length
    && candidateClusters.every((cluster, index) => index === 0
      || candidateClusters[index - 1]?.range.end === cluster.range.start)
    && candidateClusters.every((cluster) =>
      cluster.range.start < cluster.range.end
      && Number.isFinite(cluster.offsetPt)
      && Number.isFinite(cluster.advancePt))
      ? candidateClusters
      : undefined;
  if (segment.text.length > 0 && !shapedClusters) {
    throw new Error(
      'Visible text acquisition requires complete authoritative grapheme clusters from TextLayoutService',
    );
  }
  const clusters = (shapedClusters ?? []).map((cluster, index) => {
    const prefix = segment.text.slice(0, cluster.range.start);
    const text = segment.text.slice(cluster.range.start, cluster.range.end);
    const precedingScalars = [...prefix].length;
    const scalarCount = [...text].length;
    const trailingFitPad = index === (shapedClusters?.length ?? 0) - 1
      ? segment.fitTextTrailingPadPx ?? 0
      : 0;
    return {
      range: {
        start: sourceOffset + cluster.range.start,
        end: sourceOffset + cluster.range.end,
      },
      offset: {
        xPt: cluster.offsetPt * scaleX + precedingScalars * pitchPt,
        yPt: segment.position ? -segment.position : 0,
      },
      advancePt: cluster.advancePt * scaleX + scalarCount * pitchPt + trailingFitPad,
    };
  });
  const {
    origin: _origin, bounds: _bounds, advancePt: _advancePt,
    paintOps, clusters: _clusters, ...style
  } = projected;
  const basePaintOps = segment.verticalRun
    ? clusters.map((cluster) => {
        const text = segment.text.slice(
          cluster.range.start - sourceOffset,
          cluster.range.end - sourceOffset,
        );
        const template = paintOps[0]!;
        const upright = EAST_ASIAN_RE.test(text);
        return {
          ...template,
          text,
          range: cluster.range,
          // Upright vertical glyphs rotate around the retained cell centre.
          // Keep that pivot in acquisition geometry: paint must not recover it
          // from font metrics or neighboring operations.
          offset: upright
            ? { xPt: cluster.offset.xPt + cluster.advancePt / 2, yPt: cluster.offset.yPt }
            : cluster.offset,
          letterSpacingPt: 0,
          glyphOrientation: upright ? 'upright' as const : 'sideways' as const,
        };
      })
    : segment.tateChuYoko
      ? paintOps.map((operation) => ({
          ...operation,
          offset: {
            xPt: operation.offset.xPt + segment.measuredWidth / 2,
            yPt: operation.offset.yPt,
          },
          glyphOrientation: 'upright' as const,
        }))
      : paintOps;
  return {
    ...style,
    kind: 'text', measuredWidthPt: segment.measuredWidth,
    clusters,
    basePaintOps: basePaintOps.map((operation) => ({
      ...operation,
      // Measurement resolves w:spacing, docGrid character pitch, and w:fitText
      // into one authoritative per-scalar pitch. Paint retains that same value;
      // it must never reconstruct pitch from parser source or remeasure glyphs.
      letterSpacingPt: pitchPt,
    })),
    breakBefore: segment.breakBefore !== false && !segment.joinPrev,
    rtl: segment.rtl,
    digitsAsAN: segment.digitsAsAN,
    fixedPitch: segment.fitTextRegionIndex !== undefined,
    ...(retainedGeometry ? { retainedGeometry } : {}),
    ...(segment.textLayoutService ? { textLayoutService: segment.textLayoutService } : {}),
    ...(segment.textShapeRequest ? { textShapeRequest: segment.textShapeRequest } : {}),
  };
}

interface LogicalOccurrenceMap {
  readonly runStarts: readonly number[];
  readonly runLengths: readonly number[];
}

/**
 * One paragraph-local occurrence domain shared by retained ranges and flow
 * events. Text and resolved field/math fallback values use UTF-16 offsets so a
 * TextRange slices the corresponding JavaScript string without conversion.
 * Atomic controls/resources (break, tab, image/chart, shape) consume one unit;
 * a metric-only anchor host consumes zero because it contributes no selectable
 * content. This makes source-run indices an acquisition concern only.
 */
function logicalOccurrenceMap(
  paragraph: ParagraphAcquisitionInput,
  measured: MeasuredParagraph,
): LogicalOccurrenceMap {
  const measuredLengths = new Map<number, number>();
  for (const line of measured.lines) {
    for (const segment of line.layout.segments) {
      const runIndex = sourceRunIndex(segment);
      if (runIndex === undefined) continue;
      const length = 'text' in segment
        ? (segment.metricOnly ? 0 : segment.text.length)
        : 'mathNodes' in segment ? segment.fallbackText.length
          : 'isTab' in segment || 'imagePath' in segment ? 1 : 0;
      measuredLengths.set(runIndex, (measuredLengths.get(runIndex) ?? 0) + length);
    }
  }
  const runLengths = paragraph.runs.map((run, runIndex) => {
    const measuredLength = measuredLengths.get(runIndex);
    if (measuredLength !== undefined) return measuredLength;
    if (run.type === 'text') return run.text.length;
    if (run.type === 'field') return run.fallbackText.length;
    if (run.type === 'anchorHost') return 0;
    return 1;
  });
  let cursor = 0;
  const runStarts = runLengths.map((length) => {
    const start = cursor;
    cursor += length;
    return start;
  });
  return { runStarts, runLengths };
}

function segmentOccurrenceLength(segment: LayoutTextSeg | LayoutTabSeg | LayoutImageSeg | LayoutMathSeg): number {
  if ('text' in segment) return segment.metricOnly ? 0 : segment.text.length;
  if ('mathNodes' in segment) return segment.fallbackText.length;
  return 1;
}

function planMeasuredLines(
  measured: MeasuredParagraph,
  paragraph: DocParagraph,
  paragraphXPt: number,
  availableWidthPt: number,
  source: SourceRef,
  context: ParagraphLayoutContext,
  occurrences: LogicalOccurrenceMap,
  numberingPlan?: RetainedNumberingPlan,
  textService?: import('./text.js').TextLayoutService,
): readonly LineLayout[] {
  let sourceOffset = 0;
  const consumedByRun = new Map<number, number>();
  const hasExplicitTab = measured.lines.some((line) => line.layout.segments.some((segment) => 'isTab' in segment));
  const earliestTab = paragraph.tabStops?.reduce<(typeof paragraph.tabStops)[number] | undefined>(
    (earliest, stop) => !earliest || stop.pos < earliest.pos ? stop : earliest,
    undefined,
  );
  const visibleText = measured.lines.flatMap((line) => line.layout.segments.flatMap((segment) =>
    'text' in segment && !segment.metricOnly ? [segment.text] : [])).join('').trim();
  const decimalAutoTabPt = !hasExplicitTab
    && earliestTab?.alignment === 'decimal'
    && visibleText !== ''
    && /^[+\-(]?[\d., ]+\)?%?$/u.test(visibleText)
      ? earliestTab.pos - context.physicalIndentLeftPt
      : undefined;
  return measured.lines.map((measuredLine, lineIndex) => {
    const raw = measuredLine.layout;
    const baselinePt = plannedBaselinePt(measuredLine, context);
    let lineStartOffset = Number.POSITIVE_INFINITY;
    let lineEndOffset = sourceOffset;
    const segments: MeasuredLinePlanSegment[] = [];
    for (const segment of raw.segments) {
      const runIndex = sourceRunIndex(segment);
      const sourceRun = runIndex === undefined ? undefined : paragraph.runs[runIndex];
      const occurrenceLength = segmentOccurrenceLength(segment);
      const segmentOffset = runIndex === undefined
        ? sourceOffset
        : (occurrences.runStarts[runIndex] ?? sourceOffset) + (consumedByRun.get(runIndex) ?? 0);
      if (runIndex !== undefined) {
        consumedByRun.set(runIndex, (consumedByRun.get(runIndex) ?? 0) + occurrenceLength);
      }
      lineStartOffset = Math.min(lineStartOffset, segmentOffset);
      lineEndOffset = Math.max(lineEndOffset, segmentOffset + occurrenceLength);
      if ('isTab' in segment) {
        const tab = segment as LayoutTabSeg;
        const leader = tab.leader ?? 'none';
        let leaderShape: MeasuredTabPlanSegment['leaderShape'];
        if (leader !== 'none') {
          if (!textService) {
            throw new Error('Tab leader acquisition requires TextLayoutService');
          }
          const glyph = leader === 'hyphen' ? '-'
            : leader === 'underscore' || leader === 'heavy' ? '_'
              : leader === 'middleDot' ? '·' : '.';
          const textSource = sourceRun?.type === 'text' || sourceRun?.type === 'field'
            ? sourceRun
            : undefined;
          const richRun = textSource as (typeof textSource & Readonly<{
            fontSlots?: Readonly<{
              direct: import('./text.js').TextFontSlots;
              theme?: import('./text.js').TextFontSlots;
              themePresent?: import('./text.js').TextFontSlotPresence;
            }>;
            colorAuto?: boolean;
          }>) | undefined;
          const shape = textService.shape({
            text: glyph,
            fontSizePt: tab.fontSize,
            fonts: richRun?.fontSlots?.direct
              ?? (textSource?.fontFamily ? { ascii: textSource.fontFamily } : {}),
            themeFonts: richRun?.fontSlots?.theme,
            themeFontPresence: richRun?.fontSlots?.themePresent,
            weight: tab.bold ? 700 : 400,
            style: tab.italic ? 'italic' : 'normal',
            measure: true,
          });
          const span = shape.spans[0];
          if (!span || !Number.isFinite(shape.advancePt) || shape.advancePt <= 0) {
            throw new Error('Tab leader acquisition produced no shaped glyph advance');
          }
          leaderShape = {
            glyph,
            advancePt: shape.advancePt,
            fontRoute: span.fontRoute,
            fontSizePt: tab.fontSize,
            fontWeight: span.font.weight,
            fontStyle: span.font.style,
            color: textSource?.color
              ? { kind: 'explicit', color: `#${textSource.color}` }
              : richRun?.colorAuto ? { kind: 'auto' } : { kind: 'default' },
          };
        }
        segments.push({
          kind: 'tab', range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
          measuredWidthPt: tab.measuredWidth, leader,
          fontSizePt: tab.fontSize, bold: tab.bold, italic: tab.italic,
          ...(leaderShape ? { leaderShape } : {}),
        });
      } else if ('imagePath' in segment) {
        const image = segment as LayoutImageSeg;
        if (image.anchor) continue;
        const runIndex = sourceRunIndex(segment);
        const occurrence = runSource(source, runIndex ?? 0);
        const resourceKind = image.chart ? 'chart' : 'image';
        const resourceKey = image.chart ? chartResourceKey(occurrence) : imageResourceKey(occurrence, image.imagePath);
        segments.push({
          kind: 'resource', range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
          resourceKey, resourceKind, measuredWidthPt: image.measuredWidth,
          widthPt: image.widthPt, heightPt: image.heightPt, topOffsetPt: -image.heightPt,
        });
      } else if ('mathNodes' in segment) {
        const math = segment as LayoutMathSeg;
        segments.push({
          kind: 'resource',
          range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
          resourceKey: math.mathResourceKey, resourceKind: 'math',
          measuredWidthPt: math.measuredWidth, widthPt: math.measuredWidth,
          heightPt: math.mathAscent + math.mathDescent, topOffsetPt: -math.mathAscent,
        });
      } else {
        segments.push(textPlanSegment(
          segment as LayoutTextSeg, paragraph, segmentOffset,
          context.characterGrid.active ? context.characterGrid.deltaPt : 0,
          sourceRun,
        ));
      }
      sourceOffset = Math.max(sourceOffset, segmentOffset + occurrenceLength);
    }
    const onlyMath = raw.segments.length === 1 && 'mathNodes' in (raw.segments[0] ?? {} as object)
      ? raw.segments[0] as LayoutMathSeg
      : undefined;
    return planLine({
      paragraphXPt, availableWidthPt, alignment: paragraph.alignment,
      baseRtl: context.baseRtl,
      isFirstLine: lineIndex === 0,
      isLastLine: lineIndex === measured.lines.length - 1,
      stretchLastLine: context.stretchLastLine,
      firstLineIndentPt: context.firstIndentPt,
      ...(lineIndex === 0 && numberingPlan
        ? { numbering: { bodyOffsetPt: numberingPlan.bodyOffsetPt } }
        : {}),
      ...(decimalAutoTabPt === undefined ? {} : { decimalAutoTabPt }),
      ...(onlyMath?.display ? {
        displayMathJustification: onlyMath.jc ?? context.mathDefJc ?? 'centerGroup',
      } : {}),
      line: {
        range: {
          start: Number.isFinite(lineStartOffset) ? lineStartOffset : sourceOffset,
          end: lineEndOffset,
        },
        topPt: measuredLine.topYPt,
        baselinePt,
        advancePt: measuredLine.advancePt,
        xOffsetPt: raw.xOffset,
        availableWidthPt: raw.availWidth,
        endsWithBreak: raw.endsWithBreak ?? false,
        segments,
      },
    });
  });
}

/** A mark-only paragraph still owns a real line box and baseline when numbering
 * paints there. Materializing that host through `planLine` keeps the marker on
 * the same retained geometry path as a marker followed by body text. */
function numberingMarkerHostLine(
  measured: MeasuredParagraph,
  paragraph: DocParagraph,
  paragraphXPt: number,
  availableWidthPt: number,
  context: ParagraphLayoutContext,
): LineLayout {
  const advancePt = measured.contentEndYPt - measured.contentStartYPt;
  return planLine({
    paragraphXPt,
    availableWidthPt,
    alignment: paragraph.alignment,
    baseRtl: context.baseRtl,
    isFirstLine: true,
    isLastLine: true,
    stretchLastLine: context.stretchLastLine,
    line: {
      range: { start: 0, end: 0 },
      topPt: measured.contentStartYPt,
      baselinePt: measured.contentEndYPt - measured.lastLineBelowBaselinePt,
      advancePt,
      xOffsetPt: 0,
      availableWidthPt,
      endsWithBreak: false,
      segments: [],
    },
  });
}

function resolvedShapeLayoutRect(
  shape: ShapeRun,
  options: ParagraphAcquisitionOptions,
): LayoutRect {
  const xPt = shape.anchorXPt + (shape.anchorXFromMargin ? options.placement.paragraphXPt : 0);
  const yPt = shape.anchorYPt + (shape.anchorYFromPara ? options.placement.startYPt : 0);
  return { xPt, yPt, widthPt: shape.widthPt, heightPt: shape.heightPt };
}

interface PublicAnchorFacts {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly anchorXPt?: number;
  readonly anchorYPt?: number;
  readonly anchorXFromMargin?: boolean;
  readonly anchorYFromPara?: boolean;
  readonly anchorXAlign?: string | null;
  readonly anchorYAlign?: string | null;
  readonly anchorXRelativeFrom?: string | null;
  readonly anchorYRelativeFrom?: string | null;
  readonly pctPosH?: number | null;
  readonly pctPosV?: number | null;
  readonly widthPct?: number | null;
  readonly heightPct?: number | null;
  readonly widthRelativeFrom?: string | null;
  readonly heightRelativeFrom?: string | null;
  readonly distTop?: number;
  readonly distRight?: number;
  readonly distBottom?: number;
  readonly distLeft?: number;
  readonly wrapMode?: string | null;
  readonly wrapSide?: string | null;
  readonly behindDoc?: boolean;
  readonly allowOverlap?: boolean;
}

const missingAnchorEdges = (): AnchorEdgesInput => ({
  topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
});

function publicAnchorEdges(run: PublicAnchorFacts): AnchorEdgesInput {
  const edge = (value: number | undefined) => Number.isFinite(value)
    ? { value: value as number, status: 'valid' as const }
    : { value: null, status: 'missing' as const };
  const top = edge(run.distTop);
  const right = edge(run.distRight);
  const bottom = edge(run.distBottom);
  const left = edge(run.distLeft);
  return {
    topPt: top.value, topStatus: top.status,
    rightPt: right.value, rightStatus: right.status,
    bottomPt: bottom.value, bottomStatus: bottom.status,
    leftPt: left.value, leftStatus: left.status,
  };
}

function publicAnchorChoice(
  align: string | null | undefined,
  percent: number | null | undefined,
  offset: number | undefined,
): AnchorAxisChoiceInput {
  if (align !== undefined && align !== null) return { kind: 'align', value: align };
  if (percent !== undefined && percent !== null) return { kind: 'percent', fraction: percent };
  return { kind: 'offset', valuePt: offset ?? 0 };
}

/** Hand-built public runs do not carry the parser-private wire record. Preserve
 * their authored positioning choices here so the page adapter can resolve them
 * against the final occurrence instead of blessing the provisional box. */
function publicAnchorAcquisition(
  run: PublicAnchorFacts,
  occurrenceId: string,
  relativeHeight: number,
): AnchorAcquisitionInput {
  const horizontalFrom = run.anchorXRelativeFrom
    ?? (run.anchorXFromMargin ? 'margin' : 'page');
  const verticalFrom = run.anchorYRelativeFrom
    ?? (run.anchorYFromPara ? 'paragraph' : 'page');
  const wrapKind = run.wrapMode == null || run.wrapMode === 'none'
    ? 'none'
    : run.wrapMode === 'square' || run.wrapMode === 'tight'
      || run.wrapMode === 'through' || run.wrapMode === 'topAndBottom'
      ? run.wrapMode
      : 'invalid';
  const relativeSize = (
    fraction: number | null | undefined,
    relativeFrom: string | null | undefined,
  ) => fraction === undefined || fraction === null ? null : {
    relativeFrom: relativeFrom ?? null,
    relativeFromStatus: relativeFrom == null ? 'missing' as const : 'valid' as const,
    fraction,
    fractionStatus: Number.isFinite(fraction) ? 'valid' as const : 'invalid' as const,
  };
  return {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid',
      xPt: null, xStatus: 'missing', yPt: null, yStatus: 'missing',
    },
    horizontal: {
      relativeFrom: horizontalFrom, relativeFromStatus: 'valid',
      choice: publicAnchorChoice(run.anchorXAlign, run.pctPosH, run.anchorXPt),
    },
    vertical: {
      relativeFrom: verticalFrom, relativeFromStatus: 'valid',
      choice: publicAnchorChoice(run.anchorYAlign, run.pctPosV, run.anchorYPt),
    },
    extent: {
      widthPt: run.widthPt, heightPt: run.heightPt,
      widthStatus: Number.isFinite(run.widthPt) ? 'valid' : 'invalid',
      heightStatus: Number.isFinite(run.heightPt) ? 'valid' : 'invalid',
    },
    parentEffectExtent: missingAnchorEdges(),
    anchorDistances: publicAnchorEdges(run),
    relativeSize: {
      horizontal: relativeSize(run.widthPct, run.widthRelativeFrom),
      vertical: relativeSize(run.heightPct, run.heightRelativeFrom),
    },
    wrap: {
      kind: wrapKind,
      authoredKinds: run.wrapMode == null ? [] : [run.wrapMode],
      side: run.wrapSide ?? null,
      distances: missingAnchorEdges(), effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: run.behindDoc ?? false, behindDocStatus: 'valid',
      relativeHeight, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: run.allowOverlap ?? true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
}

function publicAnchorReferenceFrames(
  lines: readonly LineLayout[],
  options: ParagraphAcquisitionOptions,
  paragraphHeightPt: number,
): AnchorReferenceFramesInput {
  const paragraph = {
    xPt: options.placement.paragraphXPt,
    yPt: options.placement.startYPt,
    widthPt: options.placement.availableWidthPt,
    heightPt: Math.max(0, paragraphHeightPt),
  };
  const line = lines[0]?.bounds ?? paragraph;
  const character = lines[0]?.placements[0]?.bounds ?? line;
  return {
    page: options.anchorFrames?.page ?? null,
    margin: options.anchorFrames?.margin ?? null,
    column: options.anchorFrames?.column ?? null,
    paragraph, line, character,
    pageParity: options.anchorFrames?.pageParity ?? null,
  };
}

function acquiredAnchorNormalization(
  acquisition: AnchorAcquisitionInput,
  frames: AnchorReferenceFramesInput,
) {
  return snapshotPlainData({
    acquisition,
    pageParity: frames.pageParity,
    physicalFrames: { page: frames.page, margin: frames.margin, column: frames.column },
    logicalHostFrames: {
      paragraph: frames.paragraph as NonNullable<typeof frames.paragraph>,
      line: frames.line as NonNullable<typeof frames.line>,
      character: frames.character as NonNullable<typeof frames.character>,
    },
  }, 'Anchor normalization facts');
}

function drawingForShape(
  shape: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'shape' }>,
  rect: LayoutRect,
  options: ParagraphAcquisitionOptions,
  runIndex: number,
  frames: AnchorReferenceFramesInput,
): DrawingLayout {
  const plan = planShapeDrawing(
    shape,
    rect,
    options.environment.layoutServices?.text,
    shape.vmlTextPathInput,
  );
  const commands = [plan.command];
  return {
    kind: 'drawing', id: `${options.id}:drawing:${runIndex}`, source: runSource(options.source, runIndex),
    flowDomainId: options.flowDomainId, flowBounds: rect, inkBounds: rect, advancePt: 0,
    ordinaryFlow: false,
    commands,
    anchorLayer: {
      occurrenceId: `public-shape:${options.id}:${runIndex}`,
      coordinateSpace: 'acquired-anchor-points',
      normalization: acquiredAnchorNormalization(publicAnchorAcquisition(
        shape,
        `public-shape:${options.id}:${runIndex}`,
        Number.isFinite(shape.zOrder) ? shape.zOrder : runIndex,
      ), frames),
      behindDoc: shape.behindDoc === true,
      relativeHeight: Number.isFinite(shape.zOrder) ? shape.zOrder : runIndex,
      sourceOrder: runIndex,
      horizontalOwnership: shape.anchorXRelativeFrom === 'character' ? 'host' : 'page',
      verticalOwnership: shape.anchorYRelativeFrom === 'paragraph'
        || shape.anchorYRelativeFrom === 'line'
        || shape.anchorYRelativeFrom === 'character'
        || (!shape.anchorYRelativeFrom && shape.anchorYFromPara)
        ? 'host' : 'page',
    },
  };
}

function publicAnchoredResourceDrawing(
  run: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'image' | 'chart' }>,
  options: ParagraphAcquisitionOptions,
  runIndex: number,
  normalizationFrames: AnchorReferenceFramesInput,
): DrawingLayout | null {
  if (!run.anchor || run.anchorAcquisitionInput) return null;
  const frames = options.anchorFrames;
  const horizontalReference = run.anchorXRelativeFrom ?? (run.anchorXFromMargin ? 'margin' : 'page');
  const verticalReference = run.anchorYRelativeFrom ?? (run.anchorYFromPara ? 'paragraph' : 'page');
  const horizontalFrame = horizontalReference === 'margin' ? frames?.margin
    : horizontalReference === 'column' ? frames?.column
      : horizontalReference === 'page' ? frames?.page : null;
  const verticalFrame = verticalReference === 'paragraph' ? {
    xPt: options.placement.paragraphXPt,
    yPt: options.placement.startYPt,
    widthPt: options.placement.availableWidthPt,
    heightPt: 0,
  } : verticalReference === 'margin' ? frames?.margin
    : verticalReference === 'column' ? frames?.column
      : verticalReference === 'page' ? frames?.page : null;
  if (!horizontalFrame || !verticalFrame) return null;
  const widthPt = run.widthPt;
  const heightPt = run.heightPt;
  const alignX = run.anchorXAlign;
  const alignY = run.anchorYAlign;
  const xPt = alignX === 'right'
    ? horizontalFrame.xPt + horizontalFrame.widthPt - widthPt
    : alignX === 'center'
      ? horizontalFrame.xPt + (horizontalFrame.widthPt - widthPt) / 2
      : horizontalFrame.xPt + (run.anchorXPt ?? 0);
  const yPt = alignY === 'bottom'
    ? verticalFrame.yPt + verticalFrame.heightPt - heightPt
    : alignY === 'center'
      ? verticalFrame.yPt + (verticalFrame.heightPt - heightPt) / 2
      : verticalFrame.yPt + (run.anchorYPt ?? 0);
  const rect = { xPt, yPt, widthPt, heightPt };
  const source = runSource(options.source, runIndex);
  return {
    kind: 'drawing', id: `${options.id}:public-anchor-drawing:${runIndex}`, source,
    flowDomainId: options.flowDomainId, flowBounds: rect, inkBounds: rect,
    advancePt: 0, ordinaryFlow: false,
    commands: [{
      kind: 'resource',
      resourceKind: run.type,
      resourceKey: run.type === 'image'
        ? imageResourceKey(source, run.imagePath) : chartResourceKey(source),
      rect,
    }],
    anchorLayer: {
      occurrenceId: `public-anchor:${options.id}:${runIndex}`,
      coordinateSpace: 'acquired-anchor-points',
      normalization: acquiredAnchorNormalization(publicAnchorAcquisition(
        run,
        `public-anchor:${options.id}:${runIndex}`,
        runIndex,
      ), normalizationFrames),
      behindDoc: false,
      relativeHeight: runIndex,
      sourceOrder: runIndex,
      horizontalOwnership: 'page',
      verticalOwnership: verticalReference === 'paragraph' ? 'host' : 'page',
    },
  };
}

type AnchoredPayloadRun = Extract<
  ParagraphAcquisitionInput['runs'][number],
  { type: 'image' | 'chart' | 'shape' }
> & Readonly<{ anchorAcquisitionInput?: import('./anchor-input.js').AnchorAcquisitionInput }>;

function anchoredPayloadRun(
  run: ParagraphAcquisitionInput['runs'][number],
): run is AnchoredPayloadRun {
  return (run.type === 'image' || run.type === 'chart' || run.type === 'shape')
    && run.anchorAcquisitionInput !== undefined;
}

function rectanglePolygon(rect: LayoutRect): readonly PointPt[] {
  return [
    { xPt: rect.xPt, yPt: rect.yPt },
    { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt },
    { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt + rect.heightPt },
    { xPt: rect.xPt, yPt: rect.yPt + rect.heightPt },
  ];
}

function retainedAnchorChildFrame(
  acquisition: NonNullable<AnchoredPayloadRun['anchorAcquisitionInput']>,
  outerFrame: LayoutRect,
): LayoutRect {
  const child = acquisition.group?.resolvedChildFrame;
  if (!child) return outerFrame;
  const authoredWidthPt = acquisition.extent.widthPt;
  const authoredHeightPt = acquisition.extent.heightPt;
  if (
    acquisition.extent.widthStatus !== 'valid'
    || acquisition.extent.heightStatus !== 'valid'
    || authoredWidthPt === null
    || authoredHeightPt === null
    || authoredWidthPt <= 0
    || authoredHeightPt <= 0
  ) {
    throw new Error('resolved grouped anchor requires its authored wp:extent');
  }
  const scaleX = outerFrame.widthPt / authoredWidthPt;
  const scaleY = outerFrame.heightPt / authoredHeightPt;
  return {
    xPt: outerFrame.xPt + child.offsetXPt * scaleX,
    yPt: outerFrame.yPt + child.offsetYPt * scaleY,
    widthPt: child.widthPt * scaleX,
    heightPt: child.heightPt * scaleY,
  };
}

function anchorAxisOwnership(
  result: Extract<AnchorFrameResult, { status: 'resolved' }>,
  axis: 'horizontal' | 'vertical',
): 'page' | 'host' {
  const diagnostic = result.axes[axis];
  if (diagnostic.status !== 'resolved') return 'host';
  return diagnostic.referenceFrame === 'paragraph'
    || diagnostic.referenceFrame === 'line'
    || diagnostic.referenceFrame === 'character'
    ? 'host'
    : 'page';
}

interface AcquiredAnchorOccurrence {
  readonly result: AnchorFrameResult;
  readonly drawing?: DrawingLayout;
  readonly exclusion?: WrapExclusion;
  readonly textBoxes: readonly TextBoxLayout[];
  readonly hostLineIndex: number;
  readonly hostRange: import('./types.js').TextRange;
}

function acquireAnchorOccurrence(
  occurrenceId: string,
  payloads: readonly Readonly<{ run: AnchoredPayloadRun; runIndex: number }>[],
  lines: readonly LineLayout[],
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  paragraphHeightPt: number,
): AcquiredAnchorOccurrence | null {
  let hostLineIndex = -1;
  let host: Extract<ParagraphPlacement, { kind: 'anchor-host' }> | undefined;
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const found = lines[lineIndex]?.placements.find((placement) =>
      placement.kind === 'anchor-host' && placement.anchorOccurrenceId === occurrenceId);
    if (found?.kind === 'anchor-host') {
      hostLineIndex = lineIndex;
      host = found;
      break;
    }
  }
  if (!host || hostLineIndex < 0) return null;
  const ordered = [...payloads].sort((a, b) =>
    (a.run.anchorAcquisitionInput?.group?.sourceIndex ?? 0)
      - (b.run.anchorAcquisitionInput?.group?.sourceIndex ?? 0)
    || a.runIndex - b.runIndex);
  const outer = ordered[0];
  if (!outer?.run.anchorAcquisitionInput) return null;
  const line = lines[hostLineIndex]!;
  const baseFrames = options.anchorFrames;
  const referenceFrames: AnchorReferenceFramesInput = {
    page: baseFrames?.page ?? null,
    margin: baseFrames?.margin ?? null,
    column: baseFrames?.column ?? null,
    paragraph: {
      xPt: options.placement.paragraphXPt,
      yPt: options.placement.startYPt,
      widthPt: options.placement.availableWidthPt,
      heightPt: Math.max(0, paragraphHeightPt),
    },
    line: line.bounds,
    character: host.bounds,
    pageParity: baseFrames?.pageParity ?? null,
  };
  const result = resolveAnchorFrame({
    acquisition: outer.run.anchorAcquisitionInput,
    frames: referenceFrames,
  });
  if (result.status !== 'resolved') {
    return { result, textBoxes: [], hostLineIndex, hostRange: host.range };
  }
  const behavior = outer.run.anchorAcquisitionInput.behavior;
  if (
    behavior.behindDocStatus !== 'valid'
    || behavior.relativeHeightStatus !== 'valid'
    || behavior.behindDoc === null
    || behavior.relativeHeight === null
  ) {
    throw new Error('resolved anchor frame must retain required CT_Anchor behavior');
  }
  const authoredRect = result.geometry.objectFrame;
  const commands: DrawingPaintCommand[] = [];
  const textBoxes: TextBoxLayout[] = [];
  const textBoxIds: string[] = [];
  const acquiredShapeTextBoxes = new Map<number, TextBoxLayout>();
  let rect = authoredRect;
  if (outer.run.type === 'shape' && outer.run.anchorAcquisitionInput.group === null) {
    const source = runSource(options.source, outer.runIndex);
    const textBox = acquireShapeTextBoxLayout(outer.run, authoredRect, {
      id: `${options.id}:anchor-textbox:${occurrenceId}:${outer.runIndex}`,
      source,
      flowDomainId: options.flowDomainId,
      context: options.context,
      measurer: options.measurer,
      environment: options.environment,
      input: outer.run.textBoxInput,
    });
    if (textBox) {
      acquiredShapeTextBoxes.set(outer.runIndex, textBox);
      rect = textBox.flowBounds;
    }
  }
  const effectiveResult = resizeResolvedAnchorGeometry(result, rect);
  for (const { run, runIndex } of ordered) {
    const source = runSource(options.source, runIndex);
    const acquisition = run.anchorAcquisitionInput as NonNullable<typeof run.anchorAcquisitionInput>;
    const commandRect = retainedAnchorChildFrame(acquisition, rect);
    if (run.type === 'image') {
      commands.push({
        kind: 'resource', resourceKind: 'image',
        resourceKey: imageResourceKey(source, run.imagePath), rect: commandRect,
      });
    } else if (run.type === 'chart') {
      commands.push({
        kind: 'resource', resourceKind: 'chart',
        resourceKey: chartResourceKey(source), rect: commandRect,
      });
    } else {
      const childTransform = acquisition.group?.resolvedChildFrame;
      const plannedRun = childTransform ? {
        ...run,
        rotation: childTransform.rotationDeg,
        flipH: childTransform.flipH,
        flipV: childTransform.flipV,
      } : run;
      commands.push(planShapeDrawing(
        plannedRun,
        commandRect,
        options.environment.layoutServices?.text,
        run.vmlTextPathInput,
      ).command);
      const textBoxId = `${options.id}:anchor-textbox:${occurrenceId}:${runIndex}`;
      const textBox = acquiredShapeTextBoxes.get(runIndex) ?? acquireShapeTextBoxLayout(run, commandRect, {
        id: textBoxId,
        source,
        flowDomainId: options.flowDomainId,
        context: options.context,
        measurer: options.measurer,
        environment: options.environment,
        input: run.textBoxInput,
      });
      if (textBox) {
        textBoxes.push(textBox);
        textBoxIds.push(textBoxId);
      }
    }
  }
  const drawing: DrawingLayout = {
    kind: 'drawing',
    id: `${options.id}:anchor-drawing:${occurrenceId}`,
    source: runSource(options.source, outer.runIndex),
    flowDomainId: options.flowDomainId,
    flowBounds: rect,
    inkBounds: effectiveResult.geometry.inkBounds,
    advancePt: 0,
    ordinaryFlow: false,
    commands,
    anchorLayer: {
      occurrenceId,
      coordinateSpace: 'acquired-anchor-points',
      normalization: acquiredAnchorNormalization(
        outer.run.anchorAcquisitionInput,
        referenceFrames,
      ),
      behindDoc: behavior.behindDoc,
      relativeHeight: behavior.relativeHeight,
      sourceOrder: outer.runIndex,
      horizontalOwnership: anchorAxisOwnership(effectiveResult, 'horizontal'),
      verticalOwnership: anchorAxisOwnership(effectiveResult, 'vertical'),
    },
    ...(textBoxIds.length ? { textBoxIds } : {}),
  };
  const wrapBounds = effectiveResult.geometry.wrapBounds;
  const exclusion = wrapBounds && effectiveResult.geometry.wrap.kind !== 'none' ? {
    id: `${options.id}:anchor-exclusion:${occurrenceId}`,
    wrap: effectiveResult.geometry.wrap.kind,
    bounds: wrapBounds,
    polygon: effectiveResult.geometry.wrap.polygon?.points ?? rectanglePolygon(wrapBounds),
    anchorOccurrenceId: occurrenceId,
    verticalOwnership: anchorAxisOwnership(effectiveResult, 'vertical'),
  } satisfies WrapExclusion : undefined;
  return {
    result: effectiveResult, drawing, exclusion, textBoxes,
    hostLineIndex, hostRange: host.range,
  };
}

export interface ShapeTextBoxAcquisitionOptions {
  readonly id: string;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly context: ParagraphLayoutContext;
  readonly measurer: TextMeasurer;
  readonly environment: ParagraphMeasurementEnvironment;
  readonly input?: readonly NormalizedTextBoxParagraphInput[];
}

function textBoxParagraphContext(
  inherited: ParagraphLayoutContext,
  paragraph: ParagraphAcquisitionInput,
): ParagraphLayoutContext {
  const baseRtl = paragraph.bidi === true;
  const hasRuby = paragraph.runs.some((run) => run.type === 'text' && Boolean(run.ruby));
  const hasEastAsianText = paragraph.runs.some((run) =>
    run.type === 'text' && EAST_ASIAN_RE.test(run.text));
  return {
    ...inherited,
    physicalIndentLeftPt: baseRtl ? paragraph.indentRight : paragraph.indentLeft,
    physicalIndentRightPt: baseRtl ? paragraph.indentLeft : paragraph.indentRight,
    firstIndentPt: paragraph.indentFirst,
    lineSpacing: paragraph.lineSpacing,
    spaceBeforePt: paragraph.spaceBefore,
    spaceAfterPt: paragraph.spaceAfter,
    baseRtl,
    isJustified: jcIsFullyJustified(paragraph.alignment),
    stretchLastLine: jcStretchesLastLine(paragraph.alignment),
    tabStops: [...paragraph.tabStops],
    hasRuby,
    hasEastAsianText,
  };
}

type RetainedTextBoxVerticalMode = NonNullable<TextBoxLayout['verticalMode']>;

function retainedTextBoxVerticalMode(value: string | null | undefined): RetainedTextBoxVerticalMode | undefined {
  return value === 'vert' || value === 'vert270' || value === 'eaVert' || value === 'mongolianVert'
    ? value : undefined;
}

function orientVerticalTextBoxParagraph(
  paragraph: ParagraphLayout,
  mode: RetainedTextBoxVerticalMode,
  innerBounds: LayoutRect,
  insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>,
): ParagraphLayout {
  const eastAsianUpright = mode === 'eaVert' || mode === 'mongolianVert';
  const lines = paragraph.lines.map((line) => {
    const rubyReservePt = mode === 'mongolianVert'
      ? line.placements.reduce((reserve, placement) => placement.kind === 'text' && placement.ruby
          ? Math.max(
              reserve,
              line.baselinePt - Math.min(
                line.baselinePt,
                ...placement.ruby.paintOps.map((operation) => operation.origin.yPt),
              ),
            )
          : reserve, 0)
      : 0;
    const mirroredBaselinePt = mode === 'mongolianVert'
      ? 2 * innerBounds.yPt + innerBounds.heightPt - line.baselinePt
        + insets.bottomPt - insets.leftPt + rubyReservePt
      : line.baselinePt;
    const deltaYPt = mirroredBaselinePt - line.baselinePt;
    const mirroredY = line.bounds.yPt + deltaYPt;
    const placements = line.placements.map((placement) => {
      if (placement.kind !== 'text') {
        return 'bounds' in placement && placement.bounds
          ? { ...placement, bounds: { ...placement.bounds, yPt: placement.bounds.yPt + deltaYPt } }
          : placement;
      }
      const paintOps = eastAsianUpright
        ? placement.clusters.map((cluster) => {
            const text = placement.text.slice(
              cluster.range.start - placement.range.start,
              cluster.range.end - placement.range.start,
            );
            const template = placement.paintOps.find((operation) =>
              operation.range.start <= cluster.range.start && operation.range.end >= cluster.range.end)
              ?? placement.paintOps[0]!;
            const upright = EAST_ASIAN_RE.test(text);
            return {
              ...template,
              text,
              range: cluster.range,
              offset: upright
                ? { xPt: cluster.offset.xPt + cluster.advancePt / 2, yPt: cluster.offset.yPt }
                : cluster.offset,
              glyphOrientation: upright ? 'upright' as const : 'sideways' as const,
            };
          })
        : placement.paintOps;
      return translatePlacementY({ ...placement, paintOps }, deltaYPt);
    });
    return {
      ...line,
      bounds: { ...line.bounds, yPt: mirroredY },
      baselinePt: line.baselinePt + deltaYPt,
      placements,
    };
  });
  return { ...paragraph, lines };
}

/** Acquires a DrawingML/WPS text body through the same paragraph measurement
 * and retained layout seam used by ordinary WordprocessingML paragraphs. */
export function acquireShapeTextBoxLayout(
  shape: Readonly<ShapeRun>,
  rect: LayoutRect,
  options: ShapeTextBoxAcquisitionOptions,
): TextBoxLayout | undefined {
  if (!shape.textBlocks?.length) return undefined;
  const source = options.source;
  const verticalMode = retainedTextBoxVerticalMode(shape.textVert);
  const contentBounds: LayoutRect = verticalMode ? {
    xPt: -rect.heightPt / 2,
    yPt: -rect.widthPt / 2,
    widthPt: rect.heightPt,
    heightPt: rect.widthPt,
  } : rect;
  const normalized = options.input ?? normalizeTextBoxInput(shape, {
    story: 'textbox',
    storyInstance: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
    path: [],
  });
  const insets = {
    topPt: shape.textInsetT ?? 0, rightPt: shape.textInsetR ?? 0,
    bottomPt: shape.textInsetB ?? 0, leftPt: shape.textInsetL ?? 0,
  };
  let yPt = contentBounds.yPt + insets.topPt;
  let previousInput: NormalizedTextBoxParagraphInput | null = null;
  let paragraphs = normalized.map((input, blockIndex) => {
    const textRuns: DocRun[] = input.runs.map((run) => shapeRunToDocRun({
      text: run.text,
      fontSizePt: run.fontSizePt,
      color: run.color?.slice(1) ?? null,
      fontFamily: run.fontFamily ?? null,
      fontFamilyEastAsia: run.fontFamilyEastAsia ?? null,
      bold: run.bold,
      italic: run.italic,
      ruby: run.ruby,
    }, shape.textVert));
    const availableImageWidthPt = Math.max(
      0,
      contentBounds.widthPt - insets.leftPt - insets.rightPt
        - input.indentLeftPt - input.indentRightPt - Math.max(0, input.indentFirstPt),
    );
    const imageNaturalWidthPt = verticalMode
      ? input.image?.heightPt ?? 0 : input.image?.widthPt ?? 0;
    const imageNaturalHeightPt = verticalMode
      ? input.image?.widthPt ?? 0 : input.image?.heightPt ?? 0;
    const imageScale = imageNaturalWidthPt > availableImageWidthPt && imageNaturalWidthPt > 0
      ? availableImageWidthPt / imageNaturalWidthPt
      : 1;
    const runs: DocRun[] = input.image ? [{
      type: 'image', imagePath: input.image.imagePath, mimeType: input.image.mimeType,
      ...(input.image.svgImagePath ? { svgImagePath: input.image.svgImagePath } : {}),
      widthPt: imageNaturalWidthPt > 0 ? imageNaturalWidthPt * imageScale : availableImageWidthPt,
      heightPt: imageNaturalHeightPt > 0
        ? imageNaturalHeightPt * imageScale : availableImageWidthPt,
      anchor: false,
    } as DocRun] : textRuns;
    const paragraph: ParagraphAcquisitionInput = {
      alignment: input.alignment,
      indentLeft: input.indentLeftPt,
      indentRight: input.indentRightPt,
      indentFirst: input.indentFirstPt,
      spaceBefore: input.spacing.beforePt,
      spaceAfter: input.spacing.afterPt,
      lineSpacing: input.lineSpacing,
      numbering: input.numbering ?? null,
      numberingMarkerShapeInput: input.numberingMarkerShapeInput,
      tabStops: [...input.tabStops],
      bidi: input.bidi,
      contextualSpacing: input.contextualSpacing,
      styleId: input.styleId,
      runs: runs as ParagraphAcquisitionInput['runs'],
    };
    const context = textBoxParagraphContext(options.context, paragraph);
    const gapPt = paragraphGapPt(
      previousInput,
      input,
      previousInput?.spacing.afterPt ?? 0,
      input.spacing.beforePt,
    );
    yPt += gapPt;
    const child = acquireParagraphLayout(paragraph, {
      id: `${options.id}:paragraph:${blockIndex}`,
      source: input.source,
      flowDomainId: `${options.flowDomainId}:textbox`,
      ordinaryFlow: true,
      context,
      placement: {
        startYPt: yPt,
        paragraphXPt: contentBounds.xPt + insets.leftPt,
        availableWidthPt: Math.max(0, contentBounds.widthPt - insets.leftPt - insets.rightPt),
        maximumYPt: contentBounds.yPt + contentBounds.heightPt - insets.bottomPt,
        // The shared flow fold above owns the complete inter-paragraph gap.
        // Paragraph acquisition therefore starts at the resolved content edge.
        suppressSpaceBefore: true,
      },
      measurer: options.measurer,
      environment: options.environment,
      exclusions: [],
    });
    yPt += child.advancePt - child.spacing.afterPt;
    previousInput = input;
    const innerBounds = {
      xPt: contentBounds.xPt + insets.leftPt,
      yPt: contentBounds.yPt + insets.topPt,
      widthPt: Math.max(0, contentBounds.widthPt - insets.leftPt - insets.rightPt),
      heightPt: Math.max(0, contentBounds.heightPt - insets.topPt - insets.bottomPt),
    };
    return verticalMode ? orientVerticalTextBoxParagraph(child, verticalMode, innerBounds, insets) : child;
  });
  const fittedExtentPt = Math.max(0, yPt - contentBounds.yPt + insets.bottomPt);
  const mayAutofit = shape.textAutofit === 'sp' && normalized.length > 0
    && (!verticalMode || normalized.every((input) => input.image === undefined));
  const effectiveRect = mayAutofit && Number.isFinite(fittedExtentPt) && fittedExtentPt > 0
    ? verticalMode
      ? { ...rect, widthPt: fittedExtentPt }
      : { ...rect, heightPt: fittedExtentPt }
    : rect;
  const effectiveContentBounds: LayoutRect = verticalMode ? {
    xPt: -effectiveRect.heightPt / 2,
    yPt: -effectiveRect.widthPt / 2,
    widthPt: effectiveRect.heightPt,
    heightPt: effectiveRect.widthPt,
  } : effectiveRect;
  if (verticalMode && effectiveRect.widthPt !== rect.widthPt && verticalMode !== 'mongolianVert') {
    const deltaYPt = effectiveContentBounds.yPt - contentBounds.yPt;
    paragraphs = paragraphs.map((paragraph) => translateParagraphY(paragraph, deltaYPt));
  }
  return deepFreezePlainData({
    kind: 'textbox', id: options.id, source: normalized[0]?.source ?? source,
    flowDomainId: `${options.flowDomainId}:textbox`, flowBounds: effectiveRect, inkBounds: effectiveRect,
    advancePt: 0, ordinaryFlow: false, paragraphs,
    writingMode: shape.textVert === 'vert270' ? 'vertical-lr' : shape.textVert ? 'vertical-rl' : 'horizontal-tb',
    insets,
    contentBounds: effectiveContentBounds,
    ...(verticalMode ? { verticalMode } : {}),
  });
}

/** Single acquisition seam from public/parser paragraph input to retained geometry.
 * Existing `measureParagraph` remains the sole segment and line-break owner. */
export function acquireParagraphLayout(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
): ParagraphLayout {
  const measurementPlacement = options.placement.wrap || options.exclusions.length === 0
    ? options.placement
    : {
        ...options.placement,
        wrap: createFloatWrapOracle(options.exclusions.map((exclusion, index) => ({
          kind: 'shape' as const,
          mode: exclusion.wrap === 'topAndBottom' ? 'topAndBottom' as const : 'square' as const,
          imageKey: exclusion.id,
          imageX: exclusion.bounds.xPt,
          imageY: exclusion.bounds.yPt,
          imageW: exclusion.bounds.widthPt,
          imageH: exclusion.bounds.heightPt,
          xLeft: exclusion.bounds.xPt,
          xRight: exclusion.bounds.xPt + exclusion.bounds.widthPt,
          yTop: exclusion.bounds.yPt,
          yBottom: exclusion.bounds.yPt + exclusion.bounds.heightPt,
          side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
          paraId: index, drawn: false,
        }))),
      };
  const measured = measureParagraph(
    paragraph,
    options.context,
    measurementPlacement,
    options.measurer,
    { ...options.environment, paragraphMarkShapeInput: paragraph.paragraphMarkShapeInput },
  );
  return paragraphLayoutFromMeasurement(paragraph, options, measured);
}

export interface RetainedFrameGroupAcquisition {
  readonly box: Readonly<{
    bounds: LayoutRect;
    exclusionBounds: LayoutRect;
    exclusionId: string;
  }>;
  readonly members: readonly Readonly<{
    paragraph: DocParagraph;
    fragment: ParagraphLayout;
    source: SourceRef;
  }>[];
}

export interface RetainedFrameGroupOptions {
  readonly contexts: readonly ParagraphLayoutContext[];
  readonly inputs: readonly ParagraphAcquisitionInput[];
  readonly borderEdges: readonly (ParagraphBorderEdges | undefined)[];
  readonly borderExtentsPt: readonly number[];
  readonly measurer: TextMeasurer;
  readonly environment: ParagraphMeasurementEnvironment;
  readonly containerShading?: string | null;
  readonly anchorFrames: NonNullable<ParagraphAcquisitionOptions['anchorFrames']>;
  /** C1 still owns legacy frame placement in renderer. This point-space seam
   * lets retained acquisition choose final content geometry without depending
   * on display scale or renderer state. */
  readonly maximumWidthPt: number;
  /** Identity of the owning measurement state. Acquisitions cannot outlive or
   * leak across the session whose resource/font facts produced their geometry. */
  readonly acquisitionSession: object;
  readonly placementSignature: string;
  readonly place: (
    contentWidthPt: number,
    contentHeightPt: number,
  ) => Readonly<{ bounds: LayoutRect; exclusionBounds: LayoutRect }>;
}

const retainedFrameGroupCache = new WeakMap<object, Map<string, RetainedFrameGroupAcquisition>>();

function frameFingerprintValue(value: unknown): unknown {
  if (value === undefined) return null;
  if (value instanceof Date) return { date: value.toISOString() };
  if (value instanceof Set) return {
    set: [...value].map(frameFingerprintValue)
      .sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right))),
  };
  if (value instanceof Map) return {
    map: [...value.entries()].map(([key, item]) => [
      frameFingerprintValue(key),
      frameFingerprintValue(item),
    ]).sort((left, right) => JSON.stringify(left[0]).localeCompare(JSON.stringify(right[0]))),
  };
  if (Array.isArray(value)) return value.map(frameFingerprintValue);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [
      key,
      frameFingerprintValue(item),
    ]));
  }
  return value;
}

/** Acquire a complete adjacent frame group into immutable final-width layouts. */
export function acquireRetainedFrameGroup(
  group: BodyFrameGroup,
  options: RetainedFrameGroupOptions,
): RetainedFrameGroupAcquisition {
  if (
    options.contexts.length !== group.members.length
    || options.inputs.length !== group.members.length
    || options.borderEdges.length !== group.members.length
    || options.borderExtentsPt.length !== group.members.length
  ) throw new Error('Frame acquisition metadata must align with every group member');
  if (!Number.isFinite(options.maximumWidthPt) || options.maximumWidthPt < 0) {
    throw new RangeError('Frame maximumWidthPt must be finite and non-negative');
  }
  let cache = retainedFrameGroupCache.get(options.acquisitionSession);
  if (!cache) {
    cache = new Map();
    retainedFrameGroupCache.set(options.acquisitionSession, cache);
  }
  const cacheKey = stableFingerprint('w:frame-acquisition', [
    group.id,
    options.placementSignature,
    options.maximumWidthPt,
    options.environment.pageIndex,
    options.environment.totalPages,
    options.environment.displayPageNumber ?? null,
    options.environment.pageNumberFormat ?? null,
    options.environment.currentDateMs ?? null,
    options.environment.documentHasEastAsianText,
    options.environment.layoutServices?.text.fingerprint ?? null,
    options.environment.layoutServices?.images.fingerprint ?? null,
    options.environment.layoutServices?.math.fingerprint ?? null,
    frameFingerprintValue(options.contexts),
    frameFingerprintValue(options.inputs),
    frameFingerprintValue(options.borderEdges),
    frameFingerprintValue(options.borderExtentsPt),
    options.containerShading ?? null,
    frameFingerprintValue(options.anchorFrames),
  ]);
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  const fp = group.framePr;
  const finalWidthPt = fp.w != null
    ? Math.max(0, fp.w)
    : Math.max(0, ...group.members.map((paragraph, index) =>
        measureParagraphIntrinsicWidth(
          paragraph,
          options.contexts[index]!,
          options.maximumWidthPt,
          options.measurer,
          options.environment,
          retainedNumberingPlan(
            options.inputs[index]!,
            options.contexts[index]!,
            options,
          ),
        )));
  const layoutWidthPt = Math.max(1, finalWidthPt);

  const acquireLocalStack = (): Readonly<{
    heightPt: number;
    members: RetainedFrameGroupAcquisition['members'];
  }> => {
    let cursorPt = 0;
    let previous: DocParagraph | null = null;
    let previousAfterPt = 0;
    let previousBorderExtentPt = 0;
    const retained: Array<RetainedFrameGroupAcquisition['members'][number]> = [];
    group.members.forEach((paragraph, memberIndex) => {
      const context = options.contexts[memberIndex]!;
      const gapPt = Math.max(
        paragraphGapPt(previous, paragraph, previousAfterPt, context.spaceBeforePt),
        previousBorderExtentPt,
      );
      const placement = {
        startYPt: cursorPt + gapPt,
        paragraphXPt: 0,
        availableWidthPt: layoutWidthPt,
        maximumYPt: Number.POSITIVE_INFINITY,
        suppressSpaceBefore: true,
      };
      const measured = measureParagraph(
        paragraph, context, placement, options.measurer, options.environment,
      );
      const borderExtentPt = options.borderExtentsPt[memberIndex] ?? 0;
      const trailingExtentPt = Math.max(measured.requestedSpaceAfterPt, borderExtentPt);
      const source: SourceRef = {
        story: 'body', storyInstance: 'body', path: [group.sourceIndices[memberIndex]!],
      };
      const fragment = paragraphLayoutFromMeasurement(
        options.inputs[memberIndex]!,
        {
          id: `body-frame:${group.id}:${memberIndex}`,
          source,
          flowDomainId: `body-frame:${group.id}`,
          ordinaryFlow: false,
          context,
          placement,
          measurer: options.measurer,
          environment: options.environment,
          exclusions: [],
          containerShading: options.containerShading,
          paragraphBorderEdges: options.borderEdges[memberIndex],
          trailingExtentPt,
          anchorFrames: options.anchorFrames,
        },
        measured,
      );
      retained.push({ paragraph, fragment, source });
      cursorPt = measured.contentEndYPt;
      previous = paragraph;
      previousAfterPt = measured.requestedSpaceAfterPt;
      previousBorderExtentPt = borderExtentPt;
    });
    return {
      heightPt: Math.max(
        0,
        cursorPt + Math.max(previousAfterPt, previousBorderExtentPt),
      ),
      members: retained,
    };
  };

  const local = acquireLocalStack();
  const placed = options.place(finalWidthPt, local.heightPt);
  const members = Object.freeze(local.members.map((member) => {
    const translated = translateParagraphLayout(member.fragment, {
      xPt: placed.bounds.xPt,
      yPt: placed.bounds.yPt,
    });
    const fragment = layoutParagraph(fp.hRule === 'exact' && fp.h != null
      ? { ...translated, clipBounds: placed.bounds }
      : translated);
    return Object.freeze({ ...member, fragment });
  }));
  const acquired = Object.freeze({
    box: Object.freeze({
      bounds: placed.bounds,
      exclusionBounds: placed.exclusionBounds,
      exclusionId: `frame:${group.id}`,
    }),
    members,
  });
  cache.set(cacheKey, acquired);
  return acquired;
}

/** Projects an already-acquired line partition without measuring a second time. */
export function paragraphLayoutFromMeasurement(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  measured: MeasuredParagraph,
): ParagraphLayout {
  const planningContext = options.continuesFromPrevious
    ? { ...options.context, firstIndentPt: 0 }
    : options.context;
  const paragraphXPt = options.placement.paragraphXPt + planningContext.physicalIndentLeftPt;
  const availableWidthPt = options.placement.availableWidthPt
    - planningContext.physicalIndentLeftPt - planningContext.physicalIndentRightPt;
  const occurrences = logicalOccurrenceMap(paragraph, measured);
  const numberingPlan = options.continuesFromPrevious
    ? undefined
    : retainedNumberingPlan(paragraph, planningContext, options);
  let lines = planMeasuredLines(
    measured, paragraph, paragraphXPt, availableWidthPt, options.source, planningContext,
    occurrences, numberingPlan, options.environment.layoutServices?.text,
  );
  if (
    numberingPlan
    && measured.markOnly
    && lines.length === 0
    && (numberingPlan.markerText !== '' || paragraph.numbering?.picBulletImagePath)
  ) {
    lines = [numberingMarkerHostLine(
      measured,
      paragraph,
      paragraphXPt,
      availableWidthPt,
      planningContext,
    )];
  }
  const resources: InlineResourceLayout[] = [];
  const drawings: DrawingLayout[] = [];
  const textBoxes: TextBoxLayout[] = [];
  const anchorResults: AnchorFrameResult[] = [];
  const anchorExclusions: WrapExclusion[] = [];
  const events = paragraph.runs
    .map((run, runIndex) => run.type === 'break'
      ? { kind: 'break' as const, breakKind: run.breakType, offset: occurrences.runStarts[runIndex] ?? 0 }
      : undefined)
    .filter((event): event is NonNullable<typeof event> => event !== undefined);
  const payloadsByOccurrence = new Map<
    string,
    Array<Readonly<{ run: AnchoredPayloadRun; runIndex: number }>>
  >();
  paragraph.runs.forEach((run, runIndex) => {
    if (!anchoredPayloadRun(run)) return;
    const payloads = payloadsByOccurrence.get(run.anchorAcquisitionInput!.occurrenceId) ?? [];
    payloads.push({ run, runIndex });
    payloadsByOccurrence.set(run.anchorAcquisitionInput!.occurrenceId, payloads);
  });
  for (const [occurrenceId, payloads] of payloadsByOccurrence) {
    const acquired = acquireAnchorOccurrence(
      occurrenceId,
      payloads,
      lines,
      paragraph,
      options,
      measured.contentEndYPt - options.placement.startYPt,
    );
    if (!acquired) continue;
    anchorResults.push(acquired.result);
    if (!acquired.drawing) continue;
    drawings.push(acquired.drawing);
    textBoxes.push(...acquired.textBoxes);
    if (acquired.exclusion) anchorExclusions.push(acquired.exclusion);
    const hostLine = lines[acquired.hostLineIndex];
    if (hostLine) {
      lines = lines.map((line, lineIndex) => lineIndex === acquired.hostLineIndex ? {
        ...line,
        placements: [...line.placements, {
          kind: 'drawing', range: acquired.hostRange,
          drawingId: acquired.drawing!.id,
          bounds: acquired.drawing!.inkBounds,
          advancePt: 0,
        }],
      } : line);
    }
  }
  if (numberingPlan && lines[0]) {
    const markerPlacements = numberingMarkerPlacements(
      numberingPlan, paragraph, options.context, paragraphXPt, availableWidthPt, lines[0],
    );
    if (markerPlacements.length > 0) {
      lines = [{ ...lines[0], placements: [...markerPlacements, ...lines[0].placements] }, ...lines.slice(1)];
    }
  }
  paragraph.runs.forEach((run, runIndex) => {
    const source = runSource(options.source, runIndex);
    if (run.type === 'image') resources.push({
      kind: 'image', resourceKey: imageResourceKey(source, run.imagePath),
      intrinsicSize: { widthPt: run.widthPt, heightPt: run.heightPt },
    });
    if (run.type === 'chart') resources.push({
      kind: 'chart', resourceKey: chartResourceKey(source),
      intrinsicSize: { widthPt: run.widthPt, heightPt: run.heightPt },
    });
    if (run.type === 'math') resources.push({
      kind: 'math', resourceKey: (run as { resourceKey?: string }).resourceKey ?? stableFingerprint('math-resource', source),
      intrinsicSize: {
        widthPt: lines.flatMap((line) => line.placements).find((placement) =>
          placement.kind === 'resource' && placement.resourceKind === 'math')?.bounds?.widthPt ?? 0,
        heightPt: run.fontSize,
      },
    });
    if ((run.type === 'image' || run.type === 'chart') && !options.continuesFromPrevious) {
      const drawing = publicAnchoredResourceDrawing(
        run,
        options,
        runIndex,
        publicAnchorReferenceFrames(
          lines,
          options,
          measured.contentEndYPt - options.placement.startYPt,
        ),
      );
      if (drawing) {
        drawings.push(drawing);
        const firstLine = lines[0];
        if (firstLine) lines = [{
          ...firstLine,
          placements: [...firstLine.placements, {
            kind: 'drawing',
            range: {
              start: occurrences.runStarts[runIndex] ?? 0,
              end: (occurrences.runStarts[runIndex] ?? 0) + (occurrences.runLengths[runIndex] ?? 1),
            },
            drawingId: drawing.id, bounds: drawing.inkBounds, advancePt: 0,
          }],
        }, ...lines.slice(1)];
      }
    }
    if (run.type === 'shape' && !run.anchorAcquisitionInput && !options.continuesFromPrevious) {
      // Resolve the point-space box once. Shape panel paint, retained textbox
      // flow, and the line's drawing placement must own identical geometry.
      const authoredShapeRect = resolvedShapeLayoutRect(run, options);
      const textBoxId = `${options.id}:textbox:${runIndex}`;
      const textBox = acquireShapeTextBoxLayout(run, authoredShapeRect, {
        id: textBoxId,
        source,
        flowDomainId: options.flowDomainId,
        context: options.context,
        measurer: options.measurer,
        environment: options.environment,
        input: run.textBoxInput,
      });
      const shapeRect = textBox?.flowBounds ?? authoredShapeRect;
      let drawing = drawingForShape(
        run,
        shapeRect,
        options,
        runIndex,
        publicAnchorReferenceFrames(
          lines,
          options,
          measured.contentEndYPt - options.placement.startYPt,
        ),
      );
      if (textBox) {
        textBoxes.push(textBox);
        drawing = { ...drawing, textBoxIds: [textBoxId] };
      }
      drawings.push(drawing);
      const firstLine = lines[0];
      if (firstLine) {
        lines = [{
          ...firstLine,
          placements: [...firstLine.placements, {
          kind: 'drawing',
          range: {
            start: occurrences.runStarts[runIndex] ?? 0,
            end: (occurrences.runStarts[runIndex] ?? 0) + (occurrences.runLengths[runIndex] ?? 1),
          },
          drawingId: drawing.id,
          bounds: drawing.inkBounds, advancePt: 0,
          }],
        }, ...lines.slice(1)];
      }
    }
  });
  if (paragraph.numbering?.picBulletImagePath && !options.continuesFromPrevious) resources.push({
    kind: 'picture-bullet',
    resourceKey: imageResourceKey(options.source, paragraph.numbering.picBulletImagePath),
    intrinsicSize: {
      widthPt: paragraph.numbering.picBulletWidthPt
        ?? paragraph.numberingMarkerShapeInput?.fontSizePt ?? 0,
      heightPt: paragraph.numbering.picBulletHeightPt
        ?? paragraph.numberingMarkerShapeInput?.fontSizePt ?? 0,
    },
  });
  if (paragraph.numbering?.picBulletImagePath && lines[0] && !options.continuesFromPrevious) {
    if (!numberingPlan) {
      throw new Error('Picture-bullet acquisition requires resolved marker font geometry');
    }
    const widthPt = paragraph.numbering.picBulletWidthPt ?? numberingPlan.markerWidthPt;
    const heightPt = paragraph.numbering.picBulletHeightPt
      ?? paragraph.numberingMarkerShapeInput?.fontSizePt;
    if (heightPt === undefined) {
      throw new Error('Picture-bullet acquisition requires resolved marker height');
    }
    const markerLeftPt = numberingMarkerPhysicalLeft({
      baseRtl: options.context.baseRtl,
      paragraphXPt,
      availableWidthPt,
      authoredFirstIndentPt: paragraph.indentFirst,
      markerShiftPt: numberingPlan.markerShiftPt,
      markerWidthPt: widthPt,
    });
    lines = [{
      ...lines[0],
      placements: [{
      kind: 'resource', resourceKind: 'picture-bullet',
      range: { start: -1, end: 0 },
      resourceKey: imageResourceKey(options.source, paragraph.numbering.picBulletImagePath),
      bounds: {
        xPt: markerLeftPt,
        yPt: lines[0].baselinePt - heightPt,
        widthPt, heightPt,
      },
      advancePt: 0,
      }, ...lines[0].placements],
    }, ...lines.slice(1)];
  }
  lines = retainEffectiveTextBackground(
    lines,
    paragraph.shading,
    options.containerShading,
  );
  const contentHeightPt = measured.contentEndYPt - measured.contentStartYPt;
  const paragraphBorderEdges = options.paragraphBorderEdges ?? {
    top: 'top' as const,
    bottom: 'bottom' as const,
  };
  const borderBounds = paragraphDecorationBox(
    paragraph,
    lines,
    paragraphXPt,
    availableWidthPt,
    measured.contentStartYPt,
    contentHeightPt,
    paragraphBorderEdges,
  );
  const borderEntries: Array<readonly [
    NonNullable<import('./types.js').BorderSegment['edge']>,
    NonNullable<ParagraphAcquisitionInput['borders']>['top'],
  ]> = paragraph.borders ? [
    ...(paragraphBorderEdges.top === 'none'
      ? [] : [[paragraphBorderEdges.top, paragraph.borders[paragraphBorderEdges.top]] as const]),
    ['right', paragraph.borders.right],
    ...(paragraphBorderEdges.bottom === 'none'
      ? [] : [['bottom', paragraph.borders.bottom] as const]),
    ['left', paragraph.borders.left],
  ] : [];
  const borderSegments = paragraph.borders
    ? borderEntries.flatMap(([side, edge]) => {
        if (!visibleParagraphBorder(edge)) return [];
        const horizontal = side === 'top' || side === 'between' || side === 'bottom';
        const atEnd = side === 'right' || side === 'bottom';
        const coordinate = horizontal
          ? borderBounds.yPt + (atEnd ? borderBounds.heightPt : 0)
          : borderBounds.xPt + (atEnd ? borderBounds.widthPt : 0);
        return [{
          edge: side,
          from: horizontal
            ? { xPt: borderBounds.xPt, yPt: coordinate }
            : { xPt: coordinate, yPt: borderBounds.yPt },
          to: horizontal
            ? { xPt: borderBounds.xPt + borderBounds.widthPt, yPt: coordinate }
            : { xPt: coordinate, yPt: borderBounds.yPt + borderBounds.heightPt },
          color: edge.color ? `#${edge.color}` : '#000000',
          widthPt: edge.width,
          ...retainedBorderTreatment(edge.style, edge.width),
        }];
      })
    : [];
  const trailingExtentPt = options.trailingExtentPt ?? measured.requestedSpaceAfterPt;
  return layoutParagraph({
    kind: 'paragraph', id: options.id, source: options.source,
    flowDomainId: options.flowDomainId, ordinaryFlow: options.ordinaryFlow,
    ...(paragraph.styleId !== undefined ? { styleId: paragraph.styleId } : {}),
    flowBounds: {
      xPt: options.placement.paragraphXPt, yPt: options.placement.startYPt,
      widthPt: options.placement.availableWidthPt,
      heightPt: measured.contentEndYPt - options.placement.startYPt + trailingExtentPt,
    },
    inkBounds: {
      ...(paragraph.shading || paragraph.borders
        ? borderBounds
        : {
            xPt: paragraphXPt,
            yPt: measured.contentStartYPt,
            widthPt: Math.max(0, ...lines.map((line) => line.bounds.widthPt)),
            heightPt: contentHeightPt,
          }),
    },
    spacing: {
      beforePt: options.placement.suppressSpaceBefore ? 0 : measured.requestedSpaceBeforePt,
      afterPt: trailingExtentPt,
    },
    contextualSpacing: paragraph.contextualSpacing ?? false,
    lines, borders: borderSegments,
    shading: paragraph.shading ? { color: `#${paragraph.shading}` } : undefined,
    resources, drawings, textBoxes, events,
    exclusions: [...options.exclusions, ...anchorExclusions],
    ...(anchorResults.length ? { anchorFrames: anchorResults } : {}),
    paragraphMark: measured.markOnly ? {
      hidden: paragraph.markVanish === true,
      bounds: { xPt: paragraphXPt, yPt: measured.contentStartYPt, widthPt: 0, heightPt: contentHeightPt },
    } : undefined,
  });
}

const translatePointY = (point: PointPt, yPt: number): PointPt =>
  translatePoint(point, { xPt: 0, yPt });
const translateRectY = (rect: LayoutRect, yPt: number): LayoutRect =>
  translateRect(rect, { xPt: 0, yPt });
const translateDrawingY = (drawing: DrawingLayout, yPt: number): DrawingLayout =>
  translateDrawing(drawing, { xPt: 0, yPt });
const translatePlacementY = (placement: ParagraphPlacement, yPt: number): ParagraphPlacement =>
  translatePlacement(placement, { xPt: 0, yPt });
const translateLineY = (line: LineLayout, yPt: number): LineLayout =>
  translateLine(line, { xPt: 0, yPt });
const translateParagraphY = (paragraph: ParagraphLayout, yPt: number): ParagraphLayout =>
  translateParagraphLayout(paragraph, { xPt: 0, yPt });
const translateTextBoxY = (textBox: TextBoxLayout, yPt: number): TextBoxLayout =>
  translateTextBox(textBox, { xPt: 0, yPt });

function sliceParagraphDecoration(
  acquired: ParagraphLayout,
  selected: readonly LineLayout[],
  deltaYPt: number,
  continuation: NonNullable<ParagraphLayout['continuation']>,
): Readonly<{ box: LayoutRect; borders: ParagraphLayout['borders'] }> | null {
  if (!acquired.shading && acquired.borders.length === 0) return null;
  const first = selected[0];
  const last = selected.at(-1);
  if (!first || !last) return {
    box: translateRectY(acquired.inkBounds, deltaYPt),
    borders: [],
  };
  const decorationTopPt = acquired.inkBounds.yPt;
  const decorationBottomPt = decorationTopPt + acquired.inkBounds.heightPt;
  const ownedTopPt = continuation.continuesFromPrevious
    ? Math.max(decorationTopPt, first.bounds.yPt)
    : decorationTopPt;
  const ownedBottomPt = continuation.continuesOnNext
    ? Math.min(decorationBottomPt, last.bounds.yPt + last.advancePt)
    : decorationBottomPt;
  const box: LayoutRect = {
    xPt: acquired.inkBounds.xPt,
    yPt: ownedTopPt + deltaYPt,
    widthPt: acquired.inkBounds.widthPt,
    heightPt: Math.max(0, ownedBottomPt - ownedTopPt),
  };
  const leftPt = box.xPt;
  const rightPt = leftPt + box.widthPt;
  const topPt = box.yPt;
  const bottomPt = topPt + box.heightPt;
  const borders = acquired.borders.flatMap((border) => {
    if ((border.edge === 'top' || border.edge === 'between')
      && continuation.continuesFromPrevious) return [];
    if (border.edge === 'bottom' && continuation.continuesOnNext) return [];
    if (border.edge === 'top' || border.edge === 'between') return [{
      ...border,
      from: { xPt: leftPt, yPt: topPt },
      to: { xPt: rightPt, yPt: topPt },
    }];
    if (border.edge === 'bottom') return [{
      ...border,
      from: { xPt: leftPt, yPt: bottomPt },
      to: { xPt: rightPt, yPt: bottomPt },
    }];
    if (border.edge === 'left') return [{
      ...border,
      from: { xPt: leftPt, yPt: topPt },
      to: { xPt: leftPt, yPt: bottomPt },
    }];
    if (border.edge === 'right') return [{
      ...border,
      from: { xPt: rightPt, yPt: topPt },
      to: { xPt: rightPt, yPt: bottomPt },
    }];
    return [{
      ...border,
      from: translatePointY(border.from, deltaYPt),
      to: translatePointY(border.to, deltaYPt),
    }];
  });
  return { box, borders };
}

/** Produces a continuation without reacquiring text or touching a measurer. */
export function sliceParagraphLayout(
  acquired: ParagraphLayout,
  continuation: NonNullable<ParagraphLayout['continuation']>,
  id = `${acquired.id}:${continuation.lineStart}-${continuation.lineEnd}`,
): ParagraphLayout {
  const selected = acquired.lines.slice(continuation.lineStart, continuation.lineEnd);
  const first = selected[0];
  const last = selected.at(-1);
  // A continuation is placed in a new flow slice. Preserve the acquired x/range
  // geometry, but make its first retained line own the same local y origin as
  // the original paragraph so placement translates one coherent coordinate
  // space instead of carrying the preceding page's consumed line offset.
  const deltaYPt = continuation.continuesFromPrevious && first
    ? acquired.flowBounds.yPt - first.bounds.yPt
    : 0;
  const rebasedSelected = deltaYPt === 0
    ? selected
    : selected.map((line) => translateLineY(line, deltaYPt));
  const rebasedFirst = rebasedSelected[0];
  const rebasedLast = rebasedSelected.at(-1);
  const rebasedLines = acquired.lines.map((line, index) =>
    index >= continuation.lineStart && index < continuation.lineEnd
      ? rebasedSelected[index - continuation.lineStart]!
      : line);
  const lineInkBounds = rebasedFirst && rebasedLast ? {
    xPt: Math.min(...rebasedSelected.map((line) => line.bounds.xPt)),
    yPt: rebasedFirst.bounds.yPt,
    widthPt: Math.max(...rebasedSelected.map((line) => line.bounds.xPt + line.bounds.widthPt))
      - Math.min(...rebasedSelected.map((line) => line.bounds.xPt)),
    heightPt: rebasedLast.bounds.yPt + rebasedLast.bounds.heightPt - rebasedFirst.bounds.yPt,
  } : acquired.inkBounds;
  const decoration = sliceParagraphDecoration(
    acquired,
    selected,
    deltaYPt,
    continuation,
  );
  const drawingIds = new Set(selected.flatMap((line) => line.placements.flatMap((placement) =>
    placement.kind === 'drawing' ? [placement.drawingId] : [])));
  const drawings = acquired.drawings
    .filter((drawing) => drawingIds.has(drawing.id))
    .map((drawing) => drawing.anchorLayer?.verticalOwnership === 'page'
      ? drawing : translateDrawingY(drawing, deltaYPt));
  const resourceKeys = new Set(selected.flatMap((line) => line.placements.flatMap((placement) =>
    placement.kind === 'resource' ? [placement.resourceKey] : [])));
  for (const drawing of drawings) {
    for (const command of drawing.commands) {
      if (command.kind === 'resource') resourceKeys.add(command.resourceKey);
    }
  }
  const textBoxIds = new Set(drawings.flatMap((drawing) => [
    drawing.id.replace(':drawing:', ':textbox:'),
    ...(drawing.textBoxIds ?? []),
  ]));
  const pageOwnedTextBoxIds = new Set(drawings
    .filter((drawing) => drawing.anchorLayer?.verticalOwnership === 'page')
    .flatMap((drawing) => drawing.textBoxIds ?? []));
  const drawingSourceKeys = new Set(drawings.map((drawing) =>
    stableFingerprint('source-occurrence', drawing.source)));
  const lineRangeStart = first?.range.start;
  const lineRangeEnd = last?.range.end;
  return layoutParagraph({
    ...acquired,
    kind: 'paragraph', id,
    lines: rebasedLines,
    flowBounds: {
      ...acquired.flowBounds,
      yPt: acquired.flowBounds.yPt,
    },
    ...(acquired.clipBounds
      ? { clipBounds: translateRectY(acquired.clipBounds, deltaYPt) }
      : {}),
    spacing: {
      beforePt: continuation.continuesFromPrevious ? 0 : acquired.spacing.beforePt,
      afterPt: continuation.continuesOnNext ? 0 : acquired.spacing.afterPt,
    },
    inkBounds: decoration?.box ?? lineInkBounds,
    borders: decoration?.borders ?? acquired.borders
      .map((border) => ({
        ...border,
        from: translatePointY(border.from, deltaYPt),
        to: translatePointY(border.to, deltaYPt),
      })),
    resources: acquired.resources.filter((resource) => resourceKeys.has(resource.resourceKey)),
    drawings,
    textBoxes: acquired.textBoxes
      .filter((textBox) =>
        textBoxIds.has(textBox.id)
        || drawingSourceKeys.has(stableFingerprint('source-occurrence', textBox.source)))
      .map((textBox) => pageOwnedTextBoxIds.has(textBox.id)
        ? textBox : translateTextBoxY(textBox, deltaYPt)),
    events: lineRangeStart === undefined || lineRangeEnd === undefined
      ? []
      : acquired.events.filter((event) => event.offset >= lineRangeStart
        && (event.offset < lineRangeEnd
          || (!continuation.continuesOnNext && event.offset === lineRangeEnd))),
    exclusions: acquired.exclusions.map((exclusion) => ({
      ...exclusion,
      bounds: exclusion.verticalOwnership === 'page'
        ? exclusion.bounds : translateRectY(exclusion.bounds, deltaYPt),
      polygon: exclusion.verticalOwnership === 'page'
        ? exclusion.polygon
        : exclusion.polygon.map((point) => translatePointY(point, deltaYPt)),
    })),
    ...(continuation.continuesOnNext
      ? { paragraphMark: undefined }
      : acquired.paragraphMark
        ? { paragraphMark: {
            ...acquired.paragraphMark,
            bounds: translateRectY(acquired.paragraphMark.bounds, deltaYPt),
          } }
        : {}),
    continuation,
  });
}
