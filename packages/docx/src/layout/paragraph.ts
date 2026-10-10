import { revisionIsOmitted } from './revision-visibility.js';
import { wordKerningApplies } from './line-compatibility.js';
import type { LineGapPlan } from '../line-breaker/line-gaps.js';
import { distributeLineSlack as distributeProportionalSlack } from '@silurus/ooxml-core';
import { specifiedTextLineMetrics, specifiedTextParagraphIsHomogeneous } from './specified-line-spacing.js';
import { wordTextBoxVerticalMode } from './compatibility.js';
import { autoContrastColor, canvasFontString, createCanvasFontRoute } from '@silurus/ooxml-core';
import {
  effectiveParagraphTabStops,
  paragraphGridRightAdjustmentPt,
  type ParagraphLayoutContext,
} from '../layout-context.js';
import {
  measureParagraph,
  paragraphCharacterGrid,
  type MeasuredParagraph,
  type ParagraphMeasurementEnvironment,
  type ParagraphPlacement as MeasurementPlacement,
  type TextMeasurer,
} from '../paragraph-measure.js';
import { createFloatWrapOracle } from './float-wrap-oracle.js';
import { firstFreeGapLeftPt, polygonMeetsRect, type FloatRect } from './float-wrap.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import type {
  LayoutImageSeg,
  LayoutLine,
  LayoutMathSeg,
  LayoutTabSeg,
  LayoutTextSeg,
  DocGridCtx,
} from '../line-layout.js';
import {
  effectiveCharacterSpacingPt,
  getDefaultFontSize,
  segLetterSpacingPx,
  widthBalanceSpaceAdjustmentForTextPt,
} from '../line-layout.js';
import {
  calcEffectiveFontPx, EAST_ASIAN_RE, independentTextShapeRequest,
  replaceTextShapeRequest, sliceTextShapeRequest, sliceSemanticSlotSpans, shapeRunToDocRun,
} from './text.js';
import { eastAsianUprightPaintOps } from './vertical-glyph-orientation.js';
import { wordTrackChangeDecoration } from './paint-compatibility.js';
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
  type DistributeResult,
  type SegStretch,
} from '../text-distribute.js';
import { computeKashidaDistribution, type KashidaLevel } from '../kashida-justify.js';
import { imageResourceKey } from './source-key.js';
import { stableFingerprint } from './fingerprint.js';
import { chartResourceKey as canonicalChartResourceKey } from './source-key.js';
import {
  planShapeDrawing,
  type ShapeDrawingPlanResult,
} from './shape-drawing-plan.js';
import {
  normalizeTextBoxInput,
  type NormalizedTextBoxParagraphInput,
  type TextBoxAcquisitionInput,
} from './textbox-input.js';
import {
  numberingMarkerPhysicalLeft,
  resolveNumberingMarkerGeometry,
  shapeNumberingMarkerText,
} from './numbering-marker.js';
import { deepFreezePlainData, deepFreezePlainDataWithFrozenAliases } from './plain-data.js';
import { retainedBorderTreatment } from './border-treatment.js';
import type { ParagraphBorderEdges } from './paragraph-border-adjacency.js';
import {
  centeredLeaderGlyphOrigins,
  groupedRunBorderFragments,
  retainedEmphasisGlyphs,
  retainedTextDecorations,
  retainedWavePath,
  rubyPaintOperations,
  type RetainedEmphasisClusterInk,
  type RetainedEmphasisMarkInput,
} from './retained-typography.js';
import type { RunTypographyAcquisitionInput } from './typography-input.js';
import { alignedAnchorPlacement, resolveAnchorFrame, type AnchorFrameRect, type AnchorReferenceFramesInput, type AnchorFrameResult } from './anchor-frame.js';
import { paragraphGapPt } from './paragraph-spacing.js';
import {
  translateDrawing,
  translateLine,
  translateParagraphLayout,
  translatePlacement,
  translatePoint,
  translateRect,
  translateTextBox,
  translateTextBoxStoryTable,
} from './retained-geometry-translation.js';
export { translateParagraphLayout } from './retained-geometry-translation.js';
import { paginationFieldDependency } from './pagination-fields.js';
import {
  ExactConvergenceError,
  convergeExactState,
} from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import {
  commitParagraphWrapRegistry,
  createParagraphWrapRegistry,
} from './paragraph-wrap-registry.js';
import {
  paragraphAcquisitionCacheOf,
  type ParagraphAcquisitionRuntimeCache,
} from './runtime-state.js';
import {
  wordLayoutInCellOwnsRowContainment,
  wordPreservesLowerLayerSameParagraphComposition,
  wordTextBoxVisibleAnchorExtentPt,
  wordGridPictureLineOriginPt,
} from './anchor-compatibility.js';
import {
  wordRunVerticalAlignRaisePt,
  wordSnapToCharsEastAsianCellCount,
} from './line-compatibility.js';
import { wordFramePositionExtendsLineBox } from './body-pagination-compatibility.js';
import {
  resolveFloatPlacement,
  type FloatPlacementParticipant,
} from './floats.js';
import { unionLayoutRects } from './rect-union.js';
import type { LayoutTranslation } from './retained-geometry-translation.js';
import {
  measureParagraphIntrinsicWidth,
  type BodyFrameGroup,
} from './frame.js';
import {
  createSectionRegionCoordinateSpace,
  transformPoint,
  transformRect,
  transformRectEdges,
  uprightPhysicalExtent,
  uprightResourceOrientation,
} from './coordinate-space.js';
import { composeAffine, inverseMapAffinePoint, translationAffine } from './affine.js';
import {
  solveStoryPageFrames,
  storyPageFramesThrough,
  type StoryPageFrames,
} from './story-page-frames.js';
export {
  bodyFrameGroupFor,
  bodyParagraphBorderEdgesFor,
  collectBodyFrameGroups,
  prepareBodyFrameMetadata,
} from './frame.js';
export type { BodyFrameGroup } from './frame.js';
import type { ParagraphAcquisitionInput, ParagraphAcquisitionRun, ParagraphLayoutSource } from './text.js';
import type { VerticalGlyphMeasurementService } from './measurement-capabilities.js';
import type {
  DrawingLayout,
  DrawingPaintCommand,
  DrawingMLCollisionEntryPt,
  AcquiredParagraphLayoutInput,
  InlineResourceLayout,
  LineLayout,
  LayoutDiagnostic,
  LayoutRect,
  Matrix2DData,
  NumberingMarkerShapeInput,
  ParagraphLayout,
  ParagraphPlacement,
  PointPt,
  SourceRef,
  StoryLayout,
  FlowContainer,
  TextBoxLayout,
  TextClusterLayout,
  TextDecorationLayout,
  TextPaintOp,
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
    /** Physical advance from the segment origin through the final glyph. Word
     * retains later grid slack for layout but excludes it from a terminal underline. */
    decorationTerminalAdvancePt?: number;
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
    /** Selected face's font box, retained independently of authored decoration
     * so application overlays can use the same character-height rectangle. */
    selectedFaceFontBox?: Readonly<{ ascentPt: number; descentPt: number }>;
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
    /** Stroke colour override (markup-view revision strikes are painted in
     *  the stable author colour, not the run text colour). */
    color?: string;
  }>;
  readonly emphasis?: Readonly<{
    authored: string;
    glyph: string;
    mark: RetainedEmphasisMarkInput;
    clusterInk: readonly RetainedEmphasisClusterInk[];
  }>;
}

function retainedTypographyInput(
  run: import('./text.js').ParagraphLayoutRun | undefined,
): RunTypographyAcquisitionInput | undefined {
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
  readonly underline?: Readonly<{
    readonly base: RetainedInkMetric;
    readonly authoredStyle?: string;
    readonly color: string;
    readonly probe: RetainedInkMetric;
  }>;
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
  readonly sourceRunIndex?: number;
  readonly measuredWidthPt: number;
  readonly resourceKey: string;
  readonly resourceKind: import('./types.js').InlineResourceKind;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly topOffsetPt: number;
  readonly orientation?: import('./types.js').UprightResourceOrientation;
}

export interface MeasuredUnavailableResourcePlanSegment {
  readonly kind: 'unavailable-resource';
  readonly range: import('./types.js').TextRange;
  readonly measuredWidthPt: number;
  readonly resourceKind: 'image' | 'chart';
  readonly widthPt: number;
  readonly heightPt: number;
  readonly topOffsetPt: number;
  readonly drawingId: string;
}

export interface MeasuredInlineDrawingPlanSegment {
  readonly kind: 'inline-drawing';
  readonly range: import('./types.js').TextRange;
  readonly measuredWidthPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly topOffsetPt: number;
  readonly drawingId: string;
}

export interface MeasuredAnchorHostPlanSegment {
  readonly noteReference?: TextPlacement['noteReference'];
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
  | MeasuredUnavailableResourcePlanSegment
  | MeasuredInlineDrawingPlanSegment
  | MeasuredAnchorHostPlanSegment;

export interface MeasuredLinePlanInput {
  readonly justifiedCompressionPt?: number;
  readonly gapPlan?: LineGapPlan;
  readonly range: import('./types.js').TextRange;
  readonly topPt: number;
  readonly baselinePt: number;
  readonly advancePt: number;
  readonly xOffsetPt: number;
  readonly availableWidthPt: number;
  /** Margin extension a tab cell added to this line's band (§17.3.1.37). */
  readonly marginExtensionPt?: number;
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
  /** Exact lines paint run shading through the authored line box. */
  readonly exactLineSpacing?: boolean;
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
    ...replaceTextShapeRequest(segment.textShapeRequest, text),
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
  if ([...distribution.perSeg.values()].some(stretch => stretch.gapDeltas !== undefined)) {
    // The fit enumerator retained only atomic space opportunities, including
    // combining-mark seams. Suppressing one here would change its arithmetic.
    for (const [index, stretch] of distribution.perSeg) {
      const segment = segments[index];
      if (segment?.kind !== 'text') continue;
      const starts = new Set(segment.clusters.map(cluster => cluster.range.start - segment.range.start));
      const chars = [...segment.text];
      let utf16 = 0;
      const offsets = chars.map(character => { const start = utf16; utf16 += character.length; return start; });
      if (stretch.splitBefore.some(cut => !starts.has(offsets[cut]))) {
        throw new Error('Proportional space justification requires a retained cluster boundary');
      }
    }
    return distribution;
  }
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

/** Merge sorted gap, cluster and paint-op boundaries. Every cursor only moves
 * forward: O(code points + clusters + gaps + operations), including long runs
 * of consecutive spaces. Deltas are owned by the accepted line's gap plan. */
function retainedTextGeometry(
  segment: MeasuredTextPlanSegment,
  stretch: SegStretch | undefined,
  perGapPt: number,
): Readonly<{ clusters: readonly TextClusterLayout[]; paintOps: readonly TextPaintOp[] }> {
  if (!stretch || stretch.splitBefore.length === 0) {
    return { clusters: segment.clusters, paintOps: segment.basePaintOps };
  }
  const cuts = stretch.splitBefore;
  const deltas = stretch.gapDeltas;
  if (deltas && deltas.length !== cuts.length) {
    throw new Error('Internal paragraph justification has incomplete gap deltas');
  }
  const absoluteCuts: number[] = [];
  let cpIndex = 0;
  let utf16 = segment.range.start;
  let cutIndex = 0;
  for (const character of segment.text) {
    if (cuts[cutIndex] === cpIndex) {
      absoluteCuts.push(utf16);
      cutIndex += 1;
    }
    cpIndex += 1;
    utf16 += character.length;
  }
  if (cutIndex !== cuts.length || cuts.some((cut, index) =>
    cut <= 0 || (index > 0 && cut <= cuts[index - 1]!))) {
    throw new Error('Internal paragraph justification contains an invalid code-point cut');
  }

  let gapCursor = 0;
  let cumulativePt = 0;
  const clusters = segment.clusters.map(cluster => {
    const cut = absoluteCuts[gapCursor];
    if (cut !== undefined && cut < cluster.range.start) {
      throw new Error('Internal paragraph justification must split at shaped cluster boundaries');
    }
    if (cut === cluster.range.start) {
      cumulativePt += deltas?.[gapCursor] ?? perGapPt;
      gapCursor += 1;
    }
    return { ...cluster, offset: { ...cluster.offset, xPt: cluster.offset.xPt + cumulativePt } };
  });
  if (gapCursor !== absoluteCuts.length) {
    throw new Error('Internal paragraph justification must split at shaped cluster boundaries');
  }
  let cursor = segment.range.start;
  for (const operation of segment.basePaintOps) {
    if (operation.range.start !== cursor || operation.range.end <= cursor) {
      throw new Error('Internal paragraph justification has incomplete retained paint operations');
    }
    cursor = operation.range.end;
  }
  if (cursor !== segment.range.end || segment.basePaintOps.length === 0) {
    throw new Error('Internal paragraph justification has incomplete retained paint operations');
  }
  const baseOp = segment.basePaintOps[0]!;
  if (segment.basePaintOps.length === 1 && cuts.length === cpIndex - 1
    && cuts.every((cut, index) => cut === index + 1)
    && (!deltas || deltas.every(delta => delta === deltas[0]))) {
    // Preserve contextual shaping when uniform letter spacing can represent
    // every boundary (notably Japanese punctuation measured in context).
    return { clusters, paintOps: [{ ...baseOp,
      letterSpacingPt: baseOp.letterSpacingPt + (deltas?.[0] ?? perGapPt) }] };
  }
  const paintOps: TextPaintOp[] = [];
  let clusterCursor = 0;
  gapCursor = 0;
  cumulativePt = 0;
  for (const operation of segment.basePaintOps) {
    let start = operation.range.start;
    while (start < operation.range.end) {
      while (absoluteCuts[gapCursor] !== undefined && absoluteCuts[gapCursor]! <= start) {
        cumulativePt += deltas?.[gapCursor] ?? perGapPt;
        gapCursor += 1;
      }
      while (clusters[clusterCursor] && clusters[clusterCursor]!.range.start < start) clusterCursor += 1;
      const firstCluster = clusters[clusterCursor];
      if (!firstCluster || firstCluster.range.start !== start) {
        throw new Error('Internal paragraph justification is missing retained slice geometry');
      }
      const end = Math.min(operation.range.end, absoluteCuts[gapCursor] ?? operation.range.end);
      paintOps.push({
        ...operation,
        text: operation.text.slice(start - operation.range.start, end - operation.range.start),
        range: { start, end },
        offset: start === operation.range.start
          ? { ...operation.offset, xPt: operation.offset.xPt + cumulativePt }
          : firstCluster.offset,
      });
      start = end;
    }
  }
  return { clusters, paintOps };
}

/**
 * Keep RTL source coverage complete while preserving the trimmed word-shaped
 * operation used to anchor trailing whitespace on the physical leading edge.
 *
 * Internal justification may split a run immediately before its final space.
 * Trimming that preceding slice must therefore retain the removed whitespace
 * as an explicitly zero-ink source slice, or the immutable paint plan contains
 * an interior range hole.
 */
function retainedRtlPaintOperations(
  operations: readonly TextPaintOp[],
  clusters: readonly TextClusterLayout[],
): readonly TextPaintOp[] {
  return operations.flatMap((operation) => {
    const text = operation.text.trimEnd();
    if (text === '' || text.length === operation.text.length) return [operation];
    if (operation.sourceMapping === 'kashida') return [{ ...operation, text }];

    const trailingStart = operation.range.start + text.length;
    const trailingCluster = clusters.find((cluster) => cluster.range.start === trailingStart);
    const {
      inkBounds: _inkBounds,
      blockAxisInkBounds: _blockAxisInkBounds,
      ...zeroInkOperation
    } = operation;
    return [
      {
        ...operation,
        text,
        range: { ...operation.range, end: trailingStart },
      },
      {
        ...zeroInkOperation,
        text: operation.text.slice(text.length),
        range: { start: trailingStart, end: operation.range.end },
        offset: trailingCluster?.offset ?? operation.offset,
      },
    ];
  });
}

function sameDashPattern(
  left: readonly number[] | undefined,
  right: readonly number[] | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

/** Decoration endpoints reconstruct the same retained run seam through
 * independent justification sums. Their combined binary64 roundoff is bounded
 * by the precision of both operands; this is not a visible layout tolerance. */
function sameCoordinateWithinFloatingPrecision(left: number, right: number): boolean {
  if (left === right) return true;
  const roundoffBound = Number.EPSILON
    * Math.max(1, Math.abs(left) + Math.abs(right));
  return Math.abs(left - right) <= roundoffBound;
}

function sameContinuousDecoration(
  left: TextDecorationLayout,
  right: TextDecorationLayout,
): boolean {
  return left.kind === 'underline'
    && left.kind === right.kind
    && left.authoredStyle === right.authoredStyle
    && left.style === right.style
    && left.color === right.color
    && left.widthPt === right.widthPt
    && sameCoordinateWithinFloatingPrecision(left.to.xPt, right.from.xPt)
    && sameDashPattern(left.dashPatternPt, right.dashPatternPt);
}

function mergedContinuousDecoration(
  left: TextDecorationLayout,
  right: TextDecorationLayout,
): TextDecorationLayout {
  const yPt = Math.max(left.from.yPt, right.from.yPt);
  const from = { xPt: left.from.xPt, yPt };
  const to = { xPt: right.to.xPt, yPt };
  const { path: _discardedPath, ...withoutPath } = left;
  return {
    ...withoutPath,
    from,
    to,
    ...(left.style === 'wavy'
      ? { path: retainedWavePath(from, to, left.widthPt) }
      : {}),
  };
}

function trimTerminalUnderline(
  decoration: TextDecorationLayout,
  trimPt: number,
): TextDecorationLayout {
  const retainedTrimPt = Math.min(
    trimPt,
    Math.max(0, decoration.to.xPt - decoration.from.xPt),
  );
  const from = decoration.from;
  const to = { ...decoration.to, xPt: decoration.to.xPt - retainedTrimPt };
  const { path: _discardedPath, ...withoutPath } = decoration;
  return {
    ...withoutPath,
    from,
    to,
    ...(decoration.style === 'wavy'
      ? { path: retainedWavePath(from, to, decoration.widthPt) }
      : {}),
  };
}

/** Adjacent underlined source runs form one visual rule. Use a common
 * clearance below every glyph in the contiguous span; acquiring each run in
 * isolation instead creates stepped solid rules and restarts dash/wave phase
 * at source seams. Normalize the span before paint while keeping authored
 * style, color, and thickness boundaries intact. */
function coalesceAdjacentDecorations(placements: ParagraphPlacement[]): void {
  type ActiveDecoration = Readonly<{
    placementIndex: number;
    decorationIndex: number;
    decoration: TextDecorationLayout;
  }>;
  let active: ActiveDecoration[] = [];
  placements.forEach((placement, placementIndex) => {
    if ((placement.kind !== 'text' && placement.kind !== 'tab') || !placement.decorations) {
      active = [];
      return;
    }
    const retained: TextDecorationLayout[] = [];
    const nextActive: ActiveDecoration[] = [];
    const consumed = new Set<ActiveDecoration>();
    for (const decoration of placement.decorations) {
      const prior = active
        .filter((candidate) => !consumed.has(candidate)
          && sameContinuousDecoration(candidate.decoration, decoration))
        .sort((left, right) => Math.abs(left.decoration.from.yPt - decoration.from.yPt)
          - Math.abs(right.decoration.from.yPt - decoration.from.yPt))[0];
      if (prior) {
        consumed.add(prior);
        const owner = placements[prior.placementIndex];
        if (!owner || (owner.kind !== 'text' && owner.kind !== 'tab') || !owner.decorations) {
          throw new Error('Continuous decoration owner left the retained line');
        }
        const ownerDecorations = [...owner.decorations];
        const merged = mergedContinuousDecoration(prior.decoration, decoration);
        ownerDecorations[prior.decorationIndex] = merged;
        placements[prior.placementIndex] = { ...owner, decorations: ownerDecorations };
        nextActive.push({ ...prior, decoration: merged });
      } else {
        const decorationIndex = retained.length;
        retained.push(decoration);
        nextActive.push({ placementIndex, decorationIndex, decoration });
      }
    }
    placements[placementIndex] = { ...placement, decorations: retained };
    active = nextActive;
  });
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
  // A margin-allocated tab cell widens this line's band past the trailing
  // indent; alignment and justification slack use that same band.
  const availableWidthPt = Math.min(input.availableWidthPt, line.availableWidthPt)
    + (line.marginExtensionPt ?? 0);
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
  if (line.justifiedCompressionPt !== undefined) {
    const model = line.gapPlan;
    if (!model) throw new Error('Justified line is missing its retained gap plan');
    const retainedNaturalWidthPt = naturalWidthPt;
    naturalWidthPt = model.visibleWidthPx;
    lineSlackPt = effectiveAvailableWidthPt - physicalStartOffsetPt - naturalWidthPt;
    if (applyJustify || line.justifiedCompressionPt > 0) {
      // Opportunities and measured advances belong to the breaker. No text or
      // width reconstruction may change its accepted line in retained planning.
      const distribution = keepGraphemeSafeCuts(distributeProportionalSlack(
        [], lineSlackPt, {
          gapModel: { gaps: model.gaps, state: model.scan },
          proportional: true,
          unweightedExpansion: {
            gaps: model.expansionGaps,
            slack: effectiveAvailableWidthPt - physicalStartOffsetPt - retainedNaturalWidthPt,
          },
        },
      ), segments);
      if (distribution?.usedUnweightedExpansion) {
        naturalWidthPt = retainedNaturalWidthPt;
        lineSlackPt = effectiveAvailableWidthPt - physicalStartOffsetPt - naturalWidthPt;
      }
      stretchByIndex = distribution?.perSeg ?? null;
      perGapPt = distribution?.perGap ?? 0;
      distributedWidthPt = distributedDelta(distribution);
      if (line.justifiedCompressionPt > 0
        && Math.abs(distributedWidthPt + line.justifiedCompressionPt) > 1e-7) {
        throw new Error(`Justified fit and retained paint disagree: delta=${distributedWidthPt}, C=${line.justifiedCompressionPt}, slack=${lineSlackPt}, visible=${model.visibleWidthPx}, end=${model.lineEndSeparatorPx}`);
      }
    }
  } else if (applyJustify) {
    const distribution = keepGraphemeSafeCuts(distributeLineSlack(
      distSegments, lineSlackPt, firstContentIndex,
      bidi ? lastDrawnIndex : segments.length,
      -(line.baselinePt - line.topPt) * .25,
      lineSlackPt > 0,
      input.alignment === 'thaiDistribute' && lineSlackPt > 0,
    ), segments);
    stretchByIndex = distribution?.perSeg ?? null;
    perGapPt = distribution?.perGap ?? 0;
    distributedWidthPt = distributedDelta(distribution);
  }

  const drawnWidthPt = naturalWidthPt + distributedWidthPt;
  // WORD_FLOAT_GAP_FLOW aligns the visible word edge, while
  // the trailing separator remains source-owned. Its advance must not shift
  // centred/right-aligned gap text. Preserve ordinary lines' existing policy.
  let trailingSeparatorPt = 0;
  if (line.availableWidthPt < input.availableWidthPt && !bidi) {
    for (let index = segments.length - 1; index >= 0; index -= 1) {
      const segment = segments[index];
      if (segment?.kind !== 'text') break;
      const visibleLength = segment.text.trimEnd().length;
      trailingSeparatorPt += segment.clusters.filter(cluster =>
        cluster.range.start >= segment.range.start + visibleLength)
        .reduce((sum, cluster) => sum + cluster.advancePt, 0);
      if (visibleLength > 0) break;
    }
  }
  const alignmentSlackPt = lineSlackPt - distributedWidthPt + trailingSeparatorPt;
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
  const terminalDecorationTrims = new Map<number, number>();
  for (const segmentIndex of visual.order) {
    const segment = segments[segmentIndex];
    if (!segment) continue;
    const stretch = stretchByIndex?.get(segmentIndex);
    const internalStretchPt = stretch?.internalStretch ?? 0;
    const widthPt = segmentWidth(segment) + internalStretchPt;
    if (segment.kind === 'tab') {
      const bounds = { xPt, yPt: line.topPt, widthPt: segment.measuredWidthPt, heightPt: line.advancePt };
      const decorations = segment.underline
        ? retainedTextDecorations({
            origin: { xPt, yPt: line.baselinePt },
            advancePt: segment.measuredWidthPt,
            base: segment.underline.base,
            color: segment.underline.color,
            underline: segment.underline,
          })
        : undefined;
      placements.push({
        kind: 'tab', range: segment.range,
        bounds,
        advancePt: segment.measuredWidthPt,
        leader: segment.leader,
        ...(decorations?.length ? { decorations } : {}),
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
        ...(segment.sourceRunIndex === undefined
          ? {} : { sourceRunIndex: segment.sourceRunIndex }),
        resourceKey: segment.resourceKey, resourceKind: segment.resourceKind,
        ...(segment.orientation ? { orientation: segment.orientation } : {}),
        bounds: {
          xPt, yPt: line.baselinePt + segment.topOffsetPt,
          widthPt: segment.widthPt, heightPt: segment.heightPt,
        },
        advancePt: segment.measuredWidthPt,
      });
    } else if (segment.kind === 'unavailable-resource' || segment.kind === 'inline-drawing') {
      placements.push({
        kind: 'drawing',
        range: segment.range,
        drawingId: segment.drawingId,
        bounds: {
          xPt,
          yPt: line.baselinePt + segment.topOffsetPt,
          widthPt: segment.widthPt,
          heightPt: segment.heightPt,
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
        ...(segment.noteReference ? { noteReference: segment.noteReference } : {}),
      });
    } else {
      const {
        measuredWidthPt: _measuredWidthPt,
        breakBefore: _breakBefore,
        rtl: _rtl,
        digitsAsAN: _digitsAsAN,
        fixedPitch: _fixedPitch,
        decorationTerminalAdvancePt,
        textLayoutService: _textLayoutService,
        textShapeRequest: _textShapeRequest,
        selectedFaceFontBox,
        retainedGeometry,
        direction: _direction,
        ...style
      } = segment;
      const textGeometry = retainedTextGeometry(segment, stretch, perGapPt);
      const direction = visual.rtl[segmentIndex] ? 'rtl' : 'ltr';
      const paintOps = direction === 'rtl'
        ? retainedRtlPaintOperations(textGeometry.paintOps, textGeometry.clusters)
        : textGeometry.paintOps;
      const trailingWhitespaceStart = segment.text.trimEnd().length;
      const rtlLeadingGapPt = direction === 'rtl'
        ? (style.fitText?.trailingPadPt ?? 0) + segment.clusters
            .filter((cluster) => cluster.range.start >= segment.range.start + trailingWhitespaceStart)
            .reduce((sum, cluster) => sum + cluster.advancePt, 0)
        : 0;
      const ownedTrailingSlackPt = stretch?.trailingGap ? stretch.trailingDelta ?? perGapPt : 0;
      const origin = { xPt: xPt + rtlLeadingGapPt, yPt: line.baselinePt };
      const baselineOffsetPt = textGeometry.paintOps[0]?.offset.yPt ?? 0;
      const geometryOrigin = {
        xPt,
        yPt: line.baselinePt + baselineOffsetPt,
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
            yPt: line.baselinePt + baselineOffsetPt,
          },
          clusters: textGeometry.clusters,
          clusterInk: retainedGeometry.emphasis.clusterInk,
          mark: retainedGeometry.emphasis.mark,
          scaleX: segment.basePaintOps[0]?.scaleX ?? 1,
        }),
      } : undefined;
      const highlightFontBox = selectedFaceFontBox ?? retainedGeometry?.base;
      const highlightBounds = highlightFontBox ? {
        xPt,
        yPt: line.baselinePt + baselineOffsetPt - highlightFontBox.ascentPt,
        widthPt: widthPt + ownedTrailingSlackPt,
        heightPt: highlightFontBox.ascentPt + highlightFontBox.descentPt,
      } : {
        xPt,
        yPt: line.topPt,
        widthPt: widthPt + ownedTrailingSlackPt,
        heightPt: line.advancePt,
      };
      const placed: TextPlacement = {
        ...style,
        kind: 'text',
        origin,
        bounds: { xPt, yPt: line.topPt, widthPt, heightPt: line.advancePt },
        highlightBounds,
        advancePt: widthPt,
        clusters: textGeometry.clusters,
        paintOps: paintOps.map((operation) => ({ ...operation, direction })),
        decorations,
        ...(emphasis ? { emphasis } : {}),
        direction,
        ...(ownedTrailingSlackPt !== 0 ? { ownedTrailingSlackPt } : {}),
        ...((style.highlight || style.background) ? {
          highlightFragments: [{
            // Word for Mac PDF run shading (§17.3.2.32) follows the selected
            // font box for auto and atLeast spacing, centered inside any
            // larger line-grid allocation. With exact spacing it fills the
            // fixed line box. Highlighting
            // (§17.3.2.15) always hugs the selected font box.
            rect: style.highlight || !input.exactLineSpacing ? highlightBounds : {
              xPt,
              yPt: line.topPt,
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
      const terminalDecorationTrimPt = decorationTerminalAdvancePt === undefined
        ? 0
        : Math.max(
            0,
            widthPt + ownedTrailingSlackPt - decorationTerminalAdvancePt,
          );
      // The physical end of a bidi line is not the logical run tail. Keep this
      // compatibility path scoped to horizontal LTR text.
      if (direction === 'ltr' && terminalDecorationTrimPt > 0) {
        terminalDecorationTrims.set(placements.length, terminalDecorationTrimPt);
      }
      placements.push(placed);
    }
    xPt += widthPt;
    if (stretch?.trailingGap) xPt += stretch.trailingDelta ?? perGapPt;
  }
  for (const [placementIndex, trimPt] of terminalDecorationTrims) {
    const placement = placements[placementIndex];
    if (placement?.kind !== 'text' || !placement.decorations) continue;
    const next = placements[placementIndex + 1];
    const decorations = placement.decorations.map((decoration) => {
      if (decoration.kind !== 'underline') return decoration;
      const continues = (next?.kind === 'text' || next?.kind === 'tab')
        && next.decorations?.some((candidate) =>
          sameContinuousDecoration(decoration, candidate));
      return continues
        ? decoration
        : trimTerminalUnderline(decoration, trimPt);
    });
    placements[placementIndex] = { ...placement, decorations };
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
  coalesceAdjacentDecorations(placements);
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

/** Gap placement is complete before retention. All downstream consumers see
 * one physical line (§17.3.1.33 spacing, §17.6.8 numbering, §17.3.1.44 widows).
 * Retain disjoint allocations solely for horizontal shading; source ranges,
 * placements and vertical allocation belong to their common physical line.
 * The union visits each fragment/placement once, with no measurement. */
function retainPhysicalLines(fragments: readonly LineLayout[], physicalIds: readonly number[]): LineLayout[] {
  const lines: LineLayout[] = [];
  for (let start = 0; start < fragments.length;) {
    const first = fragments[start]!;
    let end = start + 1;
    while (end < fragments.length && physicalIds[end] === physicalIds[start]) end += 1;
    if (end === start + 1) lines.push(first);
    else {
      const group = fragments.slice(start, end);
      const { wrapBounds: _wrapBounds, ...physical } = first;
      lines.push({
        ...physical,
        range: { start: first.range.start, end: group.at(-1)!.range.end },
        bounds: unionLayoutRects(group.map(fragment => fragment.bounds))!,
        placements: group.flatMap(fragment => fragment.placements),
        wrapFragments: group.flatMap(fragment => fragment.wrapFragments
          ?? [fragment.wrapBounds ?? fragment.bounds]),
      });
    }
    start = end;
  }
  return lines;
}

export function paragraphSliceAdvance(input: Pick<AcquiredParagraphLayoutInput,
  'continuation' | 'lines' | 'spacing' | 'flowBounds' | 'paragraphMark'>): number {
  const continuation = input.continuation;
  const start = continuation?.lineStart ?? 0;
  const end = continuation?.lineEnd ?? input.lines.length;
  if (start < 0 || end < start || end > input.lines.length) {
    throw new RangeError('Paragraph continuation line range is outside the retained lines');
  }
  const beforePt = continuation?.continuesFromPrevious ? 0 : input.spacing.beforePt;
  let advancePt = beforePt;
  for (let index = start; index < end; index += 1) {
    advancePt += retainedLineAdvanceContributionPt(input, index, start, beforePt);
  }
  if (input.lines.length === 0 && input.paragraphMark) {
    advancePt += finiteNonNegative(input.paragraphMark.bounds.heightPt, 'paragraphMark.heightPt');
  }
  if (!continuation?.continuesOnNext) advancePt += input.spacing.afterPt;
  return advancePt;
}

/** One retained line's share of a slice advance: its own advance plus the
 * gap above it (from the flow top for line 0, else from the previous line
 * when that line is inside the slice). Sole formula for slice advances. */
function retainedLineAdvanceContributionPt(
  input: Pick<AcquiredParagraphLayoutInput, 'lines' | 'flowBounds'>,
  index: number,
  start: number,
  beforePt: number,
): number {
  const line = input.lines[index];
  if (!line) return 0;
  let advancePt = 0;
  if (index === 0) {
    // A remeasured body continuation starts at lineStart 0 without space
    // before; wrap may still place its first line below the flow cursor.
    advancePt += Math.max(0, line.bounds.yPt - (input.flowBounds.yPt + beforePt));
  } else if (index > start) {
    const previous = input.lines[index - 1];
    advancePt += Math.max(0,
      line.bounds.yPt - ((previous?.bounds.yPt ?? line.bounds.yPt) + (previous?.advancePt ?? 0)));
  }
  return advancePt + finiteNonNegative(line.advancePt, 'line.advancePt');
}

/**
 * Exact advance index for continuing prefixes of a retained paragraph:
 * `index[e]` equals `sliceParagraphLayout(layout, { lineStart: 0, lineEnd: e,
 * continuesFromPrevious, continuesOnNext: true }).advancePt` for 1 <= e <=
 * lines.length (`index[0]` is the leading space alone). One linear pass over
 * the same per-line formula lets pagination read prefix charges without
 * materialising a slice per candidate. A completed fragment still owns its
 * trailing spacing through its own advance.
 */
export function paragraphContinuingPrefixAdvancesPt(
  layout: Pick<ParagraphLayout, 'lines' | 'flowBounds' | 'spacing'>,
  continuesFromPrevious: boolean,
): readonly number[] {
  const beforePt = continuesFromPrevious ? 0 : layout.spacing.beforePt;
  const index = [beforePt];
  for (let line = 0; line < layout.lines.length; line += 1) {
    index.push(index[line]! + retainedLineAdvanceContributionPt(layout, line, 0, beforePt));
  }
  return index;
}

/**
 * Finalizes the parser-independent paragraph acquisition snapshot. All coordinates
 * are scale-1 points; subsequent Canvas paint is a pure viewport transform.
 */
export function layoutParagraph(input: AcquiredParagraphLayoutInput, frozenSource?: ParagraphLayout): ParagraphLayout {
  return finalizeParagraphLayout(input, frozenSource);
}

/** Slice construction supplies only admitted/rebased lines, while the original
 * vector remains the authority for absolute source ranges and advance. */
function finalizeParagraphLayout(input: AcquiredParagraphLayoutInput, frozenSource?: ParagraphLayout,
  selectedLines?: readonly LineLayout[]): ParagraphLayout {
  const lineStart = input.continuation?.lineStart ?? 0;
  const lineEnd = input.continuation?.lineEnd ?? input.lines.length;
  const lines = selectedLines ?? input.lines.slice(lineStart, lineEnd);
  const advancePt = input.continuation
    ? paragraphSliceAdvance(input)
    : finiteNonNegative(input.flowBounds.heightPt, 'flowBounds.heightPt');
  const node: ParagraphLayout = {
    kind: 'paragraph',
    id: input.id,
    source: input.source,
    ...(input.paragraphId !== undefined ? { paragraphId: input.paragraphId } : {}),
    flowDomainId: input.flowDomainId,
    ordinaryFlow: input.ordinaryFlow,
    ...(input.styleId !== undefined ? { styleId: input.styleId } : {}),
    ...(input.bookmarkStarts?.length
      ? { bookmarkStarts: input.bookmarkStarts }
      : {}),
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
    ...(input.cellContainmentBounds
      ? { cellContainmentBounds: input.cellContainmentBounds }
      : {}),
    ...(input.anchorCollisions?.length
      ? { anchorCollisions: input.anchorCollisions }
      : {}),
    ...(input.anchorFrames ? { anchorFrames: input.anchorFrames } : {}),
    ...(input.paragraphMark ? { paragraphMark: input.paragraphMark } : {}),
    ...(input.continuation ? { continuation: input.continuation } : {}),
  };
  if (frozenSource) {
    // Continuation metadata is the only retained object supplied by this call
    // rather than inherited from the already sealed acquired paragraph.
    if (input.continuation) deepFreezePlainData(input.continuation);
    return deepFreezePlainDataWithFrozenAliases(node, frozenSource);
  }
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
  /** Effective prior DrawingML objects in this flow domain. */
  readonly anchorCollisions?: readonly DrawingMLCollisionEntryPt[];
  /** Present only while acquiring a paragraph hosted by a table cell. */
  readonly anchorCellBounds?: LayoutRect;
  /** Effective enclosing fill, retained only for automatic text-color resolution. */
  readonly containerShading?: string | null;
  /** Layout-owned §17.3.1.7 edge selection for adjacent/sliced border boxes. */
  readonly paragraphBorderEdges?: ParagraphBorderEdges;
  /** Final flow reservation; may exceed w:after when a bottom border owns more space. */
  readonly trailingExtentPt?: number;
  /** The measurement starts after a consumed line boundary on another flow slice. */
  readonly continuesFromPrevious?: boolean;
  /** Exact paragraph occurrence offset corresponding to the continuation boundary. */
  readonly sourceRangeStart?: number;
  readonly anchorFrames?: Readonly<Pick<
    AnchorReferenceFramesInput,
    'page' | 'margin' | 'column' | 'pageParity'
  >>;
  readonly acquireCompleteStory?: CompleteTextBoxStoryAcquirer;
  /**
   * The translation from this paragraph's coordinates to the coordinates of
   * its `anchorFrames` page that its host-following content still receives,
   * when its acquisition knows it: none for a body paragraph, its band for a
   * header, footer or note story given one. Absent (a table cell, whose page
   * position only pagination knows) the page frames of its drawings' text box
   * stories are unknown. Page-owned drawing axes never receive it.
   */
  readonly hostFlowPageTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  /**
   * The page frames that translation reaches, when they are not the
   * `anchorFrames` page: in a text box story, the destination page carried
   * into the coordinates its page-owned anchor axes keep
   * (story-page-frames.ts storyAnchorPageFrames). Null where that story has
   * no page frames (its box's placement carries no page band into it, or it
   * is laid out before its box is placed): its drawings' text box stories
   * then get none either.
   */
  readonly hostPageFrames?: StoryPageFrames | null;
  /** WORD_MODE14_TIGHT_ANCHOR_LINE_REWRAP decisions, keyed by anchor
   * occurrence: the host line top whose layout ignores that object. */
  readonly anchorLineExemptions?: ReadonlyMap<string, number>;
  /** WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: carried first-placement object
   * frames keyed by anchor occurrence. */
  readonly frozenAnchorFrames?: ReadonlyMap<string, LayoutRect>;
  /** Cell-grid wrap preserves an empty carrier's insertion reference while
   * its paragraph flow is admitted below the grid. Paragraph-local points. */
  readonly paragraphAnchorReferenceDeltaPt?: number;
}

function runSource(source: SourceRef, runIndex: number): SourceRef {
  return { ...source, path: [...source.path, runIndex] };
}

function shapePlanDiagnostics(
  plan: ShapeDrawingPlanResult,
  source: SourceRef,
): readonly LayoutDiagnostic[] {
  if (plan.status === 'planned') return Object.freeze([]);
  const retainedSource = Object.freeze({
    ...source,
    path: Object.freeze([...source.path]),
  });
  return Object.freeze(plan.diagnostics.map((diagnostic) => Object.freeze({
    ...diagnostic,
    source: retainedSource,
  })));
}

function chartResourceKey(source: SourceRef): string {
  return canonicalChartResourceKey(source);
}

function unavailableDrawingId(source: SourceRef, runIndex: number): string {
  return stableFingerprint('unavailable-drawing', runSource(source, runIndex));
}

function unavailableDrawingDiagnostic(
  resourceKind: 'image' | 'chart',
  source: SourceRef,
): LayoutDiagnostic {
  return Object.freeze({
    code: 'MISSING_RESOURCE',
    severity: 'warning',
    source: Object.freeze({
      ...source,
      path: Object.freeze([...source.path]),
    }),
    message: `Drawing ${resourceKind} resource is unavailable`,
  });
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
  // A native reserved-separator participant retains the sides of its bounded
  // probe, in points, at the same effective metric size as its line box
  // (pass-operations performTextSegmentBox); generic metric-only hosts keep
  // their existing request.
  const shape = segment.metricOnly && segment.metricProbeText
    ? segment.textLayoutService.shape({
        ...independentTextShapeRequest(segment.textShapeRequest, segment.metricProbeText),
        fontSizePt: segment.smallCaps && !segment.vertAlign
          ? segment.fontSize : calcEffectiveFontPx(segment, 1),
        measure: true,
        clusterGeometry: false,
      })
    : segment.textLayoutService.shape({
        ...segment.textShapeRequest,
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

/** Canvas-space baseline offset for run-level vertical positioning.
 * ECMA-376 Part 1 §17.3.2.42 requires superscript/subscript above/below the
 * default baseline; the exact Office displacement is isolated as compatibility. */
function retainedBaselineOffsetPt(segment: LayoutTextSeg): number {
  const verticalAlignRaisePt = wordRunVerticalAlignRaisePt(
    segment.vertAlign,
    segment.fontSize,
  );
  const raisePt = verticalAlignRaisePt
    + (segment.lineRelativePosition ?? segment.position ?? 0);
  return raisePt === 0 ? 0 : -raisePt;
}

function textPlacement(
  segment: LayoutTextSeg,
  paragraph: ParagraphLayoutSource,
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
      ...(run?.type === 'text' && (run.noteRef?.kind === 'footnote' || run.noteRef?.kind === 'endnote')
        ? { noteReference: { kind: run.noteRef.kind, id: run.noteRef.id } } : {}),
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
        measure: true,
      })
    : undefined;
  const rubyShape = segment.ruby && segment.textLayoutService && segment.textShapeRequest
    ? segment.textLayoutService.shape({
        ...independentTextShapeRequest(segment.textShapeRequest, segment.ruby.text),
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
  const baselineOffsetPt = retainedBaselineOffsetPt(segment);
  const sourceEnd = sourceOffset + logicalTextSegmentLength(segment);
  return {
    kind: 'text',
    text: segment.text,
    ...(segment.optionalHyphenGlyph ? { optionalHyphenGlyph: true as const } : {}),
    ...(segment.semanticSlotSpans ? { semanticSlotSpans: segment.semanticSlotRange
      ? sliceSemanticSlotSpans(segment.semanticSlotSpans,
        segment.semanticSlotRange.start, segment.semanticSlotRange.end)
      : segment.semanticSlotSpans } : {}),
    ...(runIndex === undefined ? {} : { sourceRunIndex: runIndex }),
    ...(run?.type === 'field' ? { role: 'field-result' as const, dependency: fieldDependency(run) } : {}),
    ...(run?.type === 'text'
      && (run.noteRef?.kind === 'footnote' || run.noteRef?.kind === 'endnote')
      ? { noteReference: { kind: run.noteRef.kind, id: run.noteRef.id } }
      : {}),
    range: { start: sourceOffset, end: sourceEnd },
    origin: { xPt, yPt: baselinePt + baselineOffsetPt },
    bounds: { xPt, yPt: topPt, widthPt: segment.measuredWidth, heightPt },
    advancePt: segment.measuredWidth,
    clusters: [{
      range: { start: sourceOffset, end: sourceEnd },
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
    kerning: segment.textShapeRequest?.kerning ?? wordKerningApplies(segment.fontSize, segment.kerning),
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
      range: { start: sourceOffset, end: sourceEnd },
      offset: { xPt: 0, yPt: baselineOffsetPt },
      letterSpacingPt: effectiveCharacterSpacingPt(segment),
      scaleX: segment.charScale ?? 1,
      direction: segment.rtl ? 'rtl' : 'ltr',
      kerning: (segment.textShapeRequest?.kerning ?? wordKerningApplies(segment.fontSize, segment.kerning)) ? 'normal' : 'none',
      writingMode: segment.verticalRun ? 'vertical-rl' : 'horizontal-tb',
    }],
    ...(segment.hyperlink ? { hyperlink: segment.hyperlink } : {}),
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
  readonly lineBox?: import('./numbering-marker.js').NumberingMarkerLineBox;
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

function numberingAlignedLeadingEdgePt(
  plan: RetainedNumberingPlan,
  context: ParagraphLayoutContext,
  paragraphXPt: number,
  availableWidthPt: number,
  line: LineLayout,
): number {
  // Non-empty lines retain their aligned body bounds. A numbering-only line has
  // zero width and therefore retains the paragraph frame origin instead of a
  // body origin with the resolved marker suffix offset applied.
  if (line.bounds.widthPt <= 0) {
    return context.baseRtl ? paragraphXPt + availableWidthPt : paragraphXPt;
  }
  return context.baseRtl
    ? line.bounds.xPt + line.bounds.widthPt + plan.bodyOffsetPt
    : line.bounds.xPt - plan.bodyOffsetPt;
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
    alignedLeadingEdgePt: numberingAlignedLeadingEdgePt(
      plan,
      context,
      paragraphXPt,
      availableWidthPt,
      line,
    ),
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
        kerning: paragraph.numberingMarkerShapeInput?.kerning ? 'normal' : 'none',
        writingMode: 'horizontal-tb',
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
  // ECMA-376 §17.13.5 markup view (`word-track-change-decoration`): an
  // insertion/moveTo segment gains an author-coloured underline, a
  // deletion/moveFrom segment an author-coloured strikethrough. Authored
  // run decoration wins over the synthesized revision decoration on its axis.
  const markup = segment.trackChangesMarkup;
  const markupDecoration = wordTrackChangeDecoration(markup?.kind);
  if (!(segment.highlight || segment.underline || segment.strikethrough
    || segment.doubleStrikethrough || segment.emphasisMark
    || markupDecoration.underline || markupDecoration.strike)) return undefined;
  const service = segment.textLayoutService;
  const request = segment.textShapeRequest;
  if (!service || !request) {
    throw new Error('Retained typography geometry requires TextLayoutService');
  }
  const shapeProbe = (text: string) => service.shape({ ...independentTextShapeRequest(request, text), measure: true });
  const glyphProbe = (text: string): RetainedInkMetric => {
    const measured = shapeProbe(text);
    const span = measured.spans[0];
    if (!span || measured.spans.length !== 1 || span.start !== 0 || span.end !== text.length) {
      throw new Error('Retained decoration probe requires one selected-face span');
    }
    return {
      ascentPt: span.ascentPt,
      descentPt: span.descentPt,
      ...(span.inkBounds ? { inkBounds: span.inkBounds } : {}),
    };
  };
  const fontBox = segment.selectedFaceFontBox;
  if (!fontBox || !segment.selectedFaceInkBounds) {
    throw new Error('Retained typography geometry requires authoritative selected-face metrics');
  }
  const base: RetainedInkMetric = {
    ascentPt: fontBox.ascentPt,
    descentPt: fontBox.descentPt,
    inkBounds: segment.selectedFaceInkBounds,
  };
  const textColor = retainedColorString(color);
  const underline = segment.underline ? {
    ...(segment.underlineStyle ? { authoredStyle: segment.underlineStyle } : {}),
    color: segment.underlineColor && segment.underlineColor !== 'auto'
      ? `#${segment.underlineColor}` : textColor,
    probe: glyphProbe('_'),
  } : markup && markupDecoration.underline ? {
    color: markup.authorColor,
    probe: glyphProbe('_'),
  } : undefined;
  const strike = segment.strikethrough || segment.doubleStrikethrough ? {
    double: segment.doubleStrikethrough === true,
    probe: glyphProbe('-'),
    ...(segment.doubleStrikethrough ? { doubleProbe: glyphProbe('=') } : {}),
  } : markup && markupDecoration.strike ? {
    double: false,
    probe: glyphProbe('-'),
    color: markup.authorColor,
  } : undefined;
  const emphasis = segment.emphasisMark ? (() => {
    const glyph = emphasisGlyph(segment.emphasisMark);
    const markShape = shapeProbe(glyph);
    const markSpan = markShape.spans[0];
    if (!markSpan) throw new Error('Emphasis shaping produced no selected-face span');
    // §17.3.2.12 positions w:em against each base cluster's ink. Its font
    // selection must retain the same run range as the body placement.
    const clusterInk = (segment.shapedClusters ?? []).map((cluster): RetainedEmphasisClusterInk => {
      const text = segment.text.slice(cluster.range.start, cluster.range.end);
      return {
        text,
        range: {
          start: sourceOffset + cluster.range.start,
          end: sourceOffset + cluster.range.end,
        },
        ink: completeInkBounds(service.shape({
          ...sliceTextShapeRequest(request, cluster.range.start, cluster.range.end), measure: true,
        })),
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
  paragraph: ParagraphLayoutSource,
  sourceOffset: number,
  characterGrid: DocGridCtx | undefined,
  sourceRun?: import('./text.js').ParagraphLayoutRun & Readonly<{
    anchorOccurrenceId?: string;
  }>,
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  sourceRuns?: TextPlacement['sourceRuns'],
): MeasuredTextPlanSegment | MeasuredAnchorHostPlanSegment {
  if (segment.metricOnly) {
    const sourceMetrics = selectedFaceSourceMetrics(segment);
    return {
      kind: 'anchor-host', measuredWidthPt: 0,
      ...(sourceRun?.type === 'text' && (sourceRun.noteRef?.kind === 'footnote' || sourceRun.noteRef?.kind === 'endnote')
        ? { noteReference: { kind: sourceRun.noteRef.kind, id: sourceRun.noteRef.id } } : {}),
      range: { start: sourceOffset, end: sourceOffset },
      ...(sourceMetrics ? { sourceMetrics } : {}),
      ...(sourceRun?.type === 'anchorHost' && sourceRun.anchorOccurrenceId
        ? { anchorOccurrenceId: sourceRun.anchorOccurrenceId }
        : {}),
    };
  }
  const projected = { ...textPlacement(segment, paragraph, sourceOffset, 0, 0, 0, 0),
    ...(sourceRuns ? { sourceRuns } : {}) };
  if (projected.kind !== 'text') throw new Error('Visible text segment projected as anchor host');
  const pitchPt = segLetterSpacingPx(segment, characterGrid, 1);
  const scaleX = segment.charScale ?? 1;
  const baselineOffsetPt = retainedBaselineOffsetPt(segment);
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
  let clusters = (shapedClusters ?? []).map((cluster, index) => {
    const prefix = segment.text.slice(0, cluster.range.start);
    const text = segment.text.slice(cluster.range.start, cluster.range.end);
    const precedingScalars = [...prefix].length;
    const scalarCount = [...text].length;
    const trailingFitPad = index === (shapedClusters?.length ?? 0) - 1
      ? segment.fitTextTrailingPadPx ?? 0
      : 0;
    const precedingPunctuationCompression =
      segment.punctuationCompressions
        ?.filter((compression) => compression.end <= cluster.range.start)
        .reduce((sum, compression) => sum + compression.adjustmentPt, 0)
      ?? 0;
    const clusterPunctuationCompression =
      segment.punctuationCompressions
        ?.filter((compression) =>
          compression.end > cluster.range.start
          && compression.end <= cluster.range.end)
        .reduce((sum, compression) => sum + compression.adjustmentPt, 0)
      ?? 0;
    const precedingWidthBalanceAdjustment =
      widthBalanceSpaceAdjustmentForTextPt(segment, prefix, characterGrid) * scaleX;
    const clusterWidthBalanceAdjustment =
      widthBalanceSpaceAdjustmentForTextPt(segment, text, characterGrid) * scaleX;
    return {
      range: {
        start: sourceOffset + cluster.range.start,
        end: sourceOffset + cluster.range.end,
      },
      offset: {
        xPt:
          cluster.offsetPt * scaleX
          + (segment.leadingWordBoundaryPx ?? 0)
          + precedingScalars * pitchPt
          + precedingWidthBalanceAdjustment
          + precedingPunctuationCompression,
        yPt: baselineOffsetPt,
      },
      advancePt:
        cluster.advancePt * scaleX
        + scalarCount * pitchPt
        + clusterWidthBalanceAdjustment
        + trailingFitPad
        + clusterPunctuationCompression,
    };
  });
  if (segment.latinSpaceCompressionPx && segment.text.endsWith(' ') && clusters.length > 0) {
    // The fit projections remove only the segment's trailing invisible U+0020
    // advances, shrinking each by the same share: one space under
    // WORD_LATIN_INTERWORD_XAVG_FLOOR, possibly several consecutive ones under
    // WORD_COMPRESSED_SPACE_LINE_FIT. Take each share from the cluster owning
    // that space and move later clusters with it, so retained clusters and
    // source-owned fragments stay inside the same shortened segment box;
    // Canvas paint operations still draw the preceding visible glyph naturally.
    const spacesStart = sourceOffset + segment.text.replace(/ +$/u, '').length;
    const perSpacePt = segment.latinSpaceCompressionPx
      / (sourceOffset + segment.text.length - spacesStart);
    let removedPt = 0;
    clusters = clusters.map((cluster) => {
      const spaces = Math.max(0, cluster.range.end - Math.max(cluster.range.start, spacesStart));
      if (spaces === 0 && removedPt === 0) return cluster;
      const placed = {
        ...cluster,
        offset: { ...cluster.offset, xPt: cluster.offset.xPt - removedPt },
        advancePt: Math.max(0, cluster.advancePt - spaces * perSpacePt),
      };
      removedPt += spaces * perSpacePt;
      return placed;
    });
  }
  const snapLeadingPadPt = segment.snapGridLeadingPadPx ?? 0;
  let decorationTerminalAdvancePt = segment.measuredWidth
    - (segment.snapGridTrailingPadPx ?? 0);
  let griddedTerminalInkAdvancePt: number | undefined;
  if (segment.snapGridClass === 'eastAsia' && segment.snapGridCellPitchPx) {
    const cellPitchPt = segment.snapGridCellPitchPx;
    let precedingCells = 0;
    clusters = clusters.map((cluster, index) => {
      const text = segment.text.slice(
        cluster.range.start - sourceOffset,
        cluster.range.end - sourceOffset,
      );
      const cells = wordSnapToCharsEastAsianCellCount(
        cluster.advancePt,
        cellPitchPt,
      );
      const allocatedAdvancePt = cells * cellPitchPt;
      const centeredOffsetPt = precedingCells * cellPitchPt
        + (allocatedAdvancePt - cluster.advancePt) / 2;
      if (index === clusters.length - 1) {
        decorationTerminalAdvancePt = centeredOffsetPt + cluster.advancePt;
        const naturalClusterOffsetPt = shapedClusters?.[index]?.offsetPt;
        if (
          segment.selectedFaceInkBounds
          && naturalClusterOffsetPt !== undefined
          && !/\s$/u.test(segment.text)
        ) {
          // The selected-face ink bounds describe the whole unsnapped segment.
          // Translate the terminal glyph's tight extent from its natural
          // cluster origin to the independently centered final grid cell.
          griddedTerminalInkAdvancePt = centeredOffsetPt + Math.max(
            0,
            (segment.selectedFaceInkBounds.xMaxPt - naturalClusterOffsetPt) * scaleX,
          );
        }
      }
      const placed = {
        ...cluster,
        offset: {
          xPt: centeredOffsetPt,
          yPt: cluster.offset.yPt,
        },
        advancePt: allocatedAdvancePt,
      };
      precedingCells += cells;
      return placed;
    });
  } else if (snapLeadingPadPt !== 0) {
    clusters = clusters.map((cluster) => ({
      ...cluster,
      offset: {
        xPt: cluster.offset.xPt + snapLeadingPadPt,
        yPt: cluster.offset.yPt,
      },
    }));
  }
  const {
    origin: _origin, bounds: _bounds, advancePt: _advancePt,
    paintOps, clusters: _clusters, ...style
  } = projected;
  const tateChuYokoScaleY = segment.tateChuYoko && segment.tateChuYokoCompress
    ? (() => {
        if (!segment.textLayoutService || !segment.textShapeRequest) {
          throw new Error('Tate-chu-yoko compression requires TextLayoutService');
        }
        const shape = segment.textLayoutService.shape({
          ...segment.textShapeRequest,
          fontSizePt: projected.fontSizePt,
          measure: true,
          clusterGeometry: false,
        });
        const fontBoxHeightPt = shape.ascentPt + shape.descentPt;
        return fontBoxHeightPt > projected.fontSizePt && fontBoxHeightPt > 0
          ? projected.fontSizePt / fontBoxHeightPt
          : 1;
      })()
    : 1;
  const hasInternalPunctuationCompression = segment.punctuationCompressions
    ?.some((compression) => compression.end < segment.text.length) ?? false;
  const unpaddedPaintOps = segment.verticalRun
    ? (() => {
        if (!verticalGlyphMeasurement) {
          throw new Error('Vertical glyph planning capability is required for vertical text');
        }
        const template = paintOps[0]!;
        return verticalGlyphMeasurement.planRun({
          text: segment.text,
          font: canvasFontString(
            projected.fontRoute,
            projected.fontSizePt,
            projected.fontWeight,
            projected.fontStyle,
          ),
          fontKerning: template.kerning,
          fontSizePt: projected.fontSizePt,
          letterSpacingPt: pitchPt,
          charScale: scaleX,
          growTrRotateInk: true,
          writingMode: template.writingMode,
        }).map((cell) => ({
          ...template,
          text: cell.text,
          range: {
            start: sourceOffset + cell.range.start,
            end: sourceOffset + cell.range.end,
          },
          offset: {
            xPt: cell.originPt + (
              segment.punctuationCompressions
                ?.filter((compression) => compression.end <= cell.range.start)
                .reduce((sum, compression) => sum + compression.adjustmentPt, 0)
              ?? 0
            ),
            yPt: baselineOffsetPt,
          },
          letterSpacingPt: pitchPt,
          glyphOrientation: cell.orientation,
          ...(cell.verticalFeature ? { verticalFeature: true } : {}),
          ...(cell.blockAxisInkBounds ? { blockAxisInkBounds: cell.blockAxisInkBounds } : {}),
          ...(cell.drawOffsetPt.xPt !== 0 || cell.drawOffsetPt.yPt !== 0
            ? { glyphOffsetPt: cell.drawOffsetPt }
            : {}),
        }));
      })()
    : segment.tateChuYoko
      ? paintOps.map((operation) => ({
          ...operation,
          offset: {
            xPt: operation.offset.xPt + segment.measuredWidth / 2,
            yPt: operation.offset.yPt,
          },
          glyphOrientation: 'upright' as const,
          ...(tateChuYokoScaleY !== 1 ? { scaleY: tateChuYokoScaleY } : {}),
        }))
      : hasInternalPunctuationCompression
        ? (() => {
            const template = paintOps[0]!;
            const groups: Array<{
              start: number;
              end: number;
              offset: typeof clusters[number]['offset'];
              adjustmentPt: number | null;
            }> = [];
            for (const cluster of clusters) {
              const relativeEnd = cluster.range.end - sourceOffset;
              const adjustmentPt = segment.punctuationCompressions
                ?.find((compression) => compression.end === relativeEnd)
                ?.adjustmentPt ?? null;
              const previous = groups.at(-1);
              if (previous && previous.adjustmentPt === adjustmentPt) {
                previous.end = cluster.range.end;
              } else {
                groups.push({
                  start: cluster.range.start,
                  end: cluster.range.end,
                  offset: cluster.offset,
                  adjustmentPt,
                });
              }
            }
            return groups.map((group) => ({
              ...template,
              text: segment.text.slice(
                group.start - sourceOffset,
                group.end - sourceOffset,
              ),
              range: { start: group.start, end: group.end },
              offset: group.offset,
              letterSpacingPt: pitchPt + (group.adjustmentPt ?? 0),
            }));
          })()
        : paintOps;
  const basePaintOps = segment.snapGridClass === 'eastAsia'
    ? (() => {
        const template = unpaddedPaintOps[0];
        if (!template) return unpaddedPaintOps;
        return clusters.map((cluster) => ({
          ...template,
          text: segment.text.slice(
            cluster.range.start - sourceOffset,
            cluster.range.end - sourceOffset,
          ),
          range: cluster.range,
          offset: cluster.offset,
        }));
      })()
    : snapLeadingPadPt === 0
      ? unpaddedPaintOps
      : unpaddedPaintOps.map((operation) => ({
          ...operation,
          offset: {
            xPt: operation.offset.xPt + snapLeadingPadPt,
            yPt: operation.offset.yPt,
          },
        }));
  const retainedDecorationTerminalAdvancePt =
    characterGrid?.type === 'snapToChars'
    && segment.underline
    && !segment.verticalRun
    && paragraph.bidi !== true
    && segment.selectedFaceInkBounds
      ? griddedTerminalInkAdvancePt ?? (
          basePaintOps.length === 1
            ? basePaintOps[0]!.offset.xPt
              + (basePaintOps[0]!.glyphOffsetPt?.xPt ?? 0)
              + segment.selectedFaceInkBounds.xMaxPt * (basePaintOps[0]!.scaleX ?? 1)
            : decorationTerminalAdvancePt
        )
      : decorationTerminalAdvancePt;
  return {
    ...style,
    kind: 'text', measuredWidthPt: segment.measuredWidth,
    ...(segment.latinSpaceCompressionPx ? {
      trailingSpaceCompressionPt: segment.latinSpaceCompressionPx,
    } : {}),
    // A discretionary glyph has real ink and advance, but no source UTF-16
    // units. Keep its shaped geometry at the authored zero-length boundary.
    clusters: segment.optionalHyphenGlyph ? clusters.map(cluster => ({
      ...cluster, range: { start: sourceOffset, end: sourceOffset },
    })) : clusters,
    basePaintOps: basePaintOps.map((operation) => ({
      ...operation,
      ...(segment.optionalHyphenGlyph ? { range: { start: sourceOffset, end: sourceOffset } } : {}),
      offset: { ...operation.offset,
        xPt: operation.offset.xPt + (segment.leadingWordBoundaryPx ?? 0) },
      // Measurement resolves w:spacing, docGrid character pitch, and w:fitText
      // into one authoritative per-scalar pitch. A planned vertical upright or
      // rotate cell already owns that pitch in its retained origin and advance;
      // applying Canvas letterSpacing again would move a centered single glyph
      // on the physical cross axis. Contextual sideways text and horizontal
      // tate-chu-yoko retain Canvas spacing within their multi-glyph operation.
      letterSpacingPt:
        segment.verticalRun && operation.glyphOrientation !== 'sideways'
          ? 0
          : hasInternalPunctuationCompression
            ? operation.letterSpacingPt
            : pitchPt,
      ...(!segment.verticalRun && segment.selectedFaceInkBounds
        ? { inkBounds: segment.selectedFaceInkBounds }
        : {}),
      ...(!segment.verticalRun && segment.selectedFaceInkBounds
        && operation.glyphOrientation === undefined
        ? {
            blockAxisInkBounds: {
              startPt: (operation.glyphOffsetPt?.yPt ?? 0)
                - segment.selectedFaceInkBounds.ascentPt,
              endPt: (operation.glyphOffsetPt?.yPt ?? 0)
                + segment.selectedFaceInkBounds.descentPt,
            },
          }
        : {}),
    })),
    breakBefore: segment.breakBefore !== false && !segment.joinPrev,
    rtl: segment.rtl,
    digitsAsAN: segment.digitsAsAN,
    fixedPitch: segment.fitTextRegionIndex !== undefined || segment.snapGridClass !== undefined,
    ...(characterGrid?.type === 'snapToChars' && segment.underline
      ? { decorationTerminalAdvancePt: retainedDecorationTerminalAdvancePt }
      : {}),
    ...(retainedGeometry ? { retainedGeometry } : {}),
    ...(segment.selectedFaceFontBox ? { selectedFaceFontBox: segment.selectedFaceFontBox } : {}),
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
  showTrackedChanges: boolean | undefined,
): LogicalOccurrenceMap {
  const measuredLengths = new Map<number, number>();
  const sequences = new Set<NonNullable<LayoutTextSeg['sourceTextSequence']>>();
  for (const line of measured.lines) {
    for (const segment of line.layout.segments) {
      if (segment.sourceTextSequence) {
        if (!sequences.has(segment.sourceTextSequence)) {
          sequences.add(segment.sourceTextSequence);
          for (const owner of segment.sourceTextSequence) measuredLengths.set(owner.runIndex, owner.end - owner.start);
        }
        continue;
      }
      const runIndex = sourceRunIndex(segment);
      if (runIndex === undefined) continue;
      const length = 'text' in segment
        ? logicalTextSegmentLength(segment)
        : 'math' in segment ? segment.fallbackText.length
          : 'isTab' in segment || 'imagePath' in segment ? 1 : 0;
      measuredLengths.set(runIndex, (measuredLengths.get(runIndex) ?? 0) + length);
    }
  }
  const runLengths = paragraph.runs.map((run, runIndex) => {
    // Canonical sequence offsets cover displayed text only. Falling back to
    // omitted source lengths would shift a later real-format sequence when
    // the preceding visible text is partitioned into different source runs.
    const kind = (run as { revision?: { kind?: string } }).revision?.kind;
    if (revisionIsOmitted(kind, showTrackedChanges)) return 0;
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

function logicalTextSegmentLength(segment: LayoutTextSeg): number {
  return segment.metricOnly || segment.optionalHyphenGlyph ? 0 : segment.text.length;
}

function segmentOccurrenceLength(segment: LayoutTextSeg | LayoutTabSeg | LayoutImageSeg | LayoutMathSeg): number {
  if ('text' in segment) return logicalTextSegmentLength(segment);
  if ('math' in segment) return segment.fallbackText.length;
  return 1;
}

function planMeasuredLines(
  measured: MeasuredParagraph,
  paragraph: ParagraphAcquisitionInput,
  paragraphXPt: number,
  availableWidthPt: number,
  source: SourceRef,
  paragraphId: string,
  context: ParagraphLayoutContext,
  occurrences: LogicalOccurrenceMap,
  numberingPlan?: RetainedNumberingPlan,
  textService?: import('./text.js').TextLayoutService,
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  verticalPageFrame = false,
  compatibilityMode?: number,
  paragraphMarkShapeInput?: NumberingMarkerShapeInput,
  /** Owner section's counter-turn for upright inline graphics. */
  resourceOrientation?: import('./types.js').UprightResourceOrientation,
): readonly LineLayout[] {
  let sourceOffset = 0;
  const consumedByRun = new Map<number, number>();
  const sequenceOwners = new Map<NonNullable<LayoutTextSeg['sourceTextSequence']>, number>();
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
  const specifiedParagraph = specifiedTextParagraphIsHomogeneous(paragraph);
  return retainPhysicalLines(measured.lines.map((measuredLine, lineIndex) => {
    const raw = measuredLine.layout;
    const specified = specifiedParagraph && !paragraph.numbering
      ? specifiedTextLineMetrics(raw, context, paragraph, compatibilityMode, verticalPageFrame,
          paragraphMarkShapeInput)
      : null;
    const baselinePt = specified
      ? measuredLine.topYPt + specified.baselineOffsetPt
      : plannedBaselinePt(measuredLine, context);
    let lineStartOffset = Number.POSITIVE_INFINITY;
    let lineEndOffset = sourceOffset;
    const segments: MeasuredLinePlanSegment[] = [];
    for (const segment of raw.segments) {
      const runIndex = sourceRunIndex(segment);
      const sourceRun = runIndex === undefined ? undefined : paragraph.runs[runIndex];
      const occurrenceLength = segmentOccurrenceLength(segment);
      const sequence = segment.sourceTextSequence;
      const segmentOffset = sequence
        ? (occurrences.runStarts[sequence[0]!.runIndex] ?? sourceOffset) + (segment.sourceTextOffset ?? 0)
        : runIndex === undefined
        ? sourceOffset
        : (occurrences.runStarts[runIndex] ?? sourceOffset) + (consumedByRun.get(runIndex) ?? 0);
      if (runIndex !== undefined && !sequence) {
        consumedByRun.set(runIndex, (consumedByRun.get(runIndex) ?? 0) + occurrenceLength);
      }
      lineStartOffset = Math.min(lineStartOffset, segmentOffset);
      lineEndOffset = Math.max(lineEndOffset, segmentOffset + occurrenceLength);
      if ('isTab' in segment) {
        const tab = segment as LayoutTabSeg;
        const leader = tab.leader ?? 'none';
        let leaderShape: MeasuredTabPlanSegment['leaderShape'];
        let underline: MeasuredTabPlanSegment['underline'];
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
          underlineStyle?: string | null;
          underlineColor?: string | null;
        }>) | undefined;
        const shapeTabGlyph = (glyph: string) => {
          if (!textService) throw new Error('Formatted tab acquisition requires TextLayoutService');
          return textService.shape({
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
        };
        const selectedMetric = (glyph: string): RetainedInkMetric => {
          const shaped = shapeTabGlyph(glyph);
          const span = shaped.spans[0];
          if (!span || shaped.spans.length !== 1 || span.start !== 0 || span.end !== glyph.length) {
            throw new Error('Formatted tab probe requires one selected-face span');
          }
          return {
            ascentPt: span.ascentPt,
            descentPt: span.descentPt,
            ...(span.inkBounds ? { inkBounds: span.inkBounds } : {}),
          };
        };
        if (textSource?.underline) {
          const textColor = textSource.color ? `#${textSource.color}` : '#000000';
          underline = {
            base: selectedMetric('M'),
            probe: selectedMetric('_'),
            color: richRun?.underlineColor && richRun.underlineColor !== 'auto'
              ? `#${richRun.underlineColor}` : textColor,
            ...(richRun?.underlineStyle ? { authoredStyle: richRun.underlineStyle } : {}),
          };
        }
        if (leader !== 'none') {
          if (!textService) {
            throw new Error('Tab leader acquisition requires TextLayoutService');
          }
          const glyph = leader === 'hyphen' ? '-'
            : leader === 'underscore' || leader === 'heavy' ? '_'
              : leader === 'middleDot' ? '·' : '.';
          const shape = shapeTabGlyph(glyph);
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
          ...(underline ? { underline } : {}),
          ...(leaderShape ? { leaderShape } : {}),
        });
      } else if ('imagePath' in segment) {
        const image = segment as LayoutImageSeg;
        if (image.anchor) continue;
        const runIndex = sourceRunIndex(segment);
        const occurrence = runSource(source, runIndex ?? 0);
        if (image.inlineShape) {
          segments.push({
            kind: 'inline-drawing',
            range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
            drawingId: `${paragraphId}:drawing:${runIndex ?? 0}`,
            measuredWidthPt: image.measuredWidth,
            widthPt: image.widthPt,
            heightPt: image.heightPt,
            topOffsetPt: -image.heightPt,
          });
          sourceOffset = Math.max(sourceOffset, segmentOffset + occurrenceLength);
          continue;
        }
        if (image.unavailableResourceKind) {
          segments.push({
            kind: 'unavailable-resource',
            range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
            resourceKind: image.unavailableResourceKind,
            measuredWidthPt: image.measuredWidth,
            widthPt: image.widthPt,
            heightPt: image.heightPt,
            topOffsetPt: -image.heightPt,
            drawingId: unavailableDrawingId(source, runIndex ?? 0),
          });
          sourceOffset = Math.max(sourceOffset, segmentOffset + occurrenceLength);
          continue;
        }
        const resourceKind = image.chart ? 'chart' : 'image';
        const resourceKey = image.chartResourceKey
          ?? (image.chart ? chartResourceKey(occurrence) : imageResourceKey(occurrence, image.imagePath));
        segments.push({
          kind: 'resource', range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
          ...(runIndex === undefined ? {} : { sourceRunIndex: runIndex }),
          resourceKey, resourceKind, measuredWidthPt: image.measuredWidth,
          widthPt: image.widthPt, heightPt: image.heightPt, topOffsetPt: -image.heightPt,
          ...(resourceOrientation ? { orientation: resourceOrientation } : {}),
        });
      } else if ('math' in segment) {
        const math = segment as LayoutMathSeg;
        segments.push({
          kind: 'resource',
          range: { start: segmentOffset, end: segmentOffset + occurrenceLength },
          ...(sourceRunIndex(segment) === undefined
            ? {} : { sourceRunIndex: sourceRunIndex(segment) }),
          resourceKey: math.mathResourceKey, resourceKind: 'math',
          measuredWidthPt: math.measuredWidth, widthPt: math.measuredWidth,
          heightPt: math.mathAscent + math.mathDescent, topOffsetPt: -math.mathAscent,
        });
      } else {
        let sourceRuns: TextPlacement['sourceRuns'];
        if (sequence) {
          const localStart = segment.sourceTextOffset ?? 0;
          const localEnd = localStart + occurrenceLength;
          let index = sequenceOwners.get(sequence) ?? 0;
          while (index < sequence.length && sequence[index]!.end <= localStart) index++;
          sequenceOwners.set(sequence, index);
          const owners: NonNullable<TextPlacement['sourceRuns']>[number][] = [];
          for (; index < sequence.length && sequence[index]!.start < localEnd; index++) {
            const owner = sequence[index]!;
            const original = paragraph.runs[owner.runIndex];
            owners.push({ sourceRunIndex: owner.runIndex,
              range: { start: segmentOffset + Math.max(owner.start, localStart) - localStart,
                end: segmentOffset + Math.min(owner.end, localEnd) - localStart },
              ...(original?.type === 'field' ? { role: 'field-result' as const, dependency: fieldDependency(original) } : {}),
            });
          }
          sourceRuns = owners;
        }
        segments.push(textPlanSegment(
          segment as LayoutTextSeg, paragraph, segmentOffset,
          paragraphCharacterGrid(context),
          sourceRun,
          verticalGlyphMeasurement,
          sourceRuns,
        ));
      }
      sourceOffset = Math.max(sourceOffset, segmentOffset + occurrenceLength);
    }
    const onlyMath = raw.segments.length === 1 && 'math' in (raw.segments[0] ?? {} as object)
      ? raw.segments[0] as LayoutMathSeg
      : undefined;
    const planned = planLine({
      paragraphXPt, availableWidthPt, alignment: paragraph.alignment,
      baseRtl: context.baseRtl,
      isFirstLine: lineIndex === 0,
      isLastLine: (raw.physicalLineIndex ?? lineIndex)
        === (measured.lines.at(-1)!.layout.physicalLineIndex ?? measured.lines.length - 1),
      stretchLastLine: context.stretchLastLine,
      exactLineSpacing: context.lineSpacing?.rule === 'exact',
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
        ...(raw.marginExtension ? { marginExtensionPt: raw.marginExtension } : {}),
        justifiedCompressionPt: raw.justifiedCompressionPx,
        gapPlan: raw.gapPlan,
        endsWithBreak: raw.endsWithBreak ?? false,
        segments,
      },
    });
    return raw.availWidth !== undefined && (raw.availWidth < availableWidthPt || (raw.xOffset ?? 0) !== 0)
      ? { ...planned, wrapBounds: {
          xPt: paragraphXPt + (raw.xOffset ?? 0) + (lineIndex === 0 ? Math.min(0, context.firstIndentPt) : 0),
          yPt: measuredLine.topYPt,
          widthPt: raw.availWidth - (lineIndex === 0 ? Math.min(0, context.firstIndentPt) : 0),
          heightPt: measuredLine.advancePt,
        } }
      : planned;
  }), measured.lines.map((line, index) => line.layout.physicalLineIndex ?? index));
}

/** Retain §17.18.84 bar-tab rules for every laid-out line. A bar is measured
 * from the paragraph's logical leading page margin, but never participates in
 * tab advancement. Its zero-width rule is painted as a device hairline. */
export function attachBarTabRules(
  lines: readonly LineLayout[],
  columnXPt: number,
  columnWidthPt: number,
  baseRtl: boolean,
  tabStops: readonly import('../types.js').TabStop[],
): readonly LineLayout[] {
  const bars = tabStops.filter((stop) => stop.alignment === 'bar');
  if (bars.length === 0) return lines;
  return lines.map((line) => ({
    ...line,
    barTabRules: bars.map((bar) => {
      const xPt = baseRtl
        ? columnXPt + columnWidthPt - bar.pos
        : columnXPt + bar.pos;
      return {
        from: { xPt, yPt: line.bounds.yPt },
        to: { xPt, yPt: line.bounds.yPt + line.bounds.heightPt },
        color: '#000000',
        widthPt: 0,
        authoredStyle: 'single',
        style: 'solid' as const,
      };
    }),
  }));
}

function offsetRange(range: import('./types.js').TextRange, delta: number) {
  return { start: range.start + delta, end: range.end + delta };
}

function rebaseMeasuredLineRanges(
  lines: readonly LineLayout[],
  sourceRangeStart: number,
): readonly LineLayout[] {
  if (!Number.isFinite(sourceRangeStart) || sourceRangeStart < 0) {
    throw new RangeError('Paragraph continuation source range must be finite and non-negative');
  }
  const first = lines[0];
  if (!first) return lines;
  const delta = sourceRangeStart - first.range.start;
  if (delta === 0) return lines;
  return lines.map((line) => ({
    ...line,
    range: offsetRange(line.range, delta),
    placements: line.placements.map((placement) => {
      const range = offsetRange(placement.range, delta);
      if (placement.kind !== 'text') return { ...placement, range };
      return {
        ...placement,
        range,
        ...(placement.sourceRuns ? { sourceRuns: placement.sourceRuns.map(owner => ({
          ...owner, range: offsetRange(owner.range, delta),
        })) } : {}),
        clusters: placement.clusters.map((cluster) => ({
          ...cluster,
          range: offsetRange(cluster.range, delta),
        })),
        paintOps: placement.paintOps.map((operation) => ({
          ...operation,
          range: offsetRange(operation.range, delta),
        })),
      };
    }),
  }));
}

/** A mark-only paragraph still owns a real line box and baseline when numbering
 * paints there. Materializing that host through `planLine` keeps the marker on
 * the same retained geometry path as a marker followed by body text. */
function numberingMarkerHostLine(
  measured: MeasuredParagraph,
  paragraph: ParagraphAcquisitionInput,
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

type PublicAnchorPositionRun = Extract<
  ParagraphAcquisitionInput['runs'][number],
  { type: 'shape' | 'image' | 'chart' }
>;

function resolvedPublicAnchorLayoutRect(
  run: PublicAnchorPositionRun,
  options: ParagraphAcquisitionOptions,
): LayoutRect | null {
  const frames = options.anchorFrames;
  const horizontalReference = run.anchorXRelativeFrom
    ?? (run.anchorXFromMargin ? 'margin' : 'page');
  const verticalReference = run.anchorYRelativeFrom
    ?? (run.anchorYFromPara ? 'paragraph' : 'page');
  const page = frames?.page;
  const margin = frames?.margin;
  const leftMarginFrame = page && margin ? {
    xPt: page.xPt,
    yPt: page.yPt,
    widthPt: Math.max(0, margin.xPt - page.xPt),
    heightPt: page.heightPt,
  } : null;
  const rightMarginFrame = page && margin ? {
    xPt: margin.xPt + margin.widthPt,
    yPt: page.yPt,
    widthPt: Math.max(
      0,
      page.xPt + page.widthPt - margin.xPt - margin.widthPt,
    ),
    heightPt: page.heightPt,
  } : null;
  const topMarginFrame = page && margin ? {
    xPt: page.xPt,
    yPt: page.yPt,
    widthPt: page.widthPt,
    heightPt: Math.max(0, margin.yPt - page.yPt),
  } : null;
  const bottomMarginFrame = page && margin ? {
    xPt: page.xPt,
    yPt: margin.yPt + margin.heightPt,
    widthPt: page.widthPt,
    heightPt: Math.max(
      0,
      page.yPt + page.heightPt - margin.yPt - margin.heightPt,
    ),
  } : null;
  const evenPage = frames?.pageParity === 'even';
  const insideMarginFrame = evenPage ? rightMarginFrame : leftMarginFrame;
  const outsideMarginFrame = evenPage ? leftMarginFrame : rightMarginFrame;
  const horizontalFrame = horizontalReference === 'page' ? page
    : horizontalReference === 'column' || horizontalReference === 'character' ? frames?.column
      : horizontalReference === 'leftMargin' ? leftMarginFrame
        : horizontalReference === 'rightMargin' ? rightMarginFrame
          : horizontalReference === 'insideMargin' ? insideMarginFrame
            : horizontalReference === 'outsideMargin' ? outsideMarginFrame
              : margin;
  const verticalFrame = verticalReference === 'paragraph' ? {
    xPt: options.placement.paragraphXPt,
    yPt: options.placement.startYPt,
    widthPt: options.placement.availableWidthPt,
    heightPt: 0,
  } : verticalReference === 'line' || verticalReference === 'character' ? {
    xPt: options.placement.paragraphXPt,
    yPt: options.placement.startYPt,
    widthPt: options.placement.availableWidthPt,
    heightPt: 0,
  } : verticalReference === 'page' ? page
    : verticalReference === 'column' ? frames?.column
      : verticalReference === 'topMargin' ? topMarginFrame
        : verticalReference === 'bottomMargin' ? bottomMarginFrame
          : verticalReference === 'insideMargin'
            ? (evenPage ? bottomMarginFrame : topMarginFrame)
            : verticalReference === 'outsideMargin'
              ? (evenPage ? topMarginFrame : bottomMarginFrame)
              : margin;
  if (!horizontalFrame || !verticalFrame) return null;
  const widthPt = run.widthPt;
  const heightPt = run.heightPt;
  const offsetXPt = run.anchorXPt ?? 0;
  const offsetYPt = run.anchorYPt ?? 0;
  const pctPosH = run.type === 'shape' ? run.pctPosH : null;
  const pctPosV = run.type === 'shape' ? run.pctPosV : null;
  // pctPos and posOffset are an OOXML choice. Public hand-built values can
  // nevertheless carry both; preserve the established compatibility rule by
  // applying the explicit offset after the percentage.
  const xPt = pctPosH != null
    ? horizontalFrame.xPt + horizontalFrame.widthPt * pctPosH + offsetXPt
    : run.anchorXAlign === 'center'
      ? horizontalFrame.xPt + (horizontalFrame.widthPt - widthPt) / 2
      : run.anchorXAlign === 'right'
        || (run.anchorXAlign === 'outside' && !evenPage)
        || (run.anchorXAlign === 'inside' && evenPage)
        ? horizontalFrame.xPt + horizontalFrame.widthPt - widthPt
        : horizontalFrame.xPt + offsetXPt;
  const yPt = pctPosV != null
    ? verticalFrame.yPt + verticalFrame.heightPt * pctPosV + offsetYPt
    : run.anchorYAlign === 'center'
      ? verticalFrame.yPt + (verticalFrame.heightPt - heightPt) / 2
      : run.anchorYAlign === 'bottom'
        || (run.anchorYAlign === 'outside' && !evenPage)
        || (run.anchorYAlign === 'inside' && evenPage)
        ? verticalFrame.yPt + verticalFrame.heightPt - heightPt
        : verticalFrame.yPt + offsetYPt;
  return { xPt, yPt, widthPt, heightPt };
}

function resolvedShapeLayoutRect(
  shape: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'shape' }>,
  options: ParagraphAcquisitionOptions,
): LayoutRect {
  return resolvedPublicAnchorLayoutRect(shape, options) ?? {
    xPt: shape.anchorXPt + (shape.anchorXFromMargin ? options.placement.paragraphXPt : 0),
    yPt: shape.anchorYPt + (shape.anchorYFromPara ? options.placement.startYPt : 0),
    widthPt: shape.widthPt,
    heightPt: shape.heightPt,
  };
}

function drawingForShape(
  shape: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'shape' }>,
  rect: LayoutRect,
  options: ParagraphAcquisitionOptions,
  runIndex: number,
  inline = false,
): DrawingLayout {
  const source = runSource(options.source, runIndex);
  const plan = planShapeDrawing(
    shape,
    rect,
    options.environment.layoutServices?.text,
    shape.vmlTextPathInput,
    shape.fill?.fillType === 'image'
      ? imageResourceKey(source, shape.fill.imagePath)
      : undefined,
  );
  const commands = [plan.command];
  const diagnostics = shapePlanDiagnostics(plan, source);
  return {
    kind: 'drawing', id: `${options.id}:drawing:${runIndex}`, source,
    flowDomainId: options.flowDomainId, flowBounds: rect, inkBounds: rect, advancePt: 0,
    ordinaryFlow: false,
    commands,
    ...(diagnostics.length === 0 ? {} : { diagnostics }),
    ...(!inline ? { anchorLayer: {
      occurrenceId: `public-shape:${options.id}:${runIndex}`,
      behindDoc: shape.behindDoc === true,
      relativeHeight: Number.isFinite(shape.zOrder) ? shape.zOrder : runIndex,
      sourceOrder: runIndex,
      horizontalOwnership: shape.anchorXRelativeFrom === 'character'
        || shape.anchorXRelativeFrom === 'column'
        ? 'host' : 'page',
      verticalOwnership: shape.anchorYRelativeFrom === 'paragraph'
        || shape.anchorYRelativeFrom === 'line'
        || shape.anchorYRelativeFrom === 'character'
        || (!shape.anchorYRelativeFrom && shape.anchorYFromPara)
        ? 'host' : 'page',
    } } : {}),
  };
}

function publicAnchoredResourceDrawing(
  run: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'image' | 'chart' }>,
  options: ParagraphAcquisitionOptions,
  runIndex: number,
): DrawingLayout | null {
  if (!run.anchor || run.anchorAcquisitionInput) return null;
  const rect = resolvedPublicAnchorLayoutRect(run, options);
  if (!rect) return null;
  const verticalReference = run.anchorYRelativeFrom ?? (run.anchorYFromPara ? 'paragraph' : 'page');
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
      ...(options.environment.verticalPageFrame
        ? {
            orientation: uprightResourceOrientation(
              options.environment.verticalPageFrame,
              options.environment.pageWritingMode,
            )!,
          }
        : {}),
    }],
    anchorLayer: {
      occurrenceId: `public-anchor:${options.id}:${runIndex}`,
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
  { type: 'image' | 'chart' | 'shape' | 'unavailableDrawing' }
> & Readonly<{ anchorAcquisitionInput?: import('./anchor-input.js').AnchorAcquisitionInput }>;

function anchoredPayloadRun(
  run: ParagraphAcquisitionInput['runs'][number],
): run is AnchoredPayloadRun {
  return (
    run.type === 'image'
    || run.type === 'chart'
    || run.type === 'shape'
    || run.type === 'unavailableDrawing'
  )
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

function resizeDerivedAnchorRect(
  derived: LayoutRect,
  authored: LayoutRect,
  effective: LayoutRect,
): LayoutRect {
  const leftPt = authored.xPt - derived.xPt;
  const topPt = authored.yPt - derived.yPt;
  const rightPt = derived.xPt + derived.widthPt - authored.xPt - authored.widthPt;
  const bottomPt = derived.yPt + derived.heightPt - authored.yPt - authored.heightPt;
  return {
    xPt: effective.xPt - leftPt,
    yPt: effective.yPt - topPt,
    widthPt: Math.max(0, effective.widthPt + leftPt + rightPt),
    heightPt: Math.max(0, effective.heightPt + topPt + bottomPt),
  };
}

/** Keep DrawingML commands in an upright physical, drawing-local frame and map
 * them into the logical page with the section coordinate space's canonical
 * inverse around the anchor's retained centre. Shape geometry and owned text
 * then share one orientation for both vertical-rl and vertical-lr sections. */
function uprightPhysicalDrawingTransform(
  rect: LayoutRect,
  physicalToLogical: Matrix2DData,
): Matrix2DData {
  return {
    a: physicalToLogical.a,
    b: physicalToLogical.b,
    c: physicalToLogical.c,
    d: physicalToLogical.d,
    e: rect.xPt + rect.widthPt / 2,
    f: rect.yPt + rect.heightPt / 2,
  };
}

function logicalRectToUprightDrawingLocal(
  rect: LayoutRect,
  transform: Matrix2DData,
): LayoutRect {
  const corners = [
    inverseMapAffinePoint(transform, rect),
    inverseMapAffinePoint(transform, {
      xPt: rect.xPt + rect.widthPt,
      yPt: rect.yPt,
    }),
    inverseMapAffinePoint(transform, {
      xPt: rect.xPt,
      yPt: rect.yPt + rect.heightPt,
    }),
    inverseMapAffinePoint(transform, {
      xPt: rect.xPt + rect.widthPt,
      yPt: rect.yPt + rect.heightPt,
    }),
  ];
  if (corners.some((point) => point === null)) {
    throw new Error('Upright drawing transform must be invertible');
  }
  const points = corners as readonly import('./types.js').PointPt[];
  const xPt = Math.min(...points.map((point) => point.xPt));
  const yPt = Math.min(...points.map((point) => point.yPt));
  return {
    xPt,
    yPt,
    widthPt: Math.max(...points.map((point) => point.xPt)) - xPt,
    heightPt: Math.max(...points.map((point) => point.yPt)) - yPt,
  };
}

function uprightDrawingLocalRectToLogical(
  rect: LayoutRect,
  transform: Matrix2DData,
): LayoutRect {
  return transformRect(transform, rect);
}

/** Project a complete upright DrawingML anchor result into section-logical
 * coordinates through the section writing mode's canonical affine inverse. */
export function projectPhysicalAnchorResult(
  result: Extract<AnchorFrameResult, { status: 'resolved' }>,
  physicalToLogical: Matrix2DData,
): Extract<AnchorFrameResult, { status: 'resolved' }> {
  const projectEdges = (edges: Readonly<{
    topPt: number; rightPt: number; bottomPt: number; leftPt: number;
  }>) => {
    const projected = transformRectEdges(physicalToLogical, {
      top: edges.topPt,
      right: edges.rightPt,
      bottom: edges.bottomPt,
      left: edges.leftPt,
    });
    return {
      topPt: projected.top,
      rightPt: projected.right,
      bottomPt: projected.bottom,
      leftPt: projected.left,
    };
  };
  return {
    ...result,
    // `resolveAnchorFrame` reports diagnostics in the physical wp:positionH/V
    // axes. The retained vertical-page frame swaps those axes, so ownership and
    // reference-frame diagnostics must travel with the geometry they describe.
    axes: {
      horizontal: result.axes.vertical,
      vertical: result.axes.horizontal,
    },
    geometry: {
      ...result.geometry,
      objectFrame: transformRect(physicalToLogical, result.geometry.objectFrame),
      inkBounds: transformRect(physicalToLogical, result.geometry.inkBounds),
      wrapBounds: result.geometry.wrapBounds
        ? transformRect(physicalToLogical, result.geometry.wrapBounds)
        : null,
      size: {
        horizontal: result.geometry.size.vertical,
        vertical: result.geometry.size.horizontal,
      },
      parentEffectExtent: projectEdges(result.geometry.parentEffectExtent),
      wrap: {
        ...result.geometry.wrap,
        distances: projectEdges(result.geometry.wrap.distances),
        distanceSources: transformRectEdges(
          physicalToLogical,
          result.geometry.wrap.distanceSources,
        ),
        effectExtent: projectEdges(result.geometry.wrap.effectExtent),
        ...(result.geometry.wrap.polygon ? {
          polygon: {
            ...result.geometry.wrap.polygon,
            points: result.geometry.wrap.polygon.points.map((point) =>
              transformPoint(physicalToLogical, point)),
          },
        } : {}),
      },
    },
  };
}

function resizeResolvedAnchorGeometry(
  result: Extract<AnchorFrameResult, { status: 'resolved' }>,
  effectiveObjectFrame: LayoutRect,
): Extract<AnchorFrameResult, { status: 'resolved' }> {
  const authored = result.geometry.objectFrame;
  if (
    authored.xPt === effectiveObjectFrame.xPt
    && authored.yPt === effectiveObjectFrame.yPt
    && authored.widthPt === effectiveObjectFrame.widthPt
    && authored.heightPt === effectiveObjectFrame.heightPt
  ) return result;
  const scaleX = authored.widthPt === 0 ? 1 : effectiveObjectFrame.widthPt / authored.widthPt;
  const scaleY = authored.heightPt === 0 ? 1 : effectiveObjectFrame.heightPt / authored.heightPt;
  const polygon = result.geometry.wrap.polygon;
  return {
    ...result,
    geometry: {
      ...result.geometry,
      objectFrame: effectiveObjectFrame,
      inkBounds: resizeDerivedAnchorRect(result.geometry.inkBounds, authored, effectiveObjectFrame),
      wrapBounds: result.geometry.wrapBounds
        ? resizeDerivedAnchorRect(result.geometry.wrapBounds, authored, effectiveObjectFrame)
        : null,
      wrap: {
        ...result.geometry.wrap,
        polygon: polygon ? {
          ...polygon,
          points: polygon.points.map((point) => ({
            xPt: effectiveObjectFrame.xPt + (point.xPt - authored.xPt) * scaleX,
            yPt: effectiveObjectFrame.yPt + (point.yPt - authored.yPt) * scaleY,
          })),
        } : null,
      },
    },
  };
}

function retainedAnchorChildFrame(
  acquisition: NonNullable<AnchoredPayloadRun['anchorAcquisitionInput']>,
  outerFrame: LayoutRect,
  coordinateSpace?: Readonly<{
    physicalToLogical: Matrix2DData;
    logicalToPhysical: Matrix2DData;
  }>,
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
  const physicalOuter = coordinateSpace === undefined
    ? outerFrame
    : transformRect(coordinateSpace.logicalToPhysical, outerFrame);
  const scaleX = physicalOuter.widthPt / authoredWidthPt;
  const scaleY = physicalOuter.heightPt / authoredHeightPt;
  const physicalChild = {
    xPt: physicalOuter.xPt + child.offsetXPt * scaleX,
    yPt: physicalOuter.yPt + child.offsetYPt * scaleY,
    widthPt: child.widthPt * scaleX,
    heightPt: child.heightPt * scaleY,
  };
  return coordinateSpace === undefined
    ? physicalChild
    : transformRect(coordinateSpace.physicalToLogical, physicalChild);
}

function anchorAxisOwnership(
  result: Extract<AnchorFrameResult, { status: 'resolved' }>,
  axis: 'horizontal' | 'vertical',
  layoutInCell = false,
): 'page' | 'host' {
  const diagnostic = result.axes[axis];
  if (diagnostic.status !== 'resolved') return 'host';
  // ECMA-376 Part 1 §20.4.2.3: a layoutInCell object is positioned in the
  // existing cell. Cell acquisition resolves both axes in that local content
  // band, so the coordinates must travel with the host when the retained table
  // is placed on the page.
  if (layoutInCell) return 'host';
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
  readonly collision?: DrawingMLCollisionEntryPt;
  readonly cellContainmentBounds?: LayoutRect;
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
  externalExclusions: readonly WrapExclusion[],
  sameParagraphExclusions: readonly WrapExclusion[],
  externalCollisions: readonly DrawingMLCollisionEntryPt[],
  sameParagraphCollisions: readonly DrawingMLCollisionEntryPt[],
  /** Every anchor occurrence of this paragraph, computed once per paragraph. */
  paragraphOccurrenceIds: ReadonlySet<string>,
  /** First-line content top before any float moves the line down. */
  contentStartYPt: number,
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
  const behavior = outer.run.anchorAcquisitionInput.behavior;
  const layoutInCellFrame = behavior.layoutInCellStatus === 'valid'
    && behavior.layoutInCell === true
    && options.anchorCellBounds !== undefined
    ? options.anchorCellBounds
    : null;
  const acquiredResult = resolveAnchorFrame({
    acquisition: outer.run.anchorAcquisitionInput,
    frames: {
      page: baseFrames?.page
        ? layoutInCellFrame
          ? { ...baseFrames.page, ...layoutInCellFrame }
          : baseFrames.page
        : null,
      margin: baseFrames?.margin
        ? layoutInCellFrame
          ? { ...baseFrames.margin, ...layoutInCellFrame }
          : baseFrames.margin
        : null,
      column: baseFrames?.column
        ? layoutInCellFrame
          ? { ...baseFrames.column, ...layoutInCellFrame }
          : wordMode14ColumnOriginApplies(outer.run.anchorAcquisitionInput, options)
            ? {
                ...baseFrames.column,
                xPt: wordMode14ColumnLineStartOrigin(
                  baseFrames.column,
                  externalExclusions,
                  paragraphOccurrenceIds,
                  contentStartYPt,
                  lines[0]?.bounds.heightPt ?? 0,
                ),
              }
            : baseFrames.column
        : null,
      paragraph: {
        xPt: options.placement.paragraphXPt,
        yPt: options.placement.startYPt + (
          outer.run.anchorAcquisitionInput.vertical.relativeFrom === 'paragraph'
          && outer.run.anchorAcquisitionInput.wrap.kind === 'none'
          && behavior.allowOverlap === true
          && behavior.layoutInCell === true && options.anchorCellBounds !== undefined
          && options.environment.pageWritingMode === 'horizontal-tb'
            ? options.paragraphAnchorReferenceDeltaPt ?? 0 : 0
        ),
        widthPt: options.placement.availableWidthPt,
        heightPt: Math.max(0, paragraphHeightPt),
      },
      // WORD_GRID_PICTURE_LINE_ORIGIN keeps text leading separate from the
      // reference frame; paragraph ownership already precedes before-spacing.
      // The retained host index owns first-line policy. Distinct physical
      // lines can share a numeric top; equal coordinates do not transfer it.
      // Only posOffset was measured: translating the frame for wp:align would
      // also move the line's center/bottom edge, so alignment keeps the line.
      line: hostLineIndex === 0 && options.context.lineGrid.active && outer.run.type === 'image'
        && outer.run.anchorAcquisitionInput.vertical.choice.kind === 'offset'
        ? { ...line.bounds, yPt: wordGridPictureLineOriginPt(
            line.bounds.yPt, options.placement.startYPt, contentStartYPt,
          ) }
        : line.bounds,
      character: host.bounds,
      pageParity: baseFrames?.pageParity ?? null,
    },
  });
  if (acquiredResult.status !== 'resolved') {
    return { result: acquiredResult, textBoxes: [], hostLineIndex, hostRange: host.range };
  }
  // ECMA-376 §17.6.20 + §20.4.3.x: anchor positionH/V and extent describe the
  // upright physical drawing layer. Retained body flow is section-logical, so
  // project the complete resolved geometry once before wrap/collision/paint use.
  const physicalPage = options.environment.verticalPageFrame && baseFrames?.page
    ? uprightPhysicalExtent(baseFrames.page, options.environment.pageWritingMode)
    : undefined;
  const coordinateSpace = physicalPage === undefined
    ? undefined
    : createSectionRegionCoordinateSpace(
        options.environment.pageWritingMode,
        physicalPage,
      );
  const result = physicalPage === undefined
    ? acquiredResult
    : projectPhysicalAnchorResult(
        acquiredResult,
        coordinateSpace!.physicalToLogical,
      );
  if (
    behavior.behindDocStatus !== 'valid'
    || behavior.relativeHeightStatus !== 'valid'
    || behavior.behindDoc === null
    || behavior.relativeHeight === null
  ) {
    throw new Error('resolved anchor frame must retain required CT_Anchor behavior');
  }
  const authoredRect = result.geometry.objectFrame;
  let uprightTransform = physicalPage === undefined
    ? undefined
    : uprightPhysicalDrawingTransform(authoredRect, coordinateSpace!.physicalToLogical);
  const uprightEnvironment = uprightTransform ? {
    ...options.environment,
    // The section quarter turn is cancelled by the drawing frame. Text-box
    // direction now belongs only to a:bodyPr@vert, not w:sectPr@textDirection.
    verticalCJK: false,
    verticalPageFrame: false,
  } : options.environment;
  const commands: DrawingPaintCommand[] = [];
  const diagnostics: LayoutDiagnostic[] = [];
  const textBoxes: TextBoxLayout[] = [];
  const textBoxIds: string[] = [];
  if (
    behavior.allowOverlapStatus !== 'valid'
    || behavior.allowOverlap === null
    || behavior.layoutInCellStatus !== 'valid'
    || behavior.layoutInCell === null
  ) {
    throw new Error('resolved anchor frame must retain overlap and cell behavior');
  }
  const allowOverlap = behavior.allowOverlap;
  const layoutInCell = behavior.layoutInCell;
  // The outer drawing's final placement. With `storyPageFrames` its text box
  // story's page-placed content (positioned tables of its tables) resolves
  // against those page frames (stated in the box's frame), so the placement
  // is repeated below until they are the ones it implies.
  const placeOuter = (storyPageFrames?: StoryPageFrames) => {
    const acquiredShapeTextBoxes = new Map<number, TextBoxLayout>();
    let placedRect = authoredRect;
    let placedUpright = uprightTransform;
    // Translation of the outer text box from the frame it was acquired in.
    let textBoxShift = { xPt: 0, yPt: 0 };
    if (outer.run.type === 'shape' && outer.run.anchorAcquisitionInput!.group === null) {
      const source = runSource(options.source, outer.runIndex);
      const textBoxRect = placedUpright
        ? logicalRectToUprightDrawingLocal(authoredRect, placedUpright)
        : authoredRect;
      const acquired = acquireShapeTextBoxLayout(outer.run, textBoxRect, {
        id: `${options.id}:anchor-textbox:${occurrenceId}:${outer.runIndex}`,
        source,
        flowDomainId: options.flowDomainId,
        context: options.context,
        measurer: options.measurer,
        environment: uprightEnvironment,
        input: outer.run.textBoxInput,
        acquireCompleteStory: options.acquireCompleteStory,
        ...(placedUpright ? { coordinateSpace: 'upright-physical' as const } : {}),
        ...(storyPageFrames ? { pageFrames: storyPageFrames } : {}),
      });
      // The fitted text box keeps its anchor alignment. The acquired layout is
      // immutable retained geometry in textBoxRect's space (logical page, or the
      // upright drawing frame whose axes are the physical anchor axes), so the
      // alignment is a translation of that layout, not a second acquisition.
      const fitShift = acquired
        ? alignedAutofitTranslation(
            outer.run.anchorAcquisitionInput!,
            baseFrames?.pageParity ?? null,
            textBoxRect,
            acquired.flowBounds,
          )
        : { xPt: 0, yPt: 0 };
      const textBox = acquired && (fitShift.xPt !== 0 || fitShift.yPt !== 0)
        ? translateTextBox(acquired, fitShift)
        : acquired;
      textBoxShift = fitShift;
      if (textBox) {
        acquiredShapeTextBoxes.set(outer.runIndex, textBox);
        placedRect = placedUpright
          ? uprightDrawingLocalRectToLogical(textBox.flowBounds, placedUpright)
          : textBox.flowBounds;
      }
    }
    let placedResult = resizeResolvedAnchorGeometry(result, placedRect);
    const translateAnchor = (delta: Readonly<{ xPt: number; yPt: number }>): void => {
      placedRect = translateRect(placedRect, delta);
      if (placedUpright) placedUpright = {
        ...placedUpright,
        e: placedUpright.e + delta.xPt,
        f: placedUpright.f + delta.yPt,
      };
      else {
        const outerTextBox = acquiredShapeTextBoxes.get(outer.runIndex);
        if (outerTextBox) {
          acquiredShapeTextBoxes.set(
            outer.runIndex,
            translateTextBox(outerTextBox, delta),
          );
        }
        textBoxShift = {
          xPt: textBoxShift.xPt + delta.xPt,
          yPt: textBoxShift.yPt + delta.yPt,
        };
      }
      placedResult = resizeResolvedAnchorGeometry(result, placedRect);
    };
    // WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: a drawing carried to this page keeps
    // the frame resolved when its anchor paragraph was first laid out.
    const frozenFrame = options.frozenAnchorFrames?.get(occurrenceId);
    if (frozenFrame) {
      const delta = { xPt: frozenFrame.xPt - placedRect.xPt, yPt: frozenFrame.yPt - placedRect.yPt };
      if (delta.xPt !== 0 || delta.yPt !== 0) translateAnchor(delta);
    }
    const effectiveWrapBounds = placedResult.geometry.wrapBounds;
    const normativeCollision = !allowOverlap;
    if (normativeCollision) {
      // §20.4.2.3 object collision is independent of text wrapping. An
      // allowOverlap=true object keeps its resolved position: Word controls
      // (issue #1623) never move such a picture away from pictures anchored in
      // other paragraphs, in either compatibility mode, wrap kind, or reference.
      // ECMA-376 §20.4.2.3 requires displacement for every existing object
      // whose allowOverlap behavior makes it a collision participant.
      // Word has one narrower composition exception: a source-later page-owned
      // member below the already-authored layers in this SAME anchor paragraph
      // retains its authored position. Cross-paragraph entries remain blockers.
      const movingVerticalOwnership = anchorAxisOwnership(
        placedResult,
        'vertical',
        layoutInCell && options.anchorCellBounds !== undefined,
      );
      const sameParagraphBlockers = sameParagraphCollisions.filter((entry) =>
        !wordPreservesLowerLayerSameParagraphComposition(
          movingVerticalOwnership,
          behavior.relativeHeight!,
          entry.relativeHeight,
        ));
      const blockers: FloatPlacementParticipant[] = [...externalCollisions, ...sameParagraphBlockers]
        .filter((entry) => entry.occurrenceId !== occurrenceId)
        .map((entry) => ({
          occurrenceId: entry.occurrenceId,
          kind: 'drawingml',
          paragraphId: 0,
          bounds: entry.bounds,
          exclusionBounds: entry.bounds,
        }));
      const page = options.anchorFrames?.page;
      const rightBoundary = layoutInCell
        && options.anchorCellBounds
        ? options.anchorCellBounds.xPt + options.anchorCellBounds.widthPt
        : page
          ? page.xPt + page.widthPt
          : Number.POSITIVE_INFINITY;
      const displaced = resolveFloatPlacement({
        moving: {
          occurrenceId,
          kind: 'drawingml',
          paragraphId: 1,
          bounds: placedRect,
          exclusionBounds: effectiveWrapBounds ?? placedRect,
        },
        blockers,
        avoidance: { kind: 'drawingml-normative' },
        rightBoundaryPt: rightBoundary,
      });
      const delta = displaced.displacement;
      if (delta.xPt !== 0 || delta.yPt !== 0) translateAnchor(delta);
    }
    return Object.freeze({
      acquiredShapeTextBoxes,
      rect: placedRect,
      uprightTransform: placedUpright,
      effectiveResult: placedResult,
      textBoxShift,
    });
  };
  type PlacedOuter = ReturnType<typeof placeOuter>;
  // Page frames of the outer text box's story (story-page-frames.ts): the
  // destination page frames carried back through the drawing's final
  // placement into the box's frame — its upright drawing transform, autofit
  // alignment and anchor translations, then the translation host-following
  // axes still receive (hostFlowPageTranslationPt; page-owned axes receive
  // none, paint undoes it). Without that translation (a table cell before its
  // pagination places it) a host-following drawing's page is unknown here.
  const hostFlow = options.hostFlowPageTranslationPt;
  const hostBase = options.hostPageFrames !== undefined
    ? options.hostPageFrames
    : baseFrames?.page && baseFrames.margin ? { page: baseFrames.page, margin: baseFrames.margin } : null;
  const storyFramesFor = (placed: PlacedOuter): StoryPageFrames | null => {
    if (!hostBase) return null;
    const inCell = layoutInCell && options.anchorCellBounds !== undefined;
    const follows = {
      horizontal: anchorAxisOwnership(placed.effectiveResult, 'horizontal', inCell) === 'host',
      vertical: anchorAxisOwnership(placed.effectiveResult, 'vertical', inCell) === 'host',
    };
    if ((follows.horizontal || follows.vertical) && !hostFlow) return null;
    const toPage = composeAffine(
      translationAffine(
        follows.horizontal ? hostFlow!.xPt : 0,
        follows.vertical ? hostFlow!.yPt : 0,
      ),
      composeAffine(
        placed.uprightTransform ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        translationAffine(placed.textBoxShift.xPt, placed.textBoxShift.yPt),
      ),
    );
    return storyPageFramesThrough(hostBase, toPage);
  };
  let placedOuter = placeOuter();
  const ownsBandDependentStory = placedOuter.acquiredShapeTextBoxes.get(outer.runIndex)
    ?.story.bandDependent === true;
  const startFrames = ownsBandDependentStory ? storyFramesFor(placedOuter) : null;
  if (startFrames) {
    try {
      // Exact fixed point (solveStoryPageFrames); the limit is a resource
      // guard only.
      placedOuter = solveStoryPageFrames(startFrames, (frames) => {
        const placed = placeOuter(frames ?? undefined);
        return { value: placed, next: storyFramesFor(placed) };
      }, 16);
    } catch (error) {
      if (error instanceof ExactConvergenceError) {
        throw new LayoutInvariantError(
          'NON_CONVERGENCE',
          `anchored text-box page frames did not converge (${error.reason}; ${error.states.length} states)`,
        );
      }
      throw error;
    }
  }
  const { acquiredShapeTextBoxes, rect, effectiveResult } = placedOuter;
  uprightTransform = placedOuter.uprightTransform;
  for (const { run, runIndex } of ordered) {
    const source = runSource(options.source, runIndex);
    const acquisition = run.anchorAcquisitionInput as NonNullable<typeof run.anchorAcquisitionInput>;
    const commandRect = retainedAnchorChildFrame(acquisition, rect, coordinateSpace);
    const paintRect = uprightTransform
      ? logicalRectToUprightDrawingLocal(commandRect, uprightTransform)
      : commandRect;
    if (run.type === 'image') {
      commands.push({
        kind: 'resource', resourceKind: 'image',
        resourceKey: imageResourceKey(source, run.imagePath), rect: paintRect,
      });
    } else if (run.type === 'chart') {
      commands.push({
        kind: 'resource', resourceKind: 'chart',
        resourceKey: chartResourceKey(source), rect: paintRect,
      });
    } else if (run.type === 'unavailableDrawing') {
      commands.push({ kind: 'noop' });
      diagnostics.push(unavailableDrawingDiagnostic(run.resourceKind, source));
    } else {
      const childTransform = acquisition.group?.resolvedChildFrame;
      const plannedRun = childTransform ? {
        ...run,
        rotation: childTransform.rotationDeg,
        flipH: childTransform.flipH,
        flipV: childTransform.flipV,
      } : run;
      const plan = planShapeDrawing(
        plannedRun,
        paintRect,
        options.environment.layoutServices?.text,
        run.vmlTextPathInput,
        run.fill?.fillType === 'image'
          ? imageResourceKey(source, run.fill.imagePath)
          : undefined,
      );
      commands.push(plan.command);
      diagnostics.push(...shapePlanDiagnostics(plan, source));
      const textBoxId = `${options.id}:anchor-textbox:${occurrenceId}:${runIndex}`;
      const textBox = acquiredShapeTextBoxes.get(runIndex) ?? acquireShapeTextBoxLayout(run, paintRect, {
        id: textBoxId,
        source,
        flowDomainId: options.flowDomainId,
        context: options.context,
        measurer: options.measurer,
        environment: uprightEnvironment,
        input: run.textBoxInput,
        acquireCompleteStory: options.acquireCompleteStory,
        ...(uprightTransform ? { coordinateSpace: 'upright-physical' as const } : {}),
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
    ...(uprightTransform ? {
      orientation: 'upright-physical' as const,
      transform: uprightTransform,
    } : {}),
    commands,
    ...(diagnostics.length === 0 ? {} : { diagnostics: Object.freeze(diagnostics) }),
    anchorLayer: {
      occurrenceId,
      behindDoc: behavior.behindDoc,
      relativeHeight: behavior.relativeHeight,
      sourceOrder: outer.runIndex,
      horizontalOwnership: anchorAxisOwnership(
        effectiveResult,
        'horizontal',
        behavior.layoutInCell && options.anchorCellBounds !== undefined,
      ),
      verticalOwnership: anchorAxisOwnership(
        effectiveResult,
        'vertical',
        behavior.layoutInCell && options.anchorCellBounds !== undefined,
      ),
      ...(behavior.layoutInCell && options.anchorCellBounds
        ? { layoutInCell: true as const }
        : {}),
      ...(behavior.layoutInCell
        && wordLayoutInCellOwnsRowContainment(
          behavior.allowOverlap,
          effectiveResult.geometry.wrap.kind,
        )
        && options.anchorCellBounds
        ? { cellContainment: true as const }
        : {}),
    },
    ...(textBoxIds.length ? { textBoxIds } : {}),
  };
  const wrapBounds = effectiveResult.geometry.wrapBounds;
  const exclusion = wrapBounds && effectiveResult.geometry.wrap.kind !== 'none' ? {
    id: `${options.id}:anchor-exclusion:${occurrenceId}`,
    wrap: effectiveResult.geometry.wrap.kind,
    ...(effectiveResult.geometry.wrap.side
      ? { wrapSide: effectiveResult.geometry.wrap.side }
      : {}),
    bounds: wrapBounds,
    polygon: effectiveResult.geometry.wrap.polygon?.points ?? rectanglePolygon(wrapBounds),
    anchorOccurrenceId: occurrenceId,
    ...(options.anchorLineExemptions?.has(occurrenceId)
      ? { anchorLineExemptTopPt: options.anchorLineExemptions.get(occurrenceId)! }
      : {}),
    ...(effectiveResult.geometry.wrap.kind === 'tight'
      && options.environment.compatibilityMode !== undefined
      && options.environment.compatibilityMode <= 14
      ? { wordMode14TightAnchor: true }
      : {}),
    verticalOwnership: anchorAxisOwnership(
      effectiveResult,
      'vertical',
      behavior.layoutInCell && options.anchorCellBounds !== undefined,
    ),
  } satisfies WrapExclusion : undefined;
  const collision: DrawingMLCollisionEntryPt = {
    occurrenceId,
    bounds: rect,
    horizontalOwnership: anchorAxisOwnership(
      effectiveResult,
      'horizontal',
      behavior.layoutInCell && options.anchorCellBounds !== undefined,
    ),
    verticalOwnership: anchorAxisOwnership(
      effectiveResult,
      'vertical',
      behavior.layoutInCell && options.anchorCellBounds !== undefined,
    ),
    ...(behavior.relativeHeight !== null
      ? { relativeHeight: behavior.relativeHeight }
      : {}),
  };
  return {
    result: effectiveResult, drawing, exclusion, collision, textBoxes,
    ...(behavior.layoutInCell
      && wordLayoutInCellOwnsRowContainment(
        behavior.allowOverlap,
        effectiveResult.geometry.wrap.kind,
      )
      && options.anchorCellBounds
      ? { cellContainmentBounds: rect }
      : {}),
    hostLineIndex, hostRange: host.range,
  };
}

export type CompleteTextBoxStoryAcquirer = (
  request: Readonly<{
    source: SourceRef;
    container: FlowContainer;
    /** Parser-owned vertical-page drawings acquire their shape content in the
     * same upright local frame as the surrounding DrawingML geometry. */
    coordinateSpace?: 'section-logical' | 'upright-physical';
    /** The destination page and margin rectangles in the story's own
     * coordinates (StoryLayoutAcquisitionInput.pageFrames). */
    pageFrames?: StoryPageFrames;
  }>,
) => StoryLayout;

export interface ShapeTextBoxAcquisitionOptions {
  readonly id: string;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly context: ParagraphLayoutContext;
  readonly measurer: TextMeasurer;
  readonly environment: ParagraphMeasurementEnvironment;
  readonly input?: TextBoxAcquisitionInput;
  readonly acquireCompleteStory?: CompleteTextBoxStoryAcquirer;
  readonly coordinateSpace?: 'section-logical' | 'upright-physical';
  /**
   * The destination page and margin rectangles in the coordinates of `rect`
   * (the returned layout's, before the caller's own transforms), when the
   * caller knows how it places the box on the page. The story's page-placed
   * content then resolves through the box's own story placement as well
   * (acquireShapeTextBoxLayout).
   */
  readonly pageFrames?: StoryPageFrames;
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
    rightIndentGrid: {
      ...inherited.rightIndentGrid,
      paragraphAllowsAdjustment: paragraph.adjustRightInd !== false,
    },
    physicalIndentLeftPt: baseRtl ? paragraph.indentRight : paragraph.indentLeft,
    physicalIndentRightPt: baseRtl ? paragraph.indentLeft : paragraph.indentRight,
    firstIndentPt: paragraph.indentFirst,
    lineSpacing: paragraph.lineSpacing,
    spaceBeforePt: paragraph.spaceBefore,
    spaceAfterPt: paragraph.spaceAfter,
    baseRtl,
    isJustified: jcIsFullyJustified(paragraph.alignment),
    stretchLastLine: jcStretchesLastLine(paragraph.alignment),
    tabStops: effectiveParagraphTabStops(paragraph),
    hasRuby,
    hasEastAsianText,
  };
}

type RetainedTextBoxVerticalMode = NonNullable<TextBoxLayout['verticalMode']>;

function retainedTextBoxVerticalMode(value: string | null | undefined): RetainedTextBoxVerticalMode | undefined {
  return wordTextBoxVerticalMode(value);
}

/** ECMA-376 §21.1.2.1.1 `CT_TextBodyProperties@anchor`: resolve the text
 * story's block-axis position inside the inset content rectangle. The offset is
 * retained in point-space story geometry so paint does not inspect the model or
 * reconstruct text-box alignment. */
function textBoxAnchorOffsetPt(
  anchor: string | null | undefined,
  availableExtentPt: number,
  storyExtentPt: number,
): number {
  const remainingPt = Math.max(0, availableExtentPt - storyExtentPt);
  if (anchor === 'b') return remainingPt;
  if (anchor === 'ctr') return remainingPt / 2;
  return 0;
}

function orientVerticalTextBoxParagraph(
  paragraph: ParagraphLayout,
  mode: RetainedTextBoxVerticalMode,
  innerBounds: LayoutRect,
  insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>,
  wordArt = false,
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
        + (wordArt ? 0 : insets.bottomPt - insets.leftPt) + rubyReservePt
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
        ? eastAsianUprightPaintOps(placement, wordArt)
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

function orientVerticalTextBoxTable(
  table: import('./types.js').TableLayout,
  mode: RetainedTextBoxVerticalMode,
  wordArt = false,
): import('./types.js').TableLayout {
  const orientChild = (
    child: ParagraphLayout | import('./types.js').TableLayout,
    cellBounds: LayoutRect,
  ): ParagraphLayout | import('./types.js').TableLayout =>
    child.kind === 'paragraph'
      ? orientVerticalTextBoxParagraph(
          child,
          // A table cell owns its own horizontal line frame. Mongolian column
          // reflection belongs to the outer text-box story, while its glyphs
          // use the same upright/sideways rule as eaVert.
          mode === 'mongolianVert' ? 'eaVert' : mode,
          cellBounds,
          { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
          wordArt,
        )
      : orientVerticalTextBoxTable(child, mode, wordArt);
  const oriented: import('./types.js').TableLayout = {
    ...table,
    rows: table.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map((block) => ({
          ...block,
          layout: orientChild(block.layout, cell.contentBounds),
        })),
      })),
    })),
  };
  const sourceMemo = new Map<
    import('./types.js').FloatingTablePlacementLayout,
    import('./types.js').FloatingTablePlacementLayout
  >();
  const orientSource = (
    placement: import('./types.js').FloatingTablePlacementLayout,
  ): import('./types.js').FloatingTablePlacementLayout => {
    const prior = sourceMemo.get(placement);
    if (prior) return prior;
    const result = {
      ...placement,
      child: orientVerticalTextBoxTable(placement.child, mode, wordArt),
    };
    sourceMemo.set(placement, result);
    return result;
  };
  const floatingTables = table.floatingTables?.map(orientSource);
  const resolvedFloatingTables = table.resolvedFloatingTables?.map(
    (placement) => {
      const source = orientSource(placement.source);
      return { ...placement, source, child: source.child };
    },
  );
  return {
    ...oriented,
    ...(floatingTables ? { floatingTables } : {}),
    ...(resolvedFloatingTables ? { resolvedFloatingTables } : {}),
  };
}

function orientVerticalTextBoxStory(
  story: StoryLayout,
  mode: RetainedTextBoxVerticalMode,
  innerBounds: LayoutRect,
  insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>,
  wordArt = false,
): StoryLayout {
  return {
    ...story,
    blocks: story.blocks.map((block) => {
      if (block.kind === 'paragraph') {
        return orientVerticalTextBoxParagraph(block, mode, innerBounds, insets, wordArt);
      }
      if (block.kind === 'table') return orientVerticalTextBoxTable(block, mode, wordArt);
      throw new Error(`Text-box story contains unsupported retained node: ${block.kind}`);
    }),
  };
}

function translateTextBoxStory(
  story: StoryLayout,
  deltaYPt: number,
  translateClipBounds = true,
): StoryLayout {
  if (deltaYPt === 0) return story;
  const delta = { xPt: 0, yPt: deltaYPt };
  return {
    ...story,
    flowBounds: translateRect(story.flowBounds, delta),
    inkBounds: translateRect(story.inkBounds, delta),
    ...(story.clipBounds ? {
      clipBounds: translateClipBounds ? translateRect(story.clipBounds, delta) : story.clipBounds,
    } : {}),
    blocks: story.blocks.map((block) => {
      if (block.kind === 'paragraph') return translateParagraphLayout(block, delta);
      if (block.kind === 'table') return translateTextBoxStoryTable(block, delta);
      throw new Error(`Text-box story contains unsupported retained node: ${block.kind}`);
    }),
  };
}


/**
 * Translation that keeps an aligned anchor's `wp:align` values when spAutoFit
 * gives a text box a fitted extent different from its authored one
 * (ECMA-376 §20.4.2.2 wp:align / §20.4.3.2 ST_AlignV, §21.1.2.1.3 spAutoFit): the aligned edge
 * belongs to the drawn extent. A trailing value (`right`/`bottom`, or
 * `inside`/`outside` by page parity) keeps the fitted box's trailing edge on
 * the authored one and `center` keeps the centre. Offsets, percentages and
 * leading values keep the fitted box where acquisition placed it. Both frames
 * are in the same coordinate space, whose axes are the anchor's physical
 * positionH/positionV axes (the logical page, or the upright drawing frame of
 * a vertical section).
 */
export function alignedAutofitTranslation(
  input: Readonly<Pick<import('./anchor-input.js').AnchorAcquisitionInput, 'horizontal' | 'vertical'>>,
  pageParity: 'odd' | 'even' | null,
  authored: LayoutRect,
  fitted: LayoutRect,
): LayoutTranslation {
  const axisShift = (
    axis: 'horizontal' | 'vertical',
    authoredStart: number,
    authoredSize: number,
    fittedStart: number,
    fittedSize: number,
  ): number => {
    const choice = input[axis].choice;
    if (choice.kind !== 'align' || authoredSize === fittedSize) return 0;
    const placement = alignedAnchorPlacement(axis, choice.value, pageParity);
    const shift = placement === 'trailing'
      ? authoredStart + authoredSize - fittedStart - fittedSize
      : placement === 'center'
        ? authoredStart + authoredSize / 2 - fittedStart - fittedSize / 2
        : 0;
    return Number.isFinite(shift) ? shift : 0;
  };
  return {
    xPt: axisShift('horizontal', authored.xPt, authored.widthPt, fitted.xPt, fitted.widthPt),
    yPt: axisShift('vertical', authored.yPt, authored.heightPt, fitted.yPt, fitted.heightPt),
  };
}

/** Acquires a DrawingML/WPS text body through the same paragraph measurement
 * and retained layout seam used by ordinary WordprocessingML paragraphs. */
export function acquireShapeTextBoxLayout(
  shape: import('./types.js').DeepReadonly<ShapeRun>,
  rect: LayoutRect,
  options: ShapeTextBoxAcquisitionOptions,
): TextBoxLayout | undefined {
  const source = options.source;
  const acquisition: TextBoxAcquisitionInput = options.input ?? {
    kind: 'compatibility',
    source: {
      story: 'textbox',
      storyInstance: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
      path: [],
    },
    paragraphs: normalizeTextBoxInput(shape, {
      story: 'textbox',
      storyInstance: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
      path: [],
    }),
  };
  const storySource = acquisition.source;
  const blockCount = acquisition.kind === 'complete'
    ? acquisition.blockCount
    : acquisition.paragraphs.length;
  if (blockCount === 0) return undefined;
  const stackedWordArt = shape.textVert === 'wordArtVert' || shape.textVert === 'wordArtVertRtl';
  const verticalMode = retainedTextBoxVerticalMode(shape.textVert);
  const contentBounds: LayoutRect = verticalMode ? {
    xPt: -rect.heightPt / 2,
    yPt: -rect.widthPt / 2,
    widthPt: rect.heightPt,
    heightPt: rect.widthPt,
  } : rect;
  const normalized = acquisition.kind === 'compatibility'
    ? acquisition.paragraphs
    : Object.freeze([]);
  // A clockwise WordArt frame maps inline start/end to physical top/bottom,
  // and logical block start/end to physical right/left. The LTR-column
  // projection below mirrors the block axis without swapping authored insets.
  const insets = stackedWordArt ? {
    topPt: shape.textInsetR ?? 0, rightPt: shape.textInsetB ?? 0,
    bottomPt: shape.textInsetL ?? 0, leftPt: shape.textInsetT ?? 0,
  } : {
    topPt: shape.textInsetT ?? 0, rightPt: shape.textInsetR ?? 0,
    bottomPt: shape.textInsetB ?? 0, leftPt: shape.textInsetL ?? 0,
  };
  const innerBounds = {
    xPt: contentBounds.xPt + insets.leftPt,
    yPt: contentBounds.yPt + insets.topPt,
    widthPt: Math.max(0, contentBounds.widthPt - insets.leftPt - insets.rightPt),
    heightPt: Math.max(0, contentBounds.heightPt - insets.topPt - insets.bottomPt),
  };
  const acquireCompleteStory = options.acquireCompleteStory;
  if (acquisition.kind === 'complete' && !acquireCompleteStory) {
    throw new Error('Complete text-box content requires the shared story acquisition adapter');
  }
  const acquireStory = (pageFrames?: StoryPageFrames): StoryLayout => acquireCompleteStory!({
    source: storySource,
    container: {
      id: `${options.id}:story`,
      kind: 'textbox',
      bounds: innerBounds,
      ...(stackedWordArt ? { quarterTurnMath: true } : {}),
      capacity: 'unbounded',
      ...(stackedWordArt && shape.textWrap === 'none' ? { noWrap: true } : {}),
    },
    coordinateSpace: options.coordinateSpace ?? 'section-logical',
    ...(pageFrames ? { pageFrames } : {}),
  });
  let yPt = contentBounds.yPt + insets.topPt;
  let previousInput: NormalizedTextBoxParagraphInput | null = null;
  const paragraphs = normalized.map((input, blockIndex) => {
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
        ...(stackedWordArt && shape.textWrap === 'none' ? { noWrap: true } : {}),
      },
      measurer: options.measurer,
      environment: options.environment,
      exclusions: [],
    });
    yPt += child.advancePt - child.spacing.afterPt;
    previousInput = input;
    return verticalMode ? orientVerticalTextBoxParagraph(child, verticalMode, innerBounds, insets, stackedWordArt) : child;
  });
  // The box composed around its story: autofit extent, anchor offset and
  // orientation. `storyToBox` carries story coordinates to the coordinates of
  // `rect`, the frame the caller places the box in.
  const compose = (completeStory: StoryLayout | undefined): Readonly<{
    layout: TextBoxLayout;
    storyToBox: Matrix2DData;
    storyFlowPt: Readonly<{ xPt: number; yPt: number }>;
  }> => {
    const fittedExtentPt = completeStory
      ? Math.max(0, completeStory.advancePt + insets.topPt + insets.bottomPt)
      : Math.max(0, yPt - contentBounds.yPt + insets.bottomPt);
    const mayAutofit = shape.textAutofit === 'sp' && blockCount > 0
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
    const fittedParagraphs = verticalMode && effectiveRect.widthPt !== rect.widthPt
      && verticalMode !== 'mongolianVert'
      ? paragraphs.map((paragraph) => translateParagraphY(
          paragraph,
          effectiveContentBounds.yPt - contentBounds.yPt,
        ))
      : paragraphs;
    const effectiveInnerBounds = {
      xPt: effectiveContentBounds.xPt + insets.leftPt,
      yPt: effectiveContentBounds.yPt + insets.topPt,
      widthPt: Math.max(
        0,
        effectiveContentBounds.widthPt - insets.leftPt - insets.rightPt,
      ),
      heightPt: Math.max(
        0,
        effectiveContentBounds.heightPt - insets.topPt - insets.bottomPt,
      ),
    };
    const paragraphFlowBounds = unionLayoutRects(fittedParagraphs.map((paragraph) => paragraph.flowBounds))
      ?? { xPt: effectiveInnerBounds.xPt, yPt: effectiveInnerBounds.yPt, widthPt: 0, heightPt: 0 };
    const paragraphInkBounds = unionLayoutRects(fittedParagraphs.map((paragraph) => paragraph.inkBounds))
      ?? { xPt: effectiveInnerBounds.xPt, yPt: effectiveInnerBounds.yPt, widthPt: 0, heightPt: 0 };
    let story: StoryLayout = completeStory ?? {
      story: 'textbox',
      flowBounds: paragraphFlowBounds,
      inkBounds: paragraphInkBounds,
      clipBounds: effectiveInnerBounds,
      blocks: fittedParagraphs,
      advancePt: Math.max(0, fittedExtentPt - insets.topPt - insets.bottomPt),
      diagnostics: [],
    };
    // `story` is still in its logical block-axis frame here. Vertical text-box
    // orientation may mirror glyph/line geometry, but `anchor` is defined against
    // this pre-orientation text-body extent. Retain the scalar now so the later
    // physical projection cannot change vertical anchoring semantics.
    const anchorStoryExtentPt = wordTextBoxVisibleAnchorExtentPt(story);
    const fittedShiftPt = completeStory && verticalMode
      ? effectiveContentBounds.yPt - contentBounds.yPt
      : 0;
    if (completeStory && verticalMode) {
      story = orientVerticalTextBoxStory(
        translateTextBoxStory(story, fittedShiftPt),
        verticalMode,
        effectiveInnerBounds,
        insets,
        stackedWordArt,
      );
    }
    // WordArt columns advance rightwards: anchoring shifts the mirrored local
    // block axis negatively so ctr/b move toward the physical trailing edge.
    const anchorShiftPt = (stackedWordArt ? -1 : 1) * textBoxAnchorOffsetPt(
      shape.textAnchor,
      effectiveInnerBounds.heightPt,
      anchorStoryExtentPt,
    );
    story = translateTextBoxStory(story, anchorShiftPt, false);
    // Issue #1668 Word controls: 0/30/90 degree shape rotations carry the
    // WordArt text frame; flipH keeps it readable, flipV turns it 180 degrees.
    // Retain the composed transform here so paint/indexing use the same frame.
    const textRotationDeg = (shape.textUpright ? 0 : shape.rotation ?? 0)
      + (stackedWordArt && shape.flipV ? 180 : 0);
    // Exact data for quarter turns: Math.cos(π/2) is not 0, and a quarter
    // turn must stay one (axis-aligned) for its story's page frames.
    const quarterTurns = Number.isInteger(textRotationDeg / 90)
      ? (((textRotationDeg / 90) % 4) + 4) % 4
      : null;
    const sin = quarterTurns === null
      ? Math.sin(textRotationDeg * Math.PI / 180) : [0, 1, 0, -1][quarterTurns]!;
    const cos = quarterTurns === null
      ? Math.cos(textRotationDeg * Math.PI / 180) : [1, 0, -1, 0][quarterTurns]!;
    const orientation: Matrix2DData = verticalMode ? {
      a: 0,
      b: verticalMode === 'vert270' ? -1 : 1,
      c: verticalMode === 'vert270' ? 1 : -1,
      d: 0,
      e: effectiveRect.xPt + effectiveRect.widthPt / 2,
      f: effectiveRect.yPt + effectiveRect.heightPt / 2,
    } : { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    // DrawingML text follows its shape rotation unless bodyPr@upright opts out.
    // Compose once here so canvas paint and the text index share the same frame.
    const centerX = effectiveRect.xPt + effectiveRect.widthPt / 2;
    const centerY = effectiveRect.yPt + effectiveRect.heightPt / 2;
    const transform = textRotationDeg === 0 ? orientation : composeAffine({
      a: cos, b: sin, c: -sin, d: cos,
      e: centerX - cos * centerX + sin * centerY,
      f: centerY - sin * centerX - cos * centerY,
    }, orientation);
    const layout: TextBoxLayout = deepFreezePlainData({
      kind: 'textbox', id: options.id, source: normalized[0]?.source ?? storySource,
      flowDomainId: `${options.flowDomainId}:textbox`, flowBounds: effectiveRect, inkBounds: effectiveRect,
      ...(shape.defaultTextColor ? {
        defaultTextColor: `#${shape.defaultTextColor.replace(/^#/u, '')}`,
      } : {}),
      // Word issue #1668 controls retain overflow in fixed stacked WordArt
      // boxes (wrap square/none and multi-paragraph cases); do not add a body clip.
      ...(shape.textAutofit === 'none' && !stackedWordArt ? { clipBounds: effectiveInnerBounds } : {}),
      advancePt: 0, ordinaryFlow: false, story,
      transform,
      writingMode: shape.textVert === 'vert270' ? 'vertical-lr' : shape.textVert ? 'vertical-rl' : 'horizontal-tb',
      insets,
      contentBounds: effectiveContentBounds,
      ...(verticalMode ? { verticalMode } : {}),
    });
    return {
      layout,
      storyToBox: composeAffine(transform, translationAffine(0, fittedShiftPt + anchorShiftPt)),
      storyFlowPt: { xPt: 0, yPt: fittedShiftPt + anchorShiftPt },
    };
  };
  if (acquisition.kind !== 'complete') return compose(undefined).layout;
  const boxPageFrames = options.pageFrames;
  if (!boxPageFrames) return compose(acquireStory()).layout;
  // Page frames of the story's page-placed content (story-page-frames.ts):
  // carried into story coordinates through the box's own story placement,
  // which depends on the story's extent (autofit, anchor offset), which can
  // depend on that content as placed. Solved for frames the story was laid
  // out with that are exactly the ones its placement implies
  // (solveStoryPageFrames); the limit only guards resources. A story without
  // such content depends on no frames.
  const framesFor = (composed: ReturnType<typeof compose>): StoryPageFrames | null => {
    const frames = storyPageFramesThrough(boxPageFrames, composed.storyToBox);
    return frames && Object.freeze({ ...frames, storyFlowPt: composed.storyFlowPt });
  };
  const seed = compose(acquireStory());
  if (!seed.layout.story.bandDependent) return seed.layout;
  const start = framesFor(seed);
  if (!start) return seed.layout;
  try {
    return solveStoryPageFrames(start, (frames) => {
      const composed = compose(frames ? acquireStory(frames) : acquireStory());
      return { value: composed.layout, next: framesFor(composed) };
    }, 16);
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        `text-box story page frames did not converge (${error.reason}; ${error.states.length} states)`,
      );
    }
    throw error;
  }
}

/** Single acquisition seam from public/parser paragraph input to retained geometry.
 * Existing `measureParagraph` remains the sole segment and line-break owner. */
class ParagraphAnchorReflowNonConvergenceError extends LayoutInvariantError {
  readonly reason: 'cycle' | 'limit';
  readonly states: readonly string[];
  readonly occurrenceCapacity: number;

  constructor(
    reason: 'cycle' | 'limit',
    states: readonly string[],
    occurrenceCapacity: number,
  ) {
    super(
      'NON_CONVERGENCE',
      `parser-owned paragraph anchor reflow did not converge (${reason}; ${occurrenceCapacity} occurrences; ${states.length} states)`,
    );
    this.name = 'ParagraphAnchorReflowNonConvergenceError';
    this.reason = reason;
    this.states = Object.freeze([...states]);
    this.occurrenceCapacity = occurrenceCapacity;
  }
}

interface AcquiredParagraphResult {
  readonly measured: MeasuredParagraph;
  readonly layout: ParagraphLayout;
}

/** Project retained wrap exclusions onto the float-wrap geometry authority. */
function wrapExclusionFloatRects(exclusions: readonly WrapExclusion[]): FloatRect[] {
  return exclusions.map((exclusion, index) => ({
    kind: 'shape' as const,
    mode: exclusion.wrap === 'topAndBottom' ? 'topAndBottom' as const : 'square' as const,
    authoredWrap: exclusion.wrap,
    wrapPolygon: exclusion.polygon,
    imageKey: exclusion.id,
    imageX: exclusion.bounds.xPt,
    imageY: exclusion.bounds.yPt,
    imageW: exclusion.bounds.widthPt,
    imageH: exclusion.bounds.heightPt,
    xLeft: exclusion.bounds.xPt,
    xRight: exclusion.bounds.xPt + exclusion.bounds.widthPt,
    yTop: exclusion.bounds.yPt,
    yBottom: exclusion.bounds.yPt + exclusion.bounds.heightPt,
    side: exclusion.wrapSide ?? 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: index,
    ...(exclusion.anchorLineExemptTopPt === undefined
      ? {} : { exemptLineTopPt: exclusion.anchorLineExemptTopPt }),
    ...(exclusion.topEdgeInclusiveFromYPt === undefined
      ? {} : { topEdgeInclusiveFromYPt: exclusion.topEdgeInclusiveFromYPt }),
  }));
}

/** WORD_MODE14_COLUMN_LINE_START_ORIGIN: the left edge of the first free gap
 * of the anchor paragraph first-line band, around other paragraphs' floats. */
function wordMode14ColumnLineStartOrigin(
  column: AnchorFrameRect,
  externalExclusions: readonly WrapExclusion[],
  paragraphOccurrenceIds: ReadonlySet<string>,
  bandTopPt: number,
  bandHeightPt: number,
): number {
  const others = externalExclusions.filter((exclusion) =>
    exclusion.anchorOccurrenceId === undefined
      || !paragraphOccurrenceIds.has(exclusion.anchorOccurrenceId));
  if (others.length === 0) return column.xPt;
  return firstFreeGapLeftPt(
    wrapExclusionFloatRects(others),
    bandTopPt,
    bandHeightPt,
    column.xPt,
    column.xPt + column.widthPt,
  ) ?? column.xPt;
}

/** WORD_MODE14_TIGHT_ANCHOR_LINE_REWRAP: host lines (laid out without their
 * own objects) whose content the unpadded tight polygon does not meet. */
function wordMode14AnchorLineExemptions(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  layout: ParagraphLayout,
  ownedExclusions: readonly WrapExclusion[],
): ReadonlyMap<string, number> {
  const mode = options.environment.compatibilityMode;
  const exemptions = new Map<string, number>();
  if (mode === undefined || mode > 14) return exemptions;
  for (const exclusion of ownedExclusions) {
    if (exclusion.wrap !== 'tight' || exclusion.anchorOccurrenceId === undefined) continue;
    const line = layout.lines.find((candidate) => candidate.placements.some((placement) =>
      placement.kind === 'anchor-host'
        && placement.anchorOccurrenceId === exclusion.anchorOccurrenceId));
    if (!line) continue;
    let left = Number.POSITIVE_INFINITY;
    let right = Number.NEGATIVE_INFINITY;
    for (const placement of line.placements) {
      if (placement.kind !== 'text' && placement.kind !== 'resource' && placement.kind !== 'tab') continue;
      const bounds = placement.bounds;
      if (!bounds || !(bounds.widthPt > 0)) continue;
      left = Math.min(left, bounds.xPt);
      right = Math.max(right, bounds.xPt + bounds.widthPt);
    }
    if (!(right > left)) {
      // An empty line holds only the paragraph mark, one mark em wide.
      left = line.bounds.xPt;
      right = left + getDefaultFontSize(paragraph);
    }
    if (!polygonMeetsRect(exclusion.polygon, line.bounds.yPt, line.bounds.heightPt, left, right)) {
      exemptions.set(exclusion.anchorOccurrenceId, line.bounds.yPt);
    }
  }
  return exemptions;
}

function wordMode14ColumnOriginApplies(
  acquisition: AnchorAcquisitionInput,
  options: ParagraphAcquisitionOptions,
): boolean {
  const mode = options.environment.compatibilityMode;
  return mode !== undefined
    && mode <= 14
    && options.ordinaryFlow
    && acquisition.simplePosition.enabled !== true
    && acquisition.horizontal.relativeFromStatus === 'valid'
    && acquisition.horizontal.relativeFrom === 'column'
    && acquisition.horizontal.choice.kind === 'offset';
}

function measurementPlacement(
  options: ParagraphAcquisitionOptions,
  exclusions: readonly WrapExclusion[],
): MeasurementPlacement {
  if (exclusions.length === 0) return options.placement;
  if (options.placement.wrap) {
    throw new Error('Conflicting paragraph wrap authorities: placement.wrap and effective exclusions');
  }
  const pageReference = options.anchorFrames?.page;
  const exclusionOracle = createFloatWrapOracle(wrapExclusionFloatRects(exclusions), {
          xLeftPt: pageReference?.xPt ?? options.placement.paragraphXPt,
          xRightPt: pageReference
            ? pageReference.xPt + pageReference.widthPt
            : options.placement.paragraphXPt + options.placement.availableWidthPt,
          readingDirection: options.context.baseRtl ? 'rtl' : 'ltr',
        });
  return {
    ...options.placement,
    wrap: exclusionOracle,
  };
}

function canonicalOwnedExclusions(
  layout: ParagraphLayout,
  occurrenceIds: ReadonlySet<string>,
): readonly WrapExclusion[] {
  const byOccurrence = new Map<string, WrapExclusion>();
  for (const exclusion of layout.exclusions) {
    const occurrenceId = exclusion.anchorOccurrenceId;
    if (!occurrenceId || !occurrenceIds.has(occurrenceId)) continue;
    if (byOccurrence.has(occurrenceId)) {
      throw new Error(`Paragraph anchor occurrence produced duplicate exclusions: ${occurrenceId}`);
    }
    byOccurrence.set(occurrenceId, exclusion);
  }
  return Object.freeze([...byOccurrence.values()]);
}

function exclusionSetState(exclusions: readonly WrapExclusion[]): string {
  return stableFingerprint('paragraph-effective-wrap-exclusions', exclusions.map((exclusion) => ({
    id: exclusion.id,
    ...(exclusion.anchorOccurrenceId === undefined
      ? {} : { occurrenceId: exclusion.anchorOccurrenceId }),
    wrap: exclusion.wrap,
    ...(exclusion.wrapSide === undefined ? {} : { wrapSide: exclusion.wrapSide }),
    bounds: exclusion.bounds,
    polygon: exclusion.polygon,
    ...(exclusion.verticalOwnership === undefined
      ? {} : { verticalOwnership: exclusion.verticalOwnership }),
    ...(exclusion.anchorLineExemptTopPt === undefined
      ? {} : { anchorLineExemptTopPt: exclusion.anchorLineExemptTopPt }),
    ...(exclusion.topEdgeInclusiveFromYPt === undefined
      ? {} : { topEdgeInclusiveFromYPt: exclusion.topEdgeInclusiveFromYPt }),
    ...(exclusion.wordMode14TightAnchor === true ? { wordMode14TightAnchor: true } : {}),
  })));
}

function externalExclusionOccurrenceIds(
  exclusions: readonly WrapExclusion[],
): ReadonlySet<string> {
  const occurrenceIds = new Set<string>();
  for (const exclusion of exclusions) {
    const occurrenceId = exclusion.anchorOccurrenceId;
    if (!occurrenceId) continue;
    if (occurrenceIds.has(occurrenceId)) {
      throw new Error(`Duplicate external paragraph exclusion occurrence: ${occurrenceId}`);
    }
    occurrenceIds.add(occurrenceId);
  }
  return occurrenceIds;
}

function mergeParagraphExclusions(
  external: readonly WrapExclusion[],
  owned: readonly WrapExclusion[],
): readonly WrapExclusion[] {
  const externallyOwned = externalExclusionOccurrenceIds(external);
  return Object.freeze([
    ...external,
    ...owned.filter((exclusion) =>
      !exclusion.anchorOccurrenceId || !externallyOwned.has(exclusion.anchorOccurrenceId)),
  ]);
}

function mergeAnchorCollisions(
  external: readonly DrawingMLCollisionEntryPt[],
  owned: readonly DrawingMLCollisionEntryPt[],
): readonly DrawingMLCollisionEntryPt[] {
  const externalOccurrences = new Set<string>();
  for (const entry of external) {
    if (externalOccurrences.has(entry.occurrenceId)) {
      throw new Error(`Duplicate external anchor collision occurrence: ${entry.occurrenceId}`);
    }
    externalOccurrences.add(entry.occurrenceId);
  }
  return Object.freeze([
    ...external,
    ...owned.filter((entry) => !externalOccurrences.has(entry.occurrenceId)),
  ]);
}

/**
 * The line key keeps the weak-map paragraph identity (runs, paragraph mark,
 * numbering) plus every input read by measureParagraph/buildSegments/layoutLines:
 * width/X, paragraph context, measurer, text/math services, vertical glyphs,
 * document compatibility settings, continuation, and field/note values only
 * for runs that consume them. Page/flow IDs, source, border/shading/trailing
 * extent, anchor frames/collisions, source-range rebasing, and page writing
 * frame are consumed by retained placement, so only the exact v1 key owns them.
 * startYPt, maximumYPt, and suppressed space-before are also exact-placement
 * inputs. Active line grids and wrap authorities are gated below before line
 * reuse; the exact key still includes all of their original facts.
 */
function paragraphAcquisitionKey(
  cache: ParagraphAcquisitionRuntimeCache,
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  continuation?: Parameters<typeof measureParagraph>[5],
  lineOnly = false,
): string {
  const layoutServices = options.environment.layoutServices;
  const verticalGlyphMeasurement = options.environment.verticalGlyphMeasurement;
  const anchorFrames = options.anchorFrames;
  const hasAnchoredPayload = !lineOnly && paragraph.runs.some(anchoredPayloadRun);
  const hasCompleteTextBox = !lineOnly && paragraph.runs.some((run) =>
    run.type === 'shape' && run.textBoxInput?.kind === 'complete');
  const {
    wrap,
    ...plainPlacement
  } = options.placement;
  const context = options.context;
  const environment = options.environment;
  const hasFieldRun = lineOnly && paragraph.runs.some((run) => run.type === 'field');
  const hasNoteReference = lineOnly && paragraph.runs.some((run) =>
    run.type === 'text' && run.noteRef !== undefined);
  // Fixed-order tuples avoid the recursive generic fingerprint cost on this hot
  // path. A different property insertion order may conservatively miss for the
  // explicitly JSON-valued geometry below, but can never alias different facts.
  return `${lineOnly ? 'paragraph-line-breaking-v1' : 'paragraph-acquisition-v1'}:${JSON.stringify([
    lineOnly ? null : options.id,
    lineOnly ? null : [options.source.story, options.source.storyInstance, options.source.path],
    lineOnly ? null : options.flowDomainId,
    lineOnly ? null : options.ordinaryFlow,
    lineOnly
      ? [plainPlacement.paragraphXPt, plainPlacement.availableWidthPt, plainPlacement.noWrap ?? false]
      : [
          plainPlacement.startYPt,
          plainPlacement.paragraphXPt,
          plainPlacement.availableWidthPt,
          plainPlacement.maximumYPt,
          plainPlacement.suppressSpaceBefore,
          plainPlacement.noWrap ?? false,
          wrap ? cache.objectIdentity(wrap) : null,
        ],
    [
      context.lineGrid.active,
      context.lineGrid.pitchPt,
      context.characterGrid.active,
      context.characterGrid.kind,
      context.characterGrid.deltaPt,
      context.rightIndentGrid.pitchPt,
      context.rightIndentGrid.paragraphAllowsAdjustment,
      context.physicalIndentLeftPt,
      context.physicalIndentRightPt,
      context.firstIndentPt,
      context.lineSpacing
        ? [
            context.lineSpacing.value,
            context.lineSpacing.rule,
            context.lineSpacing.explicit ?? null,
          ]
        : null,
      context.spaceBeforePt,
      context.spaceAfterPt,
      context.baseRtl,
      context.isJustified,
      context.stretchLastLine,
      context.tabStops.map((stop) => [stop.pos, stop.alignment, stop.leader]),
      context.hasRuby,
      context.hasEastAsianText,
      cache.kinsokuKey(context.kinsoku),
      context.defaultTabPt,
      context.overflowPunct !== false,
      context.numberingMarkerGeometry
        ? JSON.stringify(context.numberingMarkerGeometry)
        : null,
      context.mathDefJc ?? null,
      ...(lineOnly ? [context.characterGrid.pitchPt] : []),
    ],
    [
      cache.objectIdentity(options.measurer.context),
      cache.objectIdentity(options.measurer.fontFamilyClasses),
    ],
    [
      lineOnly && !hasFieldRun ? null : environment.pageIndex,
      lineOnly && !hasFieldRun ? null : environment.totalPages,
      lineOnly && !hasFieldRun ? null : environment.displayPageNumber ?? null,
      lineOnly && !hasFieldRun ? null : environment.pageNumberFormat ?? null,
      lineOnly && !hasFieldRun ? null : environment.currentDateMs ?? null,
      (lineOnly && !hasNoteReference) || !environment.noteNumbers
        ? null
        : [...environment.noteNumbers.entries()]
          .sort(([left], [right]) => left.localeCompare(right)),
      lineOnly && !hasNoteReference ? null : environment.noteReferenceNumber ?? null,
      lineOnly && !hasNoteReference ? null : environment.noteNumbering
        ? [
          environment.noteNumbering.footnote.format,
          environment.noteNumbering.footnote.start,
          environment.noteNumbering.endnote.format,
          environment.noteNumbering.endnote.start,
        ]
        : null,
      lineOnly ? null : environment.pageWritingMode,
      environment.verticalCJK ?? null,
      lineOnly ? null : environment.verticalPageFrame ?? null,
      // Mode owns acquired kerning and justified-compression decisions.
      environment.compatibilityMode ?? null,
      environment.documentHasEastAsianText,
      environment.useFeLayout ?? null,
      environment.balanceSingleByteDoubleByteWidth ?? null,
      environment.characterSpacingControl ?? null,
      environment.lineWrapLikeWord6 ?? null,
      environment.resolvedLocalFonts
        ? cache.objectIdentity(environment.resolvedLocalFonts)
        : null,
      cache.fingerprintOrdinal(layoutServices?.text.fingerprint ?? null),
      cache.fingerprintOrdinal(layoutServices?.images.fingerprint ?? null),
      cache.fingerprintOrdinal(layoutServices?.math.fingerprint ?? null),
      cache.fingerprintOrdinal(layoutServices?.verticalGlyphFingerprint ?? null),
      cache.fingerprintOrdinal(verticalGlyphMeasurement?.fingerprint ?? null),
      ...(lineOnly ? [
        environment.showTrackedChanges === true,
        environment.revisionAuthorColor
          ? cache.objectIdentity(environment.revisionAuthorColor) : null,
        environment.enableOpenTypeFeatures === true,
        environment.positionExtendsLineBox !== false,
      ] : []),
    ],
    lineOnly ? null : JSON.stringify(options.exclusions),
    lineOnly || !hasAnchoredPayload ? null : JSON.stringify(options.anchorCollisions ?? []),
    lineOnly || !hasAnchoredPayload ? null : options.paragraphAnchorReferenceDeltaPt ?? null,
    lineOnly || !hasAnchoredPayload || !options.frozenAnchorFrames?.size
      ? null
      : JSON.stringify([...options.frozenAnchorFrames].sort(([left], [right]) => left.localeCompare(right))),
    continuation ? JSON.stringify(continuation) : null,
    lineOnly ? null : options.paragraphBorderEdges
      ? [options.paragraphBorderEdges.top, options.paragraphBorderEdges.bottom]
      : null,
    lineOnly ? null : options.trailingExtentPt ?? null,
    lineOnly ? null : options.containerShading ?? null,
    options.continuesFromPrevious ?? null,
    lineOnly ? null : options.sourceRangeStart ?? null,
    lineOnly ? null : anchorFrames ? [
      anchorFrames.page
        ? [
            anchorFrames.page.xPt,
            anchorFrames.page.yPt,
            anchorFrames.page.widthPt,
            anchorFrames.page.heightPt,
          ]
        : null,
      anchorFrames.margin
        ? [
            anchorFrames.margin.xPt,
            anchorFrames.margin.yPt,
            anchorFrames.margin.widthPt,
            anchorFrames.margin.heightPt,
          ]
        : null,
      anchorFrames.column
        ? [
            anchorFrames.column.xPt,
            anchorFrames.column.yPt,
            anchorFrames.column.widthPt,
            anchorFrames.column.heightPt,
          ]
        : null,
      anchorFrames.pageParity,
    ] : null,
    lineOnly || !hasAnchoredPayload ? null : JSON.stringify(options.anchorCellBounds ?? null),
    lineOnly || !hasCompleteTextBox || !options.acquireCompleteStory
      ? null : cache.objectIdentity(options.acquireCompleteStory),
    lineOnly || !hasCompleteTextBox || !options.hostFlowPageTranslationPt
      ? null
      : [options.hostFlowPageTranslationPt.xPt, options.hostFlowPageTranslationPt.yPt],
    lineOnly || !hasCompleteTextBox || options.hostPageFrames === undefined
      ? null
      : JSON.stringify(options.hostPageFrames),
  ])}`;
}

export function paragraphAcquisitionCacheKey(
  cache: ParagraphAcquisitionRuntimeCache,
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  continuation?: Parameters<typeof measureParagraph>[5],
): string {
  return paragraphAcquisitionKey(cache, paragraph, options, continuation);
}

/**
 * Position-sensitive line breaking is limited to float/wrap exclusions and
 * self-owned anchor exclusions, which feed absolute line windows and pageH to
 * layoutLines. Line grids remain on the exact-position path conservatively,
 * since section/page-relative snapping must not inherit a partition measured
 * at another origin.
 * With none of these, measureParagraph uses startYPt and suppressed space-before
 * only to place its already-broken lines. maximumYPt/page-bottom fitting is
 * consumed by pagination after acquisition; it cannot alter this line partition.
 * The line key above retains the remaining measurement inputs; the exact
 * acquisition key retains the placement-only inputs as well.
 */
function lineBreakingDependsOnPosition(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
): boolean {
  return options.placement.wrap !== undefined
    || options.exclusions.length > 0
    || options.context.lineGrid.active
    || paragraph.runs.some(anchoredPayloadRun);
}

/** Reapply the same cursor arithmetic as measureParagraph to immutable lines. */
function placeCachedLineBreaking(
  template: MeasuredParagraph,
  placement: MeasurementPlacement,
): MeasuredParagraph {
  let cursorPt = placement.startYPt
    + (placement.suppressSpaceBefore ? 0 : template.requestedSpaceBeforePt);
  const lines = template.lines.map((line) => {
    const placed = Object.freeze({ ...line, topYPt: cursorPt });
    cursorPt += line.advancePt;
    return placed;
  });
  return Object.freeze({
    ...template,
    lines: Object.freeze(lines),
    contentStartYPt: lines[0]!.topYPt,
    contentEndYPt: cursorPt,
    placement: Object.freeze({ ...placement }),
  });
}

type MeasuredLayoutSegment = LayoutLine['segments'][number];

function immutableMeasuredLayoutSegment(
  segment: MeasuredLayoutSegment,
): MeasuredLayoutSegment {
  const source = segment.src ? Object.freeze({ ...segment.src }) : undefined;
  if ('text' in segment) {
    return Object.freeze({
      ...segment,
      ...(source ? { src: source } : {}),
      ...(segment.shapedClusters ? {
        shapedClusters: Object.freeze(segment.shapedClusters.map((cluster) => Object.freeze({
          ...cluster,
          range: Object.freeze({ ...cluster.range }),
        }))),
      } : {}),
      ...(segment.selectedFaceInkBounds ? {
        selectedFaceInkBounds: Object.freeze({ ...segment.selectedFaceInkBounds }),
      } : {}),
      ...(segment.selectedFaceFontBox ? {
        selectedFaceFontBox: Object.freeze({ ...segment.selectedFaceFontBox }),
      } : {}),
      ...(segment.ruby ? { ruby: Object.freeze({ ...segment.ruby }) } : {}),
      ...(segment.border ? { border: Object.freeze({ ...segment.border }) } : {}),
      ...(segment.revision ? { revision: Object.freeze({ ...segment.revision }) } : {}),
      ...(segment.hyperlink ? { hyperlink: Object.freeze({ ...segment.hyperlink }) } : {}),
      ...(segment.seaBreaks ? {
        seaBreaks: Object.freeze([...segment.seaBreaks]),
      } : {}),
    });
  }
  if ('imagePath' in segment) {
    return Object.freeze({
      ...segment,
      ...(source ? { src: source } : {}),
      ...(segment.srcRect ? { srcRect: Object.freeze({ ...segment.srcRect }) } : {}),
      ...(segment.duotone ? { duotone: Object.freeze({ ...segment.duotone }) } : {}),
    });
  }
  if ('isTab' in segment) {
    return Object.freeze({
      ...segment,
      ...(source ? { src: source } : {}),
      ...(segment.ptab ? { ptab: Object.freeze({ ...segment.ptab }) } : {}),
    });
  }
  return Object.freeze({
    ...segment,
    ...(source ? { src: source } : {}),
  });
}

function immutableMeasuredLine(
  line: MeasuredParagraph['lines'][number],
): MeasuredParagraph['lines'][number] {
  return Object.freeze({
    ...line,
    layout: Object.freeze({
      ...line.layout,
      // LayoutLine predates retained acquisition and exposes a mutable array
      // type. The cached snapshot is intentionally runtime-immutable.
      segments: Object.freeze(
        line.layout.segments.map(immutableMeasuredLayoutSegment),
      ) as unknown as LayoutLine['segments'],
      ...(line.layout.consumedEnd ? {
        consumedEnd: Object.freeze({ ...line.layout.consumedEnd }),
      } : {}),
    }),
  });
}

/** @internal Acquires the measurement and retained layout as one final candidate. */
export function acquireParagraphResult(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
  continuation?: Parameters<typeof measureParagraph>[5],
): AcquiredParagraphResult {
  const cache = options.environment.layoutServices
    ? paragraphAcquisitionCacheOf(options.environment.layoutServices)
    : undefined;
  const cacheKey = cache
    ? paragraphAcquisitionCacheKey(cache, paragraph, options, continuation)
    : undefined;
  const cached = cacheKey === undefined
    ? undefined
    : cache!.get(paragraph, cacheKey) as AcquiredParagraphResult | undefined;
  if (cached) return cached;
  cache?.noteMiss();
  // Defer the second serialization until the exact placement misses. The
  // first eligible miss must still retain a template: pagination may next
  // acquire that paragraph at a different Y during convergence.
  const reusableLineKey = !lineBreakingDependsOnPosition(paragraph, options)
    && cache
    ? paragraphAcquisitionKey(cache, paragraph, options, continuation, true)
    : undefined;
  const lineTemplate = reusableLineKey === undefined
    ? undefined
    : cache!.getLineBreaking(paragraph, reusableLineKey) as MeasuredParagraph | undefined;
  const externallyOwnedOccurrenceIds = externalExclusionOccurrenceIds(options.exclusions);
  const occurrenceIds = new Set(paragraph.runs.flatMap((run) =>
    anchoredPayloadRun(run) ? [run.anchorAcquisitionInput!.occurrenceId] : []));
  for (const occurrenceId of externallyOwnedOccurrenceIds) occurrenceIds.delete(occurrenceId);
  const occurrenceCapacity = occurrenceIds.size;
  const initialOwnedExclusions: readonly WrapExclusion[] = Object.freeze([]);
  const initialExclusions = mergeParagraphExclusions(
    options.exclusions,
    initialOwnedExclusions,
  );
  const numberingPlan = continuation || options.continuesFromPrevious
    ? undefined : retainedNumberingPlan(paragraph, options.context, options);
  const acquisitionOptions = numberingPlan && !options.context.numberingMarkerGeometry
    ? { ...options, context: { ...options.context, numberingMarkerGeometry: numberingPlan } }
    : options;
  type Pass = Readonly<{
    measured: MeasuredParagraph;
    layout: ParagraphLayout;
    ownedExclusions: readonly WrapExclusion[];
    state: string;
  }>;
  // WORD_MODE14_TIGHT_ANCHOR_LINE_REWRAP: decided once from the pass that laid
  // every host line out without its own objects.
  let anchorLineExemptions: ReadonlyMap<string, number> | undefined;
  try {
    const result = convergeExactState<Pass>({
      seedState: exclusionSetState(initialExclusions),
      step: (previous) => {
        const effectiveExclusions = mergeParagraphExclusions(
          options.exclusions,
          previous?.ownedExclusions ?? initialOwnedExclusions,
        );
        const measured = lineTemplate
          ? placeCachedLineBreaking(lineTemplate, options.placement)
          : measureParagraph(
          paragraph,
          acquisitionOptions.context,
          measurementPlacement(options, effectiveExclusions),
          options.measurer,
          {
            ...options.environment,
            paragraphMarkShapeInput: paragraph.paragraphMarkShapeInput,
            ...(numberingPlan?.lineBox && numberingPlan.markerText ? {
              firstLineNumberingMarkerBox: numberingPlan.lineBox,
            } : {}),
          },
          continuation,
          );
        const layout = paragraphLayoutFromMeasurement(
          paragraph,
          anchorLineExemptions && anchorLineExemptions.size > 0
            ? { ...acquisitionOptions, anchorLineExemptions }
            : acquisitionOptions,
          measured,
        );
        const ownedExclusions = canonicalOwnedExclusions(layout, occurrenceIds);
        if (anchorLineExemptions === undefined) {
          anchorLineExemptions = wordMode14AnchorLineExemptions(
            paragraph,
            options,
            layout,
            ownedExclusions,
          );
        }
        const nextEffectiveExclusions = mergeParagraphExclusions(
          options.exclusions,
          ownedExclusions,
        );
        const state = exclusionSetState(nextEffectiveExclusions);
        if (exclusionSetState(layout.exclusions) !== state) {
          throw new Error('Paragraph retained exclusions differ from the measured exclusion authority');
        }
        return Object.freeze({ measured, layout, ownedExclusions, state });
      },
      stateOf: (pass) => pass.state,
      // Operational fail-closed resource guard. The exact-state/cycle checks
      // establish correctness; this budget prevents an all-distinct malicious
      // geometry orbit from consuming unbounded work.
      limit: 16,
    }).value;
    // A cache hit may cross convergence passes. Retain an immutable measurement
    // envelope without recursively freezing caller-owned capabilities such as
    // the wrap oracle referenced by placement.
    const immutableMeasured: MeasuredParagraph = lineTemplate
      ? result.measured
      : Object.freeze({
          ...result.measured,
          lines: Object.freeze(result.measured.lines.map(immutableMeasuredLine)),
          placement: Object.freeze({ ...result.measured.placement }),
        });
    const acquired = Object.freeze({ measured: immutableMeasured, layout: result.layout });
    if (cacheKey !== undefined) cache!.set(paragraph, cacheKey, acquired);
    if (reusableLineKey !== undefined && !lineTemplate && !immutableMeasured.markOnly) {
      cache!.setLineBreaking(paragraph, reusableLineKey, immutableMeasured);
    }
    return acquired;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new ParagraphAnchorReflowNonConvergenceError(
        error.reason,
        error.states,
        occurrenceCapacity,
      );
    }
    throw error;
  }
}

export function acquireParagraphLayout(
  paragraph: ParagraphAcquisitionInput,
  options: ParagraphAcquisitionOptions,
): ParagraphLayout {
  return acquireParagraphResult(paragraph, options).layout;
}

export interface RetainedFrameGroupAcquisition {
  readonly box: Readonly<{
    bounds: LayoutRect;
    exclusionBounds: LayoutRect;
    exclusionId: string;
  }>;
  readonly members: readonly Readonly<{
    paragraph: ParagraphLayoutSource;
    fragment: ParagraphLayout;
    source: SourceRef;
  }>[];
}

/** Maximum downward baseline shift carried by retained frame glyph paint.
 * Consumers receive immutable retained geometry and do not reconstruct run
 * positioning from parser fields. */
export function retainedFrameMaximumBaselineLoweringPt(
  acquisition: RetainedFrameGroupAcquisition,
): number {
  let maximumPt = 0;
  for (const member of acquisition.members) {
    for (const line of member.fragment.lines) {
      for (const placement of line.placements) {
        if (placement.kind !== 'text') continue;
        maximumPt = Math.max(maximumPt, -(placement.positionPt ?? 0));
      }
    }
  }
  return maximumPt;
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
  /** Owning story of the grouped paragraphs; `sourceIndices` are paths in its
   * root block list. Defaults to the main body. */
  readonly story?: Readonly<{ story: SourceRef['story']; storyInstance: string }>;
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
  const owner = options.story ?? { story: 'body' as const, storyInstance: 'body' };
  const framePrefix = owner.story === 'body'
    ? 'body-frame'
    : `${owner.story}:${owner.storyInstance}:frame`;
  const cacheKey = stableFingerprint('w:frame-acquisition', [
    owner.story,
    owner.storyInstance,
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
    options.environment.layoutServices?.verticalGlyphFingerprint ?? null,
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
    let wrapRegistry = createParagraphWrapRegistry(`${framePrefix}:${group.id}`);
    let cursorPt = 0;
    let previous: ParagraphLayoutSource | null = null;
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
      const borderExtentPt = options.borderExtentsPt[memberIndex] ?? 0;
      const source: SourceRef = {
        story: owner.story,
        storyInstance: owner.storyInstance,
        path: [group.sourceIndices[memberIndex]!],
      };
      const acquired = acquireParagraphResult(
        options.inputs[memberIndex]!,
        {
          id: `${framePrefix}:${group.id}:${memberIndex}`,
          source,
          flowDomainId: `${framePrefix}:${group.id}`,
          ordinaryFlow: false,
          context,
          placement,
          measurer: options.measurer,
          // `word-lowered-drop-cap-anchor-leading`: only a drop cap keeps its
          // authored line count/height fixed while separately retained glyph
          // lowering expands the following anchor exclusion. Ordinary frames
          // keep positioned ink in their line boxes.
          environment: {
            ...options.environment,
            positionExtendsLineBox: wordFramePositionExtendsLineBox(group.framePr.dropCap),
          },
          exclusions: wrapRegistry.exclusions,
          anchorCollisions: wrapRegistry.collisions,
          containerShading: options.containerShading,
          paragraphBorderEdges: options.borderEdges[memberIndex],
          trailingExtentPt: Math.max(context.spaceAfterPt, borderExtentPt),
          anchorFrames: options.anchorFrames,
        },
      );
      const { measured, layout: fragment } = acquired;
      wrapRegistry = commitParagraphWrapRegistry(wrapRegistry, fragment);
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
    const laidOut = layoutParagraph(fp.hRule === 'exact' && fp.h != null
      ? { ...translated, clipBounds: placed.bounds }
      : translated);
    // w:framePr box height remains retained geometry, but a positioned frame
    // contributes no block advance to the ordinary paragraph flow that anchors it.
    const fragment = Object.freeze({ ...laidOut, advancePt: 0 });
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
  const rightGridAdjustmentPt = paragraphGridRightAdjustmentPt(
    planningContext,
    options.placement.availableWidthPt,
  );
  const availableWidthPt = options.placement.availableWidthPt
    - planningContext.physicalIndentLeftPt
    - planningContext.physicalIndentRightPt
    - rightGridAdjustmentPt;
  const occurrences = logicalOccurrenceMap(paragraph, measured, options.environment.showTrackedChanges);
  const numberingPlan = options.continuesFromPrevious
    ? undefined
    : retainedNumberingPlan(paragraph, planningContext, options);
  let lines = planMeasuredLines(
    measured, paragraph, paragraphXPt, availableWidthPt, options.source, options.id, planningContext,
    occurrences, numberingPlan, options.environment.layoutServices?.text,
    options.environment.verticalGlyphMeasurement,
    options.environment.verticalPageFrame,
    options.environment.compatibilityMode,
    options.environment.paragraphMarkShapeInput,
    uprightResourceOrientation(
      options.environment.verticalPageFrame,
      options.environment.pageWritingMode,
    ),
  );
  if (options.sourceRangeStart !== undefined) {
    lines = rebaseMeasuredLineRanges(lines, options.sourceRangeStart);
  }
  lines = attachBarTabRules(
    lines,
    options.placement.paragraphXPt,
    options.placement.availableWidthPt,
    planningContext.baseRtl,
    planningContext.tabStops,
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
  const anchorCollisions: DrawingMLCollisionEntryPt[] = [];
  const cellContainmentRects: LayoutRect[] = [];
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
  const paragraphOccurrenceIds: ReadonlySet<string> = new Set(payloadsByOccurrence.keys());
  for (const [occurrenceId, payloads] of payloadsByOccurrence) {
    const acquired = acquireAnchorOccurrence(
      occurrenceId,
      payloads,
      lines,
      paragraph,
      options,
      measured.contentEndYPt - options.placement.startYPt,
      options.exclusions,
      anchorExclusions,
      options.anchorCollisions ?? [],
      anchorCollisions,
      paragraphOccurrenceIds,
      options.placement.startYPt
        + (options.placement.suppressSpaceBefore ? 0 : measured.requestedSpaceBeforePt),
    );
    if (!acquired) continue;
    anchorResults.push(acquired.result);
    if (acquired.cellContainmentBounds) {
      cellContainmentRects.push(acquired.cellContainmentBounds);
    }
    if (!acquired.drawing) continue;
    drawings.push(acquired.drawing);
    textBoxes.push(...acquired.textBoxes);
    if (acquired.exclusion) anchorExclusions.push(acquired.exclusion);
    if (acquired.collision) anchorCollisions.push(acquired.collision);
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
    if (run.type === 'unavailableDrawing' && run.anchorAcquisitionInput === undefined) {
      const drawingId = unavailableDrawingId(options.source, runIndex);
      const placement = lines
        .flatMap((line) => line.placements)
        .find((candidate) => candidate.kind === 'drawing' && candidate.drawingId === drawingId);
      if (placement?.kind === 'drawing') {
        drawings.push({
          kind: 'drawing',
          id: drawingId,
          source,
          flowDomainId: options.flowDomainId,
          flowBounds: placement.bounds,
          inkBounds: placement.bounds,
          advancePt: 0,
          ordinaryFlow: false,
          commands: Object.freeze([{ kind: 'noop' }]),
          diagnostics: Object.freeze([
            unavailableDrawingDiagnostic(run.resourceKind, source),
          ]),
        });
      }
    }
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
      const drawing = publicAnchoredResourceDrawing(run, options, runIndex);
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
      const drawingId = `${options.id}:drawing:${runIndex}`;
      const inlinePlacement = run.inline === true
        ? lines.flatMap((line) => line.placements).find((placement) =>
            placement.kind === 'drawing' && placement.drawingId === drawingId)
        : undefined;
      if (run.inline === true && !inlinePlacement) {
        throw new Error(`Inline shape ${drawingId} has no retained line placement`);
      }
      const authoredShapeRect = inlinePlacement?.bounds ?? resolvedShapeLayoutRect(run, options);
      const textBoxId = `${options.id}:textbox:${runIndex}`;
      // The box rides this paragraph's flow, so its story's page frames
      // (story-page-frames.ts) are the anchor frames' page through the
      // translation that flow still receives, when known.
      const hostFlow = options.hostFlowPageTranslationPt;
      const hostBase = options.hostPageFrames !== undefined
        ? options.hostPageFrames
        : options.anchorFrames?.page && options.anchorFrames.margin
          ? { page: options.anchorFrames.page, margin: options.anchorFrames.margin }
          : null;
      const pageFrames = hostFlow && hostBase
        ? storyPageFramesThrough(hostBase, translationAffine(hostFlow.xPt, hostFlow.yPt))
        : null;
      const textBox = acquireShapeTextBoxLayout(run, authoredShapeRect, {
        id: textBoxId,
        source,
        flowDomainId: options.flowDomainId,
        context: options.context,
        measurer: options.measurer,
        environment: options.environment,
        input: run.textBoxInput,
        acquireCompleteStory: options.acquireCompleteStory,
        ...(pageFrames ? { pageFrames } : {}),
      });
      // The wp:inline extent is the line-flow contract. A WPS text body may
      // acquire richer internal geometry, but it must not move or resize the
      // outer inline object after the line breaker has committed its advance.
      const shapeRect = run.inline === true ? authoredShapeRect : textBox?.flowBounds ?? authoredShapeRect;
      let drawing = drawingForShape(run, shapeRect, options, runIndex, run.inline === true);
      if (textBox) {
        textBoxes.push(textBox);
        drawing = { ...drawing, textBoxIds: [textBoxId] };
      }
      drawings.push(drawing);
      const firstLine = run.inline === true ? undefined : lines[0];
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
      alignedLeadingEdgePt: numberingAlignedLeadingEdgePt(
        numberingPlan,
        options.context,
        paragraphXPt,
        availableWidthPt,
        lines[0],
      ),
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
  const cellContainmentBounds = unionLayoutRects(cellContainmentRects);
  return layoutParagraph({
    kind: 'paragraph', id: options.id, source: options.source,
    ...(paragraph.paragraphId !== undefined ? { paragraphId: paragraph.paragraphId } : {}),
    flowDomainId: options.flowDomainId, ordinaryFlow: options.ordinaryFlow,
    ...(paragraph.styleId !== undefined ? { styleId: paragraph.styleId } : {}),
    ...(!options.continuesFromPrevious && paragraph.bookmarks?.length
      ? { bookmarkStarts: paragraph.bookmarks }
      : {}),
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
    exclusions: mergeParagraphExclusions(options.exclusions, anchorExclusions),
    ...(cellContainmentBounds ? { cellContainmentBounds } : {}),
    anchorCollisions: mergeAnchorCollisions(
      options.anchorCollisions ?? [],
      anchorCollisions,
    ),
    ...(anchorResults.length ? { anchorFrames: anchorResults } : {}),
    paragraphMark: measured.markOnly ? {
      ...(measured.markWrapBounds ? { wrapBounds: measured.markWrapBounds } : {}),
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
const translatePlacementY = (
  placement: import('./types.js').ParagraphPlacement,
  yPt: number,
): import('./types.js').ParagraphPlacement => translatePlacement(placement, { xPt: 0, yPt });
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
  // A continuation cut from lines measured earlier in the same flow (a table
  // cell slice starting at lineStart > 0) is placed in a new flow slice.
  // Preserve the acquired x/range geometry, but make its first retained line
  // own the same local y origin as the original paragraph so placement
  // translates one coherent coordinate space instead of carrying the
  // preceding page's consumed line offset. A body continuation is remeasured
  // at its new location (lineStart 0); its first line already sits where that
  // location's wrap places it (§20.4.2.20 topAndBottom skips the band), and
  // rebasing it would paint it inside a page-owned float's exclusion.
  const deltaYPt = continuation.continuesFromPrevious && continuation.lineStart > 0 && first
    ? acquired.flowBounds.yPt - first.bounds.yPt
    : 0;
  const rebasedSelected = deltaYPt === 0
    ? selected
    : selected.map((line) => translateLineY(line, deltaYPt));
  const rebasedFirst = rebasedSelected[0];
  const rebasedLast = rebasedSelected.at(-1);
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
  const cellContainmentBounds = unionLayoutRects(
    drawings
      .filter((drawing) => drawing.anchorLayer?.cellContainment === true)
      .map((drawing) => drawing.flowBounds),
  );
  const acquiredHostAnchorOccurrenceIds = new Set(acquired.drawings.flatMap((drawing) => {
    if (drawing.anchorLayer?.verticalOwnership !== 'host') return [];
    const occurrenceId = drawing.anchorLayer.acquisitionOccurrenceId
      ?? drawing.anchorLayer.occurrenceId;
    return occurrenceId === undefined ? [] : [occurrenceId];
  }));
  const retainedHostAnchorOccurrenceIds = new Set(drawings.flatMap((drawing) => {
    if (drawing.anchorLayer?.verticalOwnership !== 'host') return [];
    const occurrenceId = drawing.anchorLayer.acquisitionOccurrenceId
      ?? drawing.anchorLayer.occurrenceId;
    return occurrenceId === undefined ? [] : [occurrenceId];
  }));
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
  const stationaryTextBoxIds = new Set(drawings
    .filter((drawing) => drawing.anchorLayer?.verticalOwnership === 'page'
      || drawing.orientation === 'upright-physical')
    .flatMap((drawing) => drawing.textBoxIds ?? []));
  const drawingSourceKeys = new Set(drawings.map((drawing) =>
    stableFingerprint('source-occurrence', drawing.source)));
  const lineRangeStart = first?.range.start;
  const lineRangeEnd = last?.range.end;
  const {
    bookmarkStarts: acquiredBookmarkStarts,
    ...acquiredWithoutBookmarkStarts
  } = acquired;
  return finalizeParagraphLayout({
    ...acquiredWithoutBookmarkStarts,
    kind: 'paragraph', id,
    ...(!continuation.continuesFromPrevious && acquiredBookmarkStarts?.length
      ? { bookmarkStarts: acquiredBookmarkStarts }
      : {}),
    lines: acquired.lines,
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
    cellContainmentBounds: cellContainmentBounds ?? undefined,
    textBoxes: acquired.textBoxes
      .filter((textBox) =>
        textBoxIds.has(textBox.id)
        || drawingSourceKeys.has(stableFingerprint('source-occurrence', textBox.source)))
      .map((textBox) => stationaryTextBoxIds.has(textBox.id)
        ? textBox : translateTextBoxY(textBox, deltaYPt)),
    events: lineRangeStart === undefined || lineRangeEnd === undefined
      ? []
      : acquired.events.filter((event) => event.offset >= lineRangeStart
        && (event.offset < lineRangeEnd
          || (!continuation.continuesOnNext && event.offset === lineRangeEnd))),
    exclusions: acquired.exclusions
      .filter((exclusion) => exclusion.verticalOwnership === 'page'
        || exclusion.anchorOccurrenceId === undefined
        || !acquiredHostAnchorOccurrenceIds.has(exclusion.anchorOccurrenceId)
        || retainedHostAnchorOccurrenceIds.has(exclusion.anchorOccurrenceId))
      .map((exclusion) => ({
        ...exclusion,
        bounds: exclusion.verticalOwnership === 'page'
          ? exclusion.bounds : translateRectY(exclusion.bounds, deltaYPt),
        polygon: exclusion.verticalOwnership === 'page'
          ? exclusion.polygon
          : exclusion.polygon.map((point) => translatePointY(point, deltaYPt)),
      })),
    anchorCollisions: (acquired.anchorCollisions ?? [])
      .filter((entry) => entry.verticalOwnership === 'page'
        || !acquiredHostAnchorOccurrenceIds.has(entry.occurrenceId)
        || retainedHostAnchorOccurrenceIds.has(entry.occurrenceId))
      .map((entry) => ({
        ...entry,
        bounds: entry.verticalOwnership === 'page'
          ? entry.bounds : translateRectY(entry.bounds, deltaYPt),
      })),
    ...(continuation.continuesOnNext
      ? { paragraphMark: undefined }
      : acquired.paragraphMark
        ? { paragraphMark: {
            ...acquired.paragraphMark,
            bounds: translateRectY(acquired.paragraphMark.bounds, deltaYPt),
            ...(acquired.paragraphMark.wrapBounds ? { wrapBounds: translateRectY(acquired.paragraphMark.wrapBounds, deltaYPt) } : {}),
          } }
        : {}),
    continuation,
  }, acquired, rebasedSelected);
}
