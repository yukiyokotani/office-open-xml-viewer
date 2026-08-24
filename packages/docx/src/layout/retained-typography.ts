import type {
  BorderSegment,
  LayoutRect,
  PointPt,
  RetainedGlyphPaintOperation,
  RetainedRunBorderFacts,
  TextClusterLayout,
  TextDecorationLayout,
} from './types.js';
import { retainedBorderTreatment } from './border-treatment.js';
import type { GlyphInkBounds, GlyphMeasurement } from './text.js';

export interface ShapedLeaderInput extends Omit<RetainedGlyphPaintOperation, 'text' | 'origin'> {
  readonly interval: LayoutRect;
  readonly baselinePt: number;
  readonly glyph: string;
  readonly advancePt: number;
}

/** ECMA-376 §17.3.1.37 requires the selected leader character to repeat through
 * the tab interval. The selected-face advance is acquired before this function;
 * the unoccupied residual is split evenly at both ends. */
export function centeredLeaderGlyphOrigins(
  input: ShapedLeaderInput,
): readonly RetainedGlyphPaintOperation[] {
  if (!Number.isFinite(input.advancePt) || input.advancePt <= 0) {
    throw new RangeError('Tab leader glyph advance must be finite and positive');
  }
  const count = Math.floor(input.interval.widthPt / input.advancePt);
  const residualPt = input.interval.widthPt - count * input.advancePt;
  return Array.from({ length: count }, (_, index) => ({
    text: input.glyph,
    origin: {
      xPt: input.interval.xPt + residualPt / 2 + index * input.advancePt,
      yPt: input.baselinePt,
    },
    fontRoute: input.fontRoute,
    fontSizePt: input.fontSizePt,
    fontWeight: input.fontWeight,
    fontStyle: input.fontStyle,
    color: input.color,
  }));
}

export interface RubyGuideSpanInput extends Omit<RetainedGlyphPaintOperation, 'origin'> {
  readonly offsetPt: number;
}

export function rubyPaintOperations(input: Readonly<{
  baseOrigin: PointPt;
  baseAdvancePt: number;
  guideAdvancePt: number;
  raisePt?: number;
  /** Used only for the specification-independent touching-ink fallback. */
  baseInkTopPt?: number;
  guideInkBottomFromBaselinePt?: number;
  spans: readonly RubyGuideSpanInput[];
}>): readonly RetainedGlyphPaintOperation[] {
  let baselinePt: number;
  if (input.raisePt !== undefined) {
    baselinePt = input.baseOrigin.yPt - input.raisePt;
  } else if (
    input.baseInkTopPt !== undefined
    && input.guideInkBottomFromBaselinePt !== undefined
  ) {
    baselinePt = input.baseInkTopPt - input.guideInkBottomFromBaselinePt;
  } else {
    throw new Error(
      'Ruby geometry requires authored w:hpsRaise or retained base/guide ink bounds',
    );
  }
  const startXPt = input.baseOrigin.xPt
    + (input.baseAdvancePt - input.guideAdvancePt) / 2;
  return input.spans.map((span) => ({
    text: span.text,
    origin: { xPt: startXPt + span.offsetPt, yPt: baselinePt },
    fontRoute: span.fontRoute,
    fontSizePt: span.fontSizePt,
    fontWeight: span.fontWeight,
    fontStyle: span.fontStyle,
    color: span.color,
  }));
}

type RetainedInkMetric = Pick<GlyphMeasurement, 'ascentPt' | 'descentPt' | 'inkBounds'>;

function inkTop(metric: RetainedInkMetric): number {
  return -(metric.inkBounds?.ascentPt ?? metric.ascentPt);
}

function inkBottom(metric: RetainedInkMetric): number {
  return metric.inkBounds?.descentPt ?? metric.descentPt;
}

function inkStrokeWidth(metric: RetainedInkMetric): number {
  const inkHeight = metric.inkBounds
    ? metric.inkBounds.ascentPt + metric.inkBounds.descentPt
    : Math.min(metric.ascentPt, metric.descentPt);
  if (!Number.isFinite(inkHeight) || inkHeight <= 0) {
    throw new Error('Retained decoration probe requires positive selected-face ink');
  }
  return inkHeight;
}

function underlineStyle(authored: string | undefined): TextDecorationLayout['style'] {
  if (authored === 'double' || authored === 'dbl') return 'double';
  if (authored?.includes('dot')) return 'dotted';
  if (authored?.includes('dash')) return 'dashed';
  if (authored?.includes('wave')) return 'wavy';
  return 'solid';
}

export function retainedWavePath(
  from: PointPt,
  to: PointPt,
  strokeWidthPt: number,
): readonly PointPt[] {
  const widthPt = Math.max(0, to.xPt - from.xPt);
  const stepPt = strokeWidthPt * 2;
  const count = Math.max(1, Math.ceil(widthPt / stepPt));
  return Array.from({ length: count + 1 }, (_, index) => ({
    xPt: from.xPt + widthPt * index / count,
    yPt: from.yPt + (index % 2 === 0 ? -strokeWidthPt / 2 : strokeWidthPt / 2),
  }));
}

/** Retains underline and strike geometry from selected-face glyph probes. `_`
 * supplies the face's low-line position/thickness; `-` and `=` supply the
 * strike bands. The actual run ink is the collision boundary when a face's
 * native low-line would otherwise intersect a descending glyph. */
export function retainedTextDecorations(input: Readonly<{
  origin: PointPt;
  advancePt: number;
  base: RetainedInkMetric;
  color: string;
  underline?: Readonly<{ authoredStyle?: string; color: string; probe: RetainedInkMetric }>;
  strike?: Readonly<{ double: boolean; probe: RetainedInkMetric; doubleProbe?: RetainedInkMetric }>;
}>): readonly TextDecorationLayout[] {
  const decorations: TextDecorationLayout[] = [];
  const rightPt = input.origin.xPt + input.advancePt;
  if (input.underline) {
    const strokeWidthPt = inkStrokeWidth(input.underline.probe);
    const nativeCenterPt = input.origin.yPt
      + (inkTop(input.underline.probe) + inkBottom(input.underline.probe)) / 2;
    const clearCenterPt = input.origin.yPt + inkBottom(input.base) + strokeWidthPt / 2;
    const centerPt = Math.max(nativeCenterPt, clearCenterPt);
    const style = underlineStyle(input.underline.authoredStyle);
    const common = {
      kind: 'underline' as const,
      ...(input.underline.authoredStyle === undefined
        ? {} : { authoredStyle: input.underline.authoredStyle }),
      color: input.underline.color,
      widthPt: strokeWidthPt,
    };
    if (style === 'double') {
      const secondCenterPt = centerPt + strokeWidthPt * 2;
      decorations.push(
        { ...common, style: 'solid', from: { xPt: input.origin.xPt, yPt: centerPt }, to: { xPt: rightPt, yPt: centerPt } },
        { ...common, style: 'solid', from: { xPt: input.origin.xPt, yPt: secondCenterPt }, to: { xPt: rightPt, yPt: secondCenterPt } },
      );
    } else {
      const from = { xPt: input.origin.xPt, yPt: centerPt };
      const to = { xPt: rightPt, yPt: centerPt };
      decorations.push({
        ...common,
        style,
        from,
        to,
        ...(style === 'wavy' ? { path: retainedWavePath(from, to, strokeWidthPt) } : {}),
        ...(style === 'dotted' ? { dashPatternPt: [strokeWidthPt, strokeWidthPt * 2] } : {}),
        ...(style === 'dashed' ? { dashPatternPt: [strokeWidthPt * 4, strokeWidthPt * 3] } : {}),
      });
    }
  }
  if (input.strike) {
    const strokeWidthPt = inkStrokeWidth(input.strike.probe);
    if (input.strike.double && input.strike.doubleProbe) {
      const topPt = input.origin.yPt + inkTop(input.strike.doubleProbe) + strokeWidthPt / 2;
      const bottomPt = input.origin.yPt + inkBottom(input.strike.doubleProbe) - strokeWidthPt / 2;
      for (const yPt of [topPt, bottomPt]) decorations.push({
        kind: 'strikethrough', color: input.color, widthPt: strokeWidthPt, style: 'solid',
        from: { xPt: input.origin.xPt, yPt }, to: { xPt: rightPt, yPt },
      });
    } else {
      const yPt = input.origin.yPt
        + (inkTop(input.strike.probe) + inkBottom(input.strike.probe)) / 2;
      decorations.push({
        kind: 'strikethrough', color: input.color, widthPt: strokeWidthPt, style: 'solid',
        from: { xPt: input.origin.xPt, yPt }, to: { xPt: rightPt, yPt },
      });
    }
  }
  return decorations;
}

export interface RetainedEmphasisClusterInk {
  readonly text: string;
  readonly range: TextClusterLayout['range'];
  readonly ink: GlyphInkBounds;
}

export interface RetainedEmphasisMarkInput extends Omit<
  RetainedGlyphPaintOperation,
  'text' | 'origin'
> {
  readonly inkBounds: GlyphInkBounds;
}

/** ECMA-376 §17.18.24 leaves the emphasis glyph implementation-dependent.
 * Acquisition therefore retains the actual selected authored glyph once per
 * non-space source cluster. Ink is used only to place the selected glyph; its
 * outline (including ○'s hollow counter and ﹅'s comma shape) stays font-owned. */
export function retainedEmphasisGlyphs(input: Readonly<{
  authored: string;
  glyph: string;
  origin: PointPt;
  clusters: readonly TextClusterLayout[];
  clusterInk: readonly RetainedEmphasisClusterInk[];
  mark: RetainedEmphasisMarkInput;
  scaleX: number;
}>): readonly RetainedGlyphPaintOperation[] {
  const markWidthPt = input.mark.inkBounds.xMaxPt - input.mark.inkBounds.xMinPt;
  const markHeightPt = input.mark.inkBounds.ascentPt + input.mark.inkBounds.descentPt;
  if (!(markWidthPt > 0) || !(markHeightPt > 0)) {
    throw new Error('Retained emphasis glyph requires positive selected-face ink bounds');
  }
  const operations: RetainedGlyphPaintOperation[] = [];
  for (const ink of input.clusterInk) {
    if (/^\s+$/u.test(ink.text)) continue;
    const cluster = input.clusters.find((candidate) =>
      candidate.range.start === ink.range.start && candidate.range.end === ink.range.end);
    if (!cluster) throw new Error('Retained emphasis cluster ink does not match shaped cluster geometry');
    const baseLeftPt = input.origin.xPt + cluster.offset.xPt + ink.ink.xMinPt * input.scaleX;
    const baseRightPt = input.origin.xPt + cluster.offset.xPt + ink.ink.xMaxPt * input.scaleX;
    const markOriginXPt = (baseLeftPt + baseRightPt) / 2
      - (input.mark.inkBounds.xMinPt + input.mark.inkBounds.xMaxPt) / 2;
    const above = input.authored !== 'underDot';
    const markBaselinePt = above
      ? input.origin.yPt - ink.ink.ascentPt - input.mark.inkBounds.descentPt
      : input.origin.yPt + ink.ink.descentPt + input.mark.inkBounds.ascentPt;
    operations.push({
      text: input.glyph,
      origin: { xPt: markOriginXPt, yPt: markBaselinePt },
      fontRoute: input.mark.fontRoute,
      fontSizePt: input.mark.fontSizePt,
      fontWeight: input.mark.fontWeight,
      fontStyle: input.mark.fontStyle,
      color: input.mark.color,
      inkBounds: input.mark.inkBounds,
    });
  }
  return operations;
}

export interface RunBorderFragmentInput {
  readonly bounds: LayoutRect;
  readonly trailingSlackPt: number;
  readonly border: RetainedRunBorderFacts;
}

function sameRunBorder(a: RetainedRunBorderFacts, b: RetainedRunBorderFacts): boolean {
  return a.val === b.val
    && a.color === b.color
    && a.widthPt === b.widthPt
    && a.spacePt === b.spacePt
    && a.themeColor === b.themeColor
    && a.themeTint === b.themeTint
    && a.themeShade === b.themeShade
    && a.shadow === b.shadow
    && a.frame === b.frame;
}

/** Builds one four-edge rectangle for each visually adjacent equal border group.
 * Callers pass the exact justification slack owned by each visual fragment. */
export function groupedRunBorderFragments(
  inputs: readonly RunBorderFragmentInput[],
): readonly BorderSegment[] {
  const output: BorderSegment[] = [];
  let index = 0;
  while (index < inputs.length) {
    const first = inputs[index]!;
    let end = index + 1;
    let rightPt = first.bounds.xPt + first.bounds.widthPt + first.trailingSlackPt;
    while (end < inputs.length) {
      const next = inputs[end]!;
      if (!sameRunBorder(first.border, next.border)
        || Math.abs(next.bounds.xPt - rightPt) > 1e-6
        || next.bounds.yPt !== first.bounds.yPt
        || next.bounds.heightPt !== first.bounds.heightPt) break;
      rightPt = next.bounds.xPt + next.bounds.widthPt + next.trailingSlackPt;
      end += 1;
    }
    const left = first.bounds.xPt - first.border.spacePt;
    const top = first.bounds.yPt - first.border.spacePt;
    const right = rightPt + first.border.spacePt;
    const bottom = first.bounds.yPt + first.bounds.heightPt + first.border.spacePt;
    const common = {
      color: first.border.color,
      widthPt: first.border.widthPt,
      ...retainedBorderTreatment(first.border.val, first.border.widthPt),
    } as const;
    output.push(
      { ...common, edge: 'top', from: { xPt: left, yPt: top }, to: { xPt: right, yPt: top } },
      { ...common, edge: 'right', from: { xPt: right, yPt: top }, to: { xPt: right, yPt: bottom } },
      { ...common, edge: 'bottom', from: { xPt: left, yPt: bottom }, to: { xPt: right, yPt: bottom } },
      { ...common, edge: 'left', from: { xPt: left, yPt: top }, to: { xPt: left, yPt: bottom } },
    );
    index = end;
  }
  return output;
}
