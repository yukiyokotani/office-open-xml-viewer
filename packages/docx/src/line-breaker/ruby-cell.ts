import { canvasFontString, graphemeClusterOffsets } from '@silurus/ooxml-core';
import type { VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { independentTextShapeRequest } from '../layout/text.js';
import {
  WORD_RUBY_ALIGN_PLACEMENT,
  wordArabicTextKerned,
  wordRubyBaseOffset,
  wordRubyGuideSpacing,
} from '../layout/line-compatibility.js';
import type { LayoutTextSeg, RubyCellGeometry, RubyGuideUnit } from './model.js';

/** Scripts whose letters join cursively: a guide run of them is one unit,
 * shaped and painted whole, or its joining forms would be lost. */
const CURSIVE_JOINING_RE = /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Mongolian}\p{Script=Nko}\p{Script=Mandaic}\p{Script=Adlam}\p{Script=Hanifi_Rohingya}\p{Script=Sogdian}\p{Script=Manichaean}\p{Script=Psalter_Pahlavi}\p{Script=Phags_Pa}]/u;

function joinsCursively(text: string): boolean {
  return CURSIVE_JOINING_RE.test(text);
}

/** Inputs the line breaker resolves for the one geometry of a ruby cell. */
export interface RubyCellInputs {
  readonly scale: number;
  /** Laid-out advance of one base piece (spacing, grid, fitText and
   * snapToChars allocation included; ruby widening excluded). */
  readonly baseAdvancePx: (piece: LayoutTextSeg) => number;
  /** Natural advance of `text` at the guide size without a shaping service. */
  readonly fallbackGuideAdvancePx: (text: string) => number;
  /** The guide baseline above the base baseline: hpsRaise, or the retained
   * ink reserve without it (WORD_RUBY_LINE_ASCENT). */
  readonly reservePx: number;
  /** The guide's extent past its baseline toward the line head. */
  readonly headPx: number;
  /** The base is laid out on a vertical line (a vertical section run or an
   * upright frame's): the rightVertical guide axis stands half the guide
   * size nearer the base (WORD_VERTICAL_RUBY_UPRIGHT_GUIDE). */
  readonly verticalLine: boolean;
  /** Plans the guide as a vertical run (a vertical section's base). */
  readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService;
  /** snapToChars pitch: the widened cell takes whole cells. */
  readonly snapPitchPx?: number;
}

type GuideUnit = Omit<RubyGuideUnit, 'originPx'> & Readonly<{ offsetPx: number }>;

/** Horizontal guide units: graphemes at their contextual offsets; a joining
 * or pair-kerned span (WORD_ARABIC_PAIR_KERNING) is one unit. */
function horizontalUnits(
  shape: NonNullable<ReturnType<NonNullable<LayoutTextSeg['textLayoutService']>['shape']>>,
  text: string,
): GuideUnit[] {
  const units: GuideUnit[] = [];
  let whole: { span: object; index: number } | undefined;
  for (const cluster of shape.clusters ?? []) {
    const span = shape.spans.find((candidate) =>
      candidate.start <= cluster.range.start && candidate.end >= cluster.range.end) ?? shape.spans[0];
    if (!span) throw new Error('Ruby shaping produced no selected-face span');
    const kerned = wordArabicTextKerned(span.text);
    const clusterText = text.slice(cluster.range.start, cluster.range.end);
    const joined = kerned || joinsCursively(clusterText);
    const previous = units.at(-1);
    if (joined && whole?.span === span && previous && whole.index === units.length - 1) {
      units[units.length - 1] = {
        ...previous,
        text: text.slice(previous.start, cluster.range.end),
        end: cluster.range.end,
        advancePx: cluster.offsetPt + cluster.advancePt - previous.offsetPx,
      };
      continue;
    }
    units.push({
      text: clusterText,
      start: cluster.range.start,
      end: cluster.range.end,
      offsetPx: cluster.offsetPt,
      advancePx: cluster.advancePt,
      fontRoute: span.fontRoute,
      fontWeight: span.font.weight,
      fontStyle: span.font.style,
      ...(kerned ? { kerning: 'normal' as const } : {}),
    });
    whole = joined ? { span, index: units.length - 1 } : undefined;
  }
  return units;
}

/** WORD_VERTICAL_RUBY_UPRIGHT_GUIDE: the guide planned as a vertical run of
 * its own size. A turned cell spaces letter by letter at its contextual
 * prefix offsets inside the cell, so later cells keep their planned pens; a
 * joining cell stays whole. */
function verticalUnits(
  shape: NonNullable<ReturnType<NonNullable<LayoutTextSeg['textLayoutService']>['shape']>>,
  text: string,
  sizePx: number,
  service: VerticalGlyphMeasurementService,
): Readonly<{ units: GuideUnit[]; advancePx: number }> {
  const span = shape.spans[0];
  if (!span) throw new Error('Ruby shaping produced no selected-face span');
  const plan = (part: string) => service.planRun({
    text: part,
    font: canvasFontString(span.fontRoute, sizePx, span.font.weight, span.font.style),
    fontKerning: 'none',
    fontSizePt: sizePx,
    letterSpacingPt: 0,
    charScale: 1,
    growTrRotateInk: false,
    writingMode: 'vertical-rl',
  });
  const face = { fontRoute: span.fontRoute, fontWeight: span.font.weight, fontStyle: span.font.style };
  let penPx = 0;
  const units = plan(text).flatMap((cell): GuideUnit[] => {
    const cellPenPx = penPx;
    penPx += cell.advancePt;
    const common = {
      ...face,
      orientation: cell.orientation,
      ...(cell.verticalFeature ? { verticalFeature: true as const } : {}),
    };
    const bounds = cell.orientation === 'sideways' && !joinsCursively(cell.text)
      ? [...new Set([0, ...graphemeClusterOffsets(cell.text), cell.text.length])].sort((a, b) => a - b)
      : [];
    if (bounds.length <= 2) {
      return [{
        ...common,
        text: cell.text,
        start: cell.range.start,
        end: cell.range.end,
        offsetPx: cell.originPt,
        advancePx: cell.advancePt,
        ...(cell.drawOffsetPt.xPt !== 0 || cell.drawOffsetPt.yPt !== 0 ? { drawOffsetPx: cell.drawOffsetPt } : {}),
      }];
    }
    const prefixPx = (end: number) => end === 0 ? 0 : end === cell.text.length
      ? cell.advancePt
      : plan(cell.text.slice(0, end)).reduce((sum, part) => sum + part.advancePt, 0);
    const letterOrigin = cell.originPt - cellPenPx;
    return bounds.slice(0, -1).map((start, index) => {
      const end = bounds[index + 1]!;
      const startPx = prefixPx(start);
      return {
        ...common,
        text: cell.text.slice(start, end),
        start: cell.range.start + start,
        end: cell.range.start + end,
        offsetPx: cellPenPx + letterOrigin + startPx,
        advancePx: prefixPx(end) - startPx,
        ...(cell.drawOffsetPt.xPt !== 0 || cell.drawOffsetPt.yPt !== 0 ? { drawOffsetPx: cell.drawOffsetPt } : {}),
      };
    });
  });
  return { units, advancePx: penPx };
}

/** WORD_RUBY_ALIGN_PLACEMENT: the one geometry of the ruby run that `s` is a
 * piece of, computed once per ruby (and layout scale) and shared by every
 * piece: the base box, the guide units, the widened cell and the raise. */
export function rubyCellGeometry(s: LayoutTextSeg, inputs: RubyCellInputs): RubyCellGeometry | undefined {
  const group = s.rubyGroup;
  if (!group) return undefined;
  if (group.cell?.scale === inputs.scale) return group.cell;
  void WORD_RUBY_ALIGN_PLACEMENT;
  const { ruby, pieces } = group;
  const first = pieces[0]!;
  const sizePx = ruby.fontSizePt * inputs.scale;
  const shape = first.textLayoutService && first.textShapeRequest
    ? first.textLayoutService.shape({
        ...independentTextShapeRequest(first.textShapeRequest, ruby.text),
        fontSizePt: sizePx,
        measure: true,
      })
    : undefined;
  const vertical = shape !== undefined && inputs.verticalGlyphMeasurement !== undefined
    && (first.verticalRun === true || first.tateChuYoko === true);
  const planned = vertical ? verticalUnits(shape, ruby.text, sizePx, inputs.verticalGlyphMeasurement!) : undefined;
  const units = planned?.units ?? (shape ? horizontalUnits(shape, ruby.text) : []);
  const guidePx = planned?.advancePx ?? shape?.advancePt ?? inputs.fallbackGuideAdvancePx(ruby.text);
  const basePx = pieces.reduce((sum, piece) => sum + inputs.baseAdvancePx(piece), 0);
  // Only an authored alignment widens the base to a longer guide.
  const widenedPx = ruby.align ? Math.max(basePx, guidePx) : basePx;
  const pitch = inputs.snapPitchPx;
  const advancePx = pitch && pitch > 0 && widenedPx > basePx
    ? Math.ceil(widenedPx / pitch - 1e-9) * pitch
    : widenedPx;
  const overhangPx = widenedPx - basePx;
  const baseOffsetPx = wordRubyBaseOffset(ruby.align, overhangPx);
  const spacing = wordRubyGuideSpacing(ruby.align, widenedPx - guidePx, units.length);
  const cell: RubyCellGeometry = {
    scale: inputs.scale,
    basePx,
    guidePx,
    advancePx,
    baseOffsetPx,
    units: units.map(({ offsetPx, ...unit }, index) => ({
      ...unit,
      originPx: spacing.leadPt + offsetPx + index * spacing.gapPt,
    })),
    raisePx: inputs.reservePx - (inputs.verticalLine && ruby.align === 'rightVertical' ? sizePx / 2 : 0),
    headPx: inputs.headPx,
    vertical,
  };
  group.cell = cell;
  return cell;
}

/** The share of the cell's widening this piece advances by: the base offset
 * before the first piece, the rest after the last. */
export function rubyPieceWideningPx(s: LayoutTextSeg, cell: RubyCellGeometry | undefined): number {
  const group = s.rubyGroup;
  if (!group || !cell) return 0;
  const index = s.rubyPieceIndex ?? 0;
  const isFirst = index === 0;
  const isLast = index === group.pieces.length - 1;
  return (isFirst ? cell.baseOffsetPx : 0)
    + (isLast ? cell.advancePx - cell.basePx - cell.baseOffsetPx : 0);
}

/** The base's offset inside its cell, applied to the first piece's glyphs. */
export function rubyBaseOffsetPx(s: LayoutTextSeg): number {
  return s.ruby !== undefined ? s.rubyGroup?.cell?.baseOffsetPx ?? 0 : 0;
}
