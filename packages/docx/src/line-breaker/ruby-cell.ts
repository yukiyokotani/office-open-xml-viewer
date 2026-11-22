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

/** The letters a turned guide cell is spaced by: its graphemes, with each
 * run of cursively joining graphemes kept whole so it keeps its joined
 * forms, while the letters around it are spaced on their own. */
function turnedGuideLetters(text: string): readonly Readonly<{ start: number; end: number }>[] {
  const bounds = [...new Set([0, ...graphemeClusterOffsets(text), text.length])].sort((a, b) => a - b);
  const letters: { start: number; end: number; joins: boolean }[] = [];
  for (let index = 0; index + 1 < bounds.length; index++) {
    const start = bounds[index]!;
    const end = bounds[index + 1]!;
    const joins = joinsCursively(text.slice(start, end));
    const previous = letters.at(-1);
    if (joins && previous?.joins) previous.end = end;
    else letters.push({ start, end, joins });
  }
  return letters;
}

type PlannedGuideCell = ReturnType<VerticalGlyphMeasurementService['planRun']>[number];

/** WORD_VERTICAL_RUBY_UPRIGHT_GUIDE: the guide planned as a vertical run of
 * its own size, one unit per planned cell. Only when the alignment spaces
 * the guide's letters (`spaced`) is a turned cell split into its letters
 * (turnedGuideLetters), each at its contextual prefix offset inside the
 * cell, so later cells keep their planned pens. */
function verticalUnits(
  cells: readonly PlannedGuideCell[],
  face: Pick<GuideUnit, 'fontRoute' | 'fontWeight' | 'fontStyle'>,
  spaced: boolean,
  advanceOf: (text: string) => number,
): GuideUnit[] {
  let penPx = 0;
  return cells.flatMap((cell): GuideUnit[] => {
    const cellPenPx = penPx;
    penPx += cell.advancePt;
    const common = {
      ...face,
      orientation: cell.orientation,
      ...(cell.verticalFeature ? { verticalFeature: true as const } : {}),
      ...(cell.drawOffsetPt.xPt !== 0 || cell.drawOffsetPt.yPt !== 0 ? { drawOffsetPx: cell.drawOffsetPt } : {}),
    };
    const letters = spaced && cell.orientation === 'sideways' ? turnedGuideLetters(cell.text) : [];
    if (letters.length <= 1) {
      return [{
        ...common,
        text: cell.text,
        start: cell.range.start,
        end: cell.range.end,
        offsetPx: cell.originPt,
        advancePx: cell.advancePt,
      }];
    }
    // The contextual prefix advance before each letter, in linear work.
    // Library policy: a turned letter's context is its predecessor (a pair
    // adjustment; a joining run is one letter), so each prefix extends the
    // previous one by the predecessor's advance plus that predecessor's pair
    // adjustment with its own predecessor. Planning every prefix gives the
    // same offsets at quadratic cost.
    const texts = letters.map(({ start, end }) => cell.text.slice(start, end));
    const own = texts.map(advanceOf);
    const prefixPx = [0, own[0]!];
    for (let index = 1; index + 1 < texts.length; index++) {
      const pairPx = advanceOf(texts[index - 1]! + texts[index]!);
      prefixPx.push(prefixPx[index]! + pairPx - own[index - 1]!);
    }
    prefixPx.push(cell.advancePt);
    const letterOrigin = cell.originPt - cellPenPx;
    return letters.map(({ start, end }, index) => ({
      ...common,
      text: texts[index]!,
      start: cell.range.start + start,
      end: cell.range.start + end,
      offsetPx: cellPenPx + letterOrigin + prefixPx[index]!,
      advancePx: prefixPx[index + 1]! - prefixPx[index]!,
    }));
  });
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
  const service = inputs.verticalGlyphMeasurement;
  const vertical = service !== undefined && (first.verticalRun === true || first.tateChuYoko === true);
  const shape = first.textLayoutService && first.textShapeRequest
    ? first.textLayoutService.shape({
        ...independentTextShapeRequest(first.textShapeRequest, ruby.text),
        fontSizePt: sizePx,
        measure: true,
        // A vertical guide takes its geometry from its planned cells; only
        // its face is read from the shape.
        ...(vertical ? { clusterGeometry: false } : {}),
      })
    : undefined;
  const span = vertical && shape ? shape.spans[0] : undefined;
  if (vertical && shape && !span) throw new Error('Ruby shaping produced no selected-face span');
  const plan = span ? (part: string) => service!.planRun({
    text: part,
    font: canvasFontString(span.fontRoute, sizePx, span.font.weight, span.font.style),
    fontKerning: 'none',
    fontSizePt: sizePx,
    letterSpacingPt: 0,
    charScale: 1,
    growTrRotateInk: false,
    writingMode: 'vertical-rl',
  }) : undefined;
  const cells = plan?.(ruby.text);
  const guidePx = cells
    ? cells.reduce((sum, cell) => sum + cell.advancePt, 0)
    : shape?.advancePt ?? inputs.fallbackGuideAdvancePx(ruby.text);
  const basePx = pieces.reduce((sum, piece) => sum + inputs.baseAdvancePx(piece), 0);
  // Only an authored alignment widens the base to a longer guide.
  const widenedPx = ruby.align ? Math.max(basePx, guidePx) : basePx;
  // Only the distributed alignments space a longer base's guide letter by
  // letter (wordRubyGuideSpacing).
  const spaced = widenedPx > guidePx
    && (ruby.align === 'distributeLetter' || ruby.align === 'distributeSpace');
  const units = cells && span && plan
    ? verticalUnits(cells, { fontRoute: span.fontRoute, fontWeight: span.font.weight, fontStyle: span.font.style },
      spaced, (part) => plan(part).reduce((sum, cell) => sum + cell.advancePt, 0))
    : shape ? horizontalUnits(shape, ruby.text) : [];
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
    vertical: cells !== undefined,
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
