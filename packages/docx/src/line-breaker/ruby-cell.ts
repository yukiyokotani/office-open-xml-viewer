import { canvasFontString, graphemeClusterOffsets } from '@silurus/ooxml-core';
import { verticalDrawMode } from '../vertical-text.js';
import { runDisplayOrder } from '../bidi-line.js';
import type { VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { independentTextShapeRequest } from '../layout/text.js';
import {
  WORD_RUBY_ALIGN_PLACEMENT,
  wordArabicTextKerned,
  wordRubyBaseOffset,
  wordRubyGuideSpacing,
} from '../layout/line-compatibility.js';
import type { LayoutTextSeg, PlacedUnit, RubyCellGeometry, RubyGroup, RubyGuideRunFace, RubyGuideUnit } from './model.js';

/** WORD_RUBY_ALIGN_PLACEMENT: the ruby cell a segment is a piece of. Every
 * piece of a ruby run carries its cell, and this is the one test of it. */
export function rubyCellOf(segment: Pick<LayoutTextSeg, 'rubyGroup'>): RubyGroup | undefined {
  return segment.rubyGroup;
}

/** WORD_RUBY_ALIGN_PLACEMENT, §17.3.2.14: the unit a segment is admitted,
 * kept and taken back with. A ruby cell and a fitText region are each
 * indivisible, so where they share a segment neither can leave the other:
 * the unit is their union. Undefined for a segment outside any ruby cell's
 * unit (a fitText region without ruby keeps its own admission). */
export function placedUnitOf(segment: Pick<LayoutTextSeg, 'placedUnit'>): PlacedUnit | undefined {
  return segment.placedUnit;
}

/** Scripts whose letters join cursively: a guide run of them is one unit,
 * shaped and painted whole, or its joining forms would be lost. */
const CURSIVE_JOINING_RE = /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Mongolian}\p{Script=Nko}\p{Script=Mandaic}\p{Script=Adlam}\p{Script=Hanifi_Rohingya}\p{Script=Sogdian}\p{Script=Manichaean}\p{Script=Psalter_Pahlavi}\p{Script=Phags_Pa}]/u;

function joinsCursively(text: string): boolean {
  return CURSIVE_JOINING_RE.test(text);
}

/** One face a shaped guide selects, for the guide's head extent. */
export interface RubyGuideFaceBox {
  readonly boxPerEm: number;
  readonly text: string;
  readonly farEast: boolean | undefined;
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
  /** The guide's extent past its baseline toward the line head, given each
   * face the shaped guide selects: its font box per em, the text it paints
   * and its Far East single-line class when its face data is admitted
   * (undefined faces without a shaping service). */
  readonly headPx: (guideFaces: readonly RubyGuideFaceBox[] | undefined) => number;
  /** The base is laid out on a vertical line (a vertical section run or an
   * upright frame's): the rightVertical guide axis stands half the guide
   * size nearer the base (WORD_VERTICAL_RUBY_UPRIGHT_GUIDE). */
  readonly verticalLine: boolean;
  /** Plans the guide as a vertical run (a vertical section's base). */
  readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService;
  /** snapToChars pitch: the widened cell takes whole cells. */
  readonly snapPitchPx?: number;
}

type GuideUnit = Omit<RubyGuideUnit, 'originPx' | 'penPx'> & Readonly<{ offsetPx: number }>;

/** Horizontal guide units: graphemes at their contextual offsets; a joining
 * or pair-kerned span (WORD_ARABIC_PAIR_KERNING) is one unit. */
function horizontalUnits(
  shape: NonNullable<ReturnType<NonNullable<LayoutTextSeg['textLayoutService']>['shape']>>,
  text: string,
): GuideUnit[] {
  const units: GuideUnit[] = [];
  let whole: { span: object; index: number } | undefined;
  // Clusters and font spans both run in text order: one cursor finds each
  // cluster's span.
  let spanIndex = 0;
  for (const cluster of shape.clusters ?? []) {
    while (spanIndex + 1 < shape.spans.length && shape.spans[spanIndex]!.end <= cluster.range.start) spanIndex += 1;
    const span = shape.spans[spanIndex];
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
      faceBoxPx: { ascentPx: span.ascentPt, descentPx: span.descentPt },
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

/** A planned guide cell (origin and range rebased to the whole guide) with
 * the face of the font span it was planned in and that span's measurer. */
interface SpanGuideCell {
  readonly cell: PlannedGuideCell;
  readonly face: Pick<GuideUnit, 'fontRoute' | 'fontWeight' | 'fontStyle'>;
  readonly advanceOf: (text: string) => number;
}

/** WORD_VERTICAL_RUBY_UPRIGHT_GUIDE: the guide planned as a vertical run of
 * its own size, one unit per planned cell, in the face its font slot selects
 * and with the kerning it was planned with. Only when the alignment spaces
 * the guide's letters (`spaced`) is a turned cell split into its letters
 * (turnedGuideLetters), each at its contextual prefix offset inside the
 * cell, so later cells keep their planned pens. */
function verticalUnits(cells: readonly SpanGuideCell[], spaced: boolean): GuideUnit[] {
  let penPx = 0;
  return cells.flatMap(({ cell, face, advanceOf }): GuideUnit[] => {
    const cellPenPx = penPx;
    penPx += cell.advancePt;
    const common = {
      ...face,
      kerning: GUIDE_PLAN_KERNING,
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

type GuideShape = NonNullable<ReturnType<NonNullable<LayoutTextSeg['textLayoutService']>['shape']>>;

const FONT_SLOTS = ['ascii', 'highAnsi', 'eastAsia', 'complexScript'] as const;
type FontSlot = typeof FONT_SLOTS[number];

/** Guide runs shaped as one text: the font each rFonts slot takes among the
 * runs whose characters use that slot, and the one face they all select. */
interface GuideSpan {
  readonly start: number;
  readonly texts: string[];
  readonly shapes: GuideShape[];
  readonly face: string | undefined;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly rtl: boolean;
  readonly complexScript: boolean;
  readonly fontHint: RubyGuideRunFace['fontHint'];
  readonly fonts: Record<FontSlot, string | null | undefined>;
  readonly used: Set<FontSlot>;
}

/** A guide run as it is shaped: its text, and the weight and slant its
 * text takes. */
type ShapedGuideRun = Pick<RubyGuideRunFace, 'text' | 'fonts' | 'bold' | 'italic' | 'rtl' | 'complexScript' | 'fontHint'>;

/** Each position's rank among `offsets`, from the start of the line. */
function ranksOf(offsets: readonly number[]): number[] {
  const ranks = new Array<number>(offsets.length);
  offsets.map((offset, index) => [offset, index] as const)
    .sort((left, right) => left[0] - right[0] || left[1] - right[1])
    .forEach(([, index], rank) => { ranks[index] = rank; });
  return ranks;
}

/** The guide runs whose text is shaped together, none parting a grapheme
 * cluster (UAX #29) and each with one weight and slant.
 *
 * §17.3.2.2 / §17.3.2.17: complex-script text takes bCs and iCs, other text
 * b and i. A run that is complex script as a whole (§17.3.2.30 rtl,
 * §17.3.2.7 cs) takes the complex-script axes; a run whose two pairs differ
 * is divided where its text changes between the complex-script slot and the
 * others, as a base run is divided into pieces.
 *
 * A cluster belongs to the run of its first character: a combining mark or
 * variation selector, or the rest of an emoji sequence, in a later w:rt run
 * moves to the run of its base and takes that run's faces. */
function shapedGuideRuns(
  ruby: NonNullable<LayoutTextSeg['ruby']>,
  slotsOf: (run: ShapedGuideRun) => readonly Readonly<{ start: number; end: number; complexScript: boolean }>[],
): readonly ShapedGuideRun[] {
  const byAxes = (ruby.guideRuns ?? []).flatMap((run): ShapedGuideRun[] => {
    const cs = { ...run, bold: run.boldCs, italic: run.italicCs };
    if (run.complexScript) return [cs];
    if (run.boldCs === run.bold && run.italicCs === run.italic) return [run];
    const pieces: ShapedGuideRun[] = [];
    let complexScript: boolean | undefined;
    for (const span of slotsOf(run)) {
      const text = run.text.slice(span.start, span.end);
      if (span.complexScript === complexScript) pieces[pieces.length - 1] = { ...pieces.at(-1)!, text: pieces.at(-1)!.text + text };
      else pieces.push({ ...(span.complexScript ? cs : run), text });
      complexScript = span.complexScript;
    }
    return pieces;
  });
  // Each run starts at the first cluster boundary at or after its start.
  const boundaries = [0, ...graphemeClusterOffsets(ruby.text), ruby.text.length];
  let boundary = 0;
  let offset = 0;
  const starts = byAxes.map((run) => {
    while (boundaries[boundary]! < offset) boundary += 1;
    offset += run.text.length;
    return boundaries[boundary]!;
  });
  return byAxes.flatMap((run, index) => {
    const text = ruby.text.slice(starts[index]!, starts[index + 1] ?? ruby.text.length);
    return text === '' ? [] : [{ ...run, text }];
  });
}

/** The guide's w:rt runs (§17.3.3.24) shaped at the guide size, each with
 * its own faces over the base run's request: the authored runs when they
 * spell the guide (shapedGuideRuns), else the whole guide in the base run's
 * faces.
 *
 * Adjacent runs whose characters select the same face (family, weight and
 * slant) are shaped as one text, so a run seam inside a word keeps its
 * contextual (joining, kerned) forms. Each character keeps the font of its
 * own run's slot: runs join only when, for every slot the characters of
 * either use, they name the same font, so the joined text is shaped with one
 * font per slot. Nothing else separates runs: a slot no character uses, or a
 * script control that changes no character's face, does not. The joined
 * text is shaped with the first run's controls and kept only if every
 * character still selects the one face. Each run is shaped once to learn
 * its face and slots, and each joined text once: linear in the guide. */
function guideShapes(
  first: LayoutTextSeg,
  ruby: NonNullable<LayoutTextSeg['ruby']>,
  sizePx: number,
  vertical: boolean,
): readonly Readonly<{ start: number; text: string; shape: GuideShape; rtl: boolean }>[] | undefined {
  const service = first.textLayoutService;
  const request = first.textShapeRequest;
  if (!service || !request) return undefined;
  const fontsOf = (run: ShapedGuideRun) =>
    Object.fromEntries(FONT_SLOTS.map((slot) => [slot, run.fonts[slot] ?? request.fonts[slot]])) as GuideSpan['fonts'];
  const shapeText = (
    text: string,
    face?: Pick<GuideSpan, 'fonts' | 'bold' | 'italic' | 'complexScript' | 'fontHint'>,
    measure = true,
  ): GuideShape =>
    service.shape({
      ...independentTextShapeRequest(request, text),
      // The guide run's faces are resolved, theme references included, and
      // its slots are selected by its own script controls, not the base's.
      ...(face ? {
        fonts: face.fonts, themeFonts: undefined, themeFontPresence: undefined,
        weight: face.bold ? 700 : 400, style: face.italic ? 'italic' as const : 'normal' as const,
        complexScript: face.complexScript, fontHint: face.fontHint,
      } : {}),
      fontSizePt: sizePx,
      measure,
      // A vertical guide takes its geometry from its planned cells; only
      // its faces are read from the shape.
      ...(vertical || !measure ? { clusterGeometry: false } : {}),
    });
  if (!ruby.guideRuns) return [{ start: 0, text: ruby.text, shape: shapeText(ruby.text), rtl: false }];
  const faceOf = (shape: GuideShape): string | undefined => {
    const keys = new Set(shape.spans.map((span) => `${span.fontRoute.fingerprint}|${span.font.weight}|${span.font.style}`));
    return keys.size === 1 ? [...keys][0] : undefined;
  };
  const slotsOf = (shape: GuideShape): Set<FontSlot> => new Set(shape.spans.flatMap((span) =>
    [span.script, ...(span.semanticSlotSpans ?? []).map((slot) => slot.script)]));
  const runs = shapedGuideRuns(ruby, (run) =>
    shapeText(run.text, { ...run, fonts: fontsOf(run), fontHint: run.fontHint }, false).spans.flatMap((span) =>
      (span.semanticSlotSpans ?? [span]).map((slot) => ({
        start: slot.start, end: slot.end, complexScript: slot.script === 'complexScript',
      }))));
  const spans: GuideSpan[] = [];
  let start = 0;
  for (const run of runs) {
    const fonts = fontsOf(run);
    const controls = {
      bold: run.bold, italic: run.italic, rtl: run.rtl, complexScript: run.complexScript, fontHint: run.fontHint,
    };
    const shape = shapeText(run.text, { fonts, ...controls });
    const face = faceOf(shape);
    const used = slotsOf(shape);
    const previous = spans.at(-1);
    // Direction is not a control that only selects a face: runs of opposite
    // direction stand apart on the line.
    if (previous && face !== undefined && previous.face === face && previous.rtl === run.rtl
      && [...used].every((slot) => !previous.used.has(slot) || previous.fonts[slot] === fonts[slot])) {
      previous.texts.push(run.text);
      previous.shapes.push(shape);
      for (const slot of used) {
        previous.fonts[slot] = fonts[slot];
        previous.used.add(slot);
      }
    } else {
      spans.push({ start, texts: [run.text], shapes: [shape], face, ...controls, fonts, used });
    }
    start += run.text.length;
  }
  return spans.flatMap((span) => {
    if (span.shapes.length > 1) {
      const text = span.texts.join('');
      const joined = shapeText(text, span);
      // Characters are classified in context: joined runs that no longer
      // select their one face stay apart.
      if (faceOf(joined) === span.face) return [{ start: span.start, text, shape: joined, rtl: span.rtl }];
    }
    let runStart = span.start;
    return span.shapes.map((shape, index) => {
      const entry = { start: runStart, text: span.texts[index]!, shape, rtl: span.rtl };
      runStart += entry.text.length;
      return entry;
    });
  });
}

/** A vertical guide is planned, and so painted, without kerning. */
const GUIDE_PLAN_KERNING = 'none' as const;

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
  // §17.3.3.24: each w:rt run is shaped with its own faces (guideShapes).
  const shapes = guideShapes(first, ruby, sizePx, vertical);
  if (vertical && shapes?.some(({ shape }) => shape.spans.length === 0)) {
    throw new Error('Ruby shaping produced no selected-face span');
  }
  // Each font span (its rFonts slot) is planned in its own face.
  let spanPenPx = 0;
  // The guide is planned by extended grapheme clusters (UAX #29). A cluster
  // that stands upright or is rotated is one glyph cell: it is planned by
  // its first character and painted whole (its combining marks, variation
  // selector or the rest of an emoji sequence with it), in that character's
  // face, whichever font span or run holds the rest. Clusters laid sideways
  // are planned as the text they are, shaped in context.
  const clusterStarts = vertical ? new Set([0, ...graphemeClusterOffsets(ruby.text)]) : undefined;
  let lastCell: { cell: { range: { start: number; end: number }; text: string } } | undefined;
  const cells = vertical && shapes ? shapes.flatMap(({ shape, start }) => shape.spans.flatMap((span): SpanGuideCell[] => {
    const plan = (part: string) => service!.planRun({
      text: part,
      font: canvasFontString(span.fontRoute, sizePx, span.font.weight, span.font.style),
      fontKerning: GUIDE_PLAN_KERNING,
      fontSizePt: sizePx,
      letterSpacingPt: 0,
      charScale: 1,
      growTrRotateInk: false,
      writingMode: 'vertical-rl',
    });
    const advanceOf = (part: string) => plan(part).reduce((sum, cell) => sum + cell.advancePt, 0);
    const face = { fontRoute: span.fontRoute, fontWeight: span.font.weight, fontStyle: span.font.style };
    const originPx = spanPenPx;
    const spanStart = start + span.start;
    // The text the span is planned by, and where each UTF-16 unit of it
    // lies in the guide.
    let planned = '';
    const offsets: number[] = [];
    let keep = !lastCell;
    for (let offset = 0; offset < span.text.length;) {
      const char = String.fromCodePoint(span.text.codePointAt(offset)!);
      // A cluster's first character decides; the rest of it follows.
      const opensCluster = clusterStarts!.has(spanStart + offset);
      if (opensCluster) keep = true;
      if (keep) {
        for (let unit = 0; unit < char.length; unit += 1) offsets.push(spanStart + offset + unit);
        planned += char;
      }
      if (opensCluster) keep = verticalDrawMode(char.codePointAt(0)!) === 'sideways';
      offset += char.length;
    }
    const spanEnd = spanStart + span.text.length;
    // What opens the span inside a cluster belongs to the cell before it.
    const firstBase = offsets[0] ?? spanEnd;
    if (lastCell && firstBase > spanStart) {
      lastCell.cell.text += ruby.text.slice(lastCell.cell.range.end, firstBase);
      lastCell.cell.range.end = firstBase;
    }
    const spanCells = (planned === '' ? [] : plan(planned)).map((cell): SpanGuideCell => {
      const range = { start: offsets[cell.range.start]!, end: offsets[cell.range.end] ?? spanEnd };
      // The cell paints the text the planner selected for it (a vertical
      // form may stand for the source character), then the rest of its
      // cluster, which was not planned.
      const plannedEnd = offsets[cell.range.end - 1]! + 1;
      return {
        cell: { ...cell, range, text: cell.text + ruby.text.slice(plannedEnd, range.end), originPt: originPx + cell.originPt },
        face,
        advanceOf,
      };
    });
    lastCell = spanCells.at(-1) ?? lastCell;
    spanPenPx += spanCells.reduce((sum, { cell }) => sum + cell.advancePt, 0);
    return spanCells;
  })) : undefined;
  // A horizontal guide's shaped runs stand in bidirectional order
  // (runDisplayOrder, resolved over the guide's characters): right-to-left
  // letters in runs of different formatting are separate shapes, and the
  // later one stands before the earlier one on the line, while digits keep
  // their order.
  let runPenPx = 0;
  const shapePenPx = new Map<number, number>();
  if (!cells && shapes) {
    for (const index of runDisplayOrder(shapes, group.baseRtl === true)) {
      shapePenPx.set(index, runPenPx);
      runPenPx += shapes[index]!.shape.advancePt;
    }
  }
  const horizontal = !cells && shapes ? shapes.flatMap(({ shape, start, text }, index) =>
    horizontalUnits(shape, text).map((unit) => ({
      ...unit, start: start + unit.start, end: start + unit.end, offsetPx: shapePenPx.get(index)! + unit.offsetPx,
    }))) : undefined;
  const guidePx = cells
    ? spanPenPx
    : shapes ? runPenPx : inputs.fallbackGuideAdvancePx(ruby.text);
  const basePx = pieces.reduce((sum, piece) => sum + inputs.baseAdvancePx(piece), 0);
  // Only an authored alignment widens the base to a longer guide.
  const widenedPx = ruby.align ? Math.max(basePx, guidePx) : basePx;
  // Only the distributed alignments space a longer base's guide letter by
  // letter (wordRubyGuideSpacing).
  const spaced = widenedPx > guidePx
    && (ruby.align === 'distributeLetter' || ruby.align === 'distributeSpace');
  const units = cells ? verticalUnits(cells, spaced) : horizontal ?? [];
  const pitch = inputs.snapPitchPx;
  const advancePx = pitch && pitch > 0 && widenedPx > basePx
    ? Math.ceil(widenedPx / pitch - 1e-9) * pitch
    : widenedPx;
  const overhangPx = widenedPx - basePx;
  const baseOffsetPx = wordRubyBaseOffset(ruby.align, overhangPx);
  const spacing = wordRubyGuideSpacing(ruby.align, widenedPx - guidePx, units.length);
  const unitRanks = ranksOf(units.map((unit) => unit.offsetPx));
  const cell: RubyCellGeometry = {
    scale: inputs.scale,
    basePx,
    guidePx,
    advancePx,
    baseOffsetPx,
    // Units are kept in guide order; the alignment's gaps open between
    // them in the order they stand on the line.
    units: units.map(({ offsetPx, ...unit }, index) => ({
      ...unit,
      penPx: offsetPx,
      originPx: spacing.leadPt + offsetPx + unitRanks[index]! * spacing.gapPt,
    })),
    raisePx: inputs.reservePx - (inputs.verticalLine && ruby.align === 'rightVertical' ? sizePx / 2 : 0),
    // The guide's own faces bound its extent, each by its own box and class.
    headPx: inputs.headPx(shapes && shapes.length > 0 && sizePx > 0
      ? shapes.flatMap(({ shape }) => shape.spans.map((span): RubyGuideFaceBox => ({
        boxPerEm: (span.ascentPt + span.descentPt) / sizePx,
        text: span.text,
        farEast: group.guideFarEastClass?.(span.font, span.text),
      })))
      : undefined),
    vertical: cells !== undefined,
  };
  group.cell = cell;
  return cell;
}

/** The pieces of the cell that stand first and last on the line. The
 * base's runs are reordered by direction like any text (runDisplayOrder):
 * in a left-to-right paragraph consecutive right-to-left pieces stand in
 * reverse, so the last of them leads, while a left-to-right piece or a
 * number keeps its place. */
function edgePieces(group: RubyGroup): Readonly<{ leading: number; trailing: number }> {
  let edges = cellEdges.get(group);
  if (!edges) {
    const order = runDisplayOrder(group.pieces, group.baseRtl === true);
    edges = { leading: order[0] ?? 0, trailing: order.at(-1) ?? 0 };
    cellEdges.set(group, edges);
  }
  return edges;
}

/** A cell's pieces are fixed once its pass has formed it (prepareBreakQueue),
 * so their order is resolved once per cell, not per measurement. */
const cellEdges = new WeakMap<RubyGroup, Readonly<{ leading: number; trailing: number }>>();

/** The share of the cell's widening this piece advances by: the base offset
 * before the piece that leads the cell on the line, the rest after the piece
 * that ends it, so the widening lies around the base and never inside it. */
export function rubyPieceWideningPx(s: LayoutTextSeg, cell: RubyCellGeometry | undefined): number {
  const group = s.rubyGroup;
  if (!group || !cell) return 0;
  const index = s.rubyPieceIndex ?? 0;
  const { leading, trailing } = edgePieces(group);
  return (index === leading ? cell.baseOffsetPx : 0)
    + (index === trailing ? cell.advancePx - cell.basePx - cell.baseOffsetPx : 0);
}

/** WORD_RUBY_ALIGN_PLACEMENT: the paint origins of the guide units of the
 * ruby run placed as `pieces` (its first piece first), from the cell start. A widened cell is
 * its own allocation and keeps the origins it was planned with. An
 * unwidened base shares its snapToChars script block, whose allocation is
 * settled only once the block is complete, so its guide is aligned on the
 * base as placed: from the first piece's glyph start (after an automatic
 * East Asian gap and a Latin block's leading pad) to the last piece's glyph
 * end (before a Latin block's trailing pad). East Asian cells centre their
 * glyphs in whole cells, so their box is the whole share. */
export function rubyGuideOriginsPx(pieces: readonly LayoutTextSeg[], cell: RubyCellGeometry): readonly number[] {
  const first = pieces[0];
  const group = first?.rubyGroup;
  if (!first || !group || cell.advancePx > cell.basePx) return cell.units.map((unit) => unit.originPx);
  const last = pieces.at(-1)!;
  const startPx = (first.autoSpaceLeadingPx ?? 0)
    + (first.snapGridClass === 'latin' ? first.snapGridLeadingPadPx ?? 0 : 0);
  const endPx = pieces.reduce((sum, piece) => sum + piece.measuredWidth, 0)
    - (last.snapGridClass === 'latin' ? last.snapGridTrailingPadPx ?? 0 : 0);
  const spacing = wordRubyGuideSpacing(group.ruby.align, endPx - startPx - cell.guidePx, cell.units.length);
  const ranks = ranksOf(cell.units.map((unit) => unit.penPx));
  return cell.units.map((unit, index) => startPx + spacing.leadPt + unit.penPx + ranks[index]! * spacing.gapPt);
}

/** The base's offset inside its cell, applied to the glyphs of the piece
 * that leads the cell on the line. */
export function rubyBaseOffsetPx(s: LayoutTextSeg): number {
  const group = s.rubyGroup;
  return group && (s.rubyPieceIndex ?? 0) === edgePieces(group).leading ? group.cell?.baseOffsetPx ?? 0 : 0;
}
