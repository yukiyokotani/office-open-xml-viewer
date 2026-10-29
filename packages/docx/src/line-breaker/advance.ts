import {
  CONDENSED_CONTEXT_GRAPHEMES,
  condensedClusterFloorPx,
  countedSpacingAdvance,
  countSpacingCharacters,
} from '../character-spacing.js';
export { CONDENSED_CONTEXT_GRAPHEMES, condensedClusterFloorPx };
import { sliceTextBreakWindow } from './text-break-window.js';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import {
  EAST_ASIAN_RE,
  shapeRequestKeyFields,
  shapeRequestKeyFieldsArePrimitive,
  sliceTextShapeRequest,
  semanticSlotStartIndex,
  type TextLayoutService,
  type TextShapeRequest,
} from '../layout/text.js';
import { wordBalancedSpaceCellAdjustmentApplies } from '../layout/line-compatibility.js';
import { type DocGridCtx, type LayoutSeg, type LayoutTextSeg } from './model.js';

/** Shaping and line allocation are derived from the current text slice. A
 * segment split for wrapping must not retain geometry from the parent slice;
 * measurement and addToLine recompute these facts from the new text. */
export const RESET_SLICED_TEXT_MEASUREMENT = {
  leadingWordBoundaryPx: undefined,
  shapedClusters: undefined,
  shapedSpaceClusters: undefined,
  selectedFaceInkBounds: undefined,
  selectedFaceFontBox: undefined,
  snapGridClass: undefined,
  snapGridNaturalWidthPx: undefined,
  snapGridLeadingPadPx: undefined,
  snapGridTrailingPadPx: undefined,
  snapGridCellPitchPx: undefined,
} as const;


/** Code points whose presence marks a line as East Asian for docGrid line-cell
 *  rounding: CJK symbols/punctuation, Hiragana, Katakana, CJK Unified +
 *  Extension A, compatibility ideographs, Hangul, and fullwidth forms. Content
 *  test only — not a font-name heuristic (cf. packages/docx/CLAUDE.md). */
/** Per-character character-grid delta in px, before applying the grid's scope. */
export function gridCharDeltaPx(grid: DocGridCtx | undefined, scale: number): number {
  if (!grid || grid.charSpacePt == null) return 0;
  if (grid.type !== 'linesAndChars' && grid.type !== 'snapToChars') return 0;
  return grid.charSpacePt * scale;
}


/** Count of East-Asian (full-width) code points in `text` — the glyphs the
 *  character grid snaps to cells. Uses the same {@link EAST_ASIAN_RE} content
 *  predicate as docGrid line-cell rounding (no font-name heuristic). */
export function eaGlyphCount(text: string): number {
  let n = 0;
  for (const ch of text) if (EAST_ASIAN_RE.test(ch)) n++;
  return n;
}


/** Total character-grid delta gained by a segment. `linesAndChars` applies its
 *  authored pitch to every character. The non-`linesAndChars` branch preserves
 *  the renderer's pre-existing East-Asian-only fallback; full snap-to-character
 *  grid-unit allocation is handled by the block allocator below. */
export function gridSegDeltaPx(
  text: string,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  const deltaPx = gridCharDeltaPx(grid, scale);
  if (deltaPx === 0 || text.length === 0) return 0;
  const cps = [...text];
  if (grid?.type === 'linesAndChars') return cps.length * deltaPx;
  return eaGlyphCount(text) === cps.length ? cps.length * deltaPx : 0;
}


/** Resolve the per-glyph character-grid delta for one text segment. */
export function segmentCharacterGridDeltaPx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (seg.snapToCharacterGrid === false) return 0;
  // snapToChars allocates full cells/blocks in layoutLines. It is not a
  // per-glyph letter-spacing delta.
  if (grid?.type === 'snapToChars') return 0;
  if (grid?.type === 'linesAndChars' && seg.widthBalanceGridDeltaFactor !== undefined) {
    return gridCharDeltaPx(grid, scale) * seg.widthBalanceGridDeltaFactor;
  }
  const total = gridSegDeltaPx(seg.text, grid, scale);
  return total === 0 ? 0 : gridCharDeltaPx(grid, scale);
}


/** ECMA-376 §17.3.2.35 `<w:spacing>` — the per-GLYPH character-spacing pitch in
 *  px for a segment (its authored points × the paint scale). Unlike the docGrid
 *  delta this applies to EVERY code point of the run, not just East-Asian ones
 *  ("the amount of character pitch … added after each character in this run").
 *  0 when the run declares no `w:spacing`. */
export function charSpacingDeltaPx(seg: LayoutTextSeg, scale: number): number {
  // §17.3.2.14 fitText replaces cached §17.3.2.35 spacing with the resolved
  // region gap. The paint path already reads this authority.
  if (seg.fitTextPerGapPx !== undefined) return seg.fitTextPerGapPx;
  return effectiveCharacterSpacingPt(seg) * scale;
}


/** The uniform paint/measure pitch contributed by run-authored `w:spacing`.
 * Document-level punctuation compression is a one-time trailing-cell advance
 * adjustment, not a per-glyph Canvas letter-spacing value. */
export function effectiveCharacterSpacingPt(seg: LayoutTextSeg): number {
  return seg.charSpacing ?? 0;
}


/** WORD_EAST_ASIAN_AUTO_SPACE gap a measurement of `seg` from its first
 * character includes, in layout px: the placement's resolved value, or the
 * same-line value while the segment is still being fitted. */
export function autoSpaceLeadingAdvancePx(seg: LayoutTextSeg, scale: number): number {
  if (seg.autoSpaceBeforePt === undefined) return 0;
  return seg.autoSpaceLeadingPx ?? seg.autoSpaceBeforePt * scale;
}


/** {@link autoSpaceLeadingAdvancePx} under the paragraph's character grid.
 * Grid-cell allocation under `snapToChars`
 * (WORD_SNAP_TO_CHARS_SCRIPT_BLOCK_ALLOCATION) owns that grid's spacing; the
 * automatic gap is not applied there. */
export function leadingAutoSpacePx(
  state: Readonly<{ characterGrid?: DocGridCtx; scale: number }>,
  seg: LayoutTextSeg,
): number {
  if (state.characterGrid?.type === 'snapToChars') return 0;
  return autoSpaceLeadingAdvancePx(seg, state.scale);
}


export function punctuationCompressionTotalPt(seg: LayoutTextSeg): number {
  return seg.punctuationCompressions?.reduce(
    (sum, compression) => sum + compression.adjustmentPt,
    0,
  ) ?? 0;
}


/** §17.15.3.3 plus the registered Word space-sequence projection. A segment
 * created by splitTextForLayout contains at most one trailing U+0020 sequence;
 * slicing keeps the sequence flag, so prefixes/tails count only their retained
 * authored spaces. */
export function widthBalanceSpaceAdjustmentTotalPt(
  seg: LayoutTextSeg,
  characterGrid?: DocGridCtx,
): number {
  return widthBalanceSpaceAdjustmentForTextPt(seg, seg.text, characterGrid);
}


/** Retained-cluster projection of the same authored-space adjustment used by
 * {@link widthBalanceSpaceAdjustmentTotalPt}. Keeping this calculation on the
 * immutable segment fact makes line measurement, cluster hit geometry, RTL
 * whitespace anchoring, and paint-plan slicing consume one width authority. */
export function widthBalanceSpaceAdjustmentForTextPt(
  seg: LayoutTextSeg,
  text: string,
  characterGrid?: DocGridCtx,
): number {
  // Both properties replace the ordinary inline advance with their own
  // specification-defined region/cell width. Their interaction with Word's
  // proportional-space projection is outside the observation matrix, so keep
  // the preexisting override geometry intact in every retained consumer.
  if (seg.fitTextPerGapPx !== undefined || seg.tateChuYoko) return 0;
  if (!wordBalancedSpaceCellAdjustmentApplies(characterGrid?.type)) return 0;
  if (!seg.widthBalanceSpaceSequence || seg.widthBalanceSpaceAdjustmentPt === undefined) {
    return 0;
  }
  let count = 0;
  for (const character of text) if (character === ' ') count += 1;
  return count * seg.widthBalanceSpaceAdjustmentPt;
}


export function slicedPunctuationCompressions(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): LayoutTextSeg['punctuationCompressions'] {
  const sliced = seg.punctuationCompressions
    ?.filter((compression) => compression.end > start && compression.end <= end)
    .map((compression) => Object.freeze({
      end: compression.end - start,
      adjustmentPt: compression.adjustmentPt,
    }));
  return sliced && sliced.length > 0 ? Object.freeze(sliced) : undefined;
}


export function slicedNoBreakRanges(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): LayoutTextSeg['noBreakRanges'] {
  const sliced = seg.noBreakRanges
    ?.filter((range) => range.start >= start && range.end <= end)
    .map((range) => Object.freeze({
      start: range.start - start,
      end: range.end - start,
    }));
  return sliced && sliced.length > 0 ? Object.freeze(sliced) : undefined;
}


export function protectedNoBreakOffsets(seg: LayoutTextSeg): ReadonlySet<number> {
  return new Set(seg.noBreakRanges?.flatMap((range) => [range.start, range.end]) ?? []);
}


export function legalTextSplitAtOrBefore(
  seg: LayoutTextSeg,
  proposed: number,
  minimum = 0,
): number {
  const protectedOffsets = protectedNoBreakOffsets(seg);
  return [0, ...graphemeClusterOffsets(seg.text), seg.text.length]
    .filter((offset, index, all) => all.indexOf(offset) === index)
    .filter((offset) => offset >= minimum && offset <= proposed && !protectedOffsets.has(offset))
    .at(-1) ?? 0;
}


/** Smallest prefix that consumes a hard source seam. It includes the complete
 * authored noBreakHyphen range plus the first following grapheme when both
 * range edges live in this segment. */
export function hardJoinPrefixEnd(seg: LayoutTextSeg): number | undefined {
  if (seg.hardJoinPrev !== true || seg.text.length === 0) return undefined;
  const protectedOffsets = protectedNoBreakOffsets(seg);
  const firstLegal = [
    ...graphemeClusterOffsets(seg.text),
    seg.text.length,
  ].find((offset) => offset > 0 && !protectedOffsets.has(offset));
  return firstLegal ?? seg.text.length;
}


/** Single authority for metadata whose UTF-16 coordinates are relative to a
 * text segment. Every retained split path must use this projection. */
export function slicedTextMetadata(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): Pick<LayoutTextSeg,
  'punctuationCompressions' | 'noBreakRanges' | 'explicitBreaks' | 'textShapeRequest' | 'sourceTextOffset' | 'semanticSlotSpans' | 'semanticSlotRange' | 'script' | 'optionalHyphenBreaks'
  | 'autoSpaceBeforePt' | 'autoSpaceLeadingPx'
> {
  const slots = seg.semanticSlotSpans;
  const slotStart = (seg.semanticSlotRange?.start ?? 0) + start;
  const slotEnd = (seg.semanticSlotRange?.start ?? 0) + end;
  return {
    ...(slots ? { semanticSlotSpans: slots,
      semanticSlotRange: Object.freeze({ start: slotStart, end: slotEnd }),
      script: slots[semanticSlotStartIndex(slots, slotStart)]?.script ?? seg.script,
    } : {}),
    ...(seg.sourceTextOffset === undefined ? {} : { sourceTextOffset: seg.sourceTextOffset + start }),
    ...(seg.textShapeRequest
      ? { textShapeRequest: sliceTextShapeRequest(seg.textShapeRequest, start, end) }
      : {}),
    punctuationCompressions: slicedPunctuationCompressions(seg, start, end),
    // WORD_EAST_ASIAN_AUTO_SPACE: the gap precedes the first character only.
    autoSpaceBeforePt: start === 0 ? seg.autoSpaceBeforePt : undefined,
    autoSpaceLeadingPx: start === 0 ? seg.autoSpaceLeadingPx : undefined,
    noBreakRanges: slicedNoBreakRanges(seg, start, end),
    explicitBreaks: sliceTextBreakWindow(seg.explicitBreaks, start, end),
    optionalHyphenBreaks: seg.optionalHyphenBreaks?.filter(marker => marker.offset > start && marker.offset <= end)
      .map(marker => Object.freeze({ ...marker, offset: marker.offset - start })),
  };
}


export function tightHorizontalGraphemeInk(
  segment: LayoutTextSeg,
  grapheme: string,
  start = 0,
): Readonly<{ advancePt: number; xMinPt: number; xMaxPt: number }> | undefined {
  if (!segment.textLayoutService || !segment.textShapeRequest || grapheme.length === 0) {
    return undefined;
  }
  const shaped = segment.textLayoutService.shape({
    ...sliceTextShapeRequest(segment.textShapeRequest, start, start + grapheme.length),
    measure: true,
    clusterGeometry: false,
  });
  if (
    shaped.horizontalInkBoundsAreTight !== true
    || !shaped.inkBounds
    || !Number.isFinite(shaped.advancePt)
    || !Number.isFinite(shaped.inkBounds.xMinPt)
    || !Number.isFinite(shaped.inkBounds.xMaxPt)
  ) {
    return undefined;
  }
  const scale = segment.charScale ?? 1;
  return {
    advancePt: shaped.advancePt * scale,
    xMinPt: shaped.inkBounds.xMinPt * scale,
    xMaxPt: shaped.inkBounds.xMaxPt * scale,
  };
}


export function contextualHorizontalGraphemeAdvances(
  segment: LayoutTextSeg,
): ReadonlyMap<number, number> | undefined {
  if (!segment.textLayoutService || !segment.textShapeRequest || segment.text.length === 0) {
    return undefined;
  }
  const shaped = segment.textLayoutService.shape({
    ...segment.textShapeRequest,
    measure: true,
    clusterGeometry: true,
  });
  if (!shaped.clusters?.length) return undefined;
  const scale = segment.charScale ?? 1;
  const advances = new Map<number, number>();
  for (const cluster of shaped.clusters) {
    if (!Number.isFinite(cluster.advancePt)) return undefined;
    advances.set(cluster.range.end, cluster.advancePt * scale);
  }
  return advances;
}


/**
 * Bound document-level punctuation compression by the adjacent glyphs' tight
 * horizontal ink and contextual advance. Canvas shaping can already kern a
 * punctuation pair down to the retained half-cell; subtracting the isolated
 * glyph's removable sidebearing again would collapse the second mark to zero
 * advance. A following glyph's left ink edge also participates in the collision
 * equation. Resolve this after all source runs have been segmented so a
 * formatting seam cannot reintroduce the overlap.
 */
export function retainHorizontalPunctuationInkClearance(segs: LayoutSeg[]): void {
  let pending: Readonly<{
    segment: LayoutTextSeg;
    compressionIndex: number;
    ink: Readonly<{ advancePt: number; xMinPt: number; xMaxPt: number }>;
    contextualAdvancePt: number;
  }> | undefined;
  const adjustedBySegment = new Map<
    LayoutTextSeg,
    Array<{ end: number; adjustmentPt: number }>
  >();
  for (const candidate of segs) {
    if (!('text' in candidate) || candidate.verticalRun) {
      pending = undefined;
      continue;
    }
    const segment = candidate;
    const compressions = segment.punctuationCompressions ?? [];
    const compressionIndexByEnd = new Map(
      compressions.map((compression, index) => [compression.end, index]),
    );
    const contextualAdvances = compressions.length > 0
      ? contextualHorizontalGraphemeAdvances(segment)
      : undefined;
    const boundaries = [0, ...graphemeClusterOffsets(segment.text), segment.text.length];
    for (let index = 0; index < boundaries.length - 1; index += 1) {
      const start = boundaries[index]!;
      const end = boundaries[index + 1]!;
      if (end <= start) continue;
      const compressionIndex = compressionIndexByEnd.get(end);
      const currentInk = pending || compressionIndex !== undefined
        ? tightHorizontalGraphemeInk(segment, segment.text.slice(start, end), start)
        : undefined;
      if (pending && currentInk) {
        const adjustments = adjustedBySegment.get(pending.segment)
          ?? pending.segment.punctuationCompressions!.map((compression) => ({
            ...compression,
          }));
        const compression = adjustments[pending.compressionIndex]!;
        const retainedExtentPt = Math.max(
          0,
          pending.ink.advancePt + compression.adjustmentPt,
        );
        const retainedExtentAdjustmentPt = Math.min(
          0,
          retainedExtentPt - pending.contextualAdvancePt,
        );
        const collisionSafeAdjustmentPt = Math.min(
          0,
          pending.ink.xMaxPt
            - currentInk.xMinPt
            - pending.contextualAdvancePt,
        );
        const adjustmentPt = Math.max(
          compression.adjustmentPt,
          retainedExtentAdjustmentPt,
          collisionSafeAdjustmentPt,
        );
        if (adjustmentPt !== compression.adjustmentPt) {
          adjustments[pending.compressionIndex] = {
            end: compression.end,
            adjustmentPt,
          };
          adjustedBySegment.set(pending.segment, adjustments);
        }
      }
      pending = compressionIndex !== undefined && currentInk
        ? {
            segment,
            compressionIndex,
            ink: currentInk,
            contextualAdvancePt:
              contextualAdvances?.get(end) ?? currentInk.advancePt,
          }
        : undefined;
    }
  }
  for (const [segment, adjusted] of adjustedBySegment) {
    segment.punctuationCompressions = Object.freeze(
      adjusted.map((compression) => Object.freeze(compression)),
    );
  }
}


/** ECMA-376 §17.3.2.43 `<w:w>` — the horizontal glyph-width scale fraction of a
 *  segment (0.67 = 67%). 1 when the run declares no `w:w`. Multiplies the
 *  natural `measureText` width; the paint pass reproduces it with `ctx.scale`. */
export function charScaleFactor(seg: LayoutTextSeg): number {
  return seg.charScale ?? 1;
}


/** Canonical advance formula for a text string in a run: natural glyph width
 *  scaled by ECMA-376 §17.3.2.43 `<w:w>`, plus the §17.6.5 character-grid
 *  delta, plus one ECMA-376 §17.3.2.35 `<w:spacing>` pitch per character. */
export function textAdvanceWidth(
  naturalWidthPx: number,
  text: string,
  characterGridDeltaPx: number,
  charScale: number,
  charSpacingPx: number,
): number {
  const { scalars, pitchBearing } = countSpacingCharacters(text);
  // Preserve main's left-to-right additions, separately from retained geometry's
  // combined pitch. Regrouping these terms changes unaffected advances by ULPs.
  return naturalWidthPx * charScale
    + scalars * characterGridDeltaPx
    + pitchBearing * charSpacingPx;
}



/** Uniform native Canvas letter-spacing (px) a segment draws with: the docGrid
 *  cell delta (already scoped by {@link segmentCharacterGridDeltaPx}) PLUS the
 *  §17.3.2.35 character-spacing pitch. Because Canvas
 *  `ctx.letterSpacing` inserts the SAME advance after every glyph, the two are
 *  additive only when the grid delta applies to every glyph — i.e. a pure-EA
 *  segment (or none, when grid is inactive). For a mixed / Latin segment the
 *  grid delta is 0 (Latin is never snapped, §17.6.5) so only char-spacing
 *  contributes, and the value is still uniform across the segment. Measurement
 *  counts their contributions separately via textAdvanceWidth.
 *  Native engines may space standalone or spacing marks differently. Earlier
 *  measurement counted every scalar; retain the established contextual paint
 *  behavior instead of splitting shaped text to force measurement parity. */
export function segLetterSpacingPx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (seg.fitTextPerGapPx !== undefined) return seg.fitTextPerGapPx;
  return segmentCharacterGridDeltaPx(seg, grid, scale) + charSpacingDeltaPx(seg, scale);
}


/** One grapheme cluster's natural (unspaced, unscaled) advance, in the same
 * px space as the caller's natural segment width. */
export interface NaturalClusterAdvance {
  readonly text: string;
  readonly advancePx: number;
  /** Vertical cells only: the planner floored this cell. */
  readonly floored?: boolean;
}


/** Per-unit geometry {@link segAdvanceWidth} needs to apply the
 * {@link condensedSpacingMayFloor} floor: horizontal text supplies its natural
 * grapheme-cluster advances; a vertical run supplies the planned glyph cells
 * (`planVerticalRunWithCapability`), whose advances already carry the run
 * pitch and the floor, so measurement and painted cells share one authority. */
export type CondensedAdvanceUnits =
  | Readonly<{ kind: 'natural-clusters'; clusters: readonly NaturalClusterAdvance[] }>
  | Readonly<{ kind: 'vertical-cells'; cells: readonly NaturalClusterAdvance[] }>;


/** WORD_CONDENSED_SPACING_ADVANCE_FLOOR: §17.3.2.35 removes pitch "after each
 * character" but never moves the pen backwards. Word floors each character's
 * advance at zero, so an over-condensed character stacks on its successor and
 * the run, its line fit, alignment and decorations all use the floored
 * advance. The floored advance is the character's whole advance: scaled glyph,
 * grid delta, run pitch, width balance and punctuation compression (compression
 * never adds advance, so max(0, max(0, glyph + pitch) + compression) is the
 * same value). True only when some character can reach a negative advance;
 * §17.3.2.14 fitText and §17.3.2.10 tate-chu-yoko own their allocation and are
 * outside this rule. */
export function condensedSpacingMayFloor(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
): boolean {
  if (seg.fitTextPerGapPx !== undefined || seg.tateChuYoko) return false;
  const gridPx = segmentCharacterGridDeltaPx(seg, grid, scale);
  return gridPx + charSpacingDeltaPx(seg, scale) < 0 || gridPx < 0;
}


/** Punctuation compression plus width balance of a source range of
 * `seg.text` (px): the per-character adjustments the floor covers together
 * with the glyph and pitch. Built once per segment in linear time and
 * memoised, so repeated prefix probes reuse it (each compression is owned by
 * the unit containing its end offset); a build charges the segment length to
 * `reserveWork`. Vertical planning receives it as `characterAdjustmentPt`. */
export function condensedCharacterAdjustments(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
  reserveWork?: (utf16Units: number) => void,
): (start: number, end: number) => number {
  const cached = condensedAdjustmentCache.get(seg);
  if (cached && cached.text === seg.text && cached.compressions === seg.punctuationCompressions
    && cached.grid === grid && cached.scale === scale) {
    return cached.adjustmentPx;
  }
  reserveWork?.(seg.text.length);
  const compressionBefore = new Float64Array(seg.text.length + 2);
  for (const compression of seg.punctuationCompressions ?? []) {
    if (compression.end >= 0 && compression.end <= seg.text.length) {
      compressionBefore[compression.end + 1]! += compression.adjustmentPt;
    }
  }
  for (let index = 1; index < compressionBefore.length; index += 1) {
    compressionBefore[index]! += compressionBefore[index - 1]!;
  }
  const charScale = charScaleFactor(seg);
  const adjustmentPx = (start: number, end: number): number =>
    (compressionBefore[end + 1]! - compressionBefore[start + 1]!) * scale
    + widthBalanceSpaceAdjustmentForTextPt(seg, seg.text.slice(start, end), grid) * scale * charScale;
  condensedAdjustmentCache.set(seg, {
    text: seg.text,
    compressions: seg.punctuationCompressions,
    grid,
    scale,
    adjustmentPx,
  });
  return adjustmentPx;
}

const condensedAdjustmentCache = new WeakMap<LayoutTextSeg, Readonly<{
  text: string;
  compressions: LayoutTextSeg['punctuationCompressions'];
  grid: DocGridCtx | undefined;
  scale: number;
  adjustmentPx: (start: number, end: number) => number;
}>>();


/** Each unit's whole signed advance (px). Horizontal clusters scale their
 * natural advance and add grid/run pitch plus the character adjustments;
 * vertical cells were planned with pitch, adjustments and floor already.
 * Shared by measurement ({@link segAdvanceWidth}) and retained geometry. */
export function condensedUnitAdvancesPx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
  units: CondensedAdvanceUnits,
): readonly number[] {
  const gridPx = segmentCharacterGridDeltaPx(seg, grid, scale);
  const spacingPx = charSpacingDeltaPx(seg, scale);
  const charScale = charScaleFactor(seg);
  const adjustmentPx = units.kind === 'natural-clusters'
    ? condensedCharacterAdjustments(seg, grid, scale)
    : undefined;
  let offset = 0;
  const list = units.kind === 'natural-clusters' ? units.clusters : units.cells;
  const advances = list.map((unit) => {
    const start = offset;
    offset += unit.text.length;
    return adjustmentPx
      ? unit.advancePx * charScale
        + countedSpacingAdvance(
          countSpacingCharacters(unit.text), gridPx, spacingPx, gridPx + spacingPx,
        )
        + adjustmentPx(start, offset)
      : unit.advancePx;
  });
  if (offset !== seg.text.length) {
    throw new Error('Condensed advance units do not cover their segment');
  }
  return advances;
}


/** Whether {@link condensedUnitAdvancesPx} reaches the floor: a negative
 * horizontal unit, or a vertical cell the planner floored. */
export function condensedUnitsFloor(
  units: CondensedAdvanceUnits,
  advances: readonly number[],
): boolean {
  return units.kind === 'vertical-cells'
    ? units.cells.some((cell) => cell.floored === true)
    : advances.some((advance) => advance < 0);
}


/** The segment's advance as the sum of its floored unit advances, or
 * `undefined` when no unit reaches the floor (callers then keep the signed
 * whole-segment arithmetic unchanged). Summing the non-negative units directly
 * keeps a fully floored run exactly zero. */
function condensedFlooredAdvancePx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
  units: CondensedAdvanceUnits,
): number | undefined {
  const advances = condensedUnitAdvancesPx(seg, grid, scale, units);
  if (!condensedUnitsFloor(units, advances)) return undefined;
  let total = 0;
  for (const advance of advances) total += advance > 0 ? advance : 0;
  return total;
}


/** The single contextual-advance authority for
 * WORD_CONDENSED_SPACING_ADVANCE_FLOOR, used by line fitting and retained
 * geometry alike: each grapheme of `segment.text[start, end)` measured with up
 * to {@link CONDENSED_CONTEXT_GRAPHEMES} preceding graphemes of the same slice,
 * minus that context. Linear in the slice length; every measured UTF-16 unit
 * is charged to `reserveWork`. */
export function contextualGraphemeAdvances(
  segment: LayoutTextSeg,
  start: number,
  end: number,
  fontSizePx: number,
  measureFallback: (text: string) => number,
  reserveWork: (utf16Units: number) => void,
): readonly NaturalClusterAdvance[] {
  const offsets = [...new Set([
    start,
    ...graphemeClusterOffsets(segment.text.slice(start, end)).map((offset) => start + offset),
    end,
  ])].sort((a, b) => a - b);
  // Context windows overlap within a range and are re-acquired by fresh
  // segments of the same run (prefix probes, retained geometry, repeated
  // passes). Each distinct window request is shaped once per service; work is
  // still charged per window measured.
  const windows = segment.textLayoutService
    ? contextualWindowMemo(segment.textLayoutService)
    : undefined;
  const width = (from: number, to: number): number => {
    if (to <= from) return 0;
    reserveWork(to - from);
    if (segment.textLayoutService && segment.textShapeRequest) {
      if (segment.textShapeRequest.text !== segment.text) {
        throw new Error('Cluster measurement does not match its retained text range context');
      }
      const request = {
        ...sliceTextShapeRequest(segment.textShapeRequest, from, to),
        fontSizePt: fontSizePx,
        measure: true,
        clusterGeometry: false,
      };
      const key = contextualWindowKey(request);
      const runKey = key === undefined ? undefined : request.substituteContext!.text;
      const known = key === undefined || runKey === undefined
        ? undefined
        : windows!.byRun.get(runKey)?.get(key);
      if (known !== undefined) return known;
      const advancePt = segment.textLayoutService.shape(request).advancePt;
      if (key !== undefined && runKey !== undefined) rememberContextualWindow(windows!, runKey, key, advancePt);
      return advancePt;
    }
    return measureFallback(segment.text.slice(from, to));
  };
  const clusters: NaturalClusterAdvance[] = [];
  for (let index = 0; index < offsets.length - 1; index += 1) {
    const clusterStart = offsets[index]!;
    const clusterEnd = offsets[index + 1]!;
    const contextStart = offsets[Math.max(0, index - CONDENSED_CONTEXT_GRAPHEMES)]!;
    clusters.push({
      text: segment.text.slice(clusterStart, clusterEnd),
      advancePx: width(contextStart, clusterEnd) - width(contextStart, clusterStart),
    });
  }
  return clusters;
}

/** Shaped widths of context windows, per service: by the window's run
 * context text (shared, not copied), then by {@link contextualWindowKey}. */
interface ContextualWindowMemo {
  readonly byRun: Map<string, Map<string, number>>;
  entries: number;
}

/** Bound of retained window widths per service; reaching it drops them all
 * (a memoisation miss, never a different width). */
const CONTEXTUAL_WINDOW_MEMO_ENTRIES = 1 << 16;

const contextualWindowMemos = new WeakMap<TextLayoutService, ContextualWindowMemo>();

function contextualWindowMemo(service: TextLayoutService): ContextualWindowMemo {
  let memo = contextualWindowMemos.get(service);
  if (!memo) {
    memo = { byRun: new Map(), entries: 0 };
    contextualWindowMemos.set(service, memo);
  }
  return memo;
}

function rememberContextualWindow(
  memo: ContextualWindowMemo,
  runKey: string,
  key: string,
  advancePt: number,
): void {
  if (memo.entries >= CONTEXTUAL_WINDOW_MEMO_ENTRIES) {
    memo.byRun.clear();
    memo.entries = 0;
  }
  let widths = memo.byRun.get(runKey);
  if (!widths) {
    widths = new Map();
    memo.byRun.set(runKey, widths);
  }
  if (!widths.has(key)) memo.entries += 1;
  widths.set(key, advancePt);
}

/** Within its run context, the window request's identity: the fields the
 * service's shape result is keyed by ({@link shapeRequestKeyFields}) and the
 * context offset its run-scope descriptor is read at. Every value is a
 * primitive written exactly (NaN, the infinities and -0 stay distinct), so
 * equal keys are equal shape inputs; anything else is not memoised. */
function contextualWindowKey(request: Readonly<TextShapeRequest>): string | undefined {
  // A malformed request, including one whose fields cannot be read, is left
  // to shape() and its own validation.
  try {
    return contextualWindowKeyOf(request);
  } catch {
    return undefined;
  }
}

function contextualWindowKeyOf(request: Readonly<TextShapeRequest>): string | undefined {
  if (!shapeRequestKeyFieldsArePrimitive(request)
    || typeof request.substituteContext?.text !== 'string') return undefined;
  const parts: string[] = [];
  const leaf = (value: unknown): void => {
    if (value === null) parts.push('z');
    else if (value === undefined) parts.push('u');
    else if (typeof value === 'boolean') parts.push(value ? 't' : 'f');
    else if (typeof value === 'number') parts.push(`n${Object.is(value, -0) ? '-0' : String(value)}`);
    else parts.push(`s${(value as string).length}:${value as string}`);
  };
  leaf(request.substituteContext?.offset);
  // The only arrays are the helper's dense slot tuples of primitives.
  for (const field of shapeRequestKeyFields(request)) {
    if (Array.isArray(field)) {
      parts.push(`a${field.length}`);
      for (let index = 0; index < field.length; index += 1) leaf(field[index]);
    } else {
      leaf(field);
    }
  }
  return parts.join('|');
}


/** A text segment's laid-out advance including the §17.3.2.43 horizontal scale
 *  and §17.3.2.35 character spacing on top of the docGrid delta. Scale natural
 *  glyph width first, then add run pitch per rendered character;
 *  these are independent OOXML properties. Over-condensed characters take
 *  the {@link condensedSpacingMayFloor} zero floor; callers that can reach it
 *  pass the segment's {@link CondensedAdvanceUnits}. Without units the text
 *  is one unit, which is exact for a single cluster. */
export function segAdvanceWidth(
  seg: LayoutTextSeg,
  naturalWidthPx: number,
  grid: DocGridCtx | undefined,
  scale: number,
  units?: CondensedAdvanceUnits,
): number {
  if (seg.fitTextPerGapPx !== undefined) {
    const charCount = [...seg.text].length;
    const gapCount = seg.fitTextRegionEnd ? Math.max(0, charCount - 1) : charCount;
    return naturalWidthPx * charScaleFactor(seg)
      + gapCount * seg.fitTextPerGapPx
      + (seg.fitTextTrailingPadPx ?? 0);
  }
  // ECMA-376 §17.3.2.10 縦中横 (horizontal-in-vertical): the whole run is written
  // horizontally inside ONE cell of the vertical line ("keeping the text on the
  // same line"), so its advance ALONG the column is exactly one em (one cell),
  // independent of the character count and of `w:w` (which stretches the
  // side-by-side glyphs ACROSS the column, not the along-column cell height).
  // A multi-digit tate-chu-yoko run occupies one cell (PDF comparison).
  // (Because the vertical page lays out in a swapped logical frame, this
  // logical-horizontal advance IS the vertical column advance after the page
  // rotation — see vertical-text.ts and renderer's page transform.)
  if (seg.tateChuYoko) return seg.fontSize * scale;
  const segmentDelta = segmentCharacterGridDeltaPx(seg, grid, scale);
  const signed = textAdvanceWidth(
    naturalWidthPx,
    seg.text,
    segmentDelta,
    charScaleFactor(seg),
    charSpacingDeltaPx(seg, scale),
  )
    + widthBalanceSpaceAdjustmentTotalPt(seg, grid) * scale * charScaleFactor(seg)
    + punctuationCompressionTotalPt(seg) * scale;
  if (!condensedSpacingMayFloor(seg, grid, scale)) return signed;
  if (!units) return signed > 0 ? signed : 0;
  return condensedFlooredAdvancePx(seg, grid, scale, units) ?? signed;
}


export type SnapToCharsClass = 'eastAsia' | 'latin' | 'complexScript';


/** Normative §17.6.5 script class. Source/run/style boundaries are deliberately
 * absent: ascii and hAnsi form one contiguous Latin block, complex-script text
 * forms its own block, and each East-Asian character owns one cell. */
export function snapToCharsClass(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
): SnapToCharsClass | null {
  if (grid?.type !== 'snapToChars'
    || !grid.characterPitchPt
    || grid.characterPitchPt <= 0
    || seg.snapToCharacterGrid === false
    || seg.metricOnly
    || seg.fitTextRegionIndex !== undefined
    || seg.tateChuYoko) return null;
  if (seg.script === 'eastAsia') return 'eastAsia';
  if (seg.script === 'complexScript') return 'complexScript';
  return 'latin';
}


export function snapToCharsAllocatedWidthPx(
  naturalWidthPx: number,
  kind: SnapToCharsClass,
  pitchPx: number,
  eastAsianCellCount = 1,
): number {
  if (!(pitchPx > 0) || !Number.isFinite(naturalWidthPx)) return naturalWidthPx;
  if (kind === 'eastAsia') return Math.max(1, eastAsianCellCount) * pitchPx;
  return Math.max(1, Math.ceil(Math.max(0, naturalWidthPx) / pitchPx - 1e-9)) * pitchPx;
}


export function isGridLineRule(ctx: DocGridCtx | undefined): boolean {
  if (!ctx || !ctx.linePitchPt || ctx.linePitchPt <= 0) return false;
  return ctx.type === 'lines'
    || ctx.type === 'linesAndChars'
    || ctx.type === 'snapToChars';
}
