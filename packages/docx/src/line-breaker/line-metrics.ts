import type { LineSpacing } from '../types';
import type { CanvasFontRoute, ResolvedFontMetric } from '@silurus/ooxml-core';
import { measureResolvedCanvasFontBoxRatio, SYNTHETIC_ASCENT_EM, SYNTHETIC_DESCENT_EM } from '@silurus/ooxml-core';
import type { NumberingMarkerShapeInput } from '../layout/types.js';
import type { MeasurementTextContext } from '../layout/measurement-capabilities.js';
import type { ParagraphLayoutSource, TextLayoutService } from '../layout/text.js';
import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';
import { WORD_EMPTY_MARK_ASCII_FACE, wordDegenerateLineSpacingIsSingle, wordEastAsianGridLineCells, wordFarEastSingleLinePx, wordGridAtLeastLineHeightPx, wordUseFeLayoutInheritedGridHeightPx, wordUseFeLayoutParagraphMarkGridAdvancePx, wordInlinePictureAutoLineHeightPx, wordLatinDesignGridSingleHeight } from '../layout/line-compatibility.js';
import { type DocGridCtx, type LayoutLine } from './model.js';
import { isGridLineRule } from './advance.js';
import { buildFont, getDefaultFontFamily, getDefaultFontSize } from './font-routes.js';
import { allowsReferenceFontMetric, authoredReferenceVerticalMetric, indexedFontMetrics, mayUseAuthoredReferenceVerticalMetric, selectResourceMetric } from './font-metrics.js';

/**
 * ECMA-376 §17.6.5 docGrid line grid — number of whole grid CELLS a
 * single-spaced East Asian line occupies on a pitch of `pitchPx`, from the
 * line's SINGLE-LINE HEIGHT `naturalPx` (admitted design geometry, or the
 * documented generic East Asian fallback when resource geometry is absent).
 * The count is `ceil(naturalPx / pitchPx)` — the smallest number of whole
 * cells that CONTAINS the line.
 *
 * `word-east-asian-grid-line-allocation` records the compatibility formula:
 * ceil(design-line-height / pitch), independent of horizontal or vertical text
 * direction. The focused grid-allocation tests retain the adjudicated boundary
 * matrix; this production comment states only the resulting invariant.
 *
 * A line that fills k pitches exactly occupies k cells (ceil; no measured
 * point sits on the boundary — the geometric reading is that it still FITS).
 * For a mixed-size line, callers supply the tallest run's resolved height
 * (§17.3.1.33 tallest-run line box). ECMA-376 defines `linePitch` as one
 * single-spaced line; taller-line spreading is governed by
 * `word-east-asian-grid-line-allocation`. Returns at least 1 for every finite
 * `naturalPx >= 0`.
 */
export function docGridLineCells(naturalPx: number, pitchPx: number): number {
  return wordEastAsianGridLineCells(naturalPx, pitchPx);
}


/** Deterministic single-line height used to count docGrid cells for one East
 * Asian text run. Resolved font resources contribute their parsed design height.
 *
 * The `word-east-asian-grid-line-allocation` rule supplies the 1.3 × usWin-box
 * fallback measured for the Far East grid path; §17.6.5 does not define this
 * factor.
 *
 * When the font resource is unavailable, its usWin box is unknown, so the
 * compatibility fallback uses 1.3em (WORD_FAR_EAST_SINGLE_LINE_FACTOR).
 * This assumes an unknown usWin box of 1em and can choose the wrong cell near
 * a grid boundary. It is an explicit unavailable-resource fallback, not a
 * normative font-metrics claim or an error bound. Issue #1525 tracks its
 * replacement when the actual face geometry or a justified substitute exists.
 *
 * Never use a substituted Canvas box here: its integer-rounded metrics are
 * font- and scale-dependent. */
export function eastAsianGridCountSinglePx(intendedSinglePx: number, emPx: number): number {
  return wordFarEastSingleLinePx(intendedSinglePx, emPx);
}


/**
 * Compute the total line-box height in px from a line's natural font metrics
 * (fontBoundingBoxAscent + fontBoundingBoxDescent) per ECMA-376 §17.3.1.33.
 *
 *   auto    → natural × value ("single" = 1 natural line, "double" = 2).
 *             When the docGrid line axis is active, the
 *             multiplier applies against the grid pitch instead, with a
 *             floor of the natural line height.
 *   exact   → value in pt, converted to px (ignores font and grid).
 *   atLeast → max(natural, authored minimum, active grid minimum).
 *             `word-grid-at-least-tall-line-unsnapped` owns the explicit
 *             tall-line compatibility branch.
 *   null    → natural, or grid pitch if the section defines one.
 *
 * Exported for unit tests only — not part of the package API (not
 * re-exported from index.ts).
 */
export function lineBoxHeight(
  ls: LineSpacing | null,
  ascentPx: number,
  descentPx: number,
  scale: number,
  grid?: DocGridCtx,
  hasRuby?: boolean,
  intendedSinglePx = 0,
  eastAsian = false,
  // px — the line's DESIGN grid-count height: the max over segments of each
  // run's format-policy single-line height (a resolved resource's design
  // height, or the generic East Asian fallback). Used ONLY to count
  // docGrid cells for East Asian lines, so a substituted face's Canvas box
  // cannot change pagination or paint-scale cell allocation.
  gridCountSinglePx?: number,
  // px — unresolved East Asian run em used only by direct/synthetic callers that
  // cannot provide the producer-computed per-line gridCountSinglePx.
  untabledEastAsianEmPx?: number,
  uniformPositionAuto?: LayoutLine['uniformPositionAuto'],
  inlinePictureTextSinglePx = 0,
  latinGridCountSinglePx = 0,
): number {
  const glyphNatural = ascentPx + descentPx;
  // For `auto`/single spacing the multiplier applies to the intended font's
  // design line height (ECMA-376 §17.3.1.33). When the document's font is
  // substituted, the Canvas glyph extent (`glyphNatural`) can understate that.
  // An admitted resource or reference profile supplies the intended height.
  // This takes the maximum of the geometry passed to the line allocator; it
  // does not prove that every painted glyph fits a different Canvas face.
  // Grid-snapped lines are governed by the grid pitch instead.
  const natural = Math.max(glyphNatural, intendedSinglePx);
  const hasGrid = isGridLineRule(grid);
  const pitchPx = hasGrid ? grid!.linePitchPt! * scale : 0;
  // Per ECMA-376 §17.6.5, a paragraph whose `line` attribute is NOT
  // explicitly set — it only inherits from docDefault — snaps to one grid
  // pitch per text line in docGrid sections, regardless of the inherited
  // multiplier. Paragraphs that do set `line` on their pPr or a named style
  // multiply against the pitch as usual.
  //
  // A single-spaced line on a docGrid snaps to whole grid CELLS in East Asian
  // text. The number of cells is derived from the line's DESIGN single-line
  // height (`gridCountSinglePx`), per
  // `word-east-asian-grid-line-allocation`; the substituted Canvas glyph box is
  // not used because it can overstate the source resource's design height.
  // Known non-FE Latin design profiles use WORD_LATIN_DESIGN_GRID_CELLS;
  // unclassified/FE profiles retain their natural height above a one-cell floor. ECMA-376 Part 1 defines only the natural ≤ pitch case
  // (§17.6.5 / §17.3.1.32), so `word-east-asian-grid-line-allocation` gates
  // whole-cell allocation on the line's script.
  const gridSingleCell = (): number => {
    if (!eastAsian) {
      // WORD_LATIN_DESIGN_GRID_CELLS admits design metrics only; unknown
      // Canvas boxes and Far-East reference faces keep the existing floor.
      return wordLatinDesignGridSingleHeight(natural, pitchPx, latinGridCountSinglePx);
    }
    // `word-east-asian-grid-line-allocation`: count cells from the source face's
    // design single-line height, not a substituted Canvas glyph box. Prefer the
    // per-line design-grid height; direct unresolved callers may supply the run em.
    // A legacy caller with neither input gets one pitch.
    const cellCountHeight = gridCountSinglePx
      ?? (intendedSinglePx > 0
        ? intendedSinglePx
        : untabledEastAsianEmPx === undefined
          ? pitchPx
          : eastAsianGridCountSinglePx(0, untabledEastAsianEmPx));
    // ECMA-376 §17.3.3.25 places the ruby guide above the normal base text;
    // §17.6.5 defines the single-spaced grid pitch. Within the compatibility
    // whole-cell allocation, reserve the measured base + rt extent without
    // dropping the admitted design minimum: ruby may add cells, never remove
    // cells that the same line without its annotation would occupy.
    const requiredHeight = hasRuby ? Math.max(glyphNatural, cellCountHeight) : cellCountHeight;
    return docGridLineCells(requiredHeight, pitchPx) * pitchPx;
  };
  const inheritedOnly = ls !== null && ls.explicit !== true;
  if (!ls) {
    // No explicit spacing → single line. Use the intended single-line height
    // (`natural`) off-grid; on-grid, snap per gridSingleCell.
    return hasGrid ? gridSingleCell() : natural;
  }
  // A zero/negative `w:line` is degenerate input whose behavior ECMA-376
  // §17.3.1.33 does not define (read literally, an `exact` line of 0 would
  // collapse the line box to no height; some generators emit
  // `<w:spacing w:line="0" w:lineRule="exact"/>` on table cells).
  // `word-degenerate-line-spacing-single` follows the native LSPD
  // representation:
  // "exact" spacing is encoded as a negative dyaLine ("the line spacing, in
  // twips, is exactly 0x10000 minus dyaLine", so an exact 0 is unrepresentable)
  // and a non-negative dyaLine in twips mode is "dyaLine or the number of twips
  // necessary for single spacing, whichever value is greater" — i.e. a stored 0
  // resolves to exactly single spacing. `word-degenerate-line-spacing-single`
  // applies that non-collapsing interpretation to exact/auto values <= 0.
  if (wordDegenerateLineSpacingIsSingle(ls.rule, ls.value)) {
    return hasGrid ? gridSingleCell() : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      if (inheritedOnly) {
        const allocated = gridSingleCell();
        return eastAsian
          ? wordUseFeLayoutInheritedGridHeightPx(allocated, pitchPx, ls.value)
          : allocated;
      }
      return ls.value === 1 ? gridSingleCell() : Math.max(natural, pitchPx * ls.value);
    }
    if (inlinePictureTextSinglePx > 0 && ls.value >= 1) {
      // The object owns its baseline-union extent; the authored auto leading
      // belongs to the selected text/paragraph-mark face, not to the object.
      return wordInlinePictureAutoLineHeightPx(
        natural, inlinePictureTextSinglePx, ls.value,
      );
    }
    if (uniformPositionAuto) {
      // ECMA-376 §17.3.2.24 defines the signed run baseline position, and
      // §17.3.1.33 defines the auto multiplier, but neither defines their
      // joint line allocation. Word for Mac 16.112.4 controls (Arial, Calibri,
      // Times New Roman; 8/10/16 pt; -6..+12 pt; 1.0..2.0 auto multiples)
      // show that a uniformly lowered visible line keeps its normal pitch,
      // while a raise consumes hhea descent before adding to that pitch. The
      // residual displacement is NOT multiplied. Mixed-position lines and
      // exact/atLeast spacing are counterexamples and retain their own paths.
      return uniformPositionAuto.normalSinglePx * ls.value
        + Math.max(0, uniformPositionAuto.positionPx - uniformPositionAuto.designDescentPx);
    }
    return natural * ls.value;
  }
  if (ls.rule === 'exact') return ls.value * scale;
  if (ls.rule === 'atLeast') {
    // §17.18.48 establishes the authored minimum and §17.6.5 establishes the
    // grid pitch, but neither clause specifies how a tall, plain line with an
    // explicit atLeast value combines with whole-cell grid allocation.
    // `word-grid-at-least-tall-line-unsnapped` preserves the raw content height;
    // ruby and inherited-only spacing retain their established whole-cell path.
    const gridMinimum = hasGrid
      ? (hasRuby || inheritedOnly ? gridSingleCell() : pitchPx)
      : 0;
    return wordGridAtLeastLineHeightPx(natural, ls.value * scale, gridMinimum);
  }
  return natural;
}


/** Natural single-line height in px for an empty paragraph (no rendered text). */
export function emptyLineNaturalPx(fontSizePt: number, scale: number): { asc: number; desc: number } {
  return { asc: fontSizePt * scale * SYNTHETIC_ASCENT_EM, desc: fontSizePt * scale * SYNTHETIC_DESCENT_EM };
}


export function measuredLineMetrics(
  m: TextMetrics,
  fallbackEmPx: number,
): { ascent: number; descent: number } {
  return {
    ascent: m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fallbackEmPx * SYNTHETIC_ASCENT_EM,
    descent: m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fallbackEmPx * SYNTHETIC_DESCENT_EM,
  };
}


// Canvas fontBoundingBox sides are rounded to device pixels at the run's
// ordinary size. For an unprofiled native face this can lose a fraction of a
// point on every §17.3.1.33 automatic line, which accumulates across table
// rows. Probe the same browser-selected glyph at two larger sizes and use the
// high-size ratio only when both agree within their two-side quantization
// bounds (2/200 + 2/1000 em). This is measured Canvas geometry, not a claim
// about the installed OpenType tables or Word's metric selection. The shared
// core probe compares the glyph against a missing-family control and declines
// aliases that cannot be distinguished from fallback. Parsed resources and
// catalogued native references take precedence over this approximation.
export const NATIVE_BOX_PROBE_SMALL_EM = 200;

export const NATIVE_BOX_PROBE_LARGE_EM = 1000;

export const nativeBoxRatios = new WeakMap<MeasurementTextContext, WeakMap<object, Map<string, number | null>>>();


export function nativeCanvasLineRatio(
  context: MeasurementTextContext,
  cacheOwner: object,
  route: CanvasFontRoute | undefined,
  family: string | null | undefined,
  weight: number,
  style: 'normal' | 'italic',
  text: string,
): number | null {
  if (route?.scope !== 'native' || !family) return null;
  let probe = '';
  for (const scalar of text) {
    if (!/\s/u.test(scalar)) { probe = scalar; break; }
  }
  if (!probe) return null;
  let ownerCaches = nativeBoxRatios.get(context);
  if (!ownerCaches) {
    ownerCaches = new WeakMap();
    nativeBoxRatios.set(context, ownerCaches);
  }
  let cache = ownerCaches.get(cacheOwner);
  if (!cache) {
    cache = new Map();
    ownerCaches.set(cacheOwner, cache);
  }
  const key = JSON.stringify([route.fingerprint, family, weight, style, probe]);
  if (cache.has(key)) return cache.get(key) ?? null;
  const small = measureResolvedCanvasFontBoxRatio(context, family, {
    text: probe, emPx: NATIVE_BOX_PROBE_SMALL_EM, weight, style,
  });
  const large = small === null ? null : measureResolvedCanvasFontBoxRatio(context, family, {
    text: probe, emPx: NATIVE_BOX_PROBE_LARGE_EM, weight, style,
  });
  const ratio = small !== null && large !== null
    && Math.abs(small - large) <= 2 / NATIVE_BOX_PROBE_SMALL_EM + 2 / NATIVE_BOX_PROBE_LARGE_EM
    ? large : null;
  // The owner is document-scoped in production, but one document can contain
  // arbitrary authored family names and characters. Bound retained tuples.
  if (cache.size >= 128) cache.clear();
  cache.set(key, ratio);
  return ratio;
}


/**
 * Height (px) of the paragraph-mark line box for a paragraph that places no
 * inline content on any line. Per ECMA-376 §17.3.1.29 the paragraph mark always
 * produces one line box even when the paragraph has no inline runs; floating
 * objects (§20.4.2.x `wp:anchor`) are removed from the inline flow but never
 * suppress that paragraph-mark line. This is the height used both by the
 * literal empty-paragraph path and by paragraphs whose only segments are
 * wrap-float anchors (which `layoutLines` skips, yielding zero lines).
 * `effectiveLineSpacing` lets resolved paragraph context override the source
 * value; omitting it preserves the existing `para.lineSpacing` behavior.
 */
/** The natural ascent/descent (px) and the resolved line-box advance (px) of an
 *  empty paragraph's mark line. Shared by {@link paragraphMarkLineHeight} (which
 *  returns only the advance) and {@link paragraphMarkBelowBaselinePt} (which needs
 *  the ascent/descent to locate the mark baseline within the box). */
export interface MarkLineMetrics {
  readonly advancePx: number;
  /** Grid-only leading below the centered own (natural/authored) box. */
  readonly gridAfterLeadingPx: number;
  readonly ascentPx: number;
  readonly descentPx: number;
  /** The mark face's intended single-line height. */
  readonly intendedSinglePx: number;
}


export function paragraphMarkLineMetrics(
  para: ParagraphLayoutSource,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  _resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
  // px — a picture bullet on an empty list item stands on the mark line's
  // baseline: its height is a minimum ascent of the mark line.
  minimumAscentPx = 0,
): MarkLineMetrics {
  const effectiveMarkShapeInput = markShapeInput;
  // WORD_EMPTY_MARK_ASCII_FACE: an empty paragraph mark is measured with its
  // ascii face whatever its hint, language, eastAsia face or the document's
  // scripts. East-Asian docGrid cell classification (`eastAsian`) stays
  // content/document based and is unaffected.
  void WORD_EMPTY_MARK_ASCII_FACE;
  const forceCs = effectiveMarkShapeInput?.complexScript === true;
  const fs = effectiveMarkShapeInput?.fontSizePt ?? getDefaultFontSize(para);
  const authoredFamily = getDefaultFontFamily(para, false);
  const markWeight = effectiveMarkShapeInput?.weight ?? 400;
  const markStyle = effectiveMarkShapeInput?.style ?? 'normal';
  // §17.3.1.29 stores the paragraph mark's own run properties in pPr/rPr.
  // Shape this probe in its own context: the library's scoped Arabic substitute
  // policy must not borrow script proof from surrounding body runs.
  const markProbe = 'x';
  // A supplied metric map without the selecting text service cannot prove
  // which face Canvas paints. The compatibility argument above is ignored;
  // only the selected resource tuple below can lend authoritative geometry.
  let resolvedLocalFont: ResolvedFontMetric | undefined;
  // ECMA-376 §17.3.1.33: atLeast still reserves the normal single-line box
  // when its authored minimum is smaller. Reusing the automatic-line
  // OpenType projection for that natural box is a format-policy inference;
  // exact spacing bypasses the box.
  const naturalMetricAllowed = effectiveLineSpacing?.rule !== 'exact';
  const measuredFamily = authoredFamily;
  let asc: number;
  let desc: number;
  let referenceMarkMetric: ReturnType<typeof referenceFontLineMetrics> = undefined;
  let nativeMarkRatio: number | null = null;
  if (textLayoutService) {
    const bold = markWeight >= 600;
    const italic = markStyle === 'italic';
    const ascii = effectiveMarkShapeInput?.fonts.ascii ?? para.defaultFontFamily ?? authoredFamily;
    const shaped = textLayoutService.shape({
      text: markProbe,
      fontSizePt: fs * scale,
      fonts: effectiveMarkShapeInput?.fonts ?? {
        ascii,
        highAnsi: ascii,
        eastAsia: para.defaultFontFamilyEastAsia ?? ascii,
        complexScript: ascii,
      },
      themeFonts: effectiveMarkShapeInput?.themeFonts,
      themeFontPresence: effectiveMarkShapeInput?.themeFontPresence,
      weight: bold ? 700 : 400,
      style: italic ? 'italic' : 'normal',
      complexScript: forceCs,
      fontHint: effectiveMarkShapeInput?.fontHint,
      eastAsiaLanguage: effectiveMarkShapeInput?.eastAsiaLanguage,
      kerning: effectiveMarkShapeInput?.kerning,
      measure: true,
    });
    const selectedFont = shaped.spans[0]?.font;
    // The paragraph mark has no visible run, but Canvas selects its probe
    // through the same CSS face tuple as text. Resolve against the complete
    // resource set: an authored alias cannot disambiguate overlapping faces.
    // The probe selects the mark's face but is never painted, so painted-glyph
    // ownership does not apply (see metricCoversText).
    resolvedLocalFont = selectedFont
      ? selectResourceMetric(indexedFontMetrics(
          textLayoutService.fontMetrics ?? textLayoutService.localMetrics,
        ), selectedFont, markProbe, false)
      : undefined;
    if (!naturalMetricAllowed && resolvedLocalFont?.designAscentRatio != null) {
      resolvedLocalFont = undefined;
    }
    referenceMarkMetric = naturalMetricAllowed
      && !resolvedLocalFont && mayUseAuthoredReferenceVerticalMetric(selectedFont)
      && allowsReferenceFontMetric(indexedFontMetrics(
        textLayoutService.fontMetrics ?? textLayoutService.localMetrics ?? {},
      ), selectedFont)
      ? authoredReferenceVerticalMetric(selectedFont, { text: textLayoutService, probe: markProbe })
      : undefined;
    if (naturalMetricAllowed && !resolvedLocalFont && !referenceMarkMetric
      && selectedFont?.source === 'native' && ctx) {
      nativeMarkRatio = nativeCanvasLineRatio(
        ctx, fontFamilyClasses, selectedFont.route, selectedFont.resolvedFamily,
        selectedFont.weight, selectedFont.style, markProbe,
      );
    }
    const markMeasured = {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    ({ ascent: asc, descent: desc } = measuredLineMetrics(markMeasured, fs * scale));
  } else if (ctx) {
    // ECMA-376 §17.3.1.29 / §17.3.1.33: an empty paragraph's mark line reserves
    // the mark font's REAL single-line height — the SAME fontBoundingBox a text
    // line of that font and size uses (layoutLines), so an empty paragraph is
    // exactly as tall as a one-character paragraph of the same run properties.
    // The synthetic 0.8/0.2 ≈ 1em box under-measured every empty paragraph
    // whenever the (often substituted) font's real box exceeds 1em — a Latin
    // fallback reports ~1.15em — so a run of empty "spacer" paragraphs fell
    // short and following content rose into a preceding float's wrap band
    // instead of clearing the float. East Asian documents probe an EA glyph so
    // docGrid cell rounding (lineBoxHeight) reserves whole cells (a 20pt mark
    // on a 20pt pitch occupies two cells); others probe a Latin glyph.
    // fontBoundingBox is reported per
    // resolved face (not per glyph), so the probe choice does not change the box
    // for a face that contains it — and the probe is script-matched, so the mark
    // font does. A parsed resource metric, when available, is applied below by
    // the same path used for visible text.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, measuredFamily, fontFamilyClasses);
    const m = ctx.measureText(markProbe);
    ctx.font = prevFont;
    // A mark line carries no smallCaps/vertAlign, so fallback == correction size.
    ({ ascent: asc, descent: desc } = measuredLineMetrics(m, fs * scale));
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  const designSides = resolvedLocalFont?.designAscentRatio != null
    && resolvedLocalFont?.designDescentRatio != null
    ? resolvedLocalFont : referenceMarkMetric;
  if (designSides) {
    asc = designSides.designAscentRatio! * fs * scale;
    desc = designSides.designDescentRatio! * fs * scale;
  }
  // WORD_PICTURE_BULLET_LINE_PARTICIPATION: an empty list item's picture
  // bullet makes the mark line max(mark ascent, picture height) high above
  // the baseline, and that whole box then follows the paragraph's line rule.
  // Exact spacing ignores the picture (the caller passes no minimum).
  asc = Math.max(asc, minimumAscentPx);
  const resourceRatio = resolvedLocalFont?.lineHeightRatio;
  const designIntendedSingle = referenceMarkMetric
    ? referenceMarkMetric.lineHeightRatio * fs * scale
    : resourceRatio != null
      ? resourceRatio * fs * scale
      : 0;
  // A native Canvas box may set the ordinary automatic-line floor, but it
  // cannot establish the Word Far-East design height used to count grid cells.
  // Keep an untabled mark on the same 1.3em grid fallback as visible text.
  const intendedSingle = Math.max(designIntendedSingle, (nativeMarkRatio ?? 0) * fs * scale);
  const gridCountSingle = eastAsian
    ? Math.max(
        eastAsianGridCountSinglePx(designIntendedSingle, fs * scale),
        minimumAscentPx > 0 ? asc + desc : 0,
      )
    : undefined;
  const ordinaryAdvancePx = lineBoxHeight(
    effectiveLineSpacing,
    asc,
    desc,
    scale,
    grid,
    paraHasRuby,
    intendedSingle,
    eastAsian,
    gridCountSingle,
  );
  const gridAllocationActive = useFeLayout && eastAsian && isGridLineRule(grid);
  const allocatedGridAdvancePx = gridAllocationActive
    ? lineBoxHeight(
        null,
        asc,
        desc,
        scale,
        grid,
        paraHasRuby,
        intendedSingle,
        eastAsian,
        gridCountSingle,
      )
    : ordinaryAdvancePx;
  // Candidate for the observed atLeast=0 compatibility branch. Compute it
  // independently of the source's inheritance flag so the compatibility owner
  // can select it without leaking the signed boundary into this producer.
  const atLeastZeroAdvancePx = gridAllocationActive
    ? lineBoxHeight(
        { rule: 'atLeast', value: 0, explicit: true },
        asc,
        desc,
        scale,
        grid,
        paraHasRuby,
        intendedSingle,
        eastAsian,
        gridCountSingle,
      )
    : ordinaryAdvancePx;
  const advancePx = useFeLayout
    // `word-use-fe-layout-empty-mark-grid-allocation`
    ? wordUseFeLayoutParagraphMarkGridAdvancePx({
        ordinaryAdvancePx,
        allocatedGridAdvancePx,
        atLeastZeroAdvancePx,
        lineSpacing: effectiveLineSpacing,
        gridAllocationActive,
        scale,
      })
    : ordinaryAdvancePx;
  // WORD_GRID_LINE_CONTENT_FIT_ADMISSION also covers an empty mark. Resolve
  // its own box from the same selected font and spacing, without re-shaping.
  // §17.6.5 grid allocation remains the full retained advance; authored exact,
  // atLeast and automatic multiples never become an admission allowance.
  const ownAdvancePx = lineBoxHeight(
    effectiveLineSpacing, asc, desc, scale, undefined, paraHasRuby,
    intendedSingle, eastAsian, gridCountSingle,
  );
  const gridAfterLeadingPx = isGridLineRule(grid)
    ? Math.max(0, Math.min((advancePx - ownAdvancePx) / 2,
        lineBelowBaselinePx(advancePx, asc, desc) - desc))
    : 0;
  return { advancePx, gridAfterLeadingPx, ascentPx: asc, descentPx: desc, intendedSinglePx: intendedSingle };
}


export function paragraphMarkLineHeight(
  para: ParagraphLayoutSource,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
): number {
  return paragraphMarkLineMetrics(
    para, scale, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts, textLayoutService, markShapeInput, useFeLayout,
  ).advancePx;
}


/**
 * §17.3.1.29 / §17.3.1.33 — the extent (px) of a line that sits BELOW its
 * baseline (descent + half of any auto/atLeast leading), using the HALF-LEADING
 * (centred) baseline `top + (advance − (ascent + descent)) / 2 + ascent`, so the
 * portion below it is `(advance − ascent + descent) / 2`.
 *
 * Called for BOTH a paragraph's last visible line (paragraph-measure.ts, the
 * `lastLineBelowBaselinePt` field) and — via {@link paragraphMarkBelowBaselinePt}
 * — an empty paragraph's mark line. Its ONE consumer (renderer.ts
 * `trailingMarkOverflow`) reads it only for an inkless trailing MARK: the
 * whitespace such a paragraph may let overflow the bottom content edge under
 * `word-trailing-empty-mark-baseline-admission` (#981).
 *
 * NOTE: this stays the CENTRED baseline even though VISIBLE lineRule=auto content
 * lines now use a PINNED baseline (#990: multiplier leading placed entirely
 * below the glyphs). The pagination consumer is inkless
 * (mark-only), so the pinned glyph baseline never reaches it, and the #981 page
 * fit is pinned by that admission rule — changing it would move page boundaries.
 * `word-auto-multiple-baseline-pin` is therefore intentionally DRAW-ONLY.
 */
export function lineBelowBaselinePx(advancePx: number, ascentPx: number, descentPx: number): number {
  return Math.max(0, (advancePx - ascentPx + descentPx) / 2);
}


export function paragraphMarkBelowBaselinePt(
  para: ParagraphLayoutSource,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian: boolean,
  ctx: MeasurementTextContext | undefined,
  fontFamilyClasses: Record<string, string>,
  effectiveLineSpacing: LineSpacing | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
): number {
  // Measured at scale 1 so the returned px value is already in points.
  const m = paragraphMarkLineMetrics(
    para, 1, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts, textLayoutService, markShapeInput, useFeLayout,
  );
  return lineBelowBaselinePx(m.advancePx, m.ascentPx, m.descentPx);
}
