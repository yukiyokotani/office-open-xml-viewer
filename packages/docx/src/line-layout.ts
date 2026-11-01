import type { TabStop } from './types';
import type { KinsokuRules } from '@silurus/ooxml-core';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { prepareFloatWrap, type PreparedFloatWrap } from './float-layout.js';
import { cloneSegmentsForLinePass, convergeLineWrap } from './layout/line-wrap-convergence.js';
import type { MeasurementTextContext, VerticalGlyphMeasurementService } from './layout/measurement-capabilities.js';
import { type DocGridCtx, type LayoutLine, type LayoutSeg, type LineBoundary, type WrapLayoutCtx } from './line-breaker/model.js';
import { DEFAULT_TAB_PT } from './line-breaker/tabs.js';
import { runLineBreakerPass, type LineFitWidth } from './line-breaker/pass-driver.js';
export { type LineBoundary, type LayoutTextSeg, type LayoutTabSeg, type LayoutImageSeg, type LayoutMathSeg, type LayoutLineBreak, type LayoutSeg, type LayoutLine, type WrapLayoutCtx, type DocGridCtx, type LineLayoutEnvironment } from './line-breaker/model.js';
export { gridCharDeltaPx, eaGlyphCount, gridSegDeltaPx, segmentCharacterGridDeltaPx, charSpacingDeltaPx, condensedSpacingMayFloor, condensedClusterFloorPx, type NaturalClusterAdvance, type CondensedAdvanceUnits, effectiveCharacterSpacingPt, punctuationCompressionTotalPt, widthBalanceSpaceAdjustmentTotalPt, widthBalanceSpaceAdjustmentForTextPt, slicedPunctuationCompressions, charScaleFactor, segLetterSpacingPx, segAdvanceWidth, type SnapToCharsClass, snapToCharsClass, snapToCharsAllocatedWidthPx, isGridLineRule } from './line-breaker/advance.js';
export { rubyAscentReservePx } from './line-breaker/ruby-metrics.js';
export { ARABIC_SUBSTITUTE_FONTS, NASKH_SERIF_ARABIC_FONTS, isArabicSubstituteFont, quoteAll, ARABIC_TAIL_SANS, sansTail, serifTail, fontFamilyNormalizeCache, fontFamilyPitchesByClasses, fontClassesWithPitches, normalizeFontFamily, normalizeFontFamilyUncached, buildFont, segmentIntendedSingleLinePx, segmentEastAsiaFloorSingleLinePx, getDefaultFontSize, getDefaultFontFamily } from './line-breaker/font-routes.js';
export { docGridLineCells, eastAsianGridCountSinglePx, lineBoxHeight, emptyLineNaturalPx, type MarkLineMetrics, paragraphMarkLineMetrics, paragraphMarkLineHeight, lineBelowBaselinePx, paragraphMarkBelowBaselinePt } from './line-breaker/line-metrics.js';
export { splitSmallCapsCase, type NoteNumbering, formatNoteNumber, findNearbyFontSize, resolveFieldText, hasCJKBreakOpportunity, isRtlBidiLang, splitByComplexScript, splitByEastAsia, splitDigitGroups, splitTextForLayout } from './line-breaker/text-runs.js';
export { fitCJKPrefix } from './line-breaker/fit-search.js';
export { DEFAULT_TAB_PT, resolveDefaultTabPt, type BidiTabItem, type BidiTabResult, layoutBidiTabStops } from './line-breaker/tabs.js';
export { buildSegments } from './line-breaker/segment-builder.js';

export function layoutLines(
  ctx: MeasurementTextContext,
  segs: LayoutSeg[],
  maxWidth: number,
  firstIndent: number,
  scale: number,
  tabStops?: TabStop[],
  wrapCtx?: WrapLayoutCtx,
  fontFamilyClasses?: Record<string, string>,
  tabOriginPx?: number,
  kinsoku?: KinsokuRules,
  characterGrid?: DocGridCtx,
  defaultTabPt?: number,
  marginRightPx?: number,
  baseRtl?: boolean,
  isJustified?: boolean,
  stretchLastLine?: boolean,
  startBoundary?: LineBoundary,
  widthPolicy?: 'bounded' | 'intrinsic' | 'unwrapped',
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  overflowPunct?: boolean,
  justifiedCompression?: boolean,
  lineFit?: LineFitWidth,
): LayoutLine[];

export function layoutLines(
  ctx: MeasurementTextContext,
  segs: LayoutSeg[],
  maxWidth: number,
  firstIndent: number,
  scale: number,
  tabStops: TabStop[] = [],
  wrapCtx?: WrapLayoutCtx,
  fontFamilyClasses: Record<string, string> = {},
  // Paragraph left-indent in px. Tab-stop positions are measured from the text
  // margin (ECMA-376 §17.3.1.37), but layout is paraX-relative, so subtract this.
  tabOriginPx: number = 0,
  // ECMA-376 §17.15.1.58–.60 Japanese line-breaking rules. Default kinsoku is
  // ON; the CJK overflow path retracts the break to a kinsoku-legal position.
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
  // ECMA-376 §17.6.5 docGrid CHARACTER grid. The grid kind and pitch travel
  // together so measure, line breaking, and retained paint cannot disagree on
  // whether the delta applies to all characters or East Asian characters only.
  characterGrid: DocGridCtx | undefined = undefined,
  // ECMA-376 §17.15.1.25 — automatic tab-stop interval (pt). The automatic-stop
  // grid (`nextTabStop`) multiplies this by `scale`; defaults to the spec absent
  // value (720 twips = 36pt) for callers without document settings.
  defaultTabPt: number = DEFAULT_TAB_PT,
  // ECMA-376 §17.3.3.23 — paraX-relative X (px) of the TEXT-MARGIN right edge,
  // used only to resolve a `<w:ptab w:relativeTo="margin">`. Equals
  // `maxWidth + indentRightPx`; defaults to `maxWidth` (correct when the
  // paragraph has no right indent — the common footer case). The margin LEFT
  // edge is `-tabOriginPx`. `relativeTo="indent"` uses the content box
  // (`[0, maxWidth]`) and needs neither.
  marginRightPx: number = maxWidth,
  // ECMA-376 §17.3.1.6 `<w:bidi>` — the paragraph's base direction is RTL. Tab
  // stops mirror to the leading (right) edge in this case (§17.18.84 start/end
  // are logical edges): the tab widths are computed in the VISUAL frame by a
  // per-line post-pass (`layoutBidiTabStops`) instead of the LTR pen math, and
  // tabs do not trigger the LTR right/center/overflow wrap paths. Default false
  // ⇒ the LTR tab paths run unchanged (byte-identical output).
  baseRtl = false,
  // ECMA-376 §17.18.44 paragraph classification. The fit budget is gated per
  // prospective line with the same predicate the paint pass uses.
  isJustified = false,
  // `distribute`/`thaiDistribute` stretch the logical last line; `both` and
  // kashida modes leave true-last/manual-break lines non-justified.
  stretchLastLine = false,
  startBoundary?: LineBoundary,
  widthPolicy: 'bounded' | 'intrinsic' | 'unwrapped' = 'bounded',
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  overflowPunct = false,
  justifiedCompression = false,
  // WORD_KASHIDA_LINE_FIT: the part of each bounded line's width its text may
  // fill; retained paint still justifies to the full width.
  lineFit?: LineFitWidth,
  passContext?: Readonly<{
    probeHeights: readonly number[] | null;
    probeFloors?: readonly number[] | null;
    preparedFloatWrap?: PreparedFloatWrap;
  }>,
): LayoutLine[] {
  if (passContext === undefined) {
    // Keep the pass inside the existing declaration: extracting this root
    // implementation would widen the frozen migration boundary, while moving
    // it under layout/ would reverse that layer's dependency direction. The
    // public overload deliberately hides this pass-only context from .d.ts.
    const runPass = (
      probeHeights: readonly number[] | null,
      preparedFloatWrap?: PreparedFloatWrap,
      probeFloors: readonly number[] | null = null,
    ): LayoutLine[] => (layoutLines as unknown as (
      ...args: unknown[]
    ) => LayoutLine[])(
      ctx,
      cloneSegmentsForLinePass(segs),
      maxWidth,
      firstIndent,
      scale,
      tabStops,
      wrapCtx,
      fontFamilyClasses,
      tabOriginPx,
      kinsoku,
      characterGrid,
      defaultTabPt,
      marginRightPx,
      baseRtl,
      isJustified,
      stretchLastLine,
      startBoundary,
      widthPolicy,
      verticalGlyphMeasurement,
      overflowPunct,
      justifiedCompression,
      lineFit,
      { probeHeights, probeFloors, preparedFloatWrap },
    );
    // DrawingML wrap=none (§21.1.2.1.1) removes only the automatic break at the
    // text-body edge; floats of the same story (e.g. a §17.3.1.11 frame host)
    // still exclude its lines. An unwrapped line with a wrap context therefore
    // takes the same exclusion fixed point and publishes the same allocation
    // provenance as a bounded one. Intrinsic measurement has no placement.
    if (!wrapCtx || widthPolicy === 'intrinsic') return runPass(null);
    const preparedFloatWrap = wrapCtx.lineWindow
      ? undefined
      : prepareFloatWrap(wrapCtx.floats);
    const lines = convergeLineWrap(
      (probeHeights, probeFloors) => runPass(probeHeights, preparedFloatWrap, probeFloors),
      (line) => wrapCtx.lineBoxH(
        line.ascent,
        line.descent,
        line.hasRuby,
        line.intendedSingle,
        line.eastAsian,
        line.gridCountSingle,
        line.uniformPositionAuto,
        line.inlinePictureTextSingle,
        line.latinGridCountSingle,
      ),
      wrapCtx.resolveLineAdvances,
    );
    const advances = wrapCtx.resolveLineAdvances?.(lines);
    return lines.map((line, index) => ({
      ...line,
      // Publish provenance only after exact-state convergence confirms the
      // same physical partition, probes and tops, never on an exploratory pass.
      wrapAllocation: Object.freeze({
        physicalLineIndex: line.physicalLineIndex!,
        topYPt: line.topY!,
        advancePt: advances?.[index] ?? wrapCtx.lineBoxH(
          line.ascent, line.descent, line.hasRuby, line.intendedSingle,
          line.eastAsian, line.gridCountSingle, line.uniformPositionAuto,
          line.inlinePictureTextSingle, line.latinGridCountSingle,
        ),
      }),
    }));
  }
  return runLineBreakerPass({
    ctx, segs, maxWidth, firstIndent, scale, tabStops, wrapCtx,
    fontFamilyClasses, tabOriginPx, kinsoku, characterGrid, defaultTabPt,
    marginRightPx, baseRtl, isJustified, stretchLastLine, startBoundary,
    widthPolicy, verticalGlyphMeasurement, overflowPunct, justifiedCompression,
    ...(lineFit ? { lineFit } : {}), passContext,
  });
}
