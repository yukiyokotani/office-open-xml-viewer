import {
  paragraphGridRightAdjustmentPt,
  type ParagraphLayoutContext,
} from './layout-context.js';
import {
  buildSegments,
  getDefaultFontSize,
  isGridLineRule,
  layoutLines,
  lineBelowBaselinePx,
  lineBoxHeight,
  paragraphMarkBelowBaselinePt,
  paragraphMarkLineHeight,
  paragraphMarkLineMetrics,
  type DocGridCtx,
  type LineBoundary,
  type LayoutLine,
  type LineLayoutEnvironment,
  type WrapLayoutCtx,
} from './line-layout.js';
import type { ParagraphLayoutSource } from './layout/text.js';
import type { DocParagraph } from './types.js';
import type { WrapOracle } from './layout/float-wrap-oracle.js';
import type { NumberingMarkerShapeInput, WritingMode } from './layout/types.js';
import { wordEmptyMarkMinimumStartWidthPx } from './layout/compatibility.js';
import { WORD_NUMBERING_MARKER_FIRST_LINE_UNION } from './layout/line-compatibility.js';
import type { MeasurementTextContext } from './layout/measurement-capabilities.js';

export type { LineLayoutEnvironment } from './line-layout.js';
export { createFloatWrapOracle } from './layout/float-wrap-oracle.js';
export type { WrapOracle } from './layout/float-wrap-oracle.js';

export interface ParagraphMeasurementEnvironment extends LineLayoutEnvironment {
  readonly documentHasEastAsianText: boolean;
  readonly paragraphMarkShapeInput?: NumberingMarkerShapeInput;
  /** Selected-face text marker box, resolved by retained numbering before line acquisition. */
  readonly firstLineNumberingMarkerBox?: Readonly<{ ascentPt: number; descentPt: number }>;
  /** Canonical section writing mode used by retained page geometry. */
  readonly pageWritingMode: WritingMode;
  /** The paragraph is acquired in a section-logical frame that paint rotates
   * into a vertical physical page. This is independent of glyph orientation. */
  readonly verticalPageFrame?: boolean;
}

export interface TextMeasurer {
  readonly context: MeasurementTextContext;
  readonly fontFamilyClasses: Readonly<Record<string, string>>;
}

export interface ParagraphPlacement {
  readonly startYPt: number;
  readonly paragraphXPt: number;
  readonly availableWidthPt: number;
  readonly maximumYPt: number;
  readonly suppressSpaceBefore: boolean;
  readonly wrap?: WrapOracle;
}

export interface MeasuredLine {
  readonly layout: LayoutLine;
  readonly topYPt: number;
  readonly advancePt: number;
}

export interface MeasuredParagraph {
  readonly lines: readonly MeasuredLine[];
  readonly markOnly: boolean;
  readonly requestedSpaceBeforePt: number;
  readonly requestedSpaceAfterPt: number;
  /** ECMA-376 §17.3.3.25 paragraph-wide uniform line advance in points, snapped
   *  to the docGrid. Zero when the paragraph has no ruby. */
  readonly uniformRubyAdvancePt: number;
  readonly contentStartYPt: number;
  readonly contentEndYPt: number;
  /**
   * ECMA-376 §17.3.1.29 / §17.3.1.33 — the extent (pt) of the LAST line's box that
   * lies below its baseline (descent + half of any auto/atLeast leading). Word's
   * page fit is baseline-based: a line whose baseline sits within the text area may
   * let this below-baseline whitespace extend into the bottom margin. The paginator
   * uses it (for an empty paragraph, whose mark line paints no ink there) so a
   * trailing empty paragraph is not pushed to the next page merely because its
   * invisible mark box grazes past the bottom content edge.
   */
  readonly lastLineBelowBaselinePt: number;
  readonly placement: Readonly<ParagraphPlacement>;
}

/** Project the normalized paragraph policy without losing w:docGrid/@w:type. */
export function paragraphCharacterGrid(
  context: ParagraphLayoutContext,
): DocGridCtx | undefined {
  if (!context.characterGrid.active) return undefined;
  return {
    type: context.characterGrid.kind,
    linePitchPt: null,
    characterPitchPt: context.characterGrid.pitchPt,
    charSpacePt: context.characterGrid.deltaPt,
  };
}

function paragraphGrid(context: ParagraphLayoutContext): DocGridCtx {
  const characterGrid = paragraphCharacterGrid(context);
  return {
    type: characterGrid
      ? characterGrid.type
      : context.lineGrid.active ? 'lines' : null,
    linePitchPt: context.lineGrid.active ? context.lineGrid.pitchPt : null,
    characterPitchPt: context.characterGrid.active ? context.characterGrid.pitchPt : null,
    charSpacePt: context.characterGrid.active ? context.characterGrid.deltaPt : null,
  };
}

/** Preserve the renderer's paragraph-wide ruby/docGrid height calculation. */
function snapParagraphLineToGrid(heightPt: number, grid: DocGridCtx): number {
  if (!isGridLineRule(grid)) return heightPt;
  const pitchPt = grid.linePitchPt!;
  if (pitchPt <= 0) return heightPt;
  if (heightPt <= pitchPt) return pitchPt;
  return Math.ceil(heightPt / pitchPt) * pitchPt;
}

export function measureParagraph(
  paragraph: ParagraphLayoutSource,
  context: ParagraphLayoutContext,
  placement: ParagraphPlacement,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
  continuation?: {
    readonly boundary: LineBoundary;
    readonly uniformRubyAdvancePt?: number;
  },
): MeasuredParagraph {
  const grid = paragraphGrid(context);
  const rightGridAdjustmentPt = paragraphGridRightAdjustmentPt(
    context,
    placement.availableWidthPt,
  );
  const paragraphWidthPt = Math.max(
    1,
    placement.availableWidthPt
      - context.physicalIndentLeftPt
      - context.physicalIndentRightPt
      - rightGridAdjustmentPt,
  );
  const paragraphXPt = placement.paragraphXPt + context.physicalIndentLeftPt;
  const requestedSpaceBeforePt = context.spaceBeforePt;
  const requestedSpaceAfterPt = context.spaceAfterPt;
  const recordedPlacement = Object.freeze({ ...placement });
  const fontFamilyClasses = measurer.fontFamilyClasses as Record<string, string>;
  // The `w:useFELayout` compatibility projection applies the Far East docGrid
  // allocation to an otherwise content-less paragraph mark as well as to its
  // text lines. This matters when the mark's design height crosses a grid-cell
  // boundary: a 16pt East Asian mark on an 18pt grid occupies two cells even
  // when the document contains no literal CJK code point.
  const markUsesEastAsianGrid = environment.documentHasEastAsianText === true
    || environment.useFeLayout === true;

  let cursorPt = placement.startYPt
    + (placement.suppressSpaceBefore ? 0 : requestedSpaceBeforePt);
  if (placement.wrap) {
    // §20.4.2.20 / §17.6.4 column scope: pass this paragraph's COLUMN band
    // (placement.paragraphXPt / availableWidthPt = colX()/colW()), NOT the
    // indented text band, so measure agrees bit-for-bit with the paint pass,
    // which scopes the same skip against state.contentX/contentW (the column
    // band). A topAndBottom float anchored in another newspaper column is
    // filtered out in both passes.
    cursorPt = placement.wrap.skipTopAndBottomBands({
      yPt: cursorPt,
      columnXPt: placement.paragraphXPt,
      columnWidthPt: placement.availableWidthPt,
    });
  }

  const measureMarkOnly = (): MeasuredParagraph => {
    let markTopPt = cursorPt;
    const markAdvancePt = paragraphMarkLineHeight(
      paragraph,
      1,
      grid,
      context.hasRuby,
      markUsesEastAsianGrid,
      measurer.context,
      fontFamilyClasses,
      context.lineSpacing,
      environment.resolvedLocalFonts,
      environment.layoutServices?.text,
      environment.paragraphMarkShapeInput,
      environment.useFeLayout === true,
    );
    if (placement.wrap) {
      markTopPt = placement.wrap.lineWindow({
        topYPt: markTopPt,
        minimumStartWidthPt: getDefaultFontSize(paragraph),
        squareMinimumStartWidthPt: wordEmptyMarkMinimumStartWidthPx(
          getDefaultFontSize(paragraph),
          1,
        ),
        probeHeightPt: markAdvancePt,
        paragraphXPt,
        maximumWidthPt: paragraphWidthPt,
        // §20.4.2.20 / §17.6.4 column scope: the topAndBottom gate sees the raw
        // COLUMN band, not the indented mark band above.
        columnXPt: placement.paragraphXPt,
        columnWidthPt: placement.availableWidthPt,
      }).topYPt;
    }
    return {
      lines: [],
      markOnly: true,
      requestedSpaceBeforePt,
      requestedSpaceAfterPt,
      uniformRubyAdvancePt: 0,
      contentStartYPt: markTopPt,
      contentEndYPt: markTopPt + markAdvancePt,
      lastLineBelowBaselinePt: paragraphMarkBelowBaselinePt(
        paragraph,
        grid,
        context.hasRuby,
        markUsesEastAsianGrid,
        measurer.context,
        fontFamilyClasses,
        context.lineSpacing,
        environment.resolvedLocalFonts,
        environment.layoutServices?.text,
        environment.paragraphMarkShapeInput,
        environment.useFeLayout === true,
      ),
      placement: recordedPlacement,
    };
  };

  const segments = buildSegments(paragraph.runs, {
    ...environment,
    lineSpacing: context.lineSpacing,
    lineGridActive: context.lineGrid.active,
  });
  if (segments.length === 0) return measureMarkOnly();

  if (context.lineSpacing?.rule === 'auto'
    && context.lineSpacing.value > 1
    && segments.some((segment) => 'imagePath' in segment && segment.inlinePicture === true)) {
    // An image-only line has no text segment from which to obtain the authored
    // single-line height. Measure its paragraph mark through the same selected
    // face service as an empty paragraph, without the auto multiplier or grid.
    // The visible text's own selected metric takes precedence on mixed lines.
    const markSinglePx = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, markUsesEastAsianGrid,
      measurer.context, fontFamilyClasses, null, environment.resolvedLocalFonts,
      environment.layoutServices?.text, environment.paragraphMarkShapeInput,
    ).advancePx;
    for (const segment of segments) {
      if ('imagePath' in segment && segment.inlinePicture === true) {
        segment.paragraphMarkSinglePx = markSinglePx;
      }
    }
  }

  const wrapContext: WrapLayoutCtx | undefined = placement.wrap
    ? {
        startPageY: cursorPt,
        paraX: paragraphXPt,
        // Raw COLUMN band (placement) for the topAndBottom gate; paraX above is
        // the indented text band for the square side-gap math (§20.4.2.20 vs
        // §20.4.2.17). See WrapLayoutCtx.columnXPt.
        columnXPt: placement.paragraphXPt,
        columnWidthPt: placement.availableWidthPt,
        floats: [],
        paragraphMarkLineStartWidth: wordEmptyMarkMinimumStartWidthPx(
          getDefaultFontSize(paragraph),
          1,
        ),
        lineWindow: (input) => placement.wrap!.lineWindow(input),
        lineBoxH: (ascent, descent, _hasRuby, intendedSingle, eastAsian, gridCountSingle, uniformPositionAuto, inlinePictureTextSingle) => lineBoxHeight(
          context.lineSpacing,
          ascent,
          descent,
          1,
          grid,
          context.hasRuby,
          intendedSingle ?? 0,
          // §17.6.5 cell rounding follows this line's script, matching text boxes;
          // ruby paragraphs retain their established uniform paragraph resolver.
          context.hasRuby ? context.hasEastAsianText : (eastAsian ?? false),
          gridCountSingle,
          undefined,
          uniformPositionAuto,
          inlinePictureTextSingle,
        ),
        pageH: placement.maximumYPt,
      }
    : undefined;
  const lines = layoutLines(
    measurer.context,
    segments,
    paragraphWidthPt,
    // ECMA-376 §17.3.1.12: first-line and hanging indents apply only to the
    // paragraph's first line, not to a continuation measured in another column.
    continuation ? 0 : context.firstIndentPt,
    1,
    [...context.tabStops],
    wrapContext,
    fontFamilyClasses,
    context.physicalIndentLeftPt,
    context.kinsoku,
    grid,
    context.defaultTabPt,
    paragraphWidthPt + context.physicalIndentRightPt + rightGridAdjustmentPt,
    context.baseRtl,
    context.isJustified,
    context.stretchLastLine,
    continuation?.boundary,
    undefined,
    environment.verticalGlyphMeasurement,
    context.overflowPunct !== false,
  );
  if (lines.length === 0) return measureMarkOnly();

  let uniformRubyAdvancePt = context.hasRuby
    ? snapParagraphLineToGrid(
        Math.max(0, ...lines.map((line) => lineBoxHeight(
          context.lineSpacing,
          line.ascent,
          line.descent,
          1,
          grid,
          true,
          line.intendedSingle,
          context.hasEastAsianText,
        ))),
        grid,
      )
    : 0;
  if (context.hasRuby && continuation?.uniformRubyAdvancePt !== undefined) {
    uniformRubyAdvancePt = Math.max(
      uniformRubyAdvancePt,
      continuation.uniformRubyAdvancePt,
    );
  }
  const measuredLines: MeasuredLine[] = [];
  for (const [lineIndex, originalLine] of lines.entries()) {
    const markerBox = lineIndex === 0 && !continuation && !placement.wrap
      && !context.lineGrid.active && context.lineSpacing?.rule === 'auto'
      && context.lineSpacing.value >= 1 && !context.hasRuby
      && !originalLine.uniformPositionAuto && !originalLine.inlinePictureTextSingle
      ? environment.firstLineNumberingMarkerBox : undefined;
    const markerAscent = markerBox?.ascentPt;
    const markerDescent = markerBox?.descentPt;
    const line = markerAscent !== undefined && markerDescent !== undefined
      && Number.isFinite(markerAscent) && Number.isFinite(markerDescent)
      ? {
          ...originalLine,
          ascent: Math.max(originalLine.ascent, markerAscent),
          descent: Math.max(originalLine.descent, markerDescent),
          visibleAscent: Math.max(originalLine.visibleAscent ?? originalLine.ascent, markerAscent),
          visibleDescent: Math.max(originalLine.visibleDescent ?? originalLine.descent, markerDescent),
        }
      : originalLine;
    const topYPt = line.topY !== undefined && line.topY > cursorPt
      ? line.topY
      : cursorPt;
    const textSinglePt = Math.max(
      originalLine.ascent + originalLine.descent,
      originalLine.intendedSingle,
    );
    // ECMA-376 §17.9.6 supplies marker rPr and §17.3.1.33 the auto multiple,
    // but neither specifies their line-box union. This selected-face projection
    // is limited to the observed auto/non-grid text-marker class; other classes
    // retain their established allocation.
    void WORD_NUMBERING_MARKER_FIRST_LINE_UNION;
    const markerNaturalPt = line.ascent + line.descent;
    const markerRaisesBox = line !== originalLine
      && markerNaturalPt > textSinglePt;
    const advancePt = markerRaisesBox
      ? markerNaturalPt + textSinglePt * ((context.lineSpacing?.value ?? 1) - 1)
      : context.hasRuby
      ? uniformRubyAdvancePt
      : lineBoxHeight(
          context.lineSpacing,
          line.ascent,
          line.descent,
          1,
          grid,
          false,
          line.intendedSingle,
          // §17.6.5 cell rounding is gated by the line's script; a Latin-only
          // line in a CJK paragraph keeps its natural height.
          line.eastAsian ?? false,
          line.gridCountSingle,
          undefined,
          line.uniformPositionAuto,
          line.inlinePictureTextSingle,
        );
    measuredLines.push({ layout: line, topYPt, advancePt });
    cursorPt = topYPt + advancePt;
  }

  const lastLine = measuredLines[measuredLines.length - 1];
  return {
    lines: measuredLines,
    markOnly: false,
    requestedSpaceBeforePt,
    requestedSpaceAfterPt,
    uniformRubyAdvancePt,
    contentStartYPt: measuredLines[0].topYPt,
    contentEndYPt: cursorPt,
    lastLineBelowBaselinePt: lineBelowBaselinePx(
      lastLine.advancePt,
      lastLine.layout.ascent,
      lastLine.layout.descent,
    ),
    placement: recordedPlacement,
  };
}
