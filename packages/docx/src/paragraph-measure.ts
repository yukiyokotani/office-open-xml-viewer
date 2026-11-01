import { specifiedTextLineMetrics, specifiedTextParagraphIsHomogeneous } from './layout/specified-line-spacing.js';
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
  paragraphMarkLineMetrics,
  type DocGridCtx,
  type LineBoundary,
  type LayoutLine,
  type LayoutImageSeg,
  type LineLayoutEnvironment,
  type WrapLayoutCtx,
} from './line-layout.js';
import type { ParagraphLayoutSource } from './layout/text.js';
import type { WrapOracle } from './layout/float-wrap-oracle.js';
import type { NumberingMarkerShapeInput, WritingMode } from './layout/types.js';
import { wordEmptyMarkMinimumStartWidthPx } from './layout/compatibility.js';
import {
  wordTextMarkerKeepsBodyDescent,
  WORD_INLINE_OBJECT_GRID_LINE_ALLOCATION,
  WORD_PICTURE_BULLET_LINE_PARTICIPATION,
  wordJustifiedInterwordCompressionApplies,
} from './layout/line-compatibility.js';
import { LayoutInvariantError } from './layout/diagnostics.js';
import type { MeasurementTextContext } from './layout/measurement-capabilities.js';

export type { LineLayoutEnvironment } from './line-layout.js';
export { createFloatWrapOracle } from './layout/float-wrap-oracle.js';
export type { WrapOracle } from './layout/float-wrap-oracle.js';

export interface ParagraphMeasurementEnvironment extends LineLayoutEnvironment {
  readonly documentHasEastAsianText: boolean;
  readonly paragraphMarkShapeInput?: NumberingMarkerShapeInput;
  /** Selected-face text marker box, resolved by retained numbering before line acquisition. */
  /** Ascent (with line gap) of a numbered paragraph's text marker; an empty
   * item's mark line takes part with it (#1884). */
  readonly markOnlyTextMarkerAscentPt?: number;
  /** WrapLayoutCtx.numberingFirstLine for a numbered paragraph's first line. */
  readonly numberingFirstLine?: NonNullable<WrapLayoutCtx['numberingFirstLine']>;
  readonly firstLineNumberingMarkerBox?: Readonly<{ ascentPt: number; descentPt: number; intendedSinglePt?: number }>;
  /** Display height of a sized picture-bullet marker, which stands on the
   * paragraph's first baseline (see `markerMetric` in the line breaker). */
  readonly firstLinePictureMarkerHeightPt?: number;
  /** Canonical section writing mode used by retained page geometry. */
  readonly pageWritingMode: WritingMode;
  /** The paragraph is acquired in a section-logical frame that paint rotates
   * into a vertical physical page. This is independent of glyph orientation. */
  readonly verticalPageFrame?: boolean;
  /** `w:compatSetting` compatibilityMode (§17.15.3.4) of the document; absent when not authored. */
  readonly compatibilityMode?: number;
  /** §17.15.3.8 ulTrailSpace: hanging line-end spaces keep their underline. */
  readonly ulTrailSpace?: boolean;
  /** §17.15.3.5 doNotExpandShiftReturn: lines ended by a manual break stay
   * unjustified. */
  readonly doNotExpandShiftReturn?: boolean;
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
  /** DrawingML wrap=none: keep the real alignment band, allow inline overflow. */
  readonly noWrap?: boolean;
}

export interface MeasuredLine {
  readonly layout: LayoutLine;
  readonly topYPt: number;
  readonly advancePt: number;
}

export interface MeasuredParagraph {
  readonly lines: readonly MeasuredLine[];
  readonly markOnly: boolean;
  /** Grid-only leading below the empty mark's own font/spacing box. */
  readonly markGridAfterLeadingPt?: number;
  /** Selected empty-mark wrap gap, retained for measurement-free shading. */
  readonly markWrapBounds?: { readonly xPt: number; readonly yPt: number; readonly widthPt: number; readonly heightPt: number };
  /** An empty numbered item whose marker leading edge a float moved: that
   * edge, logical pt from the leading indent (placeNumberingFirstLine). */
  readonly markNumberingLeadingEdgePt?: number;
  readonly requestedSpaceBeforePt: number;
  readonly requestedSpaceAfterPt: number;
  /** Established paragraph-wide ruby allocation in points, snapped to docGrid
   *  (§17.6.5). §17.3.3.25 defines guide/base placement, not uniform advances.
   *  Zero when the paragraph has no ruby. */
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
  /** Draw-only baseline of a mark-only line that carries a picture bullet,
   * measured from `contentStartYPt` (see `measureMarkOnly`). */
  readonly markPictureBaselineOffsetPt?: number;
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

/** Allocate a line advance to whole cells of the active document grid. */
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
  // Keep that compatibility projection inside the active line grid. Part 4
  // §14.8.3.50 does not select an East Asian font for an otherwise Latin mark;
  // §17.3.2.26's explicit mark hint remains authoritative without a grid.
  // Windows Word controls with useFELayout and no grid retain the Latin
  // single-line box (Times New Roman 10pt footnote separators; Ubuntu 12pt
  // body marks and footnote/endnote separators),
  // including when the inherited East Asian font slot names a different face.
  const markUsesEastAsianGrid = environment.documentHasEastAsianText === true
    || (environment.useFeLayout === true && context.lineGrid.active);
  // WORD_PICTURE_BULLET_LINE_PARTICIPATION: a picture bullet stands on its
  // paragraph's first baseline and takes part in that line's box, except
  // under exact spacing, whose box and baseline it leaves unchanged (the
  // picture then overflows above the line).
  void WORD_PICTURE_BULLET_LINE_PARTICIPATION;
  const pictureMarkerHeight = context.lineSpacing?.rule === 'exact' && context.lineSpacing.value > 0
    ? undefined : environment.firstLinePictureMarkerHeightPt;
  // The metric segment's height. It is prepended in every measurement of the
  // paragraph, continuations included, so that line boundaries index the same
  // segment array; a continuation starts past it and never admits it.
  const pictureMarkerSegmentHeightPt = pictureMarkerHeight !== undefined
    && Number.isFinite(pictureMarkerHeight) && pictureMarkerHeight > 0
    ? pictureMarkerHeight : undefined;
  const firstLinePictureMarkerHeightPt = continuation ? undefined : pictureMarkerSegmentHeightPt;

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

  // Windows Word 16.0 #1884 controls (Arial 10pt / Times New Roman 12pt
  // marks; 6-20pt Symbol bullets and Arial numbers; exact 12/18/30, atLeast
  // 18/30/60, auto 1/1.5/2): an empty item's text marker joins its mark line
  // like a picture bullet (max of the two ascents over the mark's descent,
  // then the line rule). The baseline sits four fifths down an exact line,
  // the mark's descent above the bottom of an atLeast line, and at the union
  // ascent of an automatic line, whose added leading falls below it.
  // The controls are compatibility mode 15, exact lines within the 12-30pt
  // interval of WORD_SPECIFIED_TEXT_LINE_BOX; other classes keep their path.
  const markSpacing = context.lineSpacing;
  const textMarkerAscentPt = !continuation && firstLinePictureMarkerHeightPt === undefined
    && environment.compatibilityMode === 15
    && !context.lineGrid.active && !context.hasRuby
    && !(markSpacing?.rule === 'exact' && (markSpacing.value < 12 || markSpacing.value > 30))
    ? environment.markOnlyTextMarkerAscentPt : undefined;
  const textMarkerMarkBelowBaselinePt = (advancePt: number, ascentPt: number, descentPt: number) => {
    const spacing = context.lineSpacing;
    if (spacing?.rule === 'exact' && spacing.value > 0) return advancePt / 5;
    if (spacing?.rule === 'atLeast') return Math.min(advancePt, descentPt);
    return Math.max(0, advancePt - ascentPt);
  };
  const measureMarkOnly = (): MeasuredParagraph => {
    let markTopPt = cursorPt;
    const markMetrics = paragraphMarkLineMetrics(
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
      firstLinePictureMarkerHeightPt ?? textMarkerAscentPt ?? 0,
    );
    const markAdvancePt = markMetrics.advancePx;
    let markWrapBounds: MeasuredParagraph['markWrapBounds'];
    let markNumberingLeadingEdgePt: number | undefined;
    if (placement.wrap) {
      const window = placement.wrap.lineWindow({
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
      });
      markTopPt = window.topYPt;
      const numbering = environment.numberingFirstLine;
      if (numbering && !continuation) {
        // The marker hosts on this mark line exactly as on a first text line
        // (placeNumberingFirstLine): query the band widened by the authored
        // hanging indent and place the marker at the moved leading edge.
        const authored = numbering.authoredFirstIndentPx;
        const hanging = Math.min(0, authored);
        const band = placement.wrap.lineWindow({
          topYPt: markTopPt, minimumStartWidthPt: 0, squareMinimumStartWidthPt: 0,
          probeHeightPt: markAdvancePt,
          paragraphXPt: paragraphXPt + (context.baseRtl ? 0 : hanging),
          maximumWidthPt: paragraphWidthPt - hanging,
          columnXPt: placement.paragraphXPt, columnWidthPt: placement.availableWidthPt,
        });
        const windowEdge = context.baseRtl
          ? paragraphWidthPt - (band.xOffsetPt + band.maximumWidthPt)
          : hanging + band.xOffsetPt;
        const leadingEdge = Math.max(Math.max(0, windowEdge) + authored, windowEdge);
        if (band.topYPt === markTopPt && leadingEdge > authored) markNumberingLeadingEdgePt = leadingEdge;
      }
      if (window.xOffsetPt !== 0 || window.maximumWidthPt !== paragraphWidthPt) {
        markWrapBounds = { xPt: paragraphXPt + window.xOffsetPt, yPt: markTopPt,
          widthPt: window.maximumWidthPt, heightPt: markAdvancePt };
      }
    }
    // WORD_PICTURE_BULLET_LINE_PARTICIPATION: a picture bullet on an empty
    // item stands on the mark line's baseline, which an automatic multiple
    // pins to the top of its box like a text line's
    // (WORD_AUTO_MULTIPLE_BASELINE_PIN). This is draw-only; pagination keeps
    // the centred trailing-mark metric below.
    const lineSpacing = context.lineSpacing;
    const markPictureBaselineOffsetPt = firstLinePictureMarkerHeightPt !== undefined
      && lineSpacing?.rule === 'auto' && lineSpacing.value > 1
      && !context.hasRuby && !context.lineGrid.active
      ? (Math.max(markMetrics.ascentPx + markMetrics.descentPx, markMetrics.intendedSinglePx)
        - markMetrics.ascentPx - markMetrics.descentPx) / 2 + markMetrics.ascentPx
      : undefined;
    return {
      ...(markWrapBounds ? { markWrapBounds } : {}),
      ...(markNumberingLeadingEdgePt === undefined ? {} : { markNumberingLeadingEdgePt }),
      ...(markPictureBaselineOffsetPt === undefined ? {} : { markPictureBaselineOffsetPt }),
      lines: [],
      markOnly: true,
      requestedSpaceBeforePt,
      requestedSpaceAfterPt,
      uniformRubyAdvancePt: 0,
      contentStartYPt: markTopPt,
      contentEndYPt: markTopPt + markAdvancePt,
      lastLineBelowBaselinePt: textMarkerAscentPt === undefined
        ? lineBelowBaselinePx(markAdvancePt, markMetrics.ascentPx, markMetrics.descentPx)
        : textMarkerMarkBelowBaselinePt(markAdvancePt, markMetrics.ascentPx, markMetrics.descentPx),
      markGridAfterLeadingPt: markMetrics.gridAfterLeadingPx,
      placement: recordedPlacement,
    };
  };

  const segments = buildSegments(paragraph.runs, {
    ...environment,
    lineSpacing: context.lineSpacing,
    lineGridActive: context.lineGrid.active,
    characterGridActive: context.characterGrid.active,
    paragraphRtl: context.baseRtl,
    autoSpaceDE: paragraph.autoSpaceDE,
    autoSpaceDN: paragraph.autoSpaceDN,
  });
  if (segments.length === 0) return measureMarkOnly();
  const pictureMarkerHeightPt = firstLinePictureMarkerHeightPt;
  if (pictureMarkerSegmentHeightPt !== undefined) {
    // WORD_PICTURE_BULLET_LINE_PARTICIPATION: a picture bullet takes part in
    // its first line like an inline picture of that height standing on the
    // baseline (automatic multiples add the text's own leading only, atLeast
    // keeps the larger box, ruby forms one union with it). A metric-only
    // segment carries that height into the line breaker's
    // inline-picture accounting without taking width, a break opportunity or
    // paint; the marker itself is painted separately.
    segments.unshift({
      imagePath: '', mimeType: '', widthPt: 0, heightPt: pictureMarkerSegmentHeightPt,
      inlinePicture: true, markerMetric: true,
      anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
    } as LayoutImageSeg);
  }

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

  // Library policy retains the established paragraph-wide ruby reserve
  // (§17.3.3.25 describes the guide above the base; §17.6.5 supplies grid cells).
  // Resolve the complete physical unions before querying floats: a per-fragment
  // reserve cannot determine either a probe band or the next physical origin.
  const specifiedParagraph = specifiedTextParagraphIsHomogeneous(paragraph,
    context.lineGrid.active || environment.compatibilityMode !== 15);
  const allocateLines = (lines: readonly LayoutLine[]) => {
    // `word-ruby-paragraph-uniform-line-advance`
    let uniformRubyAdvancePt = context.hasRuby
      ? snapParagraphLineToGrid(
          lines.reduce((heightPt, line) => Math.max(heightPt, lineBoxHeight(
            context.lineSpacing,
            line.ascent,
            line.descent,
            1,
            grid,
            true,
            line.intendedSingle,
            context.hasEastAsianText,
          )), 0),
          grid,
        )
      : 0;
    if (context.hasRuby && continuation?.uniformRubyAdvancePt !== undefined) {
      uniformRubyAdvancePt = Math.max(
        uniformRubyAdvancePt,
        continuation.uniformRubyAdvancePt,
      );
    }
    const allocations: { layout: LayoutLine; advancePt: number }[] = [];
    // A float can split the first physical line into fragments; the metric
    // segment joins only the fragment that holds it. Every fragment of that
    // physical line shares the bullet's ascent so that they resolve one advance.
    const markerFragment = pictureMarkerHeightPt === undefined ? undefined : lines[0];
    const holdsPictureMarker = (line: LayoutLine, lineIndex: number): boolean =>
      markerFragment !== undefined && (lineIndex === 0
        || (markerFragment.physicalLineIndex !== undefined
          && line.physicalLineIndex === markerFragment.physicalLineIndex));
    for (const [lineIndex, fragment] of lines.entries()) {
      const firstPhysicalLine = holdsPictureMarker(fragment, lineIndex);
      const originalLine = firstPhysicalLine && lineIndex !== 0
        ? {
            ...fragment,
            ascent: Math.max(fragment.ascent, pictureMarkerHeightPt!),
            visibleAscent: Math.max(fragment.visibleAscent ?? fragment.ascent, pictureMarkerHeightPt!),
            ...(markerFragment!.inlinePictureTextSingle === undefined ? {} : {
              inlinePictureTextSingle: Math.max(
                fragment.inlinePictureTextSingle ?? 0,
                markerFragment!.inlinePictureTextSingle,
              ),
            }),
          }
        : fragment;
      // Windows Word 16.0 #1883 controls keep the same marker union beside a
      // square-wrapped float (a 42pt Symbol bullet raises the first line).
      // WORD_NUMBERING_MARKER_FIRST_LINE_UNION: the first line keeps body
      // descent in the measured modes. Marker eligibility beside floats and
      // under atLeast is still bounded to the earlier mode-15 controls.
      // Correct body allocation here; note reserves must retain their measured
      // separator/line boxes even when excess body height changes pagination.
      const windowsMarkerUnion = environment.compatibilityMode === 15;
      const markerAscentOnly = wordTextMarkerKeepsBodyDescent(environment.compatibilityMode);
      const onFirstPhysicalLine = lineIndex === 0 || (windowsMarkerUnion
        && lines[0]?.physicalLineIndex !== undefined
        && fragment.physicalLineIndex === lines[0]!.physicalLineIndex);
      const markerBox = onFirstPhysicalLine && !continuation
        && (windowsMarkerUnion || !placement.wrap)
        && !context.lineGrid.active
        && (context.lineSpacing == null
          || (windowsMarkerUnion && context.lineSpacing.rule === 'atLeast')
          || (context.lineSpacing.rule === 'auto' && context.lineSpacing.value >= 1))
        && !context.hasRuby
        && !originalLine.uniformPositionAuto && !originalLine.inlinePictureTextSingle
        ? environment.firstLineNumberingMarkerBox : undefined;
      const markerAscent = markerBox?.ascentPt;
      const markerDescent = markerBox?.descentPt;
      const line = markerAscentOnly
        ? (markerAscent !== undefined && Number.isFinite(markerAscent)
          && markerAscent > originalLine.ascent
          ? {
              ...originalLine,
              ascent: markerAscent,
              visibleAscent: Math.max(originalLine.visibleAscent ?? originalLine.ascent, markerAscent),
            }
          : originalLine)
        : markerAscent !== undefined && markerDescent !== undefined
          && Number.isFinite(markerAscent) && Number.isFinite(markerDescent)
          ? {
              ...originalLine,
              ascent: Math.max(originalLine.ascent, markerAscent),
              descent: Math.max(originalLine.descent, markerDescent),
              visibleAscent: Math.max(originalLine.visibleAscent ?? originalLine.ascent, markerAscent),
              visibleDescent: Math.max(originalLine.visibleDescent ?? originalLine.descent, markerDescent),
              intendedSingle: Math.max(originalLine.intendedSingle, markerBox?.intendedSinglePt ?? 0),
              visibleIntendedSingle: Math.max(originalLine.visibleIntendedSingle ?? originalLine.intendedSingle,
                markerBox?.intendedSinglePt ?? 0),
            }
          : originalLine;
      const textSinglePt = Math.max(
        originalLine.ascent + originalLine.descent,
        originalLine.intendedSingle,
      );
      const markerNaturalPt = markerAscentOnly ? line.ascent + line.descent
        : Math.max(line.ascent + line.descent, markerBox?.intendedSinglePt ?? 0);
      const markerRaisesBox = line !== originalLine
        && markerNaturalPt > textSinglePt;
      const specified = specifiedParagraph
        ? specifiedTextLineMetrics(line, context, paragraph, environment.compatibilityMode,
            environment.verticalPageFrame === true, environment.paragraphMarkShapeInput,
            markerRaisesBox ? markerNaturalPt : undefined)
        : null;
      // WORD_PICTURE_BULLET_LINE_PARTICIPATION: a picture bullet that supplies
      // an East Asian line's ascent counts grid cells together with the
      // text's descent.
      const pictureMarkerGridCountPt = firstPhysicalLine
        && pictureMarkerHeightPt! >= line.ascent
        ? Math.max(line.gridCountSingle, line.ascent + line.descent)
        : undefined;
      // Automatic multiples add the text's own leading to the raised box; an
      // atLeast line takes the larger of its minimum and that box (below).
      let advancePt = specified ? specified.advancePt
        : markerRaisesBox && context.lineSpacing?.rule !== 'atLeast'
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
            pictureMarkerGridCountPt ?? line.gridCountSingle,
            undefined,
            line.uniformPositionAuto,
            line.inlinePictureTextSingle,
            line.latinGridCountSingle,
          );
      // §20.4.2.8: wp:inline objects affect their line's height like glyphs.
      // §17.6.5 / §17.3.1.32 govern grid participation and the exact override.
      // WORD_INLINE_OBJECT_GRID_LINE_ALLOCATION reuses text grid allocation
      // and the existing centered baseline for pictures/charts/shapes;
      // retain the separate atLeast path and floating-anchor origin rules.
      void WORD_INLINE_OBJECT_GRID_LINE_ALLOCATION;
      if (context.lineGrid.active
        && (context.lineSpacing == null || context.lineSpacing.rule === 'auto')
        // Only an object supplying the line's ascent owns this allocation.
        // A small object beside taller text must retain that text's grid policy.
        && line.ascent + line.descent >= line.intendedSingle
        && (line.segments.some((segment) => 'imagePath' in segment && !segment.anchor
          && segment.heightPt >= line.ascent)
          // A picture-bullet marker on the first line is that inline object
          // (WORD_PICTURE_BULLET_LINE_PARTICIPATION).
          || (firstPhysicalLine && pictureMarkerHeightPt! >= line.ascent))) {
        advancePt = snapParagraphLineToGrid(advancePt, grid);
      }
      allocations.push({ layout: line, advancePt });
    }
    return { allocations, uniformRubyAdvancePt };
  };

  const wrapContext: WrapLayoutCtx | undefined = placement.wrap
    ? {
        startPageY: cursorPt,
        resolveLineAdvances: (lines) => allocateLines(lines).allocations.map(line => line.advancePt),
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
        hasExclusions: placement.wrap!.hasExclusions,
        ...(environment.numberingFirstLine && !continuation
          ? { numberingFirstLine: environment.numberingFirstLine } : {}),
        lineWindow: (input) => placement.wrap!.lineWindow(input),
        lineBoxH: (ascent, descent, _hasRuby, intendedSingle, eastAsian, gridCountSingle, uniformPositionAuto, inlinePictureTextSingle, latinGridCountSingle) => lineBoxHeight(
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
          latinGridCountSingle,
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
    placement.noWrap ? 'unwrapped' : undefined,
    environment.verticalGlyphMeasurement,
    context.overflowPunct !== false,
    wordJustifiedInterwordCompressionApplies(
      paragraph.alignment, environment.compatibilityMode,
    ) && environment.verticalCJK !== true
      ? { lineEndIdeographicSpaceHangs: paragraph.wordWrap !== false }
      : false,
  );
  if (lines.length === 0) return measureMarkOnly();

  const { allocations, uniformRubyAdvancePt } = allocateLines(lines);
  const measuredLines: MeasuredLine[] = [];
  let physicalLineIndex: number | undefined;
  let physicalTopPt = cursorPt;
  for (const { layout: line, advancePt } of allocations) {
    // A wrap top is valid only after the fixed point used this allocation for
    // this physical line. Fragment tops share that allocation; equal numeric
    // tops alone do not establish physical-line identity.
    const allocation = line.wrapAllocation;
    if (placement.wrap && (!allocation
      || allocation.physicalLineIndex !== line.physicalLineIndex
      || allocation.topYPt !== line.topY
      || allocation.advancePt !== advancePt)) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', 'line origin does not own its allocated advance');
    }
    const samePhysicalLine = line.physicalLineIndex !== undefined
      && line.physicalLineIndex === physicalLineIndex;
    const topYPt = allocation?.topYPt ?? (samePhysicalLine ? physicalTopPt : cursorPt);
    physicalLineIndex = line.physicalLineIndex;
    physicalTopPt = topYPt;
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
