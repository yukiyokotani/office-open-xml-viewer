import type { DocParagraph, DocxDocumentModel } from '../types.js';
import { segmentsHaveRtl } from '../bidi-line.js';
import type { KashidaLevel } from '../kashida-justify.js';
import {
  isGridLineRule,
  type DocGridCtx,
  type LayoutSeg,
  type LineLayoutEnvironment,
} from '../line-layout.js';
import {
  paragraphCharacterGrid,
  type ParagraphMeasurementEnvironment,
} from '../paragraph-measure.js';
import type { ParagraphLayoutContext, StoryContext } from '../layout-context.js';
import type { BodyMeasurementContext } from './acquisition-context.js';
import { writingModeFromTextDirection } from './coordinate-space.js';

function kashidaLevelOf(alignment: string | null | undefined): KashidaLevel | null {
  if (alignment === 'lowKashida') return 'low';
  if (alignment === 'mediumKashida') return 'medium';
  if (alignment === 'highKashida') return 'high';
  return null;
}

/** Whether scale-1 glyph geometry can remain the single layout authority and be
 * mapped to the paint viewport by a Canvas transform. */
export function canonicalParagraphTextScaleEligible(
  storyContext: StoryContext,
  verticalCJK: boolean | undefined,
  inFrame: boolean,
  hasWrapContext: boolean,
  paragraphContext: Pick<ParagraphLayoutContext, 'hasRuby' | 'baseRtl'>,
  paragraph: Pick<DocParagraph, 'alignment' | 'numbering'>,
  segments: readonly LayoutSeg[],
): boolean {
  const supportedBodyContainers = storyContext.story === 'body'
    && (storyContext.containers.length === 0
      || storyContext.containers.every((container) => container.kind === 'tableCell'));
  return !hasWrapContext
    && !inFrame
    && supportedBodyContainers
    && !verticalCJK
    && !paragraphContext.hasRuby
    && !paragraphContext.baseRtl
    && paragraph.numbering == null
    && !segmentsHaveRtl(segments)
    && kashidaLevelOf(paragraph.alignment) === null
    && segments.every((segment) =>
      !('isTab' in segment)
      && !('math' in segment)
      && (!('text' in segment) || segment.emphasisMark == null));
}

/** The document's default body font size used for line-number glyph metrics.
 * Parser-folded paragraph defaults win, followed by the first text run and the
 * ECMA-376 absent docDefaults size of 10 pt. */
export function docDefaultFontSizePt(document: DocxDocumentModel): number {
  for (const element of document.body) {
    if (element.type !== 'paragraph') continue;
    const paragraph = element as DocParagraph;
    if (typeof paragraph.defaultFontSize === 'number') return paragraph.defaultFontSize;
    for (const run of paragraph.runs) {
      if (run.type === 'text') return run.fontSize;
    }
  }
  return 10;
}

export function paragraphMeasurementEnvironment(
  state: BodyMeasurementContext,
): ParagraphMeasurementEnvironment {
  return {
    pageIndex: state.pageIndex,
    totalPages: state.totalPages,
    displayPageNumber: state.displayPageNumber,
    pageNumberFormat: state.pageNumberFormat,
    currentDateMs: state.currentDateMs,
    noteNumbers: state.noteNumbers,
    noteReferenceNumber: state.noteReferenceNumber,
    pageWritingMode: writingModeFromTextDirection(state.sectionLayout.textDirection),
    // §17.6.20 btLr uses the horizontal line model rotated wholesale.
    verticalCJK: state.verticalCJK && !state.verticalAllRotated,
    verticalPageFrame: state.verticalCJK === true,
    documentHasEastAsianText: state.docEastAsian,
    useFeLayout: state.layoutSettings.compat.useFeLayout,
    balanceSingleByteDoubleByteWidth:
      state.layoutSettings.compat.balanceSingleByteDoubleByteWidth,
    characterSpacingControl: state.layoutSettings.characterSpacingControl,
    resolvedLocalFonts: state.resolvedLocalFonts,
    layoutServices: state.layoutServices,
    verticalGlyphMeasurement: state.verticalGlyphMeasurement,
  };
}

/** Build-segment environment for direct acquisition measurement. `btLr` clears
 * upright-vertical grouping because its complete horizontal layout is rotated. */
export function segmentEnvironmentOf(
  state: BodyMeasurementContext,
): LineLayoutEnvironment {
  if (!state.verticalAllRotated
    && state.layoutSettings.characterSpacingControl === undefined
    && !state.layoutSettings.compat.balanceSingleByteDoubleByteWidth) return state;
  return {
    ...state,
    ...(state.verticalAllRotated ? { verticalCJK: false } : {}),
    characterSpacingControl: state.layoutSettings.characterSpacingControl,
    balanceSingleByteDoubleByteWidth:
      state.layoutSettings.compat.balanceSingleByteDoubleByteWidth,
  };
}

/** Snap a uniform paragraph line advance to an integer docGrid pitch. */
export function snapParaLineToGrid(
  heightPx: number,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (!isGridLineRule(grid)) return heightPx;
  const pitchPx = grid!.linePitchPt! * scale;
  if (pitchPx <= 0) return heightPx;
  if (heightPx <= pitchPx) return pitchPx;
  return Math.ceil(heightPx / pitchPx) * pitchPx;
}

/** Project normalized section/paragraph grid participation into line-layout's
 * legacy-shaped grid value without reacquiring parser facts. */
export function gridForParagraphContext(
  state: Pick<BodyMeasurementContext, 'sectionLayout'>,
  context: ParagraphLayoutContext,
): DocGridCtx {
  const characterGrid = paragraphCharacterGrid(context);
  return {
    type: characterGrid
      ? characterGrid.type
      : context.lineGrid.active ? state.sectionLayout.grid.kind : null,
    linePitchPt: context.lineGrid.active ? context.lineGrid.pitchPt : null,
    characterPitchPt: characterGrid?.characterPitchPt ?? null,
    charSpacePt: characterGrid?.charSpacePt ?? null,
  };
}
