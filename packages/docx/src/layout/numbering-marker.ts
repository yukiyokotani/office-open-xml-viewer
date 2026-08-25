import { symbolFontToUnicode } from '@silurus/ooxml-core';
import type { NumberingInfo, TabStop } from '../types.js';
import { wordNumberingSuffixAcceptsCoincidentListTab } from './line-compatibility.js';
import { nextTabStop, type TextLayoutService, type TextShapeResult } from './text.js';
import type { DeepReadonly, NumberingMarkerShapeInput } from './types.js';

export interface NumberingMarkerTextLayout {
  readonly shape: TextShapeResult;
  readonly fontSizePx: number;
}

export interface NumberingMarkerGeometry {
  readonly bodyOffsetPt: number;
  readonly markerText: string;
  readonly markerWidthPt: number;
  readonly markerShiftPt: number;
  readonly shape: TextShapeResult | null;
}

/** Marker interval in the paragraph's logical-leading coordinate system. */
export function numberingMarkerLogicalInterval(input: Readonly<{
  leadingIndentPt: number;
  authoredFirstIndentPt: number;
  markerShiftPt: number;
  markerWidthPt: number;
}>): Readonly<{ startPt: number; endPt: number }> {
  const startPt = input.leadingIndentPt
    + input.authoredFirstIndentPt
    + input.markerShiftPt;
  return { startPt, endPt: startPt + input.markerWidthPt };
}

/** Convert the same logical-leading marker interval to physical page X. Keeping
 * this conversion beside intrinsic marker geometry prevents RTL auto-width
 * acquisition and final retained placement from choosing different origins. */
export function numberingMarkerPhysicalLeft(input: Readonly<{
  baseRtl: boolean;
  alignedLeadingEdgePt: number;
  authoredFirstIndentPt: number;
  markerShiftPt: number;
  markerWidthPt: number;
}>): number {
  const logicalStartPt = input.authoredFirstIndentPt + input.markerShiftPt;
  return input.baseRtl
    ? input.alignedLeadingEdgePt - logicalStartPt - input.markerWidthPt
    : input.alignedLeadingEdgePt + logicalStartPt;
}

/** Apply retained marker geometry to an otherwise parser-independent paragraph context. */
export function applyNumberingBodyOffset<Context extends Readonly<{
  baseRtl: boolean;
  firstIndentPt: number;
  physicalIndentLeftPt: number;
  defaultTabPt: number;
}>>(
  context: Context,
  input: Readonly<{
    numbering: DeepReadonly<NumberingInfo> | null;
    markerInput?: NumberingMarkerShapeInput;
    authoredFirstIndentPt: number;
    tabStops: readonly TabStop[];
    defaultTabPt?: number;
    service?: TextLayoutService;
    clusterGeometry?: boolean;
  }>,
): Context {
  const { numbering, markerInput, service } = input;
  const hasMarker = numbering != null
    && (numbering.text !== '' || numbering.picBulletImagePath != null);
  const usesResolvedBodyOffset = hasMarker
    && (!context.baseRtl
      || ((numbering?.suff || 'tab') === 'tab' && input.authoredFirstIndentPt < 0));
  if (!numbering || !markerInput || !service || !usesResolvedBodyOffset) return context;
  const geometry = resolveNumberingMarkerGeometry(numbering, markerInput, {
    authoredFirstIndentPt: input.authoredFirstIndentPt,
    physicalIndentLeftPt: context.physicalIndentLeftPt,
    tabStops: input.tabStops,
    defaultTabPt: input.defaultTabPt ?? context.defaultTabPt,
  }, service, input.clusterGeometry ?? true);
  return {
    ...context,
    firstIndentPt: geometry.bodyOffsetPt,
    numberingMarkerGeometry: geometry,
  };
}

/** Shape numbering text through the document's one font authority. ECMA-376
 * §17.9.6 applies the level rPr to the marker, while §17.3.2.26 selects a
 * slot for each scalar; a mixed marker therefore cannot be represented by one
 * leading-code-point family. Older parser models still enter the service using
 * their public ascii/eastAsia projection, but selection and exact Canvas routes
 * remain owned by TextLayoutService. */
export function shapeNumberingMarkerText(
  input: NumberingMarkerShapeInput,
  text: string,
  scale: number,
  service: TextLayoutService | undefined,
  clusterGeometry = true,
): NumberingMarkerTextLayout | null {
  if (!service) return null;
  const shape = service.shape({
    text,
    fontSizePt: input.fontSizePt * scale,
    fonts: input.fonts,
    themeFonts: input.themeFonts,
    themeFontPresence: input.themeFontPresence,
    weight: input.weight,
    style: input.style,
    complexScript: input.complexScript,
    fontHint: input.fontHint,
    eastAsiaLanguage: input.eastAsiaLanguage,
    kerning: input.kerning,
    measure: true,
    clusterGeometry,
  });
  return { shape, fontSizePx: input.fontSizePt * scale };
}

/** Resolve the tab synthesized by w:suff="tab". ST_TabJc `num` identifies the
 * list tab between the marker and paragraph contents. The coincident-stop branch
 * is governed by WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB; ordinary tab
 * characters retain the strictly-forward {@link nextTabStop} rule. */
function numberingSuffixTabStop(
  currentMarginPt: number,
  tabStops: readonly TabStop[],
  defaultTabPt: number,
): Readonly<{ pos: number }> | null {
  const coincidentListStop = tabStops.find((stop) =>
    wordNumberingSuffixAcceptsCoincidentListTab(currentMarginPt, stop));
  return coincidentListStop
    ?? nextTabStop(currentMarginPt, [...tabStops], defaultTabPt);
}

/** Resolve the marker and suffix once in document points. The body offset is a
 * measurement input, not paint-time decoration: line breaking and retained
 * placement must consume this same value or a hanging list acquires a different
 * first-line partition from the one it paints. */
export function resolveNumberingMarkerGeometry(
  numbering: DeepReadonly<NumberingInfo>,
  markerInput: NumberingMarkerShapeInput,
  input: Readonly<{
    authoredFirstIndentPt: number;
    physicalIndentLeftPt: number;
    tabStops: readonly TabStop[];
    defaultTabPt: number;
  }>,
  service: TextLayoutService,
  clusterGeometry = true,
): NumberingMarkerGeometry {
  const markerText = numbering.picBulletImagePath
    ? ''
    : symbolFontToUnicode(numbering.text, numbering.fontFamily ?? null);
  const markerShape = markerText
    ? shapeNumberingMarkerText(
        markerInput,
        markerText,
        1,
        service,
        clusterGeometry,
      )?.shape ?? null
    : null;
  const markerWidthPt = numbering.picBulletImagePath
    ? numbering.picBulletWidthPt ?? markerInput.fontSizePt
    : markerShape?.advancePt ?? 0;
  const markerShiftPt = numbering.jc === 'right'
    ? -markerWidthPt
    : numbering.jc === 'center' ? -markerWidthPt / 2 : 0;
  const markerEndPt = input.authoredFirstIndentPt + markerShiftPt + markerWidthPt;
  const suffix = numbering.suff || 'tab';
  let bodyOffsetPt = markerEndPt;
  if (suffix === 'space') {
    bodyOffsetPt += shapeNumberingMarkerText(
      markerInput,
      ' ',
      1,
      service,
      clusterGeometry,
    )?.shape.advancePt ?? 0;
  } else if (suffix === 'tab') {
    bodyOffsetPt = 0;
    if (markerEndPt > 0) {
      const stop = numberingSuffixTabStop(
        input.physicalIndentLeftPt + markerEndPt,
        input.tabStops,
        input.defaultTabPt,
      );
      bodyOffsetPt = stop ? stop.pos - input.physicalIndentLeftPt : markerEndPt;
    }
  }
  // ECMA-376 §17.3.1.12 keeps paragraph content at the authored start indent.
  // The §17.9.28 suffix advances that body only when the marker plus suffix
  // crosses the start edge; a marker contained by its hanging region cannot
  // pull the paragraph body backward into that region.
  bodyOffsetPt = Math.max(0, bodyOffsetPt);
  return {
    bodyOffsetPt,
    markerText,
    markerWidthPt,
    markerShiftPt,
    shape: markerShape,
  };
}
