import { isSymbolFontFamily, symbolTextToUnicodeSegments } from '@silurus/ooxml-core';
import type { NumberingInfo, TabStop } from '../types.js';
import { wordNumberingSpaceSuffixUsesArial, wordNumberingSuffixAcceptsCoincidentListTab } from './line-compatibility.js';
import { nextTabStop, type TextLayoutService, type TextShapeResult } from './text.js';
import type { DeepReadonly, NumberingMarkerShapeInput } from './types.js';
import { allowsReferenceFontMetric, indexedFontMetrics, selectResourceMetric, selectedFontLineMetric } from '../line-breaker/font-metrics.js';

export interface NumberingMarkerTextLayout {
  readonly shape: TextShapeResult;
  readonly fontSizePx: number;
  readonly lineBox: NumberingMarkerLineBox;
}

export interface NumberingMarkerLineBox {
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly intendedSinglePt: number;
}

export interface NumberingMarkerGeometry {
  readonly bodyOffsetPt: number;
  readonly markerText: string;
  readonly markerWidthPt: number;
  readonly markerShiftPt: number;
  readonly shape: TextShapeResult | null;
  readonly lineBox?: NumberingMarkerLineBox;
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
    compatibilityMode?: number;
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
    compatibilityMode: input.compatibilityMode,
    baseRtl: context.baseRtl,
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
    kerning: input.kerning ?? false,
    measure: true,
    clusterGeometry,
  });
  const metrics = indexedFontMetrics(service.fontMetrics ?? service.localMetrics ?? {});
  let ascentPt = 0;
  let descentPt = 0;
  let intendedSinglePt = 0;
  for (const span of shape.spans) {
    const { resourceMetric, referenceMetric, lineMetric } = selectedFontLineMetric(
      span.font, selectResourceMetric(metrics, span.font, span.text), true,
      allowsReferenceFontMetric(metrics, span.font),
      { text: service, probe: span.text },
    );
    const design = resourceMetric?.lineHeightRatio != null ? resourceMetric : referenceMetric;
    // Body text admits these same selected-face sides. Using raw Canvas boxes
    // only for markers adds device-pixel rounding to otherwise precise lines,
    // including same-face/same-size numbered headings near a keepNext boundary.
    const ownsSides = design?.designAscentRatio != null && design.designDescentRatio != null;
    ascentPt = Math.max(ascentPt, ownsSides
      ? design.designAscentRatio! * input.fontSizePt * scale : span.ascentPt);
    descentPt = Math.max(descentPt, ownsSides
      ? design.designDescentRatio! * input.fontSizePt * scale : span.descentPt);
    const ratio = span.script === 'eastAsia'
      ? lineMetric?.eastAsianLineHeightRatio ?? lineMetric?.lineHeightRatio
      : lineMetric?.lineHeightRatio;
    intendedSinglePt = Math.max(intendedSinglePt, (ratio ?? 0) * input.fontSizePt * scale);
  }
  const lineBox = Object.freeze(shape.spans.length > 0 ? { ascentPt, descentPt, intendedSinglePt }
    : { ascentPt: shape.ascentPt, descentPt: shape.descentPt, intendedSinglePt: 0 });
  return { shape, fontSizePx: input.fontSizePt * scale, lineBox };
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
    /** Observed legacy false-flag numbering window. Ordinary custom/grid
     * coordinates follow this origin; the implicit indent stop does not. */
    suffixFrameOriginPt?: number;
    tabStops: readonly TabStop[];
    defaultTabPt: number;
    compatibilityMode?: number;
    /** Paragraph direction; WORD_NUMBERING_SPACE_SUFFIX_ARIAL is measured for
     * left-to-right paragraphs only, so an absent value keeps the marker face. */
    baseRtl?: boolean;
  }>,
  service: TextLayoutService,
  clusterGeometry = true,
): NumberingMarkerGeometry {
  // A symbol-font marker (§17.9.x w:lvlText in Symbol/Wingdings...) is judged
  // against the font the text service selects for it, exactly as run text
  // (core symbol-font.ts), and is then shaped as ONE run, as every marker is:
  //   - no code converts: the authored codes, in the marker's symbol slots;
  //   - codes convert and the face draws none of the marker's codes (the font
  //     is missing or Unicode-encoded, like Apple's Symbol): the converted
  //     text, in the authored slots as in 0.89 when no character of it would
  //     be drawn there as a symbol code; else with every slot cleared (a
  //     Unicode-capable generic route), because a symbol cmap would re-read
  //     e.g. a converted U+00A9 as its own code 0xA9 (a heart in Symbol). An
  //     emoji-capable target also takes the cleared route, the one its
  //     monochrome proof was made in;
  //   - codes convert while the face draws another code of the same marker
  //     (a face with only part of its encoding): one run cannot take both
  //     routes, so the marker stays as authored, as before this conversion.
  const markerFamily = numbering.fontFamily ?? markerInput.fonts.ascii ?? null;
  const slots = [markerInput.fonts, markerInput.themeFonts].flatMap((fonts) => fonts
    ? [fonts.ascii, fonts.highAnsi, fonts.eastAsia, fonts.complexScript] : []);
  const markerSlots = {
    fonts: markerInput.fonts, themeFonts: markerInput.themeFonts, themeFontPresence: markerInput.themeFontPresence,
    weight: markerInput.weight, style: markerInput.style, complexScript: markerInput.complexScript,
    fontHint: markerInput.fontHint, eastAsiaLanguage: markerInput.eastAsiaLanguage,
  };
  // Slots carry the effective level rPr. Where a slot that shaping selects
  // does not name a symbol font (an older/public model, or a complex-script
  // marker whose cs slot names a text face), the lvlText codes still belong
  // to the marker's symbol font (§17.9.x): the marker is judged, and its kept
  // codes shaped, in that family alone.
  const slotted = service.scalarFamily
    ? [...numbering.text].every((character) => isSymbolFontFamily(service.scalarFamily!(markerSlots, character)))
    : slots.some(isSymbolFontFamily);
  const symbolMarker = !numbering.picBulletImagePath && [markerFamily, ...slots].some(isSymbolFontFamily);
  const keptInput: NumberingMarkerShapeInput = slotted || !symbolMarker || !isSymbolFontFamily(markerFamily) ? markerInput : {
    ...markerInput,
    fonts: { ascii: markerFamily, highAnsi: markerFamily, eastAsia: markerFamily, complexScript: markerFamily },
    themeFonts: undefined,
    themeFontPresence: undefined,
  };
  const keptSlots = { ...markerSlots, fonts: keptInput.fonts, themeFonts: keptInput.themeFonts, themeFontPresence: keptInput.themeFontPresence };
  const judged = !symbolMarker ? [{ text: numbering.text, mapped: false }]
    : service.symbolText?.({ ...keptSlots, text: numbering.text })
      ?? symbolTextToUnicodeSegments(numbering.text, markerFamily);
  const judgedText = judged.map((part) => part.text).join('');
  const drawnAsSymbol = (slotsOf: typeof markerSlots, text: string) =>
    [...text].some((character) => service.symbolDraws?.(slotsOf, character) === true);
  const anyConverted = judged.some((part) => part.mapped);
  const converted = anyConverted && !drawnAsSymbol(keptSlots, numbering.text);
  // A wholly kept marker takes the judged text: a kept bare code may be
  // written F0-shifted there (core symbol-font.ts).
  const markerText = numbering.picBulletImagePath ? '' : converted || !anyConverted ? judgedText : numbering.text;
  const markerLayout = markerText
    ? shapeNumberingMarkerText(
        !converted ? keptInput
          : !markerText.includes('\uFE0E') && !drawnAsSymbol(markerSlots, markerText) ? markerInput
            : { ...markerInput, fonts: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null },
              themeFonts: undefined, themeFontPresence: undefined },
        markerText,
        1,
        service,
        clusterGeometry,
      ) ?? null
    : null;
  const markerShape = markerLayout?.shape ?? null;
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
    // `word-numbering-space-suffix-arial`: §17.9.28 specifies the space,
    // while this compatibility face is independent of every marker slot
    // and theme. Keep the marker size and measure through the same authority,
    // including its resource resolution / metric-compatible fallback.
    const spaceInput = wordNumberingSpaceSuffixUsesArial({
      compatibilityMode: input.compatibilityMode,
      leftToRight: input.baseRtl === false,
      markerWeight: markerInput.weight,
      markerStyle: markerInput.style,
      complexScriptMarker: markerInput.complexScript,
      // A symbol-encoded marker (Private Use Area code points, or text the
      // shared symbol-font table remaps) is not measured.
      symbolMarker: markerText !== numbering.text || /[\uE000-\uF8FF]/u.test(numbering.text),
    })
      ? {
          ...markerInput,
          fonts: { ascii: 'Arial', highAnsi: 'Arial', eastAsia: 'Arial', complexScript: 'Arial' },
          themeFonts: undefined,
          themeFontPresence: undefined,
        }
      : markerInput;
    bodyOffsetPt += shapeNumberingMarkerText(
      spaceInput,
      ' ',
      1,
      service,
      clusterGeometry,
    )?.shape.advancePt ?? 0;
  } else if (suffix === 'tab') {
    bodyOffsetPt = 0;
    // `word-numbering-marker-overflow-tab-advance`
    if (markerEndPt > 0) {
      const origin = input.suffixFrameOriginPt;
      const stop = origin === undefined
        ? numberingSuffixTabStop(input.physicalIndentLeftPt + markerEndPt, input.tabStops, input.defaultTabPt)
        : numberingSuffixTabStop(input.physicalIndentLeftPt + markerEndPt - origin,
          [...input.tabStops, { pos: Math.max(input.physicalIndentLeftPt, origin) - origin, alignment: 'num', leader: 'none' }],
          input.defaultTabPt);
      bodyOffsetPt = stop ? stop.pos + (origin ?? 0) - input.physicalIndentLeftPt : markerEndPt;
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
    ...(markerLayout ? { lineBox: markerLayout.lineBox } : {}),
  };
}
