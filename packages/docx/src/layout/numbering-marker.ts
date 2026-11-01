import { symbolFontToUnicode } from '@silurus/ooxml-core';
import type { NumberingInfo, TabStop } from '../types.js';
import { wordNumberingSuffixAcceptsCoincidentListTab } from './line-compatibility.js';
import { nextTabStop, type TextLayoutService, type TextShapeResult } from './text.js';
import type { DeepReadonly, NumberingMarkerShapeInput } from './types.js';
import type { NumberingBodyIndentRegime } from '../line-breaker/model.js';
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
    tabStops: readonly TabStop[];
    defaultTabPt: number;
  }>,
  service: TextLayoutService,
  clusterGeometry = true,
): NumberingMarkerGeometry {
  const markerText = numbering.picBulletImagePath
    ? ''
    : symbolFontToUnicode(numbering.text, numbering.fontFamily ?? null);
  const markerLayout = markerText
    ? shapeNumberingMarkerText(
        markerInput,
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
    bodyOffsetPt += shapeNumberingMarkerText(
      markerInput,
      ' ',
      1,
      service,
      clusterGeometry,
    )?.shape.advancePt ?? 0;
  } else if (suffix === 'tab') {
    bodyOffsetPt = 0;
    // `word-numbering-marker-overflow-tab-advance`
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
    ...(markerLayout ? { lineBox: markerLayout.lineBox } : {}),
  };
}

/** The regimes of {@link resolveNumberingMarkerGeometry}'s body offset over
 * the marker's authored leading edge `e` (the other inputs fixed), for a
 * line breaker that must split a sloped window at each change of form. The
 * marker extent k (shift plus width) is shaped once:
 * - suffix space / nothing: body = e + c (c = k plus the space advance)
 *   while that is positive, else 0;
 * - suffix tab: 0 while e + k <= 0; else the suffix stop (a constant) or,
 *   with no stop, e + k. The stop changes only where the margin e + k crosses
 *   a custom stop, the coincident list-stop tolerance, the last custom stop,
 *   or a default-grid position (nextTabStop's own thresholds). */
export function numberingBodyIndentRegimes(
  numbering: DeepReadonly<NumberingInfo>,
  markerInput: NumberingMarkerShapeInput,
  input: Readonly<{
    physicalIndentLeftPt: number;
    tabStops: readonly TabStop[];
    defaultTabPt: number;
  }>,
  service: TextLayoutService,
): {
  regimeAt: (edgePt: number) => NumberingBodyIndentRegime;
  breaksIn: (lowerPt: number, upperPt: number, limit: number) => readonly number[] | null;
} {
  const atZero = resolveNumberingMarkerGeometry(numbering, markerInput, { ...input, authoredFirstIndentPt: 0 }, service);
  const extent = atZero.markerShiftPt + atZero.markerWidthPt;
  const { physicalIndentLeftPt: physicalLeft, tabStops, defaultTabPt } = input;
  if ((numbering.suff || 'tab') !== 'tab') {
    const offset = atZero.bodyOffsetPt;
    return {
      regimeAt: (edge) => edge + offset > 0 ? { key: 'offset', offsetPx: offset } : { key: 'zero', constantPx: 0 },
      breaksIn: (lower, upper) => -offset > lower && -offset < upper ? [-offset] : [],
    };
  }
  const placing = tabStops.filter((stop) => stop.alignment !== 'bar');
  const lastCustom = placing.reduce((last, stop) => Math.max(last, stop.pos), 0);
  // Margins (physical, from physicalLeft) at which the suffix stop may change.
  const fixedMargins = [lastCustom];
  for (const stop of placing) {
    fixedMargins.push(stop.pos);
    if (stop.alignment === 'num') fixedMargins.push(stop.pos - 1e-6, stop.pos + 1e-6);
  }
  const gridEpsilon = 1e-6;
  return {
    regimeAt: (edge) => {
      const markerEnd = edge + atZero.markerShiftPt + atZero.markerWidthPt;
      if (!(markerEnd > 0)) return { key: 'zero', constantPx: 0 };
      const stop = numberingSuffixTabStop(physicalLeft + markerEnd, tabStops, defaultTabPt);
      return stop
        ? { key: `stop:${stop.pos}`, constantPx: Math.max(0, stop.pos - physicalLeft) }
        : { key: 'marker', offsetPx: extent };
    },
    breaksIn: (lower, upper, limit) => {
      const toEdge = (margin: number): number => margin - physicalLeft - extent;
      const edges = [-extent, ...fixedMargins.map(toEdge)];
      if (defaultTabPt > 0) {
        const from = Math.max(physicalLeft + lower + extent, lastCustom);
        const to = physicalLeft + upper + extent;
        const first = Math.ceil((from + gridEpsilon) / defaultTabPt);
        const last = Math.floor((to + gridEpsilon) / defaultTabPt);
        if (last - first + 1 > limit) return null;
        for (let index = first; index <= last; index += 1) edges.push(toEdge(index * defaultTabPt - gridEpsilon));
      }
      const inside = [...new Set(edges.filter((edge) => edge > lower && edge < upper))].sort((a, b) => a - b);
      return inside.length > limit ? null : inside;
    },
  };
}
