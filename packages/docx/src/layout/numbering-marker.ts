import { graphemeClusterOffsets, isSymbolFontFamily, symbolTextToUnicodeSegments, type SymbolTextSegment } from '@silurus/ooxml-core';
import { computeSegmentLineVisualOrder } from '@silurus/ooxml-core/internal/bidi-line';
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
  /** Paint direction of each `shape.spans` entry when a symbol marker was
   * joined from pieces in a right-to-left paragraph (UAX #9 resolved levels);
   * absent: every span paints in the paragraph direction. */
  readonly spanDirections?: readonly ('ltr' | 'rtl')[];
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

/** A marker piece: one shaping run, at logical UTF-16 `start`, painted in
 * direction `rtl`. */
interface MarkerPiece extends SymbolTextSegment {
  readonly start: number;
  readonly rtl: boolean;
}

/**
 * The pieces of a symbol marker in a right-to-left paragraph, in visual order.
 * UAX #9 (core segment-line order, as body segments are) resolves every
 * grapheme of the marker; consecutive visual graphemes that are logically
 * adjacent, share a resolved direction and a conversion disposition form one
 * piece. So no piece crosses an embedding-level boundary, a converted "("
 * paints in its resolved direction, a VS15 stays with its base, and Arabic
 * keeps its joining. Limitation: explicit embedding, override and isolate
 * controls (U+202A..U+202E, U+2066..U+2069) inside such a marker are not
 * modelled; each becomes its own piece at the paragraph level.
 */
function rtlMarkerPieces(segments: readonly SymbolTextSegment[]): MarkerPiece[] {
  // Graphemes of the whole marker (an authored VS15 or combining mark joins
  // its base across dispositions); each takes its base's disposition.
  const text = segments.map((segment) => segment.text).join('');
  const owners: SymbolTextSegment[] = [];
  for (const segment of segments) for (let unit = 0; unit < segment.text.length; unit++) owners.push(segment);
  const cuts = [0, ...graphemeClusterOffsets(text), text.length];
  const graphemes: (SymbolTextSegment & { start: number })[] = [];
  for (let index = 0; index + 1 < cuts.length; index++) {
    const from = cuts[index]!;
    if (from === cuts[index + 1]) continue;
    const owner = owners[from]!;
    graphemes.push({ text: text.slice(from, cuts[index + 1]), mapped: owner.mapped,
      ...(owner.unicodeFace ? { unicodeFace: true as const } : {}), start: from });
  }
  const resolved = computeSegmentLineVisualOrder(graphemes.map((grapheme) => ({ text: grapheme.text })), true);
  const pieces: (MarkerPiece & { last: number })[] = [];
  for (const index of resolved.order) {
    const grapheme = graphemes[index]!;
    // `rtl` is indexed by logical segment.
    const rtl = resolved.rtl[index] === true;
    const previous = pieces.at(-1);
    if (previous && previous.mapped === grapheme.mapped && previous.unicodeFace === grapheme.unicodeFace
      && previous.rtl === rtl && previous.last === (rtl ? index + 1 : index - 1)) {
      pieces[pieces.length - 1] = rtl
        ? { ...previous, text: grapheme.text + previous.text, start: grapheme.start, last: index }
        : { ...previous, text: previous.text + grapheme.text, last: index };
    } else {
      pieces.push({ ...grapheme, rtl, last: index });
    }
  }
  return pieces.map(({ last: _last, ...piece }) => piece);
}

/** Join marker pieces shaped with different slots (a symbol marker whose
 * selected face draws some codes, while others are converted) into one
 * marker shape. Pieces arrive in visual order with their logical starts:
 * ranges stay logical, while pen offsets, clusters and ink move by the
 * pieces placed before them. Marker paint draws spans left to right in array
 * order. The line box takes the largest sides. */
function joinMarkerLayouts(
  parts: readonly Readonly<{ start: number; rtl: boolean; text: string; layout: NumberingMarkerTextLayout }>[],
  withDirections: boolean,
): NumberingMarkerTextLayout & { readonly spanDirections?: readonly ('ltr' | 'rtl')[] } {
  const logicalOffset = parts.reduce((sum, part) => sum + part.text.length, 0);
  const spanDirections: ('ltr' | 'rtl')[] = [];
  const spans: TextShapeResult['spans'][number][] = [];
  const clusters: NonNullable<TextShapeResult['clusters']>[number][] = [];
  const boundaries = new Set<number>([0]);
  const diagnostics: TextShapeResult['diagnostics'][number][] = [];
  let advancePt = 0;
  let withClusters = true;
  let ink: { xMinPt: number; xMaxPt: number; ascentPt: number; descentPt: number } | undefined;
  let withInk = true;
  for (const { start: offset, rtl, layout: { shape } } of parts) {
    // A right-to-left piece paints its own spans right to left.
    for (const span of rtl ? [...shape.spans].reverse() : shape.spans) {
      spanDirections.push(rtl ? 'rtl' : 'ltr');
      spans.push({
        ...span,
        start: span.start + offset,
        end: span.end + offset,
        ...(span.semanticSlotSpans ? {
          semanticSlotSpans: span.semanticSlotSpans.map((slot) => ({ ...slot, start: slot.start + offset, end: slot.end + offset })),
        } : {}),
      });
    }
    if (shape.clusters) {
      for (const cluster of shape.clusters) {
        clusters.push({
          range: { start: cluster.range.start + offset, end: cluster.range.end + offset },
          // A right-to-left piece is laid out mirrored: its first logical
          // cluster sits at its right edge.
          offsetPt: advancePt + (rtl ? shape.advancePt - cluster.offsetPt - cluster.advancePt : cluster.offsetPt),
          advancePt: cluster.advancePt,
        });
      }
    } else withClusters = false;
    for (const boundary of shape.graphemeBoundaries) boundaries.add(boundary + offset);
    diagnostics.push(...shape.diagnostics);
    if (shape.inkBounds && withInk) {
      const moved = rtl
        ? { ...shape.inkBounds, xMinPt: advancePt + shape.advancePt - shape.inkBounds.xMaxPt, xMaxPt: advancePt + shape.advancePt - shape.inkBounds.xMinPt }
        : { ...shape.inkBounds, xMinPt: shape.inkBounds.xMinPt + advancePt, xMaxPt: shape.inkBounds.xMaxPt + advancePt };
      ink = ink ? {
        xMinPt: Math.min(ink.xMinPt, moved.xMinPt), xMaxPt: Math.max(ink.xMaxPt, moved.xMaxPt),
        ascentPt: Math.max(ink.ascentPt, moved.ascentPt), descentPt: Math.max(ink.descentPt, moved.descentPt),
      } : moved;
    } else withInk = false;
    advancePt += shape.advancePt;
  }
  boundaries.add(logicalOffset);
  const shapes = parts.map((part) => part.layout);
  const shape: TextShapeResult = {
    spans,
    advancePt,
    ascentPt: Math.max(...shapes.map((layout) => layout.shape.ascentPt)),
    descentPt: Math.max(...shapes.map((layout) => layout.shape.descentPt)),
    graphemeBoundaries: [...boundaries].sort((a, b) => a - b),
    ...(withClusters ? { clusters } : {}),
    ...(withInk && ink ? { inkBounds: ink } : {}),
    diagnostics,
  };
  return {
    shape,
    ...(withDirections ? { spanDirections } : {}),
    fontSizePx: shapes[0]!.fontSizePx,
    lineBox: {
      ascentPt: Math.max(...shapes.map((layout) => layout.lineBox.ascentPt)),
      descentPt: Math.max(...shapes.map((layout) => layout.lineBox.descentPt)),
      intendedSinglePt: Math.max(...shapes.map((layout) => layout.lineBox.intendedSinglePt)),
    },
  };
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
  // against the face the text service selects for it, exactly as run text
  // (core symbol-font.ts). Converted characters are Unicode text: shaping
  // them with the marker's symbol family would let its symbol cmap re-read
  // e.g. U+00A9 as its own code 0xA9 (a heart in Symbol), so a converted
  // marker is shaped with every authored slot cleared (a Unicode-capable
  // generic route), like a converted run segment; the codes the face keeps
  // are shaped in the marker's slots, and the pieces are joined in order.
  // Converted text from a face known to draw no symbol code (missing, or
  // Unicode-encoded like Apple's Symbol) keeps the slots when that is safe,
  // as in 0.89 (see safeInSlots below).
  const markerFamily = numbering.fontFamily ?? markerInput.fonts.ascii ?? null;
  const slots = [markerInput.fonts, markerInput.themeFonts].flatMap((fonts) => fonts
    ? [fonts.ascii, fonts.highAnsi, fonts.eastAsia, fonts.complexScript] : []);
  // Slots carry the effective level rPr; a model whose slots do not name the
  // marker's symbol family (older/public models) judges that family alone.
  // A slot whose family is not the marker's symbol font (say a complex-script
  // marker whose cs slot names a text face) would draw the codes as plain
  // text, while Word draws the lvlText codes in that symbol font; such a
  // marker is judged, and its kept codes shaped, in the marker family alone.
  const markerSlots = {
    fonts: markerInput.fonts, themeFonts: markerInput.themeFonts, themeFontPresence: markerInput.themeFontPresence,
    weight: markerInput.weight, style: markerInput.style, complexScript: markerInput.complexScript,
    fontHint: markerInput.fontHint, eastAsiaLanguage: markerInput.eastAsiaLanguage,
  };
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
  const symbolRequest = { ...keptSlots, text: numbering.text };
  const judged = !symbolMarker
    ? null
    : service.symbolText?.(symbolRequest) ?? symbolTextToUnicodeSegments(numbering.text, markerFamily);
  // Each code is converted on its own, as in body runs. A marker of one
  // disposition is one run (Canvas orders it, bidi included). A mixed one is
  // shaped piece by piece: in logical order in a left-to-right paragraph, and
  // as UAX #9 level runs in visual order in a right-to-left one
  // (rtlMarkerPieces).
  const mixed = judged !== null && judged.length > 1;
  const pieces: MarkerPiece[] | null = !judged ? null
    : mixed && input.baseRtl === true ? rtlMarkerPieces(judged)
      : (() => {
        let start = 0;
        return judged.map((part) => {
          const piece = { ...part, start, rtl: false };
          start += part.text.length;
          return piece;
        });
      })();
  const markerSegments = pieces;
  const markerText = numbering.picBulletImagePath
    ? ''
    : judged ? judged.map((part) => part.text).join('') : numbering.text;
  const convertedInput: NumberingMarkerShapeInput = {
    ...markerInput,
    fonts: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null },
    themeFonts: undefined,
    themeFontPresence: undefined,
  };
  const markerParts: readonly MarkerPiece[] = markerSegments ?? [{ text: markerText, mapped: false, start: 0, rtl: false }];
  // Converted text may keep the marker slots only when no character of it
  // takes a symbol face there that draws symbol codes: the slot shaping
  // selects for it names a non-symbol family, or a symbol family whose face
  // draws none (judged with a probe code). Otherwise (say "(" taking an
  // installed Wingdings in another slot) it is shaped without the slots, as
  // is an emoji-capable one: its monochrome proof is for the generic route.
  const safeInSlots = (text: string) => !text.includes('\uFE0E') && [...text].every((character) => {
    const family = service.scalarFamily?.(markerSlots, character);
    if (family === undefined) return false;
    if (!isSymbolFontFamily(family)) return true;
    const probe = service.symbolText?.({ ...symbolRequest, text: '\uF0B7',
      fonts: { ascii: family, highAnsi: family, eastAsia: family, complexScript: family },
      themeFonts: undefined, themeFontPresence: undefined });
    return probe?.length === 1 && probe[0]!.unicodeFace === true;
  });
  const partLayouts = markerText
    ? markerParts.map((part) => shapeNumberingMarkerText(
        !part.mapped ? keptInput : part.unicodeFace && safeInSlots(part.text) ? markerInput : convertedInput,
        part.text, 1, service, clusterGeometry,
      ))
    : [];
  const markerLayout: (NumberingMarkerTextLayout & { readonly spanDirections?: readonly ('ltr' | 'rtl')[] }) | null = partLayouts.length === 0 || partLayouts.some((layout) => !layout)
    ? null
    : partLayouts.length === 1
      ? partLayouts[0]!
      : joinMarkerLayouts(markerParts.map((part, index) => ({ ...part, layout: partLayouts[index]! })),
        input.baseRtl === true);
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
    ...(markerLayout?.spanDirections ? { spanDirections: markerLayout.spanDirections } : {}),
  };
}
