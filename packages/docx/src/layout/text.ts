import type { LayoutDiagnostic } from './types.js';
import {
  classifyFontGeneric,
  graphemeClusterOffsets,
  normalizeFontMetricFamily,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import type {
  FontResolution,
  FontResolver,
  FontStyle,
} from './font-service.js';
import type { CanvasFontRoute } from '@silurus/ooxml-core';
export type { ResolvedFontMetric, ResolvedLocalFontMetric } from '@silurus/ooxml-core';
import { stableFingerprint } from './fingerprint.js';
import type {
  DocParagraph,
  DocRun,
  DocxTextRun,
  FieldRun,
  ShapeTextRun,
  TabStop,
} from '../types.js';
import type {
  DeepReadonly,
  NumberingMarkerShapeInput,
  SourceRef,
  VmlTextPathAcquisitionInput,
} from './types.js';
import type { TextBoxAcquisitionInput } from './textbox-input.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';

/** The subset of a measured text segment needed to resolve its effective size. */
export interface EffectiveFontSegment {
  readonly fontSize: number;
  readonly smallCaps?: boolean;
  readonly vertAlign?: 'super' | 'sub' | null;
}

export type ResolvedTabStop = Readonly<{
  pos: number;
  alignment: TabStop['alignment'];
  leader?: TabStop['leader'];
}>;

/** Internal line-acquisition flags layered onto the stable public text-run model. */
export type ShapeTextDocRun = Extract<DocRun, { type: 'text' }> & Readonly<{
  textBoxLineFloor: true;
  textBoxVertical: boolean;
}>;

/** ECMA-376 §17.3.2.33 effective size in the caller's scaled coordinate space. */
export function calcEffectiveFontPx(segment: EffectiveFontSegment, scale: number): number {
  const fontSizePt = segment.smallCaps ? Math.max(segment.fontSize - 2, 1) : segment.fontSize;
  let size = fontSizePt * scale;
  if (segment.vertAlign) size *= 0.65;
  return size;
}

/** East Asian content predicate used by both legacy and retained line acquisition. */
export const EAST_ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟　-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿＀-￯]/u;

/** ECMA-376 §17.3.1.37 and §17.15.1.25 tab-stop resolution. */
export function nextTabStop(
  curMarginPx: number,
  customStopsPx: readonly ResolvedTabStop[],
  intervalPx: number,
): ResolvedTabStop | null {
  let custom: ResolvedTabStop | null = null;
  let maxCustomPx = 0;
  for (const stop of customStopsPx) {
    // ECMA-376 §17.18.84: bar draws a vertical rule but "does not result in a
    // custom tab stop" and is skipped when positioning a tab character.
    if (stop.alignment === 'bar') continue;
    if (stop.pos > maxCustomPx) maxCustomPx = stop.pos;
    if (stop.pos > curMarginPx && (custom === null || stop.pos < custom.pos)) custom = stop;
  }

  let automatic: ResolvedTabStop | null = null;
  if (intervalPx > 0) {
    const epsilon = 1e-6;
    const from = Math.max(curMarginPx, maxCustomPx);
    let pos = Math.ceil((from + epsilon) / intervalPx) * intervalPx;
    if (pos <= curMarginPx) pos += intervalPx;
    automatic = { pos, alignment: 'left' };
  }

  if (custom && automatic) return custom.pos <= automatic.pos ? custom : automatic;
  return custom ?? automatic;
}

/** RTL tab coordinates use the same distance-from-leading-edge stop grid. */
export function nextTabStopRtl(
  curMarginPx: number,
  customStopsPx: readonly ResolvedTabStop[],
  intervalPx: number,
): ResolvedTabStop | null {
  return nextTabStop(curMarginPx, customStopsPx, intervalPx);
}

/** Adapt public shape text to the neutral body-run contract before line acquisition. */
export function shapeRunToDocRun(
  run: ShapeTextRun,
  textVert?: string | null,
): ShapeTextDocRun {
  const textBoxVertical = textVert === 'vert' || textVert === 'vert270'
    || textVert === 'eaVert' || textVert === 'mongolianVert';
  return {
    type: 'text',
    text: run.text,
    bold: run.bold ?? false,
    italic: run.italic ?? false,
    underline: false,
    strikethrough: false,
    fontSize: run.fontSizePt,
    color: run.color ?? null,
    fontFamily: run.fontFamily ?? null,
    fontFamilyEastAsia: run.fontFamilyEastAsia ?? null,
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
    ruby: run.ruby ?? undefined,
    textBoxLineFloor: true,
    textBoxVertical,
  };
}

/** Plain parser-boundary snapshot used by retained line acquisition. Private
 * parser extensions are copied into named immutable fields exactly once. */
type ParagraphTextFacts = Readonly<{
  /** Parser-projected CT_R boundary constraint around an authored
   * `<w:noBreakHyphen/>`. These names are layout facts, not parser wire keys. */
  noBreakBefore?: boolean;
  noBreakAfter?: boolean;
  noBreakRanges?: readonly Readonly<{ start: number; end: number }>[];
  fontFamilyHighAnsi?: string | null;
  fontFamilyEastAsia?: string | null;
  fontHint?: 'default' | 'eastAsia' | 'cs';
  rtl?: boolean;
  cs?: boolean;
  fontFamilyCs?: string | null;
  fontSizeCs?: number;
  boldCs?: boolean;
  italicCs?: boolean;
  langBidi?: string;
  langEastAsia?: string;
  fontSlots?: Readonly<{
    direct: TextFontSlots;
    theme: TextFontSlots;
    themePresent: TextFontSlotPresence;
  }>;
}>;

export type ParagraphTextBearingRun =
  | (DeepReadonly<Extract<DocRun, { type: 'text' }>> & ParagraphTextFacts)
  | (DeepReadonly<Extract<DocRun, { type: 'field' }>> &
    DeepReadonly<Partial<DocxTextRun>> & ParagraphTextFacts);

export type ParagraphMathRun = Readonly<{
  type: 'math';
  display: boolean;
  fontSize: number;
  jc?: string;
  source: SourceRef;
  resourceKey: string;
  fallbackText: string;
}>;

export type ParagraphShapeRun = DeepReadonly<Extract<DocRun, { type: 'shape' }>> & Readonly<{
  vmlTextPathInput?: VmlTextPathAcquisitionInput;
  textBoxInput?: TextBoxAcquisitionInput;
  anchorAcquisitionInput?: AnchorAcquisitionInput;
}>;

export type ParagraphImageRun = DeepReadonly<Extract<DocRun, { type: 'image' }>> & Readonly<{
  anchorAcquisitionInput?: AnchorAcquisitionInput;
}>;

export type ParagraphChartRun = DeepReadonly<Omit<Extract<DocRun, { type: 'chart' }>, 'chart'>> & Readonly<{
  resourceKey: string;
  anchorAcquisitionInput?: AnchorAcquisitionInput;
}>;

/** Parser-private placeholder for a recognized DrawingML payload whose package
 * resource could not be resolved. Authored geometry survives acquisition so
 * surrounding flow and anchor ownership remain deterministic, but this type is
 * intentionally excluded from the public `DocRun` contract. */
export interface UnavailableDrawingAcquisitionRun {
  readonly type: 'unavailableDrawing';
  readonly resourceKind: 'image' | 'chart';
  readonly widthPt: number;
  readonly heightPt: number;
  readonly anchorAcquisitionInput?: AnchorAcquisitionInput;
}

type ParagraphAnchorHostRun = DeepReadonly<Extract<DocRun, { type: 'anchorHost' }>> & Readonly<{
  anchorOccurrenceId?: string;
}>;

export type ParagraphAcquisitionRun =
  | ParagraphTextBearingRun
  | DeepReadonly<Exclude<DocRun, { type: 'text' } | { type: 'field' } | { type: 'math' } | { type: 'shape' } | { type: 'image' } | { type: 'chart' } | { type: 'anchorHost' } | { type: 'unavailableDrawing' }>>
  | ParagraphImageRun
  | ParagraphChartRun
  | ParagraphAnchorHostRun
  | ParagraphShapeRun
  | ParagraphMathRun
  | UnavailableDrawingAcquisitionRun;

/** Immutable parser-boundary event describing one edge of a complex field's
 * cached result interval (ECMA-376 §17.16). The interval is structural only:
 * it does not add a flow segment or affect measurement until a dedicated field
 * compatibility rule explicitly consumes it. */
export interface ComplexFieldBoundaryInput {
  readonly occurrenceKey: string;
  readonly boundary: 'start' | 'end';
  readonly runIndex: number;
  readonly fieldType: 'ref' | 'pageRef' | 'other';
  readonly instruction: string;
  readonly hyperlinkAnchor?: string;
}

export type ParagraphAcquisitionInput = DeepReadonly<Omit<DocParagraph, 'runs'>> & Readonly<{
  runs: readonly ParagraphAcquisitionRun[];
  complexFieldBoundaries?: readonly ComplexFieldBoundaryInput[];
  numberingMarkerShapeInput?: NumberingMarkerShapeInput;
  paragraphMarkShapeInput?: NumberingMarkerShapeInput;
}>;

/** Read-only paragraph shape accepted by layout-only policy helpers. Public
 * hand-built paragraphs and canonical acquisition paragraphs both satisfy this
 * structural contract; production always supplies the canonical arm. */
export type ParagraphLayoutSource = DeepReadonly<Omit<DocParagraph, 'runs'>> & Readonly<{
  runs: readonly (DeepReadonly<DocRun> | ParagraphAcquisitionRun)[];
  complexFieldBoundaries?: readonly ComplexFieldBoundaryInput[];
  numberingMarkerShapeInput?: NumberingMarkerShapeInput;
  paragraphMarkShapeInput?: NumberingMarkerShapeInput;
}>;

export type ParagraphLayoutRun = ParagraphLayoutSource['runs'][number];

export type FontScriptSlot = 'ascii' | 'highAnsi' | 'eastAsia' | 'complexScript';

export interface TextFontSlots {
  readonly ascii?: string | null;
  readonly highAnsi?: string | null;
  readonly eastAsia?: string | null;
  readonly complexScript?: string | null;
}

export interface TextFontSlotPresence {
  readonly ascii?: boolean;
  readonly highAnsi?: boolean;
  readonly eastAsia?: boolean;
  readonly complexScript?: boolean;
}

export interface TextShapeRequest {
  readonly text: string;
  readonly fontSizePt: number;
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly themeFontPresence?: TextFontSlotPresence;
  readonly weight?: number;
  readonly style?: FontStyle;
  readonly complexScript?: boolean;
  /** ECMA-376 §17.3.2.26 rFonts@hint after style inheritance. */
  readonly fontHint?: 'default' | 'eastAsia' | 'cs';
  /** Resolved w:lang@eastAsia, normalized to lower case. */
  readonly eastAsiaLanguage?: string;
  /** fontTable w:charset for the selected eastAsia face (hex byte). */
  readonly eastAsiaFontCharset?: string;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly letterSpacingPt?: number;
  /** Resolved §17.3.2.19 w:kern state at this run size. Absent preserves the
   * measurement adapter's inherited kerning policy. */
  readonly kerning?: boolean;
  /** Resolve script slots and faces without touching the measurement adapter. */
  readonly measure?: boolean;
  /** Aggregate-only acquisition may omit per-grapheme contextual advances.
   * Script spans and aggregate metrics remain authoritative. */
  readonly clusterGeometry?: boolean;
}

export interface TextFontResolveRequest {
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly themeFontPresence?: TextFontSlotPresence;
  readonly slot: FontScriptSlot;
  readonly weight?: number;
  readonly style?: FontStyle;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
}

export interface GlyphMeasureRequest {
  readonly text: string;
  readonly fontRoute: CanvasFontRoute;
  readonly fontSizePt: number;
  readonly weight: number;
  readonly style: FontStyle;
  readonly letterSpacingPt: number;
  readonly kerning?: boolean;
}

export interface GlyphMeasurement {
  /** May be negative when authored character spacing intentionally overlaps glyphs. */
  readonly advancePt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
  /** Tight glyph ink relative to the run origin and alphabetic baseline.
   * Unlike advance/ascent/descent, this can describe ink from a zero-advance
   * combining mark and excludes the font's reserved ascender/descender space. */
  readonly inkBounds?: GlyphInkBounds;
  /** True only when the horizontal ink edges came from the measurement backend
   * rather than an advance-width fallback. */
  readonly horizontalInkBoundsAreTight?: boolean;
}

export interface GlyphInkBounds {
  readonly xMinPt: number;
  readonly xMaxPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
}

export interface GlyphMeasurer {
  readonly fingerprint: string;
  measure(request: Readonly<GlyphMeasureRequest>): GlyphMeasurement;
}

export interface TextShapeSpan extends GlyphMeasurement {
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly script: FontScriptSlot;
  /** False when this scalar span continues the preceding grapheme cluster. */
  readonly breakBefore: boolean;
  readonly font: FontResolution;
  readonly fontRoute: CanvasFontRoute;
}

export interface TextShapeResult extends GlyphMeasurement {
  readonly spans: readonly TextShapeSpan[];
  /** UTF-16 offsets at which line splitting may legally separate graphemes. */
  readonly graphemeBoundaries: readonly number[];
  /** Contextually measured source clusters, relative to the shaped request. */
  readonly clusters?: readonly Readonly<{
    range: Readonly<{ start: number; end: number }>;
    offsetPt: number;
    advancePt: number;
  }>[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export interface TextLayoutService {
  readonly fingerprint: string;
  /** Geometry keyed by a resolved resource identity. This is the authoritative
   * metric snapshot used by layout. */
  readonly fontMetrics?: Readonly<Record<string, Readonly<ResolvedFontMetric>>>;
  /** @deprecated Compatibility alias for {@link fontMetrics}. */
  readonly localMetrics: Readonly<Record<string, Readonly<ResolvedFontMetric>>>;
  resolve(request: Readonly<TextFontResolveRequest>): FontResolution;
  shape(request: Readonly<TextShapeRequest>): TextShapeResult;
}

export interface TextLayoutServiceInput {
  readonly fonts: FontResolver;
  readonly measurer: GlyphMeasurer;
  readonly fontMetrics?: Readonly<Record<string, Readonly<ResolvedFontMetric>>>;
  /** Exact local aliases also participate in resolution; retained separately
   * from resource-only metrics so an embedded face is never mislabeled local. */
  readonly localMetrics?: Readonly<Record<string, Readonly<ResolvedFontMetric>>>;
  readonly eastAsiaFontCharsets?: Readonly<Record<string, string>>;
  readonly genericFamilies?: Readonly<Record<string, 'serif' | 'sans-serif' | 'monospace'>>;
}

/** Generic tail for an authored DOCX face. fontTable family/pitch metadata is
 * authoritative; absent or `auto` entries use the shared, bounded face-name
 * classifier that also drives the native fallback list. */
export function classifyDocxFontGeneric(
  family: string | null | undefined,
  fontFamilyClasses: Readonly<Record<string, string>> = {},
  fontFamilyPitches: Readonly<Record<string, string>> = {},
): 'serif' | 'sans-serif' | 'monospace' {
  if (!family) return 'sans-serif';
  const tableClass = fontFamilyClasses[family];
  if (tableClass === 'roman') return 'serif';
  if (tableClass === 'swiss') return 'sans-serif';
  if (tableClass === 'modern' && fontFamilyPitches[family] === 'fixed') return 'monospace';
  const inferred = classifyFontGeneric(family);
  return inferred === 'mono' ? 'monospace' : inferred === 'serif' ? 'serif' : 'sans-serif';
}

/** Consumer choice at the §17.3.2.26 boundary where no font slot resolves.
 * Keep the established East Asian fallback stable; Word-produced evidence for
 * this change covers Latin and complex-script text only. */
function defaultGenericForSlot(
  slot: FontScriptSlot,
): 'serif' | 'sans-serif' {
  return slot === 'eastAsia' ? 'sans-serif' : 'serif';
}

const FONT_METRIC_SNAPSHOT = Symbol('docx.fontMetricSnapshot');
type FontMetricSnapshot = Readonly<Record<string, Readonly<ResolvedFontMetric>>> & {
  readonly [FONT_METRIC_SNAPSHOT]: true;
};

/** Copy successful face routes once at the document boundary. The brand lets
 * downstream services share the same deeply frozen object without retaining
 * caller-owned mutable records. */
export function snapshotFontMetrics(
  input: Readonly<Record<string, Readonly<ResolvedFontMetric>>> = {},
): Readonly<Record<string, Readonly<ResolvedFontMetric>>> {
  if ((input as Partial<FontMetricSnapshot>)[FONT_METRIC_SNAPSHOT]) return input;
  const entries = Object.entries(input)
    .map(([key, metric]) => {
      if (!metric.family?.trim()) throw new TypeError(`Font metric ${key} requires a family`);
      if (metric.lineHeightRatio !== undefined
        && (!Number.isFinite(metric.lineHeightRatio) || metric.lineHeightRatio < 0)) {
        throw new RangeError(`Font metric ${key} lineHeightRatio must be finite and non-negative`);
      }
      if (metric.eastAsianLineHeightRatio !== undefined
        && (!Number.isFinite(metric.eastAsianLineHeightRatio) || metric.eastAsianLineHeightRatio < 0)) {
        throw new RangeError(`Font metric ${key} eastAsianLineHeightRatio must be finite and non-negative`);
      }
      if (metric.fontBoxRatio !== undefined
        && (!Number.isFinite(metric.fontBoxRatio) || metric.fontBoxRatio <= 0)) {
        throw new RangeError(`Font metric ${key} fontBoxRatio must be finite and positive`);
      }
      if (metric.weight !== undefined
        && (!Number.isFinite(metric.weight) || metric.weight < 1 || metric.weight > 1000)) {
        throw new RangeError(`Font metric ${key} weight must be finite and between 1 and 1000`);
      }
      const copy: ResolvedFontMetric = {
        family: metric.family,
        ...(metric.lineHeightRatio === undefined ? {} : { lineHeightRatio: metric.lineHeightRatio }),
        ...(metric.eastAsianLineHeightRatio === undefined
          ? {}
          : { eastAsianLineHeightRatio: metric.eastAsianLineHeightRatio }),
        ...(metric.fontBoxRatio === undefined ? {} : { fontBoxRatio: metric.fontBoxRatio }),
        ...(metric.requestedFamily === undefined ? {} : { requestedFamily: metric.requestedFamily }),
        ...(metric.weight === undefined ? {} : { weight: metric.weight }),
        ...(metric.style === undefined ? {} : { style: metric.style }),
        ...(metric.sourceIdentity === undefined ? {} : { sourceIdentity: metric.sourceIdentity }),
        ...(metric.synthesized === undefined ? {} : { synthesized: metric.synthesized }),
      };
      return [normalizeFontMetricFamily(key), Object.freeze(copy)] as const;
    })
    .sort(([a], [b]) => a.localeCompare(b));
  const snapshot = Object.fromEntries(entries) as FontMetricSnapshot;
  Object.defineProperty(snapshot, FONT_METRIC_SNAPSHOT, { value: true });
  return Object.freeze(snapshot);
}

const LATIN1_EAST_ASIA = new Set([
  0x00a1, 0x00a4, 0x00a7, 0x00a8, 0x00aa, 0x00ad, 0x00af,
  0x00b0, 0x00b1, 0x00b2, 0x00b3, 0x00b4, 0x00b6, 0x00b7,
  0x00b8, 0x00b9, 0x00ba, 0x00bc, 0x00bd, 0x00be, 0x00bf, 0x00d7, 0x00f7,
]);
const LATIN1_CHINESE_EAST_ASIA = new Set([
  0x00e0, 0x00e1, 0x00e8, 0x00e9, 0x00ea, 0x00ec, 0x00ed,
  0x00f2, 0x00f3, 0x00f9, 0x00fa, 0x00fc,
]);

function scriptSlot(
  codePoint: number,
  forceComplex: boolean,
  hint: TextShapeRequest['fontHint'],
  eastAsiaLanguage: string | undefined,
  eastAsiaFontCharset: string | undefined,
): FontScriptSlot {
  const hintedEastAsia = hint === 'eastAsia';
  const chinese = eastAsiaLanguage?.split(/[-_]/, 1)[0]?.toLowerCase() === 'zh';
  const chineseCharset = /^(?:86|88)$/i.test(eastAsiaFontCharset?.trim() ?? '');
  let tableSlot: Exclude<FontScriptSlot, 'complexScript'> = 'highAnsi';
  // ECMA-376 §17.3.2.26 assigns the Hebrew/Arabic-family ranges to the ASCII
  // slot unless the run is explicitly complex-script (`w:cs` / `w:rtl`). They
  // must not fall through to highAnsi merely because their scalar is > 0x7f.
  if (codePoint <= 0x007f) tableSlot = 'ascii';
  else if (codePoint <= 0x00ff) {
    tableSlot = hintedEastAsia && (
      LATIN1_EAST_ASIA.has(codePoint)
      || (chinese && LATIN1_CHINESE_EAST_ASIA.has(codePoint))
    ) ? 'eastAsia' : 'highAnsi';
  } else if (codePoint >= 0x0100 && codePoint <= 0x02af) {
    tableSlot = hintedEastAsia && (chinese || chineseCharset) ? 'eastAsia' : 'highAnsi';
  } else if (
    (codePoint >= 0x02b0 && codePoint <= 0x02ff)
    || (codePoint >= 0x0300 && codePoint <= 0x036f)
    || (codePoint >= 0x0370 && codePoint <= 0x03cf)
    || (codePoint >= 0x0400 && codePoint <= 0x04ff)
  ) {
    tableSlot = hintedEastAsia ? 'eastAsia' : 'highAnsi';
  } else if (
    (codePoint >= 0x0590 && codePoint <= 0x07bf)
    || (codePoint >= 0xfb1d && codePoint <= 0xfdff)
    || (codePoint >= 0xfe70 && codePoint <= 0xfefe)
  ) tableSlot = 'ascii';
  else if (
    (codePoint >= 0x1100 && codePoint <= 0x11ff)
    || (codePoint >= 0x2e80 && codePoint <= 0x2eff)
    || (codePoint >= 0x2f00 && codePoint <= 0x2fdf)
    || (codePoint >= 0x2ff0 && codePoint <= 0x318f)
    || (codePoint >= 0x3190 && codePoint <= 0x319f)
    || (codePoint >= 0x3200 && codePoint <= 0x4dbf)
    || (codePoint >= 0x4e00 && codePoint <= 0x9faf)
    || (codePoint >= 0xa000 && codePoint <= 0xa48f)
    || (codePoint >= 0xa490 && codePoint <= 0xa4cf)
    || (codePoint >= 0xac00 && codePoint <= 0xd7af)
    || (codePoint >= 0xf900 && codePoint <= 0xfaff)
    || (codePoint >= 0xfe30 && codePoint <= 0xfe4f)
    || (codePoint >= 0xfe50 && codePoint <= 0xfe6f)
    || (codePoint >= 0xff00 && codePoint <= 0xffef)
    // The normative table is expressed over UTF-16 code units and assigns the
    // complete high/high-private/low-surrogate ranges to eastAsia. This shaper
    // iterates Unicode scalars, so every supplementary scalar projects through
    // one listed surrogate pair and is therefore equivalent to eastAsia.
    || (codePoint >= 0x10000 && codePoint <= 0x10ffff)
  ) tableSlot = 'eastAsia';
  else if (codePoint >= 0x1e00 && codePoint <= 0x1eff) {
    tableSlot = hintedEastAsia && chinese ? 'eastAsia' : 'highAnsi';
  } else if (
    (codePoint >= 0x2000 && codePoint <= 0x27bf)
    || (codePoint >= 0xe000 && codePoint <= 0xf8ff)
    || (codePoint >= 0xfb00 && codePoint <= 0xfb1c)
  ) tableSlot = hintedEastAsia ? 'eastAsia' : 'highAnsi';

  // §17.3.2.26 step 2: an eastAsia table result is protected from w:cs/w:rtl
  // only when rFonts@hint explicitly selects eastAsia. Otherwise cs wins.
  if (tableSlot === 'eastAsia' && hintedEastAsia) return tableSlot;
  if (forceComplex) return 'complexScript';
  return tableSlot;
}

function requestedFamily(
  request: Readonly<Pick<TextShapeRequest, 'fonts' | 'themeFonts' | 'themeFontPresence'>>,
  slot: FontScriptSlot,
): string | null | undefined {
  const selectedThemePresent = request.themeFontPresence?.[slot]
    ?? request.themeFonts?.[slot] != null;
  if (selectedThemePresent) return request.themeFonts?.[slot];
  const selected = request.fonts[slot];
  if (selected != null) return selected;
  const asciiThemePresent = request.themeFontPresence?.ascii
    ?? request.themeFonts?.ascii != null;
  if (asciiThemePresent) return request.themeFonts?.ascii;
  return request.fonts.ascii;
}

/**
 * Shape per script span because ECMA-376 §17.3.2.26 selects rFonts slots per
 * Unicode character; choosing one family for an entire mixed-script run loses
 * authored East Asian and complex-script faces.
 */
export function createTextLayoutService(input: TextLayoutServiceInput): TextLayoutService {
  const fontMetrics = snapshotFontMetrics({
    ...input.localMetrics,
    ...input.fontMetrics,
  });
  const genericFamilies = Object.freeze(Object.fromEntries(
    Object.entries(input.genericFamilies ?? {})
      .map(([family, generic]) => [family.trim().toLocaleLowerCase('en-US'), generic])
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  const eastAsiaFontCharsets = Object.freeze(Object.fromEntries(
    Object.entries(input.eastAsiaFontCharsets ?? {})
      .map(([family, charset]) => [family.trim().toLocaleLowerCase('en-US'), charset.trim()])
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  const fingerprint = stableFingerprint('text', {
    fonts: input.fonts.fingerprint,
    measurer: input.measurer.fingerprint,
    fontMetrics,
    eastAsiaFontCharsets,
    genericFamilies,
  });
  const resolve = (request: Readonly<TextFontResolveRequest>): FontResolution => {
    const authoredFamily = requestedFamily(request, request.slot);
    const genericFamily = authoredFamily
      ? genericFamilies[authoredFamily.trim().toLocaleLowerCase('en-US')]
        ?? request.genericFamily
        ?? 'sans-serif'
      // ECMA-376 §17.3.2.26 deliberately leaves the consumer's supported
      // default font unspecified when no rFonts slot resolves. This bounded
      // consumer policy changes only the script classes covered by Word output
      // evidence; every authored direct/theme face remains authoritative above.
      : request.genericFamily ?? defaultGenericForSlot(request.slot);
    return input.fonts.resolve({
      requestedFamily: authoredFamily,
      genericFamily,
      weight: request.weight,
      style: request.style,
    });
  };
  // Pagination convergence reacquires equal paragraphs under equal service
  // fingerprints. Native Canvas metrics are pure for this complete request
  // tuple, so retain one immutable document-scoped result across those passes.
  const measurementCache = new Map<string, Readonly<GlyphMeasurement>>();
  const measureGlyph = (request: Readonly<GlyphMeasureRequest>): Readonly<GlyphMeasurement> => {
    const key = JSON.stringify([
      request.text,
      request.fontRoute.familyList,
      request.fontRoute.scope,
      request.fontRoute.fingerprint,
      request.fontSizePt,
      request.weight,
      request.style,
      request.letterSpacingPt,
      request.kerning ?? null,
    ]);
    const retained = measurementCache.get(key);
    if (retained) return retained;
    const measured = input.measurer.measure(request);
    const snapshot = Object.freeze({
      ...measured,
      ...(measured.inkBounds ? {
        inkBounds: Object.freeze({ ...measured.inkBounds }),
      } : {}),
    });
    measurementCache.set(key, snapshot);
    return snapshot;
  };
  // Contextual cluster geometry measures every grapheme prefix of a script
  // span. Those prefixes are useful only while deriving this shape result:
  // retaining their text-bearing keys for the document lifetime would consume
  // quadratic space for one long run. The per-shape prefixAdvances map below
  // already deduplicates shared cluster edges within the acquisition.
  const measureContextualPrefixAdvance = (
    request: Readonly<GlyphMeasureRequest>,
  ): number => input.measurer.measure(request).advancePt;
  const shapeCache = new Map<string, TextShapeResult>();
  return Object.freeze({
    fingerprint,
    fontMetrics,
    localMetrics: fontMetrics,
    resolve,
    shape(request: Readonly<TextShapeRequest>): TextShapeResult {
      if (!Number.isFinite(request.fontSizePt) || request.fontSizePt < 0) {
        throw new RangeError('fontSizePt must be a finite non-negative number');
      }
      const shapeKey = JSON.stringify([
        request.text,
        request.fontSizePt,
        [
          request.fonts.ascii ?? null,
          request.fonts.highAnsi ?? null,
          request.fonts.eastAsia ?? null,
          request.fonts.complexScript ?? null,
        ],
        [
          request.themeFonts?.ascii ?? null,
          request.themeFonts?.highAnsi ?? null,
          request.themeFonts?.eastAsia ?? null,
          request.themeFonts?.complexScript ?? null,
        ],
        [
          request.themeFontPresence?.ascii ?? null,
          request.themeFontPresence?.highAnsi ?? null,
          request.themeFontPresence?.eastAsia ?? null,
          request.themeFontPresence?.complexScript ?? null,
        ],
        request.weight ?? null,
        request.style ?? null,
        request.complexScript ?? null,
        request.fontHint ?? null,
        request.eastAsiaLanguage ?? null,
        request.eastAsiaFontCharset ?? null,
        request.genericFamily ?? null,
        request.letterSpacingPt ?? null,
        request.kerning ?? null,
        request.measure ?? null,
        request.clusterGeometry ?? null,
      ]);
      const retainedShape = shapeCache.get(shapeKey);
      if (retainedShape) return retainedShape;
      const grouped: {
        text: string; start: number; end: number; script: FontScriptSlot; breakBefore: boolean;
      }[] = [];
      const graphemeBoundaries = Object.freeze(
        [...new Set([0, ...graphemeClusterOffsets(request.text), request.text.length])].sort((a, b) => a - b),
      );
      const graphemeStarts = new Set(graphemeBoundaries);
      let start = 0;
      for (const character of request.text) {
        const end = start + character.length;
        const eastAsiaFamily = requestedFamily(request, 'eastAsia');
        const eastAsiaCharset = request.eastAsiaFontCharset
          ?? (eastAsiaFamily
            ? eastAsiaFontCharsets[eastAsiaFamily.trim().toLocaleLowerCase('en-US')]
            : undefined);
        const script = scriptSlot(
          character.codePointAt(0) ?? 0,
          request.complexScript ?? false,
          request.fontHint,
          request.eastAsiaLanguage,
          eastAsiaCharset,
        );
        const previous = grouped.at(-1);
        if (previous?.script === script) {
          previous.text += character;
          previous.end = end;
        } else {
          grouped.push({ text: character, start, end, script, breakBefore: graphemeStarts.has(start) });
        }
        start = end;
      }

      const spans = grouped.map((group): TextShapeSpan => {
        const font = resolve({
          fonts: request.fonts,
          themeFonts: request.themeFonts,
          themeFontPresence: request.themeFontPresence,
          slot: group.script,
          weight: request.weight,
          style: request.style,
          genericFamily: request.genericFamily,
        });
        const measurement = request.measure === false ? {
          advancePt: 0,
          ascentPt: 0,
          descentPt: 0,
        } : measureGlyph({
          text: group.text,
          fontRoute: font.route,
          fontSizePt: request.fontSizePt,
          weight: font.weight,
          style: font.style,
          letterSpacingPt: request.letterSpacingPt ?? 0,
          kerning: request.kerning,
        });
        return Object.freeze({
          ...group, ...measurement, font, fontRoute: font.route,
        });
      });
      const diagnostics = spans.flatMap((span) => span.font.diagnostics);
      const inkBounds = spans.length > 0 && spans.every((span) => span.inkBounds !== undefined)
        ? (() => {
            let originPt = 0;
            let xMinPt = Number.POSITIVE_INFINITY;
            let xMaxPt = Number.NEGATIVE_INFINITY;
            let ascentPt = 0;
            let descentPt = 0;
            for (const span of spans) {
              const ink = span.inkBounds as GlyphInkBounds;
              xMinPt = Math.min(xMinPt, originPt + ink.xMinPt);
              xMaxPt = Math.max(xMaxPt, originPt + ink.xMaxPt);
              ascentPt = Math.max(ascentPt, ink.ascentPt);
              descentPt = Math.max(descentPt, ink.descentPt);
              originPt += span.advancePt;
            }
            return Object.freeze({ xMinPt, xMaxPt, ascentPt, descentPt });
          })()
        : undefined;
      const totalAdvancePt = spans.reduce((sum, span) => sum + span.advancePt, 0);
      const clusters = request.clusterGeometry === false
        ? undefined
        : (() => {
            const prefixAdvances = new Map<number, number>([
              [0, 0],
              [request.text.length, totalAdvancePt],
            ]);
            const prefixAdvance = (boundary: number): number => {
              if (request.measure === false || boundary <= 0) return 0;
              const retained = prefixAdvances.get(boundary);
              if (retained !== undefined) return retained;
              let advancePt = 0;
              for (const span of spans) {
                if (boundary >= span.end) {
                  advancePt += span.advancePt;
                  continue;
                }
                if (boundary <= span.start) break;
                advancePt += measureContextualPrefixAdvance({
                  text: span.text.slice(0, boundary - span.start),
                  fontRoute: span.fontRoute,
                  fontSizePt: request.fontSizePt,
                  weight: span.font.weight,
                  style: span.font.style,
                  letterSpacingPt: request.letterSpacingPt ?? 0,
                  kerning: request.kerning,
                });
                break;
              }
              // One boundary is the trailing edge of one cluster and the leading
              // edge of the next. Keep one contextual fact for this shape call.
              prefixAdvances.set(boundary, advancePt);
              return advancePt;
            };
            return Object.freeze(graphemeBoundaries.slice(0, -1).map((start, index) => {
              const end = graphemeBoundaries[index + 1] ?? start;
              const offsetPt = prefixAdvance(start);
              return Object.freeze({
                range: Object.freeze({ start, end }),
                offsetPt,
                advancePt: prefixAdvance(end) - offsetPt,
              });
            }));
          })();
      const result: TextShapeResult = Object.freeze({
        advancePt: totalAdvancePt,
        ascentPt: Math.max(0, ...spans.map((span) => span.ascentPt)),
        descentPt: Math.max(0, ...spans.map((span) => span.descentPt)),
        ...(inkBounds ? { inkBounds } : {}),
        ...(inkBounds && spans.every((span) => span.horizontalInkBoundsAreTight === true)
          ? { horizontalInkBoundsAreTight: true }
          : {}),
        spans: Object.freeze(spans),
        graphemeBoundaries,
        ...(clusters ? { clusters } : {}),
        diagnostics: Object.freeze(diagnostics),
      });
      shapeCache.set(shapeKey, result);
      return result;
    },
  });
}
