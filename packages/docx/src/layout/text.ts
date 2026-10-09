import { wordTextBoxVerticalMode } from './compatibility.js';
import type { LayoutDiagnostic } from './types.js';
import {
  classifyFontGeneric,
  fontSubstituteScriptCoversText,
  fontSubstituteScriptScope,
  graphemeClusterOffsets,
  normalizeFontMetricFamily,
  type CjkLang,
  type FontSubstituteScript,
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
import { containsHanScript } from '@silurus/ooxml-core/internal/script-preload-accumulator';
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
  const textBoxVertical = wordTextBoxVerticalMode(textVert) !== undefined;
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
  /** Private native reserved-separator character (MS-DOC 2.3.3): the
   * U+0003/U+0004 control or the story paragraph's content mark. It carries
   * that character's own effective CHPX and has no text; line acquisition
   * resolves its selected face as a zero-advance, inkless metric participant.
   * Its rule ink is retained separately, never as a glyph. */
  noteSeparatorCharacter?: 'rule-control' | 'paragraph-mark';
}>;

export type ParagraphTextBearingRun =
  | (DeepReadonly<Extract<DocRun, { type: 'text' }>> & ParagraphTextFacts)
  | (DeepReadonly<Extract<DocRun, { type: 'field' }>> &
    DeepReadonly<Partial<DocxTextRun>> & ParagraphTextFacts);

export type ParagraphMathRun = Readonly<{
  type: 'math';
  revision?: DeepReadonly<DocRun['revision']>;
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
  /** Resolved §17.3.2.19 w:kern state at this run size. Absence preserves the
   * measurement adapter's inherited kerning policy, matching the paint path. */
  readonly kerning?: boolean;
  /** Builder-owned proof that one registered grapheme has no per-slot
   * allocation policy. The shaper still revalidates resource and cmap facts. */
  readonly joinRegisteredGrapheme?: boolean;
  /** Resolve script slots and faces without touching the measurement adapter. */
  readonly measure?: boolean;
  /** False acquires only aggregate metrics; 'spaces' acquires contextual
   * scalar U+0020 advances for fit arithmetic. Full clusters are acquired after
   * wrapping, avoiding repeated prefix shaping of overlong words. */
  readonly clusterGeometry?: boolean | 'spaces';
  /** The same run's surrounding text, with this request's offset in it. A
   * script-scoped substitute's scope is decided over the whole contiguous
   * context, so a word of Arabic digits after proven Arabic text continues it,
   * while digits alone never enable the substitute. */
  readonly substituteContext?: Readonly<{ text: string; offset: number }>;
}

/** The measured class has a strong Latin base with attached marks. rFonts
 * ascii is a semantic slot, not proof of ASCII or bidi direction: it also
 * contains unmarked Hebrew/Arabic, which must retain their existing path. */
export function registeredLatinMarkGraphemeCandidate(text: string): boolean {
  return /^[A-Za-z]\p{M}+$/u.test(text) && graphemeClusterOffsets(text).length === 0;
}

/** Validate against the owning transformed run, not merely a self-consistent
 * fragment. Acquisition knows the full run; partial measurements inherit it. */
export function assertTextShapeRunContext(
  request: Readonly<Pick<TextShapeRequest, 'text' | 'substituteContext'>>,
  fullRunText: string,
): void {
  const context = request.substituteContext;
  if (!context || context.text !== fullRunText
    || !Number.isInteger(context.offset) || context.offset < 0
    || context.offset + request.text.length > fullRunText.length
    || fullRunText.slice(context.offset, context.offset + request.text.length) !== request.text) {
    throw new Error('Text shape request does not match its full run context; project the range explicitly');
  }
}

/** Slice retained shaping input without re-judging a fragment's script proof.
 * UTF-16 offsets match both Canvas text and the shared run-scope descriptor. */
export function sliceTextShapeRequest(
  request: Readonly<TextShapeRequest>,
  start: number,
  end: number,
): TextShapeRequest {
  const context = request.substituteContext ?? { text: request.text, offset: 0 };
  return {
    ...request,
    text: request.text.slice(start, end),
    joinRegisteredGrapheme: start === 0 && end === request.text.length
      ? request.joinRegisteredGrapheme : undefined,
    substituteContext: { text: context.text, offset: context.offset + start },
  };
}

/** A generated probe or ruby guide is independent text, even when its glyphs
 * happen to occur in the base run. It must not inherit that run's Arabic proof. */
export function independentTextShapeRequest(
  request: Readonly<TextShapeRequest>,
  text: string,
): TextShapeRequest {
  return { ...request, text, joinRegisteredGrapheme: undefined, substituteContext: { text, offset: 0 } };
}

/** Transform this range in its run (e.g. inserting justification kashidas),
 * keeping proof on either side and recomputing scope for the transformed run. */
export function replaceTextShapeRequest(
  request: Readonly<TextShapeRequest>,
  text: string,
): TextShapeRequest {
  const context = request.substituteContext ?? { text: request.text, offset: 0 };
  return { ...request, text, joinRegisteredGrapheme: undefined, substituteContext: {
    text: context.text.slice(0, context.offset) + text
      + context.text.slice(context.offset + request.text.length),
    offset: context.offset,
  } };
}

export interface TextFontResolveRequest {
  /** Text covered by this request. Language-selected regional fallback applies
   * only when the text actually contains Han; it must not capture Latin glyphs. */
  readonly text?: string;
  readonly eastAsiaLanguage?: string;
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly themeFontPresence?: TextFontSlotPresence;
  readonly slot: FontScriptSlot;
  readonly weight?: number;
  readonly style?: FontStyle;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  /** Carried decision whether a script-scoped substitute covers this text,
   * made by the shaper with the shared per-cluster scope rule. When omitted,
   * the same rule judges `text` as a whole. */
  readonly substituteScope?: boolean;
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
  /** Original ECMA-376 rFonts facts within one physical grapheme. These
   * carry no competing advance or ink authority. */
  readonly semanticSlotSpans?: readonly Readonly<{
    start: number; end: number; script: FontScriptSlot; font: FontResolution;
  }>[];
  /** Run-context substitute decision: true selects the scoped face, false
   * retains exclusion. Absence means this request has no scoped substitute. */
  readonly substituteScope?: boolean;
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
  /** Contextually measured source clusters, relative to the shaped request.
   * For clusterGeometry:'spaces' these are scalar space ranges; the caller
   * must preserve grapheme-safe cuts when allocating their advances. */
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
  /** Per-source script-substitute proof, when configured. A mixed scope stays
   * a semantic face-selection boundary; plain runs return no scope key. */
  sourceScopeKey?(request: Readonly<TextShapeRequest>): string | undefined;
}

export interface TextLayoutServiceInput {
  readonly fonts: FontResolver;
  readonly measurer: GlyphMeasurer;
  readonly cjkFallback?: CjkLang;
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

function snapshotUnicodeRanges(
  ranges: readonly (readonly [number, number])[],
): readonly (readonly [number, number])[] {
  if (ranges.length > 32_768) throw new RangeError('Font cmap coverage has too many ranges');
  const sorted = ranges.map(([start, end]) => {
    if (!Number.isInteger(start) || !Number.isInteger(end)
      || start < 0 || end > 0x10ffff || start > end) {
      throw new RangeError('Font cmap coverage contains an invalid range');
    }
    return [start, end] as const;
  }).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const merged: Array<readonly [number, number]> = [];
  for (const [start, end] of sorted) {
    const last = merged[merged.length - 1];
    if (last && start <= last[1] + 1) {
      merged[merged.length - 1] = [last[0], Math.max(last[1], end)];
    } else {
      merged.push([start, end]);
    }
  }
  return Object.freeze(merged.map((range) => Object.freeze(range)));
}

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
      if (metric.designAscentRatio !== undefined
        && (!Number.isFinite(metric.designAscentRatio) || metric.designAscentRatio < 0)) {
        throw new RangeError(`Font metric ${key} designAscentRatio must be finite and non-negative`);
      }
      if (metric.designDescentRatio !== undefined
        && (!Number.isFinite(metric.designDescentRatio) || metric.designDescentRatio < 0)) {
        throw new RangeError(`Font metric ${key} designDescentRatio must be finite and non-negative`);
      }
      if (metric.eastAsianLineHeightRatio !== undefined
        && (!Number.isFinite(metric.eastAsianLineHeightRatio) || metric.eastAsianLineHeightRatio < 0)) {
        throw new RangeError(`Font metric ${key} eastAsianLineHeightRatio must be finite and non-negative`);
      }
      if (metric.fontBoxRatio !== undefined
        && (!Number.isFinite(metric.fontBoxRatio) || metric.fontBoxRatio <= 0)) {
        throw new RangeError(`Font metric ${key} fontBoxRatio must be finite and positive`);
      }
      if (metric.averageCharWidthRatio !== undefined
        && (!Number.isFinite(metric.averageCharWidthRatio) || metric.averageCharWidthRatio <= 0)) {
        throw new RangeError(`Font metric ${key} averageCharWidthRatio must be finite and positive`);
      }
      if (metric.weight !== undefined
        && (!Number.isFinite(metric.weight) || metric.weight < 1 || metric.weight > 1000)) {
        throw new RangeError(`Font metric ${key} weight must be finite and between 1 and 1000`);
      }
      const copy: ResolvedFontMetric = {
        family: metric.family,
        ...(metric.lineHeightRatio === undefined ? {} : { lineHeightRatio: metric.lineHeightRatio }),
        ...(metric.designAscentRatio === undefined ? {} : { designAscentRatio: metric.designAscentRatio }),
        ...(metric.designDescentRatio === undefined ? {} : { designDescentRatio: metric.designDescentRatio }),
        ...(metric.eastAsianLineHeightRatio === undefined
          ? {}
          : { eastAsianLineHeightRatio: metric.eastAsianLineHeightRatio }),
        ...(metric.fontBoxRatio === undefined ? {} : { fontBoxRatio: metric.fontBoxRatio }),
        ...(metric.averageCharWidthRatio === undefined
          ? {}
          : { averageCharWidthRatio: metric.averageCharWidthRatio }),
        ...(metric.unicodeRanges === undefined
          ? {}
          : { unicodeRanges: snapshotUnicodeRanges(metric.unicodeRanges) }),
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

const SCOPE_NEUTRAL_SCALAR = /^[\s\p{Cf}\p{M}]$/u;
/** Recent run contexts whose substitute scope is retained (one per run). */
const SCOPE_CACHE_LIMIT = 64;
const SCOPE_CONFIGURATIONS_PER_RUN = 8;
const SCOPE_SLOTS = ['ascii', 'highAnsi', 'eastAsia', 'complexScript'] as const;

/** Two resolutions select the same registered face: the same resource,
 * weight, style and source. Native and generic routes name no resource, so
 * Canvas's eventual choice is unknown and they never compare equal. */
function sameEffectiveFace(a: FontResolution, b: FontResolution): boolean {
  if (a === b) return true;
  if (a.source === 'native' || a.source === 'generic') return false;
  return a.source === b.source && a.resolvedFamily === b.resolvedFamily
    && a.resourceIdentity === b.resourceIdentity && a.weight === b.weight && a.style === b.style;
}

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
    // §17.3.2.26 General Punctuation: U+2014 follows hint, not language
    // alone. Omitted/default hint selects highAnsi; eastAsia selects eastAsia.
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

/** ECMA-376 §17.3.2.26 slot family: a present theme reference (even one that
 * resolved to no name) governs its slot, then the direct slot, then the ascii
 * theme/direct fallback. Exported so resource preloading requests exactly the
 * families this shaper will ask the font resolver for. */
export function requestedFamily(
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

/** Upper bound on distinct font routes a text service keeps ordinals for;
 * reaching it retires the ordinals together with the measurement cache. */
export const TEXT_ROUTE_ORDINAL_LIMIT = 4096;
const routeOrdinalTableSizes = new WeakMap<object, () => number>();
const scopeScanStats = new WeakMap<object, () => Readonly<{ scans: number; utf16Units: number }>>();

/** Internal resource diagnostic: whole-run work, independent of wall-clock noise. */
export function textScopeScanStats(service: TextLayoutService): Readonly<{ scans: number; utf16Units: number }> | undefined {
  return scopeScanStats.get(service)?.();
}

/** Internal diagnostic: live route-ordinal entries held by a text service. */
export function textRouteOrdinalTableSize(service: TextLayoutService): number | undefined {
  return routeOrdinalTableSizes.get(service)?.();
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
  const metricTupleKey = (identity: string, family: string, weight: number, style: FontStyle) =>
    JSON.stringify([identity, family, weight, style]);
  const registeredMetrics = new Map<string, Readonly<ResolvedFontMetric>>();
  for (const metric of Object.values(fontMetrics)) {
    if (metric.sourceIdentity && metric.weight !== undefined && metric.style !== undefined) {
      registeredMetrics.set(metricTupleKey(metric.sourceIdentity, metric.family, metric.weight, metric.style), metric);
    }
  }
  // Metric resources represented in this immutable snapshot, grouped by the
  // case-insensitive Canvas family/weight/style tuple. This does not enumerate
  // registrations outside the snapshot in an external FontFaceSet.
  const canvasFaceKey = (family: string, weight: number, style: FontStyle) =>
    JSON.stringify([normalizeFontMetricFamily(family), weight, style]);
  const metricsByCanvasFace = new Map<string, Readonly<ResolvedFontMetric>[]>();
  for (const metric of Object.values(fontMetrics)) {
    const key = canvasFaceKey(metric.family, metric.weight ?? 400, metric.style ?? 'normal');
    const peers = metricsByCanvasFace.get(key);
    if (peers) peers.push(metric);
    else metricsByCanvasFace.set(key, [metric]);
  }
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
    cjkFallback: input.cjkFallback ?? null,
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
    const hasHan = containsHanScript(request.text ?? '');
    // A scoped visual substitute (core substitute-script.ts) answers only text
    // of its script. The shaper carries the shared per-cluster scope
    // (`substituteScope`); otherwise the same shared rule judges `text` as a
    // whole: a complex-script span is covered once it proves the script, any
    // other span only when all of it belongs to the proven script.
    const scoped = input.fonts.scopedSubstituteScript?.(authoredFamily, request.weight, request.style);
    const covered = request.substituteScope ?? (scoped !== undefined && fontSubstituteScriptCoversText(
      scoped,
      request.text ?? '',
      request.slot === 'complexScript' ? 'any' : 'exclusive',
    ));
    const substituteScript = scoped && covered ? scoped : undefined;
    return input.fonts.resolve({
      requestedFamily: authoredFamily,
      ...(substituteScript ? { script: substituteScript } : {}),
      cjkFallback: hasHan ? input.cjkFallback : undefined,
      language: request.slot === 'eastAsia' && hasHan
        ? request.eastAsiaLanguage
        : undefined,
      genericFamily,
      weight: request.weight,
      style: request.style,
    });
  };
  // Pagination convergence and later view variants revisit text under the same
  // service fingerprint. Keep recent pure results for those passes, but cap
  // document lifetime retention after long documents finish paginating.
  // The 295-page relayout probe used 33,024 distinct measurements and 73,436
  // shapes; a smaller cap caused sequential variant layouts to thrash.
  const measurementCacheLimit = 40960;
  const shapeCacheLimit = 81920;
  const measurementCache = new Map<string, Readonly<GlyphMeasurement>>();
  const cached = <T>(cache: Map<string, T>, key: string): T | undefined => {
    const value = cache.get(key);
    if (value !== undefined) {
      cache.delete(key);
      cache.set(key, value);
    }
    return value;
  };
  const retain = <T>(cache: Map<string, T>, key: string, value: T, limit: number): void => {
    cache.set(key, value);
    if (cache.size > limit) cache.delete(cache.keys().next().value as string);
  };
  // A route's identity is its (familyList, scope, fingerprint) triple. Spelling
  // that triple into every measurement key made each retained key ~1.8 KB for
  // a registered local face, so the measurement cache was dominated by copies
  // of three distinct routes. Keys carry a service-scoped ordinal assigned
  // one-to-one to the exact triple instead; the object map only skips
  // re-deriving the ordinal for the shared resolver routes.
  //
  // The ordinal table must not outlive the bounded cache it serves: a document
  // with ever-new routes would otherwise grow it without limit after the
  // measurement cache had evicted every key that used them. Only measurement
  // keys carry ordinals, so when the table reaches its bound, the table, the
  // object map and the measurement cache are dropped together. No key that
  // uses a retired ordinal survives, so a re-issued ordinal can never alias
  // two routes; the reset is a memoization miss, never a different result.
  let routeOrdinals = new Map<string, number>();
  let routeOrdinalByObject = new WeakMap<object, number>();
  const routeOrdinal = (route: Readonly<CanvasFontRoute>): number => {
    const known = routeOrdinalByObject.get(route);
    if (known !== undefined) return known;
    const identity = JSON.stringify([route.familyList, route.scope, route.fingerprint]);
    let ordinal = routeOrdinals.get(identity);
    if (ordinal === undefined) {
      if (routeOrdinals.size >= TEXT_ROUTE_ORDINAL_LIMIT) {
        routeOrdinals = new Map();
        routeOrdinalByObject = new WeakMap();
        measurementCache.clear();
      }
      ordinal = routeOrdinals.size;
      routeOrdinals.set(identity, ordinal);
    }
    // Only frozen routes are memoized by object: a mutable route could change
    // its triple after the ordinal was recorded.
    if (Object.isFrozen(route)) routeOrdinalByObject.set(route, ordinal);
    return ordinal;
  };
  const measureGlyph = (request: Readonly<GlyphMeasureRequest>): Readonly<GlyphMeasurement> => {
    const key = JSON.stringify([
      request.text,
      routeOrdinal(request.fontRoute),
      request.fontSizePt,
      request.weight,
      request.style,
      request.letterSpacingPt,
      request.kerning ?? null,
    ]);
    const retained = cached(measurementCache, key);
    if (retained) return retained;
    const measured = input.measurer.measure(request);
    const snapshot = Object.freeze({
      ...measured,
      ...(measured.inkBounds ? {
        inkBounds: Object.freeze({ ...measured.inkBounds }),
      } : {}),
    });
    retain(measurementCache, key, snapshot, measurementCacheLimit);
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
  const scopeCache = new Map<string, Map<string, Uint8Array>>();
  let scopeScans = 0;
  let scopeUtf16Units = 0;
  const service: TextLayoutService = Object.freeze({
    fingerprint,
    fontMetrics,
    localMetrics: fontMetrics,
    resolve,
    sourceScopeKey(request: Readonly<TextShapeRequest>): string | undefined {
      if (!SCOPE_SLOTS.some(slot => input.fonts.scopedSubstituteScript?.(
        requestedFamily(request, slot), request.weight, request.style))) return undefined;
      const spans = service.shape({ ...request, measure: false, clusterGeometry: false }).spans;
      const keys = [...new Set(spans.map(span => JSON.stringify([
        span.fontRoute.fingerprint, span.substituteScope,
      ])))];
      // The independent run-scope rule must not borrow Arabic proof from a
      // neighbouring run. Mixed scopes require their original source context.
      return keys.length === 1 ? keys[0] : 'mixed';
    },
    shape(request: Readonly<TextShapeRequest>): TextShapeResult {
      if (!Number.isFinite(request.fontSizePt) || request.fontSizePt < 0) {
        throw new RangeError('fontSizePt must be a finite non-negative number');
      }
      // Retained partial measurements must use sliceTextShapeRequest. Failing
      // closed prevents a stale offset from silently selecting a different face
      // than paint. Independent and transformed text have explicit helpers.
      const context = request.substituteContext ?? { text: request.text, offset: 0 };
      if (request.substituteContext) assertTextShapeRunContext(request, context.text);
      const configuredBySlot = new Map<FontScriptSlot, FontSubstituteScript>();
      const scopedBySlot = new Map<FontScriptSlot, FontSubstituteScript>();
      for (const slot of SCOPE_SLOTS) {
        const family = requestedFamily(request, slot);
        const configured = input.fonts.configuredSubstituteScript?.(family);
        if (configured) configuredBySlot.set(slot, configured);
        const scoped = input.fonts.scopedSubstituteScript?.(family, request.weight, request.style);
        if (scoped) scopedBySlot.set(slot, scoped);
      }
      const eastAsiaFamily = requestedFamily(request, 'eastAsia');
      const eastAsiaCharset = request.eastAsiaFontCharset
        ?? (eastAsiaFamily
          ? eastAsiaFontCharsets[eastAsiaFamily.trim().toLocaleLowerCase('en-US')]
          : undefined);
      // ECMA-376 §17.3.2.26 (and the MS-OI29500 table) selects the rFonts slot
      // per code point, including rFonts@hint. General text keeps that exactly.
      type Scalar = { text: string; start: number; end: number; slot: FontScriptSlot; inScope: boolean };
      const scalarsOf = (text: string): Scalar[] => {
        const result: Scalar[] = [];
        let offset = 0;
        for (const character of text) {
          const end = offset + character.length;
          result.push({
            text: character,
            start: offset,
            end,
            slot: scriptSlot(
              character.codePointAt(0) ?? 0,
              request.complexScript ?? false,
              request.fontHint,
              request.eastAsiaLanguage,
              eastAsiaCharset,
            ),
            inScope: false,
          });
          offset = end;
        }
        return result;
      };
      // Library substitution policy (§17.8.2), not an Office slot override:
      // core's one proof/extension/neutral rule judges clusters over the run.
      // Retain its scoped host slot at each UTF-16 position, including marks
      // and transparent controls, so word/span cuts cannot change the face.
      let scopeDescriptor: string | undefined;
      if (scopedBySlot.size > 0) {
        const key = JSON.stringify([
          [...scopedBySlot], [...configuredBySlot], request.complexScript ?? false, request.fontHint ?? null,
          request.eastAsiaLanguage ?? null, eastAsiaCharset ?? null,
        ]);
        // The raw run string is a Map key only once, never JSON-serialized or
        // embedded in each word key. Parent tokens and their resolved children
        // can alternate cs and scalar-slot classification on every word. Keep
        // both configurations: replacing one descriptor would rescan the whole
        // run per child, making layout quadratic. Both LRU levels are bounded
        // resource policy (64 runs, 8 configurations per run); retained arrays
        // remain linear in the retained run lengths, independent of word count.
        let configurations = cached(scopeCache, context.text);
        if (!configurations) {
          configurations = new Map();
          retain(scopeCache, context.text, configurations, SCOPE_CACHE_LIMIT);
        }
        let slots = cached(configurations, key);
        if (!slots) {
          scopeScans += 1;
          scopeUtf16Units += context.text.length;
          const contextScalars = scalarsOf(context.text);
          const indexAt = new Map(contextScalars.map((scalar, index) => [scalar.start, index]));
          slots = new Uint8Array(context.text.length);
          for (const script of new Set(scopedBySlot.values())) {
            const baseSlot = (start: number, end: number): FontScriptSlot | undefined => {
              let first: FontScriptSlot | undefined;
              for (let index = indexAt.get(start) ?? contextScalars.length; index < contextScalars.length
                && contextScalars[index]!.start < end; index += 1) {
                first ??= contextScalars[index]!.slot;
                if (!SCOPE_NEUTRAL_SCALAR.test(contextScalars[index]!.text)) return contextScalars[index]!.slot;
              }
              return first;
            };
            const scope = fontSubstituteScriptScope(script, context.text, (start, end) => {
              const slot = baseSlot(start, end);
              return slot !== undefined && (configuredBySlot.get(slot) ?? scopedBySlot.get(slot)) === script;
            });
            let host: FontScriptSlot | undefined;
            for (const cluster of scope) {
              if (!cluster.inScope) {
                host = undefined;
                continue;
              }
              if (cluster.cls !== 'neutral') host = baseSlot(cluster.start, cluster.end);
              // Surrounding Arabic can prove the run even when its own face
              // is unavailable or authored. Only an actually selected scoped
              // substitute may replace this cluster's normative scalar slots.
              if (host !== undefined && scopedBySlot.get(host) === script) slots.fill(SCOPE_SLOTS.indexOf(host) + 1, cluster.start, cluster.end);
            }
          }
          retain(configurations, key, slots, SCOPE_CONFIGURATIONS_PER_RUN);
        }
        // Exact, collision-free range descriptor: zero means normal scalar
        // classification; 1–4 carry the in-scope host slot. Key size is O(word),
        // and identical words with identical decisions share shapes across runs.
        scopeDescriptor = Array.from(
          slots.subarray(context.offset, context.offset + request.text.length),
          (slot) => String.fromCharCode(48 + slot),
        ).join('');
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
        request.joinRegisteredGrapheme ?? null,
        request.measure ?? null,
        request.clusterGeometry ?? null,
        ...(scopeDescriptor !== undefined ? [scopeDescriptor] : []),
      ]);
      const retainedShape = cached(shapeCache, shapeKey);
      if (retainedShape) return retainedShape;
      const grouped: {
        text: string; start: number; end: number; script: FontScriptSlot; breakBefore: boolean;
        substituteScript: boolean;
      }[] = [];
      const graphemeBoundaries = Object.freeze(
        [...new Set([0, ...graphemeClusterOffsets(request.text), request.text.length])].sort((a, b) => a - b),
      );
      const graphemeStarts = new Set(graphemeBoundaries);
      const scalars = scalarsOf(request.text);
      scalars.forEach((scalar) => {
        const scopedSlot = Number(scopeDescriptor?.[scalar.start] ?? 0);
        if (scopedSlot > 0) {
          scalar.inScope = true;
          scalar.slot = SCOPE_SLOTS[scopedSlot - 1]!;
        }
      });
      for (const scalar of scalars) {
        const previous = grouped.at(-1);
        if (previous?.script === scalar.slot && previous.substituteScript === scalar.inScope) {
          previous.text += scalar.text;
          previous.end = scalar.end;
        } else {
          grouped.push({
            text: scalar.text,
            start: scalar.start,
            end: scalar.end,
            script: scalar.slot,
            breakBefore: graphemeStarts.has(scalar.start),
            substituteScript: scalar.inScope,
          });
        }
      }

      const resolvedGroups = grouped.map((group) => ({
        ...group,
        font: resolve({
          fonts: request.fonts,
          themeFonts: request.themeFonts,
          themeFontPresence: request.themeFontPresence,
          slot: group.script,
          text: group.text,
          eastAsiaLanguage: request.eastAsiaLanguage,
          weight: request.weight,
          style: request.style,
          genericFamily: request.genericFamily,
          // Carry the shared per-cluster scope; a fragment is never re-judged.
          ...(scopedBySlot.has(group.script) ? { substituteScope: group.substituteScript } : {}),
        }),
      }));
      // Scoped text that resolves to the same effective face (same registered
      // resource, weight, style and source) is one string for Canvas, whichever
      // family name or slot requested it and whatever CSS fallback list follows.
      // Shaping it apart would break Arabic joining. Only in-scope spans
      // merge; all other spans keep main's per-slot runs.
      const merged: Array<(typeof resolvedGroups)[number] & Pick<TextShapeSpan, 'semanticSlotSpans'>> = [];
      for (const group of resolvedGroups) {
        const previous = merged.at(-1);
        if (previous && previous.substituteScript && group.substituteScript
          && sameEffectiveFace(previous.font, group.font)) {
          merged[merged.length - 1] = { ...previous, text: previous.text + group.text, end: group.end };
        } else {
          merged.push(group);
        }
      }

      // Semantic slots select faces per scalar (§17.3.2.26); they need not
      // detach a combining mark when the exact physical face is proven.
      // Keep the existing scoped Arabic merge above independent of this gate.
      if (request.joinRegisteredGrapheme === true && registeredLatinMarkGraphemeCandidate(request.text)
        && graphemeBoundaries.length === 2
        && merged.length > 1 && scopeDescriptor === undefined) {
        const first = merged[0]!;
        const face = first.font;
        const metric = face.resourceIdentity ? registeredMetrics.get(metricTupleKey(
          face.resourceIdentity, face.resolvedFamily, face.weight, face.style,
        )) : undefined;
        const scalars = Array.from(request.text, scalar => scalar.codePointAt(0)!);
        const coveredCount = (ranges: readonly (readonly [number, number])[]) => scalars.filter(cp => {
          let low = 0;
          let high = ranges.length;
          while (low < high) {
            const middle = (low + high) >>> 1;
            if (ranges[middle]![1] < cp) low = middle + 1;
            else high = middle;
          }
          return low < ranges.length && ranges[low]![0] <= cp;
        }).length;
        const ranges = metric?.unicodeRanges;
        // This join guard is stricter than selected line-metric admission:
        // every snapshot peer must explicitly cover all or none of the
        // grapheme. Unknown or partial cmap coverage declines joining;
        // geometry selection remains the separate line-metric policy.
        const covered = ranges !== undefined && coveredCount(ranges) === scalars.length
          && (metricsByCanvasFace.get(canvasFaceKey(face.resolvedFamily, face.weight, face.style)) ?? [])
            .every(peer => peer === metric || (peer.unicodeRanges !== undefined
              && [0, scalars.length].includes(coveredCount(peer.unicodeRanges))));
        if ((face.source === 'embedded' || face.source === 'local') && face.resourceIdentity
          && covered && merged.every(group => !group.substituteScript
            && (group.script === 'ascii' || group.script === 'highAnsi')
            && group.font.source === face.source && group.font.resourceIdentity === face.resourceIdentity
            && group.font.weight === face.weight && group.font.style === face.style
            && group.font.route.fingerprint === face.route.fingerprint)) {
          const semanticSlotSpans = Object.freeze(merged.map(group => Object.freeze({
            start: group.start, end: group.end, script: group.script, font: group.font,
          })));
          merged.splice(0, merged.length, { ...first, text: request.text, end: request.text.length,
            semanticSlotSpans });
        }
      }

      const spans = merged.map(({ substituteScript, font, ...group }): TextShapeSpan => {
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
          // Excluded spans also depend on the full run: a mark attached to a
          // Latin base must not become Arabic proof when measured in isolation.
          ...(scopeDescriptor !== undefined ? { substituteScope: substituteScript } : {}),
        });
      });
      const diagnostics = spans.flatMap(span => span.semanticSlotSpans
        ? span.semanticSlotSpans.flatMap(slot => slot.font.diagnostics) : span.font.diagnostics);
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
            const selected: Array<{ start: number; end: number }> = [];
            if (request.clusterGeometry === 'spaces') {
              // Scalar spaces need not start a grapheme (Prepend + SPACE).
              // Acquire their exact contextual prefix difference; gap
              // selection independently rejects cuts inside a grapheme.
              for (let start = request.text.indexOf(' '); start >= 0;
                start = request.text.indexOf(' ', start + 1)) {
                selected.push({ start, end: start + 1 });
              }
            } else {
              for (let index = 0; index < graphemeBoundaries.length - 1; index += 1) {
                selected.push({ start: graphemeBoundaries[index]!, end: graphemeBoundaries[index + 1]! });
              }
            }
            return Object.freeze(selected.map(({ start, end }) => {
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
      retain(shapeCache, shapeKey, result, shapeCacheLimit);
      return result;
    },
  });
  routeOrdinalTableSizes.set(service, () => routeOrdinals.size);
  scopeScanStats.set(service, () => Object.freeze({ scans: scopeScans, utf16Units: scopeUtf16Units }));
  return service;
}
