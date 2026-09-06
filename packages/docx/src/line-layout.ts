// DOCX line-layout engine — the pure segmentation + line-breaking + measurement
// kernel that both the paginator and the paint pass call to turn a paragraph's
// runs into laid-out lines and line-box heights (ECMA-376 §17.3.1.x line
// spacing, §17.3.1.37 tabs, §17.6.5 docGrid, §17.15.1.58–.60 kinsoku, §17.3.2.26
// script/font axes, §17.3.2.33 small caps, §17.8.3.10 font classification).
//
// Lifted verbatim out of renderer.ts along the domain phase boundary that the B2
// text-layout unification established (paragraph #684/#689, table #693, textbox
// #697): everything here MEASURES (it may touch a Canvas 2D context to call
// measureText / set ctx.font) but never DRAWS, mutates body acquisition state,
// paginates, or
// registers floats — those stay in renderer.ts, which imports this module. The
// split is one-directional at runtime: renderer.ts → line-layout.ts. Which
// Compatibility projections are owned by the reviewed modules under
// `layout/*-compatibility.ts`; this kernel consumes their narrow decisions.

import type {
  DocParagraph, DocRun, DocxTextRun, FieldRun,
  LineSpacing, TabStop, DocxRunBorder, DocSettings, EmphasisMark,
} from './types';
import type { CanvasFontRoute, KinsokuRules, HyperlinkTarget, NumberFormat, Duotone, ResolvedFontMetric } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  crossRunKinsokuRetract,
  isCjkBreakChar,
  isUax14NoBreakPair,
  containsSeaScript,
  isGraphemeFillText,
  isDictionarySeaText,
  seaMixedBreakOffsets,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  formatOrdinalNumber,
  parseFieldFormatSwitch,
  formatDateTimePicture,
  parseDateTimePictureSwitch,
  fontAdvanceBiasEm,
  normalizeFontMetricFamily,
  canvasFontString,
  intendedSingleLinePx,
  correctLineMetrics,
} from '@silurus/ooxml-core';
import { groupFitTextRegions, type FitTextRun } from './fit-text.js';
import {
  type FloatRect,
  MIN_LINE_GAP,
  prepareFloatWrap,
  computePreparedLineFloatWindow,
  wordMinLineStartPx,
  type PreparedFloatWrap,
} from './float-layout.js';
import { mathFallbackText } from './layout/math-fallback-text.js';
import {
  cloneSegmentsForLinePass,
  convergeLineWrap,
} from './layout/line-wrap-convergence.js';
import type { LayoutServices, NumberingMarkerShapeInput } from './layout/types.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './layout/measurement-capabilities.js';
import type {
  ParagraphAcquisitionRun,
  ParagraphAcquisitionInput,
  ParagraphLayoutSource,
  ParagraphLayoutRun,
  ParagraphTextBearingRun,
  GlyphInkBounds,
  TextLayoutService,
  TextShapeRequest,
  TextShapeSpan,
  FontScriptSlot,
} from './layout/text.js';
import {
  calcEffectiveFontPx,
  EAST_ASIAN_RE,
  nextTabStop,
  nextTabStopRtl,
  shapeRunToDocRun,
} from './layout/text.js';
import type { MathLayoutResource } from './layout/resources.js';
import {
  wordDegenerateLineSpacingIsSingle,
  wordEastAsianGridLineCells,
  wordFarEastSingleLinePx,
  wordGridAtLeastLineHeightPx,
  wordUseFeLayoutInheritedGridHeightPx,
  wordCandidateFitWidthPx,
  wordJustifiedCandidateFitAllowancePx,
  wordIsOverflowPunctuation,
  wordDocumentCharacterCompressionApplies,
  wordJapanesePunctuationRetainedExtentPt,
  wordMsMinchoEmptyEastAsianMarkSingleLinePx,
  wordSnapToCharsEastAsianCellCount,
  wordSourceRunSpaceContinuesSequence,
  wordBalancedConsecutiveSpaceCellApplies,
  wordBalancedLinesAndCharsGridDeltaFactor,
  wordBalancedSpaceCellAdjustmentApplies,
  wordIdeographicSpaceLineEndAllowanceCount,
  wordUniformRunPositionPaintPt,
  wordUseFeLayoutParagraphMarkGridAdvancePx,
  wordExternalLinkSyntaxBreakOffsets,
} from './layout/line-compatibility.js';
import { wordNeutralAttachesToActiveScript } from './layout/script-compatibility.js';

export interface LineBoundary {
  segIndex: number;
  charOffset: number;
}

interface LayoutSegSource {
  src?: LineBoundary;
  /** Parser-boundary run occurrence retained independently of mutable source
   * objects. Unlike `src.segIndex` (the flattened segment stream), this remains
   * the original paragraph run index through line splitting. */
  sourceRunIndex?: number;
}

export interface LayoutTextSeg extends LayoutSegSource {
  text: string;
  /** §17.3.2.26 script slot selected by the authoritative shaping service. */
  script?: FontScriptSlot;
  /** Internal §17.6.5 snapToChars allocation retained from measure to paint. */
  snapGridClass?: 'eastAsia' | 'latin' | 'complexScript';
  snapGridNaturalWidthPx?: number;
  snapGridLeadingPadPx?: number;
  snapGridTrailingPadPx?: number;
  snapGridCellPitchPx?: number;
  /** Registered Word projection of ECMA-376 §17.15.3.3 onto a `linesAndChars`
   * grid: half of `charSpace` for SBCS and the observed space classes, full
   * delta for DBCS. Present only when width balancing is enabled. */
  widthBalanceGridDeltaFactor?: 0.5 | 1;
  /** Registered compatibility projection for a two-or-more authored U+0020 sequence.
   * The flag preserves sequence membership through source/split boundaries;
   * the adjustment replaces each selected route's natural space with half of
   * its East-Asian ideographic cell. */
  widthBalanceSpaceSequence?: true;
  widthBalanceSpaceAdjustmentPt?: number;
  /** Internal marker assigned once after queue construction: this segment is
   * part of the paragraph-final U+3000-only suffix. */
  paragraphFinalIdeographicSpaceTail?: true;
  /** Number of trailing U+3000 characters owned by this segment. Kept
   * separately from the suffix-wide count so a source seam cannot move visible
   * text into the tail when the segment is split for line breaking. */
  paragraphFinalIdeographicSpaceLocalCount?: number;
  paragraphFinalIdeographicSpaceCount?: number;
  paragraphFinalIdeographicSpaceTailStart?: true;
  /** Zero-advance anchor-character placeholder: contributes run metrics to the
   * line box but paints no glyph. */
  metricOnly?: true;
  /** The run participates in Far East line-grid metrics despite containing no
   * East Asian code point. This covers an East-Asian anchor host and the
   * w:useFELayout + rFonts@hint=eastAsia compatibility path. */
  metricEastAsian?: true;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** ECMA-376 §17.3.2.40 `<w:u w:val>` — raw ST_Underline (§17.18.99) style; the
   *  renderer maps it to DrawingML §20.1.10.82 for `core.drawUnderline`. Absent
   *  ⇒ plain single rule. */
  underlineStyle?: string;
  /** ECMA-376 §17.3.2.40 `<w:u w:color>` — underline-only colour (hex 6 or
   *  `auto`). Absent ⇒ the underline follows the glyph colour. */
  underlineColor?: string;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  fontRoute?: CanvasFontRoute;
  /** Exact local face selected during async document loading. The family above
   * is its isolated alias; this measured ratio supplies the design-line floor
   * without a version-specific font constant. */
  resolvedLineHeightRatio?: number;
  resolvedEastAsianLineHeightRatio?: number;
  vertAlign: 'super' | 'sub' | null;
  measuredWidth: number;  // px (set during layout)
  /** A2 text authority captured during segmentation; production text width and
   * metrics are resolved through this same service during line layout. */
  textLayoutService?: TextLayoutService;
  textShapeRequest?: TextShapeRequest;
  /** Contextually shaped grapheme geometry from the authoritative text service. */
  shapedClusters?: readonly Readonly<{
    range: Readonly<{ start: number; end: number }>;
    offsetPt: number;
    advancePt: number;
  }>[];
  /** Tight selected-face ink retained by the authoritative shape call that
   * also produced `shapedClusters`. */
  selectedFaceInkBounds?: GlyphInkBounds;
  /** Selected-face font box retained by that same authoritative shape call.
   * Decorations and highlighting consume this instead of shaping the placed
   * run a second time. */
  selectedFaceFontBox?: Readonly<{ ascentPt: number; descentPt: number }>;
  /** False when this segment starts inside the preceding grapheme cluster. */
  breakBefore?: boolean;
  smallCaps?: boolean;
  /** This segment is GLUED to the preceding one (no inter-segment break): they
   *  are case-pieces of the same word emitted at different sizes for small caps
   *  (§17.3.2.33) — e.g. "I"(full)+"NTRODUCTION"(reduced). The line breaker must
   *  not start a new line before a glued segment; it retracts the whole glued
   *  group instead, so a small-caps word never splits across lines. */
  joinPrev?: boolean;
  /** Non-negotiable CT_R/noBreakHyphen seam. Unlike kinsoku/UAX glue, this
   * remains atomic even when either side otherwise exposes CJK/SEA breaks. */
  hardJoinPrev?: true;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten / 圏点) mark stamped on
   *  every non-space character of this segment (§17.18.24 ST_Em). The renderer
   *  paints it per glyph after the text; it does not affect layout metrics. */
  emphasisMark?: EmphasisMark;
  /** ECMA-376 §17.3.2.32 `<w:shd w:fill>` — run shading fill (hex 6). Painted as
   *  a solid rect behind the glyphs; also the effective background that an
   *  automatic text color resolves against. */
  background?: string | null;
  /** ECMA-376 §17.3.2.6 — run carries `<w:color w:val="auto"/>`. The glyph
   *  color is resolved from {@link LayoutTextSeg.background} for contrast
   *  (implementation-defined black/white pick; no normative algorithm). */
  colorAuto?: boolean;
  /** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border (box) around the text. */
  border?: DocxRunBorder | null;
  /** Ruby annotation rendered in a small font directly above this segment. */
  ruby?: { text: string; fontSizePt: number; hpsRaisePt?: number };
  /** Track-changes revision attached to this run (insertion / deletion /
   *  moveFrom / moveTo). */
  revision?: { kind: 'insertion' | 'deletion' | string; author?: string };
  /** Markup-view revision decoration facts (set only when the layout variant
   *  has `showTrackedChanges`): the revision kind plus the resolved stable
   *  author colour. Read by the retained decoration planner to synthesize the
   *  author-coloured underline (insertion/moveTo) or strikethrough
   *  (deletion/moveFrom), per the `word-track-change-decoration` rule. */
  trackChangesMarkup?: Readonly<{ kind: string; authorColor: string }>;
  /** ECMA-376 §17.3.2.30 `<w:rtl>` — run carries right-to-left characteristics.
   *  When true the segment's text is treated as a strong-RTL embedding in the
   *  per-line bidi pass (so leading digits / neutrals resolve RTL). */
  rtl?: boolean;
  /** `word-rtl-run-ambiguous-class-override`: classify this segment's European digits
   *  (U+0030–0039) as Arabic-Number (AN) in the per-line bidi pass, so a date
   *  like "28-02-2026" in an Arabic complex-script run reorders to "2026-02-28"
   *  in the registered order (ECMA-376 §17.3.2.20 w:lang w:bidi). */
  digitsAsAN?: boolean;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`) DECLARED on the
   *  originating run, retained for a line-box design floor. The floor is read
   *  from the resolved font resource; the authored family name itself carries
   *  no geometry. */
  eaFloorFamily?: string | null;
  /** Exact Canvas route for the explicit East Asian design-line probe. */
  eaFloorRoute?: CanvasFontRoute;
  resolvedEaFloorLineHeightRatio?: number;
  resolvedEaFloorEastAsianLineHeightRatio?: number;
  /** This segment belongs to a DrawingML/WPS text body whose declared
   * eastAsia face contributes a design-line floor independent of glyph slot. */
  textBoxLineFloor?: boolean;
  textBoxVertical?: boolean;
  /** IX1 — the resolved hyperlink target of the originating run (ECMA-376
   *  §17.16.22 external `r:id` URL / §17.16.23 internal `w:anchor` bookmark),
   *  computed once per run in `buildSegments`. The text-layer consumes it for
   *  the clickable region; line layout also uses external-link syntax as a
   *  preferred break opportunity for otherwise unbreakable URL text. Absent
   *  for a non-link run. */
  hyperlink?: HyperlinkTarget;
  /** Parser-independent UTF-16 ranges occupied by authored
   * `<w:noBreakHyphen/>` glyphs. Neither edge is a legal line boundary. */
  noBreakRanges?: readonly Readonly<{ start: number; end: number }>[];
  /** Registered external-URL syntax breaks, as segment-local UTF-16 offsets. */
  externalLinkBreakOffsets?: readonly number[];
  /** This segment starts after a registered external-URL syntax break. */
  externalLinkBreakBefore?: true;
  /** ECMA-376 §17.3.2.34 `<w:snapToGrid>` — false opts this run out of the
   *  section character grid without changing paragraph line-grid policy. */
  snapToCharacterGrid?: boolean;
  /** ECMA-376 §17.3.2.35 `<w:spacing>` — character-spacing pitch in POINTS
   *  (signed), added after every character of the run. Applied as a per-glyph
   *  `ctx.letterSpacing` delta on BOTH measure and paint (measure==paint), on top
   *  of any docGrid / justify delta. Absent ⇒ 0. */
  charSpacing?: number;
  /** ECMA-376 §17.15.1.18 document-level full-width character compression.
   * Each entry belongs to one shaped grapheme and adjusts the advance after its
   * UTF-16 end offset. Keeping the complete list preserves contextual shaping
   * for consecutive punctuation/kana while retained clusters, wrapping, and
   * paint all consume the same per-cluster geometry. */
  punctuationCompressions?: readonly Readonly<{
    end: number;
    adjustmentPt: number;
  }>[];
  /** Effective `w:lang/@w:eastAsia` consumed by the isolated
   *  {@link wordIsOverflowPunctuation} compatibility projection. */
  eastAsiaLanguage?: string;
  /** ECMA-376 §17.3.2.43 `<w:w>` — horizontal glyph-width scale as a FRACTION
   *  (0.67 = 67%). Measured widths are multiplied by it and the paint pass draws
   *  under `ctx.scale(charScale, 1)`; decorations follow the scaled extent.
   *  Absent ⇒ 1 (100%). */
  charScale?: number;
  /** ECMA-376 §17.3.2.14 `<w:fitText>` — target width in TWIPS and optional
   *  link id (wire strings plus numeric synthetic inputs). All segments emitted
   *  from one tab-delimited source-run fragment retain the same fragment/region
   *  indices so script and small-caps splitting cannot create a new fit region. */
  fitTextVal?: number;
  fitTextId?: number | string;
  fitTextRegionIndex?: number;
  /** Flattened tab-delimited source-fragment index (historical field name). */
  fitTextRunIndex?: number;
  /** Scale-resolved gap shared by the canonical advance and paint paths. */
  fitTextPerGapPx?: number;
  /** Region residual carried after its final glyph; scale-resolved like the gap. */
  fitTextTrailingPadPx?: number;
  fitTextRegionStart?: boolean;
  fitTextRegionEnd?: boolean;
  /** ECMA-376 §17.3.2.24 `<w:position>` — baseline raise(+)/lower(−) in POINTS,
   *  applied as a y-offset to the glyphs and decorations without changing the
   *  font size. Absent ⇒ 0. */
  position?: number;
  /** Line-relative baseline shift in points, resolved when the line is closed.
   * When every metric-bearing item has the same `position`, half of the common
   * displacement is retained so its enlarged line box shares the surplus above
   * and below the glyphs. The authored, style-resolved value remains in
   * `position` for the retained model. */
  lineRelativePosition?: number;
  /** Whether shifted ink contributes to this retained segment's line extent.
   * False only for the fixed-line-count drop-cap compatibility projection. */
  positionExtendsLineBox?: boolean;
  /** ECMA-376 §17.3.2.19 `<w:kern>` — font-kerning threshold in POINTS (smallest
   *  kerned size). Sets `ctx.fontKerning` on measure and paint when the run's
   *  font size ≥ the threshold. Absent ⇒ kerning off (`ctx.fontKerning='none'`
   *  is NOT forced globally; only a threshold-satisfied run enables it). */
  kerning?: number;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vert>` — horizontal-in-vertical
   *  (縦中横). Set by {@link buildSegments} ONLY when the run declares `w:vert`
   *  AND the page is vertical (tbRl); the property is inert in a horizontal page,
   *  so the gate is folded in here at build time and the measure/paint passes just
   *  read this flag. When set, the whole segment occupies ONE cell along the
   *  vertical column (advance = 1em, NOT the per-glyph sideways width), with its
   *  characters drawn horizontally side by side across the column (§17.3.2.10,
   *  PDF-verified on sample-26). Absent ⇒ normal per-glyph vertical advance. */
  tateChuYoko?: boolean;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vertCompress>` — set alongside
   *  {@link tateChuYoko} when the run also declares `w:vertCompress`. Compresses
   *  the horizontally-laid-out run so it fits the line height. Only meaningful
   *  when {@link tateChuYoko} is set. */
  tateChuYokoCompress?: boolean;
  /** issue #1014 — set by {@link buildSegments} when this segment is drawn by the
   *  per-glyph upright-vertical (tbRl) path (`environment.verticalCJK`, and NOT a
   *  縦中横 cell). It gates the vo=Tr rotate-fallback INK-extent advance correction
   *  (`verticalRunInkExtraPx`) in the measure passes so the layout advance matches
   *  the ink-sized cell `drawVerticalRun` paints — measure == draw. Inert (0
   *  correction) for every font that does not under-report a rotate mark's advance,
   *  which is all of them except a Chrome substitute; absent on horizontal pages. */
  verticalRun?: boolean;
  /** Issue #797 — dictionary word-break offsets (seg-local UTF-16 indices, from
   *  core `seaWordBreakOffsets`) for a Thai/Lao/Khmer segment, which has no
   *  inter-word spaces. Populated by {@link layoutLines} for SEA text; the wrap
   *  path breaks such a segment only at one of these boundaries (never mid-word)
   *  and re-queues the tail with the offsets rebased. Absent ⇒ not SEA text, or
   *  Intl.Segmenter unavailable (falls back to grapheme-safe emergency split). */
  seaBreaks?: readonly number[];
}

/** Shaping and line allocation are derived from the current text slice. A
 * segment split for wrapping must not retain geometry from the parent slice;
 * measurement and addToLine recompute these facts from the new text. */
const RESET_SLICED_TEXT_MEASUREMENT = {
  shapedClusters: undefined,
  selectedFaceInkBounds: undefined,
  selectedFaceFontBox: undefined,
  snapGridClass: undefined,
  snapGridNaturalWidthPx: undefined,
  snapGridLeadingPadPx: undefined,
  snapGridTrailingPadPx: undefined,
  snapGridCellPitchPx: undefined,
} as const;

/** ECMA-376 §17.3.3.12 defines `hpsRaise` as the “distance [...] between the
 * phonetic guide base text and the phonetic guide text.” The absent case needs
 * selected-face ink and is therefore resolved by retainedRubyAscentReservePx. */
export function rubyAscentReservePx(
  rubySizePt: number,
  hpsRaisePt: number | undefined,
  scale: number,
  segment?: LayoutTextSeg,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
): number {
  if (hpsRaisePt != null) return hpsRaisePt * scale;
  if (!segment?.ruby || !ctx) {
    throw new Error(`Ruby at ${rubySizePt}pt without hpsRaise requires retained base and guide ink`);
  }
  if (segment.textLayoutService && segment.textShapeRequest) {
    const base = segment.textLayoutService.shape({
      ...segment.textShapeRequest,
      text: segment.text,
      fontSizePt: calcEffectiveFontPx(segment, scale),
      measure: true,
      clusterGeometry: false,
    });
    const guide = segment.textLayoutService.shape({
      ...segment.textShapeRequest,
      text: segment.ruby.text,
      fontSizePt: segment.ruby.fontSizePt * scale,
      measure: true,
      clusterGeometry: false,
    });
    if (base.inkBounds && guide.inkBounds) {
      return base.inkBounds.ascentPt + guide.inkBounds.descentPt;
    }
  }
  // Isolated line-layout callers may not carry a service snapshot. Canvas
  // actual ink under the same selected route is still authoritative geometry;
  // retain it here instead of restoring the former font-size ratio.
  const previousFont = ctx.font;
  try {
    ctx.font = buildFont(
      segment.bold, segment.italic, calcEffectiveFontPx(segment, scale),
      segment.fontFamily, fontFamilyClasses, segment.fontRoute,
    );
    const base = ctx.measureText(segment.text);
    ctx.font = buildFont(
      segment.bold, segment.italic, rubySizePt * scale,
      segment.fontFamily, fontFamilyClasses, segment.fontRoute,
    );
    const guide = ctx.measureText(segment.ruby.text);
    if (
      Number.isFinite(base.actualBoundingBoxAscent)
      && Number.isFinite(guide.actualBoundingBoxDescent)
    ) {
      return base.actualBoundingBoxAscent + guide.actualBoundingBoxDescent;
    }
  } finally {
    ctx.font = previousFont;
  }
  throw new Error('Ruby without hpsRaise requires retained base and guide ink');
}

/**
 * Horizontal tab. Width is resolved during layout against paragraph tab stops
 * (or the default 36pt interval if no explicit stop is configured).
 */
export interface LayoutTabSeg extends LayoutSegSource {
  isTab: true;
  fontSize: number;  // pt — for line-height purposes
  measuredWidth: number;
  /** tab leader to fill the gap (e.g. TOC dot leaders); set during layout. */
  leader?: TabStop['leader'];
  /** Alignment selected from the effective stop during layout. */
  resolvedAlignment?: TabStop['alignment'];
  /** Bold/italic of the run carrying the tab (ECMA-376 §17.3.1.37 — the leader
   *  characters take the formatting of the tab's run, e.g. a bold TOC1 entry's
   *  dot leader is bold). Threaded so {@link drawTabLeader} can match the font. */
  bold?: boolean;
  italic?: boolean;
  /** ECMA-376 §17.3.3.23 `<w:ptab>` — when set, this is an ABSOLUTE-position tab.
   *  It ignores the paragraph's custom tab stops and the default-tab interval and
   *  advances to a position derived from `alignment` (§17.18.71) + `relativeTo`
   *  (§17.18.73). Absent ⇒ an ordinary `<w:tab>` resolved against tab stops. */
  ptab?: {
    alignment: 'left' | 'center' | 'right';
    relativeTo: 'margin' | 'indent';
  };
}

export interface LayoutImageSeg extends LayoutSegSource {
  /** Zip path of the blip — also the `'imagePath' in seg` discriminant that
   *  distinguishes an image segment from text/math/tab segments. */
  imagePath: string;
  /** MIME type of the blip at {@link LayoutImageSeg.imagePath}. */
  mimeType: string;
  widthPt: number;
  heightPt: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** true = wp:anchor: skip inline flow, draw at absolute page coords */
  anchor: boolean;
  anchorXPt: number;
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** When set, pixels matching this hex color are replaced with alpha=0 before drawing. */
  colorReplaceFrom?: string;
  /** ECMA-376 §20.1.8.23 `<a:duotone>` recolour (two endpoint colours). Part of
   *  the image cache key so the recoloured raster is looked up (draws through
   *  the same `imageKey(imagePath, colorReplaceFrom, duotone)` the prefetch used). */
  duotone?: Duotone;
  /** ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` opacity as 0..1. When < 1 the
   *  inline draw multiplies `globalAlpha` by it. `undefined` ⇒ fully opaque. */
  alpha?: number;
  /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-rectangle crop (signed fractions of
   *  the decoded bitmap). When present the draw paths use the 9-arg
   *  `drawImage` to blit only `[l, t, 1−r, 1−b]` of the bitmap into the display
   *  box. `undefined` ⇒ draw the full bitmap. */
  srcRect?: { l: number; t: number; r: number; b: number };
  /** ECMA-376 §21.2 — when set, this "image" box is actually a DrawingML chart.
   *  The box is sized like a picture (via {@link LayoutImageSeg.widthPt}/
   *  {@link LayoutImageSeg.heightPt}) and painted with the shared `renderChart`
   *  instead of blitting a bitmap: an inline chart seg flows with the text and
   *  is drawn at its flow position; an anchored chart seg (`anchor: true`,
   *  §20.4.2.3) is zero-width in the flow and retained at its absolute page
   *  box by anchor acquisition. `imagePath`/`mimeType` are
   *  empty sentinels for a chart seg — no blip is fetched (the bitmap-prefetch
   *  walk keys off `run.type === 'image'` and never sees a chart run). */
  chart?: true;
  chartResourceKey?: string;
  /** ECMA-376 §20.4.2.8 — this image-shaped line segment reserves the inline
   * WPS shape's extent; paint is owned by the paragraph's retained drawing. */
  inlineShape?: true;
  /** Parser-private retained placeholder for a recognized payload whose
   * package part is unavailable. It participates in line/anchor geometry but
   * never enters the paint resource registry. */
  unavailableResourceKind?: 'image' | 'chart';
  measuredWidth: number;
}

/** An inline OMML equation. Measured + drawn via the core math engine. */
export interface LayoutMathSeg extends LayoutSegSource {
  math: true;
  mathResourceKey: string;
  mathMetadata?: MathLayoutResource;
  display: boolean;
  fontSize: number;  // pt
  color: string | null;
  /** Plain-text fallback used when the async math renderer has not prepared an image. */
  fallbackText: string;
  measuredWidth: number;
  /** px ascent/descent of the laid-out box at scale, cached during measurement. */
  mathAscent: number;
  mathDescent: number;
  /** ECMA-376 §22.1.2.88 `m:oMathPara/m:jc` — per-instance justification of a
   *  display equation (ST_Jc math). `undefined` for inline math; the renderer
   *  resolves the document default (`mathDefJc`, spec default `centerGroup`). */
  jc?: string;
}

/** Sentinel that forces a new line when encountered in layoutLines. */
export interface LayoutLineBreak extends LayoutSegSource {
  lineBreak: true;
  fontSize: number;  // pt — used to set line height on empty lines
  measuredWidth: 0;
}

export type LayoutSeg = LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutLineBreak | LayoutTabSeg;

export interface LayoutLine {
  segments: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[];
  height: number;  // pt — max fontSize on line (for empty-line sizing fallback)
  ascent: number;  // px — fontBoundingBoxAscent (font-metric, stable per font+size)
  descent: number; // px — fontBoundingBoxDescent
  /** Baseline box contributed by segments that paint inline ink. A floating
   *  anchor host can reserve the line's ascent/descent while contributing no
   *  glyph; in that mixed case these exclude the metric-only placeholder. */
  visibleAscent?: number;
  visibleDescent?: number;
  visibleIntendedSingle?: number;
  /** px — intended single-line height (max over segments of the selected
   * resource ratio, with the established compatibility registry as fallback). */
  intendedSingle: number;
  /** px — DESIGN grid-count height: the max over segments of each run's
   *  Word-compatible single-line height (a resolved resource's design height,
   *  or the generic East Asian fallback). Feeds docGrid cell
   *  counting without depending on a substituted face's Canvas box
   *  (§17.6.5; sample-9/sample-52). */
  gridCountSingle: number;
  /** Additional horizontal offset (px) from paraX, caused by wrap-around floats. */
  xOffset: number;
  /** Effective available width (px) for this line after float exclusion. */
  availWidth: number;
  /** When wrap context is active, the absolute canvas Y where this line begins. */
  topY?: number;
  /** Set when at least one segment on this line carries a ruby annotation —
   *  enables docGrid pitch snapping in lineBoxHeight. */
  hasRuby?: boolean;
  /** §17.6.5 — a text segment on this line contains an East Asian code point
   *  (EAST_ASIAN_RE), enabling docGrid line-cell rounding. Undefined/false for
   *  synthesized textless lines. */
  eastAsian?: boolean;
  /** ECMA-376 §17.3.3.1 — this line is terminated by a MANUAL line break
   *  (`<w:br w:type="textWrapping"/>`). In a justified (`both`) paragraph it is
   *  the end of a logical line and must be left-aligned, not stretched — exactly
   *  like the paragraph's final line (§17.18.44). */
  endsWithBreak?: boolean;
  /** Issue #908 — the consumed-content END boundary of this line in the ORIGINAL
   *  `segs` stream of the layoutLines call that produced it (see LineBoundary).
   *  Break-aware: a manual-break-terminated line consumes its sentinel. Laying out
   *  the suffix from this boundary (same width, firstIndent 0) reproduces the
   *  following lines exactly; at a different width it re-wraps — the remainder
   *  re-measure seam. */
  consumedEnd?: LineBoundary;
}

/** Additional context passed to layoutLines so it can honor floats on the current page. */
export interface WrapLayoutCtx {
  startPageY: number;   // absolute canvas Y where the first line should start
  paraX: number;        // absolute canvas X of the paragraph's INDENTED text left edge
  /** Absolute canvas X of the paragraph's raw COLUMN left edge. Distinct from
   *  `paraX` when the paragraph has a left indent: the topAndBottom wrap gate
   *  (§20.4.2.20 full-column block) is scoped to the COLUMN band, while the
   *  square side-gap math (§20.4.2.17) is scoped to the indented `paraX` band. */
  columnXPt: number;
  /** Absolute px width of the paragraph's raw COLUMN band. See columnXPt. */
  columnWidthPt: number;
  floats: FloatRect[];  // legacy float geometry supplied directly by renderer paths
  /** Minimum clear side-gap for an anchor-host-only paragraph mark. Such a
   *  zero-advance metric placeholder preserves the anchor character's line box,
   *  but is not inline content and therefore keeps the pilcrow-em threshold
   *  instead of the 1-inch content-line threshold (issue #676). */
  paragraphMarkLineStartWidth?: number;
  /** Placement-aware wrap boundary used by paragraph measurement. */
  lineWindow?: (input: {
    topYPt: number;
    minimumStartWidthPt: number;
    /** `word-square-line-start-one-inch`, active only for a square object. */
    squareMinimumStartWidthPt?: number;
    probeHeightPt: number;
    paragraphXPt: number;
    maximumWidthPt: number;
    /** The paragraph's raw COLUMN band, scoping the topAndBottom gate
     *  (§20.4.2.20 / §17.6.4). Distinct from paragraphXPt/maximumWidthPt (the
     *  indented text band the square side-gap math uses). */
    columnXPt: number;
    columnWidthPt: number;
  }) => {
    topYPt: number;
    xOffsetPt: number;
    maximumWidthPt: number;
  };
  /** Page/reference band used by ST_WrapText `largest` (§20.4.3.7). */
  referenceXPt?: number;
  referenceWidthPt?: number;
  /** Reading order of the first line intersecting a centered `largest` object. */
  readingDirection?: 'ltr' | 'rtl';
  /** Per-line box-height resolver (line natural ascent+descent → total px box height).
   *  `gridCountSinglePx` (the line's design grid-count height) keeps the
   *  float-wrap advance consistent with the final render's docGrid cell count. */
  lineBoxH: (ascentPx: number, descentPx: number, hasRuby?: boolean, intendedSinglePx?: number, eastAsian?: boolean, gridCountSinglePx?: number) => number;
  /** Hard cap on Y to keep layout from running past the page. */
  pageH: number;
}

/** Document-grid context passed to line-box computation.  When the section's
 *  `w:docGrid` is "lines"/"linesAndChars" with a positive pitch (ECMA-376
 *  §17.6.5), auto line spacing multiplies against the grid pitch instead of
 *  the font's natural line height. Without this, a 56-pt heading with
 *  lineRule="auto" value=4.33 would claim 56×1.25×4.33 ≈ 303pt of vertical
 *  space; with this, it claims max(natural, 18pt × 4.33) ≈ 78pt — matching
 *  Word's rendering on grids typical of Japanese/Chinese templates. */
export interface DocGridCtx {
  /** "default" | "lines" | "linesAndChars" | "snapToChars" */
  type: string | null | undefined;
  /** Grid pitch in pt (already converted from twips in the parser). */
  linePitchPt: number | null | undefined;
  /** Full §17.6.5 character pitch in pt (Normal-style size + charSpace delta). */
  characterPitchPt?: number | null;
  /** ECMA-376 §17.6.5 `<w:docGrid w:charSpace>` divided by 4096. This is a
   *  flat-point character-pitch delta, independent of font size. `linesAndChars`
   *  adds it to every character. */
  charSpacePt?: number | null;
}

/** Page/document values that can change segment text or vertical-text behavior.
 * Canvas/font measurement belongs to the caller's TextMeasurer instead. The
 * document-level East Asian flag is used only for content-less paragraph-mark
 * metrics; content lines use ParagraphLayoutContext.hasEastAsianText. */
export interface LineLayoutEnvironment {
  readonly pageIndex: number;
  readonly totalPages: number;
  readonly displayPageNumber?: number;
  readonly pageNumberFormat?: NumberFormat;
  readonly currentDateMs?: number;
  /** ECMA-376 §17.13.5 tracked-change view. `true` = markup view: revision
   * content stays visible for author-coloured decoration. Absent/false =
   * final view: deleted (`w:del`) and moved-away (`w:moveFrom`) runs produce
   * no segments, so line breaking sees the accepted document state. */
  readonly showTrackedChanges?: boolean;
  /** Markup-view author → stable palette colour (layout/track-changes.ts
   * first-appearance policy over the compatibility palette). Present only
   * when the markup variant is being built. */
  readonly revisionAuthorColor?: (author?: string) => string;
  readonly noteNumbers?: ReadonlyMap<string, number>;
  readonly noteReferenceNumber?: number;
  readonly verticalCJK?: boolean;
  /** ECMA-376 Part 4 §14.8.3.50 w:useFELayout compatibility switch. */
  readonly useFeLayout?: boolean;
  /** §17.15.3.3 w:balanceSingleByteDoubleByteWidth document compatibility switch. */
  readonly balanceSingleByteDoubleByteWidth?: boolean;
  readonly resolvedLocalFonts?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly layoutServices?: LayoutServices;
  readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService;
  /** ECMA-376 §17.15.1.18 document-wide full-width character compression. */
  readonly characterSpacingControl?: string;
  /** False only when `w:framePr` specifies a drop cap with a fixed `w:lines`;
   * the authored frame height remains authoritative even when glyph paint is
   * lowered beyond it. Folded into retained text segments during acquisition. */
  readonly positionExtendsLineBox?: boolean;
}

// ── Math (OMML) rendering via MathJax ───────────────────────────────────────
// Each equation is converted OMML AST -> MathML -> MathJax SVG, then rasterized
// to an auxiliary Canvas once (async, before pagination). Layout reads cached
// em-extents synchronously; drawing blits the Canvas. Skipped entirely for
// math-free documents.
/** Arabic-script faces that hosts rarely ship; we substitute them with Noto
 *  Naskh/Sans Arabic web fonts (see DOCX_GOOGLE_FONTS in document.ts — this
 *  list MUST mirror the Arabic entries there). A source run whose font is one
 *  of these contains both Arabic and Latin/digit glyphs in one requested face,
 *  so the fallback chain must keep both scripts stylistically
 *  consistent (Arabic substitute first, serif Latin companion before the sans
 *  generics) rather than letting Latin/digits leak to a CJK sans face. */
export const ARABIC_SUBSTITUTE_FONTS = new Set([
  'sakkal majalla',
  'traditional arabic',
  'simplified arabic',
  'arabic typesetting',
  'univers next arabic',
  'noto naskh arabic',
  'noto sans arabic',
]);

/** Naskh-style traditional Arabic faces ship a serif Latin companion; the
 *  geometric/modern ones pair with a sans Latin. Drives whether an Arabic-font
 *  run's Latin+digits route to Noto Naskh Arabic (serif-like) or Noto Sans
 *  Arabic, and which Latin serif/sans companion follows. */
export const NASKH_SERIF_ARABIC_FONTS = new Set([
  'sakkal majalla',
  'traditional arabic',
  'simplified arabic',
  'arabic typesetting',
  'noto naskh arabic',
]);

export function isArabicSubstituteFont(family: string): boolean {
  return ARABIC_SUBSTITUTE_FONTS.has(family.toLowerCase());
}

/** Quote each family for a CSS font-family list. */
export function quoteAll(names: readonly string[]): string {
  return names.map((n) => `"${n}"`).join(', ');
}

/** Generic Arabic web-font fallbacks (loaded when `useGoogleFonts` is on). */
export const ARABIC_TAIL_SANS = ['Noto Naskh Arabic', 'Noto Sans Arabic'] as const;

/**
 * Sans fallback TAIL (everything after the requested face) for a Latin/CJK run.
 *
 * - `cjk`: the document's CJK language inferred from the font name, or `null`
 *   for a plain Latin face — in which case the existing Japanese system-font
 *   companions (Hiragino Sans / Meiryo) lead, preserving the long-standing JP
 *   default. For a non-JP CJK language the matching Noto CJK leads so shared
 *   Han glyphs take that language's shapes (see core/fonts/scripts.ts).
 *
 * Order: [CJK companions] → Arabic → non-CJK scripts (Hebrew/Thai/Devanagari,
 * Cyrillic via Noto Sans) → `sans-serif`. The non-CJK scripts have no Han
 * collision so their position is immaterial; they sit before the generic so
 * the browser's per-glyph fallback can reach them.
 */
export function sansTail(cjk: ReturnType<typeof classifyCjkFont>): string {
  const cjkPart =
    cjk && cjk !== 'jp'
      ? cjkFallbackChain(cjk, 'sans')
      : // JP / stray-CJK sans faces: historical system-font hints, then the Noto
        // CJK siblings so a CJK glyph still resolves on hosts lacking them.
        ['Noto Sans JP', 'Hiragino Sans', 'Meiryo', ...cjkFallbackChain('jp', 'sans').slice(1)];
  // A Latin (non-CJK) sans font must fall back to a LATIN sans for its
  // letters/digits — otherwise the browser grabs them from a Japanese Gothic
  // (wider, CJK-tuned Latin), widening Latin runs. Lead with Latin sans faces;
  // the CJK gothic faces follow for any stray CJK glyph. (Mirrors serifTail.)
  if (cjk == null) {
    return `${quoteAll([...NON_CJK_SANS_FALLBACKS, 'Arial', 'Helvetica', 'Liberation Sans', ...cjkPart, ...ARABIC_TAIL_SANS])}, sans-serif`;
  }
  return `${quoteAll([...cjkPart, ...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
}

/** Serif counterpart of {@link sansTail}. */
export function serifTail(cjk: ReturnType<typeof classifyCjkFont>): string {
  const cjkPart =
    cjk && cjk !== 'jp'
      ? cjkFallbackChain(cjk, 'serif')
      : // JP / stray-CJK serif faces: historical mincho system hints, then Noto
        // serif CJK siblings.
        [
          'Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho',
          'Noto Serif JP', ...cjkFallbackChain('jp', 'serif').slice(1),
        ];
  // A Latin (non-CJK) serif font (e.g. Century) must fall back to a LATIN serif
  // for its letters/digits. If the CJK mincho faces lead, the browser's
  // per-glyph fallback grabs Latin glyphs from a Japanese Mincho (e.g. Hiragino
  // Mincho ProN on macOS) whose Latin is ~15-18% wider, widening every Latin
  // run and forcing spurious line wraps. Lead with Latin serif faces; the CJK
  // mincho faces follow so a stray CJK glyph in a Latin-font run still resolves.
  if (cjk == null) {
    return `${quoteAll([...NON_CJK_SERIF_FALLBACKS, 'Times New Roman', 'Cambria', 'Liberation Serif', ...cjkPart, ...ARABIC_TAIL_SANS])}, serif`;
  }
  return `${quoteAll([...cjkPart, ...ARABIC_TAIL_SANS, ...NON_CJK_SERIF_FALLBACKS])}, serif`;
}

/** Resolve a requested font-family name to a CSS font-family string with
 *  appropriate fallback chain.
 *
 *  Classification priority:
 *  1. `fontFamilyClasses` map (from `word/fontTable.xml` §17.8.3.10):
 *     - "roman"      → serif
 *     - "swiss"      → sans-serif
 *     - "modern"     → monospace only for `pitch="fixed"` (§17.8.3.14), else
 *                      fall through to step 2
 *     - "script"/"decorative" → sans-serif fallback
 *     - "auto" / absent       → fall through to step 2
 *  2. Name-pattern matching (fallback for fonts absent from fontTable, or
 *     where fontTable says "auto"). Retained as a safety net for theme fonts
 *     and system fonts that OOXML docs do not list in fontTable.xml.
 */
/**
 * Per-document memo for {@link normalizeFontFamily}. The regex/classifier work
 * inside is a pure function of `(family, fontFamilyClasses)`, and
 * `fontFamilyClasses` is a stable per-document object (body acquisition threads
 * `doc.fontFamilyClasses` — one identity per render). Keying the outer WeakMap on
 * that object identity gives per-doc caching with zero call-site churn (both
 * callers already pass `fontFamilyClasses`) and no leak: the inner
 * family→result Map is collected with the document's classes object. Same idiom
 * as `sheetAxisCache`. Chosen over threading an explicit cache
 * param through buildFont because both call sites already carry the classes
 * object, so identity-keying needs no signature changes anywhere.
 */
export const fontFamilyNormalizeCache = new WeakMap<Record<string, string>, Map<string, string>>();

/** Companion to {@link fontFamilyNormalizeCache}: maps a `fontFamilyClasses`
 *  object (stable per-document identity) to the sibling per-font PITCH map
 *  (ECMA-376 §17.8.3.14 `<w:pitch>`: font name → "fixed" | "variable" |
 *  "default"). `normalizeFontFamily` reads it to decide whether a
 *  `family="modern"` (§17.8.3.10) face is genuinely monospace: only "fixed"
 *  (§17.18.66 Fixed Width) is. Keyed on the classes object so the pitch threads
 *  for free through every existing `fontFamilyClasses` call site — exactly like
 *  the normalize cache — with no second map plumbed through the renderer. */
export const fontFamilyPitchesByClasses = new WeakMap<
  Record<string, string>,
  Record<string, string>
>();

/** Bind the §17.8.3.14 pitch map to the §17.8.3.10 classes object and return the
 *  classes object (defaulting to `{}`). Call at each renderer site that
 *  materializes a document's `fontFamilyClasses` for threading, so the classifier
 *  can read a `modern` face's pitch without a second map plumbed through. */
export function fontClassesWithPitches(
  classes: Record<string, string> | undefined,
  pitches: Record<string, string> | undefined,
): Record<string, string> {
  const c = classes ?? {};
  if (pitches && Object.keys(pitches).length > 0) {
    fontFamilyPitchesByClasses.set(c, pitches);
  }
  return c;
}

export function normalizeFontFamily(
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
): string {
  const perDoc =
    fontFamilyNormalizeCache.get(fontFamilyClasses) ??
    (() => {
      const m = new Map<string, string>();
      fontFamilyNormalizeCache.set(fontFamilyClasses, m);
      return m;
    })();
  // `family` may be null; use a distinct sentinel key so a null lookup never
  // collides with a real family named "null".
  const key = family ?? '\0null';
  const cached = perDoc.get(key);
  if (cached !== undefined) return cached;
  // The pitch map is registered once against this stable per-document classes
  // identity, so the result remains a pure function of the memo key.
  const result = normalizeFontFamilyUncached(
    family,
    fontFamilyClasses,
    fontFamilyPitchesByClasses.get(fontFamilyClasses),
  );
  perDoc.set(key, result);
  return result;
}

export function normalizeFontFamilyUncached(
  family: string | null,
  fontFamilyClasses: Record<string, string>,
  fontFamilyPitches: Record<string, string> = {},
): string {
  if (!family) return sansTail(null);

  const escape = (s: string) => s.replace(/"/g, '\\"');
  const head = `"${escape(family)}"`;
  const lower = family.toLowerCase();

  // CJK language inferred from the font name (null for plain Latin faces). For a
  // non-JP CJK language the matching Noto CJK leads the fallback tail so shared
  // Han glyphs render with that language's shapes; see core/fonts/scripts.ts.
  const cjk = classifyCjkFont(family);

  // 0) Arabic-script faces substituted by Noto Naskh/Sans Arabic. A single
  //    Sakkal Majalla / Traditional Arabic run carries Arabic glyphs AND
  //    Latin letters/digits; the source run assigns both to that one face. The browser
  //    resolves each glyph against the chain in order, so the Arabic substitute
  //    MUST come first — otherwise the Latin/digit glyphs are grabbed by the
  //    first chain member that has them (e.g. the CJK "Noto Sans JP"), and
  //    Latin/digits render in a different, sans face than the Arabic. Keeping
  //    the Arabic substitute first makes Arabic+Latin+digits resolve from one
  //    coherent family.
  //
  //    Latin companion: traditional Naskh faces ship a serif Latin companion.
  //    Noto Naskh Arabic supplies the same script combination, so placing it
  //    first keeps Latin and digits stylistically consistent with the Arabic.
  //    "Noto Serif" is a safety net when Noto Naskh Arabic is unavailable;
  //    geometric Arabic faces instead pair with a sans Latin fallback.
  if (isArabicSubstituteFont(family)) {
    if (NASKH_SERIF_ARABIC_FONTS.has(lower)) {
      return `${head}, "Noto Naskh Arabic", "Noto Sans Arabic", "Noto Serif", "Noto Sans JP", "Hiragino Sans", serif`;
    }
    return `${head}, "Noto Sans Arabic", "Noto Naskh Arabic", "Noto Sans JP", "Hiragino Sans", sans-serif`;
  }

  // 1) Authoritative classification from word/fontTable.xml §17.8.3.10.
  const tableClass = fontFamilyClasses[family];
  if (tableClass && tableClass !== 'auto') {
    switch (tableClass) {
      case 'roman':
        return `${head}, ${serifTail(cjk)}`;
      case 'swiss':
        return `${head}, ${sansTail(cjk)}`;
      case 'modern': {
        // §17.8.3.10 `modern` is the "modern/monospace" typeface family, but the
        // family value classifies the DESIGN, not the pitch — §17.8.3.14
        // `<w:pitch>` states the actual pitch. Treat the face as monospace ONLY
        // when pitch is "fixed" (§17.18.66 Fixed Width). A "variable"
        // (proportional) modern face — e.g. Meiryo UI (`family="modern"`,
        // `pitch="variable"`), a condensed ~0.84em CJK sans — must NOT map to
        // Courier/monospace: that measures its CJK at a full 1.0em and over-wraps
        // table cells onto a spurious extra page (issue #855). "default" and an
        // omitted `<w:pitch>` (assumed "default" per §17.8.3.14) are likewise not
        // a fixed-width guarantee, so they fall through to the name-pattern /
        // CJK-sans path below. Genuine monospace faces (Courier, Consolas, 等幅)
        // are still caught there by name.
        if (fontFamilyPitches[family] === 'fixed') {
          if (cjk != null) {
            const cjkFallbacks = cjk === 'jp'
              ? ['Yu Gothic', 'YuGothic', 'Hiragino Sans', 'Meiryo', 'Noto Sans JP']
              : cjkFallbackChain(cjk, 'sans');
            return `${head}, ${quoteAll([...cjkFallbacks, 'Courier New'])}, monospace`;
          }
          return `${head}, "Courier New", monospace`;
        }
        break;
      }
      default:
        // script / decorative — fall through to name-pattern matching
        break;
    }
  }

  // 2) Name-pattern fallback for fonts absent from fontTable or classified
  //    "auto". The serif/sans/mono DECISION is the shared core classifier
  //    (`classifyFontGeneric`, §17.8.3.10-aligned name heuristic) that pptx and
  //    xlsx also route through — so all three renderers agree on the generic
  //    class. docx keeps its own richer fallback-chain construction (Latin-first
  //    ordering + per-language CJK chains + Arabic tail + JP system hints) below;
  //    only the regex-based decision is delegated here. Core's serif token set
  //    is a verified superset of docx's former serif tokens (it additionally
  //    detects e.g. Century/Palatino/Didot as serif and Consolas/Courier/等幅 as
  //    mono on the name path), so no prior serif/sans coverage is lost.
  const generic = classifyFontGeneric(family);
  if (generic === 'serif') {
    return `${head}, ${serifTail(cjk)}`;
  }
  if (generic === 'mono') {
    // Mirror the fontTable `modern` branch's monospace fallback. NEW for the
    // name path: core now detects consolas/courier/等幅 etc. as mono.
    return `${head}, "Courier New", monospace`;
  }

  // Japanese system-font hints (only meaningful for JP / Latin faces; a non-JP
  // CJK face skips these so its matching Noto CJK leads the tail).
  if (cjk == null || cjk === 'jp') {
    if (lower.includes('meiryo') || family.includes('メイリオ')) {
      return `${head}, "Meiryo UI", "Meiryo", ${sansTail(cjk)}`;
    }
    if (family.includes('游ゴシック') || /\byu\s*gothic\b/i.test(family) || lower.includes('yugothic')) {
      return `${head}, "Yu Gothic", "YuGothic", ${sansTail(cjk)}`;
    }
    if (lower.includes('ipa')) {
      return `${head}, "IPAexGothic", ${sansTail(cjk)}`;
    }
    if (lower.includes('segoe')) {
      return `${head}, "Segoe UI", ${quoteAll([...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
    }
  }
  return `${head}, ${sansTail(cjk)}`;
}

export function buildFont(
  bold: boolean,
  italic: boolean,
  sizePx: number,
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
  fontRoute?: CanvasFontRoute,
): string {
  if (fontRoute) return canvasFontString(fontRoute, sizePx, bold ? 700 : 400, italic ? 'italic' : 'normal');
  const w = bold ? 'bold' : 'normal';
  const s = italic ? 'italic' : 'normal';
  const f = normalizeFontFamily(family, fontFamilyClasses);
  return `${s} ${w} ${sizePx}px ${f}`;
}

/** Design single-line floor. Parsed resource geometry takes part through the
 * generic fields; the compatibility registry remains the fallback for authored
 * fonts whose bytes are unavailable to the browser. */
export function segmentIntendedSingleLinePx(
  segment: LayoutTextSeg,
  emPx: number,
  eastAsian = false,
): number {
  const resourceRatio = eastAsian
    ? segment.resolvedEastAsianLineHeightRatio ?? segment.resolvedLineHeightRatio ?? 0
    : segment.resolvedLineHeightRatio ?? 0;
  return Math.max(
    intendedSingleLinePx(segment.fontFamily, emPx, eastAsian),
    resourceRatio * emPx,
  );
}

export function segmentEastAsiaFloorSingleLinePx(
  segment: LayoutTextSeg,
  emPx: number,
  eastAsian = false,
): number {
  const resourceRatio = eastAsian
    ? segment.resolvedEaFloorEastAsianLineHeightRatio
      ?? segment.resolvedEaFloorLineHeightRatio
      ?? 0
    : segment.resolvedEaFloorLineHeightRatio ?? 0;
  return Math.max(
    intendedSingleLinePx(segment.eaFloorFamily, emPx, eastAsian),
    resourceRatio * emPx,
  );
}

export function getDefaultFontSize(para: ParagraphLayoutSource): number {
  for (const run of para.runs) {
    if (run.type === 'text') {
      return (run as unknown as DocxTextRun).fontSize;
    }
    if (run.type === 'field') {
      return (run as unknown as FieldRun).fontSize;
    }
  }
  if (typeof para.defaultFontSize === 'number') return para.defaultFontSize;
  return 10; // pt fallback
}

/** First text/field run's font family. Empty paragraphs fall back to the
 * paragraph's style-resolved default family. Resource lookup consumes the name
 * first; the legacy compatibility registry may use it only when unavailable. */
export function getDefaultFontFamily(
  para: ParagraphLayoutSource,
  eastAsian = false,
): string | null {
  for (const run of para.runs) {
    if (run.type === 'text') return (run as unknown as DocxTextRun).fontFamily;
    if (run.type === 'field') return (run as unknown as FieldRun).fontFamily;
  }
  if (eastAsian && para.defaultFontFamilyEastAsia) return para.defaultFontFamilyEastAsia;
  return para.defaultFontFamily ?? null;
}

/** Compatibility-registry floor for an empty paragraph. */
export function emptyIntendedSinglePx(
  para: ParagraphLayoutSource,
  scale: number,
): number {
  return intendedSingleLinePx(getDefaultFontFamily(para), getDefaultFontSize(para) * scale);
}

function emptyIntendedSingleForScriptPx(
  para: ParagraphLayoutSource,
  scale: number,
  eastAsian: boolean,
): number {
  return intendedSingleLinePx(
    getDefaultFontFamily(para, eastAsian),
    getDefaultFontSize(para) * scale,
    eastAsian,
  );
}

/** Code points whose presence marks a line as East Asian for docGrid line-cell
 *  rounding: CJK symbols/punctuation, Hiragana, Katakana, CJK Unified +
 *  Extension A, compatibility ideographs, Hangul, and fullwidth forms. Content
 *  test only — not a font-name heuristic (cf. packages/docx/CLAUDE.md). */
/** Per-character character-grid delta in px, before applying the grid's scope. */
export function gridCharDeltaPx(grid: DocGridCtx | undefined, scale: number): number {
  if (!grid || grid.charSpacePt == null) return 0;
  if (grid.type !== 'linesAndChars' && grid.type !== 'snapToChars') return 0;
  return grid.charSpacePt * scale;
}

/** Count of East-Asian (full-width) code points in `text` — the glyphs the
 *  character grid snaps to cells. Uses the same {@link EAST_ASIAN_RE} content
 *  predicate as docGrid line-cell rounding (no font-name heuristic). */
export function eaGlyphCount(text: string): number {
  let n = 0;
  for (const ch of text) if (EAST_ASIAN_RE.test(ch)) n++;
  return n;
}

/** Total character-grid delta gained by a segment. `linesAndChars` applies its
 *  authored pitch to every character. The non-`linesAndChars` branch preserves
 *  the renderer's pre-existing East-Asian-only fallback; full snap-to-character
 *  grid-unit allocation is handled by the block allocator below. */
export function gridSegDeltaPx(
  text: string,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  const deltaPx = gridCharDeltaPx(grid, scale);
  if (deltaPx === 0 || text.length === 0) return 0;
  const cps = [...text];
  if (grid?.type === 'linesAndChars') return cps.length * deltaPx;
  return eaGlyphCount(text) === cps.length ? cps.length * deltaPx : 0;
}

/** Resolve the per-glyph character-grid delta for one text segment. */
export function segmentCharacterGridDeltaPx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (seg.snapToCharacterGrid === false) return 0;
  // snapToChars allocates full cells/blocks in layoutLines. It is not a
  // per-glyph letter-spacing delta.
  if (grid?.type === 'snapToChars') return 0;
  if (grid?.type === 'linesAndChars' && seg.widthBalanceGridDeltaFactor !== undefined) {
    return gridCharDeltaPx(grid, scale) * seg.widthBalanceGridDeltaFactor;
  }
  const total = gridSegDeltaPx(seg.text, grid, scale);
  return total === 0 ? 0 : gridCharDeltaPx(grid, scale);
}

/** ECMA-376 §17.3.2.35 `<w:spacing>` — the per-GLYPH character-spacing pitch in
 *  px for a segment (its authored points × the paint scale). Unlike the docGrid
 *  delta this applies to EVERY code point of the run, not just East-Asian ones
 *  ("the amount of character pitch … added after each character in this run").
 *  0 when the run declares no `w:spacing`. */
export function charSpacingDeltaPx(seg: LayoutTextSeg, scale: number): number {
  // §17.3.2.14 fitText replaces cached §17.3.2.35 spacing with the resolved
  // region gap. The paint path already reads this authority.
  if (seg.fitTextPerGapPx !== undefined) return seg.fitTextPerGapPx;
  return effectiveCharacterSpacingPt(seg) * scale;
}

/** The uniform paint/measure pitch contributed by run-authored `w:spacing`.
 * Document-level punctuation compression is a one-time trailing-cell advance
 * adjustment, not a per-glyph Canvas letter-spacing value. */
export function effectiveCharacterSpacingPt(seg: LayoutTextSeg): number {
  return seg.charSpacing ?? 0;
}

export function punctuationCompressionTotalPt(seg: LayoutTextSeg): number {
  return seg.punctuationCompressions?.reduce(
    (sum, compression) => sum + compression.adjustmentPt,
    0,
  ) ?? 0;
}

/** §17.15.3.3 plus the registered Word space-sequence projection. A segment
 * created by splitTextForLayout contains at most one trailing U+0020 sequence;
 * slicing keeps the sequence flag, so prefixes/tails count only their retained
 * authored spaces. */
export function widthBalanceSpaceAdjustmentTotalPt(
  seg: LayoutTextSeg,
  characterGrid?: DocGridCtx,
): number {
  return widthBalanceSpaceAdjustmentForTextPt(seg, seg.text, characterGrid);
}

/** Retained-cluster projection of the same authored-space adjustment used by
 * {@link widthBalanceSpaceAdjustmentTotalPt}. Keeping this calculation on the
 * immutable segment fact makes line measurement, cluster hit geometry, RTL
 * whitespace anchoring, and paint-plan slicing consume one width authority. */
export function widthBalanceSpaceAdjustmentForTextPt(
  seg: LayoutTextSeg,
  text: string,
  characterGrid?: DocGridCtx,
): number {
  // Both properties replace the ordinary inline advance with their own
  // specification-defined region/cell width. Their interaction with Word's
  // proportional-space projection is outside the observation matrix, so keep
  // the preexisting override geometry intact in every retained consumer.
  if (seg.fitTextPerGapPx !== undefined || seg.tateChuYoko) return 0;
  if (!wordBalancedSpaceCellAdjustmentApplies(characterGrid?.type)) return 0;
  if (!seg.widthBalanceSpaceSequence || seg.widthBalanceSpaceAdjustmentPt === undefined) {
    return 0;
  }
  let count = 0;
  for (const character of text) if (character === ' ') count += 1;
  return count * seg.widthBalanceSpaceAdjustmentPt;
}

export function slicedPunctuationCompressions(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): LayoutTextSeg['punctuationCompressions'] {
  const sliced = seg.punctuationCompressions
    ?.filter((compression) => compression.end > start && compression.end <= end)
    .map((compression) => Object.freeze({
      end: compression.end - start,
      adjustmentPt: compression.adjustmentPt,
    }));
  return sliced && sliced.length > 0 ? Object.freeze(sliced) : undefined;
}

function slicedNoBreakRanges(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): LayoutTextSeg['noBreakRanges'] {
  const sliced = seg.noBreakRanges
    ?.filter((range) => range.start >= start && range.end <= end)
    .map((range) => Object.freeze({
      start: range.start - start,
      end: range.end - start,
    }));
  return sliced && sliced.length > 0 ? Object.freeze(sliced) : undefined;
}

function protectedNoBreakOffsets(seg: LayoutTextSeg): ReadonlySet<number> {
  return new Set(seg.noBreakRanges?.flatMap((range) => [range.start, range.end]) ?? []);
}

function legalTextSplitAtOrBefore(
  seg: LayoutTextSeg,
  proposed: number,
  minimum = 0,
): number {
  const protectedOffsets = protectedNoBreakOffsets(seg);
  return [0, ...graphemeClusterOffsets(seg.text), seg.text.length]
    .filter((offset, index, all) => all.indexOf(offset) === index)
    .filter((offset) => offset >= minimum && offset <= proposed && !protectedOffsets.has(offset))
    .at(-1) ?? 0;
}

/** Smallest prefix that consumes a hard source seam. It includes the complete
 * authored noBreakHyphen range plus the first following grapheme when both
 * range edges live in this segment. */
function hardJoinPrefixEnd(seg: LayoutTextSeg): number | undefined {
  if (seg.hardJoinPrev !== true || seg.text.length === 0) return undefined;
  const protectedOffsets = protectedNoBreakOffsets(seg);
  const firstLegal = [
    ...graphemeClusterOffsets(seg.text),
    seg.text.length,
  ].find((offset) => offset > 0 && !protectedOffsets.has(offset));
  return firstLegal ?? seg.text.length;
}

/** Single authority for metadata whose UTF-16 coordinates are relative to a
 * text segment. Every retained split path must use this projection. */
function slicedTextMetadata(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): Pick<LayoutTextSeg,
  'punctuationCompressions' | 'noBreakRanges' | 'externalLinkBreakOffsets'
> {
  return {
    punctuationCompressions: slicedPunctuationCompressions(seg, start, end),
    noBreakRanges: slicedNoBreakRanges(seg, start, end),
    externalLinkBreakOffsets: slicedExternalLinkBreakOffsets(seg, start, end),
  };
}

function slicedExternalLinkBreakOffsets(
  seg: LayoutTextSeg,
  start: number,
  end: number,
): readonly number[] | undefined {
  const sliced = seg.externalLinkBreakOffsets
    ?.filter((offset) => offset > start && offset < end)
    .map((offset) => offset - start);
  return sliced && sliced.length > 0 ? Object.freeze(sliced) : undefined;
}

function tightHorizontalGraphemeInk(
  segment: LayoutTextSeg,
  grapheme: string,
): Readonly<{ advancePt: number; xMinPt: number; xMaxPt: number }> | undefined {
  if (!segment.textLayoutService || !segment.textShapeRequest || grapheme.length === 0) {
    return undefined;
  }
  const shaped = segment.textLayoutService.shape({
    ...segment.textShapeRequest,
    text: grapheme,
    measure: true,
    clusterGeometry: false,
  });
  if (
    shaped.horizontalInkBoundsAreTight !== true
    || !shaped.inkBounds
    || !Number.isFinite(shaped.advancePt)
    || !Number.isFinite(shaped.inkBounds.xMinPt)
    || !Number.isFinite(shaped.inkBounds.xMaxPt)
  ) {
    return undefined;
  }
  const scale = segment.charScale ?? 1;
  return {
    advancePt: shaped.advancePt * scale,
    xMinPt: shaped.inkBounds.xMinPt * scale,
    xMaxPt: shaped.inkBounds.xMaxPt * scale,
  };
}

function contextualHorizontalGraphemeAdvances(
  segment: LayoutTextSeg,
): ReadonlyMap<number, number> | undefined {
  if (!segment.textLayoutService || !segment.textShapeRequest || segment.text.length === 0) {
    return undefined;
  }
  const shaped = segment.textLayoutService.shape({
    ...segment.textShapeRequest,
    text: segment.text,
    measure: true,
    clusterGeometry: true,
  });
  if (!shaped.clusters?.length) return undefined;
  const scale = segment.charScale ?? 1;
  const advances = new Map<number, number>();
  for (const cluster of shaped.clusters) {
    if (!Number.isFinite(cluster.advancePt)) return undefined;
    advances.set(cluster.range.end, cluster.advancePt * scale);
  }
  return advances;
}

/**
 * Bound document-level punctuation compression by the adjacent glyphs' tight
 * horizontal ink and contextual advance. Canvas shaping can already kern a
 * punctuation pair down to the retained half-cell; subtracting the isolated
 * glyph's removable sidebearing again would collapse the second mark to zero
 * advance. A following glyph's left ink edge also participates in the collision
 * equation. Resolve this after all source runs have been segmented so a
 * formatting seam cannot reintroduce the overlap.
 */
function retainHorizontalPunctuationInkClearance(segs: LayoutSeg[]): void {
  let pending: Readonly<{
    segment: LayoutTextSeg;
    compressionIndex: number;
    ink: Readonly<{ advancePt: number; xMinPt: number; xMaxPt: number }>;
    contextualAdvancePt: number;
  }> | undefined;
  const adjustedBySegment = new Map<
    LayoutTextSeg,
    Array<{ end: number; adjustmentPt: number }>
  >();
  for (const candidate of segs) {
    if (!('text' in candidate) || candidate.verticalRun) {
      pending = undefined;
      continue;
    }
    const segment = candidate;
    const compressions = segment.punctuationCompressions ?? [];
    const compressionIndexByEnd = new Map(
      compressions.map((compression, index) => [compression.end, index]),
    );
    const contextualAdvances = compressions.length > 0
      ? contextualHorizontalGraphemeAdvances(segment)
      : undefined;
    const boundaries = [0, ...graphemeClusterOffsets(segment.text), segment.text.length];
    for (let index = 0; index < boundaries.length - 1; index += 1) {
      const start = boundaries[index]!;
      const end = boundaries[index + 1]!;
      if (end <= start) continue;
      const compressionIndex = compressionIndexByEnd.get(end);
      const currentInk = pending || compressionIndex !== undefined
        ? tightHorizontalGraphemeInk(segment, segment.text.slice(start, end))
        : undefined;
      if (pending && currentInk) {
        const adjustments = adjustedBySegment.get(pending.segment)
          ?? pending.segment.punctuationCompressions!.map((compression) => ({
            ...compression,
          }));
        const compression = adjustments[pending.compressionIndex]!;
        const retainedExtentPt = Math.max(
          0,
          pending.ink.advancePt + compression.adjustmentPt,
        );
        const retainedExtentAdjustmentPt = Math.min(
          0,
          retainedExtentPt - pending.contextualAdvancePt,
        );
        const collisionSafeAdjustmentPt = Math.min(
          0,
          pending.ink.xMaxPt
            - currentInk.xMinPt
            - pending.contextualAdvancePt,
        );
        const adjustmentPt = Math.max(
          compression.adjustmentPt,
          retainedExtentAdjustmentPt,
          collisionSafeAdjustmentPt,
        );
        if (adjustmentPt !== compression.adjustmentPt) {
          adjustments[pending.compressionIndex] = {
            end: compression.end,
            adjustmentPt,
          };
          adjustedBySegment.set(pending.segment, adjustments);
        }
      }
      pending = compressionIndex !== undefined && currentInk
        ? {
            segment,
            compressionIndex,
            ink: currentInk,
            contextualAdvancePt:
              contextualAdvances?.get(end) ?? currentInk.advancePt,
          }
        : undefined;
    }
  }
  for (const [segment, adjusted] of adjustedBySegment) {
    segment.punctuationCompressions = Object.freeze(
      adjusted.map((compression) => Object.freeze(compression)),
    );
  }
}

/** ECMA-376 §17.3.2.43 `<w:w>` — the horizontal glyph-width scale fraction of a
 *  segment (0.67 = 67%). 1 when the run declares no `w:w`. Multiplies the
 *  natural `measureText` width; the paint pass reproduces it with `ctx.scale`. */
export function charScaleFactor(seg: LayoutTextSeg): number {
  return seg.charScale ?? 1;
}

/** Canonical advance formula for a text string in a run: natural glyph width
 *  scaled by ECMA-376 §17.3.2.43 `<w:w>`, plus the §17.6.5 character-grid
 *  delta, plus one ECMA-376 §17.3.2.35 `<w:spacing>` pitch per code point. */
function textAdvanceWidth(
  naturalWidthPx: number,
  text: string,
  characterGridDeltaPx: number,
  charScale: number,
  charSpacingPx: number,
): number {
  return naturalWidthPx * charScale
    + [...text].length * characterGridDeltaPx
    + [...text].length * charSpacingPx;
}

/** Total per-code-point letter-spacing (px) a segment draws with: the docGrid
 *  cell delta (already scoped by {@link segmentCharacterGridDeltaPx}) PLUS the
 *  §17.3.2.35 character-spacing pitch (all code points). Because Canvas
 *  `ctx.letterSpacing` inserts the SAME advance after every glyph, the two are
 *  additive only when the grid delta applies to every glyph — i.e. a pure-EA
 *  segment (or none, when grid is inactive). For a mixed / Latin segment the
 *  grid delta is 0 (Latin is never snapped, §17.6.5) so only char-spacing
 *  contributes, and the value is still uniform across the segment. This single
 *  value is used for BOTH the measured advance and the painted `ctx.letterSpacing`
 *  so measure==paint holds. */
export function segLetterSpacingPx(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (seg.fitTextPerGapPx !== undefined) return seg.fitTextPerGapPx;
  return segmentCharacterGridDeltaPx(seg, grid, scale) + charSpacingDeltaPx(seg, scale);
}

/** A text segment's laid-out advance including the §17.3.2.43 horizontal scale
 *  and §17.3.2.35 character spacing on top of the docGrid delta. Scale natural
 *  glyph width first, then add the fixed character-spacing pitch per code point;
 *  these are independent OOXML properties. */
export function segAdvanceWidth(
  seg: LayoutTextSeg,
  naturalWidthPx: number,
  grid: DocGridCtx | undefined,
  scale: number,
): number {
  if (seg.fitTextPerGapPx !== undefined) {
    const charCount = [...seg.text].length;
    const gapCount = seg.fitTextRegionEnd ? Math.max(0, charCount - 1) : charCount;
    return naturalWidthPx * charScaleFactor(seg)
      + gapCount * seg.fitTextPerGapPx
      + (seg.fitTextTrailingPadPx ?? 0);
  }
  // ECMA-376 §17.3.2.10 縦中横 (horizontal-in-vertical): the whole run is written
  // horizontally inside ONE cell of the vertical line ("keeping the text on the
  // same line"), so its advance ALONG the column is exactly one em (one cell),
  // independent of the character count and of `w:w` (which stretches the
  // side-by-side glyphs ACROSS the column, not the along-column cell height).
  // PDF-verified on sample-26: the "２９" run occupies exactly one 12 pt cell.
  // (Because the vertical page lays out in a swapped logical frame, this
  // logical-horizontal advance IS the vertical column advance after the page
  // rotation — see vertical-text.ts and renderer's page transform.)
  if (seg.tateChuYoko) return seg.fontSize * scale;
  const segmentDelta = segmentCharacterGridDeltaPx(seg, grid, scale);
  return textAdvanceWidth(
    naturalWidthPx,
    seg.text,
    segmentDelta,
    charScaleFactor(seg),
    charSpacingDeltaPx(seg, scale),
  )
    + widthBalanceSpaceAdjustmentTotalPt(seg, grid) * scale * charScaleFactor(seg)
    + punctuationCompressionTotalPt(seg) * scale;
}

export type SnapToCharsClass = 'eastAsia' | 'latin' | 'complexScript';

/** Normative §17.6.5 script class. Source/run/style boundaries are deliberately
 * absent: ascii and hAnsi form one contiguous Latin block, complex-script text
 * forms its own block, and each East-Asian character owns one cell. */
export function snapToCharsClass(
  seg: LayoutTextSeg,
  grid: DocGridCtx | undefined,
): SnapToCharsClass | null {
  if (grid?.type !== 'snapToChars'
    || !grid.characterPitchPt
    || grid.characterPitchPt <= 0
    || seg.snapToCharacterGrid === false
    || seg.metricOnly
    || seg.fitTextRegionIndex !== undefined
    || seg.tateChuYoko) return null;
  if (seg.script === 'eastAsia') return 'eastAsia';
  if (seg.script === 'complexScript') return 'complexScript';
  return 'latin';
}

export function snapToCharsAllocatedWidthPx(
  naturalWidthPx: number,
  kind: SnapToCharsClass,
  pitchPx: number,
  eastAsianCellCount = 1,
): number {
  if (!(pitchPx > 0) || !Number.isFinite(naturalWidthPx)) return naturalWidthPx;
  if (kind === 'eastAsia') return Math.max(1, eastAsianCellCount) * pitchPx;
  return Math.max(1, Math.ceil(Math.max(0, naturalWidthPx) / pitchPx - 1e-9)) * pitchPx;
}

export function isGridLineRule(ctx: DocGridCtx | undefined): boolean {
  if (!ctx || !ctx.linePitchPt || ctx.linePitchPt <= 0) return false;
  return ctx.type === 'lines'
    || ctx.type === 'linesAndChars'
    || ctx.type === 'snapToChars';
}

/**
 * ECMA-376 §17.6.5 docGrid line grid — number of whole grid CELLS a
 * single-spaced East Asian line occupies on a pitch of `pitchPx`, from the
 * line's SINGLE-LINE HEIGHT `naturalPx` (the resolved font resource's design
 * line height, with the established compatibility registry used only when the
 * selected resource cannot expose metrics).
 * The count is `ceil(naturalPx / pitchPx)` — the smallest number of whole
 * cells that CONTAINS the line.
 *
 * `word-east-asian-grid-line-allocation` records the compatibility formula:
 * ceil(design-line-height / pitch), independent of horizontal or vertical text
 * direction. The focused grid-allocation tests retain the adjudicated boundary
 * matrix; this production comment states only the resulting invariant.
 *
 * A line that fills k pitches exactly occupies k cells (ceil; no measured
 * point sits on the boundary — the geometric reading is that it still FITS).
 * For a mixed-size line, callers supply the tallest run's resolved height
 * (§17.3.1.33 tallest-run line box). ECMA-376 defines `linePitch` as one
 * single-spaced line; taller-line spreading is governed by
 * `word-east-asian-grid-line-allocation`. Returns at least 1 for every finite
 * `naturalPx >= 0`.
 */
export function docGridLineCells(naturalPx: number, pitchPx: number): number {
  return wordEastAsianGridLineCells(naturalPx, pitchPx);
}

/** Deterministic single-line height used to count docGrid cells for one East
 * Asian text run. Resolved font resources contribute their parsed design height.
 *
 * The `word-east-asian-grid-line-allocation` rule supplies the 1.3 × hhea-box
 * fallback measured for the Far East grid path; §17.6.5 does not define this
 * factor.
 *
 * When the font resource is unavailable, its hhea box is unknown, so the
 * compatibility fallback assumes 1.0em.
 * Whole-cell allocation bounds the error. This is an explicit fallback, not a
 * normative font-metrics claim.
 *
 * Never use a substituted Canvas box here: its integer-rounded metrics are
 * font- and scale-dependent. */
export function eastAsianGridCountSinglePx(intendedSinglePx: number, emPx: number): number {
  return wordFarEastSingleLinePx(intendedSinglePx, emPx);
}

/**
 * Compute the total line-box height in px from a line's natural font metrics
 * (fontBoundingBoxAscent + fontBoundingBoxDescent) per ECMA-376 §17.3.1.33.
 *
 *   auto    → natural × value ("single" = 1 natural line, "double" = 2).
 *             When the docGrid line axis is active, the
 *             multiplier applies against the grid pitch instead, with a
 *             floor of the natural line height.
 *   exact   → value in pt, converted to px (ignores font and grid).
 *   atLeast → max(natural, authored minimum, active grid minimum).
 *             `word-grid-at-least-tall-line-unsnapped` owns the explicit
 *             tall-line compatibility branch.
 *   null    → natural, or grid pitch if the section defines one.
 *
 * Exported for unit tests only — not part of the package API (not
 * re-exported from index.ts).
 */
export function lineBoxHeight(
  ls: LineSpacing | null,
  ascentPx: number,
  descentPx: number,
  scale: number,
  grid?: DocGridCtx,
  hasRuby?: boolean,
  intendedSinglePx = 0,
  eastAsian = false,
  // px — the line's DESIGN grid-count height: the max over segments of each
  // run's Word-compatible single-line height (a resolved resource's design
  // height, or the generic East Asian fallback). Used ONLY to count
  // docGrid cells for East Asian lines, so a substituted face's Canvas box
  // cannot change pagination or paint-scale cell allocation.
  gridCountSinglePx?: number,
  // px — unresolved East Asian run em used only by direct/synthetic callers that
  // cannot provide the producer-computed per-line gridCountSinglePx.
  untabledEastAsianEmPx?: number,
): number {
  const glyphNatural = ascentPx + descentPx;
  // For `auto`/single spacing the multiplier applies to the intended font's
  // design line height (ECMA-376 §17.3.1.33). When the document's font is
  // substituted, the Canvas glyph extent (`glyphNatural`) can understate that.
  // A resolved-resource metric or established compatibility profile restores
  // the intended height while
  // never dropping below the substituted glyph extent, so glyphs are not
  // clipped. Grid-snapped lines are governed
  // by the grid pitch instead, so the metric correction stays out of them.
  const natural = Math.max(glyphNatural, intendedSinglePx);
  const hasGrid = isGridLineRule(grid);
  const pitchPx = hasGrid ? grid!.linePitchPt! * scale : 0;
  // Per ECMA-376 §17.6.5, a paragraph whose `line` attribute is NOT
  // explicitly set — it only inherits from docDefault — snaps to one grid
  // pitch per text line in docGrid sections, regardless of the inherited
  // multiplier. Paragraphs that do set `line` on their pPr or a named style
  // multiply against the pitch as usual.
  //
  // A single-spaced line on a docGrid snaps to whole grid CELLS in East Asian
  // text. The number of cells is derived from the line's DESIGN single-line
  // height (`gridCountSinglePx`), per
  // `word-east-asian-grid-line-allocation`; the substituted Canvas glyph box is
  // not used because it can overstate the source resource's design height.
  // A Latin-only line is not cell-rounded: it keeps its natural height above a
  // one-cell floor. ECMA-376 Part 1 defines only the natural ≤ pitch case
  // (§17.6.5 / §17.3.1.32), so `word-east-asian-grid-line-allocation` gates
  // whole-cell allocation on the line's script.
  const gridSingleCell = (): number => {
    if (!eastAsian) return Math.max(glyphNatural, pitchPx);
    // Ruby lines reserve real furigana height (base + rt); honor the measured
    // glyph box so the annotation is not clipped. Plain EA lines snap their
    // design single-line height to whole cells.
    if (hasRuby) return Math.max(pitchPx, Math.ceil(glyphNatural / pitchPx) * pitchPx);
    // `word-east-asian-grid-line-allocation`: count cells from the source face's
    // design single-line height, not a substituted Canvas glyph box. Prefer the
    // per-line design-grid height; direct unresolved callers may supply the run em.
    // A legacy caller with neither input gets one pitch.
    const cellCountHeight = gridCountSinglePx
      ?? (intendedSinglePx > 0
        ? intendedSinglePx
        : untabledEastAsianEmPx === undefined
          ? pitchPx
          : eastAsianGridCountSinglePx(0, untabledEastAsianEmPx));
    return docGridLineCells(cellCountHeight, pitchPx) * pitchPx;
  };
  const inheritedOnly = ls !== null && ls.explicit !== true;
  if (!ls) {
    // No explicit spacing → single line. Use the intended single-line height
    // (`natural`) off-grid; on-grid, snap per gridSingleCell.
    return hasGrid ? gridSingleCell() : natural;
  }
  // A zero/negative `w:line` is degenerate input whose behavior ECMA-376
  // §17.3.1.33 does not define (read literally, an `exact` line of 0 would
  // collapse the line box to no height; some generators emit
  // `<w:spacing w:line="0" w:lineRule="exact"/>` on table cells, e.g. sample-7).
  // `word-degenerate-line-spacing-single` follows the native LSPD
  // representation:
  // "exact" spacing is encoded as a negative dyaLine ("the line spacing, in
  // twips, is exactly 0x10000 minus dyaLine", so an exact 0 is unrepresentable)
  // and a non-negative dyaLine in twips mode is "dyaLine or the number of twips
  // necessary for single spacing, whichever value is greater" — i.e. a stored 0
  // resolves to exactly single spacing. `word-degenerate-line-spacing-single`
  // applies that non-collapsing interpretation to exact/auto values <= 0.
  if (wordDegenerateLineSpacingIsSingle(ls.rule, ls.value)) {
    return hasGrid ? gridSingleCell() : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      if (inheritedOnly) {
        const allocated = gridSingleCell();
        return eastAsian
          ? wordUseFeLayoutInheritedGridHeightPx(allocated, pitchPx, ls.value)
          : allocated;
      }
      return Math.max(glyphNatural, pitchPx * ls.value);
    }
    return natural * ls.value;
  }
  if (ls.rule === 'exact') return ls.value * scale;
  if (ls.rule === 'atLeast') {
    // §17.18.48 establishes the authored minimum and §17.6.5 establishes the
    // grid pitch, but neither clause specifies how a tall, plain line with an
    // explicit atLeast value combines with whole-cell grid allocation.
    // `word-grid-at-least-tall-line-unsnapped` preserves the raw content height;
    // ruby and inherited-only spacing retain their established whole-cell path.
    const gridMinimum = hasGrid
      ? (hasRuby || inheritedOnly ? gridSingleCell() : pitchPx)
      : 0;
    return wordGridAtLeastLineHeightPx(natural, ls.value * scale, gridMinimum);
  }
  return natural;
}

/** Natural single-line height in px for an empty paragraph (no rendered text). */
export function emptyLineNaturalPx(fontSizePt: number, scale: number): { asc: number; desc: number } {
  return { asc: fontSizePt * scale * 0.8, desc: fontSizePt * scale * 0.2 };
}

/** Single-line ascent/descent from the selected face, with the compatibility
 * registry correcting known unavailable-font substitutions. */
export function correctedLineMetrics(
  m: TextMetrics,
  family: string | null | undefined,
  fallbackEmPx: number,
  correctionEmPx: number,
  eastAsian = false,
): { ascent: number; descent: number } {
  const rawAsc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fallbackEmPx * 0.8;
  const rawDesc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fallbackEmPx * 0.2;
  return correctLineMetrics(family, correctionEmPx, rawAsc, rawDesc, eastAsian);
}

/**
 * Height (px) of the paragraph-mark line box for a paragraph that places no
 * inline content on any line. Per ECMA-376 §17.3.1.29 the paragraph mark always
 * produces one line box even when the paragraph has no inline runs; floating
 * objects (§20.4.2.x `wp:anchor`) are removed from the inline flow but never
 * suppress that paragraph-mark line. This is the height used both by the
 * literal empty-paragraph path and by paragraphs whose only segments are
 * wrap-float anchors (which `layoutLines` skips, yielding zero lines).
 * `effectiveLineSpacing` lets resolved paragraph context override the source
 * value; omitting it preserves the existing `para.lineSpacing` behavior.
 */
/** The natural ascent/descent (px) and the resolved line-box advance (px) of an
 *  empty paragraph's mark line. Shared by {@link paragraphMarkLineHeight} (which
 *  returns only the advance) and {@link paragraphMarkBelowBaselinePt} (which needs
 *  the ascent/descent to locate the mark baseline within the box). */
export interface MarkLineMetrics {
  readonly advancePx: number;
  readonly ascentPx: number;
  readonly descentPx: number;
}

export function paragraphMarkLineMetrics(
  para: ParagraphLayoutSource,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
): MarkLineMetrics {
  const effectiveMarkShapeInput = markShapeInput;
  // ECMA-376 §17.3.2.26 `w:rFonts@w:hint`: an empty paragraph has no code
  // point from which to infer a script slot, so the paragraph-mark hint selects
  // the face used to measure the mark. It does NOT make an otherwise Latin-only
  // paragraph occupy East-Asian docGrid cells: grid-cell classification remains
  // content/document based, independently of font routing.
  const markUsesEastAsianFace = eastAsian || effectiveMarkShapeInput?.fontHint === 'eastAsia';
  const forceCs = effectiveMarkShapeInput?.complexScript === true;
  const fs = effectiveMarkShapeInput?.fontSizePt ?? getDefaultFontSize(para);
  const authoredFamily = getDefaultFontFamily(para, markUsesEastAsianFace);
  const markWeight = effectiveMarkShapeInput?.weight ?? 400;
  const markStyle = effectiveMarkShapeInput?.style ?? 'normal';
  const normalizedFamily = authoredFamily
    ? normalizeFontMetricFamily(authoredFamily)
    : null;
  const resolvedLocalFont = normalizedFamily
    ? resolvedLocalFonts[`${normalizedFamily}:${markWeight}:${markStyle}`]
      ?? (markWeight === 400 && markStyle === 'normal'
        ? resolvedLocalFonts[normalizedFamily]
        : undefined)
    : undefined;
  const measuredFamily = resolvedLocalFont?.family ?? authoredFamily;
  let asc: number;
  let desc: number;
  if (textLayoutService) {
    const bold = markWeight >= 600;
    const italic = markStyle === 'italic';
    const ascii = effectiveMarkShapeInput?.fonts.ascii ?? para.defaultFontFamily ?? authoredFamily;
    const shaped = textLayoutService.shape({
      text: markUsesEastAsianFace ? 'あ' : 'x',
      fontSizePt: fs * scale,
      fonts: effectiveMarkShapeInput?.fonts ?? {
        ascii,
        highAnsi: ascii,
        eastAsia: para.defaultFontFamilyEastAsia ?? ascii,
        complexScript: ascii,
      },
      themeFonts: effectiveMarkShapeInput?.themeFonts,
      themeFontPresence: effectiveMarkShapeInput?.themeFontPresence,
      weight: bold ? 700 : 400,
      style: italic ? 'italic' : 'normal',
      complexScript: forceCs,
      fontHint: effectiveMarkShapeInput?.fontHint,
      eastAsiaLanguage: effectiveMarkShapeInput?.eastAsiaLanguage,
      kerning: effectiveMarkShapeInput?.kerning,
      measure: true,
    });
    const face = shaped.spans[0]?.font.resolvedFamily ?? authoredFamily;
    ({ ascent: asc, descent: desc } = correctedLineMetrics(
      {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics,
      face,
      fs * scale,
      fs * scale,
      markUsesEastAsianFace,
    ));
  } else if (ctx) {
    // ECMA-376 §17.3.1.29 / §17.3.1.33: an empty paragraph's mark line reserves
    // the mark font's REAL single-line height — the SAME fontBoundingBox a text
    // line of that font and size uses (layoutLines), so an empty paragraph is
    // exactly as tall as a one-character paragraph of the same run properties.
    // The synthetic 0.8/0.2 ≈ 1em box under-measured every empty paragraph
    // whenever the (often substituted) font's real box exceeds 1em — a Latin
    // fallback reports ~1.15em — so a run of empty "spacer" paragraphs fell
    // short and following content rose into a preceding float's wrap band
    // instead of clearing the float. East Asian documents probe an EA glyph so
    // docGrid cell rounding (lineBoxHeight) reserves whole cells (a 20pt mark
    // on a 20pt pitch occupies two cells); others probe a Latin glyph.
    // fontBoundingBox is reported per
    // resolved face (not per glyph), so the probe choice does not change the box
    // for a face that contains it — and the probe is script-matched, so the mark
    // font does. A parsed resource metric, when available, is applied below by
    // the same path used for visible text.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, measuredFamily, fontFamilyClasses);
    const m = ctx.measureText(markUsesEastAsianFace ? 'あ' : 'x');
    ctx.font = prevFont;
    // A mark line carries no smallCaps/vertAlign, so fallback == correction size.
    ({ ascent: asc, descent: desc } = correctedLineMetrics(
      m, measuredFamily, fs * scale, fs * scale, markUsesEastAsianFace,
    ));
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  const resourceRatio = markUsesEastAsianFace
    ? resolvedLocalFont?.eastAsianLineHeightRatio ?? resolvedLocalFont?.lineHeightRatio
    : resolvedLocalFont?.lineHeightRatio;
  const intendedSingle = Math.max(
    (resourceRatio ?? 0) * fs * scale,
    emptyIntendedSingleForScriptPx(para, scale, markUsesEastAsianFace),
    wordMsMinchoEmptyEastAsianMarkSingleLinePx(
      authoredFamily,
      fs * scale,
      markUsesEastAsianFace,
    ),
  );
  const gridCountSingle = eastAsian
    ? eastAsianGridCountSinglePx(intendedSingle, fs * scale)
    : undefined;
  const ordinaryAdvancePx = lineBoxHeight(
    effectiveLineSpacing,
    asc,
    desc,
    scale,
    grid,
    paraHasRuby,
    intendedSingle,
    eastAsian,
    gridCountSingle,
  );
  const gridAllocationActive = useFeLayout && eastAsian && isGridLineRule(grid);
  const allocatedGridAdvancePx = gridAllocationActive
    ? lineBoxHeight(
        null,
        asc,
        desc,
        scale,
        grid,
        paraHasRuby,
        intendedSingle,
        eastAsian,
        gridCountSingle,
      )
    : ordinaryAdvancePx;
  // Candidate for the observed atLeast=0 compatibility branch. Compute it
  // independently of the source's inheritance flag so the compatibility owner
  // can select it without leaking the signed boundary into this producer.
  const atLeastZeroAdvancePx = gridAllocationActive
    ? lineBoxHeight(
        { rule: 'atLeast', value: 0, explicit: true },
        asc,
        desc,
        scale,
        grid,
        paraHasRuby,
        intendedSingle,
        eastAsian,
        gridCountSingle,
      )
    : ordinaryAdvancePx;
  const advancePx = useFeLayout
    ? wordUseFeLayoutParagraphMarkGridAdvancePx({
        ordinaryAdvancePx,
        allocatedGridAdvancePx,
        atLeastZeroAdvancePx,
        lineSpacing: effectiveLineSpacing,
        gridAllocationActive,
        scale,
      })
    : ordinaryAdvancePx;
  return { advancePx, ascentPx: asc, descentPx: desc };
}

export function paragraphMarkLineHeight(
  para: ParagraphLayoutSource,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
): number {
  return paragraphMarkLineMetrics(
    para, scale, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts, textLayoutService, markShapeInput, useFeLayout,
  ).advancePx;
}

/**
 * §17.3.1.29 / §17.3.1.33 — the extent (px) of a line that sits BELOW its
 * baseline (descent + half of any auto/atLeast leading), using the HALF-LEADING
 * (centred) baseline `top + (advance − (ascent + descent)) / 2 + ascent`, so the
 * portion below it is `(advance − ascent + descent) / 2`.
 *
 * Called for BOTH a paragraph's last visible line (paragraph-measure.ts, the
 * `lastLineBelowBaselinePt` field) and — via {@link paragraphMarkBelowBaselinePt}
 * — an empty paragraph's mark line. Its ONE consumer (renderer.ts
 * `trailingMarkOverflow`) reads it only for an inkless trailing MARK: the
 * whitespace such a paragraph may let overflow the bottom content edge under
 * `word-trailing-empty-mark-baseline-admission` (#981).
 *
 * NOTE: this stays the CENTRED baseline even though VISIBLE lineRule=auto content
 * lines now use a PINNED baseline (#990: multiplier leading placed entirely
 * below the glyphs). The pagination consumer is inkless
 * (mark-only), so the pinned glyph baseline never reaches it, and the #981 page
 * fit is pinned by that admission rule — changing it would move page boundaries.
 * `word-auto-multiple-baseline-pin` is therefore intentionally DRAW-ONLY.
 */
export function lineBelowBaselinePx(advancePx: number, ascentPx: number, descentPx: number): number {
  return Math.max(0, (advancePx - ascentPx + descentPx) / 2);
}

export function paragraphMarkBelowBaselinePt(
  para: ParagraphLayoutSource,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian: boolean,
  ctx: MeasurementTextContext | undefined,
  fontFamilyClasses: Record<string, string>,
  effectiveLineSpacing: LineSpacing | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  textLayoutService?: TextLayoutService,
  markShapeInput?: NumberingMarkerShapeInput,
  useFeLayout = false,
): number {
  // Measured at scale 1 so the returned px value is already in points.
  const m = paragraphMarkLineMetrics(
    para, 1, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts, textLayoutService, markShapeInput, useFeLayout,
  );
  return lineBelowBaselinePx(m.advancePx, m.ascentPx, m.descentPx);
}

/**
 * Resolve the formatting axis that actually governs a run's glyphs.
 *
 * ECMA-376 §17.3.2.30 `w:rtl` marks a run as complex-script. For such a run the
 * complex-script properties take effect — §17.3.2.4 `bCs` (bold), §17.3.2.6
 * `iCs` (italic), §17.3.2.26 `rFonts@cs` (typeface), §17.3.2.39 `szCs` (size) —
 * instead of the non-CS `b`/`i`/`rFonts@ascii`/`sz`, which apply to
 * non-complex (Latin/CJK) text. `bCs`/`iCs` are INDEPENDENT toggles: an absent
 * `bCs`/`iCs` does not inherit `b`/`i`'s value, so a complex-script run that
 * carries only `w:b`/`w:i` renders non-bold/upright (`csBold = boldCs ?? false`,
 * `csItalic = italicCs ?? false`). Thus `w:b` without `w:bCs` remains regular
 * weight, and `w:i` without `w:iCs` remains upright, while the corresponding
 * non-complex text uses the Latin-axis toggle.
 */
/**
 * Split a `w:smallCaps` (§17.3.2.33) run into maximal pieces by character class
 * for sizing. The spec reduces "all SMALL LETTER characters ... two points
 * smaller", so ONLY lowercase letters are `reduced`; uppercase letters AND every
 * non-alphabetic character (digits, punctuation) stay at the FULL run size.
 * So "Introduction" → "I" full + "NTRODUCTION" reduced (matching the heading's
 * "1."), and "co2" → "CO" reduced + "2" full. `reduced` flags the small-cap
 * pieces; the caller still uppercases every piece for display.
 *
 * Whitespace carries no glyph, so it EXTENDS the current piece rather than
 * opening a full-size one — otherwise an inter-word space between two small-cap
 * words would fragment into its own segment and corrupt trailing-space collapse
 * / line breaking. A leading run with no lowercase letter defaults to full size.
 */
export function splitSmallCapsCase(text: string): { text: string; reduced: boolean }[] {
  const out: { text: string; reduced: boolean }[] = [];
  for (const ch of text) {
    // A lowercase letter: unchanged by toLowerCase AND changed by toUpperCase.
    const isLowerLetter = ch.toLowerCase() === ch && ch.toUpperCase() !== ch;
    const reduced = /\s/.test(ch)
      ? (out[out.length - 1]?.reduced ?? false) // whitespace: keep with current piece
      : isLowerLetter;
    const last = out[out.length - 1];
    if (last && last.reduced === reduced) last.text += ch;
    else out.push({ text: ch, reduced });
  }
  return out.length ? out : [{ text, reduced: false }];
}

export function findNearbyFontSize(
  runs: readonly ParagraphLayoutRun[],
  idx: number,
): number {
  // Look backwards then forwards for a text or field run to get font size
  for (let i = idx - 1; i >= 0; i--) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  for (let i = idx + 1; i < runs.length; i++) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  return 10; // pt fallback
}

export function resolveFieldText(f: FieldRun, environment: LineLayoutEnvironment): string {
  if (f.fieldType === 'page') {
    // ECMA-376 §17.16.5.44 PAGE — "the number of the current page". Use the
    // per-section DISPLAY number (§17.6.12 `w:start` restart), falling back to the
    // raw physical index for a single-section document without `<w:pgNumType>`.
    const n = environment.displayPageNumber ?? environment.pageIndex + 1;
    // §17.16.4.3.1 — the field's own general-formatting switch (`\* roman`, …)
    // OVERRIDES the section format (§17.6.12 `w:fmt`); it is authored ON the field.
    // No switch ⇒ the section format (or decimal for a single-section document).
    const fmt = parseFieldFormatSwitch(f.instruction) ?? environment.pageNumberFormat ?? 'decimal';
    return formatOrdinalNumber(n, fmt);
  }
  // ECMA-376 §17.16.5.42 NUMPAGES — "the number of pages in the current document".
  // This is the DOCUMENT's physical page count and is NOT affected by §17.6.12
  // page-number restart (which only shifts the DISPLAYED number). It IS still
  // subject to the field's own `\*` format switch.
  if (f.fieldType === 'numPages') {
    const fmt = parseFieldFormatSwitch(f.instruction) ?? 'decimal';
    return formatOrdinalNumber(environment.totalPages, fmt);
  }
  // ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — display the CURRENT date/time
  // filtered through the field's `\@` date-time picture (§17.16.4.1). The
  // "current" instant is injected via `environment.currentDateMs` (default = real time,
  // set at the render entry point) so the output is deterministic under test.
  // A field with NO `\@` picture, or one whose picture uses an unimplemented
  // token, falls back to the authored cached result (§17.16.4.1: with no picture
  // the result is formatted "in an implementation-defined manner" — we keep
  // Word's cached rendering rather than invent one).
  if (f.fieldType === 'date' || f.fieldType === 'time') {
    const picture = parseDateTimePictureSwitch(f.instruction);
    if (picture) {
      const now = new Date(environment.currentDateMs ?? Date.now());
      const formatted = formatDateTimePicture(picture, now);
      if (formatted !== null) return formatted;
    }
    return f.fallbackText;
  }
  return f.fallbackText;
}

export const mathPlainText = mathFallbackText;

/** Returns true when any code point of `text` permits a line break between
 *  adjacent characters (CJK / ideographic). The canonical ranges live in core's
 *  {@link isCjkBreakChar} (single source of truth across all renderers). */
export function hasCJKBreakOpportunity(text: string): boolean {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (isCjkBreakChar(cp)) return true;
    i += cp > 0xffff ? 2 : 1;
  }
  return false;
}

// ECMA-376 §17.15.1.18 / §17.18.7 distinguishes punctuation-only compression
// from punctuation-plus-Japanese-kana compression. This is the reviewed
// supported subset: full-width dividing punctuation and closing forms verified
// by the registered Word fixture. U+3017, full-width !, and full-width ? remain
// full-cell in that same matrix and are deliberately excluded.
// JLReq classifies middle dot, colon, and semicolon together (cl-05); their
// whitespace belongs on both sides and must be resolved from the adjacent
// character classes, so they cannot use this trailing-side-only projection.
// Halfwidth U+FF61/U+FF64 are not full-width. Opening punctuation likewise
// needs line-start positioning rather than a pen-advance reduction. The
// implementation-note evidence for the full-width punctuation scope is
// registered in layout/line-compatibility.ts.
const COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION = new Set([
  '、', '。', '，', '．', '」', '』', '】', '）', '］', '｝',
]);

/** Full-width Japanese kana characters for
 * `compressPunctuationAndJapaneseKana`. The ranges follow Unicode's Hiragana,
 * Katakana, Katakana Phonetic Extensions, and supplementary Kana blocks while
 * excluding halfwidth Katakana and punctuation such as U+30FB. U+30FC is the
 * shared full-width kana prolonged-sound mark. */
function isFullWidthJapaneseKana(character: string): boolean {
  const cp = character.codePointAt(0);
  if (cp === undefined) return false;
  return (
    (cp >= 0x3041 && cp <= 0x3096)
    || (cp >= 0x309d && cp <= 0x309f)
    || (cp >= 0x30a1 && cp <= 0x30fa)
    || cp === 0x30fc
    || (cp >= 0x30fd && cp <= 0x30ff)
    || (cp >= 0x31f0 && cp <= 0x31ff)
    || (cp >= 0x1aff0 && cp <= 0x1afff)
    || (cp >= 0x1b000 && cp <= 0x1b16f)
  );
}

function characterSpacingControlCompresses(
  grapheme: string,
  setting: string | undefined,
): boolean {
  switch (setting) {
    case 'compressPunctuation':
      return COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(grapheme);
    case 'compressPunctuationAndJapaneseKana':
      return COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(grapheme)
        || isFullWidthJapaneseKana(grapheme);
    case 'doNotCompress':
    default:
      return false;
  }
}

/** Shift a SEA break-offset list (issue #797) onto a suffix that drops the first
 *  `cut` UTF-16 units: keep offsets strictly greater than `cut` and rebase them.
 *  Used when a Thai/Lao/Khmer segment is split (line wrap) or resumed at a
 *  pagination boundary. A non-SEA segment (`offsets === undefined`) stays
 *  non-SEA; a SEA segment stays SEA-flagged (returns `[]` when no dictionary
 *  boundary remains, so an over-long FINAL word still takes the SEA path and is
 *  split grapheme-safely rather than by code point). */
function rebaseSeaBreaks(offsets: readonly number[] | undefined, cut: number): readonly number[] | undefined {
  if (offsets === undefined) return undefined;
  const out: number[] = [];
  for (const o of offsets) if (o > cut) out.push(o - cut);
  return out;
}

/**
 * Binary-search the longest prefix of `text` whose rendered width fits in `maxWidth`.
 * Used for CJK overflow splitting.
 */
/** Extend an accepted split point through IMMEDIATELY FOLLOWING IDEOGRAPHIC
 *  SPACES (U+3000): the fullwidth space belongs to the line it ends, hanging
 *  past the band (JLReq line-end ideographic-space handling — the same
 *  allowance fitCJKPrefix's fit predicate applies), so a split must never
 *  strand it at the head of the next line — including the FORCE-FIT paths
 *  where the band is narrower than a single glyph (a one-glyph-wide form
 *  label column). A zero split (whole-run move / kinsoku retraction) is left
 *  untouched. */
function extendThroughTrailingIdeographicSpaces(
  chars: string[],
  split: number,
  maximum = Number.POSITIVE_INFINITY,
): number {
  if (split <= 0 || maximum <= 0) return split;
  if (Number.isFinite(maximum) && chars[split - 1] === '\u3000') return split;
  let s = split;
  let remaining = maximum;
  while (s < chars.length && chars[s] === '\u3000' && remaining > 0) {
    s++;
    remaining--;
  }
  return s;
}

/** Project the registered line-end allowance from the immediately preceding
 * visible East-Asian character, not from another character elsewhere in a
 * mixed-script segment. */
function hasEastAsianVisiblePredecessor(text: string): boolean {
  const characters = [...text];
  for (let index = characters.length - 1; index >= 0; index -= 1) {
    if (characters[index] === '\u3000') continue;
    return EAST_ASIAN_RE.test(characters[index]);
  }
  return false;
}

export function fitCJKPrefix(
  ctx: MeasurementTextContext,
  text: string,
  maxWidth: number,
  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph, 0 when inactive).
  // The fit must compare the same advance model as the line box / draw so the
  // grid's char count and run character metrics land on the same split.
  gridDeltaPx = 0,
  // WD4 — the run's §17.3.2.43 horizontal glyph scale (1 = 100%) and §17.3.2.35
  // per-code-point character-spacing pitch in px. Threaded so a CJK run that is
  // scaled/spaced splits at the SAME cell boundary the whole-segment advance
  // model uses (measure==paint). Default (1, 0) reproduces the prior behaviour.
  charScale = 1,
  charSpacingPx = 0,
  // issue #1014 — a vertical (tbRl) run whose segment is flagged `verticalRun`:
  // fold the vo=Tr rotate-fallback ink deficit into the fit predicate too, so the
  // wrap chooses a prefix whose CORRECTED advance (the same the line box measures)
  // fits — not one that only fits by the under-reported raw width. 0 for horizontal
  // / non-under-reporting runs, so the split is byte-identical there.
  verticalRun = false,
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  /** Optional caller-owned advance authority. Production layout supplies the
   * same substring measurement used by whole-segment fit so every prefix uses
   * the canonical selected-face and OOXML pitch model. */
  measureAdvance?: (text: string) => number,
  /** Maximum trailing U+3000 characters excluded from the fit width. */
  maximumIdeographicSpaceHang = Number.POSITIVE_INFINITY,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  const advanceOf = (prefix: string): number => measureAdvance?.(prefix) ?? (() => {
    let verticalExtraPx = 0;
    if (verticalRun) {
      if (!verticalGlyphMeasurement) {
        throw new Error('Vertical glyph measurement capability is required for vertical text');
      }
      verticalExtraPx = verticalGlyphMeasurement.measureRunInkExtra(prefix);
    }
    return textAdvanceWidth(
      ctx.measureText(prefix).width + verticalExtraPx,
      prefix,
      gridDeltaPx,
      charScale,
      charSpacingPx,
    );
  })();
  // Trailing IDEOGRAPHIC SPACE (U+3000) line-end allowance: a candidate that
  // overflows ONLY because it ends in fullwidth spaces still fits — those spaces
  // hang past the line end (JLReq line-end ideographic-space handling; Word
  // does the same, which is what keeps a "char + U+3000" form label at one
  // visible glyph per line instead of alternating glyph/space lines). The
  // accepted range KEEPS the trailing spaces, so the next line starts at the
  // following visible character. Scope: trailing U+3000 in the candidate only —
  // leading/interior fullwidth spaces stay width-bearing (authored indents),
  // and ASCII-space handling is a separate, untouched mechanism. The predicate
  // stays monotone in the candidate length (appending a U+3000 never changes
  // the visible advance; appending a visible char only grows it), so the
  // binary search remains valid.
  const fitsWithHang = (endExclusive: number): boolean => {
    let visibleEnd = endExclusive;
    let remainingHang = maximumIdeographicSpaceHang;
    if (remainingHang > 0) {
      while (visibleEnd > 0 && chars[visibleEnd - 1] === '\u3000') visibleEnd--;
      const trailingCount = endExclusive - visibleEnd;
      visibleEnd += Math.max(0, trailingCount - remainingHang);
    }
    const prefix = chars.slice(0, visibleEnd).join('');
    return advanceOf(prefix) <= maxWidth;
  };
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (fitsWithHang(mid)) lo = mid;
    else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
}

/**
 * Split a text run into layout-segment strings.
 * Each segment is an atomic unit for word-level fitting; CJK overflow is handled in layoutLines.
 */
/** RTL primary language subtags (ISO 639) whose complex-script context makes
 *  Word classify European digits as Arabic-Number (AN). */
export const RTL_PRIMARY_SUBTAGS = new Set([
  'ar', // Arabic
  'fa', // Persian
  'ur', // Urdu
  'he', 'iw', // Hebrew (iw = legacy code)
  'yi', 'ji', // Yiddish
  'ps', // Pashto
  'sd', // Sindhi
  'ug', // Uyghur
  'dv', // Divehi
  'syr', // Syriac
  'ckb', // Central Kurdish (Sorani)
]);

/**
 * Decide whether a `w:lang w:bidi` tag (§17.3.2.20) designates an RTL
 * complex-script language, so the run's European digits are classified AN
 * (Word's date ordering). The tag's primary subtag (before the first '-') is
 * matched against {@link RTL_PRIMARY_SUBTAGS}. When the tag is absent OR a
 * malformed/unknown value (e.g. the "ae-AR" seen in real-world files), fall
 * back to whether the run is explicitly rtl-marked — `w:rtl` already asserts
 * the run is complex-script RTL content.
 */
export function isRtlBidiLang(langBidi: string | undefined, runIsRtl: boolean): boolean {
  if (langBidi) {
    const primary = langBidi.split('-')[0].toLowerCase();
    if (RTL_PRIMARY_SUBTAGS.has(primary)) return true;
  }
  return runIsRtl;
}

/**
 * Split `text` into maximal runs that are uniformly complex-script or not, per
 * §17.3.2.26 per-character classification. Returns `[{text, cs}]` in logical
 * order. Used only when a run has NO explicit `w:rtl`/`w:cs` (which would force
 * the whole run to cs); otherwise the caller treats the entire piece as cs.
 *
 * Under `word-neutral-script-attachment`, digits / spaces / punctuation attach
 * to the PRECEDING slice so a number embedded in Arabic ("نص 12 نص") does not
 * fragment into extra segments. A leading neutral run takes the first strong
 * slice's class.
 */
export function splitByComplexScript(text: string): { text: string; cs: boolean }[] {
  const out: { text: string; cs: boolean }[] = [];
  let curCs: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    // Neutral (non-letter) characters do not switch the active class; they ride
    // with whatever script is currently open (or the next one if none yet).
    if (wordNeutralAttachesToActiveScript(ch)) {
      buf += ch;
      continue;
    }
    const cs = isComplexScriptCodePoint(cp);
    if (curCs === null) {
      curCs = cs;
      buf += ch;
    } else if (cs === curCs) {
      buf += ch;
    } else {
      out.push({ text: buf, cs: curCs });
      curCs = cs;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, cs: curCs ?? false });
  return out;
}

/**
 * Split a (non-complex-script) string into maximal runs that are uniformly
 * East-Asian (CJK) or not, per the §17.3.2.26 ascii/eastAsia axis split. Returns
 * `[{text, ea}]` in logical order. CJK classification uses the canonical
 * {@link isCjkBreakChar} from `@silurus/ooxml-core` — the SAME predicate the body
 * wrap/justify paths use. Text-box text now feeds this splitter too (its runs are
 * adapted to body runs and run through {@link buildSegments}), so the eastAsia
 * face is picked consistently across body and shape with no name heuristics. Each
 * returned slice stays single-font when emitted, preserving the
 * measure==draw / docGrid char-grid invariant.
 *
 * Boundary rule: classification is purely per code point (every CJK code point
 * opens/continues an `ea` run; every other code point a `latin` run). This is
 * intentionally simpler than {@link splitByComplexScript}'s neutral-attachment —
 * a digit between two ideographs is Latin/ascii either way (§17.3.2.26 assigns
 * ASCII digits to the ascii face), and a single fillText anchors to the cumulative
 * whole-string advance, so the visible spacing is unchanged.
 *
 * NOTE: this split decides the FONT slot only. `linesAndChars` applies its pitch
 * to both partitions. The pre-existing non-`linesAndChars` fallback uses the
 * grid's own `EAST_ASIAN_RE` purity test (see `gridSegDeltaPx`/`eaGlyphCount`),
 * not the `ea` font-slot flag here.
 */
export function splitByEastAsia(text: string): { text: string; ea: boolean }[] {
  const out: { text: string; ea: boolean }[] = [];
  let curEa: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    const ea = isCjkBreakChar(cp);
    if (curEa === null || ea === curEa) {
      curEa = ea;
      buf += ch;
    } else {
      out.push({ text: buf, ea: curEa });
      curEa = ea;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, ea: curEa ?? false });
  return out;
}

/**
 * Split a token into maximal runs of European digits (U+0030–0039) versus the
 * separators between them, so a date in an AN-classified Arabic run can be
 * reordered group-by-group by the per-line bidi pass (which works at segment
 * granularity). "28-02-2026" → ["28","-","02","-","2026"], which the RTL reorder
 * then enters the right-to-left layout pass.
 *
 * EXCEPTION — ECMA-376 relies on UAX#9 W4: a SINGLE common separator (CS) sitting
 * between two numbers of the same type joins them into ONE number. So a decimal /
 * thousands / time separator (`.`, `,`, `:`, `/`, NBSP) flanked by European
 * digits on BOTH sides stays inside the digit group: "1234.56", "1,234.56" and
 * "12:34" are one left-to-right number, not three reorderable pieces. (A European
 * separator like `-` is ES, NOT CS, and W4's ES clause is EN-only — these run
 * digits are AN — so a hyphen still splits, preserving the date case.) Splitting
 * a decimal sent "1234.56" through the RTL segment reorder and drew it "56.1234".
 */
export function splitDigitGroups(text: string): string[] {
  const isEuDigit = (c: number) => c >= 0x30 && c <= 0x39;
  // UAX#9 Common Separator (CS) subset that can join two adjacent numbers (W4).
  // The last char is NBSP (U+00A0, e.g. a French thousands separator), itself CS;
  // a plain space is WS and never reaches here (splitTextForLayout breaks on it).
  const isJoiningCS = (ch: string) =>
    ch === '.' || ch === ',' || ch === ':' || ch === '/' || ch === ' ';
  const out: string[] = [];
  let buf = '';
  let bufDigit: boolean | null = null;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    let isDigit = isEuDigit(ch.charCodeAt(0));
    // W4: a single CS between two European digits is part of the number — keep it
    // in the current digit group so the whole number stays one (LTR) segment.
    if (!isDigit && bufDigit === true && isJoiningCS(ch) && isEuDigit(text.charCodeAt(i + 1))) {
      isDigit = true;
    }
    if (bufDigit === null || isDigit === bufDigit) {
      buf += ch;
    } else {
      out.push(buf);
      buf = ch;
    }
    bufDigit = isDigit;
  }
  if (buf.length > 0) out.push(buf);
  return out.length ? out : [text];
}

export function splitTextForLayout(text: string): string[] {
  const result: string[] = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== ' ') j++;
    while (j < text.length && text[j] === ' ') j++;
    if (j > i) result.push(text.slice(i, j));
    i = j;
  }
  return result.length ? result : [text];
}

/** ECMA-376 §17.15.1.25 — the ABSENT default for `<w:defaultTabStop>`: "If this
 *  element is omitted, then automatic tab stops should be generated at 720
 *  twentieths of a point (0.5")", i.e. 36 pt. Used ONLY as the fallback when a
 *  document carries no `<w:defaultTabStop>`; a document that sets one overrides
 *  this via {@link resolveDefaultTabPt}. Shared by line layout and the
 *  numbered-list marker's retained trailing-tab advance. */
export const DEFAULT_TAB_PT = 36;

/** Knuth-Plass shrink tolerance: the fraction by which the line breaker may
 *  compress each inter-word space to keep a candidate word on the current line.
 *  ECMA-376 prescribes no line-breaking algorithm — tolerance-based fit is
 *  standard typography (TeX, InDesign, Word) and lets the layout absorb the
 *  canvas `measureText` vs Word advance-width discrepancy (~0.1–0.3 px/glyph)
 *  that would otherwise push a trailing word to the next line. Per ECMA-376
 *  §17.18.44, this tolerance is suppressed per line when the draw pass will
 *  fully justify it: non-final/non-manual-break lines of `both`/kashida, and
 *  every line of `distribute`/`thaiDistribute`. Lines the paint pass leaves
 *  non-justified keep the budget so measurement and paint agree (issue #698).
 *
 *  For eligible non-justified lines this is the ONE budget shared by both sides
 *  of the fit contract: the wrap judgment below admits a word when the line's
 *  overflow Δ ≤ SPACE_SHRINK_RATIO · Σ(trailing-space widths), and the renderer's
 *  draw pass squeezes the same spaces by the same fraction so the admitted line
 *  lands inside its box instead of overrunning the clip (see
 *  `shrinkFitCompression` in text-distribute.ts). */
export const SPACE_SHRINK_RATIO = 0.25;

/** ECMA-376 §17.15.1.25 — resolve the document's automatic tab-stop interval
 *  (pt): the explicit `<w:defaultTabStop>` value when present, else the spec
 *  absent default of 720 twips (36pt). Mirrors {@link resolveKinsokuRules}: the
 *  resolved value is threaded into both the measure and draw passes so they
 *  agree. */
export function resolveDefaultTabPt(settings: DocSettings | undefined): number {
  const v = settings?.defaultTabStop;
  // §17.15.1.25 defines automatic stops as multiples of the interval, which is
  // undefined for a non-positive interval; fall back to the documented absent
  // default (720 twips = 36pt) so the automatic grid always advances.
  return v != null && v > 0 ? v : DEFAULT_TAB_PT;
}

/** One entry in a bidi line's LOGICAL-order sequence, for {@link layoutBidiTabStops}. */
export interface BidiTabItem {
  /** True for a tab segment (its width is (re)computed); false for content. */
  isTab: boolean;
  /** Content width in px (ignored for tabs). Set by the LTR layout pass. */
  width: number;
  /** Logical advance from this item's start to the resolved decimal alignment
   * point. The bidi resolver converts the complete cell's logical prefix to a
   * physical reading-frame offset. */
  decimalOffset?: number;
}

/** Per-segment result of {@link layoutBidiTabStops}. */
export interface BidiTabResult {
  /** New measuredWidth for the tab at this LOGICAL index (non-tabs: unchanged). */
  width: number;
  /** Leader to paint across this tab's span (`'none'`/undefined ⇒ blank). */
  leader?: TabStop['leader'];
}

/** Resolve ST_TabJc aliases to the physical role used by the paragraph's
 * reading frame. `start`/`end` are the strict logical aliases of
 * `left`/`right`; `num` is the leading tab between a list marker and its text.
 * Both the LTR and mirrored bidi algorithms operate in reading-frame
 * coordinates, so the role mapping itself is direction-independent. */
function tabAlignmentRole(
  alignment: TabStop['alignment'],
): 'leading' | 'center' | 'trailing' | 'decimal' {
  if (alignment === 'center') return 'center';
  if (alignment === 'decimal') return 'decimal';
  if (alignment === 'right' || alignment === 'end') {
    return 'trailing';
  }
  return 'leading';
}

/**
 * ECMA-376 §17.3.1.37 / §17.15.1.25 / §17.18.84 — lay out ONE line of a BIDI
 * (RTL-base) paragraph's tab-aligned cells, returning each tab's width and
 * leader BY LOGICAL INDEX.
 *
 * The LTR layout pass ({@link layoutLines}) resolves tab widths against the pen
 * in LOGICAL order but in a LEFT-to-right frame, which is wrong for a bidi
 * paragraph: a tab advances the pen in READING order, which under an RTL base
 * runs RIGHT-to-LEFT, and the tab-delimited cells then reorder visually. The LTR
 * result lands the trailing content (a TOC page number, a footer field) on the
 * wrong visual side — often overflowing and wrapping to a new line (the "leaders
 * appear/disappear" and "page number on its own row" symptoms of issue #820).
 *
 * This lays the line out in the RTL READING frame: the pen starts at the right
 * TEXT MARGIN (pen 0) and moves LEFT (increasing pen). A tab stop's `pos` is its
 * distance from that margin, not from the paragraph's indented edge. Content
 * begins at `startPenPx` (the paragraph's leading-indent + first-line
 * indent); the Nth tab in reading order advances to the next stop further left
 * (larger `pos`), exactly like the LTR pen advances rightward through stops.
 * Alignment is logical (Part 4 §14.11.2): physical `left` = `start` (leading ⇒
 * following content's leading/RIGHT edge on the stop), physical `right` = `end`
 * (trailing ⇒ its trailing/LEFT edge on the stop); `center` is unchanged;
 * `bar`/`clear` advance like `start`. Automatic stops fall on the §17.15.1.25
 * grid from the margin, after all custom stops. A stop past the LEFT text
 * margin (`leftLimitPx`) invokes `word-tab-stop-page-edge-clamp`.
 *
 * The widths returned here reproduce the intended layout through the draw
 * loop's visual walk because {@link computeLineVisualOrder} classifies tabs as UAX#9 S:
 * rule L2 then reverses cells AND tabs together, so the logical tab between
 * cells k−1 and k sits visually between the mirrored cells k and k−1 — its
 * reading-frame gap IS its visual gap. (This is why results map back by logical
 * index; resolving stops against the visual sequence would reverse the tab→stop
 * assignment and paint the leader in the wrong cell gap — the #830 follow-up
 * bug where the TOC leader appeared between the title and the chapter number
 * instead of between the page number and the title.)
 *
 * @param items line segments in LOGICAL order (as `line.segments`).
 * @param customStopsPx custom tab stops in margin px (`pos * scale`).
 * @param startPenPx reading-frame pen at the line's content start = the leading
 *   (logical-left ⇒ physical-right) indent, plus any first-line indent.
 * @param leftLimitPx reading-frame position of the LEFT text margin (= the
 *   margin-to-margin text width).
 * @param intervalPx automatic-stop interval = `defaultTabPt * scale`.
 * @returns one {@link BidiTabResult} per LOGICAL index (1:1 with `items`).
 */
export function layoutBidiTabStops(
  items: BidiTabItem[],
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  startPenPx: number,
  leftLimitPx: number,
  intervalPx: number,
): BidiTabResult[] {
  const n = items.length;
  const width = items.map((it) => it.width);
  const leader: (TabStop['leader'] | undefined)[] = new Array(n).fill(undefined);

  // Width of the content run immediately FOLLOWING index `i` in reading order
  // (up to the next tab / line end) — the trailing/centered stop needs it.
  const followAlignmentWidth = (
    from: number,
    role: ReturnType<typeof tabAlignmentRole>,
  ): Readonly<{ total: number; alignment: number }> => {
    let total = 0;
    let decimal: number | undefined;
    for (let j = from; j < n; j++) {
      if (items[j].isTab) break;
      if (decimal === undefined && items[j].decimalOffset !== undefined) {
        decimal = total + items[j].decimalOffset!;
      }
      total += width[j];
    }
    return {
      total,
      alignment: role === 'center'
        ? total / 2
        : role === 'trailing'
          ? total
          : role === 'decimal'
            ? decimal === undefined ? 0 : total - decimal
            : 0,
    };
  };

  // Reading-frame walk. `pen` = distance from the right TEXT MARGIN; content and
  // tabs push it further LEFT (increasing).
  let pen = startPenPx;
  for (let i = 0; i < n; i++) {
    const it = items[i];
    if (!it.isTab) {
      pen += width[i];
      continue;
    }
    const stop = nextTabStopRtl(pen, customStopsPx, intervalPx);
    if (!stop) {
      // No stop further left: the tab collapses (following content continues).
      width[i] = 0;
      continue;
    }
    // The tab's leading (right) edge sits at the pen; its trailing (left) edge
    // is the stop-aligned target, giving the gap it fills.
    const role = tabAlignmentRole(stop.alignment);
    const following = followAlignmentWidth(i + 1, role);
    const fw = following.total;
    let target: number; // pen value after the tab (its trailing/left edge)
    if (role !== 'leading') {
      // end aligns the full following cell, center half of it, and decimal the
      // reading-frame distance through the first halfwidth period. The
      // registered no-separator fallback uses the numeric cell's physical
      // right edge, which is the reading-leading edge in this mirrored frame.
      target = stop.pos - following.alignment;
    } else {
      // start/leading (or bar/clear/left): following content's LEADING (right)
      // edge on the stop.
      target = stop.pos;
    }
    // Pin content that would fall past the left text margin onto the margin: the
    // following cell spans [target, target + fw] in reading-frame margins, so its
    // far (left) edge must stay ≤ leftLimitPx.
    if (target + fw > leftLimitPx) target = leftLimitPx - fw;
    // Never let a tab move the pen backwards (right).
    if (target < pen) target = pen;
    width[i] = target - pen;
    leader[i] = stop.leader;
    pen = target;
  }

  return items.map((_, i) => ({ width: width[i], leader: leader[i] }));
}

/** Value equivalence of two resolved kinsoku rule sets, with a reference fast
 *  path. The reuse gate cannot rely on `===` alone: `resolveKinsokuRules` builds
 *  a FRESH object (fresh Sets) on every call, and canonical layout variants
 *  resolve it independently — same `doc.settings`, different references.
 *  Both derive from the same immutable settings so they are value-equal there;
 *  this check is pure defense so a genuinely different rule set (which would
 *  change CJK retract decisions in layoutLines) can never reuse stale lines. */
export function kinsokuRulesEquivalent(a: KinsokuRules, b: KinsokuRules): boolean {
  if (a === b) return true;
  if (a.enabled !== b.enabled) return false;
  const setEq = (x: Set<number>, y: Set<number>): boolean => {
    if (x.size !== y.size) return false;
    for (const cp of x) if (!y.has(cp)) return false;
    return true;
  };
  return setEq(a.lineStartForbidden, b.lineStartForbidden) && setEq(a.lineEndForbidden, b.lineEndForbidden);
}

/** Resolve §17.3.2.14 region geometry after raw Canvas advances are available.
 *  Region membership was fixed over tab-delimited source fragments in
 *  {@link buildSegments}; width and code-point count are deliberately derived
 *  here from the emitted segments so script/case transformations cannot create
 *  a second source of truth. */
function resolveFitTextSegments(
  segments: LayoutTextSeg[],
  scale: number,
  measureNaturalWidthPx: (segment: LayoutTextSeg) => number,
): void {
  const regionSegments = new Map<number, LayoutTextSeg[]>();
  for (const segment of segments) {
    if (segment.fitTextRegionIndex === undefined) continue;
    const members = regionSegments.get(segment.fitTextRegionIndex) ?? [];
    members.push(segment);
    regionSegments.set(segment.fitTextRegionIndex, members);
  }

  for (const members of regionSegments.values()) {
    const first = members.find((segment) => segment.fitTextVal !== undefined);
    if (!first || first.fitTextVal === undefined) continue;

    let naturalWidthPx = 0;
    let charCount = 0;
    for (const segment of members) {
      naturalWidthPx += measureNaturalWidthPx(segment) * charScaleFactor(segment);
      charCount += [...segment.text].length;
    }

    const resolved = groupFitTextRegions([{
      fitTextValTwips: first.fitTextVal,
      charCount,
      naturalWidthPx,
    }], scale)[0];
    if (!resolved) continue;
    members.forEach((segment, index) => {
      segment.fitTextPerGapPx = resolved.perGapPx;
      segment.fitTextTrailingPadPx = index === members.length - 1
        ? resolved.trailingPadPx
        : undefined;
      segment.fitTextRegionStart = index === 0 ? true : undefined;
      segment.fitTextRegionEnd = index === members.length - 1 ? true : undefined;
    });
  }
}

export function buildSegments(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const resolvedFont = (
    family: string | null | undefined,
    weight = 400,
    style: 'normal' | 'italic' = 'normal',
  ): ResolvedFontMetric | undefined => {
    if (!family) return undefined;
    const normalized = normalizeFontMetricFamily(family);
    const metrics = environment.resolvedLocalFonts;
    if (!metrics) return undefined;
    const tuple = metrics[`${normalized}:${weight}:${style}`];
    if (tuple) return tuple;
    const normal = metrics[normalized];
    if (weight === 400 && style === 'normal' && normal) return normal;
    return Object.values(metrics).find((metric) =>
      normalizeFontMetricFamily(metric.requestedFamily ?? '') === normalized
      && (metric.weight ?? 400) === weight
      && (metric.style ?? 'normal') === style,
    );
  };
  // Group §17.3.2.14 adjacency over SOURCE RUNS before script/font, word, or
  // small-caps segmentation, but model each tab-delimited fragment as its own
  // source unit. A tab is a position-dependent advance rather than a glyph, so a
  // non-fit kernel entry at every tab boundary prevents same-id fragments from
  // linking across it. Width/count placeholders are resolved from emitted text
  // at layout scale below.
  const fitTextFragmentEntryByKey = new Map<string, number>();
  const fitTextRuns: FitTextRun[] = [];
  for (const [runIndex, run] of runs.entries()) {
    if (run.type !== 'text') {
      fitTextRuns.push({ charCount: 0, naturalWidthPx: 0 });
      continue;
    }
    const fragments = run.text.split('\t');
    for (let fragmentIndex = 0; fragmentIndex < fragments.length; fragmentIndex += 1) {
      fitTextFragmentEntryByKey.set(`${runIndex}:${fragmentIndex}`, fitTextRuns.length);
      fitTextRuns.push({
        fitTextValTwips: run.fitTextVal,
        fitTextId: run.fitTextId,
        charCount: [...fragments[fragmentIndex]].length,
        naturalWidthPx: 0,
        charScale: run.charScale,
      });
      if (fragmentIndex < fragments.length - 1) {
        fitTextRuns.push({ charCount: 0, naturalWidthPx: 0 });
      }
    }
  }
  const fitTextRegionByEntry = new Map<number, number>();
  groupFitTextRegions(fitTextRuns, 1).forEach((region, regionIndex) => {
    for (let entryIndex = region.start; entryIndex < region.end; entryIndex += 1) {
      fitTextRegionByEntry.set(entryIndex, regionIndex);
    }
  });
  const pushTextPiece = (
    text: string,
    base: Extract<ParagraphLayoutSource['runs'][number], { type: 'text' | 'field' }>,
    vertAlign: 'super' | 'sub' | null,
    sourceRunIndex: number,
    sourceFragmentIndex?: number,
    joinPreviousRun = false,
  ) => {
    const r: ParagraphTextBearingRun = base;
    const acquiredTypography = (r as ParagraphTextBearingRun & Readonly<{
      typographyInput?: import('./layout/typography-input.js').RunTypographyAcquisitionInput;
    }>).typographyInput;
    // ECMA-376 §17.16.18 stores a complex field's instruction/result across
    // several physical runs. The parser rebuilds recomputed PAGE/NUMPAGES as a
    // single FieldRun, while its complete effective §17.3.2 run properties live
    // on the immutable typography acquisition sidecar. Consume those effective
    // facts exactly like an ordinary text run so the field result does not lose
    // baseline position or the core character-metric axes used by measure/paint.
    const acquiredValue = <T>(
      value: import('./layout/typography-input.js').TypographyValueInput<T> | undefined,
      fallback: T | undefined,
    ): T | undefined => value?.status === 'valid' && value.value !== null
      ? value.value
      : fallback;
    const effectiveVertAlign = acquiredValue(
      acquiredTypography?.verticalAlign,
      vertAlign ?? undefined,
    ) ?? null;
    const effectivePosition = acquiredValue(
      acquiredTypography?.positionPt,
      r.position,
    );
    const effectiveCharacterSpacing = acquiredTypography?.characterSpacingPt ?? r.charSpacing;
    // ECMA-376 §17.3.2.35 gives an authored run an explicit character pitch.
    // Word observation: a positive `w:spacing` owns that expanded pitch and suppresses
    // document-level §17.15.1.18 punctuation whitespace compression for the
    // run. Combining both adjustments collapses consecutive Japanese closing
    // punctuation even though Word preserves the authored spacing.
    const documentCharacterCompressionApplies =
      wordDocumentCharacterCompressionApplies(effectiveCharacterSpacing);
    const effectiveCharacterScale = acquiredTypography?.characterScale ?? r.charScale;
    const effectiveKerningThreshold = acquiredTypography?.kerningThresholdPt ?? r.kerning;
    const effectiveSnapToGrid = acquiredTypography?.snapToGrid ?? r.snapToGrid;
    // §17.3.2.33 small caps are sized per character: lowercase LETTERS render two
    // points smaller, uppercase letters and non-alphabetic characters at the full
    // run size. `reduced` (set per case-piece in the loop below) carries that onto
    // each emitted segment; calcEffectiveFontPx shrinks only the reduced ones.
    // allCaps (§17.3.2.5) and non-caps runs are a single, non-reduced piece.
    let reduced = false;
    // Ruby annotation rides with the WHOLE base text (typically 1-2 chars).
    // Splitting on word boundaries would lose the association, so attach
    // the annotation only to the first emitted segment.
    const baseRuby = r.ruby;
    const ruby = baseRuby
      ? {
          text: baseRuby.text,
          fontSizePt: baseRuby.fontSizePt,
          ...(baseRuby.hpsRaisePt != null ? { hpsRaisePt: baseRuby.hpsRaisePt } : {}),
        }
      : undefined;
    const revision = r.revision;
    const rtl = r.rtl === true ? true : undefined;
    const fitTextFragmentEntryIndex = sourceFragmentIndex === undefined
      ? undefined
      : fitTextFragmentEntryByKey.get(`${sourceRunIndex}:${sourceFragmentIndex}`);
    const fitTextRegionIndex = fitTextFragmentEntryIndex === undefined
      ? undefined
      : fitTextRegionByEntry.get(fitTextFragmentEntryIndex);

    // IX1 — resolve the run's hyperlink target ONCE (§17.16.22 external URL /
    // §17.16.23 internal anchor). An external URL (`r.hyperlink`) wins over the
    // internal `w:anchor` when both are present, matching the parser's rule. A
    // FieldRun carries neither field, so the `as DocxTextRun` guards yield
    // undefined. Purely a callback payload — it does not touch measurement.
    const hyperlink: HyperlinkTarget | undefined = r.hyperlink
      ? { kind: 'external', url: r.hyperlink }
      : r.hyperlinkAnchor
        ? { kind: 'internal', ref: r.hyperlinkAnchor }
        : undefined;

    // ECMA-376 §17.3.2.26 content classification. w:rtl/w:cs selects the cs
    // axis except for a character assigned eastAsia while rFonts@hint=eastAsia;
    // that protected span keeps the non-cs East Asian formatting axis.
    // NOTE rFonts@cs (fontFamilyCs) alone is just a font SLOT and must NOT
    // force cs — e.g. sample-1's Heading1 (Latin) has cstheme + szCs=52 but
    // renders at w:sz=24; forcing cs blew its size up to 26pt.
    const forceCs = r.rtl === true || r.cs === true;

    // Complex-script (cs) formatting sources. SIZE (§17.3.2.39 szCs) and TYPEFACE
    // (§17.3.2.26 rFonts@cs) fall back to their Latin counterpart when absent —
    // the parser resolves szCs through the full style chain, mirroring a
    // directly-set `w:sz` per §17.3.2.18. But BOLD (§17.3.2.3 bCs) and ITALIC
    // (§17.3.2.17 iCs) are independent toggles: absent `bCs`/`iCs` defaults off
    // and must not inherit Latin-axis `w:b`/`w:i`, which govern only non-complex
    // content.
    const csFontSize = r.fontSizeCs ?? base.fontSize;
    const csFontFamily = r.fontFamilyCs ?? base.fontFamily;
    const highAnsiFontFamily = r.fontFamilyHighAnsi ?? base.fontFamily;
    const csBold = r.boldCs ?? false;
    const csItalic = r.italicCs ?? false;

    // ECMA-376 §17.3.2.26 eastAsia axis. Within a non-complex-script slice, CJK
    // code points take the eastAsia face while Latin/digits keep the ascii face
    // (`base.fontFamily`). Only `DocxTextRun` carries the axis; absent (field
    // runs / single-axis parser output) ⇒ fall back to ascii. Text-box runs feed
    // this same builder (via `shapeRunToDocRun`), so a text box's per-script face
    // is picked here too. Bold/italic/size are NOT axis-specific here — eastAsia
    // shares the Latin (non-cs) toggles, so only the family differs.
    const eaFontFamily = r.fontFamilyEastAsia ?? base.fontFamily;

    // `word-rtl-complex-script-european-digits-an`: use the bidi language's
    // primary subtag when present, otherwise fall back to an rtl-marked run.
    const digitsAsAN =
      (forceCs || Boolean(r.rtl)) && isRtlBidiLang(r.langBidi, Boolean(r.rtl));

    let firstSeg = true;
    // True while the next emitted segment should be GLUED to the previous one
    // (a small-caps case-piece that continues the same word). Consumed by the
    // first pushSeg of the piece so only that segment carries joinPrev.
    let gluePending = false;
    // Script slot for an emitted segment (§17.3.2.26): 'cs' = complex-script
    // (Arabic/Hebrew/...), 'ea' = East-Asian (CJK → eastAsia face), 'latin' =
    // Latin/digits/neutral (ascii face). Each segment stays SINGLE-FONT — one
    // family for its whole `.text` — so the measure==draw / docGrid char-grid
    // invariant holds and the draw loop needs no per-segment font switching.
    const pushSeg = (
      text: string,
      cs: boolean,
      fontFamily: string | null,
      authoritativeSpan?: TextShapeSpan,
      compressCharacterWhitespace = false,
      mappedSymbolUnicode = false,
    ) => {
      if (
        environment.balanceSingleByteDoubleByteWidth
        && !cs
        && text.includes('\u3000')
        && [...text].some((character) => character !== '\u3000')
      ) {
        // The registered width-balance projection treats U+3000 as a half-delta space while
        // other East-Asian glyphs receive the full delta. Split only at that
        // semantic boundary so Canvas can retain one uniform letterSpacing per
        // segment (measure == paint); the space itself has no contextual shape.
        for (const part of text.split(/(\u3000+)/u).filter(Boolean)) {
          pushSeg(part, cs, fontFamily, undefined, compressCharacterWhitespace, mappedSymbolUnicode);
        }
        return;
      }
      // ECMA-376 §17.15.1.18 / §17.18.7 — dispatch the exact
      // ST_CharacterSpacing value and split each eligible full-width character
      // so its selected face's tight ink bounds can define the removable
      // whitespace. Non-eligible text retains contextual shaping.
      if (
        !compressCharacterWhitespace
        && documentCharacterCompressionApplies
        && fitTextRegionIndex === undefined
      ) {
        const boundaries = [0, ...graphemeClusterOffsets(text), text.length];
        const graphemes = boundaries.slice(0, -1).map(
          (start, index) => text.slice(start, boundaries[index + 1]),
        );
        if (graphemes.some((grapheme) =>
          characterSpacingControlCompresses(
            grapheme,
            environment.characterSpacingControl,
          ))) {
          pushSeg(text, cs, fontFamily, undefined, true, mappedSymbolUnicode);
          return;
        }
      }
      const bold = cs ? csBold : base.bold;
      const italic = cs ? csItalic : base.italic;
      const weight = bold ? 700 : 400;
      const style = italic ? 'italic' as const : 'normal' as const;
      const textShapeRequest: TextShapeRequest = Object.freeze({
        text,
        fontSizePt: cs ? csFontSize : base.fontSize,
        // A successfully decoded Symbol/Wingdings code point is Unicode text,
        // not a request for the legacy font encoding. Clear every authored
        // slot so the text service resolves a Unicode-capable generic route;
        // otherwise its §17.3.2.26 slot resolver selects Symbol again and the
        // mapped character can still be drawn with the wrong cmap.
        fonts: mappedSymbolUnicode
          ? { ascii: null, highAnsi: null, eastAsia: null, complexScript: null }
          : r.fontSlots?.direct ?? {
              ascii: base.fontFamily,
              highAnsi: highAnsiFontFamily,
              eastAsia: eaFontFamily,
              complexScript: csFontFamily,
            },
        themeFonts: mappedSymbolUnicode ? undefined : r.fontSlots?.theme,
        themeFontPresence: mappedSymbolUnicode ? undefined : r.fontSlots?.themePresent,
        weight,
        style,
        complexScript: cs,
        fontHint: r.fontHint,
        eastAsiaLanguage: r.langEastAsia,
        kerning: effectiveKerningThreshold == null
          ? undefined
          : (cs ? csFontSize : base.fontSize) >= effectiveKerningThreshold,
        measure: false,
      });
      const shaped = authoritativeSpan
        ? { spans: [authoritativeSpan] }
        : environment.layoutServices?.text.shape(textShapeRequest);
      const punctuationCompressions =
        compressCharacterWhitespace
          && documentCharacterCompressionApplies
        ? (() => {
            const boundaries = [0, ...graphemeClusterOffsets(text), text.length];
            const compressions: Array<{ end: number; adjustmentPt: number }> = [];
            for (let index = 0; index < boundaries.length - 1; index += 1) {
              const start = boundaries[index]!;
              const end = boundaries[index + 1]!;
              const compressedGrapheme = text.slice(start, end);
              if (!characterSpacingControlCompresses(
                compressedGrapheme,
                environment.characterSpacingControl,
              )) continue;
            const measured = environment.layoutServices?.text.shape({
              ...textShapeRequest,
              text: compressedGrapheme,
              measure: true,
              clusterGeometry: false,
            });
            // ECMA-376 §17.15.1.18 defines only compression eligibility. The
            // registered Word observation retains half of the selected route's
            // ideographic cell for punctuation; this is not assumed to be half
            // of a proportional punctuation glyph's own advance. Kana in
            // `compressPunctuationAndJapaneseKana` has no observed cell floor,
            // so only its measured trailing sidebearing is removed.
            const removableUnscaledPt = measured?.inkBounds
              && measured.horizontalInkBoundsAreTight === true
              ? (() => {
                  const trailingWhitespacePt = Math.max(
                    0,
                    Math.min(
                      measured.advancePt,
                      measured.advancePt - measured.inkBounds.xMaxPt,
                    ),
                  );
                  if (!COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(
                    compressedGrapheme,
                  )) {
                    return trailingWhitespacePt;
                  }
                  const punctuationRoute = measured.spans[0]?.fontRoute.fingerprint;
                  const ideographicCell = environment.layoutServices?.text.shape({
                    ...textShapeRequest,
                    // U+3000 is semantically an ideographic space, but several
                    // proportional East Asian faces expose it to Canvas with
                    // the same narrow advance as their punctuation. The grid's
                    // full-width character cell is represented by an
                    // ideograph, not by that platform-specific space metric.
                    text: '\u4e00',
                    fontHint: 'eastAsia',
                    measure: true,
                    clusterGeometry: false,
                  });
                  const cellRoute = ideographicCell?.spans[0]?.fontRoute.fingerprint;
                  const cellAdvancePt = ideographicCell?.advancePt;
                  if (
                    !punctuationRoute
                    || cellRoute !== punctuationRoute
                    || cellAdvancePt === undefined
                    || !Number.isFinite(cellAdvancePt)
                    || cellAdvancePt <= 0
                  ) {
                    return 0;
                  }
                  const retainedExtentPt = wordJapanesePunctuationRetainedExtentPt({
                    punctuationAdvancePt: measured.advancePt,
                    punctuationInkEndPt: measured.inkBounds.xMaxPt,
                    ideographicCellAdvancePt: cellAdvancePt,
                  });
                  return Math.max(
                    0,
                    Math.min(
                      trailingWhitespacePt,
                      measured.advancePt - retainedExtentPt,
                    ),
                  );
                })()
              : 0;
            // §17.3.2.43 w:w scales the glyph and both of its sidebearings;
            // trim in the same post-scale coordinate space as segAdvanceWidth.
            const removablePt = removableUnscaledPt * (effectiveCharacterScale ?? 1);
              if (removablePt > 0) {
                compressions.push({ end, adjustmentPt: -removablePt });
              }
            }
            return compressions.length === 0
              ? undefined
              : Object.freeze(compressions.map((compression) =>
                  Object.freeze(compression)));
          })()
        : undefined;
      const resolvedAxisDiffers = shaped?.spans.some((span) =>
        (span.script === 'complexScript') !== cs,
      ) ?? false;
      if (shaped && (shaped.spans.length > 1 || resolvedAxisDiffers)) {
        for (let spanIndex = 0; spanIndex < shaped.spans.length; spanIndex += 1) {
          const span = shaped.spans[spanIndex]!;
          const spanCs = span.script === 'complexScript';
          const spanFamily = spanCs
            ? csFontFamily
            : span.script === 'eastAsia'
              ? eaFontFamily
              : span.script === 'highAnsi'
                ? highAnsiFontFamily
                : base.fontFamily;
          const compressedSpan = documentCharacterCompressionApplies
            && compressCharacterWhitespace
            && [...span.text].some((grapheme) =>
              characterSpacingControlCompresses(
                grapheme,
                environment.characterSpacingControl,
              ));
          pushSeg(
            span.text,
            spanCs,
            spanFamily,
            span,
            compressedSpan,
          );
        }
        return;
      }
      const resolvedSpan = shaped?.spans[0];
      const serviceMetric = (resolvedFamily: string | undefined, requestedFamily?: string) => {
        if (!resolvedFamily) return undefined;
        const candidates = Object.values(
          environment.layoutServices?.text.fontMetrics
          ?? environment.layoutServices?.text.localMetrics
          ?? {},
        ).filter((metric) =>
            normalizeFontMetricFamily(metric.family) === normalizeFontMetricFamily(resolvedFamily)
            && (metric.weight ?? 400) === weight
            && (metric.style ?? 'normal') === style,
          );
        return candidates.find((metric) => requestedFamily
          && normalizeFontMetricFamily(metric.requestedFamily ?? '')
            === normalizeFontMetricFamily(requestedFamily)) ?? candidates[0];
      };
      const localFont = resolvedSpan
        ? serviceMetric(resolvedSpan.font.resolvedFamily, resolvedSpan.font.requestedFamily)
        : resolvedFont(fontFamily, weight, style);
      const eaResolution = environment.layoutServices?.text.resolve({
        fonts: textShapeRequest.fonts,
        themeFonts: textShapeRequest.themeFonts,
        themeFontPresence: textShapeRequest.themeFontPresence,
        slot: 'eastAsia',
        weight,
        style,
      });
      const localEaFloor = eaResolution
        ? serviceMetric(eaResolution.resolvedFamily, eaResolution.requestedFamily)
        : resolvedFont(eaFontFamily, weight, style);
      const familyLineMetric = localFont ?? resolvedFont(fontFamily, weight, style);
      const eaLineMetric = localEaFloor ?? resolvedFont(eaFontFamily, weight, style);
      const resolvedEaFloorFamily = eaResolution?.resolvedFamily
        ?? localEaFloor?.family
        ?? eaFontFamily;
      const useFeEastAsianMetric = environment.useFeLayout
        && (r.fontHint === 'eastAsia' || Boolean(resolvedEaFloorFamily?.trim()));
      const resolvedScript = resolvedSpan?.script ?? authoritativeSpan?.script
        ?? (cs ? 'complexScript' : EAST_ASIAN_RE.test(text) ? 'eastAsia' : 'ascii');
      const widthBalanceGridDeltaFactor = environment.balanceSingleByteDoubleByteWidth
        ? wordBalancedLinesAndCharsGridDeltaFactor(text, resolvedScript)
        : undefined;
      segs.push({
        text,
        script: resolvedScript,
        ...(widthBalanceGridDeltaFactor !== undefined
          ? {
              // §17.15.3.3 defines the SBCS:DBCS width ratio as 1:2; the
              // registered Word matrix defines how that setting projects onto
              // linesAndChars charSpace. Production shaping has already split
              // the segment at §17.3.2.26 script-slot boundaries.
              widthBalanceGridDeltaFactor,
            }
          : {}),
        ...(useFeEastAsianMetric
          ? { metricEastAsian: true as const }
          : {}),
        bold,
        italic,
        underline: base.underline,
        // §17.3.2.40 underline style / colour — carried only on DocxTextRun (a
        // FieldRun draws single). Kept raw ST_Underline; the renderer normalizes
        // to DrawingML §20.1.10.82 at draw time.
        underlineStyle: r.underlineStyle,
        underlineColor: r.underlineColor,
        strikethrough: base.strikethrough,
        fontSize: cs ? csFontSize : base.fontSize,
        color: base.color,
        fontFamily: resolvedSpan?.font.resolvedFamily ?? localFont?.family ?? fontFamily,
        fontRoute: resolvedSpan?.fontRoute,
        resolvedLineHeightRatio: familyLineMetric?.lineHeightRatio,
        resolvedEastAsianLineHeightRatio: familyLineMetric?.eastAsianLineHeightRatio,
        vertAlign: effectiveVertAlign,
        measuredWidth: 0,
        textLayoutService: environment.layoutServices?.text,
        textShapeRequest,
        breakBefore: resolvedSpan?.breakBefore ?? authoritativeSpan?.breakBefore ?? true,
        smallCaps: reduced,
        joinPrev: (
          (firstSeg && (
            r.noBreakBefore === true
            || joinPreviousRun
          ))
          || gluePending
          || authoritativeSpan?.breakBefore === false
        ) ? true : undefined,
        hardJoinPrev: (
          firstSeg && (r.noBreakBefore === true || joinPreviousRun)
        ) ? true : undefined,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
        // §17.3.2.12 w:em — carried on both DocxTextRun and FieldRun (a field's
        // resolved/fallback text stamps the mark the same as a plain run).
        emphasisMark: base.emphasisMark,
        background: base.background ?? null,
        colorAuto: r.colorAuto ?? false,
        border: r.border ?? null,
        ruby: firstSeg ? ruby : undefined,
        revision,
        ...(revision && environment.showTrackedChanges === true ? {
          trackChangesMarkup: {
            kind: revision.kind,
            authorColor: environment.revisionAuthorColor?.(revision.author) ?? '#C00000',
          },
        } : {}),
        rtl,
        digitsAsAN: digitsAsAN ? true : undefined,
        // §17.3.2.26 declared eastAsia axis — used by text-box line floors and
        // the compatibility-owned useFELayout body metric path.
        eaFloorFamily: resolvedEaFloorFamily,
        eaFloorRoute: eaResolution?.route,
        resolvedEaFloorLineHeightRatio: eaLineMetric?.lineHeightRatio,
        resolvedEaFloorEastAsianLineHeightRatio: eaLineMetric?.eastAsianLineHeightRatio,
        textBoxLineFloor: (r as DocxTextRun & { textBoxLineFloor?: boolean }).textBoxLineFloor,
        textBoxVertical: (r as DocxTextRun & { textBoxVertical?: boolean }).textBoxVertical,
        // IX1 — resolved hyperlink target of the originating run, for the
        // text-layer clickable overlay and URL-aware line-break opportunities.
        // It does not change glyph measurement or drawing.
        hyperlink,
        snapToCharacterGrid: effectiveSnapToGrid !== false,
        // WD4 — run character metrics (§17.3.2.35 spacing / §17.3.2.43 w /
        // §17.3.2.24 position / §17.3.2.19 kern). Uniform across the run, so
        // every emitted segment carries the same values; the measure and paint
        // passes apply them identically (measure==paint).
        charSpacing: effectiveCharacterSpacing,
        punctuationCompressions,
        eastAsiaLanguage: r.langEastAsia,
        charScale: effectiveCharacterScale,
        fitTextVal: fitTextRegionIndex === undefined ? undefined : r.fitTextVal,
        fitTextId: fitTextRegionIndex === undefined ? undefined : r.fitTextId,
        fitTextRegionIndex,
        fitTextRunIndex: fitTextRegionIndex === undefined ? undefined : fitTextFragmentEntryIndex,
        position: effectivePosition,
        positionExtendsLineBox: environment.positionExtendsLineBox !== false,
        kerning: effectiveKerningThreshold,
        // ECMA-376 §17.3.2.10 eastAsianLayout — 縦中横 is meaningful ONLY in a
        // vertical (tbRl) page, so fold the vertical gate in HERE at build time
        // (buildSegments receives it through LineLayoutEnvironment). Measure/paint then read a single
        // pre-gated flag. `vertCompress` rides only when `vert` is set (spec: it
        // is ignored otherwise).
        tateChuYoko: environment.verticalCJK && r.eastAsianVert === true ? true : undefined,
        tateChuYokoCompress:
          environment.verticalCJK && r.eastAsianVert === true && r.eastAsianVertCompress === true
            ? true
            : undefined,
        // #1014 — an upright-vertical (tbRl) per-glyph segment (NOT a 縦中横 cell,
        // which is one drawTateChuYokoRun cell). Marks the segment for the vo=Tr
        // rotate-fallback ink-extent advance correction in the measure passes.
        verticalRun:
          environment.verticalCJK && r.eastAsianVert !== true ? true : undefined,
      });
      firstSeg = false;
      gluePending = false; // glue applies only to a piece's FIRST segment
    };
    const emit = (word: string, slot: 'cs' | 'ea' | 'latin') => {
      const cs = slot === 'cs';
      const fontFamily = slot === 'cs' ? csFontFamily : slot === 'ea' ? eaFontFamily : base.fontFamily;
      // ECMA-376 §17.3.2.26 + §17.3.3.30: a run whose rFonts axis is Symbol or
      // Wingdings stores glyphs as the FONT's own (private) code points — Word
      // commonly in the PUA (U+F020–U+F0FF). Those render as tofu in any
      // fallback face, so normalize each character to its Unicode equivalent
      // (core `symbolTextToUnicodeSegments`, the same table the list marker uses
      // via `symbolFontToUnicode`). The string is split at mapped/unmapped
      // boundaries: a MAPPED run is drawn in a generic fallback (fontFamily=null
      // → sans tail with the dingbat glyphs; keeping the symbol family would let
      // an installed Symbol/Wingdings re-interpret the Unicode code point as the
      // WRONG glyph), while an UNMAPPED run keeps the symbol family so a host
      // that ships Symbol/Wingdings still draws its native glyph. Done once at
      // build time so measure==draw (the seg.text is never transformed later).
      if (isSymbolFontFamily(fontFamily)) {
        for (const part of symbolTextToUnicodeSegments(word, fontFamily)) {
          pushSeg(
            part.text,
            cs,
            part.mapped ? null : fontFamily,
            undefined,
            false,
            part.mapped,
          );
        }
        return;
      }
      pushSeg(word, cs, fontFamily);
    };

    // A non-complex-script slice still mixes scripts at the CJK boundary: emit
    // its maximal CJK runs on the 'ea' (eastAsia) slot and the rest on 'latin'
    // (ascii). Keeps each emitted segment single-font (so a serif ascii digit
    // sits next to a gothic eastAsia title) without changing the cs path.
    const emitNonCs = (slice: string) => {
      if (environment.layoutServices?.text) {
        // `w:sym` is parsed as a one-run private-encoding character carrying
        // its own Symbol/Wingdings family (§17.3.3.30). Keep normalization in
        // front of the service-backed script splitter as well as the legacy
        // path; otherwise a PUA code point such as Symbol F0B0 reaches Canvas
        // unchanged and renders as tofu.
        if (isSymbolFontFamily(base.fontFamily)) {
          emit(slice, 'latin');
          return;
        }
        pushSeg(slice, false, base.fontFamily);
        return;
      }
      for (const part of splitByEastAsia(slice)) emit(part.text, part.ea ? 'ea' : 'latin');
    };

    // Small caps split the run into full-size (uppercase-origin / non-cased) and
    // reduced (lowercase-origin) case-pieces; everything else is one piece. Each
    // piece is still UPPERCASED for display (allCaps or smallCaps), and `reduced`
    // drives its segments' size — see splitSmallCapsCase / calcEffectiveFontPx.
    const casePieces = base.smallCaps
      ? splitSmallCapsCase(text)
      : [{ text, reduced: false }];
    let prevPieceText = '';
    for (const piece of casePieces) {
      reduced = piece.reduced;
      // Glue this piece's FIRST segment to the previous piece when they continue
      // the same word (the previous piece did not end at a space) — so a
      // small-caps word's full-cap initial and reduced remainder stay on one line.
      gluePending = prevPieceText.length > 0 && !/\s$/.test(prevPieceText);
      prevPieceText = piece.text;
      const displayText = (base.allCaps || base.smallCaps) ? piece.text.toUpperCase() : piece.text;
      for (const word of splitTextForLayout(displayText)) {
        if (forceCs) {
          // When the run's digits are AN-classified, split a token into maximal
          // digit-groups and the surrounding separators so the per-line bidi pass
          // (which reorders at SEGMENT granularity) can place the groups in Word's
          // order — e.g. "28-02-2026" → segments [28][-][02][-][2026] reordered to
          // 2026-02-28. Canvas only reorders WITHIN a fillText using EN semantics,
          // so a single-segment date would otherwise stay 28-02-2026.
          if (digitsAsAN) {
            for (const slice of splitDigitGroups(word)) emit(slice, 'cs');
          } else {
            emit(word, 'cs');
          }
        } else {
          // Mixed Arabic+Latin word (no w:rtl / w:cs): split at script boundaries
          // so each side gets its own (cs vs Latin) size and typeface; the non-cs
          // side then sub-splits at CJK boundaries for the eastAsia face.
          // ECMA-376 §17.3.2.26 selects the cs axis only when w:cs/w:rtl
          // forces the run. Arabic/Hebrew code points in an ordinary run stay
          // on ascii/hAnsi; the text service performs the remaining grapheme-
          // safe East Asian slot split.
          emitNonCs(word);
        }
      }
    }
  };

  let joinNextVisibleText = false;
  for (const [runIndex, run] of runs.entries()) {
    // ECMA-376 §17.13.5 final view (the default): deleted (`w:del`,
    // §17.13.5.14) and moved-away (`w:moveFrom`, §17.13.5.22) content is not
    // part of the document's final state, so no segment is produced and line
    // breaking/pagination see the accepted document state. The markup view
    // (`showTrackedChanges`) keeps every revision run visible so it can be
    // decorated. Insertions/moveTo render in both views, and revision metadata
    // remains available through the parsed model for consumer-owned review UI.
    const runRevisionKind = (run as { revision?: { kind?: string } }).revision?.kind;
    if (
      environment.showTrackedChanges !== true
      && (runRevisionKind === 'deletion' || runRevisionKind === 'moveFrom')
    ) {
      continue;
    }
    const joinFromPreviousNoBreakHyphen = joinNextVisibleText;
    joinNextVisibleText = run.type === 'text'
      && (run as ParagraphTextBearingRun).noBreakAfter === true;
    const emittedStart = segs.length;
    if (run.type === 'text') {
      const t = run as unknown as DocxTextRun & { type: 'text' };
      // ECMA-376 §17.11: substitute a footnote/endnote reference marker's glyph
      // with the note's resolved sequential number. The body `*Reference` run
      // carries the id; the in-note `*Ref` placeholder carries an empty id, so
      // we fall back to the note number currently being drawn.
      const noteText =
        t.noteRef
          ? (t.noteRef.id
              ? environment.noteNumbers?.get(`${t.noteRef.kind}:${t.noteRef.id}`)
              : environment.noteReferenceNumber)
          : undefined;
      if (t.noteRef) {
        const label = noteText != null ? String(noteText) : (t.text || '');
        if (label.length > 0) {
          pushTextPiece(
            label,
            t,
            t.vertAlign ?? 'super',
            runIndex,
            0,
            joinFromPreviousNoBreakHyphen,
          );
        }
        for (let index = emittedStart; index < segs.length; index += 1) {
          segs[index].sourceRunIndex = runIndex;
        }
        continue;
      }
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) {
          pushTextPiece(
            parts[i],
            t,
            t.vertAlign,
            runIndex,
            i,
            i === 0 && joinFromPreviousNoBreakHyphen,
          );
        }
        if (i < parts.length - 1) {
          segs.push({
            isTab: true, fontSize: t.fontSize, measuredWidth: 0,
            bold: t.bold, italic: t.italic, sourceRunIndex: runIndex,
          });
        }
      }
    } else if (run.type === 'image') {
      const img = run;
      segs.push({
        imagePath: img.imagePath,
        mimeType: img.mimeType,
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        rotation: img.rotation,
        flipH: img.flipH,
        flipV: img.flipV,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        duotone: img.duotone,
        alpha: img.alpha,
        srcRect: img.srcRect ?? undefined,
        measuredWidth: 0,
      });
    } else if (run.type === 'chart') {
      // ECMA-376 §21.2 chart. Flow it as a picture box of the `<wp:extent>`
      // natural size: the same LayoutImageSeg shape (empty `imagePath`/
      // `mimeType` sentinels so `'imagePath' in seg` routes it through the image
      // measurement/split path) with only a chart resource marker; the model
      // payload remains owned by the paint resource registry.
      //
      // A `<wp:anchor>` (floating) chart (§20.4.2.3) carries `anchor: true` and
      // its parsed page-offset fields, exactly like an anchor ImageRun: the
      // measure pass zeroes an anchor seg's width (it is not part of the inline
      // flow) and anchor acquisition retains it at the resolved absolute box.
      const chartRun = run;
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: chartRun.widthPt,
        heightPt: chartRun.heightPt,
        anchor: chartRun.anchor ?? false,
        anchorXPt: chartRun.anchorXPt ?? 0,
        anchorYPt: chartRun.anchorYPt ?? 0,
        anchorXFromMargin: chartRun.anchorXFromMargin ?? false,
        anchorYFromPara: chartRun.anchorYFromPara ?? false,
        chart: true,
        chartResourceKey: (chartRun as Partial<import('./layout/text.js').ParagraphChartRun>).resourceKey,
        measuredWidth: 0,
      });
    } else if (run.type === 'shape' && run.inline === true) {
      // `wp:inline` hosts arbitrary DrawingML, including WPS shapes (§20.4.2.8).
      // Reserve its extent in the same line-breaking path as an inline picture;
      // paragraph acquisition replaces the sentinel with a retained drawing
      // placement at the resolved pen position.
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: run.widthPt,
        heightPt: run.heightPt,
        anchor: false,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: false,
        anchorYFromPara: false,
        inlineShape: true,
        measuredWidth: 0,
      });
    } else if (run.type === 'unavailableDrawing') {
      const acquiredAnchor = 'anchorAcquisitionInput' in run
        ? run.anchorAcquisitionInput
        : undefined;
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: run.widthPt,
        heightPt: run.heightPt,
        anchor: acquiredAnchor !== undefined,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: false,
        anchorYFromPara: false,
        unavailableResourceKind: run.resourceKind,
        measuredWidth: 0,
      });
    } else if (run.type === 'break') {
      if (run.breakType === 'line') {
        // Determine font size for the line break height from surrounding text runs
        const fontSize = findNearbyFontSize(runs, runs.indexOf(run));
        segs.push({ lineBreak: true, fontSize, measuredWidth: 0 });
      }
      // page/column breaks handled at the document level (splitPages)
    } else if (run.type === 'field') {
      const f = run as unknown as FieldRun & { type: 'field' };
      const text = resolveFieldText(f, environment);
      if (text) {
        pushTextPiece(
          text,
          f,
          f.vertAlign,
          runIndex,
          undefined,
          joinFromPreviousNoBreakHyphen,
        );
      }
    } else if (run.type === 'math') {
      // The parser resolves the paragraph font size; fall back to a nearby run only
      // if it is somehow absent.
      const fontSize = run.fontSize || findNearbyFontSize(runs, runs.indexOf(run));
      const resourceKey = 'resourceKey' in run ? run.resourceKey : undefined;
      if (environment.layoutServices && !resourceKey) {
        throw new Error('Service-backed math layout requires a normalized structural resource key');
      }
      const mathMetadata = resourceKey
        ? environment.layoutServices?.math.resolve(resourceKey)
        : undefined;
      segs.push({
        math: true,
        mathResourceKey: resourceKey ?? '',
        mathMetadata,
        display: run.display,
        fontSize,
        color: null,
        fallbackText: 'fallbackText' in run ? run.fallbackText : mathFallbackText(run.nodes),
        measuredWidth: 0,
        mathAscent: 0,
        mathDescent: 0,
        jc: run.jc,
      });
    } else if (run.type === 'ptab') {
      // ECMA-376 §17.3.3.23 absolute-position tab. Emit a tab segment carrying the
      // ptab descriptor; layoutLines resolves it to an absolute X (independent of
      // the paragraph's tab stops) and fills the gap with the run's leader.
      segs.push({
        isTab: true,
        fontSize: run.fontSize || findNearbyFontSize(runs, runs.indexOf(run)),
        measuredWidth: 0,
        leader: run.leader,
        ptab: { alignment: run.alignment, relativeTo: run.relativeTo },
      });
    } else if (run.type === 'anchorHost') {
      const eastAsian = run.fontFamilyEastAsia != null;
      const bold = run.bold ?? false;
      const italic = run.italic ?? false;
      const authoredFamily = run.fontFamilyEastAsia ?? run.fontFamily ?? null;
      const weight = bold ? 700 : 400;
      const style = italic ? 'italic' as const : 'normal' as const;
      const localFont = resolvedFont(authoredFamily, weight, style);
      const localEaFloor = resolvedFont(run.fontFamilyEastAsia ?? null, weight, style);
      const familyLineMetric = localFont ?? (authoredFamily
        ? environment.resolvedLocalFonts?.[normalizeFontMetricFamily(authoredFamily)]
        : undefined);
      const eaLineMetric = localEaFloor ?? (run.fontFamilyEastAsia
        ? environment.resolvedLocalFonts?.[normalizeFontMetricFamily(run.fontFamilyEastAsia)]
        : undefined);
      segs.push({
        text: '',
        metricOnly: true,
        ...(eastAsian ? { metricEastAsian: true as const } : {}),
        bold,
        italic,
        underline: false,
        strikethrough: false,
        fontSize: run.fontSize,
        color: null,
        fontFamily: localFont?.family ?? authoredFamily,
        resolvedLineHeightRatio: familyLineMetric?.lineHeightRatio,
        resolvedEastAsianLineHeightRatio: familyLineMetric?.eastAsianLineHeightRatio,
        vertAlign: null,
        measuredWidth: 0,
        eaFloorFamily:
          localEaFloor?.family ?? run.fontFamilyEastAsia ?? null,
        resolvedEaFloorLineHeightRatio: eaLineMetric?.lineHeightRatio,
        resolvedEaFloorEastAsianLineHeightRatio: eaLineMetric?.eastAsianLineHeightRatio,
        snapToCharacterGrid: false,
      });
    }
    for (let index = emittedStart; index < segs.length; index += 1) {
      segs[index].sourceRunIndex = runIndex;
    }
  }

  // Project acquisition-owned no-break ranges through the display case
  // transform and onto the single-font layout segments produced above.
  for (const [runIndex, run] of runs.entries()) {
    if (run.type !== 'text') continue;
    const textRun = run as Extract<ParagraphTextBearingRun, { type: 'text' }>;
    const sourceRanges = textRun.noBreakRanges;
    if (!sourceRanges || sourceRanges.length === 0) continue;
    const displayedRanges = sourceRanges.map((range) => {
      const transformOffset = (offset: number) => {
        const prefix = textRun.text.slice(0, offset);
        return textRun.allCaps || textRun.smallCaps ? prefix.toUpperCase().length : prefix.length;
      };
      return { start: transformOffset(range.start), end: transformOffset(range.end) };
    });
    let displayedCursor = 0;
    for (const candidate of segs) {
      if (candidate.sourceRunIndex !== runIndex) continue;
      if (!('text' in candidate)) {
        if ('isTab' in candidate) displayedCursor += 1;
        continue;
      }
      const segmentEnd = displayedCursor + candidate.text.length;
      if (
        displayedCursor > 0
        && displayedRanges.some((range) =>
          range.start === displayedCursor || range.end === displayedCursor)
      ) {
        candidate.joinPrev = true;
        candidate.hardJoinPrev = true;
      }
      const local = displayedRanges
        .filter((range) => range.start >= displayedCursor && range.end <= segmentEnd)
        .map((range) => Object.freeze({
          start: range.start - displayedCursor,
          end: range.end - displayedCursor,
        }));
      if (local.length > 0) candidate.noBreakRanges = Object.freeze(local);
      displayedCursor = segmentEnd;
    }
  }

  // Project the registered `word-external-link-syntax-breaks` opportunities
  // across the complete semantic link and all formatting seams first, then
  // distribute them onto the existing segments.
  // Segments stay intact unless a real overflow selects one of those offsets,
  // preserving contextual shaping, decoration geometry, and paint identity on
  // lines that do not wrap.
  for (let groupStart = 0; groupStart < segs.length;) {
    const first = segs[groupStart];
    if (!('text' in first) || first.hyperlink?.kind !== 'external') {
      groupStart += 1;
      continue;
    }
    const target = first.hyperlink.url;
    let groupEnd = groupStart;
    const group: LayoutTextSeg[] = [];
    while (groupEnd < segs.length) {
      const candidate = segs[groupEnd];
      if (
        !('text' in candidate)
        || candidate.hyperlink?.kind !== 'external'
        || candidate.hyperlink.url !== target
      ) break;
      group.push(candidate);
      groupEnd += 1;
    }
    const groupText = group.map((segment) => segment.text).join('');
    const protectedOffsets = new Set<number>();
    let cursor = 0;
    for (const segment of group) {
      for (const range of segment.noBreakRanges ?? []) {
        const offsets = [range.start, range.end];
        for (const offset of offsets) {
          protectedOffsets.add(cursor + offset);
        }
      }
      cursor += segment.text.length;
    }
    const legalOffsets = new Set<number>();
    for (const match of groupText.matchAll(/\S+/gu)) {
      const token = match[0];
      const tokenStart = match.index;
      const graphemeBoundaries = new Set(
        graphemeClusterOffsets(token).map((offset) => tokenStart + offset),
      );
      const tokenProtected = new Set(
        [...protectedOffsets]
          .filter((offset) => offset > tokenStart && offset <= tokenStart + token.length)
          .map((offset) => offset - tokenStart),
      );
      const tokenGraphemes = new Set(
        [...graphemeBoundaries].map((offset) => offset - tokenStart),
      );
      for (const offset of wordExternalLinkSyntaxBreakOffsets(
        token,
        tokenGraphemes,
        tokenProtected,
      )) legalOffsets.add(tokenStart + offset);
    }
    if (legalOffsets.size === 0) {
      groupStart = groupEnd;
      continue;
    }
    cursor = 0;
    for (let index = 0; index < group.length; index += 1) {
      const segment = group[index]!;
      const segmentStart = cursor;
      const segmentEnd = segmentStart + segment.text.length;
      const localBreaks = [...legalOffsets]
        .filter((offset) => offset > segmentStart && offset < segmentEnd)
        .map((offset) => offset - segmentStart)
        .sort((a, b) => a - b);
      if (localBreaks.length > 0) {
        segment.externalLinkBreakOffsets = Object.freeze(localBreaks);
      }
      if (index > 0 && legalOffsets.has(segmentStart)) {
        segment.joinPrev = undefined;
        segment.externalLinkBreakBefore = true;
      }
      cursor = segmentEnd;
    }
    groupStart = groupEnd;
  }

  if (environment.balanceSingleByteDoubleByteWidth) {
    // ECMA-376 §17.15.3.3 normatively requests a 1:2 SBCS/DBCS width balance,
    // but does not define how a proportional inter-word separator becomes a
    // fixed-pitch half-width cell. The registered Word observation limits that
    // projection to two-or-more explicitly authored U+0020 spaces. One normal
    // separator remains at its natural proportional advance.
    const metricCache = new Map<string, number>();
    const adjustmentFor = (segment: LayoutTextSeg): number | undefined => {
      const service = segment.textLayoutService;
      const request = segment.textShapeRequest;
      if (!service || !request) return undefined;
      const effectiveFontSizePt = calcEffectiveFontPx(segment, 1);
      const key = [
        service.fingerprint,
        segment.fontRoute?.fingerprint ?? 'implicit-latin',
        segment.eaFloorRoute?.fingerprint ?? 'implicit-east-asia',
        effectiveFontSizePt,
        segment.bold ? 700 : 400,
        segment.italic ? 'italic' : 'normal',
        segment.kerning ?? 'auto',
      ].join('|');
      const cached = metricCache.get(key);
      if (cached !== undefined) return cached;
      const naturalSpace = service.shape({
        ...request,
        text: ' ',
        fontSizePt: effectiveFontSizePt,
        measure: true,
        clusterGeometry: false,
      }).advancePt;
      const ideographicCell = service.shape({
        ...request,
        text: '\u4e00',
        fontSizePt: effectiveFontSizePt,
        fontHint: 'eastAsia',
        measure: true,
        clusterGeometry: false,
      }).advancePt;
      if (
        !Number.isFinite(naturalSpace)
        || !Number.isFinite(ideographicCell)
        || naturalSpace < 0
        || ideographicCell <= 0
      ) return undefined;
      const adjustmentPt = ideographicCell / 2 - naturalSpace;
      metricCache.set(key, adjustmentPt);
      return adjustmentPt;
    };
    let sequence: LayoutTextSeg[] = [];
    let sequenceCount = 0;
    const flushSequence = () => {
      if (wordBalancedConsecutiveSpaceCellApplies(sequenceCount)) {
        for (const segment of sequence) {
          segment.widthBalanceSpaceSequence = true;
          const adjustmentPt = adjustmentFor(segment);
          if (adjustmentPt !== undefined) {
            segment.widthBalanceSpaceAdjustmentPt = adjustmentPt;
          }
        }
      }
      sequence = [];
      sequenceCount = 0;
    };
    for (const candidate of segs) {
      if (!('text' in candidate) || candidate.script === 'complexScript') {
        flushSequence();
        continue;
      }
      const trailingSpaces = candidate.text.length - candidate.text.replace(/ +$/u, '').length;
      const spaceOnly = trailingSpaces > 0 && trailingSpaces === candidate.text.length;
      if (!spaceOnly) flushSequence();
      if (trailingSpaces > 0) {
        sequence.push(candidate);
        sequenceCount += trailingSpaces;
      } else {
        flushSequence();
      }
    }
    flushSequence();
  }

  // ── UAX#14 LB13 / ECMA-376 §17.15.1.59 (行頭禁則 — line-start-forbidden) ──────
  // A closing / mid-punctuation code point (comma, period, ; : ! ? ) ] } and
  // their CJK forms) carries NO line-break opportunity before it, so it may
  // never BEGIN a line. When such a char OPENS a segment that is glued to the
  // previous text segment — no intervening whitespace, e.g. a comma authored in
  // its own run as in sample-12's "…detection system" | ", metadata" — mark it
  // `joinPrev` so the group machinery in layoutLines keeps it with the preceding
  // word and wraps "system," together instead of orphaning "," at the next
  // line's head.
  //
  // This is a UNIVERSAL Latin/Western rule (UAX#14 LB13), NOT the East-Asian
  // kinsoku feature, so it consults the application's DEFAULT forbidden table
  // UNCONDITIONALLY — independent of the document's §17.3.1.16 `w:kinsoku`
  // toggle and of any custom §17.15.1.59 `w:noLineBreaksBefore` set (which
  // REPLACES the default East-Asian table for a language and so must NOT be able
  // to drop the ASCII non-starters and re-orphan a Latin comma). The document's
  // kinsoku settings still govern the separate per-character CJK retract paths
  // (kinsokuAdjustedSplit / crossRunKinsokuRetract), which read the layout kinsoku argument.
  // The ASCII non-starters (!),.:;?]}) live in that default table (core
  // rules.ts), so one membership test covers Latin and (incidentally) CJK forms.
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev) continue;
    const firstCp = cur.text.codePointAt(0);
    if (firstCp === undefined || !DEFAULT_KINSOKU_RULES.lineStartForbidden.has(firstCp)) continue;
    const prev = segs[i - 1];
    // Only glue across a boundary that is NOT already a break opportunity: the
    // preceding unit must be text that does not end in whitespace (a trailing
    // space is a legal break, so the mark may legitimately start the line).
    if (!('text' in prev) || /\s$/.test(prev.text)) continue;
    cur.joinPrev = true;
  }

  // Preserve the established Word/JLReq line-end allowance for U+3000 across
  // internal script/font/width-balance shaping seams. U+3000 is BA in UAX #14,
  // so the break opportunity is after the space; splitting the space into its
  // own internal segment must not invent an opportunity before it. A real
  // U+0020 source-run boundary is delegated to
  // wordSourceRunSpaceContinuesSequence below.
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (
      !('text' in cur)
      || cur.joinPrev
      || (cur.text[0] !== ' ' && cur.text[0] !== '\u3000')
    ) continue;
    const prev = segs[i - 1];
    if (!('text' in prev)) continue;
    const trailingSpaceFromSameRun = cur.sourceRunIndex === prev.sourceRunIndex;
    const compatibleSourceBoundary = wordSourceRunSpaceContinuesSequence(
      prev.text,
      cur.text,
    );
    if (
      !trailingSpaceFromSameRun
      && !compatibleSourceBoundary
    ) continue;
    cur.joinPrev = true;
  }

  // ── UAX #14 no-break pairs (LB14/LB23/LB23a/LB24/LB25/LB28/LB30) ──
  // buildSegments intentionally splits at run / font-script boundaries, but
  // those formatting seams are not line-break opportunities. Mark the following
  // segment so layoutLines' existing atomic-group pre-flush selects the previous
  // real opportunity instead. The shared predicate is deliberately one-way:
  // false means unsupported/deferred, never "break allowed".
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (
      !('text' in cur)
      || cur.joinPrev
      || cur.externalLinkBreakBefore
      || cur.text.length === 0
    ) continue;
    const prev = segs[i - 1];
    if (!('text' in prev) || prev.text.length === 0) continue;

    // Whitespace is an actual wrap boundary. Check both sides because source
    // runs may start with whitespace even though ASCII spaces normally remain
    // attached to the preceding splitTextForLayout token.
    if (/\s$/u.test(prev.text) || /^\s/u.test(cur.text)) continue;

    const prevChar = [...prev.text].at(-1);
    const nextChar = [...cur.text][0];
    const prevCp = prevChar?.codePointAt(0);
    const nextCp = nextChar?.codePointAt(0);
    if (prevCp === undefined || nextCp === undefined) continue;

    // U+200B is the explicit zero-width-space opportunity from LB8 and is not
    // included in JavaScript's \s character class.
    if (prevCp === 0x200b || nextCp === 0x200b) continue;

    // SEA uses the application's dictionary tailoring, so the LB1 SA→AL default
    // must not suppress a real word boundary. CJK keeps its established
    // per-character split / kinsoku path and sparse-line safeguards.
    if (containsSeaScript(prev.text) || containsSeaScript(cur.text)) continue;
    if (hasCJKBreakOpportunity(prev.text) || hasCJKBreakOpportunity(cur.text)) continue;

    if (isUax14NoBreakPair(prevCp, nextCp)) cur.joinPrev = true;
  }

  // §17.3.2.14 fitText is a fixed-width, non-wrapping unit. Glue every segment
  // after the first in the RUN-grouped region, including script/small-caps
  // pieces emitted from the same source run.
  const seenFitTextRegions = new Set<number>();
  for (const seg of segs) {
    if (!('text' in seg) || seg.fitTextRegionIndex === undefined) continue;
    if (seenFitTextRegions.has(seg.fitTextRegionIndex)) seg.joinPrev = true;
    else {
      seg.fitTextRegionStart = true;
      seenFitTextRegions.add(seg.fitTextRegionIndex);
    }
  }

  retainHorizontalPunctuationInkClearance(segs);

  return segs;
}

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
  widthPolicy?: 'bounded' | 'intrinsic',
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  overflowPunct?: boolean,
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
  widthPolicy: 'bounded' | 'intrinsic' = 'bounded',
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  overflowPunct = false,
  passContext?: Readonly<{
    probeHeights: readonly number[] | null;
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
      { probeHeights, preparedFloatWrap },
    );
    if (!wrapCtx || widthPolicy === 'intrinsic') return runPass(null);
    const preparedFloatWrap = wrapCtx.lineWindow
      ? undefined
      : prepareFloatWrap(wrapCtx.floats);
    return convergeLineWrap(
      (probeHeights) => runPass(probeHeights, preparedFloatWrap),
      (line) => wrapCtx.lineBoxH(
        line.ascent,
        line.descent,
        line.hasRuby,
        line.intendedSingle,
        line.eastAsian,
        line.gridCountSingle,
      ),
    );
  }
  const { probeHeights, preparedFloatWrap } = passContext;
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[] = [];
  let currentWidth = 0;
  const snapPitchPx = characterGrid?.type === 'snapToChars'
    && characterGrid.characterPitchPt != null
    && characterGrid.characterPitchPt > 0
      ? characterGrid.characterPitchPt * scale
      : null;
  type SnapBlockState = {
    kind: 'latin' | 'complexScript';
    first: LayoutTextSeg;
    last: LayoutTextSeg;
    naturalWidthPx: number;
    allocatedWidthPx: number;
  };
  let snapBlock: SnapBlockState | null = null;
  // Sum of ordinary ONE-space inter-word separators on the current line.
  // A consecutive authored SP sequence is preserved as explicit spacing and
  // contributes no Knuth-Plass shrink budget. Track the pending suffix across
  // segmentation/source-run boundaries so formatting cannot change the result.
  let lineTotalTrailingW = 0;
  let pendingTrailingSpaceCount = 0;
  let pendingTrailingSpaceContribution = 0;
  // Incremental Canvas-vs-Word bias of the text already committed to this line.
  // Candidate checks add only the prospective text, avoiding a hot-loop rescan.
  let lineBiasBudget = 0;
  const lineMeasurementRoutes = new Set<string>();
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let lineDescent = 0;  // px
  let lineIntendedSingle = 0; // px — max intended single-line height on the line
  let lineGridCountSingle = 0; // px — max resolved design height or generic fallback
  let lineVisibleAscent = 0;
  let lineVisibleDescent = 0;
  let lineVisibleIntendedSingle = 0;
  let lineHasVisibleMetrics = false;
  let isFirst = true;
  // Effective width/offset for the current line after float exclusion.
  let lineMaxWidth = maxWidth;
  let lineXOffset = 0;
  let currentLineTopY = wrapCtx?.startPageY ?? 0;

  // Square-only compatibility side-space (px) a CONTENT line needs before it may
  // START beside a square object rather than flow below its band.
  // `word-square-line-start-one-inch` supplies the requirement and tolerance via
  // wordMinLineStartPx(scale),
  // tolerance), INDEPENDENT of a content line's text — the same threshold for a
  // short-token line and a long-word line (a first word that overruns the ≥1-inch
  // gap is force-broken there by the over-long-word char-break below, matching
  // Word's "AFTE"/"R-10" wrapping). This replaced a per-line first-atomic-token
  // width probe that wedged short-token lines into sub-inch gaps and refused
  // ≥1-inch gaps to long-word lines. See issue #676 (fixtures
  // private/sample-19/20/22, pdftotext bbox). Shared by the paint pass and the
  // paginator's two mirror layouts (they call layoutLines with scale 1), so the
  // flow/beside decision agrees across passes.
  //
  // NOTE — this 1-inch rule is the CONTENT-line threshold. A literally-empty
  // paragraph's pilcrow is placed by resolveEmptyMarkTop / flowMarkLine
  // (renderer.ts) against the NARROWER pilcrow-em threshold. An anchorHost-only
  // paragraph still enters layoutLines so its anchor-character metrics size the
  // mark line, but `isParagraphMarkOnlyFlow` selects the same narrow threshold
  // for its first line. `word-empty-mark-float-side-gap` supplies that narrower
  // threshold. #676 over-generalized one inch onto marks; inline
  // content (including a content paragraph's trailing-break final line) keeps
  // the square-only 1-inch rule. Tight/through are governed by their polygon
  // openings (§20.4.2.18/.19), for which there is no corresponding evidence.
  const minLineStartWidth = (): number => wordMinLineStartPx(scale);
  const isParagraphMarkOnlyFlow = segs.length > 0 && segs.every((segment) =>
    ('text' in segment && segment.metricOnly === true)
    || ('imagePath' in segment && Boolean(segment.anchor)),
  );

  // Compute wrap constraints for a new line about to start. Mutates
  // lineXOffset/lineMaxWidth/currentLineTopY. `minWidth` is the smallest clear
  // square side-space the upcoming line must have to START here. Polygon wraps
  // receive MIN_LINE_GAP separately so the compatibility policy cannot erase a
  // through opening explicitly permitted by §20.4.2.18.
  const startLine = (minWidth: number = 0): void => {
    snapBlock = null;
    lineBiasBudget = 0;
    lineMeasurementRoutes.clear();
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    const probeH = probeHeights?.[lines.length];
    // The first pass measures this line without a float window. A later pass
    // resolves it only once that exact line index has an observed line-box
    // height; newly-created lines are likewise measured before they are probed.
    if (probeH === undefined) return;
    const reference = {
      xLeftPt: wrapCtx.referenceXPt ?? wrapCtx.paraX,
      xRightPt: (wrapCtx.referenceXPt ?? wrapCtx.paraX)
        + (wrapCtx.referenceWidthPt ?? maxWidth),
      readingDirection: wrapCtx.readingDirection ?? (baseRtl ? 'rtl' : 'ltr'),
    } as const;
    if (wrapCtx.lineWindow) {
      const win = wrapCtx.lineWindow({
        topYPt: currentLineTopY,
        minimumStartWidthPt: MIN_LINE_GAP,
        squareMinimumStartWidthPt: minWidth,
        probeHeightPt: probeH,
        paragraphXPt: wrapCtx.paraX,
        maximumWidthPt: maxWidth,
        columnXPt: wrapCtx.columnXPt,
        columnWidthPt: wrapCtx.columnWidthPt,
      });
      currentLineTopY = win.topYPt;
      lineXOffset = win.xOffsetPt;
      lineMaxWidth = win.maximumWidthPt;
    } else {
      const win = computePreparedLineFloatWindow(
        currentLineTopY,
        MIN_LINE_GAP,
        probeH,
        wrapCtx.paraX,
        maxWidth,
        preparedFloatWrap ?? prepareFloatWrap(wrapCtx.floats),
        wrapCtx.columnXPt,
        wrapCtx.columnXPt + wrapCtx.columnWidthPt,
        reference,
        minWidth,
      );
      currentLineTopY = win.topY;
      lineXOffset = win.xOffset;
      lineMaxWidth = win.maxWidth;
    }
  };

  // Intrinsic acquisition deliberately disables automatic line wrapping while
  // retaining the real paragraph/anchor width for tab and alignment reference
  // frames. This is a semantic mode, not a synthetic oversized page.
  const availW = () => widthPolicy === 'intrinsic'
    ? Number.POSITIVE_INFINITY
    : lineMaxWidth - (isFirst ? firstIndent : 0);

  // ECMA-376 §17.3.1.37 tab stops in leading-edge px, for the bidi post-pass.
  const bidiCustomStopsPx = baseRtl
    ? tabStops.map((t) => ({ pos: t.pos * scale, alignment: t.alignment, leader: t.leader }))
    : [];
  const bidiIntervalPx = defaultTabPt * scale;

  // Rewrite a finalized bidi line's tab widths (+ leaders) in the VISUAL frame
  // (§17.3.1.6 base RTL). The line's tabs were laid out with provisional width 0
  // by the tab block below (the LTR pen math does not apply under an RTL base);
  // here we place each tab-delimited cell at its mirrored stop. No-op for a line
  // without tabs (LTR paragraphs skip this entirely — `baseRtl` is false).
  const applyBidiTabs = (): void => {
    if (!baseRtl) return;
    if (!currentLine.some((s) => 'isTab' in s)) return;
    // LOGICAL order — the reading-frame walk resolves the Nth tab against the
    // Nth-reachable stop in the logical reading frame. Do not feed the visual
    // sequence here: UAX#9 L2 reverses cells AND tabs together, so a
    // visual-order walk assigns the stops in reverse and paints the leader in
    // the wrong cell gap (the #830 follow-up bug — the TOC underscore leader
    // appeared between the title and the chapter number instead of between the
    // page number and the title). Because the reversal is symmetric, each
    // logical tab's reading-frame gap IS its visual gap, so widths mapped back
    // by logical index tile correctly under the draw loop's visual walk.
    const items: BidiTabItem[] = currentLine.map((s) => ({
      isTab: 'isTab' in s,
      width: s.measuredWidth,
    }));
    for (let tabIndex = 0; tabIndex < currentLine.length; tabIndex += 1) {
      if (!('isTab' in currentLine[tabIndex]!)) continue;
      let cellEnd = tabIndex + 1;
      while (cellEnd < currentLine.length && !('isTab' in currentLine[cellEnd]!)) {
        cellEnd += 1;
      }
      const cell = currentLine.slice(tabIndex + 1, cellEnd);
      const point = decimalAlignmentPoint(cell);
      if (!point) continue;
      const itemIndex = tabIndex + 1 + point.segmentIndex;
      const segment = currentLine[itemIndex]!;
      if ('text' in segment) {
        items[itemIndex]!.decimalOffset = strAdvance(
          segment,
          segment.text.slice(0, point.charOffset),
        );
      }
    }
    // Margin-anchored frame (§17.3.1.37 — stops measure from the TEXT MARGIN):
    // pen 0 = right text margin. Content starts after the leading indent — the
    // line window's RIGHT edge is paraX-relative `lineXOffset + lineMaxWidth`
    // (= maxWidth when no float narrows it), so its margin distance is
    // marginRightPx minus that — plus the first line's first-line indent
    // (which narrows the leading edge under an RTL base, mirroring the draw
    // loop's `effAvailW`). The left text margin sits tabOriginPx past the
    // paragraph box (its trailing indent).
    const startPen = marginRightPx - (lineXOffset + lineMaxWidth) + (isFirst ? firstIndent : 0);
    const leftLimit = marginRightPx + tabOriginPx;
    const res = layoutBidiTabStops(items, bidiCustomStopsPx, startPen, leftLimit, bidiIntervalPx);
    let delta = 0;
    for (let i = 0; i < currentLine.length; i++) {
      const s = currentLine[i];
      if (!('isTab' in s)) continue;
      delta += res[i].width - s.measuredWidth;
      s.measuredWidth = res[i].width;
      (s as LayoutTabSeg).leader = res[i].leader;
    }
    currentWidth += delta;
  };

  let lineHasRuby = false;
  let lineEastAsian = false;
  // Whether any committed token on the current line carries DICTIONARY-SEA
  // (Thai/Lao/Khmer) text — `seaBreaks` marks all SEA segments; the
  // grapheme-fill scripts (Myanmar/Tibetan, #961) are excluded because they use
  // the per-cluster greedy path. `word-dictionary-sea-natural-fit` gates the
  // trailing-space shrink budget for the dictionary scripts.
  let lineHasSea = false;
  const flush = (
    forceHeight?: number,
    brTerminated = false,
    nextStart?: LineBoundary,
  ) => {
    applyBidiTabs();
    // §17.3.2.24 defines `position` relative to surrounding non-positioned
    // text. A line whose every metric-bearing item shares the same inherited
    // position has no differently-positioned peer to pin the resulting extra
    // line height to one side. `word-uniform-run-position-leading` owns Word's
    // observed placement of that surplus above and below the glyphs. Keep mixed
    // lines relative to zero so their authored displacement and ink union
    // remain unchanged. Images/math
    // provide a zero-position reference; tabs do not contribute vertical
    // metrics. The fixed drop-cap path intentionally keeps its paint-only
    // lowering and therefore opts out of this normalization.
    let commonPositionPt: number | undefined;
    let hasPositionReference = false;
    for (const segment of currentLine) {
      if ('isTab' in segment) continue;
      const positionPt = 'text' in segment ? (segment.position ?? 0) : 0;
      if ('text' in segment && segment.positionExtendsLineBox === false) {
        commonPositionPt = 0;
        hasPositionReference = true;
        break;
      }
      if (!hasPositionReference) {
        commonPositionPt = positionPt;
        hasPositionReference = true;
      } else if (commonPositionPt !== positionPt) {
        commonPositionPt = 0;
        break;
      }
    }
    const linePositionReferencePt = hasPositionReference ? (commonPositionPt ?? 0) : 0;
    if (linePositionReferencePt !== 0) {
      for (const segment of currentLine) {
        if ('text' in segment) {
          segment.lineRelativePosition = wordUniformRunPositionPaintPt(
            segment.position ?? 0,
            linePositionReferencePt,
          );
        }
      }
    }
    // §17.3.3.1 — the break is one run among the line's runs: its own size
    // participates in the line height but must not override a taller peer.
    const h = forceHeight !== undefined ? Math.max(lineHeight, forceHeight) : (lineHeight || 10);
    // If the line has no measured content (empty/line-break line), synthesize
    // stable ascent/descent from the effective font size so wrap/baseline math
    // stays consistent with non-empty lines.
    const hasContent = lineAscent > 0 || lineDescent > 0;
    const asc = hasContent ? lineAscent : h * scale * 0.8;
    const desc = hasContent ? lineDescent : h * scale * 0.2;
    const visibleAscent = lineHasVisibleMetrics ? lineVisibleAscent : asc;
    const visibleDescent = lineHasVisibleMetrics ? lineVisibleDescent : desc;
    const visibleIntendedSingle = lineHasVisibleMetrics
      ? lineVisibleIntendedSingle
      : lineIntendedSingle;
    const gridCountSingle = lineGridCountSingle
      || (lineEastAsian ? eastAsianGridCountSinglePx(lineIntendedSingle, h * scale) : asc + desc);
    lines.push({
      segments: currentLine,
      height: h,
      ascent: asc,
      descent: desc,
      visibleAscent,
      visibleDescent,
      visibleIntendedSingle,
      intendedSingle: lineIntendedSingle,
      // Empty/synthetic East Asian lines use the same design-height rule as a
      // text run; their synthesized Canvas box must not reintroduce a
      // scale-dependent cell count.
      gridCountSingle,
      xOffset: lineXOffset,
      availWidth: lineMaxWidth,
      topY: wrapCtx ? currentLineTopY : undefined,
      hasRuby: lineHasRuby,
      eastAsian: lineEastAsian,
      endsWithBreak: brTerminated,
      consumedEnd: nextStart ?? queue[0]?.src ?? endBoundary,
    });
    if (wrapCtx) {
      currentLineTopY += wrapCtx.lineBoxH(
        asc,
        desc,
        lineHasRuby,
        lineIntendedSingle,
        lineEastAsian,
        gridCountSingle,
      );
    }
    currentLine = [];
    currentWidth = 0;
    lineTotalTrailingW = 0;
    pendingTrailingSpaceCount = 0;
    pendingTrailingSpaceContribution = 0;
    lineBiasBudget = 0;
    lineMeasurementRoutes.clear();
    lineHeight = 0;
    lineAscent = 0;
    lineDescent = 0;
    lineIntendedSingle = 0;
    lineGridCountSingle = 0;
    lineVisibleAscent = 0;
    lineVisibleDescent = 0;
    lineVisibleIntendedSingle = 0;
    lineHasVisibleMetrics = false;
    lineHasRuby = false;
    lineEastAsian = false;
    lineHasSea = false;
    isFirst = false;
    startLine(minLineStartWidth());
  };

  const biasBudgetContribution = (s: LayoutTextSeg, text: string = s.text): number =>
    fontAdvanceBiasEm(s.fontFamily)
      * calcEffectiveFontPx(s, scale)
      * charScaleFactor(s)
      * [...text].length;

  // A face allowance is calibrated for one resolved measurement route. Keep
  // route identity separate from whether that route currently has a non-zero
  // profile: two differently resolved Georgia routes are still mixed and must
  // not share one calibrated allowance. Font size is deliberately excluded;
  // small-caps pieces use the same face route at different sizes.
  const measurementRouteIdentity = (s: LayoutTextSeg): string => {
    const weight = s.bold ? 700 : 400;
    const style = s.italic ? 'italic' : 'normal';
    if (s.fontRoute) return `${s.fontRoute.fingerprint}|${weight}|${style}`;
    return `implicit|${buildFont(s.bold, s.italic, 1, s.fontFamily, fontFamilyClasses)}`;
  };

  const noteMeasurementRoute = (
    routes: Set<string>,
    s: LayoutTextSeg,
    text: string = s.text,
  ): void => {
    if (/\S/.test(text)) routes.add(measurementRouteIdentity(s));
  };

  const measurementRouteCountWith = (candidateRoutes: ReadonlySet<string>): number => {
    let count = lineMeasurementRoutes.size;
    for (const route of candidateRoutes) {
      if (!lineMeasurementRoutes.has(route)) count += 1;
    }
    return count;
  };

  const prospectiveSnapAdvance = (s: LayoutTextSeg, naturalWidth: number): number => {
    const kind = snapToCharsClass(s, characterGrid);
    if (!kind || snapPitchPx == null) return naturalWidth;
    if (kind === 'eastAsia') {
      const cells = eastAsianSnapCellCount(s);
      return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx, cells);
    }
    if (snapBlock?.kind === kind) {
      return snapToCharsAllocatedWidthPx(
        snapBlock.naturalWidthPx + naturalWidth,
        kind,
        snapPitchPx,
      ) - snapBlock.allocatedWidthPx;
    }
    return snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
  };

  const addToLine = (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
    trailingSpaceW: number = 0,
  ) => {
    let committedWidth = w;
    if ('text' in s) {
      const kind = snapToCharsClass(s, characterGrid);
      const naturalWidth = s.snapGridNaturalWidthPx ?? w;
      if (kind && snapPitchPx != null) {
        s.snapGridClass = kind;
        s.snapGridNaturalWidthPx = naturalWidth;
        s.snapGridCellPitchPx = snapPitchPx;
        if (kind === 'eastAsia') {
          const cellCount = eastAsianSnapCellCount(s);
          committedWidth = snapToCharsAllocatedWidthPx(
            naturalWidth,
            kind,
            snapPitchPx,
            cellCount,
          );
          s.snapGridLeadingPadPx = 0;
          s.snapGridTrailingPadPx = committedWidth - naturalWidth;
          s.measuredWidth = committedWidth;
          snapBlock = null;
        } else if (snapBlock?.kind === kind) {
          const previousLeading = snapBlock.first.snapGridLeadingPadPx ?? 0;
          const previousTrailing = snapBlock.last.snapGridTrailingPadPx ?? 0;
          const combinedNatural = snapBlock.naturalWidthPx + naturalWidth;
          const combinedAllocated = snapToCharsAllocatedWidthPx(
            combinedNatural,
            kind,
            snapPitchPx,
          );
          const slack = combinedAllocated - combinedNatural;
          const leading = kind === 'latin' ? slack / 2 : 0;
          const trailing = slack - leading;
          snapBlock.first.measuredWidth -= previousLeading;
          snapBlock.first.snapGridLeadingPadPx = leading;
          snapBlock.first.measuredWidth += leading;
          snapBlock.last.measuredWidth -= previousTrailing;
          s.snapGridLeadingPadPx = 0;
          s.snapGridTrailingPadPx = trailing;
          s.measuredWidth = naturalWidth + trailing;
          committedWidth = combinedAllocated - snapBlock.allocatedWidthPx;
          snapBlock = {
            kind,
            first: snapBlock.first,
            last: s,
            naturalWidthPx: combinedNatural,
            allocatedWidthPx: combinedAllocated,
          };
        } else {
          const allocated = snapToCharsAllocatedWidthPx(naturalWidth, kind, snapPitchPx);
          const slack = allocated - naturalWidth;
          const leading = kind === 'latin' ? slack / 2 : 0;
          const trailing = slack - leading;
          s.snapGridLeadingPadPx = leading;
          s.snapGridTrailingPadPx = trailing;
          s.measuredWidth = allocated;
          committedWidth = allocated;
          snapBlock = {
            kind,
            first: s,
            last: s,
            naturalWidthPx: naturalWidth,
            allocatedWidthPx: allocated,
          };
        }
      } else {
        s.snapGridClass = undefined;
        s.snapGridLeadingPadPx = undefined;
        s.snapGridTrailingPadPx = undefined;
        s.snapGridCellPitchPx = undefined;
        s.measuredWidth = w;
        snapBlock = null;
      }
    } else {
      snapBlock = null;
    }
    currentLine.push(s);
    currentWidth += committedWidth;
    if ('text' in s) {
      const trailingSpaceCount = s.text.length - s.text.replace(/ +$/, '').length;
      const spaceOnly = trailingSpaceCount > 0 && trailingSpaceCount === s.text.length;
      if (spaceOnly && pendingTrailingSpaceCount > 0) {
        // UAX #14 LB7 retains the entire SP sequence with the preceding text.
        // A source/style split must not turn an explicit two-space sequence
        // into two separately shrinkable one-space separators.
        lineTotalTrailingW -= pendingTrailingSpaceContribution;
        pendingTrailingSpaceCount += trailingSpaceCount;
        pendingTrailingSpaceContribution = 0;
      } else {
        pendingTrailingSpaceCount = 0;
        pendingTrailingSpaceContribution = 0;
        const preceding = currentLine[currentLine.length - 2];
        const followsVisibleText = preceding !== undefined
          && 'text' in preceding
          && /\S$/u.test(preceding.text);
        if (spaceOnly && followsVisibleText) {
          // Script/font shaping can isolate a normal separator into its own
          // segment. It remains one shrinkable inter-word space until another
          // adjacent SP extends the sequence.
          pendingTrailingSpaceCount = trailingSpaceCount;
          pendingTrailingSpaceContribution = trailingSpaceCount === 1
            ? trailingSpaceW
            : 0;
          lineTotalTrailingW += pendingTrailingSpaceContribution;
        } else if (trailingSpaceCount > 0 && !spaceOnly) {
          pendingTrailingSpaceCount = trailingSpaceCount;
          pendingTrailingSpaceContribution = trailingSpaceCount === 1
            ? trailingSpaceW
            : 0;
          lineTotalTrailingW += pendingTrailingSpaceContribution;
        }
      }
      lineBiasBudget += biasBudgetContribution(s);
      noteMeasurementRoute(lineMeasurementRoutes, s);
    }
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
    if (desc > lineDescent) lineDescent = desc;
    const paintsInlineInk = !('text' in s) || s.metricOnly !== true;
    if (paintsInlineInk) {
      lineHasVisibleMetrics = true;
      if (asc > lineVisibleAscent) lineVisibleAscent = asc;
      if (desc > lineVisibleDescent) lineVisibleDescent = desc;
    }
    // Grid-count height for docGrid cell allocation (§17.6.5). Only East Asian
    // TEXT (and tall inline objects) drives the count — a Latin run keeps its
    // natural height and is NOT cell-rounded, so it must not contribute (its
    // substituted Canvas box would otherwise inflate the count). An EA text run
    // counts from its DESIGN height when tabled, else the deterministic Word FE
    // 1.3em fallback; an image/math object counts its measured box. The line's
    // value is the max.
    let segGridCount = 0;
    if (!('isTab' in s) && !('imagePath' in s) && !('math' in s)) {
      const ts = s as LayoutTextSeg;
      if (ts.ruby) lineHasRuby = true;
      if (ts.seaBreaks !== undefined && isDictionarySeaText(ts.text)) lineHasSea = true;
      const metricEastAsian = ts.metricEastAsian === true || EAST_ASIAN_RE.test(ts.text);
      if (!lineEastAsian && metricEastAsian) lineEastAsian = true;
      // Prefer the selected resource's single-line height. The established
      // family compatibility registry remains a fallback when bytes or native
      // geometry are unavailable; this path adds no new per-family estimate.
      // Small caps (non-super/sub) keep the FULL run size here so the line box
      // follows the run size, not the 2pt-reduced glyphs (§17.3.2.33).
      const intendedEm = ts.smallCaps && !ts.vertAlign ? ts.fontSize * scale : effectiveFontPx(ts);
      // Script hint: eaOnly design heights (Word FE 1.3 × hhea, e.g. Yu Mincho)
      // apply to East Asian segments only — a Latin segment in the same font
      // keeps its Canvas box (issue #1013 / demo sample-1 footnote). Ruby
      // segments are excluded too: a ruby line reserves its MEASURED base +
      // annotation box (sample-5 calibration) and Word's FE height for a
      // ruby-bearing line is unmeasured, so the pre-#1013 metrics stand.
      const segScriptHint = metricEastAsian && !ts.ruby;
      const intended = ts.textBoxLineFloor && ts.ruby
        ? 0
        : Math.max(
            segmentIntendedSingleLinePx(ts, intendedEm, segScriptHint),
            ts.textBoxLineFloor || ts.metricEastAsian === true
              ? segmentEastAsiaFloorSingleLinePx(ts, intendedEm, segScriptHint)
              : 0,
          );
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
      if (paintsInlineInk && intended > lineVisibleIntendedSingle) {
        lineVisibleIntendedSingle = intended;
      }
      // Only East Asian text is cell-rounded. Both branches are scale-linear:
      // resolved resources use parsed metrics; unresolved faces use 1.3em.
      if (segScriptHint) segGridCount = eastAsianGridCountSinglePx(intended, intendedEm);
    } else if (!('isTab' in s)) {
      // Image/math object: a tall inline object sizes the line's cells too.
      segGridCount = asc + desc;
    }
    if (segGridCount > lineGridCountSingle) lineGridCountSingle = segGridCount;
  };

  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);

  // Measure-loop font guard: line wrapping calls measureText / strAdvance many
  // times in a row for the SAME segment (fit search, split prefixes/tails), so
  // the built font string is usually identical to the previous one. Skip the
  // redundant `ctx.font =` in that case. This tracker is written by EVERY font
  // assignment on the measure path (both helpers below route through it), so it
  // always reflects the context's current measure font — no stale skip. The
  // draw-path `ctx.font =` sites are separate and left untouched. `buildFont` is
  // now cheap (normalizeFontFamily is memoized per-doc), so this only elides the
  // setter call itself.
  let lastMeasureFont: string | null = null;
  const setMeasureFont = (font: string): void => {
    if (font !== lastMeasureFont) {
      ctx.font = font;
      lastMeasureFont = font;
    }
  };

  // ECMA-376 §17.3.2.19 `<w:kern>` — set `ctx.fontKerning` to match how the PAINT
  // pass will draw a run, so a kerned run measures exactly as it is drawn
  // (measure==paint). Returns the value to restore afterwards (only when the run
  // opts in). Kerning is enabled only when the run declares `w:kern` and its font
  // size is at or above the threshold (the spec's "smallest font size which shall
  // have its kerning automatically adjusted"). A run that does not opt in leaves
  // `ctx.fontKerning` at its inherited value rather than forcing a document-wide
  // default. Such a default is a separate unsupported policy, not part of this
  // run-level implementation. `setSegKerning` mirrors the paint-side
  // `paintSegKerning` in renderer.ts exactly.
  const setSegKerning = (s: LayoutTextSeg): CanvasFontKerning | null => {
    if (s.kerning == null) return null;
    const prev = ctx.fontKerning;
    ctx.fontKerning = s.fontSize >= s.kerning ? 'normal' : 'none';
    return prev;
  };
  const restoreKerning = (prev: CanvasFontKerning | null): void => {
    if (prev != null) ctx.fontKerning = prev;
  };

  const measureText = (s: LayoutTextSeg, clusterGeometry = false): TextMetrics => {
    if (s.textLayoutService && s.textShapeRequest) {
      const shaped = s.textLayoutService.shape({
        ...s.textShapeRequest,
        text: s.text,
        fontSizePt: effectiveFontPx(s),
        measure: true,
        clusterGeometry,
      });
      if (clusterGeometry) {
        s.shapedClusters = shaped.clusters;
        s.selectedFaceFontBox = {
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
        s.selectedFaceInkBounds = shaped.inkBounds ?? {
          xMinPt: 0,
          xMaxPt: shaped.advancePt,
          ascentPt: shaped.ascentPt,
          descentPt: shaped.descentPt,
        };
      }
      return {
        width: shaped.advancePt,
        actualBoundingBoxAscent: shaped.ascentPt,
        actualBoundingBoxDescent: shaped.descentPt,
        fontBoundingBoxAscent: shaped.ascentPt,
        fontBoundingBoxDescent: shaped.descentPt,
      } as TextMetrics;
    }
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses, s.fontRoute));
    const prevKern = setSegKerning(s);
    const m = ctx.measureText(s.text);
    restoreKerning(prevKern);
    return m;
  };
  // #1014 — extra along-column advance a vertical (tbRl) run needs so a vo=Tr
  // rotate-fallback mark (ー 〜 “” ：) whose substitute font UNDER-REPORTS its
  // advance via measureText keeps its ink inside the ink-sized cell
  // `drawVerticalRun` paints. Added to the natural advance at EVERY site that
  // measures a vertical text seg's advance (the main commit, the tab forced-commit
  // paths, the fitText gap resolver, and the wrap/split look-ahead) so the measured
  // box tracks the drawn cell (measure == draw). 0 for horizontal runs, 縦中横 cells
  // (`!verticalRun`), and every font that does not under-report — byte-identical
  // common path. The run's font must already be selected on `ctx` (the callers
  // select it via measureText / setMeasureFont immediately before).
  const verticalInkExtra = (s: LayoutTextSeg, text: string): number => {
    if (!s.verticalRun) return 0;
    if (!verticalGlyphMeasurement) {
      throw new Error('Vertical glyph measurement capability is required for vertical text');
    }
    // The format-neutral text service may measure on a different adapter and
    // restores its Canvas state. The vertical-feature probe is intentionally
    // paint-context-local, so select the same resolved face explicitly here.
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses, s.fontRoute));
    const prevKern = setSegKerning(s);
    try {
      return verticalGlyphMeasurement.measureRunInkExtra(text);
    } finally {
      restoreKerning(prevKern);
    }
  };

  const endBoundary: LineBoundary = { segIndex: segs.length, charOffset: 0 };
  const sourcedSegs = segs.map((seg, segIndex) => {
    seg.src = { segIndex, charOffset: 0 };
    // Issue #797 / #960 — attach the SEA (Thai/Lao/Khmer) break offsets ONCE per
    // segment (perf: never per line/char). Only for SEA text; non-SEA segments
    // keep `seaBreaks` absent so their wrap path is byte-identical. The set now
    // UNIONS the dictionary word boundaries (#797) with the no-space SEA↔non-SEA
    // script transitions and, for a mixed CJK+SEA run (a `<w:cs/>` run keeps CJK
    // in the same cs segment), the CJK per-character opportunities — so each
    // script keeps its own break rule inside one contiguous segment (#960). The
    // layout kinsoku set (§17.15.1.58–.60) drops positions that would orphan a
    // forbidden char at a line head/tail, replacing the CJK path's retract.
    if ('text' in seg && containsSeaScript(seg.text)) {
      const protectedOffsets = protectedNoBreakOffsets(seg);
      seg.seaBreaks = seaMixedBreakOffsets(seg.text, { cjk: true, kinsoku })
        .filter((offset) => !protectedOffsets.has(offset));
    }
    return seg;
  });
  let queue: LayoutSeg[];
  if (!startBoundary) {
    queue = sourcedSegs;
  } else if (startBoundary.segIndex >= sourcedSegs.length) {
    queue = [];
  } else {
    const first = sourcedSegs[startBoundary.segIndex];
    if (startBoundary.charOffset > 0) {
      if (!('text' in first) || startBoundary.charOffset > first.text.length) {
        queue = [];
      } else {
        const text = first.text.slice(startBoundary.charOffset);
        queue = text
          ? [
              {
                ...first,
                text,
                measuredWidth: 0,
                src: { ...startBoundary },
                // A retained resume boundary has already consumed the source
                // seam. Carrying either marker would invent new ownership at
                // the start of this suffix.
                joinPrev: undefined,
                hardJoinPrev: undefined,
                ...slicedTextMetadata(first, startBoundary.charOffset, first.text.length),
                // Rebase the SEA break offsets onto the resumed (sliced) text so
                // a paginated Thai paragraph still breaks at word boundaries.
                seaBreaks: rebaseSeaBreaks(first.seaBreaks, startBoundary.charOffset),
              },
              ...sourcedSegs.slice(startBoundary.segIndex + 1),
            ]
          : sourcedSegs.slice(startBoundary.segIndex + 1);
      }
    } else {
      queue = sourcedSegs.slice(startBoundary.segIndex);
    }
  }

  // Mark the paragraph-final U+3000 suffix once. A backwards pass avoids the
  // O(N²) suffix rescans that would result from queue.every/reduce per segment.
  let paragraphFinalIdeographicSpaceCount = 0;
  let paragraphFinalIdeographicSpaceTailStartIndex = -1;
  const markedParagraphFinalTail: Array<Readonly<{
    index: number;
    segment: LayoutTextSeg;
  }>> = [];
  for (let index = queue.length - 1; index >= 0; index -= 1) {
    const candidate = queue[index];
    if (!candidate || !('text' in candidate) || candidate.text.length === 0) break;
    // `fitText` (§17.3.2.14) and tate-chu-yoko (§17.3.2.10) are indivisible
    // layout cells. Ruby owns one base/guide pair. Paragraph-final whitespace
    // may affect how those cells measure, but must never split or clone them.
    if (
      candidate.fitTextRegionIndex !== undefined
      || candidate.tateChuYoko === true
      || candidate.ruby !== undefined
    ) {
      // buildSegments can split an atomic source run before its U+3000 tail
      // (ruby is retained only on the first emitted segment). Undo any markers
      // already assigned to trailing pieces from that same authored run.
      if (candidate.sourceRunIndex !== undefined) {
        for (let markedIndex = markedParagraphFinalTail.length - 1; markedIndex >= 0; markedIndex -= 1) {
          const marked = markedParagraphFinalTail[markedIndex];
          if (marked.segment.sourceRunIndex !== candidate.sourceRunIndex) continue;
          marked.segment.paragraphFinalIdeographicSpaceTail = undefined;
          marked.segment.paragraphFinalIdeographicSpaceLocalCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceCount = undefined;
          marked.segment.paragraphFinalIdeographicSpaceTailStart = undefined;
          markedParagraphFinalTail.splice(markedIndex, 1);
        }
        paragraphFinalIdeographicSpaceTailStartIndex =
          markedParagraphFinalTail.at(-1)?.index ?? -1;
      }
      break;
    }
    const trailingSpaces = /^\u3000+$/u.test(candidate.text);
    const visibleWithTrailingSpaces = /[^\u3000]\u3000+$/u.test(candidate.text);
    if (!trailingSpaces && !visibleWithTrailingSpaces) break;
    const localTrailingCount = trailingSpaces
      ? [...candidate.text].length
      : [...candidate.text].reverse().findIndex((character) => character !== '\u3000');
    paragraphFinalIdeographicSpaceCount += localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceTail = true;
    candidate.paragraphFinalIdeographicSpaceLocalCount = localTrailingCount;
    candidate.paragraphFinalIdeographicSpaceCount = paragraphFinalIdeographicSpaceCount;
    paragraphFinalIdeographicSpaceTailStartIndex = index;
    markedParagraphFinalTail.push({ index, segment: candidate });
    if (visibleWithTrailingSpaces) break;
  }
  if (paragraphFinalIdeographicSpaceTailStartIndex >= 0) {
    const start = queue[paragraphFinalIdeographicSpaceTailStartIndex];
    if (start && 'text' in start) start.paragraphFinalIdeographicSpaceTailStart = true;
  }

  // Resolve §17.3.2.14 from RAW natural advances at this exact layout scale.
  // The resulting per-gap is folded into segAdvanceWidth below, so the line
  // breaker and paint pen use one width authority. Cached w:spacing is ignored.
  // #1014 — the natural width includes the vo=Tr ink deficit so the resolved gap
  // (target − natural)/n, plus the ink-grown cell the paint draws, still sums to
  // the fitText target (measure == paint); 0 for non-under-reporting runs.
  resolveFitTextSegments(
    queue.filter((seg): seg is LayoutTextSeg => 'text' in seg),
    scale,
    (segment) => measureText(segment).width + verticalInkExtra(segment, segment.text),
  );

  // The segment's laid-out ADVANCE (= its measuredWidth): natural width plus the
  // character-grid delta, the §17.3.2.43 horizontal glyph scale (w:w) and the
  // §17.3.2.35 character-spacing pitch (w:spacing). This is the SINGLE source of
  // truth shared with the draw paths (segAdvanceWidth) — every line-break / fit /
  // tab measurement uses it so line wrapping packs the grid's char count and the
  // box matches what is drawn (measure==paint). `kerning` (§17.3.2.19) is applied
  // via `ctx.fontKerning` inside `withSegKerning`, wrapping the measureText call.
  // The #1014 vo=Tr ink deficit (`verticalInkExtra`, defined above) is folded into
  // the natural width so measure == paint on an under-reporting vertical run.
  const segNaturalAdvance = (s: LayoutTextSeg): number =>
    segAdvanceWidth(s, measureText(s).width + verticalInkExtra(s, s.text), characterGrid, scale);
  const standaloneSnapAdvance = (s: LayoutTextSeg, naturalWidth: number): number => {
    const kind = snapToCharsClass(s, characterGrid);
    if (!kind || snapPitchPx == null || s.text.length === 0) return naturalWidth;
    return snapToCharsAllocatedWidthPx(
      naturalWidth,
      kind,
      snapPitchPx,
      kind === 'eastAsia' ? eastAsianSnapCellCount(s) : 1,
    );
  };
  const segAdvance = (s: LayoutTextSeg): number =>
    standaloneSnapAdvance(s, segNaturalAdvance(s));
  // Grid advance of an arbitrary substring under a segment's font (for split
  // prefixes/tails). Selects the font (and the run's kerning state), then applies
  // the same width model as a whole segment BUT with the substring's own
  // text/length so char-spacing scales with the piece — the split-prefix vs
  // whole-segment advances must agree.
  const strNaturalAdvance = (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression = false,
  ): number => {
    const start = retainTrailingPunctuationCompression
      ? s.text.length - text.length
      : 0;
    const measuredSegment = {
      ...s,
      text,
      punctuationCompressions: slicedPunctuationCompressions(
        s,
        Math.max(0, start),
        Math.max(0, start) + text.length,
      ),
    };
    if (s.textLayoutService && s.textShapeRequest) {
      const shaped = s.textLayoutService.shape({
        ...s.textShapeRequest,
        text,
        fontSizePt: effectiveFontPx(s),
        measure: true,
        clusterGeometry: false,
      });
      return segAdvanceWidth(
        measuredSegment,
        shaped.advancePt + verticalInkExtra(s, text),
        characterGrid,
        scale,
      );
    }
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses, s.fontRoute));
    const prevKern = setSegKerning(s);
    const natural = ctx.measureText(text).width;
    restoreKerning(prevKern);
    return segAdvanceWidth(
      measuredSegment,
      natural + verticalInkExtra(s, text),
      characterGrid,
      scale,
    );
  };
  const eastAsianSnapCellCount = (s: LayoutTextSeg): number => {
    if (snapPitchPx == null) return 1;
    if (s.textLayoutService && s.textShapeRequest && !s.shapedClusters) {
      measureText(s, true);
    }
    const shapedClusters = s.shapedClusters?.length
      ? s.shapedClusters
      : null;
    const boundaries = shapedClusters == null
      ? [...new Set([
          0,
          ...graphemeClusterOffsets(s.text),
          s.text.length,
        ])].sort((a, b) => a - b)
      : null;
    const ranges = shapedClusters?.map((cluster) => ({
      start: cluster.range.start,
      end: cluster.range.end,
      advancePx: cluster.advancePt,
    })) ?? boundaries!.slice(0, -1).map((start, index) => ({
      start,
      end: boundaries![index + 1]!,
      advancePx: undefined,
    }));
    let cells = 0;
    for (const range of ranges) {
      const { start, end } = range;
      if (end <= start) continue;
      const text = s.text.slice(start, end);
      const measuredSegment = {
        ...s,
        text,
        punctuationCompressions: slicedPunctuationCompressions(s, start, end),
      };
      let naturalAdvancePx: number;
      if (range.advancePx != null) {
        naturalAdvancePx = segAdvanceWidth(
          measuredSegment,
          range.advancePx + verticalInkExtra(s, text),
          characterGrid,
          scale,
        );
      } else {
        setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses, s.fontRoute));
        const previousKerning = setSegKerning(s);
        const naturalWidthPx = ctx.measureText(text).width;
        restoreKerning(previousKerning);
        naturalAdvancePx = segAdvanceWidth(
          measuredSegment,
          naturalWidthPx + verticalInkExtra(s, text),
          characterGrid,
          scale,
        );
      }
      cells += wordSnapToCharsEastAsianCellCount(naturalAdvancePx, snapPitchPx);
    }
    return Math.max(1, cells);
  };
  const strAdvance = (
    s: LayoutTextSeg,
    text: string,
    retainTrailingPunctuationCompression = false,
  ): number => {
    const candidate = {
      ...s,
      text,
      shapedClusters: text === s.text ? s.shapedClusters : undefined,
    };
    return standaloneSnapAdvance(
      candidate,
      strNaturalAdvance(s, text, retainTrailingPunctuationCompression),
    );
  };

  /** Measure one text segment's canonical advance and vertical contribution.
   * Every path that commits a complete text segment to a line must use this
   * authority so font fallback, small-caps, position, ruby and grid metrics do
   * not diverge at internal segment seams. */
  const textSegmentBox = (s: LayoutTextSeg): Readonly<{
    width: number;
    height: number;
    ascent: number;
    descent: number;
  }> => {
    const measured = measureText(s, snapToCharsClass(s, characterGrid) === 'eastAsia');
    const width = segAdvanceWidth(
      s,
      measured.width + verticalInkExtra(s, s.text),
      characterGrid,
      scale,
    );
    s.snapGridNaturalWidthPx = width;

    const fullPx = s.fontSize * scale;
    let metricMeasurement = measured;
    let metricEmPx = effectiveFontPx(s);
    if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
      if (s.textLayoutService && s.textShapeRequest) {
        const shaped = s.textLayoutService.shape({
          ...s.textShapeRequest,
          text: s.text || 'X',
          fontSizePt: fullPx,
          measure: true,
          clusterGeometry: false,
        });
        metricMeasurement = {
          width: shaped.advancePt,
          actualBoundingBoxAscent: shaped.ascentPt,
          actualBoundingBoxDescent: shaped.descentPt,
          fontBoundingBoxAscent: shaped.ascentPt,
          fontBoundingBoxDescent: shaped.descentPt,
        } as TextMetrics;
      } else {
        const previousFont = ctx.font;
        ctx.font = buildFont(
          s.bold,
          s.italic,
          fullPx,
          s.fontFamily,
          fontFamilyClasses,
          s.fontRoute,
        );
        metricMeasurement = ctx.measureText(s.text || 'X');
        ctx.font = previousFont;
      }
      metricEmPx = fullPx;
    }

    const corrected = correctedLineMetrics(
      metricMeasurement,
      s.fontFamily,
      fullPx,
      metricEmPx,
      (s.metricEastAsian === true || EAST_ASIAN_RE.test(s.text)) && !s.ruby,
    );
    let ascent = corrected.ascent;
    let descent = corrected.descent;
    if (s.positionExtendsLineBox !== false) {
      const positionPx = (s.position ?? 0) * scale;
      if (positionPx > 0) ascent += positionPx;
      else if (positionPx < 0) descent -= positionPx;
    }
    if (s.ruby && (!s.textBoxLineFloor || s.textBoxVertical)) {
      ascent += rubyAscentReservePx(
        s.ruby.fontSizePt,
        s.ruby.hpsRaisePt,
        scale,
        s,
        ctx,
        fontFamilyClasses,
      );
    }
    return { width, height: s.fontSize, ascent, descent };
  };

  /** Continue the existing U+3000 line-end hanging rule across an internal
   * width-balance segment seam. The split space keeps its own font/grid advance
   * for measure == paint. UAX #14 classifies U+3000 as BA, so an internal seam
   * before it must not become an authored break opportunity. Restrict
   * consumption to the same authored run; an actual source boundary remains
   * independently modeled. */
  const appendQueuedIdeographicSpaceSegment = (
    source: LayoutTextSeg,
  ): void => {
    if (
      /\s$/u.test(source.text)
      || source.ruby !== undefined
      || source.tateChuYoko === true
      || source.fitTextRegionIndex !== undefined
    ) return;
    const follower = queue[0];
    if (
      !follower
      || !('text' in follower)
      || follower.joinPrev !== true
      || follower.text.length === 0
      || [...follower.text].some((character) => character !== '\u3000')
    ) return;
    queue.shift();
    const hangingCount = wordIdeographicSpaceLineEndAllowanceCount(
      hasEastAsianVisiblePredecessor(source.text),
      follower.paragraphFinalIdeographicSpaceCount ?? [...follower.text].length,
    );
    if (hangingCount === 0) {
      queue.unshift(follower);
      return;
    }
    const hangingText = follower.text.slice(0, hangingCount);
    const hangingSegment: LayoutTextSeg = {
      ...follower,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: hangingText,
      measuredWidth: 0,
      ...slicedTextMetadata(follower, 0, hangingText.length),
    };
    const followerBox = textSegmentBox(hangingSegment);
    hangingSegment.measuredWidth = followerBox.width;
    addToLine(
      hangingSegment,
      followerBox.width,
      followerBox.height,
      followerBox.ascent,
      followerBox.descent,
    );
    const remainder = follower.text.slice(hangingText.length);
    if (remainder.length > 0) {
      queue.unshift({
        ...follower,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: remainder,
        measuredWidth: 0,
        joinPrev: undefined,
        hardJoinPrev: undefined,
        ...slicedTextMetadata(follower, hangingText.length, follower.text.length),
        src: follower.src
          ? {
              segIndex: follower.src.segIndex,
              charOffset: follower.src.charOffset + hangingText.length,
            }
          : undefined,
      });
    }
  };

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => {
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('imagePath' in q) return q.widthPt * scale;
    if ('math' in q) return q.measuredWidth || 0;
    if ('lineBreak' in q) return 0;
    return segAdvance(q);
  };

  /** Resolve the registered decimal alignment point independently of run/style
   * seams so both LTR and mirrored bidi tab paths consume one source boundary. */
  const decimalAlignmentPoint = (
    segments: readonly LayoutSeg[],
  ): Readonly<{ segmentIndex: number; charOffset: number }> | null => {
    for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
      const segment = segments[segmentIndex]!;
      if (!('text' in segment)) continue;
      const separator = segment.text.indexOf('.');
      if (separator >= 0) return { segmentIndex, charOffset: separator };
    }

    let lastDigit: Readonly<{ segmentIndex: number; charOffset: number }> | null = null;
    let inFirstNumber = false;
    for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
      const segment = segments[segmentIndex]!;
      if (!('text' in segment)) {
        if (inFirstNumber) return lastDigit;
        continue;
      }
      let charOffset = 0;
      for (const scalar of segment.text) {
        charOffset += scalar.length;
        if (/\p{Decimal_Number}/u.test(scalar)) {
          inFirstNumber = true;
          lastDigit = { segmentIndex, charOffset };
        } else if (inFirstNumber) {
          return lastDigit;
        }
      }
    }
    return lastDigit;
  };

  const decimalAlignmentPrefixWidth = (segments: readonly LayoutSeg[]): number | undefined => {
    const point = decimalAlignmentPoint(segments);
    if (!point) return undefined;
    let width = 0;
    for (let index = 0; index < point.segmentIndex; index += 1) {
      width += tabFollowWidth(segments[index]!);
    }
    const segment = segments[point.segmentIndex]!;
    if (!('text' in segment)) return width;
    return width + strAdvance(segment, segment.text.slice(0, point.charOffset));
  };

  const tabFollowingMetrics = (): Readonly<{
    totalWidth: number;
    decimalPrefixWidth?: number;
  }> => {
    const following: LayoutSeg[] = [];
    let totalWidth = 0;
    for (const q of queue) {
      if ('isTab' in q || 'lineBreak' in q) break;
      following.push(q);
      totalWidth += tabFollowWidth(q);
    }
    const decimalPrefixWidth = decimalAlignmentPrefixWidth(following);
    return decimalPrefixWidth === undefined
      ? { totalWidth }
      : { totalWidth, decimalPrefixWidth };
  };

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph, the new line is empty but still occupies one line
  // height. Track the trailing break so it can be flushed after the loop.
  let trailingBreakFontSize: number | null = null;

  // Establish the first line's wrap window now that the content queue exists.
  startLine(
    isParagraphMarkOnlyFlow
      ? (wrapCtx?.paragraphMarkLineStartWidth ?? minLineStartWidth())
      : minLineStartWidth(),
  );

  /**
   * Return the widest grapheme-safe UTF-16 prefix that fits an emergency
   * break band. This is the single authority used whether the overlong token
   * starts on an empty line or consumes the useful remainder of the current
   * line. Keeping both cases here prevents measurement and retained paint
   * partitions from drifting apart.
   */
  const emergencyTextSplit = (
    segment: LayoutTextSeg,
    available: number,
    forceAtLeastOne = true,
  ): number => {
    const protectedOffsets = protectedNoBreakOffsets(segment);
    const graphemeOffsets = [
      0,
      ...graphemeClusterOffsets(segment.text),
      segment.text.length,
    ].filter((offset, index, all) => all.indexOf(offset) === index);
    let split = 0;
    if (available > 0) {
      const monotoneAllocation = charSpacingDeltaPx(segment, scale) >= 0
        && snapToCharsClass(segment, characterGrid) !== 'latin';
      if (monotoneAllocation) {
        setMeasureFont(buildFont(
          segment.bold,
          segment.italic,
          effectiveFontPx(segment),
          segment.fontFamily,
          fontFamilyClasses,
          segment.fontRoute,
        ));
        const prevKern = setSegKerning(segment);
        try {
          const fitted = fitCJKPrefix(
            ctx,
            segment.text,
            available,
            segmentCharacterGridDeltaPx(segment, characterGrid, scale),
            charScaleFactor(segment),
            charSpacingDeltaPx(segment, scale),
            segment.verticalRun === true,
            verticalGlyphMeasurement,
            (prefix) => strAdvance(segment, prefix),
          ).length;
          split = graphemeOffsets
            .filter((offset) => offset <= fitted && !protectedOffsets.has(offset))
            .at(-1) ?? 0;
        } finally {
          restoreKerning(prevKern);
        }
      } else {
        // Signed spacing and a Latin snap block can make prefix advances
        // non-monotone. Evaluate every legal retained candidate against the
        // exact prospective line block rather than binary-searching a
        // standalone approximation.
        for (const offset of graphemeOffsets) {
          if (offset <= 0 || protectedOffsets.has(offset)) continue;
          const natural = strNaturalAdvance(segment, segment.text.slice(0, offset));
          if (prospectiveSnapAdvance(segment, natural) <= available + 1e-9) split = offset;
        }
      }
    }
    if (split <= 0 && forceAtLeastOne) {
      split = graphemeOffsets.find((offset) => offset > 0 && !protectedOffsets.has(offset))
        ?? segment.text.length;
    }
    // Preserve the existing JLReq/Word line-end hanging rule after switching
    // the emergency splitter from code-point indexes to UTF-16 grapheme offsets.
    while (segment.text.startsWith('\u3000', split)) split += 1;
    return split;
  };

  /** Select the last semantic URL candidate that fits the prospective band.
   * Every candidate is evaluated independently: signed character spacing can
   * make prefix advances non-monotone, and snapToChars must include the current
   * line's active script block rather than treating the prefix in isolation. */
  const externalLinkSyntaxSplit = (
    segment: LayoutTextSeg,
    available: number,
  ): number => {
    if (!(available > 0) || !segment.externalLinkBreakOffsets?.length) return 0;
    let selected = 0;
    for (const offset of segment.externalLinkBreakOffsets) {
      if (offset <= 0 || offset >= segment.text.length) continue;
      const naturalAdvance = strNaturalAdvance(segment, segment.text.slice(0, offset));
      const prospectiveAdvance = prospectiveSnapAdvance(segment, naturalAdvance);
      if (prospectiveAdvance <= available + 1e-9) selected = offset;
    }
    return selected;
  };

  const queueEmergencyTail = (segment: LayoutTextSeg, split: number): void => {
    queue.unshift({
      ...segment,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: segment.text.slice(split),
      ...slicedTextMetadata(segment, split, segment.text.length),
      seaBreaks: rebaseSeaBreaks(segment.seaBreaks, split),
      measuredWidth: 0,
      // The emergency split itself is now the legal line boundary. A source-
      // boundary glue marker protects only the first retained prefix; carrying
      // it onto the tail would make the next line overflow again.
      joinPrev: undefined,
      hardJoinPrev: undefined,
      src: {
        segIndex: segment.src!.segIndex,
        charOffset: segment.src!.charOffset + split,
      },
    });
  };

  type CrossRunKinsokuRetraction =
    | { readonly kind: 'none' }
    | { readonly kind: 'blocked' }
    | { readonly kind: 'retracted'; readonly tail: LayoutTextSeg };

  /**
   * Move a legal suffix of the current line ahead of a segment whose first
   * glyph is forbidden at line start. This is the single cross-run 追い出し
   * authority for both the CJK and SEA overflow paths.
   *
   * A source seam marked by `hardJoinPrev` is indivisible: moving the complete
   * following segment would strand its owner on the previous line. Likewise,
   * a split at either edge of an authored no-break range is not legal. When
   * either constraint blocks retraction, the caller must keep the forbidden
   * leader on the current line instead of weakening the authored constraint.
   */
  const retractCurrentLineForLeadingKinsoku = (
    next: LayoutTextSeg,
  ): CrossRunKinsokuRetraction => {
    const firstCp = next.text.codePointAt(0);
    const lastSeg = currentLine[currentLine.length - 1];
    if (
      firstCp === undefined
      || !kinsoku.lineStartForbidden.has(firstCp)
      || lastSeg === undefined
      || !('text' in lastSeg)
    ) {
      return { kind: 'none' };
    }

    const lastText = lastSeg as LayoutTextSeg;
    const chars = [...lastText.text];
    const minKeep = currentLine.length > 1 ? 0 : 1;
    const retractCount = crossRunKinsokuRetract(chars, kinsoku, minKeep);
    if (retractCount <= 0) return { kind: 'none' };

    const headText = chars.slice(0, chars.length - retractCount).join('');
    const split = headText.length;
    // Moving the whole segment would cut the hard source seam immediately
    // before it. A protected no-break edge is equally indivisible.
    if (
      (split === 0 && lastText.hardJoinPrev === true)
      || protectedNoBreakOffsets(lastText).has(split)
    ) {
      return { kind: 'blocked' };
    }

    const tailText = lastText.text.slice(split);
    const tail: LayoutTextSeg = {
      ...lastText,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: tailText,
      ...slicedTextMetadata(lastText, split, lastText.text.length),
      measuredWidth: strAdvance(lastText, tailText, true),
      // The retraction creates a real line boundary. The old source seam was
      // either retained in `headText` or was soft; it must not be projected
      // onto this newly-created suffix.
      joinPrev: undefined,
      hardJoinPrev: undefined,
      src: {
        segIndex: lastText.src!.segIndex,
        charOffset: lastText.src!.charOffset + split,
      },
      seaBreaks: rebaseSeaBreaks(lastText.seaBreaks, split),
    };

    if (headText) {
      const headW = strAdvance(lastText, headText);
      currentWidth -= lastText.measuredWidth - headW;
      currentLine[currentLine.length - 1] = {
        ...lastText,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: headText,
        measuredWidth: headW,
        ...slicedTextMetadata(lastText, 0, split),
      };
    } else {
      currentWidth -= lastText.measuredWidth;
      currentLine.pop();
    }
    return { kind: 'retracted', tail };
  };

  /** Keep one otherwise-forbidden line-start grapheme with the current line.
   * This is the only legal fallback when cross-run retraction would split an
   * authored hard/no-break group. Reprocessing the tail repeats the rule for a
   * sequence of forbidden leaders while guaranteeing grapheme-safe progress. */
  const keepLeadingKinsokuWithCurrentLine = (
    segment: LayoutTextSeg,
    h: number,
    asc: number,
    desc: number,
  ): boolean => {
    const firstEnd = graphemeClusterOffsets(segment.text)[0] ?? segment.text.length;
    if (firstEnd <= 0) return false;
    const prefix = segment.text.slice(0, firstEnd);
    const prefixWidth = strNaturalAdvance(segment, prefix);
    addToLine({
      ...segment,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: prefix,
      measuredWidth: prefixWidth,
      ...slicedTextMetadata(segment, 0, firstEnd),
    }, prefixWidth, h, asc, desc);
    if (firstEnd < segment.text.length) queueEmergencyTail(segment, firstEnd);
    return true;
  };

  while (queue.length > 0) {
    const seg = queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      // The line being flushed ends at a MANUAL break (§17.3.3.1) — mark it so a
      // justified paragraph left-aligns it like its final line (§17.18.44).
      flush(seg.fontSize, true);
      trailingBreakFontSize = seg.fontSize;
      continue;
    }
    trailingBreakFontSize = null;

    // ── Tab segment ──────────────────────────────────────
    if ('isTab' in seg) {
      // ── ECMA-376 §17.3.1.6 base-RTL ordinary tab ─────────────────────────
      // The LTR pen math below resolves stops in LOGICAL order, which mis-places
      // a bidi paragraph's tab-delimited cells (they reorder visually — see
      // `layoutBidiTabStops`). Add the tab with a PROVISIONAL width of 0 and do
      // NOT wrap on it; the per-line post-pass (`applyBidiTabs`, run in `flush`)
      // recomputes every tab width in the visual frame once the line's content
      // is known. A `<w:ptab>` (absolute-position tab) keeps the LTR path for
      // now (no bidi ptab fixture; its own NOTE flags the gap).
      if (baseRtl && !seg.ptab) {
        seg.measuredWidth = 0;
        addToLine(seg, 0, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
        continue;
      }

      // Absolute position on the line measured from paraX (line origin for continuation lines)
      const absFromParaX = currentWidth + (isFirst ? firstIndent : 0);

      // ── ECMA-376 §17.3.3.23 absolute-position tab (<w:ptab>) ──────────────
      // A ptab ignores the paragraph's custom tab stops and the default-tab
      // interval; it advances to a fixed position on the line derived from its
      // `alignment` (§17.18.71) and `relativeTo` (§17.18.73). The `alignment`
      // ALSO governs how the text after the ptab aligns to that position (left /
      // centered / right). All coordinates below are paraX-relative px.
      //
      // NOTE: the ptab target is resolved in LOGICAL (LTR) coordinates — this
      // block runs before the per-line bidi reorder pass, so it has no notion
      // of the paragraph's base direction. Interaction with bidi mirroring in
      // an RTL paragraph (where "left"/"right" alignment and the box edges
      // ought to mirror) is unverified; the primary use case (an LTR footer's
      // centered/right-aligned PAGE field) is correct.
      if (seg.ptab) {
        seg.resolvedAlignment = seg.ptab.alignment;
        // Reference box: "indent" ⇒ the paragraph content box [0, maxWidth];
        // "margin" ⇒ the text-margin box [-tabOriginPx, marginRightPx].
        const boxLeft = seg.ptab.relativeTo === 'indent' ? 0 : -tabOriginPx;
        const boxRight = seg.ptab.relativeTo === 'indent' ? maxWidth : marginRightPx;
        const target =
          seg.ptab.alignment === 'left'
            ? boxLeft
            : seg.ptab.alignment === 'center'
              ? (boxLeft + boxRight) / 2
              : boxRight;
        // Width of the content that trails the ptab up to the next tab / line end
        // — needed to right-/center-align it against `target` (the trailing text
        // is what aligns to the stop, §17.18.71).
        let followW = 0;
        for (const q of queue) {
          if ('isTab' in q || 'lineBreak' in q) break;
          followW += tabFollowWidth(q);
        }
        const frac = seg.ptab.alignment === 'center' ? 0.5 : seg.ptab.alignment === 'right' ? 1 : 0;
        let tabW = target - absFromParaX - followW * frac;
        // §17.3.3.23: "If the alignment location … cannot be found on the current
        // line, because the starting location is past that point, then the tab …
        // shall advance to that location on the next available line." So when the
        // pen already sits at/after the target, wrap the ptab (and its trailing
        // content) to a fresh line — unless the line is empty (nowhere to wrap).
        if (tabW <= 0) {
          if (currentLine.length > 0) {
            flush(undefined, false, seg.src);
            queue.unshift(seg);
            continue;
          }
          // Empty line: cannot advance backwards; contribute no width but keep the
          // segment so the line-height reflects the ptab's font.
          tabW = 0;
        }
        seg.measuredWidth = tabW;
        addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
        // Commit the trailing content onto this line without a wrap re-check, so
        // it sits exactly at the aligned position (mirrors the custom right/center
        // tab path below).
        if (seg.ptab.alignment !== 'left') {
          while (queue.length > 0) {
            const q = queue[0];
            if ('isTab' in q || 'lineBreak' in q) break;
            queue.shift();
            if ('imagePath' in q) {
              const w = q.widthPt * scale;
              q.measuredWidth = w;
              addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
            } else if ('math' in q) {
              addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
            } else {
              const m = measureText(q);
              // #1014 — fold the vo=Tr ink deficit into the committed advance too.
              const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), characterGrid, scale);
              q.measuredWidth = w;
              const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
              const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
              addToLine(q, w, q.fontSize, asc, desc);
            }
          }
        }
        continue;
      }
      // ECMA-376 §17.3.1.37 / §17.15.1.25 — resolve the next stop in TEXT-MARGIN
      // coordinates (the same origin as custom stops): the current pen position is
      // `absFromParaX + tabOriginPx`, custom stops are `pos * scale`, and the
      // automatic grid interval is `defaultTabPt * scale`. Mixing paraX and margin
      // coordinates is what diverged leading-tab rows from labeled ones; computing
      // both in margin space and converting back keeps them aligned.
      const curMarginPx = absFromParaX + tabOriginPx;
      const customStopsPx = tabStops.map((t) => ({ pos: t.pos * scale, alignment: t.alignment, leader: t.leader }));
      const stop = nextTabStop(curMarginPx, customStopsPx, defaultTabPt * scale);
      seg.resolvedAlignment = stop?.alignment ?? 'left';
      // Convert the chosen margin-space stop back to paraX-relative px.
      const stopParaX = stop ? stop.pos - tabOriginPx : absFromParaX;
      // Right/center/decimal tab: place the tab + its trailing content (up to the next
      // tab / line end) so the content ends at / centers on the stop, and commit that
      // content directly so the normal wrap check doesn't push it past the stop
      // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
      // Automatic stops returned by nextTabStop are left-aligned, so they fall
      // through to the left-tab path below.
      const alignmentRole = stop ? tabAlignmentRole(stop.alignment) : 'leading';
      if (stop && alignmentRole !== 'leading') {
        const stopX = stopParaX;
        seg.leader = stop.leader;
        const following = tabFollowingMetrics();
        const alignmentWidth = alignmentRole === 'center'
          ? following.totalWidth / 2
          : alignmentRole === 'decimal'
            ? following.decimalPrefixWidth ?? following.totalWidth
            : following.totalWidth;
        let tabW = stopX - absFromParaX - alignmentWidth;
        if (tabW <= 0) tabW = 0;
        seg.measuredWidth = tabW;
        addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
        // Commit the trailing content onto this line without a wrap re-check.
        while (queue.length > 0) {
          const q = queue[0];
          if ('isTab' in q || 'lineBreak' in q) break;
          queue.shift();
          if ('imagePath' in q) {
            const w = q.widthPt * scale;
            q.measuredWidth = w;
            addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
          } else if ('math' in q) {
            addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
          } else {
            const m = measureText(q);
            // #1014 — fold the vo=Tr ink deficit into the committed advance too.
            const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), characterGrid, scale);
            q.measuredWidth = w;
            const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
            const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
            addToLine(q, w, q.fontSize, asc, desc);
          }
        }
        continue;
      }

      // Left-aligned tab (custom 'left'/'bar'/'clear' or an automatic stop): the
      // pen moves to the stop's paraX. nextTabStop already applied the §17.15.1.25
      // "after all custom stops" automatic grid, so there is no separate fallback.
      let tabWidth = stopParaX - absFromParaX;
      if (stop) seg.leader = stop.leader;
      // Clamp to avoid negative widths; if tab would overflow the line, wrap instead
      if (tabWidth <= 0) {
        flush(undefined, false, seg.src);
        queue.unshift(seg);
        continue;
      }
      if (currentWidth + tabWidth > availW() && currentLine.length > 0) {
        flush(undefined, false, seg.src);
        queue.unshift(seg);
        continue;
      }
      seg.measuredWidth = tabWidth;
      addToLine(seg, tabWidth, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
      continue;
    }

    // ── Image segment ────────────────────────────────────
    if ('imagePath' in seg) {
      if (seg.anchor) { seg.measuredWidth = 0; continue; }
      const w = seg.widthPt * scale;
      const h = seg.heightPt;
      const asc = seg.heightPt * scale;
      seg.measuredWidth = w;
      if (currentLine.length > 0 && currentWidth + w > availW()) {
        flush(undefined, false, seg.src);
      }
      addToLine(seg, w, h, asc, 0);
      continue;
    }

    // ── Math segment ─────────────────────────────────────
    if ('math' in seg) {
      const render = seg.mathMetadata;
      if (!render || render.available === false) {
        const emPx = seg.fontSize * scale;
        setMeasureFont(buildFont(false, false, emPx, null, fontFamilyClasses));
        const m = ctx.measureText(seg.fallbackText);
        const w = m.width;
        const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? emPx * 0.8;
        const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? emPx * 0.2;
        seg.measuredWidth = w;
        seg.mathAscent = asc;
        seg.mathDescent = desc;
        if (currentLine.length > 0 && currentWidth + w > availW()) {
          flush(undefined, false, seg.src);
        }
        addToLine(seg, w, seg.fontSize, Math.max(asc, emPx * 0.8), Math.max(desc, emPx * 0.2));
        continue;
      }
      const emPx = seg.fontSize * scale;
      const w = render.widthEm * emPx;
      const asc = render.ascentEm * emPx;
      const desc = render.descentEm * emPx;
      seg.measuredWidth = w;
      // Ink extents (from the MathJax SVG viewBox) position the rasterized
      // glyph relative to the baseline when drawing.
      seg.mathAscent = asc;
      seg.mathDescent = desc;
      // …but the LINE BOX must reserve at least a normal single line for the
      // run's font size. A short equation — e.g. a lone "−" — has near-zero ink
      // height; using that as the line height would collapse the line (and the
      // table row) and pin the glyph to the very top of the cell. Floor to the
      // font's natural ascent/descent so math occupies a full line like text
      // does (tall math — fractions, big operators — keeps its larger ink box).
      const lineAsc = Math.max(asc, emPx * 0.8);
      const lineDesc = Math.max(desc, emPx * 0.2);
      if (currentLine.length > 0 && currentWidth + w > availW()) {
        flush(undefined, false, seg.src);
      }
      addToLine(seg, w, seg.fontSize, lineAsc, lineDesc);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    const s = seg as LayoutTextSeg;
    const segmentBox = textSegmentBox(s);
    const w = segmentBox.width;
    const prospectiveWidth = prospectiveSnapAdvance(s, w);
    const h = segmentBox.height;
    const asc = segmentBox.ascent;
    const desc = segmentBox.descent;
    const paragraphFinalIdeographicSpaceTail =
      s.paragraphFinalIdeographicSpaceTail === true;
    const paragraphFinalIdeographicSpaceCount =
      s.paragraphFinalIdeographicSpaceCount ?? 0;
    const paragraphFinalIdeographicSpaceLocalCount =
      s.paragraphFinalIdeographicSpaceLocalCount ?? 0;
    const visibleBeforeParagraphFinalTail = paragraphFinalIdeographicSpaceTail
      ? s.text.slice(0, Math.max(0, s.text.length - paragraphFinalIdeographicSpaceLocalCount))
      : s.text;
    if (
      paragraphFinalIdeographicSpaceTail
      && paragraphFinalIdeographicSpaceCount > 1
      && visibleBeforeParagraphFinalTail.length > 0
    ) {
      const visibleSegment: LayoutTextSeg = {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: visibleBeforeParagraphFinalTail,
        paragraphFinalIdeographicSpaceTail: undefined,
        paragraphFinalIdeographicSpaceLocalCount: undefined,
        paragraphFinalIdeographicSpaceCount: undefined,
        paragraphFinalIdeographicSpaceTailStart: undefined,
        measuredWidth: 0,
        ...slicedTextMetadata(s, 0, visibleBeforeParagraphFinalTail.length),
      };
      const trailingSegment: LayoutTextSeg = {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: s.text.slice(visibleBeforeParagraphFinalTail.length),
        paragraphFinalIdeographicSpaceLocalCount,
        joinPrev: undefined,
        hardJoinPrev: undefined,
        paragraphFinalIdeographicSpaceTailStart: true,
        measuredWidth: 0,
        ...slicedTextMetadata(s, visibleBeforeParagraphFinalTail.length, s.text.length),
        src: s.src
          ? {
              segIndex: s.src.segIndex,
              charOffset: s.src.charOffset + visibleBeforeParagraphFinalTail.length,
            }
          : undefined,
      };
      queue.unshift(trailingSegment);
      queue.unshift(visibleSegment);
      continue;
    }
    if (
      paragraphFinalIdeographicSpaceTail
      && /^\u3000+$/u.test(s.text)
      && s.paragraphFinalIdeographicSpaceTailStart === true
    ) {
      const currentLineHasVisibleText = currentLine.some((candidate) =>
        'text' in candidate && /[^\u3000]/u.test(candidate.text));
      if (currentLineHasVisibleText) {
        let trailingTailWidth = w;
        for (const candidate of queue) {
          if (!('text' in candidate) || candidate.paragraphFinalIdeographicSpaceTail !== true) break;
          trailingTailWidth += segAdvance(candidate);
        }
        if (currentWidth + trailingTailWidth > availW()) {
          flush(undefined, false, s.src);
          queue.unshift(s);
          continue;
        }
      }
    }

    // ECMA-376 §17.3.2.14: a fit region is an atomic fixed-width cell. The
    // first segment judges the WHOLE resolved region; after an optional flush,
    // every member is added without entering the CJK/overlong-word split paths.
    // This also handles a target wider than the line: it overflows as one unit
    // instead of violating the required internal non-wrap boundary.
    if (s.fitTextRegionIndex !== undefined) {
      if (s.fitTextRegionStart) {
        let regionWidth = w;
        for (const queued of queue) {
          if (!('text' in queued) || queued.fitTextRegionIndex !== s.fitTextRegionIndex) break;
          regionWidth += segAdvance(queued);
        }
        if (currentLine.length > 0 && currentWidth + regionWidth > availW()) {
          flush(undefined, false, s.src);
        }
      }
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
      continue;
    }
    // Wrap-fit check uses two standard typographic allowances:
    //   1. Trailing-space collapse: if this word becomes the last on the
    //      line, its trailing space (if any) collapses. We subtract it from
    //      the width used to test fit.
    //   2. Knuth-Plass shrink tolerance: lines the paint pass leaves
    //      non-justified keep the budget. Per §17.18.44, lines the paint pass
    //      fully justifies get no budget: non-final/non-manual-break `both`/kashida
    //      lines, and every `distribute`/`thaiDistribute` line (issue #698).
    const trimmed = s.text.replace(/ +$/, '');
    // Subtract the full-model advance of the trimmed text (not the natural width)
    // so the grid delta, w:w scale and w:spacing pitch on the retained glyphs all
    // cancel and trailingSpaceW is the bare trailing-space advance — keeping `w`
    // and `wForFit` on the one advance model (`strAdvance` == the model behind `w`).
    const trailingSpaceW = snapToCharsClass(s, characterGrid)
      ? 0
      : s.text.endsWith(' ') ? w - strAdvance(s, trimmed) : 0;
    const prospectiveLineWillJustify = (next: LayoutSeg | undefined): boolean => {
      const closesLogicalLine = next === undefined || 'lineBreak' in next;
      return isJustified && (!closesLogicalLine || stretchLastLine);
    };
    const fitWidthFor = (
      widthPx: number,
      trailingSpacePx: number,
      next: LayoutSeg | undefined,
    ): number => wordCandidateFitWidthPx({
      widthPx,
      trailingSpacePx,
      lineWillJustify: prospectiveLineWillJustify(next),
      wrapNarrowed: lineMaxWidth !== maxWidth || lineXOffset !== 0,
    });
    const wForFit = fitWidthFor(
      prospectiveWidth,
      trailingSpaceW,
      queue[0],
    );
    // The two fit-tolerance roles are EXCLUSIVE per line, mirroring paint's
    // per-line predicate `isJustified && (!endsLogicalLine || stretchLastLine)`
    // (`next` is the first segment after the prospective closing candidate):
    //
    //  - A line the paint pass will justify stretches to the column edge. Admit
    //    only the backend-specific per-font measurement bias there; suppress the
    //    trailing-space allowance.
    //  - A line left NON-justified keeps the classic Knuth-Plass trailing-space
    //    shrink allowance, whose 25% promise the draw pass spends through
    //    `shrinkFitCompression`. Adding the bias would double-count tolerance.
    // Dictionary-SEA candidate (Thai/Lao/Khmer; grapheme-fill Myanmar/Tibetan
    // stays on its per-cluster greedy path). Per-codepoint scan: a rare segment
    // mixing both SEA families is not dictionary-SEA, so
    // it keeps the pre-#991 greedy path instead of moving a grapheme-fill span
    // inside an atomic chunk.
    const sDictSea = s.seaBreaks !== undefined && isDictionarySeaText(s.text);
    const candidateMeasurementRoutes = new Set<string>();
    noteMeasurementRoute(candidateMeasurementRoutes, s, trimmed);
    const shrinkBudgetFor = (
      next: LayoutSeg | undefined,
      biasBudget: number,
      measurementRoutes: ReadonlySet<string>,
    ): number => {
      const lineWillJustify = prospectiveLineWillJustify(next);
      if (lineWillJustify) return wordJustifiedCandidateFitAllowancePx({
        biasBudgetPx: biasBudget,
        resolvedMeasurementRouteCount: measurementRouteCountWith(measurementRoutes),
      });
      // `word-dictionary-sea-natural-fit`: a dictionary-SEA line does not
      // compress inter-word spaces. The candidate counts too because admitting
      // it would make the line SEA. Other scripts retain the drawable 25%
      // trailing-space budget.
      return lineHasSea || sDictSea ? 0 : lineTotalTrailingW * SPACE_SHRINK_RATIO;
    };

    // Atomic glued group: when THIS segment starts a glued group (its followers
    // in the queue are `joinPrev` pieces — small-caps case-pieces of the SAME
    // word like "I" then "NTRODUCTION", or a UAX#14 LB13 non-starter authored in
    // its own run like a trailing "," / "。"), the per-segment wrap below would
    // let the group split across lines. Pre-measure it and, if it does not fit on
    // the current (non-empty) line, flush so it starts fresh.
    //
    // ONLY when the lead segment is NOT itself CJK-breakable. A glued group whose
    // lead is a CJK run (e.g. "…通過する" + "。") is NOT atomic: the run splits at
    // an inter-CJK boundary and the trailing non-starter stays on its LAST piece
    // (§17.3.1.16 kinsoku keeps it off the next line's head when enabled — the
    // default; with kinsoku off it may lead the line, as it did before PR #602).
    // Pre-flushing the whole run instead leaves the prior line far short, which a
    // `both` line then stretches wide (sample-9). `joinPrev` stays a pure "this is
    // a non-starter" marker; the atomic-vs-breakable decision lives HERE. A
    // non-breakable Latin / small-caps lead is genuinely atomic, so the pre-flush
    // (and the over-long-word char-break path below) still applies there.
    if (
      !s.joinPrev &&
      currentLine.length > 0 &&
      (queue[0] as LayoutTextSeg | undefined)?.joinPrev &&
      (
        (queue[0] as LayoutTextSeg | undefined)?.hardJoinPrev === true
        || !hasCJKBreakOpportunity(s.text)
      ) &&
      // A SEA (Thai/Lao/Khmer) lead with usable word breaks is NOT atomic — the
      // run splits at a dictionary boundary (issue #797), mirroring the CJK gate.
      (
        (queue[0] as LayoutTextSeg | undefined)?.hardJoinPrev === true
        || !(s.seaBreaks && s.seaBreaks.length > 0)
      )
    ) {
      let groupW = w;
      let groupTrail = trailingSpaceW;
      let groupEnd = 0;
      let groupBiasBudget = lineBiasBudget;
      const groupMeasurementRoutes = new Set(candidateMeasurementRoutes);
      // Keep one pending member so only the final member is trimmed. Committing
      // each previous member left-to-right preserves the former prospective-array
      // summation order exactly, without cloning or rescanning the current line.
      let pendingGroupBiasSeg = s;
      let pendingGroupBiasText = s.text;
      const advanceGroupBias = (member: LayoutTextSeg, text: string = member.text): void => {
        groupBiasBudget += biasBudgetContribution(pendingGroupBiasSeg, pendingGroupBiasText);
        pendingGroupBiasSeg = member;
        pendingGroupBiasText = text;
      };
      for (; groupEnd < queue.length && (queue[groupEnd] as LayoutTextSeg).joinPrev; groupEnd++) {
        const f = queue[groupEnd] as LayoutTextSeg;
        const hardPrefixEnd = hardJoinPrefixEnd(f);
        if (hardPrefixEnd !== undefined) {
          const prefix = f.text.slice(0, hardPrefixEnd);
          const prefixWidth = strAdvance(f, prefix);
          groupW += prefixWidth;
          advanceGroupBias(f, prefix);
          noteMeasurementRoute(groupMeasurementRoutes, f, prefix);
          groupTrail = prefix.endsWith(' ')
            ? prefixWidth - strAdvance(f, prefix.replace(/ +$/, ''))
            : 0;
          // A whole hard member can lead into another hard member. Otherwise
          // the first legal boundary after the seam ends the atomic prefix.
          if (hardPrefixEnd < f.text.length) break;
          continue;
        }
        const firstExternalBreak = f.externalLinkBreakOffsets?.[0];
        if (firstExternalBreak !== undefined) {
          const prefix = f.text.slice(0, firstExternalBreak);
          const prefixWidth = strAdvance(f, prefix);
          groupW += prefixWidth;
          advanceGroupBias(f, prefix);
          noteMeasurementRoute(groupMeasurementRoutes, f, prefix);
          groupTrail = 0;
          break;
        }
        // A CJK-BREAKABLE follower (e.g. "Roman" + "、あるいは…用いる。") is NOT
        // atomic: only its LEADING run of line-start-forbidden chars would orphan
        // at a line head (UAX#14 LB13 / §17.3.1.16); the rest splits at an
        // inter-CJK boundary and wraps on its own. So glue only that prefix's
        // advance to the lead and STOP summing here — mirror of the CJK-lead
        // direction handled by the `!hasCJKBreakOpportunity(s.text)` gate above
        // (sample-9 fb836d6). Summing the whole breakable run instead would
        // pre-flush "Roman" down alone, leaving a `both` line stretched sparse
        // (sample-16). A Latin / small-caps follower (no CJK break opportunity —
        // the "I" + "NTRODUCTION" case) stays fully atomic: keep full-add.
        if (hasCJKBreakOpportunity(f.text)) {
          const chars = [...f.text];
          let p = 0;
          while (p < chars.length && DEFAULT_KINSOKU_RULES.lineStartForbidden.has(chars[p].codePointAt(0)!)) p++;
          if (p < chars.length) {
            // Breakable rest exists past the leading non-starters: glue only the
            // prefix (it may be empty — then "Roman" is effectively unglued and
            // wraps on its own) and end the atomic group here.
            const prefix = chars.slice(0, p).join('');
            const prefixWidth = strAdvance(f, prefix);
            groupW += prefixWidth;
            if (prefix.length > 0) {
              advanceGroupBias(f, prefix);
              noteMeasurementRoute(groupMeasurementRoutes, f, prefix);
            }
            groupTrail = 0;
            break;
          }
          // Entirely non-starters (no breakable rest): fall through to full-add.
        }
        const fw = segAdvance(f);
        groupW += fw;
        advanceGroupBias(f);
        noteMeasurementRoute(groupMeasurementRoutes, f);
        const ft = f.text.replace(/ +$/, '');
        const followerTrail = f.text.endsWith(' ') ? fw - strAdvance(f, ft) : 0;
        // UAX #14 LB7 makes a consecutive SP sequence one trailing suffix even
        // when a source-formatting boundary split it into multiple segments.
        // Accumulate space-only followers so the line-end fit allowance is
        // invariant to that non-textual boundary. A follower containing visible
        // text starts a new suffix and therefore replaces the previous value.
        groupTrail = ft.length === 0 && groupTrail > 0
          ? groupTrail + followerTrail
          : followerTrail;
      }
      groupBiasBudget += biasBudgetContribution(
        pendingGroupBiasSeg,
        pendingGroupBiasText.replace(/ +$/, ''),
      );
      if (
        currentWidth + fitWidthFor(groupW, groupTrail, queue[groupEnd])
        > availW() + shrinkBudgetFor(
          queue[groupEnd],
          groupBiasBudget,
          groupMeasurementRoutes,
        )
      ) {
        flush(undefined, false, s.src);
      }
    }

    // `word-dictionary-sea-atomic-chunk`: ECMA-376 prescribes no SEA
    // line-breaking algorithm. Treat dictionary boundaries inside a no-space
    // Thai/Lao/Khmer chunk as secondary opportunities: move a chunk that fits a
    // full line as a unit; only a full-line-overlong chunk breaks at dictionary
    // boundaries through the greedy SEA branch below.
    //
    // Judged only at chunk START: if the previously committed token is a text
    // segment glued to `s` (no trailing space), the whole chunk already passed
    // this judgment when its head was placed, so a mid-chunk segment never
    // needs it. The chunk spans `s` plus following queue segments while they
    // stay dictionary-SEA text glued without intervening spaces. Grapheme-fill
    // scripts (Myanmar/Tibetan) are excluded because their per-cluster path
    // fills the remaining width.
    if (
      sDictSea &&
      currentLine.length > 0 &&
      (() => {
        const last = currentLine[currentLine.length - 1];
        return !('text' in last) || (last as LayoutTextSeg).text.endsWith(' ');
      })()
    ) {
      let chunkW = w;
      let chunkTrail = trailingSpaceW;
      let chunkEnd = 0;
      let chunkBias = lineBiasBudget + biasBudgetContribution(s, trimmed);
      const chunkMeasurementRoutes = new Set(candidateMeasurementRoutes);
      if (!s.text.endsWith(' ')) {
        for (; chunkEnd < queue.length; chunkEnd++) {
          const f = queue[chunkEnd];
          if (!('text' in f) || (f as LayoutTextSeg).seaBreaks === undefined) break;
          if (!isDictionarySeaText((f as LayoutTextSeg).text)) break;
          const ft = f as LayoutTextSeg;
          const fw = segAdvance(ft);
          const fTrim = ft.text.replace(/ +$/, '');
          chunkW += fw;
          chunkTrail = ft.text.endsWith(' ') ? fw - strAdvance(ft, fTrim) : 0;
          chunkBias += biasBudgetContribution(ft, fTrim);
          noteMeasurementRoute(chunkMeasurementRoutes, ft, fTrim);
          if (ft.text.endsWith(' ')) { chunkEnd++; break; } // a space ends the chunk
        }
      }
      const chunkWForFit = fitWidthFor(chunkW, chunkTrail, queue[chunkEnd]);
      if (
        currentWidth + chunkWForFit > availW() + shrinkBudgetFor(
          queue[chunkEnd],
          chunkBias,
          chunkMeasurementRoutes,
        ) &&
        chunkWForFit <= lineMaxWidth
      ) {
        flush(undefined, false, s.src);
      }
    }

    const shrinkBudget = shrinkBudgetFor(
      queue[0],
      lineBiasBudget + biasBudgetContribution(s, trimmed),
      candidateMeasurementRoutes,
    );
    // §17.3.1.21 is script-neutral: if the segment would fit without its final
    // eligible punctuation character, admit that one character beyond the text
    // extent before selecting a script-specific wrap algorithm. The isolated
    // predicate owns the compatibility character sets. CJK segments that need
    // an internal split retain their separate overflowPunct-vs-kinsoku rule.
    const visibleSegmentScalars = [...trimmed];
    const trailingOverflowCharacter = visibleSegmentScalars.at(-1);
    const textBeforeTrailingOverflow = visibleSegmentScalars.slice(0, -1).join('');
    const admitsTrailingOverflowPunctuation =
      overflowPunct
      && trailingOverflowCharacter !== undefined
      && (currentLine.length > 0 || textBeforeTrailingOverflow.length > 0)
      && wordIsOverflowPunctuation(
        trailingOverflowCharacter,
        s.eastAsiaLanguage,
      )
      && currentWidth + strAdvance(s, textBeforeTrailingOverflow)
        <= availW() + shrinkBudget;

    if (currentWidth + wForFit <= availW() + shrinkBudget) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
      appendQueuedIdeographicSpaceSegment(s);
    } else if (admitsTrailingOverflowPunctuation) {
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
      appendQueuedIdeographicSpaceSegment(s);
    } else if (
      hasCJKBreakOpportunity(s.text)
      && s.seaBreaks === undefined
      && s.hardJoinPrev !== true
    ) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
      // A segment that ALSO contains SEA (a mixed CJK+SEA `<w:cs/>` run) is routed
      // to the SEA branch below instead — its `seaBreaks` already merges the CJK
      // per-character opportunities with the SEA dictionary/transition ones
      // (issue #960), so both scripts break by their own rule from one offset set.
      // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
      //  separate: it sums per-char advances, whereas this path uses substring
      //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
      const available = availW() - currentWidth;
      let rawPrefix = '';
      const maximumIdeographicSpaceHang = paragraphFinalIdeographicSpaceTail
        ? wordIdeographicSpaceLineEndAllowanceCount(
            hasEastAsianVisiblePredecessor(s.text),
            s.paragraphFinalIdeographicSpaceCount ?? 0,
          )
        : Number.POSITIVE_INFINITY;
      if (available > 0) {
        const nonMonotoneAllocation = charSpacingDeltaPx(s, scale) < 0
          || snapToCharsClass(s, characterGrid) === 'latin';
        if (nonMonotoneAllocation) {
          rawPrefix = s.text.slice(0, emergencyTextSplit(s, available, false));
        } else {
          setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses, s.fontRoute));
          const prevKern = setSegKerning(s);
          try {
            rawPrefix = fitCJKPrefix(
              ctx,
              s.text,
              available,
              segmentCharacterGridDeltaPx(s, characterGrid, scale),
              charScaleFactor(s),
              charSpacingDeltaPx(s, scale),
              s.verticalRun === true,
              verticalGlyphMeasurement,
              (prefix) => strAdvance(s, prefix),
              maximumIdeographicSpaceHang,
            );
          } finally {
            restoreKerning(prevKern);
          }
        }
      }
      // Apply kinsoku to the break position: retract leftwards so the tail
      // never begins with a 行頭禁則 char and the head never ends with a
      // 行末禁則 char (ECMA-376 §17.15.1.58–.60). When the current line
      // already has content, retracting to an empty prefix is allowed — the
      // whole run moves to the next (fresh) line, which is Word's 追い出し.
      // When the line is empty we keep at least one char (minSplit=1) so we
      // never lose forward progress.
      const allChars = [...s.text];
      const rawSplit = [...rawPrefix].length;
      const minSplit = currentLine.length > 0 ? 0 : 1;
      // ECMA-376 §17.3.1.21 permits one punctuation character beyond the
      // paragraph extents. The isolated compatibility projection resolves the
      // language-specific set and its precedence over kinsoku at this internal
      // CJK split.
      const hangingSplit = overflowPunct
        && rawSplit < allChars.length
        && (currentLine.length > 0 || rawSplit > 0)
        && wordIsOverflowPunctuation(allChars[rawSplit], s.eastAsiaLanguage)
          ? rawSplit + 1
          : null;
      const proposedSplit = extendThroughTrailingIdeographicSpaces(
        allChars,
        hangingSplit ?? kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, minSplit),
        paragraphFinalIdeographicSpaceTail && maximumIdeographicSpaceHang === 0
          ? 0
          : maximumIdeographicSpaceHang,
      );
      const proposedPrefix = allChars.slice(0, proposedSplit).join('').length;
      const protectedSplit = legalTextSplitAtOrBefore(s, proposedPrefix, minSplit > 0 ? 1 : 0);
      const prefix = s.text.slice(0, protectedSplit);
      if (prefix.length > 0) {
        // Grid advance for the head piece — the same model as the line box / draw.
        const pw = strNaturalAdvance(s, prefix);
        const headSeg: LayoutTextSeg = {
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        };
        addToLine(headSeg, pw, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) {
          queue.unshift({
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: tail,
            ...slicedTextMetadata(s, prefix.length, s.text.length),
            measuredWidth: 0,
            src: {
              segIndex: s.src!.segIndex,
              charOffset: s.src!.charOffset + prefix.length,
            },
          });
        } else {
          appendQueuedIdeographicSpaceSegment(s);
        }
      } else if (currentLine.length > 0) {
        // No prefix of `s` fits. If `s` would lead the next line with a 行頭禁則
        // char, kinsokuAdjustedSplit can't fix it from within `s` (the offending
        // char is its first); pull trailing graphemes of the current line's last
        // text segment down so they lead the next line ahead of `s` — cross-run
        // 追い出し (§17.3.1.16). See crossRunKinsokuRetract for the bounded,
        // re-validating, whitespace-guarded retraction count.
        const retraction = retractCurrentLineForLeadingKinsoku(s);
        if (retraction.kind === 'blocked') {
          keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
          continue;
        }
        flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
        queue.unshift(s);
        if (retraction.kind === 'retracted') queue.unshift(retraction.tail);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const forcedChars = [...s.text];
        const forcedSplit = forcedChars.length > 0
          ? extendThroughTrailingIdeographicSpaces(
              forcedChars,
              1,
              s.paragraphFinalIdeographicSpaceTail === true
                ? wordIdeographicSpaceLineEndAllowanceCount(
                    EAST_ASIAN_RE.test(forcedChars[0] ?? ''),
                    s.paragraphFinalIdeographicSpaceCount ?? 0,
                  )
                : Number.POSITIVE_INFINITY,
            )
          : 0;
        const forcedUtf16 = forcedChars.slice(0, forcedSplit).join('').length;
        const legalForcedUtf16 = legalTextSplitAtOrBefore(s, forcedUtf16)
          || emergencyTextSplit(s, availW(), true);
        const firstChar = s.text.slice(0, legalForcedUtf16);
        if (firstChar) {
          const fw = strNaturalAdvance(s, firstChar);
          const headSeg: LayoutTextSeg = {
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: firstChar,
            measuredWidth: fw,
            ...slicedTextMetadata(s, 0, firstChar.length),
          };
          addToLine(headSeg, fw, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) {
            queue.unshift({
              ...s,
              ...RESET_SLICED_TEXT_MEASUREMENT,
              text: tail,
              ...slicedTextMetadata(s, firstChar.length, s.text.length),
              measuredWidth: 0,
              src: {
                segIndex: s.src!.segIndex,
                charOffset: s.src!.charOffset + firstChar.length,
              },
            });
          } else {
            appendQueuedIdeographicSpaceSegment(s);
          }
        }
      }
    } else if (s.seaBreaks !== undefined && s.hardJoinPrev !== true) {
      // No-inter-word-space line wrap: Thai/Lao/Khmer dictionary words (#797) or
      // Myanmar/Tibetan grapheme clusters (#961). This ONE segment is a whole run;
      // break it only at a member of `s.seaBreaks` — the UNION (#960) of the
      // dictionary word (or grapheme-cluster) boundaries, the no-space SEA↔non-SEA
      // script transitions, and (for a mixed CJK+SEA `<w:cs/>` run) the CJK
      // per-character opportunities, already kinsoku-filtered by
      // `seaMixedBreakOffsets`. Entered for ANY such segment (even one with no
      // interior boundary — a single word/cluster wider than the column, or
      // Segmenter unavailable) so the emergency split below stays GRAPHEME-safe
      // instead of falling to the code-point path. Kinsoku 行頭/行末禁則 was applied
      // when the offsets were built (so a forbidden CJK char never heads/tails a
      // line); choosing an earlier legal offset is the only remaining adjustment,
      // which fitSeaWordPrefix already does. The run stays one contiguous draw per
      // line (measure==paint); the tail re-queues with its offsets rebased.
      const available = availW() - currentWidth;
      const measureSub = (sub: string): number => strAdvance(s, sub);
      // Grapheme-fill runs (Myanmar/Tibetan) have DENSE offsets (one per cluster),
      // so use the monotone binary-search fit — a per-line full scan would be O(n²)
      // down a long run. Dictionary runs keep the negative-spacing-safe full scan.
      const monotone = isGraphemeFillText(s.text)
        && charSpacingDeltaPx(s, scale) >= 0
        && snapToCharsClass(s, characterGrid) !== 'latin';
      const split = fitSeaWordPrefix(s.text, s.seaBreaks, 0, available, measureSub, monotone);
      if (split > 0) {
        const prefix = s.text.slice(0, split);
        const pw = strNaturalAdvance(s, prefix);
        addToLine({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        }, pw, h, asc, desc);
        const tail = s.text.slice(split);
        if (tail) {
          queue.unshift({
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: tail,
            ...slicedTextMetadata(s, split, s.text.length),
            measuredWidth: 0,
            src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + split },
            seaBreaks: rebaseSeaBreaks(s.seaBreaks, split),
          });
        }
      } else if (currentLine.length > 0) {
        // No whole word fits the remaining band — move the run to a fresh line and
        // re-process (Latin-word style). If `s` would then LEAD the next line with
        // a 行頭禁則 char (a mixed CJK+SEA run whose first glyph is a forbidden
        // leader — #960 routes it here, where the offset set cannot fix a
        // segment-initial char), pull trailing graphemes of the current line's
        // last text segment down so they lead ahead of `s` — the same cross-run
        // 追い出し (§17.3.1.16) the CJK branch does.
        const retraction = retractCurrentLineForLeadingKinsoku(s);
        if (retraction.kind === 'blocked') {
          keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
          continue;
        }
        flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
        queue.unshift(s);
        if (retraction.kind === 'retracted') queue.unshift(retraction.tail);
      } else {
        // Empty line and the first dictionary word is wider than the whole
        // column: emergency GRAPHEME-safe split (a code-point split would tear a
        // base + tone/combining mark, both BMP). Guarantee ≥1 cluster of progress.
        const firstWordEnd = s.seaBreaks[0] ?? s.text.length;
        const firstWord = s.text.slice(0, firstWordEnd);
        const graphemes = graphemeClusterOffsets(firstWord);
        let gsplit = fitSeaWordPrefix(firstWord, graphemes, 0, available, measureSub, monotone);
        if (gsplit <= 0) gsplit = graphemes.length > 0 ? graphemes[0] : firstWord.length;
        gsplit = legalTextSplitAtOrBefore(s, gsplit)
          || emergencyTextSplit(s, available, true);
        const prefix = s.text.slice(0, gsplit);
        const pw = strNaturalAdvance(s, prefix);
        addToLine({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        }, pw, h, asc, desc);
        const tail = s.text.slice(gsplit);
        if (tail) {
          queue.unshift({
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: tail,
            ...slicedTextMetadata(s, gsplit, s.text.length),
            measuredWidth: 0,
            src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + gsplit },
            seaBreaks: rebaseSeaBreaks(s.seaBreaks, gsplit),
          });
        }
      }
    } else if (currentLine.length === 0) {
      // `word-overlong-token-emergency-break`: for a single non-CJK token wider
      // than a full line, fit the widest character prefix (at least one
      // character), draw it, and re-queue the remainder. Segments are already
      // space-delimited, so this cannot bypass an ordinary space opportunity.
      const split = externalLinkSyntaxSplit(s, availW()) || emergencyTextSplit(s, availW());
      if (split >= s.text.length) {
        // The visible glyphs actually fit (only a trailing space pushed it over the
        // fit test) — place the word whole.
        s.measuredWidth = w;
        addToLine(s, w, h, asc, desc);
      } else {
        const prefix = s.text.slice(0, split);
        const pw = strNaturalAdvance(s, prefix);
        addToLine({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        }, pw, h, asc, desc);
        queueEmergencyTail(s, split);
      }
    } else {
      const semanticSplit = externalLinkSyntaxSplit(
        s,
        availW() + shrinkBudget - currentWidth,
      );
      if (semanticSplit > 0 && semanticSplit < s.text.length) {
        const prefix = s.text.slice(0, semanticSplit);
        const pw = strNaturalAdvance(s, prefix);
        addToLine({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        }, pw, h, asc, desc);
        queueEmergencyTail(s, semanticSplit);
        continue;
      }
      if (s.joinPrev) {
        // LB14 and the other UAX glue rules prohibit a line boundary at this
        // source seam. If the complete glued group is wider than the fresh
        // line, split this member at the widest legal grapheme boundary that
        // fits the actual remaining band. This bases the decision on the group
        // advance, not on the follower's standalone width.
        const remaining = availW() - currentWidth;
        const split = emergencyTextSplit(s, remaining, true);
        if ((remaining > 0 || s.hardJoinPrev === true) && split > 0 && split < s.text.length) {
          const prefix = s.text.slice(0, split);
          const pw = strNaturalAdvance(s, prefix);
          addToLine({
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: prefix,
            measuredWidth: pw,
            ...slicedTextMetadata(s, 0, prefix.length),
          }, pw, h, asc, desc);
          queueEmergencyTail(s, split);
          continue;
        }
        // A scalar span that continues the preceding grapheme (or another
        // explicitly glued piece) may overflow a pathological narrow line, but
        // it must never become a new line head and tear the cluster.
        s.measuredWidth = w;
        addToLine(s, w, h, asc, desc, trailingSpaceW);
        continue;
      }
      // Latin token does not fit on the current (non-empty) line: move it to a fresh
      // line and re-process. There it either fits, or — when it is wider than the
      // whole column — the empty-line branch above breaks it at the character level
      // (overflow-wrap). Re-queueing rather than force-adding is what lets that
      // over-long-word path run instead of letting the word spill the column.
      flush(undefined, false, s.src);
      queue.unshift(s);
    }
  }

  if (currentLine.length > 0) flush();
  // Trailing <w:br/>: emit the empty line it opened (§17.3.3.1).
  else if (trailingBreakFontSize !== null) flush(trailingBreakFontSize);

  // A3 acquisition consumes final line pieces, not the pre-wrap source
  // segments. Prefix/tail objects created by the breakers above may inherit the
  // source segment's cluster array, whose ranges describe a different string.
  // Re-shape every final visible piece through the same A2 authority so the
  // returned LayoutLine contract always carries complete, piece-relative
  // grapheme geometry. A missing service is deliberately left unshaped; the
  // retained acquisition boundary rejects that production contract violation.
  if (widthPolicy === 'bounded') {
    for (const line of lines) {
      for (const segment of line.segments) {
        if (!('text' in segment) || segment.metricOnly || segment.text.length === 0) continue;
        segment.shapedClusters = undefined;
        if (segment.textLayoutService && segment.textShapeRequest) measureText(segment, true);
      }
    }
  }

  return lines;
}
