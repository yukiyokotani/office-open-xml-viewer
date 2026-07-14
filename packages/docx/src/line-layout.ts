// DOCX line-layout engine — the pure segmentation + line-breaking + measurement
// kernel that both the paginator and the paint pass call to turn a paragraph's
// runs into laid-out lines and line-box heights (ECMA-376 §17.3.1.x line
// spacing, §17.3.1.37 tabs, §17.6.5 docGrid, §17.15.1.58–.60 kinsoku, §17.3.2.26
// script/font axes, §17.3.2.33 small caps, §17.8.3.10 font classification).
//
// Lifted verbatim out of renderer.ts along the domain phase boundary that the B2
// text-layout unification established (paragraph #684/#689, table #693, textbox
// #697): everything here MEASURES (it may touch a Canvas 2D context to call
// measureText / set ctx.font) but never DRAWS, mutates RenderState, paginates, or
// registers floats — those stay in renderer.ts, which imports this module. The
// split is one-directional at runtime: renderer.ts → line-layout.ts. The only
// back-reference is the RenderState / DecodedImage TYPE (import type, erased at
// runtime — same idiom frame-geometry.ts / anchor-geometry.ts use), so there is
// no import cycle. Which behaviours here are ECMA-376-mandated vs Word-mimicking
// is documented inline (as before the move) — see packages/docx/CLAUDE.md.

import type {
  DocParagraph, DocRun, DocxTextRun, ImageRun, ShapeTextRun, FieldRun,
  LineSpacing, TabStop, DocxRunBorder, DocSettings, EmphasisMark,
} from './types';
import type { MathNode, KinsokuRules, ChartModel, HyperlinkTarget, NumberFormat, Duotone, ResolvedLocalFontMetric, TextSourceRef } from '@silurus/ooxml-core';
import type { RenderState, DecodedImage } from './renderer.js';
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
  normalizeLocalFontMetricFamily,
  sliceTextSourceRefs,
} from '@silurus/ooxml-core';
import { intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';
import { groupFitTextRegions, type FitTextRun } from './fit-text.js';
import {
  type FloatRect,
  resolveLineFloatWindow,
  wordMinLineStartPx,
} from './float-layout.js';
import { verticalRunInkExtraPx } from './vertical-text.js';

export interface LineBoundary {
  segIndex: number;
  charOffset: number;
}

interface LayoutSegSource {
  src?: LineBoundary;
}

export interface LayoutTextSeg extends LayoutSegSource {
  text: string;
  /** Full source mapping for the unsplit layout segment. */
  sourceRefs?: TextSourceRef[];
  /** Zero-advance anchor-character placeholder: contributes run metrics to the
   * line box but paints no glyph. */
  metricOnly?: true;
  /** The anchor character uses the run's East Asian font axis, so §17.6.5 grid
   * allocation must use Far East design metrics despite `text` being empty. */
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
  /** Exact local face selected during async document loading. The family above
   * is its isolated alias; this measured ratio supplies Word's design-line floor
   * without a version-specific font constant. */
  resolvedLineHeightRatio?: number;
  vertAlign: 'super' | 'sub' | null;
  measuredWidth: number;  // px (set during layout)
  smallCaps?: boolean;
  /** This segment is GLUED to the preceding one (no inter-segment break): they
   *  are case-pieces of the same word emitted at different sizes for small caps
   *  (§17.3.2.33) — e.g. "I"(full)+"NTRODUCTION"(reduced). The line breaker must
   *  not start a new line before a glued segment; it retracts the whole glued
   *  group instead, so a small-caps word never splits across lines. */
  joinPrev?: boolean;
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
  /** Track-changes revision attached to this run (insertion / deletion). */
  revision?: { kind: 'insertion' | 'deletion' | string; author?: string };
  /** ECMA-376 §17.3.2.30 `<w:rtl>` — run carries right-to-left characteristics.
   *  When true the segment's text is treated as a strong-RTL embedding in the
   *  per-line bidi pass (so leading digits / neutrals resolve RTL). */
  rtl?: boolean;
  /** UAX#9 §4.3 HL1 / Word behaviour: classify this segment's European digits
   *  (U+0030–0039) as Arabic-Number (AN) in the per-line bidi pass, so a date
   *  like "28-02-2026" in an Arabic complex-script run reorders to "2026-02-28"
   *  exactly as Word renders it (ECMA-376 §17.3.2.20 w:lang w:bidi). */
  digitsAsAN?: boolean;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`) DECLARED on the
   *  originating run, retained purely for a line-box DESIGN-LINE FLOOR. Word
   *  reserves the declared eastAsia face's line height even when this particular
   *  segment renders Latin glyphs (the common Japanese encoding puts a tabled CJK
   *  face — Meiryo — on eastAsia while ascii stays an untabled Latin default). The
   *  BODY line breaker ignores this (its floor is `intendedSingleLinePx(fontFamily)`
   *  per segment, so body behaviour is unchanged); the TEXT-BOX metrics
   *  (`lineMetricsFor`) floor on it so a text box matches PR #640/#646/#648. */
  eaFloorFamily?: string | null;
  resolvedEaFloorLineHeightRatio?: number;
  /** IX1 — the resolved hyperlink target of the originating run (ECMA-376
   *  §17.16.22 external `r:id` URL / §17.16.23 internal `w:anchor` bookmark),
   *  computed once per run in `buildSegments`. Carried purely so the text-layer
   *  overlay can build a clickable region; it does NOT affect measurement, line
   *  breaking, or the drawn glyphs. Absent for a non-link run. */
  hyperlink?: HyperlinkTarget;
  /** ECMA-376 §17.3.2.34 `<w:snapToGrid>` — false opts this run out of the
   *  section character grid without changing paragraph line-grid policy. */
  snapToCharacterGrid?: boolean;
  /** ECMA-376 §17.3.2.35 `<w:spacing>` — character-spacing pitch in POINTS
   *  (signed), added after every character of the run. Applied as a per-glyph
   *  `ctx.letterSpacing` delta on BOTH measure and paint (measure==paint), on top
   *  of any docGrid / justify delta. Absent ⇒ 0. */
  charSpacing?: number;
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
   *  font size or the line box. Absent ⇒ 0. */
  position?: number;
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

/** ECMA-376 §17.3.3.12 defines `hpsRaise` as the “distance [...] between the
 *  phonetic guide base text and the phonetic guide text.” Use that distance as
 *  the ruby ascent reservation when present; preserve the existing 1.5×
 *  ruby-size fallback when it is absent because the property has no default. */
export function rubyAscentReservePx(
  rubySizePt: number,
  hpsRaisePt: number | undefined,
  scale: number,
): number {
  return hpsRaisePt != null ? hpsRaisePt * scale : rubySizePt * scale * 1.5;
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
  /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-rectangle crop (fractions 0..1 of
   *  the decoded bitmap). When present the draw paths use the 9-arg
   *  `drawImage` to blit only `[l, t, 1−r, 1−b]` of the bitmap into the display
   *  box. `undefined` ⇒ draw the full bitmap. */
  srcRect?: { l: number; t: number; r: number; b: number };
  /** ECMA-376 §21.2 — when set, this "image" box is actually a DrawingML chart.
   *  The box is sized like a picture (via {@link LayoutImageSeg.widthPt}/
   *  {@link LayoutImageSeg.heightPt}) and painted with the shared `renderChart`
   *  instead of blitting a bitmap: an inline chart seg flows with the text and
   *  is drawn at its flow position; an anchored chart seg (`anchor: true`,
   *  §20.4.2.3) is zero-width in the flow and the chart is drawn at its
   *  absolute page box by `renderAnchorImages`. `imagePath`/`mimeType` are
   *  empty sentinels for a chart seg — no blip is fetched (the bitmap-prefetch
   *  walk keys off `run.type === 'image'` and never sees a chart run). */
  chart?: ChartModel;
  measuredWidth: number;
}

/** An inline OMML equation. Measured + drawn via the core math engine. */
export interface LayoutMathSeg extends LayoutSegSource {
  mathNodes: import('@silurus/ooxml-core').MathNode[];
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
  /** px — intended single-line height (max over segments of the requested
   *  font's win line-height ratio × em), for fonts whose substituted Canvas
   *  metrics understate Word's line spacing. 0 when no segment needs it. */
  intendedSingle: number;
  /** px — DESIGN grid-count height: the max over segments of each run's
   *  Word-faithful single-line height (a tabled run's design height, an
   *  untabled East Asian run's 1.3em Word FE fallback). Feeds docGrid cell
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
  /** ECMA-376 §17.6.5 `<w:docGrid w:charSpace>` divided by 4096 — the per-EA-
   *  glyph character-grid delta = charSpace/4096 in FLAT POINTS (independent of
   *  font size), added to the measured glyph advance (≈1em for full-width EA
   *  glyphs). Negative tightens. `null`/`undefined` when the section declares no
   *  charSpace; the character grid is then inactive even if `type` is
   *  linesAndChars/snapToChars. See {@link gridCharDeltaPx}. */
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
  readonly noteNumbers?: ReadonlyMap<string, number>;
  readonly currentNoteNumber?: number;
  readonly verticalCJK?: boolean;
  readonly resolvedLocalFonts?: Readonly<Record<string, ResolvedLocalFontMetric>>;
}

// ── Math (OMML) rendering via MathJax ───────────────────────────────────────
// Each equation is converted OMML AST -> MathML -> MathJax SVG, then rasterized to
// an <img> once (async, before pagination). Layout reads cached em-extents
// synchronously; drawing blits the image. Skipped entirely for math-free documents.
export interface MathRender {
  img: CanvasImageSource;
  /** baseline-relative extents in em (1em = the equation's font size). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
}

// Keyed by the run's MathNode[] reference, which is stable from parse through render.
export const mathRenders = new WeakMap<MathNode[], MathRender>();

/** Arabic-script faces that hosts rarely ship; we substitute them with Noto
 *  Naskh/Sans Arabic web fonts (see DOCX_GOOGLE_FONTS in document.ts — this
 *  list MUST mirror the Arabic entries there). A run whose font is one of these
 *  contains BOTH Arabic and Latin/digit glyphs that Word renders from the same
 *  single face, so the fallback chain must keep both scripts stylistically
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
 *     - "modern"     → monospace only for `pitch="fixed"` (§17.8.3.29), else
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
 * `fontFamilyClasses` is a stable per-document object (RenderState threads
 * `doc.fontFamilyClasses` — one identity per render). Keying the outer WeakMap on
 * that object identity gives per-doc caching with zero call-site churn (both
 * callers already pass `fontFamilyClasses`) and no leak: the inner
 * family→result Map is collected with the document's classes object. Same idiom
 * as `sheetAxisCache` / `mathRenders`. Chosen over threading an explicit cache
 * param through buildFont because both call sites already carry the classes
 * object, so identity-keying needs no signature changes anywhere.
 */
export const fontFamilyNormalizeCache = new WeakMap<Record<string, string>, Map<string, string>>();

/** Companion to {@link fontFamilyNormalizeCache}: maps a `fontFamilyClasses`
 *  object (stable per-document identity) to the sibling per-font PITCH map
 *  (ECMA-376 §17.8.3.29 `<w:pitch>`: font name → "fixed" | "variable" |
 *  "default"). `normalizeFontFamily` reads it to decide whether a
 *  `family="modern"` (§17.8.3.10) face is genuinely monospace: only "fixed"
 *  (§17.18.66 Fixed Width) is. Keyed on the classes object so the pitch threads
 *  for free through every existing `fontFamilyClasses` call site — exactly like
 *  the normalize cache — with no second map plumbed through the renderer. */
export const fontFamilyPitchesByClasses = new WeakMap<
  Record<string, string>,
  Record<string, string>
>();

/** Bind the §17.8.3.29 pitch map to the §17.8.3.10 classes object and return the
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
  //    Latin letters/digits; Word draws both from that one face. The browser
  //    resolves each glyph against the chain in order, so the Arabic substitute
  //    MUST come first — otherwise the Latin/digit glyphs are grabbed by the
  //    first chain member that has them (e.g. the CJK "Noto Sans JP"), and
  //    Latin/digits render in a different, sans face than the Arabic. Keeping
  //    the Arabic substitute first makes Arabic+Latin+digits all resolve from
  //    one family, matching Word's single-face rendering.
  //
  //    Latin companion: traditional Naskh faces (Sakkal Majalla, Traditional
  //    Arabic, …) ship a SERIF Latin companion — Word's PDF export of sample-7
  //    renders the Latin "first leader name" with bracketed serifs and the
  //    "2026" digits as serif figures. Noto Naskh Arabic, our substitute, also
  //    ships a serif Latin face (verified: its Latin glyphs carry bracketed
  //    serifs and closely match the PDF), so placing it first gives Latin+digits
  //    a serif look consistent with the Arabic — matching Word. "Noto Serif" is
  //    kept as a serif safety net for the rare case Noto Naskh Arabic is
  //    unavailable, so Latin still falls to a serif rather than the CJK sans.
  //    Geometric Arabic faces (Univers Next Arabic) pair with a sans Latin.
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
        // family value classifies the DESIGN, not the pitch — §17.8.3.29
        // `<w:pitch>` states the actual pitch. Treat the face as monospace ONLY
        // when pitch is "fixed" (§17.18.66 Fixed Width). A "variable"
        // (proportional) modern face — e.g. Meiryo UI (`family="modern"`,
        // `pitch="variable"`), a condensed ~0.84em CJK sans — must NOT map to
        // Courier/monospace: that measures its CJK at a full 1.0em and over-wraps
        // table cells onto a spurious extra page (issue #855). "default" and an
        // omitted `<w:pitch>` (assumed "default" per §17.8.3.29) are likewise not
        // a fixed-width guarantee, so they fall through to the name-pattern /
        // CJK-sans path below. Genuine monospace faces (Courier, Consolas, 等幅)
        // are still caught there by name.
        if (fontFamilyPitches[family] === 'fixed') {
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
): string {
  const w = bold ? 'bold' : 'normal';
  const s = italic ? 'italic' : 'normal';
  const f = normalizeFontFamily(family, fontFamilyClasses);
  return `${s} ${w} ${sizePx}px ${f}`;
}

export function calcEffectiveFontPx(s: LayoutTextSeg, scale: number): number {
  // ECMA-376 §17.3.2.33: small-caps small letters render "in a font size TWO
  // POINTS SMALLER than the actual font size" (subtractive, not a ratio), floored
  // at the smallest renderable size when 2pt smaller is not possible. Applied in
  // pt before scaling. (At ~10pt this is ≈0.8×, but it diverges at larger sizes —
  // a 20pt heading's caps are 18pt, not 16pt.)
  const pt = s.smallCaps ? Math.max(s.fontSize - 2, 1) : s.fontSize;
  let size = pt * scale;
  if (s.vertAlign) size *= 0.65;
  return size;
}

/** Design single-line floor for a measured segment. An exact local face, when
 * resolved during document loading, supersedes the static family profile; the
 * latter remains the fallback when local() is unavailable or rejected. */
export function segmentIntendedSingleLinePx(
  segment: LayoutTextSeg,
  emPx: number,
  eastAsian = false,
): number {
  return Math.max(
    intendedSingleLinePx(segment.fontFamily, emPx, eastAsian),
    (segment.resolvedLineHeightRatio ?? 0) * emPx,
  );
}

export function segmentEastAsiaFloorSingleLinePx(
  segment: LayoutTextSeg,
  emPx: number,
  eastAsian = false,
): number {
  return Math.max(
    intendedSingleLinePx(segment.eaFloorFamily, emPx, eastAsian),
    (segment.resolvedEaFloorLineHeightRatio ?? 0) * emPx,
  );
}

export function getDefaultFontSize(para: DocParagraph): number {
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

/** First text/field run's font family — used to size empty paragraphs whose
 *  intended font (e.g. Meiryo) has a larger win line height than the fallback.
 *  Empty paragraphs (no runs) fall back to the paragraph's style-resolved
 *  default font so e.g. an empty Meiryo cell that forms a résumé "bar" reserves
 *  Meiryo's tall line box rather than the generic fallback's. */
export function getDefaultFontFamily(para: DocParagraph, eastAsian = false): string | null {
  for (const run of para.runs) {
    if (run.type === 'text') return (run as unknown as DocxTextRun).fontFamily;
    if (run.type === 'field') return (run as unknown as FieldRun).fontFamily;
  }
  if (eastAsian && para.defaultFontFamilyEastAsia) return para.defaultFontFamilyEastAsia;
  return para.defaultFontFamily ?? null;
}

/** Intended single-line height (px) for an empty paragraph, from its default
 *  font's win line-height ratio. 0 when the font is not in the metrics table. */
export function emptyIntendedSinglePx(para: DocParagraph, scale: number): number {
  return intendedSingleLinePx(getDefaultFontFamily(para), getDefaultFontSize(para) * scale);
}

/** Intended single-line height (px) for an empty paragraph in the script axis
 *  used to draw its paragraph mark. */
function emptyIntendedSingleForScriptPx(para: DocParagraph, scale: number, eastAsian: boolean): number {
  return intendedSingleLinePx(getDefaultFontFamily(para, eastAsian), getDefaultFontSize(para) * scale, eastAsian);
}

/** Code points whose presence marks a line as East Asian for docGrid line-cell
 *  rounding: CJK symbols/punctuation, Hiragana, Katakana, CJK Unified +
 *  Extension A, compatibility ideographs, Hangul, and fullwidth forms. Content
 *  test only — not a font-name heuristic (cf. packages/docx/CLAUDE.md). */
export const EAST_ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟　-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿＀-￯]/u;

/** Per-EA-glyph character-grid delta in px for a paragraph's grid, or 0 when the
 *  CHARACTER grid is inactive. Active only for docGrid type ∈ {linesAndChars,
 *  snapToChars} with a declared charSpace (ECMA-376 §17.6.5). The line grid
 *  ("lines") and a missing charSpace leave EA glyphs at natural advance. */
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

/** The total character-grid delta (px) a segment's advance gains under an active
 *  character grid. Applied ONLY to a PURE East-Asian segment (every code point
 *  is EA): then `len × deltaPx` cells the whole run, and a uniform per-cp
 *  letter-spacing of `deltaPx` reproduces it exactly on both draw paths. A mixed
 *  or pure-Latin segment returns 0 — Latin is never snapped (§17.6.5), and
 *  skipping mixed segments avoids the per-cp-vs-whole-string and justification
 *  drift that would break measure==draw. `deltaPx===0` (grid inactive) ⇒ 0. */
export function gridSegDeltaPx(text: string, deltaPx: number): number {
  if (deltaPx === 0 || text.length === 0) return 0;
  const cps = [...text];
  return eaGlyphCount(text) === cps.length ? cps.length * deltaPx : 0;
}

/** Resolve the per-glyph character-grid delta for one text segment. */
export function segmentCharacterGridDeltaPx(
  seg: LayoutTextSeg,
  gridDeltaPx: number,
): number {
  return seg.snapToCharacterGrid === false ? 0 : gridDeltaPx;
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
  return (seg.charSpacing ?? 0) * scale;
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
  gridDeltaPx: number,
  charScale: number,
  charSpacingPx: number,
): number {
  return naturalWidthPx * charScale
    + gridSegDeltaPx(text, gridDeltaPx)
    + [...text].length * charSpacingPx;
}

/** Total per-code-point letter-spacing (px) a segment draws with: the docGrid
 *  cell delta (East-Asian-only, {@link gridSegDeltaPx}'s per-cp value) PLUS the
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
  gridDeltaPx: number,
  scale: number,
): number {
  if (seg.fitTextPerGapPx !== undefined) return seg.fitTextPerGapPx;
  const segmentDelta = segmentCharacterGridDeltaPx(seg, gridDeltaPx);
  const grid = gridSegDeltaPx(seg.text, segmentDelta) === 0 ? 0 : segmentDelta;
  return grid + charSpacingDeltaPx(seg, scale);
}

/** A text segment's laid-out advance INCLUDING the §17.3.2.43 horizontal scale
 *  and the §17.3.2.35 character spacing, on top of the docGrid delta. The
 *  natural width is scaled first (w:w stretches the glyphs), then the char-spacing
 *  pitch is added per code point (w:spacing adds fixed gaps that w:w does not
 *  stretch), matching Word's independent treatment of the two axes. */
export function segAdvanceWidth(
  seg: LayoutTextSeg,
  naturalWidthPx: number,
  gridDeltaPx: number,
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
  const segmentDelta = segmentCharacterGridDeltaPx(seg, gridDeltaPx);
  return textAdvanceWidth(
    naturalWidthPx,
    seg.text,
    segmentDelta,
    charScaleFactor(seg),
    charSpacingDeltaPx(seg, scale),
  );
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
 * line's SINGLE-LINE HEIGHT `naturalPx` (the document font's design line
 * height: max of the corrected glyph box and the intendedSingleLinePx floor).
 * The count is `ceil(naturalPx / pitchPx)` — the smallest number of whole
 * cells that CONTAINS the line.
 *
 * Adjudicated by the sample-58 sweep (issue #1013; Word PDF, pdftotext -bbox;
 * 19 sections over {10.5,12,14,16,20}pt × pitch {18,24}pt × {lrTb,tbRl} ×
 * {lines,linesAndChars,none}, all Yu Mincho): with Yu Mincho's design line
 * height (1.3 × hhea box = 1.43267 em — see core line-metrics) every measured
 * point is ceil(design/pitch): on an 18pt pitch 10.5/12pt → 1 cell and
 * 14/16/20pt → 2 cells; on a 24pt pitch 12/16pt → 1 cell and 20pt → 2 cells.
 * Horizontal (lrTb) and vertical (tbRl) sections measured IDENTICAL pitches,
 * so the rule is direction-agnostic (the tbRl column pitch is this same cell
 * height), and the §17.6.5 grid type does not change the count. The pre-sweep
 * calibration points remain satisfied: sample-35's 12pt heading / 10.5pt body
 * on 18pt → 1 cell (design 17.19 / 15.04 < 18) and sample-9's 20pt title on a
 * 20pt pitch → 2 cells (design 28.65). An earlier em-based rule
 * (floor(em/pitch)+1) fit those sparse points but under-counted every
 * 14–16pt-class line whose design height exceeds the pitch (Word: 2 cells).
 *
 * A line that fills k pitches exactly occupies k cells (ceil; no measured
 * point sits on the boundary — the geometric reading is that it still FITS).
 * For a mixed-size line, callers supply the tallest run's resolved height
 * (§17.3.1.33 tallest-run line box). ECMA-376 defines `linePitch` as one
 * single-spaced line; spreading taller lines over whole cells is Word runtime
 * behaviour. Returns at least 1 for every finite `naturalPx >= 0`.
 */
export function docGridLineCells(naturalPx: number, pitchPx: number): number {
  return pitchPx > 0 ? Math.max(1, Math.ceil(naturalPx / pitchPx)) : 1;
}

/** Deterministic single-line height used to count docGrid cells for one East
 * Asian text run. Tabled fonts contribute their recorded design height.
 *
 * Word's Far-East single-line height is 1.3 × the font's hhea box: measured for
 * Yu Mincho in the issue #1013 sample-58 sweep and corroborated by Meiryo, whose
 * usWin sum is designed as 1.3 × its hhea box. The factor is font-independent
 * Word runtime behaviour, not an ECMA-376 rule (§17.6.5 does not define it).
 *
 * An untabled font's hhea box is unknown, so the fallback assumes 1.0em. That is
 * exact for the legacy full-frame Japanese fonts that dominate the untabled
 * corpus (MS Mincho / MS Gothic), and ceil-to-whole-cell allocation bounds the
 * assumption's effect. The resulting counts match the previous Word-measured
 * em rule (floor(em/pitch) + 1, before #1018) at every recorded untabled point:
 * 10.5pt and 12pt on 18pt → 1 cell (sample-35), 20pt on 20pt → 2 cells
 * (sample-9), and 11pt on 16.4pt → 1 cell. The rules diverge only in the
 * never-measured em/pitch ∈ (0.77, 1) band; here this fallback follows the
 * #1013 sweep direction (14–16pt on 18pt → 2 cells).
 *
 * Never use a substituted Canvas box here: its integer-rounded metrics are
 * font- and scale-dependent. Follow-up: replace the 1.0em assumption with
 * fontTools-extracted MS Mincho/MS Gothic hhea metrics in core WIN_METRICS when
 * those font binaries are available. */
export function eastAsianGridCountSinglePx(intendedSinglePx: number, emPx: number): number {
  return intendedSinglePx > 0 ? intendedSinglePx : emPx * 1.3;
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
 *   atLeast → max(natural, authored minimum, active grid minimum). For an
 *             explicitly authored value, the unsnapped tall-line result is an
 *             observed Windows Word compatibility behavior; it is not stated
 *             normatively by the line-grid clauses.
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
  // run's Word-faithful single-line height (a tabled run's design height, an
  // untabled East Asian run's 1.3em Word FE fallback). Used ONLY to count
  // docGrid cells for East Asian lines, so a substituted face's Canvas box
  // cannot change pagination or paint-scale cell allocation.
  gridCountSinglePx?: number,
  // px — untabled East Asian run em used only by direct/synthetic callers that
  // cannot provide the producer-computed per-line gridCountSinglePx.
  untabledEastAsianEmPx?: number,
): number {
  const glyphNatural = ascentPx + descentPx;
  // For `auto`/single spacing the multiplier applies to the intended font's
  // design line height (ECMA-376 §17.3.1.33). When the document's font is
  // substituted, the Canvas glyph extent (`glyphNatural`) understates that —
  // see font-metrics.ts. `base` restores the intended single-line height so
  // line spacing matches Word, while never dropping below the substituted
  // glyph extent (so glyphs are not clipped). Grid-snapped lines are governed
  // by the grid pitch instead, so the metric correction stays out of them.
  const natural = Math.max(glyphNatural, intendedSinglePx);
  const hasGrid = isGridLineRule(grid);
  const pitchPx = hasGrid ? grid!.linePitchPt! * scale : 0;
  // Per ECMA-376 §17.6.5, a paragraph whose `line` attribute is NOT
  // explicitly set — it only inherits from docDefault — snaps to one grid
  // pitch per text line in docGrid sections, regardless of the inherited
  // multiplier. Paragraphs that do set `line` on their pPr or a named style
  // multiply against the pitch as usual. This is what makes Word render
  // ESSAY (9 pt, no explicit line) at ~1 pitch (~18 pt) while a 1.33×
  // body paragraph with line="320" renders at pitch × 1.33 = ~24 pt.
  //
  // A single-spaced line on a docGrid snaps to whole grid CELLS in East Asian
  // text. The number of cells is derived from the line's DESIGN single-line
  // height (`gridCountSinglePx` — what Word measures: each run's real-font
  // single-line height), per the sample-58 adjudication of issue #1013; the
  // substituted Canvas glyph box is NOT used for the count (it can overstate a
  // tabled font's design height and add a spurious cell). See gridSingleCell and
  // docGridLineCells for the rule and measurements.
  // A Latin-only line is NOT cell-rounded — it keeps its natural height above
  // a one-cell floor (demo/sample-1: an 18pt heading on an 18pt pitch stays
  // ~20.7px, not 36). ECMA-376 Part 1 only defines the natural ≤ pitch case
  // (§17.6.5 / §17.3.1.32); the East-Asian cell rounding for taller lines is
  // Word runtime behaviour, so it is gated on the line's script.
  const gridSingleCell = (): number => {
    if (!eastAsian) return Math.max(glyphNatural, pitchPx);
    // Ruby lines reserve real furigana height (base + rt); honor the measured
    // glyph box so the annotation is not clipped. Plain EA lines snap their
    // design single-line height to whole cells.
    if (hasRuby) return Math.max(pitchPx, Math.ceil(glyphNatural / pitchPx) * pitchPx);
    // Word counts grid cells from the run's DESIGN single-line height (the real
    // font's metrics — §17.6.5), NOT the substituted Canvas glyph box, which can
    // overstate that height. A substitute face whose box lands a hair over the
    // pitch (e.g. Hiragino Mincho ProN standing in for Yu Mincho: an 18.17px box
    // vs Yu's 17.19px design height at 12pt on an 18pt pitch) must not push the
    // line into an extra cell — that mis-widths every tbRl column and shifts
    // vertical-section block tables by whole cells (sample-52). Prefer the
    // per-line design grid-count height when the caller supplies it — it is the
    // max over runs of each run's Word-faithful design height. Direct callers
    // may instead provide the untabled run em; tabled direct callers already
    // expose their design height through intendedSinglePx. A legacy caller with
    // neither design input gets one pitch, never the substituted Canvas box.
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
  // Word's native model has no such state: per the [MS-DOC] LSPD structure,
  // "exact" spacing is encoded as a negative dyaLine ("the line spacing, in
  // twips, is exactly 0x10000 minus dyaLine", so an exact 0 is unrepresentable)
  // and a non-negative dyaLine in twips mode is "dyaLine or the number of twips
  // necessary for single spacing, whichever value is greater" — i.e. a stored 0
  // resolves to exactly single spacing. Word's PDF export of sample-7 confirms
  // (those rows render at normal single-line height). Match that: treat
  // exact/auto line <= 0 as single spacing. (LSPD's max() rule is the twips
  // mode; applying the same fallback to a degenerate auto multiplier <= 0 is
  // the analogous non-collapsing reading.)
  if ((ls.rule === 'exact' || ls.rule === 'auto') && ls.value <= 0) {
    return hasGrid ? gridSingleCell() : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      if (inheritedOnly) return gridSingleCell();
      return Math.max(glyphNatural, pitchPx * ls.value);
    }
    return natural * ls.value;
  }
  if (ls.rule === 'exact') return ls.value * scale;
  if (ls.rule === 'atLeast') {
    // §17.18.48 establishes the authored minimum and §17.6.5 establishes the
    // grid pitch, but neither clause specifies how a tall, plain line with an
    // explicit atLeast value combines with whole-cell grid allocation. Windows
    // Word output leaves that line at its raw content height while later
    // ordinary lines remain on one pitch. Preserve that observed compatibility
    // behavior here. Ruby and inherited-only spacing retain the established
    // whole-cell path because their separate fixtures require it.
    return Math.max(
      natural,
      ls.value * scale,
      hasGrid ? (hasRuby || inheritedOnly ? gridSingleCell() : pitchPx) : 0,
    );
  }
  return natural;
}

/** Natural single-line height in px for an empty paragraph (no rendered text). */
export function emptyLineNaturalPx(fontSizePt: number, scale: number): { asc: number; desc: number } {
  return { asc: fontSizePt * scale * 0.8, desc: fontSizePt * scale * 0.2 };
}

/** Corrected single-line ascent/descent (px) from an ALREADY-measured
 *  `TextMetrics`: the Canvas `fontBoundingBox` (with the synthetic 0.8/0.2-em
 *  fallback when the engine reports none), rescaled to the document font's design
 *  line box via {@link correctLineMetrics}. The single source of truth for "how
 *  tall is one line of `family`", shared by the text-line path (layoutLines) and
 *  the empty paragraph-mark path (paragraphMarkLineHeight) so the two cannot
 *  drift — that drift (the empty path skipping `correctLineMetrics`) was the
 *  empty-paragraph under-measure bug (§17.3.1.29 / §17.3.1.33). `fallbackEmPx`
 *  sizes the synthetic box (the run's full size); `correctionEmPx` is the design
 *  size handed to `correctLineMetrics` — they differ only for smallCaps/vertAlign
 *  runs (where the text path keeps the full-size fallback) and coincide for a
 *  plain paragraph-mark line. The hhea single-line FLOOR for tabled fonts is
 *  applied separately by lineBoxHeight via {@link intendedSingleLinePx}. */
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
  para: DocParagraph,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>> = {},
): MarkLineMetrics {
  const fs = getDefaultFontSize(para);
  const authoredFamily = getDefaultFontFamily(para, eastAsian);
  const resolvedLocalFont = authoredFamily
    ? resolvedLocalFonts[normalizeLocalFontMetricFamily(authoredFamily)]
    : undefined;
  const measuredFamily = resolvedLocalFont?.family ?? authoredFamily;
  let asc: number;
  let desc: number;
  if (ctx) {
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
    // font does. correctedLineMetrics rescales a substituted font to the document
    // font's design box, identical to the text path; the hhea single-line floor
    // (intendedSingleLinePx, via emptyIntendedSinglePx below) then raises tabled
    // fonts — Latin included — to Word's line height.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, measuredFamily, fontFamilyClasses);
    const m = ctx.measureText(eastAsian ? 'あ' : 'x');
    ctx.font = prevFont;
    // A mark line carries no smallCaps/vertAlign, so fallback == correction size.
    ({ ascent: asc, descent: desc } = correctedLineMetrics(
      m, measuredFamily, fs * scale, fs * scale, eastAsian,
    ));
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  const intendedSingle = resolvedLocalFont
    ? fs * scale * resolvedLocalFont.lineHeightRatio
    : emptyIntendedSingleForScriptPx(para, scale, eastAsian);
  const gridCountSingle = eastAsian
    ? eastAsianGridCountSinglePx(intendedSingle, fs * scale)
    : undefined;
  const advancePx = lineBoxHeight(
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
  return { advancePx, ascentPx: asc, descentPx: desc };
}

export function paragraphMarkLineHeight(
  para: DocParagraph,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
  effectiveLineSpacing: LineSpacing | null = para.lineSpacing,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>> = {},
): number {
  return paragraphMarkLineMetrics(
    para, scale, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts,
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
 * whitespace such a paragraph may let overflow the bottom content edge, matching
 * Word's measured page fit (#981).
 *
 * NOTE: this stays the CENTRED baseline even though VISIBLE lineRule=auto content
 * lines now draw a PINNED baseline ({@link drawParagraphLine} / #990: multiplier
 * leading placed entirely below the glyphs). The pagination consumer is inkless
 * (mark-only), so the pinned glyph baseline never reaches it, and the #981 page
 * fit was verified against Word with this half-leading extent — changing it would
 * move page boundaries. The pin is therefore intentionally DRAW-ONLY.
 */
export function lineBelowBaselinePx(advancePx: number, ascentPx: number, descentPx: number): number {
  return Math.max(0, (advancePx - ascentPx + descentPx) / 2);
}

export function paragraphMarkBelowBaselinePt(
  para: DocParagraph,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian: boolean,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | undefined,
  fontFamilyClasses: Record<string, string>,
  effectiveLineSpacing: LineSpacing | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>> = {},
): number {
  // Measured at scale 1 so the returned px value is already in points.
  const m = paragraphMarkLineMetrics(
    para, 1, grid, paraHasRuby, eastAsian, ctx, fontFamilyClasses, effectiveLineSpacing,
    resolvedLocalFonts,
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
 * `csItalic = italicCs ?? false`). Adjudicated in issue #937 against Word: the
 * numbered Arabic headings in sample-7 carry `w:rtl`+`w:cs`+`w:b` WITHOUT
 * `w:bCs` and Word draws them at regular weight; sample-41's cs-italic Case A/C
 * carry `w:i` without `w:iCs` and render upright, matching the explicit-OFF
 * Case B (only Case D's plain Latin `w:i` renders italic).
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

export function findNearbyFontSize(runs: DocRun[], idx: number): number {
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

const MATH_SPACED_OPERATORS = new Set(['+', '-', '−', '=', '±', '×', '÷']);

function mathRunPlainText(text: string): string {
  return MATH_SPACED_OPERATORS.has(text) ? ` ${text} ` : text;
}

export function mathPlainText(nodes: MathNode[]): string {
  const renderNode = (node: MathNode): string => {
    switch (node.kind) {
      case 'run':
        return mathRunPlainText(node.text);
      case 'fraction':
        return `${mathPlainText(node.num)}/${mathPlainText(node.den)}`;
      case 'sup':
        return `${mathPlainText(node.base)}^${mathPlainText(node.sup ?? [])}`;
      case 'sub':
        return `${mathPlainText(node.base)}_${mathPlainText(node.sub ?? [])}`;
      case 'subSup':
        return `${mathPlainText(node.base)}_${mathPlainText(node.sub ?? [])}^${mathPlainText(node.sup ?? [])}`;
      case 'nary':
        return `${node.op}${mathPlainText(node.sub ?? [])}${mathPlainText(node.sup ?? [])}${mathPlainText(node.body)}`;
      case 'delimiter':
        return `${node.begChar}${node.items.map(mathPlainText).join(',')}${node.endChar}`;
      case 'radical':
        return `${node.index && node.index.length > 0 ? mathPlainText(node.index) : ''}√${mathPlainText(node.radicand)}`;
      case 'limit':
        return `${mathPlainText(node.base)}${mathPlainText(node.lower ?? [])}${mathPlainText(node.upper ?? [])}`;
      case 'array':
        return node.rows.map((row) => row.map(mathPlainText).join(' ')).join(' ');
      case 'groupChr':
        return `${node.char}${mathPlainText(node.base)}`;
      case 'bar':
      case 'box':
      case 'borderBox':
        return mathPlainText(node.base);
      case 'accent':
        return `${node.char}${mathPlainText(node.base)}`;
      case 'func':
        return `${mathPlainText(node.name)}(${mathPlainText(node.arg)})`;
      case 'group':
        return mathPlainText(node.items);
      case 'phant':
        return node.show ? mathPlainText(node.base) : '';
      case 'sPre':
        return `${mathPlainText(node.sub)}${mathPlainText(node.sup)}${mathPlainText(node.base)}`;
    }
  };
  return nodes.map(renderNode).join('').replace(/[ \t]{2,}/g, ' ');
}

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
function extendThroughTrailingIdeographicSpaces(chars: string[], split: number): number {
  if (split <= 0) return split;
  let s = split;
  while (s < chars.length && chars[s] === '\u3000') s++;
  return s;
}

export function fitCJKPrefix(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
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
): string {
  const chars = [...text]; // spread handles surrogate pairs
  // Trailing IDEOGRAPHIC SPACE (U+3000) line-end allowance: a candidate that
  // overflows ONLY because it ends in fullwidth spaces still fits — the spaces
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
    while (visibleEnd > 0 && chars[visibleEnd - 1] === '\u3000') visibleEnd--;
    const prefix = chars.slice(0, visibleEnd).join('');
    if (prefix.length === 0) return true;
    const advance = textAdvanceWidth(
      ctx.measureText(prefix).width + (verticalRun ? verticalRunInkExtraPx(ctx, prefix) : 0),
      prefix,
      gridDeltaPx,
      charScale,
      charSpacingPx,
    );
    return advance <= maxWidth;
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
 * Digits / spaces / punctuation (neutral, non-cs) attach to the PRECEDING slice
 * so a number embedded in Arabic ("نص 12 نص") does not fragment the word into
 * extra segments — Word keeps such weak/neutral characters with the script they
 * border. A leading neutral run takes the first strong slice's class.
 */
export function splitByComplexScript(text: string): { text: string; cs: boolean }[] {
  const out: { text: string; cs: boolean }[] = [];
  let curCs: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    // Neutral (non-letter) characters do not switch the active class; they ride
    // with whatever script is currently open (or the next one if none yet).
    const isLetter = /\p{L}/u.test(ch);
    if (!isLetter) {
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
 * a digit between two ideographs is Latin/ascii either way (Word renders ASCII
 * digits with the ascii face), and a single fillText anchors to the cumulative
 * whole-string advance, so the visible spacing is unchanged.
 *
 * NOTE: this split decides the FONT slot only. Whether a resulting segment is
 * snapped to the §17.6.5 character grid is decided SEPARATELY by the grid's own
 * `EAST_ASIAN_RE` purity test (see `gridSegDeltaPx`/`eaGlyphCount`), not by the
 * `ea` flag here. The two CJK predicates classify slightly different code-point
 * sets; correctness of the grid total relies on `eaGlyphCount` being additive
 * over this partition (covered by docgrid-char.test.ts's mixed-token case). Keep
 * them independent — do not "unify" the predicates without re-checking that test.
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
 * then lays out right-to-left as Word does.
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
 *  this via {@link resolveDefaultTabPt}. Shared by the line layout
 *  (`layoutLines`) and the numbered-list marker's trailing-tab advance
 *  (`renderParagraph`). */
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

/** ECMA-376 §17.3.1.37 (tabs) + §17.15.1.25 (defaultTabStop) — advance from a
 *  pen position to the next effective tab stop, ALL in text-margin px.
 *
 *  The effective stop set is the custom stops PLUS automatic left-aligned stops
 *  at every multiple of `intervalPx` that occurs AFTER all custom stops
 *  (§17.15.1.25: "Automatic tab stops refer to the tab stop locations which
 *  occur after all custom tab stops"; the spec example puts a 0.25" grid at
 *  2.5"/2.75"/3.0" past a 2.28" custom stop). A tab moves to the smallest
 *  effective stop strictly greater than `curMarginPx`.
 *
 *  @param curMarginPx pen position, measured from the text margin.
 *  @param customStopsPx custom stops in margin-px (`pos = tabStop.pos * scale`),
 *    order not assumed; each carries its own alignment + leader.
 *  @param intervalPx automatic-stop interval = `defaultTabPt * scale`.
 *  @returns the chosen stop (custom keeps its alignment/leader; an automatic
 *    stop is `'left'` with no leader), or `null` only when no stop exists. */
export function nextTabStop(
  curMarginPx: number,
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  intervalPx: number,
): { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null {
  // Candidate 1 — the nearest custom stop strictly past the pen. The spec set is
  // unordered, so scan for the minimum rather than assuming sort order.
  let custom: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null = null;
  let maxCustomPx = 0;
  for (const t of customStopsPx) {
    if (t.pos > maxCustomPx) maxCustomPx = t.pos;
    if (t.pos > curMarginPx && (custom === null || t.pos < custom.pos)) custom = t;
  }

  // Candidate 2 — the nearest automatic multiple of the interval that is BOTH
  // past the pen AND past the last custom stop (§17.15.1.25: automatic stops
  // begin only after all custom stops). EPS forces a pen sitting exactly on a
  // boundary to advance to the NEXT one.
  let auto: { pos: number; alignment: TabStop['alignment'] } | null = null;
  if (intervalPx > 0) {
    const EPS = 1e-6;
    const from = Math.max(curMarginPx, maxCustomPx);
    let pos = Math.ceil((from + EPS) / intervalPx) * intervalPx;
    // Guard against floating-point landing on/under the pen (rounding can yield a
    // boundary that is not strictly greater): step to the next multiple.
    if (pos <= curMarginPx) pos += intervalPx;
    auto = { pos, alignment: 'left' };
  }

  if (custom && auto) {
    // Ties → the custom stop wins so its alignment/leader are honoured.
    return custom.pos <= auto.pos ? custom : auto;
  }
  return custom ?? auto;
}

/** ECMA-376 §17.3.1.37 / §17.15.1.25 in a BIDI paragraph — the RTL twin of
 *  {@link nextTabStop}. In a right-to-left paragraph the tab-stop coordinate
 *  system is anchored at the LEADING (right) text edge: a stop's `pos` is the
 *  distance from that right edge, and the pen advances LEFTWARD (decreasing
 *  margin position). This finds the nearest stop strictly BELOW the pen — i.e.
 *  the next stop to the left — using the same automatic-grid rule (§17.15.1.25:
 *  automatic stops sit past all custom stops, here on the far LEFT of the line).
 *
 *  @param curMarginPx pen position measured from the right (leading) text edge.
 *  @param customStopsPx custom stops in right-edge px (`pos * scale`).
 *  @param intervalPx automatic-stop interval = `defaultTabPt * scale`.
 *  @returns the chosen stop (custom keeps its alignment/leader; automatic is
 *    `'left'`), or `null` when the pen is already at/left of every stop. */
export function nextTabStopRtl(
  curMarginPx: number,
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  intervalPx: number,
): { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null {
  // Candidate 1 — the nearest custom stop strictly PAST the pen (greater pos =
  // further from the right edge = further left). Symmetric to nextTabStop.
  let custom: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null = null;
  let maxCustomPx = 0;
  for (const t of customStopsPx) {
    if (t.pos > maxCustomPx) maxCustomPx = t.pos;
    if (t.pos > curMarginPx && (custom === null || t.pos < custom.pos)) custom = t;
  }
  // Candidate 2 — nearest automatic multiple past BOTH the pen and the last
  // custom stop (§17.15.1.25). Identical grid to nextTabStop; the mirror is
  // entirely in how `pos` is interpreted (distance from the right edge).
  let auto: { pos: number; alignment: TabStop['alignment'] } | null = null;
  if (intervalPx > 0) {
    const EPS = 1e-6;
    const from = Math.max(curMarginPx, maxCustomPx);
    let pos = Math.ceil((from + EPS) / intervalPx) * intervalPx;
    if (pos <= curMarginPx) pos += intervalPx;
    auto = { pos, alignment: 'left' };
  }
  if (custom && auto) return custom.pos <= auto.pos ? custom : auto;
  return custom ?? auto;
}

/** One entry in a bidi line's LOGICAL-order sequence, for {@link layoutBidiTabStops}. */
export interface BidiTabItem {
  /** True for a tab segment (its width is (re)computed); false for content. */
  isTab: boolean;
  /** Content width in px (ignored for tabs). Set by the LTR layout pass. */
  width: number;
}

/** Per-segment result of {@link layoutBidiTabStops}. */
export interface BidiTabResult {
  /** New measuredWidth for the tab at this LOGICAL index (non-tabs: unchanged). */
  width: number;
  /** Leader to paint across this tab's span (`'none'`/undefined ⇒ blank). */
  leader?: TabStop['leader'];
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
 * distance from that MARGIN — not from the paragraph's indented edge (verified
 * against Word: sample-28's TOC2 title tab at 1017 twips lands 50.85pt from the
 * page margin even though the paragraph carries a 36pt logical-left indent).
 * Content begins at `startPenPx` (the paragraph's leading-indent + first-line
 * indent); the Nth tab in reading order advances to the next stop further left
 * (larger `pos`), exactly like the LTR pen advances rightward through stops.
 * Alignment is logical (Part 4 §14.11.2): physical `left` = `start` (leading ⇒
 * following content's leading/RIGHT edge on the stop), physical `right` = `end`
 * (trailing ⇒ its trailing/LEFT edge on the stop); `center` is unchanged;
 * `bar`/`clear` advance like `start`. Automatic stops fall on the §17.15.1.25
 * grid from the margin, after all custom stops. A stop past the LEFT text
 * margin (`leftLimitPx`) pins its content to that margin (Word never pushes it
 * off the page — sample-28's page numbers pin at x=72).
 *
 * The widths returned here reproduce Word's layout through the draw loop's
 * VISUAL walk because {@link computeLineVisualOrder} classifies tabs as UAX#9 S:
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
  const followW = (from: number): number => {
    let w = 0;
    for (let j = from; j < n; j++) {
      if (items[j].isTab) break;
      w += width[j];
    }
    return w;
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
    const fw = followW(i + 1);
    let target: number; // pen value after the tab (its trailing/left edge)
    if (stop.alignment === 'right') {
      // end/trailing: following content's TRAILING (left) edge on the stop, i.e.
      // its right edge sits fw to the RIGHT (smaller margin) of the stop.
      target = stop.pos - fw;
    } else if (stop.alignment === 'center') {
      target = stop.pos - fw / 2;
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
 *  a FRESH object (fresh Sets) on every call, and the prebuiltPages production
 *  path (DocxDocument.renderPage) resolves it independently in paginateDocument
 *  and in renderDocumentToCanvas — same `doc.settings`, different references.
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

/** True when {@link buildSegments} would produce DIFFERENT segments for this
 *  paragraph under the paginator's measure state versus the paint state — i.e.
 *  the segment text depends on per-page render context rather than on the runs
 *  alone. Exactly two sources exist today:
 *    - `field` runs of type page / numPages: {@link resolveFieldText} returns
 *      `environment.pageIndex + 1` / `environment.totalPages`, and the measure environment is
 *      frozen at pageIndex 0 / totalPages 1 (buildMeasureState);
 *    - `noteRef` text runs: the label resolves via `environment.noteNumbers` /
 *      `environment.currentNoteNumber`, which only the paint environment carries
 *      (renderDocumentToCanvas builds the map; the measure state never does).
 *  Such a paragraph must NOT stamp its measured lines: the stamped segments
 *  would carry the measure-time text (a stale page number / note label) and the
 *  paint pass would draw it verbatim. Skipping the stamp keeps those paragraphs
 *  on the recompute path, which resolves fields against the real page context —
 *  the pre-reuse behaviour. Extend this predicate if buildSegments ever gains a
 *  new state-dependent text source. */
export function paragraphSegsStateSensitive(para: DocParagraph): boolean {
  for (const run of para.runs) {
    if (run.type === 'field') {
      const ft = (run as unknown as FieldRun).fieldType;
      // page / numPages depend on the per-page render context. DATE / TIME can
      // depend on `environment.currentDateMs`; although LineLayoutEnvironment
      // permits that value, current pagination environments may omit it. Keep
      // these fields on the paint-time recompute path so a pagination fallback
      // instant is never stamped as the final field text (§17.16.4.1).
      if (ft === 'page' || ft === 'numPages' || ft === 'date' || ft === 'time') return true;
    } else if (run.type === 'text' && (run as unknown as DocxTextRun).noteRef) {
      return true;
    }
  }
  return false;
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

export function buildSegments(runs: DocRun[], environment: LineLayoutEnvironment): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const resolvedFont = (family: string | null | undefined): ResolvedLocalFontMetric | undefined =>
    family ? environment.resolvedLocalFonts?.[normalizeLocalFontMetricFamily(family)] : undefined;
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
    base: DocxTextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
    sourceRunIndex: number,
    sourceFragmentIndex?: number,
    sourceTextStart?: number,
  ) => {
    const firstOutputIndex = segs.length;
    // §17.3.2.33 small caps are sized per character: lowercase LETTERS render two
    // points smaller, uppercase letters and non-alphabetic characters at the full
    // run size. `reduced` (set per case-piece in the loop below) carries that onto
    // each emitted segment; calcEffectiveFontPx shrinks only the reduced ones.
    // allCaps (§17.3.2.5) and non-caps runs are a single, non-reduced piece.
    let reduced = false;
    // Ruby annotation rides with the WHOLE base text (typically 1-2 chars).
    // Splitting on word boundaries would lose the association, so attach
    // the annotation only to the first emitted segment.
    const baseRuby = (base as DocxTextRun).ruby;
    const ruby = baseRuby
      ? {
          text: baseRuby.text,
          fontSizePt: baseRuby.fontSizePt,
          ...(baseRuby.hpsRaisePt != null ? { hpsRaisePt: baseRuby.hpsRaisePt } : {}),
        }
      : undefined;
    const revision = (base as DocxTextRun).revision;
    const r = base as DocxTextRun;
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

    // ECMA-376 §17.3.2.26 content classification. A run with `w:rtl`
    // (§17.3.2.30) or the `<w:cs/>` toggle (§17.3.2.7) applies complex-script
    // formatting to ALL of its characters; otherwise each character is routed by
    // its Unicode block (Arabic/Hebrew/... → cs; Latin/digits/CJK → ascii/hAnsi).
    // NOTE rFonts@cs (fontFamilyCs) alone is just a font SLOT and must NOT
    // force cs — e.g. sample-1's Heading1 (Latin) has cstheme + szCs=52 but
    // renders at w:sz=24; forcing cs blew its size up to 26pt.
    const forceCs = r.rtl === true || r.cs === true;

    // Complex-script (cs) formatting sources. SIZE (§17.3.2.39 szCs) and TYPEFACE
    // (§17.3.2.26 rFonts@cs) fall back to their Latin counterpart when absent —
    // the parser resolves szCs through the full style chain, mirroring a
    // directly-set `w:sz` per §17.3.2.18. But BOLD (§17.3.2.3 bCs) and ITALIC
    // (§17.3.2.17 iCs) are INDEPENDENT toggles: an absent `bCs`/`iCs` defaults
    // OFF and must NOT inherit the Latin-axis `w:b`/`w:i` (which govern only
    // non-complex content). Adjudicated in issue #937 against Word ground truth —
    // sample-41 Case A/C (`w:i`, no `w:iCs`) render upright like the explicit-OFF
    // Case B (contrast Case D's plain Latin `w:i` = italic); sample-7's
    // `w:rtl`+`w:cs`+`w:b` (no `w:bCs`) Arabic headings render at regular weight.
    const csFontSize = r.fontSizeCs ?? base.fontSize;
    const csFontFamily = r.fontFamilyCs ?? base.fontFamily;
    const csBold = r.boldCs ?? false;
    const csItalic = r.italicCs ?? false;

    // ECMA-376 §17.3.2.26 eastAsia axis. Within a non-complex-script slice, CJK
    // code points take the eastAsia face while Latin/digits keep the ascii face
    // (`base.fontFamily`). Only `DocxTextRun` carries the axis; absent (field
    // runs / single-axis parser output) ⇒ fall back to ascii. Text-box runs feed
    // this same builder (via `shapeRunToDocRun`), so a text box's per-script face
    // is picked here too. Bold/italic/size are NOT axis-specific here — eastAsia
    // shares the Latin (non-cs) toggles, so only the family differs.
    const eaFontFamily = (base as DocxTextRun).fontFamilyEastAsia ?? base.fontFamily;

    // Word classifies European digits in an Arabic/Hebrew complex-script run as
    // AN (§17.3.2.20 w:lang w:bidi): use the bidi language's primary subtag when
    // present, else fall back to the run being rtl-marked.
    const digitsAsAN =
      (forceCs || r.rtl === true) && isRtlBidiLang(r.langBidi, r.rtl === true);

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
    const pushSeg = (text: string, cs: boolean, fontFamily: string | null) => {
      const bold = cs ? csBold : base.bold;
      const italic = cs ? csItalic : base.italic;
      const localFont = resolvedFont(fontFamily);
      const localEaFloor = resolvedFont(eaFontFamily);
      // Local-metric aliases intentionally register the normal face only. A
      // bold/italic Canvas request against that alias would synthesize the
      // style instead of resolving the installed family's real styled face.
      // Keep the authored family for styled paint/measurement, while retaining
      // the family-level design line ratio used by Word's line-box calculation.
      const localPaintFamily = !bold && !italic ? localFont?.family : undefined;
      const localEaFloorFamily = !bold && !italic ? localEaFloor?.family : undefined;
      segs.push({
        text,
        bold,
        italic,
        underline: base.underline,
        // §17.3.2.40 underline style / colour — carried only on DocxTextRun (a
        // FieldRun draws single). Kept raw ST_Underline; the renderer normalizes
        // to DrawingML §20.1.10.82 at draw time.
        underlineStyle: (base as DocxTextRun).underlineStyle,
        underlineColor: (base as DocxTextRun).underlineColor,
        strikethrough: base.strikethrough,
        fontSize: cs ? csFontSize : base.fontSize,
        color: base.color,
        fontFamily: localPaintFamily ?? fontFamily,
        resolvedLineHeightRatio: localFont?.lineHeightRatio,
        vertAlign,
        measuredWidth: 0,
        smallCaps: reduced,
        joinPrev: gluePending ? true : undefined,
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
        rtl,
        digitsAsAN: digitsAsAN ? true : undefined,
        // §17.3.2.26 declared eastAsia axis — recorded for the text-box line-box
        // floor only (see LayoutTextSeg.eaFloorFamily). Inert for the body path.
        eaFloorFamily: localEaFloorFamily ?? eaFontFamily,
        resolvedEaFloorLineHeightRatio: localEaFloor?.lineHeightRatio,
        // IX1 — resolved hyperlink target of the originating run, for the
        // text-layer clickable overlay. Does not affect layout or drawing.
        hyperlink,
        snapToCharacterGrid: r.snapToGrid !== false,
        // WD4 — run character metrics (§17.3.2.35 spacing / §17.3.2.43 w /
        // §17.3.2.24 position / §17.3.2.19 kern). Uniform across the run, so
        // every emitted segment carries the same values; the measure and paint
        // passes apply them identically (measure==paint).
        charSpacing: r.charSpacing,
        charScale: r.charScale,
        fitTextVal: fitTextRegionIndex === undefined ? undefined : r.fitTextVal,
        fitTextId: fitTextRegionIndex === undefined ? undefined : r.fitTextId,
        fitTextRegionIndex,
        fitTextRunIndex: fitTextRegionIndex === undefined ? undefined : fitTextFragmentEntryIndex,
        position: r.position,
        kerning: r.kerning,
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
          pushSeg(part.text, cs, part.mapped ? null : fontFamily);
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
          for (const slice of splitByComplexScript(word)) {
            if (slice.cs) emit(slice.text, 'cs');
            else emitNonCs(slice.text);
          }
        }
      }
    }

    const emitted = segs
      .slice(firstOutputIndex)
      .filter((segment): segment is LayoutTextSeg => 'text' in segment);
    if (
      sourceTextStart !== undefined
      && emitted.map((segment) => segment.text).join('') === text
    ) {
      let offset = sourceTextStart;
      const refs = (base as DocxTextRun).sourceRefs;
      for (const segment of emitted) {
        const sourceRefs = sliceTextSourceRefs(refs, offset, offset + segment.text.length);
        if (sourceRefs.length > 0) segment.sourceRefs = sourceRefs;
        offset += segment.text.length;
      }
    }
  };

  for (const [runIndex, run] of runs.entries()) {
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
              : environment.currentNoteNumber)
          : undefined;
      if (t.noteRef) {
        const label = noteText != null ? String(noteText) : (t.text || '');
        if (label.length > 0) {
          pushTextPiece(label, t, t.vertAlign ?? 'super', runIndex, 0);
        }
        continue;
      }
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      let sourceTextStart = 0;
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) {
          pushTextPiece(parts[i], t, t.vertAlign, runIndex, i, sourceTextStart);
        }
        sourceTextStart += parts[i].length;
        if (i < parts.length - 1) {
          segs.push({ isTab: true, fontSize: t.fontSize, measuredWidth: 0, bold: t.bold, italic: t.italic });
          sourceTextStart += 1;
        }
      }
    } else if (run.type === 'image') {
      const img = run as unknown as ImageRun & { type: 'image' };
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
      // measurement/split path) but carrying the ChartModel, which the draw site
      // paints with the shared `renderChart`.
      //
      // A `<wp:anchor>` (floating) chart (§20.4.2.3) carries `anchor: true` and
      // its parsed page-offset fields, exactly like an anchor ImageRun: the
      // measure pass zeroes an anchor seg's width (it is not part of the inline
      // flow) and `renderAnchorImages` draws it at the resolved absolute box.
      const chartRun = run as unknown as import('./types').ChartRun & { type: 'chart' };
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
        chart: chartRun.chart,
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
      if (text) pushTextPiece(text, f, f.vertAlign, runIndex);
    } else if (run.type === 'math') {
      // The parser resolves the paragraph font size; fall back to a nearby run only
      // if it is somehow absent.
      const fontSize = run.fontSize || findNearbyFontSize(runs, runs.indexOf(run));
      segs.push({
        mathNodes: run.nodes,
        display: run.display,
        fontSize,
        color: null,
        fallbackText: mathPlainText(run.nodes),
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
      const localFont = resolvedFont(authoredFamily);
      const localEaFloor = resolvedFont(run.fontFamilyEastAsia ?? null);
      const normalFace = !bold && !italic;
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
        fontFamily: (normalFace ? localFont?.family : undefined) ?? authoredFamily,
        resolvedLineHeightRatio: localFont?.lineHeightRatio,
        vertAlign: null,
        measuredWidth: 0,
        eaFloorFamily:
          (normalFace ? localEaFloor?.family : undefined) ?? run.fontFamilyEastAsia ?? null,
        resolvedEaFloorLineHeightRatio: localEaFloor?.lineHeightRatio,
        snapToCharacterGrid: false,
      });
    }
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

  // ── UAX #14 no-break pairs (LB14/LB23/LB23a/LB24/LB25/LB28/LB30) ──
  // buildSegments intentionally splits at run / font-script boundaries, but
  // those formatting seams are not line-break opportunities. Mark the following
  // segment so layoutLines' existing atomic-group pre-flush selects the previous
  // real opportunity instead. The shared predicate is deliberately one-way:
  // false means unsupported/deferred, never "break allowed".
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev || cur.text.length === 0) continue;
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

  return segs;
}

export function layoutLines(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
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
  // ECMA-376 §17.6.5 docGrid CHARACTER grid: per-EA-glyph cell delta in px (0
  // when inactive). Folded into every text advance so line breaking packs the
  // grid's char count per line; the draw paths add the SAME delta.
  gridDeltaPx = 0,
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
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[] = [];
  let currentWidth = 0;
  // Sum of trailing-space widths of every text token on the current line.
  // Used for two things:
  //   1. Knuth-Plass-style shrink tolerance: every line may consume up to
  //      SPACE_SHRINK_RATIO of its accumulated trailing-space width.
  //   2. Trailing-space collapse at line end — the last token's trailing
  //      space disappears when no further word is added, so when deciding
  //      whether a candidate word fits we treat it as if it would become the
  //      final word (its own trailing spaces collapsible).
  let lineTotalTrailingW = 0;
  // Incremental Canvas-vs-Word bias of the text already committed to this line.
  // Candidate checks add only the prospective text, avoiding a hot-loop rescan.
  let lineBiasBudget = 0;
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let lineDescent = 0;  // px
  let lineIntendedSingle = 0; // px — max intended single-line height on the line
  let lineGridCountSingle = 0; // px — max over segments of (tabled design height | untabled box)
  let lineVisibleAscent = 0;
  let lineVisibleDescent = 0;
  let lineVisibleIntendedSingle = 0;
  let lineHasVisibleMetrics = false;
  let isFirst = true;
  // Effective width/offset for the current line after float exclusion.
  let lineMaxWidth = maxWidth;
  let lineXOffset = 0;
  let currentLineTopY = wrapCtx?.startPageY ?? 0;

  // Minimum clear side-space (px) a CONTENT line needs before it may START beside
  // a float rather than flow below the band. Word's measured rule is 1 inch
  // (wordMinLineStartPx(scale) — the 1-inch requirement less a one-twip rounding
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
  // for its first line. Word keeps either mark beside a float down to a sub-inch
  // gap and drops it below only for a full-width band. #676 over-generalized 1
  // inch onto marks and pushed following content to the next page. Inline
  // content (including a content paragraph's trailing-break final line) keeps
  // the 1-inch rule.
  const minLineStartWidth = (): number => wordMinLineStartPx(scale);
  const isParagraphMarkOnlyFlow = segs.length > 0 && segs.every((segment) =>
    ('text' in segment && segment.metricOnly === true)
    || ('imagePath' in segment && Boolean(segment.anchor)),
  );

  // Compute wrap constraints for a new line about to start. Mutates
  // lineXOffset/lineMaxWidth/currentLineTopY. `minWidth` is the smallest clear
  // side-space the upcoming line must have to START here (minLineStartWidth() —
  // Word's 1-inch rule, §676); a free gap narrower than this is treated as
  // unusable and the line is sent below the intervening float(s), which is how
  // Word flows a line that cannot start beside a floating object (there is no
  // ECMA-376 §x.x.x for this trigger — see resolveLineFloatWindow / issue #676).
  const startLine = (minWidth: number = 0): void => {
    lineBiasBudget = 0;
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    // Small fixed probe height for float intersection (matches the historical
    // wrap behaviour for the topAndBottom skip and horizontal-gap scan).
    const probeH = 10 * scale;
    if (wrapCtx.lineWindow) {
      const win = wrapCtx.lineWindow({
        topYPt: currentLineTopY,
        minimumStartWidthPt: minWidth,
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
      const win = resolveLineFloatWindow(
        currentLineTopY, minWidth, probeH, wrapCtx.paraX, maxWidth, wrapCtx.floats,
        wrapCtx.columnXPt, wrapCtx.columnXPt + wrapCtx.columnWidthPt,
      );
      currentLineTopY = win.topY;
      lineXOffset = win.xOffset;
      lineMaxWidth = win.maxWidth;
    }
  };

  const availW = () => lineMaxWidth - (isFirst ? firstIndent : 0);

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
    // Nth-reachable stop, exactly like Word's pen. Do NOT feed the VISUAL
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
  // grapheme-fill scripts (Myanmar/Tibetan, #961) are excluded because the
  // issue #991 ground truth covers only the dictionary scripts and their
  // Word-verified wrap is per-cluster greedy. Gates the trailing-space shrink
  // budget: Word-observed (issue #991 — the calibration fixture's
  // 21-paragraph overflow sweep at 5/9/13 inter-phrase spaces) shows Word
  // wraps such a line's final word at natural fit for EVERY overflow > 0,
  // i.e. it never compresses inter-word spaces on Thai lines, while the Latin
  // demo evidence for SPACE_SHRINK_RATIO (sample-1 p3/p6) and the CJK centred
  // title (sample-10) keep the 25% drawable budget.
  let lineHasSea = false;
  const flush = (
    forceHeight?: number,
    brTerminated = false,
    nextStart?: LineBoundary,
  ) => {
    applyBidiTabs();
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
    lineBiasBudget = 0;
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

  const addToLine = (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
    trailingSpaceW: number = 0,
  ) => {
    currentLine.push(s);
    currentWidth += w;
    lineTotalTrailingW += trailingSpaceW;
    if ('text' in s) {
      lineBiasBudget += biasBudgetContribution(s);
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
    if (!('isTab' in s) && !('imagePath' in s) && !('mathNodes' in s)) {
      const ts = s as LayoutTextSeg;
      if (ts.ruby) lineHasRuby = true;
      if (ts.seaBreaks !== undefined && isDictionarySeaText(ts.text)) lineHasSea = true;
      const metricEastAsian = ts.metricEastAsian === true || EAST_ASIAN_RE.test(ts.text);
      if (!lineEastAsian && metricEastAsian) lineEastAsian = true;
      // Intended single-line height for fonts whose substituted Canvas metrics
      // understate Word's line spacing (font-metrics.ts). 0 for untabled fonts.
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
      const intended = segmentIntendedSingleLinePx(ts, intendedEm, segScriptHint);
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
      if (paintsInlineInk && intended > lineVisibleIntendedSingle) {
        lineVisibleIntendedSingle = intended;
      }
      // Only East Asian text is cell-rounded. Both branches are scale-linear:
      // tabled fonts use recorded design metrics; untabled fonts use 1.3em.
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
  // opts in). Kerning is enabled only when the run declares `w:kern` AND its font
  // size is at or above the threshold (the spec's "smallest font size which shall
  // have its kerning automatically adjusted"). A run that does not opt in leaves
  // `ctx.fontKerning` at its inherited value rather than being forced off — Word's
  // hierarchy default is off, but the browser default `'auto'` already produced
  // the ±1–2px body-text behaviour the existing references were captured against;
  // forcing `'none'` document-wide is a separate decision measured against the
  // Word PDFs (see the WD4 report), not made here. `setSegKerning` mirrors the
  // paint-side `paintSegKerning` in renderer.ts EXACTLY.
  const setSegKerning = (s: LayoutTextSeg): CanvasFontKerning | null => {
    if (s.kerning == null) return null;
    const prev = ctx.fontKerning;
    ctx.fontKerning = s.fontSize >= s.kerning ? 'normal' : 'none';
    return prev;
  };
  const restoreKerning = (prev: CanvasFontKerning | null): void => {
    if (prev != null) ctx.fontKerning = prev;
  };

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
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
    const prevKern = setSegKerning(s);
    try {
      return verticalRunInkExtraPx(ctx, text);
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
      seg.seaBreaks = seaMixedBreakOffsets(seg.text, { cjk: true, kinsoku });
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
  const segAdvance = (s: LayoutTextSeg): number =>
    segAdvanceWidth(s, measureText(s).width + verticalInkExtra(s, s.text), gridDeltaPx, scale);
  // Grid advance of an arbitrary substring under a segment's font (for split
  // prefixes/tails). Selects the font (and the run's kerning state), then applies
  // the same width model as a whole segment BUT with the substring's own
  // text/length so char-spacing scales with the piece — the split-prefix vs
  // whole-segment advances must agree.
  const strAdvance = (s: LayoutTextSeg, text: string): number => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    const prevKern = setSegKerning(s);
    const natural = ctx.measureText(text).width;
    restoreKerning(prevKern);
    return segAdvanceWidth({ ...s, text }, natural + verticalInkExtra(s, text), gridDeltaPx, scale);
  };

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => {
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('imagePath' in q) return q.widthPt * scale;
    if ('mathNodes' in q) return q.measuredWidth || 0;
    if ('lineBreak' in q) return 0;
    return segAdvance(q);
  };

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph that new line is an EMPTY line that still
  // occupies one line height (Word reserves it; visible e.g. as extra table
  // row height). Track the trailing break so it can be flushed after the loop.
  let trailingBreakFontSize: number | null = null;

  // Establish the first line's wrap window now that the content queue exists.
  startLine(
    isParagraphMarkOnlyFlow
      ? (wrapCtx?.paragraphMarkLineStartWidth ?? minLineStartWidth())
      : minLineStartWidth(),
  );

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
            } else if ('mathNodes' in q) {
              addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
            } else {
              const m = measureText(q);
              // #1014 — fold the vo=Tr ink deficit into the committed advance too.
              const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), gridDeltaPx, scale);
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
      // Convert the chosen margin-space stop back to paraX-relative px.
      const stopParaX = stop ? stop.pos - tabOriginPx : absFromParaX;
      // Right/center/decimal tab: place the tab + its trailing content (up to the next
      // tab / line end) so the content ends at / centers on the stop, and commit that
      // content directly so the normal wrap check doesn't push it past the stop
      // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
      // Automatic stops returned by nextTabStop are left-aligned, so they fall
      // through to the left-tab path below.
      if (stop && stop.alignment !== 'left' && stop.alignment !== 'bar' && stop.alignment !== 'clear') {
        const stopX = stopParaX;
        seg.leader = stop.leader;
        let followW = 0;
        for (const q of queue) {
          if ('isTab' in q || 'lineBreak' in q) break;
          followW += tabFollowWidth(q);
        }
        const frac = stop.alignment === 'center' ? 0.5 : 1;
        let tabW = stopX - absFromParaX - followW * frac;
        if (tabW <= 0) tabW = seg.fontSize * scale * 0.25;
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
          } else if ('mathNodes' in q) {
            addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
          } else {
            const m = measureText(q);
            // #1014 — fold the vo=Tr ink deficit into the committed advance too.
            const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), gridDeltaPx, scale);
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
    if ('mathNodes' in seg) {
      const render = mathRenders.get(seg.mathNodes);
      if (!render) {
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
    const m = measureText(s);
    // Advance = natural width + character-grid delta (the SINGLE model shared
    // with the draw paths; 0 unless an active grid AND a pure-EA segment).
    // #1014 — plus the vo=Tr rotate-fallback ink deficit for a vertical run, so
    // this MAIN commit path (the segment's stored measuredWidth and the pen advance
    // to the next segment) matches the ink-sized cell `drawVerticalRun` paints
    // (measure == paint); 0 for horizontal / non-under-reporting runs.
    const w = segAdvanceWidth(s, m.width + verticalInkExtra(s, s.text), gridDeltaPx, scale);
    // Line-height tracks the un-scaled pt font so super/sub don't shrink the line.
    const h = s.fontSize;
    // Prefer font-metric ascent/descent (stable per font+size) so baselines and
    // line boxes do not jitter based on the specific characters on each line.
    // When the document font is substituted by one with different vertical
    // metrics, rescale to the document font's design line box so the line
    // height (and thus baseline centering, row auto-heights, cell vAlign
    // §17.4.84, and pagination) match Word — see font-metrics.ts.
    // Fallback box at the run's full size; correction at the effective size
    // (smallCaps/vertAlign shrink it). See correctedLineMetrics.
    // §17.3.2.33: a small-caps run's LINE BOX follows the FULL run size, not the
    // 2pt-reduced glyph size — so a wrapped continuation line carrying only reduced
    // (all-lowercase) small caps is not short. Measure the box metrics at the full
    // size (the advance `w` above still uses the reduced glyphs). Super/subscript
    // is excluded: it intentionally shrinks its line contribution.
    const fullPx = s.fontSize * scale;
    let metricM = m;
    let metricEmPx = effectiveFontPx(s);
    if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
      const prevFont = ctx.font;
      ctx.font = buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses);
      metricM = ctx.measureText(s.text || 'X');
      ctx.font = prevFont;
      metricEmPx = fullPx;
    }
    // FE design correction for EA segments only; ruby keeps its measured box
    // (see addToLine's segScriptHint note).
    const corrected = correctedLineMetrics(
      metricM,
      s.fontFamily,
      fullPx,
      metricEmPx,
      (s.metricEastAsian === true || EAST_ASIAN_RE.test(s.text)) && !s.ruby,
    );
    let asc = corrected.ascent;
    const desc = corrected.descent;
    // Ruby annotation: small text rendered above the base. Reserve ascent
    // room for the rt glyphs (ECMA-376 §17.3.3.25). The actual line spacing
    // in docGrid sections is set further down by the paragraph-wide
    // pitch snap — see the doc-grid-aware path in renderParagraph /
    // estimateParagraphHeight that takes max(perLineH) and snaps it to
    // an integer grid pitch. The reserve here just ensures the natural
    // line height EXCEEDS the docGrid pitch when ruby is present, so the
    // snap actually picks the next pitch slot. 1.5× rt-size is enough for
    // any rt size that fits in one extra pitch above the base.
    if (s.ruby) {
      asc = asc + rubyAscentReservePx(s.ruby.fontSizePt, s.ruby.hpsRaisePt, scale);
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
    const trailingSpaceW = s.text.endsWith(' ') ? w - strAdvance(s, trimmed) : 0;
    const wForFit = w - trailingSpaceW;
    // The two fit-tolerance roles are EXCLUSIVE per line, mirroring paint's
    // per-line predicate `isJustified && (!endsLogicalLine || stretchLastLine)`
    // (`next` is the first segment after the prospective closing candidate):
    //
    //  - A line the paint pass will JUSTIFY stretches to the column edge. Word-
    //    observed issue #698 behavior admits only the backend-agnostic Canvas-vs-
    //    Word per-font bias there; the trailing-space allowance is suppressed.
    //  - A line left NON-justified keeps the classic Knuth-Plass trailing-space
    //    shrink allowance, whose 25 % promise the draw pass actually spends
    //    (`shrinkFitCompression`). Demo/sample-1 p3/p6 space-collapse evidence
    //    shows that ADDING the bias double-counts tolerance and admits words the
    //    non-justified paint path cannot fit.
    // Dictionary-SEA candidate (Thai/Lao/Khmer; grapheme-fill Myanmar/Tibetan
    // excluded — their Word-verified wrap is per-cluster greedy, #961, and the
    // #991 ground truth covers only the dictionary scripts). Per-codepoint
    // scan: a rare segment mixing both SEA families is NOT dictionary-SEA, so
    // it keeps the pre-#991 greedy path instead of moving a grapheme-fill span
    // inside an atomic chunk.
    const sDictSea = s.seaBreaks !== undefined && isDictionarySeaText(s.text);
    const shrinkBudgetFor = (next: LayoutSeg | undefined, biasBudget: number): number => {
      const closesLogicalLine = next === undefined || 'lineBreak' in next;
      const lineWillJustify = isJustified && (!closesLogicalLine || stretchLastLine);
      if (lineWillJustify) return biasBudget;
      // Word-observed (issue #991 calibration sweep): a line carrying
      // dictionary-SEA text never compresses its inter-word spaces — Word
      // wraps at natural fit for every overflow > 0 regardless of how many
      // spaces the line holds. The candidate counts too: admitting it would
      // make the line SEA, so the same zero-shrink fit applies (the sweep's
      // committed tokens were all Thai; mixed-script GT is uncollected — this
      // takes the wrap conservatively). The 25% drawable budget below is the
      // Latin/CJK-verified behavior.
      return lineHasSea || sDictSea ? 0 : lineTotalTrailingW * SPACE_SHRINK_RATIO;
    };
    const shrinkBudget = shrinkBudgetFor(
      queue[0],
      lineBiasBudget + biasBudgetContribution(s, trimmed),
    );

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
      !hasCJKBreakOpportunity(s.text) &&
      // A SEA (Thai/Lao/Khmer) lead with usable word breaks is NOT atomic — the
      // run splits at a dictionary boundary (issue #797), mirroring the CJK gate.
      !(s.seaBreaks && s.seaBreaks.length > 0)
    ) {
      let groupW = w;
      let groupTrail = trailingSpaceW;
      let groupEnd = 0;
      let groupBiasBudget = lineBiasBudget;
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
            groupW += strAdvance(f, prefix);
            if (prefix.length > 0) advanceGroupBias(f, prefix);
            groupTrail = 0;
            break;
          }
          // Entirely non-starters (no breakable rest): fall through to full-add.
        }
        const fw = segAdvance(f);
        groupW += fw;
        advanceGroupBias(f);
        const ft = f.text.replace(/ +$/, '');
        groupTrail = f.text.endsWith(' ') ? fw - strAdvance(f, ft) : 0;
      }
      groupBiasBudget += biasBudgetContribution(
        pendingGroupBiasSeg,
        pendingGroupBiasText.replace(/ +$/, ''),
      );
      if (
        currentWidth + (groupW - groupTrail)
        > availW() + shrinkBudgetFor(queue[groupEnd], groupBiasBudget)
      ) {
        flush(undefined, false, s.src);
      }
    }

    // No-space SEA chunk placement — ECMA-376 prescribes no SEA line-breaking
    // algorithm; Word-observed (issue #991, calibration fixture parts II/II-D):
    // the dictionary boundaries inside a no-space Thai/Lao/Khmer chunk are
    // SECONDARY break opportunities. A chunk that does not fit the remaining
    // width of a non-empty line moves to the next line WHOLE when it fits a
    // full line by itself — Word never splits it mid-chunk to fill the current
    // line (invariant across remaining widths, across the chunk being authored
    // as several glued `w:r`, and with/without a leading tab). Only a chunk
    // wider than a full line breaks at the dictionary boundaries (part II-D;
    // ordinary spaceless Thai paragraphs wrap this way on every line), which
    // is the greedy SEA branch below, kept unchanged.
    //
    // Judged only at chunk START: if the previously committed token is a text
    // segment glued to `s` (no trailing space), the whole chunk already passed
    // this judgment when its head was placed, so a mid-chunk segment never
    // needs it. The chunk spans `s` plus following queue segments while they
    // stay dictionary-SEA text glued without intervening spaces. Grapheme-fill
    // scripts (Myanmar/Tibetan) are excluded: their Word-verified wrap fills
    // per cluster (#961), so a fitting chunk must NOT move whole.
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
          if (ft.text.endsWith(' ')) { chunkEnd++; break; } // a space ends the chunk
        }
      }
      const chunkWForFit = chunkW - chunkTrail;
      if (
        currentWidth + chunkWForFit > availW() + shrinkBudgetFor(queue[chunkEnd], chunkBias) &&
        chunkWForFit <= lineMaxWidth
      ) {
        flush(undefined, false, s.src);
      }
    }

    if (currentWidth + wForFit <= availW() + shrinkBudget) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
    } else if (hasCJKBreakOpportunity(s.text) && s.seaBreaks === undefined) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
      // A segment that ALSO contains SEA (a mixed CJK+SEA `<w:cs/>` run) is routed
      // to the SEA branch below instead — its `seaBreaks` already merges the CJK
      // per-character opportunities with the SEA dictionary/transition ones
      // (issue #960), so both scripts break by their own rule from one offset set.
      // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
      //  separate: it sums per-char advances, whereas this path uses substring
      //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
      const available = availW() - currentWidth;
      setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
      const prevKern = setSegKerning(s);
      let rawPrefix = '';
      try {
        rawPrefix = available > 0 ? fitCJKPrefix(ctx, s.text, available, segmentCharacterGridDeltaPx(s, gridDeltaPx), charScaleFactor(s), charSpacingDeltaPx(s, scale), s.verticalRun === true) : '';
      } finally {
        restoreKerning(prevKern);
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
      const split = extendThroughTrailingIdeographicSpaces(
        allChars,
        kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, minSplit),
      );
      const prefix = allChars.slice(0, split).join('');
      if (prefix.length > 0) {
        // Grid advance for the head piece — the same model as the line box / draw.
        const pw = strAdvance(s, prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pw };
        addToLine(headSeg, pw, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) {
          queue.unshift({
            ...s,
            text: tail,
            measuredWidth: 0,
            src: {
              segIndex: s.src!.segIndex,
              charOffset: s.src!.charOffset + prefix.length,
            },
          });
        }
      } else if (currentLine.length > 0) {
        // No prefix of `s` fits. If `s` would lead the next line with a 行頭禁則
        // char, kinsokuAdjustedSplit can't fix it from within `s` (the offending
        // char is its first); pull trailing graphemes of the current line's last
        // text segment down so they lead the next line ahead of `s` — cross-run
        // 追い出し (§17.3.1.16). See crossRunKinsokuRetract for the bounded,
        // re-validating, whitespace-guarded retraction count.
        let retracted: LayoutTextSeg | null = null;
        const sFirstCp = s.text.codePointAt(0);
        const lastSeg = currentLine[currentLine.length - 1];
        if (sFirstCp !== undefined && kinsoku.lineStartForbidden.has(sFirstCp) && 'text' in lastSeg) {
          const lastText = lastSeg as LayoutTextSeg;
          const chars = [...lastText.text];
          const minKeep = currentLine.length > 1 ? 0 : 1;
          const k = crossRunKinsokuRetract(chars, kinsoku, minKeep);
          if (k > 0) {
            const headText = chars.slice(0, chars.length - k).join('');
            const tailText = chars.slice(chars.length - k).join('');
            retracted = {
              ...lastText,
              text: tailText,
              measuredWidth: strAdvance(lastText, tailText),
              src: {
                segIndex: lastText.src!.segIndex,
                charOffset: lastText.src!.charOffset + headText.length,
              },
            };
            if (headText) {
              const headW = strAdvance(lastText, headText);
              currentWidth -= lastText.measuredWidth - headW;
              currentLine[currentLine.length - 1] = { ...lastText, text: headText, measuredWidth: headW };
            } else {
              // Whole last segment moves down. Line metrics (ascent/descent) are
              // not recomputed; the retracted graphemes share the line's font in
              // practice, so the flushed box height is unaffected.
              currentWidth -= lastText.measuredWidth;
              currentLine.pop();
            }
          }
        }
        flush(undefined, false, retracted?.src ?? s.src);
        queue.unshift(s);
        if (retracted) queue.unshift(retracted);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const forcedChars = [...s.text];
        const forcedSplit = forcedChars.length > 0
          ? extendThroughTrailingIdeographicSpaces(forcedChars, 1)
          : 0;
        const firstChar = forcedChars.slice(0, forcedSplit).join('');
        if (firstChar) {
          const fw = strAdvance(s, firstChar);
          const headSeg: LayoutTextSeg = { ...s, text: firstChar, measuredWidth: fw };
          addToLine(headSeg, fw, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) {
            queue.unshift({
              ...s,
              text: tail,
              measuredWidth: 0,
              src: {
                segIndex: s.src!.segIndex,
                charOffset: s.src!.charOffset + firstChar.length,
              },
            });
          }
        }
      }
    } else if (s.seaBreaks !== undefined) {
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
      const monotone = isGraphemeFillText(s.text);
      const split = fitSeaWordPrefix(s.text, s.seaBreaks, 0, available, measureSub, monotone);
      if (split > 0) {
        const prefix = s.text.slice(0, split);
        const pw = strAdvance(s, prefix);
        addToLine({ ...s, text: prefix, measuredWidth: pw }, pw, h, asc, desc);
        const tail = s.text.slice(split);
        if (tail) {
          queue.unshift({
            ...s,
            text: tail,
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
        let retracted: LayoutTextSeg | null = null;
        const sFirstCp = s.text.codePointAt(0);
        const lastSeg = currentLine[currentLine.length - 1];
        if (sFirstCp !== undefined && kinsoku.lineStartForbidden.has(sFirstCp) && 'text' in lastSeg) {
          const lastText = lastSeg as LayoutTextSeg;
          const chars = [...lastText.text];
          const minKeep = currentLine.length > 1 ? 0 : 1;
          const k = crossRunKinsokuRetract(chars, kinsoku, minKeep);
          if (k > 0) {
            const headText = chars.slice(0, chars.length - k).join('');
            const tailText = chars.slice(chars.length - k).join('');
            retracted = {
              ...lastText,
              text: tailText,
              measuredWidth: strAdvance(lastText, tailText),
              src: {
                segIndex: lastText.src!.segIndex,
                charOffset: lastText.src!.charOffset + headText.length,
              },
              seaBreaks: rebaseSeaBreaks(lastText.seaBreaks, headText.length),
            };
            if (headText) {
              const headW = strAdvance(lastText, headText);
              currentWidth -= lastText.measuredWidth - headW;
              currentLine[currentLine.length - 1] = { ...lastText, text: headText, measuredWidth: headW };
            } else {
              currentWidth -= lastText.measuredWidth;
              currentLine.pop();
            }
          }
        }
        flush(undefined, false, retracted?.src ?? s.src);
        queue.unshift(s);
        if (retracted) queue.unshift(retracted);
      } else {
        // Empty line and the first dictionary word is wider than the whole
        // column: emergency GRAPHEME-safe split (a code-point split would tear a
        // base + tone/combining mark, both BMP). Guarantee ≥1 cluster of progress.
        const firstWordEnd = s.seaBreaks[0] ?? s.text.length;
        const firstWord = s.text.slice(0, firstWordEnd);
        const graphemes = graphemeClusterOffsets(firstWord);
        let gsplit = fitSeaWordPrefix(firstWord, graphemes, 0, available, measureSub, monotone);
        if (gsplit <= 0) gsplit = graphemes.length > 0 ? graphemes[0] : firstWord.length;
        const prefix = s.text.slice(0, gsplit);
        const pw = strAdvance(s, prefix);
        addToLine({ ...s, text: prefix, measuredWidth: pw }, pw, h, asc, desc);
        const tail = s.text.slice(gsplit);
        if (tail) {
          queue.unshift({
            ...s,
            text: tail,
            measuredWidth: 0,
            src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + gsplit },
            seaBreaks: rebaseSeaBreaks(s.seaBreaks, gsplit),
          });
        }
      }
    } else if (currentLine.length === 0) {
      // A single non-CJK word wider than the FULL line width — e.g. a long URL in
      // a narrow newspaper column. ECMA-376 prescribes no line-break algorithm;
      // Word breaks such an over-long word at the character level (overflow-wrap)
      // so it stays inside the column instead of bleeding past the right margin /
      // into the next column. Fit the widest character prefix (≥1 char so the
      // split always advances), draw it, and re-queue the remainder. Segments are
      // already one space-delimited word (splitTextForLayout), so this never
      // breaks where a space could have wrapped.
      const available = availW();
      setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
      const allChars = [...s.text];
      const prevKern = setSegKerning(s);
      let split = 0;
      try {
        split = available > 0 ? [...fitCJKPrefix(ctx, s.text, available, segmentCharacterGridDeltaPx(s, gridDeltaPx), charScaleFactor(s), charSpacingDeltaPx(s, scale), s.verticalRun === true)].length : 0;
      } finally {
        restoreKerning(prevKern);
      }
      if (split < 1) split = 1;
      split = extendThroughTrailingIdeographicSpaces(allChars, split);
      if (split >= allChars.length) {
        // The visible glyphs actually fit (only a trailing space pushed it over the
        // fit test) — place the word whole.
        s.measuredWidth = w;
        addToLine(s, w, h, asc, desc);
      } else {
        const prefix = allChars.slice(0, split).join('');
        const pw = strAdvance(s, prefix);
        addToLine({ ...s, text: prefix, measuredWidth: pw }, pw, h, asc, desc);
        queue.unshift({
          ...s,
          text: allChars.slice(split).join(''),
          measuredWidth: 0,
          src: {
            segIndex: s.src!.segIndex,
            charOffset: s.src!.charOffset + prefix.length,
          },
        });
      }
    } else {
      // Latin word does not fit on the current (non-empty) line: move it to a fresh
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

  return lines;
}

/** Phase 4-1 B2 Stage 2 — rehydrate the paginator's scale-1 stamped lines into
 *  the paint scale while keeping the scale-1 line PARTITION. The default legacy
 *  bridge re-derives hinting-sensitive geometry at the paint scale; the optional
 *  canonical bridge scales the complete stored geometry for viewport paint.
 *
 *  Why not a pure ×scale of the scale-1 fields: a real (hinted) font's Canvas
 *  metrics are NOT scale-linear — `measureText(pt·s).width ≠ s · measureText(pt)`
 *  and `fontBoundingBoxAscent(pt·s) ≠ s · fontBoundingBoxAscent(pt)`. The draw
 *  path advances the pen by each segment's `measuredWidth` while `fillText`
 *  renders at the paint-scale font, and the line box uses `ascent/descent`; a
 *  geometric ×scale of those would drift the pen off the glyphs horizontally and
 *  the baseline off by up to ~1px vertically, accumulating down the page (seen on
 *  demo/sample-1 p.4). So RE-MEASURE every text segment at the paint scale here —
 *  the exact same per-segment measurement `layoutLines` does (advance via
 *  segAdvanceWidth, box via correctedLineMetrics, the §17.3.2.33 small-caps full-size
 *  box, the §17.3.3.25 ruby ascent reserve, and the intended-single-line floor) —
 *  so measure == draw byte-for-byte, identical to a fresh paint-scale layout of
 *  the SAME partition. Only WHICH glyphs sit on WHICH line comes from the scale-1
 *  stamp; that is the zoom-invariant part (Word lays text out in the document's
 *  coordinate space and the display scale is a viewport transform, not a
 *  re-layout — the scale-1 partition is correct at every zoom). This makes reuse
 *  a deliberate behaviour change vs. the old recompute-at-paint-scale path, which
 *  let hinting shift wrap points per zoom.
 *
 *  Geometry that IS scale-linear (a clean ×scale): the float-exclusion `xOffset` /
 *  `availWidth` / `topY` (pure page geometry, no glyph hinting) and non-text
 *  advances — image `widthPt·scale`, math `render.*Em·emPx`. Empty/synthetic
 *  lines (no text segment) fall back to the ×scale of the stamped box, matching
 *  layoutLines' `h·scale·0.8`/`0.2` synthesis.
 *
 *  Returns FRESH line + segment objects (never mutates the shared stamp — the
 *  draw path and repeated renderPage calls read the same immutable array; see
 *  layout-lines-reuse-identity.test.ts). `scale === 1` returns the stamp
 *  unchanged (identity — no copy, no re-measure), so the scale-1 reuse path stays
 *  byte-for-byte as in Stage 1.
 *
 *  When `canonicalGeometry` is true, text and line geometry are also mapped by
 *  a pure ×scale from the stamp. The caller must then paint scale-1 glyphs under
 *  the matching Canvas transform; this is the document-space body-text path that
 *  preserves both the frozen partition and its width contract. */
export function rescaleLayoutLines(
  lines: LayoutLine[],
  scale: number,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string>,
  gridDeltaPx: number,
  canonicalGeometry = false,
): LayoutLine[] {
  if (scale === 1) return lines;

  if (canonicalGeometry) {
    return lines.map((line) => ({
      ...line,
      segments: line.segments.map((segment) => {
        if ('imagePath' in segment) {
          return segment.anchor
            ? { ...segment, measuredWidth: 0 }
            : { ...segment, measuredWidth: segment.widthPt * scale };
        }
        if ('mathNodes' in segment) {
          return {
            ...segment,
            measuredWidth: segment.measuredWidth * scale,
            mathAscent: segment.mathAscent * scale,
            mathDescent: segment.mathDescent * scale,
          };
        }
        if ('text' in segment) {
          return {
            ...segment,
            measuredWidth: segment.measuredWidth * scale,
            fitTextPerGapPx: segment.fitTextPerGapPx === undefined
              ? undefined
              : segment.fitTextPerGapPx * scale,
            fitTextTrailingPadPx: segment.fitTextTrailingPadPx === undefined
              ? undefined
              : segment.fitTextTrailingPadPx * scale,
          };
        }
        return { ...segment, measuredWidth: segment.measuredWidth * scale };
      }),
      ascent: line.ascent * scale,
      descent: line.descent * scale,
      visibleAscent: (line.visibleAscent ?? line.ascent) * scale,
      visibleDescent: (line.visibleDescent ?? line.descent) * scale,
      visibleIntendedSingle: (line.visibleIntendedSingle ?? line.intendedSingle) * scale,
      intendedSingle: line.intendedSingle * scale,
      gridCountSingle: line.gridCountSingle * scale,
      xOffset: line.xOffset * scale,
      availWidth: line.availWidth * scale,
      topY: line.topY === undefined ? undefined : line.topY * scale,
    }));
  }

  const withSegKerning = <T>(s: LayoutTextSeg, measure: () => T): T => {
    if (s.kerning == null) return measure();
    const previous = ctx.fontKerning;
    ctx.fontKerning = s.fontSize >= s.kerning ? 'normal' : 'none';
    try {
      return measure();
    } finally {
      ctx.fontKerning = previous;
    }
  };

  const naturalWidth = (s: LayoutTextSeg, text: string): number =>
    withSegKerning(s, () =>
      ctx.measureText(text).width + (s.verticalRun ? verticalRunInkExtraPx(ctx, text) : 0));

  // Per-text-segment measurement at the PAINT scale — the SAME model layoutLines
  // uses (segAdvance + the text-segment metric block), so measure == draw.
  const measureTextSeg = (s: LayoutTextSeg): {
    advance: number;
    asc: number;
    desc: number;
    intended: number;
    gridCount: number;
  } => {
    const effPx = calcEffectiveFontPx(s, scale);
    ctx.font = buildFont(s.bold, s.italic, effPx, s.fontFamily, fontFamilyClasses);
    const m = withSegKerning(s, () => ctx.measureText(s.text));
    // #1014 — fold in the vo=Tr rotate-fallback ink-extent deficit for a vertical
    // run so the rescaled box matches the ink-sized cell (measure == draw); 0 on
    // horizontal pages and non-under-reporting fonts. `ctx.font` is set above.
    const advance = segAdvanceWidth(s, naturalWidth(s, s.text), gridDeltaPx, scale);
    // §17.3.2.33 — a small-caps run's LINE BOX follows the FULL run size (measure
    // the box at fullPx, not the 2pt-reduced glyphs); super/subscript keeps its
    // shrunk contribution. Mirrors layoutLines' fullPx / metricEmPx split.
    const fullPx = s.fontSize * scale;
    let metricM = m;
    let metricEmPx = effPx;
    if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
      ctx.font = buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses);
      metricM = ctx.measureText(s.text || 'X');
      metricEmPx = fullPx;
    }
    // FE design correction for EA segments only; ruby keeps its measured box
    // (see addToLine's segScriptHint note).
    const segScriptHint = (s.metricEastAsian === true || EAST_ASIAN_RE.test(s.text)) && !s.ruby;
    const corrected = correctedLineMetrics(metricM, s.fontFamily, fullPx, metricEmPx, segScriptHint);
    // §17.3.3.12 — ruby reserves hpsRaise when present, same as layoutLines.
    const asc = s.ruby
      ? corrected.ascent + rubyAscentReservePx(s.ruby.fontSizePt, s.ruby.hpsRaisePt, scale)
      : corrected.ascent;
    // Intended single-line floor (font-metrics.ts) — small caps keep the FULL run
    // size here too (addToLine's intendedEm).
    const intendedEm = s.smallCaps && !s.vertAlign ? fullPx : effPx;
    const intended = segmentIntendedSingleLinePx(s, intendedEm, segScriptHint);
    return {
      advance,
      asc,
      desc: corrected.descent,
      intended,
      gridCount: segScriptHint ? eastAsianGridCountSinglePx(intended, intendedEm) : 0,
    };
  };

  return lines.map((l) => {
    let asc = 0;
    let desc = 0;
    let intended = 0;
    let gridCount = 0; // px — max over segments of (tabled design | untabled 1.3em)
    let hasText = false;
    let visibleAsc = 0;
    let visibleDesc = 0;
    let visibleIntended = 0;
    let hasVisibleMetrics = false;
    const scaledSource = l.segments.map((segment) => ({ ...segment })) as LayoutSeg[];
    // Recompute the region gap from paint-scale natural metrics. Reusing the
    // scale-1 gap would break measure==paint because targetPx and font advances
    // are both scale-relative (and font hinting need not be linearly scalable).
    resolveFitTextSegments(
      scaledSource.filter((segment): segment is LayoutTextSeg => 'text' in segment),
      scale,
      (segment) => {
        const effPx = calcEffectiveFontPx(segment, scale);
        ctx.font = buildFont(segment.bold, segment.italic, effPx, segment.fontFamily, fontFamilyClasses);
        // #1014 — include the vo=Tr ink deficit so the rescaled fitText gap stays
        // measure==paint against the ink-grown cell; 0 for non-under-reporting runs.
        return naturalWidth(segment, segment.text);
      },
    );
    const segments = scaledSource.map((s) => {
      if ('isTab' in s) {
        // Tab advance is position-dependent (pen → next stop) and box-neutral; a
        // ×scale reproduces the scale-1 tab-to-stop advance scaled to paint space
        // (the stamp reuse is gated to no-float, non-marker paragraphs, so the
        // tab-stop grid scales cleanly with the box).
        return { ...s, measuredWidth: s.measuredWidth * scale };
      }
      if ('imagePath' in s) {
        // Anchored images live out of inline flow: layoutLines pins their
        // measuredWidth to 0 (they add no pen advance) and their anchor*Pt
        // fields stay in pt space for the draw-time position resolver.
        if (s.anchor) return { ...s, measuredWidth: 0 };
        const h = s.heightPt * scale;
        // Inline images stand fully above the baseline (descent 0) and always
        // size docGrid cells, mirroring layoutLines' addToLine contribution.
        asc = Math.max(asc, h);
        gridCount = Math.max(gridCount, h);
        visibleAsc = Math.max(visibleAsc, h);
        hasVisibleMetrics = true;
        return { ...s, measuredWidth: s.widthPt * scale };
      }
      if ('mathNodes' in s) {
        const copy = { ...s, measuredWidth: s.measuredWidth * scale } as LayoutMathSeg;
        copy.mathAscent *= scale;
        copy.mathDescent *= scale;
        asc = Math.max(asc, copy.mathAscent, s.fontSize * scale * 0.8);
        desc = Math.max(desc, copy.mathDescent, s.fontSize * scale * 0.2);
        gridCount = Math.max(gridCount, copy.mathAscent + copy.mathDescent);
        visibleAsc = Math.max(visibleAsc, copy.mathAscent, s.fontSize * scale * 0.8);
        visibleDesc = Math.max(visibleDesc, copy.mathDescent, s.fontSize * scale * 0.2);
        hasVisibleMetrics = true;
        return copy;
      }
      const t = s as LayoutTextSeg;
      const mm = measureTextSeg(t);
      hasText = true;
      if (mm.asc > asc) asc = mm.asc;
      if (mm.desc > desc) desc = mm.desc;
      if (mm.intended > intended) intended = mm.intended;
      if (t.metricOnly !== true) {
        visibleAsc = Math.max(visibleAsc, mm.asc);
        visibleDesc = Math.max(visibleDesc, mm.desc);
        visibleIntended = Math.max(visibleIntended, mm.intended);
        hasVisibleMetrics = true;
      }
      // Only East Asian text is cell-rounded (§17.6.5); a Latin run keeps its
      // natural height and must not contribute its (substituted) box to the
      // grid-count height. measureTextSeg derives both tabled and untabled
      // contributions from deterministic design quantities.
      gridCount = Math.max(gridCount, mm.gridCount);
      return { ...t, measuredWidth: mm.advance };
    });
    // Empty/synthetic line (no text/math contributing metrics): fall back to the
    // ×scale of the stamped box (layoutLines synthesises h·scale·0.8/0.2 there).
    if (!hasText && asc === 0 && desc === 0) {
      asc = l.ascent * scale;
      desc = l.descent * scale;
      intended = l.intendedSingle * scale;
      gridCount = l.gridCountSingle * scale;
    }
    return {
      ...l,
      segments,
      ascent: asc,
      descent: desc,
      visibleAscent: hasVisibleMetrics
        ? visibleAsc
        : (l.visibleAscent ?? l.ascent) * scale,
      visibleDescent: hasVisibleMetrics
        ? visibleDesc
        : (l.visibleDescent ?? l.descent) * scale,
      visibleIntendedSingle: hasVisibleMetrics
        ? visibleIntended
        : (l.visibleIntendedSingle ?? l.intendedSingle) * scale,
      intendedSingle: intended,
      gridCountSingle: gridCount || (asc + desc),
      // Pure page geometry — scale-linear (no glyph hinting).
      xOffset: l.xOffset * scale,
      availWidth: l.availWidth * scale,
      topY: l.topY === undefined ? undefined : l.topY * scale,
    };
  });
}

/** Adapt a text-box run ({@link ShapeTextRun}) to the body run model
 *  ({@link DocRun} of `type:'text'`) so text-box paragraphs feed the SAME
 *  segment builder / line breaker the body uses ({@link buildSegments} +
 *  {@link layoutLines}) — giving them kinsoku (§17.15.1.58–.60), UAX#9 bidi,
 *  §17.18.44 justification and §17.3.1.37 tab stops. A `ShapeTextRun` carries a
 *  strict SUBSET of `DocxTextRun`'s formatting (text, size, colour, the ascii +
 *  eastAsia font axes §17.3.2.26, bold, italic, ruby); every field the shape model
 *  lacks (underline/strike/highlight/border/rtl/cs/…) takes its neutral
 *  default, so the run behaves exactly like a plain body text run. */
export function shapeRunToDocRun(run: ShapeTextRun): DocRun {
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
    // §17.3.2.26 eastAsia axis — routed per CJK code point by buildSegments, the
    // SAME split the old shapeTokenFamily tokenizer performed, but now inside the
    // shared builder so measure == draw with no bespoke tokenizer.
    fontFamilyEastAsia: run.fontFamilyEastAsia ?? null,
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
    ruby: run.ruby ?? undefined,
  } as unknown as DocRun;
}

/** Synthesize the `RenderState` fields {@link buildSegments} / {@link layoutLines}
 *  need for text-box layout when the caller does not thread the document state
 *  (unit tests call {@link renderShapeText} with just ctx/scale/fonts). Only the
 *  fields those two functions actually read for PLAIN TEXT runs are populated —
 *  buildSegments touches `state` solely on `field`/`noteRef` runs (which shape
 *  runs never produce), so an empty note map + spec-default kinsoku / tab
 *  interval is exact here. The production call site passes the real state, so
 *  document-level kinsoku (§17.3.1.16) / defaultTabStop (§17.15.1.25) flow
 *  through unchanged. */
export function shapeRenderState(
  ctx: CanvasRenderingContext2D,
  scale: number,
  fontFamilyClasses: Record<string, string>,
  images: Map<string, DecodedImage>,
): RenderState {
  return {
    ctx,
    scale,
    fontFamilyClasses,
    images,
    kinsoku: DEFAULT_KINSOKU_RULES,
    defaultTabPt: DEFAULT_TAB_PT,
  } as unknown as RenderState;
}
