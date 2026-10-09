import type { LineSpacing, TabStop, DocxRunBorder, EmphasisMark } from '../types';
import type { CanvasFontRoute, HyperlinkTarget, NumberFormat, Duotone, ResolvedFontMetric } from '@silurus/ooxml-core';
import { type FloatRect } from '../float-layout.js';
import type { LayoutServices } from '../layout/types.js';
import type { VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import type { GlyphInkBounds, TextLayoutService, TextShapeRequest, FontScriptSlot } from '../layout/text.js';
import type { MathLayoutResource } from '../layout/resources.js';

/** Display format and first number of automatic note references for each
 * note kind (ECMA-376 §17.11.17/.18 numFmt, §17.11.20 numStart). */
export interface NoteNumbering {
  readonly footnote: Readonly<{ format: string; start: number }>;
  readonly endnote: Readonly<{ format: string; start: number }>;
}

export interface LineBoundary {
  segIndex: number;
  charOffset: number;
}


export interface LayoutSegSource {
  src?: LineBoundary;
  /** Parser-boundary run occurrence retained independently of mutable source
   * objects. Unlike `src.segIndex` (the flattened segment stream), this remains
   * the original paragraph run index through line splitting. */
  sourceRunIndex?: number;
  /** Original ownership is independent of the canonical text sequence. */
  sourceTextSequence?: readonly import('./text-sequence.js').TextSequenceSource[];
  sourceTextOffset?: number;
}


export interface LayoutTextSeg extends LayoutSegSource {
  text: string;
  semanticSlotSpans?: import('../layout/text.js').TextShapeSpan['semanticSlotSpans'];
  /** Shared immutable slot source plus this slice's window. Emergency suffixes
   * must not copy all remaining slots on every break. Retained placements
   * materialize only their overlapping slots, rebased to their own text. */
  semanticSlotRange?: Readonly<{ start: number; end: number }>;
  /** Authored family and selected source survive local FontFace aliases.
   * Compatibility metadata may distinguish an installed authored face from a
   * substitute without inspecting the CSS alias or changing its paint route. */
  authoredFontFamily?: string | null;
  fontSource?: import('../layout/font-service.js').FontResolutionSource;
  /** Existing selected-face reference policy; application-provided SFNT
   * resources are local inventory entries but are not installed Office faces. */
  authoredReferenceMetricAllowed?: boolean;
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
  /** Native reserved-separator participant only (MS-DOC 2.3.3 rule control or
   * content mark): the bounded selected-face probe whose vertical metrics the
   * metric-only segment keeps. Never display text, glyph ownership, source
   * range length or inline width; other metric-only segments omit it. */
  metricProbeText?: string;
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
  /** Selected-route line ratio. It may come from parsed font bytes or a bounded
   * Canvas measurement; the latter does not reveal OpenType table identity. */
  /** Admitted reference profile has no Far East code-page bits; its Latin
   * single-line design height owns whole line-grid cells (#1674). */
  resolvedLatinGridCellAllocation?: boolean;
  resolvedLineHeightRatio?: number;
  /** A selected route supplied this ratio from measured or parsed geometry. */
  resolvedResourceVerticalMetric?: true;
  /** Metadata-only vertical reference for a native authored family. It grants
   * no width, cmap, resource ownership, shaping, or paint-route authority. */
  referenceFontVerticalMetric?: true;
  resolvedDesignAscentRatio?: number;
  resolvedDesignDescentRatio?: number;
  resolvedEastAsianLineHeightRatio?: number;
  /** Selected-face OS/2 xAvgCharWidth / unitsPerEm. The bounded inter-word
   * minimum uses half this scalar; it is never a natural advance. */
  latinSpaceAverageWidthRatio?: number;
  /** Set only when the document requests the registered compression mode. */
  latinSpaceCompressionEligible?: true;
  /** Retained paint advance is shorter than the natural space by this amount. */
  latinSpaceCompressionPx?: number;
  latinNaturalTrailingSpacePx?: number;
  /** Selected-face OS/2 xAvgCharWidth / unitsPerEm, set only where
   * WORD_COMPRESSED_SPACE_LINE_FIT may shrink this segment's U+0020 on a mixed
   * East Asian / Latin line. Independent of the Latin-only projection. */
  mixedSpaceAverageWidthRatio?: number;
  /** Natural advance and count of this segment's shrinkable trailing U+0020
   * under WORD_COMPRESSED_SPACE_LINE_FIT. */
  mixedNaturalTrailingSpacePx?: number;
  mixedNaturalTrailingSpaceCount?: number;
  vertAlign: 'super' | 'sub' | null;
  measuredWidth: number;  // px (set during layout)
  /** A2 text authority captured during segmentation; production text width and
   * metrics are resolved through this same service during line layout. */
  textLayoutService?: TextLayoutService;
  textShapeRequest?: TextShapeRequest;
  /** Run-context substitute decision, including false for excluded spans;
   * intrinsic merging must preserve the context for either decision. */
  substituteScope?: boolean;
  /** Contextually shaped grapheme geometry from the authoritative text service. */
  shapedClusters?: readonly Readonly<{
    range: Readonly<{ start: number; end: number }>;
    offsetPt: number;
    advancePt: number;
  }>[];
  /** Same-line native boundary advance; also shifts retained glyph origins. */
  leadingWordBoundaryPx?: number;
  /** Sparse, contextual U+0020 cluster geometry used only during gap fitting. */
  shapedSpaceClusters?: LayoutTextSeg['shapedClusters'];
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
  /** ECMA-376 §17.18.78 foreground tile; background retains the fill color. */
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
  /** Legal ordinary-hyphen and registered URL breaks, as segment-local UTF-16 offsets. */
  explicitBreaks?: import('./text-break-window.js').TextBreakWindow;
  /** This source seam follows a legal ordinary-hyphen or URL break. */
  explicitBreakBefore?: true;
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
  /** The originating parent run contains East Asian-script content. When that
   * run has no effective East-Asian language, this provides the bounded union
   * fallback independently of the observed Latin-parent compatibility rule. */
  overflowPunctuationEastAsianRun?: true;
  /** Effective `w:lang/@w:bidi` from the originating parent run. */
  overflowPunctuationBidiLanguage?: string;
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
   *  font size ≥ the threshold. Absent at every style level disables kerning
   *  regardless of `enableOpenTypeFeatures`. WORD_KERN_THRESHOLD_AUTHORITY
   *  additionally disables zero in mode 15; unmeasured modes retain the previous
   *  zero size comparison. Canvas `auto` is not the WordprocessingML default. */
  kerning?: number;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vert>` — horizontal-in-vertical
   *  (縦中横). Set by {@link buildSegments} ONLY when the run declares `w:vert`
   *  AND the page is vertical (tbRl); the property is inert in a horizontal page,
   *  so the gate is folded in here at build time and the measure/paint passes just
   *  read this flag. When set, the whole segment occupies ONE cell along the
   *  vertical column (advance = 1em, NOT the per-glyph sideways width), with its
   *  characters drawn horizontally side by side across the column (§17.3.2.10,
   *  PDF comparison). Absent ⇒ normal per-glyph vertical advance. */
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


/**
 * Horizontal tab. Width is resolved during layout against paragraph tab stops
 * (or the default 36pt interval if no explicit stop is configured).
 */
export interface LayoutTabSeg extends LayoutSegSource {
  isTab: true;
  fontSize: number;  // pt — for line-height purposes
  measuredWidth: number;
  /** Queue-resolved reading-frame gap. The bidi post-pass must preserve the
   * same gap that ordinary text fitting consumed, including a collapsed
   * unreachable stop on an empty line. */
  readingGap?: number;
  /** tab leader to fill the gap (e.g. TOC dot leaders); set during layout. */
  leader?: TabStop['leader'];
  /** Alignment selected from the effective stop during layout. */
  resolvedAlignment?: TabStop['alignment'];
  /** Set when this aligned tab's cell was admitted past the paragraph's
   *  trailing indent into the line's exclusion-free margin extension. */
  marginAllocation?: boolean;
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
  /** The source is a literal wp:inline picture, rather than a chart/shape box. */
  inlinePicture?: true;
  /** Paragraph mark's selected-face natural single line, measured once per paragraph. */
  paragraphMarkSinglePx?: number;
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
  /** Present (including zero) for the measured proportional gap policy. The
   * natural advances stay intact; layout applies this slack once to paint. */
  justifiedCompressionPx?: number;
  gapPlan?: import('./line-gaps.js').LineGapPlan;
  /** Pass-local physical identity: gap fragments share it even if vertical
   * rounding makes distinct physical lines have equal numeric tops. */
  physicalLineIndex?: number;
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
  /** px — intended single-line height from admitted font geometry, if present. */
  intendedSingle: number;
  /** Text-face single line that supplies automatic leading to an inline picture. */
  inlinePictureTextSingle?: number;
  /** Admitted Latin design height; native fallback boxes do not establish grid cells. */
  latinGridCountSingle?: number;
  /** Registered compatibility allocation for a uniform positioned, visible run. */
  uniformPositionAuto?: Readonly<{ normalSinglePx: number; positionPx: number; designDescentPx: number }>;
  /** px — DESIGN grid-count height: the max over segments of each run's
   *  format-policy single-line height (a resolved resource's design height,
   *  or the generic East Asian fallback). Feeds docGrid cell
   *  counting without depending on a substituted face's Canvas box
   *  (§17.6.5). */
  gridCountSingle: number;
  /** Additional horizontal offset (px) from paraX, caused by wrap-around floats. */
  xOffset: number;
  /** Effective available width (px) for this line after float exclusion. */
  availWidth: number;
  /** Width (px) past `availWidth` up to the text margin that this line's
   *  margin-allocated tab cell occupies as part of its band. */
  marginExtension?: number;
  /** When wrap context is active, the absolute canvas Y where this line begins. */
  topY?: number;
  /** Confirmed fixed-point allocation that owns topY, in the same units as
   * the wrap context. Never infer ownership from a numeric top alone. */
  wrapAllocation?: Readonly<{ physicalLineIndex: number; topYPt: number; advancePt: number }>;
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
  hasExclusions?: boolean;
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
   *  like visible content admitted by its next atom (#1670). */
  paragraphMarkLineStartWidth?: number;
  /** Placement-aware wrap boundary used by paragraph measurement. */
  lineWindow?: (input: {
    topYPt: number;
    minimumStartWidthPt: number;
    /** Required atomic start width for a square-constrained gap. */
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
  /** Paragraph-wide allocation (ruby, spacing, grid and inline objects).
   * Supplies both float probes and physical-line cursor advancement. */
  resolveLineAdvances?: (lines: readonly LayoutLine[]) => readonly number[];
  /** Per-line box-height resolver for isolated line-layout callers. Paragraph
   * measurement supplies resolveLineAdvances so origins and probes include its
   * paragraph-wide allocation rather than only this fragment's metrics. */
  lineBoxH: (ascentPx: number, descentPx: number, hasRuby?: boolean, intendedSinglePx?: number, eastAsian?: boolean, gridCountSinglePx?: number, uniformPositionAuto?: LayoutLine['uniformPositionAuto'], inlinePictureTextSingle?: number, latinGridCountSingle?: number) => number;
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
  /** Effective §17.3.1.33 paragraph spacing. The Word-for-Mac OpenType
   * projection has been established for automatic spacing only. */
  readonly lineSpacing?: LineSpacing | null;
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
  /** ECMA-376 §17.11.17/.18 numFmt and §17.11.20 numStart per note kind.
   * Absent means decimal numbering from 1. */
  readonly noteNumbering?: NoteNumbering;
  readonly noteReferenceNumber?: number;
  readonly verticalCJK?: boolean;
  /** ECMA-376 Part 4 §14.8.3.50 w:useFELayout compatibility switch. */
  readonly useFeLayout?: boolean;
  /** §17.6.5 effective line-grid gate after section type, paragraph opt-out,
   *  and §17.15.3.1 table-cell compatibility have been resolved. */
  readonly lineGridActive?: boolean;
  /** Explicit paragraph character-allocation fact; omission is unknown. */
  readonly characterGridActive?: boolean;
  /** Paragraph base direction after inherited bidi resolution. */
  readonly paragraphRtl?: boolean;
  /** §17.15.3.3 w:balanceSingleByteDoubleByteWidth document compatibility switch. */
  readonly balanceSingleByteDoubleByteWidth?: boolean;
  readonly resolvedLocalFonts?: Readonly<Record<string, ResolvedFontMetric>>;
  readonly layoutServices?: LayoutServices;
  readonly verticalGlyphMeasurement?: VerticalGlyphMeasurementService;
  /** ECMA-376 §17.15.1.18 document-wide full-width character compression. */
  readonly characterSpacingControl?: string;
  /** §17.15.3.31: use full character width when deciding line fit. */
  readonly lineWrapLikeWord6?: boolean;
  /** `w:compatSetting` compatibilityMode; absent when not authored. Gates
   * WORD_COMPRESSED_SPACE_LINE_FIT. */
  readonly compatibilityMode?: number;
  /** Paragraph §17.3.1.2-3 automatic East Asian/Latin and East Asian/number
   * spacing (absent means on). The renderer does not model that spacing;
   * WORD_COMPRESSED_SPACE_LINE_FIT stays out of a paragraph it applies to. */
  readonly autoSpaceDE?: boolean;
  readonly autoSpaceDN?: boolean;
  /** See WORD_KERN_THRESHOLD_AUTHORITY for absent `w:kern`. */
  readonly enableOpenTypeFeatures?: boolean;
  /** False only when `w:framePr` specifies a drop cap with a fixed `w:lines`;
   * the authored frame height remains authoritative even when glyph paint is
   * lowered beyond it. Folded into retained text segments during acquisition. */
  readonly positionExtendsLineBox?: boolean;
}
