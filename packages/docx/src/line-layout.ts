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
  LineSpacing, TabStop, DocxRunBorder, DocSettings,
} from './types';
import type { MathNode, KinsokuRules, ChartModel } from '@silurus/ooxml-core';
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
  classifyFontGeneric,
  isComplexScriptCodePoint,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  measureAdvanceWithTrailingSpace,
} from '@silurus/ooxml-core';
import { intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';
import {
  type FloatRect,
  resolveLineFloatWindow,
  wordMinLineStartPx,
} from './float-layout.js';

export interface LayoutTextSeg {
  text: string;
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
  ruby?: { text: string; fontSizePt: number };
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
}

/**
 * Horizontal tab. Width is resolved during layout against paragraph tab stops
 * (or the default 36pt interval if no explicit stop is configured).
 */
export interface LayoutTabSeg {
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

export interface LayoutImageSeg {
  /** Zip path of the blip — also the `'imagePath' in seg` discriminant that
   *  distinguishes an image segment from text/math/tab segments. */
  imagePath: string;
  /** MIME type of the blip at {@link LayoutImageSeg.imagePath}. */
  mimeType: string;
  widthPt: number;
  heightPt: number;
  /** true = wp:anchor: skip inline flow, draw at absolute page coords */
  anchor: boolean;
  anchorXPt: number;
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** When set, pixels matching this hex color are replaced with alpha=0 before drawing. */
  colorReplaceFrom?: string;
  /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-rectangle crop (fractions 0..1 of
   *  the decoded bitmap). When present the draw paths use the 9-arg
   *  `drawImage` to blit only `[l, t, 1−r, 1−b]` of the bitmap into the display
   *  box. `undefined` ⇒ draw the full bitmap. */
  srcRect?: { l: number; t: number; r: number; b: number };
  /** ECMA-376 §21.2 — when set, this "image" box is actually a DrawingML chart.
   *  The box flows and is sized exactly like an inline picture (via
   *  {@link LayoutImageSeg.widthPt}/{@link LayoutImageSeg.heightPt}), but the
   *  draw site paints it with the shared `renderChart` instead of blitting a
   *  bitmap. `imagePath`/`mimeType` are empty sentinels for a chart seg — no
   *  blip is fetched (the bitmap-prefetch walk keys off `run.type === 'image'`
   *  and never sees a chart run). */
  chart?: ChartModel;
  measuredWidth: number;
}

/** An inline OMML equation. Measured + drawn via the core math engine. */
export interface LayoutMathSeg {
  mathNodes: import('@silurus/ooxml-core').MathNode[];
  display: boolean;
  fontSize: number;  // pt
  color: string | null;
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
export interface LayoutLineBreak {
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
  /** px — intended single-line height (max over segments of the requested
   *  font's win line-height ratio × em), for fonts whose substituted Canvas
   *  metrics understate Word's line spacing. 0 when no segment needs it. */
  intendedSingle: number;
  /** Additional horizontal offset (px) from paraX, caused by wrap-around floats. */
  xOffset: number;
  /** Effective available width (px) for this line after float exclusion. */
  availWidth: number;
  /** When wrap context is active, the absolute canvas Y where this line begins. */
  topY?: number;
  /** Set when at least one segment on this line carries a ruby annotation —
   *  enables docGrid pitch snapping in lineBoxHeight. */
  hasRuby?: boolean;
  /** ECMA-376 §17.3.3.1 — this line is terminated by a MANUAL line break
   *  (`<w:br w:type="textWrapping"/>`). In a justified (`both`) paragraph it is
   *  the end of a logical line and must be left-aligned, not stretched — exactly
   *  like the paragraph's final line (§17.18.44). */
  endsWithBreak?: boolean;
}

/** Additional context passed to layoutLines so it can honor floats on the current page. */
export interface WrapLayoutCtx {
  startPageY: number;   // absolute canvas Y where the first line should start
  paraX: number;        // absolute canvas X of the paragraph's content left edge
  floats: FloatRect[];  // floats active on the current page
  /** Per-line box-height resolver (line natural ascent+descent → total px box height). */
  lineBoxH: (ascentPx: number, descentPx: number, hasRuby?: boolean, intendedSinglePx?: number) => number;
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
 *     - "modern"     → monospace
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
  const result = normalizeFontFamilyUncached(family, fontFamilyClasses);
  perDoc.set(key, result);
  return result;
}

export function normalizeFontFamilyUncached(
  family: string | null,
  fontFamilyClasses: Record<string, string>,
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
      case 'modern':
        return `${head}, "Courier New", monospace`;
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
export function getDefaultFontFamily(para: DocParagraph): string | null {
  for (const run of para.runs) {
    if (run.type === 'text') return (run as unknown as DocxTextRun).fontFamily;
    if (run.type === 'field') return (run as unknown as FieldRun).fontFamily;
  }
  return para.defaultFontFamily ?? null;
}

/** Intended single-line height (px) for an empty paragraph, from its default
 *  font's win line-height ratio. 0 when the font is not in the metrics table. */
export function emptyIntendedSinglePx(para: DocParagraph, scale: number): number {
  return intendedSingleLinePx(getDefaultFontFamily(para), getDefaultFontSize(para) * scale);
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

/** Single source of truth for a text segment's laid-out advance: the natural
 *  `measureText` width plus the character-grid delta (0 unless an active grid
 *  AND a pure-EA segment). EVERY line-break / advance measurement and every draw
 *  path must derive the segment advance from this, so they cannot diverge. */
export function gridWidth(naturalWidthPx: number, text: string, deltaPx: number): number {
  return naturalWidthPx + gridSegDeltaPx(text, deltaPx);
}

export function isGridLineRule(ctx: DocGridCtx | undefined): boolean {
  if (!ctx || !ctx.linePitchPt || ctx.linePitchPt <= 0) return false;
  return ctx.type === 'lines' || ctx.type === 'linesAndChars';
}

/**
 * Compute the total line-box height in px from a line's natural font metrics
 * (fontBoundingBoxAscent + fontBoundingBoxDescent) per ECMA-376 §17.3.1.33.
 *
 *   auto    → natural × value ("single" = 1 natural line, "double" = 2).
 *             When docGrid type=lines|linesAndChars is active, the
 *             multiplier applies against the grid pitch instead, with a
 *             floor of the natural line height.
 *   exact   → value in pt, converted to px (ignores font and grid).
 *   atLeast → max(natural, value in pt × scale).
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
  // Note on ruby paragraphs: an earlier version of this function snapped
  // ruby lines up to the next integer grid pitch on the theory that Word
  // reserves whole grid slots for them. Empirical comparison against Word
  // PNG exports of sample-5 (13.5pt base + 8pt rt on pitch=18) showed
  // Word does NOT snap to integer pitches — the actual line height lands
  // around 2.3× the pitch, which is what `natural` produces directly when
  // the rt reserve in buildSegments is set generously enough. So the
  // ruby-aware logic now lives entirely in the segment-level ascent
  // reservation, and lineBoxHeight stays format-agnostic.
  // A single-spaced line on a docGrid snaps to whole grid CELLS in East Asian
  // text: Word rounds its height UP to the next pitch multiple, so a line taller
  // than one pitch reserves two (or more). sample-9 (pdftotext -bbox of the Word
  // PDF): a 20pt CJK title on a 20pt pitch is 40px = 2 cells; an 11pt body is
  // 20px = 1 cell. A Latin-only line is NOT cell-rounded — it keeps its natural
  // height above a one-cell floor (demo/sample-1: an 18pt heading on an 18pt
  // pitch stays ~20.7px, not 36). ECMA-376 Part 1 only defines the natural ≤
  // pitch case (§17.6.5 / §17.3.1.32); the East-Asian cell rounding for taller
  // lines is Word runtime behaviour, so it is gated on the line's script.
  const gridSingleCell = (): number => eastAsian
    ? Math.max(pitchPx, Math.ceil(glyphNatural / pitchPx) * pitchPx)
    : Math.max(glyphNatural, pitchPx);
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
  if (ls.rule === 'atLeast') return Math.max(natural, ls.value * scale);
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
): { ascent: number; descent: number } {
  const rawAsc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fallbackEmPx * 0.8;
  const rawDesc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fallbackEmPx * 0.2;
  return correctLineMetrics(family, correctionEmPx, rawAsc, rawDesc);
}

/**
 * Height (px) of the paragraph-mark line box for a paragraph that places no
 * inline content on any line. Per ECMA-376 §17.3.1.29 the paragraph mark always
 * produces one line box even when the paragraph has no inline runs; floating
 * objects (§20.4.2.x `wp:anchor`) are removed from the inline flow but never
 * suppress that paragraph-mark line. This is the height used both by the
 * literal empty-paragraph path and by paragraphs whose only segments are
 * wrap-float anchors (which `layoutLines` skips, yielding zero lines).
 */
export function paragraphMarkLineHeight(
  para: DocParagraph,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
  eastAsian = false,
  ctx?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
): number {
  const fs = getDefaultFontSize(para);
  const family = getDefaultFontFamily(para);
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
    // short and the following content rose into a preceding float's wrap band
    // (sample-12: the figure caption wrapped beside the image instead of below
    // it). East Asian documents probe an EA glyph so docGrid cell rounding
    // (lineBoxHeight) reserves whole cells (a 20pt mark on a 20pt pitch → 2
    // cells); others probe a Latin glyph. fontBoundingBox is reported per
    // resolved face (not per glyph), so the probe choice does not change the box
    // for a face that contains it — and the probe is script-matched, so the mark
    // font does. correctedLineMetrics rescales a substituted font to the document
    // font's design box, identical to the text path; the hhea single-line floor
    // (intendedSingleLinePx, via emptyIntendedSinglePx below) then raises tabled
    // fonts — Latin included — to Word's line height.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, family, fontFamilyClasses);
    const m = ctx.measureText(eastAsian ? 'あ' : 'x');
    ctx.font = prevFont;
    // A mark line carries no smallCaps/vertAlign, so fallback == correction size.
    ({ ascent: asc, descent: desc } = correctedLineMetrics(m, family, fs * scale, fs * scale));
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale), eastAsian);
}

/**
 * Resolve the formatting axis that actually governs a run's glyphs.
 *
 * ECMA-376 §17.3.2.30 `w:rtl` marks a run as complex-script. For such a run the
 * complex-script properties take effect — §17.3.2.4 `bCs` (bold), §17.3.2.6
 * `iCs` (italic), §17.3.2.26 `rFonts@cs` (typeface), §17.3.2.39 `szCs` (size) —
 * instead of the non-CS `b`/`i`/`rFonts@ascii`/`sz`, which apply to
 * non-complex (Latin/CJK) text. `bCs`/`iCs` are INDEPENDENT toggles: an absent
 * `bCs` does not inherit `b`'s value, so a complex-script run that carries only
 * `w:b` renders non-bold. (sample-7's date cells carry `b` without `bCs`, and
 * Word draws them at regular weight; the header carries both and is bold.)
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

export function resolveFieldText(f: FieldRun, state: RenderState): string {
  if (f.fieldType === 'page') return String(state.pageIndex + 1);
  if (f.fieldType === 'numPages') return String(state.totalPages);
  return f.fallbackText;
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

/**
 * Binary-search the longest prefix of `text` whose rendered width fits in `maxWidth`.
 * Used for CJK overflow splitting.
 */
export function fitCJKPrefix(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph, 0 when inactive).
  // The fit must compare CELL widths so the grid's char count lands per line —
  // the same `gridWidth` the line box / draw uses, keeping the split consistent.
  gridDeltaPx = 0,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    const prefix = chars.slice(0, mid).join('');
    if (gridWidth(ctx.measureText(prefix).width, prefix, gridDeltaPx) <= maxWidth) lo = mid;
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
 *  that would otherwise push a trailing word to the next line.
 *
 *  This is the ONE budget shared by both sides of the fit contract: the wrap
 *  judgment below admits a word when the line's overflow Δ ≤ SPACE_SHRINK_RATIO ·
 *  Σ(trailing-space widths), and the renderer's draw pass squeezes the same
 *  spaces by the same fraction so the admitted line lands inside its box instead
 *  of overrunning the clip (see `shrinkFitCompression` in text-distribute.ts). */
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
 *      `state.pageIndex + 1` / `state.totalPages`, and the measure state is
 *      frozen at pageIndex 0 / totalPages 1 (buildMeasureState);
 *    - `noteRef` text runs: the label resolves via `state.noteNumbers` /
 *      `state.currentNoteNumber`, which only the paint state carries
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
      if (ft === 'page' || ft === 'numPages') return true;
    } else if (run.type === 'text' && (run as unknown as DocxTextRun).noteRef) {
      return true;
    }
  }
  return false;
}

export function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const pushTextPiece = (
    text: string,
    base: DocxTextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
  ) => {
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
      ? { text: baseRuby.text, fontSizePt: baseRuby.fontSizePt }
      : undefined;
    const revision = (base as DocxTextRun).revision;
    const r = base as DocxTextRun;
    const rtl = r.rtl === true ? true : undefined;

    // ECMA-376 §17.3.2.26 content classification. A run with `w:rtl`
    // (§17.3.2.30) or the `<w:cs/>` toggle (§17.3.2.7) applies complex-script
    // formatting to ALL of its characters; otherwise each character is routed by
    // its Unicode block (Arabic/Hebrew/... → cs; Latin/digits/CJK → ascii/hAnsi).
    // NOTE rFonts@cs (fontFamilyCs) alone is just a font SLOT and must NOT
    // force cs — e.g. sample-1's Heading1 (Latin) has cstheme + szCs=52 but
    // renders at w:sz=24; forcing cs blew its size up to 26pt.
    const forceCs = r.rtl === true || r.cs === true;

    // Complex-script (cs) formatting sources, each falling back to its Latin
    // counterpart when the cs-specific property is absent (the parser already
    // resolves szCs through the full style chain, mirroring a directly-set
    // `w:sz` per §17.3.2.18; bCs/iCs per §17.3.2.3/§17.3.2.17).
    const csFontSize = r.fontSizeCs ?? base.fontSize;
    const csFontFamily = r.fontFamilyCs ?? base.fontFamily;
    const csBold = r.boldCs ?? base.bold;
    const csItalic = r.italicCs ?? base.italic;

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
      segs.push({
        text,
        bold: cs ? csBold : base.bold,
        italic: cs ? csItalic : base.italic,
        underline: base.underline,
        // §17.3.2.40 underline style / colour — carried only on DocxTextRun (a
        // FieldRun draws single). Kept raw ST_Underline; the renderer normalizes
        // to DrawingML §20.1.10.82 at draw time.
        underlineStyle: (base as DocxTextRun).underlineStyle,
        underlineColor: (base as DocxTextRun).underlineColor,
        strikethrough: base.strikethrough,
        fontSize: cs ? csFontSize : base.fontSize,
        color: base.color,
        fontFamily,
        vertAlign,
        measuredWidth: 0,
        smallCaps: reduced,
        joinPrev: gluePending ? true : undefined,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
        background: base.background ?? null,
        colorAuto: r.colorAuto ?? false,
        border: r.border ?? null,
        ruby: firstSeg ? ruby : undefined,
        revision,
        rtl,
        digitsAsAN: digitsAsAN ? true : undefined,
        // §17.3.2.26 declared eastAsia axis — recorded for the text-box line-box
        // floor only (see LayoutTextSeg.eaFloorFamily). Inert for the body path.
        eaFloorFamily: eaFontFamily,
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
  };

  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as DocxTextRun & { type: 'text' };
      // ECMA-376 §17.11: substitute a footnote/endnote reference marker's glyph
      // with the note's resolved sequential number. The body `*Reference` run
      // carries the id; the in-note `*Ref` placeholder carries an empty id, so
      // we fall back to the note number currently being drawn.
      const noteText =
        t.noteRef
          ? (t.noteRef.id
              ? state.noteNumbers?.get(`${t.noteRef.kind}:${t.noteRef.id}`)
              : state.currentNoteNumber)
          : undefined;
      if (t.noteRef) {
        const label = noteText != null ? String(noteText) : (t.text || '');
        if (label.length > 0) pushTextPiece(label, t, t.vertAlign ?? 'super');
        continue;
      }
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) pushTextPiece(parts[i], t, t.vertAlign);
        if (i < parts.length - 1) {
          segs.push({ isTab: true, fontSize: t.fontSize, measuredWidth: 0, bold: t.bold, italic: t.italic });
        }
      }
    } else if (run.type === 'image') {
      const img = run as unknown as ImageRun & { type: 'image' };
      segs.push({
        imagePath: img.imagePath,
        mimeType: img.mimeType,
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        srcRect: img.srcRect ?? undefined,
        measuredWidth: 0,
      });
    } else if (run.type === 'chart') {
      // ECMA-376 §21.2 inline chart. Flow it as an inline picture box of the
      // `<wp:extent>` natural size: the same LayoutImageSeg shape (empty
      // `imagePath`/`mimeType` sentinels so `'imagePath' in seg` routes it
      // through the image measurement/split path) but carrying the ChartModel,
      // which the draw site paints with the shared `renderChart`. Anchor
      // (floating) charts are not drawn yet — skip them so they occupy no box.
      const chartRun = run as unknown as import('./types').ChartRun & { type: 'chart' };
      if (chartRun.anchor) continue;
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: chartRun.widthPt,
        heightPt: chartRun.heightPt,
        anchor: false,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: false,
        anchorYFromPara: false,
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
      const text = resolveFieldText(f, state);
      if (text) pushTextPiece(text, f, f.vertAlign);
    } else if (run.type === 'math') {
      // The parser resolves the paragraph font size; fall back to a nearby run only
      // if it is somehow absent.
      const fontSize = run.fontSize || findNearbyFontSize(runs, runs.indexOf(run));
      segs.push({
        mathNodes: run.nodes,
        display: run.display,
        fontSize,
        color: null,
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
  // (kinsokuAdjustedSplit / crossRunKinsokuRetract), which read `state.kinsoku`.
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
  // when inactive). Folded into every advance via `gridWidth` so line breaking
  // packs the grid's char count per line; the draw paths add the SAME delta.
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
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[] = [];
  let currentWidth = 0;
  // Sum of trailing-space widths of every text token on the current line.
  // Used for two things:
  //   1. Knuth-Plass-style shrink tolerance: a justified line may compress
  //      inter-word spaces by up to SPACE_SHRINK_RATIO, so a candidate word
  //      whose "natural" width would overflow by less than that total shrink
  //      budget is allowed to fit. This is the standard typographic approach
  //      to line-breaking and lets us absorb the ~0.1–0.3 px/glyph advance
  //      bias between Chromium canvas and Word's internal text layout,
  //      matching Word's wrap on long paragraphs.
  //   2. Trailing-space collapse at line end — the last token's trailing
  //      space disappears when no further word is added, so when deciding
  //      whether a candidate word fits we treat it as if it would become the
  //      final word (its own trailing spaces collapsible).
  let lineTotalTrailingW = 0;
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let lineDescent = 0;  // px
  let lineIntendedSingle = 0; // px — max intended single-line height on the line
  let isFirst = true;
  // Effective width/offset for the current line after float exclusion.
  let lineMaxWidth = maxWidth;
  let lineXOffset = 0;
  let currentLineTopY = wrapCtx?.startPageY ?? 0;

  // Minimum clear side-space (px) a line needs before it may START beside a
  // float rather than flow below the band. Word's measured rule is 1 inch
  // (wordMinLineStartPx(scale) — the 1-inch requirement less a half-twip rounding
  // tolerance), INDEPENDENT of the line's content — the same threshold for an
  // empty paragraph mark, a short-token line, and a long-word line (a first word
  // that overruns the ≥1-inch gap is force-broken there by the over-long-word
  // char-break below, matching Word's "AFTE"/"R-10" wrapping). This replaces two
  // former implementation-defined branches — a per-line first-atomic-token width
  // probe for text lines, and a 1-em-of-the-mark-font reservation
  // (`wrapCtx.markEmPx`) for empty/trailing-break lines — both of which
  // mis-placed lines relative to Word (they wedged short-token lines into
  // sub-inch gaps and refused ≥1-inch gaps to long-word lines). See issue #676
  // (fixtures private/sample-19/20/22, pdftotext bbox). Shared by the paint pass
  // and the paginator's two mirror layouts (they call layoutLines with scale 1),
  // so the flow/beside decision agrees across passes.
  const minLineStartWidth = (): number => wordMinLineStartPx(scale);

  // Compute wrap constraints for a new line about to start. Mutates
  // lineXOffset/lineMaxWidth/currentLineTopY. `minWidth` is the smallest clear
  // side-space the upcoming line must have to START here (minLineStartWidth() —
  // Word's 1-inch rule, §676); a free gap narrower than this is treated as
  // unusable and the line is sent below the intervening float(s), which is how
  // Word flows a line that cannot start beside a floating object (there is no
  // ECMA-376 §x.x.x for this trigger — see resolveLineFloatWindow / issue #676).
  const startLine = (minWidth: number = 0): void => {
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    // Small fixed probe height for float intersection (matches the historical
    // wrap behaviour for the topAndBottom skip and horizontal-gap scan).
    const probeH = 10 * scale;
    const win = resolveLineFloatWindow(
      currentLineTopY, minWidth, probeH, wrapCtx.paraX, maxWidth, wrapCtx.floats,
    );
    currentLineTopY = win.topY;
    lineXOffset = win.xOffset;
    lineMaxWidth = win.maxWidth;
  };

  const availW = () => lineMaxWidth - (isFirst ? firstIndent : 0);

  let lineHasRuby = false;
  const flush = (forceHeight?: number, brTerminated = false) => {
    const h = forceHeight !== undefined ? forceHeight : (lineHeight || 10);
    // If the line has no measured content (empty/line-break line), synthesize
    // stable ascent/descent from the effective font size so wrap/baseline math
    // stays consistent with non-empty lines.
    const hasContent = lineAscent > 0 || lineDescent > 0;
    const asc = hasContent ? lineAscent : h * scale * 0.8;
    const desc = hasContent ? lineDescent : h * scale * 0.2;
    lines.push({
      segments: currentLine,
      height: h,
      ascent: asc,
      descent: desc,
      intendedSingle: lineIntendedSingle,
      xOffset: lineXOffset,
      availWidth: lineMaxWidth,
      topY: wrapCtx ? currentLineTopY : undefined,
      hasRuby: lineHasRuby,
      endsWithBreak: brTerminated,
    });
    if (wrapCtx) {
      currentLineTopY += wrapCtx.lineBoxH(asc, desc, lineHasRuby, lineIntendedSingle);
    }
    currentLine = [];
    currentWidth = 0;
    lineTotalTrailingW = 0;
    lineHeight = 0;
    lineAscent = 0;
    lineDescent = 0;
    lineIntendedSingle = 0;
    lineHasRuby = false;
    isFirst = false;
    startLine(minLineStartWidth());
  };

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
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
    if (desc > lineDescent) lineDescent = desc;
    if (!('isTab' in s) && !('imagePath' in s) && !('mathNodes' in s)) {
      const ts = s as LayoutTextSeg;
      if (ts.ruby) lineHasRuby = true;
      // Intended single-line height for fonts whose substituted Canvas metrics
      // understate Word's line spacing (font-metrics.ts). 0 for untabled fonts.
      // Small caps (non-super/sub) keep the FULL run size here so the line box
      // follows the run size, not the 2pt-reduced glyphs (§17.3.2.33).
      const intendedEm = ts.smallCaps && !ts.vertAlign ? ts.fontSize * scale : effectiveFontPx(ts);
      const intended = intendedSingleLinePx(ts.fontFamily, intendedEm);
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
    }
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

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    return ctx.measureText(s.text);
  };

  // The segment's laid-out ADVANCE (= its measuredWidth): natural width plus the
  // character-grid delta. This is the SINGLE source of truth shared with the
  // draw paths (gridWidth) — every line-break / fit / tab measurement uses it so
  // line wrapping packs the grid's char count and the box matches what is drawn.
  // `measureAdvanceWithTrailingSpace` (not the raw `ctx.measureText().width`) so a trailing
  // inter-word space keeps its advance on engines whose measureText trims trailing
  // whitespace — see the core helper (@silurus/ooxml-core → text/measure-space.ts).
  const segAdvance = (s: LayoutTextSeg): number => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    return gridWidth(measureAdvanceWithTrailingSpace(ctx, s.text), s.text, gridDeltaPx);
  };
  // Grid advance of an arbitrary string under a segment's font (for split
  // prefixes/tails). Selects the font, then applies the same gridWidth model.
  const strAdvance = (s: LayoutTextSeg, text: string): number => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    return gridWidth(measureAdvanceWithTrailingSpace(ctx, text), text, gridDeltaPx);
  };

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => {
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('imagePath' in q) return q.widthPt * scale;
    if ('mathNodes' in q) return q.measuredWidth || 0;
    if ('lineBreak' in q) return 0;
    return segAdvance(q);
  };

  // Use an explicit queue so CJK split-tails can be re-queued
  const queue: LayoutSeg[] = [...segs];

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph that new line is an EMPTY line that still
  // occupies one line height (Word reserves it; visible e.g. as extra table
  // row height). Track the trailing break so it can be flushed after the loop.
  let trailingBreakFontSize: number | null = null;

  // Establish the first line's wrap window now that the content queue exists.
  startLine(minLineStartWidth());

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
            flush();
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
              const w = gridWidth(m.width, q.text, gridDeltaPx);
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
            const w = gridWidth(m.width, q.text, gridDeltaPx);
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
        flush();
        queue.unshift(seg);
        continue;
      }
      if (currentWidth + tabWidth > availW() && currentLine.length > 0) {
        flush();
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
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, h, asc, 0);
      continue;
    }

    // ── Math segment ─────────────────────────────────────
    if ('mathNodes' in seg) {
      const render = mathRenders.get(seg.mathNodes);
      if (!render) { seg.measuredWidth = 0; continue; }
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
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, seg.fontSize, lineAsc, lineDesc);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    const s = seg as LayoutTextSeg;
    const m = measureText(s); // sets ctx.font; also feeds ascent/descent below
    // Advance = natural width + character-grid delta (the SINGLE model shared
    // with the draw paths; 0 unless an active grid AND a pure-EA segment).
    // `measureAdvanceWithTrailingSpace` (not `m.width`) so a trailing inter-word space keeps
    // its advance on engines whose measureText trims trailing whitespace.
    const w = gridWidth(measureAdvanceWithTrailingSpace(ctx, s.text), s.text, gridDeltaPx);
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
    const corrected = correctedLineMetrics(metricM, s.fontFamily, fullPx, metricEmPx);
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
      asc = asc + s.ruby.fontSizePt * scale * 1.5;
    }
    // Wrap-fit check uses two standard typographic allowances:
    //   1. Trailing-space collapse: if this word becomes the last on the
    //      line, its trailing space (if any) collapses. We subtract it from
    //      the width used to test fit.
    //   2. Knuth-Plass shrink tolerance: a line may compress each inter-word
    //      space by up to SPACE_SHRINK_RATIO (25%) of its natural width without
    //      harming readability — the module constant, shared with the draw pass
    //      so a line admitted here is squeezed by the same budget when painted
    //      (shrinkFitCompression in text-distribute.ts) rather than overrunning
    //      its box. See the SPACE_SHRINK_RATIO doc above.
    const trimmed = s.text.replace(/ +$/, '');
    // Subtract the GRID width of the trimmed text (not the natural width) so the
    // grid delta on EA glyphs cancels and trailingSpaceW is the bare space
    // advance — keeping `w` and `wForFit` on the one advance model.
    const trailingSpaceW = s.text.endsWith(' ')
      ? w - gridWidth(ctx.measureText(trimmed).width, trimmed, gridDeltaPx)
      : 0;
    const wForFit = w - trailingSpaceW;
    const shrinkBudget = lineTotalTrailingW * SPACE_SHRINK_RATIO;

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
      !hasCJKBreakOpportunity(s.text)
    ) {
      let groupW = w;
      let groupTrail = trailingSpaceW;
      for (let k = 0; k < queue.length && (queue[k] as LayoutTextSeg).joinPrev; k++) {
        const f = queue[k] as LayoutTextSeg;
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
            groupW += strAdvance(f, chars.slice(0, p).join(''));
            groupTrail = 0;
            break;
          }
          // Entirely non-starters (no breakable rest): fall through to full-add.
        }
        const fw = segAdvance(f);
        groupW += fw;
        const ft = f.text.replace(/ +$/, '');
        groupTrail = f.text.endsWith(' ') ? fw - gridWidth(ctx.measureText(ft).width, ft, gridDeltaPx) : 0;
      }
      if (currentWidth + (groupW - groupTrail) > availW() + shrinkBudget) flush();
    }

    if (currentWidth + wForFit <= availW() + shrinkBudget) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
    } else if (hasCJKBreakOpportunity(s.text)) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
      // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
      //  separate: it sums per-char advances, whereas this path uses substring
      //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
      const available = availW() - currentWidth;
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
      const rawPrefix = available > 0 ? fitCJKPrefix(ctx, s.text, available, gridDeltaPx) : '';
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
      const split = kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, minSplit);
      const prefix = allChars.slice(0, split).join('');
      if (prefix.length > 0) {
        // Grid advance for the head piece — the same model as the line box / draw.
        const pw = strAdvance(s, prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pw };
        addToLine(headSeg, pw, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
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
            retracted = { ...lastText, text: tailText, measuredWidth: strAdvance(lastText, tailText) };
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
        flush();
        queue.unshift(s);
        if (retracted) queue.unshift(retracted);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const firstChar = [...s.text][0] ?? '';
        if (firstChar) {
          const fw = strAdvance(s, firstChar);
          const headSeg: LayoutTextSeg = { ...s, text: firstChar, measuredWidth: fw };
          addToLine(headSeg, fw, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
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
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
      const allChars = [...s.text];
      let split = available > 0 ? [...fitCJKPrefix(ctx, s.text, available, gridDeltaPx)].length : 0;
      if (split < 1) split = 1;
      if (split >= allChars.length) {
        // The visible glyphs actually fit (only a trailing space pushed it over the
        // fit test) — place the word whole.
        s.measuredWidth = w;
        addToLine(s, w, h, asc, desc);
      } else {
        const prefix = allChars.slice(0, split).join('');
        const pw = strAdvance(s, prefix);
        addToLine({ ...s, text: prefix, measuredWidth: pw }, pw, h, asc, desc);
        queue.unshift({ ...s, text: allChars.slice(split).join(''), measuredWidth: 0 });
      }
    } else {
      // Latin word does not fit on the current (non-empty) line: move it to a fresh
      // line and re-process. There it either fits, or — when it is wider than the
      // whole column — the empty-line branch above breaks it at the character level
      // (overflow-wrap). Re-queueing rather than force-adding is what lets that
      // over-long-word path run instead of letting the word spill the column.
      flush();
      queue.unshift(s);
    }
  }

  if (currentLine.length > 0) flush();
  // Trailing <w:br/>: emit the empty line it opened (§17.3.3.1).
  else if (trailingBreakFontSize !== null) flush(trailingBreakFontSize);

  return lines;
}

/** Phase 4-1 B2 Stage 2 — rehydrate the paginator's scale-1 stamped lines into
 *  the paint scale, keeping the scale-1 line PARTITION but re-deriving all
 *  hinting-sensitive geometry at the paint scale.
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
 *  gridWidth, box via correctedLineMetrics, the §17.3.2.33 small-caps full-size
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
 *  byte-for-byte as in Stage 1. */
export function rescaleLayoutLines(
  lines: LayoutLine[],
  scale: number,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string>,
  gridDeltaPx: number,
): LayoutLine[] {
  if (scale === 1) return lines;

  // Per-text-segment measurement at the PAINT scale — the SAME model layoutLines
  // uses (segAdvance + the text-segment metric block), so measure == draw.
  const measureTextSeg = (s: LayoutTextSeg): { advance: number; asc: number; desc: number; intended: number } => {
    const effPx = calcEffectiveFontPx(s, scale);
    ctx.font = buildFont(s.bold, s.italic, effPx, s.fontFamily, fontFamilyClasses);
    const m = ctx.measureText(s.text);
    // `measureAdvanceWithTrailingSpace` (not `m.width`) so a trailing inter-word space keeps
    // its advance at the paint scale too, on engines whose measureText trims
    // trailing whitespace (must match the layoutLines advance so measure==draw).
    const advance = gridWidth(measureAdvanceWithTrailingSpace(ctx, s.text), s.text, gridDeltaPx);
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
    const corrected = correctedLineMetrics(metricM, s.fontFamily, fullPx, metricEmPx);
    // §17.3.3.25 — ruby reserves extra ascent room (rt size × 1.5), same as layoutLines.
    const asc = s.ruby ? corrected.ascent + s.ruby.fontSizePt * scale * 1.5 : corrected.ascent;
    // Intended single-line floor (font-metrics.ts) — small caps keep the FULL run
    // size here too (addToLine's intendedEm).
    const intendedEm = s.smallCaps && !s.vertAlign ? fullPx : effPx;
    const intended = intendedSingleLinePx(s.fontFamily, intendedEm);
    return { advance, asc, desc: corrected.descent, intended };
  };

  return lines.map((l) => {
    let asc = 0;
    let desc = 0;
    let intended = 0;
    let hasText = false;
    const segments = l.segments.map((s) => {
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
        return { ...s, measuredWidth: s.widthPt * scale };
      }
      if ('mathNodes' in s) {
        const copy = { ...s, measuredWidth: s.measuredWidth * scale } as LayoutMathSeg;
        copy.mathAscent *= scale;
        copy.mathDescent *= scale;
        asc = Math.max(asc, copy.mathAscent, s.fontSize * scale * 0.8);
        desc = Math.max(desc, copy.mathDescent, s.fontSize * scale * 0.2);
        return copy;
      }
      const t = s as LayoutTextSeg;
      const mm = measureTextSeg(t);
      hasText = true;
      if (mm.asc > asc) asc = mm.asc;
      if (mm.desc > desc) desc = mm.desc;
      if (mm.intended > intended) intended = mm.intended;
      return { ...t, measuredWidth: mm.advance };
    });
    // Empty/synthetic line (no text/math contributing metrics): fall back to the
    // ×scale of the stamped box (layoutLines synthesises h·scale·0.8/0.2 there).
    if (!hasText && asc === 0 && desc === 0) {
      asc = l.ascent * scale;
      desc = l.descent * scale;
      intended = l.intendedSingle * scale;
    }
    return {
      ...l,
      segments,
      ascent: asc,
      descent: desc,
      intendedSingle: intended,
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
 *  eastAsia font axes §17.3.2.26, bold, italic); every field the shape model
 *  lacks (underline/strike/highlight/border/ruby/rtl/cs/…) takes its neutral
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
