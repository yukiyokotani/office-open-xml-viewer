import type { CjkLang } from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import {
  ARABIC_SUBSTITUTE_FACES,
  NASKH_ARABIC_SUBSTITUTE_FACES,
  quoteCssFontFamilies,
} from '@silurus/ooxml-core/internal/font-stack';
import type { DocxTextRun, FieldRun } from '../types';
import type { CanvasFontRoute } from '@silurus/ooxml-core';
import { classifyCjkFont, cjkFallbackChain, NON_CJK_SANS_FALLBACKS, NON_CJK_SERIF_FALLBACKS, classifyFontGeneric, normalizeFontMetricFamily, canvasFontRouteFaces, canvasFontString } from '@silurus/ooxml-core';
import type { ParagraphLayoutSource } from '../layout/text.js';
import { type LayoutTextSeg } from './model.js';

// ── Math (OMML) rendering via MathJax ───────────────────────────────────────
// Each equation is converted OMML AST -> MathML -> MathJax SVG, then rasterized
// to an auxiliary Canvas once (async, before pagination). Layout reads cached
// em-extents synchronously; drawing blits the Canvas. Skipped entirely for
// math-free documents.
/** Arabic-script faces that hosts rarely ship, which the vendor-neutral
 *  substitution policy maps to Noto Naskh/Sans Arabic (core
 *  `OFFICE_FONT_SUBSTITUTIONS`). A source run whose font is
 *  one of these contains both Arabic and Latin/digit glyphs in one requested
 *  face, so the fallback chain must keep both scripts stylistically
 *  consistent (Arabic substitute first, serif Latin companion before the sans
 *  generics) rather than letting Latin/digits leak to a CJK sans face. */
export const ARABIC_SUBSTITUTE_FONTS = ARABIC_SUBSTITUTE_FACES;


/** Naskh-style traditional Arabic faces ship a serif Latin companion; the
 *  geometric/modern ones pair with a sans Latin. Drives whether an Arabic-font
 *  run's Latin+digits route to Noto Naskh Arabic (serif-like) or Noto Sans
 *  Arabic, and which Latin serif/sans companion follows. */
export const NASKH_SERIF_ARABIC_FONTS = NASKH_ARABIC_SUBSTITUTE_FACES;


export function isArabicSubstituteFont(family: string): boolean {
  return ARABIC_SUBSTITUTE_FONTS.has(family.toLowerCase());
}


/** Quote each family for a CSS font-family list. */
export const quoteAll = quoteCssFontFamilies;


/** Generic Arabic fallback families (loaded when a font source serves them). */
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
export function sansTail(cjk: ReturnType<typeof classifyCjkFont>, fallback?: CjkLang): string {
  const cjkPart =
    (cjk ?? fallback) && (cjk ?? fallback) !== 'jp'
      ? cjkFallbackChain((cjk ?? fallback) as CjkLang, 'sans')
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
export function serifTail(cjk: ReturnType<typeof classifyCjkFont>, fallback?: CjkLang): string {
  const cjkPart =
    (cjk ?? fallback) && (cjk ?? fallback) !== 'jp'
      ? cjkFallbackChain((cjk ?? fallback) as CjkLang, 'serif')
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


// fontTable names and run names refer to the same family regardless of casing.
// Normalize each immutable per-document facts map once rather than scanning it
// for every glyph's fallback chain. This uses the same key as metric resolution.
export const normalizedFontTableFacts = new WeakMap<Record<string, string>, Map<string, string | null>>();


export function fontTableFact(facts: Record<string, string>, family: string): string | undefined {
  let normalized = normalizedFontTableFacts.get(facts);
  if (!normalized) {
    normalized = new Map<string, string | null>();
    for (const [name, value] of Object.entries(facts)) {
      const key = normalizeFontMetricFamily(name);
      if (!normalized.has(key)) normalized.set(key, value);
      // Conflicting records for one case-insensitive family do not establish
      // an authoritative classification or pitch; use the ordinary fallback.
      else if (normalized.get(key) !== value) normalized.set(key, null);
    }
    normalizedFontTableFacts.set(facts, normalized);
  }
  return normalized.get(normalizeFontMetricFamily(family)) ?? undefined;
}


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
  cjkFallback?: CjkLang,
): string {
  if (!family || family === 'sans-serif') return sansTail(null, cjkFallback);
  if (family === 'serif') return serifTail(null, cjkFallback);
  const monoTail = cjkFallback
    ? `"Courier New", ${quoteAll(cjkFallbackChain(cjkFallback, 'sans'))}, monospace`
    : '"Courier New", monospace';
  if (family === 'monospace') return monoTail;

  const head = quoteCssFontFamily(family);
  const lower = family.toLowerCase();

  // CJK language inferred from the font name (null for plain Latin faces). For a
  // non-JP CJK language the matching Noto CJK leads the fallback tail so shared
  // Han glyphs render with that language's shapes; see core/fonts/scripts.ts.
  const cjk = classifyCjkFont(family);

  // 1) Authoritative classification from word/fontTable.xml §17.8.3.10.
  const tableClass = fontTableFact(fontFamilyClasses, family);
  if (tableClass && tableClass !== 'auto') {
    switch (tableClass) {
      case 'roman':
        return `${head}, ${serifTail(cjk, cjkFallback)}`;
      case 'swiss':
        return `${head}, ${sansTail(cjk, cjkFallback)}`;
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
        if (fontTableFact(fontFamilyPitches, family) === 'fixed') {
          if (cjk != null) {
            const cjkFallbacks = cjk === 'jp'
              ? ['Yu Gothic', 'YuGothic', 'Hiragino Sans', 'Meiryo', 'Noto Sans JP']
              : cjkFallbackChain(cjk, 'sans');
            return `${head}, ${quoteAll([...cjkFallbacks, 'Courier New'])}, monospace`;
          }
          return `${head}, ${monoTail}`;
        }
        break;
      }
      default:
        // script / decorative — fall through to name-pattern matching
        break;
    }
  }

  // When fontTable has no usable class, the optional Arabic substitute route
  // keeps Arabic and Latin/digits in one script-compatible face. This is a
  // fallback selection policy, not evidence that Noto has the authored face's
  // advances or Office metrics. An explicit fontTable family always wins above.
  if (isArabicSubstituteFont(family)) {
    if (NASKH_SERIF_ARABIC_FONTS.has(lower)) {
      return cjkFallback
        ? `${head}, "Noto Naskh Arabic", "Noto Sans Arabic", "Noto Serif", ${sansTail(null, cjkFallback).replace(/sans-serif$/, 'serif')}`
        : `${head}, "Noto Naskh Arabic", "Noto Sans Arabic", "Noto Serif", "Noto Sans JP", "Hiragino Sans", serif`;
    }
    return cjkFallback
      ? `${head}, "Noto Sans Arabic", "Noto Naskh Arabic", ${sansTail(null, cjkFallback)}`
      : `${head}, "Noto Sans Arabic", "Noto Naskh Arabic", "Noto Sans JP", "Hiragino Sans", sans-serif`;
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
    return `${head}, ${serifTail(cjk, cjkFallback)}`;
  }
  if (generic === 'mono') {
    // Mirror the fontTable `modern` branch's monospace fallback. NEW for the
    // name path: core now detects consolas/courier/等幅 etc. as mono.
    return `${head}, ${monoTail}`;
  }

  // Japanese system-font hints (only meaningful for JP / Latin faces; a non-JP
  // CJK face skips these so its matching Noto CJK leads the tail).
  if (cjk == null || cjk === 'jp') {
    if (lower.includes('meiryo') || family.includes('メイリオ')) {
      return `${head}, "Meiryo UI", "Meiryo", ${sansTail(cjk, cjkFallback)}`;
    }
    if (family.includes('游ゴシック') || /\byu\s*gothic\b/i.test(family) || lower.includes('yugothic')) {
      return `${head}, "Yu Gothic", "YuGothic", ${sansTail(cjk, cjkFallback)}`;
    }
    if (lower.includes('ipa')) {
      return `${head}, "IPAexGothic", ${sansTail(cjk, cjkFallback)}`;
    }
    if (lower.includes('segoe')) {
      return cjkFallback
        ? `${head}, "Segoe UI", ${sansTail(null, cjkFallback)}`
        : `${head}, "Segoe UI", ${quoteAll([...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
    }
  }
  return `${head}, ${sansTail(cjk, cjkFallback)}`;
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


/** Selected-resource or native-reference design single-line floor. */
export function segmentIntendedSingleLinePx(
  segment: LayoutTextSeg,
  emPx: number,
  eastAsian = false,
): number {
  const resourceRatio = eastAsian
    ? segment.resolvedEastAsianLineHeightRatio ?? segment.resolvedLineHeightRatio ?? 0
    : segment.resolvedLineHeightRatio ?? 0;
  return resourceRatio * emPx;
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
  return resourceRatio * emPx;
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
 * paragraph's style-resolved default family. */
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

/** Whether two segments' routes resolve to the same faces, whatever language
 * each shapes in (core canvasFontRouteFaces). The rules that read a face's
 * own metrics (line geometry, the average-width space floor) compare faces;
 * every measured or shaped comparison uses the route fingerprint itself. */
export function sameRouteFaces(
  left: Readonly<CanvasFontRoute> | undefined,
  right: Readonly<CanvasFontRoute> | undefined,
): boolean {
  return left === right
    || (left !== undefined && right !== undefined
      && canvasFontRouteFaces(left).fingerprint === canvasFontRouteFaces(right).fingerprint);
}
