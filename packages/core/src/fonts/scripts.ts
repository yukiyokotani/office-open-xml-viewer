/**
 * Shared script → Noto fallback definitions used by the docx / pptx / xlsx
 * renderers and their Google-Fonts preload maps.
 *
 * The library ships NO font binaries; these Noto families are loaded on demand
 * from the Google Fonts CDN only when the caller opts in with
 * `useGoogleFonts: true`. The renderers append the relevant Noto family names
 * to the CSS font stack so the browser's per-glyph fallback resolves CJK /
 * Cyrillic / Thai / Devanagari / Hebrew glyphs to a real web font instead of an
 * oversized OS face or tofu.
 *
 * ## Why CJK is handled differently from the other scripts
 *
 * KR / SC / TC / JP share the Han (漢字) block but render those shared
 * codepoints with DIFFERENT glyph shapes (e.g. 直/海/骨 differ between
 * Japanese, Simplified Chinese, Traditional Chinese, Korean conventions).
 * Canvas text falls back PER GLYPH down the font-family chain, so whichever
 * Noto CJK comes first in the chain decides the shape of every shared Han
 * glyph. We therefore cannot just append all four — we must put the Noto CJK
 * matching the document's CJK language FIRST. {@link classifyCjkFont} derives
 * that language from the requested Office font name (and CJK-only Unicode
 * markers in the name). The other three CJK Notos still follow, so a stray
 * codepoint missing from the primary face (rare) resolves to a sibling rather
 * than tofu.
 *
 * The non-CJK scripts (Cyrillic, Thai, Devanagari, Hebrew) do NOT share glyphs
 * with Latin or with each other, so the browser's per-glyph fallback picks the
 * right face no matter the order. They are appended unconditionally to the
 * generic sans / serif tails ({@link NON_CJK_SANS_FALLBACKS} /
 * {@link NON_CJK_SERIF_FALLBACKS}).
 *
 * Hebrew is RTL; only the FONT is supplied here — bidi/visual ordering is
 * handled by the existing bidi line logic in each package.
 */

export type CjkLang = 'kr' | 'sc' | 'tc' | 'jp';
export type FontVariant = 'sans' | 'serif';

/**
 * Classify a requested Office font name into a CJK language, or `null` when it
 * is not a known CJK face. Based on the well-known default East-Asian fonts
 * Office ships per locale (names verified against the Windows/Office font set):
 *
 * - Korean  : Malgun Gothic (맑은 고딕), Batang/Batangche, Gulim/Gulimche,
 *             Dotum/Dotumche (돋움), Gungsuh, New Gulim. Plus the Hangul
 *             Jamo/Syllables marker in the name.
 * - Simplified Chinese : SimSun (宋体), NSimSun, SimHei (黑体), Microsoft YaHei
 *             (微软雅黑), DengXian (等线), FangSong (仿宋), KaiTi (楷体), STSong,
 *             STKaiti, STFangsong, STHeiti, STXihei, LiSu, YouYuan.
 * - Traditional Chinese : PMingLiU/MingLiU (細明體/新細明體), Microsoft JhengHei
 *             (微軟正黑體), DFKai-SB (標楷體), MingLiU_HKSCS, Kaiti TC.
 * - Japanese : Yu Gothic/Mincho (游ゴシック/游明朝), Meiryo (メイリオ),
 *             MS Gothic/Mincho (ＭＳ ゴシック/明朝), Hiragino (ヒラギノ),
 *             plus the Hiragana/Katakana marker in the name.
 *
 * Detection is name + Unicode-marker based (no per-sample tuning). Where Latin
 * transliterations collide we disambiguate on the more specific token (e.g.
 * "Microsoft JhengHei" → TC vs "Microsoft YaHei" → SC; "MingLiU" → TC).
 */
export function classifyCjkFont(family: string | null | undefined): CjkLang | null {
  if (!family) return null;
  const l = family.toLowerCase();

  // --- Unicode-script markers in the name (script-exclusive ranges) ---
  // Hangul (Jamo U+1100–11FF, Compatibility Jamo U+3130–318F, Syllables
  // U+AC00–D7AF) → Korean. 돋움 / 맑은 고딕 etc.
  if (/[ᄀ-ᇿ㄰-㆏가-힯]/.test(family)) return 'kr';
  // Hiragana (U+3040–309F) / Katakana (U+30A0–30FF) → Japanese. メイリオ etc.
  if (/[぀-ヿ]/.test(family)) return 'jp';

  // --- Latin transliterations + Han names. Order matters: TC tokens that
  //     contain an SC-looking substring are matched first. ---

  // Traditional Chinese (check before SC: "jhenghei" must not be caught as a
  // generic "hei", and MingLiU family before any "ming" heuristics).
  if (
    /jhenghei|微軟正黑|新細明|細明|pmingliu|mingliu|dfkai|標楷|華康|cns11643|kaiti tc|ming\s*liu/.test(l) ||
    /新細明體|細明體|標楷體|微軟正黑體|華康/.test(family)
  ) {
    return 'tc';
  }

  // Simplified Chinese
  if (
    /simsun|nsimsun|simhei|simkai|simfang|yahei|dengxian|fangsong|kaiti|youyuan|lisu|stsong|stkaiti|stfangsong|stheiti|stxihei|stzhongsong|songti sc|heiti sc|微软雅黑/.test(l) ||
    /宋体|黑体|楷体|仿宋|等线|微软雅黑|隶书|幼圆/.test(family)
  ) {
    return 'sc';
  }

  // Korean (Latin transliterations)
  if (/malgun|batang|gulim|dotum|gungsuh|nanum|new gulim|hancom|hy(gothic|graphic|namu)?/.test(l)) {
    return 'kr';
  }

  // Japanese (Latin transliterations). Yu/MS Gothic+Mincho, Meiryo, Hiragino.
  if (
    /\bmeiryo\b|\byu\s*(gothic|mincho)\b|yugothic|yumincho|hiragino|\bms\s*(gothic|mincho|pgothic|pmincho|ui\s*gothic)\b|\bms[pg]?(gothic|mincho)\b|ipa(ex)?(gothic|mincho)|noto\s+(sans|serif)\s+jp|游ゴシック|游明朝|ＭＳ|メイリオ|ヒラギノ/.test(l) ||
    /游ゴシック|游明朝|ＭＳ ゴシック|ＭＳ 明朝|ＭＳ Ｐゴシック|メイリオ|ヒラギノ/.test(family)
  ) {
    return 'jp';
  }

  return null;
}

/** CSS generic class of a font face, inferred from its name. */
export type FontGenericClass = 'serif' | 'sans' | 'mono';

/**
 * Classify a font *name* into a CSS generic class (serif / sans / monospace).
 * The shared serif/sans name heuristic for the pptx, xlsx, and docx renderers.
 * DOCX layers word/fontTable.xml `<w:family>` priority and its format-specific
 * fallback-chain construction on top; only the final name-pattern decision is
 * shared here.
 *
 * This is the NAME-pattern fallback only. A caller that has authoritative class
 * data — Word's word/fontTable.xml `<w:family>` §17.8.3.10 (roman/swiss/modern)
 * — MUST consult that first and use this only for faces absent from the table or
 * marked "auto". The token set is the union of all three renderers' prior
 * regexes; no package-specific name guessing is introduced.
 *
 * - mono : mono / courier / consolas / 等幅 …
 * - serif: Latin serifs (Times, Cambria/Caladea, Georgia, Garamond, Century,
 *          Palatino, Antiqua (Book Antiqua & other Palatino clones — "Antiqua"
 *          is the typographic term for a Roman/serif face), Didot, Bodoni,
 *          Playfair, "… Serif", roman) AND CJK
 *          song/ming/kai/fangsong faces (SimSun 宋 / Batang / PMingLiU 細明 /
 *          KaiTi 楷 / FangSong 仿宋 / *Mincho 明朝) — so the per-language Noto CJK
 *          fallback picks its serif variant.
 * - sans : everything else (gothic, hei, kaku, round, grotesk, …).
 */
export function classifyFontGeneric(family: string | null | undefined): FontGenericClass {
  if (!family) return 'sans';
  const l = family.toLowerCase();
  if (/mono|courier|consolas|等幅|gothic_m/.test(l)) return 'mono';
  // `century(?!\s*gothic)` keeps the serif Century family (Century, Century
  // Schoolbook, …) but excludes the geometric SANS "Century Gothic" — the one
  // union token where a serif name collides with a real sans face. (The
  // authoritative split is the §17.8.3.10 <w:family> class the caller checks
  // first; this only refines the name-pattern fallback.)
  if (
    /roman|times|cambria|caladea|georgia|garamond|century(?!\s*gothic)|palatino|antiqua|didot|bodoni|playfair|source serif|noto serif|min\s*cho|明朝体|明朝|song|sung|simsun|nsimsun|batang|gungsuh|ming\s*liu|mingliu|pmingliu|fang\s*song|fangsong|kai\s*ti|kaiti|simkai|simfang|stsong|stkaiti|stfangsong|stzhongsong|新細明|細明|宋体|楷体|楷體|仿宋|標楷|游明朝|ＭＳ 明朝|ms mincho|yu mincho|hiragino mincho|ヒラギノ明朝/.test(
      l,
    ) ||
    /新細明體|細明體|宋体|明朝|楷体|楷體|仿宋|標楷體|游明朝|ＭＳ 明朝/.test(family)
  ) {
    return 'serif';
  }
  return 'sans';
}

/**
 * True when a code point belongs to a complex-script (RTL/`cs`) Unicode block —
 * Hebrew, Arabic (incl. supplements / extended / presentation forms), Syriac,
 * Thaana, NKo, Samaritan, Mandaic, and the Plane-1 RTL / Arabic-math blocks.
 * Single source of truth for ECMA-376 §17.3.2.26's per-character complex-script
 * (`cs`) font-axis split: Word routes such characters to the run's cs formatting
 * (`w:szCs`/`w:rFonts w:cs`/`w:bCs`/`w:iCs`) instead of the Latin (`ascii`/
 * `hAnsi`) formatting. Latin, European digits, punctuation and CJK are NOT cs.
 *
 * Distinct from {@link scriptPreloadNamesForText}'s ranges, which decide WHICH
 * Noto web font to load (Arabic vs Hebrew separately, only the scripts we ship)
 * — a narrower, different-purpose set. This is the broad "is this a cs-axis
 * glyph" predicate, exported so any renderer can apply the same split.
 */
export function isComplexScriptCodePoint(cp: number): boolean {
  return (
    (cp >= 0x0590 && cp <= 0x05ff) || // Hebrew
    (cp >= 0x0600 && cp <= 0x06ff) || // Arabic
    (cp >= 0x0700 && cp <= 0x074f) || // Syriac
    (cp >= 0x0750 && cp <= 0x077f) || // Arabic Supplement
    (cp >= 0x0780 && cp <= 0x07bf) || // Thaana
    (cp >= 0x07c0 && cp <= 0x07ff) || // NKo
    (cp >= 0x0800 && cp <= 0x083f) || // Samaritan
    (cp >= 0x0840 && cp <= 0x085f) || // Mandaic
    (cp >= 0x0860 && cp <= 0x08ff) || // Syriac Supp. / Arabic Extended-A/B
    (cp >= 0xfb1d && cp <= 0xfb4f) || // Hebrew presentation forms
    (cp >= 0xfb50 && cp <= 0xfdff) || // Arabic Presentation Forms-A
    (cp >= 0xfe70 && cp <= 0xfeff) || // Arabic Presentation Forms-B
    (cp >= 0x10800 && cp <= 0x10fff) || // Plane-1 RTL blocks
    (cp >= 0x1e800 && cp <= 0x1efff)    // Mende/Adlam/Arabic Math
  );
}

/**
 * Ordered Noto CJK family names for a given language, primary face first so the
 * browser resolves shared Han glyphs to that language's shapes. The other three
 * follow as a last-resort so a codepoint absent from the primary face does not
 * fall to tofu.
 */
export function cjkFallbackChain(lang: CjkLang, variant: FontVariant): string[] {
  const prefix = variant === 'serif' ? 'Noto Serif' : 'Noto Sans';
  const order: Record<CjkLang, CjkLang[]> = {
    // Primary language first; the rest in a stable order. The trailing list is
    // a tofu safety net only — shared Han glyphs are already taken by the head.
    jp: ['jp', 'sc', 'tc', 'kr'],
    sc: ['sc', 'tc', 'jp', 'kr'],
    tc: ['tc', 'sc', 'jp', 'kr'],
    kr: ['kr', 'jp', 'sc', 'tc'],
  };
  const suffix: Record<CjkLang, string> = { kr: 'KR', sc: 'SC', tc: 'TC', jp: 'JP' };
  return order[lang].map((x) => `${prefix} ${suffix[x]}`);
}

/**
 * Noto faces appended to a SANS generic tail so non-CJK, non-Latin scripts
 * resolve to a real web font via per-glyph fallback. Cyrillic is covered by the
 * un-suffixed "Noto Sans"; Thai / Devanagari / Hebrew need their script-specific
 * Noto family because "Noto Sans" does not embed those scripts.
 *
 * No glyph-shape collision with Latin or with each other, so order is
 * immaterial and these are safe to append unconditionally.
 */
export const NON_CJK_SANS_FALLBACKS = [
  'Noto Sans',          // Latin + Cyrillic + Greek
  'Noto Sans Hebrew',   // Hebrew (RTL; bidi handled separately)
  'Noto Sans Thai',
  'Noto Sans Devanagari',
] as const;

/** Serif counterpart of {@link NON_CJK_SANS_FALLBACKS}. Devanagari/Thai ship
 *  primarily as sans in Office serif contexts; Cyrillic + Hebrew get serif
 *  Noto faces. (Noto Serif Thai/Devanagari are omitted to keep the serif tail
 *  small — sans coverage above still prevents tofu for those scripts.) */
export const NON_CJK_SERIF_FALLBACKS = [
  'Noto Serif',         // Latin + Cyrillic + Greek
  'Noto Serif Hebrew',  // Hebrew (RTL; bidi handled separately)
] as const;

// ---- Google Fonts preload definitions (script Noto families) ----

const w = (q: string) =>
  `https://fonts.googleapis.com/css2?family=${q}:wght@400;700&display=swap`;

/**
 * Lower-cased Noto family name → Google Fonts CSS URL for every script face the
 * renderers may reference. Merged into each package's `*_GOOGLE_FONTS` map so a
 * single source of truth defines the URLs. `loadFamily` is omitted because
 * Google Fonts serves these under the same family name we request.
 */
export const SCRIPT_GOOGLE_FONTS: Record<string, { url: string; loadFamily?: string }> = {
  // CJK — sans + serif per language.
  'noto sans kr': { url: w('Noto+Sans+KR') },
  'noto sans sc': { url: w('Noto+Sans+SC') },
  'noto sans tc': { url: w('Noto+Sans+TC') },
  'noto sans jp': { url: w('Noto+Sans+JP') },
  'noto serif kr': { url: w('Noto+Serif+KR') },
  'noto serif sc': { url: w('Noto+Serif+SC') },
  'noto serif tc': { url: w('Noto+Serif+TC') },
  'noto serif jp': { url: w('Noto+Serif+JP') },
  // Latin/Cyrillic/Greek base (also the non-CJK sans/serif tail head).
  'noto sans': { url: w('Noto+Sans') },
  'noto serif': { url: w('Noto+Serif') },
  // Indic / SE-Asian.
  'noto sans devanagari': { url: w('Noto+Sans+Devanagari') },
  'noto sans thai': { url: w('Noto+Sans+Thai') },
  // Hebrew (RTL).
  'noto sans hebrew': { url: w('Noto+Sans+Hebrew') },
  'noto serif hebrew': { url: w('Noto+Serif+Hebrew') },
};

/**
 * Family names always queued for preload (self-referencing) when
 * `useGoogleFonts` is on, so every Noto face the renderer might append to a
 * font stack is actually loaded even when no document font maps to it by name.
 * Mirrors the existing Arabic self-referencing entries.
 */
export const SCRIPT_PRELOAD_NAMES: string[] = [
  'Noto Sans KR', 'Noto Sans SC', 'Noto Sans TC', 'Noto Sans JP',
  'Noto Serif KR', 'Noto Serif SC', 'Noto Serif TC', 'Noto Serif JP',
  'Noto Sans', 'Noto Serif',
  'Noto Sans Devanagari', 'Noto Sans Thai',
  'Noto Sans Hebrew', 'Noto Serif Hebrew',
];

/**
 * Decide WHICH of the script Noto families above actually need force-loading for
 * a document, by scanning its rendered text for the Unicode scripts that require
 * a script-specific web font.
 *
 * Why scan the text rather than always preload {@link SCRIPT_PRELOAD_NAMES}: the
 * CJK Noto families have no `text=` subset (the full multi-MB family is fetched),
 * so eagerly loading all eight CJK faces for a pure-Latin document blocks first
 * paint on megabytes of fonts that will never paint a glyph. We force-load only
 * the families whose script appears in the text. The renderer still APPENDS the
 * full Noto fallback chain to the canvas font stack — preloading is purely about
 * which faces are fetched up front before first paint; an un-preloaded face that
 * later proves needed simply loads lazily via the FontFaceSet on its first use.
 *
 * Detection is by objective Unicode block membership (no name/heuristic guessing
 * — see root CLAUDE.md). Ranges scanned per codepoint:
 *   - Hangul     U+1100–11FF (Jamo), U+3130–318F (Compat Jamo), U+AC00–D7AF
 *                (Syllables)                                   → CJK, lang 'kr'
 *   - Hira/Kana  U+3040–30FF                                   → CJK, lang 'jp'
 *   - Han        U+3400–4DBF, U+4E00–9FFF, U+F900–FAFF,
 *                U+20000–2FA1F (via codePointAt)               → CJK, shared
 *   - Arabic     U+0600–06FF, U+0750–077F, U+08A0–08FF,
 *                U+FB50–FDFF, U+FE70–FEFF      → 'Noto Naskh Arabic','Noto Sans Arabic'
 *   - Thai       U+0E00–0E7F                  → 'Noto Sans Thai'
 *   - Hebrew     U+0590–05FF, U+FB1D–FB4F      → 'Noto Sans Hebrew','Noto Serif Hebrew'
 *   - Devanagari U+0900–097F                   → 'Noto Sans Devanagari'
 *   - Cyrillic   U+0400–04FF / Greek U+0370–03FF → 'Noto Sans','Noto Serif'
 *                (the un-suffixed Latin Notos embed these)
 *   - Basic Latin / Latin-1 / everything else  → nothing
 *
 * CJK language resolution: a Hangul codepoint forces 'kr' and a Kana codepoint
 * forces 'jp' (those scripts are language-exclusive); shared Han codepoints adopt
 * the document's `cjkLang` hint (derived by the caller from the theme font via
 * {@link classifyCjkFont}), defaulting to 'jp' when no hint and no Hangul/Kana is
 * present. For each detected CJK language BOTH the Sans and Serif face are
 * emitted (the renderer may pick either depending on the run's serif-ness). When
 * multiple CJK languages appear (e.g. Hangul + Kana), each contributes its pair.
 *
 * Note the Arabic names are NOT keys of {@link SCRIPT_GOOGLE_FONTS} — they live
 * in each package's own `*_GOOGLE_FONTS` map (which also spreads SCRIPT_GOOGLE_FONTS),
 * mirroring the unconditional Arabic preload entries the callers already queue.
 *
 * The result is a deduplicated, deterministically-ordered array, so the
 * main-thread `load()` and the render worker — given the same parsed model —
 * derive an IDENTICAL preload set (required for worker/main pixel equivalence).
 *
 * Incremental collection is exposed internally through
 * {@link ScriptPreloadAccumulator}. Format-owned pull pipelines use it to
 * discard each parsed unit after contributing its script facts, while
 * {@link scriptPreloadNamesForText} preserves the one-shot contract.
 *
 * Kept off the package root API.
 */
export class ScriptPreloadAccumulator {
  private hasHan = false;
  private hasHangul = false;
  private hasKana = false;
  private hasArabic = false;
  private hasThai = false;
  private hasHebrew = false;
  private hasDevanagari = false;
  private hasCyrGreek = false;

  constructor(private readonly cjkLang: CjkLang | null) {}

  clone(): ScriptPreloadAccumulator {
    const copy = new ScriptPreloadAccumulator(this.cjkLang);
    copy.hasHan = this.hasHan;
    copy.hasHangul = this.hasHangul;
    copy.hasKana = this.hasKana;
    copy.hasArabic = this.hasArabic;
    copy.hasThai = this.hasThai;
    copy.hasHebrew = this.hasHebrew;
    copy.hasDevanagari = this.hasDevanagari;
    copy.hasCyrGreek = this.hasCyrGreek;
    return copy;
  }

  addText(text: Iterable<string>): void {
    const allFound = (): boolean =>
      (this.hasHan && this.hasHangul && this.hasKana) &&
      this.hasArabic &&
      this.hasThai &&
      this.hasHebrew &&
      this.hasDevanagari &&
      this.hasCyrGreek;

    outer: for (const chunk of text) {
      if (!chunk) continue;
      for (const ch of chunk) {
        const cp = ch.codePointAt(0);
        if (cp === undefined) continue;

        // Fast path: ASCII / Latin-1 / Latin Extended need no script font.
        if (cp <= 0x024f) continue;

        if (
          (cp >= 0x1100 && cp <= 0x11ff) ||
          (cp >= 0x3130 && cp <= 0x318f) ||
          (cp >= 0xac00 && cp <= 0xd7af)
        ) {
          this.hasHangul = true;
        } else if (cp >= 0x3040 && cp <= 0x30ff) {
          this.hasKana = true;
        } else if (
          (cp >= 0x3400 && cp <= 0x4dbf) ||
          (cp >= 0x4e00 && cp <= 0x9fff) ||
          (cp >= 0xf900 && cp <= 0xfaff) ||
          (cp >= 0x20000 && cp <= 0x2fa1f)
        ) {
          this.hasHan = true;
        } else if (
          (cp >= 0x0600 && cp <= 0x06ff) ||
          (cp >= 0x0750 && cp <= 0x077f) ||
          (cp >= 0x08a0 && cp <= 0x08ff) ||
          (cp >= 0xfb50 && cp <= 0xfdff) ||
          (cp >= 0xfe70 && cp <= 0xfeff)
        ) {
          this.hasArabic = true;
        } else if (cp >= 0x0e00 && cp <= 0x0e7f) {
          this.hasThai = true;
        } else if (
          (cp >= 0x0590 && cp <= 0x05ff) ||
          (cp >= 0xfb1d && cp <= 0xfb4f)
        ) {
          this.hasHebrew = true;
        } else if (cp >= 0x0900 && cp <= 0x097f) {
          this.hasDevanagari = true;
        } else if (
          (cp >= 0x0400 && cp <= 0x04ff) ||
          (cp >= 0x0370 && cp <= 0x03ff)
        ) {
          this.hasCyrGreek = true;
        }

        if (allFound()) break outer;
      }
    }
  }

  names(): string[] {
    const names: string[] = [];

    // CJK: each detected language contributes its Sans+Serif pair. Hangul → 'kr',
    // Kana → 'jp', shared Han → the language hint (or 'jp' default) UNLESS Hangul
    // or Kana already pinned a CJK language, in which case Han resolves to one of
    // those faces and needs no extra family.
    const cjkLangs = new Set<CjkLang>();
    if (this.hasHangul) cjkLangs.add('kr');
    if (this.hasKana) cjkLangs.add('jp');
    if (this.hasHan && cjkLangs.size === 0) {
      cjkLangs.add(this.cjkLang ?? 'jp');
    }
    // Stable order: kr, sc, tc, jp.
    for (const lang of ['kr', 'sc', 'tc', 'jp'] as const) {
      if (cjkLangs.has(lang)) {
        const suffix = { kr: 'KR', sc: 'SC', tc: 'TC', jp: 'JP' }[lang];
        names.push(`Noto Sans ${suffix}`, `Noto Serif ${suffix}`);
      }
    }

    if (this.hasCyrGreek) names.push('Noto Sans', 'Noto Serif');
    if (this.hasArabic) names.push('Noto Naskh Arabic', 'Noto Sans Arabic');
    if (this.hasThai) names.push('Noto Sans Thai');
    if (this.hasHebrew) names.push('Noto Sans Hebrew', 'Noto Serif Hebrew');
    if (this.hasDevanagari) names.push('Noto Sans Devanagari');

    return names;
  }
}

export function scriptPreloadNamesForText(
  text: Iterable<string>,
  cjkLang: CjkLang | null,
): string[] {
  const accumulator = new ScriptPreloadAccumulator(cjkLang);
  accumulator.addText(text);
  return accumulator.names();
}
