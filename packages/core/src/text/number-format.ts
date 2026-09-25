// ECMA-376 §17.18.59 ST_NumberFormat — render an ordinal integer (a list item's
// index, a page number, …) as the text a consumer displays for that value. This
// is the SHARED numbering-format kernel for every WordprocessingML surface that
// carries an ST_NumberFormat: page numbering (§17.6.12 `<w:pgNumType w:fmt>` and
// the §17.16.4.3.1 field general-formatting switches `\* roman` / `\* ALPHABETIC`
// / …), list markers (§17.9.17 `<w:numFmt>`), footnote/endnote numbering, etc.
// Keeping it in `core` means every consumer reuses ONE converter instead of
// forking a copy. (NOTE: the docx list-marker path also duplicates this table in
// Rust — `packages/docx/parser/src/numbering.rs::format_counter` — because the
// `%1.%2` multi-level composition happens at parse time; the two MUST stay in
// sync. See that file's header.)
//
// SCOPE: the numeric, algorithmically-defined formats §17.18.59 specifies with a
// concrete character set + construction steps. The families below cover the
// Latin/roman core, the CJK counting/legal systems, the RTL Hebrew/Arabic
// alphabets, and the positional digit substitutions (Thai/Hindi/full-width/…).
//
// A DOCUMENTED RESIDUAL degrades to `decimal` (except explicit `none`):
//   * The language SPELL-OUT formats — `cardinalText`, `ordinalText`, `ordinal`,
//     `thaiCounting`, `vietnameseCounting`, `hindiCounting`, `bahtText`,
//     `dollarText`, `custom` — §17.18.59 defines these only as "the textual
//     representation, in the language of the lang element, of …". There is NO
//     algorithm or character table to implement; they require a per-language
//     dictionary keyed off the run's `w:lang`, which this pure converter has no
//     access to. Implementing them would mean inventing spell-outs (a heuristic).
//   * The Chinese sexagenary-cycle ideographs (`ideographTraditional`,
//     `ideographZodiac`, `ideographZodiacTraditional`) and other rarely-authored
//     enclosed families we have not yet added.

/** ECMA-376 §17.18.59 ST_NumberFormat — the values this converter renders
 *  natively. Any other string is accepted at the call site and falls back to
 *  `decimal` (see module header). */
export type NumberFormat =
  // Latin / roman core (§17.16.4.3.1 field switches map onto these).
  | 'decimal'
  | 'upperRoman'
  | 'lowerRoman'
  | 'upperLetter'
  | 'lowerLetter'
  // Positional digit substitution (base-10, own zero glyph).
  | 'decimalFullWidth'
  | 'decimalHalfWidth'
  | 'decimalZero'
  | 'thaiNumbers'
  | 'hindiNumbers'
  | 'ideographDigital'
  | 'japaneseDigitalTenThousand'
  | 'koreanDigital'
  | 'koreanDigital2'
  | 'taiwaneseDigital'
  // Positional CJK with the 十 tens-prefix rule.
  | 'chineseCounting'
  | 'taiwaneseCounting'
  // Grouped "counting / legal" CJK (万/千/百/十 place words, 零 zero-fill).
  | 'chineseCountingThousand'
  | 'chineseLegalSimplified'
  | 'japaneseCounting'
  | 'japaneseLegal'
  | 'koreanCounting'
  | 'koreanLegal'
  | 'taiwaneseCountingThousand'
  | 'ideographLegalTraditional'
  // Repeat-letter alphabets (subtract set size, repeat the glyph).
  | 'arabicAlpha'
  | 'arabicAbjad'
  | 'russianLower'
  | 'russianUpper'
  | 'thaiLetters'
  | 'chosung'
  | 'ganada'
  | 'hindiVowels'
  | 'hindiConsonants'
  // Enclosed decimals + katakana a-i-u-e-o sequences (§17.18.59).
  | 'decimalEnclosedCircle'
  | 'aiueoFullWidth'
  | 'aiueo'
  // Hebrew (positional gematria / alphabet-with-ת-suffix — NOT the repeat scheme).
  | 'hebrew1'
  | 'hebrew2'
  // Other algorithmic systems.
  | 'hex'
  | 'numberInDash'
  | 'none'
  // Accepted but rendered as decimal (documented residual) — kept in the type so
  // callers can pass a raw parsed value without a cast for the common ones.
  | 'cardinalText'
  | 'ordinalText'
  | 'ordinal'
  | (string & {});

// ── Roman numerals ──────────────────────────────────────────────────────────
// §17.18.59 upperRoman/lowerRoman. Classic additive numerals, greedily consumed
// high→low. The four subtractive pairs (CM/CD/XC/XL/IX/IV) are inlined so
// 4/9/40/… render correctly. Values ≥ 4000 have no classical single glyph; Word
// writes repeated M (no vinculum bar), which the greedy 1000→M step reproduces
// (4000 → "MMMM").
const ROMAN_TABLE: ReadonlyArray<readonly [number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

// Resource policy, not an OOXML numeric limit: one formatted ordinal may emit
// at most 4096 UTF-16 units. Check before repeat/allocation, not after building
// the string. Huge valid section starts can otherwise expand to millions of
// Roman/alphabetic glyphs. Throw rather than silently changing the numbering.
const MAX_NUMBER_FORMAT_UNITS = 4096;
function checkNumberFormatSize(units: number): void {
  if (units > MAX_NUMBER_FORMAT_UNITS) {
    throw new RangeError('Number-format output budget exceeded');
  }
}

/** Uppercase roman numerals for a positive integer. Caller guarantees n ≥ 1. */
function toUpperRoman(n: number): string {
  let out = '';
  let rem = n;
  for (const [value, glyph] of ROMAN_TABLE) {
    const count = Math.floor(rem / value);
    checkNumberFormatSize(out.length + count * glyph.length);
    out += glyph.repeat(count);
    rem %= value;
  }
  return out;
}

// ── Repeat-letter alphabets ─────────────────────────────────────────────────
// §17.18.59 upperLetter/lowerLetter/arabicAlpha/arabicAbjad/russian*/
// thaiLetters/chosung/ganada/hindiVowels/hindiConsonants all share ONE scheme:
// a fixed ordered character set of size N; the value maps into 1..N, and for
// values > N the SAME character is REPEATED once per full N subtracted (§17.18.59
// "written once and then repeated for each time the size of the set was
// subtracted"). This is Word's letter-column-UNLIKE scheme — NOT base-N. For
// English (N=26): 27 → "aa", 53 → "aaa", 54 → "BBB". Caller guarantees n ≥ 1.
// (hebrew2 is NOT in this family — it appends a ת run instead; see toHebrew2.)
function repeatAlphabet(n: number, glyphs: readonly string[]): string {
  const size = glyphs.length;
  const repeats = Math.floor((n - 1) / size) + 1;
  const glyph = glyphs[(n - 1) % size];
  checkNumberFormatSize(repeats * glyph.length);
  return glyph.repeat(repeats);
}

/** A-Z, built once for the letter converters. */
const LATIN_UPPER: readonly string[] = Array.from({ length: 26 }, (_, i) =>
  String.fromCharCode(0x41 + i),
);

// §17.18.59 arabicAlpha "Arabic Alphabet" — positions 1–28 (spec character list).
const ARABIC_ALPHA: readonly string[] = [
  'أ', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د',
  'ذ', 'ر', 'ز', 'س', 'ش', 'ص', 'ض', 'ط',
  'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م',
  'ن', 'ه', 'و', 'ي',
];

// §17.18.59 arabicAbjad "Arabic Abjad Numerals" — positions 1–28 (spec list).
const ARABIC_ABJAD: readonly string[] = [
  'أ', 'ب', 'ج', 'د', 'ه', 'و', 'ز', 'ح',
  'ط', 'ي', 'ك', 'ل', 'م', 'ن', 'س', 'ع',
  'ف', 'ص', 'ق', 'ر', 'ش', 'ت', 'ث', 'خ',
  'ذ', 'ض', 'غ', 'ظ',
];

// §17.18.59 hebrew2 "Hebrew Alphabet" — positions 1–22 (spec list, note the
// non-contiguous ranges: aleph..yod, then kaf/lamed/mem, nun..ayin, pe, tsadi..tav).
const HEBREW_ALPHABET: readonly string[] = [
  'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח',
  'ט', 'י', // 1–10 aleph..yod
  'כ', 'ל', 'מ', // kaf, lamed, mem
  'נ', 'ס', 'ע', // nun, samekh, ayin
  'פ', // pe
  'צ', 'ק', 'ר', 'ש', 'ת', // tsadi, qof, resh, shin, tav
];

// §17.18.59 russianLower/russianUpper — positions 1–29 (the Russian alphabet
// MINUS ё, й, ъ, ь per the spec's explicit range list). Lower: а–и, к–п, р–щ, ы,
// э, ю, я.
const RUSSIAN_LOWER: readonly string[] = [
  ...rangeChars(0x0430, 0x0438), // а–и
  ...rangeChars(0x043A, 0x043F), // к–п
  ...rangeChars(0x0440, 0x0449), // р–щ
  'ы', // ы
  'э', // э
  'ю', // ю
  'я', // я
];
const RUSSIAN_UPPER: readonly string[] = [
  ...rangeChars(0x0410, 0x0418), // А–И
  ...rangeChars(0x041A, 0x041F), // К–П
  ...rangeChars(0x0420, 0x0429), // Р–Щ
  'Ы', // Ы
  'Э', // Э
  'Ю', // Ю
  'Я', // Я
];

// §17.18.59 thaiLetters — positions 1–41 (spec: U+0E01, U+0E02, U+0E04,
// U+0E07–U+0E23, U+0E25, U+0E27–U+0E2E — the Thai consonants minus a few).
const THAI_LETTERS: readonly string[] = [
  'ก', 'ข', 'ค',
  ...rangeChars(0x0E07, 0x0E23),
  'ล',
  ...rangeChars(0x0E27, 0x0E2E),
];

// §17.18.59 chosung "Korean Chosung" — positions 1–14 (spec list of jamo).
const KOREAN_CHOSUNG: readonly string[] = [
  'ㄱ', 'ㄴ', 'ㄷ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅅ', 'ㅇ',
  'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ',
];

// §17.18.59 ganada "Korean Ganada" — positions 1–14 (spec list of syllables).
const KOREAN_GANADA: readonly string[] = [
  '가', '나', '다', '라', '마', '바', '사', '아',
  '자', '차', '카', '타', '파', '하',
];

// §17.18.59 hindiVowels — positions 1–37 = U+0915–U+0939 (the Devanagari
// consonants; the spec label "vowels" is a misnomer for this range).
const HINDI_VOWELS: readonly string[] = rangeChars(0x0915, 0x0939);

// §17.18.59 hindiConsonants — positions 1–18 = U+0905–U+0914, then U+0905+U+0902
// and U+0905+U+0903 (the Devanagari independent vowels + two anusvara/visarga
// combinations; the spec label "consonants" is likewise a misnomer).
const HINDI_CONSONANTS: readonly string[] = [
  ...rangeChars(0x0905, 0x0914),
  'अं',
  'अः',
];

/** Inclusive code-point range → array of single-character strings. */
function rangeChars(fromCp: number, toCp: number): string[] {
  const out: string[] = [];
  for (let cp = fromCp; cp <= toCp; cp++) out.push(String.fromCodePoint(cp));
  return out;
}

// §17.18.59 aiueoFullWidth "AIUEO Order Full-Width Katakana" — the full-width
// katakana in the traditional a-i-u-e-o order, using the SAME repeat scheme as
// the letter alphabets above (value maps into 1..N, repeated once per full N
// subtracted). §17.18.59 ENUMERATES these 48 code points explicitly — note it
// includes the archaic ヰ (U+30F0) and ヱ (U+30F1), so wo (ヲ) / n (ン) sit at
// positions 47/48. The section's prose says "positions 1–46", but that count is
// a copy/paste artifact from the half-width `aiueo` entry (which legitimately has
// 46 because half-width forms of ヰ/ヱ do not exist); we follow the explicit
// enumerated character list, which is the concrete normative data a consumer maps
// "respectively". Only values ≥ 45 are affected by the choice. Katakana (not
// hiragana) is also what Word emits: [MS-OE376] §2.1.580 note (b) records that
// where the (1st-edition Part 4) standard said "hiragana characters", "Word uses
// these numbering formats to specify sequences that will consist of Katakana
// characters" — matching the 5th-edition code-point list implemented here.
const KATAKANA_FULLWIDTH: readonly string[] = [
  '\u{30A2}', '\u{30A4}', '\u{30A6}', '\u{30A8}', '\u{30AA}', '\u{30AB}',
  '\u{30AD}', '\u{30AF}', '\u{30B1}', '\u{30B3}', '\u{30B5}', '\u{30B7}',
  '\u{30B9}', '\u{30BB}', '\u{30BD}', '\u{30BF}', '\u{30C1}', '\u{30C4}',
  '\u{30C6}', '\u{30C8}', '\u{30CA}', '\u{30CB}', '\u{30CC}', '\u{30CD}',
  '\u{30CE}', '\u{30CF}', '\u{30D2}', '\u{30D5}', '\u{30D8}', '\u{30DB}',
  '\u{30DE}', '\u{30DF}', '\u{30E0}', '\u{30E1}', '\u{30E2}', '\u{30E4}',
  '\u{30E6}', '\u{30E8}', '\u{30E9}', '\u{30EA}', '\u{30EB}', '\u{30EC}',
  '\u{30ED}', '\u{30EF}', '\u{30F0}', '\u{30F1}', '\u{30F2}', '\u{30F3}',
];

// §17.18.59 aiueo "AIUEO Order Half-Width Katakana" — positions 1–46 =
// U+FF71–U+FF9C (ｱ..ﾜ), then U+FF66 (ｦ wo), then U+FF9D (ﾝ n). No archaic ヰ/ヱ
// (they have no half-width forms). Same repeat scheme; spec example ｱ, ｲ, …, ｦ,
// ﾝ, ｱｱ confirms wo/n at 45/46 and the wrap at 47. Katakana per [MS-OE376]
// §2.1.580 note (b) — see KATAKANA_FULLWIDTH above.
const KATAKANA_HALFWIDTH: readonly string[] = [
  ...rangeChars(0xff71, 0xff9c),
  '\u{FF66}',
  '\u{FF9D}',
];

// §17.18.59 decimalEnclosedCircle "Decimal Numbers Enclosed in a Circle": the
// spec tables 1–20 → U+2460–U+2473 (①..⑳) and states that "for values greater
// than the size of the set, the items fall back to the decimal format" (its
// example: …, ⑲, ⑳, 21, …). Unicode does carry follow-on enclosed-number blocks
// (㉑..㉟ U+3251–U+325F, ㊱..㊿ U+32B1–U+32BF), but neither §17.18.59 nor the
// Word implementation notes ([MS-OE376] §2.1.580, which records this section's
// sibling deviations in detail) documents Word continuing the circled sequence
// past 20 — so we stay with the specified decimal fallback at 21+ until primary
// evidence (real Word output or an implementation note) shows otherwise. Caller
// guarantees n ≥ 1.
function toEnclosedCircle(n: number): string {
  return n <= 20 ? String.fromCodePoint(0x2460 + (n - 1)) : String(n);
}

// ── Positional digit substitution ───────────────────────────────────────────
// §17.18.59 decimalFullWidth/thaiNumbers/hindiNumbers/ideographDigital/… : a
// base-10 positional system with a fixed 10-glyph digit set (index 0 is the zero
// glyph). The displayed text is the decimal string with each ASCII digit swapped
// for its glyph. Caller guarantees n ≥ 1 (n=0 also renders fine — "zeroGlyph").
function toPositionalDigits(n: number, digits: readonly string[]): string {
  return String(n)
    .split('')
    .map((d) => digits[d.charCodeAt(0) - 0x30])
    .join('');
}

// Digit sets for the positional formats, index 0 = zero glyph … index 9.
// Full-width Arabic ０–９ (§17.18.59 decimalFullWidth, U+FF10–U+FF19).
const DIGITS_FULLWIDTH: readonly string[] = rangeChars(0xff10, 0xff19);
// Thai numerals ๐–๙ (§17.18.59 thaiNumbers, U+0E50–U+0E59).
const DIGITS_THAI: readonly string[] = rangeChars(0x0e50, 0x0e59);
// Hindi (Devanagari) numerals ०–९ (§17.18.59 hindiNumbers, U+0966–U+096F).
const DIGITS_HINDI: readonly string[] = rangeChars(0x0966, 0x096f);
// Ideographic digits 〇一二…九 (§17.18.59 ideographDigital, zero = U+3007).
const DIGITS_IDEOGRAPH: readonly string[] = [
  '〇', '一', '二', '三', '四', '五', '六', '七',
  '八', '九',
];
// Korean hangul digits 영일이삼… (§17.18.59 koreanDigital, zero = 영 U+C601).
const DIGITS_KOREAN: readonly string[] = [
  '영', '일', '이', '삼', '사', '오', '육', '칠',
  '팔', '구',
];
// koreanDigital2: same hangul-value positions but zero = 零 (U+96F6) and the
// value glyphs are the CJK numerals 一二三… (§17.18.59 koreanDigital2).
const DIGITS_KOREAN2: readonly string[] = [
  '零', '一', '二', '三', '四', '五', '六', '七',
  '八', '九',
];
// taiwaneseDigital: CJK numerals with zero = ○ (U+25CB) (§17.18.59).
const DIGITS_TAIWANESE: readonly string[] = [
  '○', '一', '二', '三', '四', '五', '六', '七',
  '八', '九',
];

// ── chineseCounting / taiwaneseCounting (十-prefix positional) ───────────────
// §17.18.59 chineseCounting: NOT the myriad grouping system. It is base-10
// positional with ONE special rule for the tens place: "Divide the value by 10
// and write the symbol for the remainder. If the quotient is less than 10, write
// 十 to the left of the symbol." So the 十 tens-word appears ONLY while the
// running quotient is a single digit — i.e. for 2-digit values (10–99). For
// values ≥ 100 the top quotient is ≥ 10, so the rule never fires and the number
// renders as pure positional digit substitution (with 〇 for zero).
// Verified against §17.18.59 examples: 10 → 十, 11 → 十一, 20 → 二十, 99 → 九十九,
// 100 → 一〇〇, 101 → 一〇一.
function toChineseCounting(n: number, digits: readonly string[]): string {
  // digits[0] is the zero glyph (〇 for chineseCounting, ○ for taiwaneseCounting).
  if (n < 10) return digits[n];
  const TEN = '十'; // 十
  if (n < 100) {
    const tens = Math.floor(n / 10);
    const ones = n % 10;
    // Leading-one elision: 10–19 → 十… (no 一 before 十). 20+ → <digit>十.
    const head = tens === 1 ? TEN : digits[tens] + TEN;
    return ones === 0 ? head : head + digits[ones];
  }
  // ≥ 100: pure positional digit substitution (the 十 rule no longer applies).
  return toPositionalDigits(n, digits);
}

/** ECMA-376 §17.18.59 ST_NumberFormat converter dispatch. Non-native formats —
 *  and zero / negative values under a format that has no glyph for them (roman,
 *  letters) — fall back to Arabic decimal. Explicit `none` suppresses output;
 *  `undefined`/absent `fmt` is the spec default `decimal`.
 *  Resource policy: expanding formats throw above 4096 UTF-16 output units. */
export function formatOrdinalNumber(n: number, fmt: NumberFormat | undefined): string {
  if (fmt === 'none') return '';
  // The format algorithms require an exact integer. Preserve the existing
  // decimal recovery convention without entering loops on Infinity or indexing
  // an alphabet with a fraction/unsafe integer.
  if (!Number.isSafeInteger(n)) return String(n);
  switch (fmt) {
    // Roman.
    case 'upperRoman':
      return n >= 1 ? toUpperRoman(n) : String(n);
    case 'lowerRoman':
      return n >= 1 ? toUpperRoman(n).toLowerCase() : String(n);
    // Latin letters.
    case 'upperLetter':
      return n >= 1 ? repeatAlphabet(n, LATIN_UPPER) : String(n);
    case 'lowerLetter':
      return n >= 1 ? repeatAlphabet(n, LATIN_UPPER).toLowerCase() : String(n);
    // Repeat-letter non-Latin alphabets.
    case 'arabicAlpha':
      return n >= 1 ? repeatAlphabet(n, ARABIC_ALPHA) : String(n);
    case 'arabicAbjad':
      return n >= 1 ? repeatAlphabet(n, ARABIC_ABJAD) : String(n);
    case 'russianLower':
      return n >= 1 ? repeatAlphabet(n, RUSSIAN_LOWER) : String(n);
    case 'russianUpper':
      return n >= 1 ? repeatAlphabet(n, RUSSIAN_UPPER) : String(n);
    case 'thaiLetters':
      return n >= 1 ? repeatAlphabet(n, THAI_LETTERS) : String(n);
    case 'chosung':
      return n >= 1 ? repeatAlphabet(n, KOREAN_CHOSUNG) : String(n);
    case 'ganada':
      return n >= 1 ? repeatAlphabet(n, KOREAN_GANADA) : String(n);
    case 'hindiVowels':
      return n >= 1 ? repeatAlphabet(n, HINDI_VOWELS) : String(n);
    case 'hindiConsonants':
      return n >= 1 ? repeatAlphabet(n, HINDI_CONSONANTS) : String(n);
    // Katakana a-i-u-e-o sequences (repeat scheme, like the letter alphabets).
    case 'aiueoFullWidth':
      return n >= 1 ? repeatAlphabet(n, KATAKANA_FULLWIDTH) : String(n);
    case 'aiueo':
      return n >= 1 ? repeatAlphabet(n, KATAKANA_HALFWIDTH) : String(n);
    // Enclosed decimals (bounded set → decimal fallback past the range).
    case 'decimalEnclosedCircle':
      return n >= 1 ? toEnclosedCircle(n) : String(n);
    // Hebrew.
    case 'hebrew1':
      return n >= 1 ? toHebrewGematria(n) : String(n);
    case 'hebrew2':
      return n >= 1 ? toHebrew2(n) : String(n);
    // Other algorithmic systems.
    case 'hex':
      // §17.18.59 hex: base-16 positional over 0–9, A–F (U+0030–U+0039,
      // U+0041–U+0046 — uppercase). Example: …, E, F, 10, 11, …, 1E, 1F, 20.
      return n >= 1 ? n.toString(16).toUpperCase() : String(n);
    case 'numberInDash':
      // §17.18.59 numberInDash: the Arabic decimal "placed between two dashes";
      // the example pattern spaces them: - 1 -, - 2 -, …, - 10 -. §17.16.4.3.1
      // ArabicDash concurs: 'prefix of "- " and a suffix of " -"' (123 → - 123 -).
      return n >= 1 ? `- ${n} -` : String(n);
    case 'decimalZero':
      // §17.18.59 decimalZero: Arabic decimal "with a zero added to numbers one
      // through nine" — 01, 02, …, 09, 10, 11, …, 99, 100 (only 1–9 are padded).
      return n >= 1 && n <= 9 ? `0${n}` : String(n);
    // Positional digit substitution (n=0 renders the zero glyph — but list/page
    // numbering is 1-based, so we still gate ≥1 and fall back for ≤0).
    case 'decimalFullWidth':
      return n >= 1 ? toPositionalDigits(n, DIGITS_FULLWIDTH) : String(n);
    case 'decimalHalfWidth':
      return String(n); // half-width Arabic == ASCII decimal.
    case 'thaiNumbers':
      return n >= 1 ? toPositionalDigits(n, DIGITS_THAI) : String(n);
    case 'hindiNumbers':
      return n >= 1 ? toPositionalDigits(n, DIGITS_HINDI) : String(n);
    case 'ideographDigital':
    case 'japaneseDigitalTenThousand': // identical digit-substitution system.
      return n >= 1 ? toPositionalDigits(n, DIGITS_IDEOGRAPH) : String(n);
    case 'koreanDigital':
      return n >= 1 ? toPositionalDigits(n, DIGITS_KOREAN) : String(n);
    case 'koreanDigital2':
      return n >= 1 ? toPositionalDigits(n, DIGITS_KOREAN2) : String(n);
    case 'taiwaneseDigital':
      return n >= 1 ? toPositionalDigits(n, DIGITS_TAIWANESE) : String(n);
    // Positional CJK with the 十 tens-prefix.
    case 'chineseCounting':
      return n >= 1 ? toChineseCounting(n, DIGITS_IDEOGRAPH) : String(n);
    case 'taiwaneseCounting':
      return n >= 1 ? toChineseCounting(n, DIGITS_TAIWANESE) : String(n);
    // Grouped counting / legal CJK.
    case 'chineseCountingThousand':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_CHINESE) : String(n);
    case 'taiwaneseCountingThousand':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_TAIWANESE) : String(n);
    case 'chineseLegalSimplified':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_CHINESE_LEGAL) : String(n);
    case 'ideographLegalTraditional':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_TRAD_LEGAL) : String(n);
    case 'japaneseCounting':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_JAPANESE) : String(n);
    case 'japaneseLegal':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_JAPANESE_LEGAL) : String(n);
    case 'koreanCounting':
      return n >= 1 ? toMyriadGrouped(n, MYRIAD_KOREAN) : String(n);
    case 'koreanLegal':
      return n >= 1 ? toKoreanLegal(n) : String(n);
    case 'decimal':
    default:
      return String(n);
  }
}

// ── Hebrew gematria (hebrew1) ───────────────────────────────────────────────
// §17.18.59 hebrew1 "Hebrew Letters": positional letter arithmetic. Replace each
// decimal place (ones, tens, hundreds, thousands) with its Hebrew symbol, reading
// right-to-left in logical order (units first, so the string is
// hundreds+tens+ones when read left-to-right). Two special cases: 15 → טו and
// 16 → טז (to avoid spelling a divine name). Thousands reuse the ones glyphs.
const HEBREW_ONES: readonly string[] = [
  '', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז',
  'ח', 'ט', // 0..9: -, א..ט
];
const HEBREW_TENS: readonly string[] = [
  '', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע',
  'פ', 'צ', // 0..90 step 10: -, י,כ,ל,מ,נ,ס,ע,פ,צ
];
const HEBREW_HUNDREDS: readonly string[] = [
  '', 'ק', 'ר', 'ש', 'ת', 'ך', 'ם', 'ן',
  'ף', 'ץ', // 0..900 step 100: -, ק,ר,ש,ת,ך,ם,ן,ף,ץ
];

/** hebrew1 gematria for a positive integer (§17.18.59). Values are built
 *  hundreds→tens→ones (visual left-to-right). Above 999 the thousands digit is
 *  written with the ones glyphs, prepended. Caller guarantees n ≥ 1. */
function toHebrewGematria(n: number): string {
  let out = '';
  let rem = n;
  const thousands = Math.floor(rem / 1000);
  rem %= 1000;
  const hundreds = Math.floor(rem / 100);
  rem %= 100;
  // §17.18.59 step 1: thousands digit uses the ones symbols.
  if (thousands > 0) out += HEBREW_ONES[thousands % 10];
  // §17.18.59 step 2: hundreds. Values >400 combine (e.g. 500 = ת + ק), but the
  // spec's explicit hundreds table covers 100–900 directly, so use it.
  out += HEBREW_HUNDREDS[hundreds];
  // §17.18.59 step 3: 15/16 special case (avoid spelling יה / יו).
  if (rem === 15) return out + 'טו'; // טו
  if (rem === 16) return out + 'טז'; // טז
  const tens = Math.floor(rem / 10);
  const ones = rem % 10;
  out += HEBREW_TENS[tens];
  out += HEBREW_ONES[ones];
  return out;
}

// ── Hebrew alphabet with ת suffix (hebrew2) ─────────────────────────────────
// §17.18.59 hebrew2: NOT the repeat-letter scheme. For n > 22: "1. Repeatedly
// subtract the size of the set (22) from the value until the result is equal to
// or less than the size of the set. 2. Write the symbol represented by the
// result value. 3. Then the ת symbol is repeated … for each time the size of the
// set was subtracted." So the RESULT glyph appears ONCE, followed by one ת per
// subtraction — 23 → את, 24 → בת; and per the §17.16.4.3.1 field example,
// 123 \* HEBREW2 → מ + 5×ת (5 subtractions leave 13 → מ).
// Caller guarantees n ≥ 1.
function toHebrew2(n: number): string {
  const size = HEBREW_ALPHABET.length; // 22
  // Subtractions until the remainder is in 1..22 (a remainder of exactly 22
  // stops — "equal to or less than the size of the set", so 44 → ת + 1×ת).
  const subtractions = Math.floor((n - 1) / size);
  const remainder = n - size * subtractions; // 1..22
  checkNumberFormatSize(1 + subtractions);
  return HEBREW_ALPHABET[remainder - 1] + 'ת'.repeat(subtractions);
}

// ── Grouped CJK counting / legal (myriad grouping) ──────────────────────────
// §17.18.59 chineseCountingThousand / japaneseCounting / *Legal / koreanCounting:
// the East-Asian myriad (万-grouped) system. A value < 10^8 is split into two
// 4-digit myriad groups (ones-myriad and 万-myriad); each group is rendered
// digit-by-place using the 千/百/十 unit words, with the 万 unit appended to the
// upper group. A run of interior zeros collapses to a single 零 (§17.18.59 "write
// the symbol 零 instead"). The formats differ ONLY in their glyph tables and in
// whether a leading "1" is written before 十/百/千.
interface MyriadTable {
  /** Value glyphs for digits 0–9 (index 0 is the zero-fill glyph, e.g. 零/〇). */
  readonly digits: readonly string[];
  /** Unit glyphs for 10, 100, 1000. */
  readonly ten: string;
  readonly hundred: string;
  readonly thousand: string;
  /** Unit glyph for the 10^4 myriad group. */
  readonly myriad: string;
  /** Whether "1" is elided before a leading 十/百/千 (Japanese: 十, not 一十;
   *  Chinese thousand-counting: writes 一十). */
  readonly elideOne: boolean;
  /** Whether a single 零 fills an interior zero gap (§17.18.59: the Chinese
   *  counting/legal formats specify this; the Japanese/Korean/traditional-legal
   *  step lists do NOT mention 零, so they omit it — e.g. japaneseCounting
   *  10005 → 一万五, chineseCountingThousand 10005 → 一万零五). */
  readonly insertZero: boolean;
}

// Normal simplified/generic CJK digits and units.
const CJK_DIGITS: readonly string[] = DIGITS_KOREAN2; // 零 一 二 … 九
const MYRIAD_JAPANESE: MyriadTable = {
  digits: CJK_DIGITS,
  ten: '十', // 十
  hundred: '百', // 百
  thousand: '千', // 千
  myriad: '万', // 万
  elideOne: true, // §17.18.59 example: 10 → 十, 20 → 二十.
  insertZero: false, // japaneseCounting step list has no 零 rule.
};
const MYRIAD_CHINESE: MyriadTable = {
  ...MYRIAD_JAPANESE,
  elideOne: false, // §17.18.59 example: 10 → 一十.
  insertZero: true, // chineseCountingThousand specifies the 零 gap rule.
};
const MYRIAD_TAIWANESE: MyriadTable = {
  // taiwaneseCountingThousand shares the simplified units but its example shows
  // 10 → 一十 (no elision) like chineseCountingThousand, and specifies 零.
  ...MYRIAD_CHINESE,
};
const MYRIAD_KOREAN: MyriadTable = {
  digits: ['영', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'],
  ten: '십', // 십
  hundred: '백', // 백
  thousand: '천', // 천
  myriad: '만', // 만
  elideOne: true, // §17.18.59 example: 10 → 십, 십일.
  insertZero: false, // koreanCounting step list has no 零 rule.
};
const MYRIAD_CHINESE_LEGAL: MyriadTable = {
  // §17.18.59 chineseLegalSimplified digits 0–9: 零 壹 贰 叁 肆 伍 陆 柒 捌 玖.
  digits: ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'],
  ten: '拾', // 拾
  hundred: '佰', // 佰
  thousand: '仟', // 仟
  myriad: '万', // 万
  elideOne: false,
  insertZero: true, // chineseLegalSimplified specifies the 零 gap rule.
};
const MYRIAD_JAPANESE_LEGAL: MyriadTable = {
  // §17.18.59 japaneseLegal digits: 壱 弐 参 四 伍 六 七 八 九 (zero-fill reuses 零).
  digits: ['零', '壱', '弐', '参', '四', '伍', '六', '七', '八', '九'],
  ten: '拾', // 拾
  hundred: '百', // 百
  thousand: '阡', // 阡
  myriad: '萬', // 萬
  elideOne: false, // §17.18.59 example: 壱拾 (10).
  insertZero: false, // japaneseLegal step list has no 零 rule.
};
const MYRIAD_TRAD_LEGAL: MyriadTable = {
  // §17.18.59 ideographLegalTraditional digits: 壹 貳 參 肆 伍 陸 柒 捌 玖.
  digits: ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'],
  ten: '拾', // 拾
  hundred: '佰', // 佰
  thousand: '仟', // 仟
  myriad: '萬', // 萬
  elideOne: false, // §17.18.59 example: 壹拾 (10).
  insertZero: false, // ideographLegalTraditional step list has no 零 rule.
};

/** Render one 4-digit myriad group (0–9999) with 千/百/十 units. Returns '' for
 *  a zero group. A single interior 零 marks a gap between nonzero places
 *  (§17.18.59). `elideOne` drops a leading "1" before 十/百/千 when set. */
function renderMyriadGroup(group: number, t: MyriadTable, elideOne: boolean): string {
  const thousands = Math.floor(group / 1000) % 10;
  const hundreds = Math.floor(group / 100) % 10;
  const tens = Math.floor(group / 10) % 10;
  const ones = group % 10;
  const places: Array<{ digit: number; unit: string }> = [
    { digit: thousands, unit: t.thousand },
    { digit: hundreds, unit: t.hundred },
    { digit: tens, unit: t.ten },
    { digit: ones, unit: '' },
  ];
  let out = '';
  let sawNonZero = false;
  let pendingZero = false;
  for (const { digit, unit } of places) {
    if (digit === 0) {
      // Mark a zero gap only once we have already emitted a higher nonzero place
      // (leading zeros contribute nothing).
      if (sawNonZero) pendingZero = true;
      continue;
    }
    if (pendingZero) {
      // Single 零 for the collapsed interior zero run — only for formats that
      // specify it (§17.18.59: Chinese counting/legal; Japanese/Korean omit).
      if (t.insertZero) out += t.digits[0];
      pendingZero = false;
    }
    // Elide the multiplier "1" before 十/百/千 when the table asks for it
    // (Japanese/Korean: 十 not 一十, 百 not 一百, and mid-number too — 111 → 百十一,
    // 1100 → 千百). The ones place has no unit, so its "1" is always written; the
    // 万/億 myriad units sit OUTSIDE this group render, so 10000 keeps its 一 (一万).
    // Chinese-thousand / legal formats set elide=false and write 一十/一百/一千.
    if (elideOne && digit === 1 && unit) {
      out += unit;
    } else {
      out += t.digits[digit] + unit;
    }
    sawNonZero = true;
  }
  return out;
}

/** East-Asian myriad-grouped counting/legal formatter (§17.18.59). Handles values
 *  up to 10^8 − 1 precisely (two myriad groups). For ≥ 10^8 we render the 億
 *  group recursively with the same rules — the spec sketches this as "an
 *  additional symbol at one hundred million" but does not fully table it, so the
 *  高位 rendering is a faithful extension of the documented < 10^8 algorithm.
 *  Caller guarantees n ≥ 1. */
function toMyriadGrouped(n: number, t: MyriadTable): string {
  if (n >= 100000000) {
    // 億 = 10^8. Recurse: <upper>億<remainder>, remainder rendered with no leading
    // elision quirk changes. (U+5104 億 for JP/CN, U+5104 shared.)
    const upper = Math.floor(n / 100000000);
    const lower = n % 100000000;
    const OKU = '億'; // 億
    const head = toMyriadGrouped(upper, t) + OKU;
    if (lower === 0) return head;
    // A gap between the 億 group and a small remainder gets a 零 (e.g. 100000001)
    // for the formats that specify the 零 rule.
    const gap = t.insertZero && lower < 10000000 ? t.digits[0] : '';
    return head + gap + toMyriadGrouped(lower, t);
  }
  const upperGroup = Math.floor(n / 10000); // 万 group (0–9999).
  const lowerGroup = n % 10000; // ones group (0–9999).
  let out = '';
  if (upperGroup > 0) {
    out += renderMyriadGroup(upperGroup, t, t.elideOne) + t.myriad;
  }
  if (lowerGroup > 0) {
    // Insert a single 零 when the ones-group's highest place is below thousands
    // but the 万 group is present (e.g. 10005 → 一万零五): a gap exists between 万
    // and the ones group (§17.18.59 "if no groups are formed … write 零"). Only
    // for formats that specify the 零 rule.
    if (t.insertZero && upperGroup > 0 && lowerGroup < 1000) out += t.digits[0];
    out += renderMyriadGroup(lowerGroup, t, t.elideOne);
  }
  return out;
}

// ── koreanLegal (native-Korean numerals) ────────────────────────────────────
// §17.18.59 koreanLegal: NOT the myriad system — native Korean number words
// concatenated tens-word + ones-word. The spec tables ones 1–9 and tens 10..90
// explicitly (하나/둘/…, 열/스물/…), and its example shows the concatenation
// (11 → 열하나, 21 → 스물하나). The spec does NOT define values ≥ 100, so we
// render the fully-specified 1–99 range and fall back to decimal above it
// (rendering only what §17.18.59 tables — no invented composition).
const KOREAN_LEGAL_ONES: readonly string[] = [
  '', '하나', '둘', '셋', '넷', '다섯', '여섯', '일곱', '여덟', '아홉',
];
const KOREAN_LEGAL_TENS: readonly string[] = [
  '', '열', '스물', '서른', '마흔', '쉰', '예순', '일흔', '여든', '아흔',
];

/** koreanLegal for 1–99 (§17.18.59); ≥ 100 is undefined by the spec, so the
 *  caller's decimal fallback applies. Caller guarantees n ≥ 1. */
function toKoreanLegal(n: number): string {
  if (n >= 100) return String(n); // spec tables only 1–99.
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  return KOREAN_LEGAL_TENS[tens] + KOREAN_LEGAL_ONES[ones];
}
