// ── ST_NumberFormat rendering (ECMA-376 §17.18.59) ──────────────────────────
// This is the RUST twin of `packages/core/src/text/number-format.ts`
// (`formatOrdinalNumber`). List markers resolve to a final string at PARSE time
// (`resolve_text` composes `%1.%2` here in Rust), so the same numbering systems
// must be implemented on both sides and produce BYTE-IDENTICAL output. When you
// touch a format here, mirror it there (and vice versa); the TS unit tests are
// the reference values. `bullet` is a list-only concern (no §17.18.59 numeric
// meaning) and stays Rust-only.
/// Bound expansions before allocating them. Non-repeating formats below do
/// bounded work on at most ten decimal digits; their small result is checked
/// after formatting. This is a caller-owned byte budget, not a numeric cutoff.
pub(super) fn format_counter_bounded(n: u32, format: &str, limit: usize) -> Result<String, ()> {
    if n != 0 {
        let alphabet = match format {
            "arabicAlpha" => Some(ARABIC_ALPHA),
            "arabicAbjad" => Some(ARABIC_ABJAD),
            "russianLower" => Some(RUSSIAN_LOWER),
            "russianUpper" => Some(RUSSIAN_UPPER),
            "thaiLetters" => Some(THAI_LETTERS),
            "chosung" => Some(KOREAN_CHOSUNG),
            "ganada" => Some(KOREAN_GANADA),
            "hindiVowels" => Some(HINDI_VOWELS),
            "hindiConsonants" => Some(HINDI_CONSONANTS),
            "aiueoFullWidth" => Some(KATAKANA_FULLWIDTH),
            "aiueo" => Some(KATAKANA_HALFWIDTH),
            _ => None,
        };
        let expanded_bytes = if let Some(glyphs) = alphabet {
            let size = glyphs.len() as u64;
            let index = (u64::from(n) - 1) % size;
            let repeats = (u64::from(n) - 1) / size + 1;
            Some(repeats * glyphs[index as usize].len() as u64)
        } else {
            match format {
                "upperLetter" | "lowerLetter" => Some((u64::from(n) - 1) / 26 + 1),
                "hebrew2" => {
                    let repeats = (u64::from(n) - 1) / HEBREW_ALPHABET.len() as u64;
                    let index = (u64::from(n) - 1) % HEBREW_ALPHABET.len() as u64;
                    Some(HEBREW_ALPHABET[index as usize].len() as u64 + repeats * "ת".len() as u64)
                }
                "upperRoman" | "lowerRoman" => {
                    let mut remaining = n;
                    let mut bytes = 0u64;
                    for (value, glyph) in ROMAN_DIGITS {
                        bytes += u64::from(remaining / value) * glyph.len() as u64;
                        remaining %= value;
                    }
                    Some(bytes)
                }
                _ => None,
            }
        };
        if expanded_bytes.is_some_and(|bytes| bytes > limit as u64) {
            return Err(());
        }
    }
    let result = format_counter(n, format);
    if result.len() > limit {
        Err(())
    } else {
        Ok(result)
    }
}

pub fn format_counter(n: u32, format: &str) -> String {
    // ECMA-376 17.18.59: `none` suppresses the number, including start=0.
    // Mirror the shared TS field formatter instead of using decimal fallback.
    if format == "none" {
        return String::new();
    }
    if format == "bullet" {
        return "•".to_string();
    }
    // The numeric systems are 1-based; a level with start=0 (rare) or an
    // underflow falls back to the decimal string, matching the TS `n >= 1` gate.
    if n == 0 {
        return n.to_string();
    }
    match format {
        "decimal" | "decimalHalfWidth" => n.to_string(),
        // Roman.
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        // Latin + non-Latin repeat-letter alphabets (§17.18.59).
        "lowerLetter" => repeat_alphabet(n, &latin_upper()).to_lowercase(),
        "upperLetter" => repeat_alphabet(n, &latin_upper()),
        "arabicAlpha" => repeat_alphabet(n, ARABIC_ALPHA),
        "arabicAbjad" => repeat_alphabet(n, ARABIC_ABJAD),
        "russianLower" => repeat_alphabet(n, RUSSIAN_LOWER),
        "russianUpper" => repeat_alphabet(n, RUSSIAN_UPPER),
        "thaiLetters" => repeat_alphabet(n, THAI_LETTERS),
        "chosung" => repeat_alphabet(n, KOREAN_CHOSUNG),
        "ganada" => repeat_alphabet(n, KOREAN_GANADA),
        "hindiVowels" => repeat_alphabet(n, HINDI_VOWELS),
        "hindiConsonants" => repeat_alphabet(n, HINDI_CONSONANTS),
        // Katakana a-i-u-e-o sequences (repeat scheme, like the letter alphabets).
        "aiueoFullWidth" => repeat_alphabet(n, KATAKANA_FULLWIDTH),
        "aiueo" => repeat_alphabet(n, KATAKANA_HALFWIDTH),
        // Enclosed decimals (bounded set → decimal fallback past the range).
        "decimalEnclosedCircle" => to_enclosed_circle(n),
        // Hebrew: positional gematria / alphabet-with-ת-suffix (NOT repeat).
        "hebrew1" => to_hebrew_gematria(n),
        "hebrew2" => to_hebrew2(n),
        // Other algorithmic systems (§17.18.59 hex / numberInDash / decimalZero).
        "hex" => format!("{:X}", n),
        "numberInDash" => format!("- {} -", n),
        "decimalZero" => {
            if n <= 9 {
                format!("0{}", n)
            } else {
                n.to_string()
            }
        }
        // Positional digit substitution.
        "decimalFullWidth" => to_positional_digits(n, DIGITS_FULLWIDTH),
        "thaiNumbers" => to_positional_digits(n, DIGITS_THAI),
        "hindiNumbers" => to_positional_digits(n, DIGITS_HINDI),
        "ideographDigital" | "japaneseDigitalTenThousand" => {
            to_positional_digits(n, DIGITS_IDEOGRAPH)
        }
        "koreanDigital" => to_positional_digits(n, DIGITS_KOREAN),
        "koreanDigital2" => to_positional_digits(n, DIGITS_KOREAN2),
        "taiwaneseDigital" => to_positional_digits(n, DIGITS_TAIWANESE),
        // 十-prefix positional.
        "chineseCounting" => to_chinese_counting(n, DIGITS_IDEOGRAPH),
        "taiwaneseCounting" => to_chinese_counting(n, DIGITS_TAIWANESE),
        // Grouped counting / legal CJK.
        "japaneseCounting" => to_myriad_grouped(n, &MYRIAD_JAPANESE),
        "chineseCountingThousand" => to_myriad_grouped(n, &MYRIAD_CHINESE),
        "taiwaneseCountingThousand" => to_myriad_grouped(n, &MYRIAD_CHINESE),
        "chineseLegalSimplified" => to_myriad_grouped(n, &MYRIAD_CHINESE_LEGAL),
        "ideographLegalTraditional" => to_myriad_grouped(n, &MYRIAD_TRAD_LEGAL),
        "japaneseLegal" => to_myriad_grouped(n, &MYRIAD_JAPANESE_LEGAL),
        "koreanCounting" => to_myriad_grouped(n, &MYRIAD_KOREAN),
        "koreanLegal" => to_korean_legal(n),
        // Documented residual (language spell-outs / unimplemented) → decimal.
        _ => n.to_string(),
    }
}

// §17.18.59 koreanLegal — native-Korean tens-word + ones-word, tabled for 1–99
// (≥100 undefined by the spec → decimal fallback). Mirrors TS `toKoreanLegal`.
const KOREAN_LEGAL_ONES: &[&str] = &[
    "", "하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉",
];
const KOREAN_LEGAL_TENS: &[&str] = &[
    "", "열", "스물", "서른", "마흔", "쉰", "예순", "일흔", "여든", "아흔",
];

fn to_korean_legal(n: u32) -> String {
    if n >= 100 {
        return n.to_string();
    }
    let tens = n / 10;
    let ones = n % 10;
    format!(
        "{}{}",
        KOREAN_LEGAL_TENS[tens as usize], KOREAN_LEGAL_ONES[ones as usize]
    )
}

const ROMAN_DIGITS: [(u32, &str); 13] = [
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
];

fn to_roman(n: u32) -> String {
    let mut n = n;
    let mut s = String::new();
    for (v, r) in &ROMAN_DIGITS {
        while n >= *v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

// A–Z, built for the Latin letter converters (repeat scheme, not base-26).
fn latin_upper() -> Vec<&'static str> {
    const A: [&str; 26] = [
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R",
        "S", "T", "U", "V", "W", "X", "Y", "Z",
    ];
    A.to_vec()
}

// §17.18.59 arabicAlpha "Arabic Alphabet" — positions 1–28.
const ARABIC_ALPHA: &[&str] = &[
    "أ", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر", "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ",
    "ف", "ق", "ك", "ل", "م", "ن", "ه", "و", "ي",
];
// §17.18.59 arabicAbjad "Arabic Abjad Numerals" — positions 1–28.
const ARABIC_ABJAD: &[&str] = &[
    "أ", "ب", "ج", "د", "ه", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق",
    "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "غ", "ظ",
];
// §17.18.59 hebrew2 "Hebrew Alphabet" — positions 1–22.
const HEBREW_ALPHABET: &[&str] = &[
    "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ", "ק",
    "ר", "ש", "ת",
];
// §17.18.59 russianLower/Upper — positions 1–29 (alphabet minus ё, й, ъ, ь).
const RUSSIAN_LOWER: &[&str] = &[
    "а", "б", "в", "г", "д", "е", "ж", "з", "и", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у",
    "ф", "х", "ц", "ч", "ш", "щ", "ы", "э", "ю", "я",
];
const RUSSIAN_UPPER: &[&str] = &[
    "А", "Б", "В", "Г", "Д", "Е", "Ж", "З", "И", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У",
    "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ы", "Э", "Ю", "Я",
];
// §17.18.59 thaiLetters — positions 1–41 (U+0E01, U+0E02, U+0E04, U+0E07–U+0E23,
// U+0E25, U+0E27–U+0E2E).
const THAI_LETTERS: &[&str] = &[
    "ก", "ข", "ค", "ง", "จ", "ฉ", "ช", "ซ", "ฌ", "ญ", "ฎ", "ฏ", "ฐ", "ฑ", "ฒ", "ณ", "ด", "ต", "ถ",
    "ท", "ธ", "น", "บ", "ป", "ผ", "ฝ", "พ", "ฟ", "ภ", "ม", "ย", "ร", "ล", "ว", "ศ", "ษ", "ส", "ห",
    "ฬ", "อ", "ฮ",
];
// §17.18.59 chosung "Korean Chosung" — positions 1–14.
const KOREAN_CHOSUNG: &[&str] = &[
    "ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ",
];
// §17.18.59 ganada "Korean Ganada" — positions 1–14.
const KOREAN_GANADA: &[&str] = &[
    "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하",
];
// §17.18.59 hindiVowels — positions 1–37 = U+0915–U+0939 (contiguous).
const HINDI_VOWELS: &[&str] = &[
    "क", "ख", "ग", "घ", "ङ", "च", "छ", "ज", "झ", "ञ", "ट", "ठ", "ड", "ढ", "ण", "त", "थ", "द", "ध",
    "न", "ऩ", "प", "फ", "ब", "भ", "म", "य", "र", "ऱ", "ल", "ळ", "ऴ", "व", "श", "ष", "स", "ह",
];
// §17.18.59 hindiConsonants — positions 1–18 = U+0905–U+0914 then अं / अः.
const HINDI_CONSONANTS: &[&str] = &[
    "अ", "आ", "इ", "ई", "उ", "ऊ", "ऋ", "ऌ", "ऍ", "ऎ", "ए", "ऐ", "ऑ", "ऒ", "ओ", "औ", "अं", "अः",
];

/// §17.18.59 repeat-letter alphabets: map 1..N into the set; for n>N repeat the
/// SAME glyph once per full N subtracted (not base-N). Mirrors the TS
/// `repeatAlphabet`.
fn repeat_alphabet(n: u32, glyphs: &[&str]) -> String {
    let size = glyphs.len() as u32;
    let repeats = (n - 1) / size + 1;
    let glyph = glyphs[((n - 1) % size) as usize];
    glyph.repeat(repeats as usize)
}

// §17.18.59 aiueoFullWidth "AIUEO Order Full-Width Katakana" — the full-width
// katakana in a-i-u-e-o order, using the SAME repeat scheme as the letter
// alphabets. §17.18.59 ENUMERATES these 48 code points (incl. the archaic ヰ
// U+30F0 and ヱ U+30F1, so wo/n land at 47/48). The section's "positions 1–46"
// prose is a copy/paste artifact from the half-width `aiueo` entry; we follow the
// explicit enumerated character list. Katakana (not hiragana) is also what Word
// emits: [MS-OE376] §2.1.580 note (b) records that where the (1st-edition Part 4)
// standard said "hiragana characters", Word uses katakana — matching the
// 5th-edition code-point list implemented here. Mirrors TS `KATAKANA_FULLWIDTH`.
const KATAKANA_FULLWIDTH: &[&str] = &[
    "\u{30A2}", "\u{30A4}", "\u{30A6}", "\u{30A8}", "\u{30AA}", "\u{30AB}", "\u{30AD}", "\u{30AF}",
    "\u{30B1}", "\u{30B3}", "\u{30B5}", "\u{30B7}", "\u{30B9}", "\u{30BB}", "\u{30BD}", "\u{30BF}",
    "\u{30C1}", "\u{30C4}", "\u{30C6}", "\u{30C8}", "\u{30CA}", "\u{30CB}", "\u{30CC}", "\u{30CD}",
    "\u{30CE}", "\u{30CF}", "\u{30D2}", "\u{30D5}", "\u{30D8}", "\u{30DB}", "\u{30DE}", "\u{30DF}",
    "\u{30E0}", "\u{30E1}", "\u{30E2}", "\u{30E4}", "\u{30E6}", "\u{30E8}", "\u{30E9}", "\u{30EA}",
    "\u{30EB}", "\u{30EC}", "\u{30ED}", "\u{30EF}", "\u{30F0}", "\u{30F1}", "\u{30F2}", "\u{30F3}",
];

// §17.18.59 aiueo "AIUEO Order Half-Width Katakana" — positions 1–46 =
// U+FF71–U+FF9C (ｱ..ﾜ), then U+FF66 (ｦ), then U+FF9D (ﾝ). No archaic ヰ/ヱ (no
// half-width forms). Katakana per [MS-OE376] §2.1.580 note (b) — see
// `KATAKANA_FULLWIDTH` above. Mirrors TS `KATAKANA_HALFWIDTH`.
const KATAKANA_HALFWIDTH: &[&str] = &[
    "\u{FF71}", "\u{FF72}", "\u{FF73}", "\u{FF74}", "\u{FF75}", "\u{FF76}", "\u{FF77}", "\u{FF78}",
    "\u{FF79}", "\u{FF7A}", "\u{FF7B}", "\u{FF7C}", "\u{FF7D}", "\u{FF7E}", "\u{FF7F}", "\u{FF80}",
    "\u{FF81}", "\u{FF82}", "\u{FF83}", "\u{FF84}", "\u{FF85}", "\u{FF86}", "\u{FF87}", "\u{FF88}",
    "\u{FF89}", "\u{FF8A}", "\u{FF8B}", "\u{FF8C}", "\u{FF8D}", "\u{FF8E}", "\u{FF8F}", "\u{FF90}",
    "\u{FF91}", "\u{FF92}", "\u{FF93}", "\u{FF94}", "\u{FF95}", "\u{FF96}", "\u{FF97}", "\u{FF98}",
    "\u{FF99}", "\u{FF9A}", "\u{FF9B}", "\u{FF9C}", "\u{FF66}", "\u{FF9D}",
];

/// §17.18.59 decimalEnclosedCircle: the spec tables 1–20 → U+2460–U+2473 (①..⑳)
/// and states that "for values greater than the size of the set, the items fall
/// back to the decimal format" (its example: …, ⑲, ⑳, 21, …). Unicode does carry
/// follow-on enclosed-number blocks (㉑..㉟ U+3251–U+325F, ㊱..㊿ U+32B1–U+32BF),
/// but neither §17.18.59 nor the Word implementation notes ([MS-OE376] §2.1.580,
/// which records this section's sibling deviations in detail) documents Word
/// continuing the circled sequence past 20 — so we stay with the specified
/// decimal fallback at 21+ until primary evidence (real Word output or an
/// implementation note) shows otherwise. Mirrors TS `toEnclosedCircle`. Caller
/// guarantees n ≥ 1 (n = 0 is handled by the early return in `format_counter`).
fn to_enclosed_circle(n: u32) -> String {
    match n {
        1..=20 => char::from_u32(0x2460 + (n - 1))
            .map(String::from)
            .unwrap_or_else(|| n.to_string()),
        _ => n.to_string(), // 21+ : §17.18.59 decimal fallback.
    }
}

// Positional digit sets, index 0 = zero glyph … index 9 (§17.18.59).
const DIGITS_FULLWIDTH: &[&str] = &["０", "１", "２", "３", "４", "５", "６", "７", "８", "９"];
const DIGITS_THAI: &[&str] = &["๐", "๑", "๒", "๓", "๔", "๕", "๖", "๗", "๘", "๙"];
const DIGITS_HINDI: &[&str] = &["०", "१", "२", "३", "४", "५", "६", "७", "८", "९"];
const DIGITS_IDEOGRAPH: &[&str] = &["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const DIGITS_KOREAN: &[&str] = &["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
const DIGITS_KOREAN2: &[&str] = &["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const DIGITS_TAIWANESE: &[&str] = &["○", "一", "二", "三", "四", "五", "六", "七", "八", "九"];

/// §17.18.59 base-10 positional digit substitution. Mirrors TS `toPositionalDigits`.
fn to_positional_digits(n: u32, digits: &[&str]) -> String {
    n.to_string()
        .bytes()
        .map(|b| digits[(b - b'0') as usize])
        .collect()
}

/// §17.18.59 chineseCounting / taiwaneseCounting: base-10 positional with the 十
/// tens-word for 2-digit values only (10 → 十, 20 → 二十, 99 → 九十九), pure
/// positional at ≥ 100 (100 → 一〇〇). Mirrors TS `toChineseCounting`.
fn to_chinese_counting(n: u32, digits: &[&str]) -> String {
    if n < 10 {
        return digits[n as usize].to_string();
    }
    if n < 100 {
        let tens = n / 10;
        let ones = n % 10;
        let head = if tens == 1 {
            "十".to_string()
        } else {
            format!("{}十", digits[tens as usize])
        };
        return if ones == 0 {
            head
        } else {
            format!("{}{}", head, digits[ones as usize])
        };
    }
    to_positional_digits(n, digits)
}

// ── Grouped CJK counting / legal (myriad grouping) ──────────────────────────
// Mirrors the TS `MyriadTable` + `toMyriadGrouped`.
struct MyriadTable {
    digits: &'static [&'static str], // index 0 = 零/〇 zero-fill glyph … 9
    ten: &'static str,
    hundred: &'static str,
    thousand: &'static str,
    myriad: &'static str,
    elide_one: bool,   // Japanese/Korean elide the "1" before 十/百/千.
    insert_zero: bool, // Chinese counting/legal fill an interior gap with 零.
}

const CJK_DIGITS: &[&str] = DIGITS_KOREAN2; // 零 一 二 … 九
const MYRIAD_JAPANESE: MyriadTable = MyriadTable {
    digits: CJK_DIGITS,
    ten: "十",
    hundred: "百",
    thousand: "千",
    myriad: "万",
    elide_one: true,
    insert_zero: false,
};
const MYRIAD_CHINESE: MyriadTable = MyriadTable {
    elide_one: false,
    insert_zero: true,
    ..MYRIAD_JAPANESE
};
const MYRIAD_KOREAN: MyriadTable = MyriadTable {
    digits: &["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"],
    ten: "십",
    hundred: "백",
    thousand: "천",
    myriad: "만",
    elide_one: true,
    insert_zero: false,
};
const MYRIAD_CHINESE_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"],
    ten: "拾",
    hundred: "佰",
    thousand: "仟",
    myriad: "万",
    elide_one: false,
    insert_zero: true,
};
const MYRIAD_JAPANESE_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壱", "弐", "参", "四", "伍", "六", "七", "八", "九"],
    ten: "拾",
    hundred: "百",
    thousand: "阡",
    myriad: "萬",
    elide_one: false,
    insert_zero: false,
};
const MYRIAD_TRAD_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"],
    ten: "拾",
    hundred: "佰",
    thousand: "仟",
    myriad: "萬",
    elide_one: false,
    insert_zero: false,
};

/// Render one 4-digit myriad group (0–9999). Mirrors TS `renderMyriadGroup`.
fn render_myriad_group(group: u32, t: &MyriadTable) -> String {
    let thousands = group / 1000 % 10;
    let hundreds = group / 100 % 10;
    let tens = group / 10 % 10;
    let ones = group % 10;
    let places = [
        (thousands, t.thousand),
        (hundreds, t.hundred),
        (tens, t.ten),
        (ones, ""),
    ];
    let mut out = String::new();
    let mut saw_non_zero = false;
    let mut pending_zero = false;
    for (digit, unit) in places {
        if digit == 0 {
            if saw_non_zero {
                pending_zero = true;
            }
            continue;
        }
        if pending_zero {
            if t.insert_zero {
                out.push_str(t.digits[0]);
            }
            pending_zero = false;
        }
        if t.elide_one && digit == 1 && !unit.is_empty() {
            out.push_str(unit);
        } else {
            out.push_str(t.digits[digit as usize]);
            out.push_str(unit);
        }
        saw_non_zero = true;
    }
    out
}

/// East-Asian myriad-grouped counting/legal formatter. Mirrors TS
/// `toMyriadGrouped` (values ≥ 10^8 recurse through 億).
fn to_myriad_grouped(n: u32, t: &MyriadTable) -> String {
    if n >= 100_000_000 {
        let upper = n / 100_000_000;
        let lower = n % 100_000_000;
        let head = format!("{}億", to_myriad_grouped(upper, t));
        if lower == 0 {
            return head;
        }
        let gap = if t.insert_zero && lower < 10_000_000 {
            t.digits[0]
        } else {
            ""
        };
        return format!("{}{}{}", head, gap, to_myriad_grouped(lower, t));
    }
    let upper_group = n / 10000;
    let lower_group = n % 10000;
    let mut out = String::new();
    if upper_group > 0 {
        out.push_str(&render_myriad_group(upper_group, t));
        out.push_str(t.myriad);
    }
    if lower_group > 0 {
        if t.insert_zero && upper_group > 0 && lower_group < 1000 {
            out.push_str(t.digits[0]);
        }
        out.push_str(&render_myriad_group(lower_group, t));
    }
    out
}

// §17.18.59 hebrew1 gematria. Mirrors TS `toHebrewGematria`.
const HEBREW_ONES: &[&str] = &["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
const HEBREW_TENS: &[&str] = &["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
const HEBREW_HUNDREDS: &[&str] = &["", "ק", "ר", "ש", "ת", "ך", "ם", "ן", "ף", "ץ"];

fn to_hebrew_gematria(n: u32) -> String {
    let mut out = String::new();
    let mut rem = n;
    let thousands = rem / 1000;
    rem %= 1000;
    let hundreds = rem / 100;
    rem %= 100;
    if thousands > 0 {
        out.push_str(HEBREW_ONES[(thousands % 10) as usize]);
    }
    out.push_str(HEBREW_HUNDREDS[hundreds as usize]);
    if rem == 15 {
        out.push_str("טו");
        return out;
    }
    if rem == 16 {
        out.push_str("טז");
        return out;
    }
    let tens = rem / 10;
    let ones = rem % 10;
    out.push_str(HEBREW_TENS[tens as usize]);
    out.push_str(HEBREW_ONES[ones as usize]);
    out
}

/// §17.18.59 hebrew2 — NOT the repeat-letter scheme: subtract 22 until the
/// result is ≤ 22, write THAT glyph once, then append ת once per subtraction
/// (23 → את, 24 → בת; §17.16.4.3.1 field example 123 → מ + 5×ת). Mirrors TS
/// `toHebrew2`.
fn to_hebrew2(n: u32) -> String {
    let size = HEBREW_ALPHABET.len() as u32; // 22
    let subtractions = (n - 1) / size;
    let remainder = n - size * subtractions; // 1..=22
    format!(
        "{}{}",
        HEBREW_ALPHABET[(remainder - 1) as usize],
        "ת".repeat(subtractions as usize)
    )
}

#[cfg(test)]
mod budget_tests {
    use super::*;

    #[test]
    fn expanding_formats_reject_huge_values_before_allocating() {
        for format in [
            "upperLetter",
            "lowerLetter",
            "upperRoman",
            "lowerRoman",
            "arabicAlpha",
            "arabicAbjad",
            "russianLower",
            "russianUpper",
            "thaiLetters",
            "chosung",
            "ganada",
            "hindiVowels",
            "hindiConsonants",
            "aiueoFullWidth",
            "aiueo",
            "hebrew2",
        ] {
            assert!(
                format_counter_bounded(u32::MAX, format, 64).is_err(),
                "{format}"
            );
            for n in [0, 1, 22, 26, 27, 48, 49, 123, 1999] {
                let expected = format_counter(n, format);
                assert_eq!(
                    format_counter_bounded(n, format, expected.len()),
                    Ok(expected.clone())
                );
                if !expected.is_empty() {
                    assert!(format_counter_bounded(n, format, expected.len() - 1).is_err());
                }
            }
        }
    }

    #[test]
    fn budget_is_not_a_numeric_value_cutoff() {
        assert_eq!(
            format_counter_bounded(u32::MAX, "decimal", 10),
            Ok("4294967295".into())
        );
        assert_eq!(
            format_counter_bounded(u32::MAX, "none", 0),
            Ok(String::new())
        );
        assert!(format_counter_bounded(1, "bullet", 2).is_err());
        assert_eq!(format_counter_bounded(1, "bullet", 3), Ok("•".into()));
    }
}
