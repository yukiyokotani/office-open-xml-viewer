//! MS-OSHARED 2.2.1.3 MSONFC to ECMA-376 17.18.59 ST_NumberFormat.
//! Callers validate their context-specific exclusions (list levels/pages).

pub(super) fn name(value: u8) -> Result<&'static str, String> {
    const FORMATS: [&str; 60] = [
        "decimal",
        "upperRoman",
        "lowerRoman",
        "upperLetter",
        "lowerLetter",
        "ordinal",
        "cardinalText",
        "ordinalText",
        "hex",
        "chicago",
        "ideographDigital",
        "japaneseCounting",
        "aiueo",
        "iroha",
        "decimalFullWidth",
        "decimalHalfWidth",
        "japaneseLegal",
        "japaneseDigitalTenThousand",
        "decimalEnclosedCircle",
        "decimalFullWidth2",
        "aiueoFullWidth",
        "irohaFullWidth",
        "decimalZero",
        "bullet",
        "ganada",
        "chosung",
        "decimalEnclosedFullstop",
        "decimalEnclosedParen",
        "decimalEnclosedCircleChinese",
        "ideographEnclosedCircle",
        "ideographTraditional",
        "ideographZodiac",
        "ideographZodiacTraditional",
        "taiwaneseCounting",
        "ideographLegalTraditional",
        "taiwaneseCountingThousand",
        "taiwaneseDigital",
        "chineseCounting",
        "chineseLegalSimplified",
        "chineseCountingThousand",
        "decimal",
        "koreanDigital",
        "koreanCounting",
        "koreanLegal",
        "koreanDigital2",
        "hebrew1",
        "arabicAlpha",
        "hebrew2",
        "arabicAbjad",
        "hindiVowels",
        "hindiConsonants",
        "hindiNumbers",
        "hindiCounting",
        "thaiLetters",
        "thaiNumbers",
        "thaiCounting",
        "vietnameseCounting",
        "numberInDash",
        "russianLower",
        "russianUpper",
    ];
    if value == 0xff {
        return Ok("none");
    }
    FORMATS
        .get(usize::from(value))
        .copied()
        .ok_or_else(|| super::unsupported("invalid Word number format"))
}
