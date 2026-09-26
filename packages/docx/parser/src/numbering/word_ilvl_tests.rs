use super::*;

// Word 16.113 PDF controls measured every integer 0..255 with a nine-level
// definition. Each character is an independent observed class: a defined
// level 0..8, no marker (N), level-8 start without an advance (S), the special
// level-0 advance at 15 (O), or a level-8 zero marker while the digit's level
// advances. Keeping the full byte domain here makes each observed row a case.
pub(crate) const OBSERVED_BYTE_CLASSES: &str = concat!(
    "012345678NNNNSSO",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
    "012345678SSSSSSS",
);

#[test]
fn every_word_measured_byte_has_its_marker_and_counter_action() {
    assert_eq!(OBSERVED_BYTE_CLASSES.len(), 256);
    for (byte, class) in OBSERVED_BYTE_CLASSES.bytes().enumerate() {
        let use_ = word_level_use(byte as u8);
        let expected = match (byte, class) {
            (_, b'N') => (None, None, None),
            (15, b'O') => (Some(0), Some(0), None),
            (_, b'S') => (Some(8), None, Some(9)),
            (0..=8, digit) => {
                let level = u32::from(digit - b'0');
                (Some(level), Some(level), None)
            }
            (_, b'8') => (Some(8), Some(8), None),
            (_, digit @ b'0'..=b'7') => (Some(8), Some(u32::from(digit - b'0')), Some(0)),
            _ => panic!("bad golden class at {byte}"),
        };
        assert_eq!(
            (use_.marker_level, use_.counter_level, use_.fixed_counter),
            expected,
            "Word control for ilvl={byte}"
        );
    }
}

#[test]
fn word_decimal_lexical_values_wrap_without_overflow() {
    for (value, expected) in [
        ("", 0),
        ("0008", 8),
        ("+8", 8),
        ("256", 0),
        ("257", 1),
        ("511", 255),
        ("512", 0),
        ("65535", 255),
        ("65536", 0),
        ("2147483647", 255),
        ("2147483648", 0),
        ("4294967295", 255),
        ("4294967296", 0),
        ("4294967297", 1),
        ("-1", 255),
        ("-2", 254),
        ("-128", 128),
        ("-129", 127),
        ("-256", 0),
    ] {
        assert_eq!(parse_word_ilvl(value), Ok(expected), "{value}");
    }
    for value in [" 8 ", "abc"] {
        assert!(parse_word_ilvl(value).is_err(), "{value:?}");
    }
    assert_eq!(parse_word_ilvl("18446744073709551616"), Ok(0));
}

#[test]
fn maximal_start_override_does_not_overflow_the_counter() {
    let xml = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4294967295"/></w:lvlOverride></w:num></w:numbering>"#;
    let mut map = NumberingMap::parse(xml, &HashMap::new());
    assert_eq!(map.advance(1, 0), u32::MAX);
    assert_eq!(map.advance(1, 0), u32::MAX);
}

#[test]
fn every_supported_number_format_uses_the_word_measured_synthetic_zero() {
    // Each pair is an independent Word PDF with a level-8 `%9` and ilvl=16.
    for (format, zero) in [
        ("decimal", "0"),
        ("decimalHalfWidth", "0"),
        ("lowerRoman", ""),
        ("upperRoman", ""),
        ("lowerLetter", ""),
        ("upperLetter", ""),
        ("arabicAlpha", ""),
        ("arabicAbjad", ""),
        ("russianLower", ""),
        ("russianUpper", ""),
        ("thaiLetters", ""),
        ("chosung", "0"),
        ("ganada", "0"),
        ("hindiVowels", ""),
        ("hindiConsonants", ""),
        ("aiueoFullWidth", "0"),
        ("aiueo", "0"),
        ("decimalEnclosedCircle", "0"),
        ("hebrew1", ""),
        ("hebrew2", ""),
        ("hex", "0"),
        ("numberInDash", "- 0 -"),
        ("decimalZero", "00"),
        ("decimalFullWidth", "０"),
        ("thaiNumbers", "๐"),
        ("hindiNumbers", "०"),
        ("ideographDigital", "〇"),
        ("japaneseDigitalTenThousand", "〇"),
        ("koreanDigital", "영"),
        ("koreanDigital2", "零"),
        ("taiwaneseDigital", "○"),
        ("chineseCounting", "○"),
        ("taiwaneseCounting", "○"),
        ("japaneseCounting", "〇"),
        ("chineseCountingThousand", "〇"),
        ("taiwaneseCountingThousand", "零"),
        ("chineseLegalSimplified", "零"),
        ("ideographLegalTraditional", "零"),
        ("japaneseLegal", "〇"),
        ("koreanCounting", "영"),
        ("koreanLegal", "0"),
        ("none", ""),
        ("bullet", ""),
    ] {
        assert_eq!(format_word_synthetic_zero(format), zero, "{format}");
        let levels: String = (0..9)
            .map(|level| {
                let own_format = if level == 8 { format } else { "decimal" };
                format!(
                    r#"<w:lvl w:ilvl="{level}"><w:start w:val="{}"/><w:numFmt w:val="{own_format}"/><w:lvlText w:val="Z-%9."/></w:lvl>"#,
                    level + 1
                )
            })
            .collect();
        let xml = format!(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0">{levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        );
        let mut map = NumberingMap::parse(&xml, &HashMap::new());
        for level in [0, 1, 8, 8, 0] {
            map.advance(1, level);
        }
        assert_eq!(
            map.resolve_text_word_zero(1, 8),
            format!("Z-{zero}."),
            "{format}"
        );
    }
}

#[test]
fn reset_placeholders_use_zero_until_their_counter_is_live() {
    let formats = [
        "decimal",
        "upperLetter",
        "lowerLetter",
        "upperRoman",
        "lowerRoman",
        "decimalZero",
        "decimal",
        "upperLetter",
        "decimalZero",
    ];
    let levels: String = formats
        .iter()
        .enumerate()
        .map(|(i, format)| {
            format!(
                r#"<w:lvl w:ilvl="{i}"><w:start w:val="{}"/><w:numFmt w:val="{format}"/><w:lvlText w:val="P-%1/%2/%3/%4/%5/%6/%7/%8/%9."/></w:lvl>"#,
                i + 1
            )
        })
        .collect();
    let xml = format!(
        r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0">{levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    for (advancing, expected) in [
        (0, ["P-2", "", "", "", "", "00", "0", "", "00."]),
        (1, ["P-1", "C", "", "", "", "00", "0", "", "00."]),
    ] {
        let mut map = NumberingMap::parse(&xml, &HashMap::new());
        for level in [0, 1, 8, 8, advancing] {
            map.advance(1, level);
        }
        let marker = map.resolve_text_word_zero(1, 8);
        assert_eq!(
            marker.split('/').collect::<Vec<_>>(),
            expected,
            "ilvl={}",
            advancing + 16
        );
    }
}

#[test]
fn mc_ignorable_and_process_content_limit_the_validated_xml() {
    let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let xml = format!(
        r#"<w:document xmlns:w="{ns}" xmlns:mc="{mc}" xmlns:u="urn:future" mc:Ignorable="u"><w:body><u:opaque><w:ilvl w:val="abc"/></u:opaque><u:wrapper><w:ilvl w:val="0"/></u:wrapper></w:body></w:document>"#
    );
    let document = parse_guarded(&xml).expect("test XML");
    assert_eq!(validate_paragraph_ilvls(document.root_element()), Ok(()));

    let xml = format!(
        r#"<w:document xmlns:w="{ns}" xmlns:mc="{mc}" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper"><w:body><u:opaque><w:ilvl w:val="abc"/></u:opaque><u:wrapper><w:ilvl w:val="abc"/></u:wrapper></w:body></w:document>"#
    );
    let document = parse_guarded(&xml).expect("test XML");
    assert_eq!(
        validate_paragraph_ilvls(document.root_element()),
        Err("OOXML_DOCX_ILVL:cannot-open:non-decimal".to_string())
    );

    let xml = format!(
        r#"<w:document xmlns:w="{ns}" xmlns:mc="{mc}" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:*"><w:body><u:anotherWrapper><w:ilvl w:val="abc"/></u:anotherWrapper></w:body></w:document>"#
    );
    let document = parse_guarded(&xml).expect("test XML");
    assert_eq!(
        validate_paragraph_ilvls(document.root_element()),
        Err("OOXML_DOCX_ILVL:cannot-open:non-decimal".to_string())
    );

    let xml = format!(
        r#"<w:numbering xmlns:w="{ns}" xmlns:mc="{mc}" xmlns:u="urn:future"><mc:AlternateContent><mc:Choice Requires="u"><w:lvl w:ilvl="9"/></mc:Choice><mc:Fallback><w:lvl w:ilvl="8"/></mc:Fallback></mc:AlternateContent></w:numbering>"#
    );
    let document = parse_guarded(&xml).expect("test XML");
    assert_eq!(validate_level_definitions(document.root_element()), Ok(()));
}
