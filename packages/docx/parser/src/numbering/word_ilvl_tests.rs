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
