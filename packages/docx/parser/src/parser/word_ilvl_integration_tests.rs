use super::*;
use std::io::{Cursor, Read, Write};
use zip::{write::SimpleFileOptions, ZipWriter};

const NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn package(
    x: Option<&str>,
    style: bool,
    levels: usize,
    override_eight: bool,
    definition_nine: bool,
) -> Vec<u8> {
    let level_xml: String = (0..levels).map(|i| {
        let format = [
            "decimal",
            "upperLetter",
            "lowerLetter",
            "upperRoman",
            "lowerRoman",
            "decimalZero",
            "decimal",
            "upperLetter",
            "decimalZero",
        ][i];
        format!(
            r#"<w:lvl w:ilvl="{i}"><w:start w:val="{}"/><w:numFmt w:val="{format}"/><w:lvlText w:val="L{i}-%{}."/><w:lvlJc w:val="{}"/><w:pPr><w:ind w:left="{}" w:hanging="{}"/></w:pPr><w:rPr><w:rFonts w:ascii="{}"/></w:rPr></w:lvl>"#,
            i+1, i+1, ["left","center","right"][i%3], 360+180*i, 180+20*i,
            ["Times New Roman","Courier New","Arial"][i%3]
        )
    }).collect();
    let replacement = if override_eight {
        r#"<w:lvlOverride w:ilvl="8"><w:startOverride w:val="20"/><w:lvl w:ilvl="8"><w:start w:val="20"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="OV-%9."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="2400" w:hanging="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr></w:lvl></w:lvlOverride>"#
    } else {
        ""
    };
    let invalid = if definition_nine {
        r#"<w:lvlOverride w:ilvl="9"><w:lvl w:ilvl="9"><w:start w:val="1"/></w:lvl></w:lvlOverride>"#
    } else {
        ""
    };
    let numbering = format!(
        r#"<w:numbering xmlns:w="{NS}"><w:abstractNum w:abstractNumId="0">{level_xml}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/>{replacement}{invalid}</w:num></w:numbering>"#
    );
    let target = x.map_or_else(String::new, |value| format!(r#"<w:ilvl w:val="{value}"/>"#));
    let para = |label: &str, ilvl: &str, styled: bool| {
        let props = if styled {
            r#"<w:pStyle w:val="Probe"/>"#.to_string()
        } else {
            format!(r#"<w:numPr>{ilvl}<w:numId w:val="1"/></w:numPr>"#)
        };
        format!(r#"<w:p><w:pPr>{props}</w:pPr><w:r><w:t>{label}</w:t></w:r></w:p>"#)
    };
    let body = [
        para("before 0", r#"<w:ilvl w:val="0"/>"#, false),
        para("before 1", r#"<w:ilvl w:val="1"/>"#, false),
        para("before 8 A", r#"<w:ilvl w:val="8"/>"#, false),
        para("before 8 B", r#"<w:ilvl w:val="8"/>"#, false),
        para("TARGET", &target, style),
        para("after 8", r#"<w:ilvl w:val="8"/>"#, false),
        para("after 1", r#"<w:ilvl w:val="1"/>"#, false),
        para("after 8 restart", r#"<w:ilvl w:val="8"/>"#, false),
        para("after 0", r#"<w:ilvl w:val="0"/>"#, false),
        para("after 1 restart", r#"<w:ilvl w:val="1"/>"#, false),
        para("after 8 final", r#"<w:ilvl w:val="8"/>"#, false),
    ]
    .concat();
    let document = format!(r#"<w:document xmlns:w="{NS}"><w:body>{body}</w:body></w:document>"#);
    let styles = if style {
        format!(
            r#"<w:styles xmlns:w="{NS}"><w:style w:type="paragraph" w:styleId="Probe"><w:pPr><w:numPr>{target}<w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>"#
        )
    } else {
        format!(r#"<w:styles xmlns:w="{NS}"/>"#)
    };
    let mut bytes = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut bytes));
        write_test_content_types(&mut zip);
        for (name, text) in [
            ("word/document.xml", document),
            ("word/numbering.xml", numbering),
            ("word/styles.xml", styles),
        ] {
            zip.start_file(name, SimpleFileOptions::default())
                .expect("test package entry");
            zip.write_all(text.as_bytes()).expect("test package write");
        }
        zip.finish().expect("test package finish");
    }
    bytes
}

fn paragraph_markers(bytes: &[u8]) -> Vec<Option<String>> {
    parse_from_bytes(bytes)
        .expect("test package parses")
        .body
        .iter()
        .filter_map(|element| match element {
            BodyElement::Paragraph(paragraph) => {
                Some(paragraph.numbering.as_ref().map(|n| n.text.clone()))
            }
            _ => None,
        })
        .collect()
}

fn focused_counter_package(value: &str, level: usize) -> Vec<u8> {
    let source = package(Some(value), false, 9, false, false);
    let mut input = zip::ZipArchive::new(Cursor::new(source)).expect("test source archive");
    let rows = [
        ("before 0", "0".to_string()),
        ("before 1", "1".to_string()),
        ("before focus A", level.to_string()),
        ("before focus B", level.to_string()),
        ("TARGET", value.to_string()),
        ("after focus", level.to_string()),
        ("after 8", "8".to_string()),
        ("after 0", "0".to_string()),
    ];
    let body: String = rows
        .iter()
        .map(|(label, ilvl)| {
            format!(
                r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{label}</w:t></w:r></w:p>"#
            )
        })
        .collect();
    let mut output = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut output));
        for index in 0..input.len() {
            let mut part = input.by_index(index).expect("test source part");
            let name = part.name().to_string();
            let mut text = String::new();
            part.read_to_string(&mut text).expect("test XML part");
            if name == "word/document.xml" {
                text =
                    format!(r#"<w:document xmlns:w="{NS}"><w:body>{body}</w:body></w:document>"#);
            }
            zip.start_file(name, SimpleFileOptions::default())
                .expect("test output part");
            zip.write_all(text.as_bytes()).expect("test output XML");
        }
        zip.finish().expect("test output archive");
    }
    output
}

#[test]
fn malformed_levels_18_to_23_advance_the_corresponding_counter() {
    for (value, level, next_marker) in [
        ("18", 2, "L2-f."),
        ("19", 3, "L3-VII."),
        ("20", 4, "L4-viii."),
        ("21", 5, "L5-09."),
        ("22", 6, "L6-10."),
        ("23", 7, "L7-K."),
    ] {
        let markers = paragraph_markers(&focused_counter_package(value, level));
        assert_eq!(markers[4].as_deref(), Some("L8-00."), "{value} target");
        assert_eq!(markers[5].as_deref(), Some(next_marker), "{value} counter");
        assert_eq!(markers[6].as_deref(), Some("L8-09."), "{value} after 8");
        assert_eq!(markers[7].as_deref(), Some("L0-2."), "{value} after 0");
    }
}

#[test]
fn all_256_word_pdf_rows_reach_the_public_paragraph_model() {
    const DEFINED: [&str; 9] = [
        "L0-2.", "L1-C.", "L2-d.", "L3-V.", "L4-vi.", "L5-07.", "L6-8.", "L7-I.", "L8-11.",
    ];
    for (byte, class) in crate::numbering::word_ilvl_tests::OBSERVED_BYTE_CLASSES
        .bytes()
        .enumerate()
    {
        let bytes = package(Some(&byte.to_string()), false, 9, false, false);
        let document = parse_from_bytes(&bytes).expect("test package parses");
        let paragraphs: Vec<_> = document
            .body
            .iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(paragraph) => Some(paragraph.as_ref()),
                _ => None,
            })
            .collect();
        assert_eq!(paragraphs.len(), 11);
        let expected_target = if byte <= 8 {
            Some(DEFINED[byte])
        } else {
            match class {
                b'N' => None,
                b'O' => Some("L0-2."),
                b'S' => Some("L8-09."),
                b'8' => Some("L8-11."),
                b'0'..=b'7' => Some("L8-00."),
                _ => unreachable!(),
            }
        };
        let target = paragraphs[4];
        assert_eq!(
            target.numbering.as_ref().map(|number| number.text.as_str()),
            expected_target,
            "ilvl={byte} marker"
        );
        let expected_level = if byte <= 8 {
            Some(byte)
        } else if class == b'N' {
            None
        } else if class == b'O' {
            Some(0)
        } else {
            Some(8)
        };
        assert_eq!(
            target
                .numbering
                .as_ref()
                .map(|number| number.level as usize),
            expected_level,
            "ilvl={byte} formatting level"
        );
        if let Some(level) = expected_level {
            let number = target.numbering.as_ref().expect("test target marker");
            assert_eq!(
                target.indent_left,
                (18 + 9 * level) as f64,
                "ilvl={byte} left"
            );
            assert_eq!(
                target.indent_first,
                -((9 + level) as f64),
                "ilvl={byte} hanging"
            );
            assert_eq!(
                number.font_family.as_deref(),
                Some(["Times New Roman", "Courier New", "Arial"][level % 3])
            );
            assert_eq!(number.jc, ["left", "center", "right"][level % 3]);
        } else {
            assert_eq!((target.indent_left, target.indent_first), (0.0, 0.0));
        }
        let after8 = if class == b'8' {
            "L8-12."
        } else if class == b'N' || class == b'S' {
            "L8-11."
        } else {
            "L8-09."
        };
        let after1 = if class == b'1' { "L1-D." } else { "L1-C." };
        let after0 = if class == b'0' || class == b'O' {
            "L0-3."
        } else {
            "L0-2."
        };
        assert_eq!(
            paragraphs[5]
                .numbering
                .as_ref()
                .expect("after level 8 marker")
                .text,
            after8,
            "ilvl={byte} after 8"
        );
        assert_eq!(
            paragraphs[6]
                .numbering
                .as_ref()
                .expect("after level 1 marker")
                .text,
            after1,
            "ilvl={byte} after 1"
        );
        assert_eq!(
            paragraphs[8]
                .numbering
                .as_ref()
                .expect("after level 0 marker")
                .text,
            after0,
            "ilvl={byte} after 0"
        );
        for (index, expected) in [(7, "L8-09."), (9, "L1-B."), (10, "L8-09.")] {
            assert_eq!(
                paragraphs[index]
                    .numbering
                    .as_ref()
                    .expect("restart marker")
                    .text,
                expected,
                "ilvl={byte} restart paragraph {index}"
            );
        }
    }
}

#[test]
fn observed_marker_and_counter_sequences_reach_the_document_model() {
    for (value, target, next8, next1, next0) in [
        ("8", Some("L8-11."), "L8-12.", "L1-C.", "L0-2."),
        ("9", None, "L8-11.", "L1-C.", "L0-2."),
        ("13", Some("L8-09."), "L8-11.", "L1-C.", "L0-2."),
        ("15", Some("L0-2."), "L8-09.", "L1-C.", "L0-3."),
        ("16", Some("L8-00."), "L8-09.", "L1-C.", "L0-3."),
        ("17", Some("L8-00."), "L8-09.", "L1-D.", "L0-2."),
        ("24", Some("L8-11."), "L8-12.", "L1-C.", "L0-2."),
        ("25", Some("L8-09."), "L8-11.", "L1-C.", "L0-2."),
        ("127", Some("L8-09."), "L8-11.", "L1-C.", "L0-2."),
        ("128", Some("L8-00."), "L8-09.", "L1-C.", "L0-3."),
        ("255", Some("L8-09."), "L8-11.", "L1-C.", "L0-2."),
        ("256", Some("L0-2."), "L8-09.", "L1-C.", "L0-3."),
        ("-1", Some("L8-09."), "L8-11.", "L1-C.", "L0-2."),
        ("4294967296", Some("L0-2."), "L8-09.", "L1-C.", "L0-3."),
    ] {
        let markers = paragraph_markers(&package(Some(value), false, 9, false, false));
        assert_eq!(markers[4].as_deref(), target, "{value} target");
        assert_eq!(markers[5].as_deref(), Some(next8), "{value} after 8");
        assert_eq!(markers[6].as_deref(), Some(next1), "{value} after 1");
        assert_eq!(markers[8].as_deref(), Some(next0), "{value} after 0");
    }
}

#[test]
fn style_short_definition_and_override_use_the_measured_fallback_formatting() {
    let styled = parse_from_bytes(&package(Some("128"), true, 9, false, false))
        .expect("styled test package parses");
    let BodyElement::Paragraph(target) = &styled.body[4] else {
        panic!("target paragraph")
    };
    let marker = target.numbering.as_ref().expect("styled target marker");
    assert_eq!(marker.text, "L8-00.");
    assert_eq!(
        (
            marker.level,
            marker.indent_left,
            marker.jc.as_str(),
            marker.font_family.as_deref()
        ),
        (8, 90.0, "right", Some("Arial"))
    );
    assert_eq!((target.indent_left, target.indent_first), (90.0, -17.0));

    let short = paragraph_markers(&package(Some("16"), false, 2, false, false));
    assert_eq!(short[4], None);
    assert_eq!(short[8].as_deref(), Some("L0-3."));

    let overridden = paragraph_markers(&package(Some("127"), false, 9, true, false));
    assert_eq!(overridden[4].as_deref(), Some("OV-IX."));
    assert_eq!(overridden[5].as_deref(), Some("OV-XXII."));
}

#[test]
fn every_measured_extended_direct_value_matches_its_word_byte_control() {
    // Each pair is a separate Word PDF row. The byte-side expectations are
    // checked independently by all_256_word_pdf_rows_reach_the_public_paragraph_model.
    for (value, byte) in [
        ("0008", "8"),
        ("+8", "8"),
        ("256", "0"),
        ("257", "1"),
        ("511", "255"),
        ("512", "0"),
        ("65535", "255"),
        ("65536", "0"),
        ("2147483647", "255"),
        ("2147483648", "0"),
        ("4294967295", "255"),
        ("4294967296", "0"),
        ("4294967297", "1"),
        ("-1", "255"),
        ("-2", "254"),
        ("-128", "128"),
        ("-129", "127"),
        ("-256", "0"),
        ("", "0"),
    ] {
        assert_eq!(
            paragraph_markers(&package(Some(value), false, 9, false, false)),
            paragraph_markers(&package(Some(byte), false, 9, false, false)),
            "Word control for {value:?}"
        );
    }
    assert_eq!(
        paragraph_markers(&package(None, false, 9, false, false)),
        paragraph_markers(&package(Some("0"), false, 9, false, false)),
        "numId without ilvl"
    );
}

#[test]
fn every_measured_style_short_and_override_row_matches_word_markers() {
    for value in ["8", "9", "127", "128", "255"] {
        assert_eq!(
            paragraph_markers(&package(Some(value), true, 9, false, false)),
            paragraph_markers(&package(Some(value), false, 9, false, false)),
            "style-origin ilvl={value}"
        );
    }

    for (value, after_one, after_zero) in [
        ("8", "L1-C.", "L0-2."),
        ("9", "L1-C.", "L0-2."),
        ("127", "L1-C.", "L0-2."),
        ("255", "L1-C.", "L0-2."),
        ("16", "L1-C.", "L0-3."),
        ("17", "L1-D.", "L0-2."),
        ("24", "L1-C.", "L0-2."),
        ("32", "L1-C.", "L0-3."),
    ] {
        let markers = paragraph_markers(&package(Some(value), false, 2, false, false));
        assert_eq!(markers[4], None, "short ilvl={value} target");
        assert_eq!(markers[5], None, "short ilvl={value} after 8");
        assert_eq!(
            markers[6].as_deref(),
            Some(after_one),
            "short ilvl={value} after 1"
        );
        assert_eq!(
            markers[8].as_deref(),
            Some(after_zero),
            "short ilvl={value} after 0"
        );
    }

    for (value, target, after_eight, after_one, after_zero) in [
        ("8", Some("OV-XXII."), "OV-XXIII.", "L1-C.", "L0-2."),
        ("9", None, "OV-XXII.", "L1-C.", "L0-2."),
        ("127", Some("OV-IX."), "OV-XXII.", "L1-C.", "L0-2."),
        ("255", Some("OV-IX."), "OV-XXII.", "L1-C.", "L0-2."),
        ("16", Some("OV-."), "OV-XX.", "L1-C.", "L0-3."),
        ("17", Some("OV-."), "OV-XX.", "L1-D.", "L0-2."),
        ("24", Some("OV-XXII."), "OV-XXIII.", "L1-C.", "L0-2."),
        ("32", Some("OV-."), "OV-XX.", "L1-C.", "L0-3."),
    ] {
        let markers = paragraph_markers(&package(Some(value), false, 9, true, false));
        assert_eq!(
            markers[4].as_deref(),
            target,
            "override ilvl={value} target"
        );
        assert_eq!(
            markers[5].as_deref(),
            Some(after_eight),
            "override ilvl={value} after 8"
        );
        assert_eq!(
            markers[6].as_deref(),
            Some(after_one),
            "override ilvl={value} after 1"
        );
        assert_eq!(
            markers[8].as_deref(),
            Some(after_zero),
            "override ilvl={value} after 0"
        );
    }
}

#[test]
fn invalid_word_levels_reject_the_document_instead_of_degrading_it() {
    let empty = paragraph_markers(&package(Some(""), false, 9, false, false));
    assert_eq!(empty[4].as_deref(), Some("L0-2."));
    assert_eq!(empty[8].as_deref(), Some("L0-3."));
    for value in [" 8 ", "abc"] {
        let result = parse_from_bytes(&package(Some(value), false, 9, false, false));
        assert!(
            result.unwrap_err().starts_with(WORD_ILVL_ERROR_PREFIX),
            "{value:?}"
        );
    }
    let repair = parse_from_bytes(&package(Some("9"), false, 9, false, true));
    assert_eq!(
        repair.unwrap_err(),
        "OOXML_DOCX_ILVL:repair-required:level-definition"
    );
}
