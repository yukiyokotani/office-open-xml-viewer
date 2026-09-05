use super::*;

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

fn map(namespace: &str, property: &str, replacement: Option<&str>) -> NumberingMap {
    NumberingMap::parse(&xml(namespace, property, replacement), &HashMap::new())
}

fn xml(namespace: &str, property: &str, replacement: Option<&str>) -> String {
    let definitions = [("upperRoman", "3", "%1"), ("lowerLetter", "4", "%1.%2"), ("decimalZero", "5", "%1.%2.%3")]
        .iter().enumerate().map(|(level,(format,start,text))| {
            let property = if level == 2 { property } else { "" };
            format!("<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"{start}\"/><w:numFmt w:val=\"{format}\"/>{property}<w:lvlText w:val=\"{text}\"/></w:lvl>")
        }).collect::<String>();
    let replacement = replacement.map(|property| format!(
        "<w:lvlOverride w:ilvl=\"2\"><w:lvl w:ilvl=\"2\"><w:start w:val=\"5\"/><w:numFmt w:val=\"decimalZero\"/>{property}<w:lvlText w:val=\"%1.%2.%3\"/></w:lvl></w:lvlOverride>"
    )).unwrap_or_default();
    format!(
        "<w:numbering xmlns:w=\"{namespace}\" xmlns:x=\"urn:not-wordprocessingml\"><w:abstractNum w:abstractNumId=\"1\">{definitions}</w:abstractNum><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/>{replacement}</w:num></w:numbering>"
    )
}

#[test]
fn legal_numbering_uses_common_boolean_semantics_in_both_namespaces() {
    for ns in [TRANSITIONAL, STRICT] {
        for (property, legal) in [
            ("", false),
            ("<w:isLgl/>", true),
            ("<w:isLgl w:val=\"1\"/>", true),
            ("<w:isLgl w:val=\"true\"/>", true),
            ("<w:isLgl w:val=\"on\"/>", true),
            ("<w:isLgl w:val=\"0\"/>", false),
            ("<w:isLgl w:val=\"false\"/>", false),
            ("<w:isLgl w:val=\"off\"/>", false),
            ("<x:isLgl/>", false),
        ] {
            let map = map(ns, property, None);
            assert_eq!(
                map.resolve_text(2, 2, 5),
                if legal { "3.4.5" } else { "III.d.05" },
                "{ns}: {property}"
            );
            // This flag changes only this marker, not its ancestor definitions.
            assert_eq!(map.resolve_text(2, 0, 3), "III");
            assert_eq!(map.resolve_text(2, 1, 4), "III.d");
            assert_eq!(map.get_level(2, 2).unwrap().format, "decimalZero");
        }
    }
}

#[test]
fn complete_replacement_can_enable_disable_or_omit_legal_numbering() {
    for (base, replacement, expected) in [
        ("", "<w:isLgl/>", "3.4.5"),
        ("<w:isLgl/>", "<w:isLgl w:val=\"0\"/>", "III.d.05"),
        ("<w:isLgl/>", "", "III.d.05"),
    ] {
        assert_eq!(
            map(TRANSITIONAL, base, Some(replacement)).resolve_text(2, 2, 5),
            expected
        );
    }
}

#[test]
fn legal_numbering_preserves_counter_progression_and_literal_text() {
    let mut legal = map(TRANSITIONAL, "<w:isLgl/>", None);
    let mut regular = map(TRANSITIONAL, "", None);
    for level in [0, 1, 2, 2, 1, 2, 0, 1, 2] {
        assert_eq!(legal.advance(2, level), regular.advance(2, level));
    }
    let definition = &mut legal.abstract_nums.get_mut(&1).unwrap()[2];
    definition.text = "Section %1 / %1 - %3!".into();
    assert_eq!(legal.resolve_text(2, 2, 5), "Section 4 / 4 - 5!");
}

#[test]
fn legal_numbering_uses_normative_decimal_even_for_none_and_zero() {
    // ECMA-376 17.9.4 says regardless of actual numFmt. MS-OE376
    // 2.1.280(b) documents Word retaining none instead; no such exception is
    // inferred here. A zero counter remains zero, not the first positive value.
    let mut legal = map(TRANSITIONAL, "<w:isLgl/>", None);
    legal.abstract_nums.get_mut(&1).unwrap()[0].format = "none".into();
    assert_eq!(legal.resolve_text(2, 2, 0), "3.4.0");
    assert_eq!(legal.resolve_text(2, 0, 3), "");
}

#[test]
fn legal_numbering_reaches_both_document_parse_routes() {
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    let body = [0, 1, 2, 2].iter().map(|level| format!(
        "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>"
    )).collect::<String>();
    let document = format!(
        "<w:document xmlns:w=\"{TRANSITIONAL}\"><w:body>{body}<w:sectPr/></w:body></w:document>"
    );
    let numbering = xml(TRANSITIONAL, "<w:isLgl/>", None);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in [
        ("word/document.xml", document),
        ("word/numbering.xml", numbering),
    ] {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(content.as_bytes()).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();
    for parse in [
        crate::parser::parse_from_bytes_with_limits,
        crate::parser::parse_from_bytes_streamed_with_limits,
    ] {
        let value = serde_json::to_value(parse(&bytes, None, None, "parse").unwrap()).unwrap();
        let texts: Vec<&str> = value["body"]
            .as_array()
            .unwrap()
            .iter()
            .map(|p| p["numbering"]["text"].as_str().unwrap())
            .collect();
        assert_eq!(texts, ["III", "III.d", "3.4.5", "3.4.6"]);
    }
}
