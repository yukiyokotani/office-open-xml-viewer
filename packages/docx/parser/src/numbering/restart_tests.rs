use super::*;

fn numbering(namespace: &str, level: u32, restart: Option<&str>, instance: &str) -> NumberingMap {
    NumberingMap::parse(
        &numbering_xml(namespace, level, restart, instance),
        &HashMap::new(),
    )
}

fn numbering_xml(namespace: &str, level: u32, restart: Option<&str>, instance: &str) -> String {
    let levels = (0..9)
        .map(|i| {
            let restart = if i == level {
                restart.map(|v| format!("<w:lvlRestart w:val=\"{v}\"/>")).unwrap_or_default()
            } else {
                String::new()
            };
            format!("<w:lvl w:ilvl=\"{i}\"><w:start w:val=\"1\"/>{restart}<w:lvlText w:val=\"%{}\"/></w:lvl>", i + 1)
        })
        .collect::<String>();
    format!("<w:numbering xmlns:w=\"{namespace}\"><w:abstractNum w:abstractNumId=\"1\">{levels}</w:abstractNum><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/>{instance}</w:num></w:numbering>")
}

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

#[test]
fn restart_boundaries_follow_one_based_threshold_in_both_namespaces() {
    // ECMA-376 17.9.10: 0 never restarts; a valid value names the
    // one-based last parent that resets this level, including earlier parents.
    // Absent or out-of-range values use the immediately preceding level.
    for ns in [TRANSITIONAL, STRICT] {
        for level in 1..9 {
            let values: Vec<Option<String>> = std::iter::once(None)
                .chain((0..=level + 1).map(|v| Some(v.to_string())))
                .chain([
                    Some("-1".into()),
                    Some("4294967295".into()),
                    Some("4294967296".into()),
                ])
                .collect();
            for raw in values {
                let threshold = raw
                    .as_deref()
                    .and_then(|v| v.parse::<u32>().ok())
                    .filter(|&v| v <= level)
                    .unwrap_or(level);
                for parent in 0..level {
                    let mut map = numbering(ns, level, raw.as_deref(), "");
                    assert_eq!(map.advance(2, level), 1);
                    map.advance(2, parent);
                    assert_eq!(
                        map.advance(2, level),
                        if parent < threshold { 1 } else { 2 },
                        "namespace={ns}, level={level}, restart={raw:?}, parent={parent}"
                    );
                }
            }
        }
    }
}

#[test]
fn full_level_override_replaces_restart_instead_of_merging_abstract() {
    for (abstract_value, replacement, expected) in [
        ("1", "<w:lvlRestart w:val=\"0\"/>", 2),
        ("0", "", 1),
        ("0", "<w:lvlRestart w:val=\"1\"/>", 1),
    ] {
        let instance = format!("<w:lvlOverride w:ilvl=\"2\"><w:lvl w:ilvl=\"2\">{replacement}<w:lvlText w:val=\"%3\"/></w:lvl></w:lvlOverride>");
        let mut map = numbering(TRANSITIONAL, 2, Some(abstract_value), &instance);
        assert_eq!(map.advance(2, 2), 1);
        map.advance(2, 0);
        assert_eq!(map.advance(2, 2), expected);
    }
}

#[test]
fn never_restart_still_honors_explicit_start_override_once() {
    let mut map = numbering(
        TRANSITIONAL,
        2,
        Some("0"),
        "<w:lvlOverride w:ilvl=\"2\"><w:startOverride w:val=\"7\"/></w:lvlOverride>",
    );
    assert_eq!(map.advance(2, 2), 7);
    map.advance(2, 1);
    assert_eq!(map.advance(2, 2), 8);
    map.advance(2, 0);
    assert_eq!(map.advance(2, 2), 9);
}

#[test]
fn never_restart_parent_does_not_protect_its_descendants() {
    let mut map = numbering(TRANSITIONAL, 1, Some("0"), "");
    assert_eq!(map.advance(2, 1), 1);
    assert_eq!(map.advance(2, 2), 1);
    map.advance(2, 0);
    assert_eq!(map.advance(2, 1), 2);
    assert_eq!(map.advance(2, 2), 1);
}

#[test]
fn restart_text_reaches_both_document_parse_routes() {
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    // Establish both ancestors explicitly, isolating restart from the existing
    // synthetic-ancestor seeding behavior used for skipped levels.
    let body = [0, 1, 2, 1, 2, 0, 2].iter().map(|level| format!(
        "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>"
    )).collect::<String>();
    let document = format!(
        "<w:document xmlns:w=\"{TRANSITIONAL}\"><w:body>{body}<w:sectPr/></w:body></w:document>"
    );
    let numbering = numbering_xml(TRANSITIONAL, 2, Some("0"), "");
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
        let doc = parse(&bytes, None, None, "parse").unwrap();
        let value = serde_json::to_value(doc).unwrap();
        let texts: Vec<&str> = value["body"]
            .as_array()
            .unwrap()
            .iter()
            .map(|p| p["numbering"]["text"].as_str().unwrap())
            .collect();
        assert_eq!(texts, ["1", "1", "1", "2", "2", "2", "3"]);
    }
}
