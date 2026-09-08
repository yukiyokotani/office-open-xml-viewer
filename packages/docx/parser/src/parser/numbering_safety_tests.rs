use super::{parse, parse_streamed, DocxBodyCursor, StreamedDocumentUnit, Zip};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn numbering(start: u32) -> String {
    format!(
        r#"<w:numbering xmlns:w="{W}">
          <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
            <w:start w:val="{start}"/><w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1"/>
          </w:lvl></w:abstractNum>
          <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        </w:numbering>"#
    )
}

fn numbered(text: &str, level: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn numbered_without_level(text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn package(body: &str, numbering_xml: &str, extra_parts: &[(&str, &str)]) -> Vec<u8> {
    let document = format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#);
    let mut bytes = Vec::new();
    {
        let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
        let options = SimpleFileOptions::default();
        for (name, content) in [
            ("word/document.xml", document.as_str()),
            ("word/numbering.xml", numbering_xml),
        ]
        .into_iter()
        .chain(extra_parts.iter().copied())
        {
            archive.start_file(name, options).unwrap();
            archive.write_all(content.as_bytes()).unwrap();
        }
        archive.finish().unwrap();
    }
    bytes
}

fn native_error(bytes: &[u8]) -> String {
    let mut zip = Zip::new(Cursor::new(bytes.to_vec())).unwrap();
    parse(&mut zip).expect_err("invalid numbering must fail native parsing")
}

fn streamed_error(bytes: &[u8]) -> String {
    let mut zip = Zip::new(Cursor::new(bytes.to_vec())).unwrap();
    parse_streamed(&mut zip).expect_err("invalid numbering must fail streamed parsing")
}

fn assert_native_and_streamed_fail(bytes: &[u8], expected: &str) {
    let native = native_error(bytes);
    let streamed = streamed_error(bytes);
    assert!(native.contains(expected), "{native}");
    assert!(streamed.contains(expected), "{streamed}");
}

#[test]
fn rejects_out_of_supported_range_and_unrepresentable_levels() {
    for level in ["9", "4294967295", "4294967296"] {
        let bytes = package(&numbered("bad", level), &numbering(1), &[]);
        assert_native_and_streamed_fail(&bytes, "numbering level");
    }
}

#[test]
fn absent_and_whitespace_padded_valid_levels_remain_accepted() {
    for body in [numbered_without_level("default"), numbered("zero", " 0 ")] {
        let bytes = package(&body, &numbering(1), &[]);
        let mut native_zip = Zip::new(Cursor::new(bytes.clone())).unwrap();
        parse(&mut native_zip).expect("valid native numbering");
        let mut streamed_zip = Zip::new(Cursor::new(bytes)).unwrap();
        parse_streamed(&mut streamed_zip).expect("valid streamed numbering");
    }
}

#[test]
fn rejects_counter_increment_overflow_instead_of_wrapping() {
    let body = format!("{}{}", numbered("maximum", "0"), numbered("overflow", "0"));
    let bytes = package(&body, &numbering(u32::MAX), &[]);
    assert_native_and_streamed_fail(&bytes, "overflow");
}

#[test]
fn rejects_marker_expansion_before_returning_a_model() {
    for format in ["upperLetter", "upperRoman", "hebrew2"] {
        let definition = numbering(u32::MAX).replace("decimal", format);
        let bytes = package(&numbered("oversized marker", "0"), &definition, &[]);
        assert_native_and_streamed_fail(&bytes, "marker output too large");
    }
    let definition = numbering(1).replace("%1", &"x".repeat(64 * 1024 + 1));
    let bytes = package(&numbered("oversized template", "0"), &definition, &[]);
    assert_native_and_streamed_fail(&bytes, "marker output too large");
}

#[test]
fn cursor_emits_valid_prefix_but_not_the_offending_paragraph() {
    let body = format!(
        r#"<w:p><w:r><w:t>prefix</w:t></w:r></w:p>{}"#,
        numbered("must-not-emit", "9")
    );
    let bytes = package(&body, &numbering(1), &[]);
    let mut zip = Zip::new(Cursor::new(bytes)).unwrap();
    let mut cursor = DocxBodyCursor::start(&mut zip)
        .unwrap_or_else(|failure| panic!("cursor start failed: {}", failure.into_error()));
    match cursor.next_unit(&mut zip).unwrap() {
        StreamedDocumentUnit::Body { body } => assert_eq!(body.len(), 1),
        StreamedDocumentUnit::Complete { .. } => panic!("cursor completed before valid prefix"),
    }
    let error = match cursor.next_unit(&mut zip) {
        Ok(_) => panic!("offending paragraph emitted successfully"),
        Err(error) => error,
    };
    assert!(error.contains("numbering level"), "{error}");
    assert!(cursor.next_unit(&mut zip).is_err());
}

#[test]
fn numbering_failure_in_header_is_not_swallowed() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
    </Relationships>"#;
    let header = format!(
        r#"<w:hdr xmlns:w="{W}">{}</w:hdr>"#,
        numbered("header", "9")
    );
    let body = r#"<w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>"#;
    let bytes = package(
        body,
        &numbering(1),
        &[
            ("word/_rels/document.xml.rels", rels),
            ("word/header1.xml", &header),
        ],
    );
    assert_native_and_streamed_fail(&bytes, "numbering level");
}

#[test]
fn numbering_failure_in_wps_textbox_survives_numbering_map_clone() {
    let textbox = |level| {
        format!(
            r#"<w:p><w:r><w:drawing>
          <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                     behindDoc="0" relativeHeight="1">
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="127000" cy="127000"/><wp:wrapNone/><wp:docPr id="1" name="shape"/>
            <a:graphic><a:graphicData><wps:wsp>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr>
              <wps:txbx><w:txbxContent>{}</w:txbxContent></wps:txbx>
            </wps:wsp></a:graphicData></a:graphic>
          </wp:anchor>
        </w:drawing></w:r></w:p>"#,
            numbered("textbox", level)
        )
    };
    let valid = package(&textbox("0"), &numbering(1), &[]);
    let mut native_zip = Zip::new(Cursor::new(valid.clone())).unwrap();
    let document = parse(&mut native_zip).expect("valid WPS textbox numbering");
    assert!(
        serde_json::to_string(&document)
            .unwrap()
            .contains("textbox"),
        "valid control must traverse and retain WPS text content"
    );
    let mut streamed_zip = Zip::new(Cursor::new(valid)).unwrap();
    parse_streamed(&mut streamed_zip).expect("valid streamed WPS textbox numbering");

    let bytes = package(&textbox("9"), &numbering(1), &[]);
    assert_native_and_streamed_fail(&bytes, "numbering level");
}

#[test]
fn numbering_failure_in_footnote_is_checked_at_finish() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
    </Relationships>"#;
    let footnotes = |level| {
        format!(
            r#"<w:footnotes xmlns:w="{W}"><w:footnote w:id="1">{}</w:footnote></w:footnotes>"#,
            numbered("footnote", level)
        )
    };
    let body = r#"<w:p><w:r><w:t>body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#;
    let valid_footnotes = footnotes("0");
    let valid = package(
        body,
        &numbering(1),
        &[
            ("word/_rels/document.xml.rels", rels),
            ("word/footnotes.xml", &valid_footnotes),
        ],
    );
    let mut valid_zip = Zip::new(Cursor::new(valid)).unwrap();
    let document = parse(&mut valid_zip).expect("valid footnote numbering");
    assert_eq!(document.footnotes.len(), 1);

    let invalid_footnotes = footnotes("9");
    let bytes = package(
        body,
        &numbering(1),
        &[
            ("word/_rels/document.xml.rels", rels),
            ("word/footnotes.xml", &invalid_footnotes),
        ],
    );
    assert_native_and_streamed_fail(&bytes, "numbering level");
}
