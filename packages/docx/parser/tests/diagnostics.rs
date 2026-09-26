use docx_parser::parse_docx_native;
use serde_json::{json, Value};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_W_NS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_WP_NS: &str = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

fn build_docx(body: &str) -> Vec<u8> {
    build_docx_with_styles(body, None)
}

fn build_docx_with_styles(body: &str, styles: Option<&str>) -> Vec<u8> {
    build_docx_with_parts(body, styles, None)
}

fn build_docx_with_parts(body: &str, styles: Option<&str>, footnotes: Option<&str>) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
             <w:body>{body}<w:sectPr/></w:body>
           </w:document>"#
    );
    let footnotes_relationship = footnotes
        .is_some()
        .then_some(format!(
            r#"<Relationship Id="rIdFootnotes" Type="{R_NS}/footnotes" Target="footnotes.xml"/>"#
        ))
        .unwrap_or_default();
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
             <Relationship Id="rIdImage" Type="{R_NS}/image" Target="media/image.png"/>
             {footnotes_relationship}
           </Relationships>"#
    );
    let mut bytes = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
        zip.start_file("[Content_Types].xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
        )
        .unwrap();
        zip.start_file("word/document.xml", SimpleFileOptions::default())
            .expect("document entry");
        zip.write_all(document.as_bytes()).expect("document XML");
        zip.start_file("word/_rels/document.xml.rels", SimpleFileOptions::default())
            .expect("document relationships entry");
        zip.write_all(relationships.as_bytes())
            .expect("document relationships XML");
        zip.start_file("word/media/image.png", SimpleFileOptions::default())
            .expect("image entry");
        zip.write_all(PNG_1X1).expect("image bytes");
        if let Some(styles) = styles {
            zip.start_file("word/styles.xml", SimpleFileOptions::default())
                .expect("styles entry");
            zip.write_all(styles.as_bytes()).expect("styles XML");
        }
        if let Some(footnotes) = footnotes {
            zip.start_file("word/footnotes.xml", SimpleFileOptions::default())
                .expect("footnotes entry");
            zip.write_all(footnotes.as_bytes()).expect("footnotes XML");
        }
        zip.finish().expect("finish zip");
    }
    bytes
}

fn parse(body: &str) -> Value {
    let json = parse_docx_native(&build_docx(body)).expect("minimal DOCX parses");
    serde_json::from_str(&json).expect("parser output JSON")
}

fn parse_with_styles(body: &str, styles: &str) -> Value {
    let json = parse_docx_native(&build_docx_with_styles(body, Some(styles))).expect("DOCX parses");
    serde_json::from_str(&json).expect("parser output JSON")
}

fn parse_with_footnotes(body: &str, footnotes: &str) -> Value {
    let json = parse_docx_native(&build_docx_with_parts(body, None, Some(footnotes)))
        .expect("DOCX with footnotes parses");
    serde_json::from_str(&json).expect("parser output JSON")
}

fn picture_payload(extent: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline>{extent}
             <a:graphic><a:graphicData><a:blip r:embed="rIdImage"/></a:graphicData></a:graphic>
           </wp:inline></w:drawing>"#
    )
}

fn picture(extent: &str) -> String {
    format!("<w:r>{}</w:r>", picture_payload(extent))
}

fn image_run_count(value: &Value) -> usize {
    match value {
        Value::Array(values) => values.iter().map(image_run_count).sum(),
        Value::Object(values) => {
            usize::from(values.contains_key("imagePath"))
                + values.values().map(image_run_count).sum::<usize>()
        }
        Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => 0,
    }
}

fn object_key_count(value: &Value, key: &str) -> usize {
    match value {
        Value::Array(values) => values
            .iter()
            .map(|value| object_key_count(value, key))
            .sum(),
        Value::Object(values) => {
            usize::from(values.contains_key(key))
                + values
                    .values()
                    .map(|value| object_key_count(value, key))
                    .sum::<usize>()
        }
        Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => 0,
    }
}

#[test]
fn reports_stable_private_diagnostics_without_raw_values_or_document_text() {
    let document = parse(&format!(
        r#"
        <w:p>
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr><w:t>PRIVATE_SENTINEL</w:t></w:r>
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr><w:t>duplicate</w:t></w:r>
        </w:p>
        <w:p><w:r><w:rPr><w:effect w:val="producer-private-value"/></w:rPr></w:r></w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        "#,
        picture(""),
        picture(r#"<wp:extent cx="not-a-coordinate" cy="-1"/>"#),
        picture(r#"<wp:extent cx="0" cy="12700"/>"#),
    ));

    assert_eq!(
        document["diagnostics"],
        json!([
          {
            "code": "UNSUPPORTED_TEXT_EFFECT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [0]
          },
          {
            "code": "INVALID_TEXT_EFFECT_VALUE",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [1]
          },
          {
            "code": "MISSING_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [2]
          },
          {
            "code": "INVALID_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [3]
          },
          {
            "code": "DEGENERATE_DRAWING_EXTENT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [4]
          }
        ])
    );

    let diagnostics = serde_json::to_string(&document["diagnostics"]).unwrap();
    assert!(!diagnostics.contains("PRIVATE_SENTINEL"));
    assert!(!diagnostics.contains("producer-private-value"));
    assert!(!diagnostics.contains("not-a-coordinate"));
    assert_eq!(
        image_run_count(&document["body"]),
        1,
        "missing and invalid pictures are omitted; the zero-area picture is retained"
    );
}

#[test]
fn uses_the_emitted_body_cursor_across_wrappers_and_split_paragraphs() {
    let wrapped = parse(
        r#"
        <w:sdt><w:sdtContent>
          <w:p><w:r><w:t>first</w:t></w:r></w:p>
          <w:p><w:r><w:t>second</w:t></w:r></w:p>
        </w:sdtContent></w:sdt>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(wrapped["diagnostics"][0]["path"], json!([2]));

    let split = parse(
        r#"
        <w:p>
          <w:r><w:t>before</w:t></w:r>
          <w:r><w:br w:type="page"/></w:r>
          <w:r><w:t>after</w:t></w:r>
        </w:p>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(split["diagnostics"][0]["path"], json!([3]));
}

#[test]
fn table_and_nested_table_diagnostics_use_the_owning_body_index() {
    let document = parse(
        r#"
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc>
          <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
          <w:tbl><w:tr><w:tc>
            <w:p><w:r><w:rPr><w:effect w:val="invalid-value"/></w:rPr></w:r></w:p>
          </w:tc></w:tr></w:tbl>
        </w:tc></w:tr></w:tbl>
        "#,
    );

    assert_eq!(
        document["diagnostics"],
        json!([
          {
            "code": "UNSUPPORTED_TEXT_EFFECT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [1]
          },
          {
            "code": "INVALID_TEXT_EFFECT_VALUE",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [1]
          }
        ])
    );
}

#[test]
fn non_body_story_facts_do_not_claim_the_document_part() {
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
             <w:footnote w:id="1"><w:p><w:r>
               <w:rPr><w:effect w:val="sparkle"/></w:rPr>
               {}
             </w:r></w:p></w:footnote>
           </w:footnotes>"#,
        picture_payload(r#"<wp:extent cx="invalid" cy="12700"/>"#),
    );
    let document = parse_with_footnotes(
        r#"<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>"#,
        &footnotes,
    );

    assert_eq!(document["footnotes"].as_array().unwrap().len(), 1);
    assert!(
        document.get("diagnostics").is_none(),
        "facts from word/footnotes.xml must not be mislabeled as word/document.xml"
    );
}

#[test]
fn supported_and_schema_valid_control_values_do_not_report_diagnostics() {
    let document = parse(&format!(
        r#"
        <w:p>
          <w:r><w:rPr><w:effect w:val="none"/></w:rPr><w:t>control</w:t></w:r>
          {}
          <w:r><w:rPr><w:vanish/><w:effect w:val="sparkle"/></w:rPr></w:r>
          <w:r><w:drawing><wp:inline><wp:docPr hidden="1"/></wp:inline></w:drawing></w:r>
        </w:p>
        "#,
        picture(r#"<wp:extent cx="12700" cy="25400"/>"#),
    ));

    assert!(
        document.get("diagnostics").is_none(),
        "the empty private wire stays omitted"
    );
    assert_eq!(image_run_count(&document["body"]), 1);
}

#[test]
fn extent_diagnostics_match_the_actual_picture_retention_decision() {
    let document = parse(&format!(
        r#"
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        <w:p>{}</w:p>
        "#,
        picture(r#"<wp:extent cx="27273042316900" cy="+1"/>"#),
        picture(r#"<wp:extent cx=" 12700 " cy="12700"/>"#),
        picture(r#"<wp:extent cx="&#x3000;12700" cy="12700"/>"#),
        picture(r#"<wp:extent cx="-1" cy="12700"/>"#),
        picture(r#"<wp:extent cx="12700.5" cy="12700"/>"#),
        picture(r#"<wp:extent cx="1e6" cy="12700"/>"#),
        picture(r#"<wp:extent cx="27273042316901" cy="12700"/>"#),
        picture(r#"<wp:extent cx="0" cy="12700"/>"#),
        picture(""),
    ));

    let identities: Vec<_> = document["diagnostics"]
        .as_array()
        .expect("diagnostics")
        .iter()
        .map(|diagnostic| {
            (
                diagnostic["code"].as_str().unwrap(),
                diagnostic["path"][0].as_u64().unwrap(),
            )
        })
        .collect();
    assert_eq!(
        identities,
        vec![
            ("INVALID_DRAWING_EXTENT", 2),
            ("INVALID_DRAWING_EXTENT", 3),
            ("INVALID_DRAWING_EXTENT", 4),
            ("INVALID_DRAWING_EXTENT", 5),
            ("INVALID_DRAWING_EXTENT", 6),
            ("DEGENERATE_DRAWING_EXTENT", 7),
            ("MISSING_DRAWING_EXTENT", 8),
        ]
    );
    assert_eq!(
        image_run_count(&document["body"]),
        3,
        "the maximum, whitespace-collapsed, and schema-valid zero-area pictures survive"
    );
}

#[test]
fn resolved_style_vanish_and_unconsumed_mce_scope_do_not_emit_diagnostics() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}">
             <w:style w:type="character" w:styleId="HiddenDiagnosticRun">
               <w:rPr><w:vanish/></w:rPr>
             </w:style>
           </w:styles>"#
    );
    let hidden = parse_with_styles(
        &format!(
            r#"<w:p><w:r>
                 <w:rPr><w:rStyle w:val="HiddenDiagnosticRun"/><w:effect w:val="sparkle"/></w:rPr>
                 {}
               </w:r></w:p>"#,
            picture_payload(r#"<wp:extent cx="invalid" cy="12700"/>"#)
        ),
        &styles,
    );
    assert!(hidden.get("diagnostics").is_none());
    assert_eq!(image_run_count(&hidden["body"]), 0);

    let paragraph_scope = parse(&format!(
        r#"<w:p xmlns:mc="{MC_NS}" xmlns:future="urn:unsupported:drawing">
             <mc:AlternateContent>
               <mc:Choice Requires="future">
                 <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr>{}</w:r>
               </mc:Choice>
               <mc:Fallback>
                 <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr>{}</w:r>
               </mc:Fallback>
             </mc:AlternateContent>
           </w:p>"#,
        picture_payload(r#"<wp:extent cx="invalid" cy="12700"/>"#),
        picture_payload(r#"<wp:extent cx="invalid" cy="12700"/>"#),
    ));
    assert!(
        paragraph_scope.get("diagnostics").is_none(),
        "the parser does not consume paragraph-level AlternateContent"
    );
}

#[test]
fn anchored_shape_does_not_require_wp_extent() {
    let document = parse(
        r#"<w:p><w:r><w:drawing>
             <wp:anchor xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                        behindDoc="0" relativeHeight="1">
               <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
               <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
               <wp:wrapNone/><wp:docPr id="1" name="shape"/>
               <a:graphic><a:graphicData><wps:wsp><wps:spPr>
                 <a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>
                 <a:prstGeom prst="rect"/>
               </wps:spPr></wps:wsp></a:graphicData></a:graphic>
             </wp:anchor>
           </w:drawing></w:r></w:p>"#,
    );

    assert!(document.get("diagnostics").is_none());
    assert_eq!(
        object_key_count(&document["body"], "presetGeometry"),
        1,
        "the shape is retained from a:xfrm geometry"
    );
}

#[test]
fn follows_selected_mce_branch_and_remaps_removed_cover_breaks() {
    let selected_fallback = parse(&format!(
        r#"
        <w:p xmlns:mc="{MC_NS}" xmlns:future="urn:unsupported:drawing">
          <w:r><mc:AlternateContent>
            <mc:Choice Requires="future">
              <w:drawing><wp:inline/></w:drawing>
            </mc:Choice>
            <mc:Fallback>
              <w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/></wp:inline></w:drawing>
            </mc:Fallback>
          </mc:AlternateContent></w:r>
        </w:p>
        "#
    ));
    assert!(
        selected_fallback.get("diagnostics").is_none(),
        "unselected MCE markup must not produce diagnostics"
    );

    let cover = parse(
        r#"
        <w:sdt>
          <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr>
          <w:sdtContent><w:p><w:r><w:br w:type="page"/></w:r></w:p></w:sdtContent>
        </w:sdt>
        <w:p><w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r></w:p>
        "#,
    );
    assert_eq!(
        cover["diagnostics"][0]["path"],
        json!([1]),
        "removing the redundant synthetic cover break remaps later source paths"
    );
}

#[test]
fn accepts_strict_namespaces_and_positive_coordinate_boundaries() {
    let strict = parse(&format!(
        r#"
        <w:p xmlns:w="{STRICT_W_NS}" xmlns:wp="{STRICT_WP_NS}">
          <w:r><w:rPr><w:effect w:val="sparkle"/></w:rPr></w:r>
          <w:r><w:drawing><wp:inline>
            <wp:extent cx="27273042316900" cy="+1"/>
            <a:graphic><a:graphicData><a:blip r:embed="rIdImage"/></a:graphicData></a:graphic>
          </wp:inline></w:drawing></w:r>
        </w:p>
        <w:p xmlns:wp="{STRICT_WP_NS}"><w:r><w:drawing><wp:inline>
          <wp:extent cx="27273042316901" cy="1"/>
          <a:graphic><a:graphicData><a:blip r:embed="rIdImage"/></a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p>
        "#
    ));

    assert_eq!(
        strict["diagnostics"],
        json!([
          {
            "code": "UNSUPPORTED_TEXT_EFFECT",
            "severity": "warning",
            "part": "word/document.xml",
            "path": [0]
          },
          {
            "code": "INVALID_DRAWING_EXTENT",
            "severity": "error",
            "part": "word/document.xml",
            "path": [1]
            }
        ])
    );
    assert_eq!(
        image_run_count(&strict["body"]),
        1,
        "the maximum coordinate is retained while the out-of-range picture is omitted"
    );
}
