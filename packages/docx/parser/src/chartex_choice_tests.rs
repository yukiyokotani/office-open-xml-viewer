//! Synthetic package tests exercise both production parsing paths.
use crate::parser;
use std::io::{Cursor, Write};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:u="urn:unsupported" xmlns:v="urn:schemas-microsoft-com:vml""#;
fn drawing(payload: &str, uri: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="{uri}">{payload}</a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}
fn live(prefix: &str) -> String {
    drawing(
        &format!(r#"<{prefix}:chart r:id="chart"/>"#),
        "http://schemas.microsoft.com/office/drawing/2014/chartex",
    )
}
fn picture() -> String {
    picture_with_rid("image")
}
fn picture_with_rid(rid: &str) -> String {
    drawing(
        &format!(r#"<a:blip r:embed="{rid}"/>"#),
        "http://schemas.openxmlformats.org/drawingml/2006/picture",
    )
}
fn ac(choice: &str, fallback: &str) -> String {
    ac_requires("cx", choice, fallback)
}
fn ac_requires(requires: &str, choice: &str, fallback: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="{requires}">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}
fn chart(layout: &str) -> String {
    format!(
        r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">7</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="{layout}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
    )
}

fn parts_package(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    parser::write_test_content_types(&mut writer);
    for (path, data) in parts {
        writer
            .start_file(
                *path,
                zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated),
            )
            .unwrap();
        writer.write_all(data).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn document(run: &str) -> String {
    format!("<w:document {NS}><w:body><w:p><w:r>{run}</w:r></w:p></w:body></w:document>")
}
fn chart_rel(kind: &str, target: &str, mode: &str) -> String {
    format!(r#"<Relationship Id="chart" Type="{kind}" Target="{target}" TargetMode="{mode}"/>"#)
}
fn package_with(run: &str, rel: &str, part: Option<(&str, &str)>) -> Vec<u8> {
    let doc = document(run);
    let rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rel}<Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>"#
    );
    let mut parts = vec![
        ("word/document.xml", doc.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/media/image.png", b"image".as_slice()),
    ];
    if let Some((path, data)) = part {
        parts.push((path, data.as_bytes()));
    }
    parts_package(&parts)
}
fn package(run: &str, layout: &str) -> Vec<u8> {
    package_with(
        run,
        &chart_rel(
            crate::chartex_choice::CHARTEX_REL,
            "charts/chart.xml",
            "Internal",
        ),
        Some(("word/charts/chart.xml", &chart(layout))),
    )
}

// Matches the compact package shape used by the parent-comparison probe. The
// document byte length is intentional: the hidden REF case below protects the
// parent's exact streaming inflation accounting (2,562 bytes).
fn compatibility_package(run: &str, root_attrs: &str) -> Vec<u8> {
    let namespaces = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cap="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported""#;
    let doc = format!(
        r#"<w:document {namespaces} {root_attrs}><w:body><w:p><w:r>{run}</w:r></w:p></w:body></w:document>"#
    );
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rFallback" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fallback.png"/><Relationship Id="rPrimary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/primary.png"/></Relationships>"#;
    parts_package(&[
        ("word/document.xml", doc.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/media/fallback.png", b"fallback"),
        ("word/media/primary.png", b"primary"),
    ])
}
fn outcome(data: &[u8], streaming: bool) -> serde_json::Value {
    let mut zip = parser::open_document_package(data.to_vec(), None, None, None).unwrap();
    let result = zip.run_operation(
        "matrix",
        if streaming {
            parser::parse_streamed_compatible
        } else {
            parser::parse
        },
    );
    let mut record = match result {
        Ok(doc) => serde_json::json!({"model": doc}),
        Err(error) => serde_json::json!({"error": error}),
    };
    record["usage"] = serde_json::to_value(zip.usage()).unwrap();
    record["healthy"] = serde_json::json!(zip.assert_healthy().is_ok());
    record
}
fn first_type(record: &serde_json::Value) -> Option<&str> {
    record["model"]["body"][0]["runs"][0]["type"].as_str()
}

// `Parent` cases deliberately preserve each API's pre-existing general MCE,
// field and error behavior. Full model/usage/health comparison against a parent
// executable is available via DOCXLIVE_FIXTURE_EXPORT, without retaining private
// corpora or generated artifacts in the repository.
#[derive(Clone, Copy)]
enum Expected {
    Chart,
    Picture,
    Parent,
    NestedPicture,
    NestedText,
    TypedLimit,
}
struct Case {
    name: &'static str,
    data: Vec<u8>,
    expected: Expected,
}
fn cases() -> Vec<Case> {
    let mut cases = Vec::new();
    let mut add = |name, data, expected| {
        cases.push(Case {
            name,
            data,
            expected,
        })
    };
    let cx = crate::chartex_choice::CHARTEX_NS;
    let rel_type = crate::chartex_choice::CHARTEX_REL;
    let supported = chart("clusteredColumn");
    let live = live("c");
    let pic = picture();
    // 1-5: rId provenance, closed layout set, and exact direct-child payload.
    add(
        "01-c-chart",
        package(&ac(&live, &pic), "clusteredColumn"),
        Expected::Chart,
    );
    add(
        "01-cx-chart",
        package(&ac(&self::live("cx"), &pic), "clusteredColumn"),
        Expected::Chart,
    );
    add(
        "01-requires-2015",
        package(&ac_requires("cx1", &live, &pic), "clusteredColumn"),
        Expected::Chart,
    );
    add(
        "01-requires-2014-and-2015",
        package(&ac_requires("cx cx1", &live, &pic), "clusteredColumn"),
        Expected::Chart,
    );
    add(
        "01-requires-unsupported",
        package(&ac_requires("u", &live, &pic), "clusteredColumn"),
        Expected::Picture,
    );
    add(
        "02-requires-2015-unsupported-layout",
        package(&ac_requires("cx1", &live, &pic), "pie"),
        Expected::Picture,
    );
    add(
        "02-pie",
        package(&ac(&live, &pic), "pie"),
        Expected::Picture,
    );
    for (name, rel, part) in [
        (
            "03-missing-rel",
            String::new(),
            Some(("word/charts/chart.xml", supported.as_str())),
        ),
        (
            "03-classic-rel",
            chart_rel(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "charts/chart.xml",
                "Internal",
            ),
            Some(("word/charts/chart.xml", supported.as_str())),
        ),
        (
            "03-external",
            chart_rel(rel_type, "charts/chart.xml", "External"),
            Some(("word/charts/chart.xml", supported.as_str())),
        ),
        (
            "03-missing-part",
            chart_rel(rel_type, "charts/chart.xml", "Internal"),
            None,
        ),
        ("03-missing-both", String::new(), None),
        (
            "03-nonstandard-part",
            chart_rel(rel_type, "plots/plot.xml", "Internal"),
            Some(("word/plots/plot.xml", supported.as_str())),
        ),
        (
            "03-malformed",
            chart_rel(rel_type, "charts/chart.xml", "Internal"),
            Some(("word/charts/chart.xml", "<broken")),
        ),
        (
            "03-classic-root",
            chart_rel(rel_type, "charts/chart.xml", "Internal"),
            Some((
                "word/charts/chart.xml",
                r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:pieChart/></c:plotArea></c:chart></c:chartSpace>"#,
            )),
        ),
    ] {
        add(
            name,
            package_with(&ac(&live, &pic), &rel, part),
            Expected::Picture,
        );
    }
    add(
        "04-missing-chart",
        package(&ac(&drawing("", cx), &pic), "clusteredColumn"),
        Expected::Picture,
    );
    add(
        "04-missing-rid",
        package(&ac(&drawing("<cx:chart/>", cx), &pic), "clusteredColumn"),
        Expected::Picture,
    );
    add(
        "05-wps",
        package(
            &ac(&live, &pic).replace("Requires=\"cx\"", "Requires=\"wps\""),
            "pie",
        ),
        Expected::Picture,
    );
    // 6/14: events stay below the existing event-byte bound, while their sum
    // exceeds the body-block bound. Ignored fallback text must never be retained.
    let large = format!("<w:t>{}</w:t>", "x".repeat(64 * 1024)).repeat(544);
    add(
        "06-filter-large-fallback",
        package(&ac(&live, &format!("{large}{pic}")), "pie"),
        Expected::Picture,
    );
    add(
        "07-vml",
        package(
            &ac(
                &live,
                r#"<w:pict><v:shape style="width:72pt;height:72pt"><v:imagedata r:id="image"/></v:shape></w:pict>"#,
            ),
            "pie",
        ),
        Expected::Picture,
    );
    // 8-13: preserve the parent's general MCE behavior independently per API.
    add(
        "08-no-fallback",
        package(
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="cx">{live}</mc:Choice></mc:AlternateContent>"#
            ),
            "clusteredColumn",
        ),
        Expected::Parent,
    );
    add(
        "08-no-fallback-missing",
        package_with(
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="cx">{live}</mc:Choice></mc:AlternateContent>"#
            ),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "09-text-plus-drawing",
        package_with(&ac(&format!("<w:t>text</w:t>{live}"), &pic), "", None),
        Expected::Parent,
    );
    add(
        "10-nested-picture",
        package_with(&ac(&ac(&live, &pic), &picture()), "", None),
        Expected::NestedPicture,
    );
    add(
        "10-nested-text",
        package_with(&ac(&ac(&live, "<w:t>ignored</w:t>"), &pic), "", None),
        Expected::NestedText,
    );
    for (name, payload) in [
        (
            "11-extension",
            format!("<a:extLst><a:ext uri=\"urn:payload\">{live}</a:ext></a:extLst>"),
        ),
        (
            "11-ignorable-payload",
            format!("<u:payload mc:Ignorable=\"u\">{live}</u:payload>"),
        ),
    ] {
        let primary = pic.replace("</wp:inline>", &format!("{payload}</wp:inline>"));
        add(
            name,
            package_with(&ac(&primary, &pic), "", None),
            Expected::Parent,
        );
    }
    add(
        "11-ignorable-drawing",
        package_with(
            &ac(
                &format!(
                    "<w:t>choice</w:t>{}",
                    live.replace("w:drawing", "u:drawing")
                        .replace("<u:drawing>", "<u:drawing mc:Ignorable=\"u\">")
                ),
                &pic,
            ),
            "",
            None,
        ),
        Expected::Parent,
    );
    for (name, run) in [
        (
            "12-must-understand-choice",
            ac("<w:t>choice</w:t>", "<w:t>fallback</w:t>")
                .replace("Requires=\"cx\"", "Requires=\"w\" mc:MustUnderstand=\"u\""),
        ),
        (
            "12-must-understand-fallback",
            ac("<w:t>choice</w:t>", "<w:t>fallback</w:t>")
                .replace("Requires=\"cx\"", "Requires=\"u\"")
                .replace("<mc:Fallback>", "<mc:Fallback mc:MustUnderstand=\"u\">"),
        ),
        (
            "12-process-content",
            ac("<u:keep><w:t>kept</w:t></u:keep>", "<w:t>fallback</w:t>").replace(
                "Requires=\"cx\"",
                "Requires=\"w\" mc:Ignorable=\"u\" mc:ProcessContent=\"u:keep\"",
            ),
        ),
        (
            "12-ignorable-text",
            ac("<u:t>ignored</u:t><w:t>kept</w:t>", "<w:t>fallback</w:t>")
                .replace("Requires=\"cx\"", "Requires=\"w\" mc:Ignorable=\"u\""),
        ),
        (
            "12-requires-word",
            ac("<w:t>choice</w:t>", "<w:t>fallback</w:t>")
                .replace("Requires=\"cx\"", "Requires=\"w\""),
        ),
    ] {
        add(name, package_with(&run, "", None), Expected::Parent);
    }
    let page = r#"<w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>42</w:t><w:fldChar w:fldCharType="end"/>"#;
    add(
        "13-page-choice",
        package_with(
            &ac(page, "").replace("Requires=\"cx\"", "Requires=\"w\""),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "13-comment-hyphen-page",
        package_with(
            &format!(
                r#"<w:t>A</w:t><w:commentReference w:id="0"/><w:noBreakHyphen/><w:t>B</w:t>{page}"#
            ),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "14-unselected-large",
        package_with(
            &ac("<w:t>choice</w:t>", &large).replace("Requires=\"cx\"", "Requires=\"w\""),
            "",
            None,
        ),
        Expected::Parent,
    );
    // 15: same typed resource limits and poisoning on both content classes.
    let deep = format!(
        "{}<w:t>deep</w:t>{}",
        "<w:smartTag>".repeat(253),
        "</w:smartTag>".repeat(253)
    );
    let comment = format!("<!--{}-->", "x".repeat(1024 * 1024 + 1));
    add(
        "15-depth-plain",
        package_with(&deep, "", None),
        Expected::TypedLimit,
    );
    add(
        "15-depth-chartex",
        package(&format!("{deep}{}", ac(&live, &pic)), "clusteredColumn"),
        Expected::TypedLimit,
    );
    add(
        "15-comment-plain",
        package_with(&comment, "", None),
        Expected::TypedLimit,
    );
    add(
        "15-comment-chartex",
        package(&format!("{comment}{}", ac(&live, &pic)), "clusteredColumn"),
        Expected::TypedLimit,
    );
    add(
        "16-many-text-page",
        package_with(
            &format!("{}{page}", "<w:t>x</w:t>".repeat(16 * 1024)),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "17-plain",
        package_with("<w:t>plain</w:t>", "", None),
        Expected::Parent,
    );
    add(
        "17-unselected-ref",
        package_with(
            &ac(
                "<w:t>choice</w:t>",
                "<w:instrText> REF Hidden </w:instrText>",
            )
            .replace("Requires=\"cx\"", "Requires=\"w\""),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "17-supported",
        package(&ac(&live, &pic), "clusteredColumn"),
        Expected::Chart,
    );
    // 18: story-local relationships deliberately reuse the same rId for a
    // supported body chart and unsupported header/footer charts.
    for (name, root, rel_kind) in [
        ("18-header", "hdr", "header"),
        ("18-footer", "ftr", "footer"),
    ] {
        let story_path = format!("word/{rel_kind}.xml");
        let story_rels_path = format!("word/_rels/{rel_kind}.xml.rels");
        let story = format!(
            r#"<w:{root} {NS}><w:p><w:r>{}</w:r></w:p><w:p><w:r>{}</w:r></w:p></w:{root}>"#,
            ac(&live, &pic),
            ac(&drawing(r#"<cx:chart r:id="good"/>"#, cx), &pic)
        );
        let doc = document(&ac(&live, &pic)).replace("</w:body>", &format!(r#"<w:sectPr><w:{rel_kind}Reference w:type="default" r:id="story"/></w:sectPr></w:body>"#));
        let doc_rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}<Relationship Id="story" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_kind}" Target="{rel_kind}.xml"/></Relationships>"#,
            chart_rel(rel_type, "charts/chart.xml", "Internal")
        );
        let story_rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}<Relationship Id="good" Type="{rel_type}" Target="charts/chart.xml"/><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>"#,
            chart_rel(rel_type, "charts/pie.xml", "Internal")
        );
        let pie = chart("pie");
        add(
            name,
            parts_package(&[
                ("word/document.xml", doc.as_bytes()),
                ("word/_rels/document.xml.rels", doc_rels.as_bytes()),
                (&story_path, story.as_bytes()),
                (&story_rels_path, story_rels.as_bytes()),
                ("word/charts/chart.xml", supported.as_bytes()),
                ("word/charts/pie.xml", pie.as_bytes()),
                ("word/media/image.png", b"image"),
            ]),
            Expected::Chart,
        );
    }
    add(
        "19-later-choice",
        package(
            &ac(&live, &pic).replace(
                "</mc:Choice>",
                "</mc:Choice><mc:Choice Requires=\"wps\"><w:t>later</w:t></mc:Choice>",
            ),
            "pie",
        ),
        Expected::Picture,
    );
    // First matching sibling and exact namespace counterexamples protect the
    // bounded raw-path observer from borrowing a later or nested payload.
    add(
        "first-container",
        package_with(
            &ac(
                &live.replace("<wp:inline>", "<wp:inline/><wp:inline>"),
                &pic,
            ),
            "",
            None,
        ),
        Expected::Parent,
    );
    add(
        "wrong-namespace",
        package_with(&ac(&live.replace("w:drawing", "u:drawing"), &pic), "", None),
        Expected::Parent,
    );
    add(
        "malformed-classic-uri",
        package_with(
            &ac(&picture().replace("2006/picture", "2006/&invalid;"), &pic),
            "",
            None,
        ),
        Expected::Parent,
    );

    cases
}
#[test]
fn synthetic_docx_choice_matrix() {
    let export = std::env::var_os("DOCXLIVE_FIXTURE_EXPORT").map(std::path::PathBuf::from);
    if let Some(dir) = &export {
        std::fs::create_dir_all(dir).unwrap();
    }
    for case in cases() {
        let native = outcome(&case.data, false);
        let streaming = outcome(&case.data, true);
        if let Some(dir) = &export {
            std::fs::write(dir.join(format!("{}.docx", case.name)), &case.data).unwrap();
            std::fs::write(
                dir.join(format!("{}.json", case.name)),
                serde_json::to_vec(&serde_json::json!({"native": native, "stream": streaming}))
                    .unwrap(),
            )
            .unwrap();
        }
        match case.expected {
            Expected::Chart | Expected::Picture => {
                assert_eq!(native["model"], streaming["model"], "{}", case.name);
                let expected = if matches!(case.expected, Expected::Chart) {
                    "chart"
                } else {
                    "image"
                };
                assert_eq!(first_type(&native), Some(expected), "{}", case.name);
                assert_eq!(
                    native["model"]["body"][0]["runs"].as_array().unwrap().len(),
                    1,
                    "{}",
                    case.name
                );
                assert_eq!(native["healthy"], true, "{}", case.name);
                assert_eq!(streaming["healthy"], true, "{}", case.name);
                if case.name.starts_with("18-") {
                    let key = if case.name.ends_with("header") {
                        "headers"
                    } else {
                        "footers"
                    };
                    assert_eq!(
                        native["model"][key]["default"]["body"][0]["runs"][0]["type"], "image",
                        "{}",
                        case.name
                    );
                    assert_eq!(
                        native["model"][key]["default"]["body"][1]["runs"][0]["type"], "chart",
                        "{}",
                        case.name
                    );
                }
            }
            Expected::NestedPicture => {
                // The native parent only dispatches direct drawings in an AC;
                // the streaming parent resolves nested MCE before semantic parse.
                assert_eq!(first_type(&streaming), Some("image"), "{}", case.name);
                assert_eq!(
                    native["model"]["body"][0]["runs"],
                    serde_json::json!([]),
                    "{}",
                    case.name
                );
            }
            Expected::NestedText => {
                assert_eq!(native["model"]["body"][0]["runs"], serde_json::json!([]));
                assert_eq!(streaming["model"]["body"][0]["runs"], serde_json::json!([]));
            }
            Expected::TypedLimit => {
                assert!(
                    streaming["error"]
                        .as_str()
                        .is_some_and(|error| error.contains("RESOURCE_LIMIT")),
                    "{}: {}",
                    case.name,
                    streaming
                );
                assert_eq!(streaming["healthy"], false, "{}", case.name);
            }
            Expected::Parent => {
                assert_eq!(native["healthy"], true, "{}", case.name);
                assert_eq!(streaming["healthy"], true, "{}", case.name);
                let native_runs = &native["model"]["body"][0]["runs"];
                let streamed_runs = &streaming["model"]["body"][0]["runs"];
                match case.name {
                    "12-must-understand-choice" | "12-must-understand-fallback" => {
                        assert!(streaming["model"]["parseError"]
                            .as_str()
                            .unwrap()
                            .contains("MustUnderstand"));
                        assert_eq!(native_runs, &serde_json::json!([]));
                    }
                    "12-process-content"
                    | "12-ignorable-text"
                    | "12-requires-word"
                    | "14-unselected-large" => {
                        let text = if case.name == "12-process-content"
                            || case.name == "12-ignorable-text"
                        {
                            "kept"
                        } else {
                            "choice"
                        };
                        assert_eq!(native_runs, &serde_json::json!([]));
                        assert_eq!(streamed_runs.as_array().unwrap().len(), 1);
                        assert_eq!(streamed_runs[0]["text"], text);
                    }
                    "13-page-choice" => {
                        assert_eq!(native_runs, &serde_json::json!([]));
                        assert_eq!(streamed_runs[0]["type"], "field");
                        assert_eq!(streamed_runs[0]["fieldType"], "page");
                        assert_eq!(streamed_runs[0]["fallbackText"], "42");
                    }
                    "13-comment-hyphen-page" => {
                        assert_eq!(native["model"], streaming["model"]);
                        assert_eq!(native_runs[0]["text"], "A");
                        assert_eq!(native_runs[1]["text"], "-B");
                        assert_eq!(
                            native_runs[1]["__noBreakHyphenOffsets"],
                            serde_json::json!([1])
                        );
                        assert_eq!(native_runs[2]["type"], "field");
                        assert_eq!(native_runs[2]["fieldType"], "page");
                        assert_eq!(
                            native["model"]["body"][0]["commentMarks"],
                            serde_json::json!([{"id":"0", "kind":"reference", "runIndex":1}])
                        );
                    }
                    "09-text-plus-drawing" => {
                        assert_eq!(native_runs[0]["type"], "unavailableDrawing");
                        assert_eq!(streamed_runs[0]["text"], "text");
                        assert_eq!(streamed_runs[1]["type"], "unavailableDrawing");
                    }
                    "11-extension" | "11-ignorable-payload" => {
                        assert_eq!(native["model"], streaming["model"]);
                        assert_eq!(first_type(&native), Some("image"));
                    }
                    "11-ignorable-drawing" => {
                        assert_eq!(first_type(&native), Some("unavailableDrawing"));
                        assert_eq!(streamed_runs.as_array().unwrap().len(), 1);
                        assert_eq!(streamed_runs[0]["text"], "choice");
                    }
                    "first-container" => {
                        assert_eq!(native_runs, &serde_json::json!([]));
                        assert_eq!(streamed_runs, &serde_json::json!([]));
                    }
                    "wrong-namespace" => {
                        assert_eq!(native["model"], streaming["model"]);
                        assert_eq!(first_type(&native), Some("unavailableDrawing"));
                    }
                    "malformed-classic-uri" => {
                        assert!(native["model"]["parseError"].as_str().is_some());
                        assert!(streaming["model"]["parseError"].as_str().is_some());
                    }
                    "16-many-text-page" => {
                        assert_eq!(native["model"], streaming["model"]);
                        assert_eq!(native_runs.as_array().unwrap().len(), 16 * 1024 + 1);
                        assert_eq!(native_runs[16 * 1024]["fieldType"], "page");
                    }
                    _ => {}
                }
                if case.name == "08-no-fallback-missing" {
                    assert_eq!(native["model"], streaming["model"]);
                    assert_eq!(first_type(&native), Some("unavailableDrawing"));
                }
                if case.name == "17-unselected-ref" {
                    let doc_bytes = document(
                        &ac(
                            "<w:t>choice</w:t>",
                            "<w:instrText> REF Hidden </w:instrText>",
                        )
                        .replace("Requires=\"cx\"", "Requires=\"w\""),
                    )
                    .len() as u64;
                    assert_eq!(
                        streaming["usage"]["operationInflatedBytes"].as_u64(),
                        native["usage"]["operationInflatedBytes"]
                            .as_u64()
                            .map(|bytes| bytes + doc_bytes)
                    );
                }
            }
        }
    }
}

#[test]
fn capability_2015_is_local_to_single_chartex_drawing_choices() {
    let cap_ac = |choice: &str, fallback: &str| {
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="cap" >{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
        )
    };

    let ordinary_picture = compatibility_package(
        &cap_ac(
            &picture_with_rid("rPrimary"),
            &picture_with_rid("rFallback"),
        ),
        "",
    );
    for streaming in [false, true] {
        let parsed = outcome(&ordinary_picture, streaming);
        assert_eq!(
            parsed["model"]["body"][0]["runs"][0]["imagePath"],
            "word/media/fallback.png"
        );
    }

    let ordinary_text =
        compatibility_package(&cap_ac("<w:t>choice</w:t>", "<w:t>fallback</w:t>"), "");
    let native_text = outcome(&ordinary_text, false);
    let streamed_text = outcome(&ordinary_text, true);
    assert_eq!(
        native_text["model"]["body"][0]["runs"],
        serde_json::json!([])
    );
    assert_eq!(
        streamed_text["model"]["body"][0]["runs"][0]["text"],
        "fallback"
    );

    let empty_choice = compatibility_package(
        r#"<mc:AlternateContent><mc:Choice Requires="cap"/><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent>"#,
        "",
    );
    assert_eq!(
        outcome(&empty_choice, true)["model"]["body"][0]["runs"][0]["text"],
        "fallback"
    );

    let unselected_choice_mu = compatibility_package(
        &cap_ac("<w:t>choice</w:t>", "<w:t>fallback</w:t>").replace(
            "Requires=\"cap\" >",
            "Requires=\"cap\" mc:MustUnderstand=\"u\">",
        ),
        "",
    );
    let streamed_choice_mu = outcome(&unselected_choice_mu, true);
    assert!(streamed_choice_mu["model"].get("parseError").is_none());
    assert_eq!(
        streamed_choice_mu["model"]["body"][0]["runs"][0]["text"],
        "fallback"
    );

    let chart_choice_mu = package(
        &ac_requires("cx1", &live("c"), &picture()).replace(
            "Requires=\"cx1\">",
            "Requires=\"cx1\" mc:MustUnderstand=\"u\">",
        ),
        "clusteredColumn",
    );
    for streaming in [false, true] {
        assert_eq!(
            first_type(&outcome(&chart_choice_mu, streaming)),
            Some("image")
        );
    }

    let ignorable = compatibility_package(
        "<cap:t>ignored</cap:t><w:t>kept</w:t>",
        r#"mc:Ignorable="cap""#,
    );
    assert_eq!(
        outcome(&ignorable, true)["model"]["body"][0]["runs"][0]["text"],
        "kept"
    );

    let process_content = compatibility_package(
        "<cap:keep><w:t>kept</w:t></cap:keep>",
        r#"mc:Ignorable="cap" mc:ProcessContent="cap:keep""#,
    );
    assert_eq!(
        outcome(&process_content, true)["model"]["body"][0]["runs"][0]["text"],
        "kept"
    );

    let must_understand = compatibility_package("<w:t>hello</w:t>", r#"mc:MustUnderstand="cap""#);
    let streamed_mu = outcome(&must_understand, true);
    assert!(streamed_mu["model"]["parseError"]
        .as_str()
        .is_some_and(|error| error.contains("MustUnderstand")));

    let hidden_ref = compatibility_package(
        &cap_ac(
            r#"<w:fldChar w:fldCharType="begin"/><w:instrText> REF Hidden </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>42</w:t><w:fldChar w:fldCharType="end"/>"#,
            "<w:t>fallback</w:t>",
        ),
        "",
    );
    let streamed_ref = outcome(&hidden_ref, true);
    assert_eq!(
        streamed_ref["model"]["body"][0]["runs"][0]["text"],
        "fallback"
    );
    assert_eq!(streamed_ref["usage"]["operationInflatedBytes"], 2562);
}

#[test]
fn unselected_2015_picture_descendants_do_not_activate_must_understand() {
    // ECMA-376 Part 3 §§9.3–9.4: the 2015 token is understood only for
    // this library's exact ChartEx choice shape. Descendant MustUnderstand
    // directives on an ordinary picture remain inactive when it is discarded.
    // Exercise deferral before shape classification and absorption after the
    // ordinary graphicData discards the still-open provisional branch.
    let primary = picture_with_rid("rPrimary")
        .replace("<w:drawing>", "<w:drawing mc:MustUnderstand=\"u\">")
        .replace(
            "</wp:inline>",
            "<wp:docPr id=\"1\" name=\"late\" mc:MustUnderstand=\"u\"/></wp:inline>",
        );
    let data = compatibility_package(
        &ac_requires("cap", &primary, &picture_with_rid("rFallback")),
        "",
    );
    for streaming in [false, true] {
        let parsed = outcome(&data, streaming);
        assert!(parsed["model"].get("parseError").is_none(), "{parsed}");
        assert_eq!(
            parsed["model"]["body"][0]["runs"][0]["imagePath"],
            "word/media/fallback.png"
        );
    }

    // Deferral must not suppress a mismatch in an actually selected chart.
    let selected_chart = package(
        &ac_requires(
            "cx1",
            &live("c").replace("<w:drawing>", "<w:drawing mc:MustUnderstand=\"u\">"),
            &picture(),
        ),
        "clusteredColumn",
    );
    assert!(outcome(&selected_chart, true)["model"]["parseError"]
        .as_str()
        .is_some_and(|error| error.contains("MustUnderstand")));
}

#[test]
fn selected_chartex_choice_must_understand_follows_the_effective_mce_path() {
    // ECMA-376 Part 3 §§9.3–9.4: once a ChartEx Choice is selected, its own and
    // its effective descendants' MustUnderstand namespaces are processed on
    // both paths, through the 2015-local and the parent-understood selection.
    // The unrenderable "pie" part is checked before its picture substitution.
    for (run, layout) in [
        (
            ac_requires(
                "cx1",
                &live("c").replace("<w:drawing>", "<w:drawing mc:MustUnderstand=\"u\">"),
                &picture(),
            ),
            "clusteredColumn",
        ),
        (
            ac(&live("c"), &picture()).replace(
                "Requires=\"cx\">",
                "Requires=\"cx\" mc:MustUnderstand=\"u\">",
            ),
            "pie",
        ),
    ] {
        let data = package(&run, layout);
        for streaming in [false, true] {
            let parsed = outcome(&data, streaming);
            assert!(
                parsed["model"]["parseError"]
                    .as_str()
                    .is_some_and(|error| error.contains("MustUnderstand")),
                "{parsed}"
            );
            assert_eq!(parsed["model"]["body"], serde_json::json!([]));
            assert!(parsed.get("error").is_none(), "{parsed}");
            assert_eq!(parsed["healthy"], true);
        }
    }

    // Each directive is outside that path for a distinct reason: payload made
    // ignorable by the enclosing AlternateContent, an opaque extension list,
    // and an unselected nested Choice. The chart remains live on both paths.
    let inactive = concat!(
        r#"<wp:docPr id="1" name="chart"><a:extLst><a:ext uri="urn:x" mc:MustUnderstand="u"/></a:extLst></wp:docPr>"#,
        r#"<mc:AlternateContent><mc:Choice Requires="u"><wp:cNvGraphicFramePr mc:MustUnderstand="u"/></mc:Choice></mc:AlternateContent>"#,
        r#"<u:payload><wp:cNvGraphicFramePr mc:MustUnderstand="u"/></u:payload>"#,
    );
    let run = ac_requires(
        "cx1",
        &live("c").replace("<a:graphic>", &format!("{inactive}<a:graphic>")),
        &picture(),
    )
    .replacen(
        "<mc:AlternateContent>",
        "<mc:AlternateContent mc:Ignorable=\"u\">",
        1,
    );
    let data = package(&run, "clusteredColumn");
    for streaming in [false, true] {
        let parsed = outcome(&data, streaming);
        assert!(parsed["model"].get("parseError").is_none(), "{parsed}");
        assert_eq!(first_type(&parsed), Some("chart"), "{parsed}");
    }
}

#[test]
fn selected_chartex_preflight_checks_resource_fallbacks_and_target_container() {
    // Selected ChartEx processing also visits an authored resource fallback
    // chosen by a nested AC. Its drawing's MU cannot disappear just because
    // the nested chart relationship is unrenderable. The enclosing target AC
    // itself is processed before its selected branch (Part 3 §§9.1, 9.3–9.4).
    let nested = ac_requires(
        "cx1",
        &live("c").replace("r:id=\"chart\"", "r:id=\"missing\""),
        &picture().replace("<w:drawing>", "<w:drawing mc:MustUnderstand=\"u\">"),
    );
    let nested_run = ac_requires(
        "cx1",
        &live("c").replace("</wp:inline>", &format!("{nested}</wp:inline>")),
        &picture(),
    );
    let container_run = ac_requires("cx1", &live("c"), &picture()).replacen(
        "<mc:AlternateContent>",
        "<mc:AlternateContent mc:MustUnderstand=\"u\">",
        1,
    );
    for run in [nested_run, container_run] {
        let data = package(&run, "clusteredColumn");
        for streaming in [false, true] {
            let parsed = outcome(&data, streaming);
            assert!(
                parsed["model"]["parseError"]
                    .as_str()
                    .is_some_and(|error| error.contains("MustUnderstand")),
                "{parsed}"
            );
            assert_eq!(parsed["model"]["body"], serde_json::json!([]));
            assert_eq!(parsed["healthy"], true);
        }
    }
}

#[test]
fn chartex_resource_fallback_does_not_process_filtered_children() {
    // A resource substitution retains only direct w:drawing/w:pict children.
    // The text payload is excluded before MU processing; the image is kept.
    let nested = ac_requires(
        "cx1",
        &live("c").replace("r:id=\"chart\"", "r:id=\"missing\""),
        &format!("<w:t mc:MustUnderstand=\"u\">filtered</w:t>{}", picture()),
    );
    let run = ac_requires(
        "cx1",
        &live("c").replace("</wp:inline>", &format!("{nested}</wp:inline>")),
        &picture(),
    );
    let data = package(&run, "clusteredColumn");
    for streaming in [false, true] {
        let parsed = outcome(&data, streaming);
        assert!(parsed["model"].get("parseError").is_none(), "{parsed}");
        assert_eq!(first_type(&parsed), Some("chart"), "{parsed}");
        assert_eq!(parsed["healthy"], true);
    }
}

#[test]
fn large_non_chartex_2015_choice_stays_unselected_and_healthy() {
    let large = format!("<w:t>{}</w:t>", "x".repeat(64 * 1024)).repeat(544);
    let incomplete_drawing = format!("<w:drawing><wp:inline>{large}</wp:inline></w:drawing>");
    for choice in [&large, &incomplete_drawing] {
        let data = compatibility_package(&ac_requires("cap", choice, "<w:t>fallback</w:t>"), "");
        let native = outcome(&data, false);
        let streaming = outcome(&data, true);
        assert_eq!(native["healthy"], true);
        assert_eq!(streaming["healthy"], true);
        assert!(streaming.get("error").is_none(), "{streaming}");
        assert_eq!(streaming["model"]["body"][0]["runs"][0]["text"], "fallback");
    }
}

#[test]
fn unsupported_must_understand_on_substitute_fallback_keeps_parent_choice() {
    let run = ac_requires("wps", &live("c"), &picture())
        .replace("<mc:Fallback>", "<mc:Fallback mc:MustUnderstand=\"u\">");
    let data = package_with(&run, "", None);
    let native = outcome(&data, false);
    let streaming = outcome(&data, true);
    assert_eq!(native["model"], streaming["model"]);
    assert_eq!(first_type(&native), Some("unavailableDrawing"));
}

#[test]
fn chartex_allocation_limit_poisoning_matches_native_and_streaming() {
    let series =
        r#"<cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>"#.repeat(16);
    let xml = format!(
        r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="65536"><cx:pt idx="0">7</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>{series}</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#
    );
    let data = package_with(
        &ac(&live("cx"), &picture()),
        &chart_rel(
            crate::chartex_choice::CHARTEX_REL,
            "charts/chart.xml",
            "Internal",
        ),
        Some(("word/charts/chart.xml", &xml)),
    );
    for streaming in [false, true] {
        let result = outcome(&data, streaming);
        assert_eq!(result["healthy"], false);
        let error = result["error"]
            .as_str()
            .expect("chart resource limit must not select picture fallback");
        let json: serde_json::Value = serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("typed prefix"),
        )
        .expect("typed JSON");
        assert_eq!(
            json["details"]["violation"]["resource"],
            "chartex-allocation"
        );
        assert_eq!(json["details"]["violation"]["format"], "docx");
        assert_eq!(json["details"]["violation"]["metric"], "bytes");
        assert_eq!(
            json["details"]["violation"]["limit"],
            ooxml_common::resource::HARD_MAX_CHARTEX_ALLOCATION_BYTES
        );
        assert_eq!(json["details"]["violation"]["observed"], 12_582_913);
    }
}
