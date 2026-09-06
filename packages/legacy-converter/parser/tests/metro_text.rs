//! Reader foundation only: not a binary-to-XML correspondence oracle.
#[path = "../src/officeart/metro_text.rs"]
mod metro_text;
use metro_text::{read, Budget};

fn budget() -> Budget {
    Budget {
        bytes: 2 * 1024 * 1024,
        events: 10000,
        paragraphs: 1000,
    }
}
fn shape(paragraphs: &str) -> String {
    format!(
        r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txBody>{paragraphs}</p:txBody></p:sp>"#
    )
}
#[test]
fn retains_exact_typed_offsets_and_literal_placeholders_without_matching() {
    let xml = shape(
        r#"<a:p><a:pPr marL="1587" marR="0" indent="-228600" defTabSz="-1" lvl="8"/><a:r><a:t>    _  </a:t></a:r></a:p><a:p/>"#,
    );
    let result = read(xml.as_bytes(), &mut budget()).unwrap().unwrap();
    assert_eq!(result.len(), 2);
    assert_eq!(
        (
            result[0].margin_left,
            result[0].margin_right,
            result[0].indent,
            result[0].default_tab_size,
            result[0].level
        ),
        (Some(1587), Some(0), Some(-228600), Some(-1), Some(8))
    );
    assert_eq!(result[0].literal, "    _  ");
    assert_eq!(result[1].margin_left, None);
}
#[test]
fn namespaces_and_direct_ownership_prevent_foreign_property_injection() {
    let xml = shape(
        r#"<a:p xmlns:x="urn:foreign"><x:pPr marL="999"/><x:wrapper><a:pPr marL="888"/></x:wrapper><a:pPr x:indent="777" marL="0"/><a:r><a:t>ok</a:t></a:r></a:p>"#,
    );
    let result = read(xml.as_bytes(), &mut budget()).unwrap().unwrap();
    assert_eq!(result[0].margin_left, Some(0));
    assert_eq!(result[0].indent, None);
    let renamed = xml.replace("a:", "d:").replace("xmlns:a=", "xmlns:d=");
    assert_eq!(
        read(renamed.as_bytes(), &mut budget()).unwrap().unwrap(),
        result
    );
}
#[test]
fn decodes_literals_entities_cdata_and_surrogates_without_substitution() {
    let xml = shape("<a:p><a:r><a:t>A&amp;&#x1F600;<![CDATA[<日本語>]]></a:t></a:r></a:p>");
    assert_eq!(
        read(xml.as_bytes(), &mut budget()).unwrap().unwrap()[0].literal,
        "A&😀<日本語>"
    );
}
#[test]
fn unsupported_shape_roots_fields_and_breaks_are_not_projected() {
    for xml in [
        shape("<a:p><a:fld/></a:p>"),
        shape("<a:p><a:br/></a:p>"),
        "<p:grpSp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>".into(),
    ] {
        assert!(read(xml.as_bytes(), &mut budget()).unwrap().is_none());
    }
}
#[test]
fn rejects_invalid_ranges_duplicates_unbound_namespaces_and_malformed_xml() {
    for attrs in [
        "marL=\"-1\"",
        "marR=\"51206401\"",
        "indent=\"-51206401\"",
        "defTabSz=\"2147483648\"",
        "lvl=\"9\"",
        "indent=\"NaN\"",
        "marL=\"1\" marL=\"2\"",
    ] {
        assert!(read(
            shape(&format!("<a:p><a:pPr {attrs}/></a:p>")).as_bytes(),
            &mut budget()
        )
        .is_err());
    }
    for body in ["<a:p><a:pPr/><a:pPr/></a:p>", "<a:p><x:t/></a:p>", "<a:p>"] {
        assert!(read(shape(body).as_bytes(), &mut budget()).is_err());
    }
}
#[test]
fn rejects_dtd_entities_processing_instructions_and_multiple_roots() {
    for xml in [
        format!(
            "<!DOCTYPE sp [<!ENTITY x SYSTEM 'file:///secret'>]>{}",
            shape("<a:p/>")
        ),
        shape("<a:p><a:r><a:t>&unknown;</a:t></a:r></a:p>"),
        format!("<?execute data?>{}", shape("<a:p/>")),
        format!("{}{}", shape("<a:p/>"), shape("<a:p/>")),
    ] {
        assert!(read(xml.as_bytes(), &mut budget()).is_err());
    }
}
#[test]
fn aggregate_budgets_survive_calls_and_depth_and_part_size_are_bounded() {
    let xml = shape("<a:p/>");
    let mut b = budget();
    b.paragraphs = 1;
    assert!(read(xml.as_bytes(), &mut b).is_ok());
    assert!(read(xml.as_bytes(), &mut b).is_err());
    let mut b = budget();
    b.bytes = xml.len();
    assert!(read(xml.as_bytes(), &mut b).is_ok());
    assert!(read(xml.as_bytes(), &mut b).is_err());
    let mut b = budget();
    b.events = 1;
    assert!(read(xml.as_bytes(), &mut b).is_err());
    assert!(read(
        shape(&format!("{}{}", "<a:x>".repeat(70), "</a:x>".repeat(70))).as_bytes(),
        &mut budget()
    )
    .is_err());
    assert!(read(&vec![b' '; 1024 * 1024 + 1], &mut budget()).is_err());
}

#[test]
fn xml_version_encoding_characters_and_attribute_work_are_checked() {
    for prefix in [
        "<?xml version=\"1.1\"?>",
        "<?xml version=\"1.0\" encoding=\"UTF-16\"?>",
        "<?xml version=\"1.0\"?><?xml version=\"1.0\"?>",
    ] {
        assert!(read(
            format!("{prefix}{}", shape("<a:p/>")).as_bytes(),
            &mut budget()
        )
        .is_err());
    }
    for literal in ["\u{0}", "&#0;", "&#xFFFE;"] {
        assert!(read(
            shape(&format!("<a:p><a:r><a:t>{literal}</a:t></a:r></a:p>")).as_bytes(),
            &mut budget()
        )
        .is_err());
    }
    let attributes = (0..65)
        .map(|i| format!(" attr{i}=\"0\""))
        .collect::<String>();
    assert!(read(
        shape(&format!("<a:p><a:pPr{attributes}/></a:p>")).as_bytes(),
        &mut budget()
    )
    .is_err());
}

#[test]
fn legal_scalar_boundaries_and_explicit_zero_remain_exact() {
    for (key, min, max) in [
        ("marL", 0, 51_206_400),
        ("marR", 0, 51_206_400),
        ("indent", -51_206_400, 51_206_400),
        ("defTabSz", i32::MIN, i32::MAX),
        ("lvl", 0, 8),
    ] {
        for value in [min, 0, max] {
            let xml = shape(&format!("<a:p><a:pPr {key}=\"{value}\"/></a:p>"));
            let p = read(xml.as_bytes(), &mut budget())
                .unwrap()
                .unwrap()
                .remove(0);
            let actual = match key {
                "marL" => p.margin_left,
                "marR" => p.margin_right,
                "indent" => p.indent,
                "defTabSz" => p.default_tab_size,
                _ => p.level.map(i32::from),
            };
            assert_eq!(actual, Some(value));
        }
    }
}
