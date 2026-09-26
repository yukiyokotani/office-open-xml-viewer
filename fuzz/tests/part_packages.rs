use std::io::{Cursor, Read};

fn public_part(format: &str, path: &str) -> Vec<u8> {
    let source = format!("../packages/{format}/public/demo/sample-1.{format}");
    let bytes = std::fs::read(source).expect("tracked public fixture");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut xml = Vec::new();
    archive
        .by_name(path)
        .unwrap()
        .read_to_end(&mut xml)
        .unwrap();
    xml
}

fn selected(selector: u8, xml: &[u8]) -> Vec<u8> {
    let mut input = vec![selector];
    input.extend_from_slice(xml);
    input
}

fn stored_part(package: &[u8], name: &str, xml: &[u8]) {
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).unwrap();
    let main = if name.starts_with("word/") {
        "word/document.xml"
    } else if name.starts_with("xl/") {
        "xl/workbook.xml"
    } else {
        "ppt/presentation.xml"
    };
    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")
        .unwrap()
        .read_to_string(&mut content_types)
        .unwrap();
    assert!(content_types.contains(&format!("PartName=\"/{main}\"")));
    let mut root_rels = String::new();
    archive
        .by_name("_rels/.rels")
        .unwrap()
        .read_to_string(&mut root_rels)
        .unwrap();
    assert!(root_rels.contains(&format!("Target=\"{main}\"")));
    let mut entry = archive.by_name(name).unwrap();
    assert_eq!(entry.compression(), zip::CompressionMethod::Stored);
    let mut actual = Vec::new();
    entry.read_to_end(&mut actual).unwrap(); // also validates CRC-32
    assert_eq!(actual, xml);
}

#[test]
fn public_xml_parts_reach_package_entry_points() {
    let docx = [
        (0, "word/document.xml"),
        (1, "word/styles.xml"),
        (2, "word/numbering.xml"),
        (5, "word/footnotes.xml"),
    ];
    for (selector, path) in docx {
        let xml = public_part("docx", path);
        let package = ooxml_fuzz::docx_part(&selected(selector, &xml)).unwrap();
        stored_part(&package, path, &xml);
        docx_parser::parse_docx_native(&package).expect(path);
    }

    let xlsx = [
        (0, "xl/worksheets/sheet1.xml"),
        (1, "xl/sharedStrings.xml"),
        (2, "xl/styles.xml"),
        (3, "xl/workbook.xml"),
    ];
    for (selector, path) in xlsx {
        let xml = public_part("xlsx", path);
        let package = ooxml_fuzz::xlsx_part(&selected(selector, &xml)).unwrap();
        stored_part(&package, path, &xml);
        xlsx_parser::parse_workbook_native(&package).expect(path);
        xlsx_parser::parse_sheet_native(&package, 0, "Sheet1").expect(path);
    }

    let pptx = [
        (0, "ppt/slides/slide1.xml"),
        (1, "ppt/slideLayouts/slideLayout1.xml"),
        (2, "ppt/slideMasters/slideMaster1.xml"),
        (3, "ppt/theme/theme1.xml"),
    ];
    for (selector, path) in pptx {
        let xml = public_part("pptx", path);
        let package = ooxml_fuzz::pptx_part(&selected(selector, &xml)).unwrap();
        stored_part(&package, path, &xml);
        pptx_parser::parse_pptx_native(&package).expect(path);
    }

    let chart = public_part("pptx", "ppt/charts/chart1.xml");
    let package = ooxml_fuzz::chart_part(&selected(0, &chart)).unwrap();
    stored_part(&package, "ppt/charts/chart1.xml", &chart);
    let model = pptx_parser::parse_pptx_native(&package).expect("public chart");
    assert!(
        model.contains("\"type\":\"chart\""),
        "chart part was not parsed"
    );

    let chartex = br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;
    let package = ooxml_fuzz::chart_part(&selected(1, chartex)).unwrap();
    stored_part(&package, "ppt/charts/chartEx1.xml", chartex);
    let model = pptx_parser::parse_pptx_native(&package).expect("synthetic chartEx");
    assert!(
        model.contains("\"type\":\"chart\""),
        "chartEx part was not parsed"
    );
}

#[test]
fn optional_docx_stories_are_reached() {
    for (selector, root, marker) in [(3, "hdr", "HeaderSeed"), (4, "ftr", "FooterSeed")] {
        let xml = format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{marker}</w:t></w:r></w:p></w:{root}>"#
        );
        let package = ooxml_fuzz::docx_part(&selected(selector, xml.as_bytes())).unwrap();
        let model = docx_parser::parse_docx_native(&package).unwrap();
        assert!(model.contains(marker), "{root} did not reach the model");
    }

    let xml = br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>FootnoteSeed</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    let package = ooxml_fuzz::docx_part(&selected(5, xml)).unwrap();
    let model = docx_parser::parse_docx_native(&package).unwrap();
    assert!(
        model.contains("FootnoteSeed"),
        "footnote did not reach the model"
    );
}
