//! Small, deterministic OPC packages for XML-part fuzzing. ZIP's stored method
//! keeps the mutated XML intact; ZipWriter computes both CRC-32 records.
use std::io::{Cursor, Write};

const DOCX_MAIN_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const XLSX_MAIN_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const PPTX_MAIN_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DRAW_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRES_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn stored_zip(parts: &[(&str, &[u8])], main_path: &str, main_type: &str) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(Cursor::new(&mut output));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/{main_path}" ContentType="{main_type}"/></Types>"#
        );
        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();
        let root_rels = format!(
            r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rOffice" Type="{OFFICE_REL}/officeDocument" Target="{main_path}"/></Relationships>"#
        );
        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();
        for &(name, body) in parts {
            zip.start_file(name, options).unwrap();
            zip.write_all(body).unwrap();
        }
        zip.finish().unwrap();
    }
    output
}

fn replace<'a>(parts: &mut [(&'a str, &'a [u8])], selected: &str, xml: &'a [u8]) {
    let part = parts
        .iter_mut()
        .find(|(name, _)| *name == selected)
        .unwrap();
    part.1 = xml;
}

/// The first byte selects the part; all remaining bytes are the exact part body.
pub fn docx_part(data: &[u8]) -> Option<Vec<u8>> {
    let (&selector, xml) = data.split_first()?;
    let paths = [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/header1.xml",
        "word/footer1.xml",
        "word/footnotes.xml",
    ];
    let document = format!(
        r#"<w:document xmlns:w="{DOC_NS}" xmlns:r="{OFFICE_REL}"><w:body><w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:footerReference w:type="default" r:id="rFooter"/></w:sectPr></w:body></w:document>"#
    );
    let rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rStyles" Type="{OFFICE_REL}/styles" Target="styles.xml"/><Relationship Id="rNumbering" Type="{OFFICE_REL}/numbering" Target="numbering.xml"/><Relationship Id="rHeader" Type="{OFFICE_REL}/header" Target="header1.xml"/><Relationship Id="rFooter" Type="{OFFICE_REL}/footer" Target="footer1.xml"/><Relationship Id="rFootnotes" Type="{OFFICE_REL}/footnotes" Target="footnotes.xml"/></Relationships>"#
    );
    let styles = format!(r#"<w:styles xmlns:w="{DOC_NS}"/>"#);
    let numbering = format!(r#"<w:numbering xmlns:w="{DOC_NS}"/>"#);
    let header = format!(r#"<w:hdr xmlns:w="{DOC_NS}"><w:p/></w:hdr>"#);
    let footer = format!(r#"<w:ftr xmlns:w="{DOC_NS}"><w:p/></w:ftr>"#);
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="{DOC_NS}"><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>"#
    );
    let mut parts: Vec<(&str, &[u8])> = vec![
        (paths[0], document.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        (paths[1], styles.as_bytes()),
        (paths[2], numbering.as_bytes()),
        (paths[3], header.as_bytes()),
        (paths[4], footer.as_bytes()),
        (paths[5], footnotes.as_bytes()),
    ];
    replace(&mut parts, paths[selector as usize % paths.len()], xml);
    Some(stored_zip(&parts, paths[0], DOCX_MAIN_TYPE))
}

pub fn xlsx_part(data: &[u8]) -> Option<Vec<u8>> {
    let (&selector, xml) = data.split_first()?;
    let paths = [
        "xl/worksheets/sheet1.xml",
        "xl/sharedStrings.xml",
        "xl/styles.xml",
        "xl/workbook.xml",
    ];
    let workbook = format!(
        r#"<workbook xmlns="{SHEET_NS}" xmlns:r="{OFFICE_REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rSheet"/></sheets></workbook>"#
    );
    // Additional worksheet IDs let a public multi-sheet workbook seed retain
    // its original XML while resolving every sheet to this single fixed part.
    let aliases = (1..=5)
        .map(|id| format!(r#"<Relationship Id="rId{id}" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>"#))
        .collect::<String>();
    let rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rSheet" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>{aliases}<Relationship Id="rStrings" Type="{OFFICE_REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rStyles" Type="{OFFICE_REL}/styles" Target="styles.xml"/></Relationships>"#
    );
    let sheet = format!(
        r#"<worksheet xmlns="{SHEET_NS}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#
    );
    let strings =
        format!(r#"<sst xmlns="{SHEET_NS}" count="1" uniqueCount="1"><si><t>seed</t></si></sst>"#);
    let styles = format!(
        r#"<styleSheet xmlns="{SHEET_NS}"><fonts count="1"><font/></fonts><fills count="0"/><borders count="0"/><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>"#
    );
    let mut parts: Vec<(&str, &[u8])> = vec![
        (paths[3], workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
        (paths[0], sheet.as_bytes()),
        (paths[1], strings.as_bytes()),
        (paths[2], styles.as_bytes()),
    ];
    replace(&mut parts, paths[selector as usize % paths.len()], xml);
    Some(stored_zip(&parts, paths[3], XLSX_MAIN_TYPE))
}

fn pptx_package(selected: &str, xml: &[u8], chart_ex: Option<bool>) -> Vec<u8> {
    let chart_frame = match chart_ex {
        Some(true) => r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart r:id="rChart"/></a:graphicData></a:graphic></p:graphicFrame>"#.to_owned(),
        Some(false) => r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></p:graphicFrame>"#.to_owned(),
        None => String::new(),
    };
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PRES_NS}" xmlns:r="{OFFICE_REL}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rMaster"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rSlide"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#
    );
    let presentation_rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rSlide" Type="{OFFICE_REL}/slide" Target="slides/slide1.xml"/><Relationship Id="rMaster" Type="{OFFICE_REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rTheme" Type="{OFFICE_REL}/theme" Target="theme/theme1.xml"/></Relationships>"#
    );
    let slide = format!(
        r#"<p:sld xmlns:p="{PRES_NS}" xmlns:a="{DRAW_NS}" xmlns:r="{OFFICE_REL}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><p:cSld><p:spTree>{chart_frame}</p:spTree></p:cSld></p:sld>"#
    );
    let chart_rel = if chart_ex == Some(true) {
        r#"<Relationship Id="rChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/>"#.to_owned()
    } else if chart_ex == Some(false) {
        format!(
            r#"<Relationship Id="rChart" Type="{OFFICE_REL}/chart" Target="../charts/chart1.xml"/>"#
        )
    } else {
        String::new()
    };
    let slide_rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rLayout" Type="{OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>{chart_rel}</Relationships>"#
    );
    let layout = format!(
        r#"<p:sldLayout xmlns:p="{PRES_NS}" xmlns:a="{DRAW_NS}"><p:cSld><p:spTree/></p:cSld></p:sldLayout>"#
    );
    let layout_rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rMaster" Type="{OFFICE_REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#
    );
    let master = format!(
        r#"<p:sldMaster xmlns:p="{PRES_NS}" xmlns:a="{DRAW_NS}"><p:cSld><p:spTree/></p:cSld></p:sldMaster>"#
    );
    let master_rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rTheme" Type="{OFFICE_REL}/theme" Target="../theme/theme1.xml"/></Relationships>"#
    );
    let theme = format!(
        r#"<a:theme xmlns:a="{DRAW_NS}" name="seed"><a:themeElements><a:clrScheme name="seed"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme><a:fontScheme name="seed"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="seed"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#
    );
    let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
    let chartex = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData/><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>"#;
    let mut parts: Vec<(&str, &[u8])> = vec![
        ("ppt/presentation.xml", presentation.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/slide1.xml", slide.as_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes()),
        ("ppt/slideLayouts/slideLayout1.xml", layout.as_bytes()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            layout_rels.as_bytes(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", master.as_bytes()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            master_rels.as_bytes(),
        ),
        ("ppt/theme/theme1.xml", theme.as_bytes()),
    ];
    if chart_ex.is_some() {
        parts.push(("ppt/charts/chart1.xml", chart.as_bytes()));
        parts.push(("ppt/charts/chartEx1.xml", chartex.as_bytes()));
    }
    replace(&mut parts, selected, xml);
    stored_zip(&parts, "ppt/presentation.xml", PPTX_MAIN_TYPE)
}

pub fn pptx_part(data: &[u8]) -> Option<Vec<u8>> {
    let (&selector, xml) = data.split_first()?;
    let paths = [
        "ppt/slides/slide1.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/theme/theme1.xml",
    ];
    Some(pptx_package(
        paths[selector as usize % paths.len()],
        xml,
        None,
    ))
}

pub fn chart_part(data: &[u8]) -> Option<Vec<u8>> {
    let (&selector, xml) = data.split_first()?;
    let chart_ex = selector % 2 == 1;
    let path = if chart_ex {
        "ppt/charts/chartEx1.xml"
    } else {
        "ppt/charts/chart1.xml"
    };
    Some(pptx_package(path, xml, Some(chart_ex)))
}
