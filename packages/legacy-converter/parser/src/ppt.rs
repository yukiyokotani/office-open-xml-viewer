//! PowerPoint Binary File (`.ppt`) compatibility subset.
//!
//! The first converter milestone extracts text atoms from single-edit
//! PowerPoint 97-2003 slide containers. It emits one macro-free PresentationML
//! slide per binary Slide record and never executes actions, macros, hyperlinks,
//! animations, or external-resource updates. See [MS-PPT] 2.3.3, 2.4.3, and
//! 2.9.40 through 2.9.42. Unsupported/encrypted containers fail closed.

use crate::cfb::CompoundFile;
use crate::ooxml::{write_package, xml_text, ROOT_RELS_PPTX};

const DOCUMENT_CONTAINER: u16 = 1000;
const SLIDE_CONTAINER: u16 = 1006;
const TEXT_CHARS_ATOM: u16 = 4000;
const TEXT_BYTES_ATOM: u16 = 4008;
const USER_EDIT_ATOM: u16 = 0x0ff5;
const CURRENT_USER_ATOM: u16 = 0x0ff6;
const DOCUMENT_ENCRYPTION_ATOM: u16 = 12052;
const CURRENT_USER_NOT_ENCRYPTED: u32 = 0xe391_c05f;
const CURRENT_USER_ENCRYPTED: u32 = 0xf3d1_c4df;
const MAX_RECORDS: usize = 2_000_000;
const MAX_DEPTH: usize = 64;
const MAX_SLIDES: usize = 100_000;
const MAX_TEXT_BLOCKS_PER_SLIDE: usize = 100_000;

pub struct PptConversion {
    pub bytes: Vec<u8>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Copy)]
struct Record<'a> {
    version: u8,
    kind: u16,
    payload: &'a [u8],
}

pub fn convert(cfb: &CompoundFile<'_>, max_output_bytes: usize) -> Result<PptConversion, String> {
    if cfb.has_entry("EncryptedSummary") {
        return Err(unsupported(
            "encrypted PowerPoint binary documents are not supported",
        ));
    }
    let document = cfb.stream("PowerPoint Document").map_err(unsupported)?;
    let mut record_budget = MAX_RECORDS;
    validate_single_edit_generation(cfb, &document, &mut record_budget)?;
    if contains_record(&document, DOCUMENT_ENCRYPTION_ATOM, 0, &mut record_budget)? {
        return Err(unsupported(
            "encrypted PowerPoint binary documents are not supported",
        ));
    }
    let top_level = parse_records(&document, &mut record_budget)?;
    if !top_level
        .iter()
        .any(|record| record.kind == DOCUMENT_CONTAINER)
    {
        return Err(unsupported("missing PowerPoint Document container"));
    }
    let mut slides = Vec::new();
    for record in &top_level {
        if record.kind != SLIDE_CONTAINER || record.version != 0x0f {
            continue;
        }
        if slides.len() >= MAX_SLIDES {
            return Err(unsupported("too many PowerPoint slides"));
        }
        let mut text = Vec::new();
        collect_text(record.payload, 0, &mut record_budget, &mut text)?;
        slides.push(text);
    }
    if slides.is_empty() {
        return Err(unsupported(
            "no directly addressable PowerPoint slide records were found; multi-edit persist maps are not supported yet",
        ));
    }
    let bytes = build_pptx(&slides, max_output_bytes)?;
    Ok(PptConversion {
        bytes,
        warnings: vec![
            "legacy-ppt:slide-text-only".into(),
            "legacy-ppt:slides-emitted-in-stream-order".into(),
            "legacy-ppt:masters-formatting-media-and-actions-omitted".into(),
        ],
    })
}

fn validate_single_edit_generation(
    cfb: &CompoundFile<'_>,
    document: &[u8],
    budget: &mut usize,
) -> Result<(), String> {
    let current_user = cfb.stream("Current User").map_err(unsupported)?;
    let current_edit = parse_current_user_atom(&current_user, budget)?;
    let user_edit = parse_record_at(document, current_edit, budget)?;
    if user_edit.kind != USER_EDIT_ATOM || user_edit.version != 0 {
        return Err(unsupported(
            "PowerPoint current edit does not reference a UserEditAtom",
        ));
    }
    if u32_at(user_edit.payload, 8)? != 0 {
        return Err(unsupported(
            "multiple PowerPoint edit generations are not supported yet",
        ));
    }
    Ok(())
}

/// Parse the sole CurrentUserAtom from the Current User stream.
///
/// This record is handled separately from the generic record walker because
/// PowerPoint for Mac writes an empty-user-name atom with `recLen = 0x1c` but
/// can omit the final four zero bytes when the CFB stream is stored in the mini
/// stream. All normative fields through `relVersion` are still present. Keep
/// that compatibility allowance exact and bounded; other truncated records
/// remain errors. See [MS-PPT] 2.3.2.
fn parse_current_user_atom(bytes: &[u8], budget: &mut usize) -> Result<usize, String> {
    if *budget == 0 {
        return Err(unsupported("too many PowerPoint records"));
    }
    *budget -= 1;
    if bytes.len() < 8 {
        return Err(unsupported("missing PowerPoint CurrentUserAtom"));
    }
    let options = u16_at(bytes, 0)?;
    if options != 0 || u16_at(bytes, 2)? != CURRENT_USER_ATOM {
        return Err(unsupported("invalid PowerPoint Current User stream"));
    }
    let declared_len = usize::try_from(u32_at(bytes, 4)?)
        .map_err(|_| unsupported("PowerPoint CurrentUserAtom is too large"))?;
    let payload = &bytes[8..];
    if u32_at(payload, 0)? != 0x14 {
        return Err(unsupported("invalid PowerPoint CurrentUserAtom size"));
    }
    match u32_at(payload, 4)? {
        CURRENT_USER_NOT_ENCRYPTED => {}
        CURRENT_USER_ENCRYPTED => {
            return Err(unsupported(
                "encrypted PowerPoint binary documents are not supported",
            ));
        }
        _ => return Err(unsupported("invalid PowerPoint CurrentUserAtom token")),
    }

    let user_name_len = u16_at(payload, 12)? as usize;
    if user_name_len > 255 {
        return Err(unsupported("PowerPoint user name is too long"));
    }
    if u16_at(payload, 14)? != 0x03f4 || payload.get(16..18) != Some(&[3, 0]) {
        return Err(unsupported("invalid PowerPoint CurrentUserAtom version"));
    }
    let required_len = 24usize
        .checked_add(user_name_len)
        .ok_or_else(|| unsupported("PowerPoint CurrentUserAtom size overflow"))?;
    if payload.len() < required_len {
        return Err(unsupported("truncated PowerPoint CurrentUserAtom"));
    }
    let rel_version = u32_at(payload, 20 + user_name_len)?;
    if !matches!(rel_version, 8 | 9) {
        return Err(unsupported("invalid PowerPoint release version"));
    }
    let with_unicode_len = required_len
        .checked_add(
            user_name_len
                .checked_mul(2)
                .ok_or_else(|| unsupported("PowerPoint user name size overflow"))?,
        )
        .ok_or_else(|| unsupported("PowerPoint CurrentUserAtom size overflow"))?;
    let standard_length = (declared_len == required_len || declared_len == with_unicode_len)
        && payload.len() >= declared_len;
    let mac_empty_name_length = user_name_len == 0
        && declared_len == required_len + 4
        && (payload.len() == required_len
            || payload
                .get(required_len..declared_len)
                .is_some_and(|tail| tail.iter().all(|byte| *byte == 0)));
    if !standard_length && !mac_empty_name_length {
        return Err(unsupported("invalid PowerPoint CurrentUserAtom length"));
    }
    let consumed = declared_len.min(payload.len());
    if payload[consumed..].iter().any(|byte| *byte != 0) {
        return Err(unsupported(
            "unexpected data after PowerPoint CurrentUserAtom",
        ));
    }

    usize::try_from(u32_at(payload, 8)?)
        .map_err(|_| unsupported("PowerPoint current edit offset is too large"))
}

fn parse_records<'a>(bytes: &'a [u8], budget: &mut usize) -> Result<Vec<Record<'a>>, String> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        if bytes.len() - offset < 8 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(records)
            } else {
                Err(unsupported("truncated PowerPoint record header"))
            };
        }
        let options = u16_at(bytes, offset)?;
        let kind = u16_at(bytes, offset + 2)?;
        let size = u32_at(bytes, offset + 4)?;
        if options == 0 && kind == 0 && size == 0 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(records)
            } else {
                Err(unsupported("unexpected zero PowerPoint record"))
            };
        }
        let (record, end) = parse_record_with_end(bytes, offset, budget)?;
        records.push(record);
        offset = end;
    }
    Ok(records)
}

fn parse_record_at<'a>(
    bytes: &'a [u8],
    offset: usize,
    budget: &mut usize,
) -> Result<Record<'a>, String> {
    parse_record_with_end(bytes, offset, budget).map(|(record, _)| record)
}

fn parse_record_with_end<'a>(
    bytes: &'a [u8],
    offset: usize,
    budget: &mut usize,
) -> Result<(Record<'a>, usize), String> {
    if *budget == 0 {
        return Err(unsupported("too many PowerPoint records"));
    }
    *budget -= 1;
    let options = u16_at(bytes, offset)?;
    let kind = u16_at(bytes, offset + 2)?;
    let size = usize::try_from(u32_at(bytes, offset + 4)?)
        .map_err(|_| unsupported("PowerPoint record is too large"))?;
    let end = offset
        .checked_add(8 + size)
        .ok_or_else(|| unsupported("PowerPoint record range overflow"))?;
    let payload = bytes
        .get(offset + 8..end)
        .ok_or_else(|| {
            unsupported(format!(
                "truncated PowerPoint record at offset {offset}: declared {size} bytes with {} available",
                bytes.len().saturating_sub(offset + 8),
            ))
        })?;
    Ok((
        Record {
            version: (options & 0x000f) as u8,
            kind,
            payload,
        },
        end,
    ))
}

fn contains_record(
    bytes: &[u8],
    expected: u16,
    depth: usize,
    budget: &mut usize,
) -> Result<bool, String> {
    if depth > MAX_DEPTH {
        return Err(unsupported("PowerPoint record nesting is too deep"));
    }
    for record in parse_records(bytes, budget)? {
        if record.kind == expected {
            return Ok(true);
        }
        if record.version == 0x0f && contains_record(record.payload, expected, depth + 1, budget)? {
            return Ok(true);
        }
    }
    Ok(false)
}

fn collect_text(
    bytes: &[u8],
    depth: usize,
    budget: &mut usize,
    output: &mut Vec<String>,
) -> Result<(), String> {
    if depth > MAX_DEPTH {
        return Err(unsupported("PowerPoint record nesting is too deep"));
    }
    for record in parse_records(bytes, budget)? {
        match record.kind {
            TEXT_CHARS_ATOM => {
                if record.payload.len() % 2 != 0 {
                    return Err(unsupported("misaligned PowerPoint Unicode text atom"));
                }
                let units = record
                    .payload
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                let text: String = char::decode_utf16(units)
                    .map(|value| value.unwrap_or('\u{fffd}'))
                    .collect();
                push_text(output, text)?;
            }
            TEXT_BYTES_ATOM => {
                let text: String = record
                    .payload
                    .iter()
                    .map(|byte| decode_windows_1252(*byte))
                    .collect();
                push_text(output, text)?;
            }
            _ if record.version == 0x0f => {
                collect_text(record.payload, depth + 1, budget, output)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn push_text(output: &mut Vec<String>, text: String) -> Result<(), String> {
    if output.len() >= MAX_TEXT_BLOCKS_PER_SLIDE {
        return Err(unsupported("too many PowerPoint text atoms on one slide"));
    }
    if !text.is_empty() {
        output.push(text);
    }
    Ok(())
}

fn build_pptx(slides: &[Vec<String>], max_output_bytes: usize) -> Result<Vec<u8>, String> {
    let mut content_types = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#,
    );
    let mut presentation = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>"#,
    );
    let mut presentation_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#,
    );
    let mut parts = vec![
        ("_rels/.rels".into(), ROOT_RELS_PPTX.to_string()),
        ("ppt/slideMasters/slideMaster1.xml".into(), slide_master()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
            slide_master_rels(),
        ),
        ("ppt/slideLayouts/slideLayout1.xml".into(), slide_layout()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
            slide_layout_rels(),
        ),
        ("ppt/theme/theme1.xml".into(), theme()),
    ];
    for (index, slide) in slides.iter().enumerate() {
        let id = index + 1;
        content_types.push_str(&format!(
            "<Override PartName=\"/ppt/slides/slide{id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
        ));
        presentation.push_str(&format!(
            "<p:sldId id=\"{}\" r:id=\"rId{}\"/>",
            256 + index,
            id + 1
        ));
        presentation_rels.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{id}.xml\"/>",
            id + 1
        ));
        parts.push((format!("ppt/slides/slide{id}.xml"), slide_xml(slide)));
        parts.push((format!("ppt/slides/_rels/slide{id}.xml.rels"), slide_rels()));
    }
    content_types.push_str("</Types>");
    presentation.push_str("</p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>");
    presentation_rels.push_str("</Relationships>");
    parts.push(("[Content_Types].xml".into(), content_types));
    parts.push(("ppt/presentation.xml".into(), presentation));
    parts.push(("ppt/_rels/presentation.xml.rels".into(), presentation_rels));
    write_package(&parts, max_output_bytes)
}

fn slide_xml(blocks: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#,
    );
    if !blocks.is_empty() {
        xml.push_str("<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Legacy slide text\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"457200\" y=\"457200\"/><a:ext cx=\"8229600\" cy=\"5943600\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/>");
        for block in blocks {
            for line in block.split(['\r', '\n']) {
                xml.push_str("<a:p><a:r><a:rPr lang=\"ja-JP\" sz=\"1800\"/><a:t>");
                xml.push_str(&xml_text(line));
                xml.push_str("</a:t></a:r><a:endParaRPr lang=\"ja-JP\" sz=\"1800\"/></a:p>");
            }
        }
        xml.push_str("</p:txBody></p:sp>");
    }
    xml.push_str("</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
    xml
}

fn slide_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#.into()
}

fn slide_master() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>"#.into()
}

fn slide_master_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>"#.into()
}

fn slide_layout() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#.into()
}

fn slide_layout_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#.into()
}

fn theme() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Legacy conversion"><a:themeElements><a:clrScheme name="Legacy"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Legacy"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Legacy"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#.into()
}

fn decode_windows_1252(byte: u8) -> char {
    const SPECIAL: [char; 32] = [
        '\u{20ac}', '\u{0081}', '\u{201a}', '\u{0192}', '\u{201e}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02c6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008d}',
        '\u{017d}', '\u{008f}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201c}', '\u{201d}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02dc}', '\u{2122}', '\u{0161}', '\u{203a}',
        '\u{0153}', '\u{009d}', '\u{017e}', '\u{0178}',
    ];
    if (0x80..=0x9f).contains(&byte) {
        SPECIAL[(byte - 0x80) as usize]
    } else {
        char::from(byte)
    }
}

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, String> {
    let raw = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| unsupported("truncated PowerPoint integer"))?;
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let raw = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| unsupported("truncated PowerPoint integer"))?;
    Ok(u32::from_le_bytes(raw.try_into().expect("four-byte slice")))
}

#[cfg(test)]
mod tests {
    use super::{collect_text, parse_current_user_atom, parse_records, MAX_RECORDS};

    fn current_user_atom(declared_len: u32, payload: &[u8]) -> Vec<u8> {
        let mut record = Vec::new();
        record.extend_from_slice(&0u16.to_le_bytes());
        record.extend_from_slice(&0x0ff6u16.to_le_bytes());
        record.extend_from_slice(&declared_len.to_le_bytes());
        record.extend_from_slice(payload);
        record
    }

    fn empty_user_payload(current_edit: u32) -> Vec<u8> {
        let mut payload = vec![0; 24];
        payload[0..4].copy_from_slice(&0x14u32.to_le_bytes());
        payload[4..8].copy_from_slice(&0xe391_c05fu32.to_le_bytes());
        payload[8..12].copy_from_slice(&current_edit.to_le_bytes());
        payload[14..16].copy_from_slice(&0x03f4u16.to_le_bytes());
        payload[16] = 3;
        payload[20..24].copy_from_slice(&8u32.to_le_bytes());
        payload
    }

    #[test]
    fn accepts_office_mac_empty_user_atom_without_declared_zero_tail() {
        let bytes = current_user_atom(28, &empty_user_payload(1234));
        let mut budget = MAX_RECORDS;
        assert_eq!(parse_current_user_atom(&bytes, &mut budget).unwrap(), 1234);
    }

    #[test]
    fn rejects_other_truncated_current_user_atoms() {
        let bytes = current_user_atom(29, &empty_user_payload(1234));
        let mut budget = MAX_RECORDS;
        assert!(parse_current_user_atom(&bytes, &mut budget).is_err());
    }

    #[test]
    fn extracts_unicode_text_from_nested_records() {
        let mut atom = Vec::new();
        atom.extend_from_slice(&0u16.to_le_bytes());
        atom.extend_from_slice(&4000u16.to_le_bytes());
        atom.extend_from_slice(&6u32.to_le_bytes());
        for unit in "日本語".encode_utf16() {
            atom.extend_from_slice(&unit.to_le_bytes());
        }
        let mut container = Vec::new();
        container.extend_from_slice(&0x000fu16.to_le_bytes());
        container.extend_from_slice(&1036u16.to_le_bytes());
        container.extend_from_slice(&(atom.len() as u32).to_le_bytes());
        container.extend_from_slice(&atom);
        let mut output = Vec::new();
        let mut budget = MAX_RECORDS;
        collect_text(&container, 0, &mut budget, &mut output).unwrap();
        assert_eq!(output, vec!["日本語"]);
    }

    #[test]
    fn rejects_zero_records_that_hide_nonzero_trailing_data() {
        let mut budget = MAX_RECORDS;
        assert!(parse_records(&[0, 0, 0, 0, 0, 0, 0, 0, 1], &mut budget).is_err());
        let mut budget = MAX_RECORDS;
        assert!(parse_records(&[0; 9], &mut budget).unwrap().is_empty());
    }
}
