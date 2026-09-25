//! PowerPoint Binary File (`.ppt`) compatibility subset.
//!
//! The direct session resolves live slides and outline text through the
//! edit-chain persist directory and projects each live slide reference
//! straight into the presentation renderer model. It never executes actions,
//! macros, hyperlinks, animations, or external-resource updates. See [MS-PPT]
//! 2.3.3, 2.4.3, and 2.9.40 through 2.9.42. Unsupported/encrypted containers
//! fail closed.

use crate::cfb::CompoundFile;

pub(crate) mod direct_cursor;
pub(crate) mod direct_session;
mod drawing;
mod media;
mod metro;
mod paint;
mod persist;
mod ruler;
mod scheme;
mod shape_master;
mod text_style;

// MS-PPT master unit = 1/576 inch = 1587.5 EMU; round signed coordinates.
fn master_to_emu(value: i64) -> i64 {
    (value * 3175 + value.signum()) / 2
}

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
// Implementation resource policy, independent of the compressed ZIP ceiling.
const MAX_TEXT_BYTES: usize = 128 * 1024 * 1024;

use crate::officeart::{
    record_span_with_end, record_span_with_end_in, ByteSpan, Record, RecordSpan,
};

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
    let parent = ByteSpan::new(0..bytes.len(), bytes.len(), "PowerPoint")?;
    visit_record_spans(bytes, &parent, budget, |span| {
        records.push(span.view(bytes)?);
        Ok(())
    })?;
    Ok(records)
}

fn parse_record_spans(
    bytes: &[u8],
    parent: &ByteSpan,
    budget: &mut usize,
) -> Result<Vec<RecordSpan>, String> {
    let mut records = Vec::new();
    visit_record_spans(bytes, parent, budget, |span| {
        records.push(span);
        Ok(())
    })?;
    Ok(records)
}

fn visit_record_spans(
    bytes: &[u8],
    parent: &ByteSpan,
    budget: &mut usize,
    mut visitor: impl FnMut(RecordSpan) -> Result<(), String>,
) -> Result<(), String> {
    // Validate the complete parent before using its absolute range for slices.
    parent.view(bytes)?;
    let range = parent.range();
    let mut offset = range.start;
    while offset < range.end {
        if range.end - offset < 8 {
            return if bytes[offset..range.end].iter().all(|byte| *byte == 0) {
                Ok(())
            } else {
                Err(unsupported("truncated PowerPoint record header"))
            };
        }
        let options = u16_at(bytes, offset)?;
        let kind = u16_at(bytes, offset + 2)?;
        let size = u32_at(bytes, offset + 4)?;
        if options == 0 && kind == 0 && size == 0 {
            return if bytes[offset..range.end].iter().all(|byte| *byte == 0) {
                Ok(())
            } else {
                Err(unsupported("unexpected zero PowerPoint record"))
            };
        }
        let (record, end) = record_span_with_end_in(bytes, parent, offset, budget, "PowerPoint")?;
        visitor(record)?;
        offset = end;
    }
    Ok(())
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
    crate::officeart::record_with_end(bytes, offset, budget, "PowerPoint")
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

fn decode_text(record: Record<'_>) -> Result<String, String> {
    if record.kind == TEXT_CHARS_ATOM {
        if !record.payload.len().is_multiple_of(2) {
            return Err(unsupported("misaligned PowerPoint Unicode text atom"));
        }
        let units = record
            .payload
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        Ok(char::decode_utf16(units)
            .map(|c| c.unwrap_or('\u{fffd}'))
            .collect())
    } else {
        Ok(record
            .payload
            .iter()
            // MS-PPT 2.9.43: compressed UTF-16 low bytes, not Windows-1252.
            .map(|byte| char::from(*byte))
            .collect())
    }
}

fn charge_text(budget: &mut usize, bytes: usize) -> Result<(), String> {
    *budget = budget
        .checked_sub(bytes)
        .ok_or_else(|| unsupported("PowerPoint decoded text budget exceeded"))?;
    Ok(())
}

/// MS-PPT 2.5.1 / 2.6.6: only the live slide's own optional
/// SlideShowSlideInfoAtom controls visibility. Master and nested records do not.
fn slide_is_hidden(payload: &[u8], budget: &mut usize) -> Result<bool, String> {
    let mut hidden = None;
    for record in parse_records(payload, budget)? {
        if record.kind != 0x03f9 {
            continue;
        }
        if hidden.is_some()
            || record.version != 0
            || record.instance != 0
            || record.payload.len() != 16
        {
            return Err(unsupported(
                "invalid or duplicate PowerPoint SlideShowSlideInfoAtom",
            ));
        }
        // fHidden is bit 2 of the flags word following the two effect bytes.
        // Reserved bits are explicitly ignored; transitions/sounds/actions
        // remain outside this reader's subset and are never executed.
        hidden = Some(u16_at(record.payload, 10)? & 0x0004 != 0);
    }
    Ok(hidden.unwrap_or(false))
}

/// Fixed DrawingML theme part (Office 2007 default colours, Arial fonts), used
/// only by tests and the alternative-shape fuzz driver as a readable master
/// round-trip theme against which alternative shape XML resolves. Production
/// themes come from each main master's RoundTripTheme12Atom
/// (`metro::master_theme`).
#[cfg(any(test, feature = "fuzzing"))]
fn theme() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Legacy conversion"><a:themeElements><a:clrScheme name="Legacy"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Legacy"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Legacy"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#.into()
}

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

/// Fuzz driver for alternative shape XML resolution (`crate::fuzzing`):
/// `data` as a whole metroBlob package, and as the shape part of a package
/// whose relationships are well formed, against a fixed binary text shape.
#[cfg(feature = "fuzzing")]
pub(crate) fn fuzz_alternative(data: &[u8]) {
    use std::io::Write;
    let Ok(mut element) = serde_json::from_value::<pptx_model::ShapeElement>(serde_json::json!({
        "x": 0, "y": 0, "width": 1_587_500, "height": 793_750, "rotation": 0.0,
        "flipH": false, "flipV": false, "geometry": "rect",
        "fill": null, "stroke": null, "textBody": null,
        "defaultTextColor": null, "custGeom": null,
        "adj": null, "adj2": null, "adj3": null, "adj4": null,
        "adj5": null, "adj6": null, "adj7": null, "adj8": null,
        "id": "2", "name": null, "hyperlink": null,
        "placeholderType": null, "placeholderIdx": null
    })) else {
        unreachable!("fixed fuzz shape deserializes");
    };
    element.fill = Some(pptx_model::Fill::Solid {
        color: "FF0000".to_owned(),
    });
    let leaf = pptx_model::Transform {
        cx: element.width,
        cy: element.height,
        ..Default::default()
    };
    let binary = metro::BinaryShape {
        element: &element,
        leaf: &leaf,
        nested: false,
        text: Some("AB\u{b}C\rD"),
        fill: metro::RecordedFill::Stated(element.fill.clone().map(Box::new)),
        path_paint: None,
        adjust_bounds: [None; 8],
    };
    let theme = metro::Theme::Readable {
        theme_xml: theme(),
        clr_map: None,
    };
    let mut package = std::io::Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut package);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let parts: [(&str, &[u8]); 2] = [
            ("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/shapeXml" Target="drs/shapexml.xml"/></Relationships>"#),
            ("drs/shapexml.xml", data),
        ];
        for (name, bytes) in parts {
            if writer.start_file(name, options).is_err() || writer.write_all(bytes).is_err() {
                return;
            }
        }
        if writer.finish().is_err() {
            return;
        }
    }
    for blob in [data, package.get_ref().as_slice()] {
        let (mut work, mut text, mut model) = (64, 16 * 1024 * 1024, 64 * 1024 * 1024);
        let _ = metro::adopt(&binary, blob, &theme, &mut work, &mut text, &mut model);
    }
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
    #[cfg(feature = "fuzzing")]
    #[test]
    fn alternative_fuzz_driver_builds_its_fixed_shape() {
        super::fuzz_alternative(b"");
        super::fuzz_alternative(
            br#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        );
    }

    #[test]
    fn owned_and_borrowed_scans_preserve_padding_and_record_budgets() {
        for padding in 0..=16 {
            let mut bytes = super::persist::tests::record(0, 7, &[1, 2]);
            bytes.resize(bytes.len() + padding, 0);
            let bounds = super::ByteSpan::new(0..bytes.len(), bytes.len(), "test").unwrap();
            let mut borrowed_budget = 1;
            let borrowed = super::parse_records(&bytes, &mut borrowed_budget).unwrap();
            let mut owned_budget = 1;
            let owned = super::parse_record_spans(&bytes, &bounds, &mut owned_budget).unwrap();
            assert_eq!((borrowed.len(), owned.len()), (1, 1));
            assert_eq!((borrowed_budget, owned_budget), (0, 0));
            assert_eq!(owned[0].view(&bytes).unwrap().payload, &[1, 2]);
            assert!(super::parse_records(&bytes, &mut 0).is_err());
            assert!(super::parse_record_spans(&bytes, &bounds, &mut 0).is_err());
            if padding != 0 {
                *bytes.last_mut().unwrap() = 1;
                assert!(super::parse_records(&bytes, &mut 1).is_err());
                assert!(super::parse_record_spans(&bytes, &bounds, &mut 1).is_err());
            }
        }
        // An all-zero tail is padding, not a charged record, even at zero credit.
        assert!(super::parse_records(&[0; 16], &mut 0).unwrap().is_empty());
    }

    #[test]
    fn span_scan_validates_backing_before_reading_headers_or_padding() {
        let bounds = super::ByteSpan::new(3..16, 16, "test").unwrap();
        for length in 0..16 {
            let bytes = vec![0; length];
            assert!(super::parse_record_spans(&bytes, &bounds, &mut 1).is_err());
        }
        let bytes = vec![0; 16];
        assert!(super::parse_record_spans(&bytes, &bounds, &mut 0)
            .unwrap()
            .is_empty());
    }

    use super::{parse_current_user_atom, parse_records, MAX_RECORDS};

    #[test]
    fn slide_visibility_uses_only_the_hidden_flag_in_all_flag_words() {
        for flags in 0..=u16::MAX {
            let mut payload = [0; 16];
            payload[10..12].copy_from_slice(&flags.to_le_bytes());
            let record = super::persist::tests::record(0, 0x03f9, &payload);
            assert_eq!(
                super::slide_is_hidden(&record, &mut MAX_RECORDS.clone()).unwrap(),
                flags & 4 != 0,
                "flags={flags:04x}"
            );
        }
    }

    #[test]
    fn slide_visibility_scan_respects_record_budget() {
        let record = super::persist::tests::record(0, 0x03f9, &[0; 16]);
        assert!(super::slide_is_hidden(&record, &mut 0).is_err());
        let mut budget = 1;
        assert!(!super::slide_is_hidden(&record, &mut budget).unwrap());
        assert_eq!(budget, 0);
    }

    #[test]
    fn text_bytes_are_zero_high_byte_unicode_not_ansi() {
        let text = super::decode_text(super::Record {
            version: 0,
            instance: 0,
            kind: 4008,
            payload: &[0x80, 0x91, 0xe9],
        })
        .unwrap();
        assert_eq!(text, "\u{80}\u{91}é");
    }

    #[test]
    fn rejects_record_offsets_before_doing_pointer_arithmetic() {
        assert!(super::parse_record_at(&[], usize::MAX, &mut MAX_RECORDS.clone()).is_err());
    }

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
    fn rejects_zero_records_that_hide_nonzero_trailing_data() {
        let mut budget = MAX_RECORDS;
        assert!(parse_records(&[0, 0, 0, 0, 0, 0, 0, 0, 1], &mut budget).is_err());
        let mut budget = MAX_RECORDS;
        assert!(parse_records(&[0; 9], &mut budget).unwrap().is_empty());
    }
}
