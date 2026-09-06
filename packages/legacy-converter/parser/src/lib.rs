//! Purpose-built, minimum-sufficient Office 97-2003 to OOXML converter.
//!
//! This crate is intentionally independent of the OOXML renderers. It reads an
//! untrusted CFB container, extracts a documented passive subset, and creates a
//! new macro-free OOXML package. Unsupported versions and encryption fail closed.

use wasm_bindgen::prelude::*;

mod cfb;
mod doc;
mod officeart;
mod ooxml;
mod ppt;
mod xls;

pub const ENGINE_ID: &str = "silurus-legacy-office";
pub const ENGINE_VERSION: &str = env!("CARGO_PKG_VERSION");

/// Native development helper, not part of the production WASM contract.
/// Returns passive catalog images, not necessarily visible sheet objects.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_images(data: &[u8]) -> Result<Vec<(u32, &'static str, Vec<u8>)>, String> {
    xls::inspect_images(&inspection_source(data)?)
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub use xls::drawing_anchors::{CellCorner, DrawingAnchor, PictureReference};

/// Native-only raw sheet anchor evidence. Not a visibility or image-admission API.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_anchors(data: &[u8]) -> Result<Vec<DrawingAnchor>, String> {
    xls::inspect_anchors(&inspection_source(data)?)
}

/// Native development bindings. Each retained anchor references one supported
/// image here; groups, inherited visibility and layout remain uninterpreted.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub struct XlsPictureInspection {
    pub anchors: Vec<DrawingAnchor>,
    pub images: Vec<(u32, &'static str, Vec<u8>)>,
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_pictures(data: &[u8]) -> Result<XlsPictureInspection, String> {
    xls::inspect_pictures(&inspection_source(data)?)
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
fn inspection_source(data: &[u8]) -> Result<cfb::CompoundFile<'_>, String> {
    if data.len() > 256 * 1024 * 1024 {
        return Err("UNSUPPORTED:XLS inspection source byte budget exceeded".into());
    }
    cfb::CompoundFile::open(data).map_err(|e| format!("UNSUPPORTED:{e}"))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormat {
    Doc,
    Xls,
    Ppt,
}

impl LegacyFormat {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "doc" => Ok(Self::Doc),
            "xls" => Ok(Self::Xls),
            "ppt" => Ok(Self::Ppt),
            _ => Err("UNSUPPORTED:unknown legacy Office format".into()),
        }
    }
}

#[derive(Debug)]
pub struct NativeConversionOutput {
    pub bytes: Vec<u8>,
    pub warnings: Vec<String>,
}

/// Wrap arbitrary bytes in a structurally valid CFB container so the fuzz
/// target reaches the format readers as well as the outer container reader.
/// This helper is absent from production WASM builds.
#[cfg(feature = "fuzzing")]
pub fn build_fuzz_input(data: &[u8], format: LegacyFormat) -> Vec<u8> {
    match format {
        LegacyFormat::Doc => cfb::test_support::build_cfb(&[
            ("WordDocument", data.to_vec()),
            ("0Table", data.to_vec()),
        ]),
        LegacyFormat::Xls => cfb::test_support::build_cfb(&[("Workbook", data.to_vec())]),
        LegacyFormat::Ppt => cfb::test_support::build_cfb(&[
            ("PowerPoint Document", data.to_vec()),
            ("Current User", data.to_vec()),
        ]),
    }
}

pub fn convert_native(
    data: &[u8],
    format: LegacyFormat,
    max_output_bytes: usize,
) -> Result<NativeConversionOutput, String> {
    if max_output_bytes == 0 {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    let cfb = cfb::CompoundFile::open(data).map_err(|error| format!("UNSUPPORTED:{error}"))?;
    let classified = if cfb.has_entry("WordDocument") {
        Some(LegacyFormat::Doc)
    } else if cfb.has_entry("Workbook") || cfb.has_entry("Book") {
        Some(LegacyFormat::Xls)
    } else if cfb.has_entry("PowerPoint Document") {
        Some(LegacyFormat::Ppt)
    } else {
        None
    };
    if classified != Some(format) {
        return Err("UNSUPPORTED:CFB stream family does not match the requested conversion".into());
    }
    match format {
        LegacyFormat::Doc => {
            let output = doc::convert(&cfb, max_output_bytes)?;
            Ok(NativeConversionOutput {
                bytes: output.bytes,
                warnings: output.warnings,
            })
        }
        LegacyFormat::Xls => {
            let output = xls::convert(&cfb, max_output_bytes)?;
            Ok(NativeConversionOutput {
                bytes: output.bytes,
                warnings: output.warnings,
            })
        }
        LegacyFormat::Ppt => {
            let output = ppt::convert(&cfb, max_output_bytes)?;
            Ok(NativeConversionOutput {
                bytes: output.bytes,
                warnings: output.warnings,
            })
        }
    }
}

#[wasm_bindgen]
pub struct LegacyConversionOutput {
    bytes: Option<Vec<u8>>,
    warnings: String,
}

#[wasm_bindgen]
impl LegacyConversionOutput {
    /// Transfer the generated OOXML package out of WASM exactly once.
    pub fn take_bytes(&mut self) -> Result<Vec<u8>, JsValue> {
        self.bytes
            .take()
            .ok_or_else(|| JsValue::from_str("conversion output bytes were already taken"))
    }

    /// Content-free stable warning identifiers, separated by newlines.
    pub fn warnings(&self) -> String {
        self.warnings.clone()
    }
}

/// Convert one CFB-based Office 97-2003 file into a newly-created macro-free
/// OOXML package. The JS host enforces time/cancellation by terminating the
/// disposable Worker; this synchronous entry enforces input structure and the
/// output byte ceiling inside the converter boundary.
#[wasm_bindgen]
pub fn convert_legacy_office(
    data: Vec<u8>,
    from: &str,
    max_output_bytes: u32,
) -> Result<LegacyConversionOutput, JsValue> {
    console_error_panic_hook::set_once();
    let format = LegacyFormat::parse(from).map_err(|error| JsValue::from_str(&error))?;
    let output = convert_native(&data, format, max_output_bytes as usize)
        .map_err(|error| JsValue::from_str(&error))?;
    Ok(LegacyConversionOutput {
        bytes: Some(output.bytes),
        warnings: output.warnings.join("\n"),
    })
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Read};

    use super::cfb::test_support::build_cfb;
    use super::{convert_native, LegacyFormat};

    fn zip_text(bytes: Vec<u8>, name: &str) -> String {
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut output = String::new();
        archive
            .by_name(name)
            .unwrap()
            .read_to_string(&mut output)
            .unwrap();
        output
    }

    #[test]
    fn converts_a_word_97_unicode_main_story() {
        let text = "Hello 日本語\rSecond paragraph";
        let units: Vec<u16> = text.encode_utf16().collect();
        let text_offset = 0x400usize;
        let mut word = vec![0u8; text_offset + units.len() * 2];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[0x4c..0x50].copy_from_slice(&(units.len() as u32).to_le_bytes());
        word[0x1a2..0x1a6].copy_from_slice(&0u32.to_le_bytes());
        word[0x1a6..0x1aa].copy_from_slice(&21u32.to_le_bytes());
        for (index, unit) in units.iter().enumerate() {
            word[text_offset + index * 2..text_offset + index * 2 + 2]
                .copy_from_slice(&unit.to_le_bytes());
        }
        let mut table = vec![0x02];
        table.extend_from_slice(&16u32.to_le_bytes());
        table.extend_from_slice(&0u32.to_le_bytes());
        table.extend_from_slice(&(units.len() as u32).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        table.extend_from_slice(&(text_offset as u32).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        let cfb = build_cfb(&[("WordDocument", word), ("0Table", table)]);
        let output = convert_native(&cfb, LegacyFormat::Doc, 1024 * 1024).unwrap();
        let xml = zip_text(output.bytes, "word/document.xml");
        assert!(xml.contains("Hello 日本語"));
        assert!(xml.contains("Second paragraph"));
        assert!(!xml.to_ascii_lowercase().contains("vba"));
    }

    #[test]
    fn converts_biff8_scalar_and_unicode_cells() {
        fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
            let mut output = Vec::with_capacity(payload.len() + 4);
            output.extend_from_slice(&kind.to_le_bytes());
            output.extend_from_slice(&(payload.len() as u16).to_le_bytes());
            output.extend_from_slice(payload);
            output
        }
        fn bof(kind: u16) -> Vec<u8> {
            let mut payload = Vec::new();
            payload.extend_from_slice(&0x0600u16.to_le_bytes());
            payload.extend_from_slice(&kind.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            record(0x0809, &payload)
        }
        let mut workbook = bof(0x0005);
        let bound_offset = workbook.len();
        let mut bound = vec![0; 6];
        bound.push(3);
        bound.push(1);
        for unit in "表計算".encode_utf16() {
            bound.extend_from_slice(&unit.to_le_bytes());
        }
        workbook.extend(record(0x0085, &bound));
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&3u16.to_le_bytes());
        sst.push(1);
        for unit in "日本語".encode_utf16() {
            sst.extend_from_slice(&unit.to_le_bytes());
        }
        workbook.extend(record(0x00fc, &sst));
        workbook.extend(record(0x000a, &[]));
        let sheet_offset = workbook.len();
        workbook.extend(bof(0x0010));
        let mut number = vec![0; 6];
        number.extend_from_slice(&42.5f64.to_le_bytes());
        workbook.extend(record(0x0203, &number));
        let mut label = Vec::new();
        label.extend_from_slice(&1u16.to_le_bytes());
        label.extend_from_slice(&1u16.to_le_bytes());
        label.extend_from_slice(&0u16.to_le_bytes());
        label.extend_from_slice(&0u32.to_le_bytes());
        workbook.extend(record(0x00fd, &label));
        workbook.extend(record(0x000a, &[]));
        let offset = bound_offset + 4;
        workbook[offset..offset + 4].copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        let cfb = build_cfb(&[("Workbook", workbook)]);
        let output = convert_native(&cfb, LegacyFormat::Xls, 1024 * 1024).unwrap();
        let workbook_xml = zip_text(output.bytes.clone(), "xl/workbook.xml");
        let sheet_xml = zip_text(output.bytes, "xl/worksheets/sheet1.xml");
        assert!(workbook_xml.contains("表計算"));
        assert!(sheet_xml.contains("42.5"));
        assert!(sheet_xml.contains("日本語"));
    }

    #[test]
    fn converts_powerpoint_unicode_text_atoms() {
        fn ppt_record(version: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
            let mut output = Vec::new();
            output.extend_from_slice(&version.to_le_bytes());
            output.extend_from_slice(&kind.to_le_bytes());
            output.extend_from_slice(&(payload.len() as u32).to_le_bytes());
            output.extend_from_slice(payload);
            output
        }
        let mut document_atom = vec![0u8; 40];
        document_atom[..4].copy_from_slice(&5760u32.to_le_bytes());
        document_atom[4..8].copy_from_slice(&4320u32.to_le_bytes());
        let mut slide_ref = [0u8; 20];
        slide_ref[..4].copy_from_slice(&2u32.to_le_bytes());
        let document = ppt_record(
            0x000f,
            1000,
            &[
                ppt_record(1, 1001, &document_atom),
                ppt_record(15, 4080, &ppt_record(0, 1011, &slide_ref)),
            ]
            .concat(),
        );
        let mut text = Vec::new();
        for unit in "Legacy 日本語 slide".encode_utf16() {
            text.extend_from_slice(&unit.to_le_bytes());
        }
        let atom = ppt_record(0, 4000, &text);
        let slide = ppt_record(0x000f, 1006, &atom);
        let directory_offset = document.len() + slide.len();
        let directory = ppt_record(
            0,
            0x1772,
            &[
                0x00200001u32.to_le_bytes(),
                0u32.to_le_bytes(),
                (document.len() as u32).to_le_bytes(),
            ]
            .concat(),
        );
        let current_edit = directory_offset + directory.len();
        let mut user_payload = [0u8; 28];
        user_payload[12..16].copy_from_slice(&(directory_offset as u32).to_le_bytes());
        user_payload[16..20].copy_from_slice(&1u32.to_le_bytes());
        let user_edit = ppt_record(0, 0x0ff5, &user_payload);
        let mut current_user_payload = vec![0; 24];
        current_user_payload[0..4].copy_from_slice(&0x14u32.to_le_bytes());
        current_user_payload[4..8].copy_from_slice(&0xe391_c05fu32.to_le_bytes());
        current_user_payload[8..12].copy_from_slice(&(current_edit as u32).to_le_bytes());
        current_user_payload[14..16].copy_from_slice(&0x03f4u16.to_le_bytes());
        current_user_payload[16] = 3;
        current_user_payload[20..24].copy_from_slice(&8u32.to_le_bytes());
        let current_user = ppt_record(0, 0x0ff6, &current_user_payload);
        let cfb = build_cfb(&[
            (
                "PowerPoint Document",
                [document, slide, directory, user_edit].concat(),
            ),
            ("Current User", current_user),
        ]);
        let output = convert_native(&cfb, LegacyFormat::Ppt, 1024 * 1024).unwrap();
        let xml = zip_text(output.bytes, "ppt/slides/slide1.xml");
        assert!(xml.contains("Legacy 日本語 slide"));
    }

    #[test]
    fn rejects_powerpoint_invalid_edit_chain() {
        fn ppt_record(kind: u16, payload: &[u8]) -> Vec<u8> {
            let mut output = Vec::new();
            output.extend_from_slice(&0u16.to_le_bytes());
            output.extend_from_slice(&kind.to_le_bytes());
            output.extend_from_slice(&(payload.len() as u32).to_le_bytes());
            output.extend_from_slice(payload);
            output
        }
        let mut user_edit_payload = vec![0; 24];
        user_edit_payload[8..12].copy_from_slice(&1u32.to_le_bytes());
        let user_edit = ppt_record(0x0ff5, &user_edit_payload);
        let mut current_user_payload = vec![0; 24];
        current_user_payload[0..4].copy_from_slice(&0x14u32.to_le_bytes());
        current_user_payload[4..8].copy_from_slice(&0xe391_c05fu32.to_le_bytes());
        current_user_payload[8..12].copy_from_slice(&0u32.to_le_bytes());
        current_user_payload[14..16].copy_from_slice(&0x03f4u16.to_le_bytes());
        current_user_payload[16] = 3;
        current_user_payload[20..24].copy_from_slice(&8u32.to_le_bytes());
        let current_user = ppt_record(0x0ff6, &current_user_payload);
        let ppt = build_cfb(&[
            ("PowerPoint Document", user_edit),
            ("Current User", current_user),
        ]);
        assert!(convert_native(&ppt, LegacyFormat::Ppt, 1024)
            .unwrap_err()
            .contains("invalid PowerPoint UserEditAtom"));
    }

    #[test]
    fn rejects_cross_family_inputs() {
        let cfb = build_cfb(&[("Workbook", vec![0; 16])]);
        let error = convert_native(&cfb, LegacyFormat::Doc, 1024).unwrap_err();
        assert!(error.starts_with("UNSUPPORTED:"));
    }

    #[test]
    fn rejects_encrypted_legacy_office_inputs() {
        let mut word = vec![0u8; 0x01aa];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[0x0a..0x0c].copy_from_slice(&0x0100u16.to_le_bytes());
        let doc = build_cfb(&[("WordDocument", word)]);
        assert!(convert_native(&doc, LegacyFormat::Doc, 1024)
            .unwrap_err()
            .contains("encrypted"));

        let mut workbook = Vec::new();
        workbook.extend_from_slice(&0x0809u16.to_le_bytes());
        workbook.extend_from_slice(&8u16.to_le_bytes());
        workbook.extend_from_slice(&0x0600u16.to_le_bytes());
        workbook.extend_from_slice(&0x0005u16.to_le_bytes());
        workbook.extend_from_slice(&0u32.to_le_bytes());
        workbook.extend_from_slice(&0x002fu16.to_le_bytes());
        workbook.extend_from_slice(&0u16.to_le_bytes());
        let xls = build_cfb(&[("Workbook", workbook)]);
        assert!(convert_native(&xls, LegacyFormat::Xls, 1024)
            .unwrap_err()
            .contains("encrypted"));

        let mut current_user = vec![0u8; 28];
        current_user[2..4].copy_from_slice(&0x0ff6u16.to_le_bytes());
        current_user[4..8].copy_from_slice(&20u32.to_le_bytes());
        current_user[8..12].copy_from_slice(&0x14u32.to_le_bytes());
        current_user[12..16].copy_from_slice(&0xf3d1_c4dfu32.to_le_bytes());
        let ppt = build_cfb(&[
            ("PowerPoint Document", vec![0; 16]),
            ("Current User", current_user),
        ]);
        assert!(convert_native(&ppt, LegacyFormat::Ppt, 1024)
            .unwrap_err()
            .contains("encrypted"));
    }
}
