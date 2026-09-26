//! Package-level admission of an OOXML document (ECMA-376 Part 2, OPC).
//!
//! A byte stream is an OOXML document only when all of the following hold:
//!
//! 1. it is a readable ZIP archive (Part 2 §7.3, ZIP physical mapping);
//! 2. the ZIP carries the Media Types stream `[Content_Types].xml`
//!    (Part 2 §7.2.3.1 requires it for every non-MIME physical format, and
//!    §7.3.7 fixes its ZIP item name);
//! 3. the format's main part (the `officeDocument` relationship target) is
//!    present.
//!
//! Failing any of these is a whole-input rejection, not a degraded document:
//! such input is not an OOXML package, so there is no document to degrade, and
//! a placeholder would fabricate page/sheet/slide geometry that the input never
//! declared (library policy). Damage *inside* an admitted package
//! (a malformed part, a broken slide) is a different class and stays the
//! caller's per-part degradation concern.
//!
//! Library policy: the main part is checked at the conventional part name the
//! format parser reads. Resolving a non-conventional `officeDocument` target
//! from `/_rels/.rels` is not implemented by any format parser, so such a
//! package is rejected here instead of producing an empty document. The main
//! part is matched under part-name equivalence (§6.2.2.3), as every part read
//! is; the Media Types stream is a ZIP item, not a part, so it is matched by
//! its exact §7.3.7 item name. The interleaved Media Types
//! form (`[Content_Types].xml/[n].piece`, §7.3.7) is not readable by any
//! parser and is likewise rejected.
//!
//! The rejection crosses the WASM boundary as a string with
//! [`NOT_OOXML_PREFIX`]; the TypeScript layer reconstructs it as
//! `OoxmlError('not-ooxml')`. Resource-limit envelopes are never rewritten.

use crate::package_session::PackageSessionHandle;
use crate::resource::OoxmlFormat;

/// Machine-readable envelope for a package that is not an OOXML document.
/// Must stay in sync with `NOT_OOXML_PREFIX` in `packages/core/src/worker/error-wire.ts`.
pub const NOT_OOXML_PREFIX: &str = "OOXML_NOT_OOXML:";

const RESOURCE_LIMIT_PREFIX: &str = "OOXML_RESOURCE_LIMIT:";

/// OPC Media Types stream item name (ECMA-376 Part 2 §7.3.7).
pub const CONTENT_TYPES_ITEM: &str = "[Content_Types].xml";

/// Conventional main part read by each format parser.
pub fn main_part_name(format: OoxmlFormat) -> &'static str {
    match format {
        OoxmlFormat::Docx => "word/document.xml",
        OoxmlFormat::Xlsx => "xl/workbook.xml",
        OoxmlFormat::Pptx => "ppt/presentation.xml",
    }
}

fn not_ooxml(detail: impl std::fmt::Display) -> String {
    format!("{NOT_OOXML_PREFIX}{detail}")
}

/// Whether `error` is the [`NOT_OOXML_PREFIX`] envelope.
pub fn is_not_ooxml_error(error: &str) -> bool {
    error.starts_with(NOT_OOXML_PREFIX)
}

/// Map a ZIP container open failure to the not-OOXML envelope. A resource-limit
/// envelope is a policy decision about a possibly valid package and passes
/// through unchanged.
pub fn container_open_error(error: String) -> String {
    if error.starts_with(RESOURCE_LIMIT_PREFIX) || is_not_ooxml_error(&error) {
        error
    } else {
        not_ooxml(format!("the input is not a readable ZIP package: {error}"))
    }
}

/// Require an opened ZIP to be an OPC package containing the format's main part.
pub fn require_ooxml_package(
    package: &PackageSessionHandle,
    format: OoxmlFormat,
) -> Result<(), String> {
    if !package.contains_exact_entry(CONTENT_TYPES_ITEM) {
        return Err(not_ooxml(format!(
            "the ZIP archive is not an OPC package (missing {CONTENT_TYPES_ITEM})"
        )));
    }
    let main_part = main_part_name(format);
    if !package.contains_entry(main_part) {
        return Err(not_ooxml(format!(
            "the OPC package has no main document part ({main_part})"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn zip_with(entries: &[&str]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for name in entries {
            writer
                .start_file(*name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"<x/>").unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    fn admit(entries: &[&str], format: OoxmlFormat) -> Result<(), String> {
        let package =
            PackageSessionHandle::open(zip_with(entries), format, None, None, None).unwrap();
        require_ooxml_package(&package, format)
    }

    #[test]
    fn admits_each_format_main_part_and_rejects_the_rest() {
        for format in [OoxmlFormat::Docx, OoxmlFormat::Xlsx, OoxmlFormat::Pptx] {
            let main = main_part_name(format);
            assert_eq!(admit(&[CONTENT_TYPES_ITEM, main], format), Ok(()));
            let missing_main = admit(&[CONTENT_TYPES_ITEM, "_rels/.rels"], format).unwrap_err();
            assert!(is_not_ooxml_error(&missing_main), "{missing_main}");
            assert!(missing_main.contains(main), "{missing_main}");
            let not_opc = admit(&[main], format).unwrap_err();
            assert!(is_not_ooxml_error(&not_opc), "{not_opc}");
            assert!(not_opc.contains(CONTENT_TYPES_ITEM), "{not_opc}");
        }
    }

    #[test]
    fn main_part_uses_part_name_equivalence_but_media_types_stream_is_exact() {
        assert_eq!(
            admit(
                &[CONTENT_TYPES_ITEM, "Word/Document.XML"],
                OoxmlFormat::Docx
            ),
            Ok(())
        );
        let lowercase_types = admit(
            &["[content_types].xml", "word/document.xml"],
            OoxmlFormat::Docx,
        )
        .unwrap_err();
        assert!(is_not_ooxml_error(&lowercase_types), "{lowercase_types}");
    }

    #[test]
    fn container_errors_preserve_resource_envelopes() {
        let resource = format!("{RESOURCE_LIMIT_PREFIX}{{}}");
        assert_eq!(container_open_error(resource.clone()), resource);
        assert!(is_not_ooxml_error(&container_open_error(
            "(zip container): bad".to_string()
        )));
    }
}
