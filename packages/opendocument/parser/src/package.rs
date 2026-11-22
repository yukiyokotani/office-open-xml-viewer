//! OpenDocument package access (ODF Part 2).
//!
//! The package is a ZIP opened through the shared bounded package session, so
//! structural validation, CRC/size checks and the inflation governors are the
//! ones the OOXML readers use. ODF item names are case-sensitive strings
//! (Part 2 §3.7), hence [`ItemNamePolicy::ExactUtf8`]: no OPC case folding or
//! part-name equivalence.

use crate::error::{invalid, UNSUPPORTED_ENCRYPTION};
use crate::xml::{self, NS_MANIFEST};
use ooxml_common::package_session::PackageSessionHandle;
use ooxml_common::resource::OoxmlFormat;
use ooxml_common::zip::ItemNamePolicy;

/// Implementation bound: every media type a reader admits is shorter.
const MAX_MIMETYPE_BYTES: usize = 128;

/// Caller ZIP limits (the loader's normalized `resourceLimits`); `None` means
/// the caller removed the limit, leaving the governor's hard ceilings.
#[derive(Clone, Copy, Debug, Default)]
pub struct PackageLimits {
    pub max_archive_entry_bytes: Option<u64>,
    pub max_total_inflated_bytes: Option<u64>,
    pub max_archive_entries: Option<u64>,
}

pub struct OdfPackage {
    session: PackageSessionHandle,
    media_type: String,
}

impl OdfPackage {
    /// Open a package whose `mimetype` file holds one of `media_types`.
    ///
    /// Part 2 §3.3 asks producers to store `mimetype` first and uncompressed,
    /// without an extra field, so that the type can be sniffed at a fixed
    /// offset. A package re-zipped by a general archiver breaks those
    /// placement rules but keeps its content, and LibreOffice opens it; so
    /// does this reader. The type is read from the indexed `mimetype` entry
    /// wherever it is and however it is compressed, through the same bounded
    /// session (CRC, size and inflation checks) as every other file.
    pub fn open(
        data: Vec<u8>,
        format: OoxmlFormat,
        limits: PackageLimits,
        media_types: &[&str],
    ) -> Result<Self, String> {
        let session = PackageSessionHandle::open_with_name_policy(
            data,
            format,
            ItemNamePolicy::ExactUtf8,
            limits.max_archive_entry_bytes,
            limits.max_total_inflated_bytes,
            limits.max_archive_entries,
        )?;
        let media_type = read_mimetype(&session)?;
        if !media_types.contains(&media_type.as_str()) {
            return Err(invalid("the package media type does not match this reader"));
        }
        let package = Self {
            session,
            media_type,
        };
        package.check_encryption()?;
        Ok(package)
    }

    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    pub fn session(&self) -> &PackageSessionHandle {
        &self.session
    }

    pub fn contains(&self, path: &str) -> bool {
        self.session.contains_exact_entry(path)
    }

    /// One package file as UTF-8 text, through the session's accounting.
    pub fn read_text(&self, path: &str) -> Result<String, String> {
        if !self.contains(path) {
            return Err(invalid("a required package file is missing"));
        }
        self.session
            .run_operation("odf-read", |operation| operation.read_string(path))
    }

    /// Part 2 §4.8: an encrypted file carries `manifest:encryption-data` in
    /// its manifest entry, and is not decrypted. The manifest is read only
    /// for this: its version, root element and root media type are not
    /// checked (the `mimetype` file decides the type), and a package without
    /// a manifest is read from its files.
    fn check_encryption(&self) -> Result<(), String> {
        const MANIFEST: &str = "META-INF/manifest.xml";
        if !self.contains(MANIFEST) {
            return Ok(());
        }
        let text = self.read_text(MANIFEST)?;
        let document = xml::parse(&text)?;
        if document
            .descendants()
            .any(|node| xml::is(node, NS_MANIFEST, "encryption-data"))
        {
            return Err(UNSUPPORTED_ENCRYPTION.to_string());
        }
        Ok(())
    }
}

/// The media type in the package's `mimetype` file (Part 2 §3.3), which a
/// packaged ODT, ODS or ODP must carry (Part 3 §2.2.3–§2.2.5): ASCII, without
/// white space.
fn read_mimetype(session: &PackageSessionHandle) -> Result<String, String> {
    if !session.contains_exact_entry("mimetype") {
        return Err(invalid("the package has no mimetype file"));
    }
    let bytes = session.run_operation("odf-mimetype", |operation| {
        operation.read_bytes_bounded("mimetype", MAX_MIMETYPE_BYTES)
    })?;
    if bytes.is_empty() || !bytes.iter().all(|byte| (0x21..=0x7e).contains(byte)) {
        return Err(invalid("the mimetype file does not hold a media type"));
    }
    Ok(String::from_utf8_lossy(&bytes).into_owned())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::CompressionMethod;

    pub(crate) const ODT: &str = "application/vnd.oasis.opendocument.text";

    pub(crate) fn manifest(root_media_type: &str, extra: &str) -> String {
        format!(
            r#"<manifest:manifest xmlns:manifest="{NS_MANIFEST}" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{root_media_type}"/>{extra}</manifest:manifest>"#
        )
    }

    /// A package with a conforming leading `mimetype` and the given files.
    pub(crate) fn package(mimetype: &str, files: &[(&str, &str)]) -> Vec<u8> {
        let files: Vec<(&str, &[u8])> = files
            .iter()
            .map(|(name, body)| (*name, body.as_bytes()))
            .collect();
        package_bytes(mimetype, &files)
    }

    pub(crate) fn package_bytes(mimetype: &str, files: &[(&str, &[u8])]) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
            writer.start_file("mimetype", stored).unwrap();
            writer.write_all(mimetype.as_bytes()).unwrap();
            let deflated =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            for (name, body) in files {
                writer.start_file(*name, deflated).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn open(bytes: Vec<u8>) -> Result<OdfPackage, String> {
        OdfPackage::open(bytes, OoxmlFormat::Docx, PackageLimits::default(), &[ODT])
    }

    #[test]
    fn opens_a_conforming_package_and_reads_exact_names() {
        let bytes = package(
            ODT,
            &[
                ("META-INF/manifest.xml", &manifest(ODT, "")),
                ("content.xml", "<a/>"),
                ("Content.xml", "<b/>"),
            ],
        );
        let package = open(bytes).unwrap();
        assert_eq!(package.media_type(), ODT);
        assert_eq!(package.read_text("content.xml").unwrap(), "<a/>");
        assert_eq!(package.read_text("Content.xml").unwrap(), "<b/>");
        assert!(package.read_text("CONTENT.xml").is_err());
    }

    /// One archive entry: name, content, method, and whether it carries an
    /// extra field.
    type Entry<'a> = (&'a str, &'a [u8], CompressionMethod, bool);

    /// A package written entry by entry, as a general archiver might.
    fn zip_entries(entries: &[Entry<'_>]) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            for (name, body, method, extra) in entries {
                let mut options: zip::write::FullFileOptions<'_, '_> =
                    zip::write::FullFileOptions::default().compression_method(*method);
                if *extra {
                    // Info-ZIP's extended timestamp ("UT").
                    options
                        .add_extra_field(0x5455, [1u8, 0, 0, 0, 0], false)
                        .unwrap();
                }
                writer.start_file(*name, options).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    #[test]
    fn re_zipped_packages_open_by_their_mimetype_content() {
        let manifest = manifest(ODT, "");
        let content: &[u8] = b"<a/>";
        let deflated = CompressionMethod::Deflated;
        let stored = CompressionMethod::Stored;
        let cases: [&[Entry<'_>]; 4] = [
            // Compressed mimetype.
            &[
                ("mimetype", ODT.as_bytes(), deflated, false),
                (
                    "META-INF/manifest.xml",
                    manifest.as_bytes(),
                    deflated,
                    false,
                ),
                ("content.xml", content, deflated, false),
            ],
            // Not the first entry (an archiver's directory order).
            &[
                (
                    "META-INF/manifest.xml",
                    manifest.as_bytes(),
                    deflated,
                    false,
                ),
                ("content.xml", content, deflated, false),
                ("mimetype", ODT.as_bytes(), stored, false),
            ],
            // Extra fields on every entry.
            &[
                ("mimetype", ODT.as_bytes(), stored, true),
                ("META-INF/manifest.xml", manifest.as_bytes(), deflated, true),
                ("content.xml", content, deflated, true),
            ],
            // No manifest at all.
            &[
                ("content.xml", content, deflated, false),
                ("mimetype", ODT.as_bytes(), deflated, false),
            ],
        ];
        for (index, entries) in cases.iter().enumerate() {
            let package = open(zip_entries(entries)).unwrap_or_else(|e| panic!("{index}: {e}"));
            assert_eq!(package.media_type(), ODT);
            assert_eq!(package.read_text("content.xml").unwrap(), "<a/>");
        }
    }

    #[test]
    fn the_manifest_only_reports_encryption() {
        // A root entry naming another type (or none), another root element
        // and an old manifest version do not decide the type.
        for manifest in [
            manifest("application/vnd.oasis.opendocument.text-template", ""),
            format!(r#"<manifest:manifest xmlns:manifest="{NS_MANIFEST}"/>"#),
            format!(r#"<other xmlns:manifest="{NS_MANIFEST}"/>"#),
            manifest(ODT, "").replace("manifest:version=\"1.3\"", "manifest:version=\"1.1\""),
            manifest(
                ODT,
                r#"<manifest:file-entry manifest:full-path="/" manifest:media-type="x"/>"#,
            ),
        ] {
            let package = open(package(ODT, &[("META-INF/manifest.xml", &manifest)]));
            assert_eq!(package.unwrap().media_type(), ODT, "{manifest}");
        }
        let encrypted = package(
            ODT,
            &[(
                "META-INF/manifest.xml",
                &manifest(
                    ODT,
                    r#"<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data/></manifest:file-entry>"#,
                ),
            )],
        );
        assert_eq!(open(encrypted).err().unwrap(), UNSUPPORTED_ENCRYPTION);
        // A manifest that is not XML is corrupt.
        let broken = package(ODT, &[("META-INF/manifest.xml", "<manifest")]);
        assert!(open(broken).is_err());
    }

    #[test]
    fn packages_without_this_media_type_or_intact_data_reject() {
        let manifest = manifest(ODT, "");
        // Another family, a missing or empty mimetype, white space in it.
        let spreadsheet = "application/vnd.oasis.opendocument.spreadsheet";
        assert!(open(package(spreadsheet, &[])).is_err());
        let no_mimetype = zip_entries(&[(
            "META-INF/manifest.xml",
            manifest.as_bytes(),
            CompressionMethod::Deflated,
            false,
        )]);
        assert!(open(no_mimetype)
            .err()
            .unwrap()
            .contains("no mimetype file"));
        assert!(open(package("", &[])).is_err());
        assert!(open(package(&format!("{ODT}\n"), &[])).is_err());
        let long = "a".repeat(MAX_MIMETYPE_BYTES + 1);
        assert!(open(package(&long, &[])).is_err());
        // A corrupt mimetype (CRC mismatch) and a truncated archive.
        let mut corrupt = package(ODT, &[("META-INF/manifest.xml", &manifest)]);
        corrupt[30 + 8] ^= 0x01;
        assert!(open(corrupt).is_err());
        let good = package(ODT, &[("META-INF/manifest.xml", &manifest)]);
        assert!(open(good[..good.len() - 30].to_vec()).is_err());
        assert!(open(b"PK\x03\x04".to_vec()).is_err());
    }
}
