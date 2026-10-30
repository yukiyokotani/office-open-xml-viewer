//! OpenDocument package access (ODF Part 2).
//!
//! The package is a ZIP opened through the shared bounded package session, so
//! structural validation, CRC/size checks and the inflation governors are the
//! ones the OOXML readers use. ODF item names are case-sensitive strings
//! (Part 2 §3.7), hence [`ItemNamePolicy::ExactUtf8`]: no OPC case folding or
//! part-name equivalence.

use crate::error::{invalid, unsupported, UNSUPPORTED_ENCRYPTION};
use crate::xml::{self, NS_MANIFEST};
use ooxml_common::package_session::PackageSessionHandle;
use ooxml_common::resource::OoxmlFormat;
use ooxml_common::zip::ItemNamePolicy;

/// ODF versions this reader implements (Part 3 §19.388 `office:version`,
/// Part 2 §4.16.14.2 `manifest:version`). Older documents may omit the
/// attribute or use semantics this reader does not model; they are rejected.
pub const SUPPORTED_VERSIONS: [&str; 3] = ["1.2", "1.3", "1.4"];

/// Data offset of the leading `mimetype` file: a 30-byte local header and
/// the 8-byte name, without an extra field (Part 2 §3.3).
const LEADING_MIMETYPE_DATA_START: usize = 30 + 8;

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
    /// Open a package whose `mimetype` must be one of `media_types`.
    pub fn open(
        data: Vec<u8>,
        format: OoxmlFormat,
        limits: PackageLimits,
        media_types: &[&str],
    ) -> Result<Self, String> {
        let media_type = leading_mimetype(&data)?;
        if !media_types.contains(&media_type.as_str()) {
            return Err(invalid("the package media type does not match this reader"));
        }
        let session = PackageSessionHandle::open_with_name_policy(
            data,
            format,
            ItemNamePolicy::ExactUtf8,
            limits.max_archive_entry_bytes,
            limits.max_total_inflated_bytes,
            limits.max_archive_entries,
        )?;
        let package = Self {
            session,
            media_type,
        };
        package.check_indexed_mimetype()?;
        package.check_manifest()?;
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

    /// The leading local header read before the session opened must be the
    /// archive's indexed first entry (§3.3): the central directory lists
    /// `mimetype` first, its data starts right after a local header at byte
    /// 0 (30 fixed bytes, the 8-byte name, no extra field), and its
    /// CRC-checked content equals the leading bytes.
    fn check_indexed_mimetype(&self) -> Result<(), String> {
        let first = self.session.entry_paths().into_iter().next();
        if first.as_deref() != Some("mimetype") {
            return Err(invalid("the mimetype file is not the first package entry"));
        }
        if self.session.entry_data_start("mimetype") != Some(LEADING_MIMETYPE_DATA_START) {
            return Err(invalid(
                "the indexed mimetype entry is not the leading file",
            ));
        }
        let bytes = self.session.run_operation("odf-mimetype", |operation| {
            operation.read_bytes_bounded("mimetype", MAX_MIMETYPE_BYTES)
        })?;
        if bytes != self.media_type.as_bytes() {
            return Err(invalid(
                "the indexed mimetype entry differs from the leading file",
            ));
        }
        Ok(())
    }

    /// Part 2 §4: the manifest lists the package's files. Encrypted entries
    /// (`manifest:encryption-data`, §4.8) are not decrypted. The root entry's
    /// media type must equal the `mimetype` file (§3.3, §4.3).
    fn check_manifest(&self) -> Result<(), String> {
        let text = self.read_text("META-INF/manifest.xml")?;
        let document = xml::parse(&text)?;
        let root = document.root_element();
        if !xml::is(root, NS_MANIFEST, "manifest") {
            return Err(invalid(
                "the manifest root element is not manifest:manifest",
            ));
        }
        let version = root.attribute((NS_MANIFEST, "version"));
        if !version.is_some_and(|value| SUPPORTED_VERSIONS.contains(&value)) {
            return Err(unsupported("odf-version"));
        }
        let mut root_media_type = None;
        for node in root.descendants() {
            if xml::is(node, NS_MANIFEST, "encryption-data") {
                return Err(UNSUPPORTED_ENCRYPTION.to_string());
            }
            if xml::is(node, NS_MANIFEST, "file-entry")
                && node.attribute((NS_MANIFEST, "full-path")) == Some("/")
            {
                if root_media_type.is_some() {
                    return Err(invalid("the manifest lists the package root twice"));
                }
                root_media_type = node.attribute((NS_MANIFEST, "media-type"));
            }
        }
        if root_media_type != Some(self.media_type.as_str()) {
            return Err(invalid(
                "the manifest root media type differs from mimetype",
            ));
        }
        Ok(())
    }
}

/// The media type in the package's leading `mimetype` file (Part 2 §3.3): the
/// first ZIP entry, named `mimetype`, stored without compression, encryption,
/// data descriptor or extra field, holding the media type in ASCII. A packaged
/// ODT, ODS or ODP must carry it (Part 3 §2.2.3–§2.2.5).
pub fn leading_mimetype(data: &[u8]) -> Result<String, String> {
    let u16_at = |offset: usize| -> Option<u16> {
        Some(u16::from_le_bytes(
            data.get(offset..offset + 2)?.try_into().ok()?,
        ))
    };
    let u32_at = |offset: usize| -> Option<u32> {
        Some(u32::from_le_bytes(
            data.get(offset..offset + 4)?.try_into().ok()?,
        ))
    };
    let malformed = || invalid("the package does not start with a conforming mimetype file");
    if u32_at(0) != Some(0x0403_4b50) {
        return Err(malformed());
    }
    let flags = u16_at(6).ok_or_else(malformed)?;
    let method = u16_at(8).ok_or_else(malformed)?;
    let compressed = u32_at(18).ok_or_else(malformed)? as usize;
    let size = u32_at(22).ok_or_else(malformed)? as usize;
    let name_length = u16_at(26).ok_or_else(malformed)? as usize;
    let extra_length = u16_at(28).ok_or_else(malformed)?;
    if flags & 0x0009 != 0
        || method != 0
        || extra_length != 0
        || compressed != size
        || size == 0
        || size > MAX_MIMETYPE_BYTES
        || data.get(30..30 + name_length) != Some(b"mimetype".as_slice())
    {
        return Err(malformed());
    }
    let start = 30 + name_length;
    let bytes = data.get(start..start + size).ok_or_else(malformed)?;
    if !bytes.iter().all(|byte| (0x21..=0x7e).contains(byte)) {
        return Err(malformed());
    }
    Ok(String::from_utf8_lossy(bytes).into_owned())
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
                writer.write_all(body.as_bytes()).unwrap();
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

    #[test]
    fn rejects_nonconforming_mimetype_headers() {
        let good = package(ODT, &[("META-INF/manifest.xml", &manifest(ODT, ""))]);
        // Another media type, and the same file with an extra field or compressed.
        assert!(open(package(
            "application/vnd.oasis.opendocument.spreadsheet",
            &[]
        ))
        .is_err());
        let mut extra = good.clone();
        extra[28] = 1;
        assert!(leading_mimetype(&extra).is_err());
        let mut compressed = good.clone();
        compressed[8] = 8;
        assert!(leading_mimetype(&compressed).is_err());
        assert_eq!(leading_mimetype(&good).unwrap(), ODT);
        assert!(leading_mimetype(b"PK\x03\x04").is_err());
    }

    #[test]
    fn checks_the_manifest_root_version_and_encryption() {
        let mismatched = package(
            ODT,
            &[(
                "META-INF/manifest.xml",
                &manifest("application/vnd.oasis.opendocument.text-template", ""),
            )],
        );
        assert!(open(mismatched)
            .err()
            .unwrap()
            .contains("manifest root media type"));
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
        let old = package(
            ODT,
            &[(
                "META-INF/manifest.xml",
                &manifest(ODT, "").replace("manifest:version=\"1.3\"", "manifest:version=\"1.1\""),
            )],
        );
        assert_eq!(open(old).err().unwrap(), unsupported("odf-version"));
        assert!(open(package(ODT, &[])).is_err());
    }

    #[test]
    fn the_leading_mimetype_must_be_the_indexed_first_entry() {
        let mut bytes = package(ODT, &[("META-INF/manifest.xml", &manifest(ODT, ""))]);
        assert!(open(bytes.clone()).is_ok());
        // Swap the first two central directory records: the leading local
        // header is unchanged, but the index lists the manifest first.
        let eocd = bytes.len() - 22;
        let u16_at = |b: &[u8], o: usize| u16::from_le_bytes([b[o], b[o + 1]]) as usize;
        let start = u32::from_le_bytes(bytes[eocd + 16..eocd + 20].try_into().unwrap()) as usize;
        let record =
            |b: &[u8], o: usize| 46 + u16_at(b, o + 28) + u16_at(b, o + 30) + u16_at(b, o + 32);
        let first = record(&bytes, start);
        let second = record(&bytes, start + first);
        let mut swapped = bytes[start + first..start + first + second].to_vec();
        swapped.extend_from_slice(&bytes[start..start + first]);
        bytes.splice(start..start + first + second, swapped);
        assert_eq!(leading_mimetype(&bytes).unwrap(), ODT);
        assert!(open(bytes)
            .err()
            .unwrap()
            .contains("not the first package entry"));
    }

    #[test]
    fn an_unindexed_leading_record_does_not_stand_in_for_the_mimetype_entry() {
        let good = package(ODT, &[("META-INF/manifest.xml", &manifest(ODT, ""))]);
        // Prepend a copy of the leading local record (header, name, data) with
        // a wrong CRC; the archive's own entries move but stay indexed.
        let record_length = 30 + 8 + ODT.len();
        let mut prefix = good[..record_length].to_vec();
        prefix[14] ^= 0xff;
        let mut bytes = prefix;
        bytes.extend_from_slice(&good);
        assert_eq!(leading_mimetype(&bytes).unwrap(), ODT);
        let error = open(bytes).err().unwrap();
        assert!(error.contains("not the leading file"), "{error}");
    }
}
