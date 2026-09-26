//! OPC part lookup for ZIPs embedded in legacy binary records.
//!
//! ECMA-376 Part 2 §6.2.2.3 gives part names ASCII case and RFC 3986
//! percent-encoding equivalence. Reject an ambiguous embedded ZIP rather than
//! choosing whichever duplicate its index happens to return.

pub(crate) fn entry_index<R: std::io::Read + std::io::Seek>(
    archive: &zip::ZipArchive<R>,
    path: &str,
) -> Option<usize> {
    let key = ooxml_common::rels::part_name_equivalence_key(path);
    let mut found = None;
    for name in archive.file_names() {
        if ooxml_common::rels::part_name_equivalence_key(name) == key {
            if found.is_some() {
                return None;
            }
            found = archive.index_for_name(name);
        }
    }
    found
}
