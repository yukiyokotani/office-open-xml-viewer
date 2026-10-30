//! Bounded ZIP access for OOXML packages.
//!
//! This module owns archive-specific reading and metadata inspection. Persistent
//! policy, accounting, usage snapshots, and poison state live in `resource`.

use crate::resource::{
    self, OoxmlFormat, ResourceGovernor, ResourceScope, HARD_MAX_ARCHIVE_ENTRIES,
    HARD_MAX_CENTRAL_DIRECTORY_BYTES,
};
use std::collections::{HashMap, HashSet};

/// Cap eager allocation based on attacker-controlled ZIP declarations. Large
/// legitimate entries grow incrementally while reads remain resource-bounded.
const INITIAL_RESERVE_CAP: usize = 1024 * 1024;
const READ_CHUNK_BYTES: usize = 32 * 1024;
const EOCD_SIGNATURE: u32 = 0x0605_4b50;
const ZIP64_LOCATOR_SIGNATURE: u32 = 0x0706_4b50;
const ZIP64_EOCD_SIGNATURE: u32 = 0x0606_4b50;
const CENTRAL_DIRECTORY_HEADER_SIGNATURE: u32 = 0x0201_4b50;
const CENTRAL_DIRECTORY_DIGITAL_SIGNATURE: u32 = 0x0505_4b50;
const LOCAL_FILE_HEADER_SIGNATURE: u32 = 0x0403_4b50;
const CENTRAL_DIRECTORY_HEADER_BYTES: u64 = 46;

fn initial_reserve(declared_size: u64, read_limit: u64) -> usize {
    declared_size
        .min(read_limit)
        .min(INITIAL_RESERVE_CAP as u64) as usize
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn le_u64(bytes: &[u8], offset: usize) -> Option<u64> {
    Some(u64::from_le_bytes(
        bytes.get(offset..offset + 8)?.try_into().ok()?,
    ))
}

/// How a package's ZIP item names are validated and looked up. Structural
/// validation and resource accounting are identical under both policies.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ItemNamePolicy {
    /// OPC naming (ECMA-376 Part 2): ASCII item names (§7.3.3), unique under
    /// part-name equivalence (§6.2.2.3: ASCII case folding and percent-encoded
    /// unreserved characters), no part name derivable from another, and
    /// lookups by equivalent part name.
    #[default]
    Opc,
    /// Exact naming for ZIP containers whose item names are case-sensitive
    /// strings: every name is valid UTF-8 (a non-ASCII name must also carry the
    /// APPNOTE 4.4.4 bit-11 UTF-8 flag, since without it a reader decodes the
    /// bytes as CP437), names are unique octet-for-octet, and lookups never fold.
    ExactUtf8,
}

#[derive(Clone, Copy, Debug)]
struct NameRange {
    start: usize,
    len: usize,
}

#[derive(Debug)]
struct ArchivePreflight {
    entry_count: u64,
    entries: Vec<PreflightEntryIdentity>,
    archive_base: u64,
    /// What the ZIP reader may see while it builds its index.
    index_view: IndexView,
    /// (locator, record) offsets when the directory was validated through a
    /// ZIP64 footer; `None` for a classic footer.
    zip64_footer: Option<(usize, usize)>,
}

#[derive(Debug)]
struct PreflightEntryIdentity {
    name: Box<[u8]>,
    central_header_start: u64,
    local_header_start: u64,
    data_start: u64,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
    compression_method: u16,
    encrypted: bool,
}

impl ArchivePreflight {
    fn read_config(&self) -> zip::read::Config {
        zip::read::Config {
            archive_offset: zip::read::ArchiveOffset::Known(self.archive_base),
        }
    }

    /// Confirm that the ZIP crate indexed exactly the item identities validated
    /// by this preflight token. This consumes neither a second EOCD search nor
    /// another filename set.
    fn validate_archive_item_names<R: std::io::Read + std::io::Seek>(
        &self,
        archive: &mut zip::ZipArchive<R>,
    ) -> Result<(), String> {
        if self.entry_count != archive.len() as u64 {
            return Err(format!(
                "ZIP index does not match the validated central directory: expected {} items but {} were indexed",
                self.entry_count,
                archive.len()
            ));
        }

        for (index, expected) in self.entries.iter().enumerate() {
            let entry = archive
                .by_index_raw(index)
                .map_err(|error| format!("ZIP archive entry metadata error: {error}"))?;
            let matches = entry.name_raw() == expected.name.as_ref()
                && entry.central_header_start() == expected.central_header_start
                && entry.header_start() == expected.local_header_start
                && entry.data_start() == Some(expected.data_start)
                && entry.compressed_size() == expected.compressed_size
                && entry.size() == expected.uncompressed_size
                && entry.crc32() == expected.crc32
                && match expected.compression_method {
                    0 => entry.compression() == zip::CompressionMethod::STORE,
                    8 => entry.compression() == zip::CompressionMethod::DEFLATE,
                    _ => false,
                }
                && entry.encrypted() == expected.encrypted;
            if !matches {
                return Err(format!(
                    "ZIP index item {index} does not match the validated central directory"
                ));
            }
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct CentralDirectory {
    start: usize,
    end: usize,
    size: u64,
    entry_count: u64,
    archive_base: u64,
    footer_metadata_bytes: u64,
}

enum CandidateError {
    Malformed,
    Rejected(String),
}

#[derive(Clone, Copy)]
struct CentralEntry {
    header: usize,
    name: NameRange,
    extra: NameRange,
}

fn resolve_u16(classic: u16, zip64: u64) -> u64 {
    if classic == u16::MAX {
        zip64
    } else {
        classic as u64
    }
}

fn resolve_u32(classic: u32, zip64: u64) -> u64 {
    if classic == u32::MAX {
        zip64
    } else {
        classic as u64
    }
}

fn checked_directory(
    boundary: usize,
    archive_base: u64,
    relative_offset: u64,
    size: u64,
    entry_count: u64,
    footer_metadata_bytes: u64,
) -> Option<CentralDirectory> {
    let start = archive_base.checked_add(relative_offset)?;
    let end = start.checked_add(size)?;
    if end != boundary as u64 || size < entry_count.checked_mul(CENTRAL_DIRECTORY_HEADER_BYTES)? {
        return None;
    }
    Some(CentralDirectory {
        start: usize::try_from(start).ok()?,
        end: boundary,
        size,
        entry_count,
        archive_base,
        footer_metadata_bytes,
    })
}

fn classic_directory(data: &[u8], eocd: usize) -> Option<CentralDirectory> {
    let count_on_disk = le_u16(data, eocd + 8)?;
    let count = le_u16(data, eocd + 10)?;
    let size = le_u32(data, eocd + 12)?;
    let offset = le_u32(data, eocd + 16)?;
    if count_on_disk == u16::MAX
        || count == u16::MAX
        || size == u32::MAX
        || offset == u32::MAX
        || count_on_disk != count
    {
        return None;
    }
    let relative_end = (offset as u64).checked_add(size as u64)?;
    let archive_base = (eocd as u64).checked_sub(relative_end)?;
    checked_directory(
        eocd,
        archive_base,
        offset as u64,
        size as u64,
        count as u64,
        0,
    )
}

fn inspect_zip64_candidates(
    data: &[u8],
    policy: ItemNamePolicy,
    eocd: usize,
    classic_count_on_disk: u16,
    classic_count: u16,
    classic_size: u32,
    classic_offset: u32,
) -> Result<Option<ArchivePreflight>, String> {
    let Some(locator) = eocd.checked_sub(20) else {
        return Ok(None);
    };
    if le_u32(data, locator) != Some(ZIP64_LOCATOR_SIGNATURE)
        || le_u32(data, locator + 4) != Some(0)
        || le_u32(data, locator + 16) != Some(1)
    {
        return Ok(None);
    }
    let Some(relative_zip64_offset) = le_u64(data, locator + 8) else {
        return Ok(None);
    };
    if locator < 56 {
        return Ok(None);
    }

    let mut resolved_directory = None;
    for zip64 in data[..locator]
        .windows(4)
        .enumerate()
        .filter_map(|(offset, bytes)| {
            (bytes == ZIP64_EOCD_SIGNATURE.to_le_bytes()).then_some(offset)
        })
    {
        let Some(record_size) = le_u64(data, zip64 + 4) else {
            continue;
        };
        if record_size < 44
            || (zip64 as u64)
                .checked_add(12)
                .and_then(|position| position.checked_add(record_size))
                != Some(locator as u64)
            || le_u32(data, zip64 + 16) != Some(0)
            || le_u32(data, zip64 + 20) != Some(0)
        {
            continue;
        }
        let Some(archive_base) = (zip64 as u64).checked_sub(relative_zip64_offset) else {
            continue;
        };
        let Some(zip64_count_on_disk) = le_u64(data, zip64 + 24) else {
            continue;
        };
        let Some(zip64_count) = le_u64(data, zip64 + 32) else {
            continue;
        };
        let Some(zip64_size) = le_u64(data, zip64 + 40) else {
            continue;
        };
        let Some(zip64_offset) = le_u64(data, zip64 + 48) else {
            continue;
        };
        let count_on_disk = resolve_u16(classic_count_on_disk, zip64_count_on_disk);
        let count = resolve_u16(classic_count, zip64_count);
        if count_on_disk != count {
            continue;
        }
        let size = resolve_u32(classic_size, zip64_size);
        let offset = resolve_u32(classic_offset, zip64_offset);
        let Some(directory) =
            checked_directory(zip64, archive_base, offset, size, count, record_size + 12)
        else {
            continue;
        };
        // The ZIP64 extensible-data sector may contain signature-shaped bytes.
        // Count only candidates whose central and local entry structures are
        // coherent, matching the semantic filtering used for classic EOCDs.
        if validate_central_directory_structure(data, directory).is_err() {
            continue;
        }
        if resolved_directory.replace((directory, zip64)).is_some() {
            return Err("ambiguous ZIP64 end-of-central-directory records".to_string());
        }
    }
    let Some((directory, zip64)) = resolved_directory else {
        return Ok(None);
    };
    match inspect_central_directory(data, directory, policy) {
        Ok(mut preflight) => {
            preflight.zip64_footer = Some((locator, zip64));
            Ok(Some(preflight))
        }
        Err(CandidateError::Malformed) => Ok(None),
        Err(CandidateError::Rejected(error)) => Err(error),
    }
}

fn inspect_eocd_candidate(
    data: &[u8],
    eocd: usize,
    policy: ItemNamePolicy,
) -> Result<Option<ArchivePreflight>, String> {
    let Some(comment_len) = le_u16(data, eocd + 20) else {
        return Ok(None);
    };
    if eocd.checked_add(22 + comment_len as usize) != Some(data.len())
        || le_u16(data, eocd + 4) != Some(0)
        || le_u16(data, eocd + 6) != Some(0)
    {
        return Ok(None);
    }

    let Some(classic_count_on_disk) = le_u16(data, eocd + 8) else {
        return Ok(None);
    };
    let Some(classic_count) = le_u16(data, eocd + 10) else {
        return Ok(None);
    };
    let Some(classic_size) = le_u32(data, eocd + 12) else {
        return Ok(None);
    };
    let Some(classic_offset) = le_u32(data, eocd + 16) else {
        return Ok(None);
    };
    let uses_zip64 = classic_count_on_disk == u16::MAX
        || classic_count == u16::MAX
        || classic_size == u32::MAX
        || classic_offset == u32::MAX;
    let consumer_uses_zip64 = classic_count == u16::MAX || classic_offset == u32::MAX;
    if classic_count_on_disk == u16::MAX && !consumer_uses_zip64 {
        // zip 2.4.x does not consult ZIP64 for this sentinel alone and would
        // try to index 65,535 classic entries instead of the validated count.
        return Ok(None);
    }

    let has_zip64_locator = eocd >= 20 && le_u32(data, eocd - 20) == Some(ZIP64_LOCATOR_SIGNATURE);
    if uses_zip64 || has_zip64_locator {
        let zip64 = inspect_zip64_candidates(
            data,
            policy,
            eocd,
            classic_count_on_disk,
            classic_count,
            classic_size,
            classic_offset,
        )?;
        if zip64.is_some() || uses_zip64 {
            return Ok(zip64);
        }
    }

    let Some(directory) = classic_directory(data, eocd) else {
        return Ok(None);
    };
    match inspect_central_directory(data, directory, policy) {
        Ok(preflight) => Ok(Some(preflight)),
        Err(CandidateError::Malformed) => Ok(None),
        Err(CandidateError::Rejected(error)) => Err(error),
    }
}

fn walk_central_directory<F>(
    data: &[u8],
    directory: CentralDirectory,
    mut visit: F,
) -> Result<u64, CandidateError>
where
    F: FnMut(CentralEntry) -> Result<(), CandidateError>,
{
    let mut cursor = directory.start;
    let mut name_bytes = 0u64;
    for _ in 0..directory.entry_count {
        if cursor.checked_add(CENTRAL_DIRECTORY_HEADER_BYTES as usize) > Some(directory.end)
            || le_u32(data, cursor) != Some(CENTRAL_DIRECTORY_HEADER_SIGNATURE)
        {
            return Err(CandidateError::Malformed);
        }
        let name_len = le_u16(data, cursor + 28).ok_or(CandidateError::Malformed)? as usize;
        let extra_len = le_u16(data, cursor + 30).ok_or(CandidateError::Malformed)? as usize;
        let comment_len = le_u16(data, cursor + 32).ok_or(CandidateError::Malformed)? as usize;
        let name_start = cursor
            .checked_add(CENTRAL_DIRECTORY_HEADER_BYTES as usize)
            .ok_or(CandidateError::Malformed)?;
        let next = name_start
            .checked_add(name_len)
            .and_then(|position| position.checked_add(extra_len))
            .and_then(|position| position.checked_add(comment_len))
            .ok_or(CandidateError::Malformed)?;
        if next > directory.end {
            return Err(CandidateError::Malformed);
        }
        name_bytes = name_bytes
            .checked_add(name_len as u64)
            .ok_or(CandidateError::Malformed)?;
        visit(CentralEntry {
            header: cursor,
            name: NameRange {
                start: name_start,
                len: name_len,
            },
            extra: NameRange {
                start: name_start + name_len,
                len: extra_len,
            },
        })?;
        cursor = next;
    }
    if cursor != directory.end {
        // ECMA-376 Part 2, Annex B, Table B.1 requires consumers to support
        // this optional central-directory record and ignore its signature
        // payload. Its bytes remain part of the metadata safety charge.
        if le_u32(data, cursor) != Some(CENTRAL_DIRECTORY_DIGITAL_SIGNATURE) {
            return Err(CandidateError::Malformed);
        }
        let signature_len = le_u16(data, cursor + 4).ok_or(CandidateError::Malformed)? as usize;
        cursor = cursor
            .checked_add(6)
            .and_then(|position| position.checked_add(signature_len))
            .ok_or(CandidateError::Malformed)?;
        if cursor != directory.end {
            return Err(CandidateError::Malformed);
        }
    }
    Ok(name_bytes)
}

fn find_extra_field(data: &[u8], extra: NameRange, wanted_id: u16) -> Option<NameRange> {
    let mut cursor = extra.start;
    let end = cursor.checked_add(extra.len)?;
    while cursor < end {
        let id = le_u16(data, cursor)?;
        let size = le_u16(data, cursor + 2)? as usize;
        let value_start = cursor.checked_add(4)?;
        let next = value_start.checked_add(size)?;
        if next > end {
            return None;
        }
        if id == wanted_id {
            return Some(NameRange {
                start: value_start,
                len: size,
            });
        }
        cursor = next;
    }
    None
}

#[derive(Clone, Copy)]
struct CentralEntryFields {
    flags: u16,
    compression_method: u16,
    crc32: u32,
    compressed_size: u64,
    uncompressed_size: u64,
    relative_local_header: u64,
    zip64_sizes: bool,
}

#[derive(Clone, Copy)]
struct ValidatedEntryIdentity {
    name: NameRange,
    central_header_start: u64,
    local_header_start: u64,
    data_start: u64,
    fields: CentralEntryFields,
}

fn central_entry_fields(data: &[u8], entry: CentralEntry) -> Option<CentralEntryFields> {
    let compressed_32 = le_u32(data, entry.header + 20)?;
    let uncompressed_32 = le_u32(data, entry.header + 24)?;
    let disk_16 = le_u16(data, entry.header + 34)?;
    let offset_32 = le_u32(data, entry.header + 42)?;
    let needs_zip64 = compressed_32 == u32::MAX
        || uncompressed_32 == u32::MAX
        || disk_16 == u16::MAX
        || offset_32 == u32::MAX;
    let mut compressed_size = compressed_32 as u64;
    let mut uncompressed_size = uncompressed_32 as u64;
    let mut relative_local_header = offset_32 as u64;
    let mut disk = disk_16 as u32;

    if needs_zip64 {
        let zip64 = find_extra_field(data, entry.extra, 0x0001)?;
        let mut cursor = zip64.start;
        let end = cursor.checked_add(zip64.len)?;
        let mut take_u64 = || {
            let value = le_u64(data, cursor)?;
            cursor = cursor.checked_add(8)?;
            (cursor <= end).then_some(value)
        };
        if uncompressed_32 == u32::MAX {
            uncompressed_size = take_u64()?;
        }
        if compressed_32 == u32::MAX {
            compressed_size = take_u64()?;
        }
        if offset_32 == u32::MAX {
            relative_local_header = take_u64()?;
        }
        if disk_16 == u16::MAX {
            let value = le_u32(data, cursor)?;
            cursor = cursor.checked_add(4)?;
            if cursor > end {
                return None;
            }
            disk = value;
        }
    }
    if disk != 0 {
        return None;
    }
    Some(CentralEntryFields {
        flags: le_u16(data, entry.header + 8)?,
        compression_method: le_u16(data, entry.header + 10)?,
        crc32: le_u32(data, entry.header + 16)?,
        compressed_size,
        uncompressed_size,
        relative_local_header,
        zip64_sizes: compressed_32 == u32::MAX || uncompressed_32 == u32::MAX,
    })
}

fn local_entry_sizes(
    data: &[u8],
    extra: NameRange,
    compressed: u32,
    uncompressed: u32,
) -> Option<(u64, u64)> {
    if compressed != u32::MAX && uncompressed != u32::MAX {
        return Some((compressed as u64, uncompressed as u64));
    }
    let zip64 = find_extra_field(data, extra, 0x0001)?;
    let mut cursor = zip64.start;
    let end = cursor.checked_add(zip64.len)?;
    let mut take_u64 = || {
        let value = le_u64(data, cursor)?;
        cursor = cursor.checked_add(8)?;
        (cursor <= end).then_some(value)
    };
    let uncompressed = if uncompressed == u32::MAX {
        take_u64()?
    } else {
        uncompressed as u64
    };
    let compressed = if compressed == u32::MAX {
        take_u64()?
    } else {
        compressed as u64
    };
    Some((compressed, uncompressed))
}

fn descriptor_fields_match(
    data: &[u8],
    crc: usize,
    directory_start: usize,
    fields: CentralEntryFields,
) -> bool {
    if le_u32(data, crc) != Some(fields.crc32) {
        return false;
    }
    let sizes = crc + 4;
    if fields.zip64_sizes {
        le_u64(data, sizes) == Some(fields.compressed_size)
            && le_u64(data, sizes + 8) == Some(fields.uncompressed_size)
            && sizes
                .checked_add(16)
                .is_some_and(|end| end <= directory_start)
    } else {
        le_u32(data, sizes) == Some(fields.compressed_size as u32)
            && le_u32(data, sizes + 4) == Some(fields.uncompressed_size as u32)
            && sizes
                .checked_add(8)
                .is_some_and(|end| end <= directory_start)
    }
}

fn signed_descriptor_matches(
    data: &[u8],
    descriptor: usize,
    directory_start: usize,
    fields: CentralEntryFields,
) -> bool {
    le_u32(data, descriptor) == Some(0x0807_4b50)
        && descriptor
            .checked_add(4)
            .is_some_and(|crc| descriptor_fields_match(data, crc, directory_start, fields))
}

/// APPNOTE 4.4.3.2: version needed 4.5 marks ZIP64 format extensions.
const ZIP64_VERSION_NEEDED: u16 = 45;

/// ECMA-376 Part 2, Annex B.2 requires the local and central "version needed
/// to extract" values to agree. Compatibility policy (#1976): a writer that
/// reserves a local ZIP64 field before the sizes are known marks that header
/// 4.5, then fills the real 32-bit sizes and writes a classic central entry.
/// Unzip, Python's `zipfile`, LibreOffice and Word 16.114 on macOS (no repair
/// prompt, zero or matching reserved sizes) read that form. Word also ignores
/// conflicting reserved sizes; keep those excluded as in #1924. The version
/// field selects nothing here (flags, method, CRC and sizes are compared
/// separately), so accept only that explained difference:
/// - the local header says exactly 4.5 and the central entry an older version;
/// - bit 3 is clear, so the data-descriptor policy above is unchanged;
/// - neither header uses a ZIP64 sentinel, so every size and offset comes from
///   the classic fields, which must still equal each other exactly;
/// - the local extra block holds one complete ZIP64 field of the local form
///   (APPNOTE 4.5.3: both sizes, 16 bytes) and each reserved size is zero or
///   equals the central value, so no reader can derive a conflicting size.
fn local_version_needed_matches(
    data: &[u8],
    local: usize,
    local_extra: NameRange,
    entry: CentralEntry,
    fields: CentralEntryFields,
) -> bool {
    let (Some(local_version), Some(central_version)) =
        (le_u16(data, local + 4), le_u16(data, entry.header + 6))
    else {
        return false;
    };
    if local_version == central_version {
        return true;
    }
    let classic_central = [20, 24, 42]
        .into_iter()
        .all(|field| le_u32(data, entry.header + field).is_some_and(|v| v != u32::MAX))
        && le_u16(data, entry.header + 34).is_some_and(|disk| disk != u16::MAX);
    let classic_local = [18, 22]
        .into_iter()
        .all(|field| le_u32(data, local + field).is_some_and(|v| v != u32::MAX));
    if local_version != ZIP64_VERSION_NEEDED
        || central_version >= ZIP64_VERSION_NEEDED
        || fields.flags & (1 << 3) != 0
        || !classic_central
        || !classic_local
    {
        return false;
    }
    let mut cursor = local_extra.start;
    let end = local_extra.start + local_extra.len;
    let mut zip64 = None;
    while cursor < end {
        let (Some(id), Some(size)) = (le_u16(data, cursor), le_u16(data, cursor + 2)) else {
            return false;
        };
        let value = cursor + 4;
        cursor = value + usize::from(size);
        if cursor > end {
            return false;
        }
        if id == ZIP64_EXTRA_FIELD_ID && zip64.replace((value, size)).is_some() {
            return false;
        }
    }
    let Some((value, 16)) = zip64 else {
        return false;
    };
    let reserved_or_equal =
        |at: usize, size: u64| le_u64(data, at).is_some_and(|v| v == 0 || v == size);
    reserved_or_equal(value, fields.uncompressed_size)
        && reserved_or_equal(value + 8, fields.compressed_size)
}

fn validate_local_entry(
    data: &[u8],
    directory: CentralDirectory,
    entry: CentralEntry,
) -> Option<ValidatedEntryIdentity> {
    let fields = central_entry_fields(data, entry)?;
    let local = directory
        .archive_base
        .checked_add(fields.relative_local_header)
        .and_then(|offset| usize::try_from(offset).ok())?;
    if local >= directory.start || le_u32(data, local) != Some(LOCAL_FILE_HEADER_SIGNATURE) {
        return None;
    }
    let local_name_len = le_u16(data, local + 26).map(usize::from)?;
    let local_extra_len = le_u16(data, local + 28).map(usize::from)?;
    let local_name_start = local.checked_add(30)?;
    let local_extra_start = local_name_start.checked_add(local_name_len)?;
    let data_start = local_extra_start.checked_add(local_extra_len)?;
    let local_extra = NameRange {
        start: local_extra_start,
        len: local_extra_len,
    };
    if data_start > directory.start
        || data.get(local_name_start..local_extra_start)
            != data.get(entry.name.start..entry.name.start + entry.name.len)
        || !local_version_needed_matches(data, local, local_extra, entry, fields)
        || le_u16(data, local + 6) != Some(fields.flags)
        || le_u16(data, local + 8) != Some(fields.compression_method)
    {
        return None;
    }

    // ECMA-376 Part 2, Annex B.2 requires duplicated local/central fields to
    // agree, including the DOS timestamp. Some otherwise consistent packages
    // carry a timestamp-only mismatch (#1495). Accept that bounded compatibility
    // case because timestamps do not select or validate the payload; names,
    // flags, compression, CRC, sizes, descriptors, and bounds remain exact.

    let payload_end = data_start.checked_add(usize::try_from(fields.compressed_size).ok()?)?;
    if payload_end > directory.start {
        return None;
    }
    if fields.flags & (1 << 3) != 0 {
        let local_crc = le_u32(data, local + 14)?;
        let local_compressed = le_u32(data, local + 18)?;
        let local_uncompressed = le_u32(data, local + 22)?;
        let canonical_local_fields =
            local_crc == 0 && local_compressed == 0 && local_uncompressed == 0;
        let signed_descriptor =
            signed_descriptor_matches(data, payload_end, directory.start, fields);
        let valid_descriptor = signed_descriptor
            || descriptor_fields_match(data, payload_end, directory.start, fields);
        // ECMA-376 Part 2, Annex B.2/Table B.5 requires zero local fields when
        // bit 3 is set. Compatibility policy (#1428, #1924): accept zero or
        // matching values independently, only for signed descriptors with 32-bit
        // entry sizes.
        // Word 16.114 on macOS opened Word-produced DOCX entries repackaged as
        // Stored/Deflate with all eight zero/matching combinations, without a
        // repair dialog, including a typed empty Deflate entry (CRC/size zero,
        // compressed size populated). This extends the NPOI 2.3.0 form that
        // Excel accepts. It does not establish Office behavior for conflicting
        // values or populated unsigned/ZIP64 headers; keep those excluded.
        // Central sizes still bound the payload and descriptor, and the signed
        // descriptor must agree exactly with the central CRC and both sizes.
        let compatible_populated_local_fields = !fields.zip64_sizes
            && signed_descriptor
            && (local_crc == 0 || local_crc == fields.crc32)
            && (local_compressed == 0 || u64::from(local_compressed) == fields.compressed_size)
            && (local_uncompressed == 0
                || u64::from(local_uncompressed) == fields.uncompressed_size);
        if (!canonical_local_fields && !compatible_populated_local_fields) || !valid_descriptor {
            return None;
        }
    } else {
        let (local_compressed, local_uncompressed) = local_entry_sizes(
            data,
            local_extra,
            le_u32(data, local + 18)?,
            le_u32(data, local + 22)?,
        )?;
        if le_u32(data, local + 14) != Some(fields.crc32)
            || local_compressed != fields.compressed_size
            || local_uncompressed != fields.uncompressed_size
        {
            return None;
        }
    }

    Some(ValidatedEntryIdentity {
        name: entry.name,
        central_header_start: entry.header as u64,
        local_header_start: local as u64,
        data_start: data_start as u64,
        fields,
    })
}

fn validate_central_directory_structure(
    data: &[u8],
    directory: CentralDirectory,
) -> Result<u64, CandidateError> {
    walk_central_directory(data, directory, |entry| {
        validate_local_entry(data, directory, entry)
            .map(|_| ())
            .ok_or(CandidateError::Malformed)
    })
}

fn inspect_central_directory(
    data: &[u8],
    directory: CentralDirectory,
    policy: ItemNamePolicy,
) -> Result<ArchivePreflight, CandidateError> {
    // Establish semantic coherence before latching a resource error. A shallow
    // EOCD-looking byte sequence inside ordinary data or the real EOCD comment
    // must remain inert. This walk borrows input only and does not allocate.
    let name_bytes = validate_central_directory_structure(data, directory)?;

    resource::observe_archive_metadata(directory.entry_count, 0)
        .map_err(CandidateError::Rejected)?;
    if directory.entry_count > HARD_MAX_ARCHIVE_ENTRIES {
        return Err(CandidateError::Rejected(format!(
            "ZIP archive exceeds hard entry-count limit of {HARD_MAX_ARCHIVE_ENTRIES}"
        )));
    }
    let fixed_metadata_bytes = directory
        .size
        .checked_add(directory.footer_metadata_bytes)
        .ok_or(CandidateError::Malformed)?;
    resource::observe_archive_central_directory_bytes(fixed_metadata_bytes)
        .map_err(CandidateError::Rejected)?;
    if fixed_metadata_bytes > HARD_MAX_CENTRAL_DIRECTORY_BYTES {
        return Err(CandidateError::Rejected(format!(
            "ZIP archive exceeds hard central-directory metadata limit of {HARD_MAX_CENTRAL_DIRECTORY_BYTES} bytes"
        )));
    }

    let metadata_bytes = fixed_metadata_bytes
        .checked_add(name_bytes)
        .ok_or(CandidateError::Malformed)?;
    resource::observe_archive_central_directory_bytes(metadata_bytes)
        .map_err(CandidateError::Rejected)?;
    if metadata_bytes > HARD_MAX_CENTRAL_DIRECTORY_BYTES {
        return Err(CandidateError::Rejected(format!(
            "ZIP archive exceeds hard central-directory metadata limit of {HARD_MAX_CENTRAL_DIRECTORY_BYTES} bytes"
        )));
    }

    let capacity = usize::try_from(directory.entry_count).map_err(|_| CandidateError::Malformed)?;
    let mut entries = Vec::with_capacity(capacity);
    let mut exact_names = HashSet::<&[u8]>::with_capacity(capacity);
    // ECMA-376 Part 2 §6.2.2.3: part names are equivalent under ASCII case
    // folding, and (RFC 3986 §6.2.2) a percent-encoded unreserved character
    // equals the character. Keys follow `rels::part_name_equivalence_key`.
    let mut folded_names = HashMap::<String, NameRange>::with_capacity(capacity);
    walk_central_directory(data, directory, |entry| {
        let range = entry.name;
        let name = &data[range.start..range.start + range.len];
        if policy == ItemNamePolicy::ExactUtf8 {
            let flags = le_u16(data, entry.header + 8).ok_or(CandidateError::Malformed)?;
            if std::str::from_utf8(name).is_err() || (!name.is_ascii() && flags & (1 << 11) == 0) {
                return Err(CandidateError::Rejected(
                    "ZIP item names must be UTF-8 and flagged as UTF-8 when not ASCII (APPNOTE 4.4.4)"
                        .to_string(),
                ));
            }
            if !exact_names.insert(name) {
                return Err(CandidateError::Rejected(format!(
                    "ZIP item names must be unique: {}",
                    String::from_utf8_lossy(name)
                )));
            }
        } else if !name.is_ascii() {
            return Err(CandidateError::Rejected(
                "OPC ZIP item names must contain only ASCII bytes (ECMA-376 Part 2, 7.3.3)"
                    .to_string(),
            ));
        }
        if policy == ItemNamePolicy::Opc && !exact_names.insert(name) {
            return Err(CandidateError::Rejected(format!(
                "ZIP item names must be unique: {}",
                String::from_utf8_lossy(name)
            )));
        }
        let represents_file = !name.ends_with(b"/") && !name.ends_with(b"\\");
        if policy == ItemNamePolicy::Opc && represents_file {
            let key = crate::rels::part_name_equivalence_key(&String::from_utf8_lossy(name));
            if let Some(previous) = folded_names.insert(key, range) {
                let previous_name = &data[previous.start..previous.start + previous.len];
                return Err(CandidateError::Rejected(format!(
                    "OOXML part names must be unique under part-name equivalence: {} and {}",
                    String::from_utf8_lossy(previous_name),
                    String::from_utf8_lossy(name)
                )));
            }
        }
        let validated =
            validate_local_entry(data, directory, entry).ok_or(CandidateError::Malformed)?;
        entries.push(PreflightEntryIdentity {
            name: data[validated.name.start..validated.name.start + validated.name.len].into(),
            central_header_start: validated.central_header_start,
            local_header_start: validated.local_header_start,
            data_start: validated.data_start,
            compressed_size: validated.fields.compressed_size,
            uncompressed_size: validated.fields.uncompressed_size,
            crc32: validated.fields.crc32,
            compression_method: validated.fields.compression_method,
            encrypted: validated.fields.flags & 1 != 0,
        });
        Ok(())
    })?;

    // ECMA-376 Part 2, 6.2.2.3: N1 must not be equivalent to N2[S].
    // ZIP item names omit the leading slash, so every internal slash is one
    // possible segment boundary at which a shorter part could end.
    for entry in &entries {
        if policy != ItemNamePolicy::Opc
            || entry.name.ends_with(b"/")
            || entry.name.ends_with(b"\\")
        {
            continue;
        }
        for slash in entry
            .name
            .iter()
            .enumerate()
            .filter_map(|(index, byte)| (*byte == b'/').then_some(index))
        {
            let prefix = String::from_utf8_lossy(&entry.name[..slash]);
            if folded_names.contains_key(&crate::rels::part_name_equivalence_key(&prefix)) {
                return Err(CandidateError::Rejected(format!(
                    "OOXML part name must not be derivable from another part name: {}",
                    String::from_utf8_lossy(&entry.name)
                )));
            }
        }
    }

    Ok(ArchivePreflight {
        entry_count: directory.entry_count,
        entries,
        archive_base: directory.archive_base,
        index_view: IndexView {
            range: directory.start..directory.start,
            rewrites: Vec::new(),
            hidden: Vec::new(),
            footers: Vec::new(),
        },
        zip64_footer: None,
    })
}

/// ZIP64 extended information extra field id (APPNOTE 4.5.3).
const ZIP64_EXTRA_FIELD_ID: u16 = 0x0001;
/// An extra-field id the ZIP reader does not interpret.
const OPAQUE_EXTRA_FIELD_ID: u16 = 0x7a7a;

/// The extra block the ZIP reader sees: the first complete ZIP64 field, then
/// one opaque filler field (or, for fewer than four remaining bytes, zero
/// padding the reader skips). The length is unchanged, the reader keeps one
/// ignored field per record however many fields the block holds, and it
/// cannot reject unrelated or truncated metadata.
fn opaque_extra_block(extra: &[u8]) -> Box<[u8]> {
    let mut block = Vec::with_capacity(extra.len());
    let mut cursor = 0;
    while let (Some(id), Some(len)) = (le_u16(extra, cursor), le_u16(extra, cursor + 2)) {
        let Some(field) = extra.get(cursor..cursor + 4 + len as usize) else {
            break;
        };
        // Only the first ZIP64 field carries the validated sizes and offsets
        // (the preflight reads the same one); repeats join the filler.
        if id == ZIP64_EXTRA_FIELD_ID && block.is_empty() {
            block.extend_from_slice(field);
        }
        cursor += field.len();
    }
    let rest = extra.len() - block.len();
    if rest >= 4 {
        block.extend_from_slice(&OPAQUE_EXTRA_FIELD_ID.to_le_bytes());
        block.extend_from_slice(&((rest - 4) as u16).to_le_bytes());
    }
    block.resize(extra.len(), 0);
    block.into_boxed_slice()
}

/// What a ZIP reader may observe while it builds its index from a package
/// whose footer and central directory a preflight has already validated.
///
/// - Only `range` (the central directory through the EOCD fixed fields) is
///   visible; payloads and the archive comment read as zero, so inert bytes
///   there can never be taken for a footer.
/// - Every central extra field except ZIP64 (which carries the validated
///   sizes and offsets) is opaque. OPC never uses that metadata, and the
///   reader would otherwise reject some APPNOTE-valid encodings (an NTFS
///   attribute sequence, an empty timestamp) and then retry other footers.
/// - Entry comments read as zero, so their bytes are never footer-like.
/// - A fixed-size footer read (EOCD, ZIP64 locator or record) at any other
///   position than the validated ones fails, so the reader can never index an
///   unvalidated directory. Comments read as zero and non-ZIP64 extra ids
///   are remapped, so only a name that itself begins with a footer
///   signature (never a valid OPC part name in practice) can collide; such a
///   package fails closed.
#[derive(Debug, Clone)]
pub struct IndexView {
    range: std::ops::Range<usize>,
    /// Replacement bytes for each directory record's extra block, sorted by
    /// offset: its complete ZIP64 fields, then one opaque filler field.
    rewrites: Vec<(usize, Box<[u8]>)>,
    /// Entry comments: metadata OPC never uses, read as zero while indexing.
    hidden: Vec<std::ops::Range<usize>>,
    footers: Vec<usize>,
}

impl IndexView {
    /// Build the view for a validated `[directory_start, eocd)` central
    /// directory region ending in the EOCD at `eocd`; `zip64_records` are the
    /// validated ZIP64 locator/record offsets, if any.
    pub fn for_validated_directory(
        data: &[u8],
        directory_start: usize,
        eocd: usize,
        zip64_records: &[usize],
    ) -> Self {
        let mut rewrites = Vec::new();
        let mut hidden = Vec::new();
        let mut cursor = directory_start;
        while cursor + 46 <= eocd
            && le_u32(data, cursor) == Some(CENTRAL_DIRECTORY_HEADER_SIGNATURE)
        {
            let (Some(name_len), Some(extra_len), Some(comment_len)) = (
                le_u16(data, cursor + 28),
                le_u16(data, cursor + 30),
                le_u16(data, cursor + 32),
            ) else {
                break;
            };
            let name_start = cursor + 46;
            let extra_start = name_start + name_len as usize;
            let extra_end = (extra_start + extra_len as usize).min(eocd);
            let comment_start = extra_start + extra_len as usize;
            hidden.push(comment_start..comment_start + comment_len as usize);
            if extra_len > 0 && extra_end == extra_start + extra_len as usize {
                rewrites.push((
                    extra_start,
                    opaque_extra_block(&data[extra_start..extra_end]),
                ));
            }
            cursor = extra_start + extra_len as usize + comment_len as usize;
        }
        let mut footers = zip64_records.to_vec();
        footers.push(eocd);
        Self {
            range: directory_start..eocd + 22,
            rewrites,
            hidden,
            footers,
        }
    }

    /// Fill `output` with the view of `data` starting at `start`.
    pub fn read(&self, data: &[u8], start: usize, output: &mut [u8]) -> std::io::Result<usize> {
        let start = start.min(data.len());
        let count = output.len().min(data.len() - start);
        let end = start + count;
        let footer_block = match output.len() {
            22 => Some(EOCD_SIGNATURE),
            20 => Some(ZIP64_LOCATOR_SIGNATURE),
            56 => Some(ZIP64_EOCD_SIGNATURE),
            _ => None,
        };
        output[..count].fill(0);
        let visible_start = start.max(self.range.start);
        let visible_end = end.min(self.range.end);
        if visible_start < visible_end {
            output[visible_start - start..visible_end - start]
                .copy_from_slice(&data[visible_start..visible_end]);
        }
        let first = self
            .rewrites
            .partition_point(|(at, bytes)| at + bytes.len() <= start);
        for (at, bytes) in self.rewrites[first..]
            .iter()
            .take_while(|(at, _)| *at < end)
        {
            let from = (*at).max(start);
            let to = (at + bytes.len()).min(end);
            output[from - start..to - start].copy_from_slice(&bytes[from - at..to - at]);
        }
        let first_hidden = self.hidden.partition_point(|range| range.end <= start);
        for range in self.hidden[first_hidden..]
            .iter()
            .take_while(|range| range.start < end)
        {
            let hidden_start = range.start.max(start);
            let hidden_end = range.end.min(end);
            if hidden_start < hidden_end {
                output[hidden_start - start..hidden_end - start].fill(0);
            }
        }
        // Judge what the reader would see: hidden bytes are never footers.
        if footer_block.is_some_and(|signature| le_u32(output, 0) == Some(signature))
            && !self.footers.contains(&start)
        {
            return Err(std::io::Error::other("unvalidated ZIP footer record"));
        }
        Ok(count)
    }
}

/// The reader handed to the ZIP crate. While it builds its index (the only
/// time it reads footers and the central directory), just `index_range` (the
/// validated directory, any ZIP64 footer records and the EOCD fixed fields)
/// is visible: every other byte, payloads and the archive comment included,
/// reads as zero and the view's rewrites apply. The crate therefore sees exactly
/// one footer and directory, so it cannot retry an unvalidated EOCD, and NTFS
/// metadata it would reject stays opaque. Afterwards every read passes
/// through to the real bytes unchanged; lengths and offsets never differ.
#[derive(Debug)]
pub struct PackageReader<T> {
    cursor: std::io::Cursor<T>,
    view: IndexView,
    /// Shared with the opener, which clears it once the index is built.
    indexing: std::rc::Rc<std::cell::Cell<bool>>,
}

impl<T: AsRef<[u8]>> PackageReader<T> {
    /// A reader that shows `view` until the returned flag is cleared, then
    /// the real bytes. For ZIP readers outside this module that validate an
    /// embedded package with their own preflight.
    pub fn with_view(data: T, view: IndexView) -> (Self, std::rc::Rc<std::cell::Cell<bool>>) {
        let indexing = std::rc::Rc::new(std::cell::Cell::new(true));
        (
            Self {
                cursor: std::io::Cursor::new(data),
                view,
                indexing: std::rc::Rc::clone(&indexing),
            },
            indexing,
        )
    }
}

impl<T: AsRef<[u8]>> std::io::Read for PackageReader<T> {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        if !self.indexing.get() {
            return std::io::Read::read(&mut self.cursor, output);
        }
        let data = self.cursor.get_ref().as_ref();
        let start = usize::try_from(self.cursor.position()).unwrap_or(usize::MAX);
        let count = self.view.read(data, start, output)?;
        self.cursor
            .set_position((start.min(data.len()) + count) as u64);
        Ok(count)
    }
}

impl<T: AsRef<[u8]>> std::io::Seek for PackageReader<T> {
    fn seek(&mut self, position: std::io::SeekFrom) -> std::io::Result<u64> {
        std::io::Seek::seek(&mut self.cursor, position)
    }
}

/// OPC-policy preflight (tests).
#[cfg(test)]
fn preflight_archive_limits(data: &[u8]) -> Result<ArchivePreflight, String> {
    preflight_archive_with_policy(data, ItemNamePolicy::Opc)
}

/// Validate the legal EOCD tail and walk its central directory before
/// `zip::ZipArchive::new` allocates its filename index. Malformed/truncated
/// metadata is left to the ZIP crate's ordinary corruption path rather than
/// being mislabeled as a resource violation.
fn preflight_archive_with_policy(
    data: &[u8],
    policy: ItemNamePolicy,
) -> Result<ArchivePreflight, String> {
    inspect_archive(data, policy)?
        .ok_or_else(|| "ZIP central directory preflight failed".to_string())
}

#[cfg(test)]
fn inspect_archive_opc(data: &[u8]) -> Result<Option<ArchivePreflight>, String> {
    inspect_archive(data, ItemNamePolicy::Opc)
}

fn inspect_archive(
    data: &[u8],
    policy: ItemNamePolicy,
) -> Result<Option<ArchivePreflight>, String> {
    const MAX_EOCD_SEARCH: usize = 22 + u16::MAX as usize;

    if data.len() < 22 {
        return Ok(None);
    }
    let start = data.len().saturating_sub(MAX_EOCD_SEARCH);
    let mut resolved = None;
    for relative in data[start..]
        .windows(4)
        .enumerate()
        .filter_map(|(index, bytes)| (bytes == EOCD_SIGNATURE.to_le_bytes()).then_some(index))
    {
        let eocd = start + relative;
        let Some(comment_end) = le_u16(data, eocd + 20)
            .and_then(|comment_len| eocd.checked_add(22 + comment_len as usize))
        else {
            continue;
        };
        if comment_end != data.len() {
            continue;
        }
        if let Some(preflight) = inspect_eocd_candidate(data, eocd, policy)? {
            if resolved.replace((eocd, preflight)).is_some() {
                return Err("ambiguous ZIP end-of-central-directory records".to_string());
            }
        }
    }
    let Some((eocd, mut preflight)) = resolved else {
        return Ok(None);
    };
    // Only records validated by the ZIP64 path are footers besides the EOCD.
    let footers: Vec<usize> = preflight
        .zip64_footer
        .map(|(locator, record)| vec![locator, record])
        .unwrap_or_default();
    preflight.index_view =
        IndexView::for_validated_directory(data, preflight.index_view.range.start, eocd, &footers);
    Ok(Some(preflight))
}

/// Enforce the OPC package-wide part identity constraint before consumers can
/// observe the ZIP index. ECMA-376 Part 2, 6.2.2 requires unique part identity,
/// and 7.3.3 requires ZIP item names to be unique.
///
/// Run the bounded preflight and ensure the ZIP crate exposed exactly the same
/// ordered raw names. Callers that already own an [`ArchivePreflight`] should
/// use its method directly to avoid repeating validation.
#[cfg(test)]
fn validate_archive_item_names<R: std::io::Read + std::io::Seek>(
    data: &[u8],
    archive: &mut zip::ZipArchive<R>,
) -> Result<(), String> {
    preflight_archive_limits(data)?.validate_archive_item_names(archive)
}

/// Open a ZIP cursor through the exact archive-offset interpretation proven by
/// preflight, then revalidate the complete consumer index before returning it.
///
/// This is the single construction path for both legacy owned archives and
/// package sessions (`open_validated_slice` / `open_validated_zip`). It deliberately returns raw container errors; format
/// facades own their historical user-visible error context.
pub(crate) fn open_validated_slice(
    data: &[u8],
) -> Result<zip::ZipArchive<PackageReader<&[u8]>>, String> {
    open_validated_slice_with_policy(data, ItemNamePolicy::Opc)
}

/// [`open_validated_slice`] under an explicit item-name policy.
pub(crate) fn open_validated_slice_with_policy(
    data: &[u8],
    policy: ItemNamePolicy,
) -> Result<zip::ZipArchive<PackageReader<&[u8]>>, String> {
    let preflight = preflight_archive_with_policy(data, policy)?;
    open_preflighted(preflight, data)
}

fn open_preflighted<T: AsRef<[u8]>>(
    preflight: ArchivePreflight,
    data: T,
) -> Result<zip::ZipArchive<PackageReader<T>>, String> {
    let indexing = std::rc::Rc::new(std::cell::Cell::new(true));
    let reader = PackageReader {
        cursor: std::io::Cursor::new(data),
        view: preflight.index_view.clone(),
        indexing: std::rc::Rc::clone(&indexing),
    };
    let mut archive = zip::ZipArchive::with_config(preflight.read_config(), reader)
        .map_err(|error| error.to_string())?;
    // The index is built; every later read (local headers, payloads) sees
    // the real bytes.
    indexing.set(false);
    preflight.validate_archive_item_names(&mut archive)?;
    validate_archive_limits(&mut archive)?;
    Ok(archive)
}

/// Open an owned ZIP through the shared validated construction path (tests;
/// production opens borrowed bytes through `open_validated_slice`).
#[cfg(test)]
fn open_validated_zip(data: Vec<u8>) -> Result<zip::ZipArchive<PackageReader<Vec<u8>>>, String> {
    let preflight = preflight_archive_limits(&data)?;
    open_preflighted(preflight, data)
}

/// Preserve the machine-readable resource envelope while adding the historical
/// container context to ordinary ZIP errors.
pub fn tag_container_error(error: String) -> String {
    if error.starts_with("OOXML_RESOURCE_LIMIT:") {
        error
    } else {
        format!("(zip container): {error}")
    }
}

/// Create an ephemeral governor for one free-function operation.
pub fn scoped_limits(
    format: OoxmlFormat,
    operation: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> ResourceScope {
    ResourceGovernor::from_wasm(
        format,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        None,
    )
    .scope(operation)
}

/// Record accessible central-directory metadata and enforce hard entry count.
/// Declared bytes are diagnostic/early-entry evidence, never treated as proof of
/// actual total inflation.
fn validate_archive_limits<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<(), String> {
    resource::assert_healthy()?;
    let mut declared_total = 0u64;
    for index in 0..archive.len() {
        let entry = archive
            .by_index(index)
            .map_err(|e| format!("ZIP archive entry metadata error: {e}"))?;
        declared_total = declared_total.saturating_add(entry.size());
    }
    resource::observe_archive_metadata(archive.len() as u64, declared_total)
}

/// Inflate one entry from its raw stored bytes. Declared sizes only seed the
/// resource allowance; actual output is charged by `read_entry` until EOF,
/// where the CRC-32 is verified (APPNOTE 4.4.7), as the retained session's
/// decoder does. The ZIP crate's own decompressing reader is not used because
/// it rejects output past the declared size before that output is charged,
/// which would hide an actual-inflation limit crossing.
fn open_entry_content<'a, R: std::io::Read + std::io::Seek>(
    archive: &'a mut zip::ZipArchive<R>,
    part_id: usize,
    path: &str,
) -> Result<(CrcCheckedEntry<Box<dyn std::io::Read + 'a>>, u64), String> {
    let raw = archive
        .by_index_raw(part_id)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    if raw.encrypted() {
        return Err(format!(
            "entry not found: {path}: unsupported Zip archive: Password required to decrypt file"
        ));
    }
    let declared_size = raw.size();
    let expected_crc = raw.crc32();
    let content: Box<dyn std::io::Read + 'a> = match raw.compression() {
        zip::CompressionMethod::Stored => Box::new(raw),
        zip::CompressionMethod::Deflated => Box::new(flate2::read::DeflateDecoder::new(raw)),
        _ => {
            return Err(format!(
                "entry not found: {path}: unsupported Zip archive: Compression method not supported"
            ))
        }
    };
    Ok((
        CrcCheckedEntry {
            inner: content,
            hasher: crc32fast::Hasher::new(),
            expected: expected_crc,
        },
        declared_size,
    ))
}

/// Verifies the entry CRC-32 when the inflated stream ends.
struct CrcCheckedEntry<R> {
    inner: R,
    hasher: crc32fast::Hasher,
    expected: u32,
}

impl<R: std::io::Read> std::io::Read for CrcCheckedEntry<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let count = self.inner.read(buf)?;
        if count == 0 && !buf.is_empty() && self.hasher.clone().finalize() != self.expected {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "Invalid checksum",
            ));
        }
        self.hasher.update(&buf[..count]);
        Ok(count)
    }
}

fn read_entry<R: std::io::Read>(
    entry: &mut R,
    part_id: usize,
    path: &str,
    declared_size: u64,
    requested_prefix: Option<u64>,
) -> Result<Vec<u8>, String> {
    resource::assert_healthy()?;
    let resource_allowance = resource::read_allowance(part_id, path, declared_size)?;
    let caller_limit = requested_prefix.unwrap_or(u64::MAX);
    let ordinary_limit = resource_allowance.min(caller_limit);
    let resource_bound_is_tighter = resource_allowance < caller_limit;
    let mut bytes = Vec::with_capacity(initial_reserve(declared_size, ordinary_limit));
    let mut observed = 0u64;
    let mut chunk = [0u8; READ_CHUNK_BYTES];

    loop {
        if observed == ordinary_limit {
            if !resource_bound_is_tighter {
                break;
            }
            // Read exactly one more byte to distinguish EOF-at-limit from a
            // proven policy crossing. Recording it latches limit+1.
            let count = entry
                .read(&mut chunk[..1])
                .map_err(|e| format!("read error: {e}"))?;
            if count == 0 {
                break;
            }
            observed = observed.saturating_add(count as u64);
            resource::observe_inflated(part_id, path, observed, count as u64)?;
            unreachable!("resource allowance + 1 must reject");
        }

        let remaining = ordinary_limit - observed;
        let count = entry
            .read(&mut chunk[..remaining.min(READ_CHUNK_BYTES as u64) as usize])
            .map_err(|e| format!("read error: {e}"))?;
        if count == 0 {
            break;
        }
        observed = observed.saturating_add(count as u64);
        // Charge each successful decompressor delivery immediately so bytes
        // emitted before a later CRC/read failure remain accounted.
        resource::observe_inflated(part_id, path, observed, count as u64)?;
        bytes.extend_from_slice(&chunk[..count]);
    }
    Ok(bytes)
}

/// Open one package and extract a single entry under an ephemeral resource
/// session. Stateful browser paths use their retained archive governor instead.
pub fn extract_zip_entry(
    data: &[u8],
    path: &str,
    format: OoxmlFormat,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, String> {
    let _scope = scoped_limits(
        format,
        "extract",
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    );
    let mut archive = open_validated_slice(data).map_err(|error| {
        if error.starts_with("OOXML_RESOURCE_LIMIT:") {
            error
        } else {
            format!("zip open error: {error}")
        }
    })?;
    read_zip_bytes(&mut archive, path)
}

/// Index of the entry naming the same part as `path`: the exact ZIP item name
/// first, else the item equivalent under ECMA-376 Part 2 §6.2.2.3 (see
/// [`crate::rels::part_name_equivalence_key`]). Validation rejects two items
/// with one key, so the match is unique. This linear fallback serves the
/// one-shot archive helpers; retained package sessions keep a key index.
fn equivalent_entry_index<R: std::io::Read + std::io::Seek>(
    archive: &zip::ZipArchive<R>,
    path: &str,
) -> Option<usize> {
    archive.index_for_name(path).or_else(|| {
        let key = crate::rels::part_name_equivalence_key(path);
        archive
            .file_names()
            .filter_map(Result::ok)
            .find(|name| crate::rels::part_name_equivalence_key(name) == key)
            .and_then(|name| archive.index_for_name(&name))
    })
}

pub fn read_zip_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<u8>, String> {
    resource::assert_healthy()?;
    let part_id =
        equivalent_entry_index(archive, path).ok_or_else(|| format!("entry not found: {path}"))?;
    let (mut entry, declared_size) = open_entry_content(archive, part_id, path)?;
    read_entry(&mut entry, part_id, path, declared_size, None)
}

pub fn read_zip_string<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<String, String> {
    String::from_utf8(read_zip_bytes(archive, path)?).map_err(|e| format!("read error: {e}"))
}

/// Read a deliberate UTF-8 prefix. A caller prefix does not inspect the next
/// byte; a tighter resource bound does, so it cannot silently truncate.
pub fn read_zip_string_head<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    max_bytes: usize,
) -> Result<String, String> {
    resource::assert_healthy()?;
    let part_id =
        equivalent_entry_index(archive, path).ok_or_else(|| format!("entry not found: {path}"))?;
    let (mut entry, declared_size) = open_entry_content(archive, part_id, path)?;
    let mut bytes = read_entry(
        &mut entry,
        part_id,
        path,
        declared_size,
        Some(max_bytes as u64),
    )?;
    match std::str::from_utf8(&bytes) {
        Ok(text) => Ok(text.to_owned()),
        Err(error) if error.error_len().is_none() => {
            bytes.truncate(error.valid_up_to());
            Ok(String::from_utf8(bytes).expect("validated UTF-8 prefix"))
        }
        Err(error) => Err(format!("read error: {error}")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Read, Write};

    fn archive_with(name: &str, body: &[u8]) -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(body).unwrap();
            writer.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(bytes)).unwrap()
    }

    fn archive_with_two() -> zip::ZipArchive<Cursor<Vec<u8>>> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"1234").unwrap();
            writer.start_file("word/b.xml", options).unwrap();
            writer.write_all(b"5678").unwrap();
            writer.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(bytes)).unwrap()
    }

    fn details(error: &str) -> serde_json::Value {
        serde_json::from_str(
            error
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("typed resource envelope"),
        )
        .unwrap()
    }

    fn last_signature(bytes: &[u8], signature: u32) -> usize {
        bytes
            .windows(4)
            .rposition(|window| window == signature.to_le_bytes())
            .expect("signature")
    }

    fn zip_bytes(zip64: bool) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            if zip64 {
                writer.set_raw_zip64_extensible_data_sector(
                    b"metadata PK\x06\x06 with a false signature"[..].into(),
                );
            }
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"a").unwrap();
            writer.start_file("word/b.xml", options).unwrap();
            writer.write_all(b"b").unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    fn archive_bytes_with_single(name: &str, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(body).unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    fn prepend(bytes: &[u8]) -> Vec<u8> {
        let mut prefixed = b"MZ\x90\0self-extracting-prefix".to_vec();
        prefixed.extend_from_slice(bytes);
        prefixed
    }

    fn add_central_directory_signature(mut bytes: Vec<u8>, zip64: bool) -> Vec<u8> {
        let payload = b"ignored-signature-payload";
        let mut signature = CENTRAL_DIRECTORY_DIGITAL_SIGNATURE.to_le_bytes().to_vec();
        signature.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        signature.extend_from_slice(payload);

        let insertion = if zip64 {
            let locator = last_signature(&bytes, ZIP64_LOCATOR_SIGNATURE);
            le_u64(&bytes, locator + 8).unwrap() as usize
        } else {
            last_signature(&bytes, EOCD_SIGNATURE)
        };
        bytes.splice(insertion..insertion, signature.iter().copied());

        if zip64 {
            let zip64_eocd = insertion + signature.len();
            let size = le_u64(&bytes, zip64_eocd + 40).unwrap() + signature.len() as u64;
            bytes[zip64_eocd + 40..zip64_eocd + 48].copy_from_slice(&size.to_le_bytes());
            let locator = last_signature(&bytes, ZIP64_LOCATOR_SIGNATURE);
            let relative = le_u64(&bytes, locator + 8).unwrap() + signature.len() as u64;
            bytes[locator + 8..locator + 16].copy_from_slice(&relative.to_le_bytes());
            let eocd = last_signature(&bytes, EOCD_SIGNATURE);
            let classic_size = le_u32(&bytes, eocd + 12).unwrap();
            if classic_size != u32::MAX {
                bytes[eocd + 12..eocd + 16]
                    .copy_from_slice(&(classic_size + signature.len() as u32).to_le_bytes());
            }
        } else {
            let eocd = insertion + signature.len();
            let size = le_u32(&bytes, eocd + 12).unwrap() + signature.len() as u32;
            bytes[eocd + 12..eocd + 16].copy_from_slice(&size.to_le_bytes());
        }
        bytes
    }

    fn synthetic_entries(count: usize, name_len: usize) -> (Vec<u8>, u32, u32) {
        let mut bytes = Vec::new();
        let mut local_offsets = Vec::with_capacity(count);
        for _ in 0..count {
            local_offsets.push(bytes.len() as u32);
            let local = bytes.len();
            bytes.resize(local + 30, 0);
            bytes[local..local + 4].copy_from_slice(&LOCAL_FILE_HEADER_SIGNATURE.to_le_bytes());
            bytes[local + 26..local + 28].copy_from_slice(&(name_len as u16).to_le_bytes());
            bytes.resize(bytes.len() + name_len, b'a');
        }
        let central_offset = bytes.len() as u32;
        for local_offset in local_offsets {
            let central = bytes.len();
            bytes.resize(central + CENTRAL_DIRECTORY_HEADER_BYTES as usize, 0);
            bytes[central..central + 4]
                .copy_from_slice(&CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes());
            bytes[central + 28..central + 30].copy_from_slice(&(name_len as u16).to_le_bytes());
            bytes[central + 42..central + 46].copy_from_slice(&local_offset.to_le_bytes());
            bytes.resize(bytes.len() + name_len, b'a');
        }
        let central_size = bytes.len() as u32 - central_offset;
        (bytes, central_offset, central_size)
    }

    fn append_classic_eocd(
        bytes: &mut Vec<u8>,
        central_offset: u32,
        central_size: u32,
        count: u16,
        comment: &[u8],
    ) {
        bytes.extend_from_slice(&EOCD_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&count.to_le_bytes());
        bytes.extend_from_slice(&count.to_le_bytes());
        bytes.extend_from_slice(&central_size.to_le_bytes());
        bytes.extend_from_slice(&central_offset.to_le_bytes());
        bytes.extend_from_slice(&(comment.len() as u16).to_le_bytes());
        bytes.extend_from_slice(comment);
    }

    fn stored_descriptor_zip(body: &[u8]) -> Vec<u8> {
        const NAME: &[u8] = b"word/document.xml";
        let crc = crc32fast::hash(body);
        let size = body.len() as u32;
        let mut bytes = Vec::new();

        bytes.extend_from_slice(&LOCAL_FILE_HEADER_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&20u16.to_le_bytes());
        bytes.extend_from_slice(&(1u16 << 3).to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&(NAME.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(NAME);
        bytes.extend_from_slice(body);
        bytes.extend_from_slice(&0x0807_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&crc.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());

        let central_offset = bytes.len() as u32;
        bytes.extend_from_slice(&CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&20u16.to_le_bytes());
        bytes.extend_from_slice(&20u16.to_le_bytes());
        bytes.extend_from_slice(&(1u16 << 3).to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&crc.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&(NAME.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(NAME);
        let central_size = bytes.len() as u32 - central_offset;
        append_classic_eocd(&mut bytes, central_offset, central_size, 1, &[]);
        bytes
    }

    fn populated_deflate_descriptor_zip(body: &[u8]) -> Vec<u8> {
        const NAME: &str = "xl/worksheets/sheet1.xml";
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    NAME,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            writer.write_all(body).unwrap();
            writer.finish().unwrap();
        }

        let eocd = last_signature(&bytes, EOCD_SIGNATURE);
        let central = le_u32(&bytes, eocd + 16).unwrap() as usize;
        let crc = le_u32(&bytes, central + 16).unwrap();
        let compressed = le_u32(&bytes, central + 20).unwrap();
        let uncompressed = le_u32(&bytes, central + 24).unwrap();

        bytes[6..8].copy_from_slice(&(1u16 << 3).to_le_bytes());
        bytes[central + 8..central + 10].copy_from_slice(&(1u16 << 3).to_le_bytes());
        let mut descriptor = 0x0807_4b50u32.to_le_bytes().to_vec();
        descriptor.extend_from_slice(&crc.to_le_bytes());
        descriptor.extend_from_slice(&compressed.to_le_bytes());
        descriptor.extend_from_slice(&uncompressed.to_le_bytes());
        bytes.splice(central..central, descriptor);

        let shifted_eocd = eocd + 16;
        bytes[shifted_eocd + 16..shifted_eocd + 20]
            .copy_from_slice(&((central + 16) as u32).to_le_bytes());
        bytes
    }

    fn empty_zip64_with_footer_bytes(footer_bytes: usize) -> Vec<u8> {
        assert!(footer_bytes >= 56);
        let record_size = footer_bytes as u64 - 12;
        let mut bytes = Vec::with_capacity(footer_bytes + 42);
        bytes.extend_from_slice(&ZIP64_EOCD_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&record_size.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.resize(footer_bytes, 0);
        bytes.extend_from_slice(&ZIP64_LOCATOR_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());
        append_classic_eocd(&mut bytes, 0, 0, 0, &[]);
        bytes
    }

    fn replace_raw_name(bytes: &mut [u8], replacement: &[u8]) {
        const NAME: &[u8] = b"word/a.xml";
        assert_eq!(replacement.len(), NAME.len());
        let mut replacements = 0;
        for start in 0..=bytes.len() - NAME.len() {
            if &bytes[start..start + NAME.len()] == NAME {
                bytes[start..start + NAME.len()].copy_from_slice(replacement);
                replacements += 1;
            }
        }
        assert_eq!(replacements, 2, "local and central names");
    }

    #[test]
    fn extracts_by_path_and_reports_missing() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "ppt/media/image1.png",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"\x89PNGdata").unwrap();
            writer.finish().unwrap();
        }
        assert_eq!(
            extract_zip_entry(
                &bytes,
                "ppt/media/image1.png",
                OoxmlFormat::Pptx,
                Some(64),
                Some(64)
            )
            .unwrap(),
            b"\x89PNGdata"
        );
        assert!(
            extract_zip_entry(&bytes, "missing", OoxmlFormat::Pptx, Some(64), Some(64))
                .unwrap_err()
                .contains("not found")
        );
    }

    #[test]
    fn declared_entry_limit_is_typed_and_poisoned() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(4), Some(64), None);
        let _scope = governor.scope("parse");
        let mut archive = archive_with("word/document.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        let first = read_zip_bytes(&mut archive, "word/document.xml").unwrap_err();
        let json = details(&first);
        assert_eq!(json["code"], "ooxml-resource-limit");
        assert_eq!(json["details"]["stage"], "container");
        assert_eq!(
            json["details"]["violation"]["metric"],
            "declared-inflated-bytes"
        );
        assert_eq!(
            read_zip_bytes(&mut archive, "word/document.xml").unwrap_err(),
            first
        );
    }

    #[test]
    fn forged_small_declaration_is_stopped_by_actual_output() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "word/document.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"12345678").unwrap();
            writer.finish().unwrap();
        }
        // Forge both local and central uncompressed-size fields down to one.
        // The payload still delivers eight bytes, so metadata alone must not
        // authorize the read: the actual-inflation limit stops it, typed.
        bytes[22..26].copy_from_slice(&1u32.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("central-directory header");
        bytes[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(4), Some(64), None);
        let _scope = governor.scope("parse");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        validate_archive_limits(&mut archive).unwrap();
        let error = read_zip_bytes(&mut archive, "word/document.xml").unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], 4);
        assert_eq!(violation["observed"], 5);
    }

    #[test]
    fn eocd_signature_in_the_archive_comment_cannot_select_a_decoy_directory() {
        // The comment holds a decoy EOCD describing one directory entry at
        // offset 0 (a local header), which the reader alone would try first.
        let mut bytes = zip_bytes(false);
        let eocd = last_signature(&bytes, EOCD_SIGNATURE);
        let mut decoy = EOCD_SIGNATURE.to_le_bytes().to_vec();
        decoy.extend_from_slice(&[0, 0, 0, 0, 1, 0, 1, 0]);
        decoy.extend_from_slice(&46u32.to_le_bytes());
        decoy.extend_from_slice(&0u32.to_le_bytes());
        decoy.extend_from_slice(&0u16.to_le_bytes());
        bytes[eocd + 20..eocd + 22].copy_from_slice(&(decoy.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&decoy);
        let mut archive = open_validated_zip(bytes.clone()).expect("validated directory");
        assert_eq!(archive.len(), 2);
        assert_eq!(read_zip_bytes(&mut archive, "word/b.xml").unwrap(), b"b");
        let borrowed = open_validated_slice(&bytes).expect("validated directory");
        assert_eq!(borrowed.len(), 2);
    }

    #[test]
    fn index_view_exposes_only_the_validated_directory_until_indexing_ends() {
        // A stored entry whose payload holds footer signatures, plus a stray
        // EOCD signature inside a directory record's file comment.
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let stored = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer.start_file("word/a.xml", stored).unwrap();
            writer
                .write_all(b"PK\x05\x06 PK\x06\x06 PK\x06\x07")
                .unwrap();
            writer.finish().unwrap();
        }
        let preflight = preflight_archive_limits(&bytes).unwrap();
        let payload = bytes.windows(4).position(|w| w == b"PK\x05\x06").unwrap();
        let indexing = std::rc::Rc::new(std::cell::Cell::new(true));
        let mut reader = PackageReader {
            cursor: Cursor::new(bytes.as_slice()),
            view: preflight.index_view.clone(),
            indexing: std::rc::Rc::clone(&indexing),
        };
        let mut seen = Vec::new();
        reader.read_to_end(&mut seen).unwrap();
        assert_eq!(seen.len(), bytes.len());
        assert!(seen[..preflight.index_view.range.start]
            .iter()
            .all(|byte| *byte == 0));
        assert_eq!(
            seen.windows(4)
                .filter(|w| *w == EOCD_SIGNATURE.to_le_bytes())
                .count(),
            1,
            "only the validated EOCD is visible while indexing"
        );
        assert_eq!(&seen[payload..payload + 4], &[0, 0, 0, 0]);

        indexing.set(false);
        std::io::Seek::seek(&mut reader, std::io::SeekFrom::Start(0)).unwrap();
        let mut real = Vec::new();
        reader.read_to_end(&mut real).unwrap();
        assert_eq!(real, bytes, "after indexing every read is byte-exact");

        let mut archive = open_validated_zip(bytes).unwrap();
        assert_eq!(
            read_zip_bytes(&mut archive, "word/a.xml").unwrap(),
            b"PK\x05\x06 PK\x06\x06 PK\x06\x07"
        );
    }

    #[test]
    fn opaque_metadata_the_reader_would_reject_does_not_reach_payload_footers() {
        // A zero-length Extended Timestamp passes preflight; the reader would
        // reject it and then try the footer of the complete inner archive
        // stored as a payload. It stays opaque, so the validated directory
        // opens (as it did before the reader upgrade) and only it is indexed.
        let mut inner = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut inner));
            writer
                .start_file("x.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let mut outer = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut outer));
            let stored = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer.start_file("word/a.xml", stored).unwrap();
            writer.write_all(&inner).unwrap();
            writer.finish().unwrap();
        }
        let timestamp = [0x55, 0x54, 0, 0];
        let bytes = with_central_extra(&outer, &timestamp);
        assert!(preflight_archive_limits(&bytes).is_ok());
        let mut archive = open_validated_zip(bytes).expect("opaque timestamp metadata");
        assert_eq!(archive.len(), 1);
        assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), inner);
    }

    #[test]
    fn signature_valued_metadata_fields_still_index() {
        for signature in [
            EOCD_SIGNATURE,
            ZIP64_EOCD_SIGNATURE,
            ZIP64_LOCATOR_SIGNATURE,
        ] {
            let mut bytes = zip_bytes(false);
            // Forge the first entry's CRC (local and central) to a footer
            // signature value; the index must still open unchanged.
            bytes[14..18].copy_from_slice(&signature.to_le_bytes());
            let eocd = last_signature(&bytes, EOCD_SIGNATURE);
            let central = le_u32(&bytes, eocd + 16).unwrap() as usize;
            bytes[central + 16..central + 20].copy_from_slice(&signature.to_le_bytes());
            let archive = open_validated_zip(bytes).expect("signature-valued CRC indexes");
            assert_eq!(archive.len(), 2);
        }
    }

    #[test]
    fn entry_comments_shaped_like_footer_records_still_open() {
        for comment in [
            [&b"PK\x05\x06"[..], &[b'x'; 18][..]].concat(),
            [&b"PK\x06\x06"[..], &[b'x'; 52][..]].concat(),
            [&b"PK\x06\x07"[..], &[b'x'; 16][..]].concat(),
        ] {
            let mut bytes = Vec::new();
            {
                let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
                let stored = zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Stored);
                writer.start_file("word/a.xml", stored).unwrap();
                writer.write_all(b"a").unwrap();
                writer.finish().unwrap();
            }
            // Splice an entry comment into the only central record.
            let eocd = last_signature(&bytes, EOCD_SIGNATURE);
            let central = le_u32(&bytes, eocd + 16).unwrap() as usize;
            let name_len = le_u16(&bytes, central + 28).unwrap() as usize;
            let extra_len = le_u16(&bytes, central + 30).unwrap() as usize;
            let at = central + 46 + name_len + extra_len;
            let mut out = bytes[..at].to_vec();
            out.extend_from_slice(&comment);
            out.extend_from_slice(&bytes[at..]);
            out[central + 32..central + 34].copy_from_slice(&(comment.len() as u16).to_le_bytes());
            let eocd = eocd + comment.len();
            let size = le_u32(&out, eocd + 12).unwrap() + comment.len() as u32;
            out[eocd + 12..eocd + 16].copy_from_slice(&size.to_le_bytes());
            let mut archive = open_validated_zip(out.clone()).expect("entry comment is metadata");
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"a");
            assert!(open_validated_slice(&out).is_ok());
        }
    }

    #[test]
    fn a_footer_shaped_entry_name_fails_closed_instead_of_steering_the_index() {
        // A 22-byte name that begins with an EOCD signature: the reader's
        // fixed-size read there must never be accepted as a footer.
        let name = "PK\u{5}\u{6}aaaaaaaaaaaaaaaaaa";
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"a").unwrap();
            writer.finish().unwrap();
        }
        let view = preflight_archive_limits(&bytes).unwrap().index_view;
        let name_at =
            le_u32(&bytes, last_signature(&bytes, EOCD_SIGNATURE) + 16).unwrap() as usize + 46;
        let mut block = [0u8; 22];
        assert!(view.read(&bytes, name_at, &mut block).is_err());
    }

    #[test]
    fn unused_extra_metadata_is_tolerated_and_coalesced_for_the_index() {
        // A truncated Extended Timestamp (declares 9 bytes, carries 1).
        let truncated = [0x55, 0x54, 9, 0, 1];
        let bytes = with_central_extra(&zip_bytes(false), &truncated);
        let mut archive = open_validated_zip(bytes).expect("truncated unused metadata");
        assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"a");

        // Many short unknown fields become one opaque field for the reader.
        let many: Vec<u8> = (0..2_000u16)
            .flat_map(|id| [(id | 0x8000).to_le_bytes(), [0, 0]].concat())
            .collect();
        let bytes = with_central_extra(&zip_bytes(false), &many);
        let preflight = preflight_archive_limits(&bytes).unwrap();
        assert_eq!(preflight.index_view.rewrites.len(), 1);
        let mut archive = open_validated_zip(bytes).expect("many unknown fields");
        assert_eq!(read_zip_bytes(&mut archive, "word/b.xml").unwrap(), b"b");
    }

    #[test]
    fn truncated_unicode_fields_and_repeated_zip64_fields_stay_opaque() {
        for truncated in [[0x75, 0x70, 9, 0, 1], [0x75, 0x63, 9, 0, 1]] {
            let bytes = with_central_extra(&zip_bytes(false), &truncated);
            let mut archive = open_validated_zip(bytes).expect("truncated Unicode metadata");
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"a");
        }
        let repeated: Vec<u8> = (0..500).flat_map(|_| [1u8, 0, 0, 0]).collect();
        let block = opaque_extra_block(&repeated);
        assert_eq!(&block[..4], &[1, 0, 0, 0], "first ZIP64 field kept");
        assert_eq!(le_u16(&block, 4), Some(OPAQUE_EXTRA_FIELD_ID));
        assert_eq!(block.len(), repeated.len());
    }

    #[test]
    fn extended_ntfs_extra_fields_stay_opaque_metadata() {
        // APPNOTE 4.5.5: reserved u32, then attribute records. One timestamp
        // record plus a further (unknown) attribute is valid NTFS metadata
        // that OPC ignores; the package must open and read unchanged.
        let mut ntfs = 0x000au16.to_le_bytes().to_vec();
        let mut body = vec![0u8; 4];
        body.extend_from_slice(&1u16.to_le_bytes());
        body.extend_from_slice(&24u16.to_le_bytes());
        body.extend_from_slice(&[0u8; 24]);
        body.extend_from_slice(&2u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        ntfs.extend_from_slice(&(body.len() as u16).to_le_bytes());
        ntfs.extend_from_slice(&body);
        let bytes = with_central_extra(&zip_bytes(false), &ntfs);
        assert!(zip::ZipArchive::new(Cursor::new(bytes.as_slice())).is_err());
        let mut archive = open_validated_zip(bytes.clone()).expect("opaque NTFS metadata");
        assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"a");
        let mut borrowed = open_validated_slice(&bytes).expect("borrowed open");
        assert_eq!(read_zip_bytes(&mut borrowed, "word/a.xml").unwrap(), b"a");
    }

    /// Splice `extra` into the first central-directory entry.
    fn with_central_extra(zip: &[u8], extra: &[u8]) -> Vec<u8> {
        let eocd = last_signature(zip, EOCD_SIGNATURE);
        let central = le_u32(zip, eocd + 16).unwrap() as usize;
        let name_len = le_u16(zip, central + 28).unwrap() as usize;
        let extra_len = le_u16(zip, central + 30).unwrap() as usize;
        let at = central + 46 + name_len + extra_len;
        let mut out = zip[..at].to_vec();
        out.extend_from_slice(extra);
        out.extend_from_slice(&zip[at..]);
        out[central + 30..central + 32]
            .copy_from_slice(&((extra_len + extra.len()) as u16).to_le_bytes());
        let eocd = eocd + extra.len();
        let size = le_u32(&out, eocd + 12).unwrap() + extra.len() as u32;
        out[eocd + 12..eocd + 16].copy_from_slice(&size.to_le_bytes());
        out
    }

    #[test]
    fn distinct_total_counts_two_entries_and_stops_at_limit_plus_one() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(6), None);
        let _scope = governor.scope("parse");
        let mut archive = archive_with_two();
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
        let error = read_zip_bytes(&mut archive, "word/b.xml").unwrap_err();
        let json = details(&error);
        assert_eq!(
            json["details"]["violation"]["metric"],
            "distinct-inflated-bytes"
        );
        assert_eq!(json["details"]["violation"]["limit"], 6);
        assert_eq!(json["details"]["violation"]["observed"], 7);
    }

    #[test]
    fn reread_does_not_double_charge_distinct_total() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(8), None);
        let mut archive = archive_with_two();
        {
            let _scope = governor.scope("parse");
            validate_archive_limits(&mut archive).unwrap();
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
        }
        {
            let _scope = governor.scope("markdown");
            assert_eq!(read_zip_bytes(&mut archive, "word/a.xml").unwrap(), b"1234");
            assert_eq!(read_zip_bytes(&mut archive, "word/b.xml").unwrap(), b"5678");
        }
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
    }

    #[test]
    fn prefix_then_full_charges_only_the_larger_observation() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(8), None);
        let _scope = governor.scope("parse-sheet");
        let mut archive = archive_with("xl/sheet.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(
            read_zip_string_head(&mut archive, "xl/sheet.xml", 3).unwrap(),
            "123"
        );
        assert_eq!(
            read_zip_string(&mut archive, "xl/sheet.xml").unwrap(),
            "12345678"
        );
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
        assert_eq!(governor.usage().operation_inflated_bytes, 11);
    }

    #[test]
    fn full_then_prefix_adds_no_distinct_bytes() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(8), None);
        let _scope = governor.scope("parse-sheet");
        let mut archive = archive_with("xl/sheet.xml", b"12345678");
        validate_archive_limits(&mut archive).unwrap();
        read_zip_bytes(&mut archive, "xl/sheet.xml").unwrap();
        read_zip_string_head(&mut archive, "xl/sheet.xml", 2).unwrap();
        assert_eq!(governor.usage().distinct_inflated_bytes, 8);
        assert_eq!(governor.usage().operation_inflated_bytes, 10);
    }

    #[test]
    fn prefix_preserves_utf8_boundary() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(16), Some(16), None);
        let _scope = governor.scope("probe");
        let mut archive = archive_with("xl/sheet.xml", "ab€cd".as_bytes());
        validate_archive_limits(&mut archive).unwrap();
        assert_eq!(
            read_zip_string_head(&mut archive, "xl/sheet.xml", 3).unwrap(),
            "ab"
        );
        assert_eq!(governor.usage().distinct_inflated_bytes, 3);
    }

    #[test]
    fn hard_entry_count_has_no_dummy_part() {
        let mut archive = archive_with_two();
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(16), Some(16), None);
        let _scope = governor.scope("open");
        // Exercise the resource layer directly with a proven raw count beyond
        // the hard quota; ZIP preflight wiring is covered separately.
        let error = resource::observe_archive_metadata(20_001, 8).unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "entry-count");
        assert!(violation.get("part").is_none());
        assert!(validate_archive_limits(&mut archive).is_err());
    }

    #[test]
    fn eager_reserve_is_bounded() {
        assert_eq!(initial_reserve(8, 64), 8);
        assert_eq!(
            initial_reserve(512 * 1024 * 1024, 512 * 1024 * 1024),
            INITIAL_RESERVE_CAP
        );
    }

    #[test]
    fn validated_item_names_accept_classic_and_zip64_unique_archives() {
        for zip64 in [false, true] {
            let bytes = zip_bytes(zip64);
            if zip64 {
                assert!(bytes
                    .windows(4)
                    .any(|window| window == ZIP64_EOCD_SIGNATURE.to_le_bytes()));
            }
            let mut archive = zip::ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
            validate_archive_item_names(&bytes, &mut archive).unwrap();
            assert_eq!(inspect_archive_opc(&bytes).unwrap().unwrap().entry_count, 2);
        }
    }

    #[test]
    fn preflight_token_binds_consumer_entry_metadata_not_only_names() {
        let first = archive_bytes_with_single("word/document.xml", b"first");
        let second = archive_bytes_with_single("word/document.xml", b"other");
        let preflight = preflight_archive_limits(&first).unwrap();
        let mut wrong_archive = zip::ZipArchive::new(Cursor::new(second)).unwrap();
        let error = preflight
            .validate_archive_item_names(&mut wrong_archive)
            .unwrap_err();
        assert!(error.contains("does not match the validated central directory"));
    }

    #[test]
    fn stored_payload_with_shallow_eocd_signature_is_not_ambiguous() {
        let mut payload = EOCD_SIGNATURE.to_le_bytes().to_vec();
        payload.resize(22, 0);
        payload.extend_from_slice(b"ordinary stored payload");
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    "word/document.bin",
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Stored),
                )
                .unwrap();
            writer.write_all(&payload).unwrap();
            writer.finish().unwrap();
        }
        let preflight = preflight_archive_limits(&bytes).unwrap();
        let mut archive =
            zip::ZipArchive::with_config(preflight.read_config(), Cursor::new(bytes.as_slice()))
                .unwrap();
        preflight.validate_archive_item_names(&mut archive).unwrap();
    }

    #[test]
    fn preflight_accepts_local_and_central_timestamp_mismatch() {
        let mut bytes = archive_bytes_with_single("word/document.xml", b"document");
        let central = last_signature(&bytes, CENTRAL_DIRECTORY_HEADER_SIGNATURE);
        assert_eq!(le_u16(&bytes, 10), le_u16(&bytes, central + 12));
        assert_eq!(le_u16(&bytes, 12), le_u16(&bytes, central + 14));

        bytes[10..12].copy_from_slice(&1u16.to_le_bytes());
        bytes[12..14].copy_from_slice(&34u16.to_le_bytes());

        let preflight = preflight_archive_limits(&bytes)
            .expect("timestamp-only mismatch is a bounded compatibility case");
        let mut archive =
            zip::ZipArchive::with_config(preflight.read_config(), Cursor::new(bytes.as_slice()))
                .unwrap();
        preflight.validate_archive_item_names(&mut archive).unwrap();
        let mut entry = archive.by_name("word/document.xml").unwrap();
        let mut body = Vec::new();
        entry.read_to_end(&mut body).unwrap();
        assert_eq!(body, b"document");
    }

    #[test]
    fn annex_b2_rejects_local_header_and_data_descriptor_mismatches() {
        let mut local_mismatch = zip_bytes(false);
        let local = local_mismatch
            .windows(4)
            .position(|bytes| bytes == LOCAL_FILE_HEADER_SIGNATURE.to_le_bytes())
            .unwrap();
        local_mismatch[local + 8..local + 10].copy_from_slice(&0u16.to_le_bytes());
        assert!(zip::ZipArchive::new(Cursor::new(local_mismatch.as_slice())).is_ok());
        assert!(preflight_archive_limits(&local_mismatch).is_err());

        let descriptor = stored_descriptor_zip(b"descriptor body");
        let preflight = preflight_archive_limits(&descriptor).unwrap();
        let mut archive = zip::ZipArchive::with_config(
            preflight.read_config(),
            Cursor::new(descriptor.as_slice()),
        )
        .unwrap();
        preflight.validate_archive_item_names(&mut archive).unwrap();

        let mut descriptor_mismatch = descriptor;
        let descriptor_start = last_signature(&descriptor_mismatch, 0x0807_4b50);
        descriptor_mismatch[descriptor_start + 8] ^= 1;
        assert!(preflight_archive_limits(&descriptor_mismatch).is_err());
    }

    #[test]
    fn compatibility_accepts_partial_local_fields_with_signed_descriptor() {
        let body = b"partial descriptor compatibility";
        for bytes in [
            stored_descriptor_zip(body),
            populated_deflate_descriptor_zip(body),
        ] {
            let central = last_signature(&bytes, CENTRAL_DIRECTORY_HEADER_SIGNATURE);
            // Each field can independently be zero or its central value. Include
            // a mixed pair as well as the already-supported fully populated form.
            for mask in [1, 2, 4, 6, 7] {
                let mut partial = bytes.clone();
                for (bit, (local, directory)) in
                    [(14, 16), (18, 20), (22, 24)].into_iter().enumerate()
                {
                    let value = if mask & (1 << bit) != 0 {
                        le_u32(&bytes, central + directory).unwrap()
                    } else {
                        0
                    };
                    partial[local..local + 4].copy_from_slice(&value.to_le_bytes());
                }
                let mut archive = open_validated_zip(partial)
                    .unwrap_or_else(|error| panic!("local field mask {mask}: {error}"));
                let mut actual = Vec::new();
                archive
                    .by_index(0)
                    .unwrap()
                    .read_to_end(&mut actual)
                    .unwrap();
                assert_eq!(actual, body);
            }
        }
    }

    /// One Stored `word/document.xml` entry whose local header says
    /// `local_version` and carries `local_extra`; the central entry says
    /// `central_version` and has no extra field. With `descriptor`, bit 3 is set
    /// and the CRC and sizes move to a signed data descriptor.
    fn local_version_zip(
        body: &[u8],
        local_version: u16,
        central_version: u16,
        local_extra: &[u8],
        descriptor: bool,
    ) -> Vec<u8> {
        const NAME: &[u8] = b"word/document.xml";
        let crc = crc32fast::hash(body);
        let size = body.len() as u32;
        let flags: u16 = if descriptor { 1 << 3 } else { 0 };
        let local_fields = if descriptor {
            [0; 3]
        } else {
            [crc, size, size]
        };
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&LOCAL_FILE_HEADER_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&local_version.to_le_bytes());
        bytes.extend_from_slice(&flags.to_le_bytes());
        bytes.extend_from_slice(&[0; 6]);
        for field in local_fields {
            bytes.extend_from_slice(&field.to_le_bytes());
        }
        bytes.extend_from_slice(&(NAME.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&(local_extra.len() as u16).to_le_bytes());
        bytes.extend_from_slice(NAME);
        bytes.extend_from_slice(local_extra);
        bytes.extend_from_slice(body);
        if descriptor {
            bytes.extend_from_slice(&0x0807_4b50u32.to_le_bytes());
            for field in [crc, size, size] {
                bytes.extend_from_slice(&field.to_le_bytes());
            }
        }

        let central_offset = bytes.len() as u32;
        bytes.extend_from_slice(&CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&central_version.to_le_bytes());
        bytes.extend_from_slice(&flags.to_le_bytes());
        bytes.extend_from_slice(&[0; 6]);
        for field in [crc, size, size] {
            bytes.extend_from_slice(&field.to_le_bytes());
        }
        bytes.extend_from_slice(&(NAME.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&[0; 12]);
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(NAME);
        let central_size = bytes.len() as u32 - central_offset;
        append_classic_eocd(&mut bytes, central_offset, central_size, 1, &[]);
        bytes
    }

    fn zip64_local_field(uncompressed: u64, compressed: u64) -> Vec<u8> {
        let mut field = ZIP64_EXTRA_FIELD_ID.to_le_bytes().to_vec();
        field.extend_from_slice(&16u16.to_le_bytes());
        field.extend_from_slice(&uncompressed.to_le_bytes());
        field.extend_from_slice(&compressed.to_le_bytes());
        field
    }

    #[test]
    fn compatibility_accepts_reserved_zip64_local_field_with_classic_central_entry() {
        let body = b"reserved ZIP64 local field";
        let size = body.len() as u64;
        // Unrelated extra fields may surround the ZIP64 field.
        let mut surrounded = vec![0x55, 0x54, 5, 0, 1, 0, 0, 0, 0];
        surrounded.extend_from_slice(&zip64_local_field(0, 0));
        surrounded.extend_from_slice(&[0x99, 0x99, 0, 0]);
        for (central_version, extra) in [
            (20, zip64_local_field(0, 0)),
            (20, zip64_local_field(size, size)),
            (20, zip64_local_field(size, 0)),
            (20, zip64_local_field(0, size)),
            (10, zip64_local_field(0, 0)),
            (20, surrounded),
        ] {
            let bytes = local_version_zip(body, 45, central_version, &extra, false);
            let mut archive = open_validated_zip(bytes)
                .unwrap_or_else(|error| panic!("central {central_version} {extra:?}: {error}"));
            let mut actual = Vec::new();
            archive
                .by_index(0)
                .unwrap()
                .read_to_end(&mut actual)
                .unwrap();
            assert_eq!(actual, body);
        }
    }

    #[test]
    fn compatibility_rejects_unexplained_version_needed_mismatches() {
        let body = b"version needed mismatch";
        let size = body.len() as u64;
        let reserved = zip64_local_field(0, 0);
        let mut long_field = ZIP64_EXTRA_FIELD_ID.to_le_bytes().to_vec();
        long_field.extend_from_slice(&28u16.to_le_bytes());
        long_field.extend_from_slice(&[0; 28]);
        let mut short_field = ZIP64_EXTRA_FIELD_ID.to_le_bytes().to_vec();
        short_field.extend_from_slice(&8u16.to_le_bytes());
        short_field.extend_from_slice(&[0; 8]);
        let mut duplicate = reserved.clone();
        duplicate.extend_from_slice(&reserved);
        let mut truncated = reserved.clone();
        truncated.extend_from_slice(&[0x99, 0x99, 8, 0]);
        for (case, local_version, central_version, extra) in [
            (
                "conflicting uncompressed",
                45,
                20,
                zip64_local_field(size + 1, 0),
            ),
            (
                "conflicting compressed",
                45,
                20,
                zip64_local_field(0, size + 1),
            ),
            ("no ZIP64 field", 45, 20, Vec::new()),
            ("long ZIP64 field", 45, 20, long_field),
            ("short ZIP64 field", 45, 20, short_field),
            ("duplicate ZIP64 field", 45, 20, duplicate),
            ("truncated extra block", 45, 20, truncated),
            ("local newer than 4.5", 46, 20, reserved.clone()),
            ("central also ZIP64", 45, 46, reserved.clone()),
            ("local older than central", 20, 45, reserved.clone()),
            ("plain mismatch", 20, 10, Vec::new()),
        ] {
            let bytes = local_version_zip(body, local_version, central_version, &extra, false);
            assert!(preflight_archive_limits(&bytes).is_err(), "{case}");
        }

        // The data-descriptor policy stays as it is.
        let descriptor = local_version_zip(body, 45, 20, &reserved, true);
        assert!(preflight_archive_limits(&descriptor).is_err());
        assert!(preflight_archive_limits(&local_version_zip(body, 20, 20, &[], true)).is_ok());

        // A local 32-bit sentinel takes its size from the ZIP64 field, which then
        // really is needed; keep requiring the same version in that case.
        let mut sentinel = local_version_zip(body, 45, 20, &zip64_local_field(size, size), false);
        sentinel[18..26].fill(0xff);
        assert!(preflight_archive_limits(&sentinel).is_err());
        sentinel[4..6].copy_from_slice(&20u16.to_le_bytes());
        assert!(preflight_archive_limits(&sentinel).is_ok());
    }

    #[test]
    fn compatibility_reads_empty_deflated_entry_with_populated_compressed_size() {
        let mut archive = open_validated_zip(populated_deflate_descriptor_zip(b"")).expect(
            "empty Deflate entry has zero CRC/uncompressed size but nonzero compressed size",
        );
        let mut actual = Vec::new();
        archive
            .by_index(0)
            .unwrap()
            .read_to_end(&mut actual)
            .unwrap();
        assert!(actual.is_empty());
    }

    #[test]
    fn compatibility_rejects_mismatched_or_unsigned_populated_local_fields() {
        let bytes = populated_deflate_descriptor_zip(b"bounded compatibility");

        for field in [14, 18, 22] {
            let mut mismatched = bytes.clone();
            // A partially filled header is tolerated only when its populated
            // field agrees; zero placeholders must not hide a conflicting value.
            mismatched[14..26].fill(0);
            mismatched[field..field + 4].copy_from_slice(&bytes[field..field + 4]);
            mismatched[field] ^= 1;
            assert!(preflight_archive_limits(&mismatched).is_err());
        }

        let eocd = last_signature(&bytes, EOCD_SIGNATURE);
        let central = le_u32(&bytes, eocd + 16).unwrap() as usize;
        for field in [central - 12, central - 8, central - 4] {
            let mut mismatched_descriptor = bytes.clone();
            mismatched_descriptor[field] ^= 1;
            assert!(preflight_archive_limits(&mismatched_descriptor).is_err());
        }
        for field in [central + 16, central + 20, central + 24] {
            let mut mismatched_central = bytes.clone();
            mismatched_central[field] ^= 1;
            assert!(preflight_archive_limits(&mismatched_central).is_err());
        }

        let mut unsigned_descriptor = bytes;
        unsigned_descriptor.drain(central - 16..central - 12);
        let shifted_eocd = eocd - 4;
        unsigned_descriptor[shifted_eocd + 16..shifted_eocd + 20]
            .copy_from_slice(&((central - 4) as u32).to_le_bytes());
        assert!(preflight_archive_limits(&unsigned_descriptor).is_err());
    }

    #[test]
    fn part_names_must_not_be_derivable_at_ascii_case_folded_segment_boundaries() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("WORD", options).unwrap();
            writer.write_all(b"parent").unwrap();
            writer.start_file("word/document.xml", options).unwrap();
            writer.write_all(b"child").unwrap();
            writer.finish().unwrap();
        }
        let error = preflight_archive_limits(&bytes).unwrap_err();
        assert!(error.contains("must not be derivable"));

        let mut directory_bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut directory_bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("foo", options).unwrap();
            writer.write_all(b"part").unwrap();
            writer.add_directory("foo/", options).unwrap();
            writer.finish().unwrap();
        }
        preflight_archive_limits(&directory_bytes)
            .expect("directory items do not participate in OPC part derivability");
    }

    #[test]
    fn malformed_zip64_signature_inside_extensible_data_is_not_ambiguous() {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let comment = "x".repeat(128);
            writer.set_raw_zip64_extensible_data_sector(comment.into_bytes().into());
            writer
                .start_file(
                    "word/document.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"document").unwrap();
            writer.finish().unwrap();
        }

        let real_zip64 = last_signature(&bytes, ZIP64_EOCD_SIGNATURE);
        let locator = last_signature(&bytes, ZIP64_LOCATOR_SIGNATURE);
        let fake_zip64 = real_zip64 + 56;
        let fake_record_size = (locator - fake_zip64 - 12) as u64;
        bytes[fake_zip64..fake_zip64 + 4].copy_from_slice(&ZIP64_EOCD_SIGNATURE.to_le_bytes());
        bytes[fake_zip64 + 4..fake_zip64 + 12].copy_from_slice(&fake_record_size.to_le_bytes());
        bytes[fake_zip64 + 16..fake_zip64 + 24].fill(0);
        bytes[fake_zip64 + 24..fake_zip64 + 40].fill(0);
        bytes[fake_zip64 + 40..fake_zip64 + 48].copy_from_slice(&(real_zip64 as u64).to_le_bytes());
        bytes[fake_zip64 + 48..fake_zip64 + 56].fill(0);

        let preflight = preflight_archive_limits(&bytes)
            .expect("semantically malformed ZIP64 signatures are ignored");
        assert_eq!(preflight.entry_count, 1);
    }

    #[test]
    fn preflight_accepts_classic_and_zip64_sfx_archives() {
        for zip64 in [false, true] {
            let bytes = prepend(&zip_bytes(zip64));
            preflight_archive_limits(&bytes).unwrap();
            let mut archive = zip::ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
            validate_archive_item_names(&bytes, &mut archive).unwrap();
            assert_eq!(archive.len(), 2);
        }
    }

    #[test]
    fn preflight_accepts_central_directory_digital_signatures() {
        for zip64 in [false, true] {
            for sfx in [false, true] {
                let bytes = add_central_directory_signature(zip_bytes(zip64), zip64);
                let bytes = if sfx { prepend(&bytes) } else { bytes };
                let preflight = preflight_archive_limits(&bytes)
                    .unwrap_or_else(|error| panic!("zip64={zip64} sfx={sfx}: {error}"));
                let mut archive = zip::ZipArchive::with_config(
                    preflight.read_config(),
                    Cursor::new(bytes.as_slice()),
                )
                .unwrap();
                validate_archive_item_names(&bytes, &mut archive).unwrap();
                assert_eq!(archive.len(), 2);
                assert_eq!(zip::ZipArchive::new(Cursor::new(bytes)).unwrap().len(), 2);
            }
        }
    }

    #[test]
    fn trailing_garbage_is_rejected_before_zip_archive_allocation() {
        let mut bytes = zip_bytes(false);
        bytes.extend_from_slice(b"trailing-garbage");
        assert!(zip::ZipArchive::new(Cursor::new(bytes.as_slice())).is_ok());

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let error = preflight_archive_limits(&bytes).unwrap_err();
        assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(governor.first_error().is_none());
    }

    #[test]
    fn zip64_classic_sentinels_are_resolved_independently() {
        let original = zip_bytes(true);
        let eocd = last_signature(&original, EOCD_SIGNATURE);
        for (offset, width) in [(8, 2), (10, 2), (12, 4), (16, 4)] {
            let mut bytes = original.clone();
            bytes[eocd + offset..eocd + offset + width].fill(0xff);
            if offset == 8 {
                let error = preflight_archive_limits(&bytes).unwrap_err();
                assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
            } else {
                let preflight = preflight_archive_limits(&bytes).unwrap();
                assert_eq!(preflight.entry_count, 2);
                let mut archive = zip::ZipArchive::with_config(
                    preflight.read_config(),
                    Cursor::new(bytes.as_slice()),
                )
                .unwrap();
                validate_archive_item_names(&bytes, &mut archive).unwrap();
                assert_eq!(archive.len(), 2);
                assert_eq!(zip::ZipArchive::new(Cursor::new(bytes)).unwrap().len(), 2);
            }
        }
    }

    #[test]
    fn validated_item_names_reject_exact_and_ascii_case_duplicates_as_conformance_errors() {
        let mut exact = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut exact));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"first").unwrap();
            writer.start_file("word/b.xml", options).unwrap();
            writer.write_all(b"last").unwrap();
            writer.finish().unwrap();
        }
        for offset in 0..=exact.len() - b"word/b.xml".len() {
            if &exact[offset..offset + b"word/b.xml".len()] == b"word/b.xml" {
                exact[offset..offset + b"word/b.xml".len()].copy_from_slice(b"word/a.xml");
            }
        }
        let mut exact_archive = zip::ZipArchive::new(Cursor::new(exact.as_slice())).unwrap();
        let preflight_exact_error = preflight_archive_limits(&exact).unwrap_err();
        assert!(preflight_exact_error.contains("ZIP item names must be unique"));
        let exact_error = validate_archive_item_names(&exact, &mut exact_archive).unwrap_err();
        assert!(exact_error.contains("ZIP item names must be unique"));
        assert!(!exact_error.starts_with("OOXML_RESOURCE_LIMIT:"));

        let mut case_bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut case_bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"lower").unwrap();
            writer.start_file("WORD/A.XML", options).unwrap();
            writer.write_all(b"upper").unwrap();
            writer.finish().unwrap();
        }
        let mut case_archive = zip::ZipArchive::new(Cursor::new(case_bytes.as_slice())).unwrap();
        let preflight_case_error = preflight_archive_limits(&case_bytes).unwrap_err();
        assert!(preflight_case_error.contains("unique under part-name equivalence"));
        let case_error = validate_archive_item_names(&case_bytes, &mut case_archive).unwrap_err();
        assert!(case_error.contains("unique under part-name equivalence"));
        assert!(!case_error.starts_with("OOXML_RESOURCE_LIMIT:"));

        // RFC 3986 §6.2.2.2: `%41` names the same part as `A`.
        let mut percent_bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut percent_bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a.xml", options).unwrap();
            writer.write_all(b"plain").unwrap();
            writer.start_file("word/%41.xml", options).unwrap();
            writer.write_all(b"encoded").unwrap();
            writer.finish().unwrap();
        }
        let percent_error = preflight_archive_limits(&percent_bytes).unwrap_err();
        assert!(percent_error.contains("unique under part-name equivalence"));
        // Other percent-encodings stay distinct octets.
        let mut distinct = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut distinct));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("word/a b.xml", options).unwrap();
            writer.write_all(b"space").unwrap();
            writer.start_file("word/a%20b.xml", options).unwrap();
            writer.write_all(b"encoded").unwrap();
            writer.finish().unwrap();
        }
        assert!(preflight_archive_limits(&distinct).is_ok());
    }

    #[test]
    fn preflight_rejects_non_ascii_and_invalid_utf8_item_names_as_conformance_errors() {
        for replacement in [b"\xc3\xa9rd/a.xml".as_slice(), b"\xfford/a.xml".as_slice()] {
            let mut bytes = zip_bytes(false);
            replace_raw_name(&mut bytes, replacement);
            let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
            let _scope = governor.scope("open");
            let error = preflight_archive_limits(&bytes).unwrap_err();
            assert!(error.contains("ASCII"));
            assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
            assert!(governor.first_error().is_none());
        }
    }

    #[test]
    fn exact_name_policy_requires_flagged_utf8_for_non_ascii_names() {
        // Non-ASCII bytes without the APPNOTE bit-11 flag, and invalid UTF-8.
        for replacement in [b"\xc3\xa9rd/a.xml".as_slice(), b"\xfford/a.xml".as_slice()] {
            let mut bytes = zip_bytes(false);
            replace_raw_name(&mut bytes, replacement);
            let error =
                preflight_archive_with_policy(&bytes, ItemNamePolicy::ExactUtf8).unwrap_err();
            assert!(error.contains("UTF-8"), "{error}");
            assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
        }
        assert!(
            preflight_archive_with_policy(&zip_bytes(false), ItemNamePolicy::ExactUtf8).is_ok()
        );
    }

    #[test]
    fn raw_eocd_entry_count_is_rejected_before_zip_open() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let (mut eocd, central_offset, central_size) = synthetic_entries(20_001, 0);
        append_classic_eocd(
            &mut eocd,
            central_offset,
            central_size,
            20_001,
            &[
                0x50, 0x4b, 0x05, 0x06, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xff, 0xff,
            ],
        );

        // A later EOCD-looking sequence inside the legal comment is malformed.
        // Preflight must continue scanning and use the real record above.
        let error = preflight_archive_limits(&eocd).unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "entry-count");
        assert_eq!(violation["limit"], resource::STANDARD_MAX_ARCHIVE_ENTRIES);
        assert_eq!(
            violation["observed"],
            resource::STANDARD_MAX_ARCHIVE_ENTRIES + 1
        );
        assert_eq!(violation["configurable"], true);
    }

    #[test]
    fn raw_zip64_eocd_entry_count_is_rejected_before_zip_open() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Pptx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let (mut bytes, central_offset, central_size) = synthetic_entries(20_001, 0);

        // ZIP64 end of central directory record follows the synthetic central
        // directory. Only its bounds and first signature are needed pre-open.
        let zip64_offset = bytes.len() as u64;
        bytes.extend_from_slice(&0x0606_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&44u64.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&45u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&20_001u64.to_le_bytes());
        bytes.extend_from_slice(&20_001u64.to_le_bytes());
        bytes.extend_from_slice(&(central_size as u64).to_le_bytes());
        bytes.extend_from_slice(&(central_offset as u64).to_le_bytes());

        // Locator points back to the ZIP64 record above.
        bytes.extend_from_slice(&0x0706_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&zip64_offset.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());

        // Saturated classic EOCD fields require ZIP64 metadata.
        bytes.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&u16::MAX.to_le_bytes());
        bytes.extend_from_slice(&u16::MAX.to_le_bytes());
        bytes.extend_from_slice(&u32::MAX.to_le_bytes());
        bytes.extend_from_slice(&u32::MAX.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());

        let error = preflight_archive_limits(&bytes).unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "entry-count");
        assert_eq!(violation["limit"], resource::STANDARD_MAX_ARCHIVE_ENTRIES);
        assert_eq!(
            violation["observed"],
            resource::STANDARD_MAX_ARCHIVE_ENTRIES + 1
        );
        assert_eq!(violation["configurable"], true);
    }

    #[test]
    fn malformed_eocd_remains_an_ordinary_container_error() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let mut bytes = 0x0605_4b50u32.to_le_bytes().to_vec();
        bytes.resize(22, 0);
        bytes[20..22].copy_from_slice(&10u16.to_le_bytes());
        let error = preflight_archive_limits(&bytes).unwrap_err();
        assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(governor.first_error().is_none());
    }

    #[test]
    fn malformed_central_directory_remains_an_ordinary_container_error() {
        let mut bytes = zip_bytes(false);
        let central = last_signature(&bytes, CENTRAL_DIRECTORY_HEADER_SIGNATURE);
        bytes[central] = 0;
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let error = preflight_archive_limits(&bytes).unwrap_err();
        assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(governor.first_error().is_none());
        assert!(zip::ZipArchive::new(Cursor::new(bytes)).is_err());
    }

    #[test]
    fn shallow_tail_eocd_signature_with_invalid_directory_is_ignored() {
        let mut bytes = zip_bytes(false);
        let real_eocd = last_signature(&bytes, EOCD_SIGNATURE);
        bytes[real_eocd + 20..real_eocd + 22].copy_from_slice(&68u16.to_le_bytes());

        let fake_central = bytes.len();
        bytes.resize(fake_central + CENTRAL_DIRECTORY_HEADER_BYTES as usize, 0);
        bytes[fake_central..fake_central + 4]
            .copy_from_slice(&CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes());
        bytes[fake_central + 42..fake_central + 46].copy_from_slice(&u32::MAX.to_le_bytes());
        append_classic_eocd(&mut bytes, 0, 46, 1, &[]);

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Docx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let preflight = preflight_archive_limits(&bytes).unwrap();
        assert_eq!(preflight.entry_count, 2);
        assert!(governor.first_error().is_none());
    }

    #[test]
    fn non_authoritative_eocd_signatures_with_trailing_bytes_are_ignored() {
        for fake_count in [0u16, 1u16] {
            let mut bytes = zip_bytes(false);
            let real_eocd = last_signature(&bytes, EOCD_SIGNATURE);
            let fake_directory_bytes = usize::from(fake_count != 0) * 46;
            let appended = fake_directory_bytes + 22 + 3;
            bytes[real_eocd + 20..real_eocd + 22].copy_from_slice(&(appended as u16).to_le_bytes());
            if fake_count != 0 {
                let central = bytes.len();
                bytes.resize(central + 46, 0);
                bytes[central..central + 4]
                    .copy_from_slice(&CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes());
            }
            append_classic_eocd(&mut bytes, 0, fake_directory_bytes as u32, fake_count, &[]);
            bytes.extend_from_slice(b"end");

            let governor = ResourceGovernor::from_wasm(OoxmlFormat::Pptx, Some(64), Some(64), None);
            let _scope = governor.scope("open");
            let preflight = preflight_archive_limits(&bytes).unwrap();
            assert_eq!(preflight.entry_count, 2);
            assert!(governor.first_error().is_none());
        }
    }

    #[test]
    fn earlier_eocd_signature_before_a_valid_sfx_archive_is_ignored() {
        let mut bytes = EOCD_SIGNATURE.to_le_bytes().to_vec();
        bytes.resize(22, 0);
        bytes.extend_from_slice(&zip_bytes(false));

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let preflight = preflight_archive_limits(&bytes).unwrap();
        assert_eq!(preflight.entry_count, 2);
        assert!(governor.first_error().is_none());

        let mut archive =
            zip::ZipArchive::with_config(preflight.read_config(), Cursor::new(bytes.as_slice()))
                .unwrap();
        preflight.validate_archive_item_names(&mut archive).unwrap();
    }

    #[test]
    fn malformed_oversized_directory_declaration_does_not_poison() {
        let central_size = HARD_MAX_CENTRAL_DIRECTORY_BYTES as usize + 1;
        let mut bytes = vec![0; central_size];
        // An EOCD-looking suffix cannot latch a resource violation until its
        // claimed central/local structure is proven semantically coherent.
        append_classic_eocd(&mut bytes, 0, central_size as u32, 1, &[]);
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let error = preflight_archive_limits(&bytes).unwrap_err();
        assert!(!error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(governor.first_error().is_none());
    }

    #[test]
    fn aggregate_central_directory_metadata_is_bounded_before_zip_open() {
        const MAX_NAME: usize = u16::MAX as usize;
        let per_entry_charge = CENTRAL_DIRECTORY_HEADER_BYTES as usize + 2 * MAX_NAME;
        let count = (HARD_MAX_CENTRAL_DIRECTORY_BYTES as usize / per_entry_charge) + 1;
        let (mut bytes, central_offset, central_size) = synthetic_entries(count, MAX_NAME);
        append_classic_eocd(&mut bytes, central_offset, central_size, count as u16, &[]);

        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        let error = preflight_archive_limits(&bytes).unwrap_err();
        let json = details(&error);
        let violation = &json["details"]["violation"];
        assert_eq!(violation["metric"], "central-directory-bytes");
        assert_eq!(violation["limit"], HARD_MAX_CENTRAL_DIRECTORY_BYTES);
        assert_eq!(violation["observed"], HARD_MAX_CENTRAL_DIRECTORY_BYTES + 1);
        assert!(violation.get("part").is_none());
    }

    #[test]
    fn central_directory_limit_accepts_exact_limit_and_poisons_at_plus_one() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Pptx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        resource::observe_archive_central_directory_bytes(HARD_MAX_CENTRAL_DIRECTORY_BYTES)
            .unwrap();
        let first =
            resource::observe_archive_central_directory_bytes(HARD_MAX_CENTRAL_DIRECTORY_BYTES + 1)
                .unwrap_err();
        let violation = &details(&first)["details"]["violation"];
        assert_eq!(violation["metric"], "central-directory-bytes");
        assert_eq!(violation["observed"], HARD_MAX_CENTRAL_DIRECTORY_BYTES + 1);
        assert_eq!(
            resource::observe_archive_central_directory_bytes(0).unwrap_err(),
            first
        );
    }

    #[test]
    fn zip64_footer_limit_accepts_exact_and_types_plus_one() {
        let governor = ResourceGovernor::from_wasm(OoxmlFormat::Xlsx, Some(64), Some(64), None);
        let _scope = governor.scope("open");
        preflight_archive_limits(&empty_zip64_with_footer_bytes(
            HARD_MAX_CENTRAL_DIRECTORY_BYTES as usize,
        ))
        .unwrap();

        let error = preflight_archive_limits(&empty_zip64_with_footer_bytes(
            HARD_MAX_CENTRAL_DIRECTORY_BYTES as usize + 1,
        ))
        .unwrap_err();
        let violation = &details(&error)["details"]["violation"];
        assert_eq!(violation["metric"], "central-directory-bytes");
        assert_eq!(violation["limit"], HARD_MAX_CENTRAL_DIRECTORY_BYTES);
        assert_eq!(violation["observed"], HARD_MAX_CENTRAL_DIRECTORY_BYTES + 1);
    }
}
