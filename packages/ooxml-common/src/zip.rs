//! Bounded ZIP access for OOXML packages.
//!
//! This module owns archive-specific reading and metadata inspection. Persistent
//! policy, accounting, usage snapshots, and poison state live in `resource`.

use crate::resource::{
    self, OoxmlFormat, ResourceGovernor, ResourceScope, HARD_MAX_ARCHIVE_ENTRIES,
    HARD_MAX_ARCHIVE_ENTRY_BYTES, HARD_MAX_CENTRAL_DIRECTORY_BYTES,
};
use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};

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
                && entry.data_start() == expected.data_start
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

#[derive(Clone, Copy)]
struct AsciiFoldedName<'a>(&'a [u8]);

impl PartialEq for AsciiFoldedName<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(other.0)
    }
}

impl Eq for AsciiFoldedName<'_> {}

impl Hash for AsciiFoldedName<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.0.len().hash(state);
        for byte in self.0 {
            byte.to_ascii_lowercase().hash(state);
        }
    }
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
        if resolved_directory.replace(directory).is_some() {
            return Err("ambiguous ZIP64 end-of-central-directory records".to_string());
        }
    }
    let Some(directory) = resolved_directory else {
        return Ok(None);
    };
    match inspect_central_directory(data, directory) {
        Ok(preflight) => Ok(Some(preflight)),
        Err(CandidateError::Malformed) => Ok(None),
        Err(CandidateError::Rejected(error)) => Err(error),
    }
}

fn inspect_eocd_candidate(data: &[u8], eocd: usize) -> Result<Option<ArchivePreflight>, String> {
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
    match inspect_central_directory(data, directory) {
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
    if data_start > directory.start
        || data.get(local_name_start..local_extra_start)
            != data.get(entry.name.start..entry.name.start + entry.name.len)
        || le_u16(data, local + 4) != le_u16(data, entry.header + 6)
        || le_u16(data, local + 6) != Some(fields.flags)
        || le_u16(data, local + 8) != Some(fields.compression_method)
        || le_u16(data, local + 10) != le_u16(data, entry.header + 12)
        || le_u16(data, local + 12) != le_u16(data, entry.header + 14)
    {
        return None;
    }

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
        // bit 3 is set. NPOI 2.3.0 output reported in #1428 is a legacy form
        // that Excel accepts: non-ZIP64, with a signed descriptor and all three
        // local fields populated. Keep that exception bounded to exact agreement.
        let compatible_populated_local_fields = !fields.zip64_sizes
            && signed_descriptor
            && local_crc != 0
            && local_compressed != 0
            && local_uncompressed != 0
            && local_crc == fields.crc32
            && u64::from(local_compressed) == fields.compressed_size
            && u64::from(local_uncompressed) == fields.uncompressed_size;
        if (!canonical_local_fields && !compatible_populated_local_fields) || !valid_descriptor {
            return None;
        }
    } else {
        let local_extra = NameRange {
            start: local_extra_start,
            len: local_extra_len,
        };
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
    let mut folded_names = HashMap::<AsciiFoldedName<'_>, NameRange>::with_capacity(capacity);
    walk_central_directory(data, directory, |entry| {
        let range = entry.name;
        let name = &data[range.start..range.start + range.len];
        if !name.is_ascii() {
            return Err(CandidateError::Rejected(
                "OPC ZIP item names must contain only ASCII bytes (ECMA-376 Part 2, 7.3.3)"
                    .to_string(),
            ));
        }
        if !exact_names.insert(name) {
            return Err(CandidateError::Rejected(format!(
                "ZIP item names must be unique: {}",
                String::from_utf8_lossy(name)
            )));
        }
        let represents_file = !name.ends_with(b"/") && !name.ends_with(b"\\");
        if represents_file {
            if let Some(previous) = folded_names.insert(AsciiFoldedName(name), range) {
                let previous_name = &data[previous.start..previous.start + previous.len];
                return Err(CandidateError::Rejected(format!(
                    "OOXML part names must be unique ignoring ASCII case: {} and {}",
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
        if entry.name.ends_with(b"/") || entry.name.ends_with(b"\\") {
            continue;
        }
        for slash in entry
            .name
            .iter()
            .enumerate()
            .filter_map(|(index, byte)| (*byte == b'/').then_some(index))
        {
            if folded_names.contains_key(&AsciiFoldedName(&entry.name[..slash])) {
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
    })
}

/// Validate the legal EOCD tail and walk its central directory before
/// `zip::ZipArchive::new` allocates its filename index. Malformed/truncated
/// metadata is left to the ZIP crate's ordinary corruption path rather than
/// being mislabeled as a resource violation.
fn preflight_archive_limits(data: &[u8]) -> Result<ArchivePreflight, String> {
    inspect_archive(data)?.ok_or_else(|| "ZIP central directory preflight failed".to_string())
}

fn inspect_archive(data: &[u8]) -> Result<Option<ArchivePreflight>, String> {
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
        if let Some(preflight) = inspect_eocd_candidate(data, eocd)? {
            if resolved.replace(preflight).is_some() {
                return Err("ambiguous ZIP end-of-central-directory records".to_string());
            }
        }
    }
    Ok(resolved)
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
/// package sessions. It deliberately returns raw container errors; format
/// facades own their historical user-visible error context.
pub(crate) fn open_validated_cursor<T: AsRef<[u8]>>(
    cursor: std::io::Cursor<T>,
) -> Result<zip::ZipArchive<std::io::Cursor<T>>, String> {
    let preflight = preflight_archive_limits(cursor.get_ref().as_ref())?;
    let mut archive = zip::ZipArchive::with_config(preflight.read_config(), cursor)
        .map_err(|error| error.to_string())?;
    preflight.validate_archive_item_names(&mut archive)?;
    validate_archive_limits(&mut archive)?;
    Ok(archive)
}

/// Open an owned ZIP through the shared validated construction path.
pub fn open_validated_zip(
    data: Vec<u8>,
) -> Result<zip::ZipArchive<std::io::Cursor<Vec<u8>>>, String> {
    open_validated_cursor(std::io::Cursor::new(data))
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

/// Install a retained archive's persistent governor for one synchronous method.
pub fn scoped_governor(governor: &ResourceGovernor, operation: &str) -> ResourceScope {
    governor.scope(operation)
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
    let mut archive = open_validated_cursor(std::io::Cursor::new(data)).map_err(|error| {
        if error.starts_with("OOXML_RESOURCE_LIMIT:") {
            error
        } else {
            format!("zip open error: {error}")
        }
    })?;
    read_zip_bytes(&mut archive, path)
}

pub fn read_zip_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<u8>, String> {
    resource::assert_healthy()?;
    let part_id = archive
        .index_for_name(path)
        .ok_or_else(|| format!("entry not found: {path}"))?;
    let mut entry = archive
        .by_index(part_id)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    let declared_size = entry.size();
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
    let part_id = archive
        .index_for_name(path)
        .ok_or_else(|| format!("entry not found: {path}"))?;
    let mut entry = archive
        .by_index(part_id)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    let declared_size = entry.size();
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

/// Legacy fallback used only when a parser test/helper has not installed a
/// governor. Public browser paths always install a normalized session policy.
pub fn fallback_max_archive_entry_bytes() -> u64 {
    HARD_MAX_ARCHIVE_ENTRY_BYTES
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
                writer.set_zip64_comment(Some("metadata PK\u{6}\u{6} with a false signature"));
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

    fn npoi_230_descriptor_zip(body: &[u8]) -> Vec<u8> {
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
        assert_ne!(crc, 0);
        assert_ne!(compressed, 0);
        assert_ne!(uncompressed, 0);

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
        // The stored payload still delivers eight bytes, so metadata alone must
        // not authorize the read.
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
            assert_eq!(inspect_archive(&bytes).unwrap().unwrap().entry_count, 2);
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
    fn compatibility_accepts_matching_populated_local_fields_with_signed_descriptor() {
        let bytes = npoi_230_descriptor_zip(b"confirmed NPOI 2.3.0 compatibility shape");
        let preflight = preflight_archive_limits(&bytes).unwrap();
        let mut archive =
            zip::ZipArchive::with_config(preflight.read_config(), Cursor::new(bytes.as_slice()))
                .unwrap();
        preflight.validate_archive_item_names(&mut archive).unwrap();
        let mut entry = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut body = Vec::new();
        entry.read_to_end(&mut body).unwrap();
        assert_eq!(body, b"confirmed NPOI 2.3.0 compatibility shape");
    }

    #[test]
    fn compatibility_rejects_partial_or_mismatched_populated_local_fields() {
        let bytes = npoi_230_descriptor_zip(b"bounded compatibility");

        for field in [14, 18, 22] {
            let mut partial = bytes.clone();
            partial[field..field + 4].copy_from_slice(&0u32.to_le_bytes());
            assert!(preflight_archive_limits(&partial).is_err());

            let mut mismatched = bytes.clone();
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
            writer.set_zip64_comment(Some(comment));
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
                assert!(zip::ZipArchive::new(Cursor::new(bytes)).is_err());
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
        assert!(preflight_case_error.contains("unique ignoring ASCII case"));
        let case_error = validate_archive_item_names(&case_bytes, &mut case_archive).unwrap_err();
        assert!(case_error.contains("unique ignoring ASCII case"));
        assert!(!case_error.starts_with("OOXML_RESOURCE_LIMIT:"));
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
