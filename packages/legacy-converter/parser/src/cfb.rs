//! Bounded reader for the Compound File Binary container used by Office 97-2003.
//!
//! This intentionally implements only the read path needed by the converter.
//! Sector chains, DIFAT/FAT tables, MiniFAT streams, and directory entries are
//! validated before a legacy-format parser sees any bytes. The bounds follow
//! [MS-CFB] sections 2.2 through 2.6. Existing top-level lookup remains flat;
//! parent-scoped lookup additionally validates the directory hierarchy.

use std::collections::{HashMap, HashSet};

// [MS-CFB] 2.2 Compound File Header (`_abSig`). This is a byte sequence,
// not a little-endian integer; reversing it would reject every Office file.
const SIGNATURE: [u8; 8] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE_SECTOR: u32 = 0xffff_ffff;
const END_OF_CHAIN: u32 = 0xffff_fffe;
const FAT_SECTOR: u32 = 0xffff_fffd;
const DIFAT_SECTOR: u32 = 0xffff_fffc;
const NO_STREAM: u32 = 0xffff_ffff;
const HEADER_BYTES: usize = 512;
const DIRECTORY_ENTRY_BYTES: usize = 128;
const MAX_DIRECTORY_ENTRIES: usize = 1_000_000;

#[derive(Debug, Clone, Copy)]
struct DirectoryEntry {
    object_type: u8,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    start_sector: u32,
    stream_size: u64,
}

type DirectorySlot = Option<(String, DirectoryEntry)>;

pub struct CompoundFile<'a> {
    bytes: &'a [u8],
    sector_size: usize,
    mini_sector_size: usize,
    mini_stream_cutoff: u64,
    physical_sectors: usize,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    // Raw directory IDs index this vector, including unallocated slots.
    directory: Vec<DirectorySlot>,
    root_mini_stream: Vec<u8>,
}

pub struct ScopedStreams<'cfb, 'data> {
    compound: &'cfb CompoundFile<'data>,
    root: usize,
    children: Vec<HashMap<String, usize>>,
    has_non_ascii_child: Vec<bool>,
}

impl ScopedStreams<'_, '_> {
    /// Resolve an ASCII Office storage path using ASCII case folding. This
    /// intentionally does not approximate the full MS-CFB Unicode comparator;
    /// non-ASCII directory names remain retained but are outside this API.
    pub fn stream(&self, path: &[&str], maximum_bytes: usize) -> Result<Vec<u8>, String> {
        if path.is_empty() {
            return Err("empty CFB stream path".into());
        }
        let mut parent = self.root;
        for (index, expected) in path.iter().enumerate() {
            if !expected.is_ascii() {
                return Err("non-ASCII CFB scoped path".into());
            }
            if self.has_non_ascii_child[parent] {
                return Err("non-ASCII sibling in CFB scoped path".into());
            }
            let id = *self.children[parent]
                .get(&expected.to_ascii_uppercase())
                .ok_or_else(|| format!("missing CFB path entry: {expected}"))?;
            let (_, entry) = self.compound.directory[id]
                .as_ref()
                .expect("validated hierarchy entry");
            if index + 1 == path.len() {
                if entry.object_type != 2 {
                    return Err(format!("CFB path does not end in a stream: {expected}"));
                }
                if entry.stream_size > maximum_bytes as u64 {
                    return Err("CFB scoped stream exceeds its caller limit".into());
                }
                return self.compound.read_stream(*entry);
            }
            if entry.object_type != 1 {
                return Err(format!("CFB path traverses a non-storage: {expected}"));
            }
            parent = id;
        }
        unreachable!()
    }
}

impl<'a> CompoundFile<'a> {
    pub fn open(bytes: &'a [u8]) -> Result<Self, String> {
        if bytes.len() < HEADER_BYTES || bytes[..8] != SIGNATURE {
            return Err("invalid CFB signature".into());
        }
        if u16_at(bytes, 28)? != 0xfffe {
            return Err("unsupported CFB byte order".into());
        }
        let major = u16_at(bytes, 26)?;
        let sector_shift = u16_at(bytes, 30)?;
        let sector_size = match (major, sector_shift) {
            (3, 9) => 512,
            (4, 12) => 4096,
            _ => return Err("unsupported CFB version".into()),
        };
        let mini_sector_shift = u16_at(bytes, 32)?;
        if mini_sector_shift != 6 {
            return Err("unsupported CFB mini-sector size".into());
        }
        if bytes.len() < sector_size || !(bytes.len() - sector_size).is_multiple_of(sector_size) {
            return Err("truncated CFB sector area".into());
        }
        let physical_sectors = (bytes.len() - sector_size) / sector_size;
        if physical_sectors == 0 {
            return Err("empty CFB sector area".into());
        }
        let num_directory_sectors = usize_at_u32(bytes, 40)?;
        match major {
            3 if num_directory_sectors != 0 => {
                return Err("version 3 CFB declares directory sectors".into());
            }
            4 if num_directory_sectors == 0 || num_directory_sectors > physical_sectors => {
                return Err("invalid version 4 CFB directory sector count".into());
            }
            _ => {}
        }
        let num_fat_sectors = usize_at_u32(bytes, 44)?;
        if num_fat_sectors == 0 || num_fat_sectors > physical_sectors {
            return Err("invalid CFB FAT sector count".into());
        }
        let first_directory_sector = u32_at(bytes, 48)?;
        let mini_stream_cutoff = u32_at(bytes, 56)? as u64;
        if mini_stream_cutoff != 4096 {
            return Err("unsupported CFB mini-stream cutoff".into());
        }
        let first_mini_fat_sector = u32_at(bytes, 60)?;
        let num_mini_fat_sectors = usize_at_u32(bytes, 64)?;
        if num_mini_fat_sectors > physical_sectors {
            return Err("invalid CFB MiniFAT sector count".into());
        }
        let first_difat_sector = u32_at(bytes, 68)?;
        let num_difat_sectors = usize_at_u32(bytes, 72)?;
        if num_difat_sectors > physical_sectors {
            return Err("invalid CFB DIFAT sector count".into());
        }

        let initial_fat_capacity = num_fat_sectors.min(109);
        let mut fat_sector_ids = Vec::with_capacity(initial_fat_capacity);
        let mut seen_fat_sector_ids = HashSet::with_capacity(initial_fat_capacity);
        for index in 0..109 {
            let id = u32_at(bytes, 76 + index * 4)?;
            if id != FREE_SECTOR {
                push_sector_id(
                    &mut fat_sector_ids,
                    &mut seen_fat_sector_ids,
                    id,
                    physical_sectors,
                    num_fat_sectors,
                )?;
            }
        }
        let mut difat_sector = first_difat_sector;
        let mut seen_difat = HashSet::new();
        let difat_entries_per_sector = sector_size / 4 - 1;
        for _ in 0..num_difat_sectors {
            validate_sector_id(difat_sector, physical_sectors)?;
            if !seen_difat.insert(difat_sector) {
                return Err("cyclic CFB DIFAT chain".into());
            }
            let sector = sector_slice(bytes, sector_size, physical_sectors, difat_sector)?;
            for index in 0..difat_entries_per_sector {
                let id = u32_at(sector, index * 4)?;
                if id != FREE_SECTOR {
                    push_sector_id(
                        &mut fat_sector_ids,
                        &mut seen_fat_sector_ids,
                        id,
                        physical_sectors,
                        num_fat_sectors,
                    )?;
                }
            }
            difat_sector = u32_at(sector, difat_entries_per_sector * 4)?;
        }
        if num_difat_sectors == 0 && first_difat_sector != END_OF_CHAIN {
            return Err("inconsistent CFB DIFAT header".into());
        }
        if num_difat_sectors > 0 && difat_sector != END_OF_CHAIN {
            return Err("unterminated CFB DIFAT chain".into());
        }
        if fat_sector_ids.len() != num_fat_sectors {
            return Err("inconsistent CFB FAT sector count".into());
        }

        let fat_capacity = num_fat_sectors
            .checked_mul(sector_size / 4)
            .ok_or_else(|| "CFB FAT size overflow".to_string())?;
        if fat_capacity < physical_sectors {
            return Err("CFB FAT does not cover physical sectors".into());
        }
        let mut fat = Vec::with_capacity(fat_capacity.min(physical_sectors + sector_size / 4));
        for id in fat_sector_ids {
            let sector = sector_slice(bytes, sector_size, physical_sectors, id)?;
            for chunk in sector.chunks_exact(4) {
                fat.push(u32::from_le_bytes(
                    chunk.try_into().expect("four-byte chunk"),
                ));
            }
        }

        let directory_bytes = read_regular_chain(
            bytes,
            sector_size,
            physical_sectors,
            &fat,
            first_directory_sector,
            (major == 4).then_some(
                num_directory_sectors
                    .checked_mul(sector_size)
                    .ok_or_else(|| "CFB directory size overflow".to_string())?
                    as u64,
            ),
            Some(MAX_DIRECTORY_ENTRIES * DIRECTORY_ENTRY_BYTES),
        )?;
        let directory = parse_directory(&directory_bytes, major)?;
        let root = directory
            .iter()
            .filter_map(Option::as_ref)
            .find_map(|(_, entry)| (entry.object_type == 5).then_some(*entry))
            .ok_or_else(|| "missing CFB root storage".to_string())?;
        let root_mini_stream = if root.stream_size == 0 {
            Vec::new()
        } else {
            read_regular_chain(
                bytes,
                sector_size,
                physical_sectors,
                &fat,
                root.start_sector,
                Some(root.stream_size),
                None,
            )?
        };

        let mini_fat = if num_mini_fat_sectors == 0 {
            if first_mini_fat_sector != END_OF_CHAIN {
                return Err("inconsistent CFB MiniFAT header".into());
            }
            Vec::new()
        } else {
            let mini_fat_bytes = read_regular_chain(
                bytes,
                sector_size,
                physical_sectors,
                &fat,
                first_mini_fat_sector,
                Some((num_mini_fat_sectors * sector_size) as u64),
                None,
            )?;
            mini_fat_bytes
                .chunks_exact(4)
                .map(|chunk| u32::from_le_bytes(chunk.try_into().expect("four-byte chunk")))
                .collect()
        };

        Ok(Self {
            bytes,
            sector_size,
            mini_sector_size: 1usize << mini_sector_shift,
            mini_stream_cutoff,
            physical_sectors,
            fat,
            mini_fat,
            directory,
            root_mini_stream,
        })
    }

    pub fn has_entry(&self, expected: &str) -> bool {
        self.directory
            .iter()
            .filter_map(Option::as_ref)
            .any(|(name, _)| name.eq_ignore_ascii_case(expected))
    }

    pub fn stream(&self, expected: &str) -> Result<Vec<u8>, String> {
        let mut matches = self
            .directory
            .iter()
            .filter_map(Option::as_ref)
            .filter(|(name, entry)| entry.object_type == 2 && name.eq_ignore_ascii_case(expected));
        let (_, entry) = matches
            .next()
            .ok_or_else(|| format!("missing CFB stream: {expected}"))?;
        if matches.next().is_some() {
            return Err(format!("duplicate CFB stream: {expected}"));
        }
        self.read_stream(*entry)
    }

    /// Validate and index raw MS-CFB directory IDs once for repeated ASCII
    /// parent-scoped lookups. Existing flat lookup does not require this view.
    pub fn scoped_streams(&self) -> Result<ScopedStreams<'_, 'a>, String> {
        let (root, children, has_non_ascii_child) = self.validate_hierarchy()?;
        Ok(ScopedStreams {
            compound: self,
            root,
            children,
            has_non_ascii_child,
        })
    }

    fn read_stream(&self, entry: DirectoryEntry) -> Result<Vec<u8>, String> {
        if entry.stream_size == 0 {
            return Ok(Vec::new());
        }
        if entry.stream_size < self.mini_stream_cutoff {
            self.read_mini_chain(entry.start_sector, entry.stream_size)
        } else {
            read_regular_chain(
                self.bytes,
                self.sector_size,
                self.physical_sectors,
                &self.fat,
                entry.start_sector,
                Some(entry.stream_size),
                None,
            )
        }
    }

    fn validate_hierarchy(
        &self,
    ) -> Result<(usize, Vec<HashMap<String, usize>>, Vec<bool>), String> {
        let roots: Vec<_> = self
            .directory
            .iter()
            .enumerate()
            .filter_map(|(id, slot)| {
                slot.as_ref()
                    .and_then(|(_, entry)| (entry.object_type == 5).then_some(id))
            })
            .collect();
        if roots.len() != 1 {
            return Err("invalid CFB root storage count".into());
        }
        let root = roots[0];
        if root != 0 {
            return Err("CFB root storage is not directory entry 0".into());
        }
        let (_, root_entry) = self.directory[root].as_ref().expect("root was present");
        if root_entry.left_sibling != NO_STREAM || root_entry.right_sibling != NO_STREAM {
            return Err("CFB root storage has siblings".into());
        }
        let mut owner = vec![None; self.directory.len()];
        let mut children = vec![Vec::new(); self.directory.len()];
        for (parent, slot) in self.directory.iter().enumerate() {
            let Some((_, entry)) = slot else { continue };
            if entry.object_type == 2 && entry.child != NO_STREAM {
                return Err("CFB stream has a child".into());
            }
            if !matches!(entry.object_type, 1 | 5) || entry.child == NO_STREAM {
                continue;
            }
            let mut stack = vec![entry.child];
            let mut seen = HashSet::new();
            while let Some(raw) = stack.pop() {
                let id = usize::try_from(raw).map_err(|_| "CFB directory link out of range")?;
                let Some((_, child)) = self.directory.get(id).and_then(Option::as_ref) else {
                    return Err("CFB directory link references an unused or missing entry".into());
                };
                if !seen.insert(id) {
                    return Err("cyclic CFB directory sibling tree".into());
                }
                if id == root {
                    return Err("CFB root storage is owned by another entry".into());
                }
                if owner[id].replace(parent).is_some() {
                    return Err("multiply-owned CFB directory entry".into());
                }
                children[parent].push(id);
                if child.left_sibling != NO_STREAM {
                    stack.push(child.left_sibling);
                }
                if child.right_sibling != NO_STREAM {
                    stack.push(child.right_sibling);
                }
            }
        }
        for (id, slot) in self.directory.iter().enumerate() {
            if id != root && slot.is_some() && owner[id].is_none() {
                return Err("unowned CFB directory entry".into());
            }
        }
        let mut child_index = Vec::with_capacity(children.len());
        let mut has_non_ascii_child = Vec::with_capacity(children.len());
        for siblings in &children {
            let mut names = HashMap::with_capacity(siblings.len());
            let mut has_non_ascii = false;
            for id in siblings {
                let (name, _) = self.directory[*id]
                    .as_ref()
                    .expect("validated hierarchy entry");
                if !name.is_ascii() {
                    has_non_ascii = true;
                } else if names.insert(name.to_ascii_uppercase(), *id).is_some() {
                    return Err("duplicate ASCII CFB child name".into());
                }
            }
            child_index.push(names);
            has_non_ascii_child.push(has_non_ascii);
        }
        let mut reachable = HashSet::with_capacity(self.directory.len());
        let mut stack = vec![root];
        while let Some(parent) = stack.pop() {
            if !reachable.insert(parent) {
                return Err("cyclic CFB storage ancestry".into());
            }
            stack.extend(children[parent].iter().copied());
        }
        if self
            .directory
            .iter()
            .enumerate()
            .any(|(id, slot)| slot.is_some() && !reachable.contains(&id))
        {
            return Err("unreachable CFB directory entry".into());
        }
        Ok((root, child_index, has_non_ascii_child))
    }

    fn read_mini_chain(&self, start: u32, size: u64) -> Result<Vec<u8>, String> {
        let target = usize::try_from(size).map_err(|_| "CFB mini stream too large")?;
        if target > self.root_mini_stream.len() {
            return Err("CFB mini stream exceeds root stream".into());
        }
        let expected_sectors = target.div_ceil(self.mini_sector_size);
        let mut output = Vec::with_capacity(target);
        let mut current = start;
        let mut seen = HashSet::new();
        for _ in 0..expected_sectors {
            let id = usize::try_from(current).map_err(|_| "invalid CFB mini-sector id")?;
            if id >= self.mini_fat.len() || !seen.insert(current) {
                return Err("invalid or cyclic CFB MiniFAT chain".into());
            }
            let offset = id
                .checked_mul(self.mini_sector_size)
                .ok_or_else(|| "CFB mini-sector offset overflow".to_string())?;
            let end = offset
                .checked_add(self.mini_sector_size)
                .ok_or_else(|| "CFB mini-sector end overflow".to_string())?;
            if end > self.root_mini_stream.len() {
                return Err("CFB mini-sector outside root stream".into());
            }
            output.extend_from_slice(&self.root_mini_stream[offset..end]);
            current = self.mini_fat[id];
        }
        if current != END_OF_CHAIN {
            return Err("CFB mini stream chain exceeds declared size".into());
        }
        output.truncate(target);
        Ok(output)
    }
}

fn parse_directory(bytes: &[u8], major: u16) -> Result<Vec<DirectorySlot>, String> {
    if !bytes.len().is_multiple_of(DIRECTORY_ENTRY_BYTES) {
        return Err("misaligned CFB directory stream".into());
    }
    let count = bytes.len() / DIRECTORY_ENTRY_BYTES;
    if count == 0 || count > MAX_DIRECTORY_ENTRIES {
        return Err("invalid CFB directory entry count".into());
    }
    let mut entries = Vec::with_capacity(count);
    for entry_bytes in bytes.chunks_exact(DIRECTORY_ENTRY_BYTES) {
        let object_type = entry_bytes[66];
        if object_type == 0 {
            entries.push(None);
            continue;
        }
        if !matches!(object_type, 1 | 2 | 5) {
            return Err("invalid CFB directory object type".into());
        }
        let name_bytes = u16_at(entry_bytes, 64)? as usize;
        if !(2..=64).contains(&name_bytes) || !name_bytes.is_multiple_of(2) {
            return Err("invalid CFB directory name length".into());
        }
        if entry_bytes[name_bytes - 2..name_bytes] != [0, 0] {
            return Err("unterminated CFB directory name".into());
        }
        let mut utf16 = Vec::with_capacity(name_bytes / 2 - 1);
        for chunk in entry_bytes[..name_bytes - 2].chunks_exact(2) {
            utf16.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        let name = String::from_utf16(&utf16).map_err(|_| "invalid CFB directory name")?;
        if name.contains('\0') {
            return Err("invalid NUL in CFB directory name".into());
        }
        let start_sector = u32_at(entry_bytes, 116)?;
        let mut stream_size = u64_at(entry_bytes, 120)?;
        if major == 3 {
            stream_size &= 0xffff_ffff;
        }
        entries.push(Some((
            name,
            DirectoryEntry {
                object_type,
                left_sibling: u32_at(entry_bytes, 68)?,
                right_sibling: u32_at(entry_bytes, 72)?,
                child: u32_at(entry_bytes, 76)?,
                start_sector,
                stream_size,
            },
        )));
    }
    Ok(entries)
}

fn read_regular_chain(
    bytes: &[u8],
    sector_size: usize,
    physical_sectors: usize,
    fat: &[u32],
    start: u32,
    declared_size: Option<u64>,
    maximum_bytes: Option<usize>,
) -> Result<Vec<u8>, String> {
    let physical_bytes = physical_sectors
        .checked_mul(sector_size)
        .ok_or_else(|| "CFB physical size overflow".to_string())?;
    let maximum_sectors = match declared_size {
        Some(size) => {
            let size = usize::try_from(size).map_err(|_| "CFB stream too large")?;
            if size > physical_bytes {
                return Err("CFB stream size exceeds the physical file".into());
            }
            if maximum_bytes.is_some_and(|maximum| size > maximum) {
                return Err("CFB stream exceeds its implementation limit".into());
            }
            size.div_ceil(sector_size)
        }
        None => maximum_bytes
            .unwrap_or(physical_bytes)
            .div_ceil(sector_size)
            .min(physical_sectors),
    };
    if maximum_sectors == 0 {
        return Ok(Vec::new());
    }
    let capacity = maximum_sectors
        .checked_mul(sector_size)
        .ok_or_else(|| "CFB stream capacity overflow".to_string())?;
    let initial_capacity = if declared_size.is_some() {
        capacity
    } else {
        capacity.min(sector_size)
    };
    let mut output = Vec::with_capacity(initial_capacity);
    let mut current = start;
    let mut seen = HashSet::new();
    for _ in 0..maximum_sectors {
        validate_sector_id(current, physical_sectors)?;
        if !seen.insert(current) {
            return Err("cyclic CFB sector chain".into());
        }
        output.extend_from_slice(sector_slice(bytes, sector_size, physical_sectors, current)?);
        let index = current as usize;
        current = *fat
            .get(index)
            .ok_or_else(|| "CFB FAT index out of range".to_string())?;
        if current == END_OF_CHAIN {
            break;
        }
    }
    if current != END_OF_CHAIN {
        return Err("CFB stream chain exceeds its bound".into());
    }
    if let Some(size) = declared_size {
        let size = usize::try_from(size).map_err(|_| "CFB stream too large")?;
        if output.len() < size {
            return Err("truncated CFB stream".into());
        }
        output.truncate(size);
    }
    Ok(output)
}

fn push_sector_id(
    ids: &mut Vec<u32>,
    seen: &mut HashSet<u32>,
    id: u32,
    physical_sectors: usize,
    maximum: usize,
) -> Result<(), String> {
    if ids.len() >= maximum {
        return Err("too many CFB FAT sector ids".into());
    }
    validate_sector_id(id, physical_sectors)?;
    if !seen.insert(id) {
        return Err("duplicate CFB FAT sector id".into());
    }
    ids.push(id);
    Ok(())
}

fn validate_sector_id(id: u32, physical_sectors: usize) -> Result<(), String> {
    if matches!(id, FREE_SECTOR | END_OF_CHAIN | FAT_SECTOR | DIFAT_SECTOR)
        || id as usize >= physical_sectors
    {
        return Err("invalid CFB sector id".into());
    }
    Ok(())
}

fn sector_slice(
    bytes: &[u8],
    sector_size: usize,
    physical_sectors: usize,
    id: u32,
) -> Result<&[u8], String> {
    validate_sector_id(id, physical_sectors)?;
    let offset = (id as usize + 1)
        .checked_mul(sector_size)
        .ok_or_else(|| "CFB sector offset overflow".to_string())?;
    let end = offset
        .checked_add(sector_size)
        .ok_or_else(|| "CFB sector end overflow".to_string())?;
    bytes
        .get(offset..end)
        .ok_or_else(|| "truncated CFB sector".into())
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, String> {
    let raw = bytes
        .get(offset..offset + 2)
        .ok_or("truncated CFB integer")?;
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let raw = bytes
        .get(offset..offset + 4)
        .ok_or("truncated CFB integer")?;
    Ok(u32::from_le_bytes(raw.try_into().expect("four-byte slice")))
}

fn u64_at(bytes: &[u8], offset: usize) -> Result<u64, String> {
    let raw = bytes
        .get(offset..offset + 8)
        .ok_or("truncated CFB integer")?;
    Ok(u64::from_le_bytes(
        raw.try_into().expect("eight-byte slice"),
    ))
}

fn usize_at_u32(bytes: &[u8], offset: usize) -> Result<usize, String> {
    usize::try_from(u32_at(bytes, offset)?).map_err(|_| "CFB count too large".into())
}

#[cfg(any(test, feature = "fuzzing"))]
pub(crate) mod test_support {
    use super::{
        DIFAT_SECTOR, END_OF_CHAIN, FAT_SECTOR, FREE_SECTOR, HEADER_BYTES, NO_STREAM, SIGNATURE,
    };

    /// Build a version-3 CFB with regular (non-mini) streams for parser tests.
    /// Padding streams to 4096 bytes keeps the fixture writer intentionally tiny.
    pub fn build_cfb(streams: &[(&str, Vec<u8>)]) -> Vec<u8> {
        let sector_size = 512usize;
        let padded: Vec<_> = streams
            .iter()
            .map(|(name, bytes)| {
                let declared = bytes.len().max(4096);
                let sectors = declared.div_ceil(sector_size);
                (*name, bytes, declared, sectors)
            })
            .collect();
        let directory_sectors = ((streams.len() + 1) * 128).div_ceil(sector_size).max(1);
        let data_sectors: usize = padded.iter().map(|entry| entry.3).sum();
        let mut fat_sectors = 1usize;
        while data_sectors + directory_sectors + fat_sectors > fat_sectors * 128 {
            fat_sectors += 1;
        }
        assert!(fat_sectors <= 109);
        let total_sectors = data_sectors + directory_sectors + fat_sectors;
        let mut bytes = vec![0u8; HEADER_BYTES + total_sectors * sector_size];
        bytes[..8].copy_from_slice(&SIGNATURE);
        bytes[24..26].copy_from_slice(&0x003eu16.to_le_bytes());
        bytes[26..28].copy_from_slice(&3u16.to_le_bytes());
        bytes[28..30].copy_from_slice(&0xfffeu16.to_le_bytes());
        bytes[30..32].copy_from_slice(&9u16.to_le_bytes());
        bytes[32..34].copy_from_slice(&6u16.to_le_bytes());
        bytes[44..48].copy_from_slice(&(fat_sectors as u32).to_le_bytes());
        bytes[48..52].copy_from_slice(&(data_sectors as u32).to_le_bytes());
        bytes[56..60].copy_from_slice(&4096u32.to_le_bytes());
        bytes[60..64].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        bytes[68..72].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        for index in 0..109 {
            let value = if index < fat_sectors {
                (data_sectors + directory_sectors + index) as u32
            } else {
                FREE_SECTOR
            };
            bytes[76 + index * 4..80 + index * 4].copy_from_slice(&value.to_le_bytes());
        }

        let mut sector = 0usize;
        let mut stream_starts = Vec::new();
        for (_, source, declared, sectors) in &padded {
            stream_starts.push((sector as u32, *declared as u64));
            let start = HEADER_BYTES + sector * sector_size;
            bytes[start..start + source.len()].copy_from_slice(source);
            sector += sectors;
        }
        let directory_start = sector;
        sector += directory_sectors;
        let fat_start = sector;

        let directory_offset = HEADER_BYTES + directory_start * sector_size;
        write_directory_entry(
            &mut bytes[directory_offset..directory_offset + 128],
            "Root Entry",
            5,
            END_OF_CHAIN,
            0,
        );
        for (index, ((name, _, _, _), (start, size))) in
            padded.iter().zip(stream_starts.iter()).enumerate()
        {
            let offset = directory_offset + (index + 1) * 128;
            write_directory_entry(&mut bytes[offset..offset + 128], name, 2, *start, *size);
        }

        let mut fat = vec![FREE_SECTOR; fat_sectors * 128];
        let mut cursor = 0usize;
        for (_, _, _, sectors) in &padded {
            for index in 0..*sectors {
                fat[cursor + index] = if index + 1 == *sectors {
                    END_OF_CHAIN
                } else {
                    (cursor + index + 1) as u32
                };
            }
            cursor += sectors;
        }
        for index in 0..directory_sectors {
            fat[directory_start + index] = if index + 1 == directory_sectors {
                END_OF_CHAIN
            } else {
                (directory_start + index + 1) as u32
            };
        }
        for index in 0..fat_sectors {
            fat[fat_start + index] = FAT_SECTOR;
            let offset = HEADER_BYTES + (fat_start + index) * sector_size;
            for entry in 0..128 {
                bytes[offset + entry * 4..offset + entry * 4 + 4]
                    .copy_from_slice(&fat[index * 128 + entry].to_le_bytes());
            }
        }
        let _ = DIFAT_SECTOR;
        bytes
    }

    #[cfg(test)]
    pub fn build_mini_cfb(name: &str, source: &[u8]) -> Vec<u8> {
        assert!(!source.is_empty() && source.len() <= 64);
        let sector_size = 512usize;
        let mut bytes = vec![0u8; HEADER_BYTES + 4 * sector_size];
        bytes[..8].copy_from_slice(&SIGNATURE);
        bytes[24..26].copy_from_slice(&0x003eu16.to_le_bytes());
        bytes[26..28].copy_from_slice(&3u16.to_le_bytes());
        bytes[28..30].copy_from_slice(&0xfffeu16.to_le_bytes());
        bytes[30..32].copy_from_slice(&9u16.to_le_bytes());
        bytes[32..34].copy_from_slice(&6u16.to_le_bytes());
        bytes[44..48].copy_from_slice(&1u32.to_le_bytes());
        bytes[48..52].copy_from_slice(&2u32.to_le_bytes());
        bytes[56..60].copy_from_slice(&4096u32.to_le_bytes());
        bytes[60..64].copy_from_slice(&1u32.to_le_bytes());
        bytes[64..68].copy_from_slice(&1u32.to_le_bytes());
        bytes[68..72].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        bytes[76..80].copy_from_slice(&3u32.to_le_bytes());
        for index in 1..109 {
            bytes[76 + index * 4..80 + index * 4].copy_from_slice(&FREE_SECTOR.to_le_bytes());
        }

        bytes[HEADER_BYTES..HEADER_BYTES + source.len()].copy_from_slice(source);
        let mini_fat_offset = HEADER_BYTES + sector_size;
        bytes[mini_fat_offset..mini_fat_offset + 4].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        for index in 1..128 {
            bytes[mini_fat_offset + index * 4..mini_fat_offset + index * 4 + 4]
                .copy_from_slice(&FREE_SECTOR.to_le_bytes());
        }
        let directory_offset = HEADER_BYTES + 2 * sector_size;
        write_directory_entry(
            &mut bytes[directory_offset..directory_offset + 128],
            "Root Entry",
            5,
            0,
            64,
        );
        write_directory_entry(
            &mut bytes[directory_offset + 128..directory_offset + 256],
            name,
            2,
            0,
            source.len() as u64,
        );
        let fat_offset = HEADER_BYTES + 3 * sector_size;
        for (index, value) in [END_OF_CHAIN, END_OF_CHAIN, END_OF_CHAIN, FAT_SECTOR]
            .into_iter()
            .enumerate()
        {
            bytes[fat_offset + index * 4..fat_offset + index * 4 + 4]
                .copy_from_slice(&value.to_le_bytes());
        }
        for index in 4..128 {
            bytes[fat_offset + index * 4..fat_offset + index * 4 + 4]
                .copy_from_slice(&FREE_SECTOR.to_le_bytes());
        }
        bytes
    }

    #[cfg(test)]
    pub fn build_v4_cfb(name: &str, source: &[u8]) -> Vec<u8> {
        let sector_size = 4096usize;
        assert!(source.len() <= sector_size);
        let mut bytes = vec![0u8; sector_size * 4];
        bytes[..8].copy_from_slice(&SIGNATURE);
        bytes[24..26].copy_from_slice(&0x003eu16.to_le_bytes());
        bytes[26..28].copy_from_slice(&4u16.to_le_bytes());
        bytes[28..30].copy_from_slice(&0xfffeu16.to_le_bytes());
        bytes[30..32].copy_from_slice(&12u16.to_le_bytes());
        bytes[32..34].copy_from_slice(&6u16.to_le_bytes());
        bytes[40..44].copy_from_slice(&1u32.to_le_bytes());
        bytes[44..48].copy_from_slice(&1u32.to_le_bytes());
        bytes[48..52].copy_from_slice(&1u32.to_le_bytes());
        bytes[56..60].copy_from_slice(&4096u32.to_le_bytes());
        bytes[60..64].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        bytes[68..72].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        bytes[76..80].copy_from_slice(&2u32.to_le_bytes());
        for index in 1..109 {
            bytes[76 + index * 4..80 + index * 4].copy_from_slice(&FREE_SECTOR.to_le_bytes());
        }

        bytes[sector_size..sector_size + source.len()].copy_from_slice(source);
        let directory_offset = sector_size * 2;
        write_directory_entry(
            &mut bytes[directory_offset..directory_offset + 128],
            "Root Entry",
            5,
            END_OF_CHAIN,
            0,
        );
        write_directory_entry(
            &mut bytes[directory_offset + 128..directory_offset + 256],
            name,
            2,
            0,
            sector_size as u64,
        );
        let fat_offset = sector_size * 3;
        for (index, value) in [END_OF_CHAIN, END_OF_CHAIN, FAT_SECTOR]
            .into_iter()
            .enumerate()
        {
            bytes[fat_offset + index * 4..fat_offset + index * 4 + 4]
                .copy_from_slice(&value.to_le_bytes());
        }
        for index in 3..sector_size / 4 {
            bytes[fat_offset + index * 4..fat_offset + index * 4 + 4]
                .copy_from_slice(&FREE_SECTOR.to_le_bytes());
        }
        bytes
    }

    fn write_directory_entry(
        target: &mut [u8],
        name: &str,
        object_type: u8,
        start_sector: u32,
        size: u64,
    ) {
        let utf16: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        for (index, unit) in utf16.iter().enumerate() {
            target[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
        }
        target[64..66].copy_from_slice(&((utf16.len() * 2) as u16).to_le_bytes());
        target[66] = object_type;
        target[67] = 1;
        target[68..72].copy_from_slice(&NO_STREAM.to_le_bytes());
        target[72..76].copy_from_slice(&NO_STREAM.to_le_bytes());
        target[76..80].copy_from_slice(&NO_STREAM.to_le_bytes());
        target[116..120].copy_from_slice(&start_sector.to_le_bytes());
        target[120..128].copy_from_slice(&size.to_le_bytes());
    }
}

#[cfg(test)]
mod tests {
    use super::{
        test_support::{build_cfb, build_mini_cfb, build_v4_cfb},
        CompoundFile, END_OF_CHAIN,
    };

    fn directory_offset(streams: usize) -> usize {
        // build_cfb pads every synthetic regular stream to eight sectors.
        512 + streams * 8 * 512
    }

    fn link(bytes: &mut [u8], streams: usize, id: usize, at: usize, target: u32) {
        let offset = directory_offset(streams) + id * 128 + at;
        bytes[offset..offset + 4].copy_from_slice(&target.to_le_bytes());
    }

    fn object_type(bytes: &mut [u8], streams: usize, id: usize, value: u8) {
        let offset = directory_offset(streams) + id * 128;
        bytes[offset + 66] = value;
        if value == 1 {
            bytes[offset + 116..offset + 120].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
            bytes[offset + 120..offset + 128].fill(0);
        }
    }

    fn nested_presentations() -> Vec<u8> {
        let streams = [
            ("ObjectPool", Vec::new()),
            ("_1", Vec::new()),
            ("OlePres000", b"first presentation".to_vec()),
            ("_2", Vec::new()),
            ("OlePres000", b"second presentation".to_vec()),
        ];
        let mut bytes = build_cfb(&streams);
        object_type(&mut bytes, streams.len(), 1, 1);
        object_type(&mut bytes, streams.len(), 2, 1);
        object_type(&mut bytes, streams.len(), 4, 1);
        link(&mut bytes, streams.len(), 0, 76, 1);
        link(&mut bytes, streams.len(), 1, 76, 2);
        link(&mut bytes, streams.len(), 2, 72, 4);
        link(&mut bytes, streams.len(), 2, 76, 3);
        link(&mut bytes, streams.len(), 4, 76, 5);
        bytes
    }

    #[test]
    fn reads_regular_streams() {
        let source = b"hello compound file".to_vec();
        let bytes = build_cfb(&[("Workbook", source.clone())]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert!(cfb.has_entry("workbook"));
        let stream = cfb.stream("Workbook").unwrap();
        assert_eq!(&stream[..source.len()], source);
    }

    #[test]
    fn parent_scoped_lookup_distinguishes_duplicate_stream_names() {
        let bytes = nested_presentations();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let scoped = cfb.scoped_streams().unwrap();
        assert_eq!(
            &scoped
                .stream(&["objectpool", "_1", "olepres000"], 4096)
                .unwrap()[..18],
            b"first presentation"
        );
        assert_eq!(
            &scoped
                .stream(&["ObjectPool", "_2", "OlePres000"], 4096)
                .unwrap()[..19],
            b"second presentation"
        );
        assert!(scoped
            .stream(&["ObjectPool", "_1", "OlePres000"], 4095)
            .unwrap_err()
            .contains("caller limit"));
        assert!(cfb.stream("OlePres000").unwrap_err().contains("duplicate"));
    }

    #[test]
    fn scoped_lookup_rejects_ambiguous_sibling_names() {
        let streams = [
            ("ObjectPool", Vec::new()),
            ("_1", Vec::new()),
            ("OlePres000", b"first".to_vec()),
            ("OlePres000", b"second".to_vec()),
        ];
        let mut bytes = build_cfb(&streams);
        object_type(&mut bytes, streams.len(), 1, 1);
        object_type(&mut bytes, streams.len(), 2, 1);
        link(&mut bytes, streams.len(), 0, 76, 1);
        link(&mut bytes, streams.len(), 1, 76, 2);
        link(&mut bytes, streams.len(), 2, 76, 3);
        link(&mut bytes, streams.len(), 3, 72, 4);
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert!(cfb.scoped_streams().err().unwrap().contains("duplicate"));
    }

    #[test]
    fn scoped_lookup_rejects_non_ascii_only_on_the_selected_parent() {
        let streams = [
            ("ObjectPool", Vec::new()),
            ("Target", b"target".to_vec()),
            ("Other", Vec::new()),
            ("Ä", b"non-ascii sibling".to_vec()),
        ];
        let mut bytes = build_cfb(&streams);
        object_type(&mut bytes, streams.len(), 1, 1);
        object_type(&mut bytes, streams.len(), 3, 1);
        link(&mut bytes, streams.len(), 0, 76, 1);
        link(&mut bytes, streams.len(), 1, 72, 3);
        link(&mut bytes, streams.len(), 1, 76, 2);
        link(&mut bytes, streams.len(), 3, 76, 4);

        let cfb = CompoundFile::open(&bytes).unwrap();
        let scoped = cfb.scoped_streams().unwrap();
        assert_eq!(
            &scoped.stream(&["ObjectPool", "Target"], 4096).unwrap()[..6],
            b"target"
        );
        assert!(scoped
            .stream(&["Other", "missing"], 4096)
            .unwrap_err()
            .contains("non-ASCII sibling"));

        link(&mut bytes, streams.len(), 2, 72, 4);
        link(&mut bytes, streams.len(), 3, 76, u32::MAX);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let scoped = cfb.scoped_streams().unwrap();
        assert!(scoped
            .stream(&["ObjectPool", "Target"], 4096)
            .unwrap_err()
            .contains("non-ASCII sibling"));
    }

    #[test]
    fn scoped_lookup_rejects_cycles_unused_links_multiple_owners_and_stream_children() {
        for (id, field, target, expected) in [
            (2, 72, 2, "cyclic"),
            (2, 72, 7, "unused or missing"),
            (4, 76, 3, "multiply-owned"),
            (3, 76, 5, "stream has a child"),
        ] {
            let mut bytes = nested_presentations();
            link(&mut bytes, 5, id, field, target);
            let cfb = CompoundFile::open(&bytes).unwrap();
            let error = cfb.scoped_streams().err().unwrap();
            assert!(error.contains(expected), "{error}");
        }
    }

    #[test]
    fn scoped_lookup_retains_raw_ids_across_unused_slots() {
        let mut bytes = nested_presentations();
        let directory = directory_offset(5);
        let moved: [u8; 128] = bytes[directory + 5 * 128..directory + 6 * 128]
            .try_into()
            .unwrap();
        bytes[directory + 5 * 128..directory + 6 * 128].fill(0);
        bytes[directory + 6 * 128..directory + 7 * 128].copy_from_slice(&moved);
        link(&mut bytes, 5, 4, 76, 6);

        let cfb = CompoundFile::open(&bytes).unwrap();
        let scoped = cfb.scoped_streams().unwrap();
        assert_eq!(
            &scoped
                .stream(&["ObjectPool", "_2", "OlePres000"], 4096)
                .unwrap()[..19],
            b"second presentation"
        );
    }

    #[test]
    fn scoped_lookup_rejects_disconnected_storage_cycles() {
        let streams = [
            ("Visible", b"visible".to_vec()),
            ("DetachedA", Vec::new()),
            ("DetachedB", Vec::new()),
        ];
        let mut bytes = build_cfb(&streams);
        object_type(&mut bytes, streams.len(), 2, 1);
        object_type(&mut bytes, streams.len(), 3, 1);
        link(&mut bytes, streams.len(), 0, 76, 1);
        link(&mut bytes, streams.len(), 2, 76, 3);
        link(&mut bytes, streams.len(), 3, 76, 2);

        let cfb = CompoundFile::open(&bytes).unwrap();
        let error = cfb.scoped_streams().err().unwrap();
        assert!(error.contains("unreachable"), "{error}");
    }

    #[test]
    fn scoped_lookup_requires_root_at_directory_id_zero() {
        let streams = [("Storage", Vec::new()), ("Payload", b"payload".to_vec())];
        let mut bytes = build_cfb(&streams);
        object_type(&mut bytes, streams.len(), 0, 1);
        object_type(&mut bytes, streams.len(), 1, 5);
        link(&mut bytes, streams.len(), 1, 76, 2);

        let cfb = CompoundFile::open(&bytes).unwrap();
        let error = cfb.scoped_streams().err().unwrap();
        assert!(error.contains("entry 0"), "{error}");
    }

    #[test]
    fn legacy_flat_lookup_still_accepts_unlinked_synthetic_directories() {
        let bytes = build_cfb(&[("Workbook", b"legacy flat lookup".to_vec())]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert!(cfb.has_entry("Workbook"));
        assert_eq!(
            &cfb.stream("Workbook").unwrap()[..18],
            b"legacy flat lookup"
        );
        assert!(cfb.scoped_streams().is_err());
    }

    #[test]
    fn accepts_the_ms_cfb_header_signature() {
        let mut bytes = build_cfb(&[("Workbook", vec![0; 16])]);
        bytes[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
        assert!(CompoundFile::open(&bytes).is_ok());
    }

    #[test]
    fn rejects_fat_counts_beyond_the_physical_file() {
        let mut bytes = build_cfb(&[("Workbook", vec![0; 16])]);
        bytes[44..48].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(CompoundFile::open(&bytes).is_err());
    }

    #[test]
    fn reads_mini_fat_streams() {
        let source = b"small workbook stream";
        let bytes = build_mini_cfb("Workbook", source);
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert_eq!(cfb.stream("Workbook").unwrap(), source);
    }

    #[test]
    fn reads_version_4_regular_streams() {
        let source = b"version four workbook";
        let bytes = build_v4_cfb("Workbook", source);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let stream = cfb.stream("Workbook").unwrap();
        assert_eq!(&stream[..source.len()], source);
    }

    #[test]
    fn rejects_inconsistent_directory_sector_counts() {
        let mut version_3 = build_cfb(&[("Workbook", vec![0; 16])]);
        version_3[40..44].copy_from_slice(&1u32.to_le_bytes());
        assert!(CompoundFile::open(&version_3).is_err());

        let mut version_4 = build_v4_cfb("Workbook", b"workbook");
        version_4[40..44].copy_from_slice(&0u32.to_le_bytes());
        assert!(CompoundFile::open(&version_4).is_err());
    }

    #[test]
    fn rejects_declared_stream_sizes_before_large_allocation() {
        let mut bytes = build_cfb(&[("Workbook", vec![0; 16])]);
        let directory_offset = 512 + 8 * 512;
        bytes[directory_offset + 128 + 120..directory_offset + 128 + 128]
            .copy_from_slice(&u64::MAX.to_le_bytes());
        let cfb = CompoundFile::open(&bytes).unwrap();
        assert!(cfb.stream("Workbook").is_err());
    }
}
