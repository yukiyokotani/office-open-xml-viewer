//! Bounded passive image extraction for XLS drawing integration.
//! MS-XLS 2.1.7.20.3 (including implementation note 6), 2.4.58/171;
//! MS-ODRAW 2.2.12/20/22. No shape placement or active-object evaluation.
#[cfg(any(test, all(feature = "inspection", not(target_arch = "wasm32"))))]
use super::{u16_at, BOF, FILEPASS};
use super::{unsupported, Record, EOF};
use crate::officeart::{raster, record_with_end};

const MAX_DRAWING_BYTES: usize = 128 * 1024 * 1024;
const MAX_MEDIA_BYTES: usize = 128 * 1024 * 1024;

/// Resolve only admitted, owned references. Unused catalog images are neither
/// inflated nor validated; malformed *referenced* images still fail closed.
pub(super) fn selected(
    records: &[Record<'_>],
    indices: &std::collections::BTreeSet<u32>,
) -> Result<Vec<(u32, &'static str, Vec<u8>)>, String> {
    if indices.is_empty() {
        return Ok(Vec::new());
    }
    let mut work = 2_000_000;
    let bytes = assemble(records, MAX_DRAWING_BYTES, &mut work)?
        .ok_or_else(|| unsupported("missing referenced BIFF image store"))?;
    let entries = catalog(&bytes, &mut work)?;
    let mut remaining = MAX_MEDIA_BYTES;
    let mut output = Vec::new();
    for &index in indices {
        let entry = index
            .checked_sub(1)
            .and_then(|i| entries.get(i as usize))
            .ok_or_else(|| unsupported("BIFF picture index out of range"))?;
        if let Some(image) = raster::read_store_entry(*entry, None, &mut work, remaining)? {
            remaining = remaining
                .checked_sub(image.bytes.len())
                .ok_or_else(|| unsupported("BIFF retained image budget exceeded"))?;
            output.push((index, image.extension, image.bytes.into_owned()));
        }
    }
    Ok(output)
}

/// Native-only catalog inspection. Production conversion uses owned, selected
/// references instead of exposing every global catalog entry.
#[cfg(any(test, all(feature = "inspection", not(target_arch = "wasm32"))))]
pub(super) fn images(records: &[Record<'_>]) -> Result<Vec<(u32, &'static str, Vec<u8>)>, String> {
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF workbook"))?;
    if first.kind != BOF || u16_at(first.data, 0)? != 0x0600 || u16_at(first.data, 2)? != 5 {
        return Err(unsupported(
            "image inspection requires BIFF8 workbook globals",
        ));
    }
    if records.iter().any(|r| r.kind == FILEPASS) {
        return Err(unsupported("encrypted BIFF image store"));
    }
    let mut work = 2_000_000;
    let Some(bytes) = assemble(records, MAX_DRAWING_BYTES, &mut work)? else {
        return Ok(Vec::new());
    };
    extract(&bytes, &mut work, MAX_MEDIA_BYTES)
}

#[cfg(any(test, all(feature = "inspection", not(target_arch = "wasm32"))))]
fn extract(
    bytes: &[u8],
    work: &mut usize,
    mut remaining: usize,
) -> Result<Vec<(u32, &'static str, Vec<u8>)>, String> {
    let entries = catalog(bytes, work)?;
    let mut output = Vec::new();
    for (index, entry) in entries.into_iter().enumerate() {
        // XLS global BStore data is inline. A delayed reference must never be
        // reinterpreted as a workbook offset, pathname or external request.
        if let Some(image) = raster::read_store_entry(entry, None, work, remaining)? {
            remaining = remaining
                .checked_sub(image.bytes.len())
                .ok_or_else(|| unsupported("BIFF retained image budget exceeded"))?;
            output.push((index as u32 + 1, image.extension, image.bytes.into_owned()));
        }
    }
    Ok(output)
}

fn assemble(
    records: &[Record<'_>],
    max_bytes: usize,
    work: &mut usize,
) -> Result<Option<Vec<u8>>, String> {
    let end = records
        .iter()
        .position(|r| r.kind == EOF)
        .ok_or_else(|| unsupported("missing BIFF global EOF"))?;
    let globals = &records[..end];
    let Some(start) = globals.iter().position(|r| r.kind == 0x00eb) else {
        return Ok(None);
    };
    let mut next = start;
    let mut length = 0usize;
    while let Some(record) = globals.get(next) {
        let base = next == start;
        // Documented Office extension: only the FIRST continuation may be
        // serialized as another MsoDrawingGroup. Later/disjoint duplicates
        // cannot silently append a different global store.
        if !base && record.kind != 0x003c && !(next == start + 1 && record.kind == 0x00eb) {
            break;
        }
        *work = work
            .checked_sub(1)
            .ok_or_else(|| unsupported("BIFF drawing work budget exceeded"))?;
        if record.data.len() > 8224 {
            return Err(unsupported("oversized BIFF drawing fragment"));
        }
        length = length
            .checked_add(record.data.len())
            .filter(|n| *n <= max_bytes)
            .ok_or_else(|| unsupported("BIFF drawing byte budget exceeded"))?;
        next += 1;
    }
    if globals[next..].iter().any(|r| r.kind == 0x00eb) {
        return Err(unsupported("multiple BIFF drawing groups"));
    }
    // Sum/validate first, allocate once: small fragments cannot induce repeated
    // reallocations or bypass the total byte budget.
    let mut bytes = Vec::with_capacity(length);
    for record in &globals[start..next] {
        bytes.extend_from_slice(record.data);
    }
    Ok(Some(bytes))
}

fn catalog<'a>(
    bytes: &'a [u8],
    work: &mut usize,
) -> Result<Vec<crate::officeart::Record<'a>>, String> {
    let (root, end) = record_with_end(bytes, 0, work, "XLS OfficeArt")?;
    if root.kind != 0xf000 || root.version != 15 || root.instance != 0 || end != bytes.len() {
        return Err(unsupported("invalid BIFF drawing group container"));
    }
    let mut found = None;
    let mut pos = 0;
    while pos < root.payload.len() {
        let (child, end) = record_with_end(root.payload, pos, work, "XLS OfficeArt")?;
        pos = end;
        if child.kind != 0xf001 {
            continue;
        }
        if found.is_some() || child.version != 15 {
            return Err(unsupported("ambiguous BIFF image store"));
        }
        let mut entries = Vec::with_capacity(usize::from(child.instance));
        let mut at = 0;
        while at < child.payload.len() {
            let (entry, end) = record_with_end(child.payload, at, work, "XLS OfficeArt")?;
            at = end;
            if entry.kind != 0xf007 && !(0xf018..=0xf117).contains(&entry.kind) {
                return Err(unsupported("invalid BIFF image store record"));
            }
            if entries.len() >= usize::from(child.instance) {
                return Err(unsupported("BIFF image store count mismatch"));
            }
            entries.push(entry);
        }
        if entries.len() != usize::from(child.instance) {
            return Err(unsupported("BIFF image store count mismatch"));
        }
        found = Some(entries);
    }
    Ok(found.unwrap_or_default())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selected_references_are_deduplicated_and_unused_invalid_blips_are_not_decoded() {
        let good = art(0xf01e, 0x6e00, &[vec![0; 17], png()].concat());
        let invalid = art(0xf01a, 0x3d40, &[]);
        let bytes = art(0xf000, 15, &art(0xf001, 0x2f, &[good, invalid].concat()));
        let records = [
            record(BOF, &[0, 6, 5, 0]),
            record(0xeb, &bytes),
            record(EOF, &[]),
        ];
        let selected_once = std::collections::BTreeSet::from([1, 1]);
        assert_eq!(
            selected(&records, &selected_once).unwrap(),
            vec![(1, "png", png())]
        );
        assert!(selected(&records, &std::collections::BTreeSet::from([2])).is_err());
        for index in [0, 3, u32::MAX] {
            assert!(selected(&records, &std::collections::BTreeSet::from([index])).is_err());
        }
        assert!(selected(&records, &std::collections::BTreeSet::new())
            .unwrap()
            .is_empty());
        assert!(selected(
            &[record(BOF, &[0, 6, 5, 0]), record(EOF, &[])],
            &selected_once
        )
        .is_err());
    }
    fn record(kind: u16, data: &[u8]) -> Record<'_> {
        Record {
            kind,
            offset: 0,
            data,
        }
    }
    fn art(kind: u16, options: u16, body: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().to_vec(),
            kind.to_le_bytes().to_vec(),
            (body.len() as u32).to_le_bytes().to_vec(),
            body.to_vec(),
        ]
        .concat()
    }
    fn png() -> Vec<u8> {
        [
            b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".as_slice(),
            &1u32.to_be_bytes(),
            &2u32.to_be_bytes(),
            &[8, 6, 0, 0, 0, 0, 0, 0, 0],
        ]
        .concat()
    }
    fn group() -> Vec<u8> {
        let blip = art(0xf01e, 0x6e00, &[vec![0; 17], png()].concat());
        art(0xf000, 15, &art(0xf001, 31, &blip))
    }
    fn bse(blip: &[u8], embedded: bool, references: u32) -> Vec<u8> {
        let mut body = vec![0; 36];
        body[0] = 6;
        body[1] = 6;
        body[20..24].copy_from_slice(&(blip.len() as u32).to_le_bytes());
        body[24..28].copy_from_slice(&references.to_le_bytes());
        body[28..32].copy_from_slice(&12345u32.to_le_bytes());
        if embedded {
            body.extend_from_slice(blip);
        }
        art(0xf007, 0x62, &body)
    }
    #[test]
    fn embedded_store_entries_preserve_indices_without_resolving_delayed_or_unused_slots() {
        let blip = art(0xf01e, 0x6e00, &[vec![0; 17], png()].concat());
        let entries = [
            bse(&blip, false, 1), // No delayed stream exists in this path.
            bse(&blip, true, 0),  // Unreferenced slot remains unexposed.
            art(0xf01c, 0, &[]),  // Unsupported PICT is not relabeled.
            bse(&blip, true, 1),
        ]
        .concat();
        let data = art(0xf000, 15, &art(0xf001, 0x4f, &entries));
        assert_eq!(
            extract(&data, &mut 100, png().len()).unwrap(),
            vec![(4, "png", png())]
        );
        let mut malformed = bse(&blip, true, 1);
        malformed[8 + 20] ^= 1;
        assert!(extract(
            &art(0xf000, 15, &art(0xf001, 31, &malformed)),
            &mut 100,
            1000
        )
        .is_err());
    }
    #[test]
    fn compressed_metafiles_use_the_shared_passive_validator_and_total_retention_budget() {
        for (source, blip, extension) in [
            {
                let (s, b) = crate::officeart::emf_test_blip();
                (s, b, "emf")
            },
            {
                let (s, b) = crate::officeart::wmf_test_blip();
                (s, b, "wmf")
            },
        ] {
            let bytes = art(
                0xf000,
                15,
                &art(0xf001, 0x2f, &[blip.clone(), blip].concat()),
            );
            assert_eq!(
                extract(&bytes, &mut 100, 2 * source.len()).unwrap(),
                vec![
                    (1, extension, source.clone()),
                    (2, extension, source.clone())
                ]
            );
            assert!(extract(&bytes, &mut 100, 2 * source.len() - 1).is_err());
            assert!(extract(&bytes, &mut 100, 0).is_err());
            assert!(extract(&bytes, &mut 5, 2 * source.len()).is_err());
        }
    }
    #[test]
    fn extracts_identical_passive_bytes_across_every_split_and_documented_first_continuation() {
        let data = group();
        for cut in 1..data.len() {
            for kind in [0x003c, 0x00eb] {
                let records = [
                    record(BOF, &[0, 6, 5, 0]),
                    record(0xeb, &data[..cut]),
                    record(kind, &data[cut..]),
                    record(EOF, &[]),
                ];
                assert_eq!(images(&records).unwrap(), vec![(1, "png", png())]);
            }
        }
    }
    #[test]
    fn no_global_store_does_not_consume_sheet_or_unrelated_continue_payloads() {
        let data = group();
        assert!(images(&[
            record(BOF, &[0, 6, 5, 0]),
            record(EOF, &[]),
            record(0xeb, &data)
        ])
        .unwrap()
        .is_empty());
        let records = [
            record(BOF, &[0, 6, 5, 0]),
            record(0xeb, &data),
            record(0x1234, &[]),
            record(0x3c, &[255]),
            record(EOF, &[]),
        ];
        assert_eq!(images(&records).unwrap(), vec![(1, "png", png())]);
    }
    #[test]
    fn ambiguous_fragments_and_resource_budgets_fail_closed() {
        let data = group();
        for middle in [
            vec![
                record(0xeb, &data),
                record(0x1234, &[]),
                record(0xeb, &data),
            ],
            vec![
                record(0xeb, &data[..4]),
                record(0x3c, &data[4..8]),
                record(0xeb, &data[8..]),
            ],
        ] {
            let records = [
                vec![record(BOF, &[0, 6, 5, 0])],
                middle,
                vec![record(EOF, &[])],
            ]
            .concat();
            assert!(images(&records).is_err());
        }
        let records = [record(0xeb, &data), record(EOF, &[])];
        assert!(assemble(&records, data.len() - 1, &mut 10).is_err());
        assert!(assemble(&records, data.len(), &mut 0).is_err());
        assert!(assemble(
            &[record(0xeb, &vec![0; 8225]), record(EOF, &[])],
            10000,
            &mut 10
        )
        .is_err());
        for cut in 0..data.len() {
            assert!(catalog(&data[..cut], &mut 100).is_err());
        }
        assert!(catalog(&data, &mut 2).is_err());
    }
    #[test]
    fn store_count_ownership_and_encryption_are_checked() {
        let data = group();
        for bytes in [
            art(
                0xf000,
                15,
                &[art(0xf001, 15, &[]), art(0xf001, 15, &[])].concat(),
            ),
            art(0xf000, 15, &art(0xf001, 31, &[])),
            art(0xf000, 15, &art(0xf001, 31, &art(0xf011, 0, &[]))),
            [data.clone(), data.clone()].concat(),
        ] {
            assert!(catalog(&bytes, &mut 100).is_err());
        }
        assert!(images(&[
            record(BOF, &[0, 6, 5, 0]),
            record(0xeb, &data),
            record(FILEPASS, &[]),
            record(EOF, &[])
        ])
        .is_err());
        assert!(images(&[record(BOF, &[0, 5, 5, 0]), record(EOF, &[])]).is_err());
    }
}
