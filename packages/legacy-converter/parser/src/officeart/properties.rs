//! Borrowed OfficeArt property-table framing (MS-ODRAW 2.2.7-9).
//! Interpretation, duplicate policy and inheritance belong to the host format.
use super::{unsupported, Record};

pub(crate) struct Property<'a> {
    /// Full encoded property ID, including fBid and fComplex.
    pub opid: u16,
    pub value: u32,
    pub complex: Option<&'a [u8]>,
}

/// Walk every entry and validate the complete complex-data tail without
/// allocating or decoding strings/actions. Callers must discard partial state
/// on error, including an error discovered after the final callback.
pub(crate) fn visit<'a>(
    record: Record<'a>,
    budget: &mut usize,
    visitor: impl FnMut(Property<'a>) -> Result<(), String>,
) -> Result<(), String> {
    visit_kind(record, 0xf00b, budget, visitor)
}

/// Walk a shape-owned OfficeArtTertiaryFOPT table (MS-ODRAW 2.2.11).
/// Interpretation remains host-scoped; callers must explicitly select the
/// small property subset they support from tertiary options.
pub(crate) fn visit_tertiary<'a>(
    record: Record<'a>,
    budget: &mut usize,
    visitor: impl FnMut(Property<'a>) -> Result<(), String>,
) -> Result<(), String> {
    visit_kind(record, 0xf122, budget, visitor)
}

fn visit_kind<'a>(
    record: Record<'a>,
    expected_kind: u16,
    budget: &mut usize,
    mut visitor: impl FnMut(Property<'a>) -> Result<(), String>,
) -> Result<(), String> {
    if record.kind != expected_kind || record.version != 3 {
        return Err(unsupported("invalid OfficeArt property table"));
    }
    let count = usize::from(record.instance);
    *budget = budget
        .checked_sub(count)
        .ok_or_else(|| unsupported("OfficeArt property work budget exceeded"))?;
    let length = count * 6;
    let entries = record
        .payload
        .get(..length)
        .ok_or_else(|| unsupported("truncated OfficeArt properties"))?;
    let mut end = length;
    for entry in entries.chunks_exact(6) {
        let opid = u16::from_le_bytes(entry[..2].try_into().unwrap());
        let value = u32::from_le_bytes(entry[2..].try_into().unwrap());
        let complex = if opid & 0x8000 != 0 {
            let start = end;
            end = end
                .checked_add(value as usize)
                .filter(|n| *n <= record.payload.len())
                .ok_or_else(|| unsupported("truncated OfficeArt complex shape property"))?;
            Some(&record.payload[start..end])
        } else {
            None
        };
        visitor(Property {
            opid,
            value,
            complex,
        })?;
    }
    if end != record.payload.len() {
        return Err(unsupported("unexpected OfficeArt property data"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn record(payload: &[u8], count: u16) -> Record<'_> {
        Record {
            kind: 0xf00b,
            version: 3,
            instance: count,
            payload,
        }
    }
    fn entry(opid: u16, value: u32) -> Vec<u8> {
        [opid.to_le_bytes().to_vec(), value.to_le_bytes().to_vec()].concat()
    }
    #[test]
    fn preserves_bid_complex_bits_and_borrowed_data_without_reinterpreting_it() {
        let payload = [
            entry(0x4104, 7),
            entry(0xc105, 3),
            entry(0x8001, 0),
            vec![1, 2, 3],
        ]
        .concat();
        let mut values = Vec::new();
        let mut work = 3;
        visit(record(&payload, 3), &mut work, |p| {
            values.push((p.opid, p.value, p.complex));
            Ok(())
        })
        .unwrap();
        assert_eq!(
            values,
            vec![
                (0x4104, 7, None),
                (0xc105, 3, Some(&payload[18..21])),
                (0x8001, 0, Some(&payload[21..21]))
            ]
        );
        assert_eq!(work, 0);
        assert!(visit(record(&payload, 3), &mut 2, |_| Ok(())).is_err());
    }
    #[test]
    fn rejects_all_truncations_trailing_bytes_and_callback_failure() {
        let payload = [entry(0x8001, 3), vec![1, 2, 3]].concat();
        for length in 0..payload.len() {
            assert!(visit(record(&payload[..length], 1), &mut 10, |_| Ok(())).is_err());
        }
        assert!(visit(record(&[payload, vec![0]].concat(), 1), &mut 10, |_| Ok(())).is_err());
        assert!(visit(record(&entry(0x8001, u32::MAX), 1), &mut 10, |_| Ok(())).is_err());
        assert_eq!(
            visit(record(&entry(1, 2), 1), &mut 10, |_| Err("callback".into())),
            Err("callback".into())
        );
        let mut bad = record(&[], 0);
        bad.kind = 0xf122;
        assert!(visit(bad, &mut 10, |_| Ok(())).is_err());
        bad.kind = 0xf00b;
        bad.version = 15;
        assert!(visit(bad, &mut 10, |_| Ok(())).is_err());
    }

    #[test]
    fn tertiary_reuses_bounded_fopte_framing_without_weakening_primary_visit() {
        let payload = entry(0x01bf, 0x00600060);
        let tertiary = Record {
            kind: 0xf122,
            version: 3,
            instance: 1,
            payload: &payload,
        };
        let mut values = Vec::new();
        visit_tertiary(tertiary, &mut 1, |property| {
            values.push((property.opid, property.value));
            Ok(())
        })
        .unwrap();
        assert_eq!(values, [(0x01bf, 0x00600060)]);
        assert!(visit(tertiary, &mut 1, |_| Ok(())).is_err());
        assert!(visit_tertiary(tertiary, &mut 0, |_| Ok(())).is_err());

        let truncated = Record {
            payload: &payload[..5],
            ..tertiary
        };
        assert!(visit_tertiary(truncated, &mut 1, |_| Ok(())).is_err());
    }
}
