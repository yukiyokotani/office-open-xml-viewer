//! Borrowed OfficeArt property-table framing (MS-ODRAW 2.2.7-9).
//! Interpretation, duplicate policy and inheritance belong to the host format.
use std::ops::Range;

use super::{unsupported, ByteSpan, Record, RecordSpan};

pub(crate) struct Property<'a> {
    /// Full encoded property ID, including fBid and fComplex.
    pub opid: u16,
    pub value: u32,
    pub complex: Option<&'a [u8]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PropertySpan {
    /// Full encoded property ID, including fBid and fComplex.
    pub opid: u16,
    pub value: u32,
    pub complex: Option<ByteSpan>,
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

pub(crate) fn visit_span(
    record: &RecordSpan,
    backing: &[u8],
    budget: &mut usize,
    visitor: impl FnMut(PropertySpan) -> Result<(), String>,
) -> Result<(), String> {
    visit_span_kind(record, backing, 0xf00b, budget, visitor)
}

pub(crate) fn visit_tertiary_span(
    record: &RecordSpan,
    backing: &[u8],
    budget: &mut usize,
    visitor: impl FnMut(PropertySpan) -> Result<(), String>,
) -> Result<(), String> {
    visit_span_kind(record, backing, 0xf122, budget, visitor)
}

fn visit_kind<'a>(
    record: Record<'a>,
    expected_kind: u16,
    budget: &mut usize,
    mut visitor: impl FnMut(Property<'a>) -> Result<(), String>,
) -> Result<(), String> {
    decode_fopte(record, expected_kind, budget, |opid, value, complex| {
        visitor(Property {
            opid,
            value,
            complex: complex.map(|range| &record.payload[range]),
        })
    })
}

fn visit_span_kind(
    record: &RecordSpan,
    backing: &[u8],
    expected_kind: u16,
    budget: &mut usize,
    mut visitor: impl FnMut(PropertySpan) -> Result<(), String>,
) -> Result<(), String> {
    let viewed = record.view(backing)?;
    decode_fopte(viewed, expected_kind, budget, |opid, value, complex| {
        visitor(PropertySpan {
            opid,
            value,
            complex: complex
                .map(|range| {
                    record
                        .payload_span()
                        .checked_subrange(range, "OfficeArt complex shape property")
                })
                .transpose()?,
        })
    })
}

fn decode_fopte(
    record: Record<'_>,
    expected_kind: u16,
    budget: &mut usize,
    mut visitor: impl FnMut(u16, u32, Option<Range<usize>>) -> Result<(), String>,
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
            Some(start..end)
        } else {
            None
        };
        visitor(opid, value, complex)?;
    }
    if end != record.payload.len() {
        return Err(unsupported("unexpected OfficeArt property data"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::officeart::{record_span_with_end, record_span_with_end_in};
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
    fn encoded_record(kind: u16, count: u16, payload: &[u8]) -> Vec<u8> {
        let options = (count << 4) | 3;
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
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

    #[test]
    fn owned_visitation_matches_borrowed_payloads_and_budget_after_backing_moves() {
        let payload = [
            entry(0x4104, 7),
            entry(0xc105, 3),
            entry(0x8001, 0),
            vec![1, 2, 3],
        ]
        .concat();
        let mut bytes = vec![0xaa, 0xbb];
        bytes.extend(encoded_record(0xf00b, 3, &payload));
        let (span, _) = record_span_with_end(&bytes, 2, &mut 1, "property").unwrap();

        let mut borrowed = Vec::new();
        let mut borrowed_budget = 3;
        visit(
            span.view(&bytes).unwrap(),
            &mut borrowed_budget,
            |property| {
                borrowed.push((
                    property.opid,
                    property.value,
                    property.complex.map(<[u8]>::to_vec),
                ));
                Ok(())
            },
        )
        .unwrap();

        let moved = bytes;
        let mut owned = Vec::new();
        let mut span_budget = 3;
        visit_span(&span, &moved, &mut span_budget, |property| {
            owned.push((
                property.opid,
                property.value,
                property
                    .complex
                    .map(|complex| complex.view(&moved).map(<[u8]>::to_vec))
                    .transpose()?,
            ));
            Ok(())
        })
        .unwrap();
        assert_eq!(owned, borrowed);
        assert_eq!(span_budget, borrowed_budget);
        assert_eq!(owned[0].2, None);
        assert_eq!(owned[2].2, Some(Vec::new()));
    }

    #[test]
    fn owned_and_borrowed_visitors_match_all_truncation_and_callback_errors() {
        let payload = [entry(0x8001, 3), vec![1, 2, 3]].concat();
        for length in 0..payload.len() {
            let truncated = &payload[..length];
            let bytes = encoded_record(0xf00b, 1, truncated);
            let (span, _) = record_span_with_end(&bytes, 0, &mut 1, "property").unwrap();
            let mut borrowed_budget = 10;
            let borrowed = visit(span.view(&bytes).unwrap(), &mut borrowed_budget, |_| Ok(()));
            let mut span_budget = 10;
            let owned = visit_span(&span, &bytes, &mut span_budget, |_| Ok(()));
            assert_eq!(owned, borrowed, "payload length {length}");
            assert_eq!(span_budget, borrowed_budget, "payload length {length}");
        }

        let bytes = encoded_record(0xf00b, 1, &entry(1, 2));
        let (span, _) = record_span_with_end(&bytes, 0, &mut 1, "property").unwrap();
        assert_eq!(
            visit_span(&span, &bytes, &mut 10, |_| Err("callback".into())),
            Err("callback".into())
        );

        let mut borrowed_budget = 0;
        let borrowed = visit(span.view(&bytes).unwrap(), &mut borrowed_budget, |_| Ok(()));
        let mut span_budget = 0;
        let owned = visit_span(&span, &bytes, &mut span_budget, |_| Ok(()));
        assert_eq!(owned, borrowed);
        assert_eq!(span_budget, borrowed_budget);

        let trailing_bytes = encoded_record(0xf00b, 1, &[entry(1, 2), vec![9]].concat());
        let (trailing, _) = record_span_with_end(&trailing_bytes, 0, &mut 1, "property").unwrap();
        let mut callback_count = 0;
        assert!(visit_span(&trailing, &trailing_bytes, &mut 1, |_| {
            callback_count += 1;
            Ok(())
        })
        .is_err());
        assert_eq!(callback_count, 1);
    }

    #[test]
    fn owned_visitation_honors_record_kind_backing_and_parent_bounds() {
        let payload = entry(0x01bf, 0x00600060);
        let primary_bytes = encoded_record(0xf00b, 1, &payload);
        let (primary, _) = record_span_with_end(&primary_bytes, 0, &mut 1, "property").unwrap();
        assert!(visit_tertiary_span(&primary, &primary_bytes, &mut 1, |_| Ok(())).is_err());

        let tertiary_bytes = encoded_record(0xf122, 1, &payload);
        let (tertiary, _) = record_span_with_end(&tertiary_bytes, 0, &mut 1, "property").unwrap();
        assert!(visit_span(&tertiary, &tertiary_bytes, &mut 1, |_| Ok(())).is_err());
        visit_tertiary_span(&tertiary, &tertiary_bytes, &mut 1, |_| Ok(())).unwrap();
        assert!(visit_tertiary_span(
            &tertiary,
            &tertiary_bytes[..tertiary_bytes.len() - 1],
            &mut 1,
            |_| Ok(())
        )
        .is_err());

        let mut adjacent = primary_bytes.clone();
        adjacent.extend([9, 9, 9]);
        let short_parent =
            ByteSpan::new(0..primary_bytes.len() - 1, adjacent.len(), "parent").unwrap();
        assert!(record_span_with_end_in(&adjacent, &short_parent, 0, &mut 1, "property").is_err());
    }
}
