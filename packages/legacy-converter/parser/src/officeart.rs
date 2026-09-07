//! Bounded OfficeArt / PowerPoint record headers (MS-ODRAW 2.2.1).
use std::ops::Range;

pub(crate) mod geometry;
mod metafile;
pub(crate) mod properties;
#[cfg(test)]
pub(crate) use metafile::tests::{emf_test_blip, wmf_test_blip};
pub(crate) mod raster;
pub(crate) mod stroke;

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Record<'a> {
    pub version: u8,
    pub instance: u16,
    pub kind: u16,
    pub payload: &'a [u8],
}

/// An owned byte range into a parser-selected backing buffer. Range checks do
/// not establish source identity: callers must view it against the same owned
/// backing from which it was parsed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ByteSpan {
    range: Range<usize>,
}

impl ByteSpan {
    pub(crate) fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    pub(crate) fn new(
        range: Range<usize>,
        backing_len: usize,
        context: &str,
    ) -> Result<Self, String> {
        if range.start > range.end || range.end > backing_len {
            return Err(unsupported(format!("invalid {context} byte span")));
        }
        Ok(Self { range })
    }

    pub(crate) fn view<'a>(&self, backing: &'a [u8]) -> Result<&'a [u8], String> {
        backing
            .get(self.range.clone())
            .ok_or_else(|| unsupported("byte span is outside backing"))
    }
}

/// An owned record header plus an absolute payload range. The payload remains
/// in the caller-owned backing and is borrowed only for the duration of `view`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RecordSpan {
    version: u8,
    instance: u16,
    kind: u16,
    payload: ByteSpan,
}

impl RecordSpan {
    pub(crate) fn payload_span(&self) -> &ByteSpan {
        &self.payload
    }

    pub(crate) fn view<'a>(&self, backing: &'a [u8]) -> Result<Record<'a>, String> {
        Ok(Record {
            version: self.version,
            instance: self.instance,
            kind: self.kind,
            payload: self.payload_span().view(backing)?,
        })
    }
}

pub(crate) fn record_span_with_end(
    bytes: &[u8],
    offset: usize,
    budget: &mut usize,
    context: &str,
) -> Result<(RecordSpan, usize), String> {
    let backing = ByteSpan::new(0..bytes.len(), bytes.len(), context)?;
    record_span_with_end_in(bytes, &backing, offset, budget, context)
}

/// Parse a record at an absolute offset while constraining its complete header
/// and payload to `bounds`. This is the span equivalent of parsing a borrowed
/// parent payload: bytes belonging to an adjacent record are never visible.
pub(crate) fn record_span_with_end_in(
    bytes: &[u8],
    parent: &ByteSpan,
    offset: usize,
    budget: &mut usize,
    context: &str,
) -> Result<(RecordSpan, usize), String> {
    if *budget == 0 {
        return Err(unsupported(format!("too many {context} records")));
    }
    *budget -= 1;
    let bounds = parent.range();
    if bounds.end > bytes.len() || offset < bounds.start {
        return Err(unsupported(format!("invalid {context} record bounds")));
    }
    let remaining = bytes
        .get(offset..bounds.end)
        .filter(|tail| tail.len() >= 8)
        .ok_or_else(|| unsupported(format!("truncated {context} record header")))?;
    let options = u16::from_le_bytes(remaining[0..2].try_into().unwrap());
    let kind = u16::from_le_bytes(remaining[2..4].try_into().unwrap());
    let size = usize::try_from(u32::from_le_bytes(remaining[4..8].try_into().unwrap()))
        .map_err(|_| unsupported(format!("{context} record is too large")))?;
    let payload_start = offset
        .checked_add(8)
        .ok_or_else(|| unsupported(format!("{context} record range overflow")))?;
    let end = payload_start
        .checked_add(size)
        .ok_or_else(|| unsupported(format!("{context} record range overflow")))?;
    if end > bounds.end {
        return Err(unsupported(format!(
            "truncated {context} record at offset {offset}: declared {size} bytes with {} available",
            bounds.end.saturating_sub(payload_start),
        )));
    }
    Ok((
        RecordSpan {
            version: (options & 0x000f) as u8,
            instance: options >> 4,
            kind,
            payload: ByteSpan::new(payload_start..end, bytes.len(), context)?,
        },
        end,
    ))
}

pub(crate) fn record_with_end<'a>(
    bytes: &'a [u8],
    offset: usize,
    budget: &mut usize,
    context: &str,
) -> Result<(Record<'a>, usize), String> {
    let (span, end) = record_span_with_end(bytes, offset, budget, context)?;
    Ok((span.view(bytes)?, end))
}

#[cfg(test)]
mod span_tests {
    use super::*;

    fn record(options: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
    }

    #[test]
    fn owned_span_views_absolute_payload_after_backing_moves() {
        let mut bytes = vec![0xaa, 0xbb, 0xcc];
        bytes.extend(record(0x123f, 0xf00a, &[1, 2, 3, 4]));
        let mut budget = 1;
        let (span, end) = record_span_with_end(&bytes, 3, &mut budget, "test").unwrap();
        assert_eq!(end, bytes.len());
        assert_eq!(budget, 0);

        let moved = bytes;
        let view = span.view(&moved).unwrap();
        assert_eq!(view.version, 0x0f);
        assert_eq!(view.instance, 0x123);
        assert_eq!(view.kind, 0xf00a);
        assert_eq!(view.payload, &[1, 2, 3, 4]);
    }

    #[test]
    fn byte_span_rechecks_backing_bounds() {
        let span = ByteSpan::new(2..5, 5, "test").unwrap();
        assert_eq!(span.view(&[0, 1, 2, 3, 4]).unwrap(), &[2, 3, 4]);
        assert!(span
            .view(&[0, 1, 2, 3])
            .unwrap_err()
            .contains("outside backing"));
        assert!(ByteSpan::new(5..2, 5, "test").is_err());
    }

    #[test]
    fn span_parser_preserves_budget_and_record_errors() {
        let bytes = record(0, 1, &[9]);
        let mut exhausted = 0;
        assert!(record_span_with_end(&bytes, 0, &mut exhausted, "test")
            .unwrap_err()
            .contains("too many test records"));

        let mut budget = 1;
        assert!(record_span_with_end(&bytes[..8], 0, &mut budget, "test")
            .unwrap_err()
            .contains("truncated test record at offset 0"));
        assert_eq!(budget, 0);

        let mut budget = 1;
        assert!(
            record_span_with_end(&bytes, usize::MAX, &mut budget, "test")
                .unwrap_err()
                .contains("truncated test record header")
        );
        assert_eq!(budget, 0);
    }

    #[test]
    fn nested_record_uses_absolute_offsets_and_parent_payload_bounds() {
        let mut bytes = vec![0; 3];
        bytes.extend(record(15, 1000, &record(0, 7, &[1, 2])));
        let (parent, _) = record_span_with_end(&bytes, 3, &mut 1, "parent").unwrap();
        let bounds = parent.payload_span();
        let (child, end) =
            record_span_with_end_in(&bytes, bounds, bounds.range().start, &mut 1, "child").unwrap();
        assert_eq!(end, bounds.range().end);
        assert_eq!(child.view(&bytes).unwrap().payload, &[1, 2]);
        for offset in [0, bounds.range().start - 1, bounds.range().end, usize::MAX] {
            assert!(record_span_with_end_in(&bytes, bounds, offset, &mut 1, "child").is_err());
        }
        // Header itself must fit even when subsequent bytes exist in backing.
        let short = ByteSpan::new(
            bounds.range().start..bounds.range().start + 7,
            bytes.len(),
            "child",
        )
        .unwrap();
        assert!(
            record_span_with_end_in(&bytes, &short, bounds.range().start, &mut 1, "child")
                .unwrap_err()
                .contains("truncated child record header")
        );
    }

    #[test]
    fn nested_span_cannot_consume_an_adjacent_sibling() {
        let child = record(0, 7, &[1, 2, 3, 4]);
        let mut bytes = child.clone();
        bytes.extend([9, 9, 9, 9]);
        let mut budget = 1;
        let error = record_span_with_end_in(
            &bytes,
            &ByteSpan::new(0..child.len() - 2, bytes.len(), "nested").unwrap(),
            0,
            &mut budget,
            "nested",
        )
        .unwrap_err();
        assert!(error.contains("declared 4 bytes with 2 available"));
        assert_eq!(budget, 0);
    }

    #[test]
    fn owned_and_borrowed_record_parsers_have_matching_results() {
        let bytes = record(0x432a, 0x1234, &[5, 6, 7]);
        let mut borrowed_budget = 1;
        let (borrowed, borrowed_end) =
            record_with_end(&bytes, 0, &mut borrowed_budget, "parity").unwrap();
        let mut span_budget = 1;
        let (span, span_end) = record_span_with_end(&bytes, 0, &mut span_budget, "parity").unwrap();
        let owned = span.view(&bytes).unwrap();
        assert_eq!(
            (
                owned.version,
                owned.instance,
                owned.kind,
                owned.payload,
                span_end
            ),
            (
                borrowed.version,
                borrowed.instance,
                borrowed.kind,
                borrowed.payload,
                borrowed_end,
            )
        );
        assert_eq!(span_budget, borrowed_budget);

        let mut borrowed_budget = 1;
        let borrowed_error =
            record_with_end(&bytes[..8], 0, &mut borrowed_budget, "parity").unwrap_err();
        let mut span_budget = 1;
        let span_error =
            record_span_with_end(&bytes[..8], 0, &mut span_budget, "parity").unwrap_err();
        assert_eq!(span_error, borrowed_error);
        assert_eq!(span_budget, borrowed_budget);
    }
}
