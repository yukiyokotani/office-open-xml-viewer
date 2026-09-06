//! Passive EMF extraction. MS-ODRAW 2.2.24/31; MS-EMF 2.2.9/2.3.
use super::{unsupported, Record};
use std::borrow::Cow;

// Resource policy, not a format limit. Check before allocating or inflating.
const MAX_EMF_BYTES: usize = 32 * 1024 * 1024;

pub(super) fn read<'a>(
    record: Record<'a>,
    budget: &mut usize,
    remaining: usize,
) -> Result<Option<Cow<'a, [u8]>>, String> {
    let start = match (record.version, record.instance) {
        (0, 0x3d4) => 16,
        (0, 0x3d5) => 32,
        _ => return Err(unsupported("invalid OfficeArt EMF BLIP header")),
    };
    let header = record
        .payload
        .get(start..start + 34)
        .ok_or_else(|| unsupported("truncated OfficeArt metafile header"))?;
    let expanded = number(header, 0) as usize;
    let stored = number(header, 28) as usize;
    if expanded > MAX_EMF_BYTES || expanded > remaining || stored > MAX_EMF_BYTES {
        return Err(unsupported("OfficeArt metafile byte budget exceeded"));
    }
    let bytes = &record.payload[start + 34..];
    if bytes.len() != stored || header[33] != 0xfe {
        return Err(unsupported("invalid OfficeArt metafile size or filter"));
    }
    let bytes = match header[32] {
        0xfe if stored == expanded => Cow::Borrowed(bytes),
        0 => {
            // RFC 1950 zlib-wrapped DEFLATE. One spare byte detects a lying
            // cbSize without allowing the decoder to grow its destination.
            let mut output = vec![0; expanded + 1];
            let mut decoder = flate2::Decompress::new(true);
            let status = decoder
                .decompress(bytes, &mut output, flate2::FlushDecompress::Finish)
                .map_err(|_| unsupported("invalid compressed OfficeArt metafile"))?;
            if status != flate2::Status::StreamEnd
                || decoder.total_out() != expanded as u64
                || decoder.total_in() != stored as u64
            {
                return Err(unsupported(
                    "OfficeArt metafile decompression size mismatch",
                ));
            }
            output.truncate(expanded);
            Cow::Owned(output)
        }
        _ => return Err(unsupported("unsupported OfficeArt metafile compression")),
    };
    // Do not relabel another encoding as EMF, or interpret drawing commands.
    if bytes.len() < 44 || number(&bytes, 0) != 1 || number(&bytes, 40) != 0x464d4520 {
        return Ok(None);
    }
    validate(&bytes, budget)?;
    Ok(Some(bytes))
}

fn number(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn validate(bytes: &[u8], budget: &mut usize) -> Result<(), String> {
    // MS-EMF 2.2.9/2.3: validate the envelope and bounded record walk only.
    // Ordinary OOXML image handling remains responsible for drawing support.
    if bytes.len() < 88 || number(bytes, 48) as usize != bytes.len() {
        return Err(unsupported("invalid EMF header size"));
    }
    let (mut position, mut count, mut eof) = (0, 0, false);
    while position < bytes.len() {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("EMF record work budget exceeded"))?;
        let tail = &bytes[position..];
        if tail.len() < 8 || eof {
            return Err(unsupported("invalid EMF record tail"));
        }
        let kind = number(tail, 0);
        let size = number(tail, 4) as usize;
        if size < 8
            || size % 4 != 0
            || size > tail.len()
            || (position == 0 && size < 88)
            || (position != 0 && kind == 1)
        {
            return Err(unsupported("invalid EMF record size"));
        }
        eof = kind == 14;
        if eof && size < 20 {
            return Err(unsupported("truncated EMF end record"));
        }
        position += size;
        count += 1;
    }
    if !eof || count != number(bytes, 52) {
        return Err(unsupported("EMF record count or end mismatch"));
    }
    Ok(())
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    fn emf() -> Vec<u8> {
        let mut data = vec![0; 108];
        for (offset, value) in [
            (0, 1u32),
            (4, 88),
            (40, 0x464d4520),
            (44, 0x10000),
            (48, 108),
            (52, 2),
            (56, 1),
            (88, 14),
            (92, 20),
            (104, 20),
        ] {
            data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        }
        data
    }
    fn payload(data: &[u8], two: bool) -> Vec<u8> {
        let mut result = vec![0; if two { 66 } else { 50 }];
        let start = if two { 32 } else { 16 };
        result[start..start + 4].copy_from_slice(&(data.len() as u32).to_le_bytes());
        result[start + 28..start + 32].copy_from_slice(&(data.len() as u32).to_le_bytes());
        result[start + 32] = 0xfe;
        result[start + 33] = 0xfe;
        result.extend_from_slice(data);
        result
    }
    fn parse<'a>(
        data: &'a [u8],
        two: bool,
        budget: &mut usize,
        remaining: usize,
    ) -> Result<Option<Cow<'a, [u8]>>, String> {
        read(
            Record {
                kind: 0xf01a,
                instance: if two { 0x3d5 } else { 0x3d4 },
                version: 0,
                payload: data,
            },
            budget,
            remaining,
        )
    }
    #[test]
    fn retains_uncompressed_emf_bytes_and_charges_record_work() {
        for two in [false, true] {
            let source = emf();
            let data = payload(&source, two);
            let mut budget = 2;
            let output = parse(&data, two, &mut budget, source.len())
                .unwrap()
                .unwrap();
            assert_eq!(output.as_ref(), source);
            assert!(matches!(output, Cow::Borrowed(_)));
            assert_eq!(budget, 0);
            assert!(parse(&data, two, &mut 1, source.len()).is_err());
            assert!(parse(&data, two, &mut 100, source.len() - 1).is_err());
        }
    }

    fn compressed(source: &[u8], two: bool) -> Vec<u8> {
        use std::io::Write;
        let mut compressor =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        compressor.write_all(source).unwrap();
        let compressed = compressor.finish().unwrap();
        let mut result = payload(&compressed, two);
        let start = if two { 32 } else { 16 };
        result[start..start + 4].copy_from_slice(&(source.len() as u32).to_le_bytes());
        result[start + 32] = 0;
        result
    }
    pub(crate) fn emf_test_blip() -> (Vec<u8>, Vec<u8>) {
        let source = emf();
        let payload = compressed(&source, false);
        let record = [
            0x3d40u16.to_le_bytes().to_vec(),
            0xf01au16.to_le_bytes().to_vec(),
            (payload.len() as u32).to_le_bytes().to_vec(),
            payload,
        ]
        .concat();
        (source, record)
    }
    #[test]
    fn inflates_exactly_one_zlib_stream_without_rewriting_emf() {
        for two in [false, true] {
            let source = emf();
            let data = compressed(&source, two);
            let output = parse(&data, two, &mut 2, source.len()).unwrap().unwrap();
            assert_eq!(output.as_ref(), source);
            assert!(matches!(output, Cow::Owned(_)));
            let start = if two { 32 } else { 16 };
            for declared in [0, source.len() - 1, source.len() + 1, MAX_EMF_BYTES + 1] {
                let mut bad = data.clone();
                bad[start..start + 4].copy_from_slice(&(declared as u32).to_le_bytes());
                assert!(parse(&bad, two, &mut 100, usize::MAX).is_err());
            }
            let mut bad = data.clone();
            *bad.last_mut().unwrap() ^= 1; // RFC 1950 checksum is validated.
            assert!(parse(&bad, two, &mut 100, usize::MAX).is_err());
            let mut bad = data.clone();
            bad.push(0);
            let stored = (bad.len() - start - 34) as u32;
            bad[start + 28..start + 32].copy_from_slice(&stored.to_le_bytes());
            assert!(parse(&bad, two, &mut 100, usize::MAX).is_err());
        }
    }
    #[test]
    fn rejects_truncation_reserved_encodings_and_malformed_record_envelopes() {
        let source = emf();
        let data = payload(&source, false);
        for length in 0..data.len() {
            assert!(parse(&data[..length], false, &mut 100, usize::MAX).is_err());
        }
        for (offset, value) in [(48, 1u8), (49, 0)] {
            // compression/filter
            let mut bad = data.clone();
            bad[offset] = value;
            assert!(parse(&bad, false, &mut 100, usize::MAX).is_err());
        }
        for (offset, value) in [
            (4, 84u32),
            (4, 89),
            (48, 107),
            (52, 3),
            (88, 1),
            (88, 99),
            (92, 0),
            (92, 16),
            (92, 24),
            (92, u32::MAX),
        ] {
            let mut bad = source.clone();
            bad[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            assert!(parse(&payload(&bad, false), false, &mut 100, usize::MAX).is_err());
        }
        let mut wrong = source;
        wrong[40..44].fill(0);
        assert!(parse(&payload(&wrong, false), false, &mut 100, usize::MAX)
            .unwrap()
            .is_none());
    }
}
