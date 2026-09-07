//! Passive EMF/WMF extraction. MS-ODRAW 2.2.24-25/31.
use super::{unsupported, Record};
use std::borrow::Cow;

// Resource policy, not a format limit. Check before allocating or inflating.
const MAX_METAFILE_BYTES: usize = 32 * 1024 * 1024;

pub(super) fn read<'a>(
    record: Record<'a>,
    budget: &mut usize,
    remaining: usize,
) -> Result<Option<Cow<'a, [u8]>>, String> {
    let (start, format) = match (record.kind, record.version, record.instance) {
        (0xf01a, 0, 0x3d4) => (16, Format::Emf),
        (0xf01a, 0, 0x3d5) => (32, Format::Emf),
        (0xf01b, 0, 0x216) => (16, Format::Wmf),
        (0xf01b, 0, 0x217) => (32, Format::Wmf),
        _ => return Err(unsupported("invalid OfficeArt metafile BLIP header")),
    };
    let header = record
        .payload
        .get(start..start + 34)
        .ok_or_else(|| unsupported("truncated OfficeArt metafile header"))?;
    let expanded = number(header, 0) as usize;
    let stored = number(header, 28) as usize;
    if expanded > MAX_METAFILE_BYTES || expanded > remaining || stored > MAX_METAFILE_BYTES {
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
    // Preserve only the record's declared metafile encoding; never interpret it.
    match format {
        Format::Emf => {
            if bytes.len() < 44 || number(&bytes, 0) != 1 || number(&bytes, 40) != 0x464d4520 {
                return Ok(None);
            }
            validate_emf(&bytes, budget)?;
        }
        Format::Wmf => {
            if !validate_wmf(&bytes, budget)? {
                return Ok(None);
            }
        }
    }
    Ok(Some(bytes))
}

#[derive(Clone, Copy)]
enum Format {
    Emf,
    Wmf,
}

fn number(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn validate_emf(bytes: &[u8], budget: &mut usize) -> Result<(), String> {
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

fn validate_wmf(bytes: &[u8], budget: &mut usize) -> Result<bool, String> {
    // MS-WMF 2.3.2.2-3/2.3: a placeable header is optional; the META_HEADER
    // owns the file size, followed by bounded records and a terminal META_EOF.
    let start = if bytes.starts_with(&0x9ac6cdd7u32.to_le_bytes()) {
        let placeable = bytes
            .get(..22)
            .ok_or_else(|| unsupported("truncated placeable WMF header"))?;
        let checksum = (0..10).fold(0u16, |sum, i| {
            sum ^ u16::from_le_bytes(placeable[i * 2..i * 2 + 2].try_into().unwrap())
        });
        if placeable[16..20] != [0, 0, 0, 0]
            || checksum != u16::from_le_bytes(placeable[20..22].try_into().unwrap())
        {
            return Err(unsupported("invalid placeable WMF header"));
        }
        22
    } else {
        0
    };
    let header = bytes
        .get(start..start + 18)
        .ok_or_else(|| unsupported("truncated WMF header"))?;
    let word = |o| u16::from_le_bytes(header[o..o + 2].try_into().unwrap());
    let dword = |o| u32::from_le_bytes(header[o..o + 4].try_into().unwrap());
    if !matches!(word(0), 1 | 2)
        || word(2) != 9
        || !matches!(word(4), 0x100 | 0x300)
        || usize::try_from(dword(6))
            .ok()
            .and_then(|n| n.checked_mul(2))
            != Some(bytes.len() - start)
        || dword(12) < 3
    {
        return Err(unsupported("invalid WMF header"));
    }
    if start != 0 && word(0) == 2 && bytes[4..6] != [0, 0] {
        return Err(unsupported("invalid disk WMF handle"));
    }
    let mut position = start + 18;
    while position < bytes.len() {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("WMF record work budget exceeded"))?;
        let tail = &bytes[position..];
        if tail.len() < 6 {
            return Err(unsupported("invalid WMF record tail"));
        }
        let words = u32::from_le_bytes(tail[..4].try_into().unwrap()) as usize;
        let size = words
            .checked_mul(2)
            .ok_or_else(|| unsupported("invalid WMF record size"))?;
        if words < 3 || size > tail.len() {
            return Err(unsupported("invalid WMF record size"));
        }
        let eof = u16::from_le_bytes(tail[4..6].try_into().unwrap()) == 0;
        if eof && words != 3 {
            return Err(unsupported("invalid WMF end record"));
        }
        position += size;
        if eof {
            // MS-WMF 2.3.2.1 defines EOF as the end of the record stream.
            // A declared payload with trailing bytes is outside our supported
            // subset: omit the entire image, never strip, interpret or forward
            // the trailer. This keeps an unsupported image from rejecting an
            // otherwise supported document without claiming a padding rule.
            return Ok(position == bytes.len());
        }
    }
    Err(unsupported("missing WMF end record"))
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
    fn wmf() -> Vec<u8> {
        let mut data = vec![0; 24];
        data[0..2].copy_from_slice(&1u16.to_le_bytes());
        data[2..4].copy_from_slice(&9u16.to_le_bytes());
        data[4..6].copy_from_slice(&0x300u16.to_le_bytes());
        data[6..10].copy_from_slice(&12u32.to_le_bytes());
        data[12..16].copy_from_slice(&3u32.to_le_bytes());
        data[18..22].copy_from_slice(&3u32.to_le_bytes());
        data
    }
    fn placeable_wmf() -> Vec<u8> {
        let mut prefix = vec![0; 22];
        prefix[..4].copy_from_slice(&0x9ac6cdd7u32.to_le_bytes());
        prefix[14..16].copy_from_slice(&1440u16.to_le_bytes());
        let checksum = (0..10).fold(0u16, |sum, i| {
            sum ^ u16::from_le_bytes(prefix[i * 2..i * 2 + 2].try_into().unwrap())
        });
        prefix[20..22].copy_from_slice(&checksum.to_le_bytes());
        prefix.extend(wmf());
        prefix
    }
    fn parse_wmf<'a>(
        data: &'a [u8],
        two: bool,
        budget: &mut usize,
        remaining: usize,
    ) -> Result<Option<Cow<'a, [u8]>>, String> {
        read(
            Record {
                kind: 0xf01b,
                instance: if two { 0x217 } else { 0x216 },
                version: 0,
                payload: data,
            },
            budget,
            remaining,
        )
    }
    #[test]
    fn retains_raw_and_compressed_wmf_with_one_or_two_uids() {
        for two in [false, true] {
            let source = wmf();
            let raw = payload(&source, two);
            assert_eq!(
                parse_wmf(&raw, two, &mut 1, source.len())
                    .unwrap()
                    .unwrap()
                    .as_ref(),
                source
            );
            let zipped = compressed(&source, two);
            assert_eq!(
                parse_wmf(&zipped, two, &mut 1, source.len())
                    .unwrap()
                    .unwrap()
                    .as_ref(),
                source
            );
            assert!(parse_wmf(&raw, two, &mut 0, source.len()).is_err());
            assert!(parse_wmf(&raw, two, &mut 10, source.len() - 1).is_err());
        }
    }

    #[test]
    fn rejects_malformed_wmf_framing_and_end_records() {
        let source = wmf();
        for (offset, value) in [(2, 8u16), (4, 0x200), (18, 2)] {
            let mut bad = source.clone();
            bad[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
            assert!(parse_wmf(&payload(&bad, false), false, &mut 10, usize::MAX).is_err());
        }
        let mut bad = source.clone();
        bad[6..10].copy_from_slice(&11u32.to_le_bytes());
        assert!(parse_wmf(&payload(&bad, false), false, &mut 10, usize::MAX).is_err());
        let mut bad = source;
        bad[22..24].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_wmf(&payload(&bad, false), false, &mut 10, usize::MAX).is_err());
        let placeable = placeable_wmf();
        assert_eq!(
            parse_wmf(&payload(&placeable, false), false, &mut 10, usize::MAX)
                .unwrap()
                .unwrap()
                .as_ref(),
            placeable
        );
        for offset in [16, 20] {
            let mut bad = placeable.clone();
            bad[offset] ^= 1;
            assert!(parse_wmf(&payload(&bad, false), false, &mut 10, usize::MAX).is_err());
        }
        let mut memory_handle = placeable.clone();
        memory_handle[4..6].copy_from_slice(&7u16.to_le_bytes());
        let checksum = (0..10).fold(0u16, |sum, i| {
            sum ^ u16::from_le_bytes(memory_handle[i * 2..i * 2 + 2].try_into().unwrap())
        });
        memory_handle[20..22].copy_from_slice(&checksum.to_le_bytes());
        assert!(parse_wmf(&payload(&memory_handle, false), false, &mut 10, usize::MAX).is_ok());
        let mut disk_handle = memory_handle;
        disk_handle[22..24].copy_from_slice(&2u16.to_le_bytes());
        assert!(parse_wmf(&payload(&disk_handle, false), false, &mut 10, usize::MAX).is_err());
        let mut members = wmf();
        members[16..18].copy_from_slice(&9u16.to_le_bytes()); // SHOULD be zero, not MUST.
        assert!(parse_wmf(&payload(&members, false), false, &mut 10, usize::MAX).is_ok());
    }
    #[test]
    fn wmf_rejects_bad_officeart_headers_truncation_and_zlib_framing() {
        let source = wmf();
        for length in 0..18 {
            assert!(parse_wmf(
                &payload(&source[..length], false),
                false,
                &mut 10,
                usize::MAX
            )
            .is_err());
        }
        for (version, instance) in [(1, 0x216), (0, 0x215), (0, 0x218)] {
            let data = payload(&source, false);
            assert!(read(
                Record {
                    kind: 0xf01b,
                    instance,
                    version,
                    payload: &data
                },
                &mut 10,
                usize::MAX
            )
            .is_err());
        }
        let mut zipped = compressed(&source, false);
        zipped[16..20].copy_from_slice(&((source.len() + 1) as u32).to_le_bytes());
        assert!(parse_wmf(&zipped, false, &mut 10, usize::MAX).is_err());
        let mut zipped = compressed(&source, false);
        *zipped.last_mut().unwrap() ^= 1;
        assert!(parse_wmf(&zipped, false, &mut 10, usize::MAX).is_err());
        let mut zipped = compressed(&source, false);
        zipped.push(0);
        let stored = (zipped.len() - 50) as u32;
        zipped[44..48].copy_from_slice(&stored.to_le_bytes());
        assert!(parse_wmf(&zipped, false, &mut 10, usize::MAX).is_err());
    }

    #[test]
    fn wmf_preserves_unknown_records_but_requires_one_terminal_eof() {
        let mut source = wmf();
        source.splice(18..18, [4, 0, 0, 0, 0x34, 0x12, 0xaa, 0xbb]);
        source[6..10].copy_from_slice(&16u32.to_le_bytes());
        source[12..16].copy_from_slice(&4u32.to_le_bytes());
        assert_eq!(
            parse_wmf(&payload(&source, false), false, &mut 2, usize::MAX)
                .unwrap()
                .unwrap()
                .as_ref(),
            source
        );
        let mut after_eof = wmf();
        after_eof.extend_from_slice(&[3, 0, 0, 0, 1, 0]);
        after_eof[6..10].copy_from_slice(&15u32.to_le_bytes());
        assert!(parse_wmf(&payload(&after_eof, false), false, &mut 10, usize::MAX)
            .unwrap()
            .is_none());
        let mut overflow = wmf();
        overflow[18..22].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_wmf(&payload(&overflow, false), false, &mut 10, usize::MAX).is_err());
        let mut no_eof = wmf();
        no_eof[22..24].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_wmf(&payload(&no_eof, false), false, &mut 10, usize::MAX).is_err());
        let mut v100 = wmf();
        v100[4..6].copy_from_slice(&0x100u16.to_le_bytes());
        assert!(parse_wmf(&payload(&v100, false), false, &mut 10, usize::MAX).is_ok());
        let mut disk = placeable_wmf();
        disk[22..24].copy_from_slice(&2u16.to_le_bytes());
        assert!(parse_wmf(&payload(&disk, false), false, &mut 10, usize::MAX).is_ok());
    }
    #[test]
    fn wmf_omits_the_whole_image_for_any_declared_post_eof_payload() {
        for mut source in [wmf(), placeable_wmf()] {
            let start = if source.len() == 24 { 0 } else { 22 };
            let original_length = source.len();
            for length in [2, 4, 6, 8, 64] {
                for value in [0, 0xa5] {
                    source.truncate(original_length);
                    source.resize(original_length + length, value);
                    let words = ((source.len() - start) / 2) as u32;
                    source[start + 6..start + 10]
                        .copy_from_slice(&words.to_le_bytes());
                    for two in [false, true] {
                        for bytes in [payload(&source, two), compressed(&source, two)] {
                            assert!(parse_wmf(&bytes, two, &mut 1, source.len())
                                .unwrap()
                                .is_none());
                        }
                    }
                }
            }
        }
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
    pub(crate) fn wmf_test_blip() -> (Vec<u8>, Vec<u8>) {
        let source = wmf();
        let payload = compressed(&source, false);
        let record = [
            0x2160u16.to_le_bytes().to_vec(),
            0xf01bu16.to_le_bytes().to_vec(),
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
            for declared in [
                0,
                source.len() - 1,
                source.len() + 1,
                MAX_METAFILE_BYTES + 1,
            ] {
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
