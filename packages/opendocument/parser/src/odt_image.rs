//! Image resources (design §7): an image class is admitted only when the
//! reader decodes it completely. PNG, GIF and JPEG, recognised by their
//! signatures, are decoded here with bounded memory; anything else, and any
//! file a decoder rejects, rejects the document.
//!
//! Decoding alone does not make a browser and LibreOffice paint a file alike
//! (measured by the reviewer: unknown chunks and markers, damaged ancillary
//! chunks, gamma and colour profiles, EXIF orientation and partial GIF
//! frames are each treated differently). Each format's structure is
//! therefore checked first against the shape the corpus's pictures have:
//! the parts listed below and nothing after the file's end.

use crate::error::unsupported;
use std::io::Cursor;

/// Decoded-size bounds: 32768 px per side and 40 M px (the bounds the
/// legacy readers apply to their raster images).
const MAX_SIDE: u32 = 32_768;
const MAX_PIXELS: u64 = 40_000_000;
/// Decoder working memory: the largest admitted image at 8 bytes per pixel
/// (16-bit RGBA PNG), plus room for the decoders' own state.
const MAX_DECODE_BYTES: u64 = MAX_PIXELS * 8 + (1 << 20);
/// Implementation bound on the pixels a document's pictures decode, every
/// frame counted, charged before each frame is decoded.
pub(super) const MAX_DOCUMENT_PIXELS: u64 = 64_000_000;

/// The media type of `bytes` once every frame of it has decoded, charging
/// the decoded pixels against `budget` (the document's remaining pixels).
pub(super) fn decode(bytes: &[u8], budget: &mut u64) -> Result<&'static str, String> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        png(bytes, budget)?;
        Ok("image/png")
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        gif(bytes, budget)?;
        Ok("image/gif")
    } else if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        jpeg(bytes, budget)?;
        Ok("image/jpeg")
    } else {
        Err(unsupported("image-format"))
    }
}

fn check_size(width: u64, height: u64) -> Result<(), String> {
    if width == 0
        || height == 0
        || width > u64::from(MAX_SIDE)
        || height > u64::from(MAX_SIDE)
        || width * height > MAX_PIXELS
    {
        return Err(unsupported("image-size"));
    }
    Ok(())
}

fn charge(budget: &mut u64, width: u64, height: u64) -> Result<(), String> {
    *budget = budget
        .checked_sub(width * height)
        .ok_or_else(|| "OpenDocument pictures exceed the decoding limit".to_string())?;
    Ok(())
}

/// PNG chunks (ISO/IEC 15948 §5): IHDR first, then the palette and
/// transparency, physical size, text and time (which no decoder paints),
/// consecutive image data and IEND at the file's end, every CRC correct.
/// Colour-space chunks (gAMA, cHRM, sRGB, iCCP), EXIF, animation and any
/// other chunk reject. A palette belongs to indexed images only and
/// transparency samples must be in range (§11.2.3, §11.3.2.1): decoders
/// treat others differently (measured by the reviewer: an out-of-range grey
/// sample is transparent in a browser and opaque in LibreOffice).
fn png_structure(bytes: &[u8]) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let structure = || unsupported("image-structure");
    let mut at = 8;
    let mut seen_data = false;
    let mut data_ended = false;
    // Colour type and bit depth from IHDR, and the palette's entries.
    let mut header: Option<(u8, u8)> = None;
    let mut palette = 0usize;
    while at < bytes.len() {
        let length = u32::from_be_bytes(
            bytes
                .get(at..at + 4)
                .ok_or_else(decode)?
                .try_into()
                .expect("4 bytes"),
        ) as usize;
        let end = (at + 12).checked_add(length).ok_or_else(decode)?;
        let typed = bytes.get(at + 4..end - 4).ok_or_else(decode)?;
        let crc = bytes.get(end - 4..end).ok_or_else(decode)?;
        if crc32fast::hash(typed).to_be_bytes() != crc {
            return Err(decode());
        }
        let (kind, data) = typed.split_at(4);
        let in_range = |sample: &[u8], depth: u8| {
            u32::from(u16::from_be_bytes([sample[0], sample[1]])) < 1u32 << depth
        };
        match (kind, header) {
            (b"IHDR", None) => {
                if data.len() != 13 {
                    return Err(decode());
                }
                header = Some((data[9], data[8]));
            }
            (b"PLTE", Some((3, depth))) if !seen_data && palette == 0 => {
                palette = data.len() / 3;
                if data.len() % 3 != 0 || palette == 0 || palette > 1 << depth {
                    return Err(structure());
                }
            }
            (b"tRNS", Some((colour, depth))) if !seen_data => {
                let valid = match colour {
                    0 => data.len() == 2 && in_range(data, depth),
                    2 => data.len() == 6 && data.chunks(2).all(|sample| in_range(sample, depth)),
                    3 => palette > 0 && data.len() <= palette,
                    _ => false,
                };
                if !valid {
                    return Err(structure());
                }
            }
            (b"pHYs", Some(_)) if !seen_data => {}
            (b"tEXt" | b"zTXt" | b"iTXt" | b"tIME", Some(_)) => {}
            (b"IDAT", Some((colour, _))) if !data_ended => {
                if colour == 3 && palette == 0 {
                    return Err(structure());
                }
                seen_data = true;
            }
            (b"IEND", Some(_)) if seen_data => {
                return if end == bytes.len() {
                    Ok(())
                } else {
                    Err(decode())
                };
            }
            _ => return Err(structure()),
        }
        if seen_data && kind != b"IDAT" {
            data_ended = true;
        }
        at = end;
    }
    Err(decode())
}

/// GIF blocks (GIF89a §15-§27): one image covering the logical screen, with
/// graphic control, comment and application extensions, and the trailer at
/// the file's end. Plain text extensions, partial or several frames reject
/// (a browser and LibreOffice place a partial frame differently, measured
/// by the reviewer; several frames animate).
fn gif_structure(bytes: &[u8]) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let byte = |at: usize| bytes.get(at).copied().ok_or_else(decode);
    let word = |at: usize| -> Result<u16, String> {
        Ok(u16::from(byte(at)?) | u16::from(byte(at + 1)?) << 8)
    };
    let table = |flags: u8| {
        if flags & 0x80 != 0 {
            3usize << ((flags & 7) + 1)
        } else {
            0
        }
    };
    let sub_blocks = |mut at: usize| -> Result<usize, String> {
        loop {
            let size = usize::from(byte(at)?);
            at += 1 + size;
            if size == 0 {
                return Ok(at);
            }
        }
    };
    let (width, height) = (word(6)?, word(8)?);
    let mut at = 13 + table(byte(10)?);
    let mut images = 0;
    loop {
        match byte(at)? {
            0x21 => match byte(at + 1)? {
                0xf9 | 0xfe | 0xff => at = sub_blocks(at + 2)?,
                _ => return Err(unsupported("image-structure")),
            },
            0x2c => {
                images += 1;
                if images > 1 {
                    return Err(unsupported("image-animation"));
                }
                if (word(at + 1)?, word(at + 3)?, word(at + 5)?, word(at + 7)?)
                    != (0, 0, width, height)
                {
                    return Err(unsupported("image-structure"));
                }
                at = sub_blocks(at + 10 + table(byte(at + 9)?) + 1)?;
            }
            0x3b if images == 1 && at + 1 == bytes.len() => return Ok(()),
            _ => return Err(decode()),
        }
    }
}

/// JPEG markers (ITU-T T.81 Annex B): SOI, then application segments,
/// comments, quantisation and Huffman tables, restart intervals, one
/// baseline 8-bit frame of one or three components, its scans, and EOI at
/// the file's end. A colour profile (APP2 ICC_PROFILE), an EXIF orientation
/// other than 1, an Adobe colour transform other than YCbCr, other frame
/// types and any other marker reject.
fn jpeg_structure(bytes: &[u8]) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    let structure = || unsupported("image-structure");
    let mut at = 2;
    let mut frame: Option<&[u8]> = None;
    let mut scanned = false;
    let mut restart = 0u16;
    let mut tables: HuffmanTables<'_> = [[None; 2]; 2];
    loop {
        if bytes.get(at) != Some(&0xff) {
            return Err(decode());
        }
        let marker = *bytes.get(at + 1).ok_or_else(decode)?;
        // Standalone markers outside a scan (a second SOI, restarts, TEM).
        if matches!(marker, 0x01 | 0xd0..=0xd8) {
            return Err(structure());
        }
        if marker == 0xd9 {
            return if scanned && at + 2 == bytes.len() {
                Ok(())
            } else {
                Err(decode())
            };
        }
        let length = usize::from(u16::from_be_bytes(
            bytes
                .get(at + 2..at + 4)
                .ok_or_else(decode)?
                .try_into()
                .expect("2 bytes"),
        ));
        if length < 2 {
            return Err(decode());
        }
        let segment = bytes.get(at + 4..at + 2 + length).ok_or_else(decode)?;
        match marker {
            0xe2 if segment.starts_with(b"ICC_PROFILE\0") => {
                return Err(unsupported("image-colour"))
            }
            0xe1 if segment.starts_with(b"Exif\0\0") => {
                if exif_orientation(&segment[6..]).is_some_and(|o| o != 1) {
                    return Err(unsupported("image-orientation"));
                }
            }
            // The Adobe segment's transform flag (DCT Encode filter, Adobe
            // TN 5116): 1 is YCbCr, the default for three components.
            0xee if segment.starts_with(b"Adobe") => {
                if segment.get(11) != Some(&1) {
                    return Err(unsupported("image-colour"));
                }
            }
            0xe0..=0xef | 0xfe => {}
            0xdd if !scanned && segment.len() == 2 => {
                restart = u16::from_be_bytes([segment[0], segment[1]]);
            }
            0xdb if !scanned && quantisation_tables_valid(segment) => {}
            0xc4 if !scanned && huffman_tables_valid(segment) => {
                let mut rest = segment;
                while let Some((&header, tail)) = rest.split_first() {
                    let total: usize = tail[..16].iter().map(|&c| usize::from(c)).sum();
                    tables[usize::from(header >> 4)][usize::from(header & 15)] =
                        Some((&tail[..16], &tail[16..16 + total]));
                    rest = &tail[16 + total..];
                }
            }
            0xc0 if frame.is_none() && frame_header_valid(segment) => frame = Some(segment),
            // One scan holding every component: decoders fill a missing
            // scan or a cut one differently (measured by the reviewer).
            0xda if !scanned && frame.is_some() && scan_header_valid(segment) => {
                scanned = true;
                at = entropy_coded_data(
                    bytes,
                    at + 2 + length,
                    frame.expect("a frame"),
                    segment,
                    &tables,
                    restart,
                )?;
                continue;
            }
            _ => return Err(structure()),
        }
        at += 2 + length;
    }
}

/// Huffman code tables by class (DC, AC) and id: (counts, symbols).
type HuffmanTables<'b> = [[Option<(&'b [u8], &'b [u8])>; 2]; 2];

/// Entropy-coded bits of a baseline scan (T.81 F.1.2, F.2.2), read with
/// byte stuffing and restart markers.
struct Bits<'b> {
    bytes: &'b [u8],
    at: usize,
    buffer: u32,
    count: u32,
}

impl Bits<'_> {
    fn bit(&mut self) -> Option<u32> {
        if self.count == 0 {
            let byte = *self.bytes.get(self.at)?;
            if byte == 0xff {
                // A stuffed zero follows a data 0xFF; any other byte is a
                // marker, which ends the data early.
                if self.bytes.get(self.at + 1) != Some(&0) {
                    return None;
                }
                self.at += 1;
            }
            self.at += 1;
            self.buffer = u32::from(byte);
            self.count = 8;
        }
        self.count -= 1;
        Some((self.buffer >> self.count) & 1)
    }

    fn bits(&mut self, count: u8) -> Option<()> {
        for _ in 0..count {
            self.bit()?;
        }
        Some(())
    }

    /// The next symbol of a canonical code (T.81 Annex C, F.2.2.3).
    fn symbol(&mut self, (counts, symbols): (&[u8], &[u8])) -> Option<u8> {
        let (mut code, mut first, mut index) = (0u32, 0u32, 0usize);
        for &count in counts {
            code = code << 1 | self.bit()?;
            let count = u32::from(count);
            if code < first + count {
                return symbols.get(index + (code - first) as usize).copied();
            }
            index += count as usize;
            first = (first + count) << 1;
        }
        None
    }

    /// Discards the padding of a finished segment: the rest of its byte,
    /// all ones (F.1.2.3).
    fn align(&mut self) -> Option<()> {
        let mask = (1u32 << self.count) - 1;
        (self.buffer & mask == mask).then(|| self.count = 0)
    }
}

/// Walks a baseline scan's entropy-coded data from `start` (T.81 F.1.2):
/// every MCU of every component decodes with the scan's tables, restart
/// markers come in sequence after each restart interval, and the data ends
/// with its padding right before a marker. Returns that marker's offset.
fn entropy_coded_data(
    bytes: &[u8],
    start: usize,
    frame: &[u8],
    scan: &[u8],
    tables: &HuffmanTables<'_>,
    restart: u16,
) -> Result<usize, String> {
    let decode = || unsupported("image-decode");
    let height = u32::from(u16::from_be_bytes([frame[1], frame[2]]));
    let width = u32::from(u16::from_be_bytes([frame[3], frame[4]]));
    let components: Vec<&[u8]> = frame[6..].chunks(3).collect();
    let selected: Vec<&[u8]> = scan[1..1 + 2 * usize::from(scan[0])].chunks(2).collect();
    // The scan holds every frame component (one interleaved scan).
    if selected.len() != components.len()
        || components
            .iter()
            .zip(&selected)
            .any(|(component, scanned)| component[0] != scanned[0])
    {
        return Err(unsupported("image-structure"));
    }
    let (h_max, v_max) = components.iter().fold((1, 1), |(h, v), c| {
        (h.max(u32::from(c[1] >> 4)), v.max(u32::from(c[1] & 15)))
    });
    // Blocks per MCU for each component, and the MCU count (A.2).
    let (blocks, mcus): (Vec<u32>, u32) = if components.len() == 1 {
        (vec![1], width.div_ceil(8) * height.div_ceil(8))
    } else {
        (
            components
                .iter()
                .map(|c| u32::from(c[1] >> 4) * u32::from(c[1] & 15))
                .collect(),
            width.div_ceil(8 * h_max) * height.div_ceil(8 * v_max),
        )
    };
    let coding: Vec<_> = selected
        .iter()
        .map(|c| {
            Some((
                tables[0][usize::from(c[1] >> 4)]?,
                tables[1][usize::from(c[1] & 15)]?,
            ))
        })
        .collect::<Option<_>>()
        .ok_or_else(decode)?;
    let mut bits = Bits {
        bytes,
        at: start,
        buffer: 0,
        count: 0,
    };
    let mut next_restart = 0u8;
    for mcu in 0..mcus {
        if restart > 0 && mcu > 0 && mcu % u32::from(restart) == 0 {
            bits.align().ok_or_else(decode)?;
            if bytes.get(bits.at..bits.at + 2) != Some(&[0xff, 0xd0 + next_restart]) {
                return Err(decode());
            }
            bits.at += 2;
            next_restart = (next_restart + 1) % 8;
        }
        for (&count, &(dc, ac)) in blocks.iter().zip(&coding) {
            for _ in 0..count {
                let category = bits.symbol(dc).ok_or_else(decode)?;
                bits.bits(category).ok_or_else(decode)?;
                let mut k = 1;
                while k < 64 {
                    let rs = bits.symbol(ac).ok_or_else(decode)?;
                    let (run, size) = (rs >> 4, rs & 15);
                    if size == 0 {
                        if run != 15 {
                            break;
                        }
                        k += 16;
                    } else {
                        k += usize::from(run);
                        bits.bits(size).ok_or_else(decode)?;
                        k += 1;
                    }
                    if k > 64 {
                        return Err(decode());
                    }
                }
            }
        }
    }
    bits.align().ok_or_else(decode)?;
    // The marker after the data (no further restart marker or data).
    match bytes.get(bits.at..bits.at + 2) {
        Some([0xff, marker]) if !matches!(marker, 0x00 | 0xd0..=0xd7) => Ok(bits.at),
        _ => Err(decode()),
    }
}

/// DQT (T.81 B.2.4.1): 8-bit tables 0-3 of 64 non-zero values. Decoders
/// differ on other tables (measured by the reviewer).
fn quantisation_tables_valid(mut segment: &[u8]) -> bool {
    while let Some((&header, rest)) = segment.split_first() {
        if header > 3 || rest.len() < 64 || rest[..64].contains(&0) {
            return false;
        }
        segment = &rest[64..];
    }
    true
}

/// DHT (T.81 B.2.4.2, F.1.2.1.2, F.1.2.2.1, C): baseline DC or AC tables
/// 0-1 whose
/// code lengths form a prefix code that leaves the all-ones code unused,
/// with DC categories up to 11 and AC run/size symbols of sizes 1-10 (or
/// EOB and ZRL). Decoders differ on other tables (measured by the reviewer).
fn huffman_tables_valid(mut segment: &[u8]) -> bool {
    while let Some((&header, rest)) = segment.split_first() {
        let (class, id) = (header >> 4, header & 15);
        if class > 1 || id > 1 || rest.len() < 16 {
            return false;
        }
        let counts = &rest[..16];
        let total: usize = counts.iter().map(|&c| usize::from(c)).sum();
        let Some(symbols) = rest.get(16..16 + total) else {
            return false;
        };
        // Canonical codes (Annex C): each length's codes must fit, and the
        // longest length must leave its all-ones code free.
        let mut code = 0u32;
        for (length, &count) in counts.iter().enumerate() {
            code += u32::from(count);
            if code > 1 << (length + 1) || (code == 1 << (length + 1) && count > 0) {
                return false;
            }
            code <<= 1;
        }
        let symbol_valid = |&symbol: &u8| match class {
            0 => symbol <= 11,
            _ => matches!(symbol & 15, 1..=10) || symbol == 0x00 || symbol == 0xf0,
        };
        if total == 0 || !symbols.iter().all(symbol_valid) {
            return false;
        }
        segment = &rest[16 + total..];
    }
    true
}

/// SOF0 (T.81 B.2.2): 8-bit samples, one or three components with sampling
/// factors 1-4 and quantisation tables 0-3.
fn frame_header_valid(segment: &[u8]) -> bool {
    let count = segment.get(5).copied().unwrap_or(0);
    segment.first() == Some(&8)
        && matches!(count, 1 | 3)
        && segment.len() == 6 + 3 * usize::from(count)
        && segment[6..].chunks(3).all(|component| {
            matches!(component[1] >> 4, 1..=4)
                && matches!(component[1] & 15, 1..=4)
                && component[2] <= 3
        })
}

/// SOS (T.81 B.2.3) for a baseline frame: one to four components with DC
/// and AC tables 0-1, spectral selection 0-63 and no successive
/// approximation.
fn scan_header_valid(segment: &[u8]) -> bool {
    let count = segment.first().copied().unwrap_or(0);
    matches!(count, 1..=4)
        && segment.len() == 4 + 2 * usize::from(count)
        && segment[1..1 + 2 * usize::from(count)]
            .chunks(2)
            .all(|component| component[1] >> 4 <= 1 && component[1] & 15 <= 1)
        && segment[1 + 2 * usize::from(count)..] == [0, 63, 0]
}

/// The orientation tag (0x0112) of an EXIF block's first IFD (TIFF 6.0
/// §2), when present and readable.
fn exif_orientation(tiff: &[u8]) -> Option<u16> {
    let big = match tiff.get(..2)? {
        b"MM" => true,
        b"II" => false,
        _ => return None,
    };
    let u16_at = |at: usize| -> Option<u16> {
        let b: [u8; 2] = tiff.get(at..at + 2)?.try_into().ok()?;
        Some(if big {
            u16::from_be_bytes(b)
        } else {
            u16::from_le_bytes(b)
        })
    };
    let u32_at = |at: usize| -> Option<u32> {
        let b: [u8; 4] = tiff.get(at..at + 4)?.try_into().ok()?;
        Some(if big {
            u32::from_be_bytes(b)
        } else {
            u32::from_le_bytes(b)
        })
    };
    let ifd = u32_at(4)? as usize;
    let count = usize::from(u16_at(ifd)?);
    (0..count).find_map(|entry| {
        let at = ifd + 2 + 12 * entry;
        (u16_at(at)? == 0x0112).then(|| u16_at(at + 8)).flatten()
    })
}

fn png(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    png_structure(bytes)?;
    let limits = png::Limits {
        bytes: MAX_DECODE_BYTES as usize,
    };
    let mut decoder = png::Decoder::new_with_limits(Cursor::new(bytes), limits);
    // Chunk CRCs and the zlib Adler-32 are verified: LibreOffice rejects
    // image data with a wrong check value where a browser paints it.
    decoder.ignore_checksums(false);
    let mut reader = decoder.read_info().map_err(|_| decode())?;
    let info = reader.info();
    let (width, height) = (info.width, info.height);
    check_size(u64::from(width), u64::from(height))?;
    // An animated PNG plays in a browser; LibreOffice is not measured.
    if info.animation_control.is_some() {
        return Err(unsupported("image-animation"));
    }
    let palette = (info.color_type == png::ColorType::Indexed)
        .then(|| info.palette.as_ref().map_or(0, |palette| palette.len() / 3));
    charge(budget, u64::from(width), u64::from(height))?;
    let size = reader.output_buffer_size().ok_or_else(decode)?;
    let mut buffer = vec![0; size];
    let output = reader.next_frame(&mut buffer).map_err(|_| decode())?;
    // Indices past the palette paint differently (a browser's choice,
    // LibreOffice's black), so each one must name an entry.
    if let Some(entries) = palette {
        let depth = output.bit_depth as usize;
        for row in buffer.chunks(output.line_size).take(height as usize) {
            for x in 0..width as usize {
                let bit = x * depth;
                let index = (row[bit / 8] >> (8 - depth - bit % 8)) & ((1u16 << depth) - 1) as u8;
                if usize::from(index) >= entries {
                    return Err(unsupported("image-palette"));
                }
            }
        }
    }
    reader.finish().map_err(|_| decode())
}

fn gif(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    gif_structure(bytes)?;
    let mut options = gif::DecodeOptions::new();
    options.set_color_output(gif::ColorOutput::Indexed);
    options.set_memory_limit(gif::MemoryLimit::Bytes(
        std::num::NonZeroU64::new(MAX_DECODE_BYTES).expect("non-zero"),
    ));
    options.check_frame_consistency(true);
    let mut decoder = options
        .read_info(Cursor::new(bytes))
        .map_err(|_| decode())?;
    check_size(u64::from(decoder.width()), u64::from(decoder.height()))?;
    let mut frames = 0usize;
    while let Some(frame) = decoder.next_frame_info().map_err(|_| decode())? {
        // Every frame lies within the screen; each is charged before it is
        // decoded.
        charge(budget, u64::from(frame.width), u64::from(frame.height))?;
        let mut buffer = vec![0; decoder.buffer_size()];
        decoder
            .read_into_buffer(&mut buffer)
            .map_err(|_| decode())?;
        // Indices past the colour table paint differently per decoder.
        let entries = decoder.palette().map_err(|_| decode())?.len() / 3;
        if buffer.iter().any(|&index| usize::from(index) >= entries) {
            return Err(unsupported("image-palette"));
        }
        frames += 1;
    }
    if frames == 0 {
        return Err(decode());
    }
    Ok(())
}

fn jpeg(bytes: &[u8], budget: &mut u64) -> Result<(), String> {
    let decode = || unsupported("image-decode");
    jpeg_structure(bytes)?;
    let options = zune_core::options::DecoderOptions::default()
        .set_max_width(MAX_SIDE as usize)
        .set_max_height(MAX_SIDE as usize)
        .set_strict_mode(true);
    let mut decoder = zune_jpeg::JpegDecoder::new_with_options(Cursor::new(bytes), options);
    decoder.decode_headers().map_err(|_| decode())?;
    let (width, height) = decoder.dimensions().ok_or_else(decode)?;
    check_size(width as u64, height as u64)?;
    charge(budget, width as u64, height as u64)?;
    decoder.decode().map_err(|_| decode())?;
    Ok(())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    fn dec(bytes: &[u8]) -> Result<&'static str, String> {
        decode(bytes, &mut MAX_DOCUMENT_PIXELS.clone())
    }

    /// A PNG with `chunks` (type, data) after its IHDR, CRCs computed.
    fn png_chunks(ihdr: &[u8], chunks: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
        let mut out = b"\x89PNG\r\n\x1a\n".to_vec();
        let mut chunk = |kind: &[u8; 4], data: &[u8]| {
            out.extend((data.len() as u32).to_be_bytes());
            let mut crc = crc32fast::Hasher::new();
            crc.update(kind);
            crc.update(data);
            out.extend(kind);
            out.extend(data);
            out.extend(crc.finalize().to_be_bytes());
        };
        chunk(b"IHDR", ihdr);
        for (kind, data) in chunks {
            chunk(kind, data);
        }
        chunk(b"IEND", &[]);
        out
    }

    fn zlib(raw: &[u8]) -> Vec<u8> {
        use std::io::Write;
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(raw).unwrap();
        encoder.finish().unwrap()
    }

    /// A 2×1 RGB PNG (one red, one blue pixel).
    pub(crate) fn png_bytes() -> Vec<u8> {
        let mut out = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut out, 2, 1);
            encoder.set_color(png::ColorType::Rgb);
            let mut writer = encoder.write_header().unwrap();
            writer.write_image_data(&[255, 0, 0, 0, 0, 255]).unwrap();
        }
        out
    }

    #[test]
    fn decodes_png_and_rejects_damage() {
        let bytes = png_bytes();
        assert_eq!(dec(&bytes).unwrap(), "image/png");
        // A truncated stream stops inside the image data.
        assert_eq!(
            dec(&bytes[..bytes.len() - 20]).unwrap_err(),
            unsupported("image-decode")
        );
        let mut corrupt = bytes.clone();
        let data = corrupt.len() - 20;
        corrupt[data] ^= 0xff;
        assert_eq!(dec(&corrupt).unwrap_err(), unsupported("image-decode"));
    }

    #[test]
    fn verifies_png_checksums_and_palette_indices() {
        // 1×1 8-bit indexed, one palette entry.
        let ihdr = [0, 0, 0, 1, 0, 0, 0, 1, 8, 3, 0, 0, 0];
        let palette = (b"PLTE", vec![255, 0, 0]);
        let valid = png_chunks(&ihdr, &[palette.clone(), (b"IDAT", zlib(&[0, 0]))]);
        assert_eq!(dec(&valid).unwrap(), "image/png");
        // A pixel index past the palette.
        let outside = png_chunks(&ihdr, &[palette.clone(), (b"IDAT", zlib(&[0, 255]))]);
        assert_eq!(dec(&outside).unwrap_err(), unsupported("image-palette"));
        // A wrong zlib Adler-32 under a correct chunk CRC.
        let mut data = zlib(&[0, 0]);
        let end = data.len() - 1;
        data[end] ^= 1;
        let adler = png_chunks(&ihdr, &[palette.clone(), (b"IDAT", data)]);
        assert_eq!(dec(&adler).unwrap_err(), unsupported("image-decode"));
        // Colour-space, animation and unknown chunks; a damaged ancillary
        // chunk; data after IEND.
        let idat = (b"IDAT", zlib(&[0, 0]));
        for kind in [b"gAMA", b"iCCP", b"acTL", b"ABCD", b"aaaa"] {
            let extra = png_chunks(
                &ihdr,
                &[palette.clone(), (kind, vec![0, 1, 134, 160]), idat.clone()],
            );
            assert_eq!(dec(&extra).unwrap_err(), unsupported("image-structure"));
        }
        // Transparency samples out of range, a palette in an RGB image.
        let grey = [0, 0, 0, 1, 0, 0, 0, 1, 8, 0, 0, 0, 0];
        let out_of_range = png_chunks(&grey, &[(b"tRNS", vec![0xff, 0xff]), idat.clone()]);
        assert_eq!(
            dec(&out_of_range).unwrap_err(),
            unsupported("image-structure")
        );
        let in_range = png_chunks(&grey, &[(b"tRNS", vec![0, 0xff]), idat.clone()]);
        assert_eq!(dec(&in_range).unwrap(), "image/png");
        let rgb = [0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0];
        let rgb_palette = png_chunks(
            &rgb,
            &[(b"PLTE", vec![255, 0, 0]), (b"IDAT", zlib(&[0, 1, 2, 3]))],
        );
        assert_eq!(
            dec(&rgb_palette).unwrap_err(),
            unsupported("image-structure")
        );
        let text = png_chunks(
            &ihdr,
            &[palette.clone(), (b"tEXt", b"a\0b".to_vec()), idat.clone()],
        );
        assert_eq!(dec(&text).unwrap(), "image/png");
        let mut damaged = text.clone();
        let at = damaged.windows(4).position(|w| w == b"tEXt").unwrap() + 4;
        damaged[at] ^= 1;
        assert_eq!(dec(&damaged).unwrap_err(), unsupported("image-decode"));
        assert_eq!(
            dec(&[&valid[..], &[0]].concat()).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn charges_every_decoded_pixel_against_the_document() {
        let bytes = png_bytes();
        let mut budget = 3;
        assert_eq!(decode(&bytes, &mut budget).unwrap(), "image/png");
        assert_eq!(budget, 1);
        assert!(decode(&bytes, &mut budget).is_err());
    }

    /// A GIF of `frames` (left, top, width, height, indices) on a
    /// `width` × `height` screen with a two-entry table.
    fn gif_bytes(width: u16, height: u16, frames: &[(u16, u16, u16, u16, &[u8])]) -> Vec<u8> {
        let mut out = Vec::new();
        {
            let mut encoder =
                gif::Encoder::new(&mut out, width, height, &[0, 0, 0, 255, 0, 0]).unwrap();
            for (left, top, w, h, indices) in frames {
                let frame = gif::Frame {
                    left: *left,
                    top: *top,
                    width: *w,
                    height: *h,
                    buffer: std::borrow::Cow::Borrowed(indices),
                    ..Default::default()
                };
                encoder.write_frame(&frame).unwrap();
            }
        }
        out
    }

    #[test]
    fn decodes_gif_frames() {
        let one = gif_bytes(2, 2, &[(0, 0, 2, 2, &[0, 1, 1, 0])]);
        assert_eq!(dec(&one).unwrap(), "image/gif");
        // One 2×2 frame: 4 pixels.
        assert!(decode(&one, &mut 3).is_err());
        // Several frames animate; a partial frame is placed differently.
        let two = gif_bytes(
            2,
            2,
            &[(0, 0, 2, 2, &[0, 1, 1, 0]), (0, 0, 2, 2, &[1, 1, 1, 1])],
        );
        assert_eq!(dec(&two).unwrap_err(), unsupported("image-animation"));
        let partial = gif_bytes(4, 4, &[(1, 1, 1, 1, &[1])]);
        assert_eq!(dec(&partial).unwrap_err(), unsupported("image-structure"));
        // A plain text extension (GIF89a §25) before the image.
        let mut text = one.clone();
        let image = text.iter().position(|&b| b == 0x2c).unwrap();
        text.splice(image..image, [0x21, 0x01, 1, 0, 0]);
        assert_eq!(dec(&text).unwrap_err(), unsupported("image-structure"));
        // An index past the two-entry table.
        let outside = gif_bytes(1, 1, &[(0, 0, 1, 1, &[3])]);
        assert_eq!(dec(&outside).unwrap_err(), unsupported("image-palette"));
        // Truncated, or data after the trailer.
        assert_eq!(
            dec(&one[..one.len() - 3]).unwrap_err(),
            unsupported("image-decode")
        );
        assert_eq!(
            dec(&[&one[..], &[0]].concat()).unwrap_err(),
            unsupported("image-decode")
        );
    }

    #[test]
    fn decodes_jpeg_and_rejects_truncation() {
        // A 1×1 baseline JPEG.
        let bytes: &[u8] = &JPEG_1X1;
        assert_eq!(dec(bytes).unwrap(), "image/jpeg");
        // Without its end-of-image marker, or with data after it.
        assert_eq!(
            dec(&bytes[..bytes.len() - 2]).unwrap_err(),
            unsupported("image-decode")
        );
        assert_eq!(
            dec(&[bytes, &[0]].concat()).unwrap_err(),
            unsupported("image-decode")
        );
        assert_eq!(
            dec(&bytes[..bytes.len() - 6]).unwrap_err(),
            unsupported("image-decode")
        );
        // Entropy-coded data cut short before an intact end-of-image marker.
        let cut = [&bytes[..bytes.len() - 3], &bytes[bytes.len() - 2..]].concat();
        assert_eq!(dec(&cut).unwrap_err(), unsupported("image-decode"));
        // A segment inserted after SOI.
        let with = |segment: &[u8]| [&bytes[..2], segment, &bytes[2..]].concat();
        // A colour profile, a rotating EXIF orientation, a non-YCbCr Adobe
        // transform.
        let icc = with(b"\xff\xe2\x00\x10ICC_PROFILE\x00\x01\x01");
        assert_eq!(dec(&icc).unwrap_err(), unsupported("image-colour"));
        let exif = |orientation: u8| {
            with(
                &[
                    b"\xff\xe1\x00\x22Exif\x00\x00MM\x00\x2a\x00\x00\x00\x08\x00\x01".as_slice(),
                    &[
                        0x01,
                        0x12,
                        0x00,
                        0x03,
                        0,
                        0,
                        0,
                        1,
                        0,
                        orientation,
                        0,
                        0,
                        0,
                        0,
                        0,
                        0,
                    ],
                ]
                .concat(),
            )
        };
        assert_eq!(dec(&exif(1)).unwrap(), "image/jpeg");
        assert_eq!(dec(&exif(6)).unwrap_err(), unsupported("image-orientation"));
        let adobe = with(b"\xff\xee\x00\x0eAdobe\x00\x64\x00\x00\x00\x00\x00");
        assert_eq!(dec(&adobe).unwrap_err(), unsupported("image-colour"));
        // A zero quantiser, a DC category past 11, a code using all ones.
        let replace = |marker: u8, offset: usize, value: u8| {
            let at = bytes.windows(2).position(|w| w == [0xff, marker]).unwrap();
            let mut out = bytes.to_vec();
            out[at + offset] = value;
            out
        };
        assert_eq!(
            dec(&replace(0xdb, 5, 0)).unwrap_err(),
            unsupported("image-structure")
        );
        assert_eq!(
            dec(&replace(0xc4, 21, 12)).unwrap_err(),
            unsupported("image-structure")
        );
        assert_eq!(
            dec(&replace(0xc4, 5, 2)).unwrap_err(),
            unsupported("image-structure")
        );
        // An unknown marker, a second SOI, a progressive frame.
        assert_eq!(
            dec(&with(b"\xff\x60\x00\x02")).unwrap_err(),
            unsupported("image-structure")
        );
        assert_eq!(
            dec(&with(b"\xff\xd8")).unwrap_err(),
            unsupported("image-structure")
        );
        let progressive: Vec<u8> = bytes
            .windows(2)
            .position(|w| w == [0xff, 0xc0])
            .map(|at| [&bytes[..at + 1], &[0xc2], &bytes[at + 2..]].concat())
            .unwrap();
        assert_eq!(
            dec(&progressive).unwrap_err(),
            unsupported("image-structure")
        );
    }

    #[test]
    fn rejects_other_formats_and_sizes() {
        assert_eq!(dec(b"BM....").unwrap_err(), unsupported("image-format"));
        assert_eq!(
            dec(b"<svg xmlns='http://www.w3.org/2000/svg'/>").unwrap_err(),
            unsupported("image-format")
        );
        assert_eq!(check_size(0, 1).unwrap_err(), unsupported("image-size"));
        assert_eq!(
            check_size(32_769, 1).unwrap_err(),
            unsupported("image-size")
        );
        assert_eq!(
            check_size(8_000, 8_000).unwrap_err(),
            unsupported("image-size")
        );
        assert!(check_size(6_000, 6_000).is_ok());
    }

    /// 1×1 grey baseline JPEG (libjpeg output).
    const JPEG_1X1: [u8; 159] = [
        0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00,
        0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xdb, 0x00, 0x43, 0x00, 0x10, 0x0b, 0x0c, 0x0e, 0x0c,
        0x0a, 0x10, 0x0e, 0x0d, 0x0e, 0x12, 0x11, 0x10, 0x13, 0x18, 0x28, 0x1a, 0x18, 0x16, 0x16,
        0x18, 0x31, 0x23, 0x25, 0x1d, 0x28, 0x3a, 0x33, 0x3d, 0x3c, 0x39, 0x33, 0x38, 0x37, 0x40,
        0x48, 0x5c, 0x4e, 0x40, 0x44, 0x57, 0x45, 0x37, 0x38, 0x50, 0x6d, 0x51, 0x57, 0x5f, 0x62,
        0x67, 0x68, 0x67, 0x3e, 0x4d, 0x71, 0x79, 0x70, 0x64, 0x78, 0x5c, 0x65, 0x67, 0x63, 0xff,
        0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xff, 0xc4, 0x00,
        0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08,
        0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x3f, 0xff, 0xd9,
    ];
}
