//! The cells of an msofillPattern bitmap (MS-ODRAW 2.4.11), and the bitmap
//! this reader writes for a pattern on a flipped shape.
//!
//! PowerPoint stores a pattern as an 8x8 two-colour bitmap, or as a 10x10
//! one whose top-left 8x8 pixels are the pattern. White pixels take
//! fillColor and black pixels fillBackColor.

/// One byte per row, most significant bit first; a set bit is a white pixel.
pub(super) type Cells = [u8; 8];

/// Longest pattern bitmap this module decodes. A PowerPoint pattern bitmap
/// is a few hundred bytes.
const MAX_BITMAP_BYTES: usize = 4096;

/// The cells of a stored pattern bitmap. `None`: not a black-and-white 8x8
/// or 10x10 bitmap in an encoding read here (GIF, or a non-interlaced
/// palette or grayscale PNG).
pub(super) fn cells(extension: &str, bytes: &[u8]) -> Result<Option<Cells>, String> {
    if bytes.len() > MAX_BITMAP_BYTES {
        return Ok(None);
    }
    match extension {
        "gif" => Ok(super::metro::gif_pattern_cells(bytes)?.map(|(_, cells)| cells)),
        "png" => Ok(png_cells(bytes)),
        _ => Ok(None),
    }
}

/// `cells` as seen through a horizontal and/or vertical reflection.
pub(super) fn mirrored(cells: Cells, flip_h: bool, flip_v: bool) -> Cells {
    std::array::from_fn(|row| {
        let source = cells[if flip_v { 7 - row } else { row }];
        if flip_h {
            source.reverse_bits()
        } else {
            source
        }
    })
}

/// W3C PNG (Second Edition) sections 5, 9 and 11: a non-interlaced image of
/// colour type 0 (grayscale) or 3 (palette).
fn png_cells(bytes: &[u8]) -> Option<Cells> {
    let mut at = bytes.strip_prefix(b"\x89PNG\r\n\x1a\n").map(|_| 8)?;
    let mut header = None;
    let mut palette: &[u8] = &[];
    let mut transparency: &[u8] = &[];
    let mut data = Vec::new();
    loop {
        let length = u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?) as usize;
        let kind = bytes.get(at + 4..at + 8)?;
        let body = bytes.get(at + 8..(at + 8).checked_add(length)?)?;
        at += length + 12;
        match kind {
            b"IHDR" => header = Some(<[u8; 13]>::try_from(body).ok()?),
            b"PLTE" => palette = body,
            b"tRNS" => transparency = body,
            b"IDAT" => data.extend_from_slice(body),
            b"IEND" => break,
            _ => {}
        }
    }
    let header = header?;
    let size = u32::from_be_bytes(header[..4].try_into().ok()?);
    if !matches!(size, 8 | 10) || header[4..8] != header[..4] {
        return None;
    }
    let size = size as usize;
    let (depth, kind) = (usize::from(header[8]), header[9]);
    if !matches!(depth, 1 | 2 | 4 | 8) || !matches!(kind, 0 | 3) || header[10..] != [0, 0, 0] {
        return None;
    }
    let row_bytes = (size * depth).div_ceil(8);
    let expanded = size * (row_bytes + 1);
    let mut rows =
        crate::officeart::inflate::inflate_exact(&data, expanded, data.len(), expanded).ok()?;
    // Section 9.2 filters, with one byte per complete pixel at these depths.
    for row in 0..size {
        let start = row * (row_bytes + 1);
        let filter = rows[start];
        for index in 0..row_bytes {
            let at = start + 1 + index;
            let left = if index > 0 { rows[at - 1] } else { 0 };
            let (up, up_left) = if row > 0 {
                let above = at - (row_bytes + 1);
                (rows[above], if index > 0 { rows[above - 1] } else { 0 })
            } else {
                (0, 0)
            };
            let predictor = match filter {
                0 => 0,
                1 => left,
                2 => up,
                3 => ((u16::from(left) + u16::from(up)) / 2) as u8,
                4 => {
                    let estimate = i16::from(left) + i16::from(up) - i16::from(up_left);
                    let distance = |value: u8| (estimate - i16::from(value)).abs();
                    if distance(left) <= distance(up) && distance(left) <= distance(up_left) {
                        left
                    } else if distance(up) <= distance(up_left) {
                        up
                    } else {
                        up_left
                    }
                }
                _ => return None,
            };
            rows[at] = rows[at].wrapping_add(predictor);
        }
    }
    let maximum = (1usize << depth) - 1;
    let mut cells = [0u8; 8];
    for (row, cell) in cells.iter_mut().enumerate() {
        let line = &rows[row * (row_bytes + 1) + 1..];
        for column in 0..8 {
            let bit = column * depth;
            let sample = usize::from(line[bit / 8] >> (8 - depth - bit % 8)) & maximum;
            let white = if kind == 3 {
                if transparency.get(sample).is_some_and(|alpha| *alpha != 255) {
                    return None;
                }
                match palette.get(sample * 3..sample * 3 + 3)? {
                    [255, 255, 255] => true,
                    [0, 0, 0] => false,
                    _ => return None,
                }
            } else {
                // A grayscale tRNS names one fully transparent sample value.
                if transparency.len() == 2
                    && usize::from(u16::from_be_bytes([transparency[0], transparency[1]])) == sample
                {
                    return None;
                }
                match sample {
                    0 => false,
                    value if value == maximum => true,
                    _ => return None,
                }
            };
            if white {
                *cell |= 0x80 >> column;
            }
        }
    }
    Some(cells)
}

/// An 8x8 one-bit grayscale PNG of `cells` (white where a bit is set), with
/// its rows in one stored DEFLATE block (RFC 1950, RFC 1951 3.2.4).
pub(super) fn png(cells: Cells) -> Vec<u8> {
    fn chunk(output: &mut Vec<u8>, kind: &[u8; 4], body: &[u8]) {
        output.extend_from_slice(&(body.len() as u32).to_be_bytes());
        let start = output.len();
        output.extend_from_slice(kind);
        output.extend_from_slice(body);
        // PNG section 5.5: CRC-32 of the chunk type and data.
        let mut crc = u32::MAX;
        for byte in &output[start..] {
            crc ^= u32::from(*byte);
            for _ in 0..8 {
                crc = (crc >> 1) ^ (0xedb8_8320 & (crc & 1).wrapping_neg());
            }
        }
        output.extend_from_slice(&(!crc).to_be_bytes());
    }
    let mut rows = [0u8; 16];
    for (row, cell) in cells.iter().enumerate() {
        rows[row * 2 + 1] = *cell;
    }
    let (mut low, mut high) = (1u32, 0u32);
    for byte in rows {
        low = (low + u32::from(byte)) % 65_521;
        high = (high + low) % 65_521;
    }
    let mut stream = vec![0x78, 0x01, 0x01, 16, 0, !16, 0xff];
    stream.extend_from_slice(&rows);
    stream.extend_from_slice(&((high << 16) | low).to_be_bytes());
    let mut output = b"\x89PNG\r\n\x1a\n".to_vec();
    chunk(
        &mut output,
        b"IHDR",
        &[0, 0, 0, 8, 0, 0, 0, 8, 1, 0, 0, 0, 0],
    );
    chunk(&mut output, b"IDAT", &stream);
    chunk(&mut output, b"IEND", &[]);
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    const DIAGONAL: Cells = [0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66, 0x33, 0x99];

    #[test]
    fn written_bitmaps_read_back_and_reflect_their_cells() {
        // The written PNG is a valid image of the cells: this module's own
        // reader, which checks the zlib stream and its checksum, returns them.
        let image = png(DIAGONAL);
        assert_eq!(cells("png", &image).unwrap(), Some(DIAGONAL));
        // CRC-32 of "IEND" with no data (PNG section 5.5), a fixed value.
        assert!(image.ends_with(&[0, 0, 0, 0, b'I', b'E', b'N', b'D', 0xae, 0x42, 0x60, 0x82]));
        // A down diagonal seen through one reflection is an up diagonal;
        // through both it is itself again.
        let up = [0x33, 0x66, 0xcc, 0x99, 0x33, 0x66, 0xcc, 0x99];
        assert_eq!(mirrored(DIAGONAL, true, false), up);
        assert_eq!(
            mirrored(DIAGONAL, false, true),
            [0x99, 0x33, 0x66, 0xcc, 0x99, 0x33, 0x66, 0xcc]
        );
        assert_eq!(
            mirrored(mirrored(DIAGONAL, true, true), true, true),
            DIAGONAL
        );
        assert_eq!(mirrored(DIAGONAL, false, false), DIAGONAL);
    }

    /// A stored-block PNG with the given header fields and raw filtered rows.
    fn stored_png(
        size: u8,
        depth: u8,
        kind: u8,
        extra: &[(&[u8; 4], &[u8])],
        rows: &[u8],
    ) -> Vec<u8> {
        let mut output = png(DIAGONAL)[..8].to_vec();
        let mut chunk = |kind: &[u8; 4], body: &[u8]| {
            output.extend_from_slice(&(body.len() as u32).to_be_bytes());
            output.extend_from_slice(kind);
            output.extend_from_slice(body);
            output.extend_from_slice(&[0; 4]); // The CRC is not read.
        };
        chunk(
            b"IHDR",
            &[0, 0, 0, size, 0, 0, 0, size, depth, kind, 0, 0, 0],
        );
        for (kind, body) in extra {
            chunk(kind, body);
        }
        let (mut low, mut high) = (1u32, 0u32);
        for byte in rows {
            low = (low + u32::from(*byte)) % 65_521;
            high = (high + low) % 65_521;
        }
        let length = rows.len() as u16;
        let mut stream = vec![0x78, 0x01, 0x01];
        stream.extend_from_slice(&length.to_le_bytes());
        stream.extend_from_slice(&(!length).to_le_bytes());
        stream.extend_from_slice(rows);
        stream.extend_from_slice(&((high << 16) | low).to_be_bytes());
        chunk(b"IDAT", &stream);
        chunk(b"IEND", &[]);
        output
    }

    #[test]
    fn reads_palette_bitmaps_as_powerpoint_stores_them() {
        // PowerPoint 16: 8x8, two bits per pixel, palette white, black and
        // an unused transparent entry. Index 0 (white) is the foreground.
        let palette: &[u8] = &[255, 255, 255, 0, 0, 0, 0, 0, 0];
        let transparency: &[u8] = &[255, 255, 0];
        let chunks: &[(&[u8; 4], &[u8])] = &[(b"PLTE", palette), (b"tRNS", transparency)];
        // Rows of index pairs 0,0,1,1,0,0,1,1 shifted by one pixel per row;
        // the second row uses the Sub filter and the third the Up filter.
        let plain = |cells: u8| -> [u8; 2] {
            let mut packed = 0u16;
            for column in 0..8 {
                let index = u16::from(cells & (0x80 >> column) == 0);
                packed |= index << (14 - 2 * column);
            }
            packed.to_be_bytes()
        };
        let mut rows = Vec::new();
        for (row, cell) in DIAGONAL.iter().enumerate() {
            let [a, b] = plain(*cell);
            match row {
                1 => rows.extend([1, a, b.wrapping_sub(a)]),
                2 => {
                    let [up_a, up_b] = plain(DIAGONAL[1]);
                    rows.extend([2, a.wrapping_sub(up_a), b.wrapping_sub(up_b)]);
                }
                _ => rows.extend([0, a, b]),
            }
        }
        let image = stored_png(8, 2, 3, chunks, &rows);
        assert_eq!(cells("png", &image).unwrap(), Some(DIAGONAL));
        // A used transparent entry, another colour, an unknown filter, a
        // truncated stream and another size are not pattern bitmaps.
        let used: &[(&[u8; 4], &[u8])] = &[(b"PLTE", palette), (b"tRNS", &[0, 255, 0])];
        assert_eq!(
            cells("png", &stored_png(8, 2, 3, used, &rows)).unwrap(),
            None
        );
        let colored: &[(&[u8; 4], &[u8])] = &[(b"PLTE", &[255, 255, 255, 0, 0, 1])];
        assert_eq!(
            cells("png", &stored_png(8, 2, 3, colored, &rows)).unwrap(),
            None
        );
        let mut filtered = rows.clone();
        filtered[0] = 5;
        assert_eq!(
            cells("png", &stored_png(8, 2, 3, chunks, &filtered)).unwrap(),
            None
        );
        assert_eq!(
            cells("png", &stored_png(8, 2, 3, chunks, &rows[..21])).unwrap(),
            None
        );
        assert_eq!(
            cells("png", &stored_png(9, 2, 3, chunks, &rows)).unwrap(),
            None
        );
        assert_eq!(cells("png", &image[..image.len() - 13]).unwrap(), None);
        assert_eq!(cells("jpg", &image).unwrap(), None);
        // An 8x8 GIF (LZW with a clear code before every pixel pair, so
        // every code is three bits; GIF89a Appendix F).
        let mut codes = Vec::new();
        for row in DIAGONAL {
            for pair in 0..4 {
                codes.push(4u8);
                for column in [pair * 2, pair * 2 + 1] {
                    codes.push(u8::from(row & (0x80 >> column) != 0));
                }
            }
        }
        codes.push(5);
        let mut packed = vec![0u8; (codes.len() * 3).div_ceil(8)];
        for (index, code) in codes.iter().enumerate() {
            for bit in 0..3 {
                if code & (1 << bit) != 0 {
                    packed[(index * 3 + bit) / 8] |= 1 << ((index * 3 + bit) % 8);
                }
            }
        }
        // Global palette: 0 black, 1 white.
        let mut gif = b"GIF89a\x08\x00\x08\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff".to_vec();
        gif.extend(b"\x2c\x00\x00\x00\x00\x08\x00\x08\x00\x00\x02");
        gif.push(packed.len() as u8);
        gif.extend(&packed);
        gif.extend([0, 0x3b]);
        assert_eq!(cells("gif", &gif).unwrap(), Some(DIAGONAL));
        // A 10x10 grayscale bitmap contributes its top-left 8x8 pixels.
        let mut gray = Vec::new();
        for row in 0..10 {
            let cell = DIAGONAL.get(row).copied().unwrap_or(0xff);
            gray.extend([0, cell, 0xc0]);
        }
        assert_eq!(
            cells("png", &stored_png(10, 1, 0, &[], &gray)).unwrap(),
            Some(DIAGONAL)
        );
    }
}
