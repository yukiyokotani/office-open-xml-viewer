//! OfficeArt pattern fills (MS-ODRAW 2.3.7.1 msofillPattern) as DrawingML
//! preset patterns (ECMA-376 20.1.10.51 ST_PresetPatternVal).
//!
//! A pattern fill paints the fillBlip, an 8 x 8 two-tone bitmap, with the
//! fill color as its foreground and the fill back color as its background.
//! The model names a pattern by its DrawingML preset, so the bitmap is read
//! and looked up in the preset cell artwork.

// The 8 x 8 cells of each preset, row by row (`#` foreground, `.`
// background, `+` an edge cell the antialiased diagonals draw at the
// midpoint): the table of the shared shape renderer
// (`packages/core/src/shape/pattern-bitmaps.ts`, measured from Office's own
// PDF output of every preset), which the DOC reader also reads.
use crate::officeart::pattern::{artwork_cells, CELLS as PRESETS};

/// Foreground cells of an 8 x 8 pattern bitmap, one byte per row with the
/// leftmost cell in the high bit, or `None` when `png` is not such a bitmap.
/// The darker of its two tones is the foreground, as in the bitmaps Office
/// stores (black cells on white).
pub(super) fn cells(png: &[u8]) -> Option<[u8; 8]> {
    if !png.starts_with(b"\x89PNG\r\n\x1a\n") {
        return None;
    }
    let mut position = 8;
    let mut header = None;
    let mut palette: &[u8] = &[];
    let mut data = Vec::new();
    // A pattern bitmap is a few hundred bytes; chunks are read within that.
    for _ in 0..64 {
        let length = u32::from_be_bytes(png.get(position..position + 4)?.try_into().ok()?);
        let length = usize::try_from(length).ok()?;
        let kind = png.get(position + 4..position + 8)?;
        let body = png.get(position + 8..position.checked_add(8)?.checked_add(length)?)?;
        match kind {
            b"IHDR" => header = Some(body),
            b"PLTE" => palette = body,
            b"IDAT" => {
                if data.len() + body.len() > 4096 {
                    return None;
                }
                data.extend_from_slice(body);
            }
            b"IEND" => break,
            _ => {}
        }
        position += 12 + length;
    }
    let header = header.filter(|header| header.len() == 13)?;
    let dimension = |at: usize| u32::from_be_bytes(header[at..at + 4].try_into().unwrap());
    let (depth, color, interlace) = (header[8], header[9], header[12]);
    if dimension(0) != 8 || dimension(4) != 8 || interlace != 0 {
        return None;
    }
    // Grayscale or palette samples of one channel.
    if !matches!(color, 0 | 3) || !matches!(depth, 1 | 2 | 4 | 8) {
        return None;
    }
    let row_bytes = usize::from(depth);
    let mut raw = [0u8; 8 * 9];
    let size = 8 * (1 + row_bytes);
    let mut inflater = flate2::Decompress::new(true);
    let status = inflater
        .decompress(&data, &mut raw[..size], flate2::FlushDecompress::Finish)
        .ok()?;
    if inflater.total_out() != size as u64
        || !matches!(status, flate2::Status::StreamEnd | flate2::Status::Ok)
    {
        return None;
    }
    // Luminance of a sample: the palette entry, or the gray level.
    let luminance = |sample: u8| -> Option<u32> {
        if color == 3 {
            let entry = palette.get(usize::from(sample) * 3..usize::from(sample) * 3 + 3)?;
            Some(u32::from(entry[0]) + u32::from(entry[1]) + u32::from(entry[2]))
        } else {
            Some(u32::from(sample) * 3 * 255 / ((1u32 << depth) - 1))
        }
    };
    let mut output = [0u8; 8];
    let mut previous = [0u8; 8];
    for row in 0..8 {
        let line = &raw[row * (1 + row_bytes)..(row + 1) * (1 + row_bytes)];
        let filter = line[0];
        let mut current = [0u8; 8];
        // PNG filters work on bytes; a sample is at most one byte here.
        for index in 0..row_bytes {
            let left = if index == 0 { 0 } else { current[index - 1] };
            let up = previous[index];
            let upper_left = if index == 0 { 0 } else { previous[index - 1] };
            let predicted = match filter {
                0 => 0,
                1 => left,
                2 => up,
                3 => ((u16::from(left) + u16::from(up)) / 2) as u8,
                4 => {
                    let estimate = i16::from(left) + i16::from(up) - i16::from(upper_left);
                    let distance = |value: u8| (estimate - i16::from(value)).abs();
                    if distance(left) <= distance(up) && distance(left) <= distance(upper_left) {
                        left
                    } else if distance(up) <= distance(upper_left) {
                        up
                    } else {
                        upper_left
                    }
                }
                _ => return None,
            };
            current[index] = line[1 + index].wrapping_add(predicted);
        }
        for column in 0..8usize {
            let bit = column * usize::from(depth);
            let sample = (current[bit / 8] >> (8 - usize::from(depth) - bit % 8))
                & ((1u16 << depth) - 1) as u8;
            if luminance(sample)? < 3 * 128 {
                output[row] |= 0x80 >> column;
            }
        }
        previous = current;
    }
    Some(output)
}

/// The preset whose cells are `cells` at some offset of the 8 x 8 tile. A
/// pattern repeats, so the bitmap's phase does not name another pattern.
pub(super) fn preset(cells: [u8; 8]) -> Option<&'static str> {
    for (name, art) in PRESETS {
        // The midpoint edge cells of the antialiased diagonals are
        // background in a two-tone bitmap (as foreground they would be the
        // wide diagonals' cells).
        let preset = artwork_cells(art);
        for down in 0..8 {
            for right in 0..8u32 {
                if (0..8).all(|row| cells[(row + down) % 8].rotate_right(right) == preset[row]) {
                    return Some(name);
                }
            }
        }
    }
    None
}

/// The percentage preset nearest to the share of foreground cells, for a
/// bitmap that is no preset's artwork.
pub(super) fn nearest_tone(cells: [u8; 8]) -> &'static str {
    let foreground: u32 = cells.iter().map(|row| row.count_ones()).sum();
    let percent = foreground * 100 / 64;
    [
        (5, "pct5"),
        (10, "pct10"),
        (20, "pct20"),
        (25, "pct25"),
        (30, "pct30"),
        (40, "pct40"),
        (50, "pct50"),
        (60, "pct60"),
        (70, "pct70"),
        (75, "pct75"),
        (80, "pct80"),
        (90, "pct90"),
    ]
    .into_iter()
    .min_by_key(|(tone, _)| percent.abs_diff(*tone))
    .map_or("pct50", |(_, name)| name)
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    /// An 8 x 8 palette PNG of depth 2 (white, black) from rows of two-bit
    /// samples, as Excel stores a pattern bitmap.
    pub(crate) fn png(rows: [[u8; 2]; 8]) -> Vec<u8> {
        fn chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
            let mut chunk = (body.len() as u32).to_be_bytes().to_vec();
            chunk.extend(kind);
            chunk.extend(body);
            chunk.extend([0u8; 4]);
            chunk
        }
        let mut raw = Vec::new();
        for row in rows {
            raw.push(0);
            raw.extend(row);
        }
        let mut deflater = flate2::Compress::new(flate2::Compression::default(), true);
        let mut data = vec![0u8; 128];
        deflater
            .compress(&raw, &mut data, flate2::FlushCompress::Finish)
            .unwrap();
        data.truncate(deflater.total_out() as usize);
        let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
        png.extend(chunk(b"IHDR", &[0, 0, 0, 8, 0, 0, 0, 8, 2, 3, 0, 0, 0]));
        png.extend(chunk(b"PLTE", &[255, 255, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
        png.extend(chunk(b"IDAT", &data));
        png.extend(chunk(b"IEND", &[]));
        png
    }

    /// The bitmap Excel stores for its dark downward diagonal.
    pub(crate) const DARK_DOWN: [[u8; 2]; 8] = [
        [0x05, 0x05],
        [0x41, 0x41],
        [0x50, 0x50],
        [0x14, 0x14],
        [0x05, 0x05],
        [0x41, 0x41],
        [0x50, 0x50],
        [0x14, 0x14],
    ];

    #[test]
    fn pattern_bitmaps_name_their_preset_at_any_phase() {
        let cells = cells(&png(DARK_DOWN)).unwrap();
        assert_eq!(cells, [0x33, 0x99, 0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66]);
        assert_eq!(preset(cells), Some("dkDnDiag"));
        // Every preset's own artwork names it (or a preset with the same
        // cells, as `cross` and `lgGrid` have), also shifted.
        let rows = |art: &str| {
            let mut rows = [0u8; 8];
            for (row, line) in art.split('/').enumerate() {
                assert_eq!(line.len(), 8);
                for (column, cell) in line.bytes().enumerate() {
                    if cell == b'#' {
                        rows[row] |= 0x80 >> column;
                    }
                }
            }
            rows
        };
        let canonical = |rows: [u8; 8]| {
            (0..8)
                .flat_map(|down| (0..8u32).map(move |right| (down, right)))
                .map(|(down, right)| {
                    std::array::from_fn::<u8, 8, _>(|row| {
                        rows[(row + down) % 8].rotate_right(right)
                    })
                })
                .min()
                .unwrap()
        };
        for (name, art) in PRESETS {
            let own = rows(art);
            let found = preset(own).unwrap_or_else(|| panic!("{name} is not found"));
            let found_art = PRESETS
                .iter()
                .find(|(preset, _)| *preset == found)
                .unwrap()
                .1;
            assert_eq!(
                canonical(rows(found_art)),
                canonical(own),
                "{name} read as {found}"
            );
            let mut shifted = own;
            shifted.rotate_left(3);
            assert_eq!(preset(shifted.map(|row| row.rotate_left(5))), Some(found));
        }
        // A bitmap that is no preset keeps its tone.
        let odd = [0xff, 0x81, 0x81, 0x99, 0x99, 0x81, 0x81, 0xff];
        assert_eq!(preset(odd), None);
        assert_eq!(nearest_tone(odd), "pct50");
        assert_eq!(nearest_tone([0x80, 0, 0, 0, 0, 0, 0, 0]), "pct5");
        // Not an 8 x 8 bitmap, or not a PNG.
        let mut wide = png(DARK_DOWN);
        wide[19] = 16;
        assert!(super::cells(&wide).is_none());
        assert!(super::cells(b"BM").is_none());
    }
}
