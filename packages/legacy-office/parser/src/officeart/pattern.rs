//! msofillPattern bitmaps as ECMA-376 preset patterns (MS-ODRAW 2.3.7.1 and
//! 2.4.11; ECMA-376 20.1.10.51 ST_PresetPatternVal).
//!
//! MS-ODRAW stores a pattern fill as a two-colour bitmap BLIP (fillBlip),
//! fillColor and fillBackColor; it names no pattern. The DrawingML pattern
//! fill names one of 54 presets and has no bitmap. Office writes the same
//! 8x8 cells for both: the bitmap Word 16 stores for a dark downward diagonal
//! fill and the bitmaps PowerPoint 16 stores for 30% and light upward
//! diagonal fills equal the cells its PDF export paints for the DrawingML
//! presets `dkDnDiag`, `pct30` and `ltUpDiag`, with white bitmap pixels in the
//! foreground colour (fillColor) and black pixels in the background colour
//! (fillBackColor). A bitmap is therefore projected as the preset with the
//! same cells.
//!
//! `CELLS` repeats the cell artwork the shared renderer paints for each
//! preset (measured from PowerPoint's PDF export; a unit test compares the two
//! tables). The XLS reader looks its pattern bitmaps up in the same table. The three presets with half-covered edge cells have no two-colour
//! equivalent there: their bitmap takes the nearest preset like any other
//! bitmap without an exact match.

/// `#` foreground, `.` background, `+` half covered; rows separated by `/`.
pub(crate) const CELLS: [(&str, &str); 54] = [
    (
        "pct5",
        "#......./......../......../......../....#.../......../......../........",
    ),
    (
        "pct10",
        "#......./......../....#.../......../#......./......../....#.../........",
    ),
    (
        "pct20",
        "#...#.../......../..#...#./......../#...#.../......../..#...#./........",
    ),
    (
        "pct25",
        "#...#.../..#...#./#...#.../..#...#./#...#.../..#...#./#...#.../..#...#.",
    ),
    (
        "pct30",
        "#.#.#.#./.#...#../#.#.#.#./...#...#/#.#.#.#./.#...#../#.#.#.#./...#...#",
    ),
    (
        "pct40",
        "#.#.#.#./.#.#.#.#/#.#.#.#./.#.#...#/#.#.#.#./.#.#.#.#/#.#.#.#./...#.#.#",
    ),
    (
        "pct50",
        "#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#",
    ),
    (
        "pct60",
        "###.###./.#.#.#.#/#.###.##/.#.#.#.#/###.###./.#.#.#.#/#.###.##/.#.#.#.#",
    ),
    (
        "pct70",
        ".###.###/##.###.#/.###.###/##.###.#/.###.###/##.###.#/.###.###/##.###.#",
    ),
    (
        "pct75",
        ".###.###/########/##.###.#/########/.###.###/########/##.###.#/########",
    ),
    (
        "pct80",
        "###.####/########/#######./########/###.####/########/#######./########",
    ),
    (
        "pct90",
        "########/########/########/####.###/########/########/########/.#######",
    ),
    (
        "horz",
        "########/......../......../......../......../......../......../........",
    ),
    (
        "vert",
        "#......./#......./#......./#......./#......./#......./#......./#.......",
    ),
    (
        "ltHorz",
        "########/......../......../......../########/......../......../........",
    ),
    (
        "ltVert",
        "#...#.../#...#.../#...#.../#...#.../#...#.../#...#.../#...#.../#...#...",
    ),
    (
        "dkHorz",
        "########/########/......../......../########/########/......../........",
    ),
    (
        "dkVert",
        "##..##../##..##../##..##../##..##../##..##../##..##../##..##../##..##..",
    ),
    (
        "narHorz",
        "########/......../########/......../########/......../########/........",
    ),
    (
        "narVert",
        ".#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#",
    ),
    (
        "dashHorz",
        "####..../......../......../......../....####/......../......../........",
    ),
    (
        "dashVert",
        "#......./#......./#......./#......./....#.../....#.../....#.../....#...",
    ),
    (
        "cross",
        "########/#......./#......./#......./#......./#......./#......./#.......",
    ),
    (
        "dnDiag",
        "#+.....+/+#+...../.+#+..../..+#+.../...+#+../....+#+./.....+#+/+.....+#",
    ),
    (
        "upDiag",
        "+.....+#/.....+#+/....+#+./...+#+../..+#+.../.+#+..../+#+...../#+.....+",
    ),
    (
        "ltDnDiag",
        "#...#.../.#...#../..#...#./...#...#/#...#.../.#...#../..#...#./...#...#",
    ),
    (
        "ltUpDiag",
        "...#...#/..#...#./.#...#../#...#.../...#...#/..#...#./.#...#../#...#...",
    ),
    (
        "dkDnDiag",
        "##..##../.##..##./..##..##/#..##..#/##..##../.##..##./..##..##/#..##..#",
    ),
    (
        "dkUpDiag",
        "..##..##/.##..##./##..##../#..##..#/..##..##/.##..##./##..##../#..##..#",
    ),
    (
        "wdDnDiag",
        "##.....#/###...../.###..../..###.../...###../....###./.....###/#.....##",
    ),
    (
        "wdUpDiag",
        "#.....##/.....###/....###./...###../..###.../.###..../###...../##.....#",
    ),
    (
        "dashDnDiag",
        "......../......../#...#.../.#...#../..#...#./...#...#/......../........",
    ),
    (
        "dashUpDiag",
        "......../......../...#...#/..#...#./.#...#../#...#.../......../........",
    ),
    (
        "diagCross",
        "#+....+#/+#+..+#+/.+#++#+./..+##+../..+##+../.+#++#+./+#+..+#+/#+....+#",
    ),
    (
        "smCheck",
        "#..##..#/.##..##./.##..##./#..##..#/#..##..#/.##..##./.##..##./#..##..#",
    ),
    (
        "lgCheck",
        "####..../####..../####..../####..../....####/....####/....####/....####",
    ),
    (
        "smGrid",
        "########/#...#.../#...#.../#...#.../########/#...#.../#...#.../#...#...",
    ),
    (
        "lgGrid",
        "########/#......./#......./#......./#......./#......./#......./#.......",
    ),
    (
        "dotGrid",
        "#.#.#.#./......../#......./......../#......./......../#......./........",
    ),
    (
        "smConfetti",
        "#......./....#.../.#....../......#./...#..../.......#/..#...../.....#..",
    ),
    (
        "lgConfetti",
        "#.##...#/..##..../......##/...##.##/##.##.../##....../....##../#...##.#",
    ),
    (
        "horzBrick",
        "########/#......./#......./#......./########/....#.../....#.../....#...",
    ),
    (
        "diagBrick",
        ".......#/......#./.....#../....#.../...##.../..#..#../.#....#./#......#",
    ),
    (
        "solidDmnd",
        "...#..../..###.../.#####../#######./.#####../..###.../...#..../........",
    ),
    (
        "openDmnd",
        "#.....#./.#...#../..#.#.../...#..../..#.#.../.#...#../#.....#./.......#",
    ),
    (
        "dotDmnd",
        "#......./......../..#...#./......../....#.../......../..#...#./........",
    ),
    (
        "plaid",
        "#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/####..../####..../####..../####....",
    ),
    (
        "sphere",
        ".###.###/#...#..#/#...####/#...####/.###.###/#..##.../#####.../#####...",
    ),
    (
        "weave",
        "#...#.../.#.#.#../..#...#./.#...#.#/#...#.../...#.#../..#...#./.#.#...#",
    ),
    (
        "divot",
        "......../...#..../....#.../...#..../......../#......./.......#/#.......",
    ),
    (
        "shingle",
        "......##/#....#../.#..#.../..##..../....##../......#./.......#/.......#",
    ),
    (
        "wave",
        "......../...##.../..#..#.#/##....../......../...##.../..#..#.#/##......",
    ),
    (
        "trellis",
        "########/.##..##./########/#..##..#/########/.##..##./########/#..##..#",
    ),
    (
        "zigZag",
        "#......#/.#....#./..#..#../...##.../#......#/.#....#./..#..#../...##...",
    ),
];

/// The foreground cells of one `CELLS` artwork, one byte per row with the
/// leftmost cell in bit 7. Half-covered (`+`) cells count as background.
pub(crate) fn artwork_cells(artwork: &str) -> [u8; 8] {
    let mut cells = [0u8; 8];
    for (row, line) in cells.iter_mut().zip(artwork.split('/')) {
        for (column, cell) in line.bytes().enumerate() {
            *row |= u8::from(cell == b'#') << (7 - column);
        }
    }
    cells
}

/// The preset whose cells equal `rows` (bit 7 of each row is the left cell,
/// a set bit the foreground). A half-covered preset cell is the anti-aliased
/// edge of a one-cell diagonal and counts as background, which no stored
/// bitmap has confirmed yet. A bitmap that equals no preset keeps its tone:
/// it takes the percentage preset with the nearest share of foreground cells.
#[cfg(any(test, feature = "direct-doc"))]
pub(crate) fn preset(rows: [u8; 8]) -> &'static str {
    const SITE: &str = "pattern fill bitmap without an equal preset pattern";
    if let Some((name, artwork)) = CELLS
        .iter()
        .find(|(_, artwork)| artwork_cells(artwork) == rows)
    {
        if artwork.contains('+') {
            crate::rejection::approximated(SITE, ());
        }
        return name;
    }
    let coverage = |rows: [u8; 8]| rows.iter().map(|row| row.count_ones()).sum::<u32>();
    let wanted = coverage(rows);
    crate::rejection::approximated(
        SITE,
        CELLS
            .iter()
            .filter(|(name, _)| name.starts_with("pct"))
            .min_by_key(|(_, artwork)| coverage(artwork_cells(artwork)).abs_diff(wanted))
            .map_or("pct50", |(name, _)| name),
    )
}

/// The cells of a pattern bitmap: the top-left 8x8 pixels of an 8x8 PNG (as
/// Word stores it) or a 10x10 PNG (PowerPoint tiles only that area of the
/// 10x10 bitmap it stores), one bit per pixel, set for white. `None` when
/// the image is anything else or has a pixel that is not opaque black or
/// white: such a BLIP is not a two-colour pattern.
#[cfg(any(test, feature = "direct-doc"))]
pub(crate) fn cells(png: &[u8]) -> Option<[u8; 8]> {
    // A two-colour 10x10 bitmap is far below this in every PNG encoding.
    if png.len() > 4096 || !png.starts_with(b"\x89PNG\r\n\x1a\n") {
        return None;
    }
    let mut header = None;
    let mut palette: &[u8] = &[];
    let mut transparency: &[u8] = &[];
    let mut data = Vec::new();
    let mut at = 8;
    while let Some(head) = png.get(at..at + 8) {
        let length = u32::from_be_bytes(head[..4].try_into().ok()?) as usize;
        let body = png.get(at + 8..(at + 8).checked_add(length)?)?;
        match &head[4..] {
            b"IHDR" => header = Some(body),
            b"PLTE" => palette = body,
            b"tRNS" => transparency = body,
            b"IDAT" => data.extend_from_slice(body),
            b"IEND" => break,
            _ => {}
        }
        at += length.checked_add(12)?;
    }
    let header = header.filter(|header| header.len() == 13)?;
    let size = u32::from_be_bytes(header[..4].try_into().ok()?) as usize;
    if !matches!(size, 8 | 10) || header[4..8] != header[..4] || header[10..] != [0, 0, 0] {
        return None;
    }
    let (depth, kind) = (usize::from(header[8]), header[9]);
    let channels = match (kind, depth) {
        (0 | 3, 1 | 2 | 4 | 8) => 1,
        (2, 8) => 3,
        (6, 8) => 4,
        _ => return None,
    };
    let stride = (size * channels * depth).div_ceil(8);
    let step = (channels * depth).div_ceil(8);
    let expected = size * (stride + 1);
    let raw = super::inflate::inflate_exact(&data, expected, data.len(), expected).ok()?;
    // PNG filtering (ISO/IEC 15948 clause 9) over whole bytes.
    let mut image = vec![0u8; size * stride];
    for y in 0..size {
        let filter = raw[y * (stride + 1)];
        for x in 0..stride {
            let left = if x >= step {
                image[y * stride + x - step]
            } else {
                0
            };
            let up = if y > 0 {
                image[(y - 1) * stride + x]
            } else {
                0
            };
            let corner = if x >= step && y > 0 {
                image[(y - 1) * stride + x - step]
            } else {
                0
            };
            let predicted = match filter {
                0 => 0,
                1 => left,
                2 => up,
                3 => ((u16::from(left) + u16::from(up)) / 2) as u8,
                4 => {
                    let estimate = i16::from(left) + i16::from(up) - i16::from(corner);
                    let distance = |value: u8| (estimate - i16::from(value)).abs();
                    if distance(left) <= distance(up) && distance(left) <= distance(corner) {
                        left
                    } else if distance(up) <= distance(corner) {
                        up
                    } else {
                        corner
                    }
                }
                _ => return None,
            };
            image[y * stride + x] = raw[y * (stride + 1) + 1 + x].wrapping_add(predicted);
        }
    }
    let mut rows = [0u8; 8];
    for (y, row) in rows.iter_mut().enumerate() {
        let line = &image[y * stride..][..stride];
        for x in 0..8 {
            let white = if channels == 1 {
                let bit = x * depth;
                let sample =
                    usize::from(line[bit / 8] >> (8 - depth - bit % 8)) & ((1 << depth) - 1);
                if kind == 3 {
                    if transparency.get(sample).is_some_and(|alpha| *alpha != 255) {
                        return None;
                    }
                    match palette.get(sample * 3..sample * 3 + 3)? {
                        [255, 255, 255] => true,
                        [0, 0, 0] => false,
                        _ => return None,
                    }
                } else if !transparency.is_empty() {
                    return None;
                } else if sample == (1 << depth) - 1 {
                    true
                } else if sample == 0 {
                    false
                } else {
                    return None;
                }
            } else {
                let pixel = &line[x * channels..][..channels];
                if (channels == 4 && pixel[3] != 255) || !transparency.is_empty() {
                    return None;
                }
                match pixel[..3] {
                    [255, 255, 255] => true,
                    [0, 0, 0] => false,
                    _ => return None,
                }
            };
            if white {
                *row |= 0x80 >> x;
            }
        }
    }
    Some(rows)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
        // The reader does not verify chunk checksums.
        [&(body.len() as u32).to_be_bytes()[..], kind, body, &[0; 4]].concat()
    }

    fn png(size: u32, depth: u8, kind: u8, extra: &[Vec<u8>], scanlines: &[u8]) -> Vec<u8> {
        let mut encoder = flate2::Compress::new(flate2::Compression::default(), true);
        let mut data = vec![0; 256];
        encoder
            .compress(scanlines, &mut data, flate2::FlushCompress::Finish)
            .unwrap();
        data.truncate(encoder.total_out() as usize);
        let header = [
            &size.to_be_bytes()[..],
            &size.to_be_bytes(),
            &[depth, kind, 0, 0, 0],
        ]
        .concat();
        [
            b"\x89PNG\r\n\x1a\n".to_vec(),
            chunk(b"IHDR", &header),
            extra.concat(),
            chunk(b"IDAT", &data),
            chunk(b"IEND", &[]),
        ]
        .concat()
    }

    /// The 2-bit palette bitmap Word 16 stores for a dark downward diagonal
    /// pattern: index 0 white, index 1 black, index 2 transparent and unused.
    fn dark_downward_diagonal() -> Vec<u8> {
        let rows: [u16; 4] = [0x0505, 0x4141, 0x5050, 0x1414];
        let scanlines: Vec<u8> = (0..8)
            .flat_map(|y| {
                let [high, low] = rows[y % 4].to_be_bytes();
                [0, high, low]
            })
            .collect();
        png(
            8,
            2,
            3,
            &[
                chunk(b"PLTE", &[255, 255, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
                chunk(b"tRNS", &[255, 255, 0]),
            ],
            &scanlines,
        )
    }

    #[test]
    fn a_stored_pattern_bitmap_is_the_preset_with_the_same_cells() {
        let rows = cells(&dark_downward_diagonal()).unwrap();
        assert_eq!(rows, [0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66, 0x33, 0x99]);
        assert_eq!(preset(rows), "dkDnDiag");
        assert!(crate::rejection::take_best_effort().is_empty());
        // Every two-colour preset is its own bitmap (`cross` and `lgGrid`
        // share one artwork; the first name stands for both).
        for (name, artwork) in CELLS {
            if artwork.contains('+') || name == "lgGrid" {
                continue;
            }
            let mut rows = [0u8; 8];
            for (row, line) in rows.iter_mut().zip(artwork.split('/')) {
                for (column, cell) in line.bytes().enumerate() {
                    *row |= u8::from(cell == b'#') << (7 - column);
                }
            }
            assert_eq!(preset(rows), name);
        }
        assert!(crate::rejection::take_best_effort().is_empty());
    }

    #[test]
    fn a_bitmap_without_an_equal_preset_is_approximated() {
        let site = (
            crate::rejection::BestEffort::Approximated,
            "pattern fill bitmap without an equal preset pattern",
        );
        // A one-cell downward diagonal is `dnDiag` without its half-covered
        // edge cells.
        let rows = [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01];
        assert_eq!(preset(rows), "dnDiag");
        assert_eq!(crate::rejection::take_best_effort(), [site]);
        // An unknown bitmap keeps its tone: 56 of 64 cells are foreground.
        assert_eq!(
            preset([0xff, 0xff, 0xff, 0x0f, 0xf0, 0xff, 0xff, 0xff]),
            "pct75"
        );
        assert_eq!(preset([0; 8]), "pct5");
        assert_eq!(crate::rejection::take_best_effort(), [site, site]);
    }

    #[test]
    fn only_opaque_black_and_white_bitmaps_of_a_pattern_size_are_patterns() {
        // Filtered greyscale rows decode like unfiltered ones.
        let grey: Vec<u8> = (0..8)
            .flat_map(|y| if y == 0 { [0, 0xf0] } else { [2, 0] })
            .collect();
        assert_eq!(cells(&png(8, 1, 0, &[], &grey)), Some([0xf0; 8]));
        // A 10x10 bitmap contributes its top-left 8x8 pixels.
        let ten: Vec<u8> = (0..10).flat_map(|_| [0, 0xaa, 0xc0]).collect();
        assert_eq!(cells(&png(10, 1, 0, &[], &ten)), Some([0xaa; 8]));
        // A colour, a transparent pixel or another size is not a pattern.
        let truecolor: Vec<u8> = (0..8)
            .flat_map(|_| [vec![0], [255, 255, 255].repeat(7), vec![255, 0, 0]].concat())
            .collect();
        assert_eq!(cells(&png(8, 8, 2, &[], &truecolor)), None);
        let transparent = png(
            8,
            1,
            3,
            &[
                chunk(b"PLTE", &[255, 255, 255, 0, 0, 0]),
                chunk(b"tRNS", &[0]),
            ],
            &[0; 16],
        );
        assert_eq!(cells(&transparent), None);
        let sixteen: Vec<u8> = (0..16).flat_map(|_| [0, 0, 0]).collect();
        assert_eq!(cells(&png(16, 1, 0, &[], &sixteen)), None);
        assert_eq!(cells(b"GIF89a"), None);
    }

    /// `CELLS` is the artwork the shared renderer paints for the presets.
    #[test]
    fn the_cells_are_those_of_the_shared_pattern_renderer() {
        let source = include_str!("../../../../core/src/shape/pattern-bitmaps.ts");
        let painted: Vec<(&str, &str)> = source
            .lines()
            .filter_map(|line| {
                let (name, rest) = line.trim().split_once(": '")?;
                Some((name, rest.strip_suffix("',")?))
            })
            .collect();
        assert_eq!(painted, CELLS);
    }
}
