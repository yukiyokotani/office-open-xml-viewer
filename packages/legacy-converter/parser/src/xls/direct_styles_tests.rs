//! BIFF8 FONT/XF records projected through the direct session bootstrap into
//! the XLSX style model, with expectations taken from [MS-XLS] 2.4.122 Font,
//! 2.4.353 XF and 2.5.20 CellXF and ECMA-376 Part 1 18.8.

use super::direct::DirectSession;
use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};

fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
    [
        kind.to_le_bytes().as_slice(),
        &(payload.len() as u16).to_le_bytes(),
        payload,
    ]
    .concat()
}

fn font(weight: u16) -> Vec<u8> {
    let mut font = vec![0_u8; 16];
    font[..2].copy_from_slice(&220_u16.to_le_bytes());
    font[4..6].copy_from_slice(&0x7fff_u16.to_le_bytes());
    font[6..8].copy_from_slice(&weight.to_le_bytes());
    font[14] = 1;
    font.push(b'F');
    font
}

fn bootstrap_styles(fonts: &[Vec<u8>], xfs: &[[u8; 20]]) -> Result<xlsx_model::Styles, String> {
    let mut stream = record(super::BOF, &[0, 6, 5, 0]);
    let bound = stream.len() + 4;
    stream.extend(record(super::BOUNDSHEET8, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
    for font in fonts {
        stream.extend(record(0x31, font));
    }
    for xf in xfs {
        stream.extend(record(0xe0, xf));
    }
    stream.extend(record(super::EOF, &[]));
    let offset = stream.len() as u32;
    stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
    stream.extend(record(super::BOF, &[0, 6, 0x10, 0]));
    stream.extend(record(super::EOF, &[]));
    let bytes = build_scoped_cfb(&[("Workbook", stream)]);
    let cfb = CompoundFile::open(&bytes).unwrap();
    Ok(DirectSession::new(&cfb)?.bootstrap()?.styles)
}

#[test]
fn empty_and_font_only_styles_project_the_default_cell_format() {
    let empty = bootstrap_styles(&[], &[]).unwrap();
    assert_eq!(empty.fonts.len(), 1);
    assert_eq!(empty.fonts[0].name.as_deref(), Some("Calibri"));
    assert_eq!((empty.fonts[0].size, empty.fonts[0].charset), (11.0, None));

    let font_only = bootstrap_styles(&[font(400)], &[]).unwrap();
    assert_eq!(font_only.fonts.len(), 1);
    assert_eq!(font_only.fonts[0].name.as_deref(), Some("F"));
    // An authored FONT record carries its bCharSet.
    assert_eq!(
        (font_only.fonts[0].size, font_only.fonts[0].charset),
        (11.0, Some(0))
    );
    for styles in [&empty, &font_only] {
        let fills: Vec<_> = styles
            .fills
            .iter()
            .map(|f| f.pattern_type.as_str())
            .collect();
        assert_eq!(fills, ["none", "gray125"]);
        assert_eq!(styles.borders.len(), 1);
        assert!(styles.borders[0].left.is_none() && styles.borders[0].diagonal_up.is_none());
        assert_eq!(styles.cell_xfs.len(), 1);
        let xf = &styles.cell_xfs[0];
        assert_eq!(
            (xf.font_id, xf.fill_id, xf.border_id, xf.num_fmt_id),
            (0, 0, 0, 0)
        );
        assert_eq!((xf.align_h.as_deref(), xf.align_v.as_deref()), (None, None));
        assert!(!xf.wrap_text && !xf.shrink_to_fit);
        assert_eq!(
            (xf.indent, xf.text_rotation, xf.reading_order),
            (None, None, None)
        );
    }
}

#[test]
fn xf_alignment_matrix_projects_every_defined_value() {
    let horizontal = [
        None,
        Some("left"),
        Some("center"),
        Some("right"),
        Some("fill"),
        Some("justify"),
        Some("centerContinuous"),
        Some("distributed"),
    ];
    let vertical = ["top", "center", "bottom", "justify", "distributed"];
    let mut xfs = Vec::new();
    for (index, rotation, reading) in [
        (0_u8, 0_u8, 0_u8),
        (1, 1, 1),
        (2, 90, 2),
        (3, 91, 0),
        (4, 180, 1),
        (5, 255, 2),
        (6, 45, 0),
        (7, 135, 1),
    ] {
        let mut xf = [0_u8; 20];
        let wrap = index & 1;
        xf[6] = index | (wrap << 3) | ((index % 5) << 4);
        xf[7] = rotation;
        xf[8] = (reading << 6) | index | ((index & 2) << 3);
        xfs.push(xf);
    }
    let styles = bootstrap_styles(&[font(400), font(650), font(700)], &xfs).unwrap();
    // Only weight 700 is bold; 650 is not rounded up.
    let bold: Vec<_> = styles.fonts.iter().map(|f| f.bold).collect();
    assert_eq!(bold, [false, false, true]);
    for (index, xf) in styles.cell_xfs.iter().enumerate() {
        let source = &xfs[index];
        let value = u32::from(source[8] & 15);
        assert_eq!(xf.align_h.as_deref(), horizontal[index], "xf {index}");
        assert_eq!(
            xf.align_v.as_deref(),
            Some(vertical[index % 5]),
            "xf {index}"
        );
        assert_eq!(xf.wrap_text, index & 1 != 0, "xf {index}");
        assert_eq!(xf.shrink_to_fit, index & 2 != 0, "xf {index}");
        assert_eq!(xf.indent, (value != 0).then_some(value), "xf {index}");
        let rotation = u32::from(source[7]);
        assert_eq!(xf.text_rotation, (rotation != 0).then_some(rotation));
        let reading = u32::from(source[8] >> 6);
        assert_eq!(xf.reading_order, (reading != 0).then_some(reading));
    }
    for (offset, value) in [(6, 5 << 4), (7, 181), (7, 254), (8, 3 << 6)] {
        let mut invalid = [0_u8; 20];
        invalid[offset] = value;
        assert!(bootstrap_styles(&[font(400)], &[invalid]).is_err());
    }
}

#[test]
fn xf_borders_and_pattern_fills_resolve_palette_colors() {
    let mut xf = [0_u8; 20];
    // Left thick red, right slantDashDot blue, top hair, bottom medium,
    // both diagonals thin black.
    let b1: u32 = 5 | (13 << 4) | (7 << 8) | (2 << 12) | (10 << 16) | (12 << 23) | (3 << 30);
    let b2 = (8_u32 << 14) | (1 << 21) | (17 << 26);
    xf[10..14].copy_from_slice(&b1.to_le_bytes());
    xf[14..18].copy_from_slice(&b2.to_le_bytes());
    xf[18..20].copy_from_slice(&(10_u16 | (9 << 7)).to_le_bytes());
    let styles = bootstrap_styles(&[font(400)], &[xf]).unwrap();
    let xf = &styles.cell_xfs[0];
    let border = &styles.borders[xf.border_id as usize];
    let edge = |edge: &Option<xlsx_model::BorderEdge>| {
        edge.as_ref()
            .map(|e| (e.style.clone(), e.color.clone().unwrap_or_default()))
    };
    assert_eq!(edge(&border.left), Some(("thick".into(), "#FF0000".into())));
    assert_eq!(
        edge(&border.right),
        Some(("slantDashDot".into(), "#0000FF".into()))
    );
    assert_eq!(edge(&border.top).unwrap().0, "hair");
    assert_eq!(edge(&border.bottom).unwrap().0, "medium");
    let diagonal = Some(("thin".into(), "#000000".into()));
    assert_eq!(edge(&border.diagonal_up), diagonal);
    assert_eq!(edge(&border.diagonal_down), diagonal);
    // A cell pattern gray125 is its own fill, distinct from the reserved
    // second fill.
    assert_ne!(xf.fill_id, 1);
    let fill = &styles.fills[xf.fill_id as usize];
    assert_eq!(fill.pattern_type, "gray125");
    assert_eq!(fill.fg_color.as_deref(), Some("#FF0000"));
    assert_eq!(fill.bg_color.as_deref(), Some("#FFFFFF"));
}
