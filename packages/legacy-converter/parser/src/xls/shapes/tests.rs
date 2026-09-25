use super::super::drawing_anchors::GroupFrame;
use super::*;

fn font(twips: u16, flags: u16, icv: u16, weight: u16, underline: u8, name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    for value in [twips, flags, icv, weight, 0] {
        data.extend_from_slice(&value.to_le_bytes());
    }
    data.extend_from_slice(&[underline, 0, 0, 0, name.len() as u8, 0]);
    data.extend_from_slice(name.as_bytes());
    data
}

fn txo(grbit: u16, characters: u16, run_bytes: u16) -> Vec<u8> {
    let mut data = vec![0; 18];
    data[..2].copy_from_slice(&grbit.to_le_bytes());
    data[10..12].copy_from_slice(&characters.to_le_bytes());
    data[12..14].copy_from_slice(&run_bytes.to_le_bytes());
    data
}

fn runs(runs: &[(u16, u16)], characters: u16) -> Vec<u8> {
    let mut data = Vec::new();
    for &(at, font) in runs.iter().chain([(characters, 0)].iter()) {
        data.extend_from_slice(&at.to_le_bytes());
        data.extend_from_slice(&font.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
    }
    data
}

fn compressed(text: &str) -> Vec<u8> {
    [vec![0], text.as_bytes().to_vec()].concat()
}

fn records(data: &[(u16, Vec<u8>)]) -> Vec<Record<'_>> {
    let mut offset = 0;
    data.iter()
        .map(|(kind, data)| {
            let record = Record {
                kind: *kind,
                offset,
                data,
            };
            offset += 4 + data.len();
            record
        })
        .collect()
}

/// Two fonts (FontIndex 0: 12 pt bold white Meiryo UI, 1: 25 pt black) and
/// one text box: TxO, its text and its runs.
fn workbook(
    grbit: u16,
    text: &[u8],
    characters: u16,
    text_runs: &[(u16, u16)],
) -> Vec<(u16, Vec<u8>)> {
    let run_data = runs(text_runs, characters);
    vec![
        (0x0031, font(240, 1, 9, 700, 0, "Meiryo UI")),
        (0x0031, font(500, 1, 8, 400, 0, "Meiryo UI")),
        (0x01b6, txo(grbit, characters, run_data.len() as u16)),
        (0x003c, text.to_vec()),
        (0x003c, run_data),
    ]
}

const TXO: usize = 2;

fn source(owned: &[(u16, Vec<u8>)], kind: u16, properties: &[(u16, u32)]) -> ShapeSource {
    ShapeSource {
        kind,
        properties: properties.to_vec(),
        complex: Vec::new(),
        text: Some(records(owned)[TXO].offset),
    }
}

fn leaf(flags: u32) -> Leaf {
    Leaf {
        flags,
        child: false,
        order: 3,
    }
}

fn project(
    owned: &[(u16, Vec<u8>)],
    defaults: &Table,
    shape: &ShapeSource,
    flags: u32,
) -> Result<Option<xlsx_model::ShapeInfo>, String> {
    let records = records(owned);
    let styles = styles::Styles::parse(&records).unwrap();
    leaf(flags)
        .project(&records, &styles, defaults, shape)
        .map(|leaf| leaf.map(|(info, _)| info))
}

/// Excel's own drawing defaults in every corpus workbook (MS-ODRAW 2.2.12):
/// fAutoTextMargin, fill scheme color 0x41 and line scheme color 0x40.
fn excel_defaults() -> Table {
    let mut table = Table::default();
    for (id, value) in [
        (0x00bf, 0x0008_0008),
        (0x0181, 0x0800_0041),
        (0x01c0, 0x0800_0040),
    ] {
        table.add(id, value).unwrap();
    }
    table
}

/// The sample rectangle: fill #1B79AD, no line, flipped horizontally, with
/// centred white 12 pt bold text (Excel's XLSX writes the same `sp`).
const RECTANGLE: [(u16, u32); 6] = [
    (0x0087, 1),
    (0x0181, 0x00ad_791b),
    (0x01bf, 0x0010_0010),
    (0x01cb, 12_700),
    (0x01ff, 0x0008_0000),
    (0x03bf, 0x0002_0000),
];

#[test]
fn rectangle_and_text_match_excels_xlsx_counterpart() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let shape = source(&owned, 1, &RECTANGLE);
    let info = project(&owned, &excel_defaults(), &shape, 0xa40)
        .unwrap()
        .unwrap();
    assert_eq!(info.z_order, 3);
    assert_eq!(info.fill_color.as_deref(), Some("#1B79AD"));
    assert!(info.stroke_color.is_none());
    assert_eq!(info.stroke_width, 0);
    assert!(
        matches!(&info.geom, xlsx_model::ShapeGeom::Preset { name, adj } if name == "rect" && adj.is_empty())
    );
    let text = info.text.unwrap();
    assert_eq!(
        (text.anchor.as_str(), text.wrap.as_str()),
        ("ctr", "square")
    );
    // fAutoTextMargin from the drawing defaults: DrawingML default insets.
    assert_eq!(
        [text.l_ins, text.t_ins, text.r_ins, text.b_ins],
        [91_440, 45_720, 91_440, 45_720]
    );
    assert_eq!(text.paragraphs.len(), 1);
    assert_eq!(text.paragraphs[0].align, "ctr");
    let xlsx_model::ShapeTextRun::Text {
        text,
        bold,
        italic,
        size,
        color,
        font_face,
        font_face_ea,
        ..
    } = &text.paragraphs[0].runs[0]
    else {
        panic!("text run expected");
    };
    assert_eq!(text, "2018");
    assert!(*bold && !*italic);
    assert_eq!(*size, 12.0);
    assert_eq!(color.as_deref(), Some("#FFFFFF"));
    assert_eq!(font_face.as_deref(), Some("Meiryo UI"));
    assert_eq!(font_face_ea.as_deref(), Some("Meiryo UI"));
}

#[test]
fn text_splits_paragraphs_at_line_feeds_and_runs_at_font_changes() {
    let text = "ab\n\ncd";
    let owned = workbook(0x0212, &compressed(text), 6, &[(0, 0), (1, 1), (5, 0)]);
    let shape = source(&owned, 202, &[(0x01bf, 0x0010_0000), (0x01ff, 0x0008_0000)]);
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert!(info.fill.is_none() && info.stroke_color.is_none());
    let text = info.text.unwrap();
    assert_eq!(text.anchor, "t");
    let words: Vec<Vec<(String, f64)>> = text
        .paragraphs
        .iter()
        .map(|paragraph| {
            assert_eq!(paragraph.align, "l");
            paragraph
                .runs
                .iter()
                .map(|run| match run {
                    xlsx_model::ShapeTextRun::Text { text, size, .. } => (text.clone(), *size),
                    _ => panic!("text run expected"),
                })
                .collect()
        })
        .collect();
    assert_eq!(
        words,
        vec![
            vec![("a".into(), 12.0), ("b".into(), 25.0)],
            // An empty line keeps the height of the font at its position.
            vec![(String::new(), 25.0)],
            vec![("c".into(), 25.0), ("d".into(), 12.0)],
        ]
    );
    // High-byte text fragments decode as UTF-16.
    let wide: Vec<u8> = [
        vec![1],
        "予算".encode_utf16().flat_map(u16::to_le_bytes).collect(),
    ]
    .concat();
    let owned = workbook(0x0212, &wide, 2, &[(0, 0)]);
    let shape = source(&owned, 202, &[(0x01bf, 0x0010_0000), (0x01ff, 0x0008_0000)]);
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    let xlsx_model::ShapeTextRun::Text { text, .. } = &info.text.unwrap().paragraphs[0].runs[0]
    else {
        panic!("text run expected");
    };
    assert_eq!(text, "予算");
}

#[test]
fn explicit_margins_apply_only_without_automatic_margins() {
    let owned = workbook(0x0212, &compressed("x"), 1, &[(0, 0)]);
    let explicit = [
        (0x0081, 0),
        (0x0082, 12_700),
        (0x00bf, 0x0008_0000),
        (0x01bf, 0x0010_0000),
        (0x01ff, 0x0008_0000),
    ];
    let shape = source(&owned, 202, &explicit);
    let text = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap()
        .text
        .unwrap();
    assert_eq!([text.l_ins, text.t_ins], [0, 12_700]);
    let mut automatic = explicit;
    automatic[2] = (0x00bf, 0x0008_0008);
    let shape = source(&owned, 202, &automatic);
    let text = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap()
        .text
        .unwrap();
    assert_eq!([text.l_ins, text.t_ins], [91_440, 45_720]);
}

#[test]
fn solid_lines_use_their_explicit_color_width_and_dash() {
    let owned = workbook(0x0212, &compressed("x"), 1, &[(0, 0)]);
    let shape = source(
        &owned,
        1,
        &[
            (0x0181, 0x0000_ff00),
            (0x01c0, 0x0000_00ff),
            (0x01cb, 19_050),
            (0x01ce, 6),
            (0x01ff, 0x0008_0008),
        ],
    );
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert_eq!(info.fill_color.as_deref(), Some("#00FF00"));
    assert_eq!(info.stroke_color.as_deref(), Some("#FF0000"));
    assert_eq!(info.stroke_width, 19_050);
    assert_eq!(info.stroke_dash_style.as_deref(), Some("dash"));
    assert_eq!(info.stroke_line_join.as_deref(), Some("round"));
    // Translucent paint: the XLSX counterpart's `a:alpha 50000` line is
    // saved as lineOpacity 0x8080, the same alpha byte.
    let shape = source(
        &owned,
        1,
        &[
            (0x0181, 0xff),
            (0x0182, 0x4000),
            (0x01c0, 0xccff),
            (0x01c1, 0x8080),
            (0x01ff, 0x0008_0008),
        ],
    );
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert_eq!(info.fill_color.as_deref(), Some("#FF000040"));
    assert_eq!(info.stroke_color.as_deref(), Some("#FFCC0080"));
}

#[test]
fn editing_only_shape_booleans_are_accepted() {
    // fLockShapeType (Excel writes it explicitly false) and its siblings.
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let mut properties = RECTANGLE.to_vec();
    properties.push((0x033f, 0x0008_0000));
    let shape = source(&owned, 1, &properties);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .is_some());
}

#[test]
fn hidden_shapes_are_not_drawn() {
    let owned = workbook(0x0212, &compressed("x"), 1, &[(0, 0)]);
    let mut properties = RECTANGLE.to_vec();
    properties[5] = (0x03bf, 0x0002_0002);
    let shape = source(&owned, 1, &properties);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .is_none());
}

#[test]
fn unevidenced_shape_facts_fail_closed() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let reject = |properties: &[(u16, u32)], kind: u16, flags: u32, expected: &str| {
        // Each case replaces the rectangle's own value of that property.
        let mut all = RECTANGLE.to_vec();
        all.retain(|(id, _)| properties.iter().all(|(other, _)| other != id));
        all.extend_from_slice(properties);
        let shape = source(&owned, kind, &all);
        let error = project(&owned, &excel_defaults(), &shape, flags)
            .err()
            .unwrap_or_else(|| panic!("{expected} accepted"));
        assert!(error.contains(expected), "{expected}: {error}");
    };
    reject(&[], 3, 0xa00, "shape type 3");
    reject(&[], 1, 0xa10, "shape flags");
    reject(&[], 1, 0x200, "shape flags");
    reject(&[(0x023f, 0x0002_0002)], 1, 0xa00, "shadows");
    reject(&[(0x0180, 1)], 1, 0xa00, "non-solid");
    reject(&[(0x0085, 1)], 1, 0xa00, "wrapping");
    reject(&[(0x0087, 0)], 1, 0xa00, "disagrees");
    reject(&[(0x008b, 1)], 1, 0xa00, "property 0x008b");
    reject(&[(0x0304, 0)], 1, 0xa00, "property 0x0304");
    reject(&[(0x033f, 0x0001_0001)], 1, 0xa00, "background");
    reject(&[(0x033f, 0x0020_0020)], 1, 0xa00, "OLE icon");
    reject(&[(0x033f, 0x0080_0000)], 1, 0xa00, "flip overrides");
    reject(&[(0xc105, 4)], 1, 0xa00, "complex property");
    reject(&[(0x4186, 1)], 1, 0xa00, "BLIP property");
    // Primary and tertiary tables that disagree have no precedence.
    let mut conflicting = RECTANGLE.to_vec();
    conflicting.push((0x0181, 0x0800_0041));
    let shape = source(&owned, 1, &conflicting);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("conflicting"));

    // A fill or line left to Excel's scheme-colored drawing defaults.
    let shape = source(&owned, 1, &[(0x01ff, 0x0008_0000)]);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("palette or scheme"));
    let shape = source(&owned, 1, &[(0x0181, 0xff), (0x01ff, 0x0008_0008)]);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("palette or scheme"));
    let shape = source(
        &owned,
        1,
        &[
            (0x0181, 0xff),
            (0x01d1, 1),
            (0x01c0, 0),
            (0x01ff, 0x0008_0008),
        ],
    );
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("line ends"));

    let text_error = |owned: &[(u16, Vec<u8>)]| {
        let shape = source(owned, 1, &RECTANGLE);
        project(owned, &excel_defaults(), &shape, 0xa00).unwrap_err()
    };
    let mut automatic = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    automatic[0].1 = font(240, 1, 0x7fff, 700, 0, "Meiryo UI");
    assert!(text_error(&automatic).contains("automatic"));
    let mut underlined = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    underlined[0].1 = font(240, 1, 9, 700, 1, "Meiryo UI");
    assert!(text_error(&underlined).contains("underline"));
    let mut formula = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    formula[TXO].1[16] = 5;
    assert!(text_error(&formula).contains("formula"));
    assert!(text_error(&workbook(0x0244, &compressed("2018"), 4, &[(0, 0)])).contains("justified"));
    let mut rotated = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    rotated[TXO].1[2] = 3;
    assert!(text_error(&rotated).contains("rotated XLS shape text"));
    assert!(text_error(&workbook(0x0224, &compressed("2018"), 4, &[(1, 0)])).contains("runs"));
    assert!(text_error(&workbook(0x0224, &compressed("20189"), 4, &[(0, 0)])).contains("longer"));
    let hebrew = workbook(
        0x0212,
        &[
            vec![1],
            "\u{05d0}"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect(),
        ]
        .concat(),
        1,
        &[(0, 0)],
    );
    let mut properties = RECTANGLE.to_vec();
    properties[0] = (0x008b, 2);
    let shape = source(&hebrew, 1, &properties);
    assert!(project(&hebrew, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("right-to-left"));
}

#[test]
fn drawing_group_hidden_and_property_checks() {
    let defaults = excel_defaults();
    let group = |properties: &[(u16, u32)]| ShapeSource {
        kind: 0,
        properties: properties.to_vec(),
        complex: Vec::new(),
        text: None,
    };
    assert!(!group_hidden(&defaults, &group(&[(0x0004, 0), (0x03bf, 0x0002_0000)])).unwrap());
    assert!(group_hidden(&defaults, &group(&[(0x03bf, 0x0002_0002)])).unwrap());
    // Rotation belongs to the group frame and is composed with the members.
    assert!(!group_hidden(&defaults, &group(&[(0x0004, 1)])).unwrap());
    assert!(group_hidden(&defaults, &group(&[(0x0200, 1)])).is_err());
}

fn placement(
    groups: Vec<GroupFrame>,
    anchor: Option<[i32; 4]>,
    rotation: i32,
    flips: (bool, bool),
) -> Placement {
    Placement {
        groups,
        anchor,
        rotation,
        flip_h: flips.0,
        flip_v: flips.1,
    }
}

fn close(actual: f64, expected: f64) -> bool {
    (actual - expected).abs() < 1e-6
}

#[test]
fn placement_follows_the_office_rotation_and_flip_rules() {
    // A flipped shape filling its anchor keeps its box and flips.
    let flipped = placement(Vec::new(), None, 0, (true, false)).resolve(200.0, 100.0);
    assert!(close(flipped.width, 200.0) && flipped.flip_h && !flipped.flip_v);
    // 90 degrees: the stored anchor holds the rotated bounds (swap about the
    // centre); one flip negates the angle.
    let swapped = placement(Vec::new(), None, 90 * 65_536, (false, true)).resolve(200.0, 100.0);
    assert!(close(swapped.x, 50.0) && close(swapped.y, -50.0));
    assert!(close(swapped.width, 100.0) && close(swapped.height, 200.0));
    assert!(close(swapped.rotation_degrees, -90.0));
    // -45 degrees (315) is outside the half-open swap intervals: Excel's XLSX
    // of the sample keeps that freeform's child anchor unswapped.
    let diagonal = placement(Vec::new(), None, -45 * 65_536, (false, false)).resolve(200.0, 100.0);
    assert!(close(diagonal.width, 200.0) && close(diagonal.rotation_degrees, -45.0));
}

#[test]
fn placement_composes_group_frames_in_true_proportions() {
    let frame = |anchor, rect, rotation| GroupFrame {
        anchor,
        rect,
        rotation,
        flip_h: false,
        flip_v: false,
    };
    // A member in the right half of a 180-degree rotated group lands in the
    // left half, turned 180 degrees.
    let groups = vec![frame(None, [0, 0, 200, 100], 180 * 65_536)];
    let rect = placement(groups, Some([100, 0, 200, 100]), 0, (false, false)).resolve(400.0, 100.0);
    assert!(close(rect.x, 0.0) && close(rect.width, 200.0));
    assert!(close(rect.rotation_degrees, 180.0));
    // A nested group scales its own coordinates into its child anchor.
    let groups = vec![
        frame(None, [0, 0, 100, 100], 0),
        frame(Some([50, 0, 100, 50]), [0, 0, 10, 10], 0),
    ];
    let rect = placement(groups, Some([0, 0, 5, 10]), 0, (false, false)).resolve(100.0, 100.0);
    assert!(close(rect.x, 50.0) && close(rect.width, 25.0) && close(rect.height, 50.0));
}

#[test]
fn freeform_polygons_become_custom_geometry() {
    // A closed triangle in a 100x50 path space (MS-ODRAW 2.3.6.7/9 arrays).
    let vertices = [
        vec![3, 0, 3, 0, 8, 0],
        [0i32, 0, 100, 0, 50, 50]
            .iter()
            .flat_map(|v| v.to_le_bytes())
            .collect(),
    ]
    .concat();
    let segments = [
        vec![4, 0, 4, 0, 2, 0],
        [0x4000u16, 0x0002, 0x6001, 0x8000]
            .iter()
            .flat_map(|v| v.to_le_bytes())
            .collect::<Vec<u8>>(),
    ]
    .concat();
    let owned = workbook(0x0212, &compressed("x"), 1, &[(0, 0)]);
    let mut shape = source(
        &owned,
        0,
        &[
            (0x0142, 100),
            (0x0143, 50),
            (0xc145, vertices.len() as u32),
            (0xc146, segments.len() as u32),
            (0x0181, 0xff),
            (0x01c0, 0),
            (0x01ff, 0x0008_0008),
        ],
    );
    shape.text = None;
    shape.complex = vec![(0x145, vertices), (0x146, segments)];
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    let xlsx_model::ShapeGeom::Custom { paths } = &info.geom else {
        panic!("custom geometry expected");
    };
    assert_eq!(paths.len(), 1);
    assert_eq!((paths[0].w, paths[0].h), (100.0, 50.0));
    assert!(matches!(
        paths[0].commands.last(),
        Some(xlsx_model::PathCmd::Close)
    ));
    assert_eq!(paths[0].commands.len(), 4);
}

#[test]
fn grouped_pictures_reference_their_store_media_with_the_picture_crop() {
    let picture = drawing_anchors::PictureReference {
        store_index: 7,
        crop: [0x4000, 0, 0x8000, 0],
        rotation: 0,
        clipboard_format: 0,
        auto_picture: false,
    };
    let info = picture_leaf(9, &picture);
    assert_eq!(info.z_order, 9);
    let xlsx_model::ShapeGeom::Image {
        image_path,
        src_rect,
        ..
    } = &info.geom
    else {
        panic!("image leaf expected");
    };
    assert_eq!(image_path, "legacy-xls/image/7");
    let crop = src_rect.as_ref().unwrap();
    assert_eq!((crop.t, crop.b, crop.l, crop.r), (0.25, 0.0, 0.5, 0.0));
    assert!(info.fill.is_none() && info.stroke_color.is_none() && info.text.is_none());
}

#[test]
fn open_freeform_paths_are_projected_only_without_a_shape_fill() {
    // An open two-segment line in a 100x50 path space.
    let vertices = [
        vec![3, 0, 3, 0, 8, 0],
        [0i32, 0, 100, 0, 50, 50]
            .iter()
            .flat_map(|v| v.to_le_bytes())
            .collect(),
    ]
    .concat();
    let segments = [
        vec![3, 0, 3, 0, 2, 0],
        [0x4000u16, 0x0002, 0x8000]
            .iter()
            .flat_map(|v| v.to_le_bytes())
            .collect::<Vec<u8>>(),
    ]
    .concat();
    let owned = workbook(0x0212, &compressed("x"), 1, &[(0, 0)]);
    let shape = |fill: u32| {
        let mut shape = source(
            &owned,
            0,
            &[
                (0x0142, 100),
                (0x0143, 50),
                (0xc145, vertices.len() as u32),
                (0xc146, segments.len() as u32),
                (0x0181, 0xff),
                (0x01bf, fill),
                (0x01c0, 0),
                (0x01ff, 0x0008_0008),
            ],
        );
        shape.text = None;
        shape.complex = vec![(0x145, vertices.clone()), (0x146, segments.clone())];
        shape
    };
    // Whether Excel fills an open path of a filled shape is not established.
    assert!(
        project(&owned, &excel_defaults(), &shape(0x0010_0010), 0xa00)
            .unwrap_err()
            .contains("open paths")
    );
    let info = project(&owned, &excel_defaults(), &shape(0x0010_0000), 0xa00)
        .unwrap()
        .unwrap();
    assert!(info.fill.is_none());
    assert!(matches!(&info.geom, xlsx_model::ShapeGeom::Custom { paths } if paths[0].stroke));
}
