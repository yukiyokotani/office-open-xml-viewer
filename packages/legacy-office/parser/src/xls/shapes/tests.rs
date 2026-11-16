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
    let mut budget = usize::MAX;
    leaf(flags)
        .project(&records, &styles, defaults, shape, &mut budget)
        .map(|leaf| leaf.map(|projected| projected.info))
}

#[test]
fn newline_expansion_is_admitted_against_the_workbook_shape_budget() {
    let owned = workbook(
        0x0212,
        &compressed("\n".repeat(100).as_str()),
        100,
        &[(0, 0)],
    );
    let shape = source(&owned, 202, &[(0x01bf, 0x0010_0000), (0x01ff, 0x0008_0000)]);
    let records = records(&owned);
    let styles = styles::Styles::parse(&records).unwrap();
    // The encoded text fits; its 101 paragraph/run objects do not.
    let mut budget = 1024usize;
    let result = project_limited(&records, &styles, &shape, &mut budget);
    assert!(result.unwrap_err().contains("shape model retention budget"));
}

fn project_limited(
    records: &[Record<'_>],
    styles: &styles::Styles<'_>,
    shape: &ShapeSource,
    budget: &mut usize,
) -> Result<Option<(xlsx_model::ShapeInfo, i32)>, String> {
    leaf(0xa00)
        .project(records, styles, &excel_defaults(), shape, budget)
        .map(|leaf| leaf.map(|projected| (projected.info, projected.rotation)))
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

use crate::rejection::{take_best_effort, BestEffort};

/// The rectangle with these properties replacing its own values of them.
fn rectangle_with(owned: &[(u16, Vec<u8>)], kind: u16, properties: &[(u16, u32)]) -> ShapeSource {
    let mut all = RECTANGLE.to_vec();
    all.retain(|(id, _)| properties.iter().all(|(other, _)| other & 0x3fff != *id));
    all.extend_from_slice(properties);
    source(owned, kind, &all)
}

#[test]
fn shape_facts_without_a_projection_still_reject() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let reject = |properties: &[(u16, u32)], kind: u16, flags: u32, expected: &str| {
        let shape = rectangle_with(&owned, kind, properties);
        let error = project(&owned, &excel_defaults(), &shape, flags)
            .err()
            .unwrap_or_else(|| panic!("{expected} accepted"));
        assert!(error.contains(expected), "{expected}: {error}");
    };
    // OLE, master-linked and background shapes, and a shape without its
    // anchor or type flags.
    reject(&[], 1, 0xa10, "shape flags");
    reject(&[], 1, 0xa20, "shape flags");
    reject(&[], 1, 0xe00, "shape flags");
    reject(&[], 1, 0x200, "shape flags");
    // Shade-shape, title and background fills, a fill in its own
    // rectangle, and patterned lines.
    reject(&[(0x0180, 6)], 1, 0xa00, "non-solid");
    reject(&[(0x0180, 8)], 1, 0xa00, "non-solid");
    reject(&[(0x01bf, 0x0012_0012)], 1, 0xa00, "non-solid");
    reject(
        &[(0x01c4, 1), (0x01ff, 0x0008_0008)],
        1,
        0xa00,
        "non-solid XLS shape lines",
    );
    reject(&[(0x0085, 1)], 1, 0xa00, "wrapping");
    // Vertical text flow and rotated fonts.
    reject(&[(0x0088, 2)], 1, 0xa00, "property 0x0088");
    reject(&[(0x0089, 1)], 1, 0xa00, "property 0x0089");
    reject(&[(0x0087, 0)], 1, 0xa00, "disagrees");
    reject(&[(0x033f, 0x0001_0001)], 1, 0xa00, "background");
    reject(&[(0x033f, 0x0020_0020)], 1, 0xa00, "OLE icon");
    reject(&[(0x033f, 0x0080_0000)], 1, 0xa00, "flip overrides");
    // A picture fill names its picture.
    reject(&[(0x0180, 3)], 1, 0xa00, "without its picture");
    // Primary and tertiary tables that disagree have no precedence.
    let mut conflicting = RECTANGLE.to_vec();
    conflicting.push((0x0181, 0x0800_0041));
    let shape = source(&owned, 1, &conflicting);
    assert!(project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap_err()
        .contains("conflicting"));

    let text_error = |owned: &[(u16, Vec<u8>)]| {
        let shape = source(owned, 1, &RECTANGLE);
        project(owned, &excel_defaults(), &shape, 0xa00).unwrap_err()
    };
    let mut formula = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    formula[TXO].1[16] = 5;
    assert!(text_error(&formula).contains("formula"));
    assert!(text_error(&workbook(0x0244, &compressed("2018"), 4, &[(0, 0)])).contains("justified"));
    let mut rotated = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    rotated[TXO].1[2] = 3;
    assert!(text_error(&rotated).contains("rotated XLS shape text"));
    assert!(text_error(&workbook(0x0224, &compressed("2018"), 4, &[(1, 0)])).contains("runs"));
    assert!(text_error(&workbook(0x0224, &compressed("20189"), 4, &[(0, 0)])).contains("longer"));
}

#[test]
fn shape_types_are_the_presets_office_converts_them_to() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let styles_records = records(&owned);
    let styles = styles::Styles::parse(&styles_records).unwrap();
    let projected = |kind: u16, properties: &[(u16, u32)], flags: u32| {
        let shape = rectangle_with(&owned, kind, properties);
        let mut budget = usize::MAX;
        leaf(flags)
            .project(
                &styles_records,
                &styles,
                &excel_defaults(),
                &shape,
                &mut budget,
            )
            .unwrap()
            .unwrap()
    };
    let name = |projected: &Projected| match &projected.info.geom {
        xlsx_model::ShapeGeom::Preset { name, adj } => {
            assert!(adj.is_empty());
            name.clone()
        }
        _ => panic!("preset expected"),
    };
    take_best_effort();
    for (kind, expected) in [
        (2, "roundRect"),
        (3, "ellipse"),
        (5, "triangle"),
        (13, "rightArrow"),
        (61, "wedgeRectCallout"),
        (202, "rect"),
    ] {
        let shape = projected(kind, &[], 0xa00);
        assert_eq!(name(&shape), expected);
        assert!(shape.preset.is_none());
    }
    // A line and a connector carry fConnector.
    assert_eq!(name(&projected(20, &[], 0xb00)), "line");
    assert_eq!(name(&projected(34, &[], 0xb00)), "bentConnector3");
    assert!(take_best_effort().is_empty());
    // Adjust values wait for the shape's extent: a right arrow whose head
    // starts at 17280/21600 of a 2:1 shape.
    let arrow = projected(13, &[(0x0147, 17_280)], 0xa00);
    assert_eq!(
        arrow.preset,
        Some((13, {
            let mut adjust = [None; 10];
            adjust[0] = Some(17_280);
            adjust
        }))
    );
    let (kind, legacy) = arrow.preset.unwrap();
    assert_eq!(
        preset_adjustments(kind, &legacy, 2_000_000.0, 1_000_000.0),
        [None, Some(40_000.0)]
    );
    assert!(take_best_effort().is_empty());
    // Adjust values without an established conversion keep the preset's
    // default proportions.
    let mut unknown = [None; 10];
    unknown[0] = Some(1);
    assert!(preset_adjustments(96, &unknown, 100.0, 100.0).is_empty());
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "XLS drawing shape type is not projected"
        )]
    );
    // A shape type without a preset is its bounding rectangle.
    assert_eq!(name(&projected(250, &[], 0xa00)), "rect");
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "XLS drawing shape type is not projected"
        )]
    );
}

#[test]
fn a_typed_shape_with_its_own_path_is_drawn_from_the_path() {
    // The closed triangle of `freeform_polygons_become_custom_geometry` on
    // a shape of type 100, with text and a pInscribe rectangle.
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
        100,
        &[
            (0x0142, 100),
            (0x0143, 50),
            (0xc145, vertices.len() as u32),
            (0xc146, segments.len() as u32),
            (0xc157, 22),
            (0x0181, 0xff),
            (0x01c0, 0),
            (0x01ff, 0x0008_0008),
        ],
    );
    shape.complex = vec![(0x145, vertices), (0x146, segments), (0x157, vec![0; 22])];
    take_best_effort();
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    let xlsx_model::ShapeGeom::Custom { paths } = &info.geom else {
        panic!("custom geometry expected");
    };
    assert_eq!(
        (paths[0].w, paths[0].h, paths[0].commands.len()),
        (100.0, 50.0, 4)
    );
    // The text is laid out in the shape's bounds.
    assert_eq!(info.text.unwrap().paragraphs.len(), 1);
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "XLS freeform text rectangles are not projected"
        )]
    );
}

#[test]
fn lines_carry_their_ends_and_compound_style() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    take_best_effort();
    // Diamond start, large wide arrow end, a double line of 6 pt.
    let shape = rectangle_with(
        &owned,
        20,
        &[
            (0x01bf, 0x0010_0000),
            (0x01c0, 0x0082_6015),
            (0x01cb, 76_200),
            (0x01cc, 0x0008_0000),
            (0x01cd, 1),
            (0x01d0, 3),
            (0x01d1, 5),
            (0x01d4, 2),
            (0x01d5, 2),
            (0x01ff, 0x0008_0008),
        ],
    );
    let info = project(&owned, &excel_defaults(), &shape, 0xb00)
        .unwrap()
        .unwrap();
    assert!(info.fill.is_none() && info.fill_color.is_none());
    assert_eq!(info.stroke_color.as_deref(), Some("#156082"));
    assert_eq!(info.stroke_width, 76_200);
    assert_eq!(info.stroke_cmpd.as_deref(), Some("dbl"));
    let end = |end: &Option<xlsx_model::ShapeLineEnd>| {
        end.as_ref()
            .map(|end| (end.r#type.clone(), end.w.clone(), end.len.clone()))
    };
    assert_eq!(
        end(&info.stroke_head_end),
        Some(("diamond".into(), "med".into(), "med".into()))
    );
    assert_eq!(
        end(&info.stroke_tail_end),
        Some(("arrow".into(), "lg".into(), "lg".into()))
    );
    assert!(take_best_effort().is_empty());
    // A line style outside MSOLINESTYLE is the simple line.
    let shape = rectangle_with(&owned, 1, &[(0x01cd, 9), (0x01ff, 0x0008_0008)]);
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert!(info.stroke_cmpd.is_none());
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::IgnoredValue,
            "XLS drawing shape property is not projected"
        )]
    );
}

#[test]
fn scheme_colors_come_from_the_workbook_color_table() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    take_best_effort();
    // A fill and line left to Excel's drawing defaults: the window
    // background (Icv 0x41) and the window text color (0x40).
    let shape = source(&owned, 1, &[(0x01ff, 0x0008_0008)]);
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert_eq!(info.fill_color.as_deref(), Some("#FFFFFF"));
    assert_eq!(info.stroke_color.as_deref(), Some("#000000"));
    // Icv 0x0A is red in the built-in palette.
    let shape = rectangle_with(&owned, 1, &[(0x0181, 0x0800_000a)]);
    let info = project(&owned, &excel_defaults(), &shape, 0xa00)
        .unwrap()
        .unwrap();
    assert_eq!(info.fill_color.as_deref(), Some("#FF0000"));
    assert!(take_best_effort().is_empty());
    // A system-indexed color, a palette index and an Icv outside the table
    // state no color: the fill is the white and the line the black MS-ODRAW
    // gives them.
    for (fill, line) in [
        (0x1000_0005, 0x1000_0008),
        (0x0100_0002, 0x0100_0003),
        (0x0800_0045, 0x0800_0046),
    ] {
        let shape = rectangle_with(
            &owned,
            1,
            &[(0x0181, fill), (0x01c0, line), (0x01ff, 0x0008_0008)],
        );
        let info = project(&owned, &excel_defaults(), &shape, 0xa00)
            .unwrap()
            .unwrap();
        assert_eq!(info.fill_color.as_deref(), Some("#FFFFFF"));
        assert_eq!(info.stroke_color.as_deref(), Some("#000000"));
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::IgnoredValue,
                "XLS drawing palette or scheme colors are not projected"
            ); 2]
        );
    }
}

#[test]
fn shadows_and_properties_without_a_counterpart_are_not_drawn() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let plain = project(
        &owned,
        &excel_defaults(),
        &source(&owned, 1, &RECTANGLE),
        0xa00,
    )
    .unwrap()
    .unwrap();
    take_best_effort();
    for (properties, expected) in [
        // fShadow with its color and offsets.
        (
            vec![(0x0201, 0), (0x0205, 76_200), (0x023f, 0x0003_0002)],
            "XLS drawing shadows are not projected",
        ),
        // A 3-D property, a complex one and a BLIP reference other than
        // the fill's.
        (
            vec![(0x0281, 5)],
            "XLS drawing shape property is not projected",
        ),
        (
            vec![(0xc0c0, 4)],
            "XLS drawing shape complex property is not projected",
        ),
        (
            vec![(0x41c5, 1)],
            "XLS drawing shape BLIP property is not projected",
        ),
    ] {
        let shape = rectangle_with(&owned, 1, &properties);
        let info = project(&owned, &excel_defaults(), &shape, 0xa00)
            .unwrap()
            .unwrap();
        assert_eq!(
            serde_json::to_value(&info).unwrap(),
            serde_json::to_value(&plain).unwrap()
        );
        assert_eq!(take_best_effort(), [(BestEffort::Skipped, expected)]);
    }
    // A shadow that is switched off needs no decision.
    let shape = rectangle_with(&owned, 1, &[(0x0201, 0), (0x023f, 0x0002_0000)]);
    project(&owned, &excel_defaults(), &shape, 0xa00).unwrap();
    assert!(take_best_effort().is_empty());
}

/// fillShadeColors: an IMsoArray of (color, 16.16 position).
fn shade_colors(stops: &[(u32, u32)]) -> Vec<u8> {
    let mut data = Vec::new();
    for value in [stops.len() as u16, stops.len() as u16, 8] {
        data.extend(value.to_le_bytes());
    }
    for (color, position) in stops {
        data.extend(color.to_le_bytes());
        data.extend(position.to_le_bytes());
    }
    data
}

#[test]
fn shaded_fills_are_gradients() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let fill = |properties: &[(u16, u32)], colors: Option<Vec<u8>>| {
        let mut shape = rectangle_with(&owned, 1, properties);
        if let Some(colors) = colors {
            shape.properties.push((0xc197, colors.len() as u32));
            shape.complex.push((0x197, colors));
        }
        let info = project(&owned, &excel_defaults(), &shape, 0xa00)
            .unwrap()
            .unwrap();
        (
            serde_json::to_value(&info.fill).unwrap(),
            info.fill_color.clone(),
        )
    };
    take_best_effort();
    // Excel's two-color vertical gradient: red (fill) at the top, blue
    // (back) at the bottom, with the fill 40% opaque.
    let (linear, fallback) = fill(
        &[
            (0x0180, 7),
            (0x0181, 0x0000_00ff),
            (0x0182, 0x6666),
            (0x0183, 0x00ff_0000),
            (0x018c, 100),
        ],
        None,
    );
    assert_eq!(
        linear,
        serde_json::json!({
            "fillType": "gradient",
            "stops": [
                {"position": 0.0, "color": "FF000066"},
                {"position": 1.0, "color": "0000FF"}
            ],
            "angle": 90.0,
            "gradType": "linear",
            "scaled": true,
            "rotWithShape": false
        })
    );
    assert_eq!(fallback.as_deref(), Some("#0000FF"));
    // msofillShade at 225 degrees with authored stops.
    let (stops, _) = fill(
        &[(0x0180, 4), (0x018b, 225 << 16), (0x018c, 100)],
        Some(shade_colors(&[
            (0x0000_00ff, 0),
            (0x0000_c800, 0x8000),
            (0x00ff_ffff, 0x1_0000),
        ])),
    );
    assert_eq!(stops["angle"], 225.0);
    assert_eq!(stops["scaled"], false);
    assert_eq!(
        stops["stops"],
        serde_json::json!([
            {"position": 0.0, "color": "FF0000"},
            {"position": 0.5, "color": "00C800"},
            {"position": 1.0, "color": "FFFFFF"}
        ])
    );
    // msofillShadeCenter: from the fillTo rectangle to the bounds. Excel
    // writes 0.5 four times for "from center" and nothing for "from corner".
    let (center, _) = fill(
        &[
            (0x0180, 5),
            (0x0181, 0x00e6_aa78),
            (0x018c, 100),
            (0x018d, 0x8000),
            (0x018e, 0x8000),
            (0x018f, 0x8000),
            (0x0190, 0x8000),
        ],
        None,
    );
    assert_eq!(
        center,
        serde_json::json!({
            "fillType": "gradient",
            "stops": [
                {"position": 0.0, "color": "78AAE6"},
                {"position": 1.0, "color": "FFFFFF"}
            ],
            "angle": 0.0,
            "gradType": "radial",
            "path": "rect",
            "fillToRect": {"l": 0.5, "t": 0.5, "r": 0.5, "b": 0.5},
            "rotWithShape": false
        })
    );
    let (corner, _) = fill(&[(0x0180, 5), (0x0181, 0x00e6_aa78), (0x018c, 100)], None);
    assert_eq!(
        corner["fillToRect"],
        serde_json::json!({"r": 1.0, "b": 1.0})
    );
    assert!(take_best_effort().is_empty());
}

#[test]
fn picture_fills_wait_for_their_media() {
    let owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    let styles_records = records(&owned);
    let styles = styles::Styles::parse(&styles_records).unwrap();
    let projected = |properties: &[(u16, u32)]| {
        let shape = rectangle_with(&owned, 1, properties);
        let mut budget = usize::MAX;
        leaf(0xa00)
            .project(
                &styles_records,
                &styles,
                &excel_defaults(),
                &shape,
                &mut budget,
            )
            .unwrap()
            .unwrap()
    };
    take_best_effort();
    // A texture repeats and a picture is stretched; fillOpacity applies.
    let texture = projected(&[(0x0180, 2), (0x4186, 3)]);
    assert_eq!(texture.blip.as_ref().map(|blip| blip.store), Some(3));
    assert_eq!(
        serde_json::to_value(&texture.info.fill).unwrap(),
        serde_json::json!({
            "fillType": "image",
            "imagePath": "legacy-xls/image/3",
            "mimeType": "",
            "layout": "tile",
            "align": [0.0, 0.0],
            "offset": [0.0, 0.0]
        })
    );
    assert!(texture.info.fill_color.is_none());
    let picture = projected(&[(0x0180, 3), (0x0182, 0x8000), (0x4186, 4)]);
    let fill = serde_json::to_value(&picture.info.fill).unwrap();
    assert_eq!(
        (&fill["layout"], &fill["alpha"]),
        (&"stretch".into(), &0.5.into())
    );
    assert!(take_best_effort().is_empty());
    // A texture with an authored tile size is drawn at the picture's size.
    projected(&[(0x0180, 2), (0x0189, 914_400), (0x4186, 3)]);
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "non-solid XLS shape fills are not projected"
        )]
    );
    // A pattern takes its colors now and its preset from the bitmap.
    let pattern = projected(&[(0x0180, 1), (0x0181, 0), (0x0183, 0x0000_ffff), (0x4186, 1)]);
    assert_eq!(pattern.blip.as_ref().map(|blip| blip.store), Some(1));
    assert_eq!(
        serde_json::to_value(&pattern.info.fill).unwrap(),
        serde_json::json!({"fillType": "pattern", "fg": "000000", "bg": "FFFF00", "preset": ""})
    );
    assert_eq!(pattern.info.fill_color.as_deref(), Some("#000000"));
    let bitmap = pattern::tests::png(pattern::tests::DARK_DOWN);
    assert_eq!(fill::pattern_preset(&bitmap, "png"), "dkDnDiag");
    assert!(take_best_effort().is_empty());
    // A bitmap this reader does not decode is an even mix of both colors.
    assert_eq!(fill::pattern_preset(b"BM", "bmp"), "pct50");
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "non-solid XLS shape fills are not projected"
        )]
    );
}

#[test]
fn text_direction_and_fonts_follow_the_shape_model() {
    let hebrew = |text: &str| -> Vec<u8> {
        [
            vec![1],
            text.encode_utf16().flat_map(u16::to_le_bytes).collect(),
        ]
        .concat()
    };
    let paragraphs = |owned: &[(u16, Vec<u8>)], direction: u32| {
        let shape = rectangle_with(owned, 1, &[(0x0087, 0), (0x008b, direction)]);
        let info = project(owned, &excel_defaults(), &shape, 0xa00)
            .unwrap()
            .unwrap();
        info.text
            .unwrap()
            .paragraphs
            .iter()
            .map(|paragraph| (paragraph.rtl, paragraph.align.clone()))
            .collect::<Vec<_>>()
    };
    take_best_effort();
    // txdir from the context: each paragraph follows its first strong
    // character; the TxO alignment stays as stored.
    let mixed = workbook(
        0x0212,
        &hebrew("1 \u{05d0}b\nab \u{05d0}\n12"),
        12,
        &[(0, 0)],
    );
    assert_eq!(
        paragraphs(&mixed, 2),
        [
            (true, "l".to_string()),
            (false, "l".into()),
            (false, "l".into())
        ]
    );
    assert_eq!(paragraphs(&mixed, 0), vec![(false, "l".to_string()); 3]);
    assert_eq!(paragraphs(&mixed, 1), vec![(true, "l".to_string()); 3]);
    assert!(take_best_effort().is_empty());

    let run = |owned: &[(u16, Vec<u8>)]| {
        let shape = source(owned, 1, &RECTANGLE);
        let info = project(owned, &excel_defaults(), &shape, 0xa00)
            .unwrap()
            .unwrap();
        let xlsx_model::ShapeTextRun::Text { bold, color, .. } =
            info.text.unwrap().paragraphs.remove(0).runs.remove(0)
        else {
            panic!("text run expected");
        };
        (bold, color)
    };
    // Underline: the run is drawn without it.
    let mut underlined = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    underlined[0].1 = font(240, 1, 9, 700, 1, "Meiryo UI");
    assert_eq!(run(&underlined), (true, Some("#FFFFFF".into())));
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::Approximated,
            "XLS shape text underline, strikethrough or font effects are not projected"
        )]
    );
    // A weight between normal and bold is the nearer of the two.
    for (weight, bold) in [(600, true), (500, false)] {
        let mut owned = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
        owned[0].1 = font(240, 1, 9, weight, 0, "Meiryo UI");
        assert_eq!(run(&owned).0, bold);
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::Approximated,
                "XLS shape text font weight is not projected"
            )]
        );
    }
    // The automatic font color is the window text color.
    let mut automatic = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    automatic[0].1 = font(240, 1, 0x7fff, 700, 0, "Meiryo UI");
    assert_eq!(run(&automatic), (true, Some("#000000".into())));
    assert!(take_best_effort().is_empty());
    // An Icv outside the color table states no color.
    let mut unknown = workbook(0x0224, &compressed("2018"), 4, &[(0, 0)]);
    unknown[0].1 = font(240, 1, 0x48, 700, 0, "Meiryo UI");
    assert_eq!(run(&unknown), (true, Some("#000000".into())));
    assert_eq!(
        take_best_effort(),
        [(
            BestEffort::IgnoredValue,
            "XLS shape text system colors are not projected"
        )]
    );
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
