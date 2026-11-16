use super::super::Record;
use super::checksum;
use super::project::{project, Palette};
use super::reader::Cached;

fn read(records: &[Record<'_>]) -> Result<super::reader::RawChart, String> {
    super::reader::read(records, &std::cell::Cell::new(usize::MAX))
}

fn record(kind: u16, data: Vec<u8>) -> (u16, Vec<u8>) {
    (kind, data)
}

fn u16s(values: &[u16]) -> Vec<u8> {
    values.iter().flat_map(|v| v.to_le_bytes()).collect()
}

fn short_text(text: &str) -> Vec<u8> {
    let mut bytes = vec![0, 0, text.len() as u8, 0];
    bytes.extend(text.bytes());
    bytes
}

fn area3d(xti: u16, rows: (u16, u16), column: u16) -> Vec<u8> {
    let mut rgce = vec![0x3b];
    rgce.extend(u16s(&[xti, rows.0, rows.1, column, column]));
    rgce
}

fn brai(id: u8, rgce: Option<Vec<u8>>) -> Vec<u8> {
    let mut bytes = vec![id, if rgce.is_some() { 2 } else { 1 }, 0, 0, 0, 0];
    let rgce = rgce.unwrap_or_default();
    bytes.extend(u16s(&[rgce.len() as u16]));
    bytes.extend(rgce);
    bytes
}

fn line_format(pattern: u16, flags: u16, icv: u16) -> Vec<u8> {
    let mut bytes = vec![0, 0, 0, 0];
    bytes.extend(u16s(&[pattern, 0xffff, flags, icv]));
    bytes
}

fn area_format(flags: u16, icv: u16) -> Vec<u8> {
    let mut bytes = vec![0x11, 0x22, 0x33, 0, 0xff, 0xff, 0xff, 0];
    bytes.extend(u16s(&[1, flags, icv, 0x4d]));
    bytes
}

fn shape_props(context: u16, checksum: u32, xml: &str) -> Vec<u8> {
    let mut bytes = u16s(&[0x08a4, 0, 0, 0, 0, 0, context, 0]);
    bytes.extend(checksum.to_le_bytes());
    bytes.extend((xml.len() as u32).to_le_bytes());
    bytes.extend(xml.bytes());
    bytes
}

/// A one-series horizontal bar chart. `cache` adds a data cache; the series
/// format is a palette-indexed fill with optional ShapePropsStream XML.
fn bar_chart(cache: bool, shape: Option<(u32, &str)>) -> Vec<(u16, Vec<u8>)> {
    let line = line_format(0, 0, 8);
    let area = area_format(0, 10);
    let mut format = vec![
        record(0x1006, u16s(&[0xffff, 0, 0, 0])),
        record(0x1033, vec![]),
        record(0x1007, line),
        record(0x100a, area),
    ];
    if let Some((stored, xml)) = shape {
        format.push(record(0x08a4, shape_props(0, stored, xml)));
    }
    format.push(record(0x1034, vec![]));
    let mut records = vec![
        record(0x0809, u16s(&[0x0600, 0x0020, 0, 0])),
        record(0x1002, vec![0; 16]),
        record(0x1033, vec![]),
        record(0x1003, u16s(&[3, 1, 2, 2, 1, 0])),
        record(0x1033, vec![]),
        record(0x1051, brai(0, None)),
        record(0x100d, short_text("Sales")),
        record(0x1051, brai(1, Some(area3d(0, (1, 2), 1)))),
        record(0x1051, brai(2, Some(area3d(0, (1, 2), 0)))),
    ];
    records.extend(format);
    records.extend([
        record(0x1045, u16s(&[0])),
        record(0x1034, vec![]),
        record(0x1041, vec![0; 18]),
        record(0x1033, vec![]),
        record(0x1014, [vec![0; 16], u16s(&[0, 0])].concat()),
        record(0x1033, vec![]),
        record(0x1017, u16s(&[0, 150, 1])),
        record(0x1034, vec![]),
        record(0x1034, vec![]),
        record(0x1034, vec![]),
    ]);
    if cache {
        let number = |point: u16, value: f64| {
            let mut data = u16s(&[point, 0, 0]);
            data.extend(value.to_le_bytes());
            record(0x0203, data)
        };
        let label = |point: u16, text: &str| {
            let mut data = u16s(&[point, 0, 0, text.len() as u16]);
            data.push(0);
            data.extend(text.bytes());
            record(0x0204, data)
        };
        records.extend([
            record(0x1065, u16s(&[1])),
            number(0, 4.0),
            number(1, 7.5),
            record(0x1065, u16s(&[2])),
            label(0, "North"),
            label(1, "South"),
        ]);
    }
    records.push(record(0x000a, vec![]));
    records
}

fn as_records(owned: &[(u16, Vec<u8>)]) -> Vec<Record<'_>> {
    owned
        .iter()
        .map(|(kind, data)| Record {
            kind: *kind,
            offset: 0,
            data,
        })
        .collect()
}

fn palette_color(icv: u16) -> Option<String> {
    (icv == 10).then(|| "#FF0000".into())
}

fn no_font(_: u16) -> Option<super::super::styles::ChartFont> {
    None
}

fn no_decode(_: &[u8]) -> Option<super::super::styles::ChartFont> {
    None
}

fn palette() -> Palette<'static> {
    Palette {
        global_font: &no_font,
        decode_font: &no_decode,
        global_font_count: 0,
        icv: &palette_color,
        theme: std::array::from_fn(|index| (index == 4).then(|| "4472C4".into())),
    }
}

/// Axis records are scoped by their AxisParent and Begin/End blocks. The
/// expected tick/label names are the shared renderer's DrawingML contract,
/// independently specified by MS-XLS 2.4.327, not a captured parser model.
#[test]
fn primary_axis_ticks_and_label_positions_survive_biff_projection() {
    let mut owned = bar_chart(true, None);
    let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
    let axis = |kind, major, minor, labels| {
        let mut tick = vec![0; 30];
        tick[..3].copy_from_slice(&[major, minor, labels]);
        vec![
            record(0x101d, [u16s(&[kind]), vec![0; 16]].concat()),
            record(0x1033, vec![]),
            record(0x101e, tick),
            record(0x1034, vec![]),
        ]
    };
    owned.splice(at..at, [axis(0, 0, 1, 2), axis(1, 3, 2, 0)].concat());
    let chart_end = owned.iter().position(|r| r.0 == 0x1065).unwrap() - 1;
    let mut secondary = vec![
        record(0x1041, [u16s(&[1]), vec![0; 16]].concat()),
        record(0x1033, vec![]),
    ];
    secondary.extend(axis(1, 1, 3, 1));
    secondary.push(record(0x1034, vec![]));
    owned.splice(chart_end..chart_end, secondary);
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|_| None).unwrap();
    assert_eq!(model.cat_axis_major_tick_mark, "none");
    assert_eq!(model.cat_axis_minor_tick_mark.as_deref(), Some("in"));
    assert_eq!(model.cat_axis_tick_label_pos.as_deref(), Some("high"));
    assert_eq!(model.val_axis_major_tick_mark, "cross");
    assert_eq!(model.val_axis_minor_tick_mark.as_deref(), Some("out"));
    assert_eq!(model.val_axis_tick_label_pos.as_deref(), Some("none"));
}

#[test]
fn malformed_axis_tick_enumerations_fail_closed() {
    let mut owned = bar_chart(true, None);
    let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
    let mut tick = vec![0; 30];
    tick[..3].copy_from_slice(&[4, 0, 3]);
    owned.splice(
        at..at,
        [
            record(0x101d, vec![0; 18]),
            record(0x1033, vec![]),
            record(0x101e, tick),
            record(0x1034, vec![]),
        ],
    );
    assert!(read(&as_records(&owned)).is_err());
}

#[test]
fn value_axis_units_and_direction_respect_biff_automatic_flags() {
    let project_range = |flags: u16| {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
        let mut range: Vec<u8> = [0.0f64, 10.0, 2.0, 0.5, 0.0]
            .into_iter()
            .flat_map(f64::to_le_bytes)
            .collect();
        range.extend(flags.to_le_bytes());
        owned.splice(
            at..at,
            [
                record(0x101d, [u16s(&[1]), vec![0; 16]].concat()),
                record(0x1033, vec![]),
                record(0x101f, range),
                record(0x1034, vec![]),
            ],
        );
        let raw = read(&as_records(&owned)).unwrap();
        project(&raw, &palette(), &|_| None).unwrap()
    };
    let explicit = project_range(0x40);
    assert_eq!(explicit.val_axis_major_unit, Some(2.0));
    assert_eq!(explicit.val_axis_minor_unit, Some(0.5));
    assert_eq!(explicit.val_axis_orientation.as_deref(), Some("maxMin"));
    let automatic = project_range(4 | 8);
    assert_eq!(automatic.val_axis_major_unit, None);
    assert_eq!(automatic.val_axis_minor_unit, None);
    assert_eq!(automatic.val_axis_orientation, None);
    let mixed = project_range(4);
    assert_eq!(mixed.val_axis_major_unit, None);
    assert_eq!(mixed.val_axis_minor_unit, Some(0.5));
}

#[test]
fn percent_value_axis_uses_canonical_ratio_units_for_all_percent_families() {
    // BIFF Bar/Line/Area f100 uses percentage-point axis values. The canonical
    // OOXML model is fractional: a 50-point interval is 0.5 (50%). Exercise
    // production reading/projection and retain the ordinary-axis boundary.
    for (kind, data, expected_type, percent) in [
        (0x1017, u16s(&[0, 150, 6]), "stackedBarPct", true),
        (0x1017, u16s(&[0, 150, 7]), "stackedBarHPct", true),
        (0x1018, u16s(&[3]), "stackedLinePct", true),
        (0x101a, u16s(&[3]), "stackedAreaPct", true),
        (0x1017, u16s(&[0, 150, 2]), "stackedBar", false),
    ] {
        let mut owned = bar_chart(true, None);
        let group = owned.iter_mut().find(|r| r.0 == 0x1017).unwrap();
        *group = record(kind, data);
        let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
        let mut range: Vec<u8> = [-100.0f64, 100.0, 50.0, 10.0, 0.0]
            .into_iter()
            .flat_map(f64::to_le_bytes)
            .collect();
        range.extend(0u16.to_le_bytes());
        owned.splice(
            at..at,
            [
                record(0x101d, [u16s(&[1]), vec![0; 16]].concat()),
                record(0x1033, vec![]),
                record(0x101f, range),
                record(0x1034, vec![]),
            ],
        );
        let model = project(&read(&as_records(&owned)).unwrap(), &palette(), &|_| None).unwrap();
        assert_eq!(model.chart_type, expected_type);
        let factor = if percent { 0.01 } else { 1.0 };
        assert_eq!(model.val_min, Some(-100.0 * factor), "{expected_type}");
        assert_eq!(model.val_max, Some(100.0 * factor), "{expected_type}");
        assert_eq!(
            model.val_axis_major_unit,
            Some(50.0 * factor),
            "{expected_type}"
        );
        assert_eq!(
            model.val_axis_minor_unit,
            Some(10.0 * factor),
            "{expected_type}"
        );
    }
}

#[test]
fn explicit_axis_units_reject_invalid_xnum_and_interval_domains() {
    let read_units = |major: f64, minor: f64, flags: u16| {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
        let mut range: Vec<u8> = [0.0f64, 10.0, major, minor, 0.0]
            .into_iter()
            .flat_map(f64::to_le_bytes)
            .collect();
        range.extend(flags.to_le_bytes());
        owned.splice(
            at..at,
            [
                record(0x101d, [u16s(&[1]), vec![0; 16]].concat()),
                record(0x1033, vec![]),
                record(0x101f, range),
                record(0x1034, vec![]),
            ],
        );
        read(&as_records(&owned))
    };
    for (major, minor) in [
        (2.0, -0.5),
        (0.5, 2.0),
        (f64::NAN, 0.5),
        (2.0, f64::INFINITY),
        (2.0, -0.0),
        (2.0, f64::from_bits(1)),
    ] {
        assert!(read_units(major, minor, 0).is_err(), "{major:?}, {minor:?}");
    }
    assert!(read_units(0.0, 0.0, 0).is_ok());
    assert!(read_units(f64::NAN, -0.5, 4 | 8).is_ok());
    assert!(read_units(f64::NAN, 0.5, 4).is_ok());
}

#[test]
fn chart_cache_drives_the_series_and_palette_formats_color_it() {
    let owned = bar_chart(true, None);
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|_| panic!("the cache is authoritative")).unwrap();
    assert_eq!(model.chart_type, "clusteredBarH");
    assert_eq!(model.categories, ["North", "South"]);
    let series = &model.series[0];
    assert_eq!(series.name, "Sales");
    assert_eq!(series.values, [Some(4.0), Some(7.5)]);
    assert_eq!(series.series_type.as_deref(), Some("bar"));
    // AreaFormat rgbFore (112233) is ignored by Excel; icvFore selects the color.
    assert_eq!(series.color.as_deref(), Some("FF0000"));
}

#[test]
fn empty_cache_reads_the_referenced_worksheet_cells() {
    let owned = bar_chart(false, None);
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|rgce| {
        Some(match rgce[7] {
            1 => vec![Some(Cached::Number(3.0)), None],
            _ => vec![Some(Cached::Text("A".into())), Some(Cached::Number(2.0))],
        })
    })
    .unwrap();
    assert_eq!(model.series[0].values, [Some(3.0), None]);
    assert_eq!(model.categories, ["A", "2"]);
}

#[test]
fn shape_xml_is_used_only_when_its_checksum_matches_the_biff_formats() {
    let xml = r#"<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:spPr>"#;
    let mut data = checksum::line_properties(&line_format(0, 0, 8))
        .unwrap()
        .to_vec();
    let area = area_format(0, 10);
    data.extend(checksum::interior_properties(&area, &area).unwrap());
    let valid = checksum::crc(&data);
    // A verified empty stream is an empty spPr: automatic formatting, which
    // takes the theme accent instead of the explicit BIFF palette color.
    for (stored, xml, expected) in [
        (valid, xml, "4472C4"),
        (valid ^ 1, xml, "FF0000"),
        (valid, "", "4472C4"),
    ] {
        let owned = bar_chart(true, Some((stored, xml)));
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        assert_eq!(model.series[0].color.as_deref(), Some(expected));
    }
}

#[test]
fn a_maximal_shape_xml_length_leaves_the_biff_formats() {
    let mut data = checksum::line_properties(&line_format(0, 0, 8))
        .unwrap()
        .to_vec();
    let area = area_format(0, 10);
    data.extend(checksum::interior_properties(&area, &area).unwrap());
    // With its stated length the verified empty stream would win (4472C4).
    let mut owned = bar_chart(true, Some((checksum::crc(&data), "")));
    let shape = owned
        .iter_mut()
        .find(|(kind, _)| *kind == 0x08a4)
        .expect("ShapePropsStream");
    shape.1[20..24].copy_from_slice(&u32::MAX.to_le_bytes());
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|_| None).unwrap();
    assert_eq!(model.series[0].color.as_deref(), Some("FF0000"));
}

#[test]
fn a_chart_without_series_records_is_an_authored_empty_chart() {
    let owned = vec![
        record(0x0809, u16s(&[0x0600, 0x0020, 0, 0])),
        record(0x1002, vec![0; 16]),
        record(0x1033, vec![]),
        record(0x1041, vec![0; 18]),
        record(0x1033, vec![]),
        record(0x1014, [vec![0; 16], u16s(&[0, 0])].concat()),
        record(0x1033, vec![]),
        record(0x1017, u16s(&[0, 150, 1])),
        record(0x1034, vec![]),
        record(0x1034, vec![]),
        record(0x1034, vec![]),
        record(0x000a, vec![]),
    ];
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|_| None).unwrap();
    assert!(model.authored_without_series);
    assert!(model.series.is_empty());
    let budget = std::cell::Cell::new(1usize);
    assert!(
        super::project::project_bounded(&raw, &palette(), &|_| Ok(None), &budget)
            .unwrap_err()
            .contains("chart model retention budget")
    );
    // A chart with a Series record is never marked authored-empty.
    let with_series = read(&as_records(&bar_chart(true, None))).unwrap();
    assert!(
        !project(&with_series, &palette(), &|_| None)
            .unwrap()
            .authored_without_series
    );
}

#[test]
fn an_automatic_chart_area_takes_the_biff_outline_excel_writes_for_it() {
    let line = line_format(0, 0, 10);
    let area = area_format(0, 10);
    let mut data = checksum::line_properties(&line).unwrap().to_vec();
    data.extend(checksum::interior_properties(&area, &area).unwrap());
    let valid = checksum::crc(&data);
    let chart = |xml: &str| {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|(kind, _)| *kind == 0x1002).unwrap() + 2;
        owned.splice(
            at..at,
            [
                record(0x1032, u16s(&[0, 2])),
                record(0x1033, vec![]),
                record(0x1007, line.clone()),
                record(0x100a, area.clone()),
                record(0x08a4, shape_props(0, valid, xml)),
                record(0x1034, vec![]),
            ],
        );
        let raw = read(&as_records(&owned)).unwrap();
        project(&raw, &palette(), &|_| None).unwrap()
    };
    // Empty stream: the BIFF outline and fill.
    let empty = chart("");
    assert_eq!(empty.chart_border_color.as_deref(), Some("FF0000"));
    assert_eq!(empty.chart_bg.as_deref(), Some("FF0000"));
    // A stream with DrawingML still supersedes the BIFF records.
    let xml = r#"<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ln><a:noFill/></a:ln></a:spPr>"#;
    let styled = chart(xml);
    assert_eq!(styled.chart_border_color, None);
    assert_eq!(styled.chart_border_hidden, Some(true));
}

#[test]
fn referenced_shared_text_is_charged_before_every_copy() {
    let budget = std::cell::Cell::new(std::mem::size_of::<Cached>() + 4);
    let cells = super::Cells {
        sheets_by_xti: Vec::new(),
        sheets: &[],
        shared: &[],
        budget: &budget,
        budget_exceeded: std::cell::Cell::new(false),
    };
    assert!(cells.copy_text("same").is_some());
    assert!(cells.copy_text("same").is_none());
    assert!(cells.budget_exceeded.get());
}

#[test]
fn chart_context_propagates_invalid_theme_instead_of_default_colors() {
    let styles = super::super::styles::Styles::parse(&[]).unwrap();
    let records = [Record {
        kind: 0x0896,
        offset: 0,
        data: &[0; 16],
    }];
    let budget = std::cell::Cell::new(1024);
    assert!(super::with_context(&records, &[], &styles, &[], &[], &budget, |_, _| ()).is_err());
}

#[test]
fn numeric_chart_expansion_stops_at_aggregate_typed_slot_admission() {
    let owned = bar_chart(false, None);
    let raw = read(&as_records(&owned)).unwrap();
    let calls = std::cell::Cell::new(0);
    let references = |_: &[u8]| {
        calls.set(calls.get() + 1);
        Ok(Some(vec![Some(Cached::Number(0.0)); 100]))
    };
    let budget = std::cell::Cell::new(1);
    assert!(
        super::project::project_bounded(&raw, &palette(), &references, &budget)
            .unwrap_err()
            .contains("budget")
    );
    assert_eq!(calls.get(), 1); // no category/bubble expansion after failure
}

/// Chart3d is the BIFF view/group contract, not a different canonical family.
/// Exercise the record -> shared-model boundary used by both native XLS hosts.
#[test]
fn chart3d_projects_camera_and_bar_depth_grouping() {
    for (bar_flags, view_flags, expected_grouping, height) in [
        (0, 0x17, "clustered", None),
        (0, 0x10, "standard", Some(125.0)),
        (2, 0x11, "stacked", Some(125.0)),
        (6, 0x11, "percentStacked", Some(125.0)),
    ] {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
        owned[at].1 = u16s(&[0, 150, bar_flags]);
        owned.insert(
            at + 1,
            record(
                0x103a,
                u16s(&[270, (-25i16) as u16, 45, 125, 240, 80, view_flags]),
            ),
        );
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        let view = model.three_d.expect("Chart3d must reach the renderer");
        assert_eq!(view.rotation_x, Some(-25));
        assert_eq!(view.rotation_y, Some(270));
        assert_eq!(view.perspective, Some(45));
        assert_eq!(view.right_angle_axes, Some(view_flags & 1 == 0));
        assert_eq!(view.depth_percent, Some(240.0));
        assert_eq!(view.gap_depth_percent, Some(80.0));
        assert_eq!(view.height_percent, height);
        assert_eq!(view.height_percent_authored, Some(height.is_some()));
        assert_eq!(view.bar_grouping.as_deref(), Some(expected_grouping));
    }
}

#[test]
fn chart3d_bar_shapes_follow_series_format_and_ignore_2d_shapes() {
    for (riser, taper, expected) in [
        (0, 0, "box"),
        (1, 0, "cylinder"),
        (1, 1, "cone"),
        (1, 2, "coneToMax"),
        (0, 1, "pyramid"),
        (0, 2, "pyramidToMax"),
    ] {
        let mut owned = bar_chart(true, None);
        let format_at = owned.iter().position(|r| r.0 == 0x1006).unwrap() + 2;
        owned.insert(format_at, record(0x105f, vec![riser, taper]));
        let raw = read(&as_records(&owned)).unwrap();
        let flat = project(&raw, &palette(), &|_| None).unwrap();
        assert!(flat.three_d.is_none());
        assert!(flat.series[0].three_d_shape.is_none());
        let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
        owned.insert(
            at + 1,
            record(0x103a, u16s(&[20, 15, 30, 100, 100, 150, 0x17])),
        );
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        assert_eq!(model.series[0].three_d_shape.as_deref(), Some(expected));
    }
}

#[test]
fn malformed_chart3d_is_rejected_instead_of_silently_flattened() {
    let mut owned = bar_chart(true, None);
    let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
    owned.insert(at + 1, record(0x103a, vec![0; 12]));
    assert!(read(&as_records(&owned)).is_err());
}

#[test]
fn chart3d_line_area_pie_and_surface_keep_their_family_and_view() {
    for (record_kind, data, flags, expected_type) in [
        (0x1018, u16s(&[0]), 0x11, "line"),
        (0x101a, u16s(&[0]), 0x11, "area"),
        (0x1019, u16s(&[0, 0, 0]), 0, "pie"),
        (0x103f, u16s(&[0]), 0x11, "surface"),
    ] {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
        owned[at] = record(record_kind, data);
        owned.insert(
            at + 1,
            record(0x103a, u16s(&[20, 15, 30, 125, 100, 150, flags])),
        );
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        assert_eq!(model.chart_type, expected_type);
        let view = model.three_d.unwrap();
        assert_eq!(view.rotation_x, Some(15));
        if expected_type == "pie" {
            // BIFF pcHeight is pie thickness, not the OOXML hPercent multiplier.
            assert!(view.height_percent.is_none());
        }
        if expected_type == "surface" {
            assert_eq!(model.surface_wireframe, Some(true));
        }
    }
}

#[test]
fn standard_chart3d_keeps_series_axis_ticks_title_direction_and_skips() {
    let mut owned = bar_chart(true, None);
    let at = owned.iter().position(|r| r.0 == 0x1014).unwrap();
    let mut ticks = vec![0; 30];
    ticks[..3].copy_from_slice(&[1, 0, 2]);
    owned.splice(
        at..at,
        [
            record(0x101d, [u16s(&[2]), vec![0; 16]].concat()),
            record(0x1033, vec![]),
            record(0x1020, u16s(&[0, 2, 3, 4])),
            record(0x101e, ticks),
            record(0x1034, vec![]),
            record(0x1025, vec![0; 32]),
            record(0x1033, vec![]),
            record(0x1027, u16s(&[7, 0, 0])),
            record(0x100d, short_text("Series")),
            record(0x1034, vec![]),
        ],
    );
    let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
    owned[at].1 = u16s(&[0, 150, 0]);
    owned.insert(
        at + 1,
        record(0x103a, u16s(&[20, 15, 30, 125, 100, 150, 0x15])),
    );
    let raw = read(&as_records(&owned)).unwrap();
    let model = project(&raw, &palette(), &|_| None).unwrap();
    let axis = model.three_d.unwrap().series_axis.unwrap();
    assert_eq!(axis.major_tick_mark, "in");
    assert_eq!(axis.tick_label_pos.as_deref(), Some("high"));
    assert_eq!(axis.title.as_deref(), Some("Series"));
    assert_eq!(axis.orientation.as_deref(), Some("maxMin"));
    assert_eq!(axis.tick_label_skip, Some(2));
    assert_eq!(axis.tick_mark_skip, Some(3));
}

#[test]
fn bubble_3d_effect_is_series_scoped_and_ignored_for_other_families() {
    for bubbles in [false, true] {
        let mut owned = bar_chart(true, None);
        let at = owned.iter().position(|r| r.0 == 0x1006).unwrap() + 2;
        owned.insert(at, record(0x105d, u16s(&[2])));
        if bubbles {
            let at = owned.iter().position(|r| r.0 == 0x1017).unwrap();
            owned[at] = record(0x101b, u16s(&[100, 1, 1]));
        }
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        assert_eq!(model.series[0].bubble_3d, bubbles.then_some(true));
        assert!(model.three_d.is_none());
    }
}

#[test]
fn chart3d_group_shape_supplies_only_missing_series_shapes() {
    for override_shape in [None, Some(vec![1, 0])] {
        let mut owned = bar_chart(true, None);
        if let Some(shape) = override_shape.as_ref() {
            let at = owned.iter().position(|r| r.0 == 0x1006).unwrap() + 2;
            owned.insert(at, record(0x105f, shape.clone()));
        }
        let at = owned.iter().position(|r| r.0 == 0x1017).unwrap() + 1;
        owned.splice(
            at..at,
            [
                record(0x103a, u16s(&[20, 15, 30, 100, 100, 150, 0x17])),
                record(0x1006, u16s(&[0xffff, 0, 0, 0])),
                record(0x1033, vec![]),
                record(0x105f, vec![0, 1]),
                record(0x1034, vec![]),
            ],
        );
        let raw = read(&as_records(&owned)).unwrap();
        let model = project(&raw, &palette(), &|_| None).unwrap();
        let view = model.three_d.unwrap();
        assert_eq!(view.shape.as_deref(), Some("pyramid"));
        assert_eq!(
            model.series[0].three_d_shape.as_deref(),
            override_shape.map(|_| "cylinder")
        );
    }
}

#[test]
fn chart_depth_ceiling_is_resource_policy() {
    let mut records = vec![record(0x0809, u16s(&[0x0600, 0x0020]))];
    records.extend((0..33).map(|_| record(0x1033, vec![])));
    assert!(read(&as_records(&records))
        .err()
        .unwrap()
        .starts_with("LEGACY_RESOURCE_LIMIT:"));
}

#[test]
fn chart_cache_index_ceilings_keep_resource_classification() {
    for (point, series) in [(32_000, 0), (0, 255)] {
        let records = vec![
            record(0x0809, u16s(&[0x0600, 0x0020])),
            record(0x1065, u16s(&[1])),
            record(0x0203, u16s(&[point, series])),
        ];
        assert!(read(&as_records(&records))
            .err()
            .unwrap()
            .starts_with("LEGACY_RESOURCE_LIMIT:"));
    }
}

#[test]
fn first_over_limit_reference_without_cache_rejects_the_chart() {
    let mut owned = bar_chart(false, None);
    for (kind, data) in &mut owned {
        if *kind == 0x1051 && data.first() == Some(&1) {
            *data = brai(
                1,
                Some(area3d(0, (0, super::MAX_REFERENCED_CELLS as u16), 0)),
            );
        }
    }
    let sheets = vec![("Sheet".into(), super::super::SheetData::default())];
    let budget = std::cell::Cell::new(usize::MAX);
    let cells = super::Cells {
        sheets_by_xti: vec![Some(0)],
        sheets: &sheets,
        shared: &[],
        budget: &budget,
        budget_exceeded: std::cell::Cell::new(false),
    };
    // The last admissible reference succeeds; the next one must not become
    // unresolved content (empty values or generated category labels).
    assert_eq!(
        cells
            .resolve(&area3d(0, (0, super::MAX_REFERENCED_CELLS as u16 - 1), 0))
            .unwrap()
            .unwrap()
            .len(),
        super::MAX_REFERENCED_CELLS
    );
    let error = super::parse(
        &as_records(&owned),
        &palette(),
        &|rgce| cells.resolve(rgce),
        &budget,
    )
    .unwrap_err();
    assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
    assert!(error.contains("XLS chart reference retention budget exceeded"));
    assert!(cells.resolve(&area3d(0, (0, 0), 0)).is_err());
}

#[test]
fn chart_fonts_beyond_512_are_retained_in_authored_order() {
    let mut owned = bar_chart(true, None);
    let font = |size: u16| {
        let mut data = vec![0; 16];
        data[..2].copy_from_slice(&size.to_le_bytes());
        data[6..8].copy_from_slice(&400u16.to_le_bytes());
        data[14] = 1;
        data.push(b'F');
        record(0x0031, data)
    };
    owned.splice(
        1..1,
        (0..513).map(|i| font(if i == 512 { 360 } else { 200 })),
    );
    // MS-XLS FontX addresses the 513th local font for the chart title.
    owned.splice(
        516..516,
        [
            record(0x1025, vec![]),
            record(0x1033, vec![]),
            record(0x1026, u16s(&[513])),
            record(0x1027, u16s(&[1])),
            record(0x1034, vec![]),
        ],
    );
    let styles = super::super::styles::Styles::parse(&[]).unwrap();
    let decode = |bytes: &[u8]| styles.chart_font(bytes);
    let mut colors = palette();
    colors.decode_font = &decode;
    let model = project(&read(&as_records(&owned)).unwrap(), &colors, &|_| None).unwrap();
    assert_eq!(model.title_font_size_hpt, Some(1800)); // 360 twips = 18 pt
    assert_eq!(model.title_font_face.as_deref(), Some("F"));
}

#[test]
fn font_retention_exhaustion_rejects_instead_of_dropping_the_font() {
    let mut owned = bar_chart(true, None);
    owned.splice(1..1, (0..513).map(|_| record(0x0031, vec![0; 20])));
    let budget = std::cell::Cell::new(512 * (20 + std::mem::size_of::<Vec<u8>>()));
    let error = super::reader::read(&as_records(&owned), &budget).unwrap_err();
    assert_eq!(
        error,
        "LEGACY_RESOURCE_LIMIT:XLS chart font retention budget exceeded"
    );
}

/// One OfficeArt record (MS-ODRAW 2.2.1).
fn art(kind: u16, options: u16, data: &[u8]) -> Vec<u8> {
    [
        options.to_le_bytes().to_vec(),
        kind.to_le_bytes().to_vec(),
        (data.len() as u32).to_le_bytes().to_vec(),
        data.to_vec(),
    ]
    .concat()
}

/// A chart sheet whose own chart area is 400 x 200 pt, with `objects` chart
/// objects (Obj ot 5 and their chart substreams) anchored at
/// (1000, 2000)-(3000, 4000) SPRC.
fn chart_sheet_records(objects: usize) -> Vec<(u16, Vec<u8>)> {
    let mut own = bar_chart(true, None);
    let mut area = vec![0u8; 8];
    area.extend((400u32 << 16).to_le_bytes());
    area.extend((200u32 << 16).to_le_bytes());
    own[1].1 = area;
    let mut records = vec![own.remove(0)];
    for object in 0..objects {
        let mut anchor = vec![0u8, 0];
        for value in [1000i32, 2000, 3000, 4000] {
            anchor.extend(value.to_le_bytes());
        }
        let mut shape = art(
            0xf00a,
            0x0c92,
            &[
                u32::to_le_bytes(1025 + object as u32),
                0x0a00u32.to_le_bytes(),
            ]
            .concat(),
        );
        shape.extend(art(0xf010, 0, &anchor));
        shape.extend(art(0xf011, 0, &[]));
        let shape = art(0xf004, 0x000f, &shape);
        let drawing = if object == 0 {
            let mut patriarch = art(0xf009, 0x0001, &[0; 16]);
            patriarch.extend(art(
                0xf00a,
                0x0002,
                &[1024u32.to_le_bytes(), 5u32.to_le_bytes()].concat(),
            ));
            let mut group = art(0xf004, 0x000f, &patriarch);
            group.extend(&shape);
            // The containers span every object's shape container.
            let mut group = art(0xf003, 0x000f, &group);
            let extra = shape.len() * (objects - 1);
            let length = u32::from_le_bytes(group[4..8].try_into().unwrap()) + extra as u32;
            group[4..8].copy_from_slice(&length.to_le_bytes());
            let mut root = art(0xf008, 0x0010, &[0; 8]);
            root.extend(group);
            let mut root = art(0xf002, 0x000f, &root);
            let length = u32::from_le_bytes(root[4..8].try_into().unwrap()) + extra as u32;
            root[4..8].copy_from_slice(&length.to_le_bytes());
            root
        } else {
            shape
        };
        records.push(record(0x00ec, drawing));
        let mut object_record = u16s(&[0x15, 0x12, 5, 1 + object as u16, 0x4011]);
        object_record.extend([0u8; 16]);
        records.push(record(0x005d, object_record));
        records.extend(bar_chart(true, None));
    }
    records.extend(own);
    records
}

/// The chart sheet of `chart_sheet_records(objects)`: its chart area and
/// own chart, then each embedded chart's rectangle and chart, as JSON (a
/// chart model is large, so none is kept on the caller's stack).
#[inline(never)]
fn projected_chart_sheet(
    owned: &[(u16, Vec<u8>)],
) -> Result<Vec<([i64; 4], serde_json::Value)>, String> {
    let styles = super::super::styles::Styles::parse(&[]).unwrap();
    let records: Vec<_> = owned
        .iter()
        .enumerate()
        .map(|(offset, (kind, data))| Record {
            kind: *kind,
            offset,
            data,
        })
        .collect();
    let budget = std::cell::Cell::new(1 << 20);
    let sheet = super::chart_sheet(&records, 0, &[], &styles, &[], &[], &budget)?;
    let mut output = vec![(
        [sheet.x_emu, sheet.y_emu, sheet.width_emu, sheet.height_emu],
        serde_json::to_value(sheet.model.as_ref()).unwrap(),
    )];
    for embedded in &sheet.embedded {
        output.push((
            [
                embedded.x_emu,
                embedded.y_emu,
                embedded.width_emu,
                embedded.height_emu,
            ],
            serde_json::to_value(embedded.model.as_ref()).unwrap(),
        ));
    }
    Ok(output)
}

#[test]
fn chart_objects_of_a_chart_sheet_are_drawn_over_its_chart() {
    let plain = projected_chart_sheet(&chart_sheet_records(0)).unwrap();
    assert_eq!(plain.len(), 1);
    assert_eq!(plain[0].0, [0, 0, 400 * 12_700, 200 * 12_700]);
    assert_eq!(plain[0].1["chartType"], "clusteredBarH");
    for objects in [1, 2] {
        let with = projected_chart_sheet(&chart_sheet_records(objects)).unwrap();
        // The sheet's own chart is the records outside the chart objects.
        assert_eq!(with[0], plain[0]);
        assert_eq!(with.len(), 1 + objects);
        for embedded in &with[1..] {
            // 1/4000 of the 400 x 200 pt chart area.
            assert_eq!(
                embedded.0,
                [100 * 12_700, 100 * 12_700, 200 * 12_700, 100 * 12_700]
            );
            assert_eq!(embedded.1, plain[0].1);
        }
    }
    // The chart substream must follow a chart object's Obj record.
    let mut owned = chart_sheet_records(1);
    let object = owned.iter().position(|(kind, _)| *kind == 0x005d).unwrap();
    owned[object].1[4] = 2;
    assert!(projected_chart_sheet(&owned)
        .expect_err("rejected")
        .contains("invalid BIFF drawing object owner"));
    let mut owned = chart_sheet_records(1);
    owned[object].1.truncate(22);
    assert!(projected_chart_sheet(&owned)
        .expect_err("rejected")
        .contains("invalid BIFF drawing object owner"));
    // A substream no chart object owns is not a chart sheet's structure.
    let mut owned = chart_sheet_records(0);
    owned.splice(1..1, bar_chart(true, None));
    assert!(projected_chart_sheet(&owned)
        .expect_err("rejected")
        .contains("nested BIFF substream inside a chart"));
}
