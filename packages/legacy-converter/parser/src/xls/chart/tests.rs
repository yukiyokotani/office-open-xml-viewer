use super::super::Record;
use super::checksum;
use super::project::{project, Palette};
use super::reader::{read, Cached};

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
