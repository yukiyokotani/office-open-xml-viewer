//! Bounded reader for one BIFF8 chart substream (MS-XLS 2.1.7.20.1). It keeps
//! the records the projection needs in typed form and never interprets
//! formulas: series content comes from the chart data cache (2.2.3.2).
use super::super::{u16_at, u32_at, unsupported, Record};
use super::checksum;
use std::collections::BTreeMap;

const MAX_SERIES: usize = 255;
const MAX_POINTS: usize = 32_000;
const MAX_DEPTH: usize = 32;
const MAX_TEXT_BYTES: usize = 1 << 20;
const MAX_XML_BYTES: usize = 1 << 20;

/// Record identifiers (MS-XLS 2.3).
mod id {
    pub const BOF: u16 = 0x0809;
    pub const EOF: u16 = 0x000a;
    pub const CONTINUE: u16 = 0x003c;
    pub const CHART: u16 = 0x1002;
    pub const SERIES: u16 = 0x1003;
    pub const DATA_FORMAT: u16 = 0x1006;
    pub const LINE_FORMAT: u16 = 0x1007;
    pub const MARKER_FORMAT: u16 = 0x1009;
    pub const AREA_FORMAT: u16 = 0x100a;
    pub const PIE_FORMAT: u16 = 0x100b;
    pub const ATTACHED_LABEL: u16 = 0x100c;
    pub const SERIES_TEXT: u16 = 0x100d;
    pub const FONT: u16 = 0x0031;
    pub const FONT_X: u16 = 0x1026;
    pub const DEFAULT_TEXT: u16 = 0x1024;
    pub const CHART_FORMAT: u16 = 0x1014;
    pub const LEGEND: u16 = 0x1015;
    pub const BAR: u16 = 0x1017;
    pub const LINE: u16 = 0x1018;
    pub const PIE: u16 = 0x1019;
    pub const AREA: u16 = 0x101a;
    pub const SCATTER: u16 = 0x101b;
    pub const CRT_LINE: u16 = 0x101c;
    pub const AXIS: u16 = 0x101d;
    pub const VALUE_RANGE: u16 = 0x101f;
    pub const AXIS_LINE: u16 = 0x1021;
    pub const TEXT: u16 = 0x1025;
    pub const OBJECT_LINK: u16 = 0x1027;
    pub const FRAME: u16 = 0x1032;
    pub const BEGIN: u16 = 0x1033;
    pub const END: u16 = 0x1034;
    pub const PLOT_AREA: u16 = 0x1035;
    pub const CHART3D: u16 = 0x103a;
    pub const RADAR: u16 = 0x103e;
    pub const SURF: u16 = 0x103f;
    pub const RADAR_AREA: u16 = 0x1040;
    pub const AXIS_PARENT: u16 = 0x1041;
    pub const SHT_PROPS: u16 = 0x1044;
    pub const SER_TO_CRT: u16 = 0x1045;
    pub const SER_PARENT: u16 = 0x104a;
    pub const BRAI: u16 = 0x1051;
    pub const SER_FMT: u16 = 0x105d;
    pub const BOP_POP: u16 = 0x1061;
    pub const SI_INDEX: u16 = 0x1065;
    pub const GEL_FRAME: u16 = 0x1066;
    pub const BLANK: u16 = 0x0201;
    pub const NUMBER: u16 = 0x0203;
    pub const LABEL: u16 = 0x0204;
    pub const BOOL_ERR: u16 = 0x0205;
    pub const SHAPE_PROPS_STREAM: u16 = 0x08a4;
    pub const CONTINUE_FRT12: u16 = 0x087f;
}

/// Chart group type from the CRT rule (2.1.7.20.1).
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum GroupKind {
    Bar {
        horizontal: bool,
        stacked: bool,
        percent: bool,
        gap: u16,
        overlap: i16,
    },
    Line {
        stacked: bool,
        percent: bool,
    },
    Area {
        stacked: bool,
        percent: bool,
    },
    Pie {
        start_angle: u16,
        hole: u16,
    },
    Scatter {
        bubbles: bool,
    },
    Radar {
        filled: bool,
    },
    Surface,
    OfPie,
}

#[derive(Debug, Clone)]
pub(crate) struct Group {
    pub kind: GroupKind,
    /// Zero-based AxisParent ordinal (0 primary, 1 secondary).
    pub axis_group: usize,
    pub varied_colors: bool,
    pub three_d: bool,
    pub legend: Option<Legend>,
    /// Group-level default series format (SS in the CRT rule).
    pub default_format: Option<Format>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Legend {
    /// Legend.wType (2.4.152): 0 bottom, 1 corner, 2 top, 3 right, 4 left, 7 not docked.
    pub position: u8,
}

/// One chart element's formatting (FRAME / SS / AXS record sets).
#[derive(Debug, Clone, Default)]
pub(crate) struct Format {
    pub line: Option<[u8; 12]>,
    pub area: Option<[u8; 16]>,
    pub marker: Option<[u8; 20]>,
    pub gel_frame: Option<Vec<u8>>,
    /// Verified ShapePropsStream XML per wObjContext.
    pub shape_xml: BTreeMap<u16, String>,
    pub explosion: Option<u16>,
    pub smooth: bool,
    /// AttachedLabel (2.4.5) flags of a series or point data label.
    pub data_labels: Option<u16>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct Series {
    pub name: Option<String>,
    pub group: u16,
    pub series_format: Option<Format>,
    pub point_formats: BTreeMap<u16, Format>,
    pub trend_or_error: bool,
    /// BRAI (2.4.29) worksheet references by id (0 name, 1 values,
    /// 2 categories, 3 bubble sizes): the ChartParsedFormula rgce bytes.
    pub references: [Option<Vec<u8>>; 4],
}

#[derive(Debug, Clone, Default)]
pub(crate) struct Axis {
    /// Axis.wType: 0 category, 1 value, 2 series.
    pub kind: u16,
    pub axis_group: usize,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub major: Option<f64>,
    pub log: bool,
    pub reversed: bool,
    pub major_gridlines: bool,
    /// FontX.iFont of the axis labels (AXS rule).
    pub font: Option<u16>,
}

#[derive(Debug, Clone)]
pub(crate) enum Cached {
    Number(f64),
    Text(String),
}

#[derive(Debug, Default)]
pub(crate) struct RawChart {
    pub series: Vec<Series>,
    pub groups: Vec<Group>,
    pub axes: Vec<Axis>,
    pub title: Option<String>,
    pub axis_titles: BTreeMap<u16, String>,
    pub chart_format: Option<Format>,
    pub plot_format: Option<Format>,
    pub plot_visible_only: bool,
    /// FontX.iFont of the chart title, legend and chart-wide default text.
    pub title_font: Option<u16>,
    pub legend_font: Option<u16>,
    pub default_font: Option<u16>,
    /// Font records in this chart substream (after FrtFontList, 2.4.123).
    pub local_fonts: Vec<Vec<u8>>,
    /// numIndex (1 values, 2 categories, 3 bubble sizes) -> (series, point).
    pub cache: BTreeMap<u16, BTreeMap<(u16, u16), Cached>>,
}

#[derive(Clone, Copy, PartialEq)]
enum Owner {
    Chart,
    Frame,
    Series(usize),
    DataFormat,
    Text,
    Legend,
    AxisParent,
    Axis,
    ChartFormat,
    Other,
}

struct Block {
    owner: Owner,
    format: Format,
    /// Raw records of this block for the checksum (kind, data).
    records: Vec<(u16, Vec<u8>)>,
    data_format: Option<(u16, u16)>,
    text_link: Option<u16>,
    text: Option<String>,
    group: Option<Group>,
    axis: Option<Axis>,
    legend: Option<Legend>,
    font: Option<u16>,
    default_text: bool,
}

impl Block {
    fn new(owner: Owner) -> Self {
        Self {
            owner,
            format: Format::default(),
            records: Vec::new(),
            data_format: None,
            text_link: None,
            text: None,
            group: None,
            axis: None,
            legend: None,
            font: None,
            default_text: false,
        }
    }
}

fn short_string(bytes: &[u8]) -> Result<String, String> {
    // ShortXLUnicodeString (2.5.240): cch (1 byte), fHighByte, characters.
    let count = usize::from(
        *bytes
            .first()
            .ok_or_else(|| unsupported("truncated chart text"))?,
    );
    let high = bytes
        .get(1)
        .ok_or_else(|| unsupported("truncated chart text"))?
        & 1
        != 0;
    characters(bytes.get(2..).unwrap_or_default(), count, high)
}

fn long_string(bytes: &[u8]) -> Result<String, String> {
    // XLUnicodeString (2.5.294): cch (2 bytes), fHighByte, characters.
    let count = usize::from(u16_at(bytes, 0)?);
    let high = bytes
        .get(2)
        .ok_or_else(|| unsupported("truncated chart label"))?
        & 1
        != 0;
    characters(bytes.get(3..).unwrap_or_default(), count, high)
}

fn characters(bytes: &[u8], count: usize, high: bool) -> Result<String, String> {
    if high {
        let units: Vec<u16> = bytes
            .get(..count * 2)
            .ok_or_else(|| unsupported("truncated chart text characters"))?
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        Ok(String::from_utf16_lossy(&units))
    } else {
        Ok(bytes
            .get(..count)
            .ok_or_else(|| unsupported("truncated chart text characters"))?
            .iter()
            .map(|&byte| char::from(byte))
            .collect())
    }
}

fn f64_at(bytes: &[u8], offset: usize) -> Result<f64, String> {
    Ok(f64::from_le_bytes(
        bytes
            .get(offset..offset + 8)
            .ok_or_else(|| unsupported("truncated chart number"))?
            .try_into()
            .expect("slice length checked"),
    ))
}

/// Read one chart substream: `records` starts at its BOF (dt 0x0020) and ends
/// at the matching EOF.
pub(crate) fn read(records: &[Record<'_>]) -> Result<RawChart, String> {
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty chart substream"))?;
    if first.kind != id::BOF || u16_at(first.data, 2)? != 0x0020 {
        return Err(unsupported(
            "chart substream does not start with a chart BOF",
        ));
    }
    let mut chart = RawChart::default();
    let mut stack: Vec<Block> = Vec::new();
    let mut pending: Option<Owner> = None;
    let mut axis_parent = 0usize;
    let mut axis_parent_seen = false;
    let mut cache_index: Option<u16> = None;
    let mut default_text = false;
    let mut text_bytes = 0usize;
    let mut index = 1usize;
    while index < records.len() {
        let record = records[index];
        index += 1;
        match record.kind {
            id::EOF => return Ok(chart),
            id::BOF => return Err(unsupported("nested BIFF substream inside a chart")),
            id::BEGIN => {
                if stack.len() >= MAX_DEPTH {
                    return Err(unsupported("chart record nesting depth exceeded"));
                }
                let owner = pending.take().unwrap_or(Owner::Other);
                let mut block = Block::new(owner);
                if owner == Owner::Text {
                    block.default_text = std::mem::take(&mut default_text);
                }
                // The record that opens a block was read into its parent; move
                // the element it describes into the new block.
                if let Some(parent) = stack.last_mut() {
                    match owner {
                        Owner::DataFormat => block.data_format = parent.data_format,
                        Owner::ChartFormat => block.group = parent.group.take(),
                        Owner::Axis => block.axis = parent.axis.take(),
                        Owner::Legend => block.legend = parent.legend.take(),
                        _ => {}
                    }
                }
                // Carry the record that opened the block into its checksum set.
                stack.push(block);
                continue;
            }
            id::END => {
                let block = stack
                    .pop()
                    .ok_or_else(|| unsupported("unbalanced chart End"))?;
                finish_block(block, &mut chart, &mut stack, axis_parent)?;
                continue;
            }
            _ => {}
        }
        if let Some(block) = stack.last_mut() {
            if matches!(
                record.kind,
                id::LINE_FORMAT
                    | id::AREA_FORMAT
                    | id::MARKER_FORMAT
                    | id::GEL_FRAME
                    | id::AXIS_LINE
                    | id::CRT_LINE
                    | id::SHAPE_PROPS_STREAM
                    | id::CONTINUE
            ) {
                block.records.push((record.kind, record.data.to_vec()));
            }
        }
        match record.kind {
            id::CHART => pending = Some(Owner::Chart),
            id::FRAME => pending = Some(Owner::Frame),
            id::PLOT_AREA => {
                // The following FRAME belongs to the plot area (2.1.7.20.1 AXES).
                if let Some(block) = stack.last_mut() {
                    block.text_link = Some(0xfffe);
                }
            }
            id::SERIES => {
                if chart.series.len() >= MAX_SERIES {
                    return Err(unsupported("too many chart series"));
                }
                // Series (2.4.252) sdtX: the projection takes category values
                // from the cached cells whatever their type, so only the
                // record's presence is checked.
                u16_at(record.data, 0)?;
                chart.series.push(Series::default());
                pending = Some(Owner::Series(chart.series.len() - 1));
            }
            id::DATA_FORMAT => {
                let point = u16_at(record.data, 0)?;
                let series = u16_at(record.data, 2)?;
                if let Some(block) = stack.last_mut() {
                    block.data_format = Some((point, series));
                }
                pending = Some(Owner::DataFormat);
            }
            id::TEXT => pending = Some(Owner::Text),
            id::DEFAULT_TEXT => default_text = true,
            id::FONT => {
                if chart.local_fonts.len() < 512 {
                    chart.local_fonts.push(record.data.to_vec());
                }
            }
            id::FONT_X => {
                let font = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.font = Some(font);
                    if let Some(axis) = block.axis.as_mut() {
                        axis.font = Some(font);
                    }
                }
            }
            id::LEGEND => {
                let position = *record.data.get(16).unwrap_or(&7);
                if let Some(block) = stack.last_mut() {
                    block.legend = Some(Legend { position });
                }
                pending = Some(Owner::Legend);
            }
            id::AXIS_PARENT => {
                if axis_parent_seen {
                    axis_parent += 1;
                }
                axis_parent_seen = true;
                pending = Some(Owner::AxisParent);
            }
            id::AXIS => {
                let kind = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.axis = Some(Axis {
                        kind,
                        axis_group: axis_parent,
                        ..Axis::default()
                    });
                }
                pending = Some(Owner::Axis);
            }
            id::CHART_FORMAT => {
                let flags = u16_at(record.data, 16)?;
                if let Some(block) = stack.last_mut() {
                    block.group = Some(Group {
                        kind: GroupKind::Line {
                            stacked: false,
                            percent: false,
                        },
                        axis_group: axis_parent,
                        varied_colors: flags & 1 != 0,
                        three_d: false,
                        legend: None,
                        default_format: None,
                    });
                }
                pending = Some(Owner::ChartFormat);
            }
            id::BAR
            | id::LINE
            | id::PIE
            | id::AREA
            | id::SCATTER
            | id::RADAR
            | id::RADAR_AREA
            | id::SURF
            | id::BOP_POP => {
                let kind = group_kind(record)?;
                if let Some(group) = stack
                    .last_mut()
                    .filter(|b| b.owner == Owner::ChartFormat)
                    .and_then(|b| b.group.as_mut())
                {
                    group.kind = kind;
                }
            }
            id::CHART3D => {
                if let Some(group) = stack
                    .last_mut()
                    .filter(|b| b.owner == Owner::ChartFormat)
                    .and_then(|b| b.group.as_mut())
                {
                    group.three_d = true;
                }
            }
            id::SHT_PROPS => {
                let flags = u16_at(record.data, 0)?;
                chart.plot_visible_only = flags & 2 != 0;
            }
            id::SERIES_TEXT => {
                let text = short_string(record.data.get(2..).unwrap_or_default())?;
                text_bytes = text_bytes.saturating_add(text.len());
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(unsupported("chart text budget exceeded"));
                }
                if let Some(block) = stack.last_mut() {
                    block.text = Some(text);
                }
            }
            id::OBJECT_LINK => {
                let link = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.text_link = Some(link);
                }
            }
            id::SER_TO_CRT => {
                let group = u16_at(record.data, 0)?;
                if let Some(Owner::Series(series)) = stack.last().map(|b| b.owner) {
                    chart.series[series].group = group;
                }
            }
            id::BRAI => {
                let part = usize::from(*record.data.first().unwrap_or(&0xff));
                let kind = *record.data.get(1).unwrap_or(&0);
                if let (Some(Owner::Series(series)), true) =
                    (stack.last().map(|b| b.owner), part < 4 && kind == 2)
                {
                    let length = usize::from(u16_at(record.data, 6)?);
                    let rgce = record
                        .data
                        .get(8..8 + length)
                        .ok_or_else(|| unsupported("truncated chart series reference"))?;
                    chart.series[series].references[part] = Some(rgce.to_vec());
                }
            }
            id::SER_PARENT => {
                if let Some(Owner::Series(series)) = stack.last().map(|b| b.owner) {
                    chart.series[series].trend_or_error = true;
                }
            }
            id::LINE_FORMAT => set_fixed(&mut stack, record, |f, v| f.line = Some(v))?,
            id::AREA_FORMAT => set_fixed(&mut stack, record, |f, v| f.area = Some(v))?,
            id::MARKER_FORMAT => set_fixed(&mut stack, record, |f, v| f.marker = Some(v))?,
            id::ATTACHED_LABEL => {
                let flags = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.format.data_labels = Some(flags);
                }
            }
            id::PIE_FORMAT => {
                let value = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.format.explosion = Some(value);
                }
            }
            id::SER_FMT => {
                let flags = u16_at(record.data, 0)?;
                if let Some(block) = stack.last_mut() {
                    block.format.smooth = flags & 1 != 0;
                }
            }
            id::GEL_FRAME => {
                let mut bytes = record.data.to_vec();
                while records.get(index).is_some_and(|r| r.kind == id::CONTINUE) {
                    bytes.extend_from_slice(records[index].data);
                    if let Some(block) = stack.last_mut() {
                        block
                            .records
                            .push((id::CONTINUE, records[index].data.to_vec()));
                    }
                    index += 1;
                }
                if let Some(block) = stack.last_mut() {
                    block.format.gel_frame = Some(bytes);
                }
            }
            id::SHAPE_PROPS_STREAM => {
                // Verified at End, where all related formatting records are known.
                let mut data = record.data.to_vec();
                while records
                    .get(index)
                    .is_some_and(|r| r.kind == id::CONTINUE_FRT12)
                {
                    data.extend_from_slice(records[index].data.get(12..).unwrap_or_default());
                    index += 1;
                }
                if data.len() > MAX_XML_BYTES {
                    return Err(unsupported("chart shape XML budget exceeded"));
                }
                if let Some(block) = stack.last_mut() {
                    if let Some(last) = block.records.last_mut() {
                        last.1 = data;
                    }
                }
            }
            id::VALUE_RANGE => {
                let flags = u16_at(record.data, 40)?;
                let pick = |bit: u16, offset: usize| -> Result<Option<f64>, String> {
                    (flags & bit == 0)
                        .then(|| f64_at(record.data, offset))
                        .transpose()
                };
                let (min, max, major) = (pick(1, 0)?, pick(2, 8)?, pick(4, 16)?);
                if let Some(axis) = stack.last_mut().and_then(|b| b.axis.as_mut()) {
                    axis.min = min;
                    axis.max = max;
                    axis.major = major;
                    axis.log = flags & 0x20 != 0;
                    axis.reversed = flags & 0x40 != 0;
                }
            }
            id::SI_INDEX => cache_index = Some(u16_at(record.data, 0)?),
            id::NUMBER | id::LABEL | id::BOOL_ERR | id::BLANK => {
                let Some(numindex) = cache_index else {
                    continue;
                };
                let point = u16_at(record.data, 0)?;
                let series = u16_at(record.data, 2)?;
                if usize::from(point) >= MAX_POINTS || usize::from(series) >= MAX_SERIES {
                    return Err(unsupported("chart data cache index out of range"));
                }
                let value = match record.kind {
                    id::NUMBER => Some(Cached::Number(f64_at(record.data, 6)?)),
                    id::LABEL => {
                        let text = long_string(record.data.get(6..).unwrap_or_default())?;
                        text_bytes = text_bytes.saturating_add(text.len());
                        if text_bytes > MAX_TEXT_BYTES {
                            return Err(unsupported("chart text budget exceeded"));
                        }
                        Some(Cached::Text(text))
                    }
                    _ => None,
                };
                if let Some(value) = value {
                    chart
                        .cache
                        .entry(numindex)
                        .or_default()
                        .insert((series, point), value);
                }
            }
            _ => {}
        }
    }
    Err(unsupported("chart substream has no EOF"))
}

fn set_fixed<const N: usize>(
    stack: &mut [Block],
    record: Record<'_>,
    assign: impl FnOnce(&mut Format, [u8; N]),
) -> Result<(), String> {
    let bytes: [u8; N] = record
        .data
        .get(..N)
        .ok_or_else(|| unsupported("truncated chart format record"))?
        .try_into()
        .expect("slice length checked");
    if let Some(block) = stack.last_mut() {
        // AXS records AxisLine+LineFormat pairs; keep the axis line (id 0)
        // and note major gridlines, never later pairs over the axis line.
        if block.owner == Owner::Axis && N == 12 {
            let id = block
                .records
                .iter()
                .rev()
                .find(|(kind, _)| *kind == id::AXIS_LINE)
                .and_then(|(_, data)| data.get(..2))
                .map_or(0, |b| u16::from_le_bytes([b[0], b[1]]));
            if id == 1 {
                if let Some(axis) = block.axis.as_mut() {
                    axis.major_gridlines = true;
                }
            }
            if id != 0 {
                return Ok(());
            }
        }
        assign(&mut block.format, bytes);
    }
    Ok(())
}

fn group_kind(record: Record<'_>) -> Result<GroupKind, String> {
    let data = record.data;
    Ok(match record.kind {
        id::BAR => {
            let flags = u16_at(data, 4)?;
            GroupKind::Bar {
                overlap: u16_at(data, 0)? as i16,
                gap: u16_at(data, 2)?,
                horizontal: flags & 1 != 0,
                stacked: flags & 2 != 0,
                percent: flags & 4 != 0,
            }
        }
        id::LINE => {
            let flags = u16_at(data, 0)?;
            GroupKind::Line {
                stacked: flags & 1 != 0,
                percent: flags & 2 != 0,
            }
        }
        id::AREA => {
            let flags = u16_at(data, 0)?;
            GroupKind::Area {
                stacked: flags & 1 != 0,
                percent: flags & 2 != 0,
            }
        }
        id::PIE => GroupKind::Pie {
            start_angle: u16_at(data, 0)?,
            hole: u16_at(data, 2)?,
        },
        id::SCATTER => GroupKind::Scatter {
            bubbles: u16_at(data, 4)? & 1 != 0,
        },
        id::RADAR => GroupKind::Radar { filled: false },
        id::RADAR_AREA => GroupKind::Radar { filled: true },
        id::SURF => GroupKind::Surface,
        _ => GroupKind::OfPie,
    })
}

fn finish_block(
    mut block: Block,
    chart: &mut RawChart,
    stack: &mut [Block],
    _axis_parent: usize,
) -> Result<(), String> {
    verify_shape_xml(&mut block)?;
    match block.owner {
        Owner::Chart => {}
        Owner::Frame => {
            // The frame after PlotArea belongs to the plot area; the first
            // frame directly inside Chart is the chart area (2.1.7.20.1).
            let parent = stack.last_mut();
            match parent {
                Some(parent) if parent.text_link == Some(0xfffe) => {
                    parent.text_link = None;
                    chart.plot_format = Some(block.format);
                }
                Some(parent) if parent.owner == Owner::Chart && chart.chart_format.is_none() => {
                    chart.chart_format = Some(block.format);
                }
                _ => {}
            }
        }
        Owner::Series(index) => {
            chart.series[index].name = block.text.take();
        }
        Owner::DataFormat => {
            if let Some((point, series)) = block.data_format {
                match stack.last().map(|b| b.owner) {
                    Some(Owner::Series(index)) => {
                        let target = &mut chart.series[index];
                        if point == 0xffff {
                            target.series_format = Some(block.format);
                        } else {
                            target.point_formats.insert(point, block.format);
                        }
                    }
                    Some(Owner::ChartFormat) => {
                        if let Some(group) = stack.last_mut().and_then(|b| b.group.as_mut()) {
                            if point == 0xffff && series == 0 {
                                group.default_format = Some(block.format);
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        Owner::Text => {
            if block.default_text && chart.default_font.is_none() {
                chart.default_font = block.font;
            }
            if stack
                .last()
                .is_some_and(|parent| parent.owner == Owner::Legend)
            {
                chart.legend_font = block.font;
            }
            if block.text_link == Some(1) {
                chart.title_font = block.font;
            }
            if let (Some(link), Some(text)) = (block.text_link, block.text.take()) {
                match link {
                    1 => chart.title = Some(text),
                    2 | 3 | 7 => {
                        chart.axis_titles.insert(link, text);
                    }
                    _ => {}
                }
            }
        }
        Owner::Legend => {
            if let Some(group) = stack.last_mut().and_then(|parent| parent.group.as_mut()) {
                group.legend = block.legend;
            }
        }
        Owner::Axis => {
            if let Some(axis) = block.axis.take() {
                chart.axes.push(axis);
            }
        }
        Owner::ChartFormat => {
            if let Some(group) = block.group.take() {
                chart.groups.push(group);
            }
        }
        Owner::AxisParent | Owner::Other => {}
    }
    Ok(())
}

/// Keep ShapePropsStream XML only when its checksum matches the related
/// BIFF formatting records (MS-XLS 2.4.258); otherwise the BIFF records win.
fn verify_shape_xml(block: &mut Block) -> Result<(), String> {
    let owner = block.owner;
    let parent_format = &mut block.format;
    let records = &block.records;
    let area = records
        .iter()
        .find(|(k, _)| *k == id::AREA_FORMAT)
        .map(|r| r.1.as_slice());
    let marker = records
        .iter()
        .find(|(k, _)| *k == id::MARKER_FORMAT)
        .map(|r| r.1.as_slice());
    let gel = records
        .iter()
        .find(|(k, _)| *k == id::GEL_FRAME)
        .map(|r| r.1.as_slice());
    let mut lines = BTreeMap::new();
    let mut current = 0u16;
    for (kind, data) in records {
        if matches!(*kind, id::AXIS_LINE | id::CRT_LINE) {
            current = u16_at(data, 0)?;
        }
        if *kind == id::LINE_FORMAT {
            lines.entry(current).or_insert(data.as_slice());
        }
    }
    let gel_bytes = gel.map(|first| {
        let mut bytes = first.to_vec();
        let start = records
            .iter()
            .position(|(k, _)| *k == id::GEL_FRAME)
            .unwrap_or(0);
        for (kind, data) in &records[start + 1..] {
            if *kind != id::CONTINUE {
                break;
            }
            bytes.extend_from_slice(data);
        }
        bytes
    });
    for (kind, data) in records {
        if *kind != id::SHAPE_PROPS_STREAM || data.len() < 24 {
            continue;
        }
        let context = u16_at(data, 12)?;
        let stored = u32_at(data, 16)?;
        let length = u32_at(data, 20)? as usize;
        // An input u32: unchecked, it can wrap a 32-bit `usize` (wasm32).
        let Some(xml) = 24usize
            .checked_add(length)
            .and_then(|end| data.get(24..end))
        else {
            continue;
        };
        let area_auto = area.map(|a| a.get(10).is_some_and(|flags| flags & 1 != 0));
        let mut input = Vec::new();
        let include_line = match owner {
            Owner::DataFormat => context == 0,
            _ => true,
        };
        let line_key = if matches!(owner, Owner::Axis | Owner::ChartFormat) {
            context
        } else {
            0
        };
        if include_line {
            if let Some(line) = lines
                .get(&line_key)
                .and_then(|l| checksum::line_properties(l))
            {
                input.extend_from_slice(&line);
            }
        }
        let fill_scope = match owner {
            Owner::Axis => context == 3,
            Owner::DataFormat | Owner::Frame | Owner::Other => true,
            _ => false,
        };
        if fill_scope && area_auto == Some(false) {
            if let Some(gel) = gel_bytes.as_deref() {
                fill_style(gel, &mut input)?;
            } else if let Some(area) = area {
                let colors = if owner == Owner::DataFormat && context == 1 {
                    marker.unwrap_or(area)
                } else {
                    area
                };
                if let Some(interior) = checksum::interior_properties(colors, area) {
                    input.extend_from_slice(&interior);
                }
            }
        } else if owner == Owner::DataFormat && context == 1 && gel_bytes.is_none() {
            if let (Some(marker), Some(area)) = (marker, area) {
                if let Some(interior) = checksum::interior_properties(marker, area) {
                    input.extend_from_slice(&interior);
                }
            }
        }
        if checksum::crc(&input) == stored {
            if let Ok(text) = std::str::from_utf8(xml) {
                parent_format.shape_xml.insert(context, text.to_owned());
            }
        }
    }
    Ok(())
}

fn fill_style(gel_frame: &[u8], output: &mut Vec<u8>) -> Result<(), String> {
    let mut work = 1_000_000usize;
    let mut properties: BTreeMap<u16, (u32, Option<Vec<u8>>)> = BTreeMap::new();
    let mut position = 0usize;
    while position < gel_frame.len() {
        let (record, end) =
            crate::officeart::record_with_end(gel_frame, position, &mut work, "chart GelFrame")?;
        let visit = |property: crate::officeart::properties::Property<'_>| {
            properties.insert(
                property.opid & 0x3fff,
                (property.value, property.complex.map(<[u8]>::to_vec)),
            );
            Ok(())
        };
        match record.kind {
            0xf00b => crate::officeart::properties::visit(record, &mut work, visit)?,
            0xf122 => crate::officeart::properties::visit_tertiary(record, &mut work, visit)?,
            _ => {}
        }
        position = end;
    }
    checksum::fill_style_properties(
        |opid| {
            properties
                .get(&opid)
                .map(|(value, complex)| checksum::FillProperty {
                    value: *value,
                    complex: complex.as_deref(),
                })
        },
        properties.get(&0x1bf).map(|p| p.0),
        output,
    );
    Ok(())
}
