//! Excel tables (list objects) of the direct XLS model.
//!
//! MS-XLS 2.4.114 Feature11 / 2.4.115 Feature12 carry a 2.5.266
//! TableFeatureType per table (range, header and total rows, name); the
//! 2.4.157 List12 records that follow add the table style (lsd 1,
//! 2.5.176 List12TableStyleClientInfo). Custom styles are 2.4.320 TableStyle
//! records with 2.4.321 TableStyleElement children in the globals, whose
//! formats are 2.4.97 DXF records (2.5.288 XFProps). They become the XLSX
//! model's `TableInfo` exactly as the XLSX parser builds it from a table
//! part and `<tableStyles>` (ECMA-376 18.5.1, 18.8.82-83).
//!
//! Table-level and per-column differential formats (List12BlockLevel,
//! Feat11FieldDataItem) describe how Excel formats cells it adds to a table;
//! the cells already carry their formats in their XFs, and neither the XLSX
//! model nor its parser carries those formats, so they are not projected.
//! Style elements the model cannot carry reject the table when they would
//! apply (column stripes when shown, header and total corner cells, stripes
//! wider than one row).

use super::{f64_at, styles, theme, u16_at, u32_at, unsupported, Record, EOF};
use std::collections::BTreeMap;

/// Retained records of one worksheet (tables, hyperlinks), in order.
#[derive(Default)]
pub(super) struct Records {
    records: Vec<(u16, Vec<u8>)>,
    bytes: usize,
}

const MAX_SHEET_BYTES: usize = 16 * 1024 * 1024;

impl Records {
    pub(super) fn push(&mut self, kind: u16, data: &[u8]) -> Result<(), String> {
        self.bytes = self
            .bytes
            .checked_add(data.len())
            .filter(|bytes| *bytes <= MAX_SHEET_BYTES)
            .ok_or_else(|| unsupported("XLS table byte budget exceeded"))?;
        self.records.push((kind, data.to_vec()));
        Ok(())
    }

    pub(super) fn is_empty(&self) -> bool {
        self.records.is_empty()
    }

    pub(super) fn iter(&self) -> impl Iterator<Item = (u16, &[u8])> {
        self.records
            .iter()
            .map(|(kind, data)| (*kind, data.as_slice()))
    }
}

fn truncated() -> String {
    unsupported("truncated XLS table record")
}

/// TableStyleElement indexes (DXFId) of one custom table style.
#[derive(Default, Clone)]
struct Elements {
    by_type: BTreeMap<u32, (u32, u32)>,
}

/// Workbook-global table styles and the DXF records they reference.
#[derive(Default)]
pub(super) struct Styles {
    dxf_records: Vec<Vec<u8>>,
    styles: BTreeMap<String, Elements>,
    /// DXF record index -> workbook dxf id, filled as tables use them.
    projected: BTreeMap<u32, u32>,
}

fn utf16(data: &[u8], offset: usize, count: usize) -> Result<String, String> {
    let bytes = data.get(offset..offset + count * 2).ok_or_else(truncated)?;
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    String::from_utf16(&units).map_err(|_| truncated())
}

/// XLUnicodeString (2.5.294) at `offset`; returns the text and its size.
fn unicode_string(data: &[u8], offset: usize) -> Result<(String, usize), String> {
    let count = usize::from(u16_at(data, offset)?);
    match *data.get(offset + 2).ok_or_else(truncated)? {
        0 => Ok((
            data.get(offset + 3..offset + 3 + count)
                .ok_or_else(truncated)?
                .iter()
                .map(|&byte| char::from(byte))
                .collect(),
            3 + count,
        )),
        1 => Ok((utf16(data, offset + 3, count)?, 3 + count * 2)),
        _ => Err(unsupported("invalid XLS table string")),
    }
}

impl Styles {
    pub(super) fn parse(records: &[Record<'_>]) -> Result<Self, String> {
        let mut result = Self::default();
        let mut pending: Option<(String, Elements, usize)> = None;
        for record in records.iter().take_while(|r| r.kind != EOF) {
            let data = record.data;
            if let Some((_, elements, left)) = pending.as_mut() {
                if *left > 0 {
                    if record.kind != 0x0890 || u16_at(data, 0)? != 0x0890 || data.len() != 24 {
                        return Err(unsupported("XLS table style lacks its elements"));
                    }
                    let kind = u32_at(data, 12)?;
                    let size = u32_at(data, 16)?;
                    let index = u32_at(data, 20)?;
                    if kind > 0x1b || elements.by_type.insert(kind, (size, index)).is_some() {
                        return Err(unsupported("invalid XLS table style element"));
                    }
                    *left -= 1;
                    continue;
                }
                let (name, elements, _) = pending.take().expect("pending style");
                if result.styles.insert(name, elements).is_some() {
                    return Err(unsupported("duplicate XLS table style"));
                }
            }
            match record.kind {
                0x088d => {
                    if u16_at(data, 0)? != 0x088d || data.len() < 18 {
                        return Err(unsupported("invalid XLS DXF record"));
                    }
                    result.dxf_records.push(data.to_vec());
                }
                0x088f => {
                    if u16_at(data, 0)? != 0x088f {
                        return Err(unsupported("invalid XLS table style"));
                    }
                    let count = usize::try_from(u32_at(data, 14)?).map_err(|_| truncated())?;
                    let chars = usize::from(u16_at(data, 18)?);
                    if count > 28 || chars == 0 || data.len() != 20 + chars * 2 {
                        return Err(unsupported("invalid XLS table style"));
                    }
                    pending = Some((utf16(data, 20, chars)?, Elements::default(), count));
                }
                0x0890 => return Err(unsupported("orphan XLS table style element")),
                _ => {}
            }
        }
        if let Some((name, elements, left)) = pending {
            if left > 0 {
                return Err(unsupported("XLS table style lacks its elements"));
            }
            if result.styles.insert(name, elements).is_some() {
                return Err(unsupported("duplicate XLS table style"));
            }
        }
        Ok(result)
    }
}

pub(super) struct Context<'a> {
    pub(super) styles: &'a styles::Styles<'a>,
    pub(super) theme: &'a theme::Colors,
}

/// One table from Feature11/Feature12 before its List12 style arrives.
struct Pending {
    id: u32,
    range: xlsx_model::CellRange,
    header: bool,
    totals: bool,
    style: Option<(String, u8)>,
}

/// Project the tables of one worksheet; custom-style formats are appended
/// to the workbook `dxfs`.
pub(super) fn project(
    records: &Records,
    table_styles: &mut Styles,
    context: &Context<'_>,
    dxfs: &mut Vec<xlsx_model::Dxf>,
) -> Result<Vec<xlsx_model::TableInfo>, String> {
    let mut pending: Vec<Pending> = Vec::new();
    for (kind, data) in &records.records {
        match *kind {
            // FeatHdr11: common header of the sheet's table collection.
            0x0871 => {
                if u16_at(data, 0)? != 0x0871 || u16_at(data, 12)? != 5 {
                    return Err(unsupported("invalid XLS table header"));
                }
            }
            0x0872 | 0x0878 => pending.push(feature(*kind, data)?),
            0x0877 => {
                if u16_at(data, 0)? != 0x0877 {
                    return Err(unsupported("invalid XLS table List12"));
                }
                let lsd = u16_at(data, 12)?;
                let id = u32_at(data, 14)?;
                let table = pending
                    .last_mut()
                    .filter(|table| table.id == id)
                    .ok_or_else(|| unsupported("XLS List12 names no preceding table"))?;
                if lsd == 1 {
                    // List12TableStyleClientInfo: flags, then stListStyleName.
                    let flags = *data.get(18).ok_or_else(truncated)?;
                    let (name, size) = unicode_string(data, 20)?;
                    // stListStyleName MUST be non-empty, yet Excel writes an
                    // empty name for a table without a style; the Excel-saved
                    // counterpart (sample-4) has `tableStyleInfo` with only
                    // the show* flags, i.e. no style.
                    if 20 + size != data.len() || table.style.is_some() {
                        return Err(unsupported("invalid XLS table style information"));
                    }
                    table.style = Some((name, flags));
                } else if lsd > 2 {
                    return Err(unsupported("invalid XLS table List12 type"));
                }
                // lsd 0 (List12BlockLevel) and 2 (List12DisplayName): see
                // the module note; neither changes the table's display.
            }
            _ => return Err(unsupported("unexpected XLS table record")),
        }
    }
    let mut output = Vec::with_capacity(pending.len());
    for table in pending {
        output.push(table_info(table, table_styles, context, dxfs)?);
    }
    Ok(output)
}

/// Feature11/Feature12 header (2.4.114) and the fixed part of its
/// TableFeatureType (2.5.266).
fn feature(kind: u16, data: &[u8]) -> Result<Pending, String> {
    if u16_at(data, 0)? != kind || u16_at(data, 12)? != 5 {
        return Err(unsupported("invalid XLS table feature"));
    }
    let count = usize::from(u16_at(data, 19)?);
    if count != 1 {
        return Err(unsupported("XLS table with other than one range"));
    }
    let (row_first, row_last) = (u16_at(data, 27)?, u16_at(data, 29)?);
    let (col_first, col_last) = (u16_at(data, 31)?, u16_at(data, 33)?);
    if row_first > row_last || col_first > col_last || col_last > 0x00ff {
        return Err(unsupported("invalid XLS table range"));
    }
    let feat = 35;
    let id = u32_at(data, feat + 4)?;
    let header = u32_at(data, feat + 8)?;
    let totals = u32_at(data, feat + 12)?;
    if u32_at(data, feat + 20)? != 64 || header > 1 || totals > 1 || id == 0 {
        return Err(unsupported("invalid XLS table definition"));
    }
    Ok(Pending {
        id,
        range: xlsx_model::CellRange {
            top: u32::from(row_first) + 1,
            left: u32::from(col_first) + 1,
            bottom: u32::from(row_last) + 1,
            right: u32::from(col_last) + 1,
        },
        header: header == 1,
        totals: totals == 1,
        style: None,
    })
}

/// ECMA-376 18.5.1.4 built-in style accent, as the XLSX parser resolves it:
/// `TableStyle{Light|Medium|Dark}N` uses accent ((N - 1) mod 7), 0 = none.
fn builtin_accent(name: &str, theme: &theme::Colors) -> String {
    let fallback = "#808080".to_string();
    let Some(rest) = name.strip_prefix("TableStyle") else {
        return fallback;
    };
    let Some(start) = rest.find(|c: char| c.is_ascii_digit()) else {
        return fallback;
    };
    let Ok(number) = rest[start..].parse::<u32>() else {
        return fallback;
    };
    let slot = number.checked_sub(1).map(|n| n % 7).unwrap_or(0);
    if slot == 0 {
        return fallback;
    }
    // clrScheme order: accent1 is slot 4.
    theme
        .argb(3 + slot)
        .map(|[_, r, g, b]| format!("#{r:02X}{g:02X}{b:02X}"))
        .unwrap_or(fallback)
}

fn table_info(
    table: Pending,
    table_styles: &mut Styles,
    context: &Context<'_>,
    dxfs: &mut Vec<xlsx_model::Dxf>,
) -> Result<xlsx_model::TableInfo, String> {
    let (style_name, flags) = table.style.unwrap_or_default();
    let first_column = flags & 0x01 != 0;
    let last_column = flags & 0x02 != 0;
    let row_stripes = flags & 0x04 != 0;
    let column_stripes = flags & 0x08 != 0;
    let mut info = xlsx_model::TableInfo {
        range: table.range,
        accent_color: builtin_accent(&style_name, context.theme),
        style_name: style_name.clone(),
        header_row_count: u32::from(table.header),
        totals_row_count: u32::from(table.totals),
        show_row_stripes: row_stripes,
        show_column_stripes: column_stripes,
        show_first_column: first_column,
        show_last_column: last_column,
        is_custom: false,
        whole_table_dxf: None,
        header_row_dxf: None,
        total_row_dxf: None,
        first_column_dxf: None,
        last_column_dxf: None,
        band1_horizontal_dxf: None,
        band2_horizontal_dxf: None,
        columns: Vec::new(),
    };
    let Some(elements) = table_styles.styles.get(&style_name).cloned() else {
        return Ok(info);
    };
    info.is_custom = true;
    for (&kind, &(size, index)) in &elements.by_type {
        let applies = match kind {
            7 | 8 => column_stripes,
            9 | 10 => table.header,
            11 | 12 => table.totals,
            _ => false,
        };
        if applies {
            return Err(unsupported("XLS table style element is not representable"));
        }
        if matches!(kind, 5 | 6) && row_stripes && size != 1 {
            return Err(unsupported("XLS table stripe wider than one row"));
        }
        let slot = match kind {
            0 => &mut info.whole_table_dxf,
            1 => &mut info.header_row_dxf,
            2 => &mut info.total_row_dxf,
            3 => &mut info.first_column_dxf,
            4 => &mut info.last_column_dxf,
            5 => &mut info.band1_horizontal_dxf,
            6 => &mut info.band2_horizontal_dxf,
            _ => continue,
        };
        *slot = Some(table_styles.dxf(index, context, dxfs)?);
    }
    Ok(info)
}

impl Styles {
    fn dxf(
        &mut self,
        index: u32,
        context: &Context<'_>,
        dxfs: &mut Vec<xlsx_model::Dxf>,
    ) -> Result<u32, String> {
        if let Some(id) = self.projected.get(&index) {
            return Ok(*id);
        }
        // DXFId (2.5.94): zero-based over the globals' DXF records. The
        // private corpus confirms it: sample-5's styles match the Excel-saved
        // .xlsx dxf for dxf, and in sample-2, whose .xlsx reorders the style
        // formats, Excel's PDF of the .xls shows the zero-based header format
        // (dark red text, no fill) rather than the .xlsx one.
        let data = self
            .dxf_records
            .get(index as usize)
            .ok_or_else(|| unsupported("XLS table style names no DXF record"))?;
        let dxf = xfprops(data, 14, context)?;
        if dxfs.len() >= super::conditional::MAX_DXFS {
            return Err(unsupported("too many XLS differential formats"));
        }
        dxfs.push(dxf);
        let id = u32::try_from(dxfs.len() - 1).map_err(|_| truncated())?;
        self.projected.insert(index, id);
        Ok(id)
    }
}

/// XFPropColor (2.5.285): dwRgba holds the color Excel derived from the
/// type, index and tint (fValidRGBA MUST be 1); xclrType 0 is automatic.
fn color(data: &[u8], offset: usize) -> Result<Option<String>, String> {
    let head = *data.get(offset).ok_or_else(truncated)?;
    let kind = head >> 1;
    if head & 1 == 0 || kind > 3 {
        return Err(unsupported("invalid XLS DXF color"));
    }
    if kind == 0 {
        return Ok(None);
    }
    let [r, g, b] = [
        *data.get(offset + 4).ok_or_else(truncated)?,
        *data.get(offset + 5).ok_or_else(truncated)?,
        *data.get(offset + 6).ok_or_else(truncated)?,
    ];
    Ok(Some(format!("#{r:02X}{g:02X}{b:02X}")))
}

/// The dxf font, created with the XLSX parser's defaults (size 11).
fn font_entry(font: &mut Option<xlsx_model::Font>) -> &mut xlsx_model::Font {
    font.get_or_insert_with(|| xlsx_model::Font {
        size: 11.0,
        ..Default::default()
    })
}

/// XFProps (2.5.288) of a DXF record as the XLSX model `Dxf`, with the
/// XLSX parser's defaults for omitted children (font size 11, solid fill,
/// a lone background mirrored into the foreground).
fn xfprops(data: &[u8], offset: usize, context: &Context<'_>) -> Result<xlsx_model::Dxf, String> {
    let count = usize::from(u16_at(data, offset + 2)?);
    let mut at = offset + 4;
    let mut dxf = xlsx_model::Dxf::default();
    let mut font: Option<xlsx_model::Font> = None;
    let mut fill: Option<xlsx_model::Fill> = None;
    let mut border: Option<xlsx_model::Border> = None;
    let mut format_code = None;
    let mut format_id = None;
    let mut seen = std::collections::BTreeSet::new();
    let mut gradient_stops = false;
    for _ in 0..count {
        let kind = u16_at(data, at)?;
        let size = usize::from(u16_at(data, at + 2)?);
        if size < 4 || at + size > data.len() || (kind != 4 && !seen.insert(kind)) {
            return Err(unsupported("invalid XLS DXF property"));
        }
        let value = &data[at + 4..at + size];
        at += size;
        let byte = || value.first().copied().ok_or_else(truncated);
        match kind {
            0 => {
                let pattern = styles::PATTERNS
                    .get(usize::from(byte()?))
                    .ok_or_else(|| unsupported("invalid XLS DXF fill pattern"))?;
                fill.get_or_insert_with(|| xlsx_model::Fill {
                    pattern_type: "solid".into(),
                    ..Default::default()
                })
                .pattern_type = (*pattern).to_string();
            }
            1 | 2 => {
                let color = color(value, 0)?;
                let fill = fill.get_or_insert_with(|| xlsx_model::Fill {
                    pattern_type: "solid".into(),
                    ..Default::default()
                });
                if kind == 1 {
                    fill.fg_color = color;
                } else {
                    fill.bg_color = color;
                }
            }
            5 => {
                let target = font_entry(&mut font);
                target.color = color(value, 0)?;
            }
            6..=12 => {
                let style = u16_at(value, 8)?;
                let edge = if style == 0 {
                    None
                } else {
                    Some(xlsx_model::BorderEdge {
                        style: styles::BORDERS
                            .get(usize::from(style))
                            .ok_or_else(|| unsupported("invalid XLS DXF border style"))?
                            .to_string(),
                        color: color(value, 0)?,
                    })
                };
                let border = border.get_or_insert_with(Default::default);
                match kind {
                    6 => border.top = edge,
                    7 => border.bottom = edge,
                    8 => border.left = edge,
                    9 => border.right = edge,
                    // Diagonal lines need the up/down flags (0x0D/0x0E).
                    10 => {
                        border.diagonal_up = edge.clone();
                        border.diagonal_down = edge;
                    }
                    11 => border.vertical = edge,
                    _ => border.horizontal = edge,
                }
            }
            13 | 14 => {
                if byte()? > 1 {
                    return Err(unsupported("invalid XLS DXF diagonal flag"));
                }
            }
            // Alignment, merge, font family, character set, font scheme,
            // outline/shadow/condense/extend, indentation and protection:
            // not carried by the XLSX model `Dxf` (nor by its parser).
            15..=17 | 19..=23 | 0x1e..=0x23 | 0x25 | 0x2a..=0x2c => {}
            0x12 => {}
            0x18 => {
                let chars = usize::from(u16_at(value, 0)?);
                let name = utf16(value, 2, chars)?;
                font_entry(&mut font).name = Some(name);
            }
            0x19 => {
                let weight = u16_at(value, 0)?;
                font_entry(&mut font).bold = weight >= 0x02bc;
            }
            0x1a => {
                let (underline, style) = match byte()? {
                    0x00 => (false, None),
                    0x01 => (true, None),
                    0x02 => (true, Some("double")),
                    0x21 => (true, Some("singleAccounting")),
                    0x22 => (true, Some("doubleAccounting")),
                    _ => return Err(unsupported("invalid XLS DXF underline")),
                };
                let target = font_entry(&mut font);
                target.underline = underline;
                target.underline_style = style.map(str::to_string);
            }
            0x1b => {
                let target = font_entry(&mut font);
                target.vert_align = match u16_at(value, 0)? {
                    0 => None,
                    1 => Some("superscript".into()),
                    2 => Some("subscript".into()),
                    _ => return Err(unsupported("invalid XLS DXF script")),
                };
            }
            0x1c | 0x1d => {
                let on = byte()? != 0;
                let target = font_entry(&mut font);
                if kind == 0x1c {
                    target.italic = on;
                } else {
                    target.strike = on;
                }
            }
            0x24 => {
                let twips = u32_at(value, 0)?;
                font_entry(&mut font).size = f64::from(twips) / 20.0;
            }
            0x26 => format_code = Some(unicode_string(value, 0)?.0),
            0x29 => format_id = Some(u16_at(value, 0)?),
            // XFPropGradient (2.5.286) as ECMA-376 18.8.24 gradientFill; its
            // XFPropGradientStop (2.5.287) entries follow. Like the XLSX
            // parser's gradient fills, it has no pattern type.
            3 => {
                if fill.is_some() {
                    return Err(unsupported("invalid XLS DXF gradient fill"));
                }
                let kind = u32_at(value, 0)?;
                let number = |at: usize| -> Result<f64, String> {
                    let value = f64_at(value, at)?;
                    value
                        .is_finite()
                        .then_some(value)
                        .ok_or_else(|| unsupported("invalid XLS DXF gradient"))
                };
                if kind > 1 || value.len() != 44 {
                    return Err(unsupported("invalid XLS DXF gradient"));
                }
                fill = Some(xlsx_model::Fill {
                    gradient: Some(xlsx_model::GradientFillSpec {
                        gradient_type: if kind == 0 { "linear" } else { "path" }.into(),
                        degree: number(4)?,
                        left: number(12)?,
                        right: number(20)?,
                        top: number(28)?,
                        bottom: number(36)?,
                        stops: Vec::new(),
                    }),
                    ..Default::default()
                });
                gradient_stops = true;
                continue;
            }
            4 => {
                let gradient = fill
                    .as_mut()
                    .and_then(|fill| fill.gradient.as_mut())
                    .filter(|_| value.len() == 18)
                    .ok_or_else(|| unsupported("invalid XLS DXF gradient stop"))?;
                let position = f64_at(value, 2)?;
                if !(0.0..=1.0).contains(&position) {
                    return Err(unsupported("invalid XLS DXF gradient stop"));
                }
                gradient.stops.push(xlsx_model::GradientStopSpec {
                    position,
                    color: color(value, 10)?
                        .ok_or_else(|| unsupported("automatic XLS DXF gradient color"))?,
                });
            }
            _ => return Err(unsupported("unsupported XLS DXF property")),
        }
    }
    if at != data.len() {
        return Err(unsupported("unexpected XLS DXF tail"));
    }
    // XFProps: a gradient excludes a pattern and carries its stops.
    if gradient_stops
        && (seen.contains(&0)
            || fill
                .as_ref()
                .and_then(|fill| fill.gradient.as_ref())
                .is_none_or(|gradient| gradient.stops.len() < 2))
    {
        return Err(unsupported("invalid XLS DXF gradient fill"));
    }
    if format_code.is_some() || format_id.is_some() {
        let id = format_id.unwrap_or_default();
        dxf.num_fmt = Some(xlsx_model::NumFmt {
            num_fmt_id: id.into(),
            format_code: format_code
                .or_else(|| context.styles.format_code(id))
                .unwrap_or_default(),
        });
    }
    if let Some(fill) = fill.as_mut().filter(|fill| fill.gradient.is_none()) {
        if fill.fg_color.is_none() {
            fill.fg_color = fill.bg_color.clone();
        }
    }
    dxf.font = font;
    dxf.fill = fill;
    dxf.border = border;
    Ok(dxf)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn frt(kind: u16) -> Vec<u8> {
        let mut data = vec![0u8; 12];
        data[..2].copy_from_slice(&kind.to_le_bytes());
        data
    }

    fn feature(id: u32, rows: (u16, u16), cols: (u16, u16)) -> Vec<u8> {
        let mut data = frt(0x0872);
        data.extend(5u16.to_le_bytes());
        data.extend([0u8; 5]);
        data.extend(1u16.to_le_bytes());
        data.extend(0u32.to_le_bytes());
        data.extend(0u16.to_le_bytes());
        for value in [rows.0, rows.1, cols.0, cols.1] {
            data.extend(value.to_le_bytes());
        }
        // TableFeatureType: lt, idList, crwHeader, crwTotals, idFieldNext, cbFSData.
        for value in [0u32, id, 1, 0, 3, 64] {
            data.extend(value.to_le_bytes());
        }
        data
    }

    fn style_info(id: u32, flags: u8, name: &str) -> Vec<u8> {
        let mut data = frt(0x0877);
        data.extend(1u16.to_le_bytes());
        data.extend(id.to_le_bytes());
        data.extend([flags, 0]);
        data.extend((name.len() as u16).to_le_bytes());
        data.push(0);
        data.extend(name.as_bytes());
        data
    }

    fn records(list: &[Vec<u8>]) -> Records {
        let mut records = Records::default();
        for data in list {
            records
                .push(u16::from_le_bytes([data[0], data[1]]), data)
                .unwrap();
        }
        records
    }

    #[test]
    fn custom_table_styles_resolve_their_element_formats() {
        // DXF 0: solid gray fill (XFProp 0x02 background, RGBA F2F2F2).
        let mut dxf = frt(0x088d);
        dxf.extend([0, 0]);
        dxf.extend([0, 0, 1, 0]);
        dxf.extend(2u16.to_le_bytes());
        dxf.extend(12u16.to_le_bytes());
        dxf.extend([0x05, 0, 0, 0, 0xf2, 0xf2, 0xf2, 0xff]);
        let mut style = frt(0x088f);
        style.extend([0x04, 0]);
        style.extend(1u32.to_le_bytes());
        style.extend(1u16.to_le_bytes());
        style.extend(u16::from(b'S').to_le_bytes());
        let mut element = frt(0x0890);
        for value in [5u32, 1, 0] {
            element.extend(value.to_le_bytes());
        }
        let globals = [
            Record {
                kind: 0x088d,
                offset: 0,
                data: &dxf,
            },
            Record {
                kind: 0x088f,
                offset: 1,
                data: &style,
            },
            Record {
                kind: 0x0890,
                offset: 2,
                data: &element,
            },
        ];
        let mut table_styles = Styles::parse(&globals).unwrap();
        let styles = styles::Styles::parse(&[]).unwrap();
        let theme = theme::Colors::default();
        let context = Context {
            styles: &styles,
            theme: &theme,
        };
        let mut dxfs = Vec::new();
        let tables = project(
            &records(&[feature(1, (8, 35), (1, 6)), style_info(1, 0x05, "S")]),
            &mut table_styles,
            &context,
            &mut dxfs,
        )
        .unwrap();
        let table = &tables[0];
        assert_eq!(
            (
                table.range.top,
                table.range.left,
                table.range.bottom,
                table.range.right
            ),
            (9, 2, 36, 7)
        );
        assert!(table.is_custom && table.show_row_stripes && table.show_first_column);
        assert_eq!(table.band1_horizontal_dxf, Some(0));
        assert_eq!(
            dxfs[0].fill.as_ref().unwrap().fg_color.as_deref(),
            Some("#F2F2F2")
        );
        // Column stripes shown by a custom style that defines them reject.
        let mut element = frt(0x0890);
        for value in [7u32, 1, 0] {
            element.extend(value.to_le_bytes());
        }
        let globals = [
            Record {
                kind: 0x088d,
                offset: 0,
                data: &dxf,
            },
            Record {
                kind: 0x088f,
                offset: 1,
                data: &style,
            },
            Record {
                kind: 0x0890,
                offset: 2,
                data: &element,
            },
        ];
        let mut table_styles = Styles::parse(&globals).unwrap();
        assert!(project(
            &records(&[feature(1, (8, 35), (1, 6)), style_info(1, 0x08, "S")]),
            &mut table_styles,
            &context,
            &mut Vec::new(),
        )
        .is_err());
        // A List12 for another table is rejected.
        assert!(project(
            &records(&[feature(1, (8, 35), (1, 6)), style_info(2, 0, "S")]),
            &mut Styles::default(),
            &context,
            &mut Vec::new(),
        )
        .is_err());
    }
}
