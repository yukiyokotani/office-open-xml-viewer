//! PivotTable views of the direct XLS model.
//!
//! A PivotTable view is the MS-XLS 2.1.7.20.5 PIVOTVIEW record collection
//! of a worksheet substream: 2.4.313 SxView (location, axes, counts, name),
//! 2.4.292 SxIvd (the row and column fields), 2.4.298 SXPI (page fields),
//! 2.4.278 SXDI (data items), 2.4.293 SXLI (the row and column pivot lines,
//! 2.5.259 SXLIItem) and, in PIVOTADDL, 2.4.273.107
//! SXAddl_SXCView_SXDTableStyleClient (the applied table style and its
//! options). They become the XLSX model's `PivotTableMetadata` as the XLSX
//! parser builds it from a pivotTableDefinition (ECMA-376 18.10.1.73):
//! metadata plus the style and pivot lines the renderer draws the
//! PivotTable style regions from. Saved cells carry the values and cell
//! formats; the PivotCache (the `_SX_DB_CUR` storage) is not read, so the
//! metadata is partial like the XLSX parser's.
//!
//! Formatting the renderer cannot draw fails closed: a PivotTable AutoFormat
//! (fAutoFormat), OLAP views (SXViewEx), a built-in style in a theme-less
//! workbook, and page-field or column-subheading style elements that would
//! apply.

use super::{tables, u16_at, unsupported};

fn truncated() -> String {
    unsupported("truncated XLS PivotTable record")
}

fn i16_at(data: &[u8], offset: usize) -> Result<i16, String> {
    Ok(u16_at(data, offset)? as i16)
}

/// XLUnicodeStringNoCch (2.5.296) of `count` characters at `offset`.
fn string_no_cch(data: &[u8], offset: usize, count: usize) -> Result<(String, usize), String> {
    let flag = *data.get(offset).ok_or_else(truncated)?;
    match flag {
        0 => Ok((
            data.get(offset + 1..offset + 1 + count)
                .ok_or_else(truncated)?
                .iter()
                .map(|&byte| char::from(byte))
                .collect(),
            1 + count,
        )),
        1 => {
            let bytes = data
                .get(offset + 1..offset + 1 + count * 2)
                .ok_or_else(truncated)?;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            Ok((
                String::from_utf16(&units).map_err(|_| truncated())?,
                1 + count * 2,
            ))
        }
        _ => Err(unsupported("invalid XLS PivotTable string")),
    }
}

/// Worksheet records of the PivotTable views, in stream order, with each
/// Continue record of an SxIvd, SXPI or SXLI record appended to it.
pub(super) const RECORDS: [u16; 7] = [0x00b0, 0x00b4, 0x00b5, 0x00b6, 0x00c5, 0x0864, 0x080c];

/// Records whose Continue records extend them (PIVOTIVD, PIVOTPI, PIVOTLI).
pub(super) fn continued(kind: u16) -> bool {
    matches!(kind, 0x00b4..=0x00b6)
}

/// One PivotTable view's records.
#[derive(Default)]
struct View {
    view: Vec<u8>,
    ivd: Vec<Vec<u8>>,
    pages: Option<Vec<u8>>,
    data: Vec<Vec<u8>>,
    lines: Vec<Vec<u8>>,
    style: Option<Vec<u8>>,
    olap: bool,
}

fn views(records: &tables::Records) -> Result<Vec<View>, String> {
    let mut views: Vec<View> = Vec::new();
    let mut last: Option<u16> = None;
    for (kind, data) in records.iter() {
        if kind == 0x003c {
            let target = match (last, views.last_mut()) {
                (Some(0x00b4), Some(view)) => view.ivd.last_mut(),
                (Some(0x00b5), Some(view)) => view.lines.last_mut(),
                (Some(0x00b6), Some(view)) => view.pages.as_mut(),
                _ => None,
            };
            target
                .ok_or_else(|| unsupported("orphan XLS PivotTable Continue record"))?
                .extend_from_slice(data);
            continue;
        }
        last = Some(kind);
        if kind == 0x00b0 {
            views.push(View {
                view: data.to_vec(),
                ..View::default()
            });
            continue;
        }
        let view = views
            .last_mut()
            .ok_or_else(|| unsupported("XLS PivotTable record outside a view"))?;
        match kind {
            0x00b4 => view.ivd.push(data.to_vec()),
            0x00b5 => view.lines.push(data.to_vec()),
            0x00b6 => {
                if view.pages.replace(data.to_vec()).is_some() {
                    return Err(unsupported("duplicate XLS PivotTable page fields"));
                }
            }
            0x00c5 => view.data.push(data.to_vec()),
            0x080c => view.olap = true,
            // SXAddlHdr (2.5.249): frtHeaderOld, then sxc and sxd.
            0x0864 if data.get(4..6) == Some(&[0x00, 0x1e]) => {
                view.style.replace(data.to_vec()).map_or(Ok(()), |_| {
                    Err(unsupported("duplicate XLS PivotTable style"))
                })?
            }
            _ => {}
        }
    }
    Ok(views)
}

/// ST_ItemType (ECMA-376 18.18.43) of an SXLIItem itmType (2.5.259).
const ITEM_TYPES: [&str; 15] = [
    "data", "default", "sum", "countA", "count", "avg", "max", "min", "product", "stdDev",
    "stdDevP", "var", "varP", "grand", "blank",
];

/// ST_DataConsolidateFunction (18.18.17) of an SXDI iiftab.
const SUBTOTALS: [&str; 11] = [
    "sum",
    "count",
    "average",
    "max",
    "min",
    "product",
    "countNums",
    "stdDev",
    "stdDevp",
    "var",
    "varp",
];

/// ST_TableStyleType (18.18.82) of a TableStyleElement tseType (2.4.321)
/// applied to a PivotTable; 0x0A-0x0C MUST be ignored for PivotTables.
fn element_kind(tse: u32) -> Option<&'static str> {
    Some(match tse {
        0x00 => "wholeTable",
        0x01 => "headerRow",
        0x02 => "totalRow",
        0x03 => "firstColumn",
        0x04 => "lastColumn",
        0x05 => "firstRowStripe",
        0x06 => "secondRowStripe",
        0x07 => "firstColumnStripe",
        0x08 => "secondColumnStripe",
        0x09 => "firstHeaderCell",
        0x0d => "firstSubtotalColumn",
        0x0e => "secondSubtotalColumn",
        0x0f => "thirdSubtotalColumn",
        0x10 => "firstSubtotalRow",
        0x11 => "secondSubtotalRow",
        0x12 => "thirdSubtotalRow",
        0x13 => "blankRow",
        0x14 => "firstColumnSubheading",
        0x15 => "secondColumnSubheading",
        0x16 => "thirdColumnSubheading",
        0x17 => "firstRowSubheading",
        0x18 => "secondRowSubheading",
        0x19 => "thirdRowSubheading",
        0x1a => "pageFieldLabels",
        0x1b => "pageFieldValues",
        _ => return None,
    })
}

/// SXLI pivot lines (2.4.293): SXLIItem structures of `entries` pivot line
/// entries each. The depth is isxviMac less one, as the XLSX `i` element's
/// `r` plus its `x` count less one: sample-2's .xls lines match its
/// Excel-saved .xlsx rowItems (cSic = r, isxviMac - cSic = x count).
fn lines(
    data: &[u8],
    count: usize,
    entries: usize,
) -> Result<Vec<xlsx_model::PivotAxisItem>, String> {
    let size = 8 + entries * 2;
    // Two u16 counts: the product can exceed a 32-bit `usize` (wasm32).
    if count.checked_mul(size) != Some(data.len()) {
        return Err(unsupported("invalid XLS PivotTable lines"));
    }
    data.chunks_exact(size)
        .map(|item| {
            let kind = u16_at(item, 2)? & 0x7fff;
            let shown = i16_at(item, 4)?;
            let kind = ITEM_TYPES
                .get(usize::from(kind))
                .filter(|_| shown >= 0)
                .ok_or_else(|| unsupported("invalid XLS PivotTable line"))?;
            Ok(xlsx_model::PivotAxisItem {
                kind: (*kind).to_string(),
                depth: u32::from(shown.max(1).unsigned_abs()) - 1,
            })
        })
        .collect()
}

fn fields(data: &[u8], count: usize) -> Result<Vec<i32>, String> {
    if data.len() != count * 2 {
        return Err(unsupported("invalid XLS PivotTable axis fields"));
    }
    (0..count)
        .map(|i| Ok(i32::from(i16_at(data, i * 2)?)))
        .collect()
}

fn feature(name: &str) -> xlsx_model::PivotPartialReason {
    xlsx_model::PivotPartialReason::UnsupportedSemanticFeature {
        feature: name.to_string(),
    }
}

/// Project the PivotTable views of one worksheet.
pub(super) fn project(
    records: &tables::Records,
    styles: &tables::Styles,
    context: &tables::Context<'_>,
) -> Result<Vec<xlsx_model::PivotTableMetadata>, String> {
    views(records)?
        .into_iter()
        .map(|view| table(view, styles, context))
        .collect()
}

fn table(
    view: View,
    styles: &tables::Styles,
    context: &tables::Context<'_>,
) -> Result<xlsx_model::PivotTableMetadata, String> {
    if view.olap {
        return Err(unsupported("XLS OLAP PivotTables are not projected"));
    }
    let data = &view.view;
    if data.len() < 44 {
        return Err(truncated());
    }
    let word = |at: usize| u16_at(data, at);
    let (row_first, row_last, col_first, col_last) = (word(0)?, word(2)?, word(4)?, word(6)?);
    let (head, first_data_row, first_data_col) = (word(8)?, word(10)?, word(12)?);
    let cache = i16_at(data, 14)?;
    let data_axis = word(18)?;
    let data_position = i16_at(data, 20)?;
    let counts = [
        word(24)?,
        word(26)?,
        word(28)?,
        word(30)?,
        word(32)?,
        word(34)?,
    ];
    let [rows, columns, pages, data_items, row_lines, column_lines] = counts.map(usize::from);
    let flags = word(36)?;
    let (name, used) = string_no_cch(data, 44, usize::from(word(40)?))?;
    let (_, caption) = string_no_cch(data, 44 + used, usize::from(word(42)?))?;
    if 44 + used + caption != data.len()
        || cache < 0
        || row_last < row_first
        || col_last < col_first
        || head < row_first
        || first_data_row < head
        || first_data_col < col_first
        || name.trim().is_empty()
    {
        return Err(unsupported("invalid XLS PivotTable view"));
    }
    // fAutoFormat: the legacy AutoFormat formats the view by table; the
    // model has no AutoFormat tables.
    if flags & 0x0008 != 0 {
        return Err(unsupported("XLS PivotTable AutoFormat is not projected"));
    }
    let axis_records = usize::from(rows > 0) + usize::from(columns > 0);
    if view.ivd.len() != axis_records
        || view.pages.is_some() != (pages > 0)
        || view.data.len() != data_items
        || view.lines.len() != if row_lines + column_lines > 0 { 2 } else { 0 }
    {
        return Err(unsupported("inconsistent XLS PivotTable view"));
    }
    let mut ivd = view.ivd.iter();
    let row_fields = if rows > 0 {
        fields(ivd.next().expect("row fields"), rows)?
    } else {
        Vec::new()
    };
    let column_fields = if columns > 0 {
        fields(ivd.next().expect("column fields"), columns)?
    } else {
        Vec::new()
    };
    let page_fields = match &view.pages {
        Some(items) if items.len() == pages * 6 => items
            .chunks_exact(6)
            .map(|item| {
                let field = i16_at(item, 0)?;
                let selected = i16_at(item, 2)?;
                if field < 0 || !(0..=0x7ffd).contains(&selected) {
                    return Err(unsupported("invalid XLS PivotTable page field"));
                }
                Ok(xlsx_model::PivotPageField {
                    field: i32::from(field),
                    item: (selected != 0x7ffd).then_some(selected.unsigned_abs().into()),
                    name: None,
                })
            })
            .collect::<Result<Vec<_>, String>>()?,
        Some(_) => return Err(unsupported("invalid XLS PivotTable page fields")),
        None => Vec::new(),
    };
    let mut reasons = vec![feature("pivotFields")];
    let mut data_fields = Vec::new();
    for item in &view.data {
        let field = i16_at(item, 0)?;
        let function = i16_at(item, 2)?;
        let shown_as = i16_at(item, 4)?;
        let chars = u16_at(item, 12)?;
        let subtotal = usize::try_from(function)
            .ok()
            .and_then(|index| SUBTOTALS.get(index))
            .filter(|_| field >= 0)
            .ok_or_else(|| unsupported("invalid XLS PivotTable data item"))?;
        let name = if chars == 0xffff {
            if item.len() != 14 {
                return Err(unsupported("invalid XLS PivotTable data item"));
            }
            None
        } else {
            let (name, used) = string_no_cch(item, 14, usize::from(chars))?;
            if 14 + used != item.len() {
                return Err(unsupported("invalid XLS PivotTable data item"));
            }
            Some(name)
        };
        // df (display the value as a calculation) as the XLSX parser flags
        // dataField showDataAs.
        if shown_as != 0 {
            reasons.push(feature("dataField.showDataAs"));
        }
        data_fields.push(xlsx_model::PivotDataField {
            field: u32::from(field.unsigned_abs()),
            subtotal: Some((*subtotal).to_string()),
            raw_subtotal: None,
            name,
        });
    }
    let (row_items, column_items) = if view.lines.is_empty() {
        (Vec::new(), Vec::new())
    } else {
        (
            lines(&view.lines[0], row_lines, rows)?,
            lines(&view.lines[1], column_lines, columns)?,
        )
    };
    if !row_items.is_empty() {
        reasons.push(feature("rowItems"));
    }
    if !column_items.is_empty() {
        reasons.push(feature("colItems"));
    }
    // sxaxis4Data (2.5.254) on the row axis is the XLSX dataOnRows, and an
    // explicit ipos4Data its dataPosition.
    if data_axis & 0x0001 != 0 || data_position != -1 {
        reasons.push(feature("pivotTable.dataPlacement"));
    }
    // The PivotCache lives in the workbook's pivot cache storage, which is
    // not read.
    reasons.push(feature("pivotCache"));
    let style = match &view.style {
        Some(record) => Some(style(record, styles, context, pages > 0, columns > 0)?),
        None => None,
    };
    Ok(xlsx_model::PivotTableMetadata {
        name,
        cache_id: u32::from(cache.unsigned_abs()),
        location: xlsx_model::PivotLocation {
            range: xlsx_model::CellRange {
                top: u32::from(row_first) + 1,
                left: u32::from(col_first) + 1,
                bottom: u32::from(row_last) + 1,
                right: u32::from(col_last) + 1,
            },
            first_header_row: u32::from(head - row_first),
            first_data_row: u32::from(first_data_row - row_first),
            first_data_col: u32::from(first_data_col - col_first),
        },
        row_fields,
        column_fields,
        page_fields,
        data_fields,
        refresh_on_load: None,
        cache_invalid: None,
        cache_definition_part: None,
        cache_source: None,
        status: xlsx_model::PivotMetadataStatus::Partial { reasons },
        extension_uris: Vec::new(),
        style,
        row_items,
        column_items,
    })
}

/// SXAddl_SXCView_SXDTableStyleClient (2.4.273.107): the option bits and
/// the applied table style: a workbook TableStyle, else a built-in
/// PivotTable style (ECMA-376 Annex G, shared with the XLSX parser) under
/// the workbook theme. fDefaultStyle applies the workbook's default
/// PivotTable style (2.4.322 TableStyles) instead of stName.
fn style(
    record: &[u8],
    styles: &tables::Styles,
    context: &tables::Context<'_>,
    has_pages: bool,
    has_columns: bool,
) -> Result<xlsx_model::PivotTableStyle, String> {
    let flags = u16_at(record, 12)?;
    let chars = usize::from(u16_at(record, 14)?);
    let bytes = record
        .get(16..16 + chars * 2)
        .filter(|_| record.len() == 16 + chars * 2 && chars > 0)
        .ok_or_else(|| unsupported("invalid XLS PivotTable style"))?;
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    let mut name = String::from_utf16(&units).map_err(|_| truncated())?;
    if flags & 0x0040 != 0 {
        name = styles
            .default_pivot_style()
            .ok_or_else(|| unsupported("XLS PivotTable default style is not named"))?
            .to_string();
    }
    let show_last_column = flags & 0x0002 != 0;
    let show_row_stripes = flags & 0x0004 != 0;
    let show_column_stripes = flags & 0x0008 != 0;
    let show_row_headers = flags & 0x0010 != 0;
    let show_column_headers = flags & 0x0020 != 0;
    let elements: Vec<(&'static str, u32, xlsx_model::Dxf)> =
        match styles.pivot_elements(&name, context)? {
            Some(elements) => elements
                .into_iter()
                .filter_map(|(tse, size, dxf)| Some((element_kind(tse)?, size, dxf)))
                .collect(),
            None => {
                let presets = ooxml_common::spreadsheet_style_presets::pivot_style(&name)
                    .ok_or_else(|| unsupported("unknown XLS PivotTable style"))?;
                // A theme-less workbook has no colors for the preset's theme
                // references; no Office default theme is assumed.
                let theme = context
                    .theme
                    .scheme()
                    .ok_or_else(|| unsupported("XLS built-in PivotTable style lacks its theme"))?;
                presets
                    .into_iter()
                    .map(|preset| {
                        (
                            preset.kind,
                            preset.size,
                            xlsx_model::style_presets::preset_dxf(preset.dxf, &theme),
                        )
                    })
                    .collect()
            }
        };
    let mut projected = Vec::new();
    for (kind, size, dxf) in elements {
        // The renderer draws neither the page field area nor column
        // subheadings (it lacks the column header layout).
        let undrawn = match kind {
            "pageFieldLabels" | "pageFieldValues" => has_pages,
            "firstColumnSubheading" | "secondColumnSubheading" | "thirdColumnSubheading" => {
                has_columns && show_column_headers
            }
            _ => false,
        };
        if undrawn {
            return Err(unsupported(
                "XLS PivotTable style element is not representable",
            ));
        }
        projected.push(xlsx_model::PivotTableStyleElement {
            kind: kind.to_string(),
            size,
            dxf,
        });
    }
    Ok(xlsx_model::PivotTableStyle {
        name,
        show_row_headers,
        show_column_headers,
        show_row_stripes,
        show_column_stripes,
        show_last_column,
        elements: projected,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sx_view(counts: [u16; 6], flags: u16) -> Vec<u8> {
        let mut data = Vec::new();
        // B9:C42, head row 9, data row 9, data column 2 (sample-2).
        for value in [8u16, 41, 1, 2, 9, 9, 2, 0, 0, 2, 0xffff, 7] {
            data.extend(value.to_le_bytes());
        }
        for value in counts {
            data.extend(value.to_le_bytes());
        }
        data.extend(flags.to_le_bytes());
        data.extend(1u16.to_le_bytes());
        data.extend(1u16.to_le_bytes());
        data.extend(6u16.to_le_bytes());
        data.extend([0, b'P', 0]);
        data.extend(b"Values");
        data
    }

    fn line(kind: u16, shown: i16, entries: &[i16]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend(0u16.to_le_bytes());
        data.extend(kind.to_le_bytes());
        data.extend(shown.to_le_bytes());
        data.extend(0u16.to_le_bytes());
        for entry in entries {
            data.extend(entry.to_le_bytes());
        }
        data
    }

    fn records(list: &[(u16, Vec<u8>)]) -> tables::Records {
        let mut records = tables::Records::default();
        for (kind, data) in list {
            records.push(*kind, data).unwrap();
        }
        records
    }

    fn project_with(
        list: &[(u16, Vec<u8>)],
        globals: &[super::super::Record<'_>],
    ) -> Result<Vec<xlsx_model::PivotTableMetadata>, String> {
        project_themed(list, globals, super::super::theme::Colors::default())
    }

    fn project_themed(
        list: &[(u16, Vec<u8>)],
        globals: &[super::super::Record<'_>],
        theme: super::super::theme::Colors,
    ) -> Result<Vec<xlsx_model::PivotTableMetadata>, String> {
        let cell_styles = super::super::styles::Styles::parse(&[]).unwrap();
        let context = tables::Context {
            styles: &cell_styles,
            theme: &theme,
        };
        let table_styles = tables::Styles::parse(globals).unwrap();
        project(&records(list), &table_styles, &context)
    }

    fn project_view(
        list: &[(u16, Vec<u8>)],
    ) -> Result<Vec<xlsx_model::PivotTableMetadata>, String> {
        project_with(list, &[])
    }

    fn style_client(flags: u16, name: &str) -> Vec<u8> {
        let mut style = vec![0x64, 0x08, 0, 0, 0x00, 0x1e, 0, 0, 0, 0, 0, 0];
        style.extend(flags.to_le_bytes());
        style.extend((name.encode_utf16().count() as u16).to_le_bytes());
        for unit in name.encode_utf16() {
            style.extend(unit.to_le_bytes());
        }
        style
    }

    fn frt(kind: u16) -> Vec<u8> {
        let mut data = kind.to_le_bytes().to_vec();
        data.extend([0; 10]);
        data
    }

    #[test]
    fn view_location_fields_and_lines_match_the_xlsx_definition() {
        let mut data_item = Vec::new();
        for value in [3i16, 0, 0, 2, 9, 7] {
            data_item.extend(value.to_le_bytes());
        }
        data_item.extend(2u16.to_le_bytes());
        data_item.extend([0, b'G', b'x']);
        let mut row_lines = line(0, 1, &[0, 0x7fff]);
        row_lines.extend(line(0, 2, &[0, 3]));
        row_lines.extend(line(14, 1, &[0, 0x7fff]));
        let grand = line(13, 1, &[0, 0x7fff]);
        let tables = project_view(&[
            (0x00b0, sx_view([2, 0, 0, 1, 4, 1], 0x0001)),
            (
                0x00b4,
                [0u16, 4].iter().flat_map(|v| v.to_le_bytes()).collect(),
            ),
            (0x00c5, data_item),
            (0x00b5, row_lines),
            // An SXLI continued by a Continue record.
            (0x003c, grand),
            (0x00b5, line(0, 0, &[])),
        ])
        .unwrap();
        let table = &tables[0];
        assert_eq!(table.name, "P");
        let location = &table.location;
        assert_eq!(
            (
                location.range.top,
                location.range.left,
                location.range.bottom,
                location.range.right
            ),
            (9, 2, 42, 3)
        );
        assert_eq!(
            (
                location.first_header_row,
                location.first_data_row,
                location.first_data_col
            ),
            (1, 1, 1)
        );
        assert_eq!(table.row_fields, vec![0, 4]);
        assert_eq!(table.data_fields[0].field, 3);
        assert_eq!(table.data_fields[0].subtotal.as_deref(), Some("sum"));
        assert_eq!(table.data_fields[0].name.as_deref(), Some("Gx"));
        let items: Vec<_> = table
            .row_items
            .iter()
            .map(|item| (item.kind.as_str(), item.depth))
            .collect();
        assert_eq!(
            items,
            vec![("data", 0), ("data", 1), ("blank", 0), ("grand", 0)]
        );
        assert_eq!(table.column_items.len(), 1);
        assert!(table.style.is_none());
    }

    #[test]
    fn custom_styles_keep_their_options_and_cleared_edges() {
        // DXF 0: a font colour and a cleared (style 0) left edge.
        let mut dxf = frt(0x088d);
        dxf.extend([0, 0, 0, 0, 2, 0]);
        dxf.extend(5u16.to_le_bytes());
        dxf.extend(12u16.to_le_bytes());
        dxf.extend([0x05, 0, 0, 0, 0x59, 0x59, 0x59, 0xff]);
        dxf.extend(8u16.to_le_bytes());
        dxf.extend(14u16.to_le_bytes());
        dxf.extend([0x01, 0x40, 0, 0, 0, 0, 0, 0, 0, 0]);
        let mut style = frt(0x088f);
        style.extend([0x02, 0]);
        style.extend(1u32.to_le_bytes());
        style.extend(1u16.to_le_bytes());
        style.extend(u16::from(b'S').to_le_bytes());
        let mut element = frt(0x0890);
        for value in [0x17u32, 1, 0] {
            element.extend(value.to_le_bytes());
        }
        let globals = [
            super::super::Record {
                kind: 0x088d,
                offset: 0,
                data: &dxf,
            },
            super::super::Record {
                kind: 0x088f,
                offset: 1,
                data: &style,
            },
            super::super::Record {
                kind: 0x0890,
                offset: 2,
                data: &element,
            },
        ];
        let tables = project_with(
            &[
                (0x00b0, sx_view([0, 0, 0, 0, 0, 0], 0)),
                (0x0864, style_client(0x0036, "S")),
            ],
            &globals,
        )
        .unwrap();
        let style = tables[0].style.as_ref().unwrap();
        assert!(style.show_last_column && style.show_row_stripes && style.show_row_headers);
        assert!(style.show_column_headers && !style.show_column_stripes);
        let element = &style.elements[0];
        assert_eq!(element.kind, "firstRowSubheading");
        assert_eq!(
            element.dxf.font.as_ref().unwrap().color.as_deref(),
            Some("#595959")
        );
        let border = element.dxf.border.as_ref().unwrap();
        assert_eq!(border.left.as_ref().unwrap().style, "none");
        assert!(border.top.is_none());
    }

    #[test]
    fn built_in_and_default_styles_resolve_under_the_workbook_theme() {
        let scheme = [
            [0x00, 0x00, 0x00],
            [0xff, 0xff, 0xff],
            [0x44, 0x54, 0x6a],
            [0xe7, 0xe6, 0xe6],
            [0x44, 0x72, 0xc4],
            [0xed, 0x7d, 0x31],
            [0xa5, 0xa5, 0xa5],
            [0xff, 0xc0, 0x00],
            [0x5b, 0x9b, 0xd5],
            [0x70, 0xad, 0x47],
            [0x05, 0x63, 0xc1],
            [0x95, 0x4f, 0x72],
        ];
        let theme: Vec<String> = scheme
            .iter()
            .map(|[r, g, b]| format!("#{r:02X}{g:02X}{b:02X}"))
            .collect();
        let expected = serde_json::to_string(
            &xlsx_model::style_presets::pivot_style_elements("PivotStyleLight16", &theme).unwrap(),
        )
        .unwrap();
        let view = || (0x00b0, sx_view([0, 0, 0, 0, 0, 0], 0));
        let tables = project_themed(
            &[view(), (0x0864, style_client(0x0036, "PivotStyleLight16"))],
            &[],
            super::super::theme::Colors::from_scheme(scheme),
        )
        .unwrap();
        let style = tables[0].style.as_ref().unwrap();
        assert_eq!(style.name, "PivotStyleLight16");
        assert_eq!(serde_json::to_string(&style.elements).unwrap(), expected);
        // fDefaultStyle takes the TableStyles default PivotTable style.
        let mut defaults = frt(0x088e);
        defaults.extend(0u32.to_le_bytes());
        defaults.extend(1u16.to_le_bytes());
        defaults.extend(17u16.to_le_bytes());
        for unit in "TPivotStyleLight16".encode_utf16() {
            defaults.extend(unit.to_le_bytes());
        }
        let globals = [super::super::Record {
            kind: 0x088e,
            offset: 0,
            data: &defaults,
        }];
        let tables = project_themed(
            &[view(), (0x0864, style_client(0x0076, "Other"))],
            &globals,
            super::super::theme::Colors::from_scheme(scheme),
        )
        .unwrap();
        assert_eq!(tables[0].style.as_ref().unwrap().name, "PivotStyleLight16");
    }

    #[test]
    fn autoformat_olap_and_unresolvable_styles_fail_closed() {
        let base = || (0x00b0, sx_view([0, 0, 0, 0, 0, 0], 0));
        assert!(
            project_view(&[(0x00b0, sx_view([0, 0, 0, 0, 0, 0], 0x0008))])
                .unwrap_err()
                .contains("AutoFormat")
        );
        assert!(project_view(&[base(), (0x080c, vec![0; 16])])
            .unwrap_err()
            .contains("OLAP"));
        assert!(
            project_view(&[base(), (0x0864, style_client(0x0036, "Pv16"))])
                .unwrap_err()
                .contains("unknown")
        );
        // A built-in style needs the workbook theme colors.
        assert!(
            project_view(&[base(), (0x0864, style_client(0x0036, "PivotStyleLight16"))])
                .unwrap_err()
                .contains("theme")
        );
    }

    #[test]
    fn maximal_line_counts_reject_instead_of_wrapping() {
        let error = unsupported("invalid XLS PivotTable lines");
        assert_eq!(lines(&[], 0xffff, 0xffff).unwrap_err(), error);
        assert_eq!(lines(&[0; 8], 0xffff, 0xffff).unwrap_err(), error);
    }
}
