//! BIFF8 charts -> the shared chart model, without generating SpreadsheetML.
//! MS-XLS 2.1.7.20.1 (chart sheet substream), 2.1.7.20.6 (charts embedded
//! after their Obj record), 2.2.3 (chart concepts) and 2.4.258
//! (ShapePropsStream). Series content is read from the chart data cache
//! (2.2.3.2), the same authoritative source the XLSX path renders.
mod checksum;
mod project;
mod reader;
#[cfg(test)]
mod tests;

use super::drawing_anchors::{self, CellCorner};
use super::{rich, styles, theme, u16_at, unsupported, CellValue, Record, SheetData};
use std::collections::BTreeMap;

pub(crate) use project::Palette;

/// Parse one chart substream (`records` from its BOF through its EOF).
/// `Ok(None)` means the chart has no drawable series.
pub(crate) fn parse(
    records: &[Record<'_>],
    palette: &Palette<'_>,
    references: &project::References<'_>,
) -> Result<Option<ooxml_common::chart::ChartModel>, String> {
    let raw = reader::read(records)?;
    Ok(project::project(&raw, palette, references))
}

const MAX_REFERENCED_CELLS: usize = 32_000;

/// Workbook-global reference resolution for chart series parts: EXTERNSHEET
/// XTI entries (MS-XLS 2.4.105/2.5.172) that point at the workbook's own
/// SupBook (2.4.271, cch 0x0401) and single-sheet tab ranges.
struct Cells<'a> {
    /// XTI index -> converted worksheet index.
    sheets_by_xti: Vec<Option<usize>>,
    sheets: &'a [(String, SheetData)],
    shared: &'a [rich::Text],
}

impl<'a> Cells<'a> {
    fn new(
        records: &[Record<'_>],
        tab_sheets: &BTreeMap<usize, usize>,
        sheets: &'a [(String, SheetData)],
        shared: &'a [rich::Text],
    ) -> Self {
        let mut supbooks = Vec::new();
        let mut sheets_by_xti = Vec::new();
        for record in records.iter().take_while(|r| r.kind != super::EOF) {
            match record.kind {
                0x01ae => supbooks.push(u16_at(record.data, 2).is_ok_and(|cch| cch == 0x0401)),
                0x0017 => {
                    let count = u16_at(record.data, 0).map_or(0, usize::from);
                    for index in 0..count {
                        let at = 2 + index * 6;
                        let entry = (
                            u16_at(record.data, at),
                            u16_at(record.data, at + 2),
                            u16_at(record.data, at + 4),
                        );
                        sheets_by_xti.push(match entry {
                            (Ok(book), Ok(first), Ok(last)) if first == last => supbooks
                                .get(usize::from(book))
                                .copied()
                                .unwrap_or(false)
                                .then(|| tab_sheets.get(&usize::from(first)).copied())
                                .flatten(),
                            _ => None,
                        });
                    }
                }
                _ => {}
            }
        }
        Self {
            sheets_by_xti,
            sheets,
            shared,
        }
    }

    fn cell(&self, sheet: usize, row: u16, column: u16) -> Option<reader::Cached> {
        match self.sheets[sheet].1.rows.get(&row)?.get(&column)? {
            CellValue::Number(number) => Some(reader::Cached::Number(*number)),
            CellValue::Text(text) => Some(reader::Cached::Text(text.clone())),
            CellValue::SharedString(index) => {
                Some(reader::Cached::Text(self.shared.get(*index)?.text.clone()))
            }
            _ => None,
        }
    }

    /// ChartParsedFormula (2.5.49) limited to 3-D cell references, areas and
    /// their unions (ptgRef3d 2.5.198.85, ptgArea3d 2.5.198.28, ptgUnion,
    /// ptgParen). Any other token leaves the part unresolved.
    fn resolve(&self, rgce: &[u8]) -> Option<Vec<Option<reader::Cached>>> {
        let mut output = Vec::new();
        let mut at = 0usize;
        while at < rgce.len() {
            let token = rgce[at];
            let (xti, rows, columns, size) = match token {
                0x3a | 0x5a | 0x7a => {
                    let row = u16_at(rgce, at + 3).ok()?;
                    let column = u16_at(rgce, at + 5).ok()? & 0x3fff;
                    (u16_at(rgce, at + 1).ok()?, (row, row), (column, column), 7)
                }
                0x3b | 0x5b | 0x7b => (
                    u16_at(rgce, at + 1).ok()?,
                    (u16_at(rgce, at + 3).ok()?, u16_at(rgce, at + 5).ok()?),
                    (
                        u16_at(rgce, at + 7).ok()? & 0x3fff,
                        u16_at(rgce, at + 9).ok()? & 0x3fff,
                    ),
                    11,
                ),
                0x10 | 0x15 => {
                    at += 1;
                    continue;
                }
                _ => return None,
            };
            let sheet = (*self.sheets_by_xti.get(usize::from(xti))?)?;
            if rows.0 > rows.1 || columns.0 > columns.1 {
                return None;
            }
            for row in rows.0..=rows.1 {
                for column in columns.0..=columns.1 {
                    if output.len() >= MAX_REFERENCED_CELLS {
                        return None;
                    }
                    output.push(self.cell(sheet, row, column));
                }
            }
            at += size;
        }
        Some(output)
    }
}

struct PreparedChart {
    from: CellCorner,
    to: CellCorner,
    order: u64,
    model: ooxml_common::chart::ChartModel,
}

/// The palette and cell references every chart of the workbook resolves
/// through.
fn with_context<R>(
    records: &[Record<'_>],
    tabs: &[usize],
    styles: &styles::Styles<'_>,
    sheets: &[(String, SheetData)],
    shared: &[rich::Text],
    body: impl FnOnce(&Palette<'_>, &project::References<'_>) -> R,
) -> R {
    let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
    let cells = Cells::new(records, &sheet_ids, sheets, shared);
    let references = |rgce: &[u8]| cells.resolve(rgce);
    let theme = theme::Colors::parse(records).unwrap_or_default();
    let color = |icv: u16| styles.chart_color(icv);
    let global_font = |index: u16| styles.global_font(index);
    let decode_font = |data: &[u8]| styles.chart_font(data);
    let palette = Palette {
        global_font: &global_font,
        decode_font: &decode_font,
        global_font_count: styles.font_count(),
        icv: &color,
        theme: std::array::from_fn(|index| {
            theme
                .argb(index as u32)
                .map(|[_, r, g, b]| format!("{r:02X}{g:02X}{b:02X}"))
        }),
    };
    body(&palette, &references)
}

/// A chart sheet's chart (MS-XLS 2.1.7.20.1 chart sheet substream) and the
/// chart area rectangle from its Chart record (2.4.39: x, y, dx, dy as
/// FixedPoint points), in EMU.
pub(crate) struct ChartSheet {
    pub model: ooxml_common::chart::ChartModel,
    pub x_emu: i64,
    pub y_emu: i64,
    pub width_emu: i64,
    pub height_emu: i64,
}

/// Parse the chart sheet whose BOF (dt 0x0020) is `records[start]`.
pub(super) fn chart_sheet(
    records: &[Record<'_>],
    start: usize,
    tabs: &[usize],
    styles: &styles::Styles<'_>,
    sheets: &[(String, SheetData)],
    shared: &[rich::Text],
) -> Result<ChartSheet, String> {
    let bof = records
        .get(start)
        .ok_or_else(|| unsupported("XLS chart sheet outside the record stream"))?;
    if bof.kind != super::BOF || u16_at(bof.data, 2)? != 0x0020 {
        return Err(unsupported(
            "XLS chart sheet does not start with a chart BOF",
        ));
    }
    let mut depth = 0usize;
    let mut end = None;
    let mut area = None;
    for (index, record) in records.iter().enumerate().skip(start) {
        match record.kind {
            super::BOF => depth += 1,
            super::EOF => {
                depth -= 1;
                if depth == 0 {
                    end = Some(index);
                    break;
                }
            }
            // The chart sheet's own Chart record (not one of a nested
            // substream).
            0x1002 if depth == 1 && area.is_none() => {
                let value = |at| {
                    super::u32_at(record.data, at).map(|v| i64::from(v as i32) * 12_700 / 65_536)
                };
                area = Some((value(0)?, value(4)?, value(8)?, value(12)?));
            }
            _ => {}
        }
    }
    let end = end.ok_or_else(|| unsupported("XLS chart sheet lacks its EOF"))?;
    let (x_emu, y_emu, width_emu, height_emu) =
        area.ok_or_else(|| unsupported("XLS chart sheet lacks its Chart record"))?;
    if width_emu <= 0 || height_emu <= 0 {
        return Err(unsupported("empty XLS chart sheet chart area"));
    }
    let model = with_context(
        records,
        tabs,
        styles,
        sheets,
        shared,
        |palette, references| parse(&records[start..=end], palette, references),
    )?
    .ok_or_else(|| unsupported("BIFF chart without drawable series is not projected"))?;
    Ok(ChartSheet {
        model,
        x_emu,
        y_emu,
        width_emu,
        height_emu,
    })
}

/// Embedded worksheet charts, parsed once and anchored after the Normal-font
/// maximum digit width is measured (column widths depend on it).
#[derive(Default)]
pub(super) struct Charts {
    sheets: BTreeMap<usize, Vec<PreparedChart>>,
}

impl Charts {
    pub(super) fn prepare(
        records: &[Record<'_>],
        tabs: &[usize],
        styles: &styles::Styles<'_>,
        sheets: &[(String, SheetData)],
        shared: &[rich::Text],
    ) -> Result<Self, String> {
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let anchors = drawing_anchors::projectable(records)?;
        with_context(
            records,
            tabs,
            styles,
            sheets,
            shared,
            |palette, references| {
                Self::prepare_anchors(records, anchors, &sheet_ids, palette, references)
            },
        )
    }

    fn prepare_anchors(
        records: &[Record<'_>],
        anchors: Vec<drawing_anchors::DrawingAnchor>,
        sheet_ids: &BTreeMap<usize, usize>,
        palette: &Palette<'_>,
        references: &project::References<'_>,
    ) -> Result<Self, String> {
        let mut charts = Self::default();
        for anchor in anchors {
            let (Some((start, end)), Some(&sheet)) = (anchor.chart, sheet_ids.get(&anchor.sheet))
            else {
                continue;
            };
            let substream = records
                .get(start..=end)
                .ok_or_else(|| unsupported("BIFF chart substream out of range"))?;
            // A chart without Series records projects as an authored empty
            // chart; series that exist but cannot be resolved are rejected.
            let model = parse(substream, palette, references)?.ok_or_else(|| {
                unsupported("BIFF chart without drawable series is not projected")
            })?;
            charts.sheets.entry(sheet).or_default().push(PreparedChart {
                from: anchor.from,
                to: anchor.to,
                order: anchor.order,
                model,
            });
        }
        Ok(charts)
    }

    pub(super) fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }

    /// Resolve MS-XLS 2.5.193 cell fractions to DrawingML cell offsets.
    pub(super) fn resolve(
        self,
        sheets: &[(String, SheetData)],
        mdw: f64,
        warnings: &mut Vec<String>,
    ) -> BTreeMap<usize, Vec<xlsx_model::ChartAnchor>> {
        let mut output = BTreeMap::new();
        let mut omitted = false;
        for (index, charts) in self.sheets {
            let sheet = &sheets[index].1;
            if !sheet.geometry.has_sheet_defaults() || sheet.views.displays_formulas() {
                omitted = true;
                continue;
            }
            let max_row = charts
                .iter()
                .map(|c| c.from.row.max(c.to.row))
                .max()
                .unwrap_or(0);
            let max_col = charts
                .iter()
                .map(|c| c.from.column.max(c.to.column))
                .max()
                .unwrap_or(0);
            let columns = super::pictures::prefix(max_col, |c| sheet.geometry.column_emu(c, mdw));
            let rows = super::pictures::prefix(max_row, |r| sheet.geometry.row_emu(r));
            let locate = |corner: CellCorner| -> Option<(u32, i64, u32, i64)> {
                let column = usize::from(corner.column);
                let row = usize::from(corner.row);
                let x = columns[column]?;
                let y = rows[row]?;
                let dx = (columns[column + 1]? - x) * f64::from(corner.dx) / 1024.0;
                let dy = (rows[row + 1]? - y) * f64::from(corner.dy) / 256.0;
                Some((
                    u32::from(corner.column),
                    dx.round() as i64,
                    u32::from(corner.row),
                    dy.round() as i64,
                ))
            };
            let mut anchors = Vec::new();
            for chart in charts {
                let (Some(from), Some(to)) = (locate(chart.from), locate(chart.to)) else {
                    omitted = true;
                    continue;
                };
                anchors.push(xlsx_model::ChartAnchor {
                    // OfficeArt document order, shared with pictures and shapes.
                    z_order: chart.order,
                    from_col: from.0,
                    from_col_off: from.1,
                    from_row: from.2,
                    from_row_off: from.3,
                    to_col: to.0,
                    to_col_off: to.1,
                    to_row: to.2,
                    to_row_off: to.3,
                    chart: chart.model,
                });
            }
            if !anchors.is_empty() {
                output.insert(index, anchors);
            }
        }
        if omitted {
            warnings.push("legacy-xls:unresolved-chart-geometry-omitted".into());
        }
        output
    }
}
