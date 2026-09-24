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
use super::{rich, styles, theme, u16_at, CellValue, Record, SheetData};
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
    model: ooxml_common::chart::ChartModel,
}

/// Embedded worksheet charts, parsed once and anchored after the Normal-font
/// maximum digit width is measured (column widths depend on it).
#[derive(Default)]
pub(super) struct Charts {
    sheets: BTreeMap<usize, Vec<PreparedChart>>,
    unsupported: bool,
}

impl Charts {
    pub(super) fn prepare(
        records: &[Record<'_>],
        tabs: &[usize],
        styles: &styles::Styles<'_>,
        sheets: &[(String, SheetData)],
        shared: &[rich::Text],
    ) -> Self {
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let cells = Cells::new(records, &sheet_ids, sheets, shared);
        let references = |rgce: &[u8]| cells.resolve(rgce);
        let Ok(anchors) = drawing_anchors::projectable(records) else {
            return Self {
                unsupported: true,
                ..Self::default()
            };
        };
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
        let mut charts = Self::default();
        for anchor in anchors {
            let (Some((start, end)), Some(&sheet)) = (anchor.chart, sheet_ids.get(&anchor.sheet))
            else {
                continue;
            };
            let Some(substream) = records.get(start..=end) else {
                charts.unsupported = true;
                continue;
            };
            match parse(substream, &palette, &references) {
                Ok(Some(model)) => charts.sheets.entry(sheet).or_default().push(PreparedChart {
                    from: anchor.from,
                    to: anchor.to,
                    model,
                }),
                Ok(None) => {}
                Err(_) => charts.unsupported = true,
            }
        }
        charts
    }

    pub(super) fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }

    pub(super) fn has_unsupported(&self) -> bool {
        self.unsupported
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
            for (ordinal, chart) in charts.into_iter().enumerate() {
                let (Some(from), Some(to)) = (locate(chart.from), locate(chart.to)) else {
                    omitted = true;
                    continue;
                };
                anchors.push(xlsx_model::ChartAnchor {
                    z_order: ordinal as u64,
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
