//! OpenDocument spreadsheet (.ods / .ots) reader: ODF Part 3 §9 tables
//! projected directly into the XLSX renderer model (`xlsx-model`).
//!
//! Policy: every feature the XLSX model can express is rendered from its ODF
//! definition. Only corrupt, encrypted or unsafe packages and packages over
//! the resource limits are rejected. Where ODF leaves a choice to the
//! application, LibreOffice Calc is the reference.
//!
//! Geometry. ODF stores column widths and row heights as lengths; the XLSX
//! model stores column widths in maximum-digit-width units of the Normal
//! font (ECMA-376 Part 1 §18.3.1.13). The reader therefore asks the host,
//! through the XLSX host-layout capability, for the digit width of the
//! default cell font, and encodes each authored width so that the renderer's
//! own decoding yields the authored pixel width (96 DPI). Drawing anchors are
//! resolved from the same pixel grid.

use crate::error::invalid;
use crate::odf_props::{self, StyleStore};
use crate::ods_number::DataStyles;
use crate::package::OdfPackage;
pub(crate) use crate::retained_size;
use crate::xml::{self, NS_OFFICE};
use std::collections::{BTreeSet, HashMap};

#[path = "ods_cells.rs"]
pub(crate) mod cells;
#[path = "ods_draw.rs"]
pub(crate) mod draw;
#[path = "ods_extras.rs"]
pub(crate) mod extras;
#[path = "ods_shape_path.rs"]
pub(crate) mod shape_path;
#[path = "ods_table.rs"]
pub(crate) mod table;
#[path = "ods_wire.rs"]
pub mod wire;
#[cfg(test)]
#[path = "ods_tests.rs"]
mod workbook_tests;

pub use crate::odf_text::*;
pub use cells::XfBuilder;

pub const MEDIA_TYPES: [&str; 2] = [
    "application/vnd.oasis.opendocument.spreadsheet",
    "application/vnd.oasis.opendocument.spreadsheet-template",
];

pub const NS_TABLE: &str = crate::xml::NS_TABLE;
pub const NS_CONFIG: &str = crate::xml::NS_CONFIG;

/// The last row and column a SpreadsheetML grid addresses (ECMA-376 Part 1
/// §18.3.1.73 / §18.3.1.13); ODF tables beyond them are clipped, as
/// LibreOffice clips them when it saves SpreadsheetML.
pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;

/// Implementation bound on the cells (values and styled blanks) a workbook
/// materializes. Repeated empty cells with the column's default style are
/// never materialized, so this bounds authored content, not grid size.
pub const MAX_CELLS: usize = 8_000_000;
/// Implementation bound on the retained model of one workbook (strings and
/// fixed per-entry overheads), charged as content is materialized.
pub const MAX_MODEL_BYTES: usize = 512 * 1024 * 1024;
/// Implementation bound on charts and drawings per workbook.
pub const MAX_DRAWINGS: usize = 65_536;

/// ECMA-376 §18.3.1.13 fallback digit width when the host does not measure.
pub const MDW_FALLBACK: f64 = 8.0;

/// Resource accounting shared by the table and drawing readers.
pub(crate) struct Budget {
    cells: usize,
    bytes: usize,
}

impl Budget {
    pub(crate) fn new() -> Self {
        Self { cells: 0, bytes: 0 }
    }
    /// Admit one materialized cell retaining `bytes` (its
    /// [`retained_size`]).
    pub(crate) fn cell(&mut self, bytes: usize) -> Result<(), String> {
        self.cells += 1;
        self.bytes = self.bytes.saturating_add(bytes);
        if self.cells > MAX_CELLS {
            return Err(resource("cell count"));
        }
        self.check()
    }
    pub(crate) fn bytes(&mut self, bytes: usize) -> Result<(), String> {
        self.bytes = self.bytes.saturating_add(bytes);
        self.check()
    }
    fn check(&self) -> Result<(), String> {
        if self.bytes > MAX_MODEL_BYTES {
            Err(resource("model size"))
        } else {
            Ok(())
        }
    }
}

/// A resource-limit rejection (content-free).
pub(crate) fn resource(what: &str) -> String {
    format!("OpenDocument spreadsheet exceeds the {what} limit")
}

/// The workbook-wide reading context shared by the sheet readers.
pub(crate) struct Context {
    pub store: StyleStore,
    pub data: DataStyles,
    pub faces: HashMap<String, String>,
    /// Named gradients, hatches and fill images (`office:styles`).
    pub fills: crate::odf_fill::FillDefs,
    pub xfs: XfBuilder,
    /// Days from the null date (§19.636 `table:null-date`) to 1899-12-30.
    pub null_date_shift: i64,
    pub date1904: bool,
    pub budget: Budget,
    /// Package images referenced by drawings (paths).
    pub images: BTreeSet<String>,
    pub drawing_count: usize,
    pub sheet_names: Vec<String>,
}

/// One table: its worksheet model (rows held apart for the cursor) and the
/// pixel geometry the host-layout decision encodes.
pub struct OdsSheet {
    pub name: String,
    pub visibility: xlsx_model::SheetVisibility,
    pub tab_color: Option<String>,
    pub worksheet: xlsx_model::Worksheet,
    pub rows: Vec<xlsx_model::Row>,
    /// Column runs: (first, last, width px) with 1-based columns.
    pub column_px: Vec<(u32, u32, u32)>,
    pub default_col_px: u32,
    /// Cells naming a content validation (§19.616), by validation name.
    pub validation_refs: Vec<(String, xlsx_model::CellRange)>,
    /// Cell ranges whose cell style carries §16.3 `style:map` conditions,
    /// by style name.
    pub style_map_refs: Vec<(String, xlsx_model::CellRange)>,
}

pub struct OdsWorkbook {
    pub sheets: Vec<OdsSheet>,
    pub styles: Option<xlsx_model::Styles>,
    /// Per-xf left indent in points (applied once the digit width is known).
    pub xf_indent_pt: Vec<f64>,
    pub date1904: bool,
    pub defined_names: Vec<(String, Option<usize>, String)>,
    pub default_font: (String, f64, bool, bool),
    pub images: BTreeSet<String>,
    pub warnings: Vec<String>,
}

pub fn read(package: &OdfPackage) -> Result<OdsWorkbook, String> {
    // ODF Part 3 §2.2.1: `styles.xml` is optional; a package may carry its
    // spreadsheet and automatic styles in `content.xml` alone.
    let styles_text = if package.contains("styles.xml") {
        Some(package.read_text("styles.xml")?)
    } else {
        None
    };
    let content_text = package.read_text("content.xml")?;
    let styles_doc = match &styles_text {
        Some(text) => Some(xml::parse(text)?),
        None => None,
    };
    let content_doc = xml::parse(&content_text)?;
    let styles_root = styles_doc.as_ref().map(|doc| doc.root_element());
    let content_root = content_doc.root_element();
    crate::odf_text::check_space_runs(content_root)?;
    if styles_root.is_some_and(|root| !xml::is(root, NS_OFFICE, "document-styles"))
        || !xml::is(content_root, NS_OFFICE, "document-content")
    {
        return Err(invalid("styles.xml or content.xml has an unexpected root"));
    }
    let mut store = StyleStore::default();
    let mut data = DataStyles::default();
    let mut faces = HashMap::new();
    let mut fills = crate::odf_fill::FillDefs::default();
    if let Some(styles_root) = styles_root {
        odf_props::font_faces(styles_root, &mut faces);
    }
    odf_props::font_faces(content_root, &mut faces);
    for child in styles_root
        .iter()
        .flat_map(|root| root.children())
        .filter(|n| n.is_element())
    {
        if xml::is(child, NS_OFFICE, "styles") {
            store.add(child, false)?;
            fills.add(child, "styles.xml")?;
            data.add(child);
        } else if xml::is(child, NS_OFFICE, "automatic-styles") {
            data.add(child);
        }
    }
    for child in content_root.children().filter(|n| n.is_element()) {
        if xml::is(child, NS_OFFICE, "automatic-styles") {
            store.add(child, true)?;
            data.add(child);
        }
    }
    store.check_chains()?;
    data.check()?;
    let spreadsheet = content_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|body| {
            body.children()
                .find(|n| xml::is(*n, NS_OFFICE, "spreadsheet"))
        })
        .ok_or_else(|| invalid("content.xml has no office:spreadsheet body"))?;
    let (null_date_shift, date1904) = null_date(spreadsheet);
    let xfs = XfBuilder::new(&store, &mut data, &faces);
    let default_font = xfs.default_font();
    let views = if package.contains("settings.xml") {
        let text = package.read_text("settings.xml")?;
        let document = xml::parse(&text)?;
        table::read_view_settings(document.root_element())
    } else {
        HashMap::new()
    };
    let tables: Vec<_> = spreadsheet
        .children()
        .filter(|n| xml::is(*n, NS_TABLE, "table"))
        .collect();
    let sheet_names: Vec<String> = tables
        .iter()
        .map(|t| t.attribute((NS_TABLE, "name")).unwrap_or("").to_string())
        .collect();
    let mut context = Context {
        store,
        data,
        faces,
        fills,
        xfs,
        null_date_shift,
        date1904,
        budget: Budget::new(),
        images: BTreeSet::new(),
        drawing_count: 0,
        sheet_names: sheet_names.clone(),
    };
    let mut sheets = Vec::with_capacity(tables.len());
    let mut pending_charts = Vec::new();
    for (index, node) in tables.iter().enumerate() {
        let view = views.get(&sheet_names[index]);
        let (sheet, charts) = table::read_table(&mut context, package, *node, index, view)?;
        sheets.push(sheet);
        pending_charts.extend(charts);
    }
    if sheets.is_empty() {
        return Err(invalid("the spreadsheet has no table"));
    }
    let mut warnings = Vec::new();
    draw::resolve_charts(
        &mut sheets,
        pending_charts,
        package,
        &mut context,
        &mut warnings,
    )?;
    let mut defined_names = extras::named_expressions(spreadsheet, &sheet_names);
    for (index, node) in tables.iter().enumerate() {
        if let Some(ranges) = node.attribute((NS_TABLE, "print-ranges")) {
            if let Some(formula) = extras::convert_range_list(ranges) {
                defined_names.push(("_xlnm.Print_Area".into(), Some(index), formula));
            }
        }
    }
    extras::database_ranges(spreadsheet, &mut sheets, &sheet_names, &mut defined_names);
    extras::content_validations(&mut context.budget, spreadsheet, &mut sheets)?;
    extras::conditional_formats(&mut context, tables.as_slice(), &mut sheets)?;
    // Defined names in scope for each sheet (workbook-global, then local).
    // Every sheet holds its own copy of each global name: each copy is
    // admitted before it is made.
    for (index, sheet) in sheets.iter_mut().enumerate() {
        let mut names = Vec::new();
        for (name, _, formula) in defined_names
            .iter()
            .filter(|(_, scope, _)| scope.is_none())
            .chain(
                defined_names
                    .iter()
                    .filter(|(_, scope, _)| *scope == Some(index)),
            )
        {
            context.budget.bytes(
                std::mem::size_of::<xlsx_model::DefinedName>()
                    .saturating_add(name.len())
                    .saturating_add(formula.len()),
            )?;
            names.push(xlsx_model::DefinedName {
                name: name.clone(),
                formula: formula.clone(),
            });
        }
        sheet.worksheet.defined_names = names;
    }
    let (styles, xf_indent_pt) = context.xfs.finish();
    Ok(OdsWorkbook {
        sheets,
        styles: Some(styles),
        xf_indent_pt,
        date1904,
        defined_names,
        default_font,
        images: context.images,
        warnings,
    })
}

impl OdsWorkbook {
    /// Encode the pixel geometry with the host's digit width (or the
    /// §18.3.1.13 fallback): column widths, and indents measured in digits.
    pub fn apply_layout(&mut self, mdw: Option<f64>) {
        let mdw = mdw.unwrap_or(MDW_FALLBACK);
        for sheet in &mut self.sheets {
            let ws = &mut sheet.worksheet;
            ws.default_col_width = width_for_px(sheet.default_col_px, mdw);
            ws.col_width_ranges = sheet
                .column_px
                .iter()
                .filter(|(_, _, px)| *px != sheet.default_col_px)
                .map(|(first, last, px)| xlsx_model::ColumnWidthRange {
                    min: *first,
                    max: *last,
                    width: width_for_px(*px, mdw),
                })
                .collect();
        }
        if let Some(styles) = self.styles.as_mut() {
            for (xf, indent) in styles.cell_xfs.iter_mut().zip(&self.xf_indent_pt) {
                if *indent > 0.0 {
                    // The renderer insets `indent × 3 × MDW` pixels.
                    let levels = (indent * 96.0 / 72.0 / (3.0 * mdw)).round();
                    if levels >= 1.0 {
                        xf.indent = Some(levels.min(250.0) as u32);
                    }
                }
            }
        }
    }
}

/// The stored SpreadsheetML width whose decoding (`trunc(((256 w +
/// trunc(128 / mdw)) / 256) × mdw)`, §18.3.1.13) is exactly `px`.
pub fn width_for_px(px: u32, mdw: f64) -> f64 {
    if px == 0 {
        return 0.0;
    }
    let bias = (128.0 / mdw).trunc();
    // Aim at the middle of the pixel so truncation lands on `px`.
    let width = ((px as f64 + 0.5) / mdw * 256.0 - bias) / 256.0;
    (width * 256.0).round() / 256.0
}

/// A length in points as CSS pixels at 96 DPI, rounded as the renderer
/// rounds row heights.
pub fn pt_to_px(pt: f64) -> u32 {
    (pt * 96.0 / 72.0).round().max(0.0) as u32
}

/// §19.636 `table:null-date`: the day value 0. Its shift to the
/// SpreadsheetML 1900 epoch (1899-12-30, the LibreOffice default), or the
/// 1904 date system when the null date is 1904-01-01.
fn null_date(spreadsheet: roxmltree::Node<'_, '_>) -> (i64, bool) {
    let value = spreadsheet
        .children()
        .find(|n| xml::is(*n, NS_TABLE, "calculation-settings"))
        .and_then(|settings| {
            settings
                .children()
                .find(|n| xml::is(*n, NS_TABLE, "null-date"))
        })
        .and_then(|node| node.attribute((NS_TABLE, "date-value")));
    let Some(days) = value.and_then(days_from_civil_text) else {
        return (0, false);
    };
    let epoch_1904 = days_from_civil(1904, 1, 1);
    if days == epoch_1904 {
        return (0, true);
    }
    (days - days_from_civil(1899, 12, 30), false)
}

/// Days since 1970-01-01 of a proleptic Gregorian date.
pub fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let y = if month <= 2 { year - 1 } else { year };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let m = month;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn days_from_civil_text(value: &str) -> Option<i64> {
    let date = value.split('T').next()?;
    let mut parts = date.splitn(3, '-');
    let year: i64 = parts.next()?.parse().ok()?;
    let month: i64 = parts.next()?.parse().ok()?;
    let day: i64 = parts.next()?.get(..2)?.parse().ok()?;
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }
    Some(days_from_civil(year, month, day))
}

/// `office:date-value` (§19.372, an xsd:date or xsd:dateTime) as a
/// SpreadsheetML serial date.
pub(crate) fn date_serial(value: &str, null_shift: i64, date1904: bool) -> Option<f64> {
    let days = days_from_civil_text(value)?;
    let mut fraction = 0.0;
    if let Some(time) = value.split('T').nth(1) {
        let time = time.trim_end_matches('Z');
        let time = time.split(['+']).next().unwrap_or(time);
        let mut parts = time.splitn(3, ':');
        let h: f64 = parts.next()?.parse().ok()?;
        let m: f64 = parts.next().unwrap_or("0").parse().ok()?;
        let s: f64 = parts.next().unwrap_or("0").parse().ok()?;
        fraction = (h * 3600.0 + m * 60.0 + s) / 86_400.0;
    }
    let _ = null_shift;
    // Days from the 1900 system's day 0 (1899-12-30). From 1900-03-01 this
    // is the SpreadsheetML serial (§18.17.4.1); earlier dates keep the same
    // count, as LibreOffice writes them when it saves SpreadsheetML (the
    // 1900 system's nonexistent 1900-02-29 has no ODF date).
    let serial = if date1904 {
        (days - days_from_civil(1904, 1, 1)) as f64
    } else {
        (days - days_from_civil(1899, 12, 30)) as f64
    };
    Some(serial + fraction)
}

/// `office:time-value` (§19.382, an xsd:duration such as `PT12H30M05S`) as a
/// fraction of a day.
pub(crate) fn duration_days(value: &str) -> Option<f64> {
    let value = value.trim();
    let (negative, rest) = match value.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, value),
    };
    let rest = rest.strip_prefix('P')?;
    let (date_part, time_part) = match rest.split_once('T') {
        Some((d, t)) => (d, t),
        None => (rest, ""),
    };
    let mut seconds = 0.0;
    let mut number = String::new();
    for c in date_part.chars() {
        match c {
            'D' => {
                seconds += number.parse::<f64>().ok()? * 86_400.0;
                number.clear();
            }
            'Y' | 'M' | 'W' => return None,
            _ => number.push(c),
        }
    }
    for c in time_part.chars() {
        let unit = match c {
            'H' => 3600.0,
            'M' => 60.0,
            'S' => 1.0,
            _ => {
                number.push(c);
                continue;
            }
        };
        seconds += number.parse::<f64>().ok()? * unit;
        number.clear();
    }
    let days = seconds / 86_400.0;
    Some(if negative { -days } else { days })
}

/// Shared resolved data used by the wire: bootstrap model.
pub fn bootstrap(workbook: &mut OdsWorkbook) -> xlsx_model::ParsedWorkbook {
    let sheets = workbook
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| xlsx_model::SheetMeta {
            name: sheet.name.clone(),
            sheet_id: index as u32 + 1,
            r_id: format!("rId{}", index + 1),
            tab_color: sheet.tab_color.clone(),
            visibility: sheet.visibility.clone(),
        })
        .collect();
    xlsx_model::ParsedWorkbook {
        workbook: xlsx_model::Workbook {
            sheets,
            date1904: workbook.date1904,
            parse_error: None,
        },
        styles: workbook.styles.take().unwrap_or_default(),
        shared_strings: Vec::new(),
    }
}

pub(crate) fn find_cell(rows: &[xlsx_model::Row], row: u32, col: u32) -> Option<&xlsx_model::Cell> {
    let index = rows.binary_search_by_key(&row, |r| r.index).ok()?;
    let cells = &rows[index].cells;
    let at = cells.binary_search_by_key(&col, |c| c.col).ok()?;
    Some(&cells[at])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dates_times_and_addresses() {
        assert_eq!(date_serial("1900-03-01", 0, false), Some(61.0));
        assert_eq!(date_serial("1899-12-31", 0, false), Some(1.0));
        assert_eq!(date_serial("2024-01-05T12:00:00", 0, false), Some(45296.5));
        assert_eq!(date_serial("1904-01-02", 0, true), Some(1.0));
        assert_eq!(duration_days("PT12H"), Some(0.5));
        assert_eq!(duration_days("-PT36H00M00S"), Some(-1.5));
        assert_eq!(parse_cell("$AB$12"), Some((12, 28)));
        assert_eq!(column_letters(28), "AB");
        assert_eq!(
            parse_range("'It''s'.B2:'It''s'.A1"),
            Some((Some("It's".into()), (1, 1), (2, 2)))
        );
        assert_eq!(split_sheet(".C3"), (None, "C3"));
        assert_eq!(sheet_prefix("My sheet"), "'My sheet'!");
        assert_eq!(sheet_prefix("Sheet1"), "Sheet1!");
    }

    #[test]
    fn widths_decode_to_the_authored_pixels() {
        for mdw in [6.0, 7.0, 8.0, 11.0] {
            for px in [1u32, 5, 17, 64, 85, 300] {
                let w = width_for_px(px, mdw);
                let decoded = (((256.0 * w + (128.0 / mdw).trunc()) / 256.0) * mdw).trunc();
                assert_eq!(decoded as u32, px, "mdw {mdw} px {px}");
            }
        }
    }
}
