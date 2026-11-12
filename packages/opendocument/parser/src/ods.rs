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

use crate::budget::{self, Scope};
use crate::error::invalid;
pub(crate) use crate::error::resource;
use crate::odf_props::{self, StyleStore};
use crate::ods_number::DataStyles;
use crate::package::OdfPackage;
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

/// Implementation bound on the cells (values and styled blanks) and row
/// records a workbook materializes. Repeated empty cells with the column's
/// default style are never materialized, so this bounds authored content,
/// not grid size.
pub const MAX_CELLS: usize = 8_000_000;
/// The reader's memory and work bounds are the shared governor's
/// ([`crate::budget`]): measured by the allocator, checked at every step.
pub use crate::budget::MAX_MODEL_BYTES;
/// Implementation bound on charts and drawings per workbook.
pub const MAX_DRAWINGS: usize = 65_536;

/// ECMA-376 §18.3.1.13 fallback digit width when the host does not measure.
pub const MDW_FALLBACK: f64 = 8.0;

/// The workbook-wide reading context shared by the sheet readers.
pub(crate) struct Context {
    pub store: StyleStore,
    pub data: DataStyles,
    pub faces: HashMap<String, String>,
    /// Named gradients, hatches and fill images (`office:styles`).
    pub fills: crate::odf_fill::FillDefs,
    pub xfs: XfBuilder,
    pub date1904: bool,
    /// Package images referenced by drawings (paths).
    pub images: BTreeSet<String>,
    /// Each referenced package image's facts, read once.
    pub image_facts: HashMap<String, draw::ImageFact>,
    pub drawing_count: usize,
    pub sheet_names: Vec<String>,
    /// Position of each sheet name (the first table of the name).
    pub sheet_index: HashMap<String, usize>,
    /// Each embedded chart object read so far (by package path): frames
    /// that reference one object share its single read.
    pub chart_objects: HashMap<String, Option<ooxml_common::chart::ChartModel>>,
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
    /// Cell ranges whose cell style carries §16.3 `style:map` conditions.
    pub(crate) style_map_refs: Vec<(cells::CellStyle, xlsx_model::CellRange)>,
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
    read_within(package, MAX_MODEL_BYTES)
}

/// [`read`] holding at most `limit` bytes beyond the parsed package
/// documents. The documents are parsed first; everything after runs inside
/// one governed [`Scope`], so the returned workbook, and whatever was held
/// while it was built, was measured against `limit`.
pub(crate) fn read_within(package: &OdfPackage, limit: usize) -> Result<OdsWorkbook, String> {
    read_measured(package, limit).0
}

/// [`read`] under `limit`, with what the read held at its peak (bytes) and
/// the steps it took (allocations and ticks): the two quantities the
/// governor bounds, for tests and local inspection.
pub fn read_measured(
    package: &OdfPackage,
    limit: usize,
) -> (Result<OdsWorkbook, String>, usize, u64) {
    match parse_documents(package) {
        Ok(documents) => documents.read(package, limit),
        Err(error) => (Err(error), 0, 0),
    }
}

/// The package's XML parts, read and held while the workbook is built.
struct Documents {
    styles: Option<String>,
    content: String,
    settings: Option<String>,
}

fn parse_documents(package: &OdfPackage) -> Result<Documents, String> {
    // ODF Part 3 §2.2.1: `styles.xml` is optional; a package may carry its
    // spreadsheet and automatic styles in `content.xml` alone.
    let styles_text = if package.contains("styles.xml") {
        Some(package.read_text("styles.xml")?)
    } else {
        None
    };
    let content_text = package.read_text("content.xml")?;
    let settings_text = if package.contains("settings.xml") {
        Some(package.read_text("settings.xml")?)
    } else {
        None
    };
    Ok(Documents {
        styles: styles_text,
        content: content_text,
        settings: settings_text,
    })
}

impl Documents {
    fn read(
        &self,
        package: &OdfPackage,
        limit: usize,
    ) -> (Result<OdsWorkbook, String>, usize, u64) {
        let parsed = (|| {
            Ok::<_, String>((
                match &self.styles {
                    Some(text) => Some(xml::parse(text)?),
                    None => None,
                },
                xml::parse(&self.content)?,
                match &self.settings {
                    Some(text) => Some(xml::parse(text)?),
                    None => None,
                },
            ))
        })();
        let (styles_doc, content_doc, settings_doc) = match parsed {
            Ok(parsed) => parsed,
            Err(error) => return (Err(error), 0, 0),
        };
        let scope = Scope::enter(limit).with_cell_limit(MAX_CELLS);
        let workbook = read_documents(
            package,
            styles_doc.as_ref().map(|doc| doc.root_element()),
            content_doc.root_element(),
            settings_doc.as_ref().map(|doc| doc.root_element()),
        );
        let (peak, work) = (budget::peak_bytes(), budget::work());
        let workbook = workbook.and_then(|workbook| scope.finish().map(|()| workbook));
        (workbook, peak, work)
    }
}

fn read_documents(
    package: &OdfPackage,
    styles_root: Option<roxmltree::Node<'_, '_>>,
    content_root: roxmltree::Node<'_, '_>,
    settings_root: Option<roxmltree::Node<'_, '_>>,
) -> Result<OdsWorkbook, String> {
    // Style values (names, families, lengths) are bounded before any style
    // resolves: a reader may then treat one as a small constant.
    crate::odf_text::check_style_values(content_root)?;
    if let Some(root) = styles_root {
        crate::odf_text::check_style_values(root)?;
    }
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
        odf_props::font_faces(styles_root, &mut faces)?;
    }
    odf_props::font_faces(content_root, &mut faces)?;
    for child in styles_root
        .iter()
        .flat_map(|root| root.children())
        .filter(|n| n.is_element())
    {
        if xml::is(child, NS_OFFICE, "styles") {
            store.add(child, false)?;
            fills.add(child, "styles.xml")?;
            data.add(child)?;
        } else if xml::is(child, NS_OFFICE, "automatic-styles") {
            data.add(child)?;
        }
    }
    for child in content_root.children().filter(|n| n.is_element()) {
        if xml::is(child, NS_OFFICE, "automatic-styles") {
            store.add(child, true)?;
            data.add(child)?;
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
    let date1904 = uses_1904_dates(spreadsheet);
    let xfs = XfBuilder::new(&store, &mut data, &faces)?;
    let default_font = xfs.default_font();
    let views = match settings_root {
        Some(root) => table::read_view_settings(root)?,
        None => HashMap::new(),
    };
    let tables: Vec<_> = spreadsheet
        .children()
        .filter(|n| xml::is(*n, NS_TABLE, "table"))
        .collect();
    let sheet_names: Vec<String> = tables
        .iter()
        .map(|t| t.attribute((NS_TABLE, "name")).unwrap_or("").to_string())
        .collect();
    // Sheet positions by name (the first table of a name), so name lookups
    // never scan every sheet.
    let mut sheet_index: HashMap<String, usize> = HashMap::new();
    for (index, name) in sheet_names.iter().enumerate() {
        budget::tick(1)?;
        if !sheet_index.contains_key(name) {
            sheet_index.insert(name.clone(), index);
        }
    }
    let mut context = Context {
        store,
        data,
        faces,
        fills,
        xfs,
        date1904,
        images: BTreeSet::new(),
        image_facts: HashMap::new(),
        drawing_count: 0,
        sheet_names: sheet_names.clone(),
        sheet_index,
        chart_objects: HashMap::new(),
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
    let mut defined_names = extras::named_expressions(spreadsheet, &context.sheet_index)?;
    for (index, node) in tables.iter().enumerate() {
        if let Some(ranges) = node.attribute((NS_TABLE, "print-ranges")) {
            if let Some(formula) = extras::convert_range_list(ranges) {
                defined_names.push(("_xlnm.Print_Area".into(), Some(index), formula));
            }
        }
    }
    extras::database_ranges(
        spreadsheet,
        &mut sheets,
        &context.sheet_index,
        &mut defined_names,
    )?;
    extras::content_validations(spreadsheet, &mut sheets)?;
    extras::conditional_formats(&mut context, tables.as_slice(), &mut sheets)?;
    // Defined names in scope for each sheet (workbook-global, then local),
    // grouped by scope once. Every sheet holds its own copy of each global
    // name: one step per copy.
    let global: Vec<usize> = (0..defined_names.len())
        .filter(|at| defined_names[*at].1.is_none())
        .collect();
    let mut local: Vec<Vec<usize>> = Vec::new();
    for _ in 0..sheets.len() {
        budget::tick(1)?;
        local.push(Vec::new());
    }
    for (at, (_, scope, _)) in defined_names.iter().enumerate() {
        if let Some(list) = scope.and_then(|index| local.get_mut(index)) {
            list.push(at);
        }
    }
    for (index, sheet) in sheets.iter_mut().enumerate() {
        let mut names = Vec::new();
        for at in global.iter().chain(local[index].iter()) {
            budget::tick(1)?;
            let (name, _, formula) = &defined_names[*at];
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

/// §19.636 `table:null-date`: whether the day value 0 is 1904-01-01, the
/// SpreadsheetML 1904 date system. Dates are read from their calendar
/// `office:date-value`, so any other null date needs no shift.
fn uses_1904_dates(spreadsheet: roxmltree::Node<'_, '_>) -> bool {
    spreadsheet
        .children()
        .find(|n| xml::is(*n, NS_TABLE, "calculation-settings"))
        .and_then(|settings| {
            settings
                .children()
                .find(|n| xml::is(*n, NS_TABLE, "null-date"))
        })
        .and_then(|node| node.attribute((NS_TABLE, "date-value")))
        .and_then(days_from_civil_text)
        == Some(days_from_civil(1904, 1, 1))
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

/// `office:date-value` (§19.372, an xsd:date or xsd:dateTime) as the serial
/// the workbook renderer displays as that date.
///
/// In the 1904 system that is the day count from 1904-01-01. In the 1900
/// system it is the day count from 1899-12-30 from 1900-03-01 on
/// (§18.17.4.1). Before that the renderer, like Excel, counts the
/// nonexistent 1900-02-29 and shows serial `n` as the day `n + 1` after
/// 1899-12-30 (`core/src/excel-date.ts`), so an earlier date is one less:
/// 1900-01-01 is serial 1. The 1900 system has no serial for 1899-12-30 or
/// before (Excel shows serial 0 as "1900-01-00"); such a value keeps its
/// plain day count, which shows its time of day and, for a negative count,
/// the renderer's invalid-date marker.
pub(crate) fn date_serial(value: &str, date1904: bool) -> Option<f64> {
    let days = days_from_civil_text(value)?;
    let mut fraction = 0.0;
    if let Some(time) = value.split('T').nth(1) {
        let time = time.trim_end_matches('Z');
        // A zone offset (§19.372 xsd:dateTime, either sign) is dropped: the
        // authored local time is shown.
        let time = time.split(['+', '-']).next().unwrap_or(time);
        let mut parts = time.splitn(3, ':');
        let h: f64 = parts.next()?.parse().ok()?;
        let m: f64 = parts.next().unwrap_or("0").parse().ok()?;
        let s: f64 = parts.next().unwrap_or("0").parse().ok()?;
        fraction = (h * 3600.0 + m * 60.0 + s) / 86_400.0;
    }
    let serial = if date1904 {
        days - days_from_civil(1904, 1, 1)
    } else {
        let days = days - days_from_civil(1899, 12, 30);
        if (1..=60).contains(&days) {
            days - 1
        } else {
            days
        }
    };
    Some(serial as f64 + fraction)
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

/// The sheet a range names (the first of that name), or the first sheet
/// for a range without a sheet name.
pub(crate) fn sheet_named<'a>(
    sheets: &'a [OdsSheet],
    index: &HashMap<String, usize>,
    name: Option<&str>,
) -> Option<&'a OdsSheet> {
    match name {
        None => sheets.first(),
        Some(name) => index.get(name).and_then(|at| sheets.get(*at)),
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
        assert_eq!(date_serial("1900-03-01", false), Some(61.0));
        // Before 1900-03-01 the renderer shows serial n as day n + 1.
        assert_eq!(date_serial("1900-02-28", false), Some(59.0));
        assert_eq!(date_serial("1900-01-01", false), Some(1.0));
        assert_eq!(date_serial("1899-12-31T06:00:00", false), Some(0.25));
        assert_eq!(date_serial("1899-12-30T06:00:00", false), Some(0.25));
        assert_eq!(date_serial("1899-12-29", false), Some(-1.0));
        assert_eq!(date_serial("2024-01-05T12:00:00", false), Some(45296.5));
        // A zone offset of either sign keeps the authored local time.
        assert_eq!(
            date_serial("2024-01-05T12:00:00-05:00", false),
            Some(45296.5)
        );
        assert_eq!(
            date_serial("2024-01-05T12:00:00+05:00", false),
            Some(45296.5)
        );
        assert_eq!(date_serial("1904-01-02", true), Some(1.0));
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
