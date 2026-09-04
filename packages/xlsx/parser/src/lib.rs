use std::collections::{BTreeMap, HashMap, HashSet};
#[cfg(test)]
use std::io::Cursor;
use std::rc::Rc;
use wasm_bindgen::prelude::*;

use ooxml_common::depth::parse_guarded;
use ooxml_common::json_measurement::measure_json;
use ooxml_common::ns::{attr_ns, is_r_ns, is_x_ns, relationships};
use ooxml_common::package_session::{
    PackageLimitReporter, PackageOperation, PackageSessionHandle, RetainedPackageOperation,
};
use ooxml_common::resource::{
    HardResourceLimitKind, ResourceUsage, HARD_MAX_XLSX_WORKSHEET_CELLS,
    HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES, HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
    HARD_MAX_XLSX_WORKSHEET_ROWS,
};

mod markdown;
mod pivot;
use pivot::*;

mod worksheet_reference;
use worksheet_reference::{
    extend_lookup_transactionally, resolve_worksheet_reference, ReferencedCellValue,
    WorksheetCellLookup, WorksheetCellLookupBuilder, WorksheetReferenceSession,
};
#[cfg(test)]
use worksheet_reference::{extract_reference_cells, MAX_REFERENCE_CELLS};

mod worksheet_projector;
#[cfg(test)]
use worksheet_projector::stream_worksheet_rows;
use worksheet_projector::StreamedSheetData;

mod worksheet_cursor;
use worksheet_cursor::{
    WorksheetCursor, WorksheetCursorPull, WorksheetCursorTail, WORKSHEET_CURSOR_PULL_ROWS,
    WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES,
};

mod types;
pub use types::*;
mod styles;
use styles::*;
mod chart;
use chart::*;
mod drawing;
use drawing::*;
mod slicer;
use slicer::*;
mod table;
use table::*;

/// XLSX-owned adapter over one validated package session.
///
/// All reads belonging to one public parser call share the same logical
/// [`PackageOperation`]. This keeps operation accounting stable across workbook
/// dependencies, worksheet pulls, ancillary parts, and repeat reads. Focused
/// parser tests that call a low-level helper directly receive a lazily-created
/// compatibility operation with the same bounded reader semantics.
pub(crate) struct XlsxZip {
    session: PackageSessionHandle,
    operation: RetainedPackageOperation,
}

impl XlsxZip {
    #[cfg(test)]
    pub(crate) fn new(source: Cursor<Vec<u8>>) -> Result<Self, String> {
        open_zip(source.into_inner())
    }

    fn begin_operation(&mut self, name: &str) -> Result<(), String> {
        self.operation.begin(&self.session, name)
    }

    fn usage(&self) -> ResourceUsage {
        self.session.usage()
    }

    fn operation(&mut self) -> Result<&PackageOperation, String> {
        #[cfg(test)]
        let compatibility_name = Some("xlsx-parser-compat");
        #[cfg(not(test))]
        let compatibility_name = None;
        self.operation.operation(&self.session, compatibility_name)
    }

    /// Return the explicitly-started operation used by persistent production
    /// cursors. Unlike `operation`, this never creates a compatibility scope.
    fn active_operation(&self) -> Result<&PackageOperation, String> {
        self.operation.active()
    }

    fn finish_operation(&mut self) -> Result<(), String> {
        self.operation.finish()
    }

    fn cancel_operation(&mut self) {
        self.operation.cancel();
    }

    fn run_operation<T>(
        &mut self,
        name: &str,
        run: impl FnOnce(&mut Self) -> Result<T, String>,
    ) -> Result<T, String> {
        self.begin_operation(name)?;
        let result = run(self);
        self.operation.settle(&self.session, result)
    }

    fn assert_healthy(&self) -> Result<(), String> {
        self.session.assert_healthy()
    }

    fn entry_paths(&self) -> Vec<String> {
        self.session.entry_paths()
    }

    fn index_for_name(&self, path: &str) -> Option<()> {
        self.session.contains_entry(path).then_some(())
    }
}

pub(crate) fn read_zip_string(archive: &mut XlsxZip, path: &str) -> Result<String, String> {
    archive.operation()?.read_string(path)
}

pub(crate) fn read_zip_string_head(
    archive: &mut XlsxZip,
    path: &str,
    max_bytes: usize,
) -> Result<String, String> {
    let mut bytes = archive.operation()?.read_head(path, max_bytes)?;
    match std::str::from_utf8(&bytes) {
        Ok(text) => Ok(text.to_owned()),
        Err(error) if error.error_len().is_none() => {
            bytes.truncate(error.valid_up_to());
            Ok(String::from_utf8(bytes).expect("validated UTF-8 prefix"))
        }
        Err(error) => Err(format!("ZIP entry is not valid UTF-8 ({path}): {error}")),
    }
}

pub(crate) fn read_zip_bytes(archive: &mut XlsxZip, path: &str) -> Result<Vec<u8>, String> {
    archive.operation()?.read_bytes(path)
}

fn settle_xlsx_operation<T>(archive: &mut XlsxZip, result: Result<T, String>) -> Result<T, String> {
    archive.operation.settle(&archive.session, result)
}

/// Part-name tag for a whole-container degradation (#774). Already parenthesized
/// (`"(zip container)"`), symmetric with docx / pptx `"(zip container)"` — so
/// error formatting below must not wrap it in another pair of parens.
const CONTAINER_PART: &str = "(zip container)";

#[derive(Default)]
struct WorksheetModelUsage {
    rows: u64,
    cells: u64,
    owned_utf8_bytes: u64,
}

fn report_materialization_limit(
    archive: &mut XlsxZip,
    kind: HardResourceLimitKind,
    part: &str,
    limit: u64,
    observed: u64,
) -> Result<(), String> {
    archive
        .operation()?
        .limit_reporter()?
        .observe_hard_limit(kind, Some(part), limit, observed)
}

fn serialize_worksheet_bounded(
    archive: &mut XlsxZip,
    part: &str,
    worksheet: &Worksheet,
) -> Result<Vec<u8>, String> {
    let json_bytes = measure_json(worksheet)?.json_bytes;
    report_materialization_limit(
        archive,
        HardResourceLimitKind::WorksheetJsonBytes,
        part,
        HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
        json_bytes,
    )?;
    serde_json::to_vec(worksheet).map_err(|error| format!("serialize error: {error}"))
}

fn row_cell_content_utf8_bytes(
    rows: &[Row],
    shared_strings: &[SharedString],
) -> Result<u64, String> {
    rows.iter().try_fold(0u64, |row_total, row| {
        row.cells.iter().try_fold(row_total, |cell_total, cell| {
            let formula_bytes = cell
                .formula
                .as_ref()
                .map_or(0, |formula| formula.len() as u64);
            let value_bytes = match &cell.value {
                CellValue::Shared { si } => {
                    let shared = shared_strings.get(*si);
                    let shared_bytes = match shared {
                        Some(value) => measure_json(value)?.string_value_utf8_bytes,
                        None => 0,
                    };
                    // A shared reference becomes a materialized Text value. The
                    // `text` discriminator is therefore retained per cell too.
                    shared_bytes.checked_add(4)
                }
                value => measure_json(value)
                    .map(|measurement| measurement.string_value_utf8_bytes)?
                    .checked_add(0),
            }
            .ok_or_else(|| "worksheet cell string measurement overflow".to_string())?;
            cell_total
                .checked_add(formula_bytes)
                .and_then(|total| total.checked_add(value_bytes))
                .ok_or_else(|| "worksheet cell string measurement overflow".to_string())
        })
    })
}

/// Open a xlsx ZIP container, tagging a failure with the container part name.
///
/// #774 (RB7 MAJOR, symmetric with docx / pptx `open_zip`): a truncated / corrupt
/// ZIP is the MOST COMMON way a xlsx is broken (an incomplete download, a
/// byte-mangled attachment). `ZipArchive::new` maps that to an opaque
/// `zip::result::ZipError` that, if propagated, throws with no indication that the
/// CONTAINER (not some inner part) is the problem. Naming the failure lets the
/// caller build a `degraded_container_workbook` / `degraded_container_sheet`
/// tagged with the container, symmetric with how a corrupt sheet part is tagged
/// inside [`parse_sheet_with`].
///
/// `CONTAINER_PART` already carries its own parens, so this formats as
/// `"{CONTAINER_PART}: {e}"` — NOT `"({CONTAINER_PART}): {e}"`, which would
/// double-parenthesize into `"((zip container)): ..."` (docx / pptx avoid this
/// by writing the literal `"(zip container)"` directly instead of a
/// pre-parenthesized constant).
pub(crate) fn open_zip(data: Vec<u8>) -> Result<XlsxZip, String> {
    open_zip_with_limits(data, None, None)
}

fn open_zip_with_limits(
    data: Vec<u8>,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<XlsxZip, String> {
    open_zip_with_policy(
        data,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        None,
    )
}

fn open_zip_with_policy(
    data: Vec<u8>,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
    max_archive_entries: Option<u64>,
) -> Result<XlsxZip, String> {
    PackageSessionHandle::open(
        data,
        ooxml_common::resource::OoxmlFormat::Xlsx,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
        max_archive_entries,
    )
    .map(|session| XlsxZip {
        session,
        operation: RetainedPackageOperation::new("xlsx"),
    })
    .map_err(ooxml_common::zip::tag_container_error)
}

/// A placeholder [`ParsedWorkbook`] for a xlsx whose ZIP CONTAINER could not be
/// opened (truncated / corrupt / not a zip). No parts are readable, so there is
/// no styles / theme / sharedStrings to derive — surface a single placeholder
/// sheet carrying the container-tagged error so the viewer lists one tab and
/// paints a "could not be displayed" overlay. Mirrors the per-sheet
/// [`Worksheet::placeholder`] used inside [`parse_sheet_with`], but for the
/// whole-container case.
fn degraded_container_workbook(parse_error: String) -> ParsedWorkbook {
    ParsedWorkbook {
        workbook: Workbook {
            sheets: vec![SheetMeta {
                name: CONTAINER_PART.to_string(),
                sheet_id: 1,
                r_id: String::new(),
                tab_color: None,
                visibility: SheetVisibility::Visible,
            }],
            date1904: false,
            parse_error: Some(parse_error),
        },
        styles: Styles::default(),
        shared_strings: Vec::new(),
    }
}

/// The single placeholder [`Worksheet`] for the whole-container degradation
/// (#774): the viewer parses sheet 0 of a [`degraded_container_workbook`] and
/// gets this back, so it paints the same part-tagged error overlay the per-sheet
/// break uses. `name` is the placeholder tab name (`CONTAINER_PART`).
fn degraded_container_sheet(parse_error: String) -> Worksheet {
    Worksheet::placeholder(CONTAINER_PART, parse_error)
}

// Excel built-in indexed color palette (indices 0-63)
// Standard Excel 2003 color palette
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", // 0-7
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", // 8-15
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0",
    "#808080", // 16-23
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC",
    "#CCCCFF", // 24-31
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080",
    "#0000FF", // 32-39
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF",
    "#FFCC99", // 40-47
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699",
    "#969696", // 48-55
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399",
    "#333333", // 56-63
];

/// Parse a xlsx archive's workbook index and return it as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_xlsx(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    parse_xlsx_inner_with_limits(data, max_archive_entry_bytes, max_total_inflated_bytes)
        .and_then(|wb| serde_json::to_vec(&wb).map_err(|e| format!("serialize error: {e}")))
        .map_err(|error| JsValue::from_str(&error))
}

/// Workbook-level parts that every worksheet projection needs but that do NOT change
/// between sheets: the workbook.xml / workbook.xml.rels source strings, the
/// resolved sheet list, the theme palette, and the shared-string table.
///
/// Building these is the bulk of a sheet parse's fixed cost — `sharedStrings.xml`
/// in particular is decompressed and walked in full. The stateful `XlsxArchive`
/// builds them once and reuses them for every worksheet cursor.
///
/// The XML is kept as owned `String`s rather than as `roxmltree::Document`s: a
/// `Document` borrows its source, so it can't be cached across calls, but
/// re-parsing the small workbook.xml / rels strings from memory (no zip inflate)
/// is negligible next to re-decompressing + re-walking sharedStrings/theme.
struct WorkbookShared {
    workbook_xml: String,
    rels_xml: String,
    sheets: Vec<SheetMeta>,
    theme_colors: Rc<[String]>,
    /// DrawingML `fmtScheme`, parsed once per workbook and reused by every
    /// sheet's shape tree. Keeping the complete recipes avoids the previous
    /// per-sheet theme re-inflate and width-only `lnStyleLst` projection.
    theme_format_scheme: Rc<ooxml_common::theme::ThemeFormatScheme>,
    /// Image relationships owned by the workbook theme part. Chart Style
    /// `fillRef` recipes resolve `blipFill` rIds in this scope, not in the
    /// chart or style part.
    theme_chart_images: Rc<ooxml_common::chart::ChartImageRelationships>,
    /// Workbook theme `(majorFont.latin, minorFont.latin)` Latin faces
    /// (§20.1.4.2). Chart-text fallback font (CH10).
    theme_fonts: (Option<String>, Option<String>),
    /// Lightweight style projections used while materializing sheets. Full
    /// workbook styles stay owned by the full-parse path instead of being
    /// retained and deeply cloned here.
    default_font: (Option<String>, Option<f64>),
    chart_number_formats: ChartNumberFormatCache,
    shared_strings: Rc<[SharedString]>,
    /// #773: a part-tagged degradation error set when `xl/sharedStrings.xml` was
    /// PRESENT but corrupt (a broken shared-string table blanks every string cell
    /// across all sheets). `None` when the part read cleanly or is legitimately
    /// absent. Surfaced onto the workbook index's `parse_error` so the loss is
    /// visible rather than silent, without taking any sheet down.
    shared_strings_error: Option<String>,
    /// Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28).
    /// `true` = 1904 date system. Parsed once here and denormalized onto every
    /// worksheet so the renderer/cell formatter can resolve serial dates
    /// without a back-reference to the workbook.
    date1904: bool,
}

#[derive(Default)]
struct XlsxThemeData {
    colors: Vec<String>,
    format_scheme: ooxml_common::theme::ThemeFormatScheme,
    fonts: (Option<String>, Option<String>),
    chart_images: ooxml_common::chart::ChartImageRelationships,
}

impl XlsxThemeData {
    fn load(archive: &mut XlsxZip, workbook_rels_xml: &str) -> Self {
        let Some(target) = find_internal_rel_target_by_type(workbook_rels_xml, "/theme") else {
            return Self::default();
        };
        let theme_path = resolve_zip_path("xl", &target);
        let Ok(xml) = read_zip_string(archive, &theme_path) else {
            return Self::default();
        };
        let mut theme = Self::parse(&xml);
        let rels_path = ooxml_common::rels::relationship_part_path(&theme_path);
        if let Ok(rels_xml) = read_zip_string(archive, &rels_path) {
            theme.chart_images.insert_part_relationships(
                ooxml_common::chart::ChartImageSource::Theme,
                &theme_path,
                &rels_xml,
            );
        }
        theme
    }

    fn parse(xml: &str) -> Self {
        let colors = ooxml_common::theme::ThemeColorScheme::parse(xml)
            .slots_in_order()
            .into_iter()
            .flatten()
            .map(|hex| format!("#{}", hex.to_uppercase()))
            .collect();
        let theme_fonts = ooxml_common::theme::ThemeFonts::parse(xml);
        Self {
            colors,
            format_scheme: ooxml_common::theme::ThemeFormatScheme::parse(xml),
            fonts: (theme_fonts.major.latin, theme_fonts.minor.latin),
            chart_images: ooxml_common::chart::ChartImageRelationships::default(),
        }
    }
}

impl WorkbookShared {
    /// Read + parse the workbook-level shared parts from an opened archive.
    ///
    /// `workbook.xml` is mandatory (a workbook without it is unparseable), but
    /// `workbook.xml.rels` is read leniently (empty on absence): the original
    /// `parse_xlsx` tolerated a missing rels part (tab colors skipped), while
    /// worksheet projection requires it — so the mandatory-rels enforcement stays in
    /// `parse_sheet_with`, where an empty rels string fails `resolve_sheet_path`
    /// exactly as the old `?` on the rels read did.
    fn load(archive: &mut XlsxZip) -> Result<WorkbookShared, String> {
        let (shared, _) = Self::load_impl(archive, false)?;
        Ok(shared)
    }

    fn load_with_styles(
        archive: &mut XlsxZip,
    ) -> Result<(WorkbookShared, Result<Styles, String>), String> {
        let (shared, styles) = Self::load_impl(archive, true)?;
        Ok((shared, styles.expect("full style loading requested")))
    }

    fn load_impl(
        archive: &mut XlsxZip,
        include_full_styles: bool,
    ) -> Result<(WorkbookShared, Option<Result<Styles, String>>), String> {
        let workbook_xml = read_zip_string(archive, "xl/workbook.xml")?;
        let (sheets, date1904) = {
            let wb_doc = parse_guarded(&workbook_xml).map_err(|e| e.to_string())?;
            (
                parse_workbook_sheets(&wb_doc),
                parse_workbook_date1904(&wb_doc),
            )
        };
        let rels_xml = read_zip_string(archive, "xl/_rels/workbook.xml.rels").unwrap_or_default();
        let theme = XlsxThemeData::load(archive, &rels_xml);
        let theme_colors: Rc<[String]> = theme.colors.into();
        let theme_format_scheme = Rc::new(theme.format_scheme);
        let theme_fonts = theme.fonts;
        let theme_chart_images = Rc::new(theme.chart_images);
        let (default_font, chart_number_formats, styles) = if include_full_styles {
            match parse_styles(archive, theme_colors.as_ref()) {
                Ok(parsed) => (
                    parsed.default_font,
                    parsed.chart_number_formats,
                    Some(Ok(parsed.styles)),
                ),
                Err(error) => (
                    (None, None),
                    ChartNumberFormatCache::default(),
                    Some(Err(error)),
                ),
            }
        } else {
            match styles::parse_style_projection(archive) {
                Ok(parsed) => (parsed.default_font, parsed.chart_number_formats, None),
                Err(_) => ((None, None), ChartNumberFormatCache::default(), None),
            }
        };
        let (shared_strings, shared_strings_error) =
            read_shared_strings(archive, theme_colors.as_ref());
        Ok((
            WorkbookShared {
                workbook_xml,
                rels_xml,
                sheets,
                theme_colors,
                theme_format_scheme,
                theme_chart_images,
                theme_fonts,
                default_font,
                chart_number_formats,
                shared_strings: shared_strings.into(),
                shared_strings_error,
                date1904,
            },
            styles,
        ))
    }
}

/// Parse one worksheet from an opened archive, reusing already-loaded
/// [`WorkbookShared`] parts. Native MCP materialization and the cursor projection
/// both use this implementation, so worksheet semantics have one source.
///
/// `wb_doc` / `rels_doc` are re-parsed here from the cached source strings (cheap
/// in-memory roxmltree parses, no zip inflate) because a `roxmltree::Document`
/// borrows its input and so can't be stored in `WorkbookShared`.
fn parse_sheet_with(
    archive: &mut XlsxZip,
    shared: &WorkbookShared,
    sheet_index: u32,
    name: &str,
) -> Result<Vec<u8>, String> {
    // `workbook.xml.rels` is mandatory for a sheet parse (the original
    // the historical worksheet path read it with `?`). `WorkbookShared` caches it leniently for
    // the `parse_xlsx` path, so on the (defensive) missing-rels case re-read it
    // here to surface the identical "entry not found" error the old code raised.
    if shared.rels_xml.is_empty() {
        read_zip_string(archive, "xl/_rels/workbook.xml.rels")?;
    }
    let rels_doc = parse_guarded(&shared.rels_xml).map_err(|e| e.to_string())?;

    let sheet_meta = shared
        .sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;
    let sheet_part_kind = resolve_sheet_part_kind(&rels_doc, &sheet_meta.r_id);

    let theme_colors = shared.theme_colors.as_ref();
    let sheet_part = format!("xl/{}", sheet_path);
    // RB7 partial degradation: the sheet's own XML read + parse are the two
    // failure points that concern ONE sheet (the workbook-level parts above are
    // shared, cached, and already lenient). If either fails, don't abort the
    // whole workbook — return an empty placeholder sheet whose `parse_error`
    // names the offending part, so the OTHER sheets stay openable. Everything
    // after (images / charts / comments / …) is already lenient (returns empty
    // on error), so it stays outside this guard.
    let sheet_read_parse = match sheet_part_kind {
        SheetPartKind::ChartSheet => parse_chart_sheet_shell(archive, &sheet_part, name),
        SheetPartKind::DialogSheet => parse_dialog_sheet_shell(archive, &sheet_part, name),
        SheetPartKind::Worksheet => stream_sheet_data_from_archive(
            archive,
            &sheet_part,
            Rc::clone(&shared.shared_strings),
            Rc::clone(&shared.theme_colors),
        )
        .and_then(|streamed| parse_projected_worksheet(streamed, theme_colors, name)),
    };
    let parsed = match sheet_read_parse {
        Ok(parsed) => parsed,
        Err(detail) => {
            let ws = Worksheet::placeholder(name, format!("{sheet_part}: {detail}"));
            return serialize_worksheet_bounded(archive, &sheet_part, &ws);
        }
    };
    // Dialog sheets are legacy form definitions, not worksheet grids. Their
    // DrawingML/control relationships are not displayable as worksheet content,
    // so do not spend bounded package resources materializing content the
    // renderer intentionally replaces with an informational surface.
    if sheet_part_kind == SheetPartKind::DialogSheet {
        return serialize_worksheet_bounded(archive, &sheet_part, &parsed.0);
    }
    let worksheet = finalize_projected_sheet(
        archive,
        shared,
        sheet_index,
        name,
        &sheet_path,
        parsed,
        CurrentSheetLookup::BuildFromMaterializedRows,
    )?;
    serialize_worksheet_bounded(archive, &sheet_part, &worksheet)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SheetPartKind {
    Worksheet,
    ChartSheet,
    DialogSheet,
}

fn resolve_sheet_part_kind(doc: &roxmltree::Document, r_id: &str) -> SheetPartKind {
    const PACKAGE_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    let relationship_type = doc.descendants().find_map(|node| {
        (node.is_element()
            && node.tag_name().name() == "Relationship"
            && node.tag_name().namespace() == Some(PACKAGE_RELATIONSHIPS)
            && node.attribute("Id") == Some(r_id))
        .then(|| node.attribute("Type"))
        .flatten()
    });
    match relationship_type {
        Some(value) if is_office_relationship_type(value, "chartsheet") => {
            SheetPartKind::ChartSheet
        }
        Some(value) if is_office_relationship_type(value, "dialogsheet") => {
            SheetPartKind::DialogSheet
        }
        _ => SheetPartKind::Worksheet,
    }
}

fn is_office_relationship_type(value: &str, local_name: &str) -> bool {
    [relationships::TRANSITIONAL, relationships::STRICT]
        .into_iter()
        .any(|base| {
            value
                .strip_prefix(base)
                .and_then(|suffix| suffix.strip_prefix('/'))
                == Some(local_name)
        })
}

fn parse_chart_sheet_shell(
    archive: &mut XlsxZip,
    sheet_part: &str,
    name: &str,
) -> Result<(Worksheet, HyperlinkRids, String), String> {
    let xml = read_zip_string(archive, sheet_part)?;
    let document = parse_guarded(&xml).map_err(|error| error.to_string())?;
    let root = document.root_element();
    if root.tag_name().name() != "chartsheet" || !is_x_ns(root.tag_name().namespace()) {
        return Err("expected SpreadsheetML chartsheet root".to_string());
    }
    Ok((Worksheet::chart_sheet(name), Vec::new(), xml))
}

fn parse_dialog_sheet_shell(
    archive: &mut XlsxZip,
    sheet_part: &str,
    name: &str,
) -> Result<(Worksheet, HyperlinkRids, String), String> {
    let xml = read_zip_string(archive, sheet_part)?;
    let document = parse_guarded(&xml).map_err(|error| error.to_string())?;
    let root = document.root_element();
    if root.tag_name().name() != "dialogsheet" || !is_x_ns(root.tag_name().namespace()) {
        return Err("expected SpreadsheetML dialogsheet root".to_string());
    }
    Ok((Worksheet::dialog_sheet(name), Vec::new(), xml))
}

enum CurrentSheetLookup {
    BuildFromMaterializedRows,
    Seed(Option<WorksheetCellLookup>),
}

fn finalize_projected_sheet(
    archive: &mut XlsxZip,
    shared: &WorkbookShared,
    sheet_index: u32,
    name: &str,
    sheet_path: &str,
    parsed: (Worksheet, HyperlinkRids, String),
    current_lookup: CurrentSheetLookup,
) -> Result<Worksheet, String> {
    let wb_doc = parse_guarded(&shared.workbook_xml).map_err(|e| e.to_string())?;
    let rels_doc = parse_guarded(&shared.rels_xml).map_err(|e| e.to_string())?;
    let theme_colors = shared.theme_colors.as_ref();
    let (mut ws, hyperlink_rids, sheet_shell_xml) = parsed;

    // Attach any drawing-anchored images and charts for this sheet
    ws.images = load_sheet_images(archive, sheet_path, theme_colors);
    // Embedded OLE object previews (the `<oleObjects>` collection, §18.3.1.60)
    // draw through the same image
    // list; their preview parts are referenced from the worksheet XML + rels.
    ws.images
        .extend(load_sheet_ole_images(archive, sheet_path, &sheet_shell_xml));
    let mut reference_session = WorksheetReferenceSession::default();
    let materialized_rows = match current_lookup {
        CurrentSheetLookup::BuildFromMaterializedRows => Some(ws.rows.as_slice()),
        CurrentSheetLookup::Seed(lookup) => {
            reference_session.seed_current_sheet(name, lookup);
            None
        }
    };
    let defined_names = parse_defined_names_for_sheet(&wb_doc, sheet_index);
    let charts = load_sheet_charts_with_theme_images(
        archive,
        sheet_path,
        Some(ChartReferenceContext {
            materialized_rows,
            materialized_col_hidden: Some(&ws.col_hidden),
            sheet_name: name,
            sheets: &shared.sheets,
            workbook_rels: &rels_doc,
            shared_strings: shared.shared_strings.as_ref(),
            defined_names: &defined_names,
            number_formats: &shared.chart_number_formats,
            session: &mut reference_session,
        }),
        theme_colors,
        (
            shared.theme_fonts.0.as_deref(),
            shared.theme_fonts.1.as_deref(),
        ),
        Some(shared.theme_format_scheme.as_ref()),
        shared.theme_chart_images.as_ref(),
    );
    ws.charts = charts;
    ws.shape_groups = load_sheet_shape_groups(
        archive,
        sheet_path,
        theme_colors,
        shared.theme_format_scheme.as_ref(),
    );
    ws.hyperlinks = load_hyperlinks(archive, sheet_path, hyperlink_rids);
    ws.comments = load_sheet_comments(archive, sheet_path, &shared.rels_xml);
    ws.comment_refs = ws.comments.iter().map(|c| c.cell_ref.clone()).collect();
    ws.defined_names = defined_names;
    ws.tables = load_sheet_tables(archive, sheet_path, theme_colors);
    ws.slicers = load_sheet_slicers(archive, sheet_path, theme_colors);
    (ws.pivot_tables, ws.pivot_diagnostics) = load_sheet_pivots(archive, sheet_path);
    let sparkline_groups = load_sheet_sparklines(
        archive,
        &sheet_shell_xml,
        materialized_rows,
        name,
        &shared.sheets,
        &rels_doc,
        theme_colors,
        shared.shared_strings.as_ref(),
        &mut reference_session,
    );
    ws.sparkline_groups = sparkline_groups;
    ws.default_font_family = shared.default_font.0.clone();
    ws.default_font_size = shared.default_font.1;
    // Denormalize the workbook-wide date system onto this sheet so the cell
    // formatter can resolve serial dates without a workbook back-reference
    // (ECMA-376 §18.2.28 / §18.17.4.1).
    ws.date1904 = shared.date1904;

    Ok(ws)
}

fn parse_xlsx_inner_with(
    archive: &mut XlsxZip,
    shared: &WorkbookShared,
    styles: Result<Styles, String>,
) -> Result<ParsedWorkbook, String> {
    let theme_colors = shared.theme_colors.as_ref();
    let styles = styles?;

    // Surface each sheet's tab color (`<sheetPr><tabColor>`) on the workbook
    // sheet list so the viewer can paint every tab up front. `<sheetPr>` is the
    // first child of `<worksheet>` (ECMA-376 §18.3.1.99 element order), so a
    // small head read of each sheet entry is enough — we never decompress the
    // (potentially huge) `<sheetData>` body just to read the tab color.
    let mut sheets = shared.sheets.clone();
    if let Ok(rels_doc) = parse_guarded(&shared.rels_xml) {
        for sheet in sheets.iter_mut() {
            let Some(path) = resolve_sheet_path(&rels_doc, &sheet.r_id) else {
                continue;
            };
            let Ok(head) = read_zip_entry_head(archive, &format!("xl/{}", path), 16_384) else {
                continue;
            };
            sheet.tab_color = extract_tab_color_from_head(&head, theme_colors);
        }
    }

    Ok(ParsedWorkbook {
        workbook: Workbook {
            sheets,
            date1904: shared.date1904,
            // #773: a corrupt-but-present `xl/sharedStrings.xml` surfaces here as a
            // workbook-level, part-tagged error so the blanked string cells across
            // all sheets are visible rather than silent. Every sheet still opens.
            parse_error: shared.shared_strings_error.clone(),
        },
        styles,
        shared_strings: shared.shared_strings.as_ref().to_vec(),
    })
}

#[cfg(test)]
mod retained_model_limit_tests {
    use super::*;
    use std::io::Write;

    fn cell(row: u32, value: CellValue, formula: Option<&str>) -> Cell {
        Cell {
            col: 1,
            row,
            value,
            style_index: None,
            formula: formula.map(str::to_string),
            show_phonetic: false,
        }
    }

    fn row(index: u32, cells: Vec<Cell>) -> Row {
        Row {
            index,
            height: None,
            custom_height: false,
            cells,
            outline_level: 0,
            collapsed: false,
            hidden: false,
        }
    }

    #[test]
    fn counting_writer_matches_exact_serde_json_bytes() {
        let value = serde_json::json!({
            "ascii": "quote\" slash\\ newline\n",
            "unicode": "é😀",
            "array": [null, true, -0.0, "x"]
        });
        let measured = measure_json(&value).unwrap();
        assert_eq!(
            measured.json_bytes as usize,
            serde_json::to_vec(&value).unwrap().len()
        );
    }

    #[test]
    fn counting_writer_decodes_short_and_unicode_escapes_but_excludes_keys() {
        let value = serde_json::json!({
            "ignored-key-é": "\u{0008}\t\n\u{000c}\r",
            "also-ignored": "\u{0000}\u{000b}\u{001f}",
            "unicode-key-😀": "é😀"
        });
        let measured = measure_json(&value).unwrap();
        assert_eq!(
            measured.json_bytes as usize,
            serde_json::to_vec(&value).unwrap().len()
        );
        assert_eq!(
            measured.string_value_utf8_bytes,
            5 + 3 + "é😀".len() as u64,
            "only decoded string values are retained-content bytes"
        );
    }

    #[test]
    fn cell_string_measurement_expands_every_shared_reference() {
        let shared = vec![SharedString {
            text: "é".to_string(),
            ..Default::default()
        }];
        let rows = vec![
            row(1, vec![cell(1, CellValue::Shared { si: 0 }, Some("A1"))]),
            row(2, vec![cell(2, CellValue::Shared { si: 0 }, None)]),
            row(
                3,
                vec![cell(
                    3,
                    CellValue::Error {
                        error: "#VALUE!".to_string(),
                    },
                    None,
                )],
            ),
        ];
        // Each shared cell retains the resolved `text` discriminator (4) and
        // the two-byte UTF-8 value. The first also owns its formula. Error owns
        // its discriminator and payload.
        assert_eq!(
            row_cell_content_utf8_bytes(&rows, &shared).unwrap(),
            (4 + 2 + 2) + (4 + 2) + (5 + 7)
        );
    }

    #[test]
    fn escaped_controls_reach_the_hard_cell_content_boundary_and_plus_one() {
        let payload_bytes = usize::try_from(HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES - 4)
            .expect("hard limit fits usize");
        let mut text = "\0".repeat(payload_bytes);
        let shared: Vec<SharedString> = Vec::new();
        let exact = vec![row(
            1,
            vec![cell(
                1,
                CellValue::Text {
                    text: text.clone(),
                    runs: None,
                    phonetic_runs: Vec::new(),
                    phonetic_pr: None,
                },
                None,
            )],
        )];
        assert_eq!(
            row_cell_content_utf8_bytes(&exact, &shared).unwrap(),
            HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES
        );

        text.push('\0');
        let over = vec![row(
            1,
            vec![cell(
                1,
                CellValue::Text {
                    text,
                    runs: None,
                    phonetic_runs: Vec::new(),
                    phonetic_pr: None,
                },
                None,
            )],
        )];
        assert_eq!(
            row_cell_content_utf8_bytes(&over, &shared).unwrap(),
            HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES + 1
        );

        let mut package = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut package));
            writer
                .start_file(
                    "xl/worksheets/sheet1.xml",
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let mut archive = XlsxZip::new(Cursor::new(package)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let reporter = archive.operation().unwrap().limit_reporter().unwrap();
        reporter
            .observe_hard_limit(
                HardResourceLimitKind::WorksheetCellContentOwnedUtf8Bytes,
                Some("xl/worksheets/sheet1.xml"),
                HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
                HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
            )
            .unwrap();
        let error = reporter
            .observe_hard_limit(
                HardResourceLimitKind::WorksheetCellContentOwnedUtf8Bytes,
                Some("xl/worksheets/sheet1.xml"),
                HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
                HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES + 1,
            )
            .unwrap_err();
        let envelope: serde_json::Value =
            serde_json::from_str(error.strip_prefix("OOXML_RESOURCE_LIMIT:").unwrap()).unwrap();
        assert_eq!(
            envelope["details"]["violation"]["resource"],
            "worksheet-cell-content"
        );
        assert_eq!(
            envelope["details"]["violation"]["observed"],
            HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES + 1
        );
    }
}

fn parse_xlsx_inner(data: &[u8]) -> Result<ParsedWorkbook, String> {
    parse_xlsx_inner_with_limits(data, None, None)
}

fn parse_xlsx_inner_with_limits(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<ParsedWorkbook, String> {
    // #774 (RB7 MAJOR): a corrupt / truncated CONTAINER degrades to a placeholder
    // workbook (one placeholder sheet) rather than erroring, consistent with a
    // corrupt inner sheet — the viewer shows a "could not display" tab instead of
    // nothing.
    let mut archive = match open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    ) {
        Ok(zip) => zip,
        Err(error) if error.starts_with("OOXML_RESOURCE_LIMIT:") => return Err(error),
        Err(e) => return Ok(degraded_container_workbook(e)),
    };
    archive.run_operation("parse", |archive| {
        let (shared, styles) = WorkbookShared::load_with_styles(archive)?;
        parse_xlsx_inner_with(archive, &shared, styles)
    })
}

/// Read only the first `max_bytes` of a ZIP entry as text. Used to probe the
/// top of a worksheet (its `<sheetPr>`) without inflating the whole sheet.
/// Lossy UTF-8 keeps a multibyte character split at the cut from erroring; the
/// region we care about (`<sheetPr><tabColor>`) is pure ASCII near the start.
fn read_zip_entry_head(
    archive: &mut XlsxZip,
    name: &str,
    max_bytes: u64,
) -> Result<String, String> {
    let max_bytes = usize::try_from(max_bytes).unwrap_or(usize::MAX);
    read_zip_string_head(archive, name, max_bytes)
}

/// Extract the resolved tab color from the head of a worksheet XML. Locates the
/// single `<tabColor .../>` element (it lives in `<sheetPr>`, before
/// `<sheetData>`) and resolves its `rgb` / `theme`+`tint` / `indexed` attributes
/// through the same rules as cell colors. Returns `None` when no tab color is
/// declared or the tag is truncated by the head limit. A lightweight attribute
/// scan avoids any namespace-prefix assumptions in the partial document.
fn extract_tab_color_from_head(head: &str, theme_colors: &[String]) -> Option<String> {
    // Don't look past the data body — `tabColor` only appears in `<sheetPr>`.
    let scope = head.split("<sheetData").next().unwrap_or(head);
    let start = scope.find("tabColor")?;
    let rest = &scope[start..];
    // The element is self-closing (`<tabColor ... />`); read up to its `>`.
    let end = rest.find('>')?;
    let tag = &rest[..end];
    let attr = |name: &str| -> Option<&str> {
        let key = format!("{}=\"", name);
        let i = tag.find(&key)? + key.len();
        let j = tag[i..].find('"')? + i;
        Some(&tag[i..j])
    };
    resolve_color_attrs(
        attr("rgb"),
        attr("theme"),
        attr("tint"),
        attr("indexed"),
        theme_colors,
    )
}

/// Convert hex color + tint to resulting hex color using HLS model.
/// tint > 0: lighten; tint < 0: darken.
fn apply_tint(hex: &str, tint: f64) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 {
        return format!("#{}", hex);
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

    // RGB → HLS
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let s = if max == min {
        0.0
    } else if l < 0.5 {
        (max - min) / (max + min)
    } else {
        (max - min) / (2.0 - max - min)
    };
    let h = if max == min {
        0.0
    } else if max == r {
        (g - b) / (max - min) / 6.0
    } else if max == g {
        ((b - r) / (max - min) + 2.0) / 6.0
    } else {
        ((r - g) / (max - min) + 4.0) / 6.0
    };
    let h = if h < 0.0 { h + 1.0 } else { h };

    // Apply tint to luminance
    let new_l = if tint > 0.0 {
        l * (1.0 - tint) + tint
    } else {
        l * (1.0 + tint)
    };

    // HLS → RGB
    let (nr, ng, nb) = hls_to_rgb(h, new_l, s);
    format!(
        "#{:02X}{:02X}{:02X}",
        (nr * 255.0).round() as u8,
        (ng * 255.0).round() as u8,
        (nb * 255.0).round() as u8
    )
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

pub(crate) fn parse_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    resolve_color_attrs(
        node.attribute("rgb"),
        node.attribute("theme"),
        node.attribute("tint"),
        node.attribute("indexed"),
        theme_colors,
    )
}

/// Resolve a DrawingML/SpreadsheetML color from its raw attribute values
/// (`rgb` / `theme` + `tint` / `indexed`). Split out from [`parse_color`] so
/// callers that scan attributes without a roxmltree node (e.g. the bounded
/// tab-color head probe) share the exact same resolution rules.
pub(crate) fn resolve_color_attrs(
    rgb: Option<&str>,
    theme: Option<&str>,
    tint: Option<&str>,
    indexed: Option<&str>,
    theme_colors: &[String],
) -> Option<String> {
    // rgb attribute (ARGB: 8 chars, drop alpha; or 6-char RGB)
    if let Some(rgb) = rgb {
        if rgb.len() == 8 {
            return Some(format!("#{}", rgb[2..].to_uppercase()));
        }
        return Some(format!("#{}", rgb.to_uppercase()));
    }

    // theme attribute → resolve from theme color array + optional tint
    //
    // ECMA-376 §18.8.3 stores the theme clrScheme in the order
    //   dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink
    // but cell style references (c:color/@theme, c:fgColor/@theme, etc.) use
    // the Excel-internal index where dk1↔lt1 and dk2↔lt2 are SWAPPED:
    //   0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..11 unchanged.
    // This is a well-known interoperability quirk (see Open-XML-SDK issue #46
    // and ECMA-376 §22.1.2.7 where "index values of 0 and 1 are swapped").
    // This is an index→index remap, not a logical→slot-name mapping, so the
    // shared ooxml_common::color::SCHEME_DEFAULT_SLOTS table (the canonical
    // §19.3.1.6 logical→slot names) does not apply here; this stays local.
    if let Some(theme_str) = theme {
        if let Ok(idx) = theme_str.parse::<usize>() {
            let mapped = match idx {
                0 => 1,
                1 => 0,
                2 => 3,
                3 => 2,
                n => n,
            };
            if let Some(base) = theme_colors.get(mapped) {
                let tint = tint.and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                if tint == 0.0 {
                    return Some(base.clone());
                }
                return Some(apply_tint(base, tint));
            }
        }
    }

    // indexed attribute → Excel built-in palette
    if let Some(indexed_str) = indexed {
        if let Ok(idx) = indexed_str.parse::<usize>() {
            // indices 64 (foreground) and 65 (background) are special: use black/white
            let color = match idx {
                64 => "#000000",
                65 => "#FFFFFF",
                _ => INDEXED_COLORS.get(idx).copied().unwrap_or("#000000"),
            };
            return Some(color.to_string());
        }
    }

    None
}

/// Parse the workbook-level date system from `<workbookPr date1904>`
/// (ECMA-376 §18.2.28). The attribute is an xsd:boolean; `"1"` or `"true"`
/// select the 1904 date system. Absent attribute / element ⇒ false (the
/// default 1900 date system). See §18.17.4.1 for the date-system definitions.
fn parse_workbook_date1904(doc: &roxmltree::Document) -> bool {
    doc.descendants()
        .find(|n| n.tag_name().name() == "workbookPr" && is_x_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("date1904"))
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false)
}

fn parse_workbook_sheets(doc: &roxmltree::Document) -> Vec<SheetMeta> {
    let mut sheets = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "sheet" && is_x_ns(node.tag_name().namespace()) {
            let name = node.attribute("name").unwrap_or("Sheet").to_string();
            let sheet_id = node
                .attribute("sheetId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1);
            let r_id = attr_ns(
                &node,
                relationships::TRANSITIONAL,
                relationships::STRICT,
                "id",
            )
            .unwrap_or("")
            .to_string();
            let visibility = match node.attribute("state") {
                Some("hidden") => SheetVisibility::Hidden,
                Some("veryHidden") => SheetVisibility::VeryHidden,
                _ => SheetVisibility::Visible,
            };
            sheets.push(SheetMeta {
                name,
                sheet_id,
                r_id,
                tab_color: None,
                visibility,
            });
        }
    }
    sheets
}

/// Collect `<definedName>` entries from `workbook.xml`. `sheet_index` selects
/// which names are in scope: workbook-global (no `localSheetId`) plus any
/// whose `localSheetId` matches the given sheet position.
fn parse_defined_names_for_sheet(doc: &roxmltree::Document, sheet_index: u32) -> Vec<DefinedName> {
    let mut names = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() != "definedName" || !is_x_ns(node.tag_name().namespace()) {
            continue;
        }
        let local: Option<u32> = node.attribute("localSheetId").and_then(|s| s.parse().ok());
        if let Some(l) = local {
            if l != sheet_index {
                continue;
            }
        }
        let name = match node.attribute("name") {
            Some(n) => n.to_string(),
            None => continue,
        };
        let formula = node.text().unwrap_or("").to_string();
        names.push(DefinedName { name, formula });
    }
    names
}

pub(crate) fn resolve_sheet_path(doc: &roxmltree::Document, r_id: &str) -> Option<String> {
    let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship"
            && node.tag_name().namespace() == Some(ns)
            && node.attribute("Id") == Some(r_id)
        {
            let target = node.attribute("Target")?;
            // ECMA-376 / Open Packaging Conventions: Target may be a
            // package-absolute path (`/xl/worksheets/sheet1.xml`, used
            // by openpyxl and some online tools) or a path relative to
            // the .rels file's parent (`worksheets/sheet1.xml`, the
            // common Office-saved form). Callers prepend `xl/` to the
            // returned value, so strip a leading `/xl/` to convert
            // absolute paths into the relative form they expect.
            let t = target.strip_prefix('/').unwrap_or(target);
            let t = t.strip_prefix("xl/").unwrap_or(t);
            return Some(t.to_string());
        }
    }
    None
}

/// Read + parse `xl/sharedStrings.xml` (§18.4.9) into the dedup'd string table.
///
/// Returns the strings plus an optional **part-tagged degradation error** (#773).
/// The two failure modes are treated differently, on purpose:
///
///  - **Missing part** (`read_zip_string` fails): NOT an error. A workbook with no
///    string cells legitimately ships no `sharedStrings.xml`, so an absent part is
///    the normal empty-table case — `(Vec::new(), None)`.
///  - **Present but corrupt** (`parse_guarded` fails on a part that IS in the zip):
///    a real degradation. Before #773 this returned an empty table silently, so
///    EVERY string cell across ALL sheets rendered blank with no indication why.
///    Now it returns `(Vec::new(), Some("xl/sharedStrings.xml: <detail>"))` so the
///    caller can surface a workbook-level `parse_error`. We still return the empty
///    table (not an `Err`) so the workbook keeps opening and every sheet renders
///    its non-string content — partial degradation, just no longer silent.
fn read_shared_strings(
    archive: &mut XlsxZip,
    theme_colors: &[String],
) -> (Vec<SharedString>, Option<String>) {
    let Ok(xml) = read_zip_string(archive, "xl/sharedStrings.xml") else {
        // Absent part ⇒ legitimately empty table, not a degradation.
        return (Vec::new(), None);
    };
    let doc = match parse_guarded(&xml) {
        Ok(doc) => doc,
        Err(e) => {
            // Present but unparseable ⇒ surface it so the blank string cells aren't
            // a silent mystery; keep the workbook openable with an empty table.
            return (Vec::new(), Some(format!("xl/sharedStrings.xml: {e}")));
        }
    };
    let mut strings = Vec::new();
    for si in doc.descendants() {
        if si.tag_name().name() == "si" && is_x_ns(si.tag_name().namespace()) {
            strings.push(parse_si_node(&si, theme_colors));
        }
    }
    (strings, None)
}

/// Parse a `<si>` (shared) or `<is>` (inline) node into a SharedString.
/// The node may contain direct `<t>` text (plain) and/or multiple `<r>`
/// runs with per-run `<rPr>` font properties.
fn parse_si_node(node: &roxmltree::Node, theme_colors: &[String]) -> SharedString {
    let mut text = String::new();
    let mut runs: Vec<Run> = Vec::new();
    let mut has_runs = false;
    // ECMA-376 §18.4.6 `<rPh>` phonetic runs (furigana) and §18.4.3
    // `<phoneticPr>` display properties. Accumulated alongside the base text so
    // a String Item's reading rides with the string without polluting `text`.
    let mut phonetic_runs: Vec<PhoneticRun> = Vec::new();
    let mut phonetic_pr: Option<PhoneticProperties> = None;
    for child in node.children() {
        if !child.is_element() {
            continue;
        }
        match child.tag_name().name() {
            "t" if is_x_ns(child.tag_name().namespace()) => {
                if let Some(s) = child.text() {
                    text.push_str(s);
                }
            }
            "rPh" if is_x_ns(child.tag_name().namespace()) => {
                // §18.4.6: sb/eb are zero-based base-text character offsets
                // (sb < eb). The hint text sits in the child <t>.
                let sb: u32 = child
                    .attribute("sb")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
                let eb: u32 = child
                    .attribute("eb")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
                let mut rph_text = String::new();
                for rc in child.children() {
                    if rc.tag_name().name() == "t" {
                        if let Some(s) = rc.text() {
                            rph_text.push_str(s);
                        }
                    }
                }
                phonetic_runs.push(PhoneticRun {
                    sb,
                    eb,
                    text: rph_text,
                });
            }
            "phoneticPr" if is_x_ns(child.tag_name().namespace()) => {
                // §18.4.3: fontId required (0-based into styles fonts); type
                // defaults to fullwidthKatakana, alignment to left. We carry
                // the raw enum strings; the renderer applies the defaults.
                let font_id: u32 = child
                    .attribute("fontId")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
                phonetic_pr = Some(PhoneticProperties {
                    font_id,
                    r#type: child.attribute("type").map(|s| s.to_string()),
                    alignment: child.attribute("alignment").map(|s| s.to_string()),
                });
            }
            "r" if is_x_ns(child.tag_name().namespace()) => {
                has_runs = true;
                let mut run_text = String::new();
                let mut run_font: Option<RunFont> = None;
                for rc in child.children() {
                    match rc.tag_name().name() {
                        "t" => {
                            if let Some(s) = rc.text() {
                                run_text.push_str(s);
                            }
                        }
                        "rPr" => {
                            let mut f = RunFont::default();
                            for rp in rc.children() {
                                match rp.tag_name().name() {
                                    "b" => f.bold = parse_st_on_off(&rp),
                                    "i" => f.italic = parse_st_on_off(&rp),
                                    "u" => {
                                        // ECMA-376 §18.4.13 ST_UnderlineValues:
                                        // single (default) | double | singleAccounting |
                                        // doubleAccounting | none.
                                        let v = rp.attribute("val").unwrap_or("single");
                                        if v != "none" {
                                            f.underline = true;
                                            if v != "single" {
                                                f.underline_style = Some(v.to_string());
                                            }
                                        }
                                    }
                                    "strike" => f.strike = parse_st_on_off(&rp),
                                    "vertAlign" => {
                                        // ECMA-376 §18.4.6 ST_VerticalAlignRun.
                                        if let Some(v) = rp.attribute("val") {
                                            if v != "baseline" {
                                                f.vert_align = Some(v.to_string());
                                            }
                                        }
                                    }
                                    "sz" => {
                                        f.size = rp.attribute("val").and_then(|s| s.parse().ok());
                                    }
                                    "color" => {
                                        f.color = parse_color(&rp, theme_colors);
                                    }
                                    "rFont" | "name" => {
                                        f.name = rp.attribute("val").map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }
                            run_font = Some(f);
                        }
                        _ => {}
                    }
                }
                text.push_str(&run_text);
                runs.push(Run {
                    text: run_text,
                    font: run_font,
                });
            }
            _ => {}
        }
    }
    SharedString {
        text,
        runs: if has_runs { Some(runs) } else { None },
        phonetic_runs,
        phonetic_pr,
    }
}
/// Pending cell-hyperlink descriptors, awaiting rels resolution of the external
/// `r:id`. Each entry is `(col, row, rid, location, display)`:
/// - `rid`: the external relationship id (§18.3.1.47 `r:id`), if present.
/// - `location`: the inline internal target (§18.3.1.47 `location`) — a defined
///   name or cell ref like `Sheet1!A1`. No rels lookup required.
/// - `display`: the optional display text (§18.3.1.47 `display`).
///
/// A `<hyperlink>` may carry `rid`, `location`, or both, so both are optional.
type HyperlinkRids = Vec<(u32, u32, Option<String>, Option<String>, Option<String>)>;

/// Incrementally inflate and project one worksheet through the active package
/// operation. Rows are kept private and provisional until the projector emits
/// `Finished`; any XML, callback, or decoder failure drops the partial vector.
///
/// Decoder failures can be wrapped by quick-xml as an I/O error. Re-checking
/// the session at this trust boundary restores the canonical, latched resource
/// envelope instead of exposing quick-xml's wrapper text.
fn stream_sheet_data_from_archive(
    archive: &mut XlsxZip,
    part: &str,
    shared_strings: Rc<[SharedString]>,
    theme_colors: Rc<[String]>,
) -> Result<StreamedSheetData, String> {
    let reporter = archive.operation()?.limit_reporter()?;
    let mut cursor =
        archive.open_worksheet_cursor(part, Rc::clone(&shared_strings), theme_colors)?;
    let mut rows = Vec::new();
    let mut usage = WorksheetModelUsage::default();
    loop {
        match cursor.pull(
            WORKSHEET_CURSOR_PULL_ROWS,
            WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES,
        ) {
            Ok(WorksheetCursorPull::Rows {
                rows: batch,
                projected_bytes,
            }) => {
                if !batch.is_empty() && projected_bytes == 0 {
                    return Err("worksheet cursor returned an unmeasured row batch".to_string());
                }
                let next_rows = usage
                    .rows
                    .checked_add(batch.len() as u64)
                    .ok_or_else(|| "worksheet row count overflow".to_string())?;
                let batch_cells = batch.iter().try_fold(0u64, |total, row| {
                    total
                        .checked_add(row.cells.len() as u64)
                        .ok_or_else(|| "worksheet cell count overflow".to_string())
                })?;
                let next_cells = usage
                    .cells
                    .checked_add(batch_cells)
                    .ok_or_else(|| "worksheet cell count overflow".to_string())?;
                let batch_strings = row_cell_content_utf8_bytes(&batch, shared_strings.as_ref())?;
                let next_strings = usage
                    .owned_utf8_bytes
                    .checked_add(batch_strings)
                    .ok_or_else(|| "worksheet string measurement overflow".to_string())?;
                reporter.observe_hard_limit(
                    HardResourceLimitKind::WorksheetModelRows,
                    Some(part),
                    HARD_MAX_XLSX_WORKSHEET_ROWS,
                    next_rows,
                )?;
                reporter.observe_hard_limit(
                    HardResourceLimitKind::WorksheetModelCells,
                    Some(part),
                    HARD_MAX_XLSX_WORKSHEET_CELLS,
                    next_cells,
                )?;
                reporter.observe_hard_limit(
                    HardResourceLimitKind::WorksheetCellContentOwnedUtf8Bytes,
                    Some(part),
                    HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
                    next_strings,
                )?;
                usage.rows = next_rows;
                usage.cells = next_cells;
                usage.owned_utf8_bytes = next_strings;
                rows.extend(batch);
            }
            Ok(WorksheetCursorPull::Finished(tail)) => {
                archive.assert_healthy()?;
                return Ok(StreamedSheetData {
                    shell_xml: tail.shell_xml,
                    rows,
                    row_heights: tail.row_heights,
                });
            }
            Err(error) => {
                cursor.cancel();
                return Err(archive.assert_healthy().err().unwrap_or(error));
            }
        }
    }
}

#[cfg(test)]
fn parse_worksheet_with_shell(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, HyperlinkRids, String), String> {
    let streamed = worksheet_projector::stream_sheet_data(xml, shared_strings, theme_colors)?;
    parse_projected_worksheet(streamed, theme_colors, name)
}

fn parse_projected_worksheet(
    streamed: StreamedSheetData,
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, HyperlinkRids, String), String> {
    // Guard against a pathologically deep worksheet XML: the nesting-depth
    // pre-check now lives inside `parse_guarded`, which rejects an over-deep
    // part before roxmltree's tree builder can recurse and overflow the fixed
    // WASM stack. See `ooxml_common::depth::parse_guarded`.
    let doc = parse_guarded(&streamed.shell_xml).map_err(|e| e.to_string())?;

    let mut rows = streamed.rows;
    let mut col_widths: BTreeMap<u32, f64> = BTreeMap::new();
    // All authored width declarations in XML order. The boolean records
    // whether the declaration also needs the compact public wire form.
    let mut authored_col_widths: Vec<(crate::types::ColumnWidthRange, bool)> = Vec::new();
    let mut authored_col_styles: Vec<crate::types::ColumnStyleRange> = Vec::new();
    let mut row_heights = streamed.row_heights;
    // Outline (grouping) metadata — ECMA-376 §18.3.1.13 (col) / §18.3.1.73
    // (row) / §18.3.1.61 (outlinePr). Only non-default entries are recorded so
    // an outline-free sheet keeps empty maps / a `None` outlinePr (byte-stable
    // JSON, §CLAUDE "1px identical").
    let mut col_outline_levels: BTreeMap<u32, u8> = BTreeMap::new();
    let mut col_collapsed: BTreeMap<u32, bool> = BTreeMap::new();
    let mut col_hidden: BTreeMap<u32, bool> = BTreeMap::new();
    let mut outline_pr: Option<crate::types::OutlinePr> = None;
    let mut merge_cells: Vec<MergeCell> = Vec::new();
    let mut freeze_rows: u32 = 0;
    let mut freeze_cols: u32 = 0;
    let mut default_col_width = 8.43;
    // Intrinsic default row height in *points* — ECMA-376 §18.3.1.81.
    // 15 pt = 20 CSS px at 96 DPI, Excel's baseline for the Calibri 11
    // Normal style. The renderer multiplies by 4/3 at display time, so
    // both this default and per-row `<row ht="…">` values share the
    // same units across the parser/renderer boundary.
    let mut default_row_height = 15.0;
    let mut default_row_height_custom = false;
    let mut rows_hidden_by_default = false;
    let mut conditional_formats: Vec<ConditionalFormat> = Vec::new();
    let mut show_zeros = true;
    let mut show_gridlines = true;
    let mut right_to_left = false;
    let mut tab_color: Option<String> = None;
    let mut auto_filter: Option<CellRange> = None;
    let mut hyperlink_rids: HyperlinkRids = Vec::new();
    // ECMA-376 §18.3.1.73 (CT_Row, sml.xsd) makes `@r` on `<row>` optional with
    // no default; the spec does not spell out how an omitted value is resolved.
    // We follow the de-facto consumer convention (Excel, LibreOffice, SheetJS
    // agree; no competing interpretation exists): an r-less row is the previous
    // row's number + 1 (the first row is 1), and an explicit `@r` re-anchors
    // this counter. `prev_row_idx == 0` means "no row yet", so the first
    // implicit row lands at index 1.
    // Pre-scan worksheet-level extLst for x14:dataBar extension attributes.
    // Excel 2010+ stores the `gradient` flag on `<x14:dataBar>` inside
    // `<extLst>/<ext>/<x14:conditionalFormattings>/<x14:conditionalFormatting>
    // /<x14:cfRule id="{GUID}">`, linked to the SpreadsheetML cfRule via a
    // matching `<x14:id>{GUID}</x14:id>` inside the cfRule's own extLst
    // (§2.6.3). Build a GUID → gradient map so cfRule parsing can look up
    // the override.
    let mut x14_databar_gradient: HashMap<String, bool> = HashMap::new();
    for x14_rule in doc
        .descendants()
        .filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("dataBar"))
    {
        let Some(id) = x14_rule.attribute("id") else {
            continue;
        };
        for bar in x14_rule
            .children()
            .filter(|n| n.tag_name().name() == "dataBar")
        {
            if let Some(g) = bar.attribute("gradient") {
                x14_databar_gradient.insert(id.to_string(), !(g == "0" || g == "false"));
            }
        }
    }

    // Pre-scan worksheet-level extLst for x14:conditionalFormatting with
    // iconSet rules. Excel 2010+ stores custom icon sets (custom="1") here
    // with per-threshold `<x14:cfIcon iconSet="X" iconId="N"/>` overrides,
    // and cfvo values inside `<xm:f>` children instead of `val` attributes.
    // The sqref for x14 CF rules lives in a `<xm:sqref>` sibling.
    let mut x14_icon_formats: Vec<ConditionalFormat> = Vec::new();
    for x14_cf in doc.descendants().filter(|n| {
        n.tag_name().name() == "conditionalFormatting"
            && n.tag_name()
                .namespace()
                .map(|u| u.contains("/spreadsheetml/2009/9"))
                .unwrap_or(false)
    }) {
        let sqref: Vec<CellRange> = x14_cf
            .children()
            .find(|n| n.tag_name().name() == "sqref")
            .and_then(|n| n.text())
            .map(parse_sqref)
            .unwrap_or_default();
        if sqref.is_empty() {
            continue;
        }
        let mut rules: Vec<CfRule> = Vec::new();
        for x14_rule in x14_cf
            .children()
            .filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("iconSet"))
        {
            let priority: i32 = x14_rule
                .attribute("priority")
                .and_then(|s| s.parse().ok())
                .unwrap_or(0);
            let Some(icon_node) = x14_rule
                .children()
                .find(|n| n.tag_name().name() == "iconSet")
            else {
                continue;
            };
            let custom = icon_node
                .attribute("custom")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            let icon_set_name = icon_node
                .attribute("iconSet")
                .unwrap_or(if custom { "" } else { "3TrafficLights1" })
                .to_string();
            let reverse = icon_node
                .attribute("reverse")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            let mut cfvos: Vec<CfValue> = Vec::new();
            let mut custom_icons: Vec<CfIcon> = Vec::new();
            for ch in icon_node.children().filter(|n| n.is_element()) {
                match ch.tag_name().name() {
                    "cfvo" => {
                        let kind = ch.attribute("type").unwrap_or("percent").to_string();
                        // x14:cfvo stores the value in `<xm:f>` child; attribute val fallback.
                        let value = ch
                            .children()
                            .find(|n| n.tag_name().name() == "f")
                            .and_then(|n| n.text())
                            .map(|s| s.to_string())
                            .or_else(|| ch.attribute("val").map(|s| s.to_string()));
                        cfvos.push(CfValue { kind, value });
                    }
                    "cfIcon" => {
                        let set = ch.attribute("iconSet").unwrap_or("NoIcons").to_string();
                        let id = ch
                            .attribute("iconId")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        custom_icons.push(CfIcon {
                            icon_set: set,
                            icon_id: id,
                        });
                    }
                    _ => {}
                }
            }
            rules.push(CfRule::IconSet {
                icon_set: icon_set_name,
                cfvos,
                reverse,
                priority,
                custom_icons: if custom { Some(custom_icons) } else { None },
            });
        }
        if !rules.is_empty() {
            x14_icon_formats.push(ConditionalFormat { sqref, rules });
        }
    }

    for node in doc.descendants() {
        match node.tag_name().name() {
            "sheetFormatPr" if is_x_ns(node.tag_name().namespace()) => {
                if let Some(v) = node
                    .attribute("defaultColWidth")
                    .and_then(|s| s.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && *value >= 0.0)
                {
                    default_col_width = v;
                }
                // ECMA-376 §18.3.1.81 `defaultRowHeight` is the workbook
                // default row height in points. Always honor it when present
                // — `demo/sample-1` stores `defaultRowHeight="20.1"` (no
                // customHeight) and Excel uses that 20.1 pt as the default
                // for non-customized rows. `customHeight` is metadata about
                // how the height was set, not a gate on whether to honor it.
                if let Some(v) = node
                    .attribute("defaultRowHeight")
                    .and_then(|s| s.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && *value >= 0.0)
                {
                    default_row_height = v;
                }
                default_row_height_custom = attr_bool(&node, "customHeight").unwrap_or(false);
                // ECMA-376 §18.3.1.81: zeroHeight makes unspecified rows hidden
                // by default. Keep this fact separate from defaultRowHeight:
                // explicit, non-hidden rows still use the authored default
                // height even when they omit @ht.
                rows_hidden_by_default = attr_bool(&node, "zeroHeight").unwrap_or(false);
            }
            "col" if is_x_ns(node.tag_name().namespace()) => {
                let custom = attr_bool(&node, "customWidth").unwrap_or(false);
                let hidden = attr_bool(&node, "hidden").unwrap_or(false);
                let authored_style = node
                    .attribute("style")
                    .and_then(|value| value.parse::<u32>().ok());
                let authored_width = node
                    .attribute("width")
                    .and_then(|s| s.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && *value >= 0.0);
                // §18.3.1.13 outline metadata: `outlineLevel` (0-7) and the
                // summary-column `collapsed` flag. Recorded independently of the
                // width so a grouped column at the default width is still
                // surfaced to the gutter.
                let outline_level = node
                    .attribute("outlineLevel")
                    .and_then(|s| s.parse::<u8>().ok())
                    .unwrap_or(0)
                    .min(7);
                let collapsed = attr_bool(&node, "collapsed").unwrap_or(false);
                // ECMA-376 §18.3.1.13: `customWidth` records how the width was
                // established; it does not gate the authored `width` itself.
                // Also retain outline-only columns so their gutter metadata
                // reaches the viewer at the default width.
                let has_outline = outline_level > 0 || collapsed;
                if authored_width.is_none()
                    && authored_style.is_none()
                    && !custom
                    && !hidden
                    && !has_outline
                {
                    continue;
                }
                let min: u32 = node
                    .attribute("min")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(1)
                    .clamp(1, 16_384);
                let full_max: u32 = node
                    .attribute("max")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(min)
                    .clamp(1, 16_384);
                if full_max < min {
                    continue;
                }
                if let Some(style_index) = authored_style {
                    authored_col_styles.push(crate::types::ColumnStyleRange {
                        min,
                        max: full_max,
                        style_index,
                    });
                }
                let width: f64 = if hidden {
                    0.0
                } else {
                    authored_width.unwrap_or(default_col_width)
                };
                let point_max = full_max.min(min.saturating_add(255));
                let point_map_covers_width = (custom || hidden) && point_max == full_max;
                if hidden || authored_width.is_some() || custom {
                    authored_col_widths.push((
                        crate::types::ColumnWidthRange {
                            min,
                            max: full_max,
                            width,
                        },
                        !point_map_covers_width,
                    ));
                }
                // Keep the legacy point map for existing consumers while
                // bounding its wire size. Grid geometry applies the compact
                // full range first, then lets these point entries override it.
                let max = point_max;
                for c in min..=max {
                    // Only store a width entry when the column actually has a
                    // custom / hidden width; a default-width grouped column keeps
                    // the workbook default (no colWidths entry) so its rendered
                    // width is byte-identical to an ungrouped default column.
                    if custom || hidden {
                        col_widths.insert(c, width);
                    }
                    if outline_level > 0 {
                        col_outline_levels.insert(c, outline_level);
                    }
                    if collapsed {
                        col_collapsed.insert(c, true);
                    }
                    if hidden {
                        col_hidden.insert(c, true);
                    }
                }
            }
            "sheetView" if is_x_ns(node.tag_name().namespace()) => {
                show_zeros = attr_bool(&node, "showZeros").unwrap_or(true);
                show_gridlines = attr_bool(&node, "showGridLines").unwrap_or(true);
                // ECMA-376 §18.3.1.87 `rightToLeft` — mirrors the whole grid so
                // column A is on the right. Default false (left-to-right).
                right_to_left = attr_bool(&node, "rightToLeft").unwrap_or(false);
            }
            "outlinePr" if is_x_ns(node.tag_name().namespace()) => {
                // §18.3.1.61 `<sheetPr><outlinePr>`. Both flags default to true
                // (summary below/right of detail). `applyStyles` is out of scope.
                outline_pr = Some(crate::types::OutlinePr {
                    summary_below: attr_bool(&node, "summaryBelow").unwrap_or(true),
                    summary_right: attr_bool(&node, "summaryRight").unwrap_or(true),
                });
            }
            "tabColor" if is_x_ns(node.tag_name().namespace()) => {
                tab_color = parse_color(&node, theme_colors);
            }
            "autoFilter" if is_x_ns(node.tag_name().namespace()) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    auto_filter = if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        Some(CellRange {
                            top,
                            left,
                            bottom,
                            right,
                        })
                    } else {
                        let (col, row) = parse_cell_ref(parts[0]);
                        Some(CellRange {
                            top: row,
                            left: col,
                            bottom: row,
                            right: col,
                        })
                    };
                }
            }
            "hyperlinks" if is_x_ns(node.tag_name().namespace()) => {
                for hl in node.children() {
                    if !hl.is_element() || hl.tag_name().name() != "hyperlink" {
                        continue;
                    }
                    let Some(ref_str) = hl.attribute("ref") else {
                        continue;
                    };
                    // Only first cell of ref range
                    let ref_single = ref_str.split(':').next().unwrap_or(ref_str);
                    let (col, row) = parse_cell_ref(ref_single);
                    // §18.3.1.47: `r:id` is the external target (resolved later
                    // via rels); `location` is the inline internal target
                    // (defined name or cell ref). Either may be present — or
                    // both — so capture whichever exist and skip only when both
                    // are absent (nothing to navigate to).
                    let rid = hl
                        .attributes()
                        .find(|a| a.name() == "id" && is_r_ns(a.namespace()))
                        .map(|a| a.value().to_string());
                    let location = hl.attribute("location").map(|s| s.to_string());
                    let display = hl.attribute("display").map(|s| s.to_string());
                    if rid.is_some() || location.is_some() {
                        hyperlink_rids.push((col, row, rid, location, display));
                    }
                }
            }
            "pane" if is_x_ns(node.tag_name().namespace()) => {
                let state = node.attribute("state").unwrap_or("");
                if state == "frozen" || state == "frozenSplit" {
                    freeze_rows = node
                        .attribute("ySplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| (v as u32).min(1_048_576))
                        .unwrap_or(0);
                    freeze_cols = node
                        .attribute("xSplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| (v as u32).min(16_384))
                        .unwrap_or(0);
                }
            }
            "mergeCell" if is_x_ns(node.tag_name().namespace()) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        merge_cells.push(MergeCell {
                            top,
                            left,
                            bottom,
                            right,
                        });
                    }
                }
            }
            "conditionalFormatting" if is_x_ns(node.tag_name().namespace()) => {
                let sqref = node.attribute("sqref").map(parse_sqref).unwrap_or_default();
                let mut rules: Vec<CfRule> = Vec::new();
                for cf in node.children() {
                    if cf.tag_name().name() != "cfRule" {
                        continue;
                    }
                    let kind = cf.attribute("type").unwrap_or("").to_string();
                    let priority: i32 = cf
                        .attribute("priority")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(0);
                    let dxf_id: Option<u32> = cf.attribute("dxfId").and_then(|s| s.parse().ok());
                    match kind.as_str() {
                        "cellIs" => {
                            let operator = cf.attribute("operator").unwrap_or("equal").to_string();
                            let formulas: Vec<String> = cf
                                .children()
                                .filter(|n| n.tag_name().name() == "formula")
                                .filter_map(|n| n.text().map(|s| s.to_string()))
                                .collect();
                            rules.push(CfRule::CellIs {
                                operator,
                                formulas,
                                dxf_id,
                                priority,
                            });
                        }
                        "expression" | "containsBlanks" | "notContainsBlanks" | "containsText"
                        | "notContainsText" | "beginsWith" | "endsWith" | "containsErrors"
                        | "notContainsErrors" => {
                            // For `containsBlanks`/`notContainsBlanks`/`containsText` etc.,
                            // Excel serializes an equivalent boolean formula (e.g.
                            // `LEN(TRIM(C8))>0`) as the rule's `<formula>` child
                            // (ECMA-376 §18.3.1.10). Evaluate as an expression rule.
                            let formula = cf
                                .children()
                                .find(|n| n.tag_name().name() == "formula")
                                .and_then(|n| n.text())
                                .unwrap_or("")
                                .to_string();
                            let stop_if_true = cf
                                .attribute("stopIfTrue")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            rules.push(CfRule::Expression {
                                formula,
                                dxf_id,
                                priority,
                                stop_if_true,
                            });
                        }
                        "colorScale" => {
                            let scale = cf.children().find(|n| n.tag_name().name() == "colorScale");
                            let mut stop_values: Vec<(String, Option<String>)> = Vec::new();
                            let mut stop_colors: Vec<String> = Vec::new();
                            if let Some(scale_node) = scale {
                                for child in scale_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            stop_values.push((
                                                child
                                                    .attribute("type")
                                                    .unwrap_or("num")
                                                    .to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            stop_colors.push(
                                                parse_color(&child, theme_colors)
                                                    .unwrap_or_else(|| "#FFFFFF".to_string()),
                                            );
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            let stops: Vec<CfStop> = stop_values
                                .into_iter()
                                .enumerate()
                                .map(|(i, (kind, value))| CfStop {
                                    kind,
                                    value,
                                    color: stop_colors
                                        .get(i)
                                        .cloned()
                                        .unwrap_or_else(|| "#FFFFFF".to_string()),
                                })
                                .collect();
                            rules.push(CfRule::ColorScale { stops, priority });
                        }
                        "dataBar" => {
                            let bar = cf.children().find(|n| n.tag_name().name() == "dataBar");
                            let mut cfvos: Vec<(String, Option<String>)> = Vec::new();
                            let mut color = "#638EC6".to_string();
                            if let Some(bar_node) = bar {
                                for child in bar_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            cfvos.push((
                                                child
                                                    .attribute("type")
                                                    .unwrap_or("min")
                                                    .to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            if let Some(c) = parse_color(&child, theme_colors) {
                                                color = c;
                                            }
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            // Excel 2010+ x14:dataBar extension may override the
                            // gradient flag (§2.6.3, default="1"). "0" → solid
                            // fill. The override lives in a separate
                            // worksheet-level extLst and is linked via the
                            // `<x14:id>{GUID}</x14:id>` contained in this
                            // cfRule's own extLst.
                            let mut gradient = true;
                            'gradient_lookup: for ext_list in
                                cf.children().filter(|n| n.tag_name().name() == "extLst")
                            {
                                for ext in
                                    ext_list.children().filter(|n| n.tag_name().name() == "ext")
                                {
                                    for id_node in
                                        ext.descendants().filter(|n| n.tag_name().name() == "id")
                                    {
                                        if let Some(guid) = id_node.text() {
                                            if let Some(&g) = x14_databar_gradient.get(guid) {
                                                gradient = g;
                                                break 'gradient_lookup;
                                            }
                                        }
                                    }
                                    // Fallback: some files embed <x14:dataBar>
                                    // directly in the cfRule's extLst.
                                    for x14_bar in ext
                                        .descendants()
                                        .filter(|n| n.tag_name().name() == "dataBar")
                                    {
                                        if let Some(g) = x14_bar.attribute("gradient") {
                                            gradient = !(g == "0" || g == "false");
                                            break 'gradient_lookup;
                                        }
                                    }
                                }
                            }
                            let min = cfvos
                                .first()
                                .map(|(k, v)| CfValue {
                                    kind: k.clone(),
                                    value: v.clone(),
                                })
                                .unwrap_or(CfValue {
                                    kind: "min".into(),
                                    value: None,
                                });
                            let max = cfvos
                                .get(1)
                                .map(|(k, v)| CfValue {
                                    kind: k.clone(),
                                    value: v.clone(),
                                })
                                .unwrap_or(CfValue {
                                    kind: "max".into(),
                                    value: None,
                                });
                            rules.push(CfRule::DataBar {
                                color,
                                min,
                                max,
                                priority,
                                gradient,
                            });
                        }
                        "top10" => {
                            let top = !cf
                                .attribute("bottom")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let percent = cf
                                .attribute("percent")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let rank = cf
                                .attribute("rank")
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(10);
                            rules.push(CfRule::Top10 {
                                top,
                                percent,
                                rank,
                                dxf_id,
                                priority,
                            });
                        }
                        "aboveAverage" => {
                            let above_average = cf
                                .attribute("aboveAverage")
                                .map(|v| v != "0")
                                .unwrap_or(true);
                            // ECMA-376 §18.3.1.10: `equalAverage` (default
                            // false) and `stdDev` (optional, number of
                            // standard deviations for the band threshold).
                            let equal_average = cf
                                .attribute("equalAverage")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let std_dev = cf
                                .attribute("stdDev")
                                .and_then(|v| v.parse::<u32>().ok())
                                .filter(|&n| n > 0);
                            rules.push(CfRule::AboveAverage {
                                above_average,
                                equal_average,
                                std_dev,
                                dxf_id,
                                priority,
                            });
                        }
                        "iconSet" => {
                            let icon_set_node =
                                cf.children().find(|n| n.tag_name().name() == "iconSet");
                            let icon_set = icon_set_node
                                .and_then(|n| n.attribute("iconSet"))
                                .unwrap_or("3TrafficLights1")
                                .to_string();
                            let reverse = icon_set_node
                                .and_then(|n| n.attribute("reverse"))
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let cfvos: Vec<CfValue> = icon_set_node
                                .map(|n| {
                                    n.children()
                                        .filter(|c| c.is_element() && c.tag_name().name() == "cfvo")
                                        .map(|c| CfValue {
                                            kind: c
                                                .attribute("type")
                                                .unwrap_or("percent")
                                                .to_string(),
                                            value: c.attribute("val").map(|s| s.to_string()),
                                        })
                                        .collect()
                                })
                                .unwrap_or_default();
                            rules.push(CfRule::IconSet {
                                icon_set,
                                cfvos,
                                reverse,
                                priority,
                                custom_icons: None,
                            });
                        }
                        other => {
                            rules.push(CfRule::Other {
                                kind: other.to_string(),
                                priority,
                            });
                        }
                    }
                }
                conditional_formats.push(ConditionalFormat { sqref, rules });
            }
            _ => {}
        }
    }

    // Resolve the effective width of every sheet column once, from the last
    // declaration backwards. The successor DSU skips columns already claimed
    // by a later declaration, bounding work by declarations + 16,384 columns
    // rather than declarations × range length.
    let mut resolved_widths: Vec<Option<f64>> = vec![None; 16_385];
    let mut successor: Vec<usize> = (0..=16_385).collect();
    fn find_unassigned(successor: &mut [usize], start: usize) -> usize {
        let mut root = start;
        while successor[root] != root {
            root = successor[root];
        }
        let mut index = start;
        while successor[index] != index {
            let next = successor[index];
            successor[index] = root;
            index = next;
        }
        root
    }
    for (range, _) in authored_col_widths.iter().rev() {
        let mut index = find_unassigned(&mut successor, range.min as usize);
        while index <= range.max as usize {
            resolved_widths[index] = Some(range.width);
            successor[index] = find_unassigned(&mut successor, index + 1);
            index = successor[index];
        }
    }
    // Keep the legacy point map for existing consumers, but normalize its
    // entries to the final XML document-order result. Live edits still mutate
    // these points and therefore override compact ranges in grid geometry.
    for (index, width) in &mut col_widths {
        if let Some(resolved) = resolved_widths[*index as usize] {
            *width = resolved;
        }
    }
    let col_width_ranges = authored_col_widths
        .into_iter()
        .filter_map(|(range, compact)| compact.then_some(range))
        .collect();

    // ECMA-376 §18.3.1.13: a column style is the default XF for cells in that
    // column. Resolve it onto sparse authored cells only after every `<col>`
    // declaration is known; an explicit `c/@s="0"` remains `Some(0)` and wins.
    for row in &mut rows {
        for cell in &mut row.cells {
            if cell.style_index.is_some() {
                continue;
            }
            cell.style_index = authored_col_styles
                .iter()
                .rev()
                .find(|range| cell.col >= range.min && cell.col <= range.max)
                .map(|range| range.style_index);
        }
    }

    conditional_formats.extend(x14_icon_formats);

    if rows_hidden_by_default {
        let visible_default_height = default_row_height;
        for row in &mut rows {
            if !row.hidden && row.height.is_none() {
                row.height = Some(visible_default_height);
                row_heights.insert(row.index, visible_default_height);
            }
        }
        // Unspecified rows remain hidden. The sparse grid axis represents the
        // hidden default as zero while the explicit rows above are positive
        // custom bands, so it can jump over arbitrarily large hidden spans.
        default_row_height = 0.0;
    }

    let data_validations = parse_data_validations(doc.root_element());

    let worksheet = Worksheet {
        name: name.to_string(),
        is_chart_sheet: false,
        is_dialog_sheet: false,
        rows,
        col_widths,
        col_width_ranges,
        col_style_ranges: authored_col_styles,
        row_heights,
        col_outline_levels,
        col_collapsed,
        col_hidden,
        default_col_width,
        default_row_height,
        default_row_height_custom,
        merge_cells,
        freeze_rows,
        freeze_cols,
        conditional_formats,
        images: Vec::new(),
        charts: Vec::new(),
        shape_groups: Vec::new(),
        show_zeros,
        show_gridlines,
        right_to_left,
        outline_pr,
        tab_color,
        auto_filter,
        hyperlinks: Vec::new(),
        comment_refs: Vec::new(),
        comments: Vec::new(),
        data_validations,
        defined_names: Vec::new(),
        tables: Vec::new(),
        slicers: Vec::new(),
        pivot_tables: Vec::new(),
        pivot_diagnostics: Vec::new(),
        sparkline_groups: Vec::new(),
        default_font_family: None,
        default_font_size: None,
        // Set by `parse_sheet_with` from the workbook-level `<workbookPr
        // date1904>` (ECMA-376 §18.2.28); a bare `parse_worksheet` (tests)
        // defaults to the 1900 date system.
        date1904: false,
        // A successfully parsed sheet carries no error (RB7). Only the
        // `Worksheet::placeholder` path sets this.
        parse_error: None,
    };
    // `doc` borrows the shell. End that borrow before moving the small shell to
    // ancillary OLE/sparkline readers so they never rebuild a DOM over the
    // original, potentially huge `sheetData`.
    drop(doc);
    Ok((worksheet, hyperlink_rids, streamed.shell_xml))
}

/// Parse the terminal worksheet shell without pretending that it is a full
/// row-bearing stream product. Production cursors have already emitted and
/// dropped every row; only row-height metadata remains part of shell parsing.
fn parse_projected_worksheet_tail(
    tail: WorksheetCursorTail,
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, HyperlinkRids, String), String> {
    parse_projected_worksheet(
        StreamedSheetData {
            shell_xml: tail.shell_xml,
            rows: Vec::new(),
            row_heights: tail.row_heights,
        },
        theme_colors,
        name,
    )
}

#[cfg(test)]
fn parse_worksheet(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, HyperlinkRids), String> {
    parse_worksheet_with_shell(xml, shared_strings, theme_colors, name)
        .map(|(worksheet, hyperlinks, _shell)| (worksheet, hyperlinks))
}

/// Parse a .rels file into rId → Target map.
/// id → target map for a `.rels` part. Thin adapter over
/// [`ooxml_common::rels::parse_rels`] that flattens each `RelTarget` to its raw
/// target string (both Internal part names and External hyperlink URLs are kept
/// verbatim; part-name resolution happens later via [`resolve_zip_path`]),
/// preserving this parser's `HashMap<rId, Target>` shape.
pub(crate) fn parse_rels_map(xml: &str) -> HashMap<String, String> {
    ooxml_common::rels::parse_rels(xml)
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect()
}

/// Internal package-part target of the first relationship whose type ends in
/// `type_suffix`. Unlike hyperlink-oriented callers, workbook-owned resources
/// such as the theme must never treat an external URI as a ZIP part name.
fn find_internal_rel_target_by_type(rels_xml: &str, type_suffix: &str) -> Option<String> {
    let doc = parse_guarded(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if rel
            .attribute("Type")
            .is_some_and(|rel_type| rel_type.ends_with(type_suffix))
            && !rel
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
        {
            return rel.attribute("Target").map(str::to_owned);
        }
    }
    None
}

/// Parse xl/comments{N}.xml referenced from the sheet's rels and collect the
/// list of A1-style cell refs that have a `<comment>` associated. The
/// renderer draws a small red triangle in each cell's top-right corner to
/// indicate the presence of a comment (ECMA-376 §18.7.3 commentList).
/// Reads xl/commentsN.xml for the given sheet and returns each `<comment>` as
/// a structured `XlsxComment` (cell ref, resolved author name, plain text).
/// Callers can derive `comment_refs: Vec<String>` from `c.cell_ref`.
fn load_sheet_comments(
    archive: &mut XlsxZip,
    sheet_path: &str,
    workbook_rels_xml: &str,
) -> Vec<XlsxComment> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = parse_guarded(&rels_xml) else {
        return Vec::new();
    };

    // A sheet may carry classic notes (`/comments`, ECMA-376 §18.7) and/or
    // Office-365 threaded comments (`/threadedComment`, MS extension). Excel
    // also writes legacy placeholders in `xl/commentsN.xml` for threaded roots.
    // Parse both parts, then apply the explicit MS-XLSX §2.3.7.3 reconciliation
    // identity so unrelated classic notes remain visible.
    let mut classic_target: Option<String> = None;
    let mut threaded_target: Option<String> = None;
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        let rel_type = rel.attribute("Type").unwrap_or("");
        let Some(t) = rel.attribute("Target") else {
            continue;
        };
        if rel_type.ends_with("/comments") {
            classic_target.get_or_insert_with(|| t.to_string());
        } else if rel_type.ends_with("/threadedComment") {
            threaded_target.get_or_insert_with(|| t.to_string());
        }
    }

    let threaded_comments = if let Some(target) = threaded_target {
        let tc_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let comments = if let Ok(tc_xml) = read_zip_string(archive, &tc_path) {
            let persons = load_persons(archive, workbook_rels_xml);
            parse_threaded_comments_xml(&tc_xml, &persons)
        } else {
            Vec::new()
        };
        Some(comments)
    } else {
        None
    };

    let mut classic_comments = Vec::new();
    if let Some(target) = classic_target {
        let comments_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        if let Ok(comments_xml) = read_zip_string(archive, &comments_path) {
            classic_comments = parse_comments_xml(&comments_xml);
        }
    }

    merge_sheet_comments(threaded_comments, classic_comments)
}

/// Reconcile modern threads with the legacy placeholders required by MS-XLSX
/// §2.3.7.3/.3.1, while retaining unrelated ECMA-376 §18.7 notes. A recognized
/// placeholder is metadata for a thread, not a second user-visible note. `None`
/// means the worksheet has no threaded-comment relationship; `Some(empty)`
/// preserves the fact that a declared part yielded no valid thread records.
fn merge_sheet_comments(
    threaded_part: Option<Vec<XlsxComment>>,
    classic: Vec<XlsxComment>,
) -> Vec<XlsxComment> {
    let Some(threaded) = threaded_part else {
        return classic;
    };

    fn guid_key(value: &str) -> Option<String> {
        let value = value.trim().trim_start_matches('{').trim_end_matches('}');
        if value.len() != 36
            || !value.bytes().enumerate().all(|(index, byte)| {
                if matches!(index, 8 | 13 | 18 | 23) {
                    byte == b'-'
                } else {
                    byte.is_ascii_hexdigit()
                }
            })
        {
            return None;
        }
        Some(value.to_ascii_lowercase())
    }

    fn placeholder_key(comment: &XlsxComment) -> Option<String> {
        let uid = guid_key(comment.id.as_deref()?)?;
        let author = comment.author.as_deref()?;
        let marker = author.find("tc=")?;
        let candidate: String = author[marker + 3..]
            .chars()
            .take_while(|character| {
                character.is_ascii_hexdigit() || matches!(character, '-' | '{' | '}')
            })
            .collect();
        (guid_key(&candidate).as_deref() == Some(uid.as_str())).then_some(uid)
    }

    let thread_ids: HashSet<String> = threaded
        .iter()
        .filter_map(|comment| comment.id.as_deref().and_then(guid_key))
        .collect();
    let mut placeholders: HashMap<String, Vec<XlsxComment>> = HashMap::new();
    let mut notes = Vec::new();
    for comment in classic {
        let Some(key) = placeholder_key(&comment) else {
            notes.push(comment);
            continue;
        };
        // MS-XLSX §2.3.7.3.1 removes an orphan placeholder instead of showing
        // it as a note. A matching placeholder supplies the displayed ref.
        if thread_ids.contains(&key) {
            placeholders.entry(key).or_default().push(comment);
        }
    }

    let mut merged = Vec::with_capacity(threaded.len() + notes.len());
    for comment in threaded {
        let key = comment.id.as_deref().and_then(guid_key);
        let matching = key.as_ref().and_then(|id| placeholders.remove(id));
        match matching {
            Some(occurrences) => {
                // Reconciliation may copy one logical thread to multiple
                // placeholder locations. Preserve the source id while giving
                // each read-only occurrence its authored cell reference.
                for placeholder in occurrences {
                    let mut occurrence = comment.clone();
                    occurrence.cell_ref = placeholder.cell_ref;
                    merged.push(occurrence);
                }
            }
            None => merged.push(comment),
        }
    }
    merged.extend(notes);
    merged
}

/// Load the `personId` → display-name map from the Persons part targeted by the
/// workbook's implicit relationship ([MS-XLSX] §2.1.18, relationship Type
/// `http://schemas.microsoft.com/office/2017/10/relationships/person`). The
/// package part name is relationship-owned; `xl/persons/` is only a convention
/// and unreferenced entries there must not be observed. `<person displayName id/>`.
/// Returns an empty map when no internal Persons relationship or readable part exists.
fn load_persons(archive: &mut XlsxZip, workbook_rels_xml: &str) -> HashMap<String, String> {
    const PERSON_RELATIONSHIP_TYPE: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/person";
    let mut out: HashMap<String, String> = HashMap::new();
    let Some(target) =
        find_internal_rel_target_by_type(workbook_rels_xml, PERSON_RELATIONSHIP_TYPE)
    else {
        return out;
    };
    let path = resolve_zip_path("xl", &target);
    let Ok(xml) = read_zip_string(archive, &path) else {
        return out;
    };
    let Ok(doc) = parse_guarded(&xml) else {
        return out;
    };
    for p in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "person")
    {
        if let (Some(id), Some(name)) = (p.attribute("id"), p.attribute("displayName")) {
            out.insert(id.to_string(), name.to_string());
        }
    }
    out
}

/// Parse MS-XLSX CT_ThreadedComment records without flattening their reply
/// identities. Parentage follows `@parentId` only; malformed orphan/cyclic
/// records are not guessed into a thread.
fn parse_threaded_comments_xml(
    tc_xml: &str,
    persons: &HashMap<String, String>,
) -> Vec<XlsxComment> {
    let Ok(doc) = parse_guarded(tc_xml) else {
        return Vec::new();
    };
    struct RawThreadedComment {
        cell_ref: Option<String>,
        id: String,
        parent_id: Option<String>,
        person_id: String,
        author: Option<String>,
        date: Option<String>,
        text: String,
        resolved: Option<bool>,
    }

    fn xml_bool(value: Option<&str>) -> Option<bool> {
        match value?.trim().to_ascii_lowercase().as_str() {
            "1" | "true" => Some(true),
            "0" | "false" => Some(false),
            _ => None,
        }
    }

    let mut records: Vec<RawThreadedComment> = Vec::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "threadedComment")
    {
        let Some(id) = node.attribute("id") else {
            continue;
        };
        let Some(person_id) = node.attribute("personId") else {
            continue;
        };
        let author = persons.get(person_id).cloned().filter(|s| !s.is_empty());
        let text = node
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "text")
            .and_then(|t| t.text())
            .unwrap_or("")
            .to_string();
        records.push(RawThreadedComment {
            cell_ref: node.attribute("ref").map(str::to_owned),
            id: id.to_string(),
            parent_id: node.attribute("parentId").map(str::to_owned),
            person_id: person_id.to_string(),
            author,
            date: node.attribute("dT").map(str::to_owned),
            text,
            resolved: xml_bool(node.attribute("done")),
        });
    }

    // `@id` is the only authored identity. A duplicate is ambiguous, so do not
    // attach replies to whichever duplicate happened to be visited last.
    let mut by_id: HashMap<&str, usize> = HashMap::new();
    let mut duplicate_ids: HashSet<&str> = HashSet::new();
    for (index, record) in records.iter().enumerate() {
        if by_id.insert(record.id.as_str(), index).is_some() {
            duplicate_ids.insert(record.id.as_str());
        }
    }
    for id in &duplicate_ids {
        by_id.remove(id);
    }

    #[derive(Clone, Copy)]
    enum RootResolution {
        Unknown,
        Resolving,
        Resolved(Option<usize>),
    }

    fn resolve_root(
        start: usize,
        records: &[RawThreadedComment],
        by_id: &HashMap<&str, usize>,
        states: &mut [RootResolution],
    ) -> Option<usize> {
        let mut current = start;
        let mut path = Vec::new();
        let root = loop {
            match states[current] {
                RootResolution::Resolved(root) => break root,
                RootResolution::Resolving => break None,
                RootResolution::Unknown => {}
            }
            states[current] = RootResolution::Resolving;
            path.push(current);
            let Some(parent_id) = records[current].parent_id.as_deref() else {
                break Some(current);
            };
            let Some(parent) = by_id.get(parent_id).copied() else {
                break None;
            };
            current = parent;
        };
        for index in path {
            states[index] = RootResolution::Resolved(root);
        }
        root
    }

    let mut states = vec![RootResolution::Unknown; records.len()];
    let mut replies_by_root: HashMap<usize, Vec<XlsxCommentReply>> = HashMap::new();
    for (index, reply) in records.iter().enumerate() {
        if reply.parent_id.is_none() || duplicate_ids.contains(reply.id.as_str()) {
            continue;
        }
        let Some(root_index) = resolve_root(index, &records, &by_id, &mut states) else {
            continue;
        };
        let Some(parent_id) = reply.parent_id.clone() else {
            continue;
        };
        replies_by_root
            .entry(root_index)
            .or_default()
            .push(XlsxCommentReply {
                id: reply.id.clone(),
                parent_id,
                person_id: reply.person_id.clone(),
                author: reply.author.clone(),
                date: reply.date.clone(),
                text: reply.text.clone(),
                resolved: reply.resolved,
            });
    }

    records
        .iter()
        .enumerate()
        .filter(|(_, record)| {
            record.parent_id.is_none()
                && record.cell_ref.is_some()
                && !duplicate_ids.contains(record.id.as_str())
        })
        .map(|(root_index, root)| XlsxComment {
            root_text: Some(root.text.clone()),
            kind: XlsxCommentKind::Thread,
            cell_ref: root.cell_ref.clone().unwrap_or_default(),
            id: Some(root.id.clone()),
            person_id: Some(root.person_id.clone()),
            author: root.author.clone(),
            date: root.date.clone(),
            text: std::iter::once(root.text.as_str())
                .chain(
                    replies_by_root
                        .get(&root_index)
                        .into_iter()
                        .flatten()
                        .map(|reply| reply.text.as_str()),
                )
                .collect::<Vec<_>>()
                .join("\n"),
            resolved: root.resolved,
            replies: replies_by_root.remove(&root_index).unwrap_or_default(),
        })
        .collect()
}

/// Parse a `xl/commentsN.xml` document (ECMA-376 §18.7) into structured
/// `XlsxComment`s. Resolves `@authorId` against the `<authors>` block and joins
/// every `<text>/<r>/<t>` run into plain text (rich-text formatting dropped).
/// Returns an empty vec on malformed XML. Split out from `load_sheet_comments`
/// so the parse path is unit-testable without a ZIP archive.
fn parse_comments_xml(comments_xml: &str) -> Vec<XlsxComment> {
    let Ok(comments_doc) = parse_guarded(comments_xml) else {
        return Vec::new();
    };

    // Resolve <authors><author>…</author></authors> — `authorId` indexes here.
    let authors: Vec<String> = comments_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "authors")
        .map(|n| {
            n.children()
                .filter(|c| c.is_element() && c.tag_name().name() == "author")
                .map(|c| c.text().unwrap_or("").to_string())
                .collect()
        })
        .unwrap_or_default();

    let mut comments: Vec<XlsxComment> = Vec::new();
    for node in comments_doc.descendants() {
        if node.tag_name().name() != "comment" || !node.is_element() {
            continue;
        }
        let Some(cell_ref) = node.attribute("ref") else {
            continue;
        };
        let author = node
            .attribute("authorId")
            .and_then(|s| s.parse::<usize>().ok())
            .and_then(|i| authors.get(i).cloned())
            .filter(|s| !s.is_empty());
        let mut text = String::new();
        if let Some(t_node) = node
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "text")
        {
            for r in t_node.descendants() {
                if r.is_element() && r.tag_name().name() == "t" {
                    if let Some(s) = r.text() {
                        text.push_str(s);
                    }
                }
            }
        }
        comments.push(XlsxComment {
            root_text: None,
            kind: XlsxCommentKind::Note,
            cell_ref: cell_ref.to_string(),
            id: node
                .attributes()
                .find(|attribute| attribute.name() == "uid")
                .map(|attribute| attribute.value().to_string()),
            person_id: None,
            author,
            date: None,
            text,
            resolved: None,
            replies: Vec::new(),
        });
    }
    comments
}

/// ECMA-376 §18.3.1.32 — extracts `<dataValidations>` rules from the sheet
/// XML root. Returns an empty vec when the element is absent.
fn parse_data_validations(ws_root: roxmltree::Node<'_, '_>) -> Vec<DataValidation> {
    let mut out: Vec<DataValidation> = Vec::new();
    let Some(dvs) = ws_root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dataValidations")
    else {
        return out;
    };
    for dv in dvs
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dataValidation")
    {
        let sqref = dv.attribute("sqref").unwrap_or("").to_string();
        if sqref.is_empty() {
            continue;
        }
        let validation_type = dv.attribute("type").map(String::from);
        let operator = dv.attribute("operator").map(String::from);
        let allow_blank = dv
            .attribute("allowBlank")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        let prompt_title = dv
            .attribute("promptTitle")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let prompt = dv
            .attribute("prompt")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let error_title = dv
            .attribute("errorTitle")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let error_message = dv
            .attribute("error")
            .map(String::from)
            .filter(|s| !s.is_empty());

        let mut formula1: Option<String> = None;
        let mut formula2: Option<String> = None;
        for child in dv.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "formula1" => formula1 = child.text().map(String::from).filter(|s| !s.is_empty()),
                "formula2" => formula2 = child.text().map(String::from).filter(|s| !s.is_empty()),
                _ => {}
            }
        }

        out.push(DataValidation {
            sqref,
            validation_type,
            operator,
            formula1,
            formula2,
            allow_blank,
            prompt_title,
            prompt,
            error_title,
            error_message,
        });
    }
    out
}

/// Resolve hyperlink rIds to URLs from the sheet rels file.
fn load_hyperlinks(
    archive: &mut XlsxZip,
    sheet_path: &str,
    hyperlink_rids: HyperlinkRids,
) -> Vec<Hyperlink> {
    if hyperlink_rids.is_empty() {
        return Vec::new();
    }
    // Only read the sheet rels part when at least one hyperlink carries an
    // external `r:id`. A worksheet whose hyperlinks are all internal
    // (`location`-only) needs no rels lookup (§18.3.1.47).
    let needs_rels = hyperlink_rids.iter().any(|(_, _, rid, _, _)| rid.is_some());
    let rels = if needs_rels {
        match sheet_path.rsplit_once('/') {
            Some((sheet_dir, sheet_file)) => {
                let rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
                read_zip_string(archive, &rels_path)
                    .ok()
                    .map(|xml| parse_rels_map(&xml))
                    .unwrap_or_default()
            }
            None => Default::default(),
        }
    } else {
        Default::default()
    };
    hyperlink_rids
        .into_iter()
        .map(|(col, row, rid, location, display)| Hyperlink {
            col,
            row,
            url: rid.as_deref().and_then(|r| rels.get(r).cloned()),
            location,
            display,
        })
        .collect()
}

/// Resolve a relative path ("../media/image1.png") against a base dir
/// ("xl/drawings"). Thin alias for the shared
/// [`ooxml_common::rels::resolve_target`], which handles root-absolute Targets
/// (openpyxl's `/xl/...`) and `..` normalization uniformly (ECMA-376 Part 2
/// §9.3). Kept as a local name so existing call sites read unchanged.
pub(crate) fn resolve_zip_path(base_dir: &str, target: &str) -> String {
    ooxml_common::rels::resolve_target(base_dir, target)
}

pub(crate) fn resolve_fill_color(
    fill_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    // Accept either a `<a:solidFill>` directly or a `<c:spPr>` whose first
    // fill-ish child is `<a:solidFill>`. Looking at *direct* children (not
    // descendants) is intentional — chart series often carry label/axis text
    // colors under `c:dLbls`/`c:txPr` which must NOT be misread as fill.
    let solid = if fill_node.tag_name().name() == "solidFill" {
        Some(*fill_node)
    } else {
        fill_node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
    }?;
    for n in solid.children().filter(|n| n.is_element()) {
        let tag = n.tag_name().name();
        if tag == "srgbClr" {
            if let Some(v) = n.attribute("val") {
                return Some(v.to_lowercase());
            }
        }
        if tag == "schemeClr" {
            if let Some(v) = n.attribute("val") {
                // ooxml_common::color::SCHEME_DEFAULT_SLOTS is the canonical
                // logical→slot-NAME table (§19.3.1.6). xlsx instead maps the
                // logical/slot name straight to a numeric INDEX into the theme
                // color Vec (raw clrScheme order: dk1=0, lt1=1, dk2=2, lt2=3,
                // accent1..6=4..9, hlink=10, folHlink=11). Routing through the
                // shared name→name table would add an indirection without
                // changing the result, so this stays local. (Note: the cell
                // @theme path below applies the §22.1.2.7 dk1↔lt1 / dk2↔lt2
                // index swap; this drawing path indexes the array directly.)
                let idx = match v {
                    "dk1" | "tx1" => Some(0),
                    "lt1" | "bg1" => Some(1),
                    "dk2" | "tx2" => Some(2),
                    "lt2" | "bg2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    "hlink" => Some(10),
                    "folHlink" => Some(11),
                    _ => None,
                };
                if let Some(i) = idx {
                    if let Some(c) = theme_colors.get(i) {
                        return Some(c.trim_start_matches('#').to_lowercase());
                    }
                }
            }
        }
    }
    None
}

/// Build a [`CellRange`] from two corner cells, normalizing so `top <= bottom`
/// and `left <= right` regardless of which corner was written first/second in
/// the source reference.
///
/// ECMA-376 does not spell out corner order for `ST_Ref`/`ST_Sqref` (§18.18.62 /
/// §18.18.76 describe the reference grammar, not a canonicalization rule), but
/// Excel itself treats `A10:A1` and `A1:A10` as the identical range — the UI
/// re-displays a reversed typed range in top-left/bottom-right order. Callers
/// throughout this module (`extract_range_values`, dimension math) assume
/// `bottom >= top` and `right >= left` and compute spans via unsigned
/// subtraction; an un-normalized reversed range (e.g. a crafted `<xm:f>` of
/// `A10:A1`) underflows that subtraction — silently wrapping in a release
/// build (`overflow-checks` off) and panicking in debug/test builds. Defensive
/// normalization here closes that off for every `parse_sqref` consumer at the
/// source, matching Excel's actual interpretation rather than merely avoiding
/// the crash.
fn cell_range_from_corners(a: &str, b: &str) -> CellRange {
    let (col_a, row_a) = parse_cell_ref(a);
    let (col_b, row_b) = parse_cell_ref(b);
    CellRange {
        top: row_a.min(row_b),
        bottom: row_a.max(row_b),
        left: col_a.min(col_b),
        right: col_a.max(col_b),
    }
}

fn parse_sqref(s: &str) -> Vec<CellRange> {
    s.split_whitespace()
        .map(|range_str| {
            if let Some((a, b)) = range_str.split_once(':') {
                cell_range_from_corners(a, b)
            } else {
                let (col, row) = parse_cell_ref(range_str);
                CellRange {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                }
            }
        })
        .collect()
}

/// Read a worksheet XML and extract numeric `<v>` values for the cells in
/// `range`. Returns one value per cell in row-major order across the range.
/// Empty cells, non-numeric values, and cells outside the range yield `None`.
///
/// This is intentionally lighter than `parse_row_cells`: sparklines only need
/// raw numbers, no styles, formulas, or shared strings.
/// Upper bound on the number of cells a single sparkline data range may span
/// before `extract_range_values` refuses to materialize the dense value buffer.
///
/// Rationale: a real sparkline plots a handful to a few hundred points; even a
/// generous whole-column series is 1,048,576 cells (Excel's max rows, ECMA-376
/// §18.3.1.73 — `SpreadsheetML` grid is 16384 cols × 1048576 rows). We cap at
/// exactly one million cells, which:
///   • comfortably covers any legitimate range (a full single column, or a
///     1000×1000 block), and
///   • bounds the dense `Vec<Option<f64>>` to 1e6 × 16 B = 16 MiB — trivially
///     within the 512 MiB per-entry ZIP budget (`ooxml_common::zip`), so the
///     sparkline allocation can never dominate the parse.
/// A crafted `A1:XFD1048576` reference (16384 × 1048576 ≈ 1.7e10 cells ≈ 275 GB)
/// exceeds this by four orders of magnitude and is refused: we return an empty
/// `Vec` and the sparkline is simply not drawn (the renderer iterates the value
/// slice by index, so an empty slice draws nothing — graceful degradation, not
/// a hard error).
#[cfg(test)]
const MAX_SPARKLINE_CELLS: usize = MAX_REFERENCE_CELLS;

#[cfg(test)]
fn extract_range_values(sheet_xml: &str, range: &CellRange) -> Vec<Option<f64>> {
    extract_reference_cells(sheet_xml, range, &[])
        .into_iter()
        .map(|value| match value {
            ReferencedCellValue::Number(number) => Some(number),
            _ => None,
        })
        .collect()
}

/// Walk the worksheet XML's `<extLst>` and produce one `SparklineGroup` per
/// `<x14:sparklineGroup>`. Resolves cross-sheet `<xm:f>` data references by
/// reading the referenced sheet from the archive (cached per call to avoid
/// re-reads). Theme colors are flattened to `#RRGGBB` via `parse_color`.
#[allow(clippy::too_many_arguments)]
fn load_sheet_sparklines(
    archive: &mut XlsxZip,
    sheet_xml: &str,
    materialized_rows: Option<&[Row]>,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    rels_doc: &roxmltree::Document,
    theme_colors: &[String],
    shared_strings: &[SharedString],
    reference_session: &mut WorksheetReferenceSession,
) -> Vec<SparklineGroup> {
    let Ok(doc) = parse_guarded(sheet_xml) else {
        return Vec::new();
    };
    let mut groups: Vec<SparklineGroup> = Vec::new();
    let parse_bool_attr = |n: &roxmltree::Node, key: &str, default: bool| -> bool {
        match n.attribute(key) {
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
            None => default,
        }
    };
    let parse_f64_attr = |n: &roxmltree::Node, key: &str| -> Option<f64> {
        n.attribute(key).and_then(|v| v.parse::<f64>().ok())
    };

    for group_node in doc
        .descendants()
        .filter(|n| n.tag_name().name() == "sparklineGroup")
    {
        let kind = match group_node.attribute("type").unwrap_or("line") {
            "column" => SparklineType::Column,
            "stacked" => SparklineType::Stem, // historical alias
            "stem" => SparklineType::Stem,
            // ECMA-376 lists `line` and a planned `stairStep`; treat unknown
            // types as line (closest visual fallback).
            _ => SparklineType::Line,
        };

        let resolve_color = |child_name: &str| -> Option<String> {
            group_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == child_name)
                .and_then(|n| parse_color(&n, theme_colors))
        };

        let mut sparklines: Vec<Sparkline> = Vec::new();
        // <x14:sparklines> is the wrapper; <x14:sparkline> are the children.
        for sparklines_node in group_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sparklines")
        {
            for sl in sparklines_node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "sparkline")
            {
                let f_text = sl
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "f")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                let sqref_text = sl
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "sqref")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                if f_text.is_empty() || sqref_text.is_empty() {
                    continue;
                }
                let (col, row) = parse_cell_ref(sqref_text.trim());
                let values = resolve_worksheet_reference(
                    archive,
                    f_text,
                    materialized_rows,
                    current_sheet_name,
                    sheets,
                    rels_doc,
                    shared_strings,
                    reference_session,
                )
                .unwrap_or_default()
                .into_iter()
                .map(|value| match value {
                    ReferencedCellValue::Number(number) => Some(number),
                    _ => None,
                })
                .collect();

                sparklines.push(Sparkline { row, col, values });
            }
        }

        groups.push(SparklineGroup {
            kind,
            markers: parse_bool_attr(&group_node, "markers", false),
            high: parse_bool_attr(&group_node, "high", false),
            low: parse_bool_attr(&group_node, "low", false),
            first: parse_bool_attr(&group_node, "first", false),
            last: parse_bool_attr(&group_node, "last", false),
            negative: parse_bool_attr(&group_node, "negative", false),
            display_x_axis: parse_bool_attr(&group_node, "displayXAxis", false),
            display_empty_cells_as: group_node
                .attribute("displayEmptyCellsAs")
                .unwrap_or("gap")
                .to_string(),
            min_axis_type: group_node
                .attribute("minAxisType")
                .unwrap_or("individual")
                .to_string(),
            max_axis_type: group_node
                .attribute("maxAxisType")
                .unwrap_or("individual")
                .to_string(),
            manual_min: parse_f64_attr(&group_node, "manualMin"),
            manual_max: parse_f64_attr(&group_node, "manualMax"),
            line_weight: parse_f64_attr(&group_node, "lineWeight").unwrap_or(0.75),
            color_series: resolve_color("colorSeries"),
            color_negative: resolve_color("colorNegative"),
            color_axis: resolve_color("colorAxis"),
            color_markers: resolve_color("colorMarkers"),
            color_first: resolve_color("colorFirst"),
            color_last: resolve_color("colorLast"),
            color_high: resolve_color("colorHigh"),
            color_low: resolve_color("colorLow"),
            sparklines,
        });
    }
    groups
}

fn parse_row_cells(
    row_node: &roxmltree::Node,
    // The resolved 1-based row number of the containing `<row>`. Used as the row
    // coordinate for any `<c>` that omits its own `@r` (optional per ECMA-376
    // §18.3.1.4 / CT_Cell in sml.xsd).
    row_index: u32,
    // ECMA-376 §18.3.1.73 `<row ph>` — the row-level furigana display toggle.
    // A cell inherits this when it does not carry its own `<c ph>` (see the
    // `show_phonetic` resolution below).
    row_ph: bool,
    // Shared-string cells now ship an `si` reference (resolved consumer-side),
    // so this table is no longer read here. Kept in the signature for symmetry
    // with `parse_worksheet`'s threading; prefixed `_` to silence the warning.
    _shared_strings: &[SharedString],
    theme_colors: &[String],
) -> Result<Vec<Cell>, String> {
    let mut cells = Vec::new();
    // ECMA-376 §18.3.1.4 (CT_Cell, sml.xsd) makes `@r` on `<c>` optional with no
    // default; the spec does not spell out how an omitted value is resolved. We
    // follow the de-facto consumer convention (Excel, LibreOffice, SheetJS
    // agree; no competing interpretation exists): an r-less cell takes the
    // column after the previous cell in this row (the first cell starts at
    // column A / 1), and an explicit `@r` re-anchors this running column so
    // subsequent omitted cells continue from it. `prev_col == 0` means "no cell
    // yet", so the first implicit cell lands at column 1.
    let mut prev_col: u32 = 0;
    for c_node in row_node.children() {
        if c_node.tag_name().name() != "c" || !is_x_ns(c_node.tag_name().namespace()) {
            continue;
        }
        // An explicit `@r` re-anchors the running column; an omitted one takes
        // the previous cell's column + 1 and inherits the row's resolved index.
        // Both cases update `prev_col` via the shared primitive so this and the
        // sparkline data path (`extract_range_values`) cannot drift.
        let (col, row) = match c_node.attribute("r") {
            Some(cell_ref) => {
                let (raw_col, raw_row) = parse_cell_ref_checked(cell_ref)?;
                let col = resolve_implicit_ordinal(
                    Some(raw_col),
                    &mut prev_col,
                    SpreadsheetOrdinal::Column,
                )?;
                let mut row_anchor = 0;
                let row = resolve_implicit_ordinal(
                    Some(raw_row),
                    &mut row_anchor,
                    SpreadsheetOrdinal::Row,
                )?;
                (col, row)
            }
            None => (
                resolve_implicit_ordinal(None, &mut prev_col, SpreadsheetOrdinal::Column)?,
                row_index,
            ),
        };
        let cell_type = c_node.attribute("t").unwrap_or("");
        let style_index: Option<u32> = c_node.attribute("s").and_then(|s| s.parse().ok());

        // Inline string: <c t="inlineStr"><is>...</is></c>
        let is_node = c_node.children().find(|n| n.tag_name().name() == "is");

        // Formula text, if any (<f>…</f>). Kept so the renderer can
        // recompute volatile builtins (TODAY, NOW) at display time.
        let formula: Option<String> = c_node
            .children()
            .find(|n| n.tag_name().name() == "f")
            .and_then(|n| n.text())
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty());

        let v_text = c_node
            .children()
            .find(|n| n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .to_string();

        let value = if cell_type == "inlineStr" {
            match is_node {
                Some(is) => {
                    let ss = parse_si_node(&is, theme_colors);
                    CellValue::Text {
                        text: ss.text,
                        runs: ss.runs,
                        phonetic_runs: ss.phonetic_runs,
                        phonetic_pr: ss.phonetic_pr,
                    }
                }
                None => CellValue::Empty,
            }
        } else if v_text.is_empty() {
            CellValue::Empty
        } else {
            match cell_type {
                "s" => {
                    // Ship only the shared-string index; the consumer resolves
                    // it against the workbook `sharedStrings` table (once per
                    // workbook, not cloned per cell). Emit `Shared`
                    // unconditionally — an out-of-range index resolves to empty
                    // text consumer-side, matching the historical fallback.
                    let idx: usize = v_text.parse().unwrap_or(0);
                    CellValue::Shared { si: idx }
                }
                "str" => CellValue::Text {
                    text: v_text,
                    runs: None,
                    phonetic_runs: Vec::new(),
                    phonetic_pr: None,
                },
                "b" => CellValue::Bool {
                    bool: v_text == "1" || v_text == "true",
                },
                "e" => CellValue::Error { error: v_text },
                _ => {
                    if let Ok(n) = v_text.parse::<f64>() {
                        CellValue::Number { number: n }
                    } else {
                        CellValue::Text {
                            text: v_text,
                            runs: None,
                            phonetic_runs: Vec::new(),
                            phonetic_pr: None,
                        }
                    }
                }
            }
        };

        // Furigana display resolves as `cell/@ph ?? row/@ph ?? false`:
        // - ECMA-376 §18.3.1.4 `<c ph>` — per-cell toggle, wins when present
        //   (including an explicit `ph="0"` that overrides an enabled row).
        // - ECMA-376 §18.3.1.73 `<row ph>` — inherited when the cell omits `@ph`.
        // - otherwise the schema default (false): a cell whose String Item
        //   carries `<rPh>` runs still shows NO furigana unless opted in.
        let show_phonetic = attr_bool(&c_node, "ph").unwrap_or(row_ph);

        cells.push(Cell {
            col,
            row,
            value,
            style_index,
            formula,
            show_phonetic,
        });
    }
    Ok(cells)
}

/// Parse an `ST_Boolean` (ECMA-376 §22.9.2.7, xsd:boolean) attribute value.
/// Accepts `1`/`true`/`on` as true and `0`/`false`/`off` as false (case-insensitive).
/// Returns `None` when the attribute is absent so callers can apply their own default.
pub(crate) fn attr_bool(node: &roxmltree::Node, name: &str) -> Option<bool> {
    node.attribute(name)
        .map(|v| matches!(v.trim().to_ascii_lowercase().as_str(), "1" | "true" | "on"))
}

pub(crate) fn parse_cell_ref(reference: &str) -> (u32, u32) {
    // Optional metadata consumers historically split only the leading letters,
    // defaulted an unparsable suffix to row 1, and retained the parsed column.
    // Preserve that leniency while making column arithmetic overflow-safe.
    let split = reference
        .find(|character: char| !character.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (column_text, row_text) = reference.split_at(split);
    let column = checked_column_ordinal(column_text).unwrap_or(0);
    let row = row_text.parse::<u32>().unwrap_or(1);
    (column, row)
}

fn parse_cell_ref_checked(reference: &str) -> Result<(u32, u32), String> {
    try_parse_cell_ref(reference).ok_or_else(|| {
        format!("worksheet cell has invalid or overflowing r reference: {reference}")
    })
}

fn try_parse_cell_ref(reference: &str) -> Option<(u32, u32)> {
    let split = reference
        .find(|character: char| !character.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (column_text, row_text) = reference.split_at(split);
    if column_text.is_empty()
        || row_text.is_empty()
        || !row_text.chars().all(|character| character.is_ascii_digit())
    {
        return None;
    }
    let column = checked_column_ordinal(column_text)?;
    let row = row_text.parse::<u32>().ok()?;
    Some((column, row))
}

fn checked_column_ordinal(column_text: &str) -> Option<u32> {
    column_text.chars().try_fold(0u32, |value, character| {
        let digit = (character.to_ascii_uppercase() as u32)
            .checked_sub('A' as u32)?
            .checked_add(1)?;
        value.checked_mul(26)?.checked_add(digit)
    })
}

/// Resolve a 1-based ordinal (a `<row>`'s row number or a `<c>`'s column) that
/// may omit its explicit position, tracking the running previous value in place.
///
/// ECMA-376 marks `@r` `use="optional"` on both `CT_Row` (§18.3.1.73) and
/// `CT_Cell` (§18.3.1.4). The spec grants the optionality but does not spell out
/// how an omitted value resolves; the de-facto consumer convention (Excel,
/// LibreOffice, SheetJS all agree, and no competing interpretation exists) is
/// ordinal document order: an omitted value is the previous sibling's + 1 (the
/// first element, with `*prev == 0` meaning "none yet", lands at 1), and an
/// explicit value re-anchors the running counter for later omitted siblings.
///
/// This is the single primitive shared by worksheet row and cell sequencing.
/// It also enforces the SpreadsheetML grid maxima, so an explicit maximum
/// followed by an omitted ordinal fails instead of overflowing or creating an
/// unrenderable coordinate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SpreadsheetOrdinal {
    Row,
    Column,
}

impl SpreadsheetOrdinal {
    const fn max(self) -> u32 {
        match self {
            Self::Row => 1_048_576,
            Self::Column => 16_384,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Row => "row",
            Self::Column => "column",
        }
    }
}

pub(crate) fn resolve_implicit_ordinal(
    explicit: Option<u32>,
    prev: &mut u32,
    kind: SpreadsheetOrdinal,
) -> Result<u32, String> {
    let resolved = match explicit {
        Some(value) => value,
        None => prev
            .checked_add(1)
            .ok_or_else(|| format!("worksheet {} ordinal overflows u32", kind.name()))?,
    };
    if resolved == 0 || resolved > kind.max() {
        return Err(format!(
            "worksheet {} ordinal is outside SpreadsheetML range 1..={}: {resolved}",
            kind.name(),
            kind.max()
        ));
    }
    *prev = resolved;
    Ok(resolved)
}

// ===========================
//  Native (non-WASM) API
// ===========================

/// Returns workbook overview (sheet names and metadata) as JSON.
/// Native equivalent of `parse_xlsx` for use from the MCP server.
pub fn parse_workbook_native(data: &[u8]) -> Result<String, String> {
    parse_xlsx_inner(data)
        .and_then(|wb| serde_json::to_string(&wb.workbook).map_err(|e| e.to_string()))
}

/// Parse the workbook and project every sheet to GitHub-flavoured markdown:
/// `## SheetName` headings followed by a pipe table per sheet. Merged-cell
/// continuation cells are rendered as empty; the display value comes from the
/// WASM-callable markdown projection (mirrors `to_markdown_native`).
#[wasm_bindgen]
pub fn xlsx_to_markdown(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    to_markdown_impl_with_limits(data, max_archive_entry_bytes, max_total_inflated_bytes)
        .map_err(|error| JsValue::from_str(&error))
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "xl/media/image1.png") from an xlsx zip archive. Thin `wasm_bindgen` wrapper
/// over the shared [`ooxml_common::zip::extract_zip_entry`] reader; used by the
/// main thread to lazily materialize image blobs on demand.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    ooxml_common::zip::extract_zip_entry(
        data,
        path,
        ooxml_common::resource::OoxmlFormat::Xlsx,
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    )
    .map_err(|e| JsValue::from_str(&e))
}

/// A stateful handle over an opened xlsx archive.
///
/// The retained package session eliminates repeated archive and shared-part work:
///
/// 1. **Buffer copy + central-directory scan** (like docx / pptx): `new` moves
///    the WASM-owned bytes into one package session and indexes the ZIP once;
///    the workbook index, worksheet cursors, and image reads reuse that session.
/// 2. **Shared-part reuse (D3)**: the handle parses `xl/workbook.xml`,
///    `xl/sharedStrings.xml`, and the theme once, then each worksheet cursor reads
///    only that sheet's XML and drawings.
///
/// The package session and every cached part are fully owned (no borrow into the
/// input), which is what lets them live in a `#[wasm_bindgen]` struct. Limits and
/// poison state are retained by that same session across public operations.
#[wasm_bindgen]
pub struct XlsxArchive {
    /// The opened archive, or the container-open error string when the ZIP itself
    /// was truncated / corrupt (#774, RB7 MAJOR). Deferring the failure here —
    /// instead of erroring out of `new` — lets `parse()` and worksheet cursors return a
    /// degraded placeholder (symmetric with a corrupt inner sheet) rather than the
    /// constructor throwing an opaque error the viewer can't turn into a
    /// placeholder tab.
    archive: Result<XlsxZip, String>,
    /// Workbook-level parts parsed once and reused across sheet switches. Loaded
    /// lazily on the first workbook-index or worksheet operation.
    shared: Option<WorkbookShared>,
    /// At most one worksheet decoder can hold the package operation lease.
    /// The cursor owns its entry stream; it does not borrow `archive`.
    active_worksheet: Option<ActiveWorksheetCursor>,
    last_cursor_pull_terminal: bool,
    terminal_awaiting_ack: bool,
    last_cursor_usage: Option<ResourceUsage>,
}

struct ActiveWorksheetCursor {
    source: ActiveWorksheetSource,
    sheet_index: u32,
    name: String,
    sheet_path: String,
    reference_index: Option<WorksheetCellLookupBuilder>,
}

enum ActiveWorksheetSource {
    Streaming(Box<WorksheetCursor>),
    Ready(Box<Worksheet>),
    DeferredFailure(CursorOpenFailure),
    Prepared,
}

#[derive(Clone)]
enum CursorOpenFailure {
    Container(String),
    Sheet(String),
}

fn serialize_cursor_finished(
    worksheet: Worksheet,
    limit_reporter: Option<&PackageLimitReporter>,
    part: Option<&str>,
) -> Result<Vec<u8>, String> {
    #[derive(serde::Serialize)]
    struct Finished {
        kind: &'static str,
        worksheet: Worksheet,
    }
    let finished = Finished {
        kind: "finished",
        worksheet,
    };
    let json_bytes = measure_json(&finished)?.json_bytes;
    if let Some(reporter) = limit_reporter {
        reporter.observe_hard_limit(
            HardResourceLimitKind::WorksheetJsonBytes,
            part,
            HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
            json_bytes,
        )?;
    }
    serde_json::to_vec(&finished).map_err(|error| format!("serialize error: {error}"))
}

#[wasm_bindgen]
impl XlsxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// Resource limits are retained and applied on every subsequent parse
    /// method. Shared workbook parts are parsed lazily on the first
    /// `parse`/`parse_sheet`.
    ///
    /// `data` is taken by value (`Vec<u8>`): wasm-bindgen copies the JS `Uint8Array`
    /// once into a WASM-owned buffer and hands that allocation to Rust as this
    /// `Vec`, which `Cursor` then takes by value — a single copy across the
    /// JS→WASM boundary. Taking `&[u8]` would force a second `to_vec()` copy so
    /// the `Cursor` could own its backing store, transiently doubling WASM
    /// linear memory to ~2x the file size during construction.
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: Vec<u8>,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
        max_archive_entries: Option<u64>,
    ) -> Result<XlsxArchive, JsValue> {
        console_error_panic_hook::set_once();
        // #774 (RB7 MAJOR): a truncated / corrupt CONTAINER is deferred, not
        // thrown, so `parse()` / `parse_sheet()` can degrade it to a placeholder
        // instead of the constructor failing with an opaque error.
        let archive = open_zip_with_policy(
            data,
            max_archive_entry_bytes,
            max_total_inflated_bytes,
            max_archive_entries,
        );
        if let Err(error) = &archive {
            if error.starts_with("OOXML_RESOURCE_LIMIT:") {
                return Err(JsValue::from_str(error));
            }
        }
        Ok(XlsxArchive {
            archive,
            shared: None,
            active_worksheet: None,
            last_cursor_pull_terminal: false,
            terminal_awaiting_ack: false,
            last_cursor_usage: None,
        })
    }

    /// Parse (once) and return the workbook-level shared parts, caching them for
    /// reuse. Borrows `self` split so the cached `shared` and the `archive` can be
    /// used together by callers. Assumes the container opened; the corrupt-container
    /// case is short-circuited by the callers before they reach here.
    fn ensure_shared(&mut self) -> Result<(), String> {
        if self.shared.is_none() {
            let zip = self
                .archive
                .as_mut()
                .map_err(|error| format!("xlsx-parser error: {error}"))?;
            let shared = WorkbookShared::load(zip)?;
            self.shared = Some(shared);
        }
        Ok(())
    }

    /// Parse the workbook index (sheet list + styles + shared strings) and return
    /// it as UTF-8 JSON bytes. Byte-for-byte identical to `parse_xlsx`. When the
    /// CONTAINER failed to open (#774) the model is a degraded placeholder
    /// workbook tagged with the container.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        if let Err(error) = &self.archive {
            let workbook = degraded_container_workbook(error.clone());
            return serde_json::to_vec(&workbook)
                .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")));
        }
        self.archive
            .as_mut()
            .expect("container open checked above")
            .begin_operation("parse")
            .map_err(|error| JsValue::from_str(&error))?;
        let result = (|| -> Result<Vec<u8>, String> {
            let styles = if let Some(shared) = &self.shared {
                // A sheet cursor may have initialized only the lightweight style
                // projection. Parse the full style model only when the caller
                // later asks for the workbook index, and move it directly into
                // the serialized result.
                let theme_colors = Rc::clone(&shared.theme_colors);
                let zip = self.archive.as_mut().expect("container open checked above");
                parse_styles(zip, theme_colors.as_ref()).map(|parsed| parsed.styles)
            } else {
                let zip = self.archive.as_mut().expect("container open checked above");
                let (shared, styles) = WorkbookShared::load_with_styles(zip)?;
                self.shared = Some(shared);
                styles
            };
            let shared = self.shared.as_ref().expect("shared loaded above");
            let zip = self.archive.as_mut().expect("container open checked above");
            let workbook = parse_xlsx_inner_with(zip, shared, styles)?;
            serde_json::to_vec(&workbook).map_err(|error| format!("serialize error: {error}"))
        })();
        let zip = self.archive.as_mut().expect("container open checked above");
        let result = settle_xlsx_operation(zip, result);
        if result.is_err() && zip.assert_healthy().is_err() {
            self.shared = None;
        }
        result.map_err(|error| JsValue::from_str(&error))
    }

    /// Fail cached worker operations after this package session was poisoned.
    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        match &self.archive {
            Ok(archive) => archive
                .assert_healthy()
                .map_err(|error| JsValue::from_str(&error)),
            Err(_) => Ok(()),
        }
    }

    /// Session-wide archive accounting after workbook bootstrap or any later
    /// operation. This is diagnostic data, not an allocator-memory estimate.
    pub fn resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        let usage = self
            .archive
            .as_ref()
            .map(XlsxZip::usage)
            .map_err(|_| JsValue::from_str("xlsx resource usage is unavailable"))?;
        serde_json::to_vec(&usage)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Open the resumable production worksheet pipeline. Only one cursor may
    /// hold the archive's decoder lease at a time.
    pub fn open_sheet_cursor(&mut self, sheet_index: u32, name: &str) -> Result<(), JsValue> {
        if self.active_worksheet.is_some() {
            return Err(JsValue::from_str("worksheet cursor is already open"));
        }
        self.last_cursor_pull_terminal = false;
        self.terminal_awaiting_ack = false;
        self.last_cursor_usage = None;
        if let Err(error) = &self.archive {
            self.active_worksheet = Some(ActiveWorksheetCursor {
                source: ActiveWorksheetSource::DeferredFailure(CursorOpenFailure::Container(
                    error.clone(),
                )),
                sheet_index,
                name: CONTAINER_PART.to_string(),
                sheet_path: String::new(),
                reference_index: None,
            });
            return Ok(());
        }
        let zip = self
            .archive
            .as_mut()
            .map_err(|error| JsValue::from_str(&format!("xlsx-parser error: {error}")))?;
        zip.begin_operation("worksheet-cursor")
            .map_err(|error| JsValue::from_str(&error))?;
        let result = (|| -> Result<ActiveWorksheetCursor, String> {
            self.ensure_shared()?;
            let shared = self.shared.as_ref().expect("shared loaded above");
            let rels_doc = parse_guarded(&shared.rels_xml).map_err(|error| error.to_string())?;
            let sheet = shared
                .sheets
                .get(sheet_index as usize)
                .ok_or_else(|| format!("sheet index {sheet_index} out of range"))?;
            let sheet_path = resolve_sheet_path(&rels_doc, &sheet.r_id)
                .ok_or_else(|| format!("rId {} not found in rels", sheet.r_id))?;
            let part = format!("xl/{sheet_path}");
            let zip = self.archive.as_mut().expect("container open checked above");
            let sheet_part_kind = resolve_sheet_part_kind(&rels_doc, &sheet.r_id);
            let source = match sheet_part_kind {
                SheetPartKind::ChartSheet => {
                    match parse_chart_sheet_shell(zip, &part, name).and_then(|parsed| {
                        finalize_projected_sheet(
                            zip,
                            shared,
                            sheet_index,
                            name,
                            &sheet_path,
                            parsed,
                            CurrentSheetLookup::BuildFromMaterializedRows,
                        )
                    }) {
                        Ok(worksheet) => ActiveWorksheetSource::Ready(Box::new(worksheet)),
                        Err(error) => {
                            zip.assert_healthy()?;
                            ActiveWorksheetSource::DeferredFailure(CursorOpenFailure::Sheet(error))
                        }
                    }
                }
                SheetPartKind::DialogSheet => match parse_dialog_sheet_shell(zip, &part, name) {
                    Ok((worksheet, _, _)) => ActiveWorksheetSource::Ready(Box::new(worksheet)),
                    Err(error) => {
                        zip.assert_healthy()?;
                        ActiveWorksheetSource::DeferredFailure(CursorOpenFailure::Sheet(error))
                    }
                },
                SheetPartKind::Worksheet => match zip.open_worksheet_cursor(
                    &part,
                    Rc::clone(&shared.shared_strings),
                    Rc::clone(&shared.theme_colors),
                ) {
                    Ok(cursor) => ActiveWorksheetSource::Streaming(Box::new(cursor)),
                    Err(error) => {
                        zip.assert_healthy()?;
                        ActiveWorksheetSource::DeferredFailure(CursorOpenFailure::Sheet(error))
                    }
                },
            };
            Ok(ActiveWorksheetCursor {
                source,
                sheet_index,
                name: name.to_string(),
                sheet_path,
                reference_index: (sheet_part_kind == SheetPartKind::Worksheet)
                    .then(WorksheetCellLookupBuilder::bounded),
            })
        })();
        match result {
            Ok(cursor) => {
                self.active_worksheet = Some(cursor);
                Ok(())
            }
            Err(error) => {
                let zip = self.archive.as_mut().expect("container open checked above");
                zip.cancel_operation();
                Err(JsValue::from_str(&error))
            }
        }
    }

    /// Pull rows or the terminal row-free worksheet product as UTF-8 JSON.
    /// The worker adapter applies hard serialized-byte credit to these bytes.
    pub fn pull_sheet_cursor(&mut self, row_credit: usize) -> Result<Vec<u8>, JsValue> {
        self.pull_sheet_cursor_inner(row_credit)
            .map_err(|error| JsValue::from_str(&error))
    }

    fn pull_sheet_cursor_inner(&mut self, row_credit: usize) -> Result<Vec<u8>, String> {
        if self.terminal_awaiting_ack {
            return Err(
                "worksheet terminal product must be acknowledged before pulling".to_string(),
            );
        }
        self.last_cursor_pull_terminal = false;
        let deferred = match &self
            .active_worksheet
            .as_ref()
            .ok_or_else(|| "worksheet cursor is not open".to_string())?
            .source
        {
            ActiveWorksheetSource::DeferredFailure(failure) => Some(failure.clone()),
            ActiveWorksheetSource::Streaming(_) | ActiveWorksheetSource::Ready(_) => None,
            ActiveWorksheetSource::Prepared => {
                return Err("worksheet terminal product is prepared".to_string());
            }
        };
        if let Some(failure) = deferred {
            let active = self
                .active_worksheet
                .as_ref()
                .expect("cursor checked above");
            let part = (!active.sheet_path.is_empty()).then(|| format!("xl/{}", active.sheet_path));
            let worksheet = match failure {
                CursorOpenFailure::Container(error) => degraded_container_sheet(error),
                CursorOpenFailure::Sheet(error) => Worksheet::placeholder(
                    &active.name,
                    format!("xl/{}: {error}", active.sheet_path),
                ),
            };
            let reporter = match self.archive.as_ref() {
                Ok(zip) => Some(zip.active_operation()?.limit_reporter()?),
                Err(_) => None,
            };
            let bytes = serialize_cursor_finished(worksheet, reporter.as_ref(), part.as_deref())?;
            self.active_worksheet
                .as_mut()
                .expect("cursor checked above")
                .source = ActiveWorksheetSource::Prepared;
            self.last_cursor_pull_terminal = true;
            self.terminal_awaiting_ack = true;
            return Ok(bytes);
        }
        let ready = {
            let active = self
                .active_worksheet
                .as_mut()
                .ok_or_else(|| "worksheet cursor is not open".to_string())?;
            match std::mem::replace(&mut active.source, ActiveWorksheetSource::Prepared) {
                ActiveWorksheetSource::Ready(worksheet) => Some(worksheet),
                other => {
                    active.source = other;
                    None
                }
            }
        };
        if let Some(worksheet) = ready {
            let active = self
                .active_worksheet
                .as_ref()
                .expect("cursor checked above");
            let part = format!("xl/{}", active.sheet_path);
            let reporter = self
                .archive
                .as_ref()
                .expect("container open checked above")
                .active_operation()?
                .limit_reporter()?;
            let bytes = serialize_cursor_finished(*worksheet, Some(&reporter), Some(&part))?;
            self.last_cursor_pull_terminal = true;
            self.terminal_awaiting_ack = true;
            return Ok(bytes);
        }
        let shared_strings = Rc::clone(
            &self
                .shared
                .as_ref()
                .ok_or_else(|| "worksheet cursor shared state is missing".to_string())?
                .shared_strings,
        );
        let source = &mut self
            .active_worksheet
            .as_mut()
            .ok_or_else(|| "worksheet cursor is not open".to_string())?
            .source;
        let pull = match source {
            ActiveWorksheetSource::Streaming(cursor) => {
                cursor.pull(row_credit, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES)
            }
            _ => unreachable!("source checked above"),
        };
        match pull {
            Ok(WorksheetCursorPull::Rows { rows, .. }) => {
                extend_lookup_transactionally(
                    &mut self
                        .active_worksheet
                        .as_mut()
                        .expect("cursor checked above")
                        .reference_index,
                    &rows,
                    shared_strings.as_ref(),
                );
                #[derive(serde::Serialize)]
                struct Rows<'a> {
                    kind: &'static str,
                    rows: &'a [Row],
                }
                serde_json::to_vec(&Rows {
                    kind: "rows",
                    rows: &rows,
                })
                .map_err(|error| format!("serialize error: {error}"))
            }
            Ok(WorksheetCursorPull::Finished(tail)) => {
                let active = self
                    .active_worksheet
                    .as_mut()
                    .expect("cursor checked above");
                let result = (|| -> Result<Vec<u8>, String> {
                    let shared = self.shared.as_ref().expect("shared loaded above");
                    let parsed = parse_projected_worksheet_tail(
                        tail,
                        shared.theme_colors.as_ref(),
                        &active.name,
                    )?;
                    let current_index = active.reference_index.take().and_then(|mut builder| {
                        builder.mark_hidden_columns(&parsed.0.col_hidden)?;
                        Some(builder.finish())
                    });
                    let zip = self.archive.as_mut().expect("container open checked above");
                    let worksheet = finalize_projected_sheet(
                        zip,
                        shared,
                        active.sheet_index,
                        &active.name,
                        &active.sheet_path,
                        parsed,
                        CurrentSheetLookup::Seed(current_index),
                    )?;
                    // Ancillary parsers intentionally degrade malformed parts to
                    // placeholders, but a resource-limit violation is terminal
                    // for the package operation and must never be downgraded.
                    zip.assert_healthy()?;
                    let reporter = zip.active_operation()?.limit_reporter()?;
                    let part = format!("xl/{}", active.sheet_path);
                    serialize_cursor_finished(worksheet, Some(&reporter), Some(&part))
                })();
                let bytes = match result {
                    Ok(bytes) => bytes,
                    Err(error) => {
                        let zip = self.archive.as_mut().expect("container open checked above");
                        if let Err(resource_error) = zip.assert_healthy() {
                            self.active_worksheet.take();
                            zip.cancel_operation();
                            return Err(resource_error);
                        }
                        let active = self
                            .active_worksheet
                            .as_ref()
                            .expect("cursor checked above");
                        let part = format!("xl/{}", active.sheet_path);
                        let reporter = self
                            .archive
                            .as_ref()
                            .expect("container open checked above")
                            .active_operation()?
                            .limit_reporter()?;
                        serialize_cursor_finished(
                            Worksheet::placeholder(&active.name, format!("{part}: {error}")),
                            Some(&reporter),
                            Some(&part),
                        )?
                    }
                };
                self.active_worksheet
                    .as_mut()
                    .expect("cursor checked above")
                    .source = ActiveWorksheetSource::Prepared;
                self.last_cursor_pull_terminal = true;
                self.terminal_awaiting_ack = true;
                Ok(bytes)
            }
            Err(error) => {
                let zip = self.archive.as_mut().expect("container open checked above");
                if let Err(resource_error) = zip.assert_healthy() {
                    self.active_worksheet.take();
                    zip.cancel_operation();
                    return Err(resource_error);
                }
                let active = self
                    .active_worksheet
                    .as_ref()
                    .expect("cursor checked above");
                let part = format!("xl/{}", active.sheet_path);
                let reporter = self
                    .archive
                    .as_ref()
                    .expect("container open checked above")
                    .active_operation()?
                    .limit_reporter()?;
                let bytes = serialize_cursor_finished(
                    Worksheet::placeholder(&active.name, format!("{part}: {error}")),
                    Some(&reporter),
                    Some(&part),
                )?;
                self.active_worksheet
                    .as_mut()
                    .expect("cursor checked above")
                    .source = ActiveWorksheetSource::Prepared;
                self.last_cursor_pull_terminal = true;
                self.terminal_awaiting_ack = true;
                Ok(bytes)
            }
        }
    }

    /// Whether the immediately preceding successful pull produced the terminal
    /// row-free worksheet product.
    pub fn sheet_cursor_pull_finished(&self) -> bool {
        self.last_cursor_pull_terminal
    }

    /// Current operation accounting checkpoint for the shared pull protocol.
    pub fn sheet_cursor_resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        let usage = self
            .archive
            .as_ref()
            .ok()
            .and_then(|zip| zip.operation.usage())
            .or(self.last_cursor_usage)
            .ok_or_else(|| JsValue::from_str("worksheet cursor usage is unavailable"))?;
        serde_json::to_vec(&usage)
            .map_err(|error| JsValue::from_str(&format!("serialize error: {error}")))
    }

    /// Commit a prepared terminal product. The worker calls this only from the
    /// shared pull host's acknowledgement of the final sequence.
    pub fn acknowledge_sheet_cursor_terminal(&mut self) -> Result<(), JsValue> {
        self.acknowledge_sheet_cursor_terminal_inner()
            .map_err(|error| JsValue::from_str(&error))
    }

    fn acknowledge_sheet_cursor_terminal_inner(&mut self) -> Result<(), String> {
        if !self.terminal_awaiting_ack {
            return Err("worksheet terminal product is not awaiting acknowledgement".to_string());
        }
        if self.archive.is_err() {
            self.active_worksheet.take();
            self.terminal_awaiting_ack = false;
            return Ok(());
        }
        let zip = self.archive.as_mut().expect("container open checked above");
        self.last_cursor_usage = zip.operation.usage();
        let result = zip.finish_operation();
        if let Err(resource_error) = zip.assert_healthy() {
            zip.cancel_operation();
            self.active_worksheet.take();
            self.terminal_awaiting_ack = false;
            return Err(resource_error);
        }
        result?;
        self.active_worksheet.take();
        self.terminal_awaiting_ack = false;
        Ok(())
    }

    /// Cancel an open cursor and release its decoder lease. Idempotent.
    pub fn cancel_sheet_cursor(&mut self) {
        if let Some(mut active) = self.active_worksheet.take() {
            if let ActiveWorksheetSource::Streaming(cursor) = &mut active.source {
                cursor.cancel();
            }
        }
        self.terminal_awaiting_ack = false;
        self.last_cursor_pull_terminal = false;
        if let Ok(zip) = self.archive.as_mut() {
            self.last_cursor_usage = zip.operation.usage().or(self.last_cursor_usage);
            zip.cancel_operation();
        }
    }

    /// Close an open cursor and release its decoder lease. Idempotent.
    pub fn close_sheet_cursor(&mut self) {
        if self.terminal_awaiting_ack {
            self.cancel_sheet_cursor();
        } else if let Some(mut active) = self.active_worksheet.take() {
            if let ActiveWorksheetSource::Streaming(cursor) = &mut active.source {
                cursor.close();
            }
            if let Ok(zip) = self.archive.as_mut() {
                self.last_cursor_usage = zip.operation.usage().or(self.last_cursor_usage);
                zip.cancel_operation();
            }
            self.last_cursor_pull_terminal = false;
        }
    }

    /// Extract raw bytes for one embedded image entry (e.g.
    /// "xl/media/image1.png") from the retained archive. Twin of the free
    /// `extract_image`, but reads through the already-open archive. A corrupt
    /// container has no entries, so this surfaces the container-open error.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let zip = self
            .archive
            .as_mut()
            .map_err(|e| JsValue::from_str(&format!("xlsx-parser error: {e}")))?;
        zip.run_operation("extract-image", |zip| read_zip_bytes(zip, path))
            .map_err(|error| JsValue::from_str(&error))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `xlsx_to_markdown`. A corrupt container degrades to an empty document.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        let zip = self
            .archive
            .as_mut()
            .map_err(|error| JsValue::from_str(&format!("xlsx-parser error: {error}")))?;
        zip.run_operation("markdown", to_markdown_from_archive)
            .map_err(|error| JsValue::from_str(&error))
    }
}

/// cached `<v>` so formula formulas show their results, not the formula text.
/// Designed for AI agents that need to read the spreadsheet content
/// efficiently — drops styling, formatting, charts, sparklines, drawings.
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    to_markdown_impl(data)
}

/// Shared implementation between `to_markdown_native` (mcp-server) and
/// `xlsx_to_markdown` (browser / Node WASM).
fn to_markdown_impl(data: &[u8]) -> Result<String, String> {
    to_markdown_impl_with_limits(data, None, None)
}

fn to_markdown_impl_with_limits(
    data: &[u8],
    max_archive_entry_bytes: Option<u64>,
    max_total_inflated_bytes: Option<u64>,
) -> Result<String, String> {
    // #774: a corrupt CONTAINER has no sheets to render — degrade to an empty
    // markdown document instead of erroring, symmetric with the JSON path.
    let mut archive = match open_zip_with_limits(
        data.to_vec(),
        max_archive_entry_bytes,
        max_total_inflated_bytes,
    ) {
        Ok(zip) => zip,
        Err(error) if error.starts_with("OOXML_RESOURCE_LIMIT:") => return Err(error),
        Err(_) => return Ok(String::new()),
    };
    archive.run_operation("markdown", to_markdown_from_archive)
}

/// Render every sheet of an opened archive to markdown. Shared by the free
/// `xlsx_to_markdown` / `to_markdown_native` and `XlsxArchive::to_markdown`;
/// loads the workbook-level [`WorkbookShared`] parts once and renders each sheet
/// through the same `parse_sheet_with` pipeline as the JSON path (so markdown and
/// JSON never diverge on cell values).
fn to_markdown_from_archive(archive: &mut XlsxZip) -> Result<String, String> {
    let shared = WorkbookShared::load(archive)?;
    let mut out = String::new();
    let mut review_comments = String::new();
    let mut has_comments = false;
    for (idx, sheet_meta) in shared.sheets.iter().enumerate() {
        let sheet_json =
            parse_sheet_with(archive, &shared, idx as u32, &sheet_meta.name).map_err(|error| {
                format!(
                    "sheet '{}' (#{}) parse failed: {}",
                    sheet_meta.name, idx, error
                )
            })?;
        let sheet: serde_json::Value =
            serde_json::from_slice(&sheet_json).map_err(|e| e.to_string())?;
        markdown::render_sheet(&sheet, &shared.shared_strings, &mut out);
        markdown::render_sheet_comments(&sheet, &mut review_comments, &mut has_comments);
    }
    out.push_str(&review_comments);
    Ok(out)
}

/// Parses a single worksheet by 0-based index and returns it as JSON.
/// Native equivalent of `parse_sheet` for use from the MCP server. Shares the
/// exact per-sheet pipeline (`WorkbookShared::load` + `parse_sheet_with`) with
/// the WASM `parse_sheet`, then decodes the JSON bytes to a `String` — so the
/// native and WASM paths can never drift.
pub fn parse_sheet_native(data: &[u8], sheet_index: u32, name: &str) -> Result<String, String> {
    // #774: mirror the WASM `parse_sheet` — a corrupt CONTAINER degrades to the
    // container-tagged placeholder sheet rather than erroring.
    let mut archive = match open_zip(data.to_vec()) {
        Ok(zip) => zip,
        Err(e) => {
            let ws = degraded_container_sheet(e);
            return serde_json::to_string(&ws).map_err(|e| e.to_string());
        }
    };
    archive.run_operation("parse-sheet", |archive| {
        let shared = WorkbookShared::load(archive)?;
        let json = parse_sheet_with(archive, &shared, sheet_index, name)?;
        String::from_utf8(json).map_err(|error| error.to_string())
    })
}

#[cfg(test)]
mod tab_color_tests {
    use super::extract_tab_color_from_head;

    const THEME: &[&str] = &[
        "#000000", "#FFFFFF", "#44546A", "#E7E6E6", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
    ];

    fn theme() -> Vec<String> {
        THEME.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn tab_color_rgb() {
        let head = r#"<?xml version="1.0"?><worksheet><sheetPr><tabColor rgb="FFFF0000"/></sheetPr><dimension ref="A1"/><sheetData>"#;
        assert_eq!(
            extract_tab_color_from_head(head, &theme()).as_deref(),
            Some("#FF0000")
        );
    }

    #[test]
    fn tab_color_theme_with_tint() {
        // theme="4" (Excel-internal accent1) resolves to #4472C4; a tint just
        // needs to produce *something* different — exact value covered by apply_tint.
        let head = r#"<worksheet><sheetPr codeName="S1"><tabColor theme="4" tint="-0.249977111117893"/></sheetPr><sheetData/></worksheet>"#;
        let got = extract_tab_color_from_head(head, &theme());
        assert!(got.is_some(), "theme tab color should resolve");
        assert_ne!(
            got.as_deref(),
            Some("#4472C4"),
            "tint should darken the base"
        );
    }

    #[test]
    fn tab_color_absent() {
        let head = r#"<worksheet><sheetPr/><dimension ref="A1"/><sheetData><row/></sheetData></worksheet>"#;
        assert_eq!(extract_tab_color_from_head(head, &theme()), None);
    }

    #[test]
    fn tab_color_not_searched_past_sheetdata() {
        // A stray "tabColor" token inside the body must not be misread.
        let head =
            r#"<worksheet><sheetPr/><sheetData><c><is><t>tabColor rgb="00FF00"</t></is></c>"#;
        assert_eq!(extract_tab_color_from_head(head, &theme()), None);
    }
}

#[cfg(test)]
mod sheet_view_tests {
    use super::parse_worksheet;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// ECMA-376 §18.3.1.87 `<sheetView rightToLeft="1">` mirrors the entire
    /// grid (column A on the right).
    #[test]
    fn sheet_view_right_to_left() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert!(ws.right_to_left, "rightToLeft=\"1\" → right_to_left true");
    }

    /// Absent `@rightToLeft` defaults to false (left-to-right).
    #[test]
    fn sheet_view_right_to_left_defaults_false() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert!(
            !ws.right_to_left,
            "absent @rightToLeft → right_to_left false"
        );
    }

    /// ECMA-376 §22.9.2.7 `ST_Boolean` allows `true`/`false` as well as `1`/`0`.
    /// LibreOffice writes `<col customWidth="true" .../>`; the parser must honor
    /// the recorded width instead of skipping the `<col>` (which would leave the
    /// column at `defaultColWidth`).
    #[test]
    fn col_custom_width_accepts_true_literal() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col customWidth="true" min="1" max="1" width="22"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(
            ws.col_widths.get(&1).copied(),
            Some(22.0),
            "customWidth=\"true\" → width 22 recorded for column 1"
        );
    }

    /// `customWidth="1"` (Excel's spelling) must keep working after the helper change.
    #[test]
    fn col_custom_width_accepts_one_literal() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col customWidth="1" min="2" max="2" width="10"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.col_widths.get(&2).copied(), Some(10.0));
    }

    /// ECMA-376 §18.3.1.13 defines `customWidth` as metadata indicating that
    /// the width differs from the default; it is not a condition for applying
    /// an authored `width`. Excel can omit the flag on style-wide ranges.
    #[test]
    fn col_width_without_custom_width_is_preserved_as_a_compact_range() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col min="2" max="16384" width="10.83203125" style="44"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert!(
            ws.col_widths.is_empty(),
            "wide ranges must not expand into 16K JSON entries"
        );
        assert_eq!(
            ws.col_width_ranges,
            vec![crate::types::ColumnWidthRange {
                min: 2,
                max: 16_384,
                width: 10.83203125,
            }],
        );
    }

    #[test]
    fn column_style_is_preserved_without_expanding_the_range() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col min="1" max="16384" style="17"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(
            ws.col_style_ranges,
            vec![crate::types::ColumnStyleRange {
                min: 1,
                max: 16_384,
                style_index: 17,
            }],
        );
    }

    #[test]
    fn explicit_zero_cell_style_remains_distinct_from_column_inheritance() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col min="1" max="1" style="17"/></cols><sheetData><row r="1"><c r="A1" s="0"/></row></sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(ws.rows[0].cells[0].style_index, Some(0));
    }

    #[test]
    fn cell_without_authored_style_inherits_its_column_style() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col min="1" max="2" style="17"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>2</v></c></row></sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(ws.rows[0].cells[0].style_index, Some(17));
        assert_eq!(ws.rows[0].cells[1].style_index, None);
    }

    #[test]
    fn later_compact_col_width_range_overrides_legacy_point_projection() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols>
                <col min="2" max="2" width="20" customWidth="1"/>
                <col min="2" max="16384" width="10"/>
            </cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(ws.col_widths.get(&2).copied(), Some(10.0));
        assert_eq!(
            ws.col_width_ranges,
            vec![crate::types::ColumnWidthRange {
                min: 2,
                max: 16_384,
                width: 10.0,
            }],
        );
    }

    #[test]
    fn later_legacy_point_col_width_overrides_compact_range() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols>
                <col min="2" max="16384" width="10"/>
                <col min="2" max="2" width="20" customWidth="1"/>
            </cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(ws.col_widths.get(&2).copied(), Some(20.0));
        assert_eq!(ws.col_width_ranges[0].width, 10.0);
    }

    #[test]
    fn many_overlapping_full_col_width_ranges_are_normalized_in_bounded_work() {
        let mut cols = String::new();
        for chunk in 0..64 {
            let min = chunk * 256 + 1;
            let max = min + 255;
            cols.push_str(&format!(
                r#"<col min="{min}" max="{max}" width="11" customWidth="1"/>"#
            ));
        }
        for index in 0..50_000 {
            let width = 9 + (index % 2);
            cols.push_str(&format!(r#"<col min="1" max="16384" width="{width}"/>"#));
        }
        let xml = format!(r#"<worksheet xmlns="{NS}"><cols>{cols}</cols><sheetData/></worksheet>"#);
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        assert_eq!(ws.col_widths.len(), 16_384);
        assert!(ws.col_widths.values().all(|width| *width == 10.0));
        assert_eq!(ws.col_width_ranges.len(), 50_000);
    }

    /// ECMA-376 §18.3.1.81 `zeroHeight` hides rows by default, including rows
    /// not materialized in `<sheetData>`. The renderer represents that rule as
    /// a zero default axis size and skips the sparse ordinal run.
    #[test]
    fn sheet_format_zero_height_sets_zero_default_row_height() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetFormatPr defaultRowHeight="22" zeroHeight="1"/><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.default_row_height, 0.0);
    }

    #[test]
    fn sheet_format_preserves_manual_default_row_height_fact() {
        let automatic_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>"#
        );
        let manual_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetFormatPr defaultRowHeight="30" customHeight="1"/><sheetData/></worksheet>"#
        );
        let (automatic, _) =
            parse_worksheet(&automatic_xml, &[], &[], "Automatic").expect("worksheet parses");
        let (manual, _) =
            parse_worksheet(&manual_xml, &[], &[], "Manual").expect("worksheet parses");

        assert!(!automatic.default_row_height_custom);
        assert!(manual.default_row_height_custom);
        assert_eq!(manual.default_row_height, 30.0);
        assert!(
            serde_json::to_value(automatic)
                .expect("serializes")
                .get("defaultRowHeightCustom")
                .is_none(),
            "the false schema default stays wire-compatible"
        );
        assert_eq!(
            serde_json::to_value(manual)
                .expect("serializes")
                .get("defaultRowHeightCustom")
                .and_then(|value| value.as_bool()),
            Some(true)
        );
    }

    #[test]
    fn sheet_format_zero_height_preserves_explicit_visible_rows() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}">
                 <sheetFormatPr defaultRowHeight="22" zeroHeight="1"/>
                 <sheetData><row r="3" hidden="0"><c r="A3"/></row></sheetData>
               </worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.default_row_height, 0.0, "unspecified rows are hidden");
        assert_eq!(ws.row_heights.get(&3).copied(), Some(22.0));
        assert_eq!(ws.rows[0].height, Some(22.0));
        assert!(!ws.rows[0].hidden);
    }

    /// Malformed non-finite/negative dimensions never cross the parser/model
    /// boundary. Keep the schema defaults rather than poisoning cumulative
    /// geometry with NaN or a negative band size.
    #[test]
    fn malformed_sheet_and_row_dimensions_fall_back_safely() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}">
                 <sheetFormatPr defaultColWidth="NaN" defaultRowHeight="-2"/>
                 <cols><col customWidth="1" min="1" max="1" width="NaN"/></cols>
                 <sheetData><row r="1" ht="-4"/><row r="2" ht="NaN"/></sheetData>
               </worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.default_col_width, 8.43);
        assert_eq!(ws.default_row_height, 15.0);
        assert_eq!(ws.col_widths.get(&1).copied(), Some(8.43));
        assert!(ws.row_heights.is_empty());
    }

    /// The serialized worksheet JSON is deterministic: `colWidths` keys come out
    /// in ascending column order regardless of `<col>` declaration order, and
    /// two serializations of the same parse are byte-identical. This is the
    /// BTreeMap guarantee — with the former `HashMap` field the key order
    /// followed the randomized hash seed, so identical input could serialize to
    /// different byte streams across runs.
    #[test]
    fn worksheet_json_is_deterministic_and_key_ordered() {
        // Columns declared out of order (3, then 1, then 2).
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols>
                 <col customWidth="1" min="3" max="3" width="30"/>
                 <col customWidth="1" min="1" max="1" width="10"/>
                 <col customWidth="1" min="2" max="2" width="20"/>
               </cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        let json = serde_json::to_string(&ws).expect("serialize");
        // Two serializations of the same value are byte-identical.
        assert_eq!(json, serde_json::to_string(&ws).expect("serialize"));

        // colWidths keys appear in ascending column order in the JSON string.
        let widths = &json[json.find("\"colWidths\"").expect("colWidths present")..];
        let p1 = widths.find("\"1\"").expect("col 1 key");
        let p2 = widths.find("\"2\"").expect("col 2 key");
        let p3 = widths.find("\"3\"").expect("col 3 key");
        assert!(
            p1 < p2 && p2 < p3,
            "colWidths keys must serialize in ascending order (1,2,3), got positions {p1},{p2},{p3} in {widths}"
        );
    }

    // ── Outline grouping (ECMA-376 §18.3.1.13 / §18.3.1.61 / §18.3.1.73) ──

    /// The row-outline example from §18.3.1.73 (middle + lowest level collapsed):
    /// rows 6-8 are collapsed-hidden detail at levels 3/3/2, and row 9 is the
    /// level-1 summary carrying `collapsed="1"`. The parser must surface each
    /// row's `outlineLevel`, `collapsed`, and `hidden` flags verbatim.
    #[test]
    fn row_outline_levels_collapsed_and_hidden() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
                 <row r="6" hidden="1" outlineLevel="3"/>
                 <row r="7" hidden="1" outlineLevel="3"/>
                 <row r="8" hidden="1" outlineLevel="2"/>
                 <row r="9" hidden="1" outlineLevel="1" collapsed="1"/>
                 <row r="10" collapsed="1"/>
               </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        let by_idx = |i: u32| ws.rows.iter().find(|r| r.index == i).expect("row present");
        assert_eq!(by_idx(6).outline_level, 3);
        assert!(by_idx(6).hidden);
        assert!(!by_idx(6).collapsed);
        assert_eq!(by_idx(8).outline_level, 2);
        assert!(by_idx(8).hidden);
        assert_eq!(by_idx(9).outline_level, 1);
        assert!(by_idx(9).hidden);
        assert!(by_idx(9).collapsed);
        // Row 10 is the top-level summary: collapsed but visible, level 0.
        assert_eq!(by_idx(10).outline_level, 0);
        assert!(!by_idx(10).hidden);
        assert!(by_idx(10).collapsed);
    }

    /// `outlineLevel` is clamped to the §18.3.1.73 range max of 7.
    #[test]
    fn row_outline_level_clamped_to_seven() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData><row r="1" outlineLevel="9"/></sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.rows[0].outline_level, 7);
    }

    /// A grouped column at the *default* width (no `customWidth`, not hidden) must
    /// still be surfaced: its outline level reaches `col_outline_levels` even
    /// though no `colWidths` entry is recorded (so its rendered width stays the
    /// workbook default). `collapsed` and `hidden` map likewise.
    #[test]
    fn col_outline_level_recorded_without_custom_width() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols>
                 <col min="2" max="3" outlineLevel="1"/>
                 <col min="4" max="4" outlineLevel="1" collapsed="1"/>
                 <col min="2" max="2" hidden="1" outlineLevel="1"/>
               </cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.col_outline_levels.get(&2).copied(), Some(1));
        assert_eq!(ws.col_outline_levels.get(&3).copied(), Some(1));
        assert_eq!(ws.col_outline_levels.get(&4).copied(), Some(1));
        assert_eq!(ws.col_collapsed.get(&4).copied(), Some(true));
        // The last <col> hides column 2.
        assert_eq!(ws.col_hidden.get(&2).copied(), Some(true));
        // A default-width grouped column gets NO colWidths entry (col 3 was never
        // custom-width nor hidden), so its width stays the workbook default.
        assert_eq!(ws.col_widths.get(&3).copied(), None);
    }

    /// `<sheetPr><outlinePr>` flags parse with the §18.3.1.61 defaults (both
    /// `true`) and honor explicit `false`.
    #[test]
    fn outline_pr_summary_flags() {
        let default_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetPr><outlinePr/></sheetPr><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&default_xml, &[], &[], "Sheet1").expect("parses");
        let pr = ws.outline_pr.expect("outlinePr present");
        assert!(pr.summary_below);
        assert!(pr.summary_right);

        let above_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr><sheetData/></worksheet>"#
        );
        let (ws2, _) = parse_worksheet(&above_xml, &[], &[], "Sheet1").expect("parses");
        let pr2 = ws2.outline_pr.expect("outlinePr present");
        assert!(!pr2.summary_below);
        assert!(!pr2.summary_right);
    }

    /// A sheet with no outlining (no `<outlinePr>`, all `outlineLevel="0"`, as
    /// LibreOffice emits) serializes byte-for-byte as before: no `outlinePr`,
    /// `colOutlineLevels`, `colCollapsed`, `colHidden`, and no per-row
    /// `outlineLevel` / `collapsed` / `hidden` keys.
    #[test]
    fn outline_free_sheet_is_wire_stable() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}">
                 <cols><col customWidth="1" min="1" max="1" width="22" outlineLevel="0" collapsed="0"/></cols>
                 <sheetData>
                   <row r="1" hidden="0" outlineLevel="0" collapsed="0"><c r="A1"/></row>
                 </sheetData>
               </worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parses");
        let v = serde_json::to_value(&ws).unwrap();
        let obj = v.as_object().unwrap();
        assert!(!obj.contains_key("outlinePr"), "no outlinePr key");
        assert!(
            !obj.contains_key("colOutlineLevels"),
            "no colOutlineLevels key"
        );
        assert!(!obj.contains_key("colCollapsed"), "no colCollapsed key");
        assert!(!obj.contains_key("colHidden"), "no colHidden key");
        let row0 = &v["rows"][0];
        let row_obj = row0.as_object().unwrap();
        assert!(
            !row_obj.contains_key("outlineLevel"),
            "no row outlineLevel key"
        );
        assert!(!row_obj.contains_key("collapsed"), "no row collapsed key");
        assert!(!row_obj.contains_key("hidden"), "no row hidden key");
    }
}

#[cfg(test)]
mod hyperlink_tests {
    use super::parse_worksheet;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// ECMA-376 §18.3.1.47: `<hyperlink ref r:id>` is an *external* target
    /// (`r:id` resolved via the sheet rels, populating `url`), while
    /// `<hyperlink ref location>` is an *internal* target captured inline. The
    /// parse step must record the pending `r:id` for the external link and the
    /// inline `location` for the internal one. `parse_worksheet` returns the
    /// pending `(col, row, rid, location, display)` descriptors before rels
    /// resolution; both attributes must be threaded through.
    #[test]
    fn captures_external_rid_and_internal_location() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}" xmlns:r="{R_NS}"><sheetData/>
                 <hyperlinks>
                   <hyperlink ref="A1" r:id="rId1" display="Anthropic"/>
                   <hyperlink ref="B2" location="Sheet2!A1" display="Go to Sheet2"/>
                 </hyperlinks>
               </worksheet>"#
        );
        let (_ws, rids) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");

        // `parse_cell_ref` yields 1-based (col, row): A1 → (1, 1), B2 → (2, 2).
        // A1 → external: rid present (resolves to url later), no location.
        let a1 = rids
            .iter()
            .find(|(c, r, ..)| *c == 1 && *r == 1)
            .expect("A1 hyperlink captured");
        assert_eq!(a1.2.as_deref(), Some("rId1"), "external r:id captured");
        assert_eq!(a1.3, None, "external hyperlink has no inline location");
        assert_eq!(a1.4.as_deref(), Some("Anthropic"), "display captured");

        // B2 → internal: location present, no rid.
        let b2 = rids
            .iter()
            .find(|(c, r, ..)| *c == 2 && *r == 2)
            .expect("B2 hyperlink captured");
        assert_eq!(b2.2, None, "internal hyperlink has no external r:id");
        assert_eq!(
            b2.3.as_deref(),
            Some("Sheet2!A1"),
            "internal location captured"
        );
        assert_eq!(b2.4.as_deref(), Some("Go to Sheet2"), "display captured");
    }

    /// A `<hyperlink>` with neither `r:id` nor `location` is not navigable and
    /// must be skipped (nothing to record).
    #[test]
    fn skips_hyperlink_without_target() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData/>
                 <hyperlinks><hyperlink ref="C3" display="dead"/></hyperlinks>
               </worksheet>"#
        );
        let (_ws, rids) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert!(
            rids.is_empty(),
            "hyperlink with no r:id/location is skipped"
        );
    }
}

#[cfg(test)]
mod resolve_zip_path_tests {
    use super::resolve_zip_path;

    /// A relative Target resolves against the base directory, honoring `..`.
    #[test]
    fn relative_target_resolves_against_base() {
        assert_eq!(
            resolve_zip_path("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_zip_path("xl/drawings", "../media/image1.png"),
            "xl/media/image1.png"
        );
    }

    /// An absolute Target (leading "/", as openpyxl writes for drawings) is
    /// package-root-relative and ignores the base directory.
    #[test]
    fn absolute_target_ignores_base() {
        assert_eq!(
            resolve_zip_path("xl/worksheets", "/xl/drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_zip_path("xl/drawings", "/xl/charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }
}

#[cfg(test)]
mod conditional_format_tests {
    use super::parse_worksheet;
    use crate::types::CfRule;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse_cf_rules(cf_xml: &str) -> Vec<CfRule> {
        let xml = format!(r#"<worksheet xmlns="{NS}"><sheetData/>{cf_xml}</worksheet>"#);
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        ws.conditional_formats
            .into_iter()
            .flat_map(|cf| cf.rules)
            .collect()
    }

    /// ECMA-376 §18.3.1.10: an `aboveAverage` rule with no extra attributes
    /// defaults to `aboveAverage=true`, `equalAverage=false`, no `stdDev`.
    #[test]
    fn above_average_defaults() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage {
                above_average,
                equal_average,
                std_dev,
                ..
            }] => {
                assert!(*above_average, "aboveAverage defaults to true");
                assert!(!*equal_average, "equalAverage defaults to false");
                assert_eq!(*std_dev, None, "no stdDev by default");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }

    /// `aboveAverage="0"` flips to below-average; `equalAverage="1"` is honored.
    #[test]
    fn below_average_with_equal_average() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" aboveAverage="0" equalAverage="1" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage {
                above_average,
                equal_average,
                ..
            }] => {
                assert!(!*above_average, "aboveAverage=\"0\" → false");
                assert!(*equal_average, "equalAverage=\"1\" → true");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }

    /// `stdDev="2"` is captured as a band multiplier (ECMA-376 §18.3.1.10).
    #[test]
    fn above_average_std_dev() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" stdDev="2" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage { std_dev, .. }] => {
                assert_eq!(*std_dev, Some(2), "stdDev=\"2\" captured");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod data_validation_tests {
    use super::parse_worksheet;
    use crate::types::DataValidation;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse_dvs(dv_xml: &str) -> Vec<DataValidation> {
        let xml = format!(r#"<worksheet xmlns="{NS}"><sheetData/>{dv_xml}</worksheet>"#);
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        ws.data_validations
    }

    /// ECMA-376 §18.3.1.33 — a `list` rule captures type, sqref and the
    /// `<formula1>` literal list. `allowBlank="1"` is honored.
    #[test]
    fn list_validation_captures_formula_and_sqref() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation sqref="B2:B5" type="list" allowBlank="1"><formula1>"Pending,Shipped,Delivered"</formula1></dataValidation></dataValidations>"#,
        );
        assert_eq!(dvs.len(), 1, "one rule parsed");
        let dv = &dvs[0];
        assert_eq!(dv.sqref, "B2:B5");
        assert_eq!(dv.validation_type.as_deref(), Some("list"));
        assert_eq!(
            dv.formula1.as_deref(),
            Some("\"Pending,Shipped,Delivered\"")
        );
        assert!(dv.allow_blank, "allowBlank=\"1\" → true");
    }

    /// A `whole`/`between` rule keeps both operands and the operator.
    #[test]
    fn whole_between_keeps_both_operands() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation sqref="C2:C5" type="whole" operator="between"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>"#,
        );
        let dv = &dvs[0];
        assert_eq!(dv.validation_type.as_deref(), Some("whole"));
        assert_eq!(dv.operator.as_deref(), Some("between"));
        assert_eq!(dv.formula1.as_deref(), Some("1"));
        assert_eq!(dv.formula2.as_deref(), Some("100"));
        assert!(!dv.allow_blank, "absent allowBlank → false");
    }

    /// A rule without a `@sqref` is dropped (nothing to anchor it to).
    #[test]
    fn rule_without_sqref_is_skipped() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation type="list"><formula1>"A,B"</formula1></dataValidation></dataValidations>"#,
        );
        assert!(dvs.is_empty(), "missing sqref → rule dropped");
    }

    /// Multiple rules in one block are all captured, preserving order.
    #[test]
    fn multiple_rules_preserved() {
        let dvs = parse_dvs(
            r#"<dataValidations count="2"><dataValidation sqref="B2:B5" type="list"><formula1>"A,B"</formula1></dataValidation><dataValidation sqref="C2:C5" type="whole" operator="between"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>"#,
        );
        assert_eq!(dvs.len(), 2);
        assert_eq!(dvs[0].sqref, "B2:B5");
        assert_eq!(dvs[1].sqref, "C2:C5");
    }

    /// Absent `<dataValidations>` yields an empty vec (no panic).
    #[test]
    fn absent_block_yields_empty() {
        let dvs = parse_dvs("");
        assert!(dvs.is_empty());
    }
}

#[cfg(test)]
mod comment_tests {
    use super::{parse_comments_xml, XlsxCommentKind};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// ECMA-376 §18.7 — each `<comment>` yields its cell ref, the author
    /// resolved from `@authorId`, and the joined `<t>` text.
    #[test]
    fn resolves_ref_author_and_text() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>Reviewer</author><author>Ops Team</author></authors><commentList><comment ref="B1" authorId="0"><text><t>Set the order status.</t></text></comment><comment ref="C3" authorId="1"><text><t>Verify qty.</t></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs.len(), 2);
        assert!(matches!(cs[0].kind, XlsxCommentKind::Note));
        assert_eq!(cs[0].cell_ref, "B1");
        assert_eq!(cs[0].author.as_deref(), Some("Reviewer"));
        assert_eq!(cs[0].text, "Set the order status.");
        assert_eq!(cs[1].cell_ref, "C3");
        assert_eq!(cs[1].author.as_deref(), Some("Ops Team"));
    }

    /// Multiple `<r><t>` runs in one comment are concatenated into plain text.
    #[test]
    fn joins_multiple_runs() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Hello </t></r><r><t>world</t></r></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs[0].text, "Hello world");
    }

    /// An out-of-range or absent `@authorId` leaves the author as None.
    #[test]
    fn missing_author_is_none() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="9"><text><t>orphan</t></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs.len(), 1);
        assert_eq!(cs[0].author, None, "authorId out of range → None");
        assert_eq!(cs[0].text, "orphan");
    }

    /// Malformed XML returns an empty vec instead of panicking.
    #[test]
    fn malformed_xml_yields_empty() {
        assert!(parse_comments_xml("<comments><not closed").is_empty());
    }
}

#[cfg(test)]
mod threaded_comment_tests {
    use super::{
        merge_sheet_comments, parse_comments_xml, parse_sheet_native, parse_threaded_comments_xml,
        to_markdown_native, XlsxCommentKind,
    };
    use std::collections::HashMap;
    use std::io::{Cursor, Write};

    const TC_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    fn persons() -> HashMap<String, String> {
        let mut m = HashMap::new();
        m.insert("{p1}".to_string(), "Reviewer".to_string());
        m.insert("{p2}".to_string(), "Ops Team".to_string());
        m
    }

    fn workbook_with_threaded_comment(
        person_relationship: &str,
        person_parts: &[(&str, &str)],
    ) -> Vec<u8> {
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rSheet1"/></sheets></workbook>"#;
        let workbook_rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>{person_relationship}</Relationships>"#,
        );
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let worksheet_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rThread" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>"#;
        let threaded_comments = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{p1}}" id="thread-1"><text>Review this.</text></threadedComment></ThreadedComments>"#,
        );

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, body) in [
                ("xl/workbook.xml", workbook),
                ("xl/_rels/workbook.xml.rels", workbook_rels.as_str()),
                ("xl/worksheets/sheet1.xml", worksheet),
                ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels),
                (
                    "xl/threadedComments/threadedComment1.xml",
                    threaded_comments.as_str(),
                ),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            for (path, body) in person_parts {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn parsed_thread_author(package: &[u8]) -> Option<String> {
        let json = parse_sheet_native(package, 0, "Sheet1").expect("sheet parses");
        let sheet: serde_json::Value = serde_json::from_str(&json).expect("sheet JSON");
        sheet["comments"][0]["author"].as_str().map(str::to_owned)
    }

    #[test]
    fn resolves_persons_from_unconventional_workbook_relationship_target() {
        let package = workbook_with_threaded_comment(
            r#"<Relationship Id="rPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="reviewers/custom-person-list.xml"/>"#,
            &[(
                "xl/reviewers/custom-person-list.xml",
                r#"<personList><person id="{p1}" displayName="Referenced Reviewer"/></personList>"#,
            )],
        );

        assert_eq!(
            parsed_thread_author(&package).as_deref(),
            Some("Referenced Reviewer")
        );
    }

    #[test]
    fn markdown_collects_cell_comments_after_all_sheet_data() {
        let package = workbook_with_threaded_comment(
            r#"<Relationship Id="rPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="reviewers/custom-person-list.xml"/>"#,
            &[(
                "xl/reviewers/custom-person-list.xml",
                r#"<personList><person id="{p1}" displayName="Referenced Reviewer"/></personList>"#,
            )],
        );

        let markdown = to_markdown_native(&package).expect("markdown projects");

        assert!(
            markdown.find("## Sheet1").unwrap() < markdown.find("## Review comments").unwrap(),
            "{markdown}"
        );
        assert!(markdown.contains("### Sheet1 — A1"), "{markdown}");
        assert!(
            markdown.contains("> **Referenced Reviewer**\n>\n> Review this."),
            "{markdown}"
        );
    }

    #[test]
    fn ignores_unreferenced_conventional_persons_part_poison() {
        let package = workbook_with_threaded_comment(
            r#"<Relationship Id="rPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="reviewers/custom-person-list.xml"/>"#,
            &[
                (
                    "xl/reviewers/custom-person-list.xml",
                    r#"<personList><person id="{p1}" displayName="Referenced Reviewer"/></personList>"#,
                ),
                (
                    "xl/persons/person.xml",
                    r#"<personList><person id="{p1}" displayName="Unreferenced Poison"/></personList>"#,
                ),
            ],
        );

        assert_eq!(
            parsed_thread_author(&package).as_deref(),
            Some("Referenced Reviewer")
        );
    }

    #[test]
    fn external_person_relationship_is_not_treated_as_a_package_part() {
        let package = workbook_with_threaded_comment(
            r#"<Relationship Id="rPersons" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="https://example.invalid/person-list.xml" TargetMode="External"/>"#,
            &[(
                "xl/persons/person.xml",
                r#"<personList><person id="{p1}" displayName="Unreferenced Poison"/></personList>"#,
            )],
        );

        assert_eq!(parsed_thread_author(&package), None);
    }

    /// MS-XLSX top-level comments retain their source identity and metadata.
    #[test]
    fn resolves_ref_person_and_text() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="B1" personId="{{p1}}" id="a"><text>Set the status.</text></threadedComment><threadedComment ref="C3" personId="{{p2}}" id="b"><text>Verify qty.</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &persons());
        assert_eq!(cs.len(), 2);
        assert!(matches!(cs[0].kind, XlsxCommentKind::Thread));
        assert_eq!(cs[0].cell_ref, "B1");
        assert_eq!(cs[0].id.as_deref(), Some("a"));
        assert_eq!(cs[0].person_id.as_deref(), Some("{p1}"));
        assert_eq!(cs[0].author.as_deref(), Some("Reviewer"));
        assert_eq!(cs[0].text, "Set the status.");
        assert_eq!(cs[1].author.as_deref(), Some("Ops Team"));
    }

    /// Structured records are additive: the historical flattened `text` wire
    /// field remains available while `root_text` and `replies` preserve identity.
    #[test]
    fn replies_remain_structured_in_one_thread() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{p1}}" id="a" dT="2026-08-20T09:00:00Z" done="1"><text>Question?</text></threadedComment><threadedComment personId="{{p2}}" id="b" parentId="a" dT="2026-08-20T09:01:00Z"><text>Answer.</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &persons());
        assert_eq!(cs.len(), 1, "one comment per cell");
        assert_eq!(cs[0].cell_ref, "A1");
        assert_eq!(
            cs[0].author.as_deref(),
            Some("Reviewer"),
            "first author kept"
        );
        assert_eq!(cs[0].root_text.as_deref(), Some("Question?"));
        assert_eq!(cs[0].text, "Question?\nAnswer.");
        assert_eq!(cs[0].date.as_deref(), Some("2026-08-20T09:00:00Z"));
        assert_eq!(cs[0].resolved, Some(true));
        assert_eq!(cs[0].replies.len(), 1);
        assert_eq!(cs[0].replies[0].id, "b");
        assert_eq!(cs[0].replies[0].parent_id, "a");
        assert_eq!(cs[0].replies[0].author.as_deref(), Some("Ops Team"));
        assert_eq!(cs[0].replies[0].text, "Answer.");
    }

    /// An unknown `personId` leaves the author as None (no persons part).
    #[test]
    fn unknown_person_is_none() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{zzz}}" id="a"><text>hi</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &HashMap::new());
        assert_eq!(cs[0].author, None);
        assert_eq!(cs[0].text, "hi");
    }

    /// Invalid identity graphs are not repaired by document-order guesses.
    /// A valid independent thread remains available.
    #[test]
    fn ambiguous_or_cyclic_threads_fail_closed() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{p1}}" id="duplicate"><text>first</text></threadedComment><threadedComment ref="B1" personId="{{p2}}" id="duplicate"><text>second</text></threadedComment><threadedComment personId="{{p1}}" id="cycle-a" parentId="cycle-b"><text>a</text></threadedComment><threadedComment personId="{{p2}}" id="cycle-b" parentId="cycle-a"><text>b</text></threadedComment><threadedComment ref="C1" personId="{{p1}}" id="valid"><text>kept</text></threadedComment><threadedComment personId="{{p2}}" id="reply" parentId="valid"><text>reply</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &persons());
        assert_eq!(cs.len(), 1);
        assert_eq!(cs[0].id.as_deref(), Some("valid"));
        assert_eq!(cs[0].replies.len(), 1);
        assert_eq!(cs[0].replies[0].id, "reply");
    }

    /// A compatibility note for a threaded cell is suppressed, while an
    /// independent classic note on the same worksheet remains available.
    #[test]
    fn threaded_and_classic_comments_reconcile_by_authored_identity() {
        const ID: &str = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
        let threaded_xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="Z9" personId="{{p1}}" id="{ID}"><text>structured</text></threadedComment></ThreadedComments>"#
        );
        let classic_xml = format!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><authors><author>tc={ID}</author><author>Legacy</author></authors><commentList><comment ref="A1" authorId="0" x15ac:uid="{ID}"><text><t>compatibility copy</t></text></comment><comment ref="B2" authorId="1"><text><t>real note</t></text></comment></commentList></comments>"#
        );
        let comments = merge_sheet_comments(
            Some(parse_threaded_comments_xml(&threaded_xml, &persons())),
            parse_comments_xml(&classic_xml),
        );
        assert_eq!(comments.len(), 2);
        assert!(matches!(comments[0].kind, XlsxCommentKind::Thread));
        assert_eq!(comments[0].cell_ref, "A1");
        assert!(matches!(comments[1].kind, XlsxCommentKind::Note));
        assert_eq!(comments[1].cell_ref, "B2");
    }

    #[test]
    fn copied_placeholders_create_occurrences_and_orphans_stay_hidden() {
        const ID: &str = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
        const ORPHAN: &str = "{11111111-2222-3333-4444-555555555555}";
        let threaded_xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="Z9" personId="{{p1}}" id="{ID}"><text>structured</text></threadedComment></ThreadedComments>"#
        );
        let classic_xml = format!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><authors><author>tc={ID}</author><author>tc={ORPHAN}</author><author>ordinary</author></authors><commentList><comment ref="A1" authorId="0" x15ac:uid="{ID}"><text><t>copy one</t></text></comment><comment ref="C3" authorId="0" x15ac:uid="{ID}"><text><t>copy two</t></text></comment><comment ref="D4" authorId="1" x15ac:uid="{ORPHAN}"><text><t>orphan</t></text></comment><comment ref="E5" authorId="2" x15ac:uid="{ID}"><text><t>same uid, not a placeholder</t></text></comment></commentList></comments>"#
        );

        let comments = merge_sheet_comments(
            Some(parse_threaded_comments_xml(&threaded_xml, &persons())),
            parse_comments_xml(&classic_xml),
        );
        assert_eq!(
            comments
                .iter()
                .map(|comment| comment.cell_ref.as_str())
                .collect::<Vec<_>>(),
            vec!["A1", "C3", "E5"]
        );
        assert!(matches!(comments[0].kind, XlsxCommentKind::Thread));
        assert!(matches!(comments[1].kind, XlsxCommentKind::Thread));
        assert!(matches!(comments[2].kind, XlsxCommentKind::Note));
    }

    #[test]
    fn empty_threaded_part_hides_only_compatibility_placeholders() {
        const ORPHAN: &str = "{11111111-2222-3333-4444-555555555555}";
        let classic_xml = format!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><authors><author>tc={ORPHAN}</author><author>ordinary</author></authors><commentList><comment ref="A1" authorId="0" x15ac:uid="{ORPHAN}"><text><t>orphan compatibility record</t></text></comment><comment ref="B2" authorId="1"><text><t>real note</t></text></comment></commentList></comments>"#
        );

        let comments = merge_sheet_comments(Some(Vec::new()), parse_comments_xml(&classic_xml));
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell_ref, "B2");
        assert_eq!(comments[0].text, "real note");
    }

    #[test]
    fn invalid_threaded_records_still_hide_compatibility_placeholders() {
        const ORPHAN: &str = "{11111111-2222-3333-4444-555555555555}";
        let threaded_xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{p1}}"><text>missing identity</text></threadedComment></ThreadedComments>"#
        );
        let classic_xml = format!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><authors><author>tc={ORPHAN}</author><author>ordinary</author></authors><commentList><comment ref="A1" authorId="0" x15ac:uid="{ORPHAN}"><text><t>orphan compatibility record</t></text></comment><comment ref="B2" authorId="1"><text><t>real note</t></text></comment></commentList></comments>"#
        );

        let parsed = parse_threaded_comments_xml(&threaded_xml, &persons());
        assert!(parsed.is_empty());
        let comments = merge_sheet_comments(Some(parsed), parse_comments_xml(&classic_xml));
        assert_eq!(
            comments
                .iter()
                .map(|comment| comment.cell_ref.as_str())
                .collect::<Vec<_>>(),
            vec!["B2"]
        );
    }

    #[test]
    fn classic_only_sheet_does_not_apply_threaded_placeholder_rules() {
        const ID: &str = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
        let classic_xml = format!(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><authors><author>tc={ID}</author></authors><commentList><comment ref="A1" authorId="0" x15ac:uid="{ID}"><text><t>classic-only authored record</t></text></comment></commentList></comments>"#
        );

        let comments = merge_sheet_comments(None, parse_comments_xml(&classic_xml));
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].text, "classic-only authored record");
    }
}

#[cfg(test)]
mod extract_image_tests {
    use super::extract_image;

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("xl/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(
            extract_image(&buf, "xl/media/i.png", None, None).unwrap(),
            b"X"
        );
    }
}

#[cfg(test)]
mod sheet_visibility_tests {
    use super::*;

    #[test]
    fn sheet_state_attr_maps_to_visibility() {
        // ECMA-376 §18.2.19 ST_SheetState: visible (default) | hidden | veryHidden.
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2" state="hidden"/><sheet name="C" sheetId="3" r:id="rId3" state="veryHidden"/><sheet name="D" sheetId="4" r:id="rId4" state="visible"/></sheets></workbook>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sheets = parse_workbook_sheets(&doc);
        assert_eq!(sheets[0].visibility, SheetVisibility::Visible); // absent ⇒ visible
        assert_eq!(sheets[1].visibility, SheetVisibility::Hidden);
        assert_eq!(sheets[2].visibility, SheetVisibility::VeryHidden);
        assert_eq!(sheets[3].visibility, SheetVisibility::Visible);
    }
}

#[cfg(test)]
mod workbook_theme_tests {
    use super::*;
    use std::io::{Cursor, Write};

    #[test]
    fn loads_the_theme_target_declared_by_workbook_relationships_once() {
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="themes/custom.xml"/></Relationships>"#;
        let custom = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="custom"><a:dk1><a:srgbClr val="010203"/></a:dk1></a:clrScheme><a:fontScheme name="custom"><a:majorFont><a:latin typeface="Major Custom"/></a:majorFont><a:minorFont><a:latin typeface="Minor Custom"/></a:minorFont></a:fontScheme><a:fmtScheme name="custom"><a:fillStyleLst><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#;
        let decoy = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="decoy"><a:dk1><a:srgbClr val="FFFFFF"/></a:dk1></a:clrScheme></a:themeElements></a:theme>"#;
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, body) in [
                ("xl/themes/custom.xml", custom),
                ("xl/theme/theme1.xml", decoy),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }
        let mut archive = XlsxZip::new(Cursor::new(bytes)).unwrap();
        archive.begin_operation("theme-test").unwrap();
        let theme = XlsxThemeData::load(&mut archive, rels);

        assert_eq!(theme.colors.first().map(String::as_str), Some("#010203"));
        assert_eq!(theme.fonts.0.as_deref(), Some("Major Custom"));
        assert_eq!(theme.fonts.1.as_deref(), Some("Minor Custom"));
        assert!(matches!(
            theme.format_scheme.lookup_fill_ref(1),
            ooxml_common::theme::StyleMatrixLookup::Entry(_)
        ));
    }

    #[test]
    fn external_theme_relationship_is_not_treated_as_a_package_part() {
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="https://example.invalid/theme.xml" TargetMode="External"/></Relationships>"#;
        assert_eq!(find_internal_rel_target_by_type(rels, "/theme"), None);
    }
}

#[cfg(test)]
mod date1904_tests {
    use super::*;

    fn parse(xml: &str) -> bool {
        let doc = roxmltree::Document::parse(xml).unwrap();
        parse_workbook_date1904(&doc)
    }

    #[test]
    fn workbook_pr_date1904_true() {
        // ECMA-376 §18.2.28: date1904="1" ⇒ 1904 date system.
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr date1904="1"/></workbook>"#;
        assert!(parse(xml));
    }

    #[test]
    fn workbook_pr_date1904_true_word() {
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr date1904="true"/></workbook>"#;
        assert!(parse(xml));
    }

    #[test]
    fn workbook_pr_date1904_false() {
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr date1904="0"/></workbook>"#;
        assert!(!parse(xml));
    }

    #[test]
    fn workbook_pr_absent_attr_defaults_false() {
        // §18.2.28: absent attribute ⇒ 1900 date system (false).
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><workbookPr showObjects="all"/></workbook>"#;
        assert!(!parse(xml));
    }

    #[test]
    fn workbook_pr_absent_element_defaults_false() {
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"#;
        assert!(!parse(xml));
    }
}

#[cfg(test)]
mod chartsheet_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    fn chartsheet_bytes() -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = SimpleFileOptions::default();
            let files = [
                (
                    "xl/workbook.xml",
                    r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Map" sheetId="1" r:id="rSheet"/></sheets></workbook>"#,
                ),
                (
                    "xl/_rels/workbook.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#,
                ),
                (
                    "xl/chartsheets/sheet1.xml",
                    r#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rDrawing"/></chartsheet>"#,
                ),
                (
                    "xl/chartsheets/_rels/sheet1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
                ),
                (
                    "xl/drawings/drawing1.xml",
                    r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="5334000" cy="3302000"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Map"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5334000" cy="3302000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#,
                ),
                (
                    "xl/drawings/_rels/drawing1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
                ),
                (
                    "xl/charts/chart1.xml",
                    r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
                ),
            ];
            for (path, content) in files {
                zip.start_file(path, options).unwrap();
                zip.write_all(content.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        bytes
    }

    fn chartsheet_archive() -> XlsxZip {
        XlsxZip::new(Cursor::new(chartsheet_bytes())).expect("chartsheet zip opens")
    }

    /// A chartsheet has no CT_Worksheet/sheetData. It must still finalize its
    /// drawing relationships and expose the absolute-anchored chart instead of
    /// degrading to an empty worksheet parse-error placeholder.
    #[test]
    fn chart_sheet_without_sheet_data_reaches_its_absolute_chart() {
        let mut archive = chartsheet_archive();
        let shared = WorkbookShared::load(&mut archive).expect("shared workbook parts");
        let bytes =
            parse_sheet_with(&mut archive, &shared, 0, "Map").expect("chartsheet materializes");
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_eq!(
            value.get("isChartSheet").and_then(|v| v.as_bool()),
            Some(true)
        );
        assert!(value.get("parseError").is_none());
        assert_eq!(
            value.get("charts").and_then(|v| v.as_array()).map(Vec::len),
            Some(1),
        );
    }

    /// The browser viewer opens worksheets through the resumable cursor API,
    /// not `parse_sheet_with`. A chartsheet is already fully materialized after
    /// its drawing relationships are resolved, so the first pull must return a
    /// terminal worksheet model instead of trying to stream CT_Worksheet rows.
    #[test]
    fn chart_sheet_production_cursor_returns_terminal_model() {
        let mut archive = XlsxArchive::new(chartsheet_bytes(), None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Map").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_eq!(value.get("kind").and_then(|v| v.as_str()), Some("finished"));
        let worksheet = value.get("worksheet").expect("terminal worksheet model");
        assert_eq!(
            worksheet.get("isChartSheet").and_then(|v| v.as_bool()),
            Some(true)
        );
        assert!(worksheet.get("parseError").is_none());
        assert_eq!(
            worksheet
                .get("charts")
                .and_then(|v| v.as_array())
                .map(Vec::len),
            Some(1),
        );
    }
}

#[cfg(test)]
mod dialogsheet_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    fn dialogsheet_bytes(relationship_base: &str, sheet_xml: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = SimpleFileOptions::default();
            let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Dialog" sheetId="1" r:id="rSheet"/></sheets></workbook>"#;
            let relationships = format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet" Type="{relationship_base}/dialogsheet" Target="dialogsheets/sheet1.xml"/></Relationships>"#,
            );
            for (path, content) in [
                ("xl/workbook.xml", workbook),
                ("xl/_rels/workbook.xml.rels", relationships.as_str()),
                ("xl/dialogsheets/sheet1.xml", sheet_xml),
            ] {
                zip.start_file(path, options).unwrap();
                zip.write_all(content.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        bytes
    }

    fn valid_dialogsheet_xml() -> &'static str {
        r#"<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></dialogsheet>"#
    }

    fn assert_dialogsheet_model(worksheet: &serde_json::Value) {
        assert_eq!(
            worksheet
                .get("isDialogSheet")
                .and_then(|value| value.as_bool()),
            Some(true)
        );
        assert_eq!(
            worksheet
                .get("rows")
                .and_then(|value| value.as_array())
                .map(Vec::len),
            Some(0)
        );
        assert!(
            worksheet.get("parseError").is_none(),
            "a valid dialogsheet is not a broken worksheet: {worksheet}"
        );
    }

    /// ECMA-376 Part 1 §12.3.7 / §18.3.1.34 defines a Dialogsheet as a
    /// distinct Workbook target whose root is `dialogsheet`, not `worksheet`.
    #[test]
    fn dialog_sheet_is_a_normal_row_free_model() {
        let mut archive = XlsxZip::new(Cursor::new(dialogsheet_bytes(
            TRANSITIONAL_RELATIONSHIPS,
            valid_dialogsheet_xml(),
        )))
        .expect("dialogsheet zip opens");
        let shared = WorkbookShared::load(&mut archive).expect("shared workbook parts");
        let bytes =
            parse_sheet_with(&mut archive, &shared, 0, "Dialog").expect("dialogsheet materializes");
        let worksheet: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_dialogsheet_model(&worksheet);
    }

    /// The browser viewer uses the resumable cursor path. It must produce the
    /// same non-error terminal model as the monolithic/native path.
    #[test]
    fn dialog_sheet_production_cursor_returns_terminal_model() {
        let mut archive = XlsxArchive::new(
            dialogsheet_bytes(TRANSITIONAL_RELATIONSHIPS, valid_dialogsheet_xml()),
            None,
            None,
            None,
        )
        .unwrap();
        archive.open_sheet_cursor(0, "Dialog").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_eq!(
            value.get("kind").and_then(|item| item.as_str()),
            Some("finished")
        );
        assert_dialogsheet_model(value.get("worksheet").expect("terminal worksheet model"));
    }

    /// Strict packages use the purl relationship base but the same part/root
    /// contract. Exact recognition must cover both conformance classes.
    #[test]
    fn strict_dialog_sheet_relationship_is_recognized() {
        let mut archive = XlsxArchive::new(
            dialogsheet_bytes(STRICT_RELATIONSHIPS, valid_dialogsheet_xml()),
            None,
            None,
            None,
        )
        .unwrap();
        archive.open_sheet_cursor(0, "Dialog").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        assert_dialogsheet_model(value.get("worksheet").expect("terminal worksheet model"));
    }

    /// Relationship type determines the part kind, but the corresponding root
    /// still has to satisfy the Dialogsheet host schema.
    #[test]
    fn dialog_sheet_relationship_with_wrong_root_remains_a_parse_error() {
        let bytes = dialogsheet_bytes(
            TRANSITIONAL_RELATIONSHIPS,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
        );
        let mut archive = XlsxArchive::new(bytes, None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Dialog").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        let worksheet = value.get("worksheet").expect("terminal worksheet model");
        assert!(worksheet.get("isDialogSheet").is_none());
        assert!(worksheet["parseError"]
            .as_str()
            .is_some_and(|error| error.contains("expected SpreadsheetML dialogsheet root")));
    }

    #[test]
    fn dialog_sheet_root_in_a_foreign_namespace_remains_a_parse_error() {
        let bytes = dialogsheet_bytes(
            TRANSITIONAL_RELATIONSHIPS,
            r#"<dialogsheet xmlns="urn:foreign"><sheetViews/></dialogsheet>"#,
        );
        let mut archive = XlsxArchive::new(bytes, None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Dialog").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        let worksheet = value.get("worksheet").expect("terminal worksheet model");
        assert!(worksheet.get("isDialogSheet").is_none());
        assert!(worksheet["parseError"]
            .as_str()
            .is_some_and(|error| error.contains("expected SpreadsheetML dialogsheet root")));
    }

    #[test]
    fn foreign_relationship_suffix_is_not_a_dialog_sheet() {
        let bytes = dialogsheet_bytes("urn:foreign", valid_dialogsheet_xml());
        let mut archive = XlsxArchive::new(bytes, None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Dialog").unwrap();

        let bytes = archive.pull_sheet_cursor_inner(1).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
        let worksheet = value.get("worksheet").expect("terminal worksheet model");
        assert!(worksheet.get("isDialogSheet").is_none());
        assert!(worksheet["parseError"]
            .as_str()
            .is_some_and(|error| error.contains("MCE-processed worksheet root")));
    }

    #[test]
    fn strict_chart_sheet_relationship_keeps_the_chart_part_kind() {
        let relationships = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet" Type="{STRICT_RELATIONSHIPS}/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>"#,
        );
        let document = parse_guarded(&relationships).unwrap();
        assert_eq!(
            resolve_sheet_part_kind(&document, "rSheet"),
            SheetPartKind::ChartSheet
        );
    }
}

#[cfg(test)]
mod date1904_wire_shape_tests {
    // Wire-parity guard for the `date1904` field on `Workbook` / `Worksheet`:
    // it must be dropped from the JSON when false (default 1900 system, keeps
    // existing snapshots byte-stable) and present when true. Mirrors the
    // `chart_model_serializes_canonical_shape` approach in ooxml-common.
    use super::*;

    fn workbook(date1904: bool) -> Workbook {
        Workbook {
            date1904,
            ..Default::default()
        }
    }

    fn worksheet(date1904: bool) -> Worksheet {
        // Parse a minimal sheet so every non-date1904 field is default-populated
        // (robust to future field additions), then set the flag under test.
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let (mut ws, _) = parse_worksheet(xml, &[], &[], "Sheet1").expect("worksheet parses");
        ws.date1904 = date1904;
        ws
    }

    #[test]
    fn workbook_date1904_false_is_omitted_from_wire() {
        let v = serde_json::to_value(workbook(false)).unwrap();
        let obj = v.as_object().unwrap();
        assert!(!obj.contains_key("date1904"));
    }

    #[test]
    fn workbook_date1904_true_is_serialized() {
        let v = serde_json::to_value(workbook(true)).unwrap();
        assert_eq!(v.get("date1904").and_then(|d| d.as_bool()), Some(true));
    }

    #[test]
    fn worksheet_date1904_false_is_omitted_from_wire() {
        let v = serde_json::to_value(worksheet(false)).unwrap();
        let obj = v.as_object().unwrap();
        assert!(!obj.contains_key("date1904"));
    }

    #[test]
    fn worksheet_date1904_true_is_serialized() {
        let v = serde_json::to_value(worksheet(true)).unwrap();
        assert_eq!(v.get("date1904").and_then(|d| d.as_bool()), Some(true));
    }
}

/// ISO/IEC 29500 Strict-conformance fixture (`fix(xlsx): accept Strict
/// namespace URIs across the parser` routed `parse_row_cells`'s `<c>`/`<v>`
/// element matching through `is_x_ns`). Before that conversion every
/// `<row>`/`<c>`/`<v>` lookup was pinned to the Transitional `x:` URI, so a
/// Strict worksheet — `xmlns="http://purl.oclc.org/ooxml/spreadsheetml/
/// main"` — parsed to zero rows; this pins that cell values (shared-string
/// text, an inline string, and a numeric literal) and each cell's `s` style
/// index resolve identically to the Transitional case.
#[cfg(test)]
mod strict_namespace_cell_tests {
    use super::*;

    const X_NS_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    #[test]
    fn strict_worksheet_resolves_cell_values_and_style_index() {
        let shared = vec![SharedString {
            text: "Shared Hello".to_string(),
            runs: None,
            ..Default::default()
        }];
        let xml = format!(
            r#"<worksheet xmlns="{ns}">
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="2"><v>0</v></c>
      <c r="B1" t="inlineStr"><is><t>Inline Hi</t></is></c>
      <c r="C1"><v>42.5</v></c>
    </row>
  </sheetData>
</worksheet>"#,
            ns = X_NS_STRICT,
        );

        let (ws, _) =
            parse_worksheet(&xml, &shared, &[], "Sheet1").expect("Strict worksheet must parse");
        assert_eq!(ws.rows.len(), 1, "Strict <row> must be found via is_x_ns");
        let cells = &ws.rows[0].cells;
        assert_eq!(cells.len(), 3, "Strict <c> must be found via is_x_ns");

        // The wire now ships an `si` reference for `t="s"`; the text
        // ("Shared Hello") resolves consumer-side from `shared[0]`.
        match &cells[0].value {
            CellValue::Shared { si } => assert_eq!(*si, 0),
            other => panic!("expected shared-string reference, got {other:?}"),
        }
        assert_eq!(
            cells[0].style_index,
            Some(2),
            "the `s` style index must round-trip"
        );

        match &cells[1].value {
            CellValue::Text { text, .. } => assert_eq!(text, "Inline Hi"),
            other => panic!("expected inline string text, got {other:?}"),
        }

        match &cells[2].value {
            CellValue::Number { number } => assert_eq!(*number, 42.5),
            other => panic!("expected a number, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod phonetic_tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// ECMA-376 §18.4.6 / §18.4.3: a `<si>` with `<rPh>` runs and a
    /// `<phoneticPr>` must parse the furigana runs (sb/eb + hint text) and the
    /// display properties, while `text` stays the base string only.
    #[test]
    fn parse_si_node_reads_rph_and_phonetic_pr() {
        let xml = format!(
            r#"<si xmlns="{ns}"><t>課長</t><rPh sb="0" eb="1"><t>カ</t></rPh><rPh sb="1" eb="2"><t>チョウ</t></rPh><phoneticPr fontId="2" type="Hiragana" alignment="center"/></si>"#,
            ns = NS,
        );
        let doc = roxmltree::Document::parse(&xml).expect("parse");
        let ss = parse_si_node(&doc.root_element(), &[]);
        assert_eq!(ss.text, "課長", "base text excludes the furigana");
        assert_eq!(ss.phonetic_runs.len(), 2, "two rPh runs");
        assert_eq!(ss.phonetic_runs[0].sb, 0);
        assert_eq!(ss.phonetic_runs[0].eb, 1);
        assert_eq!(ss.phonetic_runs[0].text, "カ");
        assert_eq!(ss.phonetic_runs[1].sb, 1);
        assert_eq!(ss.phonetic_runs[1].eb, 2);
        assert_eq!(ss.phonetic_runs[1].text, "チョウ");
        let pr = ss.phonetic_pr.expect("phoneticPr present");
        assert_eq!(pr.font_id, 2);
        assert_eq!(pr.r#type.as_deref(), Some("Hiragana"));
        assert_eq!(pr.alignment.as_deref(), Some("center"));
    }

    /// A `<phoneticPr>` with only the required `fontId` leaves `type` /
    /// `alignment` absent so the consumer applies the schema defaults
    /// (fullwidthKatakana / left) rather than a wrong hard-coded value.
    #[test]
    fn phonetic_pr_omits_optional_attrs_when_absent() {
        let xml = format!(
            r#"<si xmlns="{ns}"><t>山</t><rPh sb="0" eb="1"><t>ヤマ</t></rPh><phoneticPr fontId="1"/></si>"#,
            ns = NS,
        );
        let doc = roxmltree::Document::parse(&xml).expect("parse");
        let ss = parse_si_node(&doc.root_element(), &[]);
        let pr = ss.phonetic_pr.expect("phoneticPr present");
        assert_eq!(pr.font_id, 1);
        assert!(pr.r#type.is_none(), "type absent → consumer defaults");
        assert!(
            pr.alignment.is_none(),
            "alignment absent → consumer defaults"
        );
    }

    /// A `<si>` with NO phonetic markup yields empty phonetic_runs and no
    /// phonetic_pr, so non-Japanese workbooks stay byte-identical on the wire.
    #[test]
    fn plain_si_has_no_phonetic_data() {
        let xml = format!(r#"<si xmlns="{ns}"><t>Hello</t></si>"#, ns = NS);
        let doc = roxmltree::Document::parse(&xml).expect("parse");
        let ss = parse_si_node(&doc.root_element(), &[]);
        assert!(ss.phonetic_runs.is_empty());
        assert!(ss.phonetic_pr.is_none());
    }

    /// ECMA-376 §18.3.1.4 `<c ph="1">` sets the cell's show_phonetic flag;
    /// a cell without `ph` (or with `ph="0"`) stays false (schema default).
    #[test]
    fn cell_ph_attribute_drives_show_phonetic() {
        let xml = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1">
              <c r="A1" t="s" ph="1"><v>0</v></c>
              <c r="B1" t="s" ph="0"><v>0</v></c>
              <c r="C1" t="s"><v>0</v></c>
            </row></sheetData></worksheet>"#,
            ns = NS,
        );
        let shared = vec![SharedString {
            text: "課長".to_string(),
            ..Default::default()
        }];
        let (ws, _) = parse_worksheet(&xml, &shared, &[], "Sheet1").expect("parse");
        let cells = &ws.rows[0].cells;
        assert!(cells[0].show_phonetic, "ph=1 → show");
        assert!(!cells[1].show_phonetic, "ph=0 → hide");
        assert!(
            !cells[2].show_phonetic,
            "no ph → hide (schema default false)"
        );
    }

    /// ECMA-376 §18.3.1.73 `<row ph="1">` turns on furigana display for every
    /// cell in the row, resolved as `cell/@ph ?? row/@ph ?? false`: a cell
    /// without its own `ph` inherits the row flag, while a cell that sets
    /// `ph="0"` explicitly overrides the row back to hidden.
    #[test]
    fn row_ph_attribute_drives_show_phonetic_with_cell_override() {
        let xml = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1" ph="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1" t="s" ph="0"><v>0</v></c>
              <c r="C1" t="s" ph="1"><v>0</v></c>
            </row></sheetData></worksheet>"#,
            ns = NS,
        );
        let shared = vec![SharedString {
            text: "課長".to_string(),
            ..Default::default()
        }];
        let (ws, _) = parse_worksheet(&xml, &shared, &[], "Sheet1").expect("parse");
        let cells = &ws.rows[0].cells;
        assert!(cells[0].show_phonetic, "no cell ph → inherits row ph=1");
        assert!(
            !cells[1].show_phonetic,
            "cell ph=0 overrides row ph=1 → hide"
        );
        assert!(cells[2].show_phonetic, "cell ph=1 agrees with row ph=1");
    }

    /// A row WITHOUT `ph` keeps the schema default (false) for its cells, so a
    /// non-Japanese sheet stays byte-identical. A cell may still opt in per-cell.
    #[test]
    fn row_without_ph_leaves_cells_at_schema_default() {
        let xml = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1" t="s" ph="1"><v>0</v></c>
            </row></sheetData></worksheet>"#,
            ns = NS,
        );
        let shared = vec![SharedString {
            text: "課長".to_string(),
            ..Default::default()
        }];
        let (ws, _) = parse_worksheet(&xml, &shared, &[], "Sheet1").expect("parse");
        let cells = &ws.rows[0].cells;
        assert!(!cells[0].show_phonetic, "no row/cell ph → hide (default)");
        assert!(cells[1].show_phonetic, "cell ph=1 still opts in");
    }

    /// Worker-boundary contract: the RESOLVED `show_phonetic` crosses the wire
    /// as `showPhonetic` (camelCase). A row-inherited `true` serializes the field
    /// so the TS renderer gate (`cell.showPhonetic`) sees it, while a cell that
    /// resolved to `false` (the `ph="0"` override) is omitted (serde skips false)
    /// and reads back as `showPhonetic ?? false`. No new row-level field is added
    /// to the JSON — resolving at parse time keeps the boundary schema stable.
    #[test]
    fn resolved_show_phonetic_serializes_to_json_boundary() {
        let xml = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1" ph="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1" t="s" ph="0"><v>0</v></c>
            </row></sheetData></worksheet>"#,
            ns = NS,
        );
        let shared = vec![SharedString {
            text: "課長".to_string(),
            ..Default::default()
        }];
        let (ws, _) = parse_worksheet(&xml, &shared, &[], "Sheet1").expect("parse");
        let json = serde_json::to_value(&ws.rows[0].cells).expect("serialize cells");
        assert_eq!(
            json[0].get("showPhonetic"),
            Some(&serde_json::Value::Bool(true)),
            "row-inherited cell serializes showPhonetic:true"
        );
        assert!(
            json[1].get("showPhonetic").is_none(),
            "override cell (resolved false) omits showPhonetic — reads back as ?? false"
        );
    }

    /// An inline string (`t="inlineStr"`) carries its own `<rPh>` runs straight
    /// onto the resolved `CellValue::Text` (no shared-string indirection).
    #[test]
    fn inline_string_carries_phonetic_runs() {
        let xml = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1">
              <c r="A1" t="inlineStr" ph="1"><is><t>森</t><rPh sb="0" eb="1"><t>モリ</t></rPh><phoneticPr fontId="1"/></is></c>
            </row></sheetData></worksheet>"#,
            ns = NS,
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        let cell = &ws.rows[0].cells[0];
        assert!(cell.show_phonetic);
        match &cell.value {
            CellValue::Text {
                text,
                phonetic_runs,
                phonetic_pr,
                ..
            } => {
                assert_eq!(text, "森");
                assert_eq!(phonetic_runs.len(), 1);
                assert_eq!(phonetic_runs[0].text, "モリ");
                assert_eq!(phonetic_pr.as_ref().expect("pr").font_id, 1);
            }
            other => panic!("expected text cell, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod sparkline_range_cap_tests {
    use super::*;

    /// A malicious `<xm:f>` referencing a whole-sheet range (`A1:XFD1048576`,
    /// 16384 × 1048576 ≈ 1.7e10 cells → ~275 GB of `Vec<Option<f64>>`) must NOT
    /// attempt the dense allocation. The cap fires and `extract_range_values`
    /// returns an empty `Vec`, so the sparkline simply is not drawn (the
    /// downstream renderer iterates `values` by index, so an empty slice draws
    /// nothing — no panic, no OOM).
    #[test]
    fn oversized_full_sheet_range_returns_empty_without_allocating() {
        // parse_cell_ref("A1") = (col 1, row 1); ("XFD1048576") = (col 16384,
        // row 1048576). This is Excel's entire grid — the worst case an
        // attacker can express.
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1_048_576,
            right: 16_384,
        };
        // Minimal well-formed worksheet with a single real value inside the
        // range. Before the fix, building `vec![None; 1.7e10]` OOMs/aborts here.
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>3.5</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert!(
            values.is_empty(),
            "an over-cap sparkline range must yield an empty Vec (no dense alloc), got len {}",
            values.len()
        );
    }

    /// A normal small sparkline range (a handful of cells) must still resolve
    /// its numeric values in row-major order, unaffected by the cap.
    #[test]
    fn normal_small_range_still_resolves_values() {
        // B2:B4 — three cells in one column.
        let range = CellRange {
            top: 2,
            left: 2,
            bottom: 4,
            right: 2,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2"><v>10</v></c></row><row r="3"><c r="B3"><v>20</v></c></row><row r="4"><c r="B4" t="s"><v>0</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert_eq!(values.len(), 3, "3-cell range must yield 3 slots");
        assert_eq!(values[0], Some(10.0));
        assert_eq!(values[1], Some(20.0));
        assert_eq!(values[2], None, "string cell (t=s) must map to None");
    }

    /// A range exactly at the cap must still allocate; one cell over must not.
    /// Guards the boundary condition of `MAX_SPARKLINE_CELLS`.
    #[test]
    fn range_at_cap_allocates_over_cap_does_not() {
        // 1000 columns × 1000 rows = 1_000_000 cells = exactly MAX_SPARKLINE_CELLS.
        let at_cap = CellRange {
            top: 1,
            left: 1,
            bottom: 1000,
            right: 1000,
        };
        let empty_xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let at = extract_range_values(empty_xml, &at_cap);
        assert_eq!(
            at.len(),
            MAX_SPARKLINE_CELLS,
            "a range exactly at the cap must allocate all slots"
        );

        // One column wider → 1001 × 1000 = 1_001_000 > cap → empty.
        let over_cap = CellRange {
            top: 1,
            left: 1,
            bottom: 1000,
            right: 1001,
        };
        assert!(
            extract_range_values(empty_xml, &over_cap).is_empty(),
            "a range one cell over the cap must yield empty"
        );
    }

    // ── #851 mirror: implicit cell/row references in sparkline data ranges ─────
    //
    // ECMA-376 marks `@r` `use="optional"` on both `CT_Cell` (§18.3.1.4) and
    // `CT_Row` (§18.3.1.73). PR #851 taught the *main* cell path
    // (`parse_worksheet` / `parse_row_cells`) to resolve omitted references by
    // ordinal document order, but `extract_range_values` (the sparkline data
    // path) still skipped any `<c>` without `@r`, so a sparkline whose source
    // worksheet uses the minimal r-less form rendered blank. These tests pin the
    // mirror resolution: r-less `<row>` = previous row + 1 (first = 1); r-less
    // `<c>` = previous cell's column + 1 within that row (first = column A); an
    // explicit `@r` on either re-anchors the running counter.

    /// A worksheet whose `<row>` and `<c>` both omit `@r` entirely must still
    /// resolve numeric values in row-major order — the counters supply A1, A2, …
    #[test]
    fn all_implicit_refs_resolve_row_major() {
        // Column A, rows 1..=3 — all implicit. Range A1:A3.
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 3,
            right: 1,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c><v>10</v></c></row><row><c><v>20</v></c></row><row><c><v>30</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert_eq!(
            values,
            vec![Some(10.0), Some(20.0), Some(30.0)],
            "all-implicit row/cell refs must resolve to A1,A2,A3 in row-major order"
        );
    }

    /// A single implicit row with several implicit cells must fill columns
    /// A,B,C,… left-to-right off the running per-row column counter.
    #[test]
    fn implicit_cells_fill_columns_left_to_right() {
        // Row 1, columns A..=C — all implicit. Range A1:C1.
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 3,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert_eq!(
            values,
            vec![Some(1.0), Some(2.0), Some(3.0)],
            "consecutive implicit cells must land in A1,B1,C1"
        );
    }

    /// An explicit `@r` on a `<c>` must re-anchor the running column so a
    /// following implicit cell continues from the explicit anchor, and an
    /// explicit `<row r>` must re-anchor the row counter for later implicit rows.
    #[test]
    fn explicit_ref_reanchors_running_counters() {
        // Range spans A1:D2. Row 1: implicit A1=5, then explicit C1=7, then
        // implicit D1=8 (continues from C). Row 2 is implicit (previous row 1 +
        // 1 = 2): implicit A2=9.
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 2,
            right: 4,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c><v>5</v></c><c r="C1"><v>7</v></c><c><v>8</v></c></row><row><c><v>9</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        // Row-major over A1:D2, row_span = 4:
        //   idx0 A1=5, idx1 B1=None, idx2 C1=7, idx3 D1=8,
        //   idx4 A2=9, idx5..7 None
        assert_eq!(
            values,
            vec![
                Some(5.0),
                None,
                Some(7.0),
                Some(8.0),
                Some(9.0),
                None,
                None,
                None,
            ],
            "explicit @r on a cell re-anchors the column; implicit row after r=1 is row 2"
        );
    }

    /// An explicit `<row r>` re-anchors the implicit row counter: a later r-less
    /// row is the explicit row + 1, not a naive +1 off document order.
    #[test]
    fn explicit_row_ref_reanchors_row_counter() {
        // Explicit row 5, then an implicit row (→ row 6). Range A5:A6.
        let range = CellRange {
            top: 5,
            left: 1,
            bottom: 6,
            right: 1,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="5"><c><v>50</v></c></row><row><c><v>60</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert_eq!(
            values,
            vec![Some(50.0), Some(60.0)],
            "implicit row after explicit r=5 must resolve to row 6"
        );
    }

    /// Implicit refs must coexist with the existing type filter: a string cell
    /// (`t="s"`) in the running sequence still maps to None while numeric
    /// neighbors resolve, and the column counter advances past it.
    #[test]
    fn implicit_cells_respect_type_filter() {
        // Row 1, A1=1 (implicit), B1 t="s" (implicit, → None), C1=3 (implicit).
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 3,
        };
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c><v>1</v></c><c t="s"><v>0</v></c><c><v>3</v></c></row></sheetData></worksheet>"#;
        let values = extract_range_values(xml, &range);
        assert_eq!(
            values,
            vec![Some(1.0), None, Some(3.0)],
            "a t=s cell in an implicit run maps to None but still advances the column counter"
        );
    }
}

#[cfg(test)]
mod reversed_range_normalization_tests {
    use super::*;

    /// A malicious (or merely hand-typed) `<xm:f>` reversed-row range like
    /// `A10:A1` must normalize to `top=1, bottom=10` — matching Excel's own
    /// interpretation of a backwards-typed range — rather than leaving
    /// `bottom < top`. Before the fix, `parse_sqref` copied the corners
    /// verbatim and `extract_range_values`'s `bottom - top` underflowed
    /// (`u32` subtraction), wrapping silently in release and panicking
    /// (`should_panic`-worthy, exit 101) in debug/test builds.
    #[test]
    fn parse_sqref_normalizes_reversed_row_range() {
        let ranges = parse_sqref("A10:A1");
        assert_eq!(ranges.len(), 1);
        let r = &ranges[0];
        assert_eq!(r.top, 1, "top must be the smaller row");
        assert_eq!(r.bottom, 10, "bottom must be the larger row");
        assert_eq!(r.left, 1);
        assert_eq!(r.right, 1);
    }

    /// Same as above but for a reversed COLUMN range (`B1:A1`): `left`/`right`
    /// must normalize independently of `top`/`bottom`.
    #[test]
    fn parse_sqref_normalizes_reversed_column_range() {
        let ranges = parse_sqref("B1:A1");
        assert_eq!(ranges.len(), 1);
        let r = &ranges[0];
        assert_eq!(r.left, 1, "left must be the smaller column");
        assert_eq!(r.right, 2, "right must be the larger column");
        assert_eq!(r.top, 1);
        assert_eq!(r.bottom, 1);
    }

    /// A range reversed on BOTH axes (`B10:A1`) normalizes on both.
    #[test]
    fn parse_sqref_normalizes_reversed_both_axes() {
        let ranges = parse_sqref("B10:A1");
        let r = &ranges[0];
        assert_eq!((r.top, r.bottom, r.left, r.right), (1, 10, 1, 2));
    }

    /// A normal, already-ordered range is unaffected (identical to what
    /// `parse_sqref` produced before this fix).
    #[test]
    fn parse_sqref_leaves_ordered_range_unchanged() {
        let ranges = parse_sqref("A1:A10");
        let r = &ranges[0];
        assert_eq!((r.top, r.bottom, r.left, r.right), (1, 10, 1, 1));
    }

    /// End-to-end: a reversed-row sparkline data range must not panic and
    /// must resolve the SAME cell values as the equivalent ordered range —
    /// proving normalization, not just crash-avoidance. Guards against a
    /// regression to the pre-fix unsigned-subtraction underflow in
    /// `extract_range_values` (`(range.bottom - range.top + 1)`), which
    /// panics in debug/test builds (`overflow-checks` on) and silently wraps
    /// in release, in both cases producing wrong/no data instead of the
    /// correct cell set.
    #[test]
    fn reversed_range_resolves_same_values_as_ordered_range() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#;

        let ordered = parse_sqref("A1:A3");
        let reversed = parse_sqref("A3:A1");
        assert_eq!(ordered.len(), 1);
        assert_eq!(reversed.len(), 1);
        assert_eq!(
            (
                ordered[0].top,
                ordered[0].bottom,
                ordered[0].left,
                ordered[0].right
            ),
            (
                reversed[0].top,
                reversed[0].bottom,
                reversed[0].left,
                reversed[0].right
            ),
            "reversed and ordered refs to the same cells must normalize identically"
        );

        let ordered_values = extract_range_values(xml, &ordered[0]);
        let reversed_values = extract_range_values(xml, &reversed[0]);
        assert_eq!(
            ordered_values,
            vec![Some(1.0), Some(2.0), Some(3.0)],
            "ordered range resolves row-major"
        );
        assert_eq!(
            reversed_values, ordered_values,
            "a reversed-row range must resolve to the identical cell values, not empty/wrong"
        );
    }

    /// A reversed column range (`B1:A1`) likewise does not panic and resolves
    /// the same values as its ordered equivalent.
    #[test]
    fn reversed_column_range_resolves_same_values_as_ordered_range() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>8</v></c></row></sheetData></worksheet>"#;

        let ordered = parse_sqref("A1:B1");
        let reversed = parse_sqref("B1:A1");
        let ordered_values = extract_range_values(xml, &ordered[0]);
        let reversed_values = extract_range_values(xml, &reversed[0]);
        assert_eq!(ordered_values, vec![Some(7.0), Some(8.0)]);
        assert_eq!(reversed_values, ordered_values);
    }
}

#[cfg(test)]
mod package_streaming_integration_tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    pub(super) fn forged_worksheet_package() -> Vec<u8> {
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let padding = "x".repeat(2 * 1024);
        let worksheet = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{padding}</t></is></c></row></sheetData></worksheet>"#
        );
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rSheet1"/></sheets></workbook>"#;
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
            for (path, body) in [
                (SHEET_PART, worksheet.as_bytes()),
                ("xl/workbook.xml", workbook.as_bytes()),
                ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }

        // The worksheet is the first local and central entry. Forge both
        // uncompressed-size declarations to one byte while leaving its stored
        // payload and compressed size intact.
        bytes[22..26].copy_from_slice(&1u32.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("worksheet central-directory header");
        bytes[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());
        bytes
    }

    #[test]
    fn production_worksheet_stream_restores_canonical_actual_overrun_and_poisons_session() {
        const LIMIT: u64 = 1024;
        let mut archive =
            open_zip_with_limits(forged_worksheet_package(), Some(LIMIT), Some(16 * 1024))
                .expect("forged declarations pass metadata preflight");

        let first = archive
            .run_operation("parse-sheet", |archive| {
                let shared = WorkbookShared::load(archive)?;
                parse_sheet_with(archive, &shared, 0, "Sheet1")
            })
            .expect_err("actual worksheet output must cross the configured limit");
        assert!(first.starts_with("OOXML_RESOURCE_LIMIT:"), "{first}");
        let details: serde_json::Value = serde_json::from_str(
            first
                .strip_prefix("OOXML_RESOURCE_LIMIT:")
                .expect("canonical envelope prefix"),
        )
        .expect("resource envelope is JSON");
        let violation = &details["details"]["violation"];
        assert_eq!(violation["operation"], "parse-sheet");
        assert_eq!(violation["part"], "xl/worksheets/sheet1.xml");
        assert_eq!(violation["metric"], "actual-inflated-bytes");
        assert_eq!(violation["limit"], LIMIT);
        assert_eq!(violation["observed"], LIMIT + 1);

        let later = archive
            .run_operation("parse", |_| Ok(()))
            .expect_err("poisoned package rejects later operations");
        assert_eq!(later, first);
    }
}

#[cfg(test)]
mod rb7_partial_degradation_tests {
    //! RB7: one corrupt sheet must not fail the whole workbook. `parse_sheet`
    //! degrades a sheet whose XML can't be read/parsed into an empty placeholder
    //! carrying a part-tagged `parseError`, so the other sheets stay openable.
    use super::*;
    use std::io::{Cursor, Write};

    /// Build a 3-sheet workbook. `broken` (0-based) sheet gets `broken_xml` as its
    /// worksheet part; pass malformed XML to simulate corruption, or `None` to
    /// omit the part entirely (an unreadable sheet). Healthy sheets carry one
    /// cell so a real parse is distinguishable from a placeholder.
    fn build_three_sheet_workbook(broken: usize, broken_xml: Option<&str>) -> Vec<u8> {
        let good_sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="str"><v>ok</v></c></row></sheetData></worksheet>"#;
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/><sheet name="Beta" sheetId="2" r:id="rId2"/><sheet name="Gamma" sheetId="3" r:id="rId3"/></sheets></workbook>"#;
        let wb_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>"#;

        let mut entries: Vec<(String, String)> = vec![
            ("xl/workbook.xml".into(), workbook.into()),
            ("xl/_rels/workbook.xml.rels".into(), wb_rels.into()),
        ];
        for i in 0..3 {
            if i == broken {
                // `None` ⇒ omit the part entirely → its read fails → placeholder.
                if let Some(xml) = broken_xml {
                    entries.push((format!("xl/worksheets/sheet{}.xml", i + 1), xml.into()));
                }
            } else {
                entries.push((
                    format!("xl/worksheets/sheet{}.xml", i + 1),
                    good_sheet.into(),
                ));
            }
        }
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in &entries {
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    fn parse_sheet_json(data: &[u8], idx: u32, name: &str) -> serde_json::Value {
        let json = parse_sheet_native(data, idx, name)
            .unwrap_or_else(|e| panic!("sheet {idx} ({name}) must parse or degrade, got: {e}"));
        serde_json::from_str(&json).unwrap()
    }

    /// NEUTRALIZATION: a workbook whose middle sheet XML is malformed still opens.
    /// The healthy sheets parse; the broken one is an empty placeholder whose
    /// `parseError` names the offending part (`xl/worksheets/sheet2.xml`).
    #[test]
    fn rb7_one_broken_sheet_degrades_rest_parse() {
        // Unterminated element → parse_worksheet fails.
        let data = build_three_sheet_workbook(1, Some("<worksheet><sheetData><row>"));

        // Healthy sheets: real cell, no parseError.
        for (idx, name) in [(0u32, "Alpha"), (2, "Gamma")] {
            let ws = parse_sheet_json(&data, idx, name);
            assert!(
                ws["parseError"].is_null(),
                "healthy sheet {name} must carry no parseError; got {ws}"
            );
            assert!(
                !ws["rows"].as_array().unwrap().is_empty(),
                "healthy sheet {name} keeps its cell data"
            );
        }

        // Broken sheet: placeholder with a part-tagged error and no rows.
        let broken = parse_sheet_json(&data, 1, "Beta");
        let err = broken["parseError"]
            .as_str()
            .expect("broken sheet carries a parseError string");
        assert!(
            err.starts_with("xl/worksheets/sheet2.xml:"),
            "error must name the offending part; got {err:?}"
        );
        assert!(
            broken["rows"].as_array().unwrap().is_empty(),
            "placeholder sheet has no rows"
        );
        // Name is preserved so the tab still shows.
        assert_eq!(broken["name"].as_str(), Some("Beta"));
    }

    /// A sheet whose part is entirely missing from the archive also degrades to a
    /// placeholder rather than failing the whole workbook.
    #[test]
    fn rb7_missing_sheet_part_degrades() {
        let data = build_three_sheet_workbook(2, None); // sheet3.xml omitted
                                                        // Healthy sheets still parse.
        assert!(parse_sheet_json(&data, 0, "Alpha")["parseError"].is_null());
        // Missing sheet degrades.
        let broken = parse_sheet_json(&data, 2, "Gamma");
        let err = broken["parseError"]
            .as_str()
            .expect("missing sheet part yields a placeholder + error");
        assert!(
            err.starts_with("xl/worksheets/sheet3.xml:"),
            "error names the missing part; got {err:?}"
        );
    }

    // ── #774: whole-container degradation ────────────────────────────────────

    /// #774 MAJOR: a truncated / corrupt ZIP CONTAINER — the most common way a
    /// xlsx is broken — degrades to a placeholder workbook (one tab) tagged with
    /// the container, rather than throwing an opaque `ZipArchive::new` error before
    /// any part is read. Symmetric with docx / pptx container degradation.
    #[test]
    fn corrupt_zip_container_degrades_to_placeholder_workbook() {
        // Truncated container: a valid workbook cut off partway is not a readable zip.
        let full = build_three_sheet_workbook(9, None); // 9 ⇒ no sheet is broken
        let truncated = &full[..full.len() / 2];

        // Workbook index opens with a single placeholder sheet + a container error.
        let wb_json =
            parse_workbook_native(truncated).expect("a corrupt container must open, not error out");
        let wb: serde_json::Value = serde_json::from_str(&wb_json).unwrap();
        let sheets = wb["sheets"]
            .as_array()
            .expect("placeholder workbook has sheets");
        assert_eq!(sheets.len(), 1, "one placeholder tab for the whole file");
        let wb_err = wb["parseError"]
            .as_str()
            .expect("degraded workbook carries a container-tagged parseError");
        assert!(
            wb_err.starts_with("(zip container): "),
            "workbook error is tagged with the container exactly once (one paren pair); got {wb_err:?}"
        );
        assert_eq!(
            wb_err.matches("zip container").count(),
            1,
            "the container tag must not be doubled; got {wb_err:?}"
        );

        // The lazily-parsed sheet 0 is the container-tagged placeholder overlay.
        let ws = parse_sheet_json(truncated, 0, "(zip container)");
        let ws_err = ws["parseError"]
            .as_str()
            .expect("placeholder sheet carries a parseError");
        assert!(
            ws_err.starts_with("(zip container): "),
            "sheet error is tagged with the container exactly once (one paren pair); got {ws_err:?}"
        );
        assert_eq!(
            ws_err.matches("zip container").count(),
            1,
            "the container tag must not be doubled; got {ws_err:?}"
        );
        assert!(
            ws["rows"].as_array().unwrap().is_empty(),
            "placeholder sheet has no rows"
        );

        // Not-a-zip-at-all also degrades (no local file header).
        let garbage =
            parse_workbook_native(b"this is definitely not a zip file").expect("non-zip opens");
        let gv: serde_json::Value = serde_json::from_str(&garbage).unwrap();
        let garbage_err = gv["parseError"]
            .as_str()
            .expect("non-zip degrades with a container-tagged error");
        assert!(
            garbage_err.starts_with("(zip container): "),
            "error is tagged with the container exactly once (one paren pair); got {garbage_err:?}"
        );
        assert_eq!(
            garbage_err.matches("zip container").count(),
            1,
            "the container tag must not be doubled; got {garbage_err:?}"
        );
    }

    // ── #832 / #833-1: implicit references through the whole-archive path ─────

    /// Build a 1-sheet workbook that OMITS `@r` on both `<row>` and every `<c>`
    /// (the minimal enterprise-exporter shape from #832 / #833-1), backed by a
    /// real `sharedStrings.xml`. Exercised end-to-end through `parse_sheet_native`
    /// — the same code the WASM `parse_sheet` entry runs — so this proves the fix
    /// survives ZIP extraction + shared-string resolution, not just the isolated
    /// `parse_worksheet` unit path.
    fn build_implicit_ref_workbook() -> Vec<u8> {
        let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        // Two rows, no @r anywhere; row 1 = three shared strings, row 2 = a
        // shared string then two numbers. Positions must fill A1:C2.
        let sheet = format!(
            r#"<worksheet xmlns="{ns}"><sheetData>
              <row><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c></row>
              <row><c t="s"><v>3</v></c><c t="n"><v>42.5</v></c><c t="n"><v>100</v></c></row>
            </sheetData></worksheet>"#
        );
        let shared = format!(
            r#"<sst xmlns="{ns}" count="4" uniqueCount="4"><si><t>Alpha</t></si><si><t>Beta</t></si><si><t>Gamma</t></si><si><t>Delta</t></si></sst>"#
        );
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let wb_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>"#;
        let entries: Vec<(String, String)> = vec![
            ("xl/workbook.xml".into(), workbook.into()),
            ("xl/_rels/workbook.xml.rels".into(), wb_rels.into()),
            ("xl/worksheets/sheet1.xml".into(), sheet),
            ("xl/sharedStrings.xml".into(), shared),
            ("xl/styles.xml".into(), styles.into()),
        ];
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in &entries {
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    fn build_missing_sheet_workbook() -> Vec<u8> {
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/></Relationships>"#;
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, body) in [
                ("xl/workbook.xml", workbook),
                ("xl/_rels/workbook.xml.rels", rels),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn build_sheet_xml_workbook(sheet: &str) -> Vec<u8> {
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for (path, body) in [
                ("xl/worksheets/sheet1.xml", sheet),
                ("xl/workbook.xml", workbook),
                ("xl/_rels/workbook.xml.rels", rels),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn build_forged_ancillary_workbook() -> Vec<u8> {
        let padding = "x".repeat(2 * 1024);
        let drawing = format!(
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><!--{padding}--></xdr:wsDr>"#
        );
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rDrawing"/></worksheet>"#;
        let sheet_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let workbook_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for (path, body) in [
                ("xl/drawings/drawing1.xml", drawing.as_str()),
                ("xl/worksheets/sheet1.xml", sheet),
                ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
                ("xl/workbook.xml", workbook),
                ("xl/_rels/workbook.xml.rels", workbook_rels),
            ] {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }

        // The ancillary drawing is first. Its declared uncompressed size passes
        // metadata preflight, but reading it exceeds the per-entry limit.
        bytes[22..26].copy_from_slice(&1u32.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("drawing central-directory header");
        bytes[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());
        bytes
    }

    fn corrupt_first_entry_crc(bytes: &mut [u8]) {
        let wrong_crc = u32::from_le_bytes(bytes[14..18].try_into().unwrap()) ^ u32::MAX;
        bytes[14..18].copy_from_slice(&wrong_crc.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .unwrap();
        bytes[central + 16..central + 20].copy_from_slice(&wrong_crc.to_le_bytes());
    }

    fn drain_cursor_model(data: Vec<u8>) -> serde_json::Value {
        let mut archive = XlsxArchive::new(data, None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Sheet1").unwrap();
        let mut rows = Vec::new();
        loop {
            let payload = archive.pull_sheet_cursor_inner(1).unwrap();
            let mut envelope: serde_json::Value = serde_json::from_slice(&payload).unwrap();
            if envelope["kind"] == "rows" {
                rows.append(envelope["rows"].as_array_mut().unwrap());
                continue;
            }
            let mut worksheet = envelope["worksheet"].take();
            if worksheet["parseError"].is_null() {
                worksheet["rows"] = serde_json::Value::Array(rows);
            }
            archive.acknowledge_sheet_cursor_terminal_inner().unwrap();
            return worksheet;
        }
    }

    fn build_chart_workbook(sheet1: &str, chart: &str, sheet2: Option<&str>) -> Vec<u8> {
        let sheets = if sheet2.is_some() {
            r#"<sheet name="Sheet1" sheetId="1" r:id="rSheet1"/><sheet name="Data" sheetId="2" r:id="rSheet2"/>"#
        } else {
            r#"<sheet name="Sheet1" sheetId="1" r:id="rSheet1"/>"#
        };
        let workbook = format!(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#
        );
        let second_rel = if sheet2.is_some() {
            r#"<Relationship Id="rSheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>"#
        } else {
            ""
        };
        let workbook_rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>{second_rel}</Relationships>"#
        );
        let sheet_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
        let drawing = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let drawing_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#;
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="165" formatCode="0.0000"/></numFmts><fonts count="1"><font><sz val="13"/><name val="Cursor Test Font"/></font></fonts><fills count="0"/><borders count="0"/><cellStyleXfs count="1"><xf fontId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="165" fontId="0"/></cellXfs></styleSheet>"#;
        let mut entries = vec![
            ("xl/workbook.xml", workbook.as_str()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_str()),
            ("xl/styles.xml", styles),
            ("xl/worksheets/sheet1.xml", sheet1),
            ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
            ("xl/drawings/drawing1.xml", drawing),
            ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
            ("xl/charts/chart1.xml", chart),
        ];
        if let Some(sheet2) = sheet2 {
            entries.push(("xl/worksheets/sheet2.xml", sheet2));
        }
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, body) in entries {
                writer.start_file(path, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn assert_cursor_parity(data: Vec<u8>) -> serde_json::Value {
        let legacy: serde_json::Value =
            serde_json::from_str(&parse_sheet_native(&data, 0, "Sheet1").unwrap()).unwrap();
        let streamed = drain_cursor_model(data);
        assert_eq!(streamed, legacy);
        streamed
    }

    #[test]
    fn wasm_cursor_commits_operation_only_after_terminal_ack() {
        let mut archive =
            XlsxArchive::new(build_implicit_ref_workbook(), None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Sheet1").unwrap();
        loop {
            let payload = archive.pull_sheet_cursor(1).unwrap();
            let value: serde_json::Value = serde_json::from_slice(&payload).unwrap();
            if value["kind"] == "finished" {
                assert!(archive.terminal_awaiting_ack);
                assert!(archive.archive.as_ref().unwrap().operation.is_active());
                assert_eq!(value["worksheet"]["rows"], serde_json::json!([]));
                let usage: serde_json::Value =
                    serde_json::from_slice(&archive.sheet_cursor_resource_usage().unwrap())
                        .unwrap();
                assert!(usage["operationInflatedBytes"].as_u64().unwrap() > 0);
                break;
            }
        }
        archive.acknowledge_sheet_cursor_terminal().unwrap();
        assert!(!archive.terminal_awaiting_ack);
        assert!(archive.active_worksheet.is_none());
        assert!(!archive.archive.as_ref().unwrap().operation.is_active());
        assert!(archive.sheet_cursor_resource_usage().is_ok());
        archive.close_sheet_cursor();
    }

    #[test]
    fn full_parse_after_sheet_cursor_matches_fresh_styles_and_chart_formats() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Series</t></is></c></row><row r="2"><c r="A2" s="1"><v>1.25</v></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" s="1"><v>2.5</v></c><c r="B3"><v>20</v></c></row></sheetData><drawing r:id="rDrawing"/></worksheet>"#;
        let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!A1</c:f></c:strRef></c:tx><c:cat><c:numRef><c:f>Sheet1!A2:A3</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>Sheet1!B2:B3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let data = build_chart_workbook(sheet, chart, None);

        let mut fresh = XlsxArchive::new(data.clone(), None, None, None).unwrap();
        let fresh_workbook = fresh.parse().unwrap();

        let mut after_cursor = XlsxArchive::new(data, None, None, None).unwrap();
        after_cursor.open_sheet_cursor(0, "Sheet1").unwrap();
        let terminal = loop {
            let payload = after_cursor.pull_sheet_cursor(128).unwrap();
            let value: serde_json::Value = serde_json::from_slice(&payload).unwrap();
            if value["kind"] == "finished" {
                break value;
            }
        };
        assert_eq!(
            terminal["worksheet"]["defaultFontFamily"],
            "Cursor Test Font"
        );
        assert_eq!(
            terminal["worksheet"]["charts"][0]["chart"]["series"][0]["catFormatBuiltinId"],
            165
        );
        after_cursor.acknowledge_sheet_cursor_terminal().unwrap();

        let after_cursor_workbook = after_cursor.parse().unwrap();
        assert_eq!(after_cursor_workbook, fresh_workbook);
        let workbook: serde_json::Value = serde_json::from_slice(&after_cursor_workbook).unwrap();
        assert_eq!(workbook["styles"]["fonts"][0]["name"], "Cursor Test Font");
        assert_eq!(workbook["styles"]["numFmts"][0]["formatCode"], "0.0000");
    }

    #[test]
    fn wasm_cursor_close_before_terminal_ack_cancels_idempotently() {
        let mut archive =
            XlsxArchive::new(build_implicit_ref_workbook(), None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Sheet1").unwrap();
        while !archive.sheet_cursor_pull_finished() {
            archive.pull_sheet_cursor(128).unwrap();
        }
        let usage = archive.sheet_cursor_resource_usage().unwrap();
        archive.close_sheet_cursor();
        archive.close_sheet_cursor();
        archive.cancel_sheet_cursor();
        assert!(!archive.terminal_awaiting_ack);
        assert!(archive.active_worksheet.is_none());
        assert!(!archive.archive.as_ref().unwrap().operation.is_active());
        assert_eq!(archive.sheet_cursor_resource_usage().unwrap(), usage);
    }

    #[test]
    fn wasm_cursor_missing_sheet_is_a_provisional_placeholder_until_ack() {
        let mut archive =
            XlsxArchive::new(build_missing_sheet_workbook(), None, None, None).unwrap();
        archive.open_sheet_cursor(0, "Sheet1").unwrap();
        let payload = archive.pull_sheet_cursor(128).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&payload).unwrap();
        assert_eq!(value["kind"], "finished");
        assert!(value["worksheet"]["parseError"]
            .as_str()
            .unwrap()
            .starts_with("xl/worksheets/missing.xml: "));
        assert!(archive.archive.as_ref().unwrap().operation.is_active());
        archive.acknowledge_sheet_cursor_terminal().unwrap();
        assert!(!archive.archive.as_ref().unwrap().operation.is_active());
    }

    #[test]
    fn wasm_cursor_corrupt_container_matches_legacy_placeholder_and_commits_on_ack() {
        let mut archive = XlsxArchive::new(b"not a zip".to_vec(), None, None, None).unwrap();
        let container_error = match &archive.archive {
            Err(error) => error.clone(),
            Ok(_) => panic!("corrupt container must be deferred"),
        };
        archive.open_sheet_cursor(0, "ignored").unwrap();
        assert!(!archive.sheet_cursor_pull_finished());
        let payload = archive.pull_sheet_cursor(128).unwrap();
        let value: serde_json::Value = serde_json::from_slice(&payload).unwrap();
        assert_eq!(value["kind"], "finished");
        assert_eq!(
            value["worksheet"],
            serde_json::to_value(degraded_container_sheet(container_error)).unwrap()
        );
        assert!(archive.terminal_awaiting_ack);
        archive.acknowledge_sheet_cursor_terminal_inner().unwrap();
        assert!(archive.active_worksheet.is_none());
        assert!(!archive.terminal_awaiting_ack);
    }

    #[test]
    fn wasm_cursor_resource_poison_never_prepares_or_allows_terminal_ack() {
        let data = super::package_streaming_integration_tests::forged_worksheet_package();
        let mut archive = XlsxArchive::new(data, Some(1024), Some(16 * 1024), None).unwrap();
        archive.open_sheet_cursor(0, "Sheet1").unwrap();
        assert!(archive.pull_sheet_cursor_inner(128).is_err());
        assert!(!archive.sheet_cursor_pull_finished());
        assert!(!archive.terminal_awaiting_ack);
        assert!(archive.active_worksheet.is_none());
        assert!(!archive.archive.as_ref().unwrap().operation.is_active());
        assert!(archive.acknowledge_sheet_cursor_terminal_inner().is_err());
        archive.cancel_sheet_cursor();
        archive.close_sheet_cursor();
    }

    #[test]
    fn wasm_cursor_ancillary_resource_poison_is_terminal_and_clears_state() {
        let mut archive = XlsxArchive::new(
            build_forged_ancillary_workbook(),
            Some(1024),
            Some(16 * 1024),
            None,
        )
        .expect("forged ancillary declaration passes metadata preflight");
        archive.open_sheet_cursor(0, "Sheet1").unwrap();

        let error = loop {
            match archive.pull_sheet_cursor_inner(128) {
                Ok(payload) => {
                    let envelope: serde_json::Value = serde_json::from_slice(&payload).unwrap();
                    assert_eq!(envelope["kind"], "rows");
                }
                Err(error) => break error,
            }
        };
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"), "{error}");
        assert!(!archive.sheet_cursor_pull_finished());
        assert!(!archive.terminal_awaiting_ack);
        assert!(archive.active_worksheet.is_none());
        assert!(!archive.archive.as_ref().unwrap().operation.is_active());
        assert!(archive.acknowledge_sheet_cursor_terminal_inner().is_err());
    }

    #[test]
    fn wasm_cursor_malformed_tail_placeholder_matches_legacy_parse_sheet() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><broken>"#;
        let data = build_sheet_xml_workbook(sheet);
        let legacy: serde_json::Value =
            serde_json::from_str(&parse_sheet_native(&data, 0, "Sheet1").unwrap()).unwrap();
        assert_eq!(drain_cursor_model(data), legacy);
        assert!(!legacy["parseError"].is_null());
    }

    #[test]
    fn wasm_cursor_crc_placeholder_matches_legacy_parse_sheet() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let mut data = build_sheet_xml_workbook(sheet);
        corrupt_first_entry_crc(&mut data);
        let legacy: serde_json::Value =
            serde_json::from_str(&parse_sheet_native(&data, 0, "Sheet1").unwrap()).unwrap();
        assert_eq!(drain_cursor_model(data), legacy);
        assert!(!legacy["parseError"].is_null());
    }

    #[test]
    fn wasm_cursor_current_sheet_cacheless_chart_matches_legacy_drain() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Two</t></is></c><c r="B3"><v>20</v></c></row></sheetData><drawing r:id="rDrawing"/></worksheet>"#;
        let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!A1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!A2:A3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!B2:B3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let model = assert_cursor_parity(build_chart_workbook(sheet, chart, None));
        assert_eq!(
            model["charts"][0]["chart"]["series"][0]["values"],
            serde_json::json!([10.0, 20.0])
        );
    }

    #[test]
    fn wasm_cursor_cross_sheet_cacheless_chart_matches_legacy_drain() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rDrawing"/></worksheet>"#;
        let data = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>30</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Two</t></is></c><c r="B3"><v>40</v></c></row></sheetData></worksheet>"#;
        let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!A1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>Data!A2:A3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!B2:B3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let model = assert_cursor_parity(build_chart_workbook(sheet, chart, Some(data)));
        assert_eq!(
            model["charts"][0]["chart"]["series"][0]["values"],
            serde_json::json!([30.0, 40.0])
        );
    }

    #[test]
    fn wasm_cursor_authored_chart_cache_precedence_matches_legacy_drain() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rDrawing"/></worksheet>"#;
        let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!A1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let model = assert_cursor_parity(build_chart_workbook(sheet, chart, None));
        assert_eq!(
            model["charts"][0]["chart"]["series"][0]["values"],
            serde_json::json!([99.0])
        );
    }

    #[test]
    fn wasm_cursor_current_sheet_sparkline_matches_legacy_drain() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1"><v>5</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData><extLst><ext uri="spark"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:A2</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>"#;
        let model = assert_cursor_parity(build_sheet_xml_workbook(sheet));
        assert_eq!(
            model["sparklineGroups"][0]["sparklines"][0]["values"],
            serde_json::json!([5.0, 7.0])
        );
    }

    #[test]
    fn wasm_cursor_cross_sheet_sparkline_with_nonempty_current_matches_legacy_drain() {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1"><v>99</v></c></row></sheetData><drawing r:id="rDrawing"/><extLst><ext uri="spark"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Data!A1:A2</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>"#;
        let data = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>11</v></c></row><row r="2"><c r="A2"><v>22</v></c></row></sheetData></worksheet>"#;
        let chart =
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let model = assert_cursor_parity(build_chart_workbook(sheet, chart, Some(data)));
        assert_eq!(
            model["sparklineGroups"][0]["sparklines"][0]["values"],
            serde_json::json!([11.0, 22.0])
        );
    }

    /// End-to-end: an implicit-reference workbook parses to a full 2×3 grid —
    /// not a single A1 cell. Cell coordinates (col/row) and the shared-string
    /// `si` indices survive ZIP extraction + the shared-string load. (The `si`
    /// index → text mapping is resolved consumer-side and covered elsewhere;
    /// here `<v>0..3</v>` map to si 0..3 by insertion order.) This is the #832
    /// reproduction driven through the real archive path (== the WASM entry).
    #[test]
    fn implicit_refs_resolve_full_grid_through_archive() {
        let data = build_implicit_ref_workbook();
        let ws = parse_sheet_json(&data, 0, "Sheet1");
        assert!(
            ws["parseError"].is_null(),
            "healthy implicit-ref sheet must carry no parseError; got {ws}"
        );
        let rows = ws["rows"].as_array().expect("rows array");
        assert_eq!(rows.len(), 2, "two rows must materialize");
        assert_eq!(rows[0]["index"].as_u64(), Some(1), "first <row> → 1");
        assert_eq!(rows[1]["index"].as_u64(), Some(2), "second <row> → 2");

        let cell_at = |r: usize, c: usize| -> &serde_json::Value { &ws["rows"][r]["cells"][c] };

        // CellValue is internally tagged (`tag = "type"`): a shared reference
        // serializes as { "type": "shared", "si": N }, a number as
        // { "type": "number", "number": X }.
        // Row 1: three shared strings at columns A, B, C (si 0, 1, 2).
        for (i, col) in [1u64, 2, 3].iter().enumerate() {
            let cell = cell_at(0, i);
            assert_eq!(cell["col"].as_u64(), Some(*col), "row1 cell {i} col");
            assert_eq!(cell["row"].as_u64(), Some(1), "row1 cell {i} row");
            assert_eq!(cell["value"]["type"].as_str(), Some("shared"));
            assert_eq!(
                cell["value"]["si"].as_u64(),
                Some(i as u64),
                "row1 cell {i} shared si"
            );
        }

        // Row 2: A2 = shared si 3, B2 = 42.5, C2 = 100.
        let a2 = cell_at(1, 0);
        assert_eq!((a2["col"].as_u64(), a2["row"].as_u64()), (Some(1), Some(2)));
        assert_eq!(a2["value"]["si"].as_u64(), Some(3));
        let b2 = cell_at(1, 1);
        assert_eq!((b2["col"].as_u64(), b2["row"].as_u64()), (Some(2), Some(2)));
        assert_eq!(b2["value"]["number"].as_f64(), Some(42.5));
        let c2 = cell_at(1, 2);
        assert_eq!((c2["col"].as_u64(), c2["row"].as_u64()), (Some(3), Some(2)));
        assert_eq!(c2["value"]["number"].as_f64(), Some(100.0));
    }

    // ── #773: corrupt sharedStrings surfaces (not silent) ────────────────────

    /// Build a 1-sheet workbook whose one string cell (`A1`, `t="s"`) references
    /// shared-string index 0. `shared_strings_xml` becomes `xl/sharedStrings.xml`
    /// verbatim; pass malformed XML to simulate corruption, or `None` to omit it.
    fn build_workbook_with_shared_strings(shared_strings_xml: Option<&str>) -> Vec<u8> {
        let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#;
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let wb_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        // `parse_xlsx_inner_with` (the workbook-index path, unlike the per-sheet
        // path) reads `xl/styles.xml` with `?`, so a minimal styles part is needed.
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>"#;
        let mut entries: Vec<(String, String)> = vec![
            ("xl/workbook.xml".into(), workbook.into()),
            ("xl/_rels/workbook.xml.rels".into(), wb_rels.into()),
            ("xl/worksheets/sheet1.xml".into(), sheet.into()),
            ("xl/styles.xml".into(), styles.into()),
        ];
        if let Some(ss) = shared_strings_xml {
            entries.push(("xl/sharedStrings.xml".into(), ss.into()));
        }
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in &entries {
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    /// Build a synthetic workbook whose one shared string ("課長") carries a
    /// `<phoneticPr>` + two `<rPh>` runs, and whose sheet has an A1 cell with
    /// `ph="1"` (opts into the furigana) and a B1 cell with the same string but
    /// no `ph` (furigana off). Mirrors the ph=true/ph=false split of the
    /// private fixtures. Styles include a small phonetic font at index 2.
    fn build_phonetic_workbook() -> Vec<u8> {
        let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let sheet = format!(
            r#"<worksheet xmlns="{ns}"><sheetData><row r="1"><c r="A1" t="s" ph="1"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>"#
        );
        let ss = format!(
            r#"<sst xmlns="{ns}" count="2" uniqueCount="1"><si><t>課長</t><rPh sb="0" eb="1"><t>カ</t></rPh><rPh sb="1" eb="2"><t>チョウ</t></rPh><phoneticPr fontId="2" alignment="center"/></si></sst>"#
        );
        let workbook = format!(
            r#"<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Alpha" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        let wb_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let styles = format!(
            r#"<styleSheet xmlns="{ns}"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/></font><font><sz val="6"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>"#
        );
        let entries: Vec<(String, String)> = vec![
            ("xl/workbook.xml".into(), workbook),
            ("xl/_rels/workbook.xml.rels".into(), wb_rels.into()),
            ("xl/worksheets/sheet1.xml".into(), sheet),
            ("xl/styles.xml".into(), styles),
            ("xl/sharedStrings.xml".into(), ss),
        ];
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in &entries {
                w.start_file(name.as_str(), o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    /// End-to-end (real zip → JSON): a `<si>` with `<rPh>`/`<phoneticPr>` surfaces
    /// on the shared-string table, and the cell `ph` attribute flows onto the
    /// per-cell `showPhonetic` flag. B1 (no `ph`) stays false even though it
    /// references the SAME phonetic string — the reading is display-off there,
    /// exactly like the private fixtures (rPh present, no cell opts in).
    #[test]
    fn phonetic_workbook_round_trips_rph_and_cell_ph() {
        let data = build_phonetic_workbook();
        // The full `ParsedWorkbook` (what `parse_xlsx` ships to TS) carries the
        // phonetic shared string in its `sharedStrings` table.
        let parsed = parse_xlsx_inner(&data).expect("workbook opens");
        let wb_json = serde_json::to_string(&parsed).unwrap();
        let wb: serde_json::Value = serde_json::from_str(&wb_json).unwrap();
        let si0 = &wb["sharedStrings"][0];
        assert_eq!(si0["text"].as_str(), Some("課長"), "base text only");
        let rph = si0["phoneticRuns"]
            .as_array()
            .expect("phoneticRuns present");
        assert_eq!(rph.len(), 2);
        assert_eq!(rph[0]["sb"].as_u64(), Some(0));
        assert_eq!(rph[0]["eb"].as_u64(), Some(1));
        assert_eq!(rph[0]["text"].as_str(), Some("カ"));
        assert_eq!(si0["phoneticPr"]["fontId"].as_u64(), Some(2));
        assert_eq!(si0["phoneticPr"]["alignment"].as_str(), Some("center"));
        // type absent → the consumer applies the fullwidthKatakana default.
        assert!(si0["phoneticPr"].get("type").is_none());

        // The sheet's A1 opts in (ph=1); B1 does not (schema default false).
        let ws = parse_sheet_json(&data, 0, "Alpha");
        let cells = ws["rows"][0]["cells"].as_array().unwrap();
        let a1 = cells.iter().find(|c| c["col"].as_u64() == Some(1)).unwrap();
        let b1 = cells.iter().find(|c| c["col"].as_u64() == Some(2)).unwrap();
        assert_eq!(a1["showPhonetic"].as_bool(), Some(true), "A1 ph=1 → show");
        assert!(
            b1.get("showPhonetic").is_none() || b1["showPhonetic"].as_bool() == Some(false),
            "B1 has no ph → showPhonetic omitted/false; got {b1}"
        );
    }

    /// #773: a PRESENT-but-corrupt `xl/sharedStrings.xml` (§18.4.9) silently
    /// blanked every string cell before this fix. Now the workbook still opens (no
    /// sheet is taken down) but the loss is SURFACED as a workbook-level,
    /// part-tagged `parseError` — no longer silent.
    #[test]
    fn corrupt_shared_strings_surfaces_workbook_error() {
        // Unterminated element → parse_guarded fails on a part that IS present.
        let data = build_workbook_with_shared_strings(Some("<sst><si><t>hi"));
        let wb_json = parse_workbook_native(&data).expect("workbook still opens");
        let wb: serde_json::Value = serde_json::from_str(&wb_json).unwrap();
        let err = wb["parseError"]
            .as_str()
            .expect("corrupt sharedStrings surfaces a workbook-level parseError");
        assert!(
            err.starts_with("xl/sharedStrings.xml:"),
            "error names the offending part; got {err:?}"
        );
        // The sheet itself is NOT taken down — it still opens as a real sheet
        // (no per-sheet parseError), only its string cell is blank.
        let ws = parse_sheet_json(&data, 0, "Alpha");
        assert!(
            ws["parseError"].is_null(),
            "the sheet must still open (partial degradation, not a placeholder)"
        );
    }

    /// A HEALTHY sharedStrings.xml leaves NO workbook-level `parseError` (the
    /// silent-degradation surfacing is inert for valid files — wire-unchanged).
    #[test]
    fn healthy_shared_strings_no_workbook_error() {
        let data = build_workbook_with_shared_strings(Some(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>hi</t></si></sst>"#,
        ));
        let wb_json = parse_workbook_native(&data).expect("workbook opens");
        let wb: serde_json::Value = serde_json::from_str(&wb_json).unwrap();
        assert!(
            wb["parseError"].is_null(),
            "a healthy sharedStrings must not surface any parseError; got {wb}"
        );
        assert!(
            !wb_json.contains("parseError"),
            "healthy workbook JSON must not carry a parseError key"
        );
    }

    /// An ABSENT sharedStrings.xml is legitimate (a workbook with no string cells)
    /// and must NOT surface a `parseError` — only a present-but-corrupt part does.
    #[test]
    fn absent_shared_strings_no_workbook_error() {
        let data = build_workbook_with_shared_strings(None);
        let wb_json = parse_workbook_native(&data).expect("workbook opens");
        let wb: serde_json::Value = serde_json::from_str(&wb_json).unwrap();
        assert!(
            wb["parseError"].is_null(),
            "an absent sharedStrings is normal, not a degradation; got {wb}"
        );
    }
}

#[cfg(test)]
mod pivot_metadata_tests {
    use super::parse_sheet_native;
    use serde_json::Value;
    use std::io::{Cursor, Write};

    fn workbook_with_pivot(
        pivot_xml: &str,
        pivot_rels: Option<&str>,
        cache_xml: Option<&str>,
    ) -> Vec<u8> {
        let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#;
        let workbook_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
        let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="3"><v>42</v></c></row></sheetData></worksheet>"#;
        let worksheet_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#;

        let mut entries = vec![
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
            ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels),
            ("xl/pivotTables/pivotTable1.xml", pivot_xml),
        ];
        if let Some(rels) = pivot_rels {
            entries.push(("xl/pivotTables/_rels/pivotTable1.xml.rels", rels));
        }
        if let Some(cache) = cache_xml {
            entries.push(("xl/pivotCache/pivotCacheDefinition7.xml", cache));
        }

        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, xml) in entries {
                zip.start_file(path, options).unwrap();
                zip.write_all(xml.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        bytes
    }

    fn workbook_with_worksheet_rels(worksheet_rels: &[u8]) -> Vec<u8> {
        workbook_with_relationship_parts(worksheet_rels, PIVOT_RELS.as_bytes())
    }

    fn workbook_with_relationship_parts(worksheet_rels: &[u8], pivot_rels: &[u8]) -> Vec<u8> {
        let entries: [(&str, &[u8]); 7] = [
            (
                "xl/workbook.xml",
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>"#,
            ),
            ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels),
            ("xl/pivotTables/pivotTable1.xml", COMPLETE_PIVOT.as_bytes()),
            (
                "xl/pivotTables/_rels/pivotTable1.xml.rels",
                pivot_rels,
            ),
            (
                "xl/pivotCache/pivotCacheDefinition7.xml",
                COMPLETE_CACHE.as_bytes(),
            ),
        ];
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            for (path, content) in entries {
                zip.start_file(path, options).unwrap();
                zip.write_all(content).unwrap();
            }
            zip.finish().unwrap();
        }
        bytes
    }

    fn parse(data: &[u8]) -> Value {
        serde_json::from_str(&parse_sheet_native(data, 0, "Report").unwrap()).unwrap()
    }

    const COMPLETE_PIVOT: &str = r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="SalesPivot" cacheId="99" dataCaption="Values"><location ref="A1:D8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><rowFields count="2"><field x="-2"/><field x="0"/></rowFields><colFields count="1"><field x="1"/></colFields><pageFields count="1"><pageField fld="-1" item="4" name="Region"/></pageFields><dataFields count="1"><dataField name="Sum of Sales" fld="3"/></dataFields><extLst><ext uri="urn:example:future-pivot-feature"/></extLst></pivotTableDefinition>"#;
    const PIVOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#;
    const COMPLETE_CACHE: &str = r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1"><cacheSource type="worksheet"><worksheetSource ref="A1:D20" sheet="Data"/></cacheSource><cacheFields count="4"><cacheField name="Category"/><cacheField name="Month"/><cacheField name="Region"/><cacheField name="Sales"/></cacheFields></pivotCacheDefinition>"#;

    #[test]
    fn preserves_saved_cells_and_attaches_complete_pivot_facts() {
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));

        // Characterization: saved worksheet values and styles are authoritative.
        assert_eq!(
            sheet["rows"][0]["cells"][0]["value"]["number"].as_f64(),
            Some(42.0)
        );
        assert_eq!(sheet["rows"][0]["cells"][0]["styleIndex"], 3);

        let pivot = &sheet["pivotTables"][0];
        assert_eq!(pivot["name"], "SalesPivot");
        assert_eq!(pivot["cacheId"], 99);
        assert_eq!(pivot["location"]["top"], 1);
        assert_eq!(pivot["location"]["right"], 4);
        // ECMA-376 §18.10 CT_Field@x and CT_PageField@fld are signed; retain sentinels.
        assert_eq!(pivot["rowFields"], serde_json::json!([-2, 0]));
        assert_eq!(pivot["columnFields"], serde_json::json!([1]));
        assert_eq!(pivot["pageFields"][0]["field"], -1);
        assert_eq!(pivot["dataFields"][0]["field"], 3);
        assert_eq!(pivot["dataFields"][0]["subtotal"], "sum");
        assert_eq!(pivot["refreshOnLoad"], true);
        assert_eq!(pivot["cacheSource"]["kind"], "worksheet");
        assert_eq!(pivot["cacheSource"]["sheet"], "Data");
        assert_eq!(pivot["status"]["state"], "complete");
        assert_eq!(
            pivot["extensionUris"],
            serde_json::json!(["urn:example:future-pivot-feature"])
        );
    }

    #[test]
    fn missing_cache_relationship_is_partial() {
        let sheet = parse(&workbook_with_pivot(COMPLETE_PIVOT, None, None));
        assert_eq!(sheet["pivotTables"][0]["status"]["state"], "partial");
        assert!(sheet["pivotTables"][0].get("refreshOnLoad").is_none());
        assert!(sheet["pivotTables"][0].get("cacheDefinitionPart").is_none());
        assert_eq!(
            sheet["pivotTables"][0]["status"]["reasons"][0]["kind"],
            "missingCacheRelationship"
        );
    }

    #[test]
    fn resolved_cache_exposes_schema_default_refresh_as_known_false() {
        let cache = COMPLETE_CACHE.replace(" refreshOnLoad=\"1\"", "");
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&cache),
        ));
        assert_eq!(sheet["pivotTables"][0]["refreshOnLoad"], false);
        assert_eq!(
            sheet["pivotTables"][0]["cacheDefinitionPart"],
            "xl/pivotCache/pivotCacheDefinition7.xml"
        );
    }

    #[test]
    fn rejects_extension_namespace_elements_as_core_pivot_identity_or_location() {
        let pivot = r#"<pivotTableDefinition xmlns="urn:not-spreadsheetml" name="Fake" cacheId="7"><location ref="A1:B2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>"#;
        let sheet = parse(&workbook_with_pivot(
            pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(
            sheet["pivotDiagnostics"][0]["reason"]["kind"],
            "malformedXml"
        );
    }

    #[test]
    fn malformed_field_fact_is_partial_instead_of_silently_omitted() {
        let pivot = COMPLETE_PIVOT.replace("<field x=\"0\"/>", "<field x=\"not-an-int\"/>");
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        let reasons = sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap();
        assert!(reasons.iter().any(|r| r["kind"] == "malformedField"));

        let bad_item = COMPLETE_PIVOT.replace("item=\"4\"", "item=\"not-unsigned\"");
        let sheet = parse(&workbook_with_pivot(
            &bad_item,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        let reasons = sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap();
        assert!(reasons.iter().any(|r| r["field"] == "pageFields"));
    }

    #[test]
    fn ambiguous_cache_relationship_is_partial_and_does_not_choose_a_target() {
        let rels = PIVOT_RELS.replace(
            "</Relationships>",
            r#"<Relationship Id="rIdCache2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition8.xml"/></Relationships>"#,
        );
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(&rels),
            Some(COMPLETE_CACHE),
        ));
        let pivot = &sheet["pivotTables"][0];
        assert!(pivot.get("cacheDefinitionPart").is_none());
        assert_eq!(
            pivot["status"]["reasons"][0]["kind"],
            "ambiguousCacheRelationship"
        );
    }

    #[test]
    fn invalid_refresh_boolean_is_unknown_and_partial() {
        let cache = COMPLETE_CACHE.replace("refreshOnLoad=\"1\"", "refreshOnLoad=\"TRUE\"");
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&cache),
        ));
        let pivot = &sheet["pivotTables"][0];
        assert!(pivot.get("refreshOnLoad").is_none());
        assert_eq!(pivot["status"]["reasons"][0]["kind"], "malformedField");
    }

    #[test]
    fn resolved_cache_part_is_retained_when_the_part_is_malformed() {
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some("<bad"),
        ));
        assert_eq!(
            sheet["pivotTables"][0]["cacheDefinitionPart"],
            "xl/pivotCache/pivotCacheDefinition7.xml"
        );
        assert_eq!(
            sheet["pivotTables"][0]["status"]["reasons"][0]["kind"],
            "malformedCacheDefinition"
        );
    }

    #[test]
    fn presentation_formats_stay_complete_but_cache_field_group_is_partial() {
        let formatted = COMPLETE_PIVOT.replace(
            "<extLst>",
            "<formats count=\"0\"/><conditionalFormats count=\"0\"/><chartFormats count=\"0\"/><extLst>",
        );
        let complete = parse(&workbook_with_pivot(
            &formatted,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(complete["pivotTables"][0]["status"]["state"], "complete");

        let grouped_cache = COMPLETE_CACHE.replace(
            "<cacheField name=\"Category\"/>",
            "<cacheField name=\"Category\"><fieldGroup/></cacheField>",
        );
        let partial = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&grouped_cache),
        ));
        let reasons = partial["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap();
        assert!(reasons.iter().any(|r| r["feature"] == "fieldGroup"));
    }

    #[test]
    fn recognized_unsupported_core_feature_is_partial_but_unknown_extension_is_not() {
        let pivot = COMPLETE_PIVOT.replace(
            "<extLst>",
            "<filters count=\"1\"><filter fld=\"0\" type=\"captionEqual\"/></filters><extLst>",
        );
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(sheet["pivotTables"][0]["status"]["state"], "partial");
        assert_eq!(
            sheet["pivotTables"][0]["status"]["reasons"][0]["kind"],
            "unsupportedSemanticFeature"
        );
        assert_eq!(
            sheet["pivotTables"][0]["status"]["reasons"][0]["feature"],
            "filters"
        );
    }

    #[test]
    fn malformed_identity_or_location_skips_only_that_pivot_with_typed_diagnostic() {
        for (pivot, reason) in [
            ("<pivotTableDefinition", "malformedXml"),
            (
                r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" cacheId="7"><location ref="A1:B2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>"#,
                "missingIdentity",
            ),
            (
                r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Bad" cacheId="7" dataCaption="Values"><location ref="not-a-range" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>"#,
                "invalidLocation",
            ),
        ] {
            let sheet = parse(&workbook_with_pivot(
                pivot,
                Some(PIVOT_RELS),
                Some(COMPLETE_CACHE),
            ));
            assert!(sheet.get("pivotTables").is_none());
            assert_eq!(sheet["pivotDiagnostics"][0]["reason"]["kind"], reason);
            assert_eq!(
                sheet["rows"][0]["cells"][0]["value"]["number"].as_f64(),
                Some(42.0)
            );
        }
    }

    #[test]
    fn unsupported_cache_source_reason_serializes_camel_case_payload() {
        for source_type in ["external", "consolidation", "futureSource"] {
            let cache =
                COMPLETE_CACHE.replace("type=\"worksheet\"", &format!("type=\"{source_type}\""));
            let sheet = parse(&workbook_with_pivot(
                COMPLETE_PIVOT,
                Some(PIVOT_RELS),
                Some(&cache),
            ));
            let reason = sheet["pivotTables"][0]["status"]["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .find(|reason| reason["kind"] == "unsupportedCacheSource")
                .unwrap();
            assert_eq!(reason["sourceType"], source_type);
            assert!(reason.get("source_type").is_none());
        }
    }

    #[test]
    fn missing_required_cache_source_is_partial() {
        let cache = COMPLETE_CACHE.replace(
            r#"<cacheSource type="worksheet"><worksheetSource ref="A1:D20" sheet="Data"/></cacheSource>"#,
            "",
        );
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&cache),
        ));
        assert_eq!(sheet["pivotTables"][0]["status"]["state"], "partial");
        assert!(sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["kind"] == "malformedCacheDefinition"));
    }

    #[test]
    fn worksheet_source_relationship_is_exposed_and_partial_until_resolved() {
        let cache = COMPLETE_CACHE.replace(
            r#"<worksheetSource ref="A1:D20" sheet="Data"/>"#,
            r#"<worksheetSource xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdExternalSheet"/>"#,
        );
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&cache),
        ));
        let pivot = &sheet["pivotTables"][0];
        assert_eq!(pivot["cacheSource"]["relationshipId"], "rIdExternalSheet");
        assert!(pivot["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["kind"] == "unresolvedWorksheetSourceRelationship"));
    }

    #[test]
    fn empty_worksheet_source_is_not_complete() {
        for replacement in ["", "<worksheetSource/>"] {
            let cache = COMPLETE_CACHE.replace(
                r#"<worksheetSource ref="A1:D20" sheet="Data"/>"#,
                replacement,
            );
            let sheet = parse(&workbook_with_pivot(
                COMPLETE_PIVOT,
                Some(PIVOT_RELS),
                Some(&cache),
            ));
            assert_eq!(sheet["pivotTables"][0]["status"]["state"], "partial");
        }
    }

    #[test]
    fn malformed_pivot_relationships_are_not_reported_as_missing() {
        for rels in [
            "<Relationships",
            r#"<Relationships xmlns="urn:not-package-relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#,
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#,
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><NotRelationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#,
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target=""/></Relationships>"#,
        ] {
            let sheet = parse(&workbook_with_pivot(COMPLETE_PIVOT, Some(rels), None));
            assert!(sheet["pivotTables"][0].get("cacheDefinitionPart").is_none());
            assert_eq!(
                sheet["pivotTables"][0]["status"]["reasons"][0]["kind"],
                "malformedCacheRelationships"
            );
        }
    }

    #[test]
    fn location_rejects_out_of_bounds_and_misplaced_dollar_markers() {
        for bad_ref in ["XFE1", "A1048577", "A$1$", "$$A1", "$A1$"] {
            let pivot = COMPLETE_PIVOT.replace("A1:D8", bad_ref);
            let sheet = parse(&workbook_with_pivot(
                &pivot,
                Some(PIVOT_RELS),
                Some(COMPLETE_CACHE),
            ));
            assert_eq!(
                sheet["pivotDiagnostics"][0]["reason"]["kind"],
                "invalidLocation"
            );
        }
        for valid_ref in ["XFD1048576", "$A$1:$D$8"] {
            let pivot = COMPLETE_PIVOT.replace("A1:D8", valid_ref);
            let sheet = parse(&workbook_with_pivot(
                &pivot,
                Some(PIVOT_RELS),
                Some(COMPLETE_CACHE),
            ));
            assert_eq!(sheet["pivotTables"][0]["name"], "SalesPivot");
        }
    }

    #[test]
    fn external_cache_relationship_is_partial_without_package_path() {
        let rels = PIVOT_RELS.replace("Target=", "TargetMode=\"External\" Target=");
        let sheet = parse(&workbook_with_pivot(COMPLETE_PIVOT, Some(&rels), None));
        let pivot = &sheet["pivotTables"][0];
        assert!(pivot.get("cacheDefinitionPart").is_none());
        assert_eq!(
            pivot["status"]["reasons"][0]["kind"],
            "externalCacheRelationship"
        );
    }

    #[test]
    fn extension_uri_collection_ignores_non_spreadsheetml_ext_elements() {
        let pivot = COMPLETE_PIVOT.replace(
            "</extLst>",
            r#"<foreign:ext xmlns:foreign="urn:foreign" uri="urn:foreign:extension"/></extLst>"#,
        );
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(
            sheet["pivotTables"][0]["extensionUris"],
            serde_json::json!(["urn:example:future-pivot-feature"])
        );
    }

    #[test]
    fn non_default_data_field_show_data_as_is_partial() {
        let pivot = COMPLETE_PIVOT.replace("fld=\"3\"", "fld=\"3\" showDataAs=\"percentOfTotal\"");
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert!(sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["feature"] == "dataField.showDataAs"));
    }

    #[test]
    fn cache_invalid_fact_is_known_only_after_valid_cache_parse() {
        for (attribute, expected) in [("", false), (" invalid=\"1\"", true)] {
            let cache = COMPLETE_CACHE.replace(
                " refreshOnLoad=\"1\"",
                &format!("{attribute} refreshOnLoad=\"1\""),
            );
            let sheet = parse(&workbook_with_pivot(
                COMPLETE_PIVOT,
                Some(PIVOT_RELS),
                Some(&cache),
            ));
            assert_eq!(sheet["pivotTables"][0]["cacheInvalid"], expected);
        }
        let malformed = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some("<bad"),
        ));
        assert!(malformed["pivotTables"][0].get("cacheInvalid").is_none());
    }

    #[test]
    fn missing_required_data_caption_skips_pivot_and_cache_fields_make_partial() {
        let no_caption = COMPLETE_PIVOT.replace(" dataCaption=\"Values\"", "");
        let sheet = parse(&workbook_with_pivot(
            &no_caption,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(
            sheet["pivotDiagnostics"][0]["reason"]["kind"],
            "missingIdentity"
        );

        let no_fields = COMPLETE_CACHE.replace(
            r#"<cacheFields count="4"><cacheField name="Category"/><cacheField name="Month"/><cacheField name="Region"/><cacheField name="Sales"/></cacheFields>"#,
            "",
        );
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&no_fields),
        ));
        assert!(sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["kind"] == "malformedCacheDefinition"));
    }

    #[test]
    fn invalid_data_subtotal_is_raw_unknown_and_partial() {
        let pivot = COMPLETE_PIVOT.replace("fld=\"3\"", "fld=\"3\" subtotal=\"notAFunction\"");
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        let field = &sheet["pivotTables"][0]["dataFields"][0];
        assert!(field.get("subtotal").is_none());
        assert_eq!(field["rawSubtotal"], "notAFunction");
        assert!(sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["field"] == "dataFields.subtotal"));

        let bad_field = COMPLETE_PIVOT.replace("fld=\"3\"", "fld=\"not-unsigned\"");
        let sheet = parse(&workbook_with_pivot(
            &bad_field,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert!(sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .any(|reason| reason["field"] == "dataFields.field"));
    }

    #[test]
    fn missing_required_cache_source_type_is_malformed_not_unsupported() {
        let cache = COMPLETE_CACHE.replace(" type=\"worksheet\"", "");
        let sheet = parse(&workbook_with_pivot(
            COMPLETE_PIVOT,
            Some(PIVOT_RELS),
            Some(&cache),
        ));
        let reasons = sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap();
        assert!(reasons
            .iter()
            .any(|reason| reason["kind"] == "malformedCacheDefinition"));
        assert!(!reasons
            .iter()
            .any(|reason| reason["kind"] == "unsupportedCacheSource"));
    }

    #[test]
    fn omitted_saved_pivot_state_structures_are_named_partial_features() {
        let pivot = COMPLETE_PIVOT.replace(
            "<rowFields",
            r#"<pivotFields count="1"><pivotField><items count="1"><item x="0"/></items></pivotField></pivotFields><rowItems count="1"><i/></rowItems><colItems count="1"><i/></colItems><rowFields"#,
        );
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        let features: Vec<_> = sheet["pivotTables"][0]["status"]["reasons"]
            .as_array()
            .unwrap()
            .iter()
            .filter_map(|reason| reason["feature"].as_str())
            .collect();
        assert!(features.contains(&"pivotFields"));
        assert!(features.contains(&"rowItems"));
        assert!(features.contains(&"colItems"));
    }

    #[test]
    fn worksheet_relationship_failures_are_diagnostic_and_keep_saved_cells() {
        for (rels, reason) in [
            (
                b"<Relationships".as_slice(),
                "malformedWorksheetRelationships",
            ),
            (&[0xff][..], "unreadableWorksheetRelationships"),
        ] {
            let sheet = parse(&workbook_with_worksheet_rels(rels));
            assert_eq!(sheet["pivotDiagnostics"][0]["reason"]["kind"], reason);
            assert_eq!(
                sheet["rows"][0]["cells"][0]["value"]["number"].as_f64(),
                Some(42.0)
            );
        }
    }

    #[test]
    fn bad_worksheet_pivot_relationship_does_not_suppress_valid_sibling() {
        let valid = r#"<Relationship Id="rIdValid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>"#;
        for bad in [
            r#"<Relationship Id="rIdBad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"/>"#,
            r#"<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/ignored.xml"/>"#,
            r#"<Relationship Id="rIdBad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" TargetMode="Bogus" Target="../pivotTables/ignored.xml"/>"#,
        ] {
            let rels = format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{bad}{valid}</Relationships>"#
            );
            let sheet = parse(&workbook_with_worksheet_rels(rels.as_bytes()));
            assert_eq!(sheet["pivotTables"].as_array().unwrap().len(), 1);
            assert_eq!(
                sheet["pivotDiagnostics"][0]["reason"]["kind"],
                "malformedPivotRelationship"
            );
        }

        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" TargetMode="External" Target="https://example.invalid/pivot.xml"/>{valid}</Relationships>"#
        );
        let sheet = parse(&workbook_with_worksheet_rels(rels.as_bytes()));
        assert_eq!(sheet["pivotTables"].as_array().unwrap().len(), 1);
        assert_eq!(
            sheet["pivotDiagnostics"][0]["reason"]["kind"],
            "externalPivotRelationship"
        );

        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMalformed" Target="ignored.xml"/>{valid}</Relationships>"#
        );
        let sheet = parse(&workbook_with_worksheet_rels(rels.as_bytes()));
        assert_eq!(sheet["pivotTables"].as_array().unwrap().len(), 1);
        assert_eq!(
            sheet["pivotDiagnostics"][0]["reason"]["kind"],
            "malformedWorksheetRelationships"
        );
    }

    #[test]
    fn exact_strict_pivot_relationship_types_resolve_and_suffix_impostors_do_not() {
        let strict_sheet_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#;
        let data = workbook_with_worksheet_rels(strict_sheet_rels);
        let sheet = parse(&data);
        assert_eq!(sheet["pivotTables"].as_array().unwrap().len(), 1);

        let impostor = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="urn:foreign/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#;
        let sheet = parse(&workbook_with_worksheet_rels(impostor));
        assert!(sheet.get("pivotTables").is_none());

        let strict_cache_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#;
        let sheet = parse(&workbook_with_relationship_parts(
            strict_sheet_rels,
            strict_cache_rels,
        ));
        assert_eq!(sheet["pivotTables"][0]["status"]["state"], "complete");

        let impostor_cache_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="urn:foreign/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition7.xml"/></Relationships>"#;
        let sheet = parse(&workbook_with_relationship_parts(
            strict_sheet_rels,
            impostor_cache_rels,
        ));
        assert_eq!(
            sheet["pivotTables"][0]["status"]["reasons"][0]["kind"],
            "missingCacheRelationship"
        );
    }

    #[test]
    fn worksheet_source_ref_uses_bounded_a1_range_grammar() {
        for valid_ref in ["$A$1:$XFD$1048576", "A1:D20"] {
            let cache = COMPLETE_CACHE.replace("A1:D20", valid_ref);
            let sheet = parse(&workbook_with_pivot(
                COMPLETE_PIVOT,
                Some(PIVOT_RELS),
                Some(&cache),
            ));
            assert_eq!(
                sheet["pivotTables"][0]["cacheSource"]["reference"],
                valid_ref
            );
        }
        for invalid_ref in ["XFE1", "A1048577", "A$1$", "D20:A1"] {
            let cache = COMPLETE_CACHE.replace("A1:D20", invalid_ref);
            let sheet = parse(&workbook_with_pivot(
                COMPLETE_PIVOT,
                Some(PIVOT_RELS),
                Some(&cache),
            ));
            let pivot = &sheet["pivotTables"][0];
            assert!(pivot["cacheSource"].get("reference").is_none());
            assert!(pivot["status"]["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|reason| reason["field"] == "cacheSource.worksheetSource.ref"));
        }
    }

    #[test]
    fn pivot_axis_data_placement_semantics_are_never_silently_complete() {
        for attributes in [" dataOnRows=\"true\"", " dataPosition=\"2\""] {
            let pivot = COMPLETE_PIVOT.replace(
                " dataCaption=\"Values\"",
                &format!(" dataCaption=\"Values\"{attributes}"),
            );
            let sheet = parse(&workbook_with_pivot(
                &pivot,
                Some(PIVOT_RELS),
                Some(COMPLETE_CACHE),
            ));
            assert!(sheet["pivotTables"][0]["status"]["reasons"]
                .as_array()
                .unwrap()
                .iter()
                .any(|reason| reason["feature"] == "pivotTable.dataPlacement"));
        }
        let pivot = COMPLETE_PIVOT.replace(
            " dataCaption=\"Values\"",
            " dataCaption=\"Values\" dataOnRows=\"false\"",
        );
        let sheet = parse(&workbook_with_pivot(
            &pivot,
            Some(PIVOT_RELS),
            Some(COMPLETE_CACHE),
        ));
        assert_eq!(sheet["pivotTables"][0]["status"]["state"], "complete");
    }
}

/// Implicit (omitted) cell/row references — ECMA-376 §18.3.1.4 (`c`) and
/// §18.3.1.73 (`row`). Both `@r` attributes are `use="optional"` in the schema
/// (CT_Cell / CT_Row, sml.xsd) with no default; that optionality is all the
/// spec mandates — it does not spell out how an omitted reference is resolved.
/// The resolution below is the de-facto consumer convention, on which Excel,
/// LibreOffice, and SheetJS agree (no competing interpretation exists):
///
///   * `<c>` without `@r` → the next column after the previous cell in the same
///     row (the first cell in a row starts at column A / 1); an explicit `@r`
///     resets the running column so subsequent omitted cells continue from it.
///   * `<row>` without `@r` → the previous row's number + 1 (the first row is 1).
///
/// Enterprise exporters (Dynamics, SAP, Oracle, SSRS) emit this minimal form to
/// shrink files; Excel/Sheets/LibreOffice/SheetJS all accept it (#832, #833-1).
#[cfg(test)]
mod implicit_reference_tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// #832: every `<c>` omits `@r`. Columns must run A, B, C per ordinal
    /// position within each row (reset at each `<row>`), not all collapse to A1.
    #[test]
    fn cells_without_r_get_sequential_columns() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c></row>
              <row r="2"><c t="s"><v>3</v></c><c t="n"><v>42.5</v></c><c t="n"><v>100</v></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        assert_eq!(ws.rows.len(), 2);

        let r1 = &ws.rows[0].cells;
        assert_eq!((r1[0].col, r1[0].row), (1, 1), "first cell of row 1 → A1");
        assert_eq!((r1[1].col, r1[1].row), (2, 1), "second cell → B1");
        assert_eq!((r1[2].col, r1[2].row), (3, 1), "third cell → C1");

        let r2 = &ws.rows[1].cells;
        assert_eq!((r2[0].col, r2[0].row), (1, 2), "first cell of row 2 → A2");
        assert_eq!((r2[1].col, r2[1].row), (2, 2), "second cell → B2");
        assert_eq!((r2[2].col, r2[2].row), (3, 2), "third cell → C2");
        match &r2[1].value {
            CellValue::Number { number } => assert_eq!(*number, 42.5),
            other => panic!("expected number 42.5 at B2, got {other:?}"),
        }
    }

    /// #833-1: every `<row>` omits `@r`. Rows must number 1, 2, 3 by document
    /// order, not all collapse to index 0.
    #[test]
    fn rows_without_r_get_sequential_indices() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row><c r="A1" t="s"><v>0</v></c></row>
              <row><c r="A2" t="s"><v>1</v></c></row>
              <row><c r="A3" t="s"><v>2</v></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        assert_eq!(ws.rows.len(), 3);
        assert_eq!(ws.rows[0].index, 1, "first <row> → 1");
        assert_eq!(ws.rows[1].index, 2, "second <row> → 2");
        assert_eq!(ws.rows[2].index, 3, "third <row> → 3");
    }

    /// Both `<row>` and `<c>` omit `@r` simultaneously (the common enterprise
    /// export shape). Positions must fill A1:C2 exactly.
    #[test]
    fn both_row_and_cell_omit_r() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c></row>
              <row><c t="s"><v>3</v></c><c t="n"><v>42.5</v></c><c t="n"><v>100</v></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        assert_eq!(ws.rows[0].index, 1);
        assert_eq!(ws.rows[1].index, 2);
        let coords: Vec<(u32, u32)> = ws
            .rows
            .iter()
            .flat_map(|r| r.cells.iter().map(|c| (c.col, c.row)))
            .collect();
        assert_eq!(
            coords,
            vec![(1, 1), (2, 1), (3, 1), (1, 2), (2, 2), (3, 2)],
            "row+cell implicit refs must fill A1:C2"
        );
    }

    /// Mixed: some cells carry an explicit `@r`, some don't. Under the de-facto
    /// convention (the spec grants only the optionality), an explicit reference
    /// re-anchors the running column, so omitted cells after it continue from
    /// that column, not from the ordinal count.
    #[test]
    fn explicit_r_reanchors_running_column() {
        // A1 (implicit) → col 1; then jump to D1 (explicit) → col 4; the next
        // implicit cell must be E1 (col 5), and the last implicit → F1 (col 6).
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row r="1"><c t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c><c t="s"><v>2</v></c><c t="s"><v>3</v></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        let cols: Vec<u32> = ws.rows[0].cells.iter().map(|c| c.col).collect();
        assert_eq!(
            cols,
            vec![1, 4, 5, 6],
            "implicit → A(1); explicit D(4) re-anchors; then E(5), F(6)"
        );
    }

    /// A `<row>` with an explicit `@r` re-anchors the running row index, so a
    /// following `<row>` without `@r` is that number + 1 (not a blind counter).
    #[test]
    fn explicit_row_r_reanchors_running_index() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row><c r="A1"><v>1</v></c></row>
              <row r="5"><c r="A5"><v>2</v></c></row>
              <row><c r="A6"><v>3</v></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("parse");
        assert_eq!(ws.rows[0].index, 1, "first implicit row → 1");
        assert_eq!(ws.rows[1].index, 5, "explicit r=5 honored");
        assert_eq!(ws.rows[2].index, 6, "implicit after r=5 → 6");
    }

    #[test]
    fn spreadsheet_ordinals_reject_zero_overflow_and_implicit_past_maximum() {
        let max_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData><row r="1048576"><c r="XFD1048576"/></row></sheetData></worksheet>"#
        );
        let (max_sheet, _) =
            parse_worksheet(&max_xml, &[], &[], "Sheet1").expect("grid maxima are valid");
        assert_eq!(max_sheet.rows[0].index, 1_048_576);
        assert_eq!(max_sheet.rows[0].cells[0].col, 16_384);

        for (body, expected) in [
            (r#"<row r="0"/>"#, "row ordinal"),
            (r#"<row r="1048577"/>"#, "row ordinal"),
            (r#"<row r="1048576"/><row/>"#, "row ordinal"),
            (r#"<row r="1"><c r="A0"/></row>"#, "row ordinal"),
            (r#"<row r="1"><c r="XFE1"/></row>"#, "column ordinal"),
            (r#"<row r="1"><c r="XFD1"/><c/></row>"#, "column ordinal"),
        ] {
            let xml =
                format!(r#"<worksheet xmlns="{NS}"><sheetData>{body}</sheetData></worksheet>"#);
            let error = parse_worksheet(&xml, &[], &[], "Sheet1")
                .expect_err("out-of-range worksheet ordinal must be a normal parse error");
            assert!(error.contains(expected), "{error}");
        }

        let mut previous = u32::MAX;
        let error = resolve_implicit_ordinal(None, &mut previous, SpreadsheetOrdinal::Row)
            .expect_err("u32 overflow must not panic or wrap");
        assert!(error.contains("overflows u32"), "{error}");

        let huge_reference = format!("{}1", "A".repeat(128));
        assert_eq!(parse_cell_ref("A"), (1, 1));
        assert_eq!(parse_cell_ref("A1x"), (1, 1));
        assert_eq!(parse_cell_ref("Afoo"), (22_037, 1));
        assert_eq!(
            parse_cell_ref(&huge_reference),
            (0, 1),
            "lenient metadata parsing preserves its fallback without overflow"
        );
        let strict_xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="{huge_reference}"/></row></sheetData></worksheet>"#
        );
        let error = parse_worksheet(&strict_xml, &[], &[], "Sheet1")
            .expect_err("strict worksheet cell references reject arithmetic overflow");
        assert!(error.contains("invalid or overflowing"), "{error}");
    }

    /// Implicit references must not disturb the other minimal-exporter
    /// constructs from #833: an inline string (`t="inlineStr"`) with rich
    /// runs and a shared-string reference resolve correctly even when `@r` is
    /// omitted on both the row and the cells.
    #[test]
    fn implicit_refs_coexist_with_inline_and_shared_strings() {
        let shared = vec![SharedString {
            text: "Shared".to_string(),
            ..Default::default()
        }];
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetData>
              <row><c t="s"><v>0</v></c><c t="inlineStr"><is><r><rPr><b/></rPr><t>Bold</t></r><r><t> tail</t></r></is></c></row>
            </sheetData></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &shared, &[], "Sheet1").expect("parse");
        let cells = &ws.rows[0].cells;
        assert_eq!((cells[0].col, cells[0].row), (1, 1));
        assert_eq!((cells[1].col, cells[1].row), (2, 1));
        match &cells[0].value {
            CellValue::Shared { si } => assert_eq!(*si, 0),
            other => panic!("expected shared ref, got {other:?}"),
        }
        match &cells[1].value {
            CellValue::Text { text, .. } => assert_eq!(text, "Bold tail"),
            other => panic!("expected concatenated inline rich text, got {other:?}"),
        }
    }
}
