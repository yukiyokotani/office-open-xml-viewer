//! BIFF8 (`.xls`) compatibility subset.
//!
//! The converter preserves worksheet names, scalar/string/boolean/error values,
//! cached formula results, merged-cell ranges, BIFF8 cell styles, shared-string
//! character formatting and geometry.
//! Formula token programs, charts, external links, and macros are never evaluated
//! or copied. Passive picture projection requires host-measured Normal-font metrics.
//! See [MS-XLS] 2.4 for record structures and 2.5.293 for BIFF8
//! Unicode strings. FILEPASS and pre-BIFF8 workbooks fail closed.

use std::collections::{BTreeMap, HashSet};

use crate::cfb::CompoundFile;
use crate::ooxml::{write_package, xml_attr, xml_text, ROOT_RELS_XLSX};

mod cell_formulas;
mod chart;
mod conditional;
pub(crate) mod direct;
#[cfg(all(test, not(target_arch = "wasm32")))]
mod direct_corpus_tests;
#[cfg(all(test, not(target_arch = "wasm32")))]
mod direct_strings_tests;
#[cfg(all(test, not(target_arch = "wasm32")))]
mod direct_styles_tests;
#[cfg(any(test, feature = "direct-xls"))]
pub(crate) mod direct_wire;
pub(crate) mod drawing_anchors;
mod drawing_media;
mod filters;
mod geometry;
mod hyperlinks;
mod names;
mod pictures;
mod print;
mod rich;
mod styles;
mod tables;
mod theme;
mod validation;
mod views;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const FILEPASS: u16 = 0x002f;
const BOUNDSHEET8: u16 = 0x0085;
const SST: u16 = 0x00fc;
const CONTINUE: u16 = 0x003c;
const NUMBER: u16 = 0x0203;
const RK: u16 = 0x027e;
const MULRK: u16 = 0x00bd;
const LABELSST: u16 = 0x00fd;
const LABEL: u16 = 0x0204;
const BOOLERR: u16 = 0x0205;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const MERGEDCELLS: u16 = 0x00e5;
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const MAX_RECORDS: usize = 2_000_000;
const MAX_SHEETS: usize = 65_536;
const MAX_CELLS: usize = 10_000_000;

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub(crate) fn inspect_images(
    cfb: &CompoundFile<'_>,
) -> Result<Vec<(u32, &'static str, Vec<u8>)>, String> {
    let workbook = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(unsupported)?;
    drawing_media::images(&records(&workbook)?)
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub(crate) fn inspect_anchors(
    cfb: &CompoundFile<'_>,
) -> Result<Vec<drawing_anchors::DrawingAnchor>, String> {
    let workbook = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(unsupported)?;
    drawing_anchors::workbook(&records(&workbook)?)
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub(crate) fn inspect_pictures(
    cfb: &CompoundFile<'_>,
) -> Result<crate::XlsPictureInspection, String> {
    let workbook = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(unsupported)?;
    let records = records(&workbook)?;
    let mut anchors = drawing_anchors::workbook(&records)?;
    let indices = anchors
        .iter()
        .filter_map(|a| a.picture.map(|p| p.store_index))
        .collect();
    let images = drawing_media::selected(&records, &indices)?;
    let supported: HashSet<u32> = images.iter().map(|i| i.0).collect();
    anchors.retain(|a| {
        a.picture
            .is_some_and(|p| supported.contains(&p.store_index))
    });
    Ok(crate::XlsPictureInspection { anchors, images })
}

pub struct XlsConversion {
    pub bytes: Vec<u8>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone)]
struct BoundSheet {
    offset: usize,
    name: String,
    sheet_type: u8,
    visibility: SheetVisibility,
}

#[derive(Debug, Clone, Copy, Default)]
enum SheetVisibility {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

impl SheetVisibility {
    fn attribute(self) -> &'static str {
        // ECMA-376 18.2.19: omitted sheet/@state means visible.
        match self {
            Self::Visible => "",
            Self::Hidden => " state=\"hidden\"",
            Self::VeryHidden => " state=\"veryHidden\"",
        }
    }
}

#[derive(Debug, Clone)]
enum CellValue {
    Blank,
    Number(f64),
    Text(String),
    SharedString(usize),
    Bool(bool),
    Error(String),
}

#[derive(Default)]
struct SheetData {
    visibility: SheetVisibility,
    rows: BTreeMap<u16, BTreeMap<u16, CellValue>>,
    cell_styles: BTreeMap<(u16, u16), u16>,
    geometry: geometry::Geometry,
    print: print::PrintSettings,
    views: views::SheetViews,
    merged: Vec<(u16, u16, u16, u16)>,
    formula_results: bool,
    custom_views_omitted: bool,
    /// MS-XLS 2.4.56 CondFmt, 2.4.42 CF, 2.4.57 CondFmt12, 2.4.43 CF12 and
    /// 2.4.44 CFEx records of the worksheet substream, in stream order.
    conditional_records: conditional::Records,
    /// Their XLSX-model projection (direct path only).
    conditional_formats: Vec<xlsx_model::ConditionalFormat>,
    /// MS-XLS 2.4.113 FeatHdr11, 2.4.114 Feature11, 2.4.115 Feature12 and
    /// 2.4.157 List12 records (tables), in stream order.
    table_records: tables::Records,
    /// Their XLSX-model projection (direct path only).
    tables: Vec<xlsx_model::TableInfo>,
    /// MS-XLS 2.4.259 SheetExt (tab color), if present.
    sheet_ext: Option<Vec<u8>>,
    /// Its resolved tab color (direct path only).
    tab_color: Option<String>,
    /// MS-XLS 2.4.140 HLink and 2.4.141 HLinkTooltip records, in order.
    hyperlink_records: tables::Records,
    /// Their XLSX-model projection (direct path only).
    hyperlinks: Vec<xlsx_model::Hyperlink>,
    /// MS-XLS 2.4.8 AutoFilterInfo: the sheet has an AutoFilter.
    autofilter_info: bool,
    /// Its `_FilterDatabase` range (direct path only).
    auto_filter: Option<xlsx_model::CellRange>,
    /// MS-XLS 2.4.96 DVal and 2.4.95 Dv records, in order.
    validation_records: tables::Records,
    /// Their XLSX-model projection (direct path only).
    data_validations: Vec<xlsx_model::DataValidation>,
    /// Defined names visible on this sheet (direct path only).
    defined_names: Vec<xlsx_model::DefinedName>,
    /// A cell, row or column shows phonetic guides (MS-XLS 2.4.192
    /// PhoneticInfo sqref, ROW/COLINFO fPhonetic).
    shows_phonetic: bool,
    /// FORMULA, SHRFMLA, ARRAY and TABLE records, in order.
    formula_records: tables::Records,
    /// Formula text by cell (direct path only).
    formulas: BTreeMap<(u16, u16), String>,
}

pub fn convert(cfb: &CompoundFile<'_>, max_output_bytes: usize) -> Result<XlsConversion, String> {
    prepare(cfb, false)?.finish(max_output_bytes, None)
}

/// Owned parse result: no CFB/BIFF slices survive preparation. This boundary
/// permits host font measurement without retaining or parsing the source again.
pub(crate) struct PreparedXls {
    sheets: Vec<(String, SheetData)>,
    styles: styles::ResolvedStyleSheet,
    shared_strings: Vec<rich::Text>,
    date1904: bool,
    window_count: usize,
    warnings: Vec<String>,
    pub(crate) font: Option<styles::NormalFont>,
    pictures: pictures::Pictures,
    charts: chart::Charts,
}

impl PreparedXls {
    pub(crate) fn finish(
        mut self,
        max_output_bytes: usize,
        mdw: Option<f64>,
    ) -> Result<XlsConversion, String> {
        if mdw.is_some_and(|v| !v.is_finite() || v.fract() != 0.0 || !(1.0..=4096.0).contains(&v)) {
            return Err(unsupported("invalid measured XLS maximum digit width"));
        }
        let mdw = mdw.filter(|_| self.font.is_some());
        if mdw.is_none() && !self.pictures.is_empty() {
            self.warnings
                .push("legacy-xls:unmeasured-pictures-omitted".into());
        }
        let drawings = mdw.map(|m| self.pictures.emit(&self.sheets, m, &mut self.warnings));
        if drawings.as_ref().is_some_and(|d| !d.sheets.is_empty()) {
            self.warnings[0] =
                "legacy-xls:unsupported-drawings-conditional-formatting-and-external-links-omitted"
                    .into();
        }
        let bytes = build_xlsx_with_drawings(
            &self.sheets,
            &self.styles,
            self.shared_strings,
            self.date1904,
            self.window_count,
            max_output_bytes,
            mdw,
            drawings.as_ref(),
        )?;
        Ok(XlsConversion {
            bytes,
            warnings: self.warnings,
        })
    }
}

pub(crate) fn prepare(cfb: &CompoundFile<'_>, with_pictures: bool) -> Result<PreparedXls, String> {
    let workbook = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(unsupported)?;
    prepare_workbook(&workbook, with_pictures, false)
}

fn prepare_workbook(
    workbook: &[u8],
    with_pictures: bool,
    direct: bool,
) -> Result<PreparedXls, String> {
    let records = records(workbook)?;
    if direct {
        validate_direct_retention(&records)?;
    }
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF workbook"))?;
    if first.kind != BOF
        || u16_at(first.data, 0)? != BIFF8
        || u16_at(first.data, 2)? != WORKBOOK_GLOBALS
    {
        return Err(unsupported("only BIFF8 workbook globals are supported"));
    }
    if records.iter().any(|record| record.kind == FILEPASS) {
        return Err(unsupported("encrypted BIFF workbooks are not supported"));
    }

    let styles = styles::Styles::parse(&records)?;
    let mut date1904 = false;
    let mut window_count = 0;
    let mut sheets = Vec::new();
    let mut shared_strings = Vec::new();
    let mut saw_sst = false;
    for (index, record) in records.iter().enumerate() {
        if record.kind == EOF {
            break;
        }
        match record.kind {
            views::WINDOW1 => views::read_window(record.data, &mut window_count)?,
            0x0022 => date1904 = u16_at(record.data, 0)? != 0,
            BOUNDSHEET8 => {
                if sheets.len() >= MAX_SHEETS {
                    return Err(unsupported("too many BIFF worksheets"));
                }
                sheets.push(parse_bound_sheet(record.data)?);
            }
            SST => {
                if saw_sst {
                    return Err(unsupported("multiple BIFF shared string tables"));
                }
                let mut fragments = vec![record.data];
                let mut continued = index + 1;
                while let Some(next) = records.get(continued) {
                    if next.kind != CONTINUE {
                        break;
                    }
                    fragments.push(next.data);
                    continued += 1;
                }
                shared_strings = parse_sst_elements(&fragments)?;
                saw_sst = true;
            }
            _ => {}
        }
    }
    if sheets.is_empty() {
        return Err(unsupported("BIFF workbook contains no sheets"));
    }

    let mut sheet_offsets = HashSet::with_capacity(sheets.len());
    if sheets
        .iter()
        .any(|sheet| !sheet_offsets.insert(sheet.offset))
    {
        return Err(unsupported("duplicate BIFF worksheet offsets"));
    }

    // Validate global rich-text references before allocating worksheet cells.
    // This replaces the old SST encoder's eager FontIndex validation without
    // making XML generation part of source admission.
    let mut resolved_styles = styles.resolve()?;
    for text in &shared_strings {
        text.validate_fonts(&resolved_styles)?;
    }

    let mut converted = Vec::new();
    let mut skipped_non_worksheets = false;
    let mut formula_results = false;
    let mut incomplete_print_margins = false;
    let mut custom_views_omitted = false;
    let mut tabs = Vec::new();
    let mut conditional_theme = None;
    let mut dxfs = Vec::new();
    let mut table_styles = None;
    let mut filter_databases = None;
    let mut defined_names = None;
    let has_names = records
        .iter()
        .take_while(|record| record.kind != EOF)
        .any(|record| record.kind == 0x0018);
    for (tab, sheet) in sheets.into_iter().enumerate() {
        if sheet.sheet_type != 0 {
            // MS-XLS 2.4.28 BoundSheet8.dt: 1 macro sheet, 2 chart sheet, 6 VB
            // module. A VB module has no sheet content to display; the direct
            // reader rejects macro and chart sheets rather than dropping tabs
            // whose cells or chart Excel shows.
            if direct && sheet.sheet_type != 6 {
                return Err(unsupported(match sheet.sheet_type {
                    1 => "XLS macro sheets are not projected",
                    2 => "XLS chart sheets are not projected",
                    _ => "unknown XLS sheet type",
                }));
            }
            skipped_non_worksheets = true;
            continue;
        }
        let mut data = parse_sheet(&records, &sheet, &shared_strings)?;
        // Only the direct model projects conditional formatting; the byte
        // converter keeps its documented omission warning.
        if direct {
            if let Some(ext) = data.sheet_ext.as_deref() {
                if conditional_theme.is_none() {
                    conditional_theme = Some((
                        theme::Colors::parse(&records)?,
                        conditional::Externs::parse(&records)?,
                    ));
                }
                let (theme, _) = conditional_theme.as_ref().expect("parsed theme");
                data.tab_color = tab_color(ext, &styles, theme)?;
            }
        }
        if direct {
            for (kind, record) in data.hyperlink_records.iter() {
                if kind == 0x01b8 {
                    data.hyperlinks.push(hyperlinks::hlink(record)?);
                } else {
                    hyperlinks::tooltip(record)?;
                }
            }
        }
        // Phonetic guides (ExtRst runs in the shared strings) are not
        // projected; a sheet that displays them fails closed.
        if direct && data.shows_phonetic {
            return Err(unsupported("XLS phonetic guides are not projected"));
        }
        if direct && !data.formula_records.is_empty() {
            if conditional_theme.is_none() {
                conditional_theme = Some((
                    theme::Colors::parse(&records)?,
                    conditional::Externs::parse(&records)?,
                ));
            }
            let (_, externs) = conditional_theme.as_ref().expect("parsed theme");
            data.formulas = cell_formulas::project(data.formula_records.iter(), externs)?;
        }
        if direct && has_names {
            if conditional_theme.is_none() {
                conditional_theme = Some((
                    theme::Colors::parse(&records)?,
                    conditional::Externs::parse(&records)?,
                ));
            }
            if defined_names.is_none() {
                let (_, externs) = conditional_theme.as_ref().expect("parsed theme");
                defined_names = Some(names::Names::parse(&records, externs)?);
            }
            data.defined_names = defined_names.as_ref().expect("parsed names").for_sheet(tab);
        }
        if direct && !data.validation_records.is_empty() {
            if conditional_theme.is_none() {
                conditional_theme = Some((
                    theme::Colors::parse(&records)?,
                    conditional::Externs::parse(&records)?,
                ));
            }
            let (_, externs) = conditional_theme.as_ref().expect("parsed theme");
            data.data_validations = validation::project(data.validation_records.iter(), externs)?;
        }
        if direct && data.autofilter_info {
            if filter_databases.is_none() {
                filter_databases = Some(filters::Databases::parse(&records)?);
            }
            let range = filter_databases
                .as_ref()
                .expect("parsed filter databases")
                .range(tab)
                .ok_or_else(|| unsupported("BIFF AutoFilter lacks its filter database"))?;
            data.auto_filter = Some(xlsx_model::CellRange { ..*range });
        }
        if direct && !data.table_records.is_empty() {
            if conditional_theme.is_none() {
                conditional_theme = Some((
                    theme::Colors::parse(&records)?,
                    conditional::Externs::parse(&records)?,
                ));
            }
            if table_styles.is_none() {
                table_styles = Some(tables::Styles::parse(&records)?);
            }
            let (theme, _) = conditional_theme.as_ref().expect("parsed theme");
            let context = tables::Context {
                styles: &styles,
                theme,
            };
            data.tables = tables::project(
                &data.table_records,
                table_styles.as_mut().expect("parsed table styles"),
                &context,
                &mut dxfs,
            )?;
        }
        if direct && !data.conditional_records.is_empty() {
            if conditional_theme.is_none() {
                conditional_theme = Some((
                    theme::Colors::parse(&records)?,
                    conditional::Externs::parse(&records)?,
                ));
            }
            let (theme, externs) = conditional_theme.as_ref().expect("parsed theme");
            let context = conditional::Context {
                styles: &styles,
                theme,
                externs,
            };
            data.conditional_formats =
                conditional::project(&data.conditional_records, &context, &mut dxfs)?;
        }
        data.views.validate_count(window_count)?;
        for index in data.cell_styles.values() {
            styles.validate_xf(*index)?;
        }
        data.geometry.validate_styles(&styles)?;
        incomplete_print_margins |= data.print.incomplete_margins();
        custom_views_omitted |= data.custom_views_omitted;
        formula_results |= data.formula_results;
        converted.push((sheet.name, data));
        tabs.push(tab);
    }
    if converted.is_empty() {
        return Err(unsupported(
            "BIFF workbook contains no supported worksheets",
        ));
    }
    resolved_styles.set_dxfs(dxfs);
    // The byte converter omits drawings and conditional formatting and says
    // so. The direct reader projects pictures, charts and conditional
    // formatting or rejects the workbook (below), and never evaluates external
    // links, whose cached cell values it shows, so it makes no such claim.
    let mut warnings: Vec<String> = if direct {
        Vec::new()
    } else {
        vec!["legacy-xls:drawings-conditional-formatting-and-external-links-omitted".into()]
    };
    // The direct model carries print areas and titles as defined names;
    // page setup, headers and footers (including 2.4.136 HeaderFooter) only
    // affect printing, which neither the XLSX model nor its viewer has.
    // Phonetic strings (ExtRst) display only in cells marked by
    // PhoneticInfo, ROW or COLINFO, which the direct path rejects instead.
    if !direct {
        warnings.push(
            "legacy-xls:phonetic-data-print-areas-titles-and-extended-headers-omitted".into(),
        );
    }
    // The direct model projects XFExt colors, indentation and gradient
    // fills; StyleExt (2.4.270) only extends the cell-style gallery entries,
    // which cells reach through their XFs. Anything else fails closed.
    if direct && styles.extensions_unrepresented() {
        return Err(unsupported(
            "XLS extended cell formatting is not representable",
        ));
    }
    if styles.extensions_omitted && !direct {
        warnings.push("legacy-xls:extended-styles-omitted".into());
    }
    if incomplete_print_margins {
        warnings.push("legacy-xls:incomplete-print-margins-omitted".into());
    }
    if custom_views_omitted {
        warnings.push("legacy-xls:saved-custom-views-omitted".into());
    }
    // The direct model carries formula text with the cached results, as the
    // XLSX model does, so only the byte converter replaces formulas.
    if formula_results && !direct {
        warnings.push("legacy-xls:formulas-replaced-with-cached-results".into());
    }
    if skipped_non_worksheets {
        warnings.push("legacy-xls:non-worksheet-tabs-omitted".into());
    }
    if with_pictures && direct {
        validate_direct_drawings(&records, &tabs)?;
    }
    let pictures = if with_pictures {
        match pictures::Pictures::prepare(&records, &tabs) {
            Ok(value) => {
                if value.has_unsupported_images() {
                    if direct {
                        return Err(unsupported("BIFF picture BLIP is not a supported image"));
                    }
                    warnings.push("legacy-xls:invalid-or-unsupported-pictures-omitted".into());
                }
                value
            }
            Err(error) if direct => return Err(format!("{error} (XLS picture)")),
            Err(_) => {
                // Optional passive content fails closed without discarding
                // otherwise valid cells. Never copy the rejected image bytes.
                warnings.push("legacy-xls:invalid-or-unsupported-pictures-omitted".into());
                pictures::Pictures::default()
            }
        }
    } else {
        pictures::Pictures::default()
    };
    // Only the direct model projects charts; the byte converter keeps its
    // documented drawing omission.
    let charts = if with_pictures && direct {
        chart::Charts::prepare(&records, &tabs, &styles, &converted, &shared_strings)?
    } else {
        chart::Charts::default()
    };
    let font = if with_pictures && (!pictures.is_empty() || !charts.is_empty()) {
        styles.normal_font()
    } else {
        None
    };
    Ok(PreparedXls {
        sheets: converted,
        styles: resolved_styles,
        shared_strings,
        date1904,
        window_count,
        warnings,
        font,
        pictures,
        charts,
    })
}

/// The direct reader must not drop drawn content: every sheet-anchored
/// drawing object on a projected worksheet is an embedded chart (MS-XLS
/// 2.4.181 ot 5 with its chart substream) or a picture (ot 8 with a BLIP
/// reference); any other object (lines, shapes, text boxes, controls, cell
/// comments), a grouped shape, or a chart/picture without its data is rejected
/// with the reason.
fn validate_direct_drawings(records: &[Record<'_>], tabs: &[usize]) -> Result<(), String> {
    let projected: std::collections::BTreeSet<_> = tabs.iter().copied().collect();
    for anchor in drawing_anchors::strict(records)? {
        if projected.contains(&anchor.sheet) {
            admit_direct_object(&anchor)?;
        }
    }
    Ok(())
}

fn admit_direct_object(anchor: &drawing_anchors::DrawingAnchor) -> Result<(), String> {
    {
        match anchor.object_type {
            5 if anchor.chart.is_some() => {}
            5 => return Err(unsupported("BIFF chart object without its chart substream")),
            8 if anchor.picture.is_some() => {}
            8 => return Err(unsupported("BIFF picture object without a BLIP reference")),
            25 => return Err(unsupported("BIFF cell comments are not projected")),
            kind => {
                return Err(unsupported(format!(
                    "BIFF drawing object type {kind} is not projected"
                )))
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod direct_drawing_tests {
    use super::drawing_anchors::{CellCorner, DrawingAnchor};

    #[test]
    fn direct_objects_are_charts_with_data_or_referenced_pictures() {
        let corner = CellCorner {
            column: 0,
            row: 0,
            dx: 0,
            dy: 0,
        };
        let anchor = |object_type, chart, picture| DrawingAnchor {
            sheet: 0,
            shape_id: 1,
            shape_flags: 0,
            object_id: 1,
            object_type,
            object_flags: 0,
            group_depth: 1,
            behavior: 0,
            from: corner,
            to: corner,
            picture,
            chart,
        };
        let picture = || {
            Some(super::drawing_anchors::PictureReference {
                store_index: 1,
                crop: [0; 4],
                rotation: 0,
                clipboard_format: 0,
                auto_picture: false,
            })
        };
        assert!(super::admit_direct_object(&anchor(5, Some((1, 2)), None)).is_ok());
        assert!(super::admit_direct_object(&anchor(8, None, picture())).is_ok());
        for (kind, chart, expected) in [
            (5, None, "without its chart substream"),
            (8, None, "without a BLIP reference"),
            (25, None, "cell comments"),
            (2, None, "type 2 is not projected"),
            (6, None, "type 6 is not projected"),
        ] {
            let error = super::admit_direct_object(&anchor(kind, chart, None)).unwrap_err();
            assert!(error.contains(expected), "{expected}: {error}");
        }
    }
}

/// MS-XLS 2.4.259 SheetExt: icvPlain (0x7F = no color), refined by the
/// SheetExtOptional CFColor when its icvPlain12 agrees with icvPlain.
fn tab_color(
    data: &[u8],
    styles: &styles::Styles<'_>,
    theme: &theme::Colors,
) -> Result<Option<String>, String> {
    let size = u32_at(data, 12)?;
    if u16_at(data, 0)? != 0x0862 || !matches!((size, data.len()), (20, 20) | (40, 40)) {
        return Err(unsupported("invalid BIFF sheet extension"));
    }
    let icv = (u32_at(data, 16)? & 0x7f) as u16;
    if icv == 0x7f {
        return Ok(None);
    }
    if size == 40 && (u32_at(data, 20)? & 0x7f) as u16 == icv {
        return conditional::cf_color(data, 24, styles, theme).map(Some);
    }
    Ok(Some(styles.chart_color(icv).ok_or_else(|| {
        unsupported("invalid BIFF sheet tab color")
    })?))
}

fn prepare_direct(cfb: &CompoundFile<'_>) -> Result<PreparedXls, String> {
    const MAX_DIRECT_WORKBOOK_BYTES: usize = 256 * 1024 * 1024;
    let streams = cfb.scoped_streams().map_err(unsupported)?;
    let workbook = match streams.optional_stream(&["Workbook"], MAX_DIRECT_WORKBOOK_BYTES) {
        Ok(Some(value)) => value,
        Ok(None) => streams
            .stream(&["Book"], MAX_DIRECT_WORKBOOK_BYTES)
            .map_err(unsupported)?,
        Err(error) => return Err(unsupported(error)),
    };
    prepare_workbook(&workbook, true, true)
}

fn validate_direct_retention(records: &[Record<'_>]) -> Result<(), String> {
    const MAX_DIRECT_NEUTRAL_BYTES: usize = 256 * 1024 * 1024;
    const MAX_DIRECT_CELLS: usize = 1_000_000;
    const MAX_DIRECT_ROWS: usize = 1_000_000;
    let mut cells = 0usize;
    let mut rows = 0usize;
    let mut decoded_text_bytes = 0usize;
    for record in records {
        match record.kind {
            NUMBER | RK | LABELSST | LABEL | BOOLERR | FORMULA | 0x0201 => {
                cells = cells.saturating_add(1)
            }
            MULRK => cells = cells.saturating_add(record.data.len().saturating_sub(6) / 6),
            0x00be => cells = cells.saturating_add(record.data.len().saturating_sub(6) / 2),
            0x0208 => rows = rows.saturating_add(1),
            _ => {}
        }
        let chars = match record.kind {
            LABEL => record.data.get(6..8),
            STRING => record.data.get(..2),
            _ => None,
        };
        if let Some(chars) = chars {
            let units = usize::from(u16::from_le_bytes([chars[0], chars[1]]));
            decoded_text_bytes = decoded_text_bytes.saturating_add(units.saturating_mul(3));
        }
    }
    let possible_rows = rows.saturating_add(cells);
    if cells > MAX_DIRECT_CELLS || possible_rows > MAX_DIRECT_ROWS {
        return Err(unsupported("XLS direct neutral entry-count limit exceeded"));
    }
    // Logical payload accounting only; allocator bookkeeping is intentionally
    // excluded and independently bounded by the entry-count limits above.
    // UTF-16 source units may expand to three UTF-8 bytes; count that decoded
    // upper bound separately before allocating LABEL/formula STRING values.
    let cell_bytes = std::mem::size_of::<u16>()
        + std::mem::size_of::<CellValue>()
        + std::mem::size_of::<((u16, u16), u16)>();
    let row_bytes = std::mem::size_of::<u16>() + std::mem::size_of::<BTreeMap<u16, CellValue>>();
    let retained = cells
        .checked_mul(cell_bytes)
        .and_then(|value| value.checked_add(possible_rows.checked_mul(row_bytes)?))
        .and_then(|value| value.checked_add(decoded_text_bytes))
        .ok_or_else(|| unsupported("XLS direct neutral retention size overflow"))?;
    if retained > MAX_DIRECT_NEUTRAL_BYTES {
        return Err(unsupported("XLS direct neutral retention budget exceeded"));
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct Record<'a> {
    kind: u16,
    offset: usize,
    data: &'a [u8],
}

fn records(bytes: &[u8]) -> Result<Vec<Record<'_>>, String> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        if output.len() >= MAX_RECORDS {
            return Err(unsupported("too many BIFF records"));
        }
        if bytes.len() - offset < 4 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(output)
            } else {
                Err(unsupported("truncated BIFF record header"))
            };
        }
        let kind = u16_at(bytes, offset)?;
        let size = u16_at(bytes, offset + 2)? as usize;
        if kind == 0 && size == 0 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(output)
            } else {
                Err(unsupported("unexpected zero BIFF record"))
            };
        }
        let end = offset
            .checked_add(4 + size)
            .ok_or_else(|| unsupported("BIFF record range overflow"))?;
        let data = bytes
            .get(offset + 4..end)
            .ok_or_else(|| unsupported("truncated BIFF record"))?;
        output.push(Record { kind, offset, data });
        offset = end;
    }
    Ok(output)
}

fn parse_bound_sheet(data: &[u8]) -> Result<BoundSheet, String> {
    if data.len() < 8 {
        return Err(unsupported("truncated BOUNDSHEET8 record"));
    }
    let offset = usize::try_from(u32_at(data, 0)?)
        .map_err(|_| unsupported("BIFF sheet offset is too large"))?;
    let sheet_type = data[5];
    // MS-XLS 2.4.28: hsState occupies two bits; the other six MUST be ignored.
    let visibility = match data[4] & 3 {
        0 => SheetVisibility::Visible,
        1 => SheetVisibility::Hidden,
        2 => SheetVisibility::VeryHidden,
        _ => return Err(unsupported("invalid BIFF sheet visibility")),
    };
    let chars = data[6] as usize;
    let high_byte = (data[7] & 0x01) != 0;
    let name = decode_biff_chars(data, 8, chars, high_byte)?.0;
    if name.is_empty() {
        return Err(unsupported("empty BIFF sheet name"));
    }
    Ok(BoundSheet {
        offset,
        name,
        sheet_type,
        visibility,
    })
}

#[cfg(test)]
fn parse_sst(fragments: &[&[u8]]) -> Result<Vec<String>, String> {
    Ok(parse_sst_elements(fragments)?
        .into_iter()
        .map(|s| s.text)
        .collect())
}

fn parse_sst_elements(fragments: &[&[u8]]) -> Result<Vec<rich::Text>, String> {
    let total_bytes = fragments.iter().try_fold(0usize, |total, fragment| {
        total
            .checked_add(fragment.len())
            .ok_or_else(|| unsupported("BIFF shared string table size overflow"))
    })?;
    let mut cursor = SstCursor::new(fragments);
    let counts = cursor.read_fixed(8, "truncated SST record")?;
    let unique = usize::try_from(u32_at(counts, 4)?)
        .map_err(|_| unsupported("BIFF shared string count is too large"))?;
    // Resource policy, separate from BIFF's cell/record limits. Retain only the
    // neutral text/run table here; XML or model expansion is route-local later.
    if unique > 1_000_000 || unique > total_bytes.saturating_sub(8) / 3 {
        return Err(unsupported("too many BIFF shared strings"));
    }
    let mut retention_budget = rich::MAX_MODEL_BYTES;
    retention_budget = retention_budget
        .checked_sub(
            unique
                .checked_mul(std::mem::size_of::<rich::Text>())
                .ok_or_else(|| unsupported("BIFF shared string retention size overflow"))?,
        )
        .ok_or_else(|| unsupported("BIFF shared string retention budget exceeded"))?;
    let mut strings = Vec::new();
    strings
        .try_reserve_exact(unique)
        .map_err(|_| unsupported("BIFF shared string retention allocation failed"))?;
    let extra_slots = strings.capacity().saturating_sub(unique);
    retention_budget = retention_budget
        .checked_sub(
            extra_slots
                .checked_mul(std::mem::size_of::<rich::Text>())
                .ok_or_else(|| unsupported("BIFF shared string retention size overflow"))?,
        )
        .ok_or_else(|| unsupported("BIFF shared string retention budget exceeded"))?;
    let mut run_budget = 1_000_000usize;
    for _ in 0..unique {
        let header = cursor.read_fixed(3, "split or truncated BIFF string header")?;
        let header_fragment = cursor.fragment;
        let chars = u16_at(header, 0)? as usize;
        let flags = header[2];
        let rich_runs = if (flags & 0x08) != 0 {
            let count = cursor.read_fixed(2, "split or truncated BIFF rich string header")?;
            u16_at(count, 0)? as usize
        } else {
            0
        };
        let ext_size = if (flags & 0x04) != 0 {
            let count = cursor.read_fixed(4, "split or truncated BIFF extended string header")?;
            usize::try_from(u32_at(count, 0)?)
                .map_err(|_| unsupported("BIFF extended string data is too large"))?
        } else {
            0
        };
        if cursor.fragment != header_fragment || ext_size > i32::MAX as usize {
            return Err(unsupported(
                "split BIFF string header or negative extension size",
            ));
        }
        run_budget = run_budget
            .checked_sub(rich_runs)
            .ok_or_else(|| unsupported("BIFF rich-text run budget exceeded"))?;
        let units = cursor.read_characters(chars, (flags & 0x01) != 0)?;
        let mut runs = Vec::with_capacity(rich_runs);
        for _ in 0..rich_runs {
            // Unlike continued character data, formatting bytes have no
            // compression option. Keep the variable-field cursor separate.
            let bytes = cursor.read_variable_four()?;
            runs.push((u16_at(&bytes, 0)?, u16_at(&bytes, 2)?));
        }
        cursor.skip_variable(ext_size)?;
        let value = rich::Text::new(&units, &runs)?;
        retention_budget = retention_budget
            .checked_sub(value.retained_bytes()?)
            .ok_or_else(|| unsupported("BIFF shared string retention budget exceeded"))?;
        strings.push(value);
    }
    Ok(strings)
}

struct SstCursor<'a, 'b> {
    fragments: &'a [&'b [u8]],
    fragment: usize,
    offset: usize,
}

impl<'a, 'b> SstCursor<'a, 'b> {
    fn new(fragments: &'a [&'b [u8]]) -> Self {
        Self {
            fragments,
            fragment: 0,
            offset: 0,
        }
    }

    /// SST non-variable fields cannot straddle BIFF record boundaries.
    fn read_fixed(&mut self, count: usize, message: &str) -> Result<&'b [u8], String> {
        self.advance_empty_fragments();
        let fragment = self
            .fragments
            .get(self.fragment)
            .ok_or_else(|| unsupported(message))?;
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| unsupported("BIFF shared string offset overflow"))?;
        let bytes = fragment
            .get(self.offset..end)
            .ok_or_else(|| unsupported(message))?;
        self.offset = end;
        Ok(bytes)
    }

    /// [MS-XLS] 2.5.293 permits the character array to cross CONTINUE records.
    /// Each continued character fragment starts with its own compression flag.
    fn read_characters(
        &mut self,
        mut remaining: usize,
        mut high_byte: bool,
    ) -> Result<Vec<u16>, String> {
        let mut output = Vec::with_capacity(remaining);
        while remaining > 0 {
            if self.current_fragment_exhausted() {
                self.fragment += 1;
                self.offset = 0;
                let option = *self
                    .fragments
                    .get(self.fragment)
                    .and_then(|fragment| fragment.first())
                    .ok_or_else(|| unsupported("truncated continued BIFF string"))?;
                if option > 1 {
                    return Err(unsupported(
                        "invalid continued BIFF string compression flag",
                    ));
                }
                high_byte = option == 1;
                self.offset = 1;
            }

            let fragment = self
                .fragments
                .get(self.fragment)
                .ok_or_else(|| unsupported("truncated continued BIFF string"))?;
            let width = if high_byte { 2 } else { 1 };
            let available_bytes = fragment.len() - self.offset;
            if high_byte && remaining > available_bytes / 2 && !available_bytes.is_multiple_of(2) {
                return Err(unsupported(
                    "continued BIFF Unicode string splits a double-byte character",
                ));
            }
            let take = remaining.min(available_bytes / width);
            if take == 0 {
                return Err(unsupported("empty continued BIFF string fragment"));
            }
            let byte_count = take * width;
            let bytes = &fragment[self.offset..self.offset + byte_count];
            if high_byte {
                let units = bytes
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                output.extend(units);
            } else {
                output.extend(bytes.iter().map(|byte| u16::from(*byte)));
            }
            self.offset += byte_count;
            remaining -= take;
        }
        Ok(output)
    }

    fn read_variable_four(&mut self) -> Result<[u8; 4], String> {
        let mut result = [0; 4];
        for byte in &mut result {
            *byte = self.read_fixed(1, "truncated BIFF format run")?[0];
        }
        Ok(result)
    }

    fn skip_variable(&mut self, mut remaining: usize) -> Result<(), String> {
        while remaining > 0 {
            self.advance_empty_fragments();
            let fragment = self
                .fragments
                .get(self.fragment)
                .ok_or_else(|| unsupported("truncated BIFF rich or extended string data"))?;
            let take = remaining.min(fragment.len() - self.offset);
            self.offset += take;
            remaining -= take;
        }
        Ok(())
    }

    fn advance_empty_fragments(&mut self) {
        while self.current_fragment_exhausted() && self.fragment < self.fragments.len() {
            self.fragment += 1;
            self.offset = 0;
        }
    }

    fn current_fragment_exhausted(&self) -> bool {
        self.fragments
            .get(self.fragment)
            .is_none_or(|fragment| self.offset == fragment.len())
    }
}

fn parse_biff_string(data: &[u8]) -> Result<(String, usize), String> {
    let chars = u16_at(data, 0)? as usize;
    let flags = *data
        .get(2)
        .ok_or_else(|| unsupported("truncated BIFF string flags"))?;
    let high_byte = (flags & 0x01) != 0;
    let rich_runs = if (flags & 0x08) != 0 {
        u16_at(data, 3)? as usize
    } else {
        0
    };
    let ext_size_offset = 3 + usize::from((flags & 0x08) != 0) * 2;
    let ext_size = if (flags & 0x04) != 0 {
        usize::try_from(u32_at(data, ext_size_offset)?)
            .map_err(|_| unsupported("BIFF extended string data is too large"))?
    } else {
        0
    };
    let chars_offset = ext_size_offset + usize::from((flags & 0x04) != 0) * 4;
    let (value, char_bytes) = decode_biff_chars(data, chars_offset, chars, high_byte)?;
    let consumed = chars_offset
        .checked_add(char_bytes)
        .and_then(|value| value.checked_add(rich_runs * 4))
        .and_then(|value| value.checked_add(ext_size))
        .ok_or_else(|| unsupported("BIFF string size overflow"))?;
    if consumed > data.len() {
        return Err(unsupported("truncated BIFF rich or extended string data"));
    }
    Ok((value, consumed))
}

fn decode_biff_chars(
    data: &[u8],
    offset: usize,
    chars: usize,
    high_byte: bool,
) -> Result<(String, usize), String> {
    let byte_count = chars
        .checked_mul(if high_byte { 2 } else { 1 })
        .ok_or_else(|| unsupported("BIFF string size overflow"))?;
    let bytes = data
        .get(offset..offset + byte_count)
        .ok_or_else(|| unsupported("truncated BIFF string"))?;
    let value = if high_byte {
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        char::decode_utf16(units)
            .map(|value| value.unwrap_or('\u{fffd}'))
            .collect()
    } else {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, byte_count))
}

fn parse_sheet(
    all_records: &[Record<'_>],
    sheet: &BoundSheet,
    shared_strings: &[rich::Text],
) -> Result<SheetData, String> {
    let start_index = all_records
        .binary_search_by_key(&sheet.offset, |record| record.offset)
        .map_err(|_| unsupported("BOUNDSHEET8 points outside the BIFF record stream"))?;
    let bof = all_records[start_index];
    if bof.kind != BOF || u16_at(bof.data, 0)? != BIFF8 || u16_at(bof.data, 2)? != WORKSHEET {
        return Err(unsupported(
            "BOUNDSHEET8 does not point to a BIFF8 worksheet",
        ));
    }
    let mut output = SheetData {
        visibility: sheet.visibility,
        ..SheetData::default()
    };
    let mut cell_count = 0usize;
    let mut pending_formula_string = None;
    let mut nested_substreams = 0usize;
    let mut found_eof = false;
    let mut custom_view = false;
    let mut previous_kind = 0u16;
    for record in &all_records[start_index + 1..] {
        let prior_kind = std::mem::replace(&mut previous_kind, record.kind);
        // [MS-XLS] 2.1.7: an embedded chart has its own BOF/EOF
        // substream. Its records are not worksheet cells or geometry.
        if matches!(record.kind, BOF | EOF) && pending_formula_string.is_some() {
            return Err(unsupported(
                "BIFF formula string result lacks its STRING record",
            ));
        }
        if record.kind == BOF {
            nested_substreams += 1;
            continue;
        }
        if record.kind == EOF {
            if nested_substreams != 0 {
                nested_substreams -= 1;
                continue;
            }
            found_eof = true;
            break;
        }
        if nested_substreams != 0 {
            continue;
        }
        // [MS-XLS] 2.1.7.20.6 CUSTOMVIEW: its print settings belong
        // to a saved view, not the currently displayed worksheet.
        if record.kind == 0x01aa {
            if custom_view {
                return Err(unsupported("nested BIFF custom view"));
            }
            custom_view = true;
            output.custom_views_omitted = true;
            continue;
        }
        if record.kind == 0x01ab {
            if !custom_view {
                return Err(unsupported("orphan BIFF custom view end"));
            }
            custom_view = false;
            continue;
        }
        if custom_view {
            continue;
        }
        // [MS-XLS] 2.1.7.20.6: FORMULA = [Uncalced] Formula [Array / Table /
        // ShrFmla / SUB] [String *Continue] -- the cached string result may
        // follow the formula's Array, Table or ShrFmla record.
        let formula_part = prior_kind == FORMULA && matches!(record.kind, 0x0221 | 0x0236 | 0x04bc);
        if record.kind != STRING && !formula_part && pending_formula_string.is_some() {
            return Err(unsupported(
                "BIFF formula string result lacks its STRING record",
            ));
        }
        if matches!(
            record.kind,
            NUMBER | RK | LABELSST | LABEL | BOOLERR | FORMULA | 0x0201
        ) {
            let position = cell_position(record.data)?;
            output.cell_styles.insert(position, u16_at(record.data, 4)?);
        }
        output.geometry.read(record)?;
        output.print.read(record)?;
        output.views.read(record)?;
        if record.kind == 0x0208 {
            output.rows.entry(u16_at(record.data, 0)?).or_default();
        }
        match record.kind {
            0x0201 => {
                let (row, column) = cell_position(record.data)?;
                insert_cell(&mut output, row, column, CellValue::Blank, &mut cell_count)?;
            }
            0x00be => parse_mul_blank(record.data, &mut output, &mut cell_count)?,
            NUMBER => {
                let (row, column) = cell_position(record.data)?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Number(f64_at(record.data, 6)?),
                    &mut cell_count,
                )?;
            }
            RK => {
                let (row, column) = cell_position(record.data)?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Number(decode_rk(u32_at(record.data, 6)?)),
                    &mut cell_count,
                )?;
            }
            MULRK => parse_mul_rk(record.data, &mut output, &mut cell_count)?,
            LABELSST => {
                let (row, column) = cell_position(record.data)?;
                let index = usize::try_from(u32_at(record.data, 6)?)
                    .map_err(|_| unsupported("BIFF shared string index is too large"))?;
                shared_strings
                    .get(index)
                    .ok_or_else(|| unsupported("BIFF shared string index is out of range"))?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::SharedString(index),
                    &mut cell_count,
                )?;
            }
            LABEL => {
                let (row, column) = cell_position(record.data)?;
                let (value, _) = parse_biff_string(
                    record
                        .data
                        .get(6..)
                        .ok_or_else(|| unsupported("truncated LABEL record"))?,
                )?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Text(value),
                    &mut cell_count,
                )?;
            }
            BOOLERR => {
                let (row, column) = cell_position(record.data)?;
                let value = *record
                    .data
                    .get(6)
                    .ok_or_else(|| unsupported("truncated BOOLERR record"))?;
                let is_error = *record
                    .data
                    .get(7)
                    .ok_or_else(|| unsupported("truncated BOOLERR record"))?
                    != 0;
                let value = if is_error {
                    CellValue::Error(error_text(value).into())
                } else {
                    CellValue::Bool(value != 0)
                };
                insert_cell(&mut output, row, column, value, &mut cell_count)?;
            }
            FORMULA => {
                output.formula_records.push(record.kind, record.data)?;
                let (row, column) = cell_position(record.data)?;
                match formula_cached_value(record.data)? {
                    FormulaResult::Value(value) => {
                        insert_cell(&mut output, row, column, value, &mut cell_count)?;
                    }
                    FormulaResult::String => pending_formula_string = Some((row, column)),
                    FormulaResult::Empty => {
                        insert_cell(&mut output, row, column, CellValue::Blank, &mut cell_count)?;
                    }
                }
                output.formula_results = true;
            }
            STRING => {
                if let Some((row, column)) = pending_formula_string.take() {
                    let (value, _) = parse_biff_string(record.data)?;
                    insert_cell(
                        &mut output,
                        row,
                        column,
                        CellValue::Text(value),
                        &mut cell_count,
                    )?;
                }
            }
            MERGEDCELLS => parse_merged_cells(record.data, &mut output.merged)?,
            0x01b0 | 0x01b1 | 0x0879 | 0x087a | 0x087b => {
                output.conditional_records.push(record.kind, record.data)?
            }
            0x01b8 | 0x0800 => output.hyperlink_records.push(record.kind, record.data)?,
            0x009d => output.autofilter_info = true,
            0x04bc | 0x0221 | 0x0236 => output.formula_records.push(record.kind, record.data)?,
            0x00ef => output.shows_phonetic |= u16_at(record.data, 4)? != 0,
            0x0208 => output.shows_phonetic |= u16_at(record.data, 14)? & 0x4000 != 0,
            0x007d => output.shows_phonetic |= u16_at(record.data, 8)? & 0x0008 != 0,
            0x01b2 | 0x01be => output.validation_records.push(record.kind, record.data)?,
            0x0862 => {
                if output.sheet_ext.replace(record.data.to_vec()).is_some() {
                    return Err(unsupported("duplicate BIFF sheet extension"));
                }
            }
            0x0871 | 0x0872 | 0x0877 | 0x0878 => {
                output.table_records.push(record.kind, record.data)?
            }
            // ContinueFrt11 (2.4.60) of a table record is not reassembled.
            0x0875 => return Err(unsupported("continued XLS table record")),
            // A continued conditional formatting record is not reassembled.
            CONTINUE
                if !output.conditional_records.is_empty()
                    && matches!(prior_kind, 0x01b0 | 0x01b1 | 0x0879 | 0x087a | 0x087b) =>
            {
                return Err(unsupported("continued XLS conditional formatting record"));
            }
            FILEPASS => return Err(unsupported("encrypted BIFF worksheet")),
            _ => {}
        }
    }
    if !found_eof || custom_view {
        return Err(unsupported("unterminated BIFF worksheet substream"));
    }
    Ok(output)
}

fn insert_cell(
    sheet: &mut SheetData,
    row: u16,
    column: u16,
    value: CellValue,
    count: &mut usize,
) -> Result<(), String> {
    if column > 255 {
        return Err(unsupported("BIFF cell column exceeds the BIFF8 limit"));
    }
    let row_values = sheet.rows.entry(row).or_default();
    if row_values.insert(column, value).is_none() {
        *count += 1;
        if *count > MAX_CELLS {
            return Err(unsupported("too many BIFF worksheet cells"));
        }
    }
    Ok(())
}

fn parse_mul_rk(data: &[u8], sheet: &mut SheetData, count: &mut usize) -> Result<(), String> {
    if data.len() < 12 || !(data.len() - 6).is_multiple_of(6) {
        return Err(unsupported("invalid MULRK record"));
    }
    let row = u16_at(data, 0)?;
    let first_column = u16_at(data, 2)?;
    let last_column = u16_at(data, data.len() - 2)?;
    let values = (data.len() - 6) / 6;
    if last_column < first_column || usize::from(last_column - first_column) + 1 != values {
        return Err(unsupported("inconsistent MULRK column range"));
    }
    for index in 0..values {
        let column = first_column
            .checked_add(index as u16)
            .ok_or_else(|| unsupported("MULRK column overflow"))?;
        let raw = u32_at(data, 6 + index * 6)?;
        sheet
            .cell_styles
            .insert((row, column), u16_at(data, 4 + index * 6)?);
        insert_cell(sheet, row, column, CellValue::Number(decode_rk(raw)), count)?;
    }
    Ok(())
}

fn parse_mul_blank(data: &[u8], sheet: &mut SheetData, count: &mut usize) -> Result<(), String> {
    if data.len() < 8 || !data.len().is_multiple_of(2) {
        return Err(unsupported("invalid MULBLANK record"));
    }
    let row = u16_at(data, 0)?;
    let first = u16_at(data, 2)?;
    let last = u16_at(data, data.len() - 2)?;
    if last < first || last > 255 || usize::from(last - first) + 1 != (data.len() - 6) / 2 {
        return Err(unsupported("invalid MULBLANK column range"));
    }
    for column in first..=last {
        sheet.cell_styles.insert(
            (row, column),
            u16_at(data, 4 + usize::from(column - first) * 2)?,
        );
        insert_cell(sheet, row, column, CellValue::Blank, count)?;
    }
    Ok(())
}

fn parse_merged_cells(data: &[u8], output: &mut Vec<(u16, u16, u16, u16)>) -> Result<(), String> {
    let count = u16_at(data, 0)? as usize;
    let required = 2usize
        .checked_add(count * 8)
        .ok_or_else(|| unsupported("MERGEDCELLS size overflow"))?;
    if required > data.len() || output.len() + count > MAX_CELLS {
        return Err(unsupported("invalid or excessive MERGEDCELLS record"));
    }
    for index in 0..count {
        let offset = 2 + index * 8;
        let first_row = u16_at(data, offset)?;
        let last_row = u16_at(data, offset + 2)?;
        let first_column = u16_at(data, offset + 4)?;
        let last_column = u16_at(data, offset + 6)?;
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return Err(unsupported("invalid BIFF merged-cell range"));
        }
        output.push((first_row, last_row, first_column, last_column));
    }
    Ok(())
}

enum FormulaResult {
    Value(CellValue),
    String,
    Empty,
}

fn formula_cached_value(data: &[u8]) -> Result<FormulaResult, String> {
    let raw = data
        .get(6..14)
        .ok_or_else(|| unsupported("truncated FORMULA result"))?;
    if raw[6] == 0xff && raw[7] == 0xff {
        return Ok(match raw[0] {
            0 => FormulaResult::String,
            1 => FormulaResult::Value(CellValue::Bool(raw[2] != 0)),
            2 => FormulaResult::Value(CellValue::Error(error_text(raw[2]).into())),
            3 => FormulaResult::Empty,
            _ => return Err(unsupported("invalid FORMULA cached result type")),
        });
    }
    Ok(FormulaResult::Value(CellValue::Number(f64::from_le_bytes(
        raw.try_into().expect("eight-byte slice"),
    ))))
}

fn cell_position(data: &[u8]) -> Result<(u16, u16), String> {
    let row = u16_at(data, 0)?;
    let column = u16_at(data, 2)?;
    if column > 255 {
        return Err(unsupported("BIFF cell column exceeds the BIFF8 limit"));
    }
    Ok((row, column))
}

fn decode_rk(raw: u32) -> f64 {
    let mut value = if (raw & 0x02) != 0 {
        ((raw as i32) >> 2) as f64
    } else {
        f64::from_bits(u64::from(raw & 0xffff_fffc) << 32)
    };
    if (raw & 0x01) != 0 {
        value /= 100.0;
    }
    value
}

fn build_xlsx_with_drawings(
    sheets: &[(String, SheetData)],
    styles: &styles::ResolvedStyleSheet,
    shared_strings: Vec<rich::Text>,
    date1904: bool,
    window_count: usize,
    max_output_bytes: usize,
    mdw: Option<f64>,
    drawings: Option<&pictures::Parts>,
) -> Result<Vec<u8>, String> {
    let mut shared_xml_budget = 256 * 1024 * 1024usize;
    let mut shared_xml = Vec::new();
    shared_xml
        .try_reserve_exact(shared_strings.len())
        .map_err(|_| unsupported("BIFF shared string XML cache allocation failed"))?;
    let mut shared_encoder = rich::XmlEncoder::new(styles, &mut shared_xml_budget);
    for value in shared_strings {
        // Drop each neutral entry after encoding it so the byte route does not
        // retain both complete text tables until package serialization ends.
        shared_xml.push(shared_encoder.encode(&value)?);
    }
    let mut workbook = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    workbook.push_str(&format!(
        "<workbookPr date1904=\"{}\"/>",
        u8::from(date1904)
    ));
    // Window1/Window2 associate by ordinal position, not by sheet selection.
    // Preserve that identity without guessing window geometry or an active tab
    // after unsupported chart/macro tabs have been omitted.
    if window_count != 0 {
        workbook.push_str("<bookViews>");
        for _ in 0..window_count {
            workbook.push_str("<workbookView/>");
        }
        workbook.push_str("</bookViews>");
    }
    workbook.push_str("<sheets>");
    let mut workbook_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut content_types = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    let mut parts = vec![
        ("_rels/.rels".into(), ROOT_RELS_XLSX.to_string()),
        ("xl/styles.xml".into(), styles.xml()),
    ];
    // Resource policy: rich run markup must not multiply without a bound when
    // the same SST entry is referenced by many cells or sheets.
    let mut remaining_sheet_xml = 256 * 1024 * 1024usize;
    for (index, (name, sheet)) in sheets.iter().enumerate() {
        let id = index + 1;
        workbook.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"{}/>",
            xml_attr(name),
            id,
            id,
            sheet.visibility.attribute()
        ));
        workbook_rels.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>",
            id, id
        ));
        content_types.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            id
        ));
        let has_drawing = drawings.is_some_and(|d| d.sheets.contains(&index));
        let sheet_xml = build_sheet_xml_with_drawings(
            sheet,
            &shared_xml,
            remaining_sheet_xml,
            if has_drawing { mdw } else { None },
            has_drawing,
        )?;
        remaining_sheet_xml = remaining_sheet_xml
            .checked_sub(sheet_xml.len())
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
        parts.push((format!("xl/worksheets/sheet{id}.xml"), sheet_xml));
    }
    workbook.push_str("</sheets></workbook>");
    workbook_rels.push_str(&format!(
        "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>",
        sheets.len() + 1
    ));
    if let Some(drawings) = drawings {
        content_types.push_str(&drawings.types);
    }
    content_types.push_str("</Types>");
    parts.push(("xl/workbook.xml".into(), workbook));
    parts.push(("xl/_rels/workbook.xml.rels".into(), workbook_rels));
    parts.push(("[Content_Types].xml".into(), content_types));
    if let Some(drawings) = drawings {
        crate::ooxml::write_package_bytes(
            parts
                .iter()
                .chain(drawings.xml.iter())
                .map(|(n, b)| (n.as_str(), b.as_bytes()))
                .chain(
                    drawings
                        .media
                        .iter()
                        .map(|(n, b)| (n.as_str(), b.as_slice())),
                ),
            max_output_bytes,
        )
    } else {
        write_package(&parts, max_output_bytes)
    }
}

#[cfg(test)]
fn build_sheet_xml(sheet: &SheetData, max_bytes: usize) -> Result<String, String> {
    build_sheet_xml_with_drawings(sheet, &[], max_bytes, None, false)
}

fn build_sheet_xml_with_drawings(
    sheet: &SheetData,
    shared_strings: &[String],
    max_bytes: usize,
    mdw: Option<f64>,
    drawing: bool,
) -> Result<String, String> {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );
    xml.push_str(&sheet.print.sheet_properties());
    xml.push_str(&sheet.views.xml());
    xml.push_str(&sheet.geometry.xml_with_metrics(mdw));
    xml.push_str("<sheetData>");
    for (row, cells) in &sheet.rows {
        xml.push_str(&format!(
            "<row r=\"{}\"{}>",
            u32::from(*row) + 1,
            sheet.geometry.row_attributes(*row)
        ));
        for (column, value) in cells {
            let reference = cell_reference(*row, *column);
            let style = sheet
                .cell_styles
                .get(&(*row, *column))
                .copied()
                .unwrap_or(0);
            let cell = format!("<c r=\"{reference}\" s=\"{style}\"");
            match value {
                CellValue::Blank => xml.push_str(&format!("{cell}/>")),
                CellValue::Number(value) if value.is_finite() => {
                    xml.push_str(&format!("{cell}><v>{value}</v></c>"));
                }
                CellValue::Number(_) => {
                    xml.push_str(&format!("{cell} t=\"e\"><v>#NUM!</v></c>"));
                }
                CellValue::Text(value) => {
                    xml.push_str(&format!(
                        "{cell} t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                        xml_text(value)
                    ));
                }
                CellValue::SharedString(index) => {
                    let value = shared_strings
                        .get(*index)
                        .ok_or_else(|| unsupported("BIFF shared string index is out of range"))?;
                    if value.len() > max_bytes.saturating_sub(xml.len()) {
                        return Err("OUTPUT_TOO_LARGE".into());
                    }
                    xml.push_str(&format!("{cell} t=\"inlineStr\"><is>"));
                    xml.push_str(value);
                    xml.push_str("</is></c>");
                }
                CellValue::Bool(value) => {
                    xml.push_str(&format!("{cell} t=\"b\"><v>{}</v></c>", u8::from(*value)));
                }
                CellValue::Error(value) => {
                    xml.push_str(&format!("{cell} t=\"e\"><v>{}</v></c>", xml_text(value)));
                }
            }
            if xml.len() > max_bytes {
                return Err("OUTPUT_TOO_LARGE".into());
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");
    if !sheet.merged.is_empty() {
        xml.push_str(&format!("<mergeCells count=\"{}\">", sheet.merged.len()));
        for (first_row, last_row, first_column, last_column) in &sheet.merged {
            xml.push_str(&format!(
                "<mergeCell ref=\"{}:{}\"/>",
                cell_reference(*first_row, *first_column),
                cell_reference(*last_row, *last_column)
            ));
        }
        xml.push_str("</mergeCells>");
    }
    xml.push_str(&sheet.print.xml());
    if drawing {
        xml.push_str("<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"legacyDrawing\"/>");
    }
    xml.push_str("</worksheet>");
    if xml.len() > max_bytes {
        return Err("OUTPUT_TOO_LARGE".into());
    }
    Ok(xml)
}

fn minimal_styles() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.into()
}

fn cell_reference(row: u16, column: u16) -> String {
    let mut value = u32::from(column) + 1;
    let mut letters = Vec::new();
    while value > 0 {
        let remainder = ((value - 1) % 26) as u8;
        letters.push(char::from(b'A' + remainder));
        value = (value - 1) / 26;
    }
    letters.reverse();
    format!(
        "{}{}",
        letters.into_iter().collect::<String>(),
        u32::from(row) + 1
    )
}

fn error_text(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        _ => "#VALUE!",
    }
}

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, String> {
    let raw = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| unsupported("truncated BIFF integer"))?;
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let raw = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| unsupported("truncated BIFF integer"))?;
    Ok(u32::from_le_bytes(raw.try_into().expect("four-byte slice")))
}

fn f64_at(bytes: &[u8], offset: usize) -> Result<f64, String> {
    let raw = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| unsupported("truncated BIFF number"))?;
    Ok(f64::from_le_bytes(
        raw.try_into().expect("eight-byte slice"),
    ))
}

#[cfg(test)]
mod tests {
    use super::{cell_reference, decode_rk, parse_biff_string, parse_sst, records};

    #[test]
    fn worksheet_window_flags_survive_without_modifying_cells() {
        use super::*;
        let bof = [0, 6, 0x10, 0];
        let mut window = [0u8; 18];
        window[..2].copy_from_slice(&0x0040u16.to_le_bytes());
        let records = [
            Record {
                kind: BOF,
                offset: 0,
                data: &bof,
            },
            Record {
                kind: 0x023e,
                offset: 8,
                data: &window,
            },
            Record {
                kind: EOF,
                offset: 30,
                data: &[],
            },
        ];
        let bound = BoundSheet {
            offset: 0,
            name: "A".into(),
            sheet_type: 0,
            visibility: SheetVisibility::Visible,
        };
        let sheet = parse_sheet(&records, &bound, &[]).unwrap();
        let xml = build_sheet_xml(&sheet, 10000).unwrap();
        assert!(xml.contains("showGridLines=\"0\""));
        assert!(xml.contains("showZeros=\"0\""));
        assert!(xml.contains("rightToLeft=\"1\""));
        assert!(xml.contains("<sheetData></sheetData>"));
        assert_eq!(
            build_sheet_xml(&sheet, xml.len() - 1).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }

    #[test]
    fn formula_string_results_follow_shared_formula_records() {
        use super::*;
        let bof = [0, 6, 0x10, 0];
        // Row 0, column 0, XF 15; cached string result; PtgExp to itself.
        let mut formula = vec![0, 0, 0, 0, 15, 0, 0, 0, 0, 0, 0, 0, 0xff, 0xff];
        formula.extend([8, 0, 0, 0, 0, 0, 5, 0, 1, 0, 0, 0, 0]);
        let shared = [0u8; 10];
        let string = [2, 0, 0, b'o', b'k'];
        let record = |kind, offset, data| Record { kind, offset, data };
        let bound = BoundSheet {
            offset: 0,
            name: "A".into(),
            sheet_type: 0,
            visibility: SheetVisibility::Visible,
        };
        let records = [
            record(BOF, 0, &bof[..]),
            record(FORMULA, 10, &formula),
            record(0x04bc, 50, &shared),
            record(STRING, 70, &string),
            record(EOF, 80, &[]),
        ];
        let sheet = parse_sheet(&records, &bound, &[]).unwrap();
        assert!(matches!(
            sheet.rows.get(&0).and_then(|row| row.get(&0)),
            Some(CellValue::Text(text)) if text == "ok"
        ));
        // A string result without its STRING record is malformed.
        let records = [
            record(BOF, 0, &bof[..]),
            record(FORMULA, 10, &formula),
            record(0x04bc, 50, &shared),
            record(EOF, 80, &[]),
        ];
        assert!(parse_sheet(&records, &bound, &[]).is_err());
    }

    #[test]
    fn sheet_visibility_uses_only_the_two_defined_bits() {
        for flags in 0..=u8::MAX {
            let parsed = super::parse_bound_sheet(&[0, 0, 0, 0, flags, 0, 1, 0, b'A']);
            if flags & 3 == 3 {
                assert!(parsed
                    .unwrap_err()
                    .contains("invalid BIFF sheet visibility"));
            } else {
                let expected = ["", " state=\"hidden\"", " state=\"veryHidden\""];
                assert_eq!(
                    parsed.unwrap().visibility.attribute(),
                    expected[usize::from(flags & 3)]
                );
            }
        }
    }

    #[test]
    fn saved_custom_view_cannot_override_the_current_print_settings() {
        use super::*;
        let bof = [0, 6, 0x10, 0];
        let records = [
            Record {
                kind: BOF,
                offset: 0,
                data: &bof,
            },
            Record {
                kind: 0x14,
                offset: 8,
                data: &[],
            },
            Record {
                kind: 0x1aa,
                offset: 12,
                data: &[0; 64],
            },
            Record {
                kind: 0x14,
                offset: 80,
                data: &[1, 0, 0, b'X'],
            },
            Record {
                kind: 0x1ab,
                offset: 88,
                data: &[0; 2],
            },
            Record {
                kind: EOF,
                offset: 94,
                data: &[],
            },
        ];
        let sheet = BoundSheet {
            offset: 0,
            name: "Sheet".into(),
            sheet_type: 0,
            visibility: SheetVisibility::Visible,
        };
        let output = parse_sheet(&records, &sheet, &[]).unwrap();
        assert!(output.print.xml().contains("<oddHeader></oddHeader>"));
        assert!(parse_sheet(&records[..4], &sheet, &[]).is_err());
    }

    #[test]
    fn worksheet_ignores_cells_in_nested_chart_substreams() {
        use super::*;
        let bof = [0, 6, 0x10, 0];
        let chart_bof = [0, 6, 0x20, 0];
        let number = [0u8; 14];
        let records = [
            Record {
                kind: BOF,
                offset: 0,
                data: &bof,
            },
            Record {
                kind: BOF,
                offset: 8,
                data: &chart_bof,
            },
            Record {
                kind: NUMBER,
                offset: 16,
                data: &number,
            },
            Record {
                kind: EOF,
                offset: 34,
                data: &[],
            },
            Record {
                kind: EOF,
                offset: 38,
                data: &[],
            },
        ];
        let sheet = BoundSheet {
            offset: 0,
            name: "S".into(),
            sheet_type: 0,
            visibility: SheetVisibility::Visible,
        };
        assert!(parse_sheet(&records, &sheet, &[]).unwrap().rows.is_empty());
        assert!(parse_sheet(&records[..4], &sheet, &[]).is_err());
    }

    #[test]
    fn preserves_cell_xf_blank_borders_and_sheet_geometry() {
        use crate::{cfb::test_support::build_cfb, convert_native, LegacyFormat};
        use std::io::{Cursor, Read};
        fn rec(kind: u16, data: &[u8]) -> Vec<u8> {
            [
                kind.to_le_bytes().as_slice(),
                &(data.len() as u16).to_le_bytes(),
                data,
            ]
            .concat()
        }
        let mut stream = rec(0x0809, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(rec(0x0085, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
        stream.extend(rec(0x0022, &[1, 0])); // Date1904
        let mut font = vec![0; 16];
        font[0..2].copy_from_slice(&360u16.to_le_bytes());
        font[2] = 2; // italic
        font[4..6].copy_from_slice(&10u16.to_le_bytes());
        font[6..8].copy_from_slice(&700u16.to_le_bytes());
        font[14..16].copy_from_slice(&[5, 0]);
        font.extend(b"Arial");
        stream.extend(rec(0x0031, &font));
        let mut xf = [0u8; 20];
        xf[6] = 0x2a; // center, wrap, bottom
        xf[10..14].copy_from_slice(&(1u32 | (10 << 16)).to_le_bytes()); // thin red left
        xf[14..18].copy_from_slice(&(1u32 << 26).to_le_bytes()); // solid fill
        xf[18..20].copy_from_slice(&(13u16 | (65 << 7)).to_le_bytes());
        stream.extend(rec(0x00e0, &xf));
        xf[2..4].copy_from_slice(&14u16.to_le_bytes()); // date format
        stream.extend(rec(0x00e0, &xf));
        stream.extend(rec(0x000a, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(rec(0x0809, &[0, 6, 0x10, 0]));
        // Print layout must survive independently of cell styles.
        stream.extend(rec(0x0081, &[0, 1])); // fit to pages
        for kind in 0x0026..=0x0029 {
            stream.extend(rec(kind, &0.5f64.to_le_bytes()));
        }
        let mut setup = vec![0u8; 34];
        setup[..16].copy_from_slice(&[9, 0, 75, 0, 3, 0, 2, 0, 0, 0, 0x89, 0, 88, 2, 88, 2]);
        setup[16..24].copy_from_slice(&0.25f64.to_le_bytes());
        setup[24..32].copy_from_slice(&0.3f64.to_le_bytes());
        setup[32] = 1;
        stream.extend(rec(0x00a1, &setup));
        stream.extend(rec(0x0014, &[4, 0, 0, b'&', b'L', b'&', b'P']));
        stream.extend(rec(0x001b, &[1, 0, 4, 0, 0, 0, 255, 63]));
        stream.extend(rec(0x0225, &[0, 0, 0x2c, 1])); // default 15 pt
        stream.extend(rec(0x007d, &[0, 0, 2, 0, 0, 20, 1, 0, 3, 0, 0, 0]));
        let mut row = [0u8; 16];
        row[0] = 2; // empty third row, preserve height/hidden state
        row[6..8].copy_from_slice(&600u16.to_le_bytes());
        row[12] = 0x60;
        stream.extend(rec(0x0208, &row));
        let mut number = vec![0, 0, 0, 0, 1, 0];
        number.extend(1f64.to_le_bytes());
        stream.extend(rec(0x0203, &number));
        stream.extend(rec(0x0201, &[0, 0, 1, 0, 1, 0])); // styled empty cell
        stream.extend(rec(0x00be, &[1, 0, 0, 0, 1, 0, 1, 0, 1, 0])); // two blanks
        stream.extend(rec(0x000a, &[]));
        let data = build_cfb(&[("Workbook", stream)]);
        let output = convert_native(&data, LegacyFormat::Xls, 1024 * 1024).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(output.bytes)).unwrap();
        let mut part = |name| {
            let mut xml = String::new();
            archive
                .by_name(name)
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            xml
        };
        let styles = part("xl/styles.xml");
        assert!(styles.contains("<sz val=\"18\"/>"));
        assert!(styles.contains("<b/>") && styles.contains("<i/>"));
        assert!(styles.contains("numFmtId=\"14\""));
        assert!(styles.contains("patternType=\"solid\""));
        assert!(styles.contains("<left style=\"thin\">"));
        assert!(styles.contains("horizontal=\"center\"") && styles.contains("wrapText=\"1\""));
        let sheet = part("xl/worksheets/sheet1.xml");
        assert!(sheet.contains("fitToPage=\"1\""));
        assert!(sheet.contains("<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.5\" bottom=\"0.5\" header=\"0.25\" footer=\"0.3\"/>"));
        assert!(sheet.contains("orientation=\"landscape\"") && sheet.contains("scale=\"75\""));
        assert!(
            sheet.contains("firstPageNumber=\"3\"") && sheet.contains("pageOrder=\"overThenDown\"")
        );
        assert!(sheet.contains("<oddHeader>&amp;L&amp;P</oddHeader>"));
        assert!(sheet.contains("<brk id=\"4\" min=\"0\" max=\"16383\" man=\"1\"/>"));
        assert!(sheet.contains("<c r=\"A1\" s=\"1\"><v>1</v></c>"));
        assert!(sheet.contains("<c r=\"B1\" s=\"1\"/>"));
        assert!(sheet.contains("<c r=\"A2\" s=\"1\"/>") && sheet.contains("<c r=\"B2\" s=\"1\"/>"));
        assert!(sheet.contains("width=\"20\""));
        assert!(sheet.contains("<row r=\"3\" ht=\"30\" hidden=\"1\""));
        assert!(part("xl/workbook.xml").contains("date1904=\"1\""));
    }

    #[test]
    fn decodes_biff8_unicode_strings() {
        let mut raw = vec![3, 0, 1];
        for unit in "日本語".encode_utf16() {
            raw.extend_from_slice(&unit.to_le_bytes());
        }
        assert_eq!(parse_biff_string(&raw).unwrap().0, "日本語");
    }

    #[test]
    fn decodes_integer_and_scaled_rk_values() {
        assert_eq!(decode_rk((42u32 << 2) | 2), 42.0);
        assert_eq!(decode_rk((1234u32 << 2) | 3), 12.34);
    }

    #[test]
    fn formats_biff8_cell_references() {
        assert_eq!(cell_reference(0, 0), "A1");
        assert_eq!(cell_reference(65535, 255), "IV65536");
    }

    #[test]
    fn rejects_impossible_shared_string_counts_before_allocation() {
        let mut raw = vec![0; 8];
        raw[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_sst(&[&raw]).is_err());
    }

    #[test]
    fn decodes_sst_text_across_continue_records() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&4u16.to_le_bytes());
        sst.push(0);
        sst.extend_from_slice(b"ab");
        let continued = [1, 0x2d, 0x4e, 0x87, 0x65];

        assert_eq!(
            parse_sst(&[&sst, &continued]).unwrap(),
            ["ab\u{4e2d}\u{6587}"]
        );
    }

    #[test]
    fn decodes_unicode_sst_text_after_switching_to_compressed_continue() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&4u16.to_le_bytes());
        sst.push(1);
        for value in ['日', '本'] {
            sst.extend_from_slice(&(value as u16).to_le_bytes());
        }
        let continued = [0, b'a', b'b'];

        assert_eq!(parse_sst(&[&sst, &continued]).unwrap(), ["日本ab"]);
    }

    #[test]
    fn rejects_sst_string_headers_split_across_continue_records() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u16.to_le_bytes());
        let continued = [0, b'a'];

        assert!(parse_sst(&[&sst, &continued]).is_err());
    }

    #[test]
    fn stops_unicode_text_before_the_next_sst_string_header() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&2u32.to_le_bytes());
        sst.extend_from_slice(&2u32.to_le_bytes());
        for value in ['\u{4e2d}', '\u{6587}'] {
            sst.extend_from_slice(&1u16.to_le_bytes());
            sst.push(1);
            sst.extend_from_slice(&(value as u16).to_le_bytes());
        }

        assert_eq!(parse_sst(&[&sst]).unwrap(), ["\u{4e2d}", "\u{6587}"]);
    }

    #[test]
    fn rejects_zero_records_that_hide_nonzero_trailing_data() {
        assert!(records(&[0, 0, 0, 0, 1]).is_err());
        assert!(records(&[0; 5]).unwrap().is_empty());
    }
}
