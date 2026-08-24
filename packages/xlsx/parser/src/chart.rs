use crate::read_zip_string;
use crate::types::*;
use crate::worksheet_reference::{
    parse_a1_range, resolve_worksheet_reference_with_visibility, resolve_worksheet_visibility,
    split_sheet_ref, ReferencedCellValue, ResolvedWorksheetReference, WorksheetReferenceSession,
};
use crate::{
    find_rel_target_by_type, parse_rels_map, resolve_fill_color, resolve_sheet_path,
    resolve_zip_path,
};
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::{is_c_ns, is_r_ns, is_x_ns, is_xdr_ns};
use std::collections::{BTreeMap, HashMap};

const CHARTEX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const MAX_CHART_STYLE_FORMATS: usize = 1_048_576;
const MAX_CHART_CUSTOM_NUMBER_FORMATS: usize = 65_536;
const MAX_CHART_NUMBER_FORMAT_BYTES: usize = 16 * 1024 * 1024;

/// Workbook-wide projection of the two style tables chart source-linked
/// number formats need. The complete cell style parser owns all other style
/// semantics; this bounded index avoids reinflating/reparsing `styles.xml` for
/// every chart or worksheet projection.
#[derive(Default)]
pub(crate) struct ChartNumberFormatCache {
    style_num_fmt_ids: Vec<Option<u32>>,
    custom_num_formats: HashMap<u32, String>,
}

impl ChartNumberFormatCache {
    pub(crate) fn from_document(document: &roxmltree::Document<'_>) -> Self {
        let style_num_fmt_ids = document
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().name() == "cellXfs"
                    && is_x_ns(node.tag_name().namespace())
            })
            .into_iter()
            .flat_map(|cell_xfs| {
                cell_xfs.children().filter(|node| {
                    node.is_element()
                        && node.tag_name().name() == "xf"
                        && is_x_ns(node.tag_name().namespace())
                })
            })
            .take(MAX_CHART_STYLE_FORMATS)
            .map(|xf| {
                xf.attribute("numFmtId")
                    .and_then(|value| value.parse::<u32>().ok())
            })
            .collect();

        let mut custom_num_formats = HashMap::new();
        let mut remaining_bytes = MAX_CHART_NUMBER_FORMAT_BYTES;
        if let Some(num_fmts) = document.descendants().find(|node| {
            node.is_element()
                && node.tag_name().name() == "numFmts"
                && is_x_ns(node.tag_name().namespace())
        }) {
            for num_fmt in num_fmts
                .children()
                .filter(|node| {
                    node.is_element()
                        && node.tag_name().name() == "numFmt"
                        && is_x_ns(node.tag_name().namespace())
                })
                .take(MAX_CHART_CUSTOM_NUMBER_FORMATS)
            {
                let Some(id) = num_fmt
                    .attribute("numFmtId")
                    .and_then(|value| value.parse::<u32>().ok())
                else {
                    continue;
                };
                let Some(code) = num_fmt.attribute("formatCode") else {
                    continue;
                };
                if code.len() > remaining_bytes {
                    break;
                }
                remaining_bytes -= code.len();
                custom_num_formats.insert(id, code.to_string());
            }
        }

        Self {
            style_num_fmt_ids,
            custom_num_formats,
        }
    }

    #[cfg(test)]
    pub(crate) fn from_styles(styles: &Styles) -> Self {
        let style_num_fmt_ids = styles
            .cell_xfs
            .iter()
            .take(MAX_CHART_STYLE_FORMATS)
            .map(|xf| Some(xf.num_fmt_id))
            .collect();
        let mut custom_num_formats = HashMap::new();
        let mut remaining_bytes = MAX_CHART_NUMBER_FORMAT_BYTES;
        for format in styles.num_fmts.iter().take(MAX_CHART_CUSTOM_NUMBER_FORMATS) {
            let code = &format.format_code;
            if code.len() > remaining_bytes {
                break;
            }
            remaining_bytes -= code.len();
            custom_num_formats.insert(format.num_fmt_id, code.clone());
        }
        Self {
            style_num_fmt_ids,
            custom_num_formats,
        }
    }
}

/// Collect live graphic frames while applying MCE branch selection at any
/// nesting depth. Excel may place a chartEx `AlternateContent` directly under
/// an anchor or inside an `xdr:grpSp`; walking plain descendants would also
/// visit the fallback preview, so MCE nodes must be handled explicitly.
fn collect_selected_graphic_frames<'a, 'input>(
    node: roxmltree::Node<'a, 'input>,
    out: &mut Vec<roxmltree::Node<'a, 'input>>,
) {
    if !node.is_element() {
        return;
    }
    if node.tag_name().name() == "AlternateContent" {
        if let Some(selected) =
            ooxml_common::mce::select_alternate_content(node, &crate::drawing::xlsx_understands_ns)
        {
            collect_selected_graphic_frames(selected, out);
        }
        return;
    }
    if node.tag_name().name() == "graphicFrame" && is_xdr_ns(node.tag_name().namespace()) {
        out.push(node);
        return;
    }
    for child in node.children().filter(|child| child.is_element()) {
        collect_selected_graphic_frames(child, out);
    }
}

pub(crate) struct ChartReferenceContext<'a, 'input, 'session> {
    pub(crate) materialized_rows: Option<&'a [Row]>,
    pub(crate) materialized_col_hidden: Option<&'a BTreeMap<u32, bool>>,
    pub(crate) sheet_name: &'a str,
    pub(crate) sheets: &'a [SheetMeta],
    pub(crate) workbook_rels: &'a roxmltree::Document<'input>,
    pub(crate) shared_strings: &'a [SharedString],
    pub(crate) defined_names: &'a [DefinedName],
    pub(crate) number_formats: &'a ChartNumberFormatCache,
    pub(crate) session: &'session mut WorksheetReferenceSession,
}

struct XlsxChartReferenceResolver<'archive, 'data, 'input, 'session> {
    archive: &'archive mut crate::XlsxZip,
    materialized_rows: Option<&'data [Row]>,
    materialized_col_hidden: Option<&'data BTreeMap<u32, bool>>,
    sheet_name: &'data str,
    sheets: &'data [SheetMeta],
    workbook_rels: &'data roxmltree::Document<'input>,
    shared_strings: &'data [SharedString],
    defined_names: &'data [DefinedName],
    number_formats: &'data ChartNumberFormatCache,
    session: &'session mut WorksheetReferenceSession,
    visibility_cache: HashMap<String, Vec<bool>>,
}

impl XlsxChartReferenceResolver<'_, '_, '_, '_> {
    fn resolve_cells(&mut self, formula: &str) -> Option<(String, ResolvedWorksheetReference)> {
        let expanded = self.expand_defined_name(formula);
        let resolved = resolve_worksheet_reference_with_visibility(
            self.archive,
            &expanded,
            self.materialized_rows,
            self.materialized_col_hidden,
            self.sheet_name,
            self.sheets,
            self.workbook_rels,
            self.shared_strings,
            self.session,
        )?;
        Some((expanded, resolved))
    }

    fn expand_defined_name(&self, formula: &str) -> String {
        let trimmed = formula.trim();
        self.defined_names
            .iter()
            .find(|defined| defined.name.eq_ignore_ascii_case(trimmed))
            .map(|defined| defined.formula.clone())
            .unwrap_or_else(|| trimmed.to_string())
    }

    fn source_style_index(&mut self, formula: &str) -> Option<u32> {
        let (formula_sheet, reference) = split_sheet_ref(formula)?;
        let range = parse_a1_range(&reference)?;
        let target_sheet = formula_sheet.as_deref().unwrap_or(self.sheet_name);
        if target_sheet.eq_ignore_ascii_case(self.sheet_name) {
            if let Some(rows) = self.materialized_rows {
                return rows
                    .iter()
                    .find(|row| row.index == range.top)
                    .and_then(|row| {
                        row.cells
                            .iter()
                            .find(|cell| cell.row == range.top && cell.col == range.left)
                    })
                    .map(|cell| cell.style_index.unwrap_or(0));
            }
        }

        let sheet = self
            .sheets
            .iter()
            .find(|sheet| sheet.name.eq_ignore_ascii_case(target_sheet))?;
        let relative_path = resolve_sheet_path(self.workbook_rels, &sheet.r_id)?;
        let part = format!("xl/{relative_path}");
        let xml = read_zip_string(self.archive, &part).ok()?;
        let document = parse_guarded(&xml).ok()?;
        document
            .descendants()
            .find(|node| {
                if !node.is_element() || node.tag_name().name() != "c" {
                    return false;
                }
                node.attribute("r")
                    .and_then(parse_a1_range)
                    .is_some_and(|cell| cell.top == range.top && cell.left == range.left)
            })
            .map(|cell| {
                cell.attribute("s")
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(0)
            })
    }

    fn number_format_id_for_style(&mut self, style_index: u32) -> Option<u32> {
        self.number_formats
            .style_num_fmt_ids
            .get(style_index as usize)
            .copied()
            .flatten()
    }

    fn number_format_for_style(&mut self, style_index: u32) -> Option<String> {
        let num_fmt_id = self.number_format_id_for_style(style_index)?;
        if let Some(format_code) = self.number_formats.custom_num_formats.get(&num_fmt_id) {
            return Some(format_code.clone());
        }
        // ECMA-376 §18.8.30 built-in number formats. Chart labels only need
        // the format code; returning it keeps the shared renderer independent
        // of XLSX style-index semantics.
        let built_in = match num_fmt_id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            14 => "m/d/yy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yy h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            48 => "##0.0E+0",
            49 => "@",
            _ => return None,
        };
        Some(built_in.to_string())
    }
}

impl ooxml_common::chart::ChartReferenceResolver for XlsxChartReferenceResolver<'_, '_, '_, '_> {
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
        self.resolve_cells(formula).map(|(formula, resolved)| {
            self.visibility_cache.insert(formula, resolved.hidden);
            resolved
                .values
                .into_iter()
                .map(|value| match value {
                    ReferencedCellValue::Text(text) => text,
                    ReferencedCellValue::Number(number) => number.to_string(),
                    ReferencedCellValue::Empty => String::new(),
                })
                .collect()
        })
    }

    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
        self.resolve_cells(formula).map(|(formula, resolved)| {
            self.visibility_cache.insert(formula, resolved.hidden);
            resolved
                .values
                .into_iter()
                .map(|value| match value {
                    ReferencedCellValue::Number(number) => Some(number),
                    _ => None,
                })
                .collect()
        })
    }

    fn resolve_hidden(&mut self, formula: &str) -> Option<Vec<bool>> {
        let expanded = self.expand_defined_name(formula);
        if let Some(hidden) = self.visibility_cache.get(&expanded) {
            return Some(hidden.clone());
        }
        let hidden = resolve_worksheet_visibility(
            self.archive,
            &expanded,
            self.materialized_rows,
            self.materialized_col_hidden,
            self.sheet_name,
            self.sheets,
            self.workbook_rels,
            self.shared_strings,
            self.session,
        )?;
        self.visibility_cache.insert(expanded, hidden.clone());
        Some(hidden)
    }

    fn resolve_number_format(&mut self, formula: &str) -> Option<String> {
        let formula = self.expand_defined_name(formula);
        let style_index = self.source_style_index(&formula)?;
        self.number_format_for_style(style_index)
    }

    fn resolve_number_format_id(&mut self, formula: &str) -> Option<u32> {
        let formula = self.expand_defined_name(formula);
        let style_index = self.source_style_index(&formula)?;
        self.number_format_id_for_style(style_index)
    }

    fn resolve_string_levels(&mut self, formula: &str) -> Option<Vec<Vec<String>>> {
        let expanded = self.expand_defined_name(formula);
        let (_, reference) = split_sheet_ref(&expanded)?;
        let range = parse_a1_range(&reference)?;
        let column_count = usize::try_from(range.right - range.left + 1).ok()?;
        let values = self.resolve_strings(&expanded)?;
        if column_count == 0 || values.len() % column_count != 0 {
            return None;
        }
        let row_count = values.len() / column_count;
        let mut levels = vec![Vec::with_capacity(row_count); column_count];
        for row in values.chunks_exact(column_count) {
            for (column, value) in row.iter().enumerate() {
                levels[column].push(value.clone());
            }
        }
        // Worksheet hierarchy ranges are authored root→leaf by columns, while
        // chartEx `<cx:lvl>` document order is deepest→root.
        levels.reverse();
        Some(levels)
    }
}

/// Read the chartStyle part (`styleN.xml`) associated with a chart part at
/// `chart_path` (e.g. `xl/charts/chart1.xml`), following that part's own
/// relationships (`xl/charts/_rels/chart1.xml.rels`) to the
/// `.../2011/relationships/chartStyle` target. Returns `None` when the chart
/// has no chartStyle relationship or the part cannot be read (the chartEx
/// title then falls back to its inline size, or the renderer's default).
fn load_chart_style_xml(archive: &mut crate::XlsxZip, chart_path: &str) -> Option<String> {
    load_chart_sidecar_xml(
        archive,
        chart_path,
        ooxml_common::chart::CHART_STYLE_REL_TYPE_SUFFIX,
    )
}

fn load_chart_color_style_xml(archive: &mut crate::XlsxZip, chart_path: &str) -> Option<String> {
    load_chart_sidecar_xml(
        archive,
        chart_path,
        ooxml_common::chart::CHART_COLOR_STYLE_REL_TYPE_SUFFIX,
    )
}

fn load_chart_image_relationships(
    archive: &mut crate::XlsxZip,
    chart_path: &str,
) -> ooxml_common::chart::ChartImageRelationships {
    let mut images = ooxml_common::chart::ChartImageRelationships::default();
    let Some((dir, file)) = chart_path.rsplit_once('/') else {
        return images;
    };
    let rels_path = format!("{dir}/_rels/{file}.rels");
    let Ok(rels_xml) = read_zip_string(archive, &rels_path) else {
        return images;
    };
    images.insert_part_relationships(
        ooxml_common::chart::ChartImageSource::Chart,
        chart_path,
        &rels_xml,
    );
    if let Some(target) =
        find_rel_target_by_type(&rels_xml, ooxml_common::chart::CHART_STYLE_REL_TYPE_SUFFIX)
    {
        let style_path = resolve_zip_path(dir, &target);
        let style_rels_path = ooxml_common::rels::relationship_part_path(&style_path);
        if let Ok(style_rels_xml) = read_zip_string(archive, &style_rels_path) {
            images.insert_part_relationships(
                ooxml_common::chart::ChartImageSource::Style,
                &style_path,
                &style_rels_xml,
            );
        }
    }
    images
}

fn load_chart_sidecar_xml(
    archive: &mut crate::XlsxZip,
    chart_path: &str,
    relationship_suffix: &str,
) -> Option<String> {
    let (dir, file) = chart_path.rsplit_once('/')?;
    let rels_path = format!("{}/_rels/{}.rels", dir, file);
    let rels_xml = read_zip_string(archive, &rels_path).ok()?;
    let target = find_rel_target_by_type(&rels_xml, relationship_suffix)?;
    let style_path = resolve_zip_path(dir, &target);
    read_zip_string(archive, &style_path).ok()
}

/// Follow the owning legacy chart's `<c:userShapes r:id>` relationship to its
/// Chart Drawing part (`cdr:CT_Drawing`, ECMA-376 dml-chartDrawing.xsd).
fn load_chart_user_shapes_xml(
    archive: &mut crate::XlsxZip,
    chart_path: &str,
    chart_xml: &str,
) -> Option<String> {
    let chart_doc = parse_guarded(chart_xml).ok()?;
    let rid = chart_doc
        .root_element()
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "userShapes")?
        .attributes()
        .find(|attribute| attribute.name() == "id" && is_r_ns(attribute.namespace()))?
        .value()
        .to_string();
    let (dir, file) = chart_path.rsplit_once('/')?;
    let rels_path = format!("{}/_rels/{}.rels", dir, file);
    let rels_xml = read_zip_string(archive, &rels_path).ok()?;
    let target = parse_rels_map(&rels_xml).remove(&rid)?;
    let user_shapes_path = resolve_zip_path(dir, &target);
    read_zip_string(archive, &user_shapes_path).ok()
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s) for chart anchors (`<xdr:graphicFrame>` elements).
#[cfg(test)]
pub(crate) fn load_sheet_charts(
    archive: &mut crate::XlsxZip,
    sheet_path: &str,
    reference_context: Option<ChartReferenceContext<'_, '_, '_>>,
    theme_colors: &[String],
    theme_fonts: (Option<&str>, Option<&str>),
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
) -> Vec<ChartAnchor> {
    let theme_images = ooxml_common::chart::ChartImageRelationships::default();
    load_sheet_charts_with_theme_images(
        archive,
        sheet_path,
        reference_context,
        theme_colors,
        theme_fonts,
        theme_format_scheme,
        &theme_images,
    )
}

pub(crate) fn load_sheet_charts_with_theme_images(
    archive: &mut crate::XlsxZip,
    sheet_path: &str,
    mut reference_context: Option<ChartReferenceContext<'_, '_, '_>>,
    theme_colors: &[String],
    theme_fonts: (Option<&str>, Option<&str>),
    theme_format_scheme: Option<&ooxml_common::theme::ThemeFormatScheme>,
    theme_images: &ooxml_common::chart::ChartImageRelationships,
) -> Vec<ChartAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = parse_guarded(&sheet_rels_xml) else {
        return Vec::new();
    };

    // Collect all drawing relationship targets
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() {
        return Vec::new();
    }

    let mut all_charts: Vec<ChartAnchor> = Vec::new();

    for target in drawing_targets {
        // Resolve drawing path relative to the sheet directory
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        let Ok(draw_doc) = parse_guarded(&drawing_xml) else {
            continue;
        };

        // Load drawing rels (to resolve chart rIds)
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
            continue;
        };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_string(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        // Iterate over chart anchors. Charts may be saved either as a
        // `<xdr:twoCellAnchor>` (from + to cells — Excel's default) or a
        // `<xdr:oneCellAnchor>` (from cell + a saved `<xdr:ext cx cy>` EMU
        // size, ECMA-376 §20.5.2.24), or `<xdr:absoluteAnchor>` (absolute
        // `<xdr:pos x y>` + extent, §20.5.2.1; the normal chart-sheet form).
        // All three must produce the same bounded ChartAnchor wire shape.
        for anchor in draw_doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
        {
            let anchor_tag = anchor.tag_name().name();
            let is_one_cell = anchor_tag == "oneCellAnchor";
            let is_absolute = anchor_tag == "absoluteAnchor";
            if (anchor_tag != "twoCellAnchor" && !is_one_cell && !is_absolute)
                || !is_xdr_ns(anchor.tag_name().namespace())
            {
                continue;
            }

            let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
                (0u32, 0i64, 0u32, 0i64);
            let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
            // oneCellAnchor size: `<xdr:ext cx cy>` in EMU.
            let (mut ext_cx, mut ext_cy) = (0i64, 0i64);
            // absoluteAnchor position: `<xdr:pos x y>` in EMU.
            let (mut pos_x, mut pos_y) = (0i64, 0i64);

            for child in anchor.children() {
                if !child.is_element() {
                    continue;
                }
                match child.tag_name().name() {
                    "from" | "to" => {
                        let is_from = child.tag_name().name() == "from";
                        let mut col: u32 = 0;
                        let mut col_off: i64 = 0;
                        let mut row: u32 = 0;
                        let mut row_off: i64 = 0;
                        for c in child.children() {
                            match (c.tag_name().name(), c.text()) {
                                ("col", Some(t)) => col = t.trim().parse().unwrap_or(0),
                                ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                                ("row", Some(t)) => row = t.trim().parse().unwrap_or(0),
                                ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                                _ => {}
                            }
                        }
                        if is_from {
                            from_col = col;
                            from_col_off = col_off;
                            from_row = row;
                            from_row_off = row_off;
                        } else {
                            to_col = col;
                            to_col_off = col_off;
                            to_row = row;
                            to_row_off = row_off;
                        }
                    }
                    "ext" => {
                        // oneCellAnchor's `<xdr:ext cx cy>` size in EMU.
                        ext_cx = child
                            .attribute("cx")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        ext_cy = child
                            .attribute("cy")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }
                    "pos" => {
                        pos_x = child
                            .attribute("x")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        pos_y = child
                            .attribute("y")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }
                    _ => {}
                }
            }

            // For a oneCellAnchor the `<to>` corner is absent; the chart's
            // size is the saved EMU extent (ECMA-376 §20.5.2.24). Encode it as
            // a `to` corner pinned to the `from` cell plus the extent offset so
            // the renderer's from/to → pixel math yields exactly `ext` px.
            if is_one_cell {
                to_col = from_col;
                to_row = from_row;
                to_col_off = from_col_off + ext_cx;
                to_row_off = from_row_off + ext_cy;
            } else if is_absolute {
                // ChartAnchor offsets are signed EMU and the renderer already
                // turns a same-cell from/to pair into an exact pixel rectangle.
                // Reuse that representation instead of fabricating worksheet
                // cell indices for a chart sheet, which has no cell grid.
                from_col = 0;
                from_row = 0;
                from_col_off = pos_x;
                from_row_off = pos_y;
                to_col = 0;
                to_row = 0;
                to_col_off = pos_x.saturating_add(ext_cx);
                to_row_off = pos_y.saturating_add(ext_cy);
            }

            // Excel wraps chartEx in MCE either directly under the anchor or
            // inside an `xdr:grpSp`. Select the live Choice and
            // recurse through groups without ever visiting its Fallback.
            let mut graphic_frames = Vec::new();
            collect_selected_graphic_frames(anchor, &mut graphic_frames);
            for graphic_frame in graphic_frames {
                // ECMA-376 §20.1.2.2.8 CT_NonVisualDrawingProps@hidden: a hidden
                // chart's own graphicFrame is not rendered.
                if crate::drawing::xdr_node_hidden(&graphic_frame) {
                    continue;
                }

                // `<a:graphicData uri>` and its child namespace distinguish modern
                // `<cx:chart>` from legacy `<c:chart>`.
                let Some(graphic_data) = graphic_frame
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
                else {
                    continue;
                };
                let is_chartex = graphic_data
                    .attribute("uri")
                    .is_some_and(|uri| uri == CHARTEX_NS);
                let chart_rid = graphic_data
                    .descendants()
                    .find(|n| {
                        n.is_element()
                            && n.tag_name().name() == "chart"
                            && if is_chartex {
                                n.tag_name().namespace() == Some(CHARTEX_NS)
                            } else {
                                is_c_ns(n.tag_name().namespace())
                            }
                    })
                    .and_then(|chart| {
                        chart
                            .attributes()
                            .find(|a| a.name() == "id" && is_r_ns(a.namespace()))
                            .map(|a| a.value().to_string())
                    });

                let Some(rid) = chart_rid else {
                    continue;
                };
                let Some(chart_target) = drawing_rels.get(&rid) else {
                    continue;
                };
                let chart_path = resolve_zip_path(drawing_dir, chart_target);
                let Ok(chart_xml) = read_zip_string(archive, &chart_path) else {
                    continue;
                };
                // A chartEx part reads its title font size from the associated
                // chartStyle sidecar (`styleN.xml`), reached via the chart part's
                // OWN rels (`xl/charts/_rels/chartN.xml.rels`,
                // `.../2011/relationships/chartStyle`). Read it best-effort now
                // (before the chart doc is parsed, since both borrow `archive`);
                // legacy `<c:>` charts ignore it (their title size is inline).
                let style_xml = load_chart_style_xml(archive, &chart_path);
                let color_style_xml = load_chart_color_style_xml(archive, &chart_path);
                let image_relationships = load_chart_image_relationships(archive, &chart_path);
                let image_resolver = ooxml_common::chart::ChartImageResolverChain::new(
                    &image_relationships,
                    theme_images,
                );
                let user_shapes_xml = if is_chartex {
                    None
                } else {
                    load_chart_user_shapes_xml(archive, &chart_path, &chart_xml)
                };
                // Parse the chart directly through the shared `parse_chart_part`
                // (the single superset parser for pptx + xlsx). The xlsx theme
                // palette + major/minor Latin faces ride on the `XlsxColorResolver`,
                // so no `ChartData` intermediate / `From` adapter is needed.
                //
                // A chartEx part (`is_chartex`) has a `<cx:chartSpace>` root instead
                // of `<c:chartSpace>` and uses the shared
                // `parse_chartex_part` structure walk (waterfall / boxWhisker /
                // treemap / sunburst / … — ECMA-376 does not cover these; they are
                // the Microsoft 2014 chartex extension). Same `ColorResolver`.
                let Ok(chart_doc) = parse_guarded(&chart_xml) else {
                    continue;
                };
                let resolver = XlsxColorResolver {
                    theme_colors,
                    theme_major_font_latin: theme_fonts.0,
                    theme_minor_font_latin: theme_fonts.1,
                    theme_format_scheme,
                };
                let chart_opt = if let Some(context) = reference_context.as_mut() {
                    let mut references = XlsxChartReferenceResolver {
                        archive,
                        materialized_rows: context.materialized_rows,
                        materialized_col_hidden: context.materialized_col_hidden,
                        sheet_name: context.sheet_name,
                        sheets: context.sheets,
                        workbook_rels: context.workbook_rels,
                        shared_strings: context.shared_strings,
                        defined_names: context.defined_names,
                        number_formats: context.number_formats,
                        session: context.session,
                        visibility_cache: HashMap::new(),
                    };
                    if is_chartex {
                        ooxml_common::chart::parse_chartex_part_with_references_style_parts_and_images(
                            chart_doc.root_element(),
                            &resolver,
                            style_xml.as_deref(),
                            color_style_xml.as_deref(),
                            &mut references,
                            &image_resolver,
                        )
                    } else {
                        ooxml_common::chart::parse_chart_part_with_references_style_parts_and_images(
                            chart_doc.root_element(),
                            &resolver,
                            style_xml.as_deref(),
                            color_style_xml.as_deref(),
                            &mut references,
                            &image_resolver,
                        )
                    }
                } else if is_chartex {
                    ooxml_common::chart::parse_chartex_part_with_style_parts_and_images(
                        chart_doc.root_element(),
                        &resolver,
                        style_xml.as_deref(),
                        color_style_xml.as_deref(),
                        &image_resolver,
                    )
                } else {
                    ooxml_common::chart::parse_chart_part_with_style_parts_and_images(
                        chart_doc.root_element(),
                        &resolver,
                        style_xml.as_deref(),
                        color_style_xml.as_deref(),
                        &image_resolver,
                    )
                };
                let Some(mut chart) = chart_opt else {
                    continue;
                };
                if let Some(user_shapes_xml) = user_shapes_xml.as_deref() {
                    if let Ok(user_shapes_doc) = parse_guarded(user_shapes_xml) {
                        let text_boxes = ooxml_common::chart::parse_chart_user_shapes_for_chart(
                            chart_doc.root_element(),
                            user_shapes_doc.root_element(),
                            &resolver,
                        );
                        if !text_boxes.is_empty() {
                            chart.chart_text_boxes = Some(text_boxes);
                        }
                    }
                }

                all_charts.push(ChartAnchor {
                    z_order: graphic_frame.range().start as u64,
                    from_col,
                    from_col_off,
                    from_row,
                    from_row_off,
                    to_col,
                    to_col_off,
                    to_row,
                    to_row_off,
                    chart,
                });
            }
        }
    }
    all_charts
}

/// xlsx `ColorResolver` used by the shared [`ooxml_common::chart::parse_chart_part`].
/// Carries the workbook theme palette (clrScheme document order) plus the
/// theme's major/minor Latin font faces so the shared parser can supply the
/// chart-text font fallbacks without a separate `theme_fonts` parameter.
struct XlsxColorResolver<'a> {
    theme_colors: &'a [String],
    theme_major_font_latin: Option<&'a str>,
    theme_minor_font_latin: Option<&'a str>,
    theme_format_scheme: Option<&'a ooxml_common::theme::ThemeFormatScheme>,
}

impl ooxml_common::chart::ColorResolver for XlsxColorResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node<'_, '_>) -> Option<String> {
        resolve_fill_color(&node, self.theme_colors)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        ooxml_common::color::ThemeResolver::resolve_scheme_color(
            &crate::drawing::XlsxSchemeResolver {
                theme_colors: self.theme_colors,
            },
            name,
        )
        .map(|hex| hex.to_lowercase())
    }

    /// Shape fills (series / marker / dPt / errBars) resolve through the FULL
    /// DrawingML color grammar (transforms included), matching the historical
    /// `extract_solid_fill_in_drawingml` path so a scheme-color marker with a
    /// `lumMod`/`lumOff` tint renders at the right strength. This is deliberately
    /// heavier than [`Self::resolve_solid_fill`] (which xlsx keeps for callers
    /// that pass an already-selected solidFill node).
    fn resolve_shape_fill(&self, parent: roxmltree::Node<'_, '_>) -> Option<String> {
        extract_solid_fill_in_drawingml(&parent, self.theme_colors)
    }

    /// Default series fill: `theme.accent[(idx % 6) + 1]`. The palette is stored
    /// in clrScheme document order (dk1@0, lt1@1, dk2@2, lt2@3, accent1@4 …
    /// accent6@9), so accent1 is index 4.
    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        self.theme_colors
            .get(4 + (idx % 6))
            .map(|c| c.trim_start_matches('#').to_lowercase())
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        self.theme_major_font_latin.map(|s| s.to_string())
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        self.theme_minor_font_latin.map(|s| s.to_string())
    }

    fn theme_format_scheme(&self) -> Option<&ooxml_common::theme::ThemeFormatScheme> {
        self.theme_format_scheme
    }

    /// Excel paints an opaque-white chart area when the file omits
    /// `<c:chartSpace><c:spPr>` entirely (the historical `has_chart_sp_pr=false`
    /// white default).
    fn default_chart_bg(&self) -> Option<String> {
        Some("FFFFFF".to_string())
    }

    /// Excel's automatic plot-area paint is opaque white even when the chart
    /// area itself has a different authored fill. A direct plot-area fill or
    /// noFill remains authoritative in the shared parser.
    fn default_plot_area_bg(&self) -> Option<String> {
        Some("FFFFFF".to_string())
    }

    fn implicit_outline_only_negative_column_style(&self) -> bool {
        true
    }
}
/// Locate the first resolvable `<a:solidFill>` among `parent`'s direct children
/// (children only, not deep descendants — chart spPr is structured shallowly)
/// and resolve its color to hex **without** `#` (uppercase). The chart wire
/// model prepends `#` on the TS side, so this matches every other chart color
/// field.
///
/// Delegates the DrawingML color grammar (`srgbClr`/`sysClr`/`prstClr`/
/// `schemeClr` + `lumMod`/`lumOff`/`tint`/`shade` transforms) to the
/// shared [`ooxml_common::color::parse_color_node`] via the crate-wide
/// [`XlsxSchemeResolver`], so scheme slots resolve through the §20.1.6.2 default
/// clrMap (`tx2`→`dk2`, `bg2`→`lt2`) and luminance transforms apply in HLS space
/// (§20.1.2.3.20/.21). The prior private copy in this module mapped `bg2`/`tx2`
/// to the wrong slots and multiplied `lumMod`/`lumOff` in RGB space.
pub(crate) fn extract_solid_fill_in_drawingml(
    parent: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    parent
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .find_map(|fill| {
            ooxml_common::color::parse_color_node(
                fill,
                &crate::drawing::XlsxSchemeResolver { theme_colors },
                ooxml_common::color::TintMode::PowerPointLinear,
            )
        })
}

#[cfg(test)]
mod solid_fill_color_tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Theme in clrScheme document order: dk1@0, lt1@1, dk2@2, lt2@3,
    // accent1@4 … folHlink@11. Distinct hexes so a mis-index is obvious.
    fn theme() -> Vec<String> {
        vec![
            "#111111".into(), // dk1 @0
            "#FEFEFE".into(), // lt1 @1
            "#222222".into(), // dk2 @2
            "#EEEEEE".into(), // lt2 @3
            "#4472C4".into(), // accent1 @4
            "#00AA00".into(), // accent2 @5
            "#0000AA".into(), // accent3 @6
            "#AAAA00".into(), // accent4 @7
            "#00AAAA".into(), // accent5 @8
            "#AA00AA".into(), // accent6 @9
            "#0563C1".into(), // hlink @10
            "#954F72".into(), // folHlink @11
        ]
    }

    fn solid_fill(inner: &str) -> String {
        format!(r#"<a:spPr xmlns:a="{A_NS}"><a:solidFill>{inner}</a:solidFill></a:spPr>"#)
    }

    /// §20.1.6.2 default clrMap: `tx2` → `dk2` (theme slot 2), NOT `lt1`.
    #[test]
    fn scheme_tx2_resolves_to_dk2_slot() {
        let xml = solid_fill(r#"<a:schemeClr val="tx2"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        // tx2 → dk2 → theme[2] = "222222" (uppercase, no `#`).
        assert_eq!(out.as_deref(), Some("222222"));
    }

    /// §20.1.6.2 default clrMap: `bg2` → `lt2` (theme slot 3), NOT `dk1`.
    #[test]
    fn scheme_bg2_resolves_to_lt2_slot() {
        let xml = solid_fill(r#"<a:schemeClr val="bg2"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        // bg2 → lt2 → theme[3] = "EEEEEE".
        assert_eq!(out.as_deref(), Some("EEEEEE"));
    }

    /// tx1 → dk1 and bg1 → lt1 (unchanged, but pinned so a refactor can't drift).
    #[test]
    fn scheme_tx1_bg1_resolve_to_dk1_lt1() {
        let tx1 = solid_fill(r#"<a:schemeClr val="tx1"/>"#);
        let doc = Document::parse(&tx1).unwrap();
        assert_eq!(
            extract_solid_fill_in_drawingml(&doc.root_element(), &theme()).as_deref(),
            Some("111111") // dk1 @0
        );
        let bg1 = solid_fill(r#"<a:schemeClr val="bg1"/>"#);
        let doc = Document::parse(&bg1).unwrap();
        assert_eq!(
            extract_solid_fill_in_drawingml(&doc.root_element(), &theme()).as_deref(),
            Some("FEFEFE") // lt1 @1
        );
    }

    /// `lumMod` is a luminance modulation applied to the HLS `L` channel
    /// (§20.1.2.3.20), NOT a per-RGB-component multiply. For `4472C4` at
    /// `lumMod 50000`, the HLS result is `203864`; the (wrong) RGB-space
    /// multiply would give `223962`.
    #[test]
    fn lummod_applies_in_hls_space_not_rgb() {
        let xml = solid_fill(r#"<a:srgbClr val="4472C4"><a:lumMod val="50000"/></a:srgbClr>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        assert_eq!(out.as_deref(), Some("203864"));
    }

    /// A plain srgbClr with no transforms passes through (uppercased, no `#`).
    #[test]
    fn plain_srgb_passthrough() {
        let xml = solid_fill(r#"<a:srgbClr val="ff8000"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        assert_eq!(out.as_deref(), Some("FF8000"));
    }

    /// A chart series is a DrawingML shape too: its `<c:spPr>` fill must retain
    /// color transforms such as `lumMod` (§20.1.2.3.20). Excel commonly writes
    /// a grey series as `bg1` (white) with a 65% luminance modulation.
    #[test]
    fn chart_series_scheme_fill_applies_lummod() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="{A_NS}">
                 <c:chart><c:plotArea><c:barChart>
                   <c:barDir val="col"/>
                   <c:grouping val="clustered"/>
                   <c:ser>
                     <c:idx val="0"/><c:order val="0"/>
                     <c:spPr><a:solidFill><a:schemeClr val="bg1">
                       <a:lumMod val="65000"/>
                     </a:schemeClr></a:solidFill></c:spPr>
                     <c:val><c:numLit><c:ptCount val="1"/>
                       <c:pt idx="0"><c:v>1</c:v></c:pt>
                     </c:numLit></c:val>
                   </c:ser>
                 </c:barChart></c:plotArea></c:chart>
               </c:chartSpace>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let mut colors = theme();
        colors[1] = "#FFFFFF".into();
        let resolver = XlsxColorResolver {
            theme_colors: &colors,
            theme_major_font_latin: None,
            theme_minor_font_latin: None,
            theme_format_scheme: None,
        };
        let chart = ooxml_common::chart::parse_chart_part(doc.root_element(), &resolver).unwrap();
        assert_eq!(chart.series[0].color.as_deref(), Some("A6A6A6"));
    }

    #[test]
    fn chart_part_honors_chart_local_color_map_override() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="{A_NS}">
              <c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent2" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink"/>
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let palette = theme();
        let resolver = XlsxColorResolver {
            theme_colors: &palette,
            theme_major_font_latin: None,
            theme_minor_font_latin: None,
            theme_format_scheme: None,
        };
        let doc = Document::parse(&xml).expect("chartSpace fixture");

        let chart = ooxml_common::chart::parse_chart_part(doc.root_element(), &resolver)
            .expect("chart should parse");

        assert_eq!(chart.series[0].color.as_deref(), Some("00AA00"));
    }

    #[test]
    fn xlsx_host_enables_the_bounded_outline_only_negative_column_default() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="{A_NS}">
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="col"/><c:varyColors val="0"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  <c:val><c:numLit><c:pt idx="0"><c:v>-24000</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>"#
        );
        let palette = theme();
        let resolver = XlsxColorResolver {
            theme_colors: &palette,
            theme_major_font_latin: None,
            theme_minor_font_latin: None,
            theme_format_scheme: None,
        };
        let doc = Document::parse(&xml).expect("chartSpace fixture");
        let chart = ooxml_common::chart::parse_chart_part(doc.root_element(), &resolver)
            .expect("chart should parse");
        let series = &chart.series[0];
        assert_eq!(series.invert_if_negative, None);
        assert_eq!(series.automatic_negative_style, Some(true));
        assert_eq!(series.inverted_fill_hidden, Some(true));
        assert_eq!(series.inverted_line_color.as_deref(), Some("000000"));
        assert_eq!(series.inverted_line_width_emu, Some(9_525));
    }
}

/// §20.1.2.2.8 — an `<xdr:cNvPr hidden="1">` graphicFrame is not rendered.
/// `load_sheet_charts` walks `<xdr:graphicFrame>` independently of the shared
/// shape walker in `drawing.rs::collect_shapes`, so it needs its own hidden
/// check — this covers that walk specifically (full sheet → drawing → chart
/// zip round trip, since `load_sheet_charts` reads from the archive).
#[cfg(test)]
mod hidden_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const NS: &str = concat!(
        r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
    );

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    fn minimal_chart_xml() -> String {
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:valAx><c:axId val="1"/><c:axPos val="l"/></c:valAx>
      <c:catAx><c:axId val="2"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
            .to_string()
    }

    fn picture_marker_chart_xml() -> String {
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:lineChart><c:ser>
    <c:idx val="0"/><c:order val="0"/>
    <c:marker><c:symbol val="picture"/><c:size val="9"/><c:spPr><a:blipFill>
      <a:blip r:embed="rIdMarker"/><a:srcRect l="10000"/><a:stretch/>
    </a:blipFill></c:spPr></c:marker>
    <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
    <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
  </c:ser></c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#.to_owned()
    }

    fn drawing_xml(hidden_attr: &str) -> String {
        format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"{hidden}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></xdr:xfrm>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
                </a:graphicData></a:graphic>
              </xdr:graphicFrame>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
            hidden = hidden_attr,
        )
    }

    /// Builds a minimal zip archive wiring `xl/worksheets/sheet1.xml`'s rels →
    /// `xl/drawings/drawing1.xml` → its own rels → `xl/charts/chart1.xml`, the
    /// same part chain `load_sheet_charts` walks in production.
    fn archive_with_chart(hidden_attr: &str) -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();

            zw.start_file("xl/worksheets/_rels/sheet1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(drawing_xml(hidden_attr).as_bytes()).unwrap();

            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/charts/chart1.xml", o).unwrap();
            zw.write_all(minimal_chart_xml().as_bytes()).unwrap();

            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    fn archive_with_picture_marker() -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();
            zw.start_file("xl/worksheets/_rels/sheet1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(drawing_xml("").as_bytes()).unwrap();
            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/charts/chart1.xml", o).unwrap();
            zw.write_all(picture_marker_chart_xml().as_bytes()).unwrap();
            zw.start_file("xl/charts/_rels/chart1.xml.rels", o).unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/marker.png"/></Relationships>"#).unwrap();
            zw.start_file("xl/media/marker.png", o).unwrap();
            zw.write_all(b"png").unwrap();
            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    fn archive_with_custom_drawing(drawing: &str, sheet_dir: &str) -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();
            zw.start_file(format!("xl/{sheet_dir}/_rels/sheet1.xml.rels"), o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(drawing.as_bytes()).unwrap();
            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/charts/chart1.xml", o).unwrap();
            zw.write_all(minimal_chart_xml().as_bytes()).unwrap();
            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    fn archive_with_chart_user_shapes() -> crate::XlsxZip {
        let mut chart_xml = minimal_chart_xml();
        chart_xml = chart_xml.replace(
            "</c:chartSpace>",
            r#"<c:userShapes xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdUserShapes"/></c:chartSpace>"#,
        );
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();
            zw.start_file("xl/worksheets/_rels/sheet1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(drawing_xml("").as_bytes()).unwrap();
            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/charts/chart1.xml", o).unwrap();
            zw.write_all(chart_xml.as_bytes()).unwrap();
            zw.start_file("xl/charts/_rels/chart1.xml.rels", o).unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUserShapes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/chartDrawing1.xml"/></Relationships>"#).unwrap();
            zw.start_file("xl/drawings/chartDrawing1.xml", o).unwrap();
            zw.write_all(br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>0.1</cdr:y></cdr:to><cdr:sp><cdr:nvSpPr><cdr:cNvPr id="1" name="TitleBox"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr><cdr:spPr/><cdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2000" b="1"/><a:t>User-shape title</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>"#).unwrap();
            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    #[test]
    fn hidden_chart_graphicframe_is_not_emitted() {
        for attr in [r#" hidden="1""#, r#" hidden="true""#] {
            let mut archive = archive_with_chart(attr);
            let charts = load_sheet_charts(
                &mut archive,
                "worksheets/sheet1.xml",
                None,
                &theme(),
                (None, None),
                None,
            );
            assert!(charts.is_empty(), "hidden chart emitted (attr={attr})");
        }
    }

    #[test]
    fn visible_chart_graphicframe_is_emitted_unchanged() {
        for attr in ["", r#" hidden="0""#, r#" hidden="false""#] {
            let mut archive = archive_with_chart(attr);
            let charts = load_sheet_charts(
                &mut archive,
                "worksheets/sheet1.xml",
                None,
                &theme(),
                (None, None),
                None,
            );
            assert_eq!(charts.len(), 1, "visible chart dropped (attr={attr})");
        }
    }

    #[test]
    fn picture_marker_relationship_reaches_the_shared_chart_model() {
        let mut archive = archive_with_picture_marker();
        let charts = load_sheet_charts(
            &mut archive,
            "worksheets/sheet1.xml",
            None,
            &theme(),
            (None, None),
            None,
        );
        let series = &charts[0].chart.series[0];
        assert_eq!(series.marker_symbol.as_deref(), Some("picture"));
        assert!(matches!(
            series.marker_fill_paint.as_ref(),
            Some(ooxml_common::chart::ChartStyleFill::Image {
                image_path,
                mime_type,
                src_rect: Some(src_rect),
                ..
            }) if image_path == "xl/media/marker.png"
                && mime_type == "image/png"
                && (src_rect.l - 0.1).abs() < 1e-9
        ));
    }

    /// ECMA-376 §20.5.2.1 `absoluteAnchor` is the chart-sheet placement form:
    /// one absolute EMU position plus one absolute EMU extent. It must reach
    /// the same ChartAnchor wire model without inventing a cell span.
    #[test]
    fn absolute_anchor_chart_uses_position_and_extent_offsets() {
        let drawing = format!(
            r#"<xdr:wsDr {NS}><xdr:absoluteAnchor>
              <xdr:pos x="914400" y="457200"/><xdr:ext cx="3657600" cy="2743200"/>
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="914400" y="457200"/><a:ext cx="3657600" cy="2743200"/></xdr:xfrm>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
                </a:graphicData></a:graphic>
              </xdr:graphicFrame><xdr:clientData/>
            </xdr:absoluteAnchor></xdr:wsDr>"#,
        );
        let mut archive = archive_with_custom_drawing(&drawing, "chartsheets");
        let charts = load_sheet_charts(
            &mut archive,
            "chartsheets/sheet1.xml",
            None,
            &theme(),
            (None, None),
            None,
        );
        assert_eq!(charts.len(), 1);
        let anchor = &charts[0];
        assert_eq!((anchor.from_col, anchor.from_row), (0, 0));
        assert_eq!((anchor.to_col, anchor.to_row), (0, 0));
        assert_eq!((anchor.from_col_off, anchor.from_row_off), (914400, 457200));
        assert_eq!((anchor.to_col_off, anchor.to_row_off), (4572000, 3200400));
    }

    #[test]
    fn chart_user_shapes_relationship_populates_shared_text_box_model() {
        let mut archive = archive_with_chart_user_shapes();
        let charts = load_sheet_charts(
            &mut archive,
            "worksheets/sheet1.xml",
            None,
            &theme(),
            (None, None),
            None,
        );
        let boxes = charts[0]
            .chart
            .chart_text_boxes
            .as_ref()
            .expect("chart user-shape text boxes");
        assert_eq!(boxes.len(), 1);
        assert_eq!(boxes[0].paragraphs[0].runs[0].text, "User-shape title");
        assert_eq!(boxes[0].paragraphs[0].runs[0].font_size_hpt, Some(2000));
        assert_eq!(boxes[0].paragraphs[0].runs[0].bold, Some(true));
        assert_eq!(boxes[0].l_ins, ooxml_common::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].r_ins, ooxml_common::text::DEFAULT_INS_LR_EMU);
        assert_eq!(boxes[0].t_ins, ooxml_common::text::DEFAULT_INS_TB_EMU);
        assert_eq!(boxes[0].b_ins, ooxml_common::text::DEFAULT_INS_TB_EMU);
    }
}

#[cfg(test)]
mod worksheet_reference_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const DATA_XML: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="C1" t="inlineStr"><is><t>المبيعات</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>أحمد</t></is></c><c r="C2" s="1"><v>5000</v></c><c r="D2"><v>10</v></c><c r="E2"><v>3</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>سارة</t></is></c><c r="C3" s="1"><v>6200</v></c><c r="D3"><v>20</v></c><c r="E3"><v>5</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>خالد</t></is></c><c r="C4" s="1"><v>7500</v></c><c r="D4"><v>30</v></c><c r="E4"><v>7</v></c></row></sheetData></worksheet>"#;

    fn chart_xml(with_cache: bool) -> String {
        let name_cache = if with_cache {
            r#"<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cached</c:v></c:pt></c:strCache>"#
        } else {
            ""
        };
        let category_cache = if with_cache {
            r#"<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cached category</c:v></c:pt></c:strCache>"#
        } else {
            ""
        };
        let value_cache = if with_cache {
            r#"<c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache>"#
        } else {
            ""
        };
        format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>'التقرير'!C1</c:f>{name_cache}</c:strRef></c:tx><c:cat><c:strRef><c:f>'التقرير'!$A$2:$A$4</c:f>{category_cache}</c:strRef></c:cat><c:val><c:numRef><c:f>'التقرير'!$C$2:$C$4</c:f>{value_cache}</c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="100"/></c:barChart><c:catAx><c:axId val="10"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="100"/><c:axPos val="l"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )
    }

    fn cacheless_bubble_chart_xml() -> String {
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>'التقرير'!C1</c:f></c:strRef></c:tx><c:xVal><c:numRef><c:f>'التقرير'!C2:C4</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>'التقرير'!D2:D4</c:f></c:numRef></c:yVal><c:bubbleSize><c:numRef><c:f>'التقرير'!E2:E4</c:f></c:numRef></c:bubbleSize></c:ser></c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#.into()
    }

    fn archive_with_chart_and_custom_data(chart_xml: &str, data_xml: &str) -> crate::XlsxZip {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = SimpleFileOptions::default();
            writer
                .start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
                .unwrap();
            writer.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();
            writer
                .start_file("xl/drawings/drawing1.xml", options)
                .unwrap();
            writer.write_all(br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#).unwrap();
            writer
                .start_file("xl/drawings/_rels/drawing1.xml.rels", options)
                .unwrap();
            writer.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();
            writer.start_file("xl/charts/chart1.xml", options).unwrap();
            writer.write_all(chart_xml.as_bytes()).unwrap();
            writer
                .start_file("xl/worksheets/sheet2.xml", options)
                .unwrap();
            writer.write_all(data_xml.as_bytes()).unwrap();
            writer.start_file("xl/styles.xml", options).unwrap();
            writer.write_all(br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="3"/></cellXfs></styleSheet>"#).unwrap();
            writer.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(bytes)).unwrap()
    }

    fn archive_with_chart_and_data(chart_xml: &str) -> crate::XlsxZip {
        archive_with_chart_and_custom_data(chart_xml, DATA_XML)
    }

    fn sheets() -> Vec<SheetMeta> {
        vec![
            SheetMeta {
                name: "Dashboard".into(),
                sheet_id: 1,
                r_id: "rDashboard".into(),
                tab_color: None,
                visibility: SheetVisibility::Visible,
            },
            SheetMeta {
                name: "التقرير".into(),
                sheet_id: 2,
                r_id: "rData".into(),
                tab_color: None,
                visibility: SheetVisibility::Visible,
            },
        ]
    }

    fn workbook_rels_xml() -> &'static str {
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDashboard" Target="worksheets/sheet1.xml"/><Relationship Id="rData" Target="worksheets/sheet2.xml"/></Relationships>"#
    }

    fn load_model(xml: &str) -> ooxml_common::chart::ChartModel {
        let mut archive = archive_with_chart_and_data(xml);
        let rels = parse_guarded(workbook_rels_xml()).unwrap();
        let sheet_metas = sheets();
        let theme = vec!["#4472C4".into(); 12];
        let mut session = WorksheetReferenceSession::default();
        let styles = crate::styles::parse_styles(&mut archive, &theme)
            .expect("styles parse for chart references");
        let number_formats = ChartNumberFormatCache::from_styles(&styles.styles);
        session.seed_current_sheet("Dashboard", None);
        let charts = load_sheet_charts(
            &mut archive,
            "worksheets/sheet1.xml",
            Some(ChartReferenceContext {
                materialized_rows: None,
                materialized_col_hidden: None,
                sheet_name: "Dashboard",
                sheets: &sheet_metas,
                workbook_rels: &rels,
                shared_strings: &[],
                defined_names: &[],
                number_formats: &number_formats,
                session: &mut session,
            }),
            &theme,
            (None, None),
            None,
        );
        assert_eq!(charts.len(), 1);
        charts.into_iter().next().unwrap().chart
    }

    #[test]
    fn cacheless_unicode_chart_resolves_cross_sheet_series() {
        let xml = chart_xml(false);
        let chart = load_model(&xml);

        assert_eq!(chart.series[0].name, "المبيعات");
        assert_eq!(chart.categories, vec!["أحمد", "سارة", "خالد"]);
        assert_eq!(chart.series[0].categories, None);
        assert_eq!(
            chart.series[0].values,
            vec![Some(5000.0), Some(6200.0), Some(7500.0)],
        );
    }

    #[test]
    fn zero_point_chart_caches_resolve_cross_sheet_series() {
        let xml = chart_xml(false)
            .replace(
                "</c:strRef>",
                "<c:strCache><c:ptCount val=\"0\"/></c:strCache></c:strRef>",
            )
            .replace(
                "</c:numRef>",
                "<c:numCache><c:ptCount val=\"0\"/></c:numCache></c:numRef>",
            );
        let chart = load_model(&xml);

        assert_eq!(chart.series[0].name, "المبيعات");
        assert_eq!(chart.categories, vec!["أحمد", "سارة", "خالد"]);
        assert_eq!(
            chart.series[0].values,
            vec![Some(5000.0), Some(6200.0), Some(7500.0)],
        );
    }

    #[test]
    fn plot_visible_only_keeps_cached_values_and_resolves_hidden_rows_and_columns() {
        let chart_xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
          <c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>'التقرير'!A2:A4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>'التقرير'!D2:D4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
          <c:ser><c:idx val="1"/><c:order val="1"/><c:cat><c:strRef><c:f>'التقرير'!A2:A4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>'التقرير'!C2:C4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>50</c:v></c:pt><c:pt idx="1"><c:v>60</c:v></c:pt><c:pt idx="2"><c:v>70</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
        </c:lineChart></c:plotArea><c:plotVisOnly/></c:chart></c:chartSpace>"#;
        let data_xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="3" max="3" hidden="1"/></cols><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="C2"><v>500</v></c><c r="D2"><v>100</v></c></row><row r="3" hidden="1"/><row r="4"><c r="A4" t="inlineStr"><is><t>C</t></is></c><c r="C4"><v>700</v></c><c r="D4"><v>300</v></c></row></sheetData></worksheet>"#;
        let mut archive = archive_with_chart_and_custom_data(chart_xml, data_xml);
        let rels = parse_guarded(workbook_rels_xml()).unwrap();
        let sheet_metas = sheets();
        let theme = vec!["#4472C4".into(); 12];
        let mut session = WorksheetReferenceSession::default();
        let styles = crate::styles::parse_styles(&mut archive, &theme).expect("styles");
        let number_formats = ChartNumberFormatCache::from_styles(&styles.styles);
        session.seed_current_sheet("Dashboard", None);
        let charts = load_sheet_charts(
            &mut archive,
            "worksheets/sheet1.xml",
            Some(ChartReferenceContext {
                materialized_rows: None,
                materialized_col_hidden: None,
                sheet_name: "Dashboard",
                sheets: &sheet_metas,
                workbook_rels: &rels,
                shared_strings: &[],
                defined_names: &[],
                number_formats: &number_formats,
                session: &mut session,
            }),
            &theme,
            (None, None),
            None,
        );
        let chart = &charts[0].chart;
        assert_eq!(chart.plot_visible_only, Some(true));
        assert_eq!(chart.categories, vec!["A", "B", "C"]);
        assert_eq!(chart.category_source_hidden, Some(vec![false, true, false]));
        assert_eq!(
            chart.series[0].values,
            vec![Some(10.0), Some(20.0), Some(30.0)]
        );
        assert_eq!(
            chart.series[0].source_hidden,
            Some(vec![false, true, false])
        );
        assert_eq!(
            chart.series[1].values,
            vec![Some(50.0), Some(60.0), Some(70.0)]
        );
        assert_eq!(chart.series[1].source_hidden, Some(vec![true, true, true]));
    }

    #[test]
    fn chart_reference_resolves_source_linked_builtin_number_format() {
        let mut archive = archive_with_chart_and_data(&chart_xml(false));
        let rels = parse_guarded(workbook_rels_xml()).unwrap();
        let sheet_metas = sheets();
        let mut session = WorksheetReferenceSession::default();
        let theme = vec!["#4472C4".into(); 12];
        let styles = crate::styles::parse_styles(&mut archive, &theme)
            .expect("styles parse for chart references");
        let number_formats = ChartNumberFormatCache::from_styles(&styles.styles);
        session.seed_current_sheet("Dashboard", None);
        let mut resolver = XlsxChartReferenceResolver {
            archive: &mut archive,
            materialized_rows: None,
            materialized_col_hidden: None,
            sheet_name: "Dashboard",
            sheets: &sheet_metas,
            workbook_rels: &rels,
            shared_strings: &[],
            defined_names: &[],
            number_formats: &number_formats,
            session: &mut session,
            visibility_cache: HashMap::new(),
        };
        assert_eq!(
            ooxml_common::chart::ChartReferenceResolver::resolve_number_format(
                &mut resolver,
                "'التقرير'!$C$2:$C$4",
            )
            .as_deref(),
            Some("#,##0"),
        );
        assert_eq!(
            ooxml_common::chart::ChartReferenceResolver::resolve_number_format_id(
                &mut resolver,
                "'التقرير'!$C$2:$C$4",
            ),
            Some(3),
        );
    }

    #[test]
    fn authored_chart_caches_take_precedence_over_live_cells() {
        let xml = chart_xml(true);
        let chart = load_model(&xml);

        assert_eq!(chart.series[0].name, "Cached");
        assert_eq!(chart.categories, vec!["Cached category"]);
        assert_eq!(chart.series[0].values, vec![Some(99.0)]);
    }

    #[test]
    fn cacheless_bubble_chart_resolves_all_series_fields_through_loader() {
        let chart = load_model(&cacheless_bubble_chart_xml());

        assert_eq!(chart.categories, vec!["5000", "6200", "7500"]);
        assert_eq!(chart.series[0].name, "المبيعات");
        assert_eq!(
            chart.series[0].values,
            vec![Some(10.0), Some(20.0), Some(30.0)]
        );
        assert_eq!(
            chart.series[0].bubble_sizes,
            Some(vec![Some(3.0), Some(5.0), Some(7.0)])
        );
    }
}

/// CH14 — chartEx (Microsoft 2014 `cx:` namespace) recognition for xlsx.
/// `xdr:graphicFrame` wires a chartEx part through `<cx:chart r:id>`; the
/// `<a:graphicData@uri>` and the child's namespace both distinguish it
/// (`http://schemas.microsoft.com/office/drawing/2014/chartex` vs the
/// DrawingML chart URI). This exercises the full zip → drawing → chartEx-part
/// round trip with a self-contained inline waterfall fixture, mirroring
/// `parse_chartex_part_waterfall_full_contract` in `ooxml-common`'s chart tests.
#[cfg(test)]
mod chartex_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const NS: &str = concat!(
        r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
    );
    const CX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    /// A minimal waterfall chartEx part: one category dimension, one value
    /// dimension with a negative point, and the `cx:series layoutId` that
    /// `parse_chartex_part` reads as the chart type.
    fn waterfall_chartex_xml() -> String {
        format!(
            r#"<cx:chartSpace xmlns:cx="{CX_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <cx:chartData>
                <cx:data id="0">
                  <cx:strDim type="cat">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">Start</cx:pt>
                      <cx:pt idx="1">Change</cx:pt>
                      <cx:pt idx="2">End</cx:pt>
                    </cx:lvl>
                  </cx:strDim>
                  <cx:numDim type="val">
                    <cx:lvl ptCount="3">
                      <cx:pt idx="0">50</cx:pt>
                      <cx:pt idx="1">-15</cx:pt>
                      <cx:pt idx="2">35</cx:pt>
                    </cx:lvl>
                  </cx:numDim>
                </cx:data>
              </cx:chartData>
              <cx:chart>
                <cx:plotArea>
                  <cx:plotAreaRegion>
                    <cx:series layoutId="waterfall"/>
                  </cx:plotAreaRegion>
                </cx:plotArea>
              </cx:chart>
            </cx:chartSpace>"#
        )
    }

    /// `<xdr:graphicFrame>` for a chartEx part. Structurally identical to the
    /// legacy `drawing_xml` fixture in `hidden_tests` except for the
    /// `graphicData@uri`, which is exactly the wire-format signal
    /// `load_sheet_charts` now checks.
    fn chartex_drawing_xml() -> String {
        format!(
            r#"<xdr:wsDr {NS}
              xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
              xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex">
            <xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:grpSp>
                <xdr:nvGrpSpPr><xdr:cNvPr id="1" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
                <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="3000000"/></a:xfrm></xdr:grpSpPr>
              <mc:AlternateContent>
                <mc:Choice Requires="cx1">
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></xdr:xfrm>
                <a:graphic><a:graphicData uri="{CX_NS}">
                  <cx:chart xmlns:cx="{CX_NS}" r:id="rIdChart"/>
                </a:graphicData></a:graphic>
              </xdr:graphicFrame>
                </mc:Choice>
                <mc:Fallback><xdr:pic/></mc:Fallback>
              </mc:AlternateContent>
              </xdr:grpSp>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
            CX_NS = CX_NS,
        )
    }

    fn classic_drawing_xml() -> String {
        format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
              <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></xdr:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
              </a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
        )
    }

    fn classic_line_chart_xml() -> &'static str {
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:dropLines/></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#
    }

    /// Builds the same `sheet1.xml.rels` → `drawing1.xml` → `drawing1.xml.rels`
    /// → `charts/chartEx1.xml` chain Excel uses, including the Microsoft
    /// `.../2014/relationships/chartEx` relationship type.
    fn archive_with_chartex_chart() -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();

            zw.start_file("xl/worksheets/_rels/sheet1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(chartex_drawing_xml().as_bytes()).unwrap();

            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/charts/chartEx1.xml", o).unwrap();
            zw.write_all(waterfall_chartex_xml().as_bytes()).unwrap();

            zw.start_file("xl/charts/_rels/chartEx1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/charts/style1.xml", o).unwrap();
            zw.write_all(br#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:dataPoint><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:spPr><a:pattFill prst="diagCross"><a:fgClr><a:schemeClr val="phClr"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></cs:spPr></cs:dataPoint><cs:dataPointMarker><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef></cs:dataPointMarker><cs:dataLabelCallout><cs:defRPr><a:noFill/></cs:defRPr><cs:bodyPr/></cs:dataLabelCallout><cs:trendlineLabel><cs:defRPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></cs:defRPr></cs:trendlineLabel></cs:chartStyle>"#).unwrap();
            zw.start_file("xl/charts/_rels/style1.xml.rels", o).unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/style-marker.png"/></Relationships>"#).unwrap();
            zw.start_file("xl/media/style-marker.png", o).unwrap();
            zw.write_all(b"png").unwrap();

            zw.start_file("xl/charts/colors1.xml", o).unwrap();
            zw.write_all(br#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="336699"/></cs:colorStyle>"#).unwrap();

            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    fn archive_with_classic_chart_style() -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();
            for (path, xml) in [
                (
                    "xl/worksheets/_rels/sheet1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
                ),
                ("xl/drawings/drawing1.xml", &classic_drawing_xml()),
                (
                    "xl/drawings/_rels/drawing1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
                ),
                ("xl/charts/chart1.xml", classic_line_chart_xml()),
                (
                    "xl/charts/_rels/chart1.xml.rels",
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>"#,
                ),
                (
                    "xl/charts/style1.xml",
                    r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:dropLine><cs:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr></cs:dropLine></cs:chartStyle>"#,
                ),
                (
                    "xl/charts/colors1.xml",
                    r#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="336699"/></cs:colorStyle>"#,
                ),
            ] {
                zw.start_file(path, o).unwrap();
                zw.write_all(xml.as_bytes()).unwrap();
            }
            zw.finish().unwrap();
        }
        crate::XlsxZip::new(Cursor::new(buf)).unwrap()
    }

    #[test]
    fn chartex_graphicframe_parses_through_parse_chartex_part() {
        let mut archive = archive_with_chartex_chart();
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:themeElements><a:fmtScheme name="Theme"><a:fillStyleLst><a:blipFill><a:blip r:embed="rIdThemeMarker"/><a:stretch/></a:blipFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#;
        let theme_scheme = ooxml_common::theme::ThemeFormatScheme::parse(theme_xml);
        let mut theme_images = ooxml_common::chart::ChartImageRelationships::default();
        theme_images.insert_part_relationships(
            ooxml_common::chart::ChartImageSource::Theme,
            "xl/theme/theme1.xml",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdThemeMarker" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/theme-marker.png"/></Relationships>"#,
        );
        let charts = load_sheet_charts_with_theme_images(
            &mut archive,
            "worksheets/sheet1.xml",
            None,
            &theme(),
            (None, None),
            Some(&theme_scheme),
            &theme_images,
        );
        assert_eq!(
            charts.len(),
            1,
            "chartEx graphicFrame did not produce a chart"
        );
        let chart = &charts[0].chart;
        assert_eq!(chart.chart_type, "waterfall");
        assert_eq!(chart.series.len(), 1, "expected exactly one chartEx series");
        assert_eq!(
            chart.categories,
            vec!["Start".to_string(), "Change".to_string(), "End".to_string()]
        );
        assert_eq!(
            chart.chartex_color_palette.as_deref(),
            Some(&[Some("336699".to_string())][..]),
        );
        assert!(matches!(
            chart
                .chartex_data_point_style
                .as_ref()
                .and_then(|style| style.fill_paints.as_ref())
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ooxml_common::chart::ChartStyleFill::Pattern { fg, bg, preset })
                if fg == "336699" && bg == "FFFFFF" && preset == "diagCross"
        ));
        let roles = chart.chart_style_roles.as_ref().expect("linked role table");
        assert_eq!(roles["dataLabelCallout"].font_hidden, Some(true));
        assert_eq!(roles["dataLabelCallout"].text_body_authored, Some(true));
        assert_eq!(
            roles["trendlineLabel"].font_color.as_deref(),
            Some("112233")
        );
        assert!(matches!(
            roles["dataPointMarker"]
                .fill_paints
                .as_ref()
                .and_then(|paints| paints.first())
                .and_then(Option::as_ref),
            Some(ooxml_common::chart::ChartStyleFill::Image { image_path, .. })
                if image_path == "xl/media/theme-marker.png"
        ));
    }

    #[test]
    fn classic_graphicframe_loads_linked_chart_style_roles() {
        let mut archive = archive_with_classic_chart_style();
        let charts = load_sheet_charts(
            &mut archive,
            "worksheets/sheet1.xml",
            None,
            &theme(),
            (None, None),
            None,
        );
        let chart = &charts.first().expect("classic chart").chart;
        assert_eq!(chart.chart_type, "line");
        assert_eq!(
            chart
                .chart_style_roles
                .as_ref()
                .and_then(|roles| roles.get("dropLine"))
                .and_then(|style| style.line_colors.as_ref())
                .and_then(|colors| colors.first())
                .and_then(Option::as_deref),
            Some("336699"),
        );
    }
}
