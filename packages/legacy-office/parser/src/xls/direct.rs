//! Owned BIFF workbook-to-renderer-model boundary. Indexed models are projected
//! on demand, with one cursor-owned projection, without SpreadsheetML or ZIP parts.

use super::*;
use std::collections::BTreeMap;

const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;

pub(crate) struct DirectSession {
    pending_sheets: Option<Vec<(String, SheetData)>>,
    sheets: Vec<SheetSlot>,
    active_projection: Option<(usize, Box<ProjectedSheet>, usize)>,
    sheet_meta: Vec<(String, SheetVisibility)>,
    /// Resolved SheetExt tab colors, by sheet.
    tab_colors: Vec<Option<String>>,
    styles: Option<styles::ResolvedStyleSheet>,
    shared_strings: Vec<rich::Text>,
    date1904: bool,
    pictures: pictures::Pictures,
    native_pictures: pictures::NativePictures,
    charts: chart::Charts,
    native_charts: BTreeMap<usize, Vec<xlsx_model::ChartAnchor>>,
    shapes: shapes::Shapes,
    native_shapes: BTreeMap<usize, Vec<xlsx_model::ShapeAnchor>>,
    /// The next sheet of the sequential test cursor.
    #[cfg(test)]
    sheet_index: usize,
    measurement_font: Option<styles::NormalFont>,
    default_font: Option<(String, f64)>,
    mdw: Option<f64>,
    warnings: Vec<String>,
    model_budget: usize,
    bootstrapped: bool,
    poisoned: bool,
}

enum SheetSlot {
    Neutral {
        name: String,
        sheet: Box<SheetData>,
    },
    #[cfg(test)]
    Consumed,
}

struct ProjectedSheet {
    worksheet: xlsx_model::Worksheet,
    rows: Vec<xlsx_model::Row>,
}

pub(crate) struct ProjectedSheetRef<'a> {
    pub(crate) worksheet: &'a xlsx_model::Worksheet,
    pub(crate) rows: &'a [xlsx_model::Row],
}

impl DirectSession {
    pub(crate) fn new(cfb: &CompoundFile<'_>) -> Result<Self, String> {
        Self::from_prepared(prepare(cfb)?)
    }

    fn from_prepared(mut prepared: PreparedXls) -> Result<Self, String> {
        let mut meta_bytes = prepared
            .sheets
            .iter()
            .try_fold(
                prepared
                    .sheets
                    .len()
                    .saturating_mul(std::mem::size_of::<(String, SheetVisibility)>()),
                |total, (name, _)| total.checked_add(name.len()),
            )
            .ok_or_else(model_error)?;
        if let Some((name, _)) = prepared.styles.default_font() {
            meta_bytes = meta_bytes.checked_add(name.len()).ok_or_else(model_error)?;
        }
        if meta_bytes > MAX_MODEL_BYTES {
            return Err(model_error());
        }
        let mut sheet_meta = Vec::new();
        sheet_meta
            .try_reserve_exact(prepared.sheets.len())
            .map_err(|_| model_error())?;
        sheet_meta.extend(
            prepared
                .sheets
                .iter()
                .map(|(name, sheet)| (name.clone(), sheet.visibility)),
        );
        let tab_colors = prepared
            .sheets
            .iter()
            .map(|(_, sheet)| sheet.tab_color.clone())
            .collect();
        let default_font = prepared
            .styles
            .default_font()
            .map(|(name, size)| (name.to_owned(), size));
        let mut session = Self {
            pending_sheets: Some(std::mem::take(&mut prepared.sheets)),
            sheets: Vec::new(),
            active_projection: None,
            sheet_meta,
            tab_colors,
            styles: Some(prepared.styles),
            shared_strings: prepared.shared_strings,
            date1904: prepared.date1904,
            pictures: prepared.pictures,
            native_pictures: pictures::NativePictures {
                sheets: BTreeMap::new(),
                resources: BTreeMap::new(),
            },
            charts: std::mem::take(&mut prepared.charts),
            native_charts: BTreeMap::new(),
            shapes: std::mem::take(&mut prepared.shapes),
            native_shapes: BTreeMap::new(),
            #[cfg(test)]
            sheet_index: 0,
            measurement_font: prepared.font,
            default_font,
            warnings: prepared.warnings,
            model_budget: MAX_MODEL_BYTES - meta_bytes,
            mdw: None,
            bootstrapped: false,
            poisoned: false,
        };
        if session.pictures.is_empty() && session.charts.is_empty() && session.shapes.is_empty() {
            session.initialize_sheet_slots()?;
        }
        Ok(session)
    }

    pub(crate) fn measurement_font(&self) -> Option<&styles::NormalFont> {
        self.measurement_font.as_ref()
    }

    pub(crate) fn requires_measurement_decision(&self) -> bool {
        self.pending_sheets.is_some()
    }

    /// The host's single layout decision before bootstrap. While drawings
    /// await the Normal font's maximum digit width, `Some` resolves their
    /// anchors and `None` omits them with a warning. With no decision pending
    /// `None` changes nothing and a measured width fails closed.
    pub(crate) fn configure_host_layout(&mut self, mdw: Option<f64>) -> Result<(), String> {
        self.healthy()?;
        let Some(pending) = self.pending_sheets.as_ref() else {
            if mdw.is_none() {
                return Ok(());
            }
            return self.fail_protocol("XLS direct host layout decision is not pending");
        };
        if mdw.is_some_and(|v| !v.is_finite() || v.fract() != 0.0 || !(1.0..=4096.0).contains(&v)) {
            return self.fail_protocol("invalid measured XLS maximum digit width");
        }
        if let Some(mdw) = mdw {
            self.native_pictures = std::mem::take(&mut self.pictures)
                .resolve(pending, mdw, &mut self.warnings)
                .into_models(&mut self.model_budget)
                .inspect_err(|_| self.poisoned = true)?;
            self.native_charts =
                std::mem::take(&mut self.charts).resolve(pending, mdw, &mut self.warnings);
            self.native_shapes =
                std::mem::take(&mut self.shapes).resolve(pending, mdw, &mut self.warnings);
            // Canonical drawing metadata is retained alongside neutral BIFF
            // facts even when no cursor is active. Account it separately from
            // the cursor-owned copies, including nested chart/shape storage.
            let admission = (|| {
                charge_copy(&mut self.model_budget, &self.native_charts)?;
                charge_copy(&mut self.model_budget, &self.native_shapes)
            })();
            if let Err(error) = admission {
                self.poisoned = true;
                return Err(error);
            }
            // The resolvers report geometry they cannot place (a sheet without
            // stored defaults, a formula-display window or an anchor past the
            // resolved grid) as an omission; drawn content is never dropped,
            // so such a workbook rejects the session instead.
            if let Some(omitted) = self.warnings.iter().find(|warning| {
                matches!(
                    warning.as_str(),
                    "legacy-xls:unresolved-picture-geometry-omitted"
                        | "legacy-xls:unresolved-chart-geometry-omitted"
                        | "legacy-xls:unresolved-shape-geometry-omitted"
                )
            }) {
                let message = if omitted.contains("chart") {
                    "unresolved XLS chart anchor geometry"
                } else if omitted.contains("shape") {
                    "unresolved XLS shape anchor geometry"
                } else {
                    "unresolved XLS picture anchor geometry"
                };
                return self.fail(message);
            }
        } else {
            // The caller declined to measure the Normal font's maximum digit
            // width, which every chart, picture and shape anchor needs.
            self.warnings
                .push("legacy-xls:unmeasured-drawings-omitted".into());
            self.pictures = pictures::Pictures::default();
            self.charts = chart::Charts::default();
            self.shapes = shapes::Shapes::default();
        }
        self.mdw = mdw;
        self.initialize_sheet_slots()
    }

    pub(crate) fn bootstrap(&mut self) -> Result<xlsx_model::ParsedWorkbook, String> {
        self.healthy()?;
        if self.pending_sheets.is_some() {
            return self.fail_protocol(
                "XLS direct pictures require an explicit font measurement decision",
            );
        }
        if self.bootstrapped {
            return self.fail_protocol("XLS direct bootstrap already consumed");
        }
        let result = self.build_bootstrap();
        if result.is_err() {
            self.poisoned = true;
        } else {
            self.bootstrapped = true;
        }
        result
    }

    fn build_bootstrap(&mut self) -> Result<xlsx_model::ParsedWorkbook, String> {
        charge(
            &mut self.model_budget,
            std::mem::size_of::<xlsx_model::ParsedWorkbook>(),
        )?;
        let mut sheets = Vec::new();
        reserve_model(&mut sheets, self.sheet_meta.len(), &mut self.model_budget)?;
        for (index, (name, visibility)) in self.sheet_meta.iter().enumerate() {
            charge(&mut self.model_budget, name.len())?;
            let digits = (index + 1).ilog10() as usize + 1;
            charge(&mut self.model_budget, "rId".len() + digits)?;
            let tab_color = self.tab_colors.get(index).cloned().flatten();
            charge(
                &mut self.model_budget,
                tab_color.as_ref().map_or(0, String::len),
            )?;
            sheets.push(xlsx_model::SheetMeta {
                name: name.clone(),
                sheet_id: u32::try_from(index + 1).map_err(|_| model_error())?,
                r_id: format!("rId{}", index + 1),
                tab_color,
                visibility: visibility.model(),
            });
        }
        let styles = self.styles.as_ref().ok_or_else(model_error)?;
        let mut shared_strings = Vec::new();
        reserve_model(
            &mut shared_strings,
            self.shared_strings.len(),
            &mut self.model_budget,
        )?;
        for value in &self.shared_strings {
            shared_strings.push(value.model_into_reserved_slot(styles, &mut self.model_budget)?);
        }
        let styles = self
            .styles
            .take()
            .ok_or_else(model_error)?
            .into_model_bounded(&mut self.model_budget)?;
        self.shared_strings.clear();
        Ok(xlsx_model::ParsedWorkbook {
            workbook: xlsx_model::Workbook {
                sheets,
                date1904: self.date1904,
                parse_error: None,
            },
            styles,
            shared_strings,
        })
    }

    #[cfg(test)]
    pub(crate) fn next_sheet(&mut self) -> Result<Option<xlsx_model::Worksheet>, String> {
        self.healthy()?;
        if !self.bootstrapped {
            return self.fail_protocol("XLS direct bootstrap must be consumed before worksheets");
        }
        if self.pending_sheets.is_some() {
            return self.fail_protocol(
                "XLS direct pictures require an explicit font measurement decision",
            );
        }
        if self.sheet_index >= self.sheets.len() {
            return Ok(None);
        }
        let index = self.sheet_index;
        let name = self.sheet_meta[index].0.clone();
        self.projected_sheet(index, &name)?;
        self.sheet_index += 1;
        let (_, projected, charged) = self.active_projection.take().expect("projected above");
        self.model_budget += charged;
        self.sheets[index] = SheetSlot::Consumed;
        let ProjectedSheet {
            mut worksheet,
            rows,
        } = *projected;
        worksheet.rows = rows;
        Ok(Some(worksheet))
    }

    /// Native projection belongs only to the active cursor. Browser/Node retain
    /// models through the ordinary XLSX admission and lease/LRU paths; keeping
    /// another inactive decoded cache here would defeat their eviction. Neutral
    /// BIFF facts remain available for page-local reprojection after cancellation
    /// or eviction. Refund only the charge saved at successful projection.
    pub(crate) fn release_projection(&mut self) {
        if let Some((_, _, charged)) = self.active_projection.take() {
            self.model_budget += charged;
        }
    }

    pub(crate) fn projected_sheet(
        &mut self,
        index: usize,
        name: &str,
    ) -> Result<ProjectedSheetRef<'_>, String> {
        self.healthy()?;
        if !self.bootstrapped {
            return self.fail_protocol("XLS direct bootstrap must be consumed before worksheets");
        }
        let Some((expected, _)) = self.sheet_meta.get(index) else {
            return Err("XLS direct sheet index is out of range".to_string());
        };
        if expected != name {
            return Err("XLS direct sheet name does not match its index".to_string());
        }
        #[cfg(test)]
        if matches!(self.sheets.get(index), Some(SheetSlot::Consumed)) {
            return Err("XLS direct sheet was already consumed".to_string());
        }
        if self
            .active_projection
            .as_ref()
            .is_none_or(|(active, _, _)| *active != index)
        {
            self.release_projection();
            let (name, sheet) = match &self.sheets[index] {
                SheetSlot::Neutral { name, sheet } => (name, sheet),
                #[cfg(test)]
                SheetSlot::Consumed => unreachable!(),
            };
            let before = self.model_budget;
            let result = project_sheet(
                name.clone(),
                sheet,
                self.date1904,
                self.mdw,
                self.default_font.as_ref(),
                &mut self.model_budget,
            );
            let mut worksheet = match result {
                Ok(value) => value,
                Err(error) => {
                    self.poisoned = true;
                    return Err(error);
                }
            };
            // Canonical drawing metadata remains neutral. Charge each cursor's
            // owned copies before cloning; no projection is retained after ACK.
            let pictures = self.native_pictures.sheets.get(&index);
            let charts = self.native_charts.get(&index);
            let shapes = self.native_shapes.get(&index);
            let copies = (|| {
                if let Some(values) = pictures {
                    charge_copy(&mut self.model_budget, values)?;
                    worksheet.images = values.clone();
                }
                if let Some(values) = charts {
                    charge_copy(&mut self.model_budget, values)?;
                    worksheet.charts.extend(values.clone());
                }
                if let Some(values) = shapes {
                    charge_copy(&mut self.model_budget, values)?;
                    worksheet.shape_groups = values.clone();
                }
                Ok::<_, String>(())
            })();
            if let Err(error) = copies {
                self.poisoned = true;
                return Err(error);
            }
            let rows = std::mem::take(&mut worksheet.rows);
            self.active_projection = Some((
                index,
                Box::new(ProjectedSheet { worksheet, rows }),
                before - self.model_budget,
            ));
        }
        let (_, projected, _) = self.active_projection.as_ref().expect("projected above");
        Ok(ProjectedSheetRef {
            worksheet: &projected.worksheet,
            rows: &projected.rows,
        })
    }

    #[cfg(test)]
    pub(crate) fn warnings(&self) -> &[String] {
        &self.warnings
    }

    pub(crate) fn resource(&self, key: &str) -> Result<&[u8], String> {
        self.healthy()?;
        let Some(id) = key.strip_prefix("legacy-xls/image/") else {
            return Err("invalid XLS direct image key".to_string());
        };
        if id.is_empty()
            || !id.bytes().all(|byte| byte.is_ascii_digit())
            || (id.len() > 1 && id.starts_with('0'))
            || id.parse::<u32>().is_err()
        {
            return Err("invalid XLS direct image key".to_string());
        }
        self.native_pictures
            .resources
            .get(key)
            .map(Vec::as_slice)
            .ok_or_else(|| "unadmitted XLS direct image key".to_string())
    }

    pub(crate) fn assert_healthy(&self) -> Result<(), String> {
        self.healthy()
    }

    fn healthy(&self) -> Result<(), String> {
        if self.poisoned {
            Err("XLS direct session is poisoned".to_string())
        } else {
            Ok(())
        }
    }

    // Caller-owned sequencing/configuration and closed/poisoned state are not
    // input representability. In particular, a prior quota failure must never
    // turn into unsupported-content merely because the session is poisoned.
    fn fail_protocol<T>(&mut self, message: &str) -> Result<T, String> {
        self.poisoned = true;
        Err(message.to_owned())
    }

    fn fail<T>(&mut self, message: &str) -> Result<T, String> {
        self.poisoned = true;
        Err(unsupported(message))
    }

    fn initialize_sheet_slots(&mut self) -> Result<(), String> {
        let pending = self
            .pending_sheets
            .take()
            .expect("sheet slots initialized once");
        let mut sheets = Vec::new();
        if let Err(error) = reserve_model(&mut sheets, pending.len(), &mut self.model_budget) {
            self.poisoned = true;
            return Err(error);
        }
        sheets.extend(pending.into_iter().map(|(name, sheet)| SheetSlot::Neutral {
            name,
            sheet: Box::new(sheet),
        }));
        self.sheets = sheets;
        Ok(())
    }
}

/// Charge structural retained storage before cloning; serialized JSON is not
/// a heap bound (numeric arrays and omitted fields can be much larger in Rust).
fn charge_copy<T: ooxml_common::chart::RetainedBytes>(
    budget: &mut usize,
    value: &T,
) -> Result<(), String> {
    charge(
        budget,
        usize::try_from(value.heap_bytes()).map_err(|_| model_error())?,
    )
}

fn project_sheet(
    name: String,
    sheet: &SheetData,
    date1904: bool,
    mdw: Option<f64>,
    default_font: Option<&(String, f64)>,
    budget: &mut usize,
) -> Result<xlsx_model::Worksheet, String> {
    charge(
        budget,
        std::mem::size_of::<xlsx_model::Worksheet>() + name.len(),
    )?;
    if let Some((name, _)) = default_font {
        charge(budget, name.len())?;
    }
    let mut worksheet = empty_worksheet(name, date1904, default_font.cloned());
    if let Some(chart) = &sheet.chart_sheet {
        // A chart sheet (ECMA-376 §18.3.1.12 CT_Chartsheet in the XLSX model)
        // has no grid; its chart is placed as the model's absolute anchor
        // (same-cell corners carrying EMU offsets) at the Chart record's
        // chart-area rectangle.
        worksheet.is_chart_sheet = true;
        worksheet.show_gridlines = false;
        charge_copy(budget, &chart.model)?;
        reserve_model(&mut worksheet.charts, 1, budget)?;
        worksheet.charts.push(xlsx_model::ChartAnchor {
            z_order: 0,
            from_col: 0,
            from_col_off: chart.x_emu,
            from_row: 0,
            from_row_off: chart.y_emu,
            to_col: 0,
            to_col_off: chart.x_emu + chart.width_emu,
            to_row: 0,
            to_row_off: chart.y_emu + chart.height_emu,
            chart: chart.model.clone(),
        });
        sheet.views.project(&mut worksheet);
        return Ok(worksheet);
    }
    reserve_model(&mut worksheet.rows, sheet.rows.len(), budget)?;
    for (&row_index, cells) in &sheet.rows {
        let mut row = xlsx_model::Row {
            index: u32::from(row_index) + 1,
            height: None,
            custom_height: false,
            cells: Vec::new(),
            outline_level: 0,
            collapsed: false,
            hidden: false,
        };
        reserve_model(&mut row.cells, cells.len(), budget)?;
        for (&column, value) in cells {
            row.cells.push(xlsx_model::Cell {
                col: u32::from(column) + 1,
                row: u32::from(row_index) + 1,
                value: cell_value(value, budget)?,
                style_index: sheet
                    .cell_styles
                    .get(&(row_index, column))
                    .map(|v| u32::from(*v)),
                formula: match sheet.formulas.get(&(row_index, column)) {
                    Some(text) => {
                        charge(budget, text.len())?;
                        Some(text.clone())
                    }
                    None => None,
                },
                // BIFF shared formulas are not expanded here (no anchor/ref
                // tracking in this path): followers keep cached values under
                // recalc, matching the pre-extension OOXML behavior.
                formula_type: None,
                formula_ref: None,
                formula_si: None,
                show_phonetic: false,
            });
        }
        worksheet.rows.push(row);
    }
    reserve_model(&mut worksheet.merge_cells, sheet.merged.len(), budget)?;
    for &(first_row, last_row, first_column, last_column) in &sheet.merged {
        worksheet.merge_cells.push(xlsx_model::MergeCell {
            top: u32::from(first_row) + 1,
            left: u32::from(first_column) + 1,
            bottom: u32::from(last_row) + 1,
            right: u32::from(last_column) + 1,
        });
    }
    charge_copy(budget, &sheet.conditional_formats)?;
    worksheet.conditional_formats = sheet.conditional_formats.clone();
    charge(
        budget,
        sheet
            .tables
            .iter()
            .map(|table| {
                std::mem::size_of::<xlsx_model::TableInfo>()
                    + table.style_name.len()
                    + table.accent_color.len()
            })
            .sum(),
    )?;
    worksheet.tables = sheet.tables.clone();
    charge(
        budget,
        sheet
            .hyperlinks
            .iter()
            .map(|link| {
                std::mem::size_of::<xlsx_model::Hyperlink>()
                    + [&link.url, &link.location, &link.display]
                        .iter()
                        .map(|text| text.as_ref().map_or(0, String::len))
                        .sum::<usize>()
            })
            .sum(),
    )?;
    worksheet.hyperlinks = sheet.hyperlinks.clone();
    worksheet.auto_filter = sheet.auto_filter.clone();
    charge(
        budget,
        sheet
            .data_validations
            .iter()
            .map(|dv| {
                std::mem::size_of::<xlsx_model::DataValidation>()
                    + dv.sqref.len()
                    + [
                        &dv.validation_type,
                        &dv.operator,
                        &dv.formula1,
                        &dv.formula2,
                        &dv.prompt_title,
                        &dv.prompt,
                        &dv.error_title,
                        &dv.error_message,
                    ]
                    .iter()
                    .map(|text| text.as_ref().map_or(0, String::len))
                    .sum::<usize>()
            })
            .sum(),
    )?;
    worksheet.data_validations = sheet.data_validations.clone();
    charge(
        budget,
        sheet
            .defined_names
            .iter()
            .map(|name| {
                std::mem::size_of::<xlsx_model::DefinedName>()
                    + name.name.len()
                    + name.formula.len()
            })
            .sum(),
    )?;
    worksheet.defined_names = sheet.defined_names.clone();
    charge_copy(budget, &sheet.pivot_tables)?;
    worksheet.pivot_tables = sheet.pivot_tables.clone();
    if let Some(color) = &sheet.tab_color {
        charge(budget, color.len())?;
        worksheet.tab_color = Some(color.clone());
    }
    sheet.geometry.project(&mut worksheet, mdw, budget)?;
    sheet.views.project(&mut worksheet);
    Ok(worksheet)
}

fn cell_value(value: &CellValue, budget: &mut usize) -> Result<xlsx_model::CellValue, String> {
    Ok(match value {
        CellValue::Blank => xlsx_model::CellValue::Empty,
        CellValue::Number(number) if number.is_finite() => {
            xlsx_model::CellValue::Number { number: *number }
        }
        CellValue::Number(_) => {
            charge(budget, "#NUM!".len())?;
            xlsx_model::CellValue::Error {
                error: "#NUM!".into(),
            }
        }
        CellValue::Text(text) => {
            charge(budget, text.len())?;
            xlsx_model::CellValue::Text {
                text: text.clone(),
                runs: None,
                phonetic_runs: Vec::new(),
                phonetic_pr: None,
            }
        }
        CellValue::SharedString(si) => xlsx_model::CellValue::Shared { si: *si },
        CellValue::Bool(bool) => xlsx_model::CellValue::Bool { bool: *bool },
        CellValue::Error(error) => {
            charge(budget, error.len())?;
            xlsx_model::CellValue::Error {
                error: error.clone(),
            }
        }
    })
}

fn empty_worksheet(
    name: String,
    date1904: bool,
    default_font: Option<(String, f64)>,
) -> xlsx_model::Worksheet {
    let (default_font_family, default_font_size) =
        default_font.map_or((None, None), |(name, size)| (Some(name), Some(size)));
    xlsx_model::Worksheet {
        name,
        is_chart_sheet: false,
        is_dialog_sheet: false,
        rows: Vec::new(),
        col_widths: BTreeMap::new(),
        col_width_ranges: Vec::new(),
        col_style_ranges: Vec::new(),
        row_heights: BTreeMap::new(),
        col_outline_levels: BTreeMap::new(),
        col_collapsed: BTreeMap::new(),
        col_hidden: BTreeMap::new(),
        default_col_width: 0.0,
        base_col_width: None,
        default_col_width_is_ui: false,
        default_row_height: 0.0,
        default_row_height_custom: false,
        merge_cells: Vec::new(),
        freeze_rows: 0,
        freeze_cols: 0,
        conditional_formats: Vec::new(),
        images: Vec::new(),
        charts: Vec::new(),
        shape_groups: Vec::new(),
        show_zeros: true,
        show_gridlines: true,
        right_to_left: false,
        outline_pr: None,
        tab_color: None,
        auto_filter: None,
        hyperlinks: Vec::new(),
        comment_refs: Vec::new(),
        comments: Vec::new(),
        data_validations: Vec::new(),
        defined_names: Vec::new(),
        tables: Vec::new(),
        slicers: Vec::new(),
        pivot_tables: Vec::new(),
        pivot_diagnostics: Vec::new(),
        sparkline_groups: Vec::new(),
        default_font_family,
        default_font_size,
        default_font_bold: None,
        default_font_italic: None,
        // BIFF Normal XF is the source of default_font (MS-XLS 2.2.6.1.2.2).
        normal_font_size: default_font_size,
        theme_japanese_major_font: None,
        theme_japanese_minor_font: None,
        date1904,
        parse_error: None,
    }
}

impl SheetVisibility {
    fn model(self) -> xlsx_model::SheetVisibility {
        match self {
            Self::Visible => xlsx_model::SheetVisibility::Visible,
            Self::Hidden => xlsx_model::SheetVisibility::Hidden,
            Self::VeryHidden => xlsx_model::SheetVisibility::VeryHidden,
        }
    }
}

fn reserve_model<T>(values: &mut Vec<T>, count: usize, budget: &mut usize) -> Result<(), String> {
    charge(
        budget,
        count
            .checked_mul(std::mem::size_of::<T>())
            .ok_or_else(model_error)?,
    )?;
    values.try_reserve_exact(count).map_err(|_| model_error())
}

fn charge(budget: &mut usize, bytes: usize) -> Result<(), String> {
    *budget = budget.checked_sub(bytes).ok_or_else(model_error)?;
    Ok(())
}

fn model_error() -> String {
    crate::rejection::resource("XLS direct model byte budget exceeded")
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;
    use crate::cfb::test_support::{build_cfb, build_scoped_cfb};

    fn record(kind: u16, data: &[u8]) -> Vec<u8> {
        [
            kind.to_le_bytes().as_slice(),
            &(data.len() as u16).to_le_bytes(),
            data,
        ]
        .concat()
    }

    fn workbook() -> Vec<u8> {
        let mut stream = record(BOF, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(record(BOUNDSHEET8, &[0, 0, 0, 0, 2, 0, 1, 0, b'S']));
        stream.extend(record(0x0022, &[1, 0]));
        stream.extend(record(EOF, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(record(BOF, &[0, 6, 0x10, 0]));
        let mut number = vec![0, 0, 2, 0, 0, 0];
        number.extend(42.5f64.to_le_bytes());
        stream.extend(record(NUMBER, &number));
        stream.extend(record(MERGEDCELLS, &[1, 0, 0, 0, 0, 0, 2, 0, 2, 0]));
        stream.extend(record(EOF, &[]));
        scoped_cfb(stream)
    }

    /// A one-stream compound file whose directory tree the scoped stream
    /// reader accepts.
    fn scoped_cfb(stream: Vec<u8>) -> Vec<u8> {
        let mut bytes = build_cfb(&[("Workbook", stream)]);
        let directory_sector = u32::from_le_bytes(bytes[48..52].try_into().unwrap()) as usize;
        let directory = 512 + directory_sector * 512;
        bytes[directory + 68..directory + 76].fill(0xff);
        bytes[directory + 76..directory + 80].copy_from_slice(&1u32.to_le_bytes());
        let workbook = directory + 128;
        bytes[workbook + 68..workbook + 80].fill(0xff);
        bytes
    }

    fn picture_session() -> (DirectSession, &'static str, Vec<u8>) {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut prepared = prepare(&cfb).unwrap();
        let (pictures, sheet, key, resource) = pictures::session_fixture();
        prepared.sheets = vec![("S".into(), sheet)];
        prepared.pictures = pictures;
        prepared.font = Some(styles::NormalFont {
            name: "Calibri".into(),
            size_points: 11.0,
            bold: false,
            italic: false,
        });
        (
            DirectSession::from_prepared(prepared).unwrap(),
            key,
            resource,
        )
    }

    pub(crate) fn wire_fixture() -> DirectSession {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        DirectSession::new(&cfb).unwrap()
    }

    pub(crate) fn wire_picture_fixture() -> DirectSession {
        picture_session().0
    }

    fn indexed_session() -> DirectSession {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut prepared = prepare(&cfb).unwrap();
        prepared.sheets = ["First", "Second", "Third"]
            .into_iter()
            .map(|name| (name.into(), SheetData::default()))
            .collect();
        DirectSession::from_prepared(prepared).unwrap()
    }

    #[test]
    fn cursor_projection_releases_budget_for_reopen() {
        let mut session = indexed_session();
        session.bootstrap().unwrap();
        // Capacity for one empty model, but not two: cache navigation must
        // release native projection ownership before another sheet is admitted.
        session.model_budget = std::mem::size_of::<xlsx_model::Worksheet>() + 100;
        for index in (0..3).cycle().take(64) {
            let name = ["First", "Second", "Third"][index];
            assert_eq!(
                session.projected_sheet(index, name).unwrap().worksheet.name,
                name
            );
            session.release_projection();
        }
        assert!(session.assert_healthy().is_ok());
    }

    #[test]
    fn rejects_macro_and_malformed_chart_sheets_and_makes_no_blanket_omission_claim() {
        // Drawings and conditional formatting are projected or the workbook
        // is rejected, so no blanket omission is ever reported.
        assert!(!wire_fixture()
            .warnings()
            .iter()
            .any(|warning| warning.contains("drawings-conditional-formatting")));
        // Chart sheets are projected; one without its Chart record fails closed.
        for (dt, expected) in [(2u8, "lacks its Chart record"), (1, "macro sheets")] {
            let mut stream = record(BOF, &[0, 6, 5, 0]);
            let first = stream.len() + 4;
            stream.extend(record(BOUNDSHEET8, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
            let second = stream.len() + 4;
            stream.extend(record(BOUNDSHEET8, &[0, 0, 0, 0, 0, dt, 1, 0, b'C']));
            stream.extend(record(EOF, &[]));
            let sheet = stream.len() as u32;
            stream[first..first + 4].copy_from_slice(&sheet.to_le_bytes());
            stream.extend(record(BOF, &[0, 6, 0x10, 0]));
            stream.extend(record(EOF, &[]));
            let other = stream.len() as u32;
            stream[second..second + 4].copy_from_slice(&other.to_le_bytes());
            stream.extend(record(BOF, &[0, 6, 0x20, 0]));
            stream.extend(record(EOF, &[]));
            let bytes = scoped_cfb(stream);
            let cfb = CompoundFile::open(&bytes).unwrap();
            let error = DirectSession::new(&cfb).err().expect("rejected");
            assert!(error.contains(expected), "{expected}: {error}");
        }
    }

    #[test]
    fn owns_bootstrap_and_moves_each_sheet_once_without_xml() {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        let bootstrap = session.bootstrap().unwrap();
        assert!(bootstrap.workbook.date1904);
        assert_eq!(bootstrap.workbook.sheets.len(), 1);
        assert_eq!(
            bootstrap.workbook.sheets[0].visibility,
            xlsx_model::SheetVisibility::VeryHidden
        );
        assert_eq!(bootstrap.shared_strings.len(), 0);
        let sheet = session.next_sheet().unwrap().unwrap();
        assert_eq!(sheet.name, "S");
        assert!(sheet.date1904);
        assert_eq!(sheet.default_font_family.as_deref(), Some("Calibri"));
        assert_eq!(sheet.default_font_size, Some(11.0));
        assert_eq!(sheet.rows[0].index, 1);
        assert_eq!(sheet.rows[0].cells[0].col, 3);
        assert_eq!(sheet.rows[0].cells[0].style_index, Some(0));
        assert!(matches!(
            sheet.rows[0].cells[0].value,
            xlsx_model::CellValue::Number { number: 42.5 }
        ));
        assert_eq!(
            (sheet.merge_cells[0].top, sheet.merge_cells[0].right),
            (1, 3)
        );
        assert!(session.next_sheet().unwrap().is_none());
    }

    #[test]
    fn indexed_projection_is_random_access_reusable_and_charged_once() {
        let mut session = indexed_session();
        session.bootstrap().unwrap();

        let before = session.model_budget;
        let third = session.projected_sheet(2, "Third").unwrap();
        assert_eq!(third.worksheet.name, "Third");
        assert!(third.worksheet.rows.is_empty());
        assert!(third.rows.is_empty());
        let after_first_projection = session.model_budget;
        assert!(after_first_projection < before);

        let third_again = session.projected_sheet(2, "Third").unwrap();
        assert_eq!(third_again.worksheet.name, "Third");
        assert_eq!(session.model_budget, after_first_projection);
        assert_eq!(
            session.projected_sheet(0, "First").unwrap().worksheet.name,
            "First"
        );
        assert!(session.projected_sheet(3, "missing").is_err());
        assert!(session.projected_sheet(1, "wrong-name").is_err());
        assert_eq!(
            session.projected_sheet(1, "Second").unwrap().worksheet.name,
            "Second"
        );
    }

    #[test]
    fn sequential_compatibility_consumes_the_indexed_slot_without_cloning() {
        let mut session = indexed_session();
        session.bootstrap().unwrap();
        session.projected_sheet(0, "First").unwrap();
        assert_eq!(session.next_sheet().unwrap().unwrap().name, "First");
        assert!(session.projected_sheet(0, "First").is_err());
        assert_eq!(session.next_sheet().unwrap().unwrap().name, "Second");
    }

    #[test]
    fn state_errors_poison_the_owned_session() {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        assert!(session.next_sheet().is_err());
        assert!(session.bootstrap().is_err());
    }

    #[test]
    fn model_budget_failure_poisoning_and_resource_keys_fail_closed() {
        let bytes = workbook();
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        for key in [
            "",
            "legacy-xls/image/",
            "legacy-xls/image/not-a-number",
            "other/1",
        ] {
            assert!(session.resource(key).is_err());
        }
        session.model_budget = 0;
        assert!(session.bootstrap().is_err());
        assert!(session.next_sheet().is_err());
    }

    #[test]
    fn cell_projection_keeps_cached_values_without_inventing_formulas() {
        for (source, expected) in [
            (CellValue::Blank, xlsx_model::CellValue::Empty),
            (
                CellValue::SharedString(7),
                xlsx_model::CellValue::Shared { si: 7 },
            ),
            (
                CellValue::Bool(true),
                xlsx_model::CellValue::Bool { bool: true },
            ),
        ] {
            let mut budget = 1024;
            assert_eq!(
                serde_json::to_value(cell_value(&source, &mut budget).unwrap()).unwrap(),
                serde_json::to_value(expected).unwrap()
            );
        }
        let mut budget = 1024;
        assert!(
            matches!(cell_value(&CellValue::Number(f64::INFINITY), &mut budget).unwrap(), xlsx_model::CellValue::Error { error } if error == "#NUM!")
        );
        let mut budget = 1024;
        assert!(
            matches!(cell_value(&CellValue::Text("cached".into()), &mut budget).unwrap(), xlsx_model::CellValue::Text { text, .. } if text == "cached")
        );
    }

    #[test]
    fn measured_picture_geometry_and_resources_share_one_session_decision() {
        let (mut session, key, expected) = picture_session();
        assert!(session.requires_measurement_decision());
        assert_eq!(session.measurement_font().unwrap().name, "Calibri");
        session.configure_host_layout(Some(7.0)).unwrap();
        assert!(!session.requires_measurement_decision());
        session.bootstrap().unwrap();
        let sheet = session.next_sheet().unwrap().unwrap();
        assert_eq!(sheet.images.len(), 1);
        assert_eq!(sheet.images[0].image_path, key);
        assert_eq!(session.resource(key).unwrap(), expected);
        assert!(session.resource("legacy-xls/image/07").is_err());

        let (mut wider, _, _) = picture_session();
        wider.configure_host_layout(Some(9.0)).unwrap();
        wider.bootstrap().unwrap();
        let wider = wider.next_sheet().unwrap().unwrap();
        assert_ne!(sheet.default_col_width, wider.default_col_width);
        assert_ne!(sheet.images[0].native_ext_cx, wider.images[0].native_ext_cx);
    }

    #[test]
    fn picture_measurement_errors_are_terminal_and_decisions_are_once_only() {
        let (mut session, _, _) = picture_session();
        assert!(session.bootstrap().is_err());
        assert!(session.configure_host_layout(Some(7.0)).is_err());

        let (mut session, _, _) = picture_session();
        assert!(session.configure_host_layout(Some(f64::NAN)).is_err());
        assert!(session.bootstrap().is_err());

        let (mut session, _, _) = picture_session();
        session.configure_host_layout(None).unwrap();
        assert!(session
            .warnings()
            .iter()
            .any(|warning| warning == "legacy-xls:unmeasured-drawings-omitted"));
        assert!(session.configure_host_layout(Some(7.0)).is_err());
    }

    #[test]
    fn projects_biff8_scalar_and_unicode_cells() {
        fn bof(kind: u16) -> Vec<u8> {
            record(BOF, &[0, 6, kind as u8, (kind >> 8) as u8, 0, 0, 0, 0])
        }
        let mut stream = bof(WORKBOOK_GLOBALS);
        let bound = stream.len() + 4;
        let mut sheet = vec![0; 6];
        sheet.extend([3, 1]);
        sheet.extend("表計算".encode_utf16().flat_map(u16::to_le_bytes));
        stream.extend(record(BOUNDSHEET8, &sheet));
        let mut sst = [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        sst.extend(3u16.to_le_bytes());
        sst.push(1);
        sst.extend("日本語".encode_utf16().flat_map(u16::to_le_bytes));
        stream.extend(record(SST, &sst));
        stream.extend(record(EOF, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(bof(WORKSHEET));
        let mut number = vec![0; 6];
        number.extend(42.5f64.to_le_bytes());
        stream.extend(record(NUMBER, &number));
        stream.extend(record(LABELSST, &[1, 0, 1, 0, 0, 0, 0, 0, 0, 0]));
        stream.extend(record(EOF, &[]));
        let bytes = build_scoped_cfb(&[("Workbook", stream)]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        let bootstrap = session.bootstrap().unwrap();
        assert_eq!(bootstrap.workbook.sheets.len(), 1);
        assert_eq!(bootstrap.workbook.sheets[0].name, "表計算");
        assert_eq!(bootstrap.shared_strings.len(), 1);
        assert_eq!(bootstrap.shared_strings[0].text, "日本語");
        let sheet = session.next_sheet().unwrap().unwrap();
        assert_eq!(sheet.name, "表計算");
        let a1 = &sheet.rows[0].cells[0];
        assert_eq!((a1.row, a1.col), (1, 1));
        assert!(matches!(
            a1.value,
            xlsx_model::CellValue::Number { number: 42.5 }
        ));
        let b2 = &sheet.rows[1].cells[0];
        assert_eq!((b2.row, b2.col), (2, 2));
        assert!(matches!(b2.value, xlsx_model::CellValue::Shared { si: 0 }));
    }

    #[test]
    fn rejects_encrypted_workbooks() {
        let mut stream = record(BOF, &[0, 6, 5, 0, 0, 0, 0, 0]);
        stream.extend(record(FILEPASS, &[]));
        let bytes = build_scoped_cfb(&[("Workbook", stream)]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let error = DirectSession::new(&cfb).err().expect("encrypted workbook");
        assert!(error.contains("encrypted"), "{error}");
    }

    #[test]
    fn projects_cell_xf_blank_cells_and_sheet_geometry_and_admits_print_records() {
        let mut stream = record(BOF, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(record(BOUNDSHEET8, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
        stream.extend(record(0x0022, &[1, 0])); // Date1904
        let mut font = vec![0; 16];
        font[0..2].copy_from_slice(&360u16.to_le_bytes());
        font[2] = 2; // italic
        font[4..6].copy_from_slice(&10u16.to_le_bytes());
        font[6..8].copy_from_slice(&700u16.to_le_bytes());
        font[14..16].copy_from_slice(&[5, 0]);
        font.extend(b"Arial");
        stream.extend(record(0x0031, &font));
        let mut xf = [0u8; 20];
        xf[6] = 0x2a; // center, wrap, bottom
        xf[10..14].copy_from_slice(&(1u32 | (10 << 16)).to_le_bytes()); // thin red left
        xf[14..18].copy_from_slice(&(1u32 << 26).to_le_bytes()); // solid fill
        xf[18..20].copy_from_slice(&(13u16 | (65 << 7)).to_le_bytes());
        stream.extend(record(0x00e0, &xf));
        xf[2..4].copy_from_slice(&14u16.to_le_bytes()); // date format
        stream.extend(record(0x00e0, &xf));
        stream.extend(record(EOF, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(record(BOF, &[0, 6, 0x10, 0]));
        // Print records carry no model data but are validated for admission;
        // with all four margins present no incomplete-margin warning arises.
        stream.extend(record(0x0081, &[0, 1])); // fit to pages
        for kind in 0x0026..=0x0029 {
            stream.extend(record(kind, &0.5f64.to_le_bytes()));
        }
        let mut setup = vec![0u8; 34];
        setup[..16].copy_from_slice(&[9, 0, 75, 0, 3, 0, 2, 0, 0, 0, 0x89, 0, 88, 2, 88, 2]);
        setup[16..24].copy_from_slice(&0.25f64.to_le_bytes());
        setup[24..32].copy_from_slice(&0.3f64.to_le_bytes());
        setup[32] = 1;
        stream.extend(record(0x00a1, &setup));
        stream.extend(record(0x0014, &[4, 0, 0, b'&', b'L', b'&', b'P']));
        stream.extend(record(0x001b, &[1, 0, 4, 0, 0, 0, 255, 63]));
        stream.extend(record(0x0225, &[0, 0, 0x2c, 1])); // default 15 pt
                                                         // Columns A:C, width 20, XF 1, hidden and custom width.
        stream.extend(record(0x007d, &[0, 0, 2, 0, 0, 20, 1, 0, 3, 0, 0, 0]));
        let mut row = [0u8; 16];
        row[0] = 2; // empty third row, 30 pt, hidden with a custom height
        row[6..8].copy_from_slice(&600u16.to_le_bytes());
        row[12] = 0x60;
        stream.extend(record(0x0208, &row));
        let mut number = vec![0, 0, 0, 0, 1, 0];
        number.extend(1f64.to_le_bytes());
        stream.extend(record(NUMBER, &number));
        stream.extend(record(0x0201, &[0, 0, 1, 0, 1, 0])); // styled empty cell
        stream.extend(record(0x00be, &[1, 0, 0, 0, 1, 0, 1, 0, 1, 0])); // two blanks
        stream.extend(record(EOF, &[]));
        let bytes = build_scoped_cfb(&[("Workbook", stream)]);
        let cfb = CompoundFile::open(&bytes).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        assert!(session.warnings().is_empty(), "{:?}", session.warnings());
        let bootstrap = session.bootstrap().unwrap();
        assert!(bootstrap.workbook.date1904);
        let styles = bootstrap.styles;
        assert_eq!(styles.fonts.len(), 1);
        let font = &styles.fonts[0];
        assert_eq!((font.name.as_deref(), font.size), (Some("Arial"), 18.0));
        assert!(font.bold && font.italic);
        assert_eq!(font.color.as_deref(), Some("#FF0000"));
        assert_eq!(styles.cell_xfs.len(), 2);
        let (normal, date) = (&styles.cell_xfs[0], &styles.cell_xfs[1]);
        assert_eq!((normal.num_fmt_id, date.num_fmt_id), (0, 14));
        assert_eq!(date.align_h.as_deref(), Some("center"));
        assert_eq!(date.align_v.as_deref(), Some("bottom"));
        assert!(date.wrap_text);
        let fill = &styles.fills[date.fill_id as usize];
        assert_eq!(fill.pattern_type, "solid");
        assert_eq!(fill.fg_color.as_deref(), Some("#FFFF00"));
        let border = &styles.borders[date.border_id as usize];
        let left = border.left.as_ref().unwrap();
        assert_eq!(
            (left.style.as_str(), left.color.as_deref()),
            ("thin", Some("#FF0000"))
        );
        assert!(border.right.is_none() && border.top.is_none() && border.bottom.is_none());

        let sheet = session.next_sheet().unwrap().unwrap();
        assert!(sheet.date1904);
        assert_eq!(sheet.default_row_height, 15.0);
        let cells: Vec<_> = sheet
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .map(|cell| (cell.row, cell.col, cell.style_index))
            .collect();
        assert_eq!(
            cells,
            [
                (1, 1, Some(1)),
                (1, 2, Some(1)),
                (2, 1, Some(1)),
                (2, 2, Some(1))
            ]
        );
        assert!(matches!(
            sheet.rows[0].cells[0].value,
            xlsx_model::CellValue::Number { number } if number == 1.0
        ));
        assert!(matches!(
            sheet.rows[0].cells[1].value,
            xlsx_model::CellValue::Empty
        ));
        let third = sheet.rows.iter().find(|row| row.index == 3).unwrap();
        assert!(third.cells.is_empty() && third.hidden && third.custom_height);
        assert_eq!(third.height, Some(0.0));
        for column in 1..=3 {
            assert_eq!(sheet.col_widths.get(&column), Some(&0.0));
            assert_eq!(sheet.col_hidden.get(&column), Some(&true));
        }
        assert!(sheet
            .col_style_ranges
            .iter()
            .all(|range| range.style_index == 1));
        assert_eq!(sheet.col_style_ranges.len(), 3);
    }
}
