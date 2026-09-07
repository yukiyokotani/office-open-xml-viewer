//! Owned BIFF workbook-to-renderer-model boundary. Indexed models are projected
//! once on demand, without creating SpreadsheetML or ZIP parts.

use super::*;
use std::collections::BTreeMap;

const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;

pub(crate) struct DirectSession {
    pending_sheets: Option<Vec<(String, SheetData)>>,
    sheets: Vec<SheetSlot>,
    sheet_meta: Vec<(String, SheetVisibility)>,
    styles: Option<styles::ResolvedStyleSheet>,
    shared_strings: Vec<rich::Text>,
    date1904: bool,
    pictures: pictures::Pictures,
    native_pictures: pictures::NativePictures,
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
    Neutral { name: String, sheet: SheetData },
    Projected(ProjectedSheet),
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
        Self::from_prepared(prepare_direct(cfb)?)
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
        let default_font = prepared
            .styles
            .default_font()
            .map(|(name, size)| (name.to_owned(), size));
        let mut session = Self {
            pending_sheets: Some(std::mem::take(&mut prepared.sheets)),
            sheets: Vec::new(),
            sheet_meta,
            styles: Some(prepared.styles),
            shared_strings: prepared.shared_strings,
            date1904: prepared.date1904,
            pictures: prepared.pictures,
            native_pictures: pictures::NativePictures {
                sheets: BTreeMap::new(),
                resources: BTreeMap::new(),
            },
            sheet_index: 0,
            measurement_font: prepared.font,
            default_font,
            warnings: prepared.warnings,
            model_budget: MAX_MODEL_BYTES - meta_bytes,
            mdw: None,
            bootstrapped: false,
            poisoned: false,
        };
        if session.pictures.is_empty() {
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

    pub(crate) fn configure_mdw(&mut self, mdw: Option<f64>) -> Result<(), String> {
        self.healthy()?;
        let Some(pending) = self.pending_sheets.as_ref() else {
            return self.fail("XLS direct measurement already configured");
        };
        if mdw.is_some_and(|v| !v.is_finite() || v.fract() != 0.0 || !(1.0..=4096.0).contains(&v)) {
            return self.fail("invalid measured XLS maximum digit width");
        }
        if let Some(mdw) = mdw {
            self.native_pictures = std::mem::take(&mut self.pictures)
                .resolve(pending, mdw, &mut self.warnings)
                .into_models(&mut self.model_budget)
                .map_err(|error| {
                    self.poisoned = true;
                    error
                })?;
        } else {
            self.warnings
                .push("legacy-xls:unmeasured-pictures-omitted".into());
            self.pictures = pictures::Pictures::default();
        }
        self.mdw = mdw;
        self.initialize_sheet_slots()
    }

    pub(crate) fn bootstrap(&mut self) -> Result<xlsx_model::ParsedWorkbook, String> {
        self.healthy()?;
        if self.pending_sheets.is_some() {
            return self.fail("XLS direct pictures require an explicit font measurement decision");
        }
        if self.bootstrapped {
            return self.fail("XLS direct bootstrap already consumed");
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
            sheets.push(xlsx_model::SheetMeta {
                name: name.clone(),
                sheet_id: u32::try_from(index + 1).map_err(|_| model_error())?,
                r_id: format!("rId{}", index + 1),
                tab_color: None,
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

    pub(crate) fn next_sheet(&mut self) -> Result<Option<xlsx_model::Worksheet>, String> {
        self.healthy()?;
        if !self.bootstrapped {
            return self.fail("XLS direct bootstrap must be consumed before worksheets");
        }
        if self.pending_sheets.is_some() {
            return self.fail("XLS direct pictures require an explicit font measurement decision");
        }
        if self.sheet_index >= self.sheets.len() {
            return Ok(None);
        }
        let index = self.sheet_index;
        let name = self.sheet_meta[index].0.clone();
        self.projected_sheet(index, &name)?;
        self.sheet_index += 1;
        let SheetSlot::Projected(projected) =
            std::mem::replace(&mut self.sheets[index], SheetSlot::Consumed)
        else {
            return self.fail("XLS direct sheet was already consumed");
        };
        let mut worksheet = projected.worksheet;
        worksheet.rows = projected.rows;
        Ok(Some(worksheet))
    }

    /// Lazily project one worksheet into an indexed, reusable model slot.
    /// The row-free shell and row sidecar are borrowed by the cursor wire layer;
    /// projection is charged once and never repeated after cancellation.
    pub(crate) fn projected_sheet(
        &mut self,
        index: usize,
        name: &str,
    ) -> Result<ProjectedSheetRef<'_>, String> {
        self.healthy()?;
        if !self.bootstrapped {
            return self.fail("XLS direct bootstrap must be consumed before worksheets");
        }
        if self.pending_sheets.is_some() {
            return self.fail("XLS direct pictures require an explicit font measurement decision");
        }
        let Some((expected, _)) = self.sheet_meta.get(index) else {
            return Err(unsupported("XLS direct sheet index is out of range"));
        };
        if expected != name {
            return Err(unsupported(
                "XLS direct sheet name does not match its index",
            ));
        }
        if matches!(self.sheets.get(index), Some(SheetSlot::Consumed)) {
            return Err(unsupported("XLS direct sheet was already consumed"));
        }
        if matches!(self.sheets.get(index), Some(SheetSlot::Neutral { .. })) {
            let SheetSlot::Neutral { name, sheet } =
                std::mem::replace(&mut self.sheets[index], SheetSlot::Consumed)
            else {
                unreachable!("neutral slot checked above")
            };
            let projected = project_sheet(
                name,
                sheet,
                self.date1904,
                self.mdw,
                self.default_font.as_ref(),
                &mut self.model_budget,
            );
            let mut worksheet = match projected {
                Ok(worksheet) => worksheet,
                Err(error) => {
                    self.poisoned = true;
                    return Err(error);
                }
            };
            worksheet.images = self
                .native_pictures
                .sheets
                .remove(&index)
                .unwrap_or_default();
            let rows = std::mem::take(&mut worksheet.rows);
            self.sheets[index] = SheetSlot::Projected(ProjectedSheet { worksheet, rows });
        }
        let SheetSlot::Projected(projected) = &self.sheets[index] else {
            unreachable!("consumed slot rejected above")
        };
        Ok(ProjectedSheetRef {
            worksheet: &projected.worksheet,
            rows: &projected.rows,
        })
    }

    pub(crate) fn warnings(&self) -> &[String] {
        &self.warnings
    }

    pub(crate) fn resource(&self, key: &str) -> Result<&[u8], String> {
        self.healthy()?;
        let Some(id) = key.strip_prefix("legacy-xls/image/") else {
            return Err(unsupported("invalid XLS direct image key"));
        };
        if id.is_empty()
            || !id.bytes().all(|byte| byte.is_ascii_digit())
            || (id.len() > 1 && id.starts_with('0'))
            || id.parse::<u32>().is_err()
        {
            return Err(unsupported("invalid XLS direct image key"));
        }
        self.native_pictures
            .resources
            .get(key)
            .map(Vec::as_slice)
            .ok_or_else(|| unsupported("unadmitted XLS direct image key"))
    }

    pub(crate) fn assert_healthy(&self) -> Result<(), String> {
        self.healthy()
    }

    fn healthy(&self) -> Result<(), String> {
        if self.poisoned {
            Err(unsupported("XLS direct session is poisoned"))
        } else {
            Ok(())
        }
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
        sheets.extend(
            pending
                .into_iter()
                .map(|(name, sheet)| SheetSlot::Neutral { name, sheet }),
        );
        self.sheets = sheets;
        Ok(())
    }
}

fn project_sheet(
    name: String,
    sheet: SheetData,
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
    reserve_model(&mut worksheet.rows, sheet.rows.len(), budget)?;
    for (row_index, cells) in sheet.rows {
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
        for (column, value) in cells {
            row.cells.push(xlsx_model::Cell {
                col: u32::from(column) + 1,
                row: u32::from(row_index) + 1,
                value: cell_value(value, budget)?,
                style_index: sheet
                    .cell_styles
                    .get(&(row_index, column))
                    .map(|v| u32::from(*v)),
                formula: None,
                show_phonetic: false,
            });
        }
        worksheet.rows.push(row);
    }
    reserve_model(&mut worksheet.merge_cells, sheet.merged.len(), budget)?;
    for (first_row, last_row, first_column, last_column) in sheet.merged {
        worksheet.merge_cells.push(xlsx_model::MergeCell {
            top: u32::from(first_row) + 1,
            left: u32::from(first_column) + 1,
            bottom: u32::from(last_row) + 1,
            right: u32::from(last_column) + 1,
        });
    }
    sheet.geometry.project(&mut worksheet, mdw, budget)?;
    sheet.views.project(&mut worksheet);
    Ok(worksheet)
}

fn cell_value(value: CellValue, budget: &mut usize) -> Result<xlsx_model::CellValue, String> {
    Ok(match value {
        CellValue::Blank => xlsx_model::CellValue::Empty,
        CellValue::Number(number) if number.is_finite() => xlsx_model::CellValue::Number { number },
        CellValue::Number(_) => {
            charge(budget, "#NUM!".len())?;
            xlsx_model::CellValue::Error {
                error: "#NUM!".into(),
            }
        }
        CellValue::Text(text) => {
            charge(budget, text.len())?;
            xlsx_model::CellValue::Text {
                text,
                runs: None,
                phonetic_runs: Vec::new(),
                phonetic_pr: None,
            }
        }
        CellValue::SharedString(si) => xlsx_model::CellValue::Shared { si },
        CellValue::Bool(bool) => xlsx_model::CellValue::Bool { bool },
        CellValue::Error(error) => {
            charge(budget, error.len())?;
            xlsx_model::CellValue::Error { error }
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
        rows: Vec::new(),
        col_widths: BTreeMap::new(),
        col_width_ranges: Vec::new(),
        col_style_ranges: Vec::new(),
        row_heights: BTreeMap::new(),
        col_outline_levels: BTreeMap::new(),
        col_collapsed: BTreeMap::new(),
        col_hidden: BTreeMap::new(),
        default_col_width: 0.0,
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
    unsupported("XLS direct model byte budget exceeded")
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;
    use crate::cfb::test_support::build_cfb;

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
        let mut prepared = prepare(&cfb, false).unwrap();
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
        let mut prepared = prepare(&cfb, false).unwrap();
        prepared.sheets = ["First", "Second", "Third"]
            .into_iter()
            .map(|name| (name.into(), SheetData::default()))
            .collect();
        DirectSession::from_prepared(prepared).unwrap()
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
                serde_json::to_value(cell_value(source, &mut budget).unwrap()).unwrap(),
                serde_json::to_value(expected).unwrap()
            );
        }
        let mut budget = 1024;
        assert!(
            matches!(cell_value(CellValue::Number(f64::INFINITY), &mut budget).unwrap(), xlsx_model::CellValue::Error { error } if error == "#NUM!")
        );
        let mut budget = 1024;
        assert!(
            matches!(cell_value(CellValue::Text("cached".into()), &mut budget).unwrap(), xlsx_model::CellValue::Text { text, .. } if text == "cached")
        );
    }

    #[test]
    fn measured_picture_geometry_and_resources_share_one_session_decision() {
        let (mut session, key, expected) = picture_session();
        assert!(session.requires_measurement_decision());
        assert_eq!(session.measurement_font().unwrap().name, "Calibri");
        session.configure_mdw(Some(7.0)).unwrap();
        assert!(!session.requires_measurement_decision());
        session.bootstrap().unwrap();
        let sheet = session.next_sheet().unwrap().unwrap();
        assert_eq!(sheet.images.len(), 1);
        assert_eq!(sheet.images[0].image_path, key);
        assert_eq!(session.resource(key).unwrap(), expected);
        assert!(session.resource("legacy-xls/image/07").is_err());

        let (mut wider, _, _) = picture_session();
        wider.configure_mdw(Some(9.0)).unwrap();
        wider.bootstrap().unwrap();
        let wider = wider.next_sheet().unwrap().unwrap();
        assert_ne!(sheet.default_col_width, wider.default_col_width);
        assert_ne!(sheet.images[0].native_ext_cx, wider.images[0].native_ext_cx);
    }

    #[test]
    fn picture_measurement_errors_are_terminal_and_decisions_are_once_only() {
        let (mut session, _, _) = picture_session();
        assert!(session.bootstrap().is_err());
        assert!(session.configure_mdw(Some(7.0)).is_err());

        let (mut session, _, _) = picture_session();
        assert!(session.configure_mdw(Some(f64::NAN)).is_err());
        assert!(session.bootstrap().is_err());

        let (mut session, _, _) = picture_session();
        session.configure_mdw(None).unwrap();
        assert!(session
            .warnings()
            .iter()
            .any(|warning| warning == "legacy-xls:unmeasured-pictures-omitted"));
        assert!(session.configure_mdw(Some(7.0)).is_err());
    }
}
