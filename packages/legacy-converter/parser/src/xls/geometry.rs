//! [MS-XLS] Row (2.4.221), ColInfo (2.4.53), DefaultRowHeight
//! (2.4.87), DefColWidth (2.4.89) -> ECMA-376 18.3.1 row/col/sheetFormatPr.
use super::{styles::Styles, u16_at, u32_at, unsupported, Record};
use std::collections::BTreeMap;

struct RowFacts {
    height: f64,
    hidden: bool,
    custom: bool,
    outline: u8,
    collapsed: bool,
    thick_top: bool,
    thick_bottom: bool,
    style: Option<u16>,
}

struct ColumnFacts {
    width: f64,
    style: u16,
    hidden: bool,
    custom: bool,
    best_fit: bool,
    outline: u8,
    collapsed: bool,
}

fn row_facts(height: u16, flags: u32) -> RowFacts {
    RowFacts {
        height: f64::from(height) / 20.0,
        hidden: flags & (1 << 5) != 0,
        custom: flags & (1 << 6) != 0,
        outline: (flags & 7) as u8,
        collapsed: flags & (1 << 4) != 0,
        thick_top: flags & (1 << 28) != 0,
        thick_bottom: flags & (1 << 29) != 0,
        style: (flags & 0x80 != 0).then_some(((flags >> 16) & 0xfff) as u16),
    }
}

fn column_facts(width: u16, style: u16, flags: u16) -> ColumnFacts {
    ColumnFacts {
        width: f64::from(width) / 256.0,
        style,
        hidden: flags & 1 != 0,
        custom: flags & 2 != 0,
        best_fit: flags & 4 != 0,
        outline: ((flags >> 8) & 7) as u8,
        collapsed: flags & (1 << 12) != 0,
    }
}

#[derive(Default)]
pub(super) struct Geometry {
    rows: BTreeMap<u16, (u16, u32)>,
    columns: BTreeMap<u16, (u16, u16, u16)>,
    default_row: Option<(u16, u16)>,
    base_width: Option<u16>,
    default_column: Option<(u16, u16, u16)>,
    digit_width: Option<u16>,
    unknown_digit_width: bool,
}

impl Geometry {
    /// Apply already-validated BIFF geometry to the renderer worksheet model.
    /// The 8.43/15.0 values are the existing XLSX parser's model defaults when
    /// SpreadsheetML omits sheetFormatPr; they are not additional BIFF facts.
    pub(super) fn project(
        &self,
        worksheet: &mut xlsx_model::Worksheet,
        mdw: Option<f64>,
        budget: &mut usize,
    ) -> Result<(), String> {
        let width_ranges = self
            .columns
            .values()
            .filter(|(_, _, flags)| flags & 3 == 0)
            .count();
        let width_points = self.columns.len() - width_ranges;
        let outlines = self
            .columns
            .values()
            .filter(|(_, _, flags)| flags & 0x700 != 0)
            .count();
        let collapsed = self
            .columns
            .values()
            .filter(|(_, _, flags)| flags & 0x1000 != 0)
            .count();
        let hidden = self
            .columns
            .values()
            .filter(|(_, _, flags)| flags & 1 != 0)
            .count();
        let row_heights = worksheet
            .rows
            .iter()
            .filter(|row| {
                row.index
                    .checked_sub(1)
                    .and_then(|value| u16::try_from(value).ok())
                    .is_some_and(|index| self.rows.contains_key(&index))
                    || (self.default_row.is_some_and(|(_, flags)| flags & 2 != 0)
                        && !row.hidden
                        && row.height.is_none())
            })
            .count();
        // Allocator node overhead is implementation-private. Charge logical
        // key/value payload; BIFF bounds column maps to 256 entries and the
        // already-admitted worksheet bounds row heights.
        let payload = checked_product(
            self.columns.len(),
            std::mem::size_of::<xlsx_model::ColumnStyleRange>(),
        )?
        .checked_add(checked_product(
            width_ranges,
            std::mem::size_of::<xlsx_model::ColumnWidthRange>(),
        )?)
        .and_then(|n| n.checked_add(width_points * std::mem::size_of::<(u32, f64)>()))
        .and_then(|n| n.checked_add(outlines * std::mem::size_of::<(u32, u8)>()))
        .and_then(|n| n.checked_add((collapsed + hidden) * std::mem::size_of::<(u32, bool)>()))
        .and_then(|n| n.checked_add(row_heights * std::mem::size_of::<(u32, f64)>()))
        .ok_or_else(model_budget_error)?;
        *budget = budget.checked_sub(payload).ok_or_else(model_budget_error)?;
        worksheet
            .col_style_ranges
            .try_reserve_exact(self.columns.len())
            .map_err(|_| model_budget_error())?;
        worksheet
            .col_width_ranges
            .try_reserve_exact(width_ranges)
            .map_err(|_| model_budget_error())?;
        worksheet.default_col_width = 8.43;
        worksheet.default_row_height = 15.0;
        worksheet.default_row_height_custom = false;

        if let Some((height, flags)) = self.default_row {
            let authored_height = f64::from(height) / 20.0;
            worksheet.default_row_height = if flags & 2 != 0 { 0.0 } else { authored_height };
            worksheet.default_row_height_custom = flags & 1 != 0;
            if flags & 2 != 0 {
                for row in &mut worksheet.rows {
                    if !row.hidden && row.height.is_none() {
                        row.height = Some(authored_height);
                        worksheet.row_heights.insert(row.index, authored_height);
                    }
                }
            }
            if let Some(width) = self.sheet_default_width(mdw) {
                worksheet.default_col_width = width / 256.0;
            }
        }

        for (&column, &(width, style, flags)) in &self.columns {
            let column = u32::from(column) + 1;
            let facts = column_facts(width, style, flags);
            let width = if facts.hidden { 0.0 } else { facts.width };
            worksheet
                .col_style_ranges
                .push(xlsx_model::ColumnStyleRange {
                    min: column,
                    max: column,
                    style_index: u32::from(facts.style),
                });
            if facts.custom || facts.hidden {
                worksheet.col_widths.insert(column, width);
            } else {
                worksheet
                    .col_width_ranges
                    .push(xlsx_model::ColumnWidthRange {
                        min: column,
                        max: column,
                        width,
                    });
            }
            if facts.outline != 0 {
                worksheet.col_outline_levels.insert(column, facts.outline);
            }
            if facts.collapsed {
                worksheet.col_collapsed.insert(column, true);
            }
            if facts.hidden {
                worksheet.col_hidden.insert(column, true);
            }
        }

        for row in &mut worksheet.rows {
            let Some(index) = row
                .index
                .checked_sub(1)
                .and_then(|value| u16::try_from(value).ok())
            else {
                continue;
            };
            let Some(&(height, flags)) = self.rows.get(&index) else {
                continue;
            };
            let facts = row_facts(height, flags);
            row.hidden = facts.hidden;
            row.height = Some(if row.hidden { 0.0 } else { facts.height });
            row.custom_height = facts.custom;
            row.outline_level = facts.outline;
            row.collapsed = facts.collapsed;
            worksheet.row_heights.insert(row.index, row.height.unwrap());
        }
        Ok(())
    }

    pub fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        let data = record.data;
        match record.kind {
            // MS-XLS 2.4.98, record 153: two-byte DxGCol, already in
            // 1/256 Normal digit widths. Do not interpret an unknown layout.
            0x0099 => {
                if data.len() == 2 {
                    self.digit_width = Some(u16_at(data, 0)?);
                } else {
                    self.unknown_digit_width = true;
                }
            }
            0x0208 => {
                let row = u16_at(data, 0)?;
                let height = u16_at(data, 6)?;
                if u16_at(data, 2)? > 255 || u16_at(data, 4)? > 256 || !(2..=8192).contains(&height)
                {
                    return Err(unsupported("invalid BIFF row dimensions"));
                }
                self.rows.insert(row, (height, u32_at(data, 12)?));
            }
            0x007d => {
                let first = u16_at(data, 0)?;
                let last = u16_at(data, 2)?;
                if first > last || last > 256 {
                    return Err(unsupported("invalid BIFF column range"));
                }
                let value = (u16_at(data, 4)?, u16_at(data, 6)?, u16_at(data, 8)?);
                // Col256U (2.5.44): 256 is the default formatting sentinel,
                // not a real BIFF8 column. Never synthesize column IW.
                if last == 256 {
                    self.default_column = Some(value);
                }
                for column in first..=last.min(255) {
                    self.columns.insert(column, value);
                }
            }
            0x0225 => {
                let flags = u16_at(data, 0)?;
                let height = u16_at(data, 2)?;
                if height > 8179 || (height == 0 && flags & 2 == 0) {
                    return Err(unsupported("invalid BIFF default row height"));
                }
                self.default_row = Some((height, flags));
            }
            0x0055 => {
                let width = u16_at(data, 0)?;
                if width > 255 {
                    return Err(unsupported("invalid BIFF default column width"));
                }
                self.base_width = Some(width);
            }
            _ => {}
        }
        Ok(())
    }

    pub fn validate_styles(&self, styles: &Styles<'_>) -> Result<(), String> {
        for (_, style, _) in self.columns.values() {
            styles.validate_xf(*style)?;
        }
        for (_, flags) in self.rows.values() {
            if flags & 0x80 != 0 {
                styles.validate_xf(((flags >> 16) & 0xfff) as u16)?;
            }
        }
        Ok(())
    }

    pub fn row_attributes(&self, row: u16) -> String {
        let Some((height, flags)) = self.rows.get(&row) else {
            return String::new();
        };
        let facts = row_facts(*height, *flags);
        let mut xml = format!(" ht=\"{}\" hidden=\"{}\" customHeight=\"{}\" outlineLevel=\"{}\" collapsed=\"{}\" thickTop=\"{}\" thickBot=\"{}\"", facts.height, u8::from(facts.hidden), u8::from(facts.custom), facts.outline, u8::from(facts.collapsed), u8::from(facts.thick_top), u8::from(facts.thick_bottom));
        if let Some(style) = facts.style {
            xml.push_str(&format!(" s=\"{style}\" customFormat=\"1\""));
        }
        xml
    }

    #[cfg(test)]
    pub fn xml(&self) -> String {
        self.xml_with_metrics(None)
    }

    pub(super) fn xml_with_metrics(&self, mdw: Option<f64>) -> String {
        let mut xml = String::new();
        if let Some((height, flags)) = self.default_row {
            xml.push_str(&format!("<sheetFormatPr defaultRowHeight=\"{}\" customHeight=\"{}\" zeroHeight=\"{}\" thickTop=\"{}\" thickBottom=\"{}\"", f64::from(height) / 20.0, flags & 1, (flags >> 1) & 1, (flags >> 2) & 1, (flags >> 3) & 1));
            // DefColWidth excludes margin padding, unlike CT_Col.width. Keep
            // it as baseColWidth; inventing a font-dependent padding loses fidelity.
            if let Some(width) = self.base_width {
                xml.push_str(&format!(" baseColWidth=\"{width}\""));
            }
            if let Some(width) = self.sheet_default_width(mdw) {
                xml.push_str(&format!(" defaultColWidth=\"{}\"", width / 256.0));
            }
            xml.push_str("/>");
        }
        if !self.columns.is_empty() {
            xml.push_str("<cols>");
            for (column, (width, style, flags)) in &self.columns {
                let column = u32::from(*column) + 1;
                let facts = column_facts(*width, *style, *flags);
                xml.push_str(&format!("<col min=\"{column}\" max=\"{column}\" width=\"{}\" style=\"{}\" hidden=\"{}\" customWidth=\"{}\" bestFit=\"{}\" outlineLevel=\"{}\" collapsed=\"{}\"/>", facts.width, facts.style, u8::from(facts.hidden), u8::from(facts.custom), u8::from(facts.best_fit), facts.outline, u8::from(facts.collapsed)));
            }
            xml.push_str("</cols>");
        }
        xml
    }

    fn default_width(&self, mdw: f64) -> Option<f64> {
        if self.unknown_digit_width {
            return None;
        }
        self.digit_width
            .map(f64::from)
            .or_else(|| self.default_column.map(|c| f64::from(c.0)))
            // ECMA-376 18.3.1.13/81: base character count excludes the
            // normative four margin pixels and one gridline pixel.
            .or_else(|| {
                self.base_width
                    .map(|w| ((f64::from(w) + 5.0 / mdw) * 256.0).trunc())
            })
    }

    /// Width stored by DxGCol and ColInfo is already in 1/256 Normal digit
    /// units. Only DefColWidth's character count needs a measured digit width.
    /// A non-two-byte DxGCol keeps the existing conservative unknown-layout
    /// policy: ignore both its value and the font-dependent base-width route,
    /// while retaining an independently authored default ColInfo sentinel.
    fn sheet_default_width(&self, mdw: Option<f64>) -> Option<f64> {
        if !self.unknown_digit_width {
            if let Some(width) = self.digit_width {
                return Some(f64::from(width));
            }
        }
        self.default_column
            .map(|column| f64::from(column.0))
            .or_else(|| mdw.and_then(|value| self.default_width(value)))
    }

    pub(super) fn column_emu(&self, col: u16, mdw: f64) -> Option<f64> {
        let width = if let Some(&(width, _, flags)) = self.columns.get(&col) {
            if flags & 1 != 0 {
                return Some(0.0);
            }
            f64::from(width)
        } else {
            self.default_width(mdw)?
        };
        Some(((width + (128.0 / mdw).trunc()) / 256.0 * mdw).trunc() * 9525.0)
    }

    pub(super) fn has_sheet_defaults(&self) -> bool {
        self.default_row.is_some()
    }

    pub(super) fn row_emu(&self, row: u16) -> Option<f64> {
        if let Some(&(height, flags)) = self.rows.get(&row) {
            Some(if flags & 32 != 0 {
                0.0
            } else {
                f64::from(height) * 635.0
            })
        } else {
            self.default_row.map(|(height, flags)| {
                if flags & 2 != 0 {
                    0.0
                } else {
                    f64::from(height) * 635.0
                }
            })
        }
    }
}

fn checked_product(count: usize, size: usize) -> Result<usize, String> {
    count.checked_mul(size).ok_or_else(model_budget_error)
}

fn model_budget_error() -> String {
    unsupported("XLS direct geometry model byte budget exceeded")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn worksheet() -> xlsx_model::Worksheet {
        xlsx_model::Worksheet::placeholder("S", "test".into())
    }

    fn row(index: u32) -> xlsx_model::Row {
        xlsx_model::Row {
            index,
            height: None,
            custom_height: false,
            cells: Vec::new(),
            outline_level: 0,
            collapsed: false,
            hidden: false,
        }
    }
    #[test]
    fn stored_digit_width_does_not_require_host_font_measurement() {
        // [MS-XLS] 2.4.98: DxGCol already stores width in 1/256 Normal
        // digit units. Only conversion to pixels needs a measured digit width.
        for width in [0_u16, 1, 2560, 65535] {
            let mut geometry = Geometry::default();
            geometry.default_row = Some((300, 0));
            geometry.default_column = Some((3072, 0, 0));
            geometry
                .read(&Record {
                    kind: 0x0099,
                    offset: 0,
                    data: &width.to_le_bytes(),
                })
                .unwrap();
            for mdw in [None, Some(7.0), Some(9.0)] {
                let mut model = worksheet();
                let mut budget = usize::MAX;
                geometry.project(&mut model, mdw, &mut budget).unwrap();
                assert_eq!(model.default_col_width, f64::from(width) / 256.0);
                assert!(geometry
                    .xml_with_metrics(mdw)
                    .contains(&format!("defaultColWidth=\"{}\"", f64::from(width) / 256.0)));
            }
        }
    }

    #[test]
    fn unknown_digit_width_keeps_default_column_and_explicit_column_policy() {
        let mut geometry = Geometry::default();
        geometry.default_row = Some((300, 0));
        geometry.default_column = Some((3072, 0, 0));
        geometry.base_width = Some(8);
        geometry.digit_width = Some(2560);
        // A non-two-byte DxGCol selects the existing unknown-layout policy. It
        // invalidates the stored digit width and font-dependent base route,
        // but not the independent default ColInfo record.
        geometry
            .read(&Record {
                kind: 0x0099,
                offset: 0,
                data: &[0, 10, 0, 0],
            })
            .unwrap();
        geometry.columns.insert(0, (4096, 0, 2));

        for mdw in [None, Some(7.0), Some(9.0)] {
            let mut model = worksheet();
            let mut budget = usize::MAX;
            geometry.project(&mut model, mdw, &mut budget).unwrap();
            assert_eq!(model.default_col_width, 12.0);
            assert_eq!(model.col_widths.get(&1), Some(&16.0));
            assert!(geometry
                .xml_with_metrics(mdw)
                .contains("defaultColWidth=\"12\""));
        }
    }

    #[test]
    fn stored_digit_width_without_default_column_and_explicit_override_are_distinct() {
        let mut geometry = Geometry::default();
        geometry.default_row = Some((300, 0));
        geometry.digit_width = Some(2560);
        geometry.columns.insert(0, (4096, 0, 2));
        let mut model = worksheet();
        let mut budget = usize::MAX;
        geometry.project(&mut model, None, &mut budget).unwrap();
        assert_eq!(model.default_col_width, 10.0);
        assert_eq!(model.col_widths.get(&1), Some(&16.0));
        assert!(geometry.xml().contains("defaultColWidth=\"10\""));
    }

    #[test]
    fn unknown_digit_width_without_default_column_omits_width_even_when_measured() {
        let mut geometry = Geometry::default();
        geometry.default_row = Some((300, 0));
        geometry.base_width = Some(8);
        geometry
            .read(&Record {
                kind: 0x0099,
                offset: 0,
                data: &[0, 10, 0, 0],
            })
            .unwrap();
        let mut model = worksheet();
        let mut budget = usize::MAX;
        geometry
            .project(&mut model, Some(7.0), &mut budget)
            .unwrap();
        assert_eq!(model.default_col_width, 8.43);
        assert!(!geometry
            .xml_with_metrics(Some(7.0))
            .contains("defaultColWidth"));
    }

    #[test]
    fn measured_width_uses_normative_padding_and_explicit_digit_width_precedence() {
        let mut geometry = Geometry::default();
        assert_eq!(geometry.column_emu(0, 7.0), None);
        geometry
            .read(&Record {
                kind: 0x55,
                offset: 0,
                data: &[8, 0],
            })
            .unwrap();
        geometry
            .read(&Record {
                kind: 0x225,
                offset: 0,
                data: &[0, 0, 44, 1],
            })
            .unwrap();
        assert_eq!(geometry.column_emu(0, 7.0), Some(61.0 * 9525.0));
        assert!(geometry
            .xml_with_metrics(Some(7.0))
            .contains("defaultColWidth=\"8.7109375\""));
        assert!(!geometry.xml().contains("defaultColWidth"));
        geometry
            .read(&Record {
                kind: 0x99,
                offset: 0,
                data: &[0, 10],
            })
            .unwrap();
        assert_eq!(geometry.column_emu(0, 7.0), Some(70.0 * 9525.0));
        assert_eq!(geometry.column_emu(0, 9.0), Some(90.0 * 9525.0));
        assert_eq!(geometry.row_emu(0), Some(300.0 * 635.0));
        geometry
            .read(&Record {
                kind: 0x7d,
                offset: 0,
                data: &[0, 0, 0, 0, 0, 12, 0, 0, 1, 0],
            })
            .unwrap();
        assert_eq!(geometry.column_emu(0, 7.0), Some(0.0));
        assert_eq!(geometry.column_emu(1, 7.0), Some(70.0 * 9525.0));
    }
    #[test]
    fn column_256_sets_default_width_without_creating_an_extra_column() {
        let mut geometry = Geometry::default();
        geometry
            .read(&Record {
                kind: 0x0225,
                offset: 0,
                data: &[0, 0, 44, 1],
            })
            .unwrap();
        geometry
            .read(&Record {
                kind: 0x007d,
                offset: 0,
                data: &[255, 0, 0, 1, 0, 12, 0, 0, 0, 0, 0, 0],
            })
            .unwrap();
        let xml = geometry.xml();
        assert!(xml.contains("defaultColWidth=\"12\""));
        assert!(xml.contains("min=\"256\" max=\"256\""));
        assert!(!xml.contains("min=\"257\""));
    }
    #[test]
    fn preserves_default_hidden_rows_and_rejects_out_of_range_geometry() {
        let mut geometry = Geometry::default();
        geometry
            .read(&Record {
                kind: 0x0225,
                offset: 0,
                data: &[2, 0, 44, 1],
            })
            .unwrap();
        assert!(geometry.xml().contains("zeroHeight=\"1\""));
        assert!(geometry
            .read(&Record {
                kind: 0x007d,
                offset: 0,
                data: &[0, 0, 1, 1, 0, 0, 0, 0, 0, 0]
            })
            .is_err());
        assert!(geometry
            .read(&Record {
                kind: 0x0208,
                offset: 0,
                data: &[0; 16]
            })
            .is_err());
    }

    #[test]
    fn model_projection_preserves_defaults_rows_columns_and_flags() {
        let mut geometry = Geometry::default();
        geometry
            .read(&Record {
                kind: 0x0225,
                offset: 0,
                data: &[3, 0, 44, 1], // custom default, unspecified rows hidden
            })
            .unwrap();
        geometry
            .read(&Record {
                kind: 0x007d,
                offset: 0,
                // column A, width 12, style 2, hidden+custom+outline3+collapsed
                data: &[0, 0, 0, 0, 0, 12, 2, 0, 3, 0x13],
            })
            .unwrap();
        let row_flags = (2_u32) | (1 << 4) | (1 << 6);
        let mut row_record = [0_u8; 16];
        row_record[0..2].copy_from_slice(&1_u16.to_le_bytes());
        row_record[2..4].copy_from_slice(&0_u16.to_le_bytes());
        row_record[4..6].copy_from_slice(&1_u16.to_le_bytes());
        row_record[6..8].copy_from_slice(&400_u16.to_le_bytes());
        row_record[12..16].copy_from_slice(&row_flags.to_le_bytes());
        geometry
            .read(&Record {
                kind: 0x0208,
                offset: 0,
                data: &row_record,
            })
            .unwrap();

        let mut rejected = worksheet();
        rejected.rows = vec![row(1), row(2)];
        assert!(geometry.project(&mut rejected, None, &mut 0).is_err());
        assert!(rejected.col_style_ranges.is_empty());
        assert!(rejected.row_heights.is_empty());

        let mut model = worksheet();
        model.rows = vec![row(1), row(2)];
        let mut budget = usize::MAX;
        geometry.project(&mut model, None, &mut budget).unwrap();
        assert_eq!(model.default_col_width, 8.43);
        assert_eq!(model.default_row_height, 0.0);
        assert!(model.default_row_height_custom);
        assert_eq!(model.col_widths.get(&1), Some(&0.0));
        assert_eq!(model.col_style_ranges[0].style_index, 2);
        assert_eq!(model.col_outline_levels.get(&1), Some(&3));
        assert_eq!(model.col_collapsed.get(&1), Some(&true));
        assert_eq!(model.col_hidden.get(&1), Some(&true));
        assert_eq!(
            (model.rows[0].height, model.rows[0].outline_level),
            (Some(15.0), 0)
        );
        assert_eq!(
            (model.rows[1].height, model.rows[1].outline_level),
            (Some(20.0), 2)
        );
        assert!(model.rows[1].custom_height && model.rows[1].collapsed);
        assert_eq!(model.row_heights.get(&1), Some(&15.0));
        assert_eq!(model.row_heights.get(&2), Some(&20.0));
    }

    #[test]
    fn native_geometry_matches_parser_across_flag_combinations() {
        use super::super::{build_xlsx_with_drawings, styles, CellValue, SheetData};
        for default_flags in 0..4 {
            for flags in 0..128_u32 {
                let mut source = SheetData::default();
                source.geometry.default_row = Some((420, default_flags));
                source.geometry.default_column = Some((2560, 0, 0));
                source.geometry.rows.insert(1, (360, flags));
                let col_flags = (flags as u16 & 7)
                    | (((flags as u16 >> 3) & 7) << 8)
                    | (((flags as u16 >> 6) & 1) << 12);
                source.geometry.columns.insert(0, (3072, 0, col_flags));
                source
                    .rows
                    .insert(0, BTreeMap::from([(0, CellValue::Number(1.0))]));
                source
                    .rows
                    .insert(1, BTreeMap::from([(0, CellValue::Number(2.0))]));
                let mut actual = worksheet();
                actual.rows = vec![row(1), row(2)];
                let mut budget = usize::MAX;
                source
                    .geometry
                    .project(&mut actual, None, &mut budget)
                    .unwrap();
                let bytes = build_xlsx_with_drawings(
                    &[("S".into(), source)],
                    &styles::minimal_resolved(),
                    Vec::new(),
                    false,
                    1,
                    1024 * 1024,
                    None,
                    None,
                )
                .unwrap();
                let expected: serde_json::Value =
                    serde_json::from_str(&xlsx_parser::parse_sheet_native(&bytes, 0, "S").unwrap())
                        .unwrap();
                let actual = serde_json::to_value(actual).unwrap();
                for field in [
                    "defaultColWidth",
                    "defaultRowHeight",
                    "defaultRowHeightCustom",
                    "colWidths",
                    "colWidthRanges",
                    "colStyleRanges",
                    "colOutlineLevels",
                    "colCollapsed",
                    "colHidden",
                    "rowHeights",
                ] {
                    assert_eq!(
                        actual[field], expected[field],
                        "{field}: default={default_flags}, flags={flags}"
                    );
                }
                for i in 0..2 {
                    for field in [
                        "height",
                        "customHeight",
                        "outlineLevel",
                        "collapsed",
                        "hidden",
                    ] {
                        assert_eq!(
                            actual["rows"][i][field], expected["rows"][i][field],
                            "row {i} {field}: default={default_flags}, flags={flags}"
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn absent_sheet_format_uses_existing_parser_model_defaults() {
        let mut model = worksheet();
        let mut budget = usize::MAX;
        Geometry::default()
            .project(&mut model, Some(7.0), &mut budget)
            .unwrap();
        assert_eq!(
            (model.default_col_width, model.default_row_height),
            (8.43, 15.0)
        );
        assert!(!model.default_row_height_custom);
    }
}
