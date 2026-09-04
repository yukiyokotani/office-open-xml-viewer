//! [MS-XLS] Row (2.4.221), ColInfo (2.4.53), DefaultRowHeight
//! (2.4.87), DefColWidth (2.4.89) -> ECMA-376 18.3.1 row/col/sheetFormatPr.
use super::{styles::Styles, u16_at, u32_at, unsupported, Record};
use std::collections::BTreeMap;

#[derive(Default)]
pub(super) struct Geometry {
    rows: BTreeMap<u16, (u16, u32)>,
    columns: BTreeMap<u16, (u16, u16, u16)>,
    default_row: Option<(u16, u16)>,
    base_width: Option<u16>,
    default_column: Option<(u16, u16, u16)>,
}

impl Geometry {
    pub fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        let data = record.data;
        match record.kind {
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
        let mut xml = format!(" ht=\"{}\" hidden=\"{}\" customHeight=\"{}\" outlineLevel=\"{}\" collapsed=\"{}\" thickTop=\"{}\" thickBot=\"{}\"", f64::from(*height) / 20.0, (flags >> 5) & 1, (flags >> 6) & 1, flags & 7, (flags >> 4) & 1, (flags >> 28) & 1, (flags >> 29) & 1);
        if flags & 0x80 != 0 {
            xml.push_str(&format!(
                " s=\"{}\" customFormat=\"1\"",
                (flags >> 16) & 0xfff
            ));
        }
        xml
    }

    pub fn xml(&self) -> String {
        let mut xml = String::new();
        if let Some((height, flags)) = self.default_row {
            xml.push_str(&format!("<sheetFormatPr defaultRowHeight=\"{}\" customHeight=\"{}\" zeroHeight=\"{}\" thickTop=\"{}\" thickBottom=\"{}\"", f64::from(height) / 20.0, flags & 1, (flags >> 1) & 1, (flags >> 2) & 1, (flags >> 3) & 1));
            // DefColWidth excludes margin padding, unlike CT_Col.width. Keep
            // it as baseColWidth; inventing a font-dependent padding loses fidelity.
            if let Some(width) = self.base_width {
                xml.push_str(&format!(" baseColWidth=\"{width}\""));
            }
            if let Some((width, _, _)) = self.default_column {
                xml.push_str(&format!(
                    " defaultColWidth=\"{}\"",
                    f64::from(width) / 256.0
                ));
            }
            xml.push_str("/>");
        }
        if !self.columns.is_empty() {
            xml.push_str("<cols>");
            for (column, (width, style, flags)) in &self.columns {
                let column = u32::from(*column) + 1;
                xml.push_str(&format!("<col min=\"{column}\" max=\"{column}\" width=\"{}\" style=\"{style}\" hidden=\"{}\" customWidth=\"{}\" bestFit=\"{}\" outlineLevel=\"{}\" collapsed=\"{}\"/>", f64::from(*width) / 256.0, flags & 1, (flags >> 1) & 1, (flags >> 2) & 1, (flags >> 8) & 7, (flags >> 12) & 1));
            }
            xml.push_str("</cols>");
        }
        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;
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
}
