//! [MS-XLS] 2.4.345 Window1 / 2.4.346 Window2 -> ECMA-376
//! 18.3.1.87 sheetView and 18.2.30 workbookView.
//! Only display booleans are projected: no formula-token reconstruction,
//! pane/selection, zoom, scroll-position, or window-geometry inference.

use super::{u16_at, unsupported, Record};

pub(super) const WINDOW1: u16 = 0x003d;
const WINDOW2: u16 = 0x023e;
// Resource policy, not a BIFF format limit. Bound retained views and XML fanout.
const MAX_WINDOWS: usize = 1024;

pub(super) fn read_window(data: &[u8], count: &mut usize) -> Result<(), String> {
    if data.len() != 18 || *count >= MAX_WINDOWS {
        return Err(unsupported("invalid or excessive BIFF workbook windows"));
    }
    *count += 1;
    Ok(())
}

#[derive(Default)]
pub(super) struct SheetViews(Vec<u16>);

impl SheetViews {
    pub(super) fn displays_formulas(&self) -> bool {
        self.0.iter().any(|flags| flags & 1 != 0)
    }
    pub(super) fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        if record.kind != WINDOW2 {
            return Ok(());
        }
        // Worksheet Window2 is 18 bytes. Chart Window2 is a different record
        // layout and is excluded by the caller's BOF/EOF ownership guard.
        if record.data.len() != 18 || self.0.len() >= MAX_WINDOWS {
            return Err(unsupported("invalid or excessive BIFF worksheet windows"));
        }
        self.0.push(u16_at(record.data, 0)?);
        Ok(())
    }

    pub(super) fn validate_count(&self, count: usize) -> Result<(), String> {
        if self.0.len() != count {
            return Err(unsupported(
                "BIFF worksheet windows do not match workbook windows",
            ));
        }
        Ok(())
    }

    pub(super) fn xml(&self) -> String {
        if self.0.is_empty() {
            return String::new();
        }
        let mut xml = String::from("<sheetViews>");
        for (id, flags) in self.0.iter().enumerate() {
            // Explicit zeros matter: OOXML defaults grid/headers/zeros to true.
            // Reserved bits are ignored, as required by Window2.
            xml.push_str(&format!(
                "<sheetView workbookViewId=\"{id}\" showGridLines=\"{}\" showRowColHeaders=\"{}\" showZeros=\"{}\" rightToLeft=\"{}\"/>",
                u8::from(flags & 0x02 != 0), u8::from(flags & 0x04 != 0),
                u8::from(flags & 0x10 != 0), u8::from(flags & 0x40 != 0),
            ));
        }
        xml.push_str("</sheetViews>");
        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn all_flag_combinations_project_only_the_four_display_bits() {
        for flags in 0..=u16::MAX {
            let xml = SheetViews(vec![flags]).xml();
            for (attr, bit) in [
                ("showGridLines", 2),
                ("showRowColHeaders", 4),
                ("showZeros", 16),
                ("rightToLeft", 64),
            ] {
                assert!(xml.contains(&format!("{attr}=\"{}\"", u8::from(flags & bit != 0))));
            }
            assert_eq!(xml, SheetViews(vec![flags & 0x56]).xml());
        }
    }

    #[test]
    fn bounds_and_window_association_are_checked() {
        let mut count = 0;
        for _ in 0..MAX_WINDOWS {
            read_window(&[0; 18], &mut count).unwrap();
        }
        assert!(read_window(&[0; 18], &mut count).is_err());
        assert!(read_window(&[0; 17], &mut 0).is_err());
        let mut views = SheetViews::default();
        assert!(views.xml().is_empty());
        for len in [0, 10, 17, 19] {
            assert!(views
                .read(&Record {
                    kind: WINDOW2,
                    offset: 0,
                    data: &vec![0; len]
                })
                .is_err());
        }
        for _ in 0..MAX_WINDOWS {
            views
                .read(&Record {
                    kind: WINDOW2,
                    offset: 0,
                    data: &[0; 18],
                })
                .unwrap();
        }
        assert!(views
            .read(&Record {
                kind: WINDOW2,
                offset: 0,
                data: &[0; 18]
            })
            .is_err());
        views.validate_count(MAX_WINDOWS).unwrap();
        assert!(views.validate_count(MAX_WINDOWS - 1).is_err());
        assert!(views.xml().contains("workbookViewId=\"1023\""));
    }
}
