//! [MS-XLS] 2.4.345 Window1 / 2.4.346 Window2 -> the XLSX model's
//! sheet-view display flags (ECMA-376 18.3.1.87 sheetView).
//! Only display booleans and frozen panes (2.4.189 Pane, as the XLSX parser
//! reads `pane state="frozen|frozenSplit"`) are projected: no formula-token
//! reconstruction, pane/selection, zoom, scroll-position, or window-geometry
//! inference.

use super::{u16_at, unsupported, Record};

pub(super) const WINDOW1: u16 = 0x003d;
const WINDOW2: u16 = 0x023e;
const PANE: u16 = 0x0041;
// Resource policy, not a BIFF format limit. Bound retained views.
const MAX_WINDOWS: usize = 1024;

/// Window2 fDspGrid, fDspZeros and fRightToLeft; reserved bits are ignored.
fn display_flags(flags: u16) -> (bool, bool, bool) {
    (flags & 0x02 != 0, flags & 0x10 != 0, flags & 0x40 != 0)
}

pub(super) fn read_window(data: &[u8], count: &mut usize) -> Result<(), String> {
    if data.len() != 18 || *count >= MAX_WINDOWS {
        return Err(unsupported("invalid or excessive BIFF workbook windows"));
    }
    *count += 1;
    Ok(())
}

#[derive(Default)]
pub(super) struct SheetViews(Vec<u16>, Vec<Option<(u16, u16)>>);

impl SheetViews {
    /// Apply the display flags of the last Window2, as the XLSX parser keeps
    /// the last of a worksheet's sheetViews (one per workbook view, in
    /// order). This is compatibility behavior, not a claim that the last BIFF
    /// window is normatively the active window.
    pub(super) fn project(&self, worksheet: &mut xlsx_model::Worksheet) {
        let Some(&flags) = self.0.last() else {
            return;
        };
        let (gridlines, zeros, right_to_left) = display_flags(flags);
        worksheet.show_gridlines = gridlines;
        worksheet.show_zeros = zeros;
        worksheet.right_to_left = right_to_left;
        // Window2 fFrozen (bit 3): the Pane split counts cells. An unfrozen
        // split (twips) is a window arrangement the XLSX model does not
        // carry, as the XLSX parser reads only frozen panes.
        if flags & 0x0008 != 0 {
            if let Some(Some((x, y))) = self.1.last() {
                worksheet.freeze_cols = u32::from(*x);
                worksheet.freeze_rows = u32::from(*y);
            }
        }
    }

    pub(super) fn displays_formulas(&self) -> bool {
        self.0.iter().any(|flags| flags & 1 != 0)
    }
    pub(super) fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        if record.kind == PANE {
            // WINDOW = Window2 [PLV] [Scl] [Pane] *Selection (2.1.7.20.5).
            let pane = self
                .1
                .last_mut()
                .filter(|pane| pane.is_none())
                .ok_or_else(|| unsupported("BIFF pane without its window"))?;
            if record.data.len() != 10 {
                return Err(unsupported("invalid BIFF pane"));
            }
            let (x, y) = (u16_at(record.data, 0)?, u16_at(record.data, 2)?);
            let frozen = self.0.last().is_some_and(|flags| flags & 0x0008 != 0);
            if frozen && x > 255 {
                return Err(unsupported("invalid BIFF frozen pane"));
            }
            *pane = Some((x, y));
            return Ok(());
        }
        if record.kind != WINDOW2 {
            return Ok(());
        }
        // Worksheet Window2 is 18 bytes. Chart Window2 is a different record
        // layout and is excluded by the caller's BOF/EOF ownership guard.
        if record.data.len() != 18 || self.0.len() >= MAX_WINDOWS {
            return Err(unsupported("invalid or excessive BIFF worksheet windows"));
        }
        self.0.push(u16_at(record.data, 0)?);
        self.1.push(None);
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
}

#[cfg(test)]
mod tests {
    use super::*;

    fn projected(views: &SheetViews) -> (bool, bool, bool) {
        let mut model = xlsx_model::Worksheet::placeholder("S", "test".into());
        model.show_gridlines = true;
        model.show_zeros = true;
        model.right_to_left = false;
        views.project(&mut model);
        (model.show_gridlines, model.show_zeros, model.right_to_left)
    }

    #[test]
    fn all_flag_combinations_project_only_the_three_display_bits() {
        for flags in 0..=u16::MAX {
            let views = SheetViews(vec![flags], vec![None]);
            assert_eq!(
                projected(&views),
                (flags & 0x02 != 0, flags & 0x10 != 0, flags & 0x40 != 0)
            );
            assert_eq!(
                projected(&views),
                projected(&SheetViews(vec![flags & 0x52], vec![None]))
            );
        }
    }

    #[test]
    fn frozen_panes_follow_their_window() {
        let record = |kind, data: &'static [u8]| Record {
            kind,
            offset: 0,
            data,
        };
        let mut window = [0u8; 18];
        window[0] = 0x08;
        let window: &'static [u8] = Box::leak(Box::new(window));
        let mut views = SheetViews::default();
        views.read(&record(WINDOW2, window)).unwrap();
        views
            .read(&record(PANE, &[2, 0, 7, 0, 7, 0, 2, 0, 0, 0]))
            .unwrap();
        let mut model = xlsx_model::Worksheet::placeholder("S", "test".into());
        views.project(&mut model);
        assert_eq!((model.freeze_rows, model.freeze_cols), (7, 2));
        // A second pane for the same window, or one without a window, rejects.
        assert!(views
            .read(&record(PANE, &[2, 0, 7, 0, 7, 0, 2, 0, 0, 0]))
            .is_err());
        assert!(SheetViews::default().read(&record(PANE, &[0; 10])).is_err());
        // An unfrozen split is not a frozen pane.
        let mut views = SheetViews::default();
        views.read(&record(WINDOW2, &[0; 18])).unwrap();
        views
            .read(&record(PANE, &[0x40, 0x1f, 0, 0, 0, 0, 0, 0, 0, 0]))
            .unwrap();
        let mut model = xlsx_model::Worksheet::placeholder("S", "test".into());
        views.project(&mut model);
        assert_eq!((model.freeze_rows, model.freeze_cols), (0, 0));
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
    }

    #[test]
    fn model_projection_uses_the_last_view_and_keeps_defaults_without_one() {
        let views = SheetViews(vec![0x52, 0x04], vec![None, None]);
        assert_eq!(projected(&views), (false, false, false));
        let views = SheetViews(vec![0x04, 0x52], vec![None, None]);
        assert_eq!(projected(&views), (true, true, true));
        assert_eq!(projected(&SheetViews::default()), (true, true, false));
    }
}
