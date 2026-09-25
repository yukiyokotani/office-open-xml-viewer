//! Passive BIFF8 print metadata admission, never printer execution.
//! [MS-XLS] Setup 2.4.257, WsBool 2.4.351, Header/Footer 2.4.136/124,
//! page breaks 2.4.142/343 and 2.5.160/276.
//!
//! The XLSX renderer model has no page setup, print options, headers,
//! footers or page breaks (print areas and titles are defined names), so
//! these records are only validated: a malformed one fails closed and the
//! remaining facts are not retained.
use super::{f64_at, parse_biff_string, u16_at, unsupported, Record};

#[derive(Default)]
pub(super) struct PrintSettings {
    /// A Setup record (whose header/footer margins it carries) was read.
    setup: bool,
    /// Which of the Left/Right/Top/BottomMargin records were read.
    margins: [bool; 4],
}

impl PrintSettings {
    pub fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        let data = record.data;
        match record.kind {
            0x00a1 => {
                if data.len() != 34 {
                    return Err(unsupported("invalid BIFF Setup length"));
                }
                let flags = u16_at(data, 10)?;
                let page = u16_at(data, 4)? as i16;
                let fit = (u16_at(data, 6)?, u16_at(data, 8)?);
                margin(data, 16, false)?;
                margin(data, 24, false)?;
                if fit.0 > 32767 || fit.1 > 32767 {
                    return Err(unsupported("invalid BIFF fit-to-page limits"));
                }
                // SpreadsheetML firstPageNumber is unsigned; a negative BIFF
                // starting page (fUsePage) has no representable meaning.
                if flags & 128 != 0 && page < 0 {
                    return Err(unsupported("negative BIFF starting page number"));
                }
                self.setup = true;
            }
            0x0026..=0x0029 => {
                if data.len() != 8 {
                    return Err(unsupported("invalid BIFF margin length"));
                }
                margin(data, 0, true)?;
                self.margins[usize::from(record.kind - 0x26)] = true;
            }
            0x0083 | 0x0084 | 0x002a | 0x002b => {
                if data.len() != 2 || u16_at(data, 0)? > 1 {
                    return Err(unsupported("invalid BIFF print option"));
                }
            }
            0x0081 => {
                u16_at(data, 0)?;
            }
            0x0014 | 0x0015 => {
                if !data.is_empty() {
                    if u16_at(data, 0)? > 255 {
                        return Err(unsupported(
                            "BIFF header/footer text exceeds 255 characters",
                        ));
                    }
                    let (_, consumed) = parse_biff_string(data)?;
                    if consumed != data.len() {
                        return Err(unsupported("unexpected BIFF header/footer tail"));
                    }
                }
            }
            0x001b => breaks(data, true)?,
            0x001a => breaks(data, false)?,
            _ => {}
        }
        Ok(())
    }

    /// Setup or margin records are present but the four page margins and
    /// the Setup header/footer margins are not all authored.
    pub fn incomplete_margins(&self) -> bool {
        (self.setup || self.margins.iter().any(|m| *m))
            && (!self.setup || self.margins.iter().any(|m| !*m))
    }
}

fn margin(data: &[u8], offset: usize, inclusive: bool) -> Result<(), String> {
    let value = f64_at(data, offset)?;
    if !value.is_finite() || !(0.0..=49.0).contains(&value) || (!inclusive && value == 49.0) {
        return Err(unsupported("invalid BIFF print margin"));
    }
    Ok(())
}

fn breaks(data: &[u8], horizontal: bool) -> Result<(), String> {
    let count = usize::from(u16_at(data, 0)?);
    if count > if horizontal { 1026 } else { 255 } || data.len() != 2 + count * 6 {
        return Err(unsupported("invalid BIFF page break count or length"));
    }
    let mut previous = None;
    for entry in data[2..].chunks_exact(6) {
        let value = (u16_at(entry, 0)?, u16_at(entry, 2)?, u16_at(entry, 4)?);
        if value.1 >= value.2
            || (horizontal && value.2 > 16383)
            || (!horizontal && value.0 > 255)
            || previous.is_some_and(|(id, end)| value.0 < id || (value.0 == id && value.1 <= end))
        {
            return Err(unsupported(
                "invalid, unsorted or overlapping BIFF page break",
            ));
        }
        previous = Some((value.0, value.2));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn setup(flags: u16) -> Vec<u8> {
        let mut bytes = vec![0xff; 34];
        bytes[6..10].copy_from_slice(&[1, 0, 0, 0]);
        bytes[10..12].copy_from_slice(&flags.to_le_bytes());
        bytes[16..24].copy_from_slice(&0.3f64.to_le_bytes());
        bytes[24..32].copy_from_slice(&0.4f64.to_le_bytes());
        bytes
    }
    fn read(value: &mut PrintSettings, kind: u16, data: &[u8]) -> Result<(), String> {
        value.read(&Record {
            kind,
            offset: 0,
            data,
        })
    }
    #[test]
    fn setup_without_page_margins_is_incomplete() {
        let mut value = PrintSettings::default();
        assert!(!value.incomplete_margins());
        // fNoPls: the undefined printer fields are ignored, not validated.
        read(&mut value, 0xa1, &setup(4)).unwrap();
        assert!(value.incomplete_margins());
        for kind in 0x26..=0x29 {
            read(&mut value, kind, &0.5f64.to_le_bytes()).unwrap();
        }
        assert!(!value.incomplete_margins());
        let mut margins_only = PrintSettings::default();
        for kind in 0x26..=0x29 {
            read(&mut margins_only, kind, &0.5f64.to_le_bytes()).unwrap();
        }
        assert!(margins_only.incomplete_margins());
    }
    #[test]
    fn validates_margins_and_unsupported_signed_page_numbers() {
        let mut value = PrintSettings::default();
        for margin in [f64::NAN, f64::INFINITY, -0.1, 49.1] {
            assert!(read(&mut value, 0x26, &margin.to_le_bytes()).is_err());
        }
        assert!(read(&mut value, 0xa1, &setup(4 | 128)).is_err());
        assert!(read(&mut value, 0x26, &49f64.to_le_bytes()).is_ok());
        assert!(read(&mut value, 0xa1, &setup(4)[..33]).is_err());
    }
    #[test]
    fn admits_full_span_manual_breaks_and_rejects_overlaps() {
        let mut value = PrintSettings::default();
        read(&mut value, 0x1a, &[1, 0, 2, 0, 0, 0, 255, 255]).unwrap();
        assert!(read(
            &mut value,
            0x1b,
            &[2, 0, 5, 0, 0, 0, 2, 0, 5, 0, 2, 0, 3, 0]
        )
        .is_err());
        assert!(read(&mut value, 0x1a, &[0, 1]).is_err());
    }
    #[test]
    fn admits_empty_headers_and_commands_without_executing_them() {
        let mut value = PrintSettings::default();
        read(&mut value, 0x14, &[]).unwrap();
        read(&mut value, 0x15, &[4, 0, 0, b'&', b'P', b'<', b'&']).unwrap();
        assert!(read(&mut value, 0x14, &[0, 1, 0]).is_err());
        assert!(read(&mut value, 0x15, &[1, 0, 0, b'A', 0]).is_err());
    }
}
