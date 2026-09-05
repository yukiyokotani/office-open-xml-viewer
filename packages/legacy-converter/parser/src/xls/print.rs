//! Passive BIFF8 print metadata -> SpreadsheetML, never printer execution.
//! [MS-XLS] Setup 2.4.257, WsBool 2.4.351, Header/Footer 2.4.136/124,
//! page breaks 2.4.142/343 and 2.5.160/276; ECMA-376 18.3.1.
use super::{f64_at, parse_biff_string, u16_at, unsupported, Record};
use crate::ooxml::xml_text;

#[derive(Default)]
pub(super) struct PrintSettings {
    setup: Option<Setup>,
    margins: [Option<f64>; 4],
    options: [Option<bool>; 4],
    fit_to_page: Option<bool>,
    header: Option<String>,
    footer: Option<String>,
    row_breaks: Option<Vec<(u16, u16, u16)>>,
    col_breaks: Option<Vec<(u16, u16, u16)>>,
}

struct Setup {
    flags: u16,
    paper: u16,
    scale: u16,
    page: i16,
    fit: (u16, u16),
    dpi: (u16, u16),
    header: f64,
    footer: f64,
    copies: u16,
}

impl PrintSettings {
    pub fn read(&mut self, record: &Record<'_>) -> Result<(), String> {
        let data = record.data;
        match record.kind {
            0x00a1 => {
                if data.len() != 34 {
                    return Err(unsupported("invalid BIFF Setup length"));
                }
                let setup = Setup {
                    flags: u16_at(data, 10)?,
                    paper: u16_at(data, 0)?,
                    scale: u16_at(data, 2)?,
                    page: u16_at(data, 4)? as i16,
                    fit: (u16_at(data, 6)?, u16_at(data, 8)?),
                    dpi: (u16_at(data, 12)?, u16_at(data, 14)?),
                    header: margin(data, 16, false)?,
                    footer: margin(data, 24, false)?,
                    copies: u16_at(data, 32)?,
                };
                if setup.fit.0 > 32767 || setup.fit.1 > 32767 {
                    return Err(unsupported("invalid BIFF fit-to-page limits"));
                }
                // OOXML firstPageNumber is unsigned. Do not silently wrap an
                // unsupported negative BIFF starting page into a huge number.
                if setup.flags & 128 != 0 && setup.page < 0 {
                    return Err(unsupported("negative BIFF starting page number"));
                }
                self.setup = Some(setup);
            }
            0x0026..=0x0029 => {
                if data.len() != 8 {
                    return Err(unsupported("invalid BIFF margin length"));
                }
                self.margins[usize::from(record.kind - 0x26)] = Some(margin(data, 0, true)?);
            }
            0x0083 | 0x0084 | 0x002a | 0x002b => {
                if data.len() != 2 || u16_at(data, 0)? > 1 {
                    return Err(unsupported("invalid BIFF print option"));
                }
                let index = match record.kind {
                    0x83 => 0,
                    0x84 => 1,
                    0x2a => 2,
                    _ => 3,
                };
                self.options[index] = Some(data[0] != 0);
            }
            0x0081 => self.fit_to_page = Some(u16_at(data, 0)? & 0x100 != 0),
            0x0014 | 0x0015 => {
                let value = if data.is_empty() {
                    String::new()
                } else {
                    if u16_at(data, 0)? > 255 {
                        return Err(unsupported(
                            "BIFF header/footer text exceeds 255 characters",
                        ));
                    }
                    let (text, consumed) = parse_biff_string(data)?;
                    if consumed != data.len() {
                        return Err(unsupported("unexpected BIFF header/footer tail"));
                    }
                    text
                };
                if record.kind == 0x14 {
                    self.header = Some(value);
                } else {
                    self.footer = Some(value);
                }
            }
            0x001b => self.row_breaks = Some(breaks(data, true)?),
            0x001a => self.col_breaks = Some(breaks(data, false)?),
            _ => {}
        }
        Ok(())
    }

    pub fn sheet_properties(&self) -> String {
        self.fit_to_page.map_or_else(String::new, |fit| {
            format!(
                "<sheetPr><pageSetUpPr fitToPage=\"{}\"/></sheetPr>",
                u8::from(fit)
            )
        })
    }

    pub fn incomplete_margins(&self) -> bool {
        (self.setup.is_some() || self.margins.iter().any(Option::is_some))
            && (self.setup.is_none() || self.margins.iter().any(Option::is_none))
    }

    pub fn xml(&self) -> String {
        let mut xml = String::new();
        if self.options.iter().any(Option::is_some) {
            xml.push_str("<printOptions");
            for (name, value) in [
                "horizontalCentered",
                "verticalCentered",
                "headings",
                "gridLines",
            ]
            .iter()
            .zip(self.options)
            {
                if let Some(value) = value {
                    xml.push_str(&format!(" {name}=\"{}\"", u8::from(value)));
                }
            }
            xml.push_str("/>");
        }
        if let Some(setup) = &self.setup {
            // All six attributes are required by CT_PageMargins. Do not invent
            // defaults for absent BIFF margins from one corpus or one printer.
            if let [Some(left), Some(right), Some(top), Some(bottom)] = self.margins {
                xml.push_str(&format!("<pageMargins left=\"{left}\" right=\"{right}\" top=\"{top}\" bottom=\"{bottom}\" header=\"{}\" footer=\"{}\"/>", setup.header, setup.footer));
            }
            xml.push_str(&setup.xml());
        }
        if self.header.is_some() || self.footer.is_some() {
            xml.push_str("<headerFooter>");
            for (name, value) in [("oddHeader", &self.header), ("oddFooter", &self.footer)] {
                if let Some(value) = value {
                    xml.push_str(&format!("<{name}>{}</{name}>", xml_text(value)));
                }
            }
            xml.push_str("</headerFooter>");
        }
        for (name, values) in [
            ("rowBreaks", &self.row_breaks),
            ("colBreaks", &self.col_breaks),
        ] {
            if let Some(values) = values {
                xml.push_str(&format!(
                    "<{name} count=\"{}\" manualBreakCount=\"{}\">",
                    values.len(),
                    values.len()
                ));
                for (id, min, max) in values {
                    xml.push_str(&format!(
                        "<brk id=\"{id}\" min=\"{min}\" max=\"{max}\" man=\"1\"/>"
                    ));
                }
                xml.push_str(&format!("</{name}>"));
            }
        }
        xml
    }
}

impl Setup {
    fn xml(&self) -> String {
        let mut xml = format!("<pageSetup fitToWidth=\"{}\" fitToHeight=\"{}\" pageOrder=\"{}\" blackAndWhite=\"{}\" draft=\"{}\" useFirstPageNumber=\"{}\" cellComments=\"{}\" errors=\"{}\"",
            self.fit.0, self.fit.1, if self.flags & 1 != 0 { "overThenDown" } else { "downThenOver" },
            (self.flags >> 3) & 1, (self.flags >> 4) & 1, (self.flags >> 7) & 1,
            if self.flags & 32 == 0 { "none" } else if self.flags & 512 != 0 { "atEnd" } else { "asDisplayed" },
            ["displayed", "blank", "dash", "NA"][usize::from((self.flags >> 10) & 3)]);
        if self.flags & 128 != 0 {
            xml.push_str(&format!(" firstPageNumber=\"{}\"", self.page));
        }
        // fNoPls is NOT an instruction to use arbitrary bytes as print settings.
        // Office writes uninitialized values here. [MS-XLS] 2.4.257 requires
        // ignoring paper, scale, resolutions, copies and orientation when set.
        if self.flags & 4 == 0 {
            xml.push_str(&format!(" paperSize=\"{}\" scale=\"{}\" horizontalDpi=\"{}\" verticalDpi=\"{}\" copies=\"{}\" orientation=\"{}\" usePrinterDefaults=\"0\"", self.paper, self.scale, self.dpi.0, self.dpi.1, self.copies,
                if self.flags & (2 | 64) != 0 { "portrait" } else { "landscape" }));
        }
        xml.push_str("/>");
        xml
    }
}

fn margin(data: &[u8], offset: usize, inclusive: bool) -> Result<f64, String> {
    let value = f64_at(data, offset)?;
    if !value.is_finite() || !(0.0..=49.0).contains(&value) || (!inclusive && value == 49.0) {
        return Err(unsupported("invalid BIFF print margin"));
    }
    Ok(value)
}

fn breaks(data: &[u8], horizontal: bool) -> Result<Vec<(u16, u16, u16)>, String> {
    let count = usize::from(u16_at(data, 0)?);
    if count > if horizontal { 1026 } else { 255 } || data.len() != 2 + count * 6 {
        return Err(unsupported("invalid BIFF page break count or length"));
    }
    let mut result = Vec::with_capacity(count);
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
        result.push(value);
    }
    Ok(result)
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
    fn undefined_printer_fields_never_leak_into_ooxml() {
        let mut value = PrintSettings::default();
        read(&mut value, 0xa1, &setup(4)).unwrap();
        let xml = value.xml();
        for name in [
            "paperSize",
            "scale",
            "orientation",
            "firstPageNumber",
            "horizontalDpi",
            "verticalDpi",
            "copies",
        ] {
            assert!(!xml.contains(name), "unexpected {name}: {xml}");
        }
        assert!(xml.contains("fitToHeight=\"0\""));
        assert!(value.incomplete_margins());
        assert!(!xml.contains("pageMargins"));
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
    fn keeps_full_span_manual_breaks_and_rejects_overlaps() {
        let mut value = PrintSettings::default();
        read(&mut value, 0x1a, &[1, 0, 2, 0, 0, 0, 255, 255]).unwrap();
        assert!(value
            .xml()
            .contains("id=\"2\" min=\"0\" max=\"65535\" man=\"1\""));
        assert!(read(
            &mut value,
            0x1b,
            &[2, 0, 5, 0, 0, 0, 2, 0, 5, 0, 2, 0, 3, 0]
        )
        .is_err());
        assert!(read(&mut value, 0x1a, &[0, 1]).is_err());
    }
    #[test]
    fn keeps_empty_headers_and_escapes_commands_without_executing_them() {
        let mut value = PrintSettings::default();
        read(&mut value, 0x14, &[]).unwrap();
        read(&mut value, 0x15, &[4, 0, 0, b'&', b'P', b'<', b'&']).unwrap();
        assert!(value
            .xml()
            .contains("<oddHeader></oddHeader><oddFooter>&amp;P&lt;&amp;</oddFooter>"));
        assert!(read(&mut value, 0x14, &[0, 1, 0]).is_err());
    }
}
