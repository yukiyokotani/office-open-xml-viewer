//! BIFF8 rich shared strings (MS-XLS 2.5.132 and 2.5.293).

use super::{styles::Styles, unsupported};
use crate::ooxml::xml_text;
use std::{collections::BTreeMap, rc::Rc};

pub(super) struct Text {
    pub text: String,
    // UTF-8 boundaries, resolved once from the source UTF-16 indices.
    runs: Vec<(usize, u16)>,
}

impl Text {
    pub fn new(units: &[u16], runs: &[(u16, u16)]) -> Result<Self, String> {
        let text = String::from_utf16_lossy(units);
        if runs.is_empty() {
            return Ok(Self {
                text,
                runs: Vec::new(),
            });
        }
        let mut offsets = vec![None; units.len() + 1];
        let mut unit = 0;
        for (byte, ch) in text.char_indices() {
            offsets[unit] = Some(byte);
            unit += ch.len_utf16();
        }
        offsets[units.len()] = Some(text.len());
        let mut result = Vec::with_capacity(runs.len());
        let mut previous = None;
        for &(start, font) in runs {
            if previous.is_some_and(|p| start <= p) {
                return Err(unsupported("BIFF format runs are not strictly ordered"));
            }
            previous = Some(start);
            let start = usize::from(start);
            let byte = offsets
                .get(start)
                .copied()
                .flatten()
                .ok_or_else(|| unsupported("BIFF format run splits or exceeds string"))?;
            // MS-XLS 2.5.132: ifnt at exactly the end is undefined/ignored.
            if start < units.len() {
                result.push((byte, font));
            }
        }
        Ok(Self { text, runs: result })
    }
}

/// Encode each SST entry once and share immutable fragments across cells.
/// The caller owns both this table and its FontIndex cache for one conversion.
pub(super) struct Encoder<'a, 'b> {
    styles: &'a Styles<'b>,
    fonts: BTreeMap<u16, String>,
    budget: usize,
}

impl<'a, 'b> Encoder<'a, 'b> {
    pub fn new(styles: &'a Styles<'b>) -> Self {
        Self {
            styles,
            fonts: BTreeMap::new(),
            budget: 256 * 1024 * 1024,
        }
    }

    pub fn encode(&mut self, string: Text) -> Result<Rc<str>, String> {
        let mut xml = String::new();
        if string.runs.is_empty() {
            append(
                &mut xml,
                &format!("<t xml:space=\"preserve\">{}</t>", xml_text(&string.text)),
                &mut self.budget,
            )?;
        } else {
            let first = string.runs[0].0;
            if first != 0 {
                append_run(&mut xml, "", &string.text[..first], &mut self.budget)?;
            }
            for (index, &(start, font)) in string.runs.iter().enumerate() {
                if let std::collections::btree_map::Entry::Vacant(entry) = self.fonts.entry(font) {
                    entry.insert(self.styles.run_font(font)?);
                }
                let end = string
                    .runs
                    .get(index + 1)
                    .map_or(string.text.len(), |r| r.0);
                append_run(
                    &mut xml,
                    &self.fonts[&font],
                    &string.text[start..end],
                    &mut self.budget,
                )?;
            }
        }
        Ok(Rc::from(xml))
    }
}

fn append_run(
    xml: &mut String,
    properties: &str,
    text: &str,
    budget: &mut usize,
) -> Result<(), String> {
    append(
        xml,
        &format!(
            "<r>{properties}<t xml:space=\"preserve\">{}</t></r>",
            xml_text(text)
        ),
        budget,
    )
}
fn append(xml: &mut String, part: &str, budget: &mut usize) -> Result<(), String> {
    *budget = budget
        .checked_sub(part.len())
        .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
    xml.push_str(part);
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::{cfb::test_support::build_cfb, convert_native, LegacyFormat};
    use std::io::{Cursor, Read};

    fn record(kind: u16, data: &[u8]) -> Vec<u8> {
        [
            kind.to_le_bytes().as_slice(),
            &(data.len() as u16).to_le_bytes(),
            data,
        ]
        .concat()
    }

    fn workbook(text: &str, runs: &[(u16, u16)]) -> Vec<u8> {
        let mut stream = record(0x0809, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(record(0x85, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
        for i in 0..5 {
            let mut font = vec![0; 16];
            font[..2].copy_from_slice(&(if i == 4 { 480u16 } else { 220u16 }).to_le_bytes());
            font[4..6].copy_from_slice(&(if i == 4 { 10u16 } else { 0x7fffu16 }).to_le_bytes());
            font[6..8].copy_from_slice(&(if i == 1 { 700u16 } else { 400u16 }).to_le_bytes());
            font[14] = 5;
            font.extend(b"Arial");
            stream.extend(record(0x31, &font));
        }
        let mut xf = [0u8; 20];
        xf[0] = 1; // Cell is bold; a normal rich run must explicitly reset it.
        stream.extend(record(0xe0, &xf));
        let units: Vec<u16> = text.encode_utf16().collect();
        let mut sst = [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        sst.extend((units.len() as u16).to_le_bytes());
        sst.push(9);
        sst.extend((runs.len() as u16).to_le_bytes());
        for unit in units {
            sst.extend(unit.to_le_bytes());
        }
        for (start, font) in runs {
            sst.extend(start.to_le_bytes());
            sst.extend(font.to_le_bytes());
        }
        stream.extend(record(0xfc, &sst));
        stream.extend(record(0x0a, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(record(0x0809, &[0, 6, 0x10, 0]));
        stream.extend(record(0xfd, &[0; 10]));
        stream.extend(record(0x0a, &[]));
        build_cfb(&[("Workbook", stream)])
    }

    #[test]
    fn emits_run_fonts_and_normal_reset_without_validating_the_end_sentinel_font() {
        let input = workbook("base RED normal", &[(5, 5), (9, 0), (15, u16::MAX)]);
        let output = convert_native(&input, LegacyFormat::Xls, 1_000_000).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(output.bytes)).unwrap();
        let mut xml = String::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains("<r><t xml:space=\"preserve\">base </t></r>"));
        assert!(xml.contains("<rFont val=\"Arial\"/>") && xml.contains("<sz val=\"24\"/>"));
        assert!(xml.contains("<color indexed=\"10\"/>") && xml.contains("<b val=\"0\"/>"));
        assert!(xml.contains(">RED </t></r>") && xml.contains(">normal</t></r>"));
    }

    #[test]
    fn rejects_live_invalid_fonts_unsorted_runs_and_surrogate_splits() {
        for runs in [&[(0, 4)][..], &[(0, 1023)], &[(2, 0), (1, 0)], &[(4, 0)]] {
            assert!(convert_native(&workbook("abc", runs), LegacyFormat::Xls, 1_000_000).is_err());
        }
        assert!(
            convert_native(&workbook("A😀B", &[(2, 0)]), LegacyFormat::Xls, 1_000_000).is_err()
        );
    }

    #[test]
    fn keeps_utf16_pairs_and_format_runs_across_separate_continuations() {
        let mut first = [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        first.extend([4, 0, 9, 2, 0]);
        first.extend([0x41, 0, 0x3d, 0xd8]);
        let second = [1, 0, 0xde, 0x42, 0, 1, 0];
        let third = [5, 0, 3, 0, 0, 0];
        let strings = super::super::parse_sst_elements(&[&first, &second, &third]).unwrap();
        assert_eq!(strings[0].text, "A😀B");
        assert_eq!(strings[0].runs, [(1, 5), (5, 0)]);
        assert!(super::super::parse_sst_elements(&[&first, &second, &third[..5]]).is_err());
    }

    #[test]
    fn rejects_split_fixed_headers_and_negative_extension_sizes() {
        let mut first = [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        first.extend([1, 0, 8]);
        assert!(super::super::parse_sst_elements(&[&first, &[0, 0, b'A']]).is_err());
        first.truncate(8);
        first.extend([1, 0, 4, 255, 255, 255, 255, b'A']);
        assert!(super::super::parse_sst_elements(&[&first]).is_err());
    }

    #[test]
    fn markup_budget_counts_escaping_and_repeated_cell_expansion() {
        let mut xml = String::new();
        assert!(super::append_run(&mut xml, "", "<&", &mut 10).is_err());
        assert!(xml.is_empty());
        let mut sheet = super::super::SheetData::default();
        let value: std::rc::Rc<str> = std::rc::Rc::from("<t>shared</t>");
        for row in 0..100 {
            sheet
                .rows
                .entry(row)
                .or_default()
                .insert(0, super::super::CellValue::SharedString(value.clone()));
        }
        assert_eq!(std::rc::Rc::strong_count(&value), 101);
        assert_eq!(
            super::super::build_sheet_xml(&sheet, 512).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert!(super::super::build_sheet_xml(&sheet, 20_000)
            .unwrap()
            .contains("<t>shared</t>"));
    }
}
