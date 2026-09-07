//! BIFF8 rich shared strings (MS-XLS 2.5.132 and 2.5.293).

use super::{styles::ResolvedStyleSheet, unsupported};
use crate::ooxml::xml_text;
use std::collections::BTreeMap;

pub(super) const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;

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

    /// Bytes retained by this neutral SST entry, excluding the owning Vec's
    /// `Text` slot. The parser charges this immediately after decoding each
    /// individually bounded BIFF string and before admitting the next entry.
    pub(super) fn retained_bytes(&self) -> Result<usize, String> {
        self.text
            .capacity()
            .checked_add(
                self.runs
                    .capacity()
                    .checked_mul(std::mem::size_of::<(usize, u16)>())
                    .ok_or_else(retained_budget_error)?,
            )
            .ok_or_else(retained_budget_error)
    }

    pub(super) fn validate_fonts(&self, styles: &ResolvedStyleSheet) -> Result<(), String> {
        for &(_, index) in &self.runs {
            styles.validate_run_font(index)?;
        }
        Ok(())
    }

    /// Materialize one SpreadsheetML string item for the byte-conversion route.
    /// The caller owns the route-wide output budget so repeated cell expansion
    /// remains bounded without retaining generated XML in the neutral SST.
    #[cfg(test)]
    pub(super) fn xml(
        &self,
        styles: &ResolvedStyleSheet,
        budget: &mut usize,
    ) -> Result<String, String> {
        XmlEncoder::new(styles, budget).encode(self)
    }

    /// Project one neutral BIFF SST entry directly to the renderer model.
    /// Every owned allocation is charged before reservation or cloning.
    #[allow(dead_code)] // Consumed by the direct XLS session in the next unit.
    pub(super) fn model(
        &self,
        styles: &ResolvedStyleSheet,
        budget: &mut usize,
    ) -> Result<xlsx_model::SharedString, String> {
        charge_model(budget, std::mem::size_of::<xlsx_model::SharedString>())?;
        let text = clone_bounded(&self.text, budget)?;
        let runs = if self.runs.is_empty() {
            None
        } else {
            let count = self.runs.len() + usize::from(self.runs[0].0 != 0);
            charge_model(
                budget,
                count
                    .checked_mul(std::mem::size_of::<xlsx_model::Run>())
                    .ok_or_else(model_budget_error)?,
            )?;
            let mut output = Vec::new();
            output
                .try_reserve_exact(count)
                .map_err(|_| model_budget_error())?;
            let first = self.runs[0].0;
            if first != 0 {
                output.push(xlsx_model::Run {
                    text: clone_bounded(&self.text[..first], budget)?,
                    font: None,
                });
            }
            for (index, &(start, font)) in self.runs.iter().enumerate() {
                let end = self
                    .runs
                    .get(index + 1)
                    .map_or(self.text.len(), |run| run.0);
                output.push(xlsx_model::Run {
                    text: clone_bounded(&self.text[start..end], budget)?,
                    font: Some(styles.run_font_model(font, budget)?),
                });
            }
            Some(output)
        };
        Ok(xlsx_model::SharedString {
            text,
            runs,
            phonetic_runs: Vec::new(),
            phonetic_pr: None,
        })
    }
}

/// Byte-route-only adapter. Its font cache exists only while `finish` creates
/// the unique SST XML fragments; neutral preparation and direct projection do
/// not retain OOXML strings.
pub(super) struct XmlEncoder<'a, 'b> {
    styles: &'a ResolvedStyleSheet,
    fonts: BTreeMap<u16, String>,
    budget: &'b mut usize,
}

impl<'a, 'b> XmlEncoder<'a, 'b> {
    pub(super) fn new(styles: &'a ResolvedStyleSheet, budget: &'b mut usize) -> Self {
        Self {
            styles,
            fonts: BTreeMap::new(),
            budget,
        }
    }

    pub(super) fn encode(&mut self, value: &Text) -> Result<String, String> {
        let mut xml = String::new();
        if value.runs.is_empty() {
            append(
                &mut xml,
                &format!("<t xml:space=\"preserve\">{}</t>", xml_text(&value.text)),
                self.budget,
            )?;
        } else {
            let first = value.runs[0].0;
            if first != 0 {
                append_run(&mut xml, "", &value.text[..first], self.budget)?;
            }
            for (index, &(start, font)) in value.runs.iter().enumerate() {
                if !self.fonts.contains_key(&font) {
                    self.fonts.insert(font, self.styles.run_font_xml(font)?);
                }
                let end = value
                    .runs
                    .get(index + 1)
                    .map_or(value.text.len(), |run| run.0);
                append_run(
                    &mut xml,
                    &self.fonts[&font],
                    &value.text[start..end],
                    self.budget,
                )?;
            }
        }
        Ok(xml)
    }
}

fn retained_budget_error() -> String {
    unsupported("BIFF retained shared string byte budget exceeded")
}

#[allow(dead_code)] // Used by the direct-model adapter above.
fn model_budget_error() -> String {
    unsupported("BIFF shared string model byte budget exceeded")
}

#[allow(dead_code)] // Used by the direct-model adapter above.
fn charge_model(budget: &mut usize, bytes: usize) -> Result<(), String> {
    *budget = budget.checked_sub(bytes).ok_or_else(model_budget_error)?;
    Ok(())
}

#[allow(dead_code)] // Used by the direct-model adapter above.
fn clone_bounded(value: &str, budget: &mut usize) -> Result<String, String> {
    charge_model(budget, value.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| model_budget_error())?;
    output.push_str(value);
    Ok(output)
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
    use crate::{
        cfb::{test_support::build_cfb, CompoundFile},
        convert_native, LegacyFormat,
    };
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
    fn direct_model_matches_the_parser_visible_rich_run_semantics() {
        let input = workbook("base RED normal", &[(5, 5), (9, 0), (15, u16::MAX)]);
        let cfb = CompoundFile::open(&input).unwrap();
        let prepared = super::super::prepare(&cfb, false).unwrap();
        let mut budget = super::MAX_MODEL_BYTES;
        let value = prepared.shared_strings[0]
            .model(&prepared.styles, &mut budget)
            .unwrap();
        let required = super::MAX_MODEL_BYTES - budget;
        assert_eq!(
            required,
            std::mem::size_of::<xlsx_model::SharedString>()
                + 3 * std::mem::size_of::<xlsx_model::Run>()
                + 2 * "base RED normal".len()
                + 2 * "Arial".len()
                + "#FF0000".len()
        );
        assert_eq!(value.text, "base RED normal");
        let runs = value.runs.unwrap();
        assert_eq!(runs.len(), 3);
        assert_eq!(runs[0].text, "base ");
        assert!(runs[0].font.is_none());
        let red = runs[1].font.as_ref().unwrap();
        assert_eq!(runs[1].text, "RED ");
        assert_eq!(
            (red.bold, red.italic, red.underline, red.strike),
            (false, false, false, false)
        );
        assert_eq!((red.size, red.name.as_deref()), (Some(24.0), Some("Arial")));
        assert_eq!(red.color.as_deref(), Some("#FF0000"));
        assert_eq!(
            (red.underline_style.as_deref(), red.vert_align.as_deref()),
            (None, None)
        );
        let normal = runs[2].font.as_ref().unwrap();
        assert_eq!(runs[2].text, "normal");
        assert_eq!(
            (normal.bold, normal.italic, normal.underline, normal.strike),
            (false, false, false, false)
        );
        assert_eq!(normal.vert_align, None);
        let mut exact = required;
        prepared.shared_strings[0]
            .model(&prepared.styles, &mut exact)
            .unwrap();
        assert_eq!(exact, 0);
        assert!(prepared.shared_strings[0]
            .model(&prepared.styles, &mut (required - 1))
            .is_err());

        let converted = convert_native(&input, LegacyFormat::Xls, 1_000_000).unwrap();
        let parsed: serde_json::Value = serde_json::from_str(
            &xlsx_parser::parse_sheet_native(&converted.bytes, 0, "S").unwrap(),
        )
        .unwrap();
        let mut parsed_value = parsed["rows"][0]["cells"][0]["value"].clone();
        assert_eq!(
            parsed_value
                .as_object_mut()
                .unwrap()
                .remove("type")
                .unwrap(),
            "text"
        );
        let mut comparison_budget = super::MAX_MODEL_BYTES;
        assert_eq!(
            serde_json::to_value(
                &prepared.shared_strings[0]
                    .model(&prepared.styles, &mut comparison_budget)
                    .unwrap()
            )
            .unwrap(),
            parsed_value
        );
    }

    #[test]
    fn native_model_preserves_biff_controls_and_line_endings_without_xml_normalization() {
        let source = "a\rb\r\nc\u{1}d";
        let text = super::Text::new(&source.encode_utf16().collect::<Vec<_>>(), &[]).unwrap();
        let styles = super::super::styles::minimal_resolved();
        let mut budget = super::MAX_MODEL_BYTES;
        assert_eq!(text.model(&styles, &mut budget).unwrap().text, source);

        let mut xml_budget = usize::MAX;
        let xml = text.xml(&styles, &mut xml_budget).unwrap();
        assert!(xml.contains("a\rb\r\nc�d"));
    }

    #[test]
    fn plain_model_charges_before_allocating_at_the_exact_boundary() {
        let text = super::Text::new(&"abc".encode_utf16().collect::<Vec<_>>(), &[]).unwrap();
        let styles = super::super::styles::minimal_resolved();
        let required = std::mem::size_of::<xlsx_model::SharedString>() + 3;
        let mut exact = required;
        assert_eq!(text.model(&styles, &mut exact).unwrap().text, "abc");
        assert_eq!(exact, 0);
        assert!(text.model(&styles, &mut (required - 1)).is_err());
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
    fn synthetic_minimal_style_font_is_not_a_live_biff_run_font() {
        let text = super::Text::new(&['x' as u16], &[(0, 0)]).unwrap();
        assert!(text
            .validate_fonts(&super::super::styles::minimal_resolved())
            .is_err());
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
        let value = super::Text::new(&"shared".encode_utf16().collect::<Vec<_>>(), &[]).unwrap();
        for row in 0..100 {
            sheet
                .rows
                .entry(row)
                .or_default()
                .insert(0, super::super::CellValue::SharedString(0));
        }
        let styles = super::super::styles::minimal_resolved();
        let mut budget = 256 * 1024 * 1024;
        let strings = [value.xml(&styles, &mut budget).unwrap()];
        assert_eq!(
            super::super::build_sheet_xml_with_drawings(&sheet, &strings, 512, None, false)
                .unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert!(
            super::super::build_sheet_xml_with_drawings(&sheet, &strings, 20_000, None, false)
                .unwrap()
                .contains("<t xml:space=\"preserve\">shared</t>")
        );
    }
}
