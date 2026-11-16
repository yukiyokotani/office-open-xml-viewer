//! BIFF8 rich shared strings (MS-XLS 2.5.132 and 2.5.293) and their
//! phonetic strings (2.5.109 ExtRst).

use super::{styles::ResolvedStyleSheet, unsupported};

pub(super) const MAX_MODEL_BYTES: usize = 256 * 1024 * 1024;

pub(super) struct Text {
    pub text: String,
    // UTF-8 boundaries, resolved once from the source UTF-16 indices.
    runs: Vec<(usize, u16)>,
    /// The phonetic string, read only for a workbook that shows one.
    phonetic: Option<Box<Phonetic>>,
}

/// MS-XLS 2.5.109 ExtRst: the phonetic string of a shared string, as
/// ECMA-376 18.4.3 `phoneticPr` and 18.4.6 `rPh` state it.
struct Phonetic {
    /// 2.5.201 Phs ifnt (a FontIndex), phType and alcH.
    font: u16,
    kind: u8,
    alignment: u8,
    /// Base-text range in Unicode scalar offsets (start, end) and the
    /// phonetic text shown over it.
    runs: Vec<(u32, u32, String)>,
}

impl Phonetic {
    /// `None` when the ExtRst holds no readable phonetic run. `units` is
    /// the base string.
    fn parse(ext: &[u8], units: &[u16]) -> Option<Self> {
        let word = |at: usize| -> Option<u16> {
            Some(u16::from_le_bytes([*ext.get(at)?, *ext.get(at + 1)?]))
        };
        // reserved (ignored), cb, then phs, rphssub and rgphruns.
        let size = usize::from(word(2)?);
        let data = ext.get(4..4 + size)?;
        let word = |at: usize| -> Option<u16> {
            Some(u16::from_le_bytes([*data.get(at)?, *data.get(at + 1)?]))
        };
        let font = word(0)?;
        let flags = word(2)?;
        // RPHSSub: crun, cch, st (an LPWideString: its own count, then
        // UTF-16 characters).
        let count = usize::from(word(4)?);
        let chars = usize::from(word(6)?);
        if usize::from(word(8)?) != chars || count > 0x7fff || chars > 0x7fff {
            return None;
        }
        let text: Vec<u16> = (0..chars)
            .map(|index| word(10 + index * 2))
            .collect::<Option<_>>()?;
        let runs_at = 10 + chars * 2;
        // PhRuns (2.5.200): ichFirst, ichMom, cchMom. "If crun is zero,
        // there is one phonetic text run": the whole string over the whole
        // base text.
        let mut raw = Vec::with_capacity(count.max(1));
        if count == 0 {
            raw.push((0usize, 0usize, units.len()));
        }
        for index in 0..count {
            let at = runs_at + index * 6;
            raw.push((
                usize::from(word(at)?),
                usize::from(word(at + 2)?),
                usize::from(word(at + 4)?),
            ));
        }
        // Scalar offset of each UTF-16 unit of the base text; a position
        // inside a surrogate pair belongs to that pair.
        let mut scalars = Vec::with_capacity(units.len() + 1);
        let mut scalar = 0u32;
        let mut high = false;
        for &unit in units {
            let low = high && (0xdc00..0xe000).contains(&unit);
            if low {
                scalars.push(scalar - 1);
            } else {
                scalars.push(scalar);
                scalar += 1;
            }
            high = !low && (0xd800..0xdc00).contains(&unit);
        }
        scalars.push(scalar);
        let mut runs = Vec::with_capacity(raw.len());
        for (index, &(first, base, length)) in raw.iter().enumerate() {
            let next = raw.get(index + 1).map_or(chars, |run| run.0);
            let end = base.checked_add(length)?;
            // Each run starts after the one before it, in both strings.
            if first > next || next > chars || end > units.len() {
                return None;
            }
            let (start, end) = (scalars[base], scalars[end]);
            if start >= end || first == next {
                continue;
            }
            runs.push((start, end, String::from_utf16_lossy(&text[first..next])));
        }
        (!runs.is_empty()).then_some(Self {
            font,
            kind: (flags & 3) as u8,
            alignment: ((flags >> 2) & 3) as u8,
            runs,
        })
    }
}

impl Text {
    pub fn new(units: &[u16], runs: &[(u16, u16)]) -> Result<Self, String> {
        let text = String::from_utf16_lossy(units);
        if runs.is_empty() {
            return Ok(Self {
                text,
                runs: Vec::new(),
                phonetic: None,
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
        Ok(Self {
            text,
            runs: result,
            phonetic: None,
        })
    }

    /// Attach the phonetic string of this string's ExtRst (`units` is the
    /// base string `new` was built from). An ExtRst is framed by the
    /// string's cbExtRst, so one whose content cannot be read leaves the
    /// string without phonetic text.
    pub(super) fn with_phonetic(mut self, units: &[u16], ext: &[u8]) -> Self {
        match Phonetic::parse(ext, units) {
            Some(phonetic) => self.phonetic = Some(Box::new(phonetic)),
            None => crate::rejection::ignored_value("XLS phonetic guides are not projected"),
        }
        self
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
            .and_then(|bytes| {
                let phonetic = self.phonetic.as_ref().map_or(0, |phonetic| {
                    std::mem::size_of::<Phonetic>()
                        + phonetic
                            .runs
                            .iter()
                            .map(|run| std::mem::size_of_val(run) + run.2.capacity())
                            .sum::<usize>()
                });
                bytes.checked_add(phonetic)
            })
            .ok_or_else(retained_budget_error)
    }

    pub(super) fn validate_fonts(&self, styles: &ResolvedStyleSheet) -> Result<(), String> {
        for &(_, index) in &self.runs {
            styles.validate_run_font(index)?;
        }
        Ok(())
    }

    /// Project one neutral BIFF SST entry directly to the renderer model.
    /// Every owned allocation is charged before reservation or cloning; the
    /// owning vector has already charged and reserved this SharedString slot.
    pub(super) fn model_into_reserved_slot(
        &self,
        styles: &ResolvedStyleSheet,
        budget: &mut usize,
    ) -> Result<xlsx_model::SharedString, String> {
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
        let (phonetic_runs, phonetic_pr) = match &self.phonetic {
            Some(phonetic) => {
                charge_model(
                    budget,
                    phonetic
                        .runs
                        .len()
                        .checked_mul(std::mem::size_of::<xlsx_model::PhoneticRun>())
                        .ok_or_else(model_budget_error)?,
                )?;
                let mut output = Vec::new();
                output
                    .try_reserve_exact(phonetic.runs.len())
                    .map_err(|_| model_budget_error())?;
                for (start, end, text) in &phonetic.runs {
                    output.push(xlsx_model::PhoneticRun {
                        sb: *start,
                        eb: *end,
                        text: clone_bounded(text, budget)?,
                    });
                }
                // The schema defaults (fullwidthKatakana, left) are left
                // unstated, as Excel's own SpreadsheetML leaves them.
                let kind = match phonetic.kind {
                    0 => Some("halfwidthKatakana"),
                    1 => None,
                    2 => Some("Hiragana"),
                    _ => Some("noConversion"),
                };
                let alignment = match phonetic.alignment {
                    0 => Some("noControl"),
                    1 => None,
                    2 => Some("center"),
                    _ => Some("distributed"),
                };
                charge_model(
                    budget,
                    kind.map_or(0, str::len) + alignment.map_or(0, str::len),
                )?;
                (
                    output,
                    Some(xlsx_model::PhoneticProperties {
                        font_id: styles.phonetic_font(phonetic.font),
                        r#type: kind.map(str::to_owned),
                        alignment: alignment.map(str::to_owned),
                    }),
                )
            }
            None => (Vec::new(), None),
        };
        Ok(xlsx_model::SharedString {
            text,
            runs,
            phonetic_runs,
            phonetic_pr,
        })
    }
}

fn retained_budget_error() -> String {
    crate::rejection::resource("BIFF retained shared string byte budget exceeded")
}

fn model_budget_error() -> String {
    crate::rejection::resource("BIFF shared string model byte budget exceeded")
}

fn charge_model(budget: &mut usize, bytes: usize) -> Result<(), String> {
    *budget = budget.checked_sub(bytes).ok_or_else(model_budget_error)?;
    Ok(())
}

fn clone_bounded(value: &str, budget: &mut usize) -> Result<String, String> {
    charge_model(budget, value.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| model_budget_error())?;
    output.push_str(value);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};
    use crate::xls::direct::DirectSession;

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
        build_scoped_cfb(&[("Workbook", stream)])
    }

    /// ExtRst (2.5.109) with Phs `font` and `flags`, the phonetic string
    /// `reading` and these PhRuns (ichFirst, ichMom, cchMom).
    fn ext_rst(font: u16, flags: u16, reading: &str, runs: &[(u16, u16, u16)]) -> Vec<u8> {
        let units: Vec<u16> = reading.encode_utf16().collect();
        let mut data = font.to_le_bytes().to_vec();
        data.extend(flags.to_le_bytes());
        data.extend((runs.len() as u16).to_le_bytes());
        data.extend((units.len() as u16).to_le_bytes());
        data.extend((units.len() as u16).to_le_bytes());
        for unit in units {
            data.extend(unit.to_le_bytes());
        }
        for (first, base, length) in runs {
            for value in [first, base, length] {
                data.extend(value.to_le_bytes());
            }
        }
        let mut ext = 1u16.to_le_bytes().to_vec();
        ext.extend((data.len() as u16).to_le_bytes());
        ext.extend(data);
        ext
    }

    /// A workbook whose only shared string `text` (in A1 and C1) has the
    /// ExtRst `ext`; `ranges` are the sheet's PhoneticInfo ranges.
    fn phonetic_workbook(text: &str, ext: &[u8], ranges: &[[u16; 4]]) -> Vec<u8> {
        let mut stream = record(0x0809, &[0, 6, 5, 0]);
        let bound = stream.len() + 4;
        stream.extend(record(0x85, &[0, 0, 0, 0, 0, 0, 1, 0, b'S']));
        for _ in 0..6 {
            let mut font = vec![0; 16];
            font[..2].copy_from_slice(&220u16.to_le_bytes());
            font[4..6].copy_from_slice(&0x7fffu16.to_le_bytes());
            font[6..8].copy_from_slice(&400u16.to_le_bytes());
            font[14] = 5;
            font.extend(b"Arial");
            stream.extend(record(0x31, &font));
        }
        stream.extend(record(0xe0, &[0u8; 20]));
        let units: Vec<u16> = text.encode_utf16().collect();
        let mut sst = [2u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        sst.extend((units.len() as u16).to_le_bytes());
        // fHighByte and fExtSt.
        sst.push(0x05);
        sst.extend((ext.len() as u32).to_le_bytes());
        for unit in units {
            sst.extend(unit.to_le_bytes());
        }
        sst.extend(ext);
        stream.extend(record(0xfc, &sst));
        stream.extend(record(0x0a, &[]));
        let offset = stream.len() as u32;
        stream[bound..bound + 4].copy_from_slice(&offset.to_le_bytes());
        stream.extend(record(0x0809, &[0, 6, 0x10, 0]));
        stream.extend(record(0xfd, &[0, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
        stream.extend(record(0xfd, &[0, 0, 2, 0, 0, 0, 0, 0, 0, 0]));
        let mut info = vec![0x06, 0x00, 0x35, 0x00];
        info.extend((ranges.len() as u16).to_le_bytes());
        for range in ranges {
            for value in range {
                info.extend(value.to_le_bytes());
            }
        }
        stream.extend(record(0xef, &info));
        stream.extend(record(0x0a, &[]));
        build_scoped_cfb(&[("Workbook", stream)])
    }

    /// The bootstrap's shared string and the (column, showPhonetic) of the
    /// first row's cells.
    fn phonetic_projection(input: &[u8]) -> (serde_json::Value, Vec<(u32, bool)>) {
        let cfb = CompoundFile::open(input).unwrap();
        let mut session = DirectSession::new(&cfb).unwrap();
        let bootstrap = session.bootstrap().unwrap();
        let string = serde_json::to_value(&bootstrap.shared_strings[0]).unwrap();
        let sheet = session.projected_sheet(0, "S").unwrap();
        let cells = sheet.rows[0]
            .cells
            .iter()
            .map(|cell| (cell.col, cell.show_phonetic))
            .collect();
        (string, cells)
    }

    #[test]
    fn phonetic_strings_are_shown_in_the_cells_of_the_phonetic_ranges() {
        use crate::rejection::{take_best_effort, BestEffort};
        take_best_effort();
        // Wide katakana (phType 1), left aligned (alcH 1), font index 6:
        // "カンジ" over the first two characters, "テ" over the last.
        let ext = ext_rst(6, 0x0005, "カンジテ", &[(0, 0, 2), (3, 4, 1)]);
        // A1 (row 0, column 0) shows its phonetic string; C1 does not.
        let (string, cells) =
            phonetic_projection(&phonetic_workbook("漢字テスト", &ext, &[[0, 0, 0, 0]]));
        assert_eq!(
            string,
            serde_json::json!({
                "text": "漢字テスト",
                "phoneticRuns": [
                    {"sb": 0, "eb": 2, "text": "カンジ"},
                    {"sb": 4, "eb": 5, "text": "テ"}
                ],
                // FontIndex 6 is the sixth font (index 4 is not used).
                "phoneticPr": {"fontId": 5}
            })
        );
        assert_eq!(cells, [(1, true), (3, false)]);
        // Hiragana (2), distributed (3); a range over both cells.
        let ext = ext_rst(0, 0x000e, "かんじ", &[(0, 0, 2)]);
        let (string, cells) = phonetic_projection(&phonetic_workbook(
            "漢字",
            &ext,
            &[[5, 9, 0, 0], [0, 3, 1, 4]],
        ));
        assert_eq!(
            string["phoneticPr"],
            serde_json::json!({"fontId": 0, "type": "Hiragana", "alignment": "distributed"})
        );
        assert_eq!(cells, [(1, false), (3, true)]);
        // crun 0: the whole string over the whole base text. A base range
        // inside a surrogate pair covers that pair.
        let ext = ext_rst(0, 0x0005, "ヨミ", &[]);
        let (string, _) =
            phonetic_projection(&phonetic_workbook("読\u{20bb7}", &ext, &[[0, 0, 0, 0]]));
        assert_eq!(
            string["phoneticRuns"],
            serde_json::json!([{"sb": 0, "eb": 2, "text": "ヨミ"}])
        );
        assert!(take_best_effort().is_empty());
        // A workbook without a phonetic range reads no phonetic string.
        let ext = ext_rst(6, 0x0005, "カンジ", &[(0, 0, 2)]);
        let (string, cells) = phonetic_projection(&phonetic_workbook("漢字", &ext, &[]));
        assert_eq!(string, serde_json::json!({"text": "漢字"}));
        assert_eq!(cells, [(1, false), (3, false)]);
        assert!(take_best_effort().is_empty());
        // An ExtRst whose runs leave the base text shows no phonetic string;
        // the cell still shows its text.
        let ext = ext_rst(6, 0x0005, "カンジ", &[(0, 1, 9)]);
        let (string, cells) =
            phonetic_projection(&phonetic_workbook("漢字", &ext, &[[0, 0, 0, 0]]));
        assert_eq!(string, serde_json::json!({"text": "漢字"}));
        assert_eq!(cells, [(1, true), (3, false)]);
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::IgnoredValue,
                "XLS phonetic guides are not projected"
            )]
        );
        // A range list the record does not hold is malformed.
        let mut input = phonetic_workbook("漢字", &ext, &[[0, 0, 0, 0]]);
        let at = input
            .windows(4)
            .position(|window| window == [0xef, 0x00, 0x0e, 0x00])
            .unwrap();
        input[at + 8] = 2;
        let cfb = CompoundFile::open(&input).unwrap();
        assert!(DirectSession::new(&cfb)
            .err()
            .expect("rejected")
            .contains("invalid BIFF phonetic ranges"));
    }

    /// Standalone projection charge: the owning vector's slot plus payload.
    fn model(
        text: &super::Text,
        styles: &super::ResolvedStyleSheet,
        budget: &mut usize,
    ) -> Result<xlsx_model::SharedString, String> {
        super::charge_model(budget, std::mem::size_of::<xlsx_model::SharedString>())?;
        text.model_into_reserved_slot(styles, budget)
    }

    #[test]
    fn direct_model_projects_run_fonts_and_resets_without_the_end_sentinel_font() {
        let input = workbook("base RED normal", &[(5, 5), (9, 0), (15, u16::MAX)]);
        let cfb = CompoundFile::open(&input).unwrap();
        let prepared = super::super::prepare(&cfb).unwrap();
        let mut budget = super::MAX_MODEL_BYTES;
        let value = model(&prepared.shared_strings[0], &prepared.styles, &mut budget).unwrap();
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
        let runs = value.runs.clone().unwrap();
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
        // The normal run explicitly resets the cell's bold XF font.
        let normal = runs[2].font.as_ref().unwrap();
        assert_eq!(runs[2].text, "normal");
        assert_eq!(
            (normal.bold, normal.italic, normal.underline, normal.strike),
            (false, false, false, false)
        );
        assert_eq!(normal.vert_align, None);
        let mut exact = required;
        model(&prepared.shared_strings[0], &prepared.styles, &mut exact).unwrap();
        assert_eq!(exact, 0);
        assert!(model(
            &prepared.shared_strings[0],
            &prepared.styles,
            &mut (required - 1)
        )
        .is_err());

        // The session bootstrap carries the same projection.
        let mut session = DirectSession::new(&cfb).unwrap();
        let bootstrap = session.bootstrap().unwrap();
        assert_eq!(
            serde_json::to_value(&bootstrap.shared_strings[0]).unwrap(),
            serde_json::to_value(value).unwrap()
        );
    }

    #[test]
    fn native_model_preserves_biff_controls_and_line_endings() {
        let source = "a\rb\r\nc\u{1}d";
        let text = super::Text::new(&source.encode_utf16().collect::<Vec<_>>(), &[]).unwrap();
        let styles = super::super::styles::minimal_resolved();
        let mut budget = super::MAX_MODEL_BYTES;
        assert_eq!(model(&text, &styles, &mut budget).unwrap().text, source);
    }

    #[test]
    fn plain_model_charges_before_allocating_at_the_exact_boundary() {
        let text = super::Text::new(&"abc".encode_utf16().collect::<Vec<_>>(), &[]).unwrap();
        let styles = super::super::styles::minimal_resolved();
        let required = 3;
        let mut exact = required;
        assert_eq!(
            text.model_into_reserved_slot(&styles, &mut exact)
                .unwrap()
                .text,
            "abc"
        );
        assert_eq!(exact, 0);
        assert!(text
            .model_into_reserved_slot(&styles, &mut (required - 1))
            .is_err());
    }

    #[test]
    fn rejects_live_invalid_fonts_unsorted_runs_and_surrogate_splits() {
        for (text, runs, expected) in [
            ("abc", &[(0, 4)][..], "font index out of range"),
            ("abc", &[(0, 1023)], "font index out of range"),
            ("abc", &[(2, 0), (1, 0)], "not strictly ordered"),
            ("abc", &[(4, 0)], "splits or exceeds string"),
            ("A😀B", &[(2, 0)], "splits or exceeds string"),
        ] {
            let input = workbook(text, runs);
            let cfb = CompoundFile::open(&input).unwrap();
            let error = DirectSession::new(&cfb).err().expect("rejected run");
            assert!(error.contains(expected), "{expected}: {error}");
        }
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
        let strings = super::super::parse_sst_elements(&[&first, &second, &third], false).unwrap();
        assert_eq!(strings[0].text, "A😀B");
        assert_eq!(strings[0].runs, [(1, 5), (5, 0)]);
        assert!(super::super::parse_sst_elements(&[&first, &second, &third[..5]], false).is_err());
    }

    #[test]
    fn rejects_split_fixed_headers_and_negative_extension_sizes() {
        let mut first = [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat();
        first.extend([1, 0, 8]);
        assert!(super::super::parse_sst_elements(&[&first, &[0, 0, b'A']], false).is_err());
        first.truncate(8);
        first.extend([1, 0, 4, 255, 255, 255, 255, b'A']);
        assert!(super::super::parse_sst_elements(&[&first], false).is_err());
    }
}
