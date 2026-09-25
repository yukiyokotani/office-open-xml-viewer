//! MS-DOC 2.5.6 FibRgFcLcb97.fcDop/lcbDop and 2.7.1-3 Dop/DopBase.
//! All supported DOP versions begin with the same 84-byte DopBase.
use super::{u16_at, u32_at, unsupported};

#[derive(Debug, PartialEq, Eq)]
pub(super) struct Properties {
    pub default_tab_twips: u16,
    pub even_and_odd_headers: bool,
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub notes: NoteProperties,
    /// ECMA-376 Part 1 17.15.3.1 adjustLineHeightInTable, the inverse of
    /// MS-DOC 2.7.13 Copts.fDontAdjustLineHeightInTable.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub adjust_line_height_in_table: bool,
    /// ECMA-376 Part 1 17.15.3.3 balanceSingleByteDoubleByteWidth, the
    /// inverse of MS-DOC 2.7.11 Copts60.fDntBlnSbDbWid.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub balance_single_byte_double_byte_width: bool,
    /// ECMA-376 Part 1 17.15.1.18 characterSpacingControl from MS-DOC 2.7.16
    /// DopTypography.iJustification; `None` when the DOP predates Dop97.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub character_spacing_control: Option<&'static str>,
    /// MS-DOC 2.7.2 DopBase fRMView / fRMPrint.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub revision_markup: RevisionMarkup,
}

/// MS-DOC 2.7.2 DopBase: fRMView "whether to show any revision markup that is
/// present in this document" and fRMPrint "whether to print" it (note <166>:
/// they can differ). A Word-exported PDF of a document with both set shows
/// insertions underlined with margin change bars.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
pub(crate) struct RevisionMarkup {
    pub(crate) on_screen: bool,
    pub(crate) in_print: bool,
    /// Whether the projected model carries any revision mark at all; set by
    /// the direct model, never read from the DOP.
    pub(crate) has_marks: bool,
}

/// MS-DOC 2.7.2 DopBase fpc/rncFtn/nFtn/rncEdn/nEdn/epc and 2.7.4 Dop97
/// nfcFtnRef/nfcEdnRef, retained raw. Except epc, MS-DOC scopes them to
/// documents whose nFib is at most 0x00D9; later documents use section
/// properties. `formats` is `None` when the DOP predates Dop97.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
pub(super) struct NoteProperties {
    pub footnote_position: u8,
    pub footnote_restart: u8,
    pub footnote_start: u16,
    pub endnote_restart: u8,
    pub endnote_start: u16,
    pub endnote_position: u8,
    pub formats: Option<(u16, u16)>,
}

pub(super) fn read(word: &[u8], table: &[u8]) -> Result<Option<Properties>, String> {
    let size = u32_at(word, 0x196)? as usize;
    // Recovery policy for incomplete documents and minimal synthetic fixtures:
    // the normative lcbDop is nonzero. The caller warns and retains the OOXML
    // default when the entire DOP is absent; no offset is dereferenced then.
    if size == 0 {
        return Ok(None);
    }
    let offset = u32_at(word, 0x192)? as usize;
    let dop = table
        .get(offset..)
        .and_then(|tail| tail.get(..size))
        .filter(|dop| dop.len() >= 84)
        .ok_or_else(|| unsupported("truncated Word document properties"))?;
    // Two four-byte flag words and copts60 (two bytes) precede dxaTab.
    let interval = u16_at(dop, 10)?;
    // ECMA-376 17.15.1.25 requires a positive interval. Do not silently invent
    // spacing for a zero interval or allow a non-progressing automatic-tab loop.
    if interval == 0 {
        return Err(unsupported("zero Word default tab interval"));
    }
    let footnotes = u16_at(dop, 2)?;
    let endnotes = u16_at(dop, 52)?;
    let notes = NoteProperties {
        footnote_position: (dop[0] >> 5) & 3,
        footnote_restart: (footnotes & 3) as u8,
        footnote_start: footnotes >> 2,
        endnote_restart: (endnotes & 3) as u8,
        endnote_start: endnotes >> 2,
        endnote_position: (u16_at(dop, 54)? & 3) as u8,
        // Dop97 (500 bytes) ends with nfcFtnRef, nfcEdnRef and two ignored
        // display values; every later Dop embeds Dop97 at its start.
        formats: if dop.len() >= 500 {
            Some((u16_at(dop, 492)?, u16_at(dop, 494)?))
        } else {
            None
        },
    };
    // MS-DOC 2.7.5 Dop2000 (544 bytes) places its 32-byte Copts after the
    // 500-byte Dop97 and eight bytes of Dop2000 fields; 2.7.13 Copts begins
    // with a four-byte Copts80, followed by the flag word whose fourth bit is
    // fDontAdjustLineHeightInTable. 2.6.4 sprmSDyaLinePitch excludes table
    // lines only "in case the fDontAdjustLineHeightInTable flag is set in the
    // document Dop2000"; an older DOP has no such flag, so it is not set.
    let adjust_line_height_in_table = if dop.len() >= 544 {
        u32_at(dop, 512)? & (1 << 3) == 0
    } else {
        true
    };
    // MS-DOC 2.7.2 DopBase.copts60 (offset 8), 2.7.11 bit P.
    let balance_single_byte_double_byte_width = u16_at(dop, 8)? & (1 << 15) == 0;
    // MS-DOC 2.7.4 Dop97: dop95 (88 bytes) and adt (2 bytes) precede the
    // DopTypography, whose bits 1-2 are iJustification.
    let character_spacing_control = if dop.len() >= 500 {
        Some(match (u16_at(dop, 90)? >> 1) & 3 {
            0 => "doNotCompress",
            1 => "compressPunctuation",
            2 => "compressPunctuationAndJapaneseKana",
            _ => return Err(unsupported("invalid Word character spacing control")),
        })
    } else {
        None
    };
    // DopBase second flag word (offset 4): h = fRMView (bit 27), i = fRMPrint
    // (bit 28), counting from the least significant bit as for fpc above.
    let flags = u32_at(dop, 4)?;
    let revision_markup = RevisionMarkup {
        on_screen: flags & (1 << 27) != 0,
        in_print: flags & (1 << 28) != 0,
        has_marks: false,
    };
    // MS-DOC 2.7.3 DopBase.fFacingPages explicitly maps to evenAndOddHeaders.
    Ok(Some(Properties {
        revision_markup,
        default_tab_twips: interval,
        even_and_odd_headers: dop[0] & 1 != 0,
        notes,
        adjust_line_height_in_table,
        balance_single_byte_double_byte_width,
        character_spacing_control,
    }))
}

impl Properties {
    pub(super) fn xml(&self) -> String {
        let interval = self.default_tab_twips;
        let facing = if self.even_and_odd_headers {
            "<w:evenAndOddHeaders/>"
        } else {
            ""
        };
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="{interval}"/>{facing}</w:settings>"#
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn default_tab_twips(word: &[u8], table: &[u8]) -> Result<Option<u16>, String> {
        Ok(read(word, table)?.map(|p| p.default_tab_twips))
    }

    fn fixture(size: u32, interval: u16) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0; 0x19a];
        word[0x192..0x196].copy_from_slice(&7u32.to_le_bytes());
        word[0x196..0x19a].copy_from_slice(&size.to_le_bytes());
        let mut table = vec![0; 7 + size as usize];
        if size >= 12 {
            table[17..19].copy_from_slice(&interval.to_le_bytes());
        }
        (word, table)
    }

    #[test]
    fn preserves_unsigned_interval_from_common_dop_prefix() {
        for size in [84, 500, 544, 594, 616, 674, 690, 694] {
            for interval in [1, 360, 720, 2160, u16::MAX] {
                let (word, table) = fixture(size, interval);
                assert_eq!(default_tab_twips(&word, &table).unwrap(), Some(interval));
                assert!(read(&word, &table)
                    .unwrap()
                    .unwrap()
                    .xml()
                    .contains(&format!("w:val=\"{interval}\"")));
            }
        }
    }

    #[test]
    fn revision_markup_display_flags_come_from_dopbase() {
        for (flags, on_screen, in_print) in [
            (0u32, false, false),
            (1 << 27, true, false),
            (1 << 28, false, true),
            ((1 << 27) | (1 << 28), true, true),
            (!((1 << 27) | (1 << 28)), false, false),
        ] {
            let (word, mut table) = fixture(84, 720);
            // Dop at table offset 7; the second DopBase flag word at offset 4.
            table[7 + 4..7 + 8].copy_from_slice(&flags.to_le_bytes());
            let markup = read(&word, &table).unwrap().unwrap().revision_markup;
            assert_eq!(
                (markup.on_screen, markup.in_print),
                (on_screen, in_print),
                "{flags:#x}"
            );
        }
    }

    #[test]
    fn dop2000_copts_controls_table_line_grid_adjustment() {
        let adjust = |size: u32, flags: u32| {
            let (word, mut table) = fixture(size, 720);
            if size >= 544 {
                // Dop at table offset 7; Copts flag word at Dop offset 512.
                table[7 + 512..7 + 516].copy_from_slice(&flags.to_le_bytes());
            }
            read(&word, &table)
                .unwrap()
                .unwrap()
                .adjust_line_height_in_table
        };
        for size in [544, 616, 694] {
            assert!(adjust(size, 0));
            assert!(!adjust(size, 1 << 3));
            // Neighboring Copts bits are independent.
            assert!(adjust(size, !(1 << 3)));
        }
        // Dop97 and older carry no fDontAdjustLineHeightInTable flag.
        assert!(adjust(500, 0));
        assert!(adjust(84, 0));
    }

    #[test]
    fn dop_typography_and_copts60_project_east_asian_spacing() {
        let facts = |size: u32, copts60: u16, typography: u16| {
            let (word, mut table) = fixture(size, 720);
            table[7 + 8..7 + 10].copy_from_slice(&copts60.to_le_bytes());
            if size >= 500 {
                table[7 + 90..7 + 92].copy_from_slice(&typography.to_le_bytes());
            }
            read(&word, &table).map(|value| {
                let value = value.unwrap();
                (
                    value.balance_single_byte_double_byte_width,
                    value.character_spacing_control,
                )
            })
        };
        assert_eq!(facts(694, 0, 0).unwrap(), (true, Some("doNotCompress")));
        assert_eq!(
            facts(694, 1 << 15, 1 << 1).unwrap(),
            (false, Some("compressPunctuation"))
        );
        assert_eq!(
            facts(694, 0x7fff, (2 << 1) | 1 | (3 << 3)).unwrap(),
            (true, Some("compressPunctuationAndJapaneseKana"))
        );
        assert!(facts(694, 0, 3 << 1).is_err());
        assert_eq!(facts(88, 1 << 15, 0).unwrap(), (false, None));
    }

    #[test]
    fn missing_dop_is_distinct_from_truncated_or_zero_interval() {
        let (mut word, mut table) = fixture(0, 0);
        word[0x192..0x196].copy_from_slice(&u32::MAX.to_le_bytes());
        assert_eq!(default_tab_twips(&word, &table).unwrap(), None);
        for size in [1, 11, 12, 83] {
            (word, table) = fixture(size, 720);
            assert!(default_tab_twips(&word, &table).is_err());
        }
        (word, table) = fixture(84, 0);
        assert!(default_tab_twips(&word, &table).is_err());
        (word, table) = fixture(84, 720);
        table.pop();
        assert!(default_tab_twips(&word, &table).is_err());
        word[0x192..0x196].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(default_tab_twips(&word, &table).is_err());
        assert!(default_tab_twips(&word[..0x198], &table).is_err());
    }
}
