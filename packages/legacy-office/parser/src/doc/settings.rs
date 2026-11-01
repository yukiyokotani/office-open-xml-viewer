//! MS-DOC 2.5.6 FibRgFcLcb97.fcDop/lcbDop and 2.7.1-3 Dop/DopBase.
//! All supported DOP versions begin with the same 84-byte DopBase.
use super::{u16_at, u32_at, unsupported};

#[derive(Debug, PartialEq, Eq)]
pub(super) struct Properties {
    pub default_tab_twips: u16,
    pub even_and_odd_headers: bool,
    pub notes: NoteProperties,
    /// ECMA-376 Part 1 17.15.3.1 adjustLineHeightInTable, the inverse of
    /// MS-DOC 2.7.13 Copts.fDontAdjustLineHeightInTable.
    pub adjust_line_height_in_table: bool,
    /// MS-DOC §2.7.13 Copts.b, mapped to ECMA-376 Part 4 §14.8.3.38.
    pub split_pg_break_and_para_mark: Option<bool>,
    /// MS-DOC §2.7.13 Copts.B (bit1), mapped to ECMA-376 Part4 §14.8.3.20.
    /// None when the selected DOP owner predates Dop2000.
    pub footnote_layout_like_ww8: Option<bool>,
    /// MS-DOC 2.7.13 Copts.fDontUseHTMLParagraphAutoSpacing, mapped to
    /// ECMA-376 Part 4 14.8.3.15. None when the DOP predates Dop2000.
    pub do_not_use_html_paragraph_auto_spacing: Option<bool>,
    /// ECMA-376 Part 1 17.15.3.3 balanceSingleByteDoubleByteWidth, the
    /// inverse of MS-DOC 2.7.11 Copts60.fDntBlnSbDbWid.
    pub balance_single_byte_double_byte_width: bool,
    /// ECMA-376 Part 1 17.15.3.5 doNotExpandShiftReturn, the inverse of
    /// MS-DOC 2.7.11 Copts60.fExpShRtn.
    pub do_not_expand_shift_return: bool,
    /// ECMA-376 Part 1 17.15.1.18 characterSpacingControl from MS-DOC 2.7.16
    /// DopTypography.iJustification; `None` when the DOP predates Dop97.
    pub character_spacing_control: Option<&'static str>,
    /// MS-DOC 2.7.2 DopBase fRMView / fRMPrint.
    pub revision_markup: RevisionMarkup,
}

/// MS-DOC 2.7.2 DopBase: fRMView "whether to show any revision markup that is
/// present in this document" and fRMPrint "whether to print" it (note <166>:
/// they can differ). A Word-exported PDF of a document with both set shows
/// insertions underlined with margin change bars.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
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
pub(super) struct NoteProperties {
    pub footnote_position: u8,
    pub footnote_restart: u8,
    pub footnote_start: u16,
    pub endnote_restart: u8,
    pub endnote_start: u16,
    pub endnote_position: u8,
    pub formats: Option<(u16, u16)>,
}

pub(super) fn read(
    word: &[u8],
    table: &[u8],
    schema: super::fib::KnownFibSchema,
) -> Result<Option<Properties>, String> {
    let size = u32_at(word, 0x196)? as usize;
    // Recovery policy for incomplete documents and minimal synthetic fixtures:
    // the normative lcbDop is nonzero. The caller warns and retains the OOXML
    // default when the entire DOP is absent; no offset is dereferenced then.
    if size == 0 {
        if schema.requires_legacy_dop97_prefix() {
            return Err(unsupported("missing legacy Word Dop97 prefix"));
        }
        return Ok(None);
    }
    let offset = u32_at(word, 0x192)? as usize;
    let dop = table
        .get(offset..)
        .and_then(|tail| tail.get(..size))
        .filter(|dop| dop.len() >= 84)
        .ok_or_else(|| unsupported("truncated Word document properties"))?;
    // The legacy forward reading profile has a published Dop97 prefix,
    // including typography and note formats. A short blob cannot acquire
    // this layout by filling missing facts with defaults.
    if schema.requires_legacy_dop97_prefix() && dop.len() < 500 {
        return Err(unsupported("truncated legacy Word Dop97 prefix"));
    }
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
    // MS-DOC 2.7.1 chooses Dop97 when cswNew is zero. A bounded excess DOP
    // tail does not authorize later Copts facts. The Microsoft 2007 legacy
    // pre-2000 FIB reading family also owns Dop97 despite a present extension;
    // its physical tail remains uninterpreted. We retain the existing
    // prefix-only DOP tolerance, not a claim of conformant producer length;
    // Section 2.5.15 FIB zero-initialization is not applied to missing DOP fields.
    let owns_copts = schema.has_dop2000_fields() && dop.len() >= 544;
    let adjust_line_height_in_table = if owns_copts {
        u32_at(dop, 512)? & (1 << 3) == 0
    } else {
        true
    };
    // Copts.B is independent of neighboring shape and HTML-spacing flags.
    let footnote_layout_like_ww8 = if owns_copts {
        Some(u32_at(dop, 512)? & (1 << 1) != 0)
    } else {
        None
    };
    let split_pg_break_and_para_mark = if owns_copts {
        Some(u32_at(dop, 512)? & (1 << 27) != 0)
    } else {
        None
    };
    // Copts bit C (bit2), immediately before the independent table-grid bit.
    // Older/absent DOPs have no acquired fact; the existing shared consumer
    // retains its default rather than inventing an authored compatibility flag.
    let do_not_use_html_paragraph_auto_spacing = if owns_copts {
        Some(u32_at(dop, 512)? & (1 << 2) != 0)
    } else {
        None
    };
    // MS-DOC 2.7.2 DopBase.copts60 (offset 8), 2.7.11 bit P.
    let balance_single_byte_double_byte_width = u16_at(dop, 8)? & (1 << 15) == 0;
    // 2.7.11 bit N, two bits before P.
    let do_not_expand_shift_return = u16_at(dop, 8)? & (1 << 13) == 0;
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
        split_pg_break_and_para_mark,
        footnote_layout_like_ww8,
        do_not_use_html_paragraph_auto_spacing,
        balance_single_byte_double_byte_width,
        do_not_expand_shift_return,
        character_spacing_control,
    }))
}

impl Properties {}

#[cfg(test)]
mod tests {
    use super::*;
    // Existing settings unit cases exercise a modern DOP owner. Whole-CFB
    // tests separately establish absent-extension semantics through acquisition.
    fn read(word: &[u8], table: &[u8]) -> Result<Option<Properties>, String> {
        let mut header = vec![0; 0x388];
        super::super::fib::write_minimal_word97_header(&mut header);
        header[0x382..0x384].copy_from_slice(&2u16.to_le_bytes());
        header[0x384..0x386].copy_from_slice(&0x00d9u16.to_le_bytes());
        let schema = super::super::fib::known_schema(&header)?;
        super::read(word, table, schema)
    }
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
    fn legacy_actual_selector_requires_dop97_and_never_promotes_opaque_tail() {
        let mut header = vec![0; 0x500];
        super::super::fib::write_minimal_word97_header(&mut header);
        header[0x98..0x9a].copy_from_slice(&0x006cu16.to_le_bytes());
        let extension = 0x9a + 0x006c * 8;
        header[extension..extension + 2].copy_from_slice(&4u16.to_le_bytes());
        header[extension + 2..extension + 4].copy_from_slice(&0x00c3u16.to_le_bytes());
        header[extension + 4..extension + 8].fill(0xff);
        let schema = super::super::fib::known_schema(&header).unwrap();
        for size in [0, 84, 499] {
            let (word, table) = fixture(size, 720);
            assert!(super::read(&word, &table, schema).is_err());
        }
        for size in [500, 544] {
            let (word, mut table) = fixture(size, 720);
            if size >= 544 {
                table[7 + 500..].fill(0xff);
            }
            let properties = super::read(&word, &table, schema).unwrap().unwrap();
            assert!(properties.adjust_line_height_in_table);
            assert_eq!(properties.do_not_use_html_paragraph_auto_spacing, None);
            assert_eq!(properties.footnote_layout_like_ww8, None);
            assert_eq!(properties.character_spacing_control, Some("doNotCompress"));
            assert_eq!(properties.notes.formats, Some((0, 0)));
        }
    }

    #[test]
    fn preserves_unsigned_interval_from_common_dop_prefix() {
        for size in [84, 500, 544, 594, 616, 674, 690, 694] {
            for interval in [1, 360, 720, 2160, u16::MAX] {
                let (word, table) = fixture(size, interval);
                assert_eq!(default_tab_twips(&word, &table).unwrap(), Some(interval));
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
    fn copts60_exp_sh_rtn_projects_do_not_expand_shift_return() {
        let do_not_expand = |size: u32, copts60: u16| {
            let (word, mut table) = fixture(size, 720);
            table[7 + 8..7 + 10].copy_from_slice(&copts60.to_le_bytes());
            read(&word, &table).map(|value| value.unwrap().do_not_expand_shift_return)
        };
        // MS-DOC 2.7.11 bit N fExpShRtn is the inverse of doNotExpandShiftReturn.
        assert!(do_not_expand(694, 0).unwrap());
        assert!(!do_not_expand(694, 1 << 13).unwrap());
        // Neighbouring bits M (fLeaveBackslashAlone) and O (fDntULTrlSpc) are independent.
        assert!(do_not_expand(694, (1 << 12) | (1 << 14)).unwrap());
        // The DopBase prefix carries Copts60 in every DOP version.
        assert!(!do_not_expand(88, 1 << 13).unwrap());
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
