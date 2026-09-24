//! MS-DOC 2.5.6 FibRgFcLcb97.fcDop/lcbDop and 2.7.1-3 Dop/DopBase.
//! All supported DOP versions begin with the same 84-byte DopBase.
use super::{u16_at, u32_at, unsupported};

#[derive(Debug, PartialEq, Eq)]
pub(super) struct Properties {
    pub default_tab_twips: u16,
    pub even_and_odd_headers: bool,
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub notes: NoteProperties,
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
    // MS-DOC 2.7.3 DopBase.fFacingPages explicitly maps to evenAndOddHeaders.
    Ok(Some(Properties {
        default_tab_twips: interval,
        even_and_odd_headers: dop[0] & 1 != 0,
        notes,
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
