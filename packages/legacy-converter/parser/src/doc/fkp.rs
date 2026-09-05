//! Physical text offsets, not logical piece order, select FKP properties.
//! [MS-DOC] 2.4.6.1/2, ChpxFkp, PapxFkp, PapxInFkp and PlcBteChpx/Papx.

use super::{u16_at, u32_at, unsupported};

#[derive(Clone, Copy)]
pub enum Kind {
    Character,
    Paragraph,
}

pub struct Run<'a> {
    pub start: usize,
    pub end: usize,
    pub properties: &'a [u8],
}

#[derive(Default)]
pub struct Index<'a>(Vec<Run<'a>>);

impl<'a> Index<'a> {
    pub fn read(word: &'a [u8], table: &[u8], kind: Kind) -> Result<Self, String> {
        let offset = match kind {
            Kind::Character => 0xfa,
            Kind::Paragraph => 0x102,
        };
        let plc = table_part(word, table, offset)?;
        if plc.is_empty() {
            return Ok(Self::default());
        }
        if plc.len() < 4 || (plc.len() - 4) % 8 != 0 {
            return Err(unsupported("invalid Word formatting page table"));
        }
        let count = (plc.len() - 4) / 8;
        // Resource policy, not a file-format limit. Borrow properties instead
        // of copying each overlapping FKP payload into every run.
        if count > 65_536 {
            return Err(unsupported("Word formatting page budget exceeded"));
        }
        let mut runs = Vec::new();
        for i in 0..count {
            let lower = u32_at(plc, i * 4)? as usize;
            let upper = u32_at(plc, i * 4 + 4)? as usize;
            if lower >= upper {
                return Err(unsupported("unordered Word formatting page ranges"));
            }
            let pn = u32_at(plc, (count + 1) * 4 + i * 4)? & 0x003f_ffff;
            let start = pn as usize * 512;
            let page = word
                .get(start..start + 512)
                .ok_or_else(|| unsupported("Word formatting page outside WordDocument"))?;
            let n = page[511] as usize;
            let (maximum, stride) = match kind {
                Kind::Character => (101, 1),
                Kind::Paragraph => (29, 13),
            };
            if n == 0 || n > maximum || runs.len() + n > 1_000_000 {
                return Err(unsupported(
                    "invalid or excessive Word formatting run count",
                ));
            }
            let pointers = (n + 1) * 4;
            let payload_start = pointers + n * stride;
            for j in 0..n {
                let start = u32_at(page, j * 4)? as usize;
                let end = u32_at(page, j * 4 + 4)? as usize;
                if start >= end || start < lower || end > upper {
                    return Err(unsupported("invalid Word formatting run range"));
                }
                let p = page[pointers + j * stride] as usize * 2;
                let properties = if p == 0 {
                    &[][..]
                } else {
                    if p < payload_start || p >= 511 {
                        return Err(unsupported("Word formatting payload overlaps its index"));
                    }
                    let cb = page[p] as usize;
                    let (begin, size) = match kind {
                        Kind::Character => (p + 1, cb),
                        Kind::Paragraph if cb != 0 => (p + 1, cb * 2 - 1),
                        Kind::Paragraph => {
                            let size = *page.get(p + 1).filter(|v| **v != 0).ok_or_else(|| {
                                unsupported("invalid extended Word paragraph payload")
                            })?;
                            (p + 2, size as usize * 2)
                        }
                    };
                    let bytes = page
                        .get(begin..begin + size)
                        .filter(|_| begin + size <= 511)
                        .ok_or_else(|| unsupported("Word formatting payload exceeds its page"))?;
                    if matches!(kind, Kind::Paragraph) && bytes.len() < 2 {
                        return Err(unsupported("missing Word paragraph style index"));
                    }
                    bytes
                };
                runs.push(Run {
                    start,
                    end,
                    properties,
                });
            }
        }
        Ok(Self(runs))
    }

    pub fn at(&self, fc: usize) -> Option<(usize, &Run<'a>)> {
        let i = self
            .0
            .partition_point(|run| run.start <= fc)
            .checked_sub(1)?;
        let run = &self.0[i];
        (fc < run.end).then_some((i, run))
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

pub fn table_part<'a>(word: &[u8], table: &'a [u8], field: usize) -> Result<&'a [u8], String> {
    let length = u32_at(word, field + 4)? as usize;
    if length == 0 {
        return Ok(&[]);
    }
    let start = u32_at(word, field)? as usize;
    let end = start
        .checked_add(length)
        .ok_or_else(|| unsupported("Word table range overflow"))?;
    table
        .get(start..end)
        .ok_or_else(|| unsupported("Word formatting table outside table stream"))
}

pub fn paragraph_style(index: &Index<'_>, fc: usize) -> Result<usize, String> {
    match index.at(fc).map(|(_, run)| run.properties) {
        None | Some([]) => Ok(0),
        Some(bytes) => Ok(u16_at(bytes, 0)? as usize),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(kind: Kind) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0; 1024];
        let field = if matches!(kind, Kind::Character) {
            0xfa
        } else {
            0x102
        };
        word[field + 4..field + 8].copy_from_slice(&12u32.to_le_bytes());
        let mut table = Vec::new();
        for value in [100u32, 110, 1] {
            table.extend(value.to_le_bytes());
        }
        word[512..516].copy_from_slice(&100u32.to_le_bytes());
        word[516..520].copy_from_slice(&110u32.to_le_bytes());
        word[520] = 20;
        word[1023] = 1;
        (word, table)
    }

    #[test]
    fn character_lookup_uses_half_open_physical_ranges() {
        let (mut word, table) = fixture(Kind::Character);
        word[552..556].copy_from_slice(&[3, 0x35, 8, 1]);
        let index = Index::read(&word, &table, Kind::Character).unwrap();
        assert!(index.at(99).is_none());
        assert_eq!(index.at(100).unwrap().1.properties, [0x35, 8, 1]);
        assert!(index.at(109).is_some());
        assert!(index.at(110).is_none());
    }

    #[test]
    fn supports_both_paragraph_payload_length_encodings() {
        let (mut word, table) = fixture(Kind::Paragraph);
        word[552..556].copy_from_slice(&[2, 7, 0, 0]);
        assert_eq!(
            paragraph_style(&Index::read(&word, &table, Kind::Paragraph).unwrap(), 109).unwrap(),
            7
        );
        word[552..556].copy_from_slice(&[0, 1, 9, 0]);
        assert_eq!(
            paragraph_style(&Index::read(&word, &table, Kind::Paragraph).unwrap(), 109).unwrap(),
            9
        );
    }

    #[test]
    fn rejects_overlapping_indices_and_truncated_pages() {
        let (mut word, table) = fixture(Kind::Character);
        word[520] = 1;
        assert!(Index::read(&word, &table, Kind::Character).is_err());
        word[520] = 20;
        word.truncate(1023);
        assert!(Index::read(&word, &table, Kind::Character).is_err());
    }
}
