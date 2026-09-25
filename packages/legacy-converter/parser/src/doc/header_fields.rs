//! MS-DOC 2.8.25 Plcfld field-marker tables (2.9.88 Fld, 2.9.89 FldCh,
//! 2.9.110 grffldEnd), with CPs relative to their document part. The direct
//! model evaluates the passive fields it supports; nothing executes.
use super::{u32_at, unsupported, MAX_STORY_CONTROLS};
use std::collections::BTreeMap;

pub(super) struct Table(BTreeMap<usize, (u8, u8)>);
impl Table {
    pub fn read(word: &[u8], table: &[u8], length: usize) -> Result<Self, String> {
        Self::read_at(word, table, 0x122, length)
    }

    /// MS-DOC 2.8.25: each document part has its own Plcfld with CPs relative
    /// to that part. `fib_offset` selects the FibRgFcLcb97 fc/lcb pair
    /// (fcPlcfFldMom 0x11A, fcPlcfFldHdr 0x122, fcPlcfFldFtn 0x12A,
    /// fcPlcfFldEdn 0x21A).
    pub fn read_at(
        word: &[u8],
        table: &[u8],
        fib_offset: usize,
        length: usize,
    ) -> Result<Self, String> {
        let size = u32_at(word, fib_offset + 4)? as usize;
        let mut entries = BTreeMap::new();
        if size == 0 {
            return Ok(Self(entries));
        }
        if size < 4 || !(size - 4).is_multiple_of(6) || (size - 4) / 6 > MAX_STORY_CONTROLS {
            return Err(unsupported("invalid Word header field table length"));
        }
        let offset = u32_at(word, fib_offset)? as usize;
        let plc = table
            .get(offset..)
            .and_then(|b| b.get(..size))
            .ok_or_else(|| unsupported("Word header field table outside its stream"))?;
        let count = (size - 4) / 6;
        let mut previous = None;
        for i in 0..=count {
            let cp = u32_at(plc, i * 4)? as usize;
            // MS-DOC 2.8.25: the final CP is only an ordering sentinel,
            // not a field location. Its value is otherwise undefined.
            if (i < count && cp >= length) || previous.is_some_and(|p| cp <= p) {
                return Err(unsupported("invalid Word header field position"));
            }
            previous = Some(cp);
            if i < count {
                let offset = (count + 1) * 4 + i * 2;
                let ch = plc[offset] & 0x1f; // fldch reserved bits MUST be ignored.
                if !matches!(ch, 0x13..=0x15) {
                    return Err(unsupported("invalid Word field marker"));
                }
                entries.insert(cp, (ch, plc[offset + 1]));
            }
        }
        Ok(Self(entries))
    }

    /// The field character and its Fld.grffld byte at `cp`, if listed.
    pub(in crate::doc) fn get(&self, cp: usize) -> Option<(u8, u8)> {
        self.0.get(&cp).copied()
    }

    pub(in crate::doc) fn len(&self) -> usize {
        self.0.len()
    }

    #[cfg(test)]
    pub(in crate::doc) fn for_test(entries: &[(u32, u8, u8)], _final_cp: u32) -> Self {
        Self(
            entries
                .iter()
                .map(|(cp, ch, grffld)| (*cp as usize, (*ch, *grffld)))
                .collect(),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_field_table_boundaries_and_ignores_reserved_marker_bits() {
        let mut word = vec![0; 0x12a];
        word[0x126..0x12a].copy_from_slice(&10u32.to_le_bytes());
        let mut table = Vec::from(2u32.to_le_bytes());
        table.extend(4u32.to_le_bytes());
        table.extend([0xf3, 0]);
        assert_eq!(Table::read(&word, &table, 4).unwrap().0[&2].0, 0x13);
        table[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(Table::read(&word, &table, 4).is_ok());
        for cp in [2u32, 1] {
            table[4..8].copy_from_slice(&cp.to_le_bytes());
            assert!(Table::read(&word, &table, 4).is_err());
        }
        table[0..4].copy_from_slice(&4u32.to_le_bytes());
        table[4..8].copy_from_slice(&6u32.to_le_bytes());
        assert!(Table::read(&word, &table, 4).is_err());
        word[0x126..0x12a].copy_from_slice(&0u32.to_le_bytes());
        word[0x122..0x126].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(Table::read(&word, &[], 4).unwrap().0.is_empty());
    }
}
