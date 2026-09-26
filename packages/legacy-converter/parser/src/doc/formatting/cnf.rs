//! Shared framing for table-style conditional formatting operands.
//! [MS-DOC] 2.9.41 requires an exact `cb`, one listed `cnfc`, and a remaining
//! `grpprl`. Property-family readers decide whether and how to inspect that
//! group; this parser never recursively expands nested CNF records.

use crate::doc::{u16_at, unsupported};

pub(super) const CONDITIONS: [u16; 12] = [
    0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080, 0x0100, 0x0200, 0x0400, 0x0800,
];

pub(super) struct Operand<'a> {
    pub(super) condition: u16,
    pub(super) grpprl: &'a [u8],
}

pub(super) fn parse(bytes: &[u8]) -> Result<Operand<'_>, String> {
    if bytes.len() < 3 || usize::from(bytes[0]) + 1 != bytes.len() {
        return Err(unsupported("invalid Word conditional formatting operand"));
    }
    let condition = u16_at(bytes, 1)?;
    if !CONDITIONS.contains(&condition) {
        return Err(unsupported("invalid Word table style condition"));
    }
    Ok(Operand {
        condition,
        grpprl: &bytes[3..],
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_exact_framing_and_each_single_condition_without_expanding_group() {
        let nested_cnf_bytes = [0x85, 0xca, 1];
        for condition in CONDITIONS {
            let mut bytes = vec![5];
            bytes.extend(condition.to_le_bytes());
            bytes.extend(nested_cnf_bytes);
            let parsed = parse(&bytes).unwrap();
            assert_eq!(parsed.condition, condition);
            assert_eq!(parsed.grpprl, nested_cnf_bytes);
        }
    }

    #[test]
    fn rejects_short_inexact_and_combined_conditions() {
        for bytes in [&[][..], &[2, 1][..], &[2, 1, 0, 0][..], &[3, 1, 0][..]] {
            assert!(parse(bytes).is_err());
        }
        for condition in [0, 0x0003, 0x1000, 0xffff] {
            let bytes = [2, condition as u8, (condition >> 8) as u8];
            assert!(parse(&bytes).is_err());
        }
    }
}
