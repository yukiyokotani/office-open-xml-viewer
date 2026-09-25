//! Cell-margin operands retained for the native table-style cascade.
//!
//! [MS-DOC] 2.6.3 and 2.9.45/2.9.46 permit an authored `ftsNil` value whose
//! numeric size is ignored. Office controls distinguish that value from an
//! omitted property, so the native path keeps it until the effective margin is
//! resolved.

use crate::doc::{u16_at, unsupported};

pub(super) const DEFAULTS: [u16; 4] = [0, 108, 0, 108];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum Value {
    Nil,
    Dxa(u16),
}

impl Value {
    pub(in crate::doc) fn resolved(self) -> u16 {
        match self {
            Self::Nil => 0,
            Self::Dxa(value) => value,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(in crate::doc) struct Patch {
    sides: [Option<Value>; 4],
}

impl Patch {
    pub(super) fn apply(&mut self, cssa: Cssa) {
        for side in 0..4 {
            if cssa.sides & (1 << side) != 0 {
                self.sides[side] = Some(cssa.value);
            }
        }
    }

    pub(in crate::doc) fn get(self, side: usize) -> Option<Value> {
        self.sides[side]
    }

    pub(in crate::doc) fn overlay(&mut self, patch: Self) {
        for side in 0..4 {
            if patch.sides[side].is_some() {
                self.sides[side] = patch.sides[side];
            }
        }
    }

    pub(in crate::doc) fn apply_style(&mut self, code: u16, bytes: &[u8]) -> Result<u8, String> {
        let cssa = read(bytes)?;
        if (cssa.first, cssa.limit) != (0, 1) {
            return Err(unsupported("invalid Word style cell margin range"));
        }
        if code == 0xd63e && !matches!(cssa.value, Value::Dxa(_)) {
            return Err(unsupported("nil Word style cell margin"));
        }
        self.apply(cssa);
        Ok(cssa.sides)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct Cssa {
    pub(super) first: usize,
    pub(super) limit: usize,
    pub(super) sides: u8,
    pub(super) value: Value,
}

pub(super) fn read(bytes: &[u8]) -> Result<Cssa, String> {
    if bytes.len() != 7 || bytes[0] != 6 {
        return Err(unsupported("invalid Word cell margin operand length"));
    }
    if bytes[3] & !0x0f != 0 {
        return Err(unsupported("invalid Word cell margin sides"));
    }
    let width = u16_at(bytes, 5)?;
    let value = match bytes[4] {
        0 if width == 0 => Value::Nil,
        0 => return Err(unsupported("nonzero Word nil cell margin")),
        3 if width <= 31_680 => Value::Dxa(width),
        3 => return Err(unsupported("oversized Word cell margin")),
        _ => return Err(unsupported("invalid Word cell margin unit")),
    };
    Ok(Cssa {
        first: usize::from(bytes[1]),
        limit: usize::from(bytes[2]),
        sides: bytes[3],
        value,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cssa(unit: u8, width: u16) -> [u8; 7] {
        let [lo, hi] = width.to_le_bytes();
        [6, 0, 1, 0x0f, unit, lo, hi]
    }

    #[test]
    fn cssa_retains_nil_separately_from_dxa_zero() {
        assert_eq!(read(&cssa(0, 0)).unwrap().value, Value::Nil);
        assert_eq!(read(&cssa(3, 0)).unwrap().value, Value::Dxa(0));
        assert_eq!(read(&cssa(3, 31_680)).unwrap().value, Value::Dxa(31_680));
    }

    #[test]
    fn cssa_rejects_bad_framing_units_sides_and_widths() {
        assert!(read(&[]).is_err());
        assert!(read(&[5, 0, 1, 0x0f, 3, 0, 0]).is_err());
        assert!(read(&[6, 0, 1, 0x10, 3, 0, 0]).is_err());
        assert!(read(&cssa(0, 1)).is_err());
        assert!(read(&cssa(1, 0)).is_err());
        assert!(read(&cssa(3, 31_681)).is_err());
    }

    #[test]
    fn style_margin_range_and_d63e_unit_are_exact() {
        let mut patch = Patch::default();
        assert!(patch.apply_style(0xd63e, &cssa(0, 0)).is_err());
        assert_eq!(patch.apply_style(0xd63e, &cssa(3, 0)).unwrap(), 0x0f);
        assert_eq!(patch.get(0), Some(Value::Dxa(0)));

        let mut bad_range = cssa(3, 10);
        bad_range[2] = 2;
        assert!(Patch::default().apply_style(0xd634, &bad_range).is_err());
    }
}
