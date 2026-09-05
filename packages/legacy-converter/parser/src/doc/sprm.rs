//! Bounded framing shared by Word character and paragraph property readers.
//! [MS-DOC] 2.2.5 and 2.6 (Sprm); variable-length tab/table exceptions.

use super::{u16_at, unsupported};

pub struct Budget(usize);
impl Default for Budget {
    fn default() -> Self {
        Self(4_000_000)
    }
}
impl Budget {
    pub fn take(&mut self) -> Result<(), String> {
        self.0 = self
            .0
            .checked_sub(1)
            .ok_or_else(|| unsupported("Word formatting operation budget exceeded"))?;
        Ok(())
    }
}

pub struct Sprms<'a> {
    bytes: &'a [u8],
}
impl<'a> Sprms<'a> {
    pub fn new(bytes: &'a [u8]) -> Self {
        Self { bytes }
    }
    pub fn next(&mut self, budget: &mut Budget) -> Result<Option<(u16, &'a [u8])>, String> {
        if self.bytes.is_empty() {
            return Ok(None);
        }
        budget.take()?;
        let code = u16_at(self.bytes, 0)?;
        let bytes = &self.bytes[2..];
        let size = match code >> 13 {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            7 => 3,
            _ if code == 0xd608 => {
                let cb = u16_at(bytes, 0)? as usize;
                if cb == 0 {
                    return Err(unsupported("invalid Word table property size"));
                }
                cb + 1
            }
            _ if code == 0xc615 && bytes.first() == Some(&255) => {
                let deleted = *bytes
                    .get(1)
                    .ok_or_else(|| unsupported("truncated Word tab property"))?
                    as usize;
                let added = *bytes
                    .get(2 + deleted * 4)
                    .ok_or_else(|| unsupported("truncated Word tab additions"))?
                    as usize;
                3 + deleted * 4 + added * 3
            }
            _ => {
                1 + *bytes
                    .first()
                    .ok_or_else(|| unsupported("truncated Word variable property"))?
                    as usize
            }
        };
        let operand = bytes
            .get(..size)
            .ok_or_else(|| unsupported("truncated Word formatting operand"))?;
        self.bytes = &bytes[size..];
        Ok(Some((code, operand)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn skips_unknown_variable_operands_and_rejects_truncation() {
        let mut budget = Budget::default();
        let mut p = Sprms::new(&[0x71, 0xca, 2, 9, 8, 0x35, 8, 1]);
        assert_eq!(p.next(&mut budget).unwrap(), Some((0xca71, &[2, 9, 8][..])));
        assert_eq!(p.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
        assert!(p.next(&mut budget).unwrap().is_none());
        assert!(Sprms::new(&[0x71, 0xca, 3, 9]).next(&mut budget).is_err());
        assert!(Sprms::new(&[1]).next(&mut budget).is_err());
    }
}
