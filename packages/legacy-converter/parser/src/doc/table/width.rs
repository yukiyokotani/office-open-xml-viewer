//! Preferred table widths from [MS-DOC] 2.9.101, 2.9.103, and 2.9.104.
//! These are layout preferences, not the physical edges carried by TDefTable.

use super::super::unsupported;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum PreferredWidth {
    Auto,
    Percent(u16),
    Dxa(u16),
}

#[derive(Clone, Copy)]
enum Scope {
    Table,
    Part,
    Tc80,
}

impl PreferredWidth {
    pub(super) fn table(bytes: &[u8]) -> Result<Option<Self>, String> {
        read(bytes, Scope::Table)
    }

    pub(super) fn part(bytes: &[u8]) -> Result<Option<Self>, String> {
        read(bytes, Scope::Part)
    }

    pub(super) fn tc80(flags: u16, value: u16) -> Result<Option<Self>, String> {
        // MS-DOC 2.9.313 requires a nonnegative integer but does not establish
        // how a high-bit value is interpreted. Reject it rather than reinterpret
        // or clamp an unproven representation.
        if value > i16::MAX as u16 {
            return Err(unsupported(
                "unsupported high-bit Word TC80 preferred width",
            ));
        }
        read_parts(((flags >> 9) & 7) as u8, value, Scope::Tc80)
    }

    pub(in crate::doc) fn kind(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Percent(_) => "pct",
            Self::Dxa(_) => "dxa",
        }
    }

    pub(in crate::doc) fn value(self) -> u16 {
        match self {
            Self::Auto => 0,
            Self::Percent(value) | Self::Dxa(value) => value,
        }
    }
}

fn read(bytes: &[u8], scope: Scope) -> Result<Option<PreferredWidth>, String> {
    if bytes.len() != 3 {
        return Err(unsupported("invalid Word preferred width length"));
    }
    read_parts(bytes[0], u16::from_le_bytes([bytes[1], bytes[2]]), scope)
}

fn read_parts(fts: u8, value: u16, scope: Scope) -> Result<Option<PreferredWidth>, String> {
    match fts {
        // ftsNil is an absent preference in MS-DOC, not ECMA ST_TblWidth nil.
        // Table requires zero; TablePart explicitly says its value is ignored.
        0 if matches!(scope, Scope::Table) && value != 0 => {
            Err(unsupported("nonzero Word nil table width"))
        }
        0 => Ok(None),
        1 if value == 0 || matches!(scope, Scope::Tc80) => Ok(Some(PreferredWidth::Auto)),
        1 => Err(unsupported("nonzero Word automatic table width")),
        2 if matches!(scope, Scope::Table) && value <= 30_000 => {
            Ok(Some(PreferredWidth::Percent(value)))
        }
        2 if matches!(scope, Scope::Part) && value <= 5_000 => {
            Ok(Some(PreferredWidth::Percent(value)))
        }
        2 if matches!(scope, Scope::Tc80) => Ok(Some(PreferredWidth::Percent(value))),
        2 => Err(unsupported("Word percentage table width outside range")),
        3 if matches!(scope, Scope::Tc80) || value <= 31_680 => {
            Ok(Some(PreferredWidth::Dxa(value)))
        }
        3 => Err(unsupported("Word absolute table width outside range")),
        _ => Err(unsupported("invalid Word table width unit")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_table_and_part_boundaries_without_narrowing_tc80() {
        assert_eq!(PreferredWidth::table(&[0, 0, 0]).unwrap(), None);
        assert!(PreferredWidth::table(&[0, 1, 0]).is_err());
        assert_eq!(PreferredWidth::part(&[0, 0xff, 0xff]).unwrap(), None);
        assert_eq!(
            PreferredWidth::table(&[1, 0, 0]).unwrap(),
            Some(PreferredWidth::Auto)
        );
        assert!(PreferredWidth::part(&[1, 1, 0]).is_err());
        assert_eq!(
            PreferredWidth::table(&[2, 0x30, 0x75]).unwrap(),
            Some(PreferredWidth::Percent(30_000))
        );
        assert!(PreferredWidth::table(&[2, 0x31, 0x75]).is_err());
        assert_eq!(
            PreferredWidth::part(&[2, 0x88, 0x13]).unwrap(),
            Some(PreferredWidth::Percent(5_000))
        );
        assert!(PreferredWidth::part(&[2, 0x89, 0x13]).is_err());
        assert_eq!(
            PreferredWidth::table(&[3, 0xc0, 0x7b]).unwrap(),
            Some(PreferredWidth::Dxa(31_680))
        );
        assert!(PreferredWidth::table(&[3, 0xc1, 0x7b]).is_err());
        assert_eq!(
            PreferredWidth::tc80(3 << 9, i16::MAX as u16).unwrap(),
            Some(PreferredWidth::Dxa(i16::MAX as u16))
        );
        assert!(PreferredWidth::tc80(3 << 9, 0x8000).is_err());
        assert_eq!(
            PreferredWidth::tc80(1 << 9, 1234).unwrap(),
            Some(PreferredWidth::Auto)
        );
        assert!(PreferredWidth::table(&[4, 0, 0]).is_err());
        assert!(PreferredWidth::table(&[0, 0]).is_err());
    }
}
