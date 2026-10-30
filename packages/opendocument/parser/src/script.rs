//! Script classes for script-dependent text properties (ODF Part 3 §20.358,
//! Table 22 "Unicode Code Point to Script Type Mapping").

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ScriptClass {
    Latin,
    Asian,
    Complex,
}

/// Table 22 ranges, sorted by start. Code points outside every range are
/// implementation-dependent in ODF (§20.358) and return `None`.
const TABLE_22: [(u32, u32, ScriptClass); 28] = {
    use ScriptClass::{Asian as A, Complex as C, Latin as L};
    [
        (0x0003, 0x001F, L),
        (0x0021, 0x009F, L),
        (0x00A1, 0x04FF, L),
        (0x0530, 0x058F, L),
        (0x0590, 0x074F, C),
        (0x0780, 0x07BF, C),
        (0x0900, 0x109F, C),
        (0x10A0, 0x10FF, L),
        (0x1100, 0x11FF, A),
        (0x1200, 0x137F, C),
        (0x13A0, 0x16FF, L),
        (0x1780, 0x18AF, C),
        (0x1E00, 0x1FFF, L),
        (0x2C60, 0x2C7F, L),
        (0x2C80, 0x2CE3, L),
        (0x2E80, 0x31BF, A),
        (0x31C0, 0x31EF, A),
        (0x3200, 0x4DBF, A),
        (0x4E00, 0xA4CF, A),
        (0xA720, 0xA7FF, L),
        (0xAC00, 0xD7AF, A),
        (0xF900, 0xFAFF, A),
        (0xFB50, 0xFDFF, C),
        (0xFE30, 0xFE4F, A),
        (0xFE70, 0xFEFF, C),
        (0xFF00, 0xFFEF, A),
        (0x20000, 0x2A6DF, A),
        (0x2F800, 0x2FA1F, A),
    ]
};

/// The Table 22 class of `c`, or `None` where ODF leaves it to the consumer.
pub fn table_22(c: char) -> Option<ScriptClass> {
    let code = c as u32;
    let index = TABLE_22.partition_point(|(start, _, _)| *start <= code);
    let (start, end, class) = *TABLE_22.get(index.checked_sub(1)?)?;
    (start..=end).contains(&code).then_some(class)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn follows_table_22_and_leaves_unmapped_points_open() {
        assert_eq!(table_22('A'), Some(ScriptClass::Latin));
        assert_eq!(table_22('1'), Some(ScriptClass::Latin));
        assert_eq!(table_22('é'), Some(ScriptClass::Latin));
        assert_eq!(table_22('ש'), Some(ScriptClass::Complex));
        assert_eq!(table_22('ก'), Some(ScriptClass::Complex));
        assert_eq!(table_22('日'), Some(ScriptClass::Asian));
        assert_eq!(table_22('の'), Some(ScriptClass::Asian));
        assert_eq!(table_22('、'), Some(ScriptClass::Asian));
        assert_eq!(table_22('Ａ'), Some(ScriptClass::Asian));
        assert_eq!(table_22('한'), Some(ScriptClass::Asian));
        assert_eq!(table_22('\u{2C65}'), Some(ScriptClass::Latin));
        for unmapped in [
            ' ',
            '\u{A0}',
            '\u{0500}',
            '\u{2014}',
            '\u{2CE4}',
            '\u{1F600}',
            '\u{E000}',
        ] {
            assert_eq!(table_22(unmapped), None, "{unmapped:?}");
        }
        assert!(TABLE_22.windows(2).all(|pair| pair[0].1 < pair[1].0));
    }
}
