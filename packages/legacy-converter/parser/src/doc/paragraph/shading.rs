//! Paragraph and character shading (MS-DOC 2.6.1 sprmCShd/sprmCShd80,
//! 2.6.2 sprmPShd/sprmPShd80) reduced to the fill-only shading that the DOCX
//! model represents (`DocParagraph::shading`, `TextRun::background`,
//! ECMA-376 17.3.1.31 / 17.3.2.32 `w:shd@w:fill`).

use super::super::{table::Shading, unsupported};

#[derive(Clone, Debug, PartialEq, Eq)]
pub(in crate::doc) enum ShadingFill {
    /// ShdAuto / ShdNil / a clear pattern over an automatic background:
    /// MS-DOC 2.9.247 states that no shading is applied.
    None,
    /// Lowercase RGB hex of the completely covering color.
    Rgb(String),
}

/// Decode an SHDOperand (`modern`, MS-DOC 2.9.249: cb = 10 then Shd) or a
/// Shd80 (MS-DOC 2.9.248) and project it. `Ok(None)` means the value is valid
/// but not representable by a single fill color, so the caller must keep the
/// property unsupported rather than approximate a pattern with a tint.
pub(in crate::doc) fn fill(operand: &[u8], modern: bool) -> Result<Option<ShadingFill>, String> {
    let shading = if modern {
        if operand.first() != Some(&10) || operand.len() != 11 {
            return Err(unsupported("invalid Word shading operand"));
        }
        Shading::read(&operand[1..], false)?
    } else {
        Shading::read(operand, true)?
    };
    // Patterns without an ST_Shd counterpart remain unrepresentable.
    let Some(shading) = shading else {
        return Ok(None);
    };
    let facts = shading.direct_facts();
    Ok(match facts.pattern {
        // ShdNil outside a table style specifies no shading (MS-DOC 2.9.247).
        "nil" => Some(ShadingFill::None),
        // ipatAuto/clear shows only the background color; an automatic
        // background is ShdAuto, which applies no shading.
        "clear" => Some(match shading.direct_background() {
            Some(fill) => ShadingFill::Rgb(fill),
            None => ShadingFill::None,
        }),
        // ipatSolid (ST_Shd solid) is a 100% foreground pattern. An automatic
        // foreground has no documented concrete color, so it stays gated.
        "solid" => match facts.foreground {
            super::super::table::Color::Rgb([r, g, b]) => {
                Some(ShadingFill::Rgb(format!("{r:02x}{g:02x}{b:02x}")))
            }
            super::super::table::Color::Auto => None,
        },
        // Percentage and hatch patterns mix two colors; the fill-only model
        // cannot represent them without inventing a blend.
        _ => None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn modern(fore: [u8; 4], back: [u8; 4], ipat: u16) -> Vec<u8> {
        let mut operand = vec![10];
        operand.extend(fore);
        operand.extend(back);
        operand.extend(ipat.to_le_bytes());
        operand
    }

    #[test]
    fn modern_shading_projects_only_single_color_patterns() {
        let auto = [0, 0, 0, 0xff];
        assert_eq!(
            fill(&modern(auto, [0xdd, 0xdd, 0xdd, 0], 0), true).unwrap(),
            Some(ShadingFill::Rgb("dddddd".into()))
        );
        assert_eq!(
            fill(&modern(auto, [0, 0, 0, 0], 0), true).unwrap(),
            Some(ShadingFill::Rgb("000000".into()))
        );
        // Automatic background (including the all-ones COLORREF) = ShdAuto.
        assert_eq!(
            fill(&modern(auto, [0xff; 4], 0), true).unwrap(),
            Some(ShadingFill::None)
        );
        assert_eq!(
            fill(&modern(auto, auto, 0), true).unwrap(),
            Some(ShadingFill::None)
        );
        // ShdNil.
        assert_eq!(
            fill(&modern([0xff; 4], [0xff; 4], 0), true).unwrap(),
            Some(ShadingFill::None)
        );
        assert_eq!(
            fill(
                &modern([0x12, 0x34, 0x56, 0], [0xff, 0xff, 0xff, 0], 1),
                true
            )
            .unwrap(),
            Some(ShadingFill::Rgb("123456".into()))
        );
        assert_eq!(fill(&modern(auto, [1, 2, 3, 0], 1), true).unwrap(), None);
        for ipat in [2u16, 0x26, 0x3d] {
            assert_eq!(
                fill(&modern(auto, [0xff, 0xff, 0xff, 0], ipat), true).unwrap(),
                None
            );
        }
        assert!(fill(&modern(auto, auto, 0x3e), true).is_err());
        assert!(fill(&modern([0, 0, 0, 1], auto, 0), true).is_err());
        let mut short = modern(auto, auto, 0);
        short[0] = 9;
        assert!(fill(&short, true).is_err());
        assert!(fill(&short[..10], true).is_err());
    }

    #[test]
    fn shd80_uses_the_ico_palette_and_nil_sentinel() {
        // icoBack = 8 (white), clear.
        assert_eq!(
            fill(&0x0100u16.to_le_bytes(), false).unwrap(),
            Some(ShadingFill::Rgb("ffffff".into()))
        );
        // icoBack = 16 (C0C0C0).
        assert_eq!(
            fill(&0x0200u16.to_le_bytes(), false).unwrap(),
            Some(ShadingFill::Rgb("c0c0c0".into()))
        );
        // icoBack = 1 (black).
        assert_eq!(
            fill(&0x0020u16.to_le_bytes(), false).unwrap(),
            Some(ShadingFill::Rgb("000000".into()))
        );
        assert_eq!(fill(&[0, 0], false).unwrap(), Some(ShadingFill::None));
        assert_eq!(fill(&[0xff, 0xff], false).unwrap(), Some(ShadingFill::None));
        // pct15 over white: two-color mix.
        assert_eq!(fill(&0x9900u16.to_le_bytes(), false).unwrap(), None);
        // solid, icoFore = 6 (red).
        assert_eq!(
            fill(&(0x0400u16 | 6).to_le_bytes(), false).unwrap(),
            Some(ShadingFill::Rgb("ff0000".into()))
        );
        assert!(fill(&[0], false).is_err());
        assert!(fill(&(17u16).to_le_bytes(), false).is_err());
    }
}
