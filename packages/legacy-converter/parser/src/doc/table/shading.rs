//! MS-DOC 2.9.43/119/121/247/248 to ECMA-376 17.3.5 / ST_Shd.
//! Preserve the two colors and pattern, never approximate coverage with a tint.
use super::super::{border::ICO_COLORS, u16_at, unsupported};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Shading {
    pattern: &'static str,
    foreground: Color,
    background: Color,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Color {
    Auto,
    Rgb([u8; 3]),
}
impl std::fmt::Display for Color {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => f.write_str("auto"),
            Self::Rgb([r, g, b]) => write!(f, "{r:02X}{g:02X}{b:02X}"),
        }
    }
}
fn color(b: &[u8]) -> Result<Color, String> {
    match b[3] {
        0 => Ok(Color::Rgb([b[0], b[1], b[2]])),
        255 => Ok(Color::Auto),
        _ => Err(unsupported("invalid Word shading COLORREF")),
    }
}

fn pattern(ipat: u16) -> Result<Option<&'static str>, String> {
    const BASIC: [&str; 26] = [
        "clear",
        "solid",
        "pct5",
        "pct10",
        "pct20",
        "pct25",
        "pct30",
        "pct40",
        "pct50",
        "pct60",
        "pct70",
        "pct75",
        "pct80",
        "pct90",
        "horzStripe",
        "vertStripe",
        "reverseDiagStripe",
        "diagStripe",
        "horzCross",
        "diagCross",
        "thinHorzStripe",
        "thinVertStripe",
        "thinReverseDiagStripe",
        "thinDiagStripe",
        "thinHorzCross",
        "thinDiagCross",
    ];
    Ok(Some(match ipat {
        0..=25 => BASIC[usize::from(ipat)],
        0x25 => "pct12",
        0x26 => "pct15",
        0x2b => "pct35",
        0x2c => "pct37",
        0x2e => "pct45",
        0x31 => "pct55",
        0x33 => "pct62",
        0x34 => "pct65",
        0x39 => "pct85",
        0x3a => "pct87",
        0x3c => "pct95",
        0xffff => "nil",
        // These documented binary patterns have no ST_Shd counterpart.
        // Return an unsupported-property warning rather than fitting a tint.
        0x23..=0x3d => return Ok(None),
        _ => return Err(unsupported("invalid Word shading pattern")),
    }))
}

impl Shading {
    pub fn read(bytes: &[u8], old: bool) -> Result<Option<Self>, String> {
        if bytes.len() != if old { 2 } else { 10 } {
            return Err(unsupported("invalid Word shading length"));
        }
        let (ipat, foreground, background) = if old {
            let bits = u16_at(bytes, 0)?;
            if bits == 0xffff {
                return Ok(Some(Self {
                    pattern: "nil",
                    foreground: Color::Auto,
                    background: Color::Auto,
                }));
            }
            let palette = |index: u16| -> Result<Color, String> {
                let text = ICO_COLORS
                    .get(usize::from(index))
                    .ok_or_else(|| unsupported("invalid Word shading palette index"))?;
                Ok(if index == 0 {
                    Color::Auto
                } else {
                    let value = u32::from_str_radix(text, 16).expect("constant Ico palette");
                    Color::Rgb([(value >> 16) as u8, (value >> 8) as u8, value as u8])
                })
            };
            (bits >> 10, palette(bits & 31)?, palette((bits >> 5) & 31)?)
        } else {
            // ShdNil is a distinct sentinel (both COLORREFs all ones, ipatAuto),
            // not a literal white foreground/background. Raw style inheritance
            // operands are deliberately not handled by the fallback reader.
            if bytes[..8].iter().all(|b| *b == 255) && u16_at(bytes, 8)? == 0 {
                return Ok(Some(Self {
                    pattern: "nil",
                    foreground: Color::Auto,
                    background: Color::Auto,
                }));
            }
            (u16_at(bytes, 8)?, color(&bytes[..4])?, color(&bytes[4..8])?)
        };
        Ok(pattern(ipat)?.map(|pattern| Self {
            pattern,
            foreground,
            background,
        }))
    }

    pub fn xml(&self) -> String {
        format!(
            "<w:shd w:val=\"{}\" w:color=\"{}\" w:fill=\"{}\"/>",
            self.pattern, self.foreground, self.background
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn modern(ipat: u16) -> [u8; 10] {
        let [a, b] = ipat.to_le_bytes();
        [0x12, 0x34, 0x56, 0, 0x98, 0x76, 0x54, 0, a, b]
    }
    #[test]
    fn shading_preserves_both_colors_and_all_mapped_patterns() {
        let mut mapped = 0;
        let mut unmapped = 0;
        for ipat in 0..=u16::MAX {
            match Shading::read(&modern(ipat), false) {
                Ok(Some(value)) => {
                    mapped += 1;
                    assert!(value.xml().contains("w:color=\"123456\" w:fill=\"987654\""));
                }
                Ok(None) => {
                    unmapped += 1;
                    assert!((0x23..=0x3d).contains(&ipat));
                }
                Err(_) => assert!(!(ipat <= 25 || (0x23..=0x3d).contains(&ipat) || ipat == 0xffff)),
            }
        }
        assert_eq!((mapped, unmapped), (38, 16));
        assert_eq!(
            Shading::read(&modern(0x25), false)
                .unwrap()
                .unwrap()
                .pattern,
            "pct12"
        );
        assert_eq!(
            Shading::read(&modern(0x2c), false)
                .unwrap()
                .unwrap()
                .pattern,
            "pct37"
        );
    }
    #[test]
    fn shading_auto_and_nil_are_not_literal_white() {
        let auto = [0, 0, 0, 255, 0, 0, 0, 255, 0, 0];
        let nil = [255, 255, 255, 255, 255, 255, 255, 255, 0, 0];
        assert_eq!(
            Shading::read(&auto, false).unwrap().unwrap().xml(),
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"auto\"/>"
        );
        assert_eq!(Shading::read(&nil, false).unwrap().unwrap().pattern, "nil");
        assert_eq!(
            Shading::read(&[255, 255], true).unwrap().unwrap().pattern,
            "nil"
        );
        for index in 0u16..17 {
            let bytes = (index | (index << 5)).to_le_bytes();
            let shade = Shading::read(&bytes, true).unwrap().unwrap();
            assert_eq!(shade.foreground.to_string(), ICO_COLORS[index as usize]);
            assert_eq!(shade.background.to_string(), ICO_COLORS[index as usize]);
        }
    }
    #[test]
    fn shading_rejects_bad_lengths_palette_indices_and_color_flags() {
        for old in [false, true] {
            let expected = if old { 2 } else { 10 };
            for n in 0..=12 {
                if n != expected {
                    assert!(Shading::read(&vec![0; n], old).is_err());
                }
            }
        }
        for flag in 1..255 {
            let mut bytes = modern(0);
            bytes[3] = flag;
            assert!(Shading::read(&bytes, false).is_err());
            bytes[3] = 0;
            bytes[7] = flag;
            assert!(Shading::read(&bytes, false).is_err());
        }
        for index in 17u16..32 {
            assert!(Shading::read(&index.to_le_bytes(), true).is_err());
            assert!(Shading::read(&(index << 5).to_le_bytes(), true).is_err());
        }
        // Retained shading uses fixed-size values, no per-cell color allocation.
        assert!(std::mem::size_of::<Shading>() <= 24);
    }
}
