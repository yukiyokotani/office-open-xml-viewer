//! [MS-DOC] Brc/Brc80/BrcType/Ico to ECMA-376 ST_Border, no pixel fitting.
use super::{u32_at, unsupported};

#[derive(Clone, Default)]
pub struct Border {
    attributes: String,
}
impl Border {
    pub fn read(bytes: &[u8], old: bool) -> Result<Self, String> {
        let size = if old { 4 } else { 8 };
        let b = bytes
            .get(..size)
            .ok_or_else(|| unsupported("short Word border"))?;
        if u32_at(b, size - 4)? == u32::MAX {
            return Ok(Self {
                attributes: "w:val=\"nil\"".into(),
            });
        }
        let (width, kind, color, flags) = if old {
            // Ico is a fixed binary-format palette, not a document theme.
            let colors = [
                "auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00",
                "FFFFFF", "000080", "008080", "008000", "800080", "800080", "808000", "808080",
                "C0C0C0",
            ];
            (
                b[0],
                b[1],
                colors
                    .get(b[2] as usize)
                    .ok_or_else(|| unsupported("invalid Word border palette index"))?
                    .to_string(),
                b[3],
            )
        } else {
            (
                b[4],
                b[5],
                if b[3] == 0xff {
                    "auto".into()
                } else {
                    format!("{:02X}{:02X}{:02X}", b[0], b[1], b[2])
                },
                b[6],
            )
        };
        let style = match kind {
            0 => "none",
            1 | 5 => "single",
            3 => "double",
            6 => "dotted",
            7 => "dashed",
            8 => "dotDash",
            9 => "dotDotDash",
            10 => "triple",
            11 => "thinThickSmallGap",
            12 => "thickThinSmallGap",
            13 => "thinThickThinSmallGap",
            14 => "thinThickMediumGap",
            15 => "thickThinMediumGap",
            16 => "thinThickThinMediumGap",
            17 => "thinThickLargeGap",
            18 => "thickThinLargeGap",
            19 => "thinThickThinLargeGap",
            20 => "wave",
            21 => "doubleWave",
            22 => "dashSmallGap",
            23 => "dashDotStroked",
            24 => "threeDEmboss",
            25 => "threeDEngrave",
            26 => "outset",
            27 => "inset",
            _ => return Err(unsupported("invalid Word cell border type")),
        };
        // Brc widths below 2 are normatively treated as 2 eighth-points.
        Ok(Self {attributes:format!("w:val=\"{style}\" w:sz=\"{}\" w:color=\"{color}\" w:space=\"{}\" w:shadow=\"{}\" w:frame=\"{}\"",width.max(2),flags&31,u8::from(flags&32!=0),u8::from(flags&64!=0))})
    }
    pub fn xml(&self, side: &str) -> String {
        format!("<w:{side} {}/>", self.attributes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn preserves_nil_palette_and_colorref_instead_of_byte_order_swapping() {
        assert!(Border::read(&[255; 4], true)
            .unwrap()
            .xml("top")
            .contains("w:val=\"nil\""));
        let b = Border::read(&[0, 6, 2, 0], true).unwrap().xml("top");
        assert!(b.contains("w:color=\"0000FF\""));
        assert!(b.contains("w:sz=\"2\""));
        assert!(Border::read(&[0x12, 0x34, 0x56, 0, 8, 3, 0, 0], false)
            .unwrap()
            .xml("left")
            .contains("w:color=\"123456\""));
    }
}
