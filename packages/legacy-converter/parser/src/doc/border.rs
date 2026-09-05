//! [MS-DOC] Brc/Brc80/BrcType/Ico to ECMA-376 ST_Border, no pixel fitting.
use super::{u32_at, unsupported};

#[derive(Clone, Default)]
pub struct Border {
    attributes: String,
}
impl Border {
    /// Paragraph Brc80/Brc values, including the documented no-border sentinel.
    /// `side` is top/logical-left/bottom/logical-right/between, before bidi.
    pub fn paragraph(bytes: &[u8], old: bool, side: usize) -> Result<Self, String> {
        let size = if old { 4 } else { 8 };
        if bytes.len() != size {
            return Err(unsupported("invalid Word paragraph border"));
        }
        // MS-DOC 2.9.157 NilBrc and 2.9.18 Brc80MayBeNil define explicit
        // no-border sentinels. Office paragraph operands can contain them too.
        // Normalize this known sentinel using its documented meaning before
        // interpreting ordinary Brc fields (or masking reserved/effect bits).
        // BrcOperand names Brc, not MayBeNil; sentinel acceptance here is
        // deliberate input recovery, not a claim that all Brc fields allow it.
        if u32_at(bytes, size - 4)? == u32::MAX {
            return Self::read(bytes, old);
        }
        // MS-DOC 2.9.17 excludes outset/inset from Brc80.
        if old && matches!(bytes[1], 0x1a | 0x1b) {
            return Err(unsupported("invalid Word paragraph Brc80 type"));
        }
        let mut value = [0u8; 8];
        value[..size].copy_from_slice(bytes);
        let flags = if old { 3 } else { 6 };
        // MS-DOC 2.9.16/17 explicitly make these effects inert on these
        // logical sides. Resolve before translating sides into physical OOXML.
        if side == 0 || side == 1 || (side == 4 && !old) {
            value[flags] &= !0x20;
        }
        if side == 0 || side == 1 || side == 4 {
            value[flags] &= !0x40;
        }
        Self::read(&value[..size], old)
    }

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
    fn paragraph_border_effects_follow_logical_side_and_record_version() {
        for old in [false, true] {
            let bytes = if old {
                vec![8, 1, 2, 0xff]
            } else {
                vec![0, 0, 255, 0, 8, 1, 0xff, 0xff]
            };
            for side in 0..5 {
                let xml = Border::paragraph(&bytes, old, side).unwrap().xml("test");
                let shadow = side == 2 || side == 3 || (old && side == 4);
                let frame = side == 2 || side == 3;
                assert!(xml.contains(&format!("w:shadow=\"{}\"", u8::from(shadow))));
                assert!(xml.contains(&format!("w:frame=\"{}\"", u8::from(frame))));
                assert!(xml.contains("w:space=\"31\""));
            }
        }
    }

    #[test]
    fn paragraph_brc_preserves_nil_sentinels_and_rejects_invalid_old_types() {
        for old in [false, true] {
            let size = if old { 4 } else { 8 };
            assert_eq!(
                Border::paragraph(&vec![255; size], old, 0)
                    .unwrap()
                    .xml("top"),
                "<w:top w:val=\"nil\"/>"
            );
            assert!(Border::read(&vec![255; size], old).is_ok());
            for length in 0..size {
                assert!(Border::paragraph(&vec![0; length], old, 0).is_err());
            }
        }
        // NilBrc.colorref is unused, irrespective of the bytes preceding it.
        assert_eq!(
            Border::paragraph(&[0, 1, 2, 3, 255, 255, 255, 255], false, 3)
                .unwrap()
                .xml("right"),
            "<w:right w:val=\"nil\"/>"
        );
        for kind in [0x1a, 0x1b] {
            assert!(Border::paragraph(&[8, kind, 0, 0], true, 0).is_err());
            assert!(Border::paragraph(&[0, 0, 0, 0, 8, kind, 0, 0], false, 0).is_ok());
        }
    }

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
