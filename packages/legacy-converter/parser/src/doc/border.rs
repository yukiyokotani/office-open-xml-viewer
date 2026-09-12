//! [MS-DOC] Brc/Brc80/BrcType/Ico to ECMA-376 ST_Border, no pixel fitting.
use super::{u32_at, unsupported};

// MS-DOC 2.9.119 Ico: the fixed palette is shared by Brc80 and Shd80.
pub(super) const ICO_COLORS: [&str; 17] = [
    "auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF",
    "000080", "008080", "008000", "800080", "800080", "808000", "808080", "C0C0C0",
];

#[derive(Clone, Default)]
pub struct Border {
    facts: Option<BorderFacts>,
}

#[derive(Clone)]
struct BorderFacts {
    style: String,
    color: Option<String>,
    width_eighth_points: Option<u8>,
    space_points: Option<u8>,
    shadow: Option<bool>,
    frame: Option<bool>,
}
impl Border {
    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn retained_bytes(&self) -> Result<usize, String> {
        let Some(facts) = &self.facts else {
            return Ok(0);
        };
        facts
            .style
            .capacity()
            .checked_add(facts.color.as_ref().map_or(0, String::capacity))
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())
    }

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
                facts: Some(BorderFacts {
                    style: "nil".into(),
                    color: None,
                    width_eighth_points: None,
                    space_points: None,
                    shadow: None,
                    frame: None,
                }),
            });
        }
        let (width, kind, color, flags) = if old {
            // Ico is a fixed binary-format palette, not a document theme.
            (
                b[0],
                b[1],
                ICO_COLORS
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
        let width = width.max(2);
        let space = flags & 31;
        let shadow = flags & 32 != 0;
        let frame = flags & 64 != 0;
        Ok(Self {
            facts: Some(BorderFacts {
                style: style.into(),
                color: Some(color),
                width_eighth_points: Some(width),
                space_points: Some(space),
                shadow: Some(shadow),
                frame: Some(frame),
            }),
        })
    }
    pub fn xml(&self, side: &str) -> String {
        let Some(value) = &self.facts else {
            return format!("<w:{side} />");
        };
        if value.style == "nil" {
            return format!("<w:{side} w:val=\"nil\"/>");
        }
        format!(
            "<w:{side} w:val=\"{}\" w:sz=\"{}\" w:color=\"{}\" w:space=\"{}\" w:shadow=\"{}\" w:frame=\"{}\"/>",
            value.style,
            value.width_eighth_points.expect("decoded border width"),
            value.color.as_deref().expect("decoded border color"),
            value.space_points.expect("decoded border space"),
            u8::from(value.shadow.expect("decoded border shadow")),
            u8::from(value.frame.expect("decoded border frame")),
        )
    }

    /// Table/cell border projection matching the current DOCX parser's
    /// `BorderSpec` contract. Spacing, shadow and frame remain available to
    /// paragraph typography, but `BorderSpec` has no fields for them.
    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn direct_spec(&self) -> docx_model::BorderSpec {
        let Some(value) = &self.facts else {
            return docx_model::BorderSpec {
                width: 0.5,
                color: None,
                style: "none".into(),
            };
        };
        docx_model::BorderSpec {
            width: value
                .width_eighth_points
                .map_or(0.5, |width| f64::from(width) / 8.0),
            color: value
                .color
                .as_deref()
                .filter(|color| *color != "auto")
                .map(str::to_ascii_lowercase),
            style: value.style.clone(),
        }
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn direct_edge(&self) -> docx_model::ParaBorderEdge {
        let Some(value) = &self.facts else {
            return docx_model::ParaBorderEdge {
                style: "none".into(),
                color: None,
                width: 0.5,
                space: 1.0,
            };
        };
        let cleared = matches!(value.style.as_str(), "none" | "nil");
        docx_model::ParaBorderEdge {
            style: if cleared {
                "none".into()
            } else {
                value.style.clone()
            },
            color: (!cleared)
                .then(|| value.color.as_deref())
                .flatten()
                .filter(|color| *color != "auto")
                .map(str::to_ascii_lowercase),
            width: if cleared {
                0.0
            } else {
                value
                    .width_eighth_points
                    .map_or(0.5, |v| f64::from(v) / 8.0)
            },
            space: if cleared {
                0.0
            } else {
                value.space_points.map_or(1.0, f64::from)
            },
        }
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn direct_typography(&self) -> docx_model::CtBorderTypographyWire {
        use docx_model::{TypographyValueStatusWire::Valid, TypographyValueWire};
        let Some(value) = &self.facts else {
            return docx_model::CtBorderTypographyWire::default();
        };
        let string = |raw: String, normalized: String| TypographyValueWire {
            status: Valid,
            raw: Some(raw),
            value: Some(normalized),
        };
        let number = |raw: u8, divisor: f64| TypographyValueWire {
            status: Valid,
            raw: Some(raw.to_string()),
            value: Some(f64::from(raw) / divisor),
        };
        let boolean = |raw: bool| TypographyValueWire {
            status: Valid,
            raw: Some(u8::from(raw).to_string()),
            value: Some(raw),
        };
        docx_model::CtBorderTypographyWire {
            val: string(value.style.clone(), value.style.clone()),
            color: value
                .color
                .as_ref()
                .map(|raw| string(raw.clone(), raw.to_ascii_lowercase()))
                .unwrap_or_default(),
            size_pt: value
                .width_eighth_points
                .map(|raw| number(raw, 8.0))
                .unwrap_or_default(),
            space_pt: value
                .space_points
                .map(|raw| number(raw, 1.0))
                .unwrap_or_default(),
            shadow: value.shadow.map(boolean).unwrap_or_default(),
            frame: value.frame.map(boolean).unwrap_or_default(),
            ..docx_model::CtBorderTypographyWire::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(feature = "direct-doc")]
    use std::io::{Cursor, Write};
    #[cfg(feature = "direct-doc")]
    use zip::write::SimpleFileOptions;

    #[cfg(feature = "direct-doc")]
    fn parsed_table(xml: &str) -> serde_json::Value {
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr>{xml}</w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
        );
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            archive
                .start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            archive.write_all(document.as_bytes()).unwrap();
            archive.finish().unwrap();
        }
        let json: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&bytes).unwrap()).unwrap();
        json["body"][0].clone()
    }
    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_table_border_matches_docx_parser_contract_at_boundaries() {
        let default = Border::default().direct_spec();
        assert_eq!(
            (default.style.as_str(), default.width, default.color),
            ("none", 0.5, None)
        );
        let nil = Border::read(&[255; 8], false).unwrap().direct_spec();
        assert_eq!(
            (nil.style.as_str(), nil.width, nil.color),
            ("nil", 0.5, None)
        );
        let value = Border::read(&[0xAB, 0xCD, 0xEF, 0, 255, 27, 31, 0], false)
            .unwrap()
            .direct_spec();
        assert_eq!(
            (value.style.as_str(), value.width, value.color.as_deref()),
            ("inset", 31.875, Some("abcdef"))
        );
        let parsed = parsed_table(&format!(
            "<w:tblBorders>{}</w:tblBorders>",
            Border::read(&[0xAB, 0xCD, 0xEF, 0, 255, 27, 31, 0], false)
                .unwrap()
                .xml("top")
        ));
        assert_eq!(
            serde_json::to_value(value).unwrap(),
            parsed["borders"]["top"]
        );
    }
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

    #[test]
    fn typed_storage_preserves_exact_xml_at_binary_boundaries() {
        assert_eq!(Border::default().xml("top"), "<w:top />");
        assert_eq!(
            Border::read(&[0xAB, 0xCD, 0xEF, 0, 255, 1, 31, 0], false)
                .unwrap()
                .xml("bottom"),
            "<w:bottom w:val=\"single\" w:sz=\"255\" w:color=\"ABCDEF\" w:space=\"31\" w:shadow=\"0\" w:frame=\"0\"/>"
        );
        assert_eq!(
            Border::read(&[0, 0, 0, 255, 255, 1, 31, 0], false)
                .unwrap()
                .xml("bottom"),
            "<w:bottom w:val=\"single\" w:sz=\"255\" w:color=\"auto\" w:space=\"31\" w:shadow=\"0\" w:frame=\"0\"/>"
        );
        assert_eq!(
            Border::read(&[0, 0, 0, 0, 8, 0, 0, 0], false)
                .unwrap()
                .xml("bottom"),
            "<w:bottom w:val=\"none\" w:sz=\"8\" w:color=\"000000\" w:space=\"0\" w:shadow=\"0\" w:frame=\"0\"/>"
        );
        assert_eq!(
            Border::read(&[255; 8], false).unwrap().xml("bottom"),
            "<w:bottom w:val=\"nil\"/>"
        );
    }
}
