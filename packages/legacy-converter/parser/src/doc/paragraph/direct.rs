//! Typed projection of already-resolved binary DOC paragraph properties.
//!
//! Runs, paragraph-mark formatting, numbering activation and source identity
//! remain owned by their acquisition layers; this module projects layout facts.

use super::Properties;
use docx_model::{
    DocParagraph, LineSpacing, ParagraphBorders, ParagraphTypographyBordersWire,
    ParagraphTypographyWire,
};

impl Properties {
    pub(in crate::doc) fn direct_paragraph(&self) -> DocParagraph {
        let [left, right] = self.logical_indents();
        let first = self
            .protected_list_indent
            .map_or(self.first, |(_, first)| first);
        let line_value = match self.line.1 {
            "auto" => f64::from(self.line.0) / 240.0,
            _ => f64::from(self.line.0) / 20.0,
        };
        let bidi = self.flags.get("bidi").copied();
        let borders = self.direct_borders(bidi == Some(true));
        let paragraph_typography_acquisition = self.direct_border_typography(bidi == Some(true));

        DocParagraph {
            paragraph_id: None,
            alignment: self.normalized_model_alignment().to_string(),
            indent_left: f64::from(left) / 20.0,
            indent_right: f64::from(right) / 20.0,
            indent_first: f64::from(first) / 20.0,
            space_before: f64::from(self.before) / 20.0,
            space_after: f64::from(self.after) / 20.0,
            line_spacing: Some(LineSpacing {
                value: line_value,
                rule: self.line.1.to_string(),
                explicit: true,
            }),
            numbering: None,
            tab_stops: self.tabs.direct(),
            runs: Vec::new(),
            run_revisions: Vec::new(),
            complex_field_boundaries: Vec::new(),
            bookmarks: Vec::new(),
            comment_marks: Vec::new(),
            shading: None,
            page_break_before: self.flag("pageBreakBefore", false),
            contextual_spacing: false,
            keep_next: self.flag("keepNext", false),
            keep_lines: self.flag("keepLines", false),
            mark_vanish: false,
            widow_control: self.flag("widowControl", true),
            overflow_punct: self.flag("overflowPunct", true),
            adjust_right_ind: self.flag("adjustRightInd", true),
            borders,
            style_id: None,
            default_font_size: None,
            default_font_family: None,
            default_font_family_east_asia: None,
            paragraph_mark_font_facts: None,
            paragraph_mark_color: None,
            outline_level: None,
            bidi,
            snap_to_grid: self.flags.get("snapToGrid").copied(),
            frame_pr: None,
            paragraph_typography_acquisition,
        }
    }

    fn flag(&self, name: &str, default: bool) -> bool {
        self.flags.get(name).copied().unwrap_or(default)
    }

    fn direct_borders(&self, bidi: bool) -> Option<ParagraphBorders> {
        let [top, logical_left, bottom, logical_right, between] = &self.borders;
        let (left, right) = if bidi {
            (logical_right, logical_left)
        } else {
            (logical_left, logical_right)
        };
        let borders = ParagraphBorders {
            top: top.as_ref().map(|edge| edge.direct_edge()),
            bottom: bottom.as_ref().map(|edge| edge.direct_edge()),
            left: left.as_ref().map(|edge| edge.direct_edge()),
            right: right.as_ref().map(|edge| edge.direct_edge()),
            between: between.as_ref().map(|edge| edge.direct_edge()),
        };
        (borders.top.is_some()
            || borders.bottom.is_some()
            || borders.left.is_some()
            || borders.right.is_some()
            || borders.between.is_some())
        .then_some(borders)
    }

    fn direct_border_typography(&self, bidi: bool) -> Option<ParagraphTypographyWire> {
        let [top, logical_left, bottom, logical_right, between] = &self.borders;
        let (left, right) = if bidi {
            (logical_right, logical_left)
        } else {
            (logical_left, logical_right)
        };
        let borders = ParagraphTypographyBordersWire {
            top: top.as_ref().map(|edge| edge.direct_typography()),
            right: right.as_ref().map(|edge| edge.direct_typography()),
            bottom: bottom.as_ref().map(|edge| edge.direct_typography()),
            left: left.as_ref().map(|edge| edge.direct_typography()),
            between: between.as_ref().map(|edge| edge.direct_typography()),
            bar: None,
        };
        (borders.top.is_some()
            || borders.right.is_some()
            || borders.bottom.is_some()
            || borders.left.is_some()
            || borders.between.is_some())
        .then_some(ParagraphTypographyWire { borders })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    fn parsed(properties: &Properties) -> serde_json::Value {
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr>{}</w:pPr></w:p></w:body></w:document>"#,
            properties.xml()
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

    fn assert_property_parity(properties: &Properties) {
        let mut expected = parsed(properties);
        // Owned by later direct acquisition seams, not paragraph properties.
        let object = expected.as_object_mut().unwrap();
        for key in [
            "type",
            "styleId",
            "defaultFontSize",
            "defaultFontFamily",
            "defaultFontFamilyEastAsia",
            "paragraphMarkFontFacts",
            "paragraphMarkColor",
        ] {
            object.remove(key);
        }
        assert_eq!(
            serde_json::to_value(properties.direct_paragraph()).unwrap(),
            expected
        );
    }

    #[test]
    fn full_resolved_property_projection_matches_docx_parser_semantics() {
        let mut properties = Properties::default();
        for (code, operand) in [
            (0x2405, vec![1]),
            (0x2406, vec![1]),
            (0x2407, vec![1]),
            (0x2431, vec![0]),
            (0x2435, vec![0]),
            (0x2441, vec![1]),
            (0x2447, vec![0]),
            (0x2448, vec![0]),
            (0xa413, 240u16.to_le_bytes().to_vec()),
            (0xa414, 480u16.to_le_bytes().to_vec()),
            (
                0x6412,
                [(-360i16).to_le_bytes(), 0u16.to_le_bytes()].concat(),
            ),
            (0x840f, 720i16.to_le_bytes().to_vec()),
            (0x840e, 360i16.to_le_bytes().to_vec()),
            (0x8411, (-240i16).to_le_bytes().to_vec()),
            (0x2403, vec![0]),
            (0xc60d, vec![5, 0, 1, 0xd0, 2, 9]),
            (0x6424, vec![8, 3, 2, 0]),
            (0x6425, vec![16, 1, 6, 0]),
            (0x6427, vec![24, 7, 1, 0]),
            (0x6428, vec![8, 6, 0, 0]),
        ] {
            assert!(properties.apply(code, &operand).unwrap());
        }
        assert_property_parity(&properties);
    }

    #[test]
    fn defaults_and_line_rule_units_match_docx_parser_semantics() {
        assert_property_parity(&Properties::default());
        for (line, multiple) in [(360i16, 1u16), (360, 0), (-360, 0)] {
            let mut properties = Properties::default();
            properties
                .apply(
                    0x6412,
                    &[line.to_le_bytes(), multiple.to_le_bytes()].concat(),
                )
                .unwrap();
            assert_property_parity(&properties);
        }
    }

    #[test]
    fn parser_model_limits_for_extended_spacing_indent_and_alignment_are_explicit() {
        let mut properties = Properties::default();
        for (code, operand) in [
            (0x4458, 120i16.to_le_bytes().to_vec()),
            (0x4459, (-20i16).to_le_bytes().to_vec()),
            (0x245b, vec![1]),
            (0x245c, vec![1]),
            (0x4455, 120u16.to_le_bytes().to_vec()),
            (0x4456, 240u16.to_le_bytes().to_vec()),
            (0x4457, 360u16.to_le_bytes().to_vec()),
            (0x4439, 2u16.to_le_bytes().to_vec()),
        ] {
            assert!(properties.apply(code, &operand).unwrap());
        }
        // The current DOCX model has no fields for line-unit/automatic spacing,
        // character-unit indents, or textAlignment. The existing DOCX parser
        // also omits them, so this projection preserves model parity rather
        // than inventing renderer-facing equivalents.
        assert_property_parity(&properties);

        for alignment in 0..=9 {
            let mut properties = Properties::default();
            assert!(properties.apply(0x2461, &[alignment]).unwrap());
            assert_property_parity(&properties);
        }
        for alignment in 0..=5 {
            let mut properties = Properties::default();
            assert!(properties.apply(0x2403, &[alignment]).unwrap());
            assert_property_parity(&properties);
            assert!(properties.apply(0x2441, &[1]).unwrap());
            assert_property_parity(&properties);
        }
    }

    #[test]
    fn protected_negative_list_indent_has_existing_adapter_precedence() {
        let mut properties = Properties::default();
        properties.ilfo = -1;
        properties.first = 120;
        properties.protected_list_indent = Some((720, -360));
        let direct = properties.direct_paragraph();
        assert_eq!(direct.indent_left, 36.0);
        assert_eq!(direct.indent_first, -18.0);
        assert_property_parity(&properties);
    }

    #[test]
    fn border_binary_boundaries_match_docx_parser_semantics() {
        for operand in [
            vec![8, 0xAB, 0xCD, 0xEF, 0, 255, 1, 31, 0],
            vec![8, 0, 0, 0, 0xff, 255, 1, 31, 0],
            vec![8, 0, 0, 0, 0, 8, 0, 0, 0],
            vec![8, 255, 255, 255, 255, 255, 255, 255, 255],
        ] {
            let mut properties = Properties::default();
            assert!(properties.apply(0xc64e, &operand).unwrap());
            assert_property_parity(&properties);
        }
    }

    #[test]
    fn every_valid_tab_alignment_and_leader_matches_docx_parser() {
        let alignments = [0u8, 1, 2, 3, 4, 6];
        let leaders = [0u8, 1, 2, 3, 4, 5, 7];
        let descriptors: Vec<_> = alignments
            .into_iter()
            .flat_map(|alignment| {
                leaders
                    .into_iter()
                    .map(move |leader| alignment | (leader << 3))
            })
            .collect();
        let positions: Vec<i16> = (0..descriptors.len())
            .map(|index| -1_000 + index as i16 * 40)
            .collect();
        let mut operand = vec![0, 0, descriptors.len() as u8];
        for position in &positions {
            operand.extend(position.to_le_bytes());
        }
        operand.extend(&descriptors);
        operand[0] = (operand.len() - 1) as u8;

        let mut properties = Properties::default();
        assert!(properties.apply(0xc60d, &operand).unwrap());
        assert_property_parity(&properties);
        assert!(properties
            .direct_paragraph()
            .tab_stops
            .first()
            .is_some_and(|tab| tab.pos < 0.0));
    }

    #[test]
    fn old_and_modern_border_sides_and_bidi_swaps_match_docx_parser() {
        for bidi in [false, true] {
            for side in 0u16..5 {
                for modern in [false, true] {
                    let mut properties = Properties::default();
                    if bidi {
                        assert!(properties.apply(0x2441, &[1]).unwrap());
                    }
                    let (code, operand) = if modern {
                        (0xc64e + side, vec![8, 0x12, 0x34, 0x56, 0, 255, 1, 0x7f, 0])
                    } else {
                        (0x6424 + side, vec![255, 1, 6, 0x7f])
                    };
                    assert!(properties.apply(code, &operand).unwrap());
                    assert_property_parity(&properties);
                }
            }
        }
    }
}
