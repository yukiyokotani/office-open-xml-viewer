//! Feature-gated direct projection from typed MS-DOC section facts.
//!
//! This module deliberately does not round-trip through WordprocessingML.
//! Values follow MS-DOC 2.6.4 defaults and the existing `docx-model` section
//! contract. Unknown installation-dependent facts fail instead of borrowing
//! OOXML defaults or renderer heuristics.

use super::{Properties, Section};
use crate::doc::unsupported;
use docx_model::{
    ColSpec, ColumnsSpec, PageNumType, SectionGeom, SectionPageGeometryWire, SectionPlacementWire,
    SectionProps,
};

pub(in crate::doc) struct EndingSectionProjection {
    /// Decoded MS-DOC `sprmSBkc` fact. The caller selects the carrier according
    /// to the established `docx-model` section-boundary contract. MS-DOC calls
    /// this a break terminating a section, while ECMA-376 describes `w:type`
    /// relative to the preceding section; this projection does not infer or
    /// shift ownership between adjacent sections.
    pub(in crate::doc) kind: String,
    pub(in crate::doc) columns: Option<ColumnsSpec>,
    pub(in crate::doc) title_page: bool,
    pub(in crate::doc) geom: Box<SectionGeom>,
    pub(in crate::doc) page_num_type: Option<PageNumType>,
    pub(in crate::doc) text_direction: Option<String>,
    pub(in crate::doc) placement: Box<SectionPlacementWire>,
}

impl Section {
    pub(in crate::doc) fn project_final(
        &self,
        ordinal: usize,
        even_and_odd_headers: bool,
    ) -> Result<SectionProps, String> {
        let geom = self.direct_geometry()?;
        Ok(SectionProps {
            page_width: geom.page_width,
            page_height: geom.page_height,
            margin_top: geom.margin_top,
            margin_right: geom.margin_right,
            margin_bottom: geom.margin_bottom,
            margin_left: geom.margin_left,
            header_distance: geom.header_distance,
            footer_distance: geom.footer_distance,
            title_page: self.properties.title,
            even_and_odd_headers,
            section_start: Some(self.properties.break_kind().to_string()),
            text_direction: self.properties.text_flow.map(str::to_string),
            doc_grid_type: self.properties.grid_type().map(str::to_string),
            doc_grid_line_pitch: self.properties.grid_line_pitch(),
            doc_grid_char_space: (self.properties.grid != 0)
                .then_some(f64::from(self.properties.char_space)),
            columns: self.direct_columns()?,
            page_num_type: self.properties.page_num_type(),
            page_borders: None,
            line_numbering: None,
            v_align: self.properties.vertical_alignment().map(str::to_string),
            section_placement: Some(Box::new(self.direct_placement(ordinal)?)),
        })
    }

    pub(in crate::doc) fn project_ending(
        &self,
        ordinal: usize,
    ) -> Result<EndingSectionProjection, String> {
        Ok(EndingSectionProjection {
            kind: self.properties.break_kind().to_string(),
            columns: self.direct_columns()?,
            title_page: self.properties.title,
            geom: Box::new(self.direct_geometry()?),
            page_num_type: self.properties.page_num_type(),
            text_direction: self.properties.text_flow.map(str::to_string),
            placement: Box::new(self.direct_placement(ordinal)?),
        })
    }

    fn direct_geometry(&self) -> Result<SectionGeom, String> {
        let [top, right, bottom, left, header, footer] = self.properties.margins;
        let missing = || unsupported("unresolved Word section margin for direct model");
        Ok(SectionGeom {
            page_width: twips_to_pt(self.properties.size.0),
            page_height: twips_to_pt(self.properties.size.1),
            margin_top: twips_to_pt(top.ok_or_else(missing)?),
            margin_right: twips_to_pt(right.ok_or_else(missing)?),
            margin_bottom: twips_to_pt(bottom.ok_or_else(missing)?),
            margin_left: twips_to_pt(left.ok_or_else(missing)?),
            header_distance: twips_to_pt(header.ok_or_else(missing)?),
            footer_distance: twips_to_pt(footer.ok_or_else(missing)?),
        })
    }

    fn direct_columns(&self) -> Result<Option<ColumnsSpec>, String> {
        let count = usize::from(self.properties.columns);
        if count < 2 {
            return Ok(None);
        }
        let default_space = || {
            self.default_column_spacing
                .ok_or_else(|| unsupported("unresolved Word equal-column spacing for direct model"))
        };
        if self.properties.equal {
            let spacing = self
                .properties
                .spacing
                .map(Ok)
                .unwrap_or_else(default_space)?;
            return Ok(Some(ColumnsSpec {
                count,
                space_pt: twips_to_pt(spacing),
                equal_width: true,
                sep: self.properties.separator,
                cols: Vec::new(),
            }));
        }
        let mut cols = Vec::with_capacity(count);
        for index in 0..count {
            cols.push(ColSpec {
                width_pt: twips_to_pt(self.properties.widths[index].expect("validated")),
                space_pt: twips_to_pt(self.properties.spaces[index]),
            });
        }
        Ok(Some(ColumnsSpec {
            count,
            // Unused for explicit columns; zero avoids inventing a second source.
            space_pt: self.properties.spacing.map(twips_to_pt).unwrap_or(0.0),
            equal_width: false,
            sep: self.properties.separator,
            cols,
        }))
    }

    fn direct_placement(&self, ordinal: usize) -> Result<SectionPlacementWire, String> {
        let geom = self.direct_geometry()?;
        Ok(SectionPlacementWire {
            section_id: format!("section:{ordinal}"),
            section_bidi: self.properties.bidi,
            v_align: self.properties.vertical_alignment().map(str::to_string),
            line_numbering: None,
            doc_grid_type: self.properties.grid_type().map(str::to_string),
            doc_grid_line_pitch: self.properties.grid_line_pitch(),
            doc_grid_char_space: (self.properties.grid != 0)
                .then_some(f64::from(self.properties.char_space)),
            gutter_pt: Some(twips_to_pt(self.properties.gutter)),
            rtl_gutter: Some(self.properties.rtl_gutter),
            page_borders_authored: None,
            page_borders: None,
            page_geometry: Some(Box::new(SectionPageGeometryWire {
                page_width: Some(geom.page_width),
                page_height: Some(geom.page_height),
                margin_top: Some(geom.margin_top),
                margin_right: Some(geom.margin_right),
                margin_bottom: Some(geom.margin_bottom),
                margin_left: Some(geom.margin_left),
                header_distance: Some(geom.header_distance),
                footer_distance: Some(geom.footer_distance),
            })),
        })
    }
}

impl Properties {
    fn break_kind(&self) -> &'static str {
        [
            "continuous",
            "nextColumn",
            "nextPage",
            "evenPage",
            "oddPage",
        ][usize::from(self.kind)]
    }

    fn vertical_alignment(&self) -> Option<&'static str> {
        match self.vertical {
            0 => None,
            1 => Some("center"),
            2 => Some("both"),
            3 => Some("bottom"),
            _ => unreachable!("validated vertical alignment"),
        }
    }

    fn grid_type(&self) -> Option<&'static str> {
        match self.grid {
            0 => None,
            1 => Some("linesAndChars"),
            2 => Some("lines"),
            3 => Some("snapToChars"),
            _ => unreachable!("validated document grid"),
        }
    }

    fn grid_line_pitch(&self) -> Option<f64> {
        (self.grid != 0)
            .then(|| self.line_pitch.map(twips_to_pt))
            .flatten()
    }

    fn page_num_type(&self) -> Option<PageNumType> {
        (self.page_restart || self.page_format != "decimal").then(|| PageNumType {
            start: self.page_restart.then_some(i64::from(self.page_start)),
            fmt: Some(self.page_format.to_string()),
        })
    }
}

fn twips_to_pt(value: impl Into<f64>) -> f64 {
    value.into() / 20.0
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::Value;

    fn section(properties: Properties, default_column_spacing: Option<u16>) -> Section {
        Section {
            end: 1,
            incomplete_margins: false,
            properties,
            header_footer_references: [None; 6],
            default_column_spacing,
        }
    }

    fn complete_properties() -> Properties {
        let mut properties = Properties::parse(&[], &mut 100).unwrap();
        properties.margins = [
            Some(-1440),
            Some(1440),
            Some(2160),
            Some(2880),
            Some(720),
            Some(708),
        ];
        properties
    }

    #[test]
    fn projects_complete_final_geometry_and_placement_without_xml() {
        let mut properties = complete_properties();
        properties.gutter = 240;
        properties.title = true;
        properties.bidi = true;
        properties.rtl_gutter = false;
        let projected = section(properties, Some(720))
            .project_final(3, true)
            .unwrap();

        assert_eq!(
            (projected.page_width, projected.page_height),
            (612.0, 792.0)
        );
        assert_eq!(
            (
                projected.margin_top,
                projected.margin_right,
                projected.margin_bottom,
                projected.margin_left
            ),
            (-72.0, 72.0, 108.0, 144.0)
        );
        assert_eq!(
            (projected.header_distance, projected.footer_distance),
            (36.0, 35.4)
        );
        assert!(projected.title_page && projected.even_and_odd_headers);
        assert_eq!(projected.section_start.as_deref(), Some("nextPage"));
        let placement = projected.section_placement.unwrap();
        assert_eq!(placement.section_id, "section:3");
        assert!(placement.section_bidi);
        assert_eq!(placement.gutter_pt, Some(12.0));
        assert_eq!(placement.rtl_gutter, Some(false));
        assert_eq!(placement.page_geometry.unwrap().margin_top, Some(-72.0));
    }

    #[test]
    fn projects_every_break_grid_direction_and_numbering_variant() {
        for (kind, expected) in [
            (0, "continuous"),
            (1, "nextColumn"),
            (2, "nextPage"),
            (3, "evenPage"),
            (4, "oddPage"),
        ] {
            let mut properties = complete_properties();
            properties.kind = kind;
            assert_eq!(
                section(properties, Some(720))
                    .project_ending(0)
                    .unwrap()
                    .kind,
                expected
            );
        }
        for (grid, expected) in [(1, "linesAndChars"), (2, "lines"), (3, "snapToChars")] {
            let mut properties = complete_properties();
            properties.grid = grid;
            properties.line_pitch = Some(360);
            properties.char_space = -4096;
            let projected = section(properties, Some(720))
                .project_final(0, false)
                .unwrap();
            assert_eq!(projected.doc_grid_type.as_deref(), Some(expected));
            assert_eq!(projected.doc_grid_line_pitch, Some(18.0));
            assert_eq!(projected.doc_grid_char_space, Some(-4096.0));
        }
        for (vertical, expected) in [
            (0, None),
            (1, Some("center")),
            (2, Some("both")),
            (3, Some("bottom")),
        ] {
            let mut properties = complete_properties();
            properties.vertical = vertical;
            properties.text_flow = Some("tbRl");
            let projected = section(properties, Some(720)).project_ending(0).unwrap();
            assert_eq!(projected.placement.v_align.as_deref(), expected);
            assert_eq!(projected.text_direction.as_deref(), Some("tbRl"));
        }
        let mut properties = complete_properties();
        properties.page_restart = true;
        properties.page_start = 0;
        properties.page_format = "lowerRoman";
        let page = section(properties, Some(720))
            .project_final(0, false)
            .unwrap()
            .page_num_type
            .unwrap();
        assert_eq!(page.start, Some(0));
        assert_eq!(page.fmt.as_deref(), Some("lowerRoman"));
    }

    #[test]
    fn resolves_normative_column_defaults_and_preserves_explicit_zero() {
        for (default, expected) in [(720, 36.0), (708, 35.4), (1296, 64.8)] {
            let mut properties = complete_properties();
            properties.columns = 2;
            let columns = section(properties, Some(default))
                .direct_columns()
                .unwrap()
                .unwrap();
            assert_eq!(columns.space_pt, expected);
        }
        let mut properties = complete_properties();
        properties.columns = 2;
        properties.spacing = Some(0);
        let columns = section(properties, None).direct_columns().unwrap().unwrap();
        assert_eq!(columns.space_pt, 0.0);

        let mut properties = complete_properties();
        properties.columns = 2;
        properties.equal = false;
        properties.widths[0] = Some(718);
        properties.widths[1] = Some(31680);
        properties.spaces[0] = 0;
        properties.spaces[1] = 20;
        let columns = section(properties, None).direct_columns().unwrap().unwrap();
        assert!(!columns.equal_width);
        assert_eq!(columns.cols[0].width_pt, 35.9);
        assert_eq!(columns.cols[0].space_pt, 0.0);
        assert_eq!(columns.cols[1].width_pt, 1584.0);
        assert_eq!(columns.cols[1].space_pt, 1.0);

        let documented = [
            (1025, 720),
            (1026, 708),
            (1027, 708),
            (1028, 720),
            (1029, 708),
            (1030, 708),
            (1031, 720),
            (1032, 720),
            (1033, 720),
            (1034, 720),
            (1035, 708),
            (1036, 720),
            (1037, 720),
            (1038, 708),
            (1039, 708),
            (1040, 720),
            (1041, 720),
            (1042, 720),
            (1043, 708),
            (1044, 708),
            (1045, 708),
            (1046, 720),
            (1048, 708),
            (1049, 720),
            (1050, 720),
            (1051, 708),
            (1053, 720),
            (1055, 708),
            (1058, 720),
            (1059, 720),
            (1060, 708),
            (1061, 708),
            (1062, 720),
            (1063, 1296),
            (1067, 720),
            (1068, 720),
            (1069, 708),
            (1078, 708),
            (1079, 720),
            (1086, 720),
            (1087, 720),
            (1088, 708),
            (1089, 708),
            (1092, 720),
            (1104, 720),
            (2052, 720),
            (2070, 720),
            (2074, 708),
        ];
        for (lid, expected) in documented {
            assert_eq!(
                super::super::column_spacing_for_install_lid(lid),
                Some(expected),
                "LCID {lid}"
            );
        }
        assert_eq!(super::super::column_spacing_for_install_lid(9999), None);
    }

    #[test]
    fn unresolved_required_facts_fail_instead_of_guessing() {
        for index in 0..6 {
            let mut properties = complete_properties();
            properties.margins[index] = None;
            assert!(section(properties, Some(720))
                .project_final(0, false)
                .is_err());
        }
        let mut properties = complete_properties();
        properties.columns = 2;
        let result = section(properties, None).project_final(0, false);
        assert!(result.is_err());
    }

    #[test]
    fn dormant_line_pitch_does_not_leak_when_grid_is_disabled() {
        let mut properties = complete_properties();
        properties.line_pitch = Some(360);
        properties.char_space = 4096;
        let projected = section(properties, Some(720))
            .project_final(0, false)
            .unwrap();
        assert_eq!(projected.doc_grid_type, None);
        assert_eq!(projected.doc_grid_line_pitch, None);
        assert_eq!(projected.doc_grid_char_space, None);
        let placement = projected.section_placement.unwrap();
        assert_eq!(placement.doc_grid_type, None);
        assert_eq!(placement.doc_grid_line_pitch, None);
        assert_eq!(placement.doc_grid_char_space, None);
    }

    fn xml_projection(properties: &Properties) -> Value {
        let section = properties.xml().unwrap();
        let document = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/>{section}</w:body></w:document>"#
        );
        let parts = vec![("word/document.xml".to_string(), document)];
        let package = crate::ooxml::write_package(&parts, 1_000_000).unwrap();
        let json = docx_parser::parse_docx_native(&package).unwrap();
        serde_json::from_str::<Value>(&json).unwrap()["section"].clone()
    }

    #[test]
    fn direct_final_matches_existing_xml_projection_for_representable_facts() {
        let mut properties = complete_properties();
        properties.kind = 4;
        properties.gutter = 240;
        properties.columns = 2;
        properties.spacing = Some(0);
        properties.separator = true;
        properties.vertical = 2;
        properties.title = true;
        properties.bidi = true;
        properties.grid = 2;
        properties.line_pitch = Some(360);
        properties.char_space = -4096;
        properties.text_flow = Some("tbRl");
        properties.page_restart = true;
        properties.page_start = 7;
        properties.page_format = "lowerRoman";

        let xml = xml_projection(&properties);
        let direct = serde_json::to_value(
            section(properties, Some(720))
                .project_final(0, false)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(direct, xml);
    }

    #[test]
    fn direct_grid_off_with_dormant_pitch_matches_existing_xml_projection() {
        let mut properties = complete_properties();
        properties.line_pitch = Some(360);
        properties.char_space = 4096;
        let xml = xml_projection(&properties);
        let direct = serde_json::to_value(
            section(properties, Some(720))
                .project_final(0, false)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(direct, xml);
    }

    fn read_fixture(lid: u16, properties: &[u8]) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0; 512];
        word[6..8].copy_from_slice(&lid.to_le_bytes());
        word[0xca..0xce].copy_from_slice(&0u32.to_le_bytes());
        word[0xce..0xd2].copy_from_slice(&20u32.to_le_bytes());
        word[300..302].copy_from_slice(&(properties.len() as u16).to_le_bytes());
        word[302..302 + properties.len()].copy_from_slice(properties);
        let mut table = vec![0; 20];
        table[4..8].copy_from_slice(&3u32.to_le_bytes());
        table[10..14].copy_from_slice(&300u32.to_le_bytes());
        (word, table)
    }

    fn prl(sprm: u16, value: i16) -> Vec<u8> {
        [sprm.to_le_bytes().to_vec(), value.to_le_bytes().to_vec()].concat()
    }

    #[test]
    fn full_read_retains_lcid_default_and_explicit_column_spacing_priority() {
        let required_margins = [
            prl(0x9023, 1440),
            prl(0xb021, 1440),
            prl(0x9024, 1440),
            prl(0xb022, 1440),
            prl(0x500b, 1),
        ]
        .concat();
        let (word, table) = read_fixture(1063, &required_margins);
        let sections = super::super::read(&word, &table, 3).unwrap();
        assert_eq!(
            sections[0].direct_columns().unwrap().unwrap().space_pt,
            64.8
        );

        let explicit = [required_margins, prl(0x900c, 0)].concat();
        let (word, table) = read_fixture(9999, &explicit);
        let sections = super::super::read(&word, &table, 3).unwrap();
        assert_eq!(sections[0].direct_columns().unwrap().unwrap().space_pt, 0.0);
    }
}
