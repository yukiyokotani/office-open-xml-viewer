//! Typed DOC character projection for the direct DOC model path.
//!
//! The binary decoder has already resolved style inheritance and validated every
//! value stored by [`Properties`]. This module projects that state without
//! manufacturing WordprocessingML and parsing it back through the DOCX parser.

use super::Properties;
use crate::doc::unsupported;
use docx_model::{
    RunFontAxisPresence, RunFontAxisValues, RunFontFacts, RunFontSlots, RunTypographyWire, TextRun,
    TypographyValueStatusWire, TypographyValueWire, UnderlineTypographyWire,
};

type FontAxes = [Option<String>; 4];

impl Properties {
    pub(in crate::doc) fn direct_font_size_pt(&self) -> Result<f64, String> {
        // Resolved properties already include the MS-DOC 2.6.1 sprmCHps
        // default (20 half-points). A sparse patch is not a resolved font.
        self.half_points("sz")?
            .ok_or_else(|| unsupported("Word resolved font size is absent"))
    }

    /// Project one visible text span. Hidden text is absent from the normal/print
    /// DOC model, matching the existing DOC-to-OOXML-to-DOCX-parser route.
    pub(in crate::doc) fn direct_text_run(
        &self,
        text: String,
        fonts: &[String],
    ) -> Result<Option<TextRun>, String> {
        // Validate referenced fonts before applying visibility, matching the
        // legacy XML construction step that precedes DOCX run filtering.
        let axes = self.direct_font_axes(fonts)?;
        if self.bool_value("vanish").unwrap_or(false) {
            return Ok(None);
        }
        let underline_token = self.values.get("u").map(String::as_str);
        let underline = underline_token.is_some_and(|value| value != "none");
        let underline_color = underline
            .then(|| {
                self.values
                    .get("uColor")
                    .map(|value| value.to_ascii_lowercase())
            })
            .flatten();
        let color_token = self.values.get("color").map(String::as_str);
        let color_auto = color_token == Some("auto");
        let vertical_token = self.values.get("vertAlign").cloned();
        let position_raw = self.values.get("position").cloned();
        let font_size = self.half_points("sz")?.unwrap_or(10.0);
        let font_size_cs = self.half_points("szCs")?.or(Some(font_size));
        let char_spacing = self.signed_twentieth_points("spacing")?;
        let char_scale = self.unsigned_number("w")?.map(|value| value / 100.0);
        let position = self.signed_number("position")?.map(|value| value / 2.0);
        let kerning = self.half_points("kern")?;
        let highlight = self
            .values
            .get("highlight")
            .filter(|value| value.as_str() != "none")
            .cloned();

        let mut run = TextRun {
            text,
            bold: self.bool_value("b").unwrap_or(false),
            italic: self.bool_value("i").unwrap_or(false),
            underline,
            underline_style: underline_token
                .filter(|value| !matches!(*value, "none" | "single"))
                .map(str::to_string),
            underline_color: underline_color.clone(),
            strikethrough: self.bool_value("strike").unwrap_or(false),
            font_size,
            color: color_token
                .filter(|value| *value != "auto")
                .map(str::to_ascii_lowercase),
            font_family: axes[0].clone().or_else(|| axes[1].clone()),
            font_family_east_asia: axes[1].clone(),
            font_family_high_ansi: axes[2].clone(),
            font_family_cs: axes[3].clone(),
            font_slots: font_slots(&axes),
            font_hint: self.font_hint.map(|hint| hint.xml_value().to_string()),
            color_auto,
            vert_align: vertical_token.as_deref().and_then(|value| match value {
                "superscript" => Some("super".to_string()),
                "subscript" => Some("sub".to_string()),
                _ => None,
            }),
            all_caps: self.bool_value("caps").unwrap_or(false),
            small_caps: self.bool_value("smallCaps").unwrap_or(false),
            double_strikethrough: self.bool_value("dstrike").unwrap_or(false),
            rtl: self.bool_value("rtl"),
            cs: self.bool_value("cs"),
            font_size_cs,
            bold_cs: self.bool_value("bCs"),
            italic_cs: self.bool_value("iCs"),
            char_spacing,
            char_scale,
            position,
            kerning,
            highlight,
            ..TextRun::default()
        };
        run.typography_acquisition = Some(RunTypographyWire {
            underline: match (
                underline_token,
                self.values.get("uColor").map(String::as_str),
            ) {
                (None, None) => None,
                // MS-DOC 2.6.1 sprmCCvUl and MS-OI29500 2.1.100(c): retain a
                // color-only underline acquisition, while the public run stays
                // un-underlined until an effective underline value exists.
                (token, color) => Some(underline_wire(token, color)),
            },
            strike: run.strikethrough,
            double_strike: run.double_strikethrough,
            caps: run.all_caps,
            small_caps: run.small_caps,
            color_auto,
            vertical_align: match (vertical_token, run.vert_align.clone()) {
                (Some(raw), Some(value)) => valid_wire(raw, value),
                // A valid DOC baseline clears super/sub. The current normalized
                // wire only represents those two values, so retain `baseline`
                // as raw Invalid for compatibility with the DOCX parser contract.
                (Some(raw), None) => invalid_wire(raw),
                _ => TypographyValueWire::default(),
            },
            position_pt: number_wire(position_raw, run.position),
            character_spacing_pt: run.char_spacing,
            character_scale: run.char_scale,
            kerning_threshold_pt: run.kerning,
            ..RunTypographyWire::default()
        });
        Ok(Some(run))
    }

    /// Effective font facts for paragraph marks and numbering markers. These are
    /// projected independently of visible text and therefore remain available
    /// even when the mark itself is hidden.
    pub(in crate::doc) fn direct_font_facts(
        &self,
        fonts: &[String],
    ) -> Result<RunFontFacts, String> {
        let axes = self.direct_font_axes(fonts)?;
        let font_size = self.half_points("sz")?;
        Ok(RunFontFacts {
            font_family: axes[0].clone().or_else(|| axes[1].clone()),
            font_family_high_ansi: axes[2].clone(),
            font_slots: font_slots(&axes),
            font_family_east_asia: axes[1].clone(),
            font_hint: self.font_hint.map(|hint| hint.xml_value().to_string()),
            rtl: self.bool_value("rtl"),
            cs: self.bool_value("cs"),
            font_family_cs: axes[3].clone(),
            font_size,
            font_size_cs: self.half_points("szCs")?.or(font_size),
            bold: self.bool_value("b").unwrap_or(false),
            italic: self.bool_value("i").unwrap_or(false),
            bold_cs: self.bool_value("bCs"),
            italic_cs: self.bool_value("iCs"),
            kerning: self.half_points("kern")?,
            ..RunFontFacts::default()
        })
    }

    pub(in crate::doc) fn direct_vanish(&self) -> bool {
        self.bool_value("vanish").unwrap_or(false)
    }

    pub(in crate::doc) fn direct_color(&self) -> Option<String> {
        self.values
            .get("color")
            .filter(|value| value.as_str() != "auto")
            .map(|value| value.to_ascii_lowercase())
    }

    fn direct_font_axes(&self, fonts: &[String]) -> Result<FontAxes, String> {
        if fonts.is_empty() {
            if self.fonts.iter().flatten().any(|index| *index != 0) {
                return Err(unsupported("Word font index outside empty font table"));
            }
            return Ok([None, None, None, None]);
        }
        let mut axes: FontAxes = [None, None, None, None];
        for (axis, index) in axes.iter_mut().zip(self.fonts) {
            if let Some(index) = index {
                *axis = Some(
                    fonts
                        .get(index)
                        .ok_or_else(|| unsupported("Word font index outside font table"))?
                        .clone(),
                );
            }
        }
        Ok(axes)
    }

    fn bool_value(&self, key: &str) -> Option<bool> {
        self.values.get(key).map(|value| value == "1")
    }

    fn unsigned_number(&self, key: &str) -> Result<Option<f64>, String> {
        self.values
            .get(key)
            .map(|value| {
                value
                    .parse::<u16>()
                    .map(f64::from)
                    .map_err(|_| unsupported("invalid resolved Word character number"))
            })
            .transpose()
    }

    fn signed_number(&self, key: &str) -> Result<Option<f64>, String> {
        self.values
            .get(key)
            .map(|value| {
                value
                    .parse::<i16>()
                    .map(f64::from)
                    .map_err(|_| unsupported("invalid resolved Word character number"))
            })
            .transpose()
    }

    fn half_points(&self, key: &str) -> Result<Option<f64>, String> {
        Ok(self.unsigned_number(key)?.map(|value| value / 2.0))
    }

    fn signed_twentieth_points(&self, key: &str) -> Result<Option<f64>, String> {
        Ok(self.signed_number(key)?.map(|value| value / 20.0))
    }
}

fn font_slots(axes: &FontAxes) -> Option<RunFontSlots> {
    axes.iter().any(Option::is_some).then(|| RunFontSlots {
        direct: RunFontAxisValues {
            ascii: axes[0].clone(),
            east_asia: axes[1].clone(),
            high_ansi: axes[2].clone(),
            complex_script: axes[3].clone(),
        },
        theme: RunFontAxisValues::default(),
        theme_present: RunFontAxisPresence::default(),
    })
}

fn valid_wire<T>(raw: String, value: T) -> TypographyValueWire<T> {
    TypographyValueWire {
        status: TypographyValueStatusWire::Valid,
        raw: Some(raw),
        value: Some(value),
    }
}

fn invalid_wire<T>(raw: String) -> TypographyValueWire<T> {
    TypographyValueWire {
        status: TypographyValueStatusWire::Invalid,
        raw: Some(raw),
        value: None,
    }
}

fn number_wire(raw: Option<String>, value: Option<f64>) -> TypographyValueWire<f64> {
    match (raw, value) {
        (Some(raw), Some(value)) => valid_wire(raw, value),
        _ => TypographyValueWire::default(),
    }
}

fn underline_wire(token: Option<&str>, color: Option<&str>) -> UnderlineTypographyWire {
    UnderlineTypographyWire {
        val: token.map_or_else(TypographyValueWire::default, |value| {
            valid_wire(value.to_string(), value.to_string())
        }),
        color: color.map_or_else(TypographyValueWire::default, |value| {
            valid_wire(value.to_string(), value.to_ascii_lowercase())
        }),
        ..UnderlineTypographyWire::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::border::ICO_COLORS;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    fn applied(entries: &[(u16, Vec<u8>)]) -> Properties {
        let base = Properties::default();
        let mut properties = base.clone();
        for (code, operand) in entries {
            assert!(properties.apply(*code, operand, &base).unwrap());
        }
        properties
    }

    fn parsed_run(properties: &Properties, fonts: &[String]) -> serde_json::Value {
        let document_xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>{}<w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
            properties.xml(fonts).unwrap(),
        );
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            archive
                .start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            archive.write_all(document_xml.as_bytes()).unwrap();
            archive.finish().unwrap();
        }
        let parsed: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&bytes).unwrap()).unwrap();
        let mut run = parsed["body"][0]["runs"][0].clone();
        run.as_object_mut().unwrap().remove("type");
        run
    }

    fn assert_parser_parity(properties: &Properties, fonts: &[String]) {
        let direct = serde_json::to_value(
            properties
                .direct_text_run("x".into(), fonts)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(direct, parsed_run(properties, fonts));
    }

    #[test]
    fn projects_units_flags_color_and_private_typography_without_ids() {
        let properties = applied(&[
            (0x0835, vec![1]),
            (0x0836, vec![1]),
            (0x0837, vec![1]),
            (0x2a53, vec![1]),
            (0x083a, vec![1]),
            (0x083b, vec![1]),
            (0x085a, vec![1]),
            (0x0882, vec![1]),
            (0x085c, vec![1]),
            (0x085d, vec![0]),
            (0x4a43, 24u16.to_le_bytes().to_vec()),
            (0x4a61, 30u16.to_le_bytes().to_vec()),
            (0x8840, (-30i16).to_le_bytes().to_vec()),
            (0x4845, (-3i16).to_le_bytes().to_vec()),
            (0x484b, 20u16.to_le_bytes().to_vec()),
            (0x4852, 67u16.to_le_bytes().to_vec()),
            (0x2a48, vec![2]),
            (0x2a3e, vec![11]),
            (0x6870, vec![0x12, 0x34, 0x56, 0]),
        ]);
        let run = properties
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        assert!(run.bold && run.italic && run.strikethrough && run.double_strikethrough);
        assert!(run.small_caps && run.all_caps);
        assert_eq!((run.font_size, run.font_size_cs), (12.0, Some(15.0)));
        assert_eq!((run.char_spacing, run.char_scale), (Some(-1.5), Some(0.67)));
        assert_eq!((run.position, run.kerning), (Some(-1.5), Some(10.0)));
        assert_eq!(
            (run.vert_align.as_deref(), run.underline_style.as_deref()),
            (Some("sub"), Some("wave"))
        );
        assert_eq!(run.color.as_deref(), Some("123456"));
        assert!(run.fit_text_id.is_none() && run.revision.is_none());
        let wire = run.typography_acquisition.unwrap();
        assert_eq!(wire.position_pt, valid_wire("-3".into(), -1.5));
        assert_eq!(
            wire.vertical_align,
            valid_wire("subscript".into(), "sub".into())
        );
        assert_eq!(wire.underline.unwrap().val.value.as_deref(), Some("wave"));
    }

    #[test]
    fn indexed_text_colors_match_the_existing_docx_parser_projection() {
        for index in 0..17u8 {
            let properties = applied(&[(0x2a42, vec![index])]);
            assert_parser_parity(&properties, &[]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            let expected = (index != 0).then(|| ICO_COLORS[index as usize].to_ascii_lowercase());
            assert_eq!(run.color, expected);
        }
    }

    #[test]
    fn highlights_match_the_existing_docx_parser_projection() {
        let expected = [
            None,
            Some("black"),
            Some("blue"),
            Some("cyan"),
            Some("green"),
            Some("magenta"),
            Some("red"),
            Some("yellow"),
            Some("white"),
            Some("darkBlue"),
            Some("darkCyan"),
            Some("darkGreen"),
            Some("darkMagenta"),
            Some("darkRed"),
            Some("darkYellow"),
            Some("darkGray"),
            Some("lightGray"),
        ];
        for (index, expected) in expected.into_iter().enumerate() {
            let properties = applied(&[(0x2a0c, vec![index as u8])]);
            assert_parser_parity(&properties, &[]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.highlight.as_deref(), expected);
        }
    }

    #[test]
    fn adjacent_highlights_remain_distinct_across_xml_and_direct_projection() {
        let magenta = applied(&[(0x2a0c, vec![12])]);
        let cleared = applied(&[(0x2a0c, vec![0])]);
        let red = applied(&[(0x2a0c, vec![13])]);
        let runs = [&magenta, &cleared, &red]
            .into_iter()
            .map(|properties| {
                properties
                    .direct_text_run("x".into(), &[])
                    .unwrap()
                    .unwrap()
                    .highlight
            })
            .collect::<Vec<_>>();
        assert_eq!(
            runs,
            vec![Some("darkMagenta".into()), None, Some("darkRed".into())]
        );

        let document_xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>{}<w:t>a</w:t></w:r><w:r>{}<w:t>b</w:t></w:r><w:r>{}<w:t>c</w:t></w:r></w:p></w:body></w:document>"#,
            magenta.xml(&[]).unwrap(),
            cleared.xml(&[]).unwrap(),
            red.xml(&[]).unwrap(),
        );
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            archive
                .start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            archive.write_all(document_xml.as_bytes()).unwrap();
            archive.finish().unwrap();
        }
        let parsed: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&bytes).unwrap()).unwrap();
        assert_eq!(parsed["body"][0]["runs"][0]["highlight"], "darkMagenta");
        assert!(parsed["body"][0]["runs"][1].get("highlight").is_none());
        assert_eq!(parsed["body"][0]["runs"][2]["highlight"], "darkRed");
    }

    #[test]
    fn projects_four_literal_font_axes_and_hint_without_theme_facts() {
        let properties = applied(&[
            (0x4a4f, 0u16.to_le_bytes().to_vec()),
            (0x4a50, 1u16.to_le_bytes().to_vec()),
            (0x4a51, 2u16.to_le_bytes().to_vec()),
            (0x4a5e, 3u16.to_le_bytes().to_vec()),
            (0x286f, vec![2]),
        ]);
        let fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"].map(String::from);
        let run = properties
            .direct_text_run("x".into(), &fonts)
            .unwrap()
            .unwrap();
        assert_eq!(run.font_family.as_deref(), Some("ASCII"));
        assert_eq!(run.font_family_east_asia.as_deref(), Some("East Asia"));
        assert_eq!(run.font_family_high_ansi.as_deref(), Some("High ANSI"));
        assert_eq!(run.font_family_cs.as_deref(), Some("Complex Script"));
        assert_eq!(run.font_hint.as_deref(), Some("cs"));
        let slots = run.font_slots.unwrap();
        assert_eq!(slots.direct.ascii.as_deref(), Some("ASCII"));
        assert_eq!(slots.direct.east_asia.as_deref(), Some("East Asia"));
        assert_eq!(slots.direct.high_ansi.as_deref(), Some("High ANSI"));
        assert_eq!(
            slots.direct.complex_script.as_deref(),
            Some("Complex Script")
        );
        assert_eq!(slots.theme, RunFontAxisValues::default());
        assert_eq!(slots.theme_present, RunFontAxisPresence::default());
    }

    #[test]
    fn hidden_text_is_absent_but_mark_font_facts_remain_available() {
        let properties = applied(&[(0x083c, vec![1]), (0x0835, vec![1])]);
        assert!(properties
            .direct_text_run("hidden".into(), &[])
            .unwrap()
            .is_none());
        assert!(properties.direct_vanish());
        let facts = properties.direct_font_facts(&[]).unwrap();
        assert!(facts.bold);
        assert_eq!(facts.font_size, Some(10.0));
    }

    #[test]
    fn underline_mapping_is_exact_and_single_needs_no_style_hint() {
        for (operand, token) in [
            (0, "none"),
            (1, "single"),
            (2, "words"),
            (3, "double"),
            (4, "dotted"),
            (6, "thick"),
            (7, "dash"),
            (9, "dotDash"),
            (10, "dotDotDash"),
            (11, "wave"),
            (20, "dottedHeavy"),
            (23, "dashedHeavy"),
            (25, "dashDotHeavy"),
            (26, "dashDotDotHeavy"),
            (27, "wavyHeavy"),
            (39, "dashLong"),
            (43, "wavyDouble"),
            (55, "dashLongHeavy"),
        ] {
            let run = applied(&[(0x2a3e, vec![operand])])
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.underline, token != "none", "{token}");
            assert_eq!(
                run.underline_style.as_deref(),
                (!matches!(token, "none" | "single")).then_some(token),
                "{token}"
            );
            assert_eq!(
                run.typography_acquisition
                    .unwrap()
                    .underline
                    .unwrap()
                    .val
                    .value
                    .as_deref(),
                Some(token)
            );
            assert_parser_parity(&applied(&[(0x2a3e, vec![operand])]), &[]);
        }
    }

    #[test]
    fn underline_color_matches_the_xml_parser_route() {
        for (underline, color, expected) in [
            (0, vec![0x12, 0x34, 0x56, 0], None),
            (1, vec![0x12, 0x34, 0x56, 0], Some("123456")),
            (11, vec![0, 0, 0, 0xff], Some("auto")),
        ] {
            let properties = applied(&[(0x2a3e, vec![underline]), (0x6877, color)]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.underline_color.as_deref(), expected);
            assert_parser_parity(&properties, &[]);
        }
        let color_only = applied(&[(0x6877, vec![0x12, 0x34, 0x56, 0])]);
        let run = color_only
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        assert!(!run.underline);
        assert_eq!(run.underline_color, None);
        assert_eq!(
            run.typography_acquisition
                .as_ref()
                .and_then(|wire| wire.underline.as_ref())
                .and_then(|wire| wire.color.value.as_deref()),
            Some("123456"),
        );
        assert_parser_parity(&color_only, &[]);
    }

    #[test]
    fn underline_color_tracks_effective_underline_style_and_clearing() {
        let base = Properties::default();
        let mut style = base.clone();
        style.apply(0x2a3e, &[4], &base).unwrap();
        style.apply(0x6877, &[0, 0, 0xff, 0], &base).unwrap();
        let inherited = style.direct_text_run("x".into(), &[]).unwrap().unwrap();
        assert_eq!(inherited.underline_style.as_deref(), Some("dotted"));
        assert_eq!(inherited.underline_color.as_deref(), Some("0000ff"));

        let mut cleared = style.clone();
        cleared.apply(0x2a3e, &[0], &style).unwrap();
        let run = cleared.direct_text_run("x".into(), &[]).unwrap().unwrap();
        assert!(!run.underline);
        assert_eq!(run.underline_color, None);
        assert_parser_parity(&cleared, &[]);
    }

    #[test]
    fn numeric_and_color_boundaries_match_the_xml_parser_route() {
        for (code, operands) in [
            (0x4a43, vec![2u16, 3276]),
            (0x4a61, vec![0u16, 3276]),
            (0x8840, vec![i16::MIN as u16, 0, i16::MAX as u16]),
            (0x4845, vec![(-3168i16) as u16, 0, 3168]),
            (0x484b, vec![0u16, 3276]),
            (0x4852, vec![1u16, 600]),
        ] {
            for operand in operands {
                assert_parser_parity(&applied(&[(code, operand.to_le_bytes().to_vec())]), &[]);
            }
        }
        for operand in [vec![0xab, 0xcd, 0xef, 0], vec![0, 0, 0, 0xff]] {
            assert_parser_parity(&applied(&[(0x6870, operand)]), &[]);
        }
        assert!(Properties::default()
            .apply(0x4a43, &0u16.to_le_bytes(), &Properties::default())
            .is_err());
    }

    #[test]
    fn vertical_alignment_boundaries_preserve_current_parser_provenance() {
        for (operand, raw, normalized, status) in [
            (0, "baseline", None, TypographyValueStatusWire::Invalid),
            (
                1,
                "superscript",
                Some("super"),
                TypographyValueStatusWire::Valid,
            ),
            (
                2,
                "subscript",
                Some("sub"),
                TypographyValueStatusWire::Valid,
            ),
        ] {
            let properties = applied(&[(0x2a48, vec![operand])]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.vert_align.as_deref(), normalized);
            let wire = run.typography_acquisition.unwrap().vertical_align;
            assert_eq!(wire.status, status);
            assert_eq!(wire.raw.as_deref(), Some(raw));
            assert_parser_parity(&properties, &[]);
        }
    }

    #[test]
    fn font_hints_cancellation_and_partial_axes_match_the_xml_parser_route() {
        for hint in [0, 1, 2, 0xff] {
            assert_parser_parity(&applied(&[(0x286f, vec![hint])]), &[]);
        }
        assert_parser_parity(&Properties::default(), &[]);

        let fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"].map(String::from);
        for (code, index) in [(0x4a4f, 0u16), (0x4a50, 1), (0x4a51, 2), (0x4a5e, 3)] {
            assert_parser_parity(&applied(&[(code, index.to_le_bytes().to_vec())]), &fonts);
        }
    }

    #[test]
    fn invalid_font_references_error_before_vanish_filtering() {
        let invalid = applied(&[(0x4a4f, 1u16.to_le_bytes().to_vec())]);
        assert!(invalid.direct_text_run("x".into(), &[]).is_err());
        let vanished = applied(&[(0x083c, vec![1]), (0x4a4f, 1u16.to_le_bytes().to_vec())]);
        assert!(vanished.direct_text_run("x".into(), &[]).is_err());
        assert!(vanished.xml(&[]).is_err());

        let hinted = applied(&[(0x286f, vec![1]), (0x4a4f, 1u16.to_le_bytes().to_vec())]);
        assert!(hinted.xml(&[]).is_ok());
        assert!(hinted.direct_text_run("x".into(), &[]).is_err());
    }

    #[test]
    fn malformed_resolved_numbers_are_invariant_errors_not_missing_values() {
        let mut properties = Properties::default();
        properties.values.insert("position", "corrupt".into());
        assert!(properties.direct_text_run("x".into(), &[]).is_err());
        properties.values.remove("position");
        properties.values.insert("sz", "corrupt".into());
        assert!(properties.direct_font_facts(&[]).is_err());
    }

    #[test]
    fn typed_projection_matches_the_existing_xml_parser_route() {
        let properties = applied(&[
            (0x0835, vec![1]),
            (0x0836, vec![1]),
            (0x0837, vec![1]),
            (0x2a53, vec![1]),
            (0x083a, vec![1]),
            (0x083b, vec![1]),
            (0x085a, vec![1]),
            (0x0882, vec![1]),
            (0x085c, vec![1]),
            (0x085d, vec![0]),
            (0x4a4f, 0u16.to_le_bytes().to_vec()),
            (0x4a50, 1u16.to_le_bytes().to_vec()),
            (0x4a51, 2u16.to_le_bytes().to_vec()),
            (0x4a5e, 3u16.to_le_bytes().to_vec()),
            (0x286f, vec![1]),
            (0x4a43, 24u16.to_le_bytes().to_vec()),
            (0x4a61, 30u16.to_le_bytes().to_vec()),
            (0x8840, (-30i16).to_le_bytes().to_vec()),
            (0x4845, (-3i16).to_le_bytes().to_vec()),
            (0x484b, 20u16.to_le_bytes().to_vec()),
            (0x4852, 67u16.to_le_bytes().to_vec()),
            (0x2a48, vec![2]),
            (0x2a3e, vec![11]),
            (0x6877, vec![0x65, 0x43, 0x21, 0]),
            (0x6870, vec![0x12, 0x34, 0x56, 0]),
        ]);
        let fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"].map(String::from);
        assert_parser_parity(&properties, &fonts);
    }
}
