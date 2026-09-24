//! Typed DOC character projection for the direct DOC model path.
//!
//! The binary decoder has already resolved style inheritance and validated every
//! value stored by [`Properties`]. This module projects that state without
//! manufacturing WordprocessingML and parsing it back through the DOCX parser.

use super::Properties;
use crate::doc::border::Border;
use crate::doc::paragraph::ShadingFill;
use crate::doc::unsupported;
use docx_model::{
    BodyElement, CellElement, CtBorderTypographyWire, DocRun, EastAsianLayoutTypographyWire,
    FitTextSpecWire, RunBorder, RunFontAxisPresence, RunFontAxisValues, RunFontFacts, RunFontSlots,
    RunTypographyWire, TextRun, TypographyLanguagesWire, TypographyValueStatusWire,
    TypographyValueWire, UnderlineTypographyWire,
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
        let languages = self.resolved_languages()?;
        let lang_default = languages.default.map(str::to_ascii_lowercase);
        let lang_east_asia = languages.east_asia.map(str::to_ascii_lowercase);
        let lang_bidi = languages.bidi.map(str::to_ascii_lowercase);
        if self.bool_value("vanish").unwrap_or(false) {
            return Ok(None);
        }
        if self.field_vanish == Some(true) {
            // Word sets fFldVanish on field instruction characters, which the
            // direct field projection never emits as text. Its effect on
            // ordinary visible text is not specified beyond "hidden".
            return Err(unsupported("Word field-hidden property on visible text"));
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
        let (border, border_wire) = self.direct_border()?;
        let symbol = self.direct_symbol(fonts)?;
        let fit_text = self.direct_only.fit_text;

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
            lang_bidi: lang_bidi.clone(),
            lang_default: lang_default.clone(),
            lang_east_asia: lang_east_asia.clone(),
            char_spacing,
            char_scale,
            position,
            kerning,
            highlight,
            snap_to_grid: self.direct_only.snap_to_grid,
            // MS-DOC 2.6.1 sprmCShd/sprmCShd80: a single-color fill only.
            background: match &self.direct_only.shading {
                Some(ShadingFill::Rgb(fill)) => Some(fill.clone()),
                Some(ShadingFill::None) | None => None,
            },
            border,
            // MS-DOC 2.9.31: dxaFitText twips and FitTextID. Contiguous runs
            // sharing an ID form one region, exactly as ECMA-376 17.3.2.14
            // fitText@w:id links consecutive runs.
            fit_text_val: fit_text.map(|(width, _)| f64::from(width)),
            fit_text_id: fit_text.map(|(_, id)| id.to_string()),
            ..TextRun::default()
        };
        if let Some((vertical, compress)) = self.direct_only.east_asian {
            run.east_asian_vert = Some(vertical);
            run.east_asian_vert_compress = Some(compress);
        }
        if let Some((glyph, font)) = symbol {
            // ECMA-376 17.3.3.30 sym: a one-glyph run in the symbol's font,
            // exactly as the DOCX parser projects `w:sym` (font on the
            // ascii, high-ANSI and East Asian axes, no slot provenance).
            run.text = glyph;
            run.font_family = Some(font.clone());
            run.font_family_high_ansi = Some(font.clone());
            run.font_family_east_asia = Some(font);
            run.font_slots = None;
        }
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
            snap_to_grid: run.snap_to_grid,
            fit_text: fit_text.map(|(width, id)| FitTextSpecWire {
                val_twips: f64::from(width),
                id: Some(id.to_string()),
            }),
            border: border_wire,
            east_asian_layout: EastAsianLayoutTypographyWire {
                vert: run.east_asian_vert,
                vert_compress: run.east_asian_vert_compress,
                ..EastAsianLayoutTypographyWire::default()
            },
            kerning_threshold_pt: run.kerning,
            languages: TypographyLanguagesWire {
                bidi: lang_bidi,
                default: lang_default,
                east_asia: lang_east_asia,
            },
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
        let languages = self.resolved_languages()?;
        let lang_default = languages.default.map(str::to_ascii_lowercase);
        let lang_east_asia = languages.east_asia.map(str::to_ascii_lowercase);
        let lang_bidi = languages.bidi.map(str::to_ascii_lowercase);
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
            lang_bidi,
            lang_default,
            lang_east_asia,
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

    pub(in crate::doc) fn direct_color_auto(&self) -> bool {
        self.values
            .get("color")
            .is_some_and(|value| value == "auto")
    }

    /// ECMA-376 17.3.2.10 horizontal-in-vertical text is rendered only in
    /// vertical (tbRl) body flow; horizontal flow and table cells would drop
    /// the MS-DOC 2.9.332 fTNY rotation. Returns true when such a run occurs
    /// where the renderer cannot honor it.
    pub(in crate::doc) fn unrenderable_east_asian_vertical(
        elements: &[BodyElement],
        vertical_flow: bool,
    ) -> bool {
        fn paragraph(value: &docx_model::DocParagraph) -> bool {
            value
                .runs
                .iter()
                .any(|run| matches!(run, DocRun::Text(text) if text.east_asian_vert == Some(true)))
        }
        fn table(value: &docx_model::DocTable) -> bool {
            value.rows.iter().flat_map(|row| &row.cells).any(|cell| {
                cell.content.iter().any(|block| match block {
                    CellElement::Paragraph(value) => paragraph(value),
                    CellElement::Table(value) => table(value),
                })
            })
        }
        elements.iter().any(|element| match element {
            BodyElement::Paragraph(value) => !vertical_flow && paragraph(value),
            BodyElement::Table(value) => table(value),
            _ => false,
        })
    }

    /// MS-DOC 2.6.1 sprmCSymbol "designates the character as a symbol", and
    /// sprmCFSpec lists U+0028 as "a symbol, see sprmCSymbol". The symbol
    /// placeholder is therefore U+0028 (checked in `direct_run_text`).
    /// sprmCSymbol does not require sprmCFSpec: a Word-produced corpus
    /// document applies it to non-fSpec U+0028 characters and Word's own PDF
    /// shows the symbol glyph there. Picture/OLE/object characters are never
    /// symbols.
    fn direct_symbol(&self, fonts: &[String]) -> Result<Option<(String, String)>, String> {
        let Some((font, code)) = self.symbol else {
            return Ok(None);
        };
        if self.picture.data || self.picture.ole || self.picture.object {
            return Err(unsupported("Word symbol on a picture or object character"));
        }
        let glyph = char::from_u32(u32::from(code))
            .ok_or_else(|| unsupported("invalid Word symbol character code"))?;
        let font = fonts
            .get(usize::from(font))
            .ok_or_else(|| unsupported("Word symbol font outside font table"))?;
        Ok(Some((glyph.to_string(), font.clone())))
    }

    /// Text for one visible span resolved by [`Self::direct_text_run`]. A
    /// symbol run carries its glyph; every source character of the span MUST
    /// then be the U+0028 symbol placeholder, one glyph each.
    pub(in crate::doc) fn direct_run_text(run: &mut TextRun, part: &str) -> Result<String, String> {
        if run.text.is_empty() {
            return Ok(part.to_string());
        }
        if part.is_empty() || part.chars().any(|character| character != '(') {
            return Err(unsupported(
                "Word symbol property on a non-symbol character",
            ));
        }
        let glyph = std::mem::take(&mut run.text);
        Ok(glyph.repeat(part.chars().count()))
    }

    /// MS-DOC 2.6.1 sprmCBrc/sprmCBrc80 as an ECMA-376 17.3.2.4 run border.
    /// "Brc.dptSpace MUST be ignored when applied to character borders", so
    /// the projected spacing is zero and the raw spacing is not acquired.
    /// Shadowed borders and asymmetric frame effects were rejected at apply.
    fn direct_border(&self) -> Result<(Option<RunBorder>, Option<CtBorderTypographyWire>), String> {
        let Some((old, raw)) = self.direct_only.border else {
            return Ok((None, None));
        };
        let border = Border::read(&raw[..if old { 4 } else { 8 }], old)?;
        let mut wire = border.direct_typography();
        wire.space_pt = TypographyValueWire::default();
        let edge = border.direct_edge();
        let run = (edge.style != "none").then_some(RunBorder {
            style: edge.style,
            color: edge.color,
            width: edge.width,
            space: 0.0,
        });
        Ok((run, Some(wire)))
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

    /// The DOCX parser's in-TOC link display: take the paragraph-level color
    /// and underline state, keeping this run's underline style when both are
    /// underlined.
    pub(in crate::doc) fn take_link_display_from(&mut self, base: &Properties) {
        match base.values.get("color") {
            Some(color) => {
                self.values.insert("color", color.clone());
            }
            None => {
                self.values.remove("color");
            }
        }
        let underlined = |properties: &Properties| {
            properties
                .values
                .get("u")
                .filter(|value| value.as_str() != "none")
                .cloned()
        };
        match (underlined(base), underlined(self)) {
            (None, _) => {
                self.values.remove("u");
            }
            (Some(token), None) => {
                self.values.insert("u", token);
            }
            (Some(_), Some(_)) => {}
        }
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
    fn revision_session_ids_do_not_create_direct_revision_state() {
        let baseline = serde_json::to_value(
            Properties::default()
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        for code in 0x6815..=0x6817 {
            let properties = applied(&[(code, 0x7856_3412u32.to_le_bytes().to_vec())]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert!(run.revision.is_none());
            assert_eq!(serde_json::to_value(run).unwrap(), baseline);
            assert_parser_parity(&properties, &[]);
        }
    }

    #[test]
    fn proofing_only_differences_do_not_change_the_native_display_model() {
        let baseline = serde_json::to_value(
            Properties::default()
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        for operand in [0, 1, 0x80, 0x81] {
            let properties = applied(&[(0x0875, vec![operand])]);
            assert_parser_parity(&properties, &[]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(serde_json::to_value(run).unwrap(), baseline);
        }
    }

    #[test]
    fn complex_script_language_matches_xml_parser_without_family_inference() {
        for (lid, expected) in [(0x0401u16, "ar-sa"), (0x0411, "ja-jp"), (0x0409, "en-us")] {
            let properties = applied(&[(0x485f, lid.to_le_bytes().to_vec())]);
            assert_parser_parity(&properties, &[]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.lang_bidi.as_deref(), Some(expected));
            assert_eq!(
                run.typography_acquisition
                    .as_ref()
                    .and_then(|wire| wire.languages.bidi.as_deref()),
                Some(expected),
            );
            assert_eq!(
                properties
                    .direct_font_facts(&[])
                    .unwrap()
                    .lang_bidi
                    .as_deref(),
                Some(expected),
            );
        }

        for lid in [u16::MAX, 0x1000, 0x0400, 0x007f, 0x0467, 0x040a] {
            let unresolved = applied(&[(0x485f, lid.to_le_bytes().to_vec())]);
            assert!(unresolved.xml(&[]).is_err(), "LID {lid:04x}");
            assert!(unresolved.direct_text_run("x".into(), &[]).is_err());
            assert!(unresolved.direct_font_facts(&[]).is_err());
        }

        // Projection validates the retained language before visibility, just as
        // the XML writer validates run properties before the DOCX parser drops
        // a vanished run.
        let hidden_unknown =
            applied(&[(0x485f, u16::MAX.to_le_bytes().to_vec()), (0x083c, vec![1])]);
        assert!(hidden_unknown.xml(&[]).is_err());
        assert!(hidden_unknown.direct_text_run("x".into(), &[]).is_err());
    }

    #[test]
    fn modern_default_and_east_asian_languages_match_xml_parser() {
        let properties = applied(&[
            (0x486d, 0x0411u16.to_le_bytes().to_vec()),
            (0x486e, 0x0412u16.to_le_bytes().to_vec()),
            (0x4873, 0x040cu16.to_le_bytes().to_vec()),
            (0x4874, 0x0404u16.to_le_bytes().to_vec()),
            (0x485f, 0x0401u16.to_le_bytes().to_vec()),
        ]);
        assert_parser_parity(&properties, &[]);
        let run = properties
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        assert_eq!(run.lang_default.as_deref(), Some("fr-fr"));
        assert_eq!(run.lang_east_asia.as_deref(), Some("zh-tw"));
        assert_eq!(run.lang_bidi.as_deref(), Some("ar-sa"));
        let languages = &run.typography_acquisition.as_ref().unwrap().languages;
        assert_eq!(languages.default.as_deref(), Some("fr-fr"));
        assert_eq!(languages.east_asia.as_deref(), Some("zh-tw"));
        assert_eq!(languages.bidi.as_deref(), Some("ar-sa"));

        let facts = properties.direct_font_facts(&[]).unwrap();
        assert_eq!(facts.lang_default.as_deref(), Some("fr-fr"));
        assert_eq!(facts.lang_east_asia.as_deref(), Some("zh-tw"));
        assert_eq!(facts.lang_bidi.as_deref(), Some("ar-sa"));

        for (code, lid) in [(0x4873, 0x0400u16), (0x4874, 0xffff)] {
            let unresolved = applied(&[(code, lid.to_le_bytes().to_vec())]);
            assert!(unresolved.xml(&[]).is_err());
            assert!(unresolved.direct_text_run("x".into(), &[]).is_err());
            assert!(unresolved.direct_font_facts(&[]).is_err());
        }
    }

    #[test]
    fn modern_languages_win_compatibility_metadata_in_both_orders() {
        for entries in [
            vec![(0x486d, 0x0411u16), (0x4873, 0x040cu16)],
            vec![(0x4873, 0x040cu16), (0x486d, 0x0411u16)],
        ] {
            let entries = entries
                .into_iter()
                .map(|(code, lid)| (code, lid.to_le_bytes().to_vec()))
                .collect::<Vec<_>>();
            let properties = applied(&entries);
            assert_parser_parity(&properties, &[]);
            let run = properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap();
            assert_eq!(run.lang_default.as_deref(), Some("fr-fr"));
            assert_eq!(
                properties
                    .direct_font_facts(&[])
                    .unwrap()
                    .lang_default
                    .as_deref(),
                Some("fr-fr")
            );
        }

        let compatibility_only = applied(&[
            (0x486d, 0x0411u16.to_le_bytes().to_vec()),
            (0x486e, 0x0412u16.to_le_bytes().to_vec()),
        ]);
        assert_parser_parity(&compatibility_only, &[]);
        let run = compatibility_only
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        assert_eq!(run.lang_default, None);
        assert_eq!(run.lang_east_asia, None);
        assert_eq!(run.typography_acquisition.unwrap().languages.default, None);
        let facts = compatibility_only.direct_font_facts(&[]).unwrap();
        assert_eq!(facts.lang_default, None);
        assert_eq!(facts.lang_east_asia, None);
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

    fn parsed_rpr(rpr: &str) -> serde_json::Value {
        let document_xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr>{rpr}</w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
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
        let object = run.as_object_mut().unwrap();
        object.remove("type");
        object.remove("__typographyAcquisition");
        run
    }

    fn public_run(properties: &Properties) -> serde_json::Value {
        let mut run = serde_json::to_value(
            properties
                .direct_text_run("x".into(), &[])
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        run.as_object_mut()
            .unwrap()
            .remove("__typographyAcquisition");
        run
    }

    fn cshd(fore: [u8; 4], back: [u8; 4], ipat: u16) -> Vec<u8> {
        let mut operand = vec![10];
        operand.extend(fore);
        operand.extend(back);
        operand.extend(ipat.to_le_bytes());
        operand
    }

    #[test]
    fn character_shading_border_and_fit_text_match_the_docx_model_semantics() {
        let properties = applied(&[
            (0xca71, cshd([0, 0, 0, 0xff], [0xdd, 0xdd, 0xdd, 0], 0)),
            (0xca72, vec![8, 0x12, 0x34, 0x56, 0, 6, 1, 0x05, 0]),
            (0xca76, vec![8, 0x60, 0x09, 0, 0, 0, 0xb5, 0xad, 0xaa]),
        ]);
        assert!(properties.has_direct_only_properties());
        // MS-DOC 2.6.1: Brc.dptSpace MUST be ignored for character borders.
        assert_eq!(
            public_run(&properties),
            parsed_rpr(
                r#"<w:sz w:val="20"/><w:bdr w:val="single" w:sz="6" w:space="0" w:color="123456"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/><w:fitText w:val="2400" w:id="-1431456512"/>"#
            )
        );
        let run = properties
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        let wire = run.typography_acquisition.unwrap();
        let fit = wire.fit_text.unwrap();
        assert_eq!(
            (fit.val_twips, fit.id.as_deref()),
            (2400.0, Some("-1431456512"))
        );
        let border = wire.border.unwrap();
        assert_eq!(border.val.value.as_deref(), Some("single"));
        assert_eq!(border.space_pt, TypographyValueWire::default());

        // Brc80 with an automatic color, then a black Shd80 background.
        let properties = applied(&[
            (0x6865, vec![4, 1, 0, 0]),
            (0x4866, 0x0020u16.to_le_bytes().to_vec()),
        ]);
        assert_eq!(
            public_run(&properties),
            parsed_rpr(
                r#"<w:sz w:val="20"/><w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:shd w:val="clear" w:color="auto" w:fill="000000"/>"#
            )
        );
    }

    #[test]
    fn character_border_none_nil_and_resets_remove_the_border() {
        for operand in [vec![0, 0, 0, 0x40], vec![0xff; 4]] {
            let run = public_run(&applied(&[(0x6865, operand)]));
            assert!(run.get("border").is_none());
        }
        let nil = applied(&[(
            0xca72,
            vec![8, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff],
        )]);
        assert!(public_run(&nil).get("border").is_none());
        // The frame effect is invisible on a single stroke; shadows and
        // asymmetric frame effects cannot be represented.
        let framed = applied(&[(0x6865, vec![4, 1, 0, 0x40])]);
        assert!(public_run(&framed).get("border").is_some());
        let base = Properties::default();
        for (code, operand) in [
            (0x6865, vec![4, 1, 0, 0x20]),
            (0x6865, vec![4, 12, 0, 0x40]),
            (0xca72, vec![8, 0, 0, 0, 0, 4, 1, 0x20, 0]),
        ] {
            assert!(!base.clone().apply(code, &operand, &base).unwrap());
        }
        for (code, operand) in [
            (0xca72, vec![7, 0, 0, 0, 0, 4, 1, 0, 0]),
            (0xca72, vec![8, 0, 0, 0, 0, 4, 1, 0]),
            (0x6865, vec![4, 1, 17, 0]),
            (0x6865, vec![4, 0x1a, 0, 0]),
        ] {
            assert!(base.clone().apply(code, &operand, &base).is_err());
        }
        // Neither CPlain nor CIstd preserves border, shading or fit text.
        let mut reset = applied(&[
            (0x6865, vec![4, 1, 0, 0]),
            (0x4866, 0x0100u16.to_le_bytes().to_vec()),
            (0xca76, vec![8, 0x60, 0x09, 0, 0, 1, 0, 0, 0]),
        ]);
        reset.reset_to(&base, false);
        assert!(!reset.has_direct_only_properties());
        assert_eq!(public_run(&reset), public_run(&base));
    }

    #[test]
    fn fit_text_zero_is_ignored_and_negative_widths_stay_unsupported() {
        let base = Properties::default();
        let mut value = applied(&[(0xca76, vec![8, 0xb5, 0x04, 0, 0, 1, 0, 0, 0])]);
        assert!(value
            .apply(0xca76, &[8, 0, 0, 0, 0, 2, 0, 0, 0], &base)
            .unwrap());
        let run = value.direct_text_run("x".into(), &[]).unwrap().unwrap();
        assert_eq!(
            (run.fit_text_val, run.fit_text_id.as_deref()),
            (Some(1205.0), Some("1"))
        );
        assert!(!base
            .clone()
            .apply(0xca76, &[8, 0xff, 0xff, 0xff, 0xff, 1, 0, 0, 0], &base)
            .unwrap());
        assert!(base
            .clone()
            .apply(0xca76, &[7, 0, 0, 0, 0, 1, 0, 0, 0], &base)
            .is_err());

        // A sparse style patch overlays fit text and shading only when set.
        let mut inherited = applied(&[(0xca76, vec![8, 0xb5, 0x04, 0, 0, 1, 0, 0, 0])]);
        inherited.overlay_visible(&Properties::sparse());
        assert!(inherited.has_direct_only_properties());
        let mut patch = Properties::sparse();
        patch
            .apply(0x4866, &0x0100u16.to_le_bytes(), &base)
            .unwrap();
        inherited.overlay_visible(&patch);
        let run = inherited.direct_text_run("x".into(), &[]).unwrap().unwrap();
        assert_eq!(run.background.as_deref(), Some("ffffff"));
        assert_eq!(run.fit_text_val, Some(1205.0));
    }

    #[test]
    fn percentage_character_shading_blends_and_other_patterns_stay_unsupported() {
        let base = Properties::default();
        for operand in [
            cshd([0, 0, 0, 0xff], [0xff, 0xff, 0xff, 0], 14),
            cshd([0, 0, 0, 0xff], [0, 0, 0, 0xff], 0x26),
            cshd([0, 0, 0, 0xff], [0xff, 0xff, 0xff, 0], 1),
        ] {
            assert!(!base.clone().apply(0xca71, &operand, &base).unwrap());
        }
        // Word paints pct15 with an automatic foreground over white as #D9D9D9.
        let pct = applied(&[(0x4866, 0x9900u16.to_le_bytes().to_vec())]);
        assert_eq!(public_run(&pct)["background"], "d9d9d9");
        let pct = applied(&[(0xca71, cshd([0, 0, 0, 0xff], [0xff, 0xff, 0xff, 0], 0x26))]);
        assert_eq!(public_run(&pct)["background"], "d9d9d9");
        let cleared = applied(&[(0xca71, cshd([0, 0, 0, 0xff], [0xff; 4], 0))]);
        assert!(public_run(&cleared)["background"].is_null());
    }

    #[test]
    fn symbol_characters_match_the_docx_sym_projection() {
        let fonts = ["Times New Roman", "Symbol"].map(String::from);
        let properties = applied(&[(0x0855, vec![1]), (0x6a09, vec![1, 0, 0xb0, 0xf0])]);
        assert!(properties.has_direct_only_properties());
        let mut run = properties
            .direct_text_run(String::new(), &fonts)
            .unwrap()
            .unwrap();
        let text = Properties::direct_run_text(&mut run, "((").unwrap();
        assert_eq!(text, "\u{f0b0}\u{f0b0}");
        run.text = "\u{f0b0}".into();
        let mut direct = serde_json::to_value(&run).unwrap();
        direct
            .as_object_mut()
            .unwrap()
            .remove("__typographyAcquisition");
        assert_eq!(
            direct,
            parsed_rpr(
                r#"<w:sz w:val="20"/></w:rPr><w:sym w:font="Symbol" w:char="F0B0"/><w:rPr>"#
            )
        );

        // CPlain/CIstd preserve the symbol designation.
        let mut reset = properties.clone();
        reset.reset_to(&Properties::default(), false);
        assert_eq!(reset.symbol, Some((1, 0xf0b0)));

        // Only the special U+0028 placeholder is a symbol.
        let mut run = properties
            .direct_text_run(String::new(), &fonts)
            .unwrap()
            .unwrap();
        assert!(Properties::direct_run_text(&mut run, "(x").is_err());
        // Word renders the symbol without sprmCFSpec as well.
        let plain = applied(&[(0x6a09, vec![1, 0, 0xb0, 0xf0])]);
        let mut run = plain
            .direct_text_run(String::new(), &fonts)
            .unwrap()
            .unwrap();
        assert_eq!(
            Properties::direct_run_text(&mut run, "(").unwrap(),
            "\u{f0b0}"
        );
        let object = applied(&[(0x0856, vec![1]), (0x6a09, vec![1, 0, 0xb0, 0xf0])]);
        assert!(object.direct_text_run(String::new(), &fonts).is_err());
        let outside = applied(&[(0x0855, vec![1]), (0x6a09, vec![2, 0, 0xb0, 0xf0])]);
        assert!(outside.direct_text_run(String::new(), &fonts).is_err());
        let surrogate = applied(&[(0x0855, vec![1]), (0x6a09, vec![1, 0, 0x00, 0xd8])]);
        assert!(surrogate.direct_text_run(String::new(), &fonts).is_err());
        let base = Properties::default();
        assert!(base.clone().apply(0x6a09, &[1, 0, 0], &base).is_err());
        // Ordinary runs keep their text.
        let mut ordinary = base
            .direct_text_run(String::new(), &fonts)
            .unwrap()
            .unwrap();
        assert_eq!(
            Properties::direct_run_text(&mut ordinary, "(a").unwrap(),
            "(a"
        );
    }

    #[test]
    fn horizontal_in_vertical_layout_matches_the_docx_projection_and_is_gated_by_flow() {
        // UFEL 0x1001: fTNY + fTNYCompress, plus an ignored must-be-zero bit.
        let properties = applied(&[(0xca78, vec![6, 0x05, 0x10, 0, 0xc4, 0x3c, 7])]);
        assert!(properties.has_direct_only_properties());
        assert_eq!(
            public_run(&properties),
            parsed_rpr(
                r#"<w:sz w:val="20"/><w:eastAsianLayout w:id="1" w:vert="1" w:vertCompress="1"/>"#
            )
        );
        let run = properties
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        let wire = run.typography_acquisition.as_ref().unwrap();
        assert_eq!(
            (
                wire.east_asian_layout.vert,
                wire.east_asian_layout.vert_compress
            ),
            (Some(true), Some(true))
        );
        // Compression alone is meaningless without fTNY.
        let plain = applied(&[(0xca78, vec![6, 0x00, 0x10, 0, 0, 0, 0])]);
        let plain_run = plain.direct_text_run("x".into(), &[]).unwrap().unwrap();
        assert_eq!(
            (
                plain_run.east_asian_vert,
                plain_run.east_asian_vert_compress
            ),
            (Some(false), Some(false))
        );
        // Two lines in one has no renderer projection.
        let base = Properties::default();
        assert!(!base
            .clone()
            .apply(0xca78, &[6, 2, 0, 0, 0, 0, 0], &base)
            .unwrap());
        assert!(base
            .clone()
            .apply(0xca78, &[5, 1, 0, 0, 0, 0, 0], &base)
            .is_err());

        let paragraph = docx_model::DocParagraph {
            runs: vec![DocRun::Text(Box::new(run))],
            ..Default::default()
        };
        let body = vec![BodyElement::Paragraph(Box::new(paragraph.clone()))];
        assert!(!Properties::unrenderable_east_asian_vertical(&body, true));
        assert!(Properties::unrenderable_east_asian_vertical(&body, false));
        let table = docx_model::DocTable {
            rows: vec![docx_model::DocTableRow {
                cells: vec![docx_model::DocTableCell {
                    content: vec![CellElement::Paragraph(Box::new(paragraph))],
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        let body = vec![BodyElement::Table(Box::new(table))];
        assert!(Properties::unrenderable_east_asian_vertical(&body, true));
        let plain_body = vec![BodyElement::Paragraph(Box::new(docx_model::DocParagraph {
            runs: vec![DocRun::Text(Box::new(plain_run))],
            ..Default::default()
        }))];
        assert!(!Properties::unrenderable_east_asian_vertical(
            &plain_body,
            false
        ));
    }

    #[test]
    fn document_grid_participation_and_field_hiding_are_style_relative_toggles() {
        let base = Properties::default();
        let properties = applied(&[(0x0868, vec![0])]);
        assert!(properties.has_direct_only_properties());
        assert_eq!(
            public_run(&properties),
            parsed_rpr(r#"<w:snapToGrid w:val="0"/><w:sz w:val="20"/>"#)
        );
        let run = properties
            .direct_text_run("x".into(), &[])
            .unwrap()
            .unwrap();
        assert_eq!(
            run.typography_acquisition.unwrap().snap_to_grid,
            Some(false)
        );
        // Toggles are relative to the style; the default is "uses the grid".
        let mut style = Properties::sparse();
        style.apply(0x0868, &[0], &base).unwrap();
        let mut value = style.clone();
        value.apply(0x0868, &[0x81], &style).unwrap();
        assert_eq!(value.direct_only.snap_to_grid, Some(true));
        let mut inherited = base.clone();
        inherited.apply(0x0868, &[0x81], &base).unwrap();
        assert_eq!(inherited.direct_only.snap_to_grid, Some(false));
        assert!(base.clone().apply(0x0868, &[2], &base).is_err());

        // Field hiding survives CPlain/CIstd and only fails for visible text.
        let mut hidden = applied(&[(0x0802, vec![1])]);
        hidden.reset_to(&base, false);
        assert_eq!(hidden.field_vanish, Some(true));
        assert!(hidden.direct_text_run("x".into(), &[]).is_err());
        let mut vanished = hidden.clone();
        vanished.apply(0x083c, &[1], &base).unwrap();
        assert!(vanished.direct_text_run("x".into(), &[]).unwrap().is_none());
        let shown = applied(&[(0x0802, vec![0])]);
        assert!(shown.direct_text_run("x".into(), &[]).unwrap().is_some());
        assert!(!shown.has_direct_only_properties());
    }
}
