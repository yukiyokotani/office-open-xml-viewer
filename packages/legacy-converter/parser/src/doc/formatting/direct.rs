//! Direct model assembly from the same resolved DOC formatting cascade used by
//! the legacy WordprocessingML adapter. Numbering remains deferred to its owner.

use super::{numbering, Formatting, Properties, TableFormattingKey};
use docx_model::AnchorHostMetrics;
use docx_model::{DocParagraph, NumberingInfo, TextRun};

pub(in crate::doc) struct DirectResolvedParagraph {
    pub(in crate::doc) paragraph: DocParagraph,
    pub(in crate::doc) numbering: Option<(numbering::Reference, Properties)>,
    /// The paragraph carries frame properties that `frame_pr` cannot
    /// represent. The caller decides: outside tables this is unsupported; a
    /// table paragraph may instead repeat its table's position.
    pub(in crate::doc) frame_gap: bool,
    pub(in crate::doc) table_frame: Option<crate::doc::paragraph::TableParagraphFrame>,
}

pub(in crate::doc) struct DirectInlinePictureFacts {
    pub(in crate::doc) location: Option<usize>,
    pub(in crate::doc) vanish: bool,
}

impl Formatting<'_> {
    pub(in crate::doc) fn direct_numbering(
        &mut self,
        store: &mut numbering::direct::Store,
        reference: numbering::Reference,
        marker: &Properties,
        paragraph: &DocParagraph,
    ) -> Result<NumberingInfo, String> {
        store.activate(&self.numbering, reference, marker, paragraph, &self.fonts)
    }

    pub(in crate::doc) fn direct_anchor_host_metrics(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Option<AnchorHostMetrics>, String> {
        let properties =
            self.run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?;
        if !properties.picture.passive_special() {
            return Err(super::super::unsupported(
                "floating picture character is not passive-special",
            ));
        }
        if properties.direct_vanish() {
            return Ok(None);
        }
        let facts = properties.direct_font_facts(&self.fonts)?;
        Ok(Some(AnchorHostMetrics {
            font_size: facts.font_size.ok_or_else(|| {
                super::super::unsupported("floating picture host has no resolved font size")
            })?,
            font_family: facts.font_family,
            font_family_east_asia: facts.font_family_east_asia,
            bold: facts.bold,
            italic: facts.italic,
            anchor_occurrence_id: None,
        }))
    }

    pub(in crate::doc) fn direct_normal_style_font_size_pt(&mut self) -> Result<f64, String> {
        // MS-DOC 2.6.4 sprmSDxtCharSpace is relative to the Normal style,
        // not the paragraph mark or any visible body run's direct formatting.
        self.paragraph_base(0)?.direct_font_size_pt()
    }

    pub(in crate::doc) fn direct_paragraph(
        &mut self,
        style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<DirectResolvedParagraph, String> {
        let mut resolved = self.resolve_paragraph_with_table(style, table_style, fc, prm, prcs)?;
        // Numbered resolution already acquired the unmodified paragraph mark;
        // plain paragraphs acquire it once here through the same run cascade.
        let mark = match resolved.paragraph_mark.take() {
            Some(mark) => mark,
            None => self.run_properties_with_table(style, table_style, fc, prm, prcs)?,
        };
        let mut paragraph = resolved.properties.direct_paragraph();
        paragraph.outline_level = resolved.properties.direct_outline_level(style);
        // ECMA-376 17.3.1.9 compares paragraph styles; the DOCX renderer does
        // so through `style_id`. The DOC paragraph istd is that identity.
        paragraph.style_id = Some(style.to_string());
        let frame_gap = match resolved.properties.direct_frame() {
            Ok(frame) => {
                paragraph.frame_pr = frame.map(Box::new);
                false
            }
            Err(_) => true,
        };
        let table_frame = resolved.properties.table_paragraph_frame();
        paragraph.mark_vanish = mark.direct_vanish();
        let mark_facts = mark.direct_font_facts(&self.fonts)?;
        paragraph.default_font_size = mark_facts.font_size;
        paragraph.default_font_family = mark_facts.font_family.clone();
        paragraph.default_font_family_east_asia = mark_facts.font_family_east_asia.clone();
        paragraph.paragraph_mark_font_facts = Some(mark_facts);
        paragraph.paragraph_mark_color = mark.direct_color();
        Ok(DirectResolvedParagraph {
            paragraph,
            numbering: resolved.numbering,
            frame_gap,
            table_frame,
        })
    }

    pub(in crate::doc) fn direct_text_run(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
        text: String,
    ) -> Result<Option<TextRun>, String> {
        self.run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?
            .direct_text_run(text, &self.fonts)
    }

    /// MS-DOC 2.6.1 sprmCLbcCRJ on a U+000B line break (see
    /// `Properties::direct_line_break_clears`).
    pub(in crate::doc) fn direct_line_break_clears(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<bool, String> {
        Ok(self
            .run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?
            .direct_line_break_clears())
    }

    /// Resolve the CHPX cascade once for an inline-picture character. Visibility
    /// is a run property just as it is for text; picture acquisition must not
    /// create a resource for a vanished run.
    pub(in crate::doc) fn direct_inline_picture_facts(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<DirectInlinePictureFacts, String> {
        let properties =
            self.run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?;
        Ok(DirectInlinePictureFacts {
            location: properties.picture.inline_location()?,
            vanish: properties.direct_vanish(),
        })
    }
}

impl<'a> Formatting<'a> {
    /// A link run (see `direct_model::fields::Link`). Inside a TOC result the
    /// DOCX parser replaces the run's color and underline state with the
    /// paragraph-level (style and table style) character properties.
    pub(in crate::doc) fn direct_link_text_run(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
        in_toc: bool,
    ) -> Result<Option<TextRun>, String> {
        let mut properties =
            self.run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?;
        if in_toc {
            let base = self.paragraph_base_with_table(paragraph_style, table_style)?;
            properties.take_link_display_from(&base);
        }
        properties.direct_text_run(String::new(), &self.fonts)
    }

    /// MS-DOC 2.6.1 sprmCFSpec + sprmCFData + sprmCPicLocation and 2.9.158
    /// NilPICFAndBinData: the binData of a binary-data character (a form
    /// field, hyperlink or add-in field payload), bounded by its lcb.
    pub(in crate::doc) fn direct_binary_data(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<&'a [u8], String> {
        let invalid = || super::super::unsupported("invalid Word binary-data character");
        let picture = self
            .run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?
            .picture;
        if !picture.special || !picture.data || picture.ole || picture.object {
            return Err(invalid());
        }
        let offset = picture
            .location
            .and_then(|location| usize::try_from(location).ok())
            .ok_or_else(invalid)?;
        let data = self.data;
        let header = data.get(offset..offset + 6).ok_or_else(invalid)?;
        let length = i32::from_le_bytes([header[0], header[1], header[2], header[3]]);
        if u16::from_le_bytes([header[4], header[5]]) != 0x44 {
            return Err(invalid());
        }
        usize::try_from(length)
            .ok()
            .filter(|length| *length >= 0x44)
            .and_then(|length| data.get(offset + 0x44..offset + length))
            .ok_or_else(invalid)
    }
}
