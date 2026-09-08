//! Direct model assembly from the same resolved DOC formatting cascade used by
//! the legacy WordprocessingML adapter. Numbering remains deferred to its owner.

use super::{numbering, Formatting, Properties};
use docx_model::AnchorHostMetrics;
use docx_model::{DocParagraph, NumberingInfo, TextRun};

pub(in crate::doc) struct DirectResolvedParagraph {
    pub(in crate::doc) paragraph: DocParagraph,
    pub(in crate::doc) numbering: Option<(numbering::Reference, Properties)>,
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
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Option<AnchorHostMetrics>, String> {
        let properties = self.run_properties(paragraph_style, fc, prm, prcs)?;
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
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<DirectResolvedParagraph, String> {
        let mut resolved = self.resolve_paragraph(style, fc, prm, prcs)?;
        // Numbered resolution already acquired the unmodified paragraph mark;
        // plain paragraphs acquire it once here through the same run cascade.
        let mark = match resolved.paragraph_mark.take() {
            Some(mark) => mark,
            None => self.run_properties(style, fc, prm, prcs)?,
        };
        let mut paragraph = resolved.properties.direct_paragraph();
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
        })
    }

    pub(in crate::doc) fn direct_text_run(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
        text: String,
    ) -> Result<Option<TextRun>, String> {
        self.run_properties(paragraph_style, fc, prm, prcs)?
            .direct_text_run(text, &self.fonts)
    }

    /// Resolve the CHPX cascade once for an inline-picture character. Visibility
    /// is a run property just as it is for text; picture acquisition must not
    /// create a resource for a vanished run.
    pub(in crate::doc) fn direct_inline_picture_facts(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<DirectInlinePictureFacts, String> {
        let properties = self.run_properties(paragraph_style, fc, prm, prcs)?;
        Ok(DirectInlinePictureFacts {
            location: properties.picture.inline_location()?,
            vanish: properties.direct_vanish(),
        })
    }
}
