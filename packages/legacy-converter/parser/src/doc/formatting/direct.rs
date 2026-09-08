//! Direct model assembly from the same resolved DOC formatting cascade used by
//! the legacy WordprocessingML adapter. Numbering remains deferred to its owner.

use super::{numbering, Formatting, Properties};
use docx_model::{DocParagraph, TextRun};

pub(in crate::doc) struct DirectResolvedParagraph {
    pub(in crate::doc) paragraph: DocParagraph,
    pub(in crate::doc) numbering: Option<(numbering::Reference, Properties)>,
}

impl Formatting<'_> {
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
}
