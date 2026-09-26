//! Direct model assembly from the resolved DOC formatting cascade. Numbering
//! remains deferred to its owner.

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
        if mark.direct_insertion().is_some() {
            // An inserted paragraph mark has no DOCX-model run to carry the
            // revision; keep it unsupported rather than drop the markup.
            self.unsupported_character_properties = true;
        }
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
        let properties =
            self.run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?;
        let Some(mut run) = properties.direct_text_run(text, &self.fonts)? else {
            return Ok(None);
        };
        if let Some((author, date)) = properties.direct_insertion() {
            let revision = self.direct_insertion_revision(author, date)?;
            if let Some(wire) = run.typography_acquisition.as_mut() {
                wire.revision = Some(docx_model::RevisionTypographyWire {
                    kind: revision.kind.clone(),
                    id: docx_model::TypographyValueWire::default(),
                    author: revision.author.clone(),
                    date: revision.date.clone(),
                });
            }
            run.revision = Some(revision);
        }
        Ok(Some(run))
    }

    /// A visible phonetic-guide run and its resolved default language, the
    /// language a DOCX `w:rt` run states in `w:lang/@w:val` (ECMA-376
    /// 17.3.2.20) and the only place the DOCX model carries that axis.
    pub(in crate::doc) fn direct_ruby_guide_run(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Option<(TextRun, Option<&'static str>)>, String> {
        let Some(run) =
            self.direct_text_run(paragraph_style, table_style, fc, prm, prcs, String::new())?
        else {
            return Ok(None);
        };
        let languages = self
            .run_properties_with_table(paragraph_style, table_style, fc, prm, prcs)?
            .resolved_languages()?;
        Ok(Some((run, languages.default)))
    }

    /// MS-DOC 2.6.1 sprmCFRMarkIns with sprmCIbstRMark (index into
    /// SttbfRMark, 2.9.290) and sprmCDttmRMark (DTTM, 2.9.65) as the ECMA-376
    /// 17.13.5.18 `w:ins` provenance the DOCX model carries. MS-DOC has no
    /// revision identifier, so `id` stays absent.
    fn direct_insertion_revision(
        &self,
        author: Option<u16>,
        date: Option<u32>,
    ) -> Result<docx_model::RunRevision, String> {
        let authors = revision_authors(self.revision_authors.ok_or_else(|| {
            super::super::unsupported("Word revision author table outside table stream")
        })?)?;
        // "By default, this index is zero, which is the index of the
        // 'unknown' author."
        let index = usize::from(author.unwrap_or(0));
        let author = authors.get(index).cloned().ok_or_else(|| {
            super::super::unsupported("Word revision author index outside SttbfRMark")
        })?;
        Ok(docx_model::RunRevision {
            kind: "insertion".into(),
            id: None,
            author: Some(author),
            date: date.map(dttm).transpose()?.flatten(),
            typography_id: docx_model::TypographyValueWire::default(),
        })
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

/// MS-DOC 2.9.290 SttbfRMark: an extended (UTF-16) STTB without extra data.
fn revision_authors(bytes: &[u8]) -> Result<Vec<String>, String> {
    use super::super::{u16_at, unsupported};
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    if u16_at(bytes, 0)? != 0xffff || u16_at(bytes, 4)? != 0 {
        return Err(unsupported("invalid Word revision author table"));
    }
    let count = usize::from(u16_at(bytes, 2)?);
    let mut offset = 6;
    let mut authors = Vec::with_capacity(count.min(bytes.len() / 2));
    for _ in 0..count {
        let length = usize::from(u16_at(bytes, offset)?);
        offset += 2;
        let units = bytes
            .get(offset..offset + length * 2)
            .ok_or_else(|| unsupported("truncated Word revision author name"))?
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        authors.push(
            String::from_utf16(&units)
                .map_err(|_| unsupported("invalid Word revision author name"))?,
        );
        offset += length * 2;
    }
    Ok(authors)
}

/// MS-DOC 2.9.65 DTTM as an ECMA-376 `w:date` xsd:dateTime. All-zero means
/// no recorded date.
fn dttm(value: u32) -> Result<Option<String>, String> {
    if value == 0 {
        return Ok(None);
    }
    let minute = value & 0x3f;
    let hour = (value >> 6) & 0x1f;
    let day = (value >> 11) & 0x1f;
    let month = (value >> 16) & 0x0f;
    let year = 1900 + ((value >> 20) & 0x1ff);
    if minute > 59 || hour > 23 || !(1..=31).contains(&day) || !(1..=12).contains(&month) {
        return Err(super::super::unsupported("invalid Word revision date"));
    }
    Ok(Some(format!(
        "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:00Z"
    )))
}

#[cfg(test)]
mod revision_tests {
    use super::*;

    #[test]
    fn revision_author_table_and_dttm_decode_per_ms_doc() {
        let mut table = vec![0xff, 0xff, 2, 0, 0, 0];
        for name in ["Unknown", "Jim"] {
            table.extend((name.len() as u16).to_le_bytes());
            for unit in name.encode_utf16() {
                table.extend(unit.to_le_bytes());
            }
        }
        assert_eq!(revision_authors(&table).unwrap(), ["Unknown", "Jim"]);
        assert!(revision_authors(&table[..table.len() - 1]).is_err());
        assert!(revision_authors(&[0, 0, 0, 0, 0, 0]).is_err());
        assert!(revision_authors(&[]).unwrap().is_empty());
        // 0x86A12D2E: 2006-01-05 20:46 (Thursday).
        assert_eq!(
            dttm(0x86a1_2d2e).unwrap().as_deref(),
            Some("2006-01-05T20:46:00Z")
        );
        assert_eq!(dttm(0).unwrap(), None);
        assert!(dttm(0x0000_003c).is_err());
    }
}
