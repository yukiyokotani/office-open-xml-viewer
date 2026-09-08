//! Direct projection of authored header/footer stories.
//!
//! MS-DOC 2.3.3 and 2.8.22 encode six slots per section after six separator
//! stories. A zero-width slot inherits its prior entry; an authored blank range
//! is a real replacement and therefore projects to an empty paragraph.

use super::{story, ModelBudget};
use crate::doc::{formatting, headers, pictures, tokenize_with_fields, Fields};
use docx_model::{HeaderFooter, HeadersFooters};

pub(super) struct Resolver<'a, 'h> {
    source: Option<&'h headers::Headers<'a>>,
    inherited: [Option<&'h headers::Entry>; 6],
}

impl<'a, 'h> Resolver<'a, 'h> {
    pub(super) fn new(source: Option<&'h headers::Headers<'a>>) -> Self {
        Self {
            source,
            inherited: [None; 6],
        }
    }

    /// Advance to one section, retaining source-entry identity rather than
    /// cloning already-projected model trees. Each required retained instance is
    /// projected under the document's single output budget.
    pub(super) fn project_section(
        &mut self,
        section: usize,
        formatting: &mut formatting::Formatting<'a>,
        pictures: &mut pictures::Store<'a>,
        budget: &mut ModelBudget,
    ) -> Result<(HeadersFooters, HeadersFooters), String> {
        if let Some(source) = self.source {
            for (slot, inherited) in self.inherited.iter_mut().enumerate() {
                if let Some(entry) = source.entry(section * 6 + slot) {
                    *inherited = Some(entry);
                }
            }
        }

        let mut projected = Vec::with_capacity(6);
        for entry in self.inherited {
            projected.push(
                entry
                    .map(|entry| self.project_entry(entry, formatting, pictures, budget))
                    .transpose()?,
            );
        }
        Ok((
            HeadersFooters {
                even: projected[0].take(),
                default: projected[1].take(),
                first: projected[4].take(),
            },
            HeadersFooters {
                even: projected[2].take(),
                default: projected[3].take(),
                first: projected[5].take(),
            },
        ))
    }

    fn project_entry(
        &self,
        entry: &headers::Entry,
        formatting: &mut formatting::Formatting<'a>,
        pictures: &mut pictures::Store<'a>,
        budget: &mut ModelBudget,
    ) -> Result<HeaderFooter, String> {
        let source = self.source.expect("entry belongs to a header source");
        let text = source.entry_text(entry);
        // Inherited entries are intentionally re-projected so no infallible
        // deep clone can allocate before admission. Charge their source bytes as
        // cumulative scratch work before each tokenization as well: a long
        // hidden-only header otherwise has tiny retained output but can multiply
        // scanning and temporary allocation by the section count.
        budget.charge(text.len())?;
        let mut paragraphs = tokenize_with_fields(text, &mut Fields::default(), entry.cp, true);
        source.restore_fields(text, entry.cp, &mut paragraphs);
        let mut body = Vec::new();
        story::project(
            &source.story,
            paragraphs,
            formatting,
            pictures,
            budget,
            &mut body,
            None,
        )?;
        Ok(HeaderFooter { body })
    }
}
