//! Native reserved-story acquisition, separate from shared note placement.
//! MS-DOC 2.3.3/2.8.22 define slots and guards; 2.6.1 sprmCFSpec gives U+0003
//! and U+0004 rule semantics only after the actual CHPX/style/PCD cascade.
//! Role does not select kind. Numeric rule geometry is a shared layout policy,
//! not acquired here. Supported CHPX/PAPX move into the shared reserved
//! separator wire; the shared consumer uses visible rule/control mark metrics.
//! This module owns the MS-DOC story shape (CPs, guard, PAPX/CHPX cascade) and
//! admits only the closed shapes that wire represents. Acquisition alone does
//! not admit custom content or establish Office-exact width/placement fidelity.

use super::super::{payload, ModelBudget};
use crate::doc::{formatting, headers, unsupported};
use docx_model::{
    NoteSeparatorMark, ReservedNoteSeparatorCharacterWire, ReservedNoteSeparatorParagraphWire,
    ReservedNoteSeparatorRuleWire, ReservedNoteSeparatorStoryWire,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum RuleKind {
    Short,
    Full,
}

/// CP relative to the header document. It stays in this module: a native
/// address must not be mistaken for a shared model source path.
pub(super) struct SourceAddress {
    pub(super) cp: usize,
}

pub(super) struct RuleFacts {
    pub(super) kind: Option<RuleKind>,
    pub(super) character: formatting::DirectRuleControlFacts,
    pub(super) source: SourceAddress,
}

pub(super) struct ParagraphFacts {
    pub(super) paragraph: docx_model::DocParagraph,
    pub(super) mark: Option<formatting::DirectRuleControlFacts>,
    pub(super) source: SourceAddress,
    numbered: bool,
    framed: bool,
}

pub(super) struct AcquiredSeparator<'a> {
    pub(super) source: headers::SeparatorStory<'a>,
    pub(super) rule: Option<RuleFacts>,
    pub(super) paragraph: Option<ParagraphFacts>,
}

impl RuleKind {
    fn mark(self) -> NoteSeparatorMark {
        match self {
            Self::Short => NoteSeparatorMark::Short,
            Self::Full => NoteSeparatorMark::Full,
        }
    }
}

impl AcquiredSeparator<'_> {
    /// Preserve explicit native rule kind independently of its destination role.
    pub(super) fn mark(&self) -> Option<NoteSeparatorMark> {
        if self.empty_story() {
            return Some(NoteSeparatorMark::None);
        }
        self.rule.as_ref()?.kind.map(RuleKind::mark)
    }

    /// Move the acquired facts of an admitted story into the shared wire.
    /// Owned paragraph/run payloads move without a clone, so their acquisition
    /// charges stand; the caller charges the new retained container. Borrowed
    /// header text is never copied, and native CPs stay here.
    ///
    /// MS-DOC 2.3.3 story shapes map onto the closed wire shapes: an empty or
    /// guard-only story has no paragraph; a paragraph-only story is exactly
    /// one paragraph mark before its guard; a rule story is exactly its
    /// control, the content paragraph mark and the guard. Anything else is a
    /// custom separator.
    /// Whether the story is one of the closed shapes of the shared wire
    /// (see `into_wire`).
    pub(super) fn wire_shape(&self) -> bool {
        let content = &self.source.content_cp;
        let mark_ends_content = |paragraph: &ParagraphFacts| paragraph.source.cp + 1 == content.end;
        match (self.source.guard_cp, self.source.content) {
            (_, "") => self.rule.is_none() && self.paragraph.is_none(),
            (Some(_), "\r") => {
                self.rule.is_none() && self.paragraph.as_ref().is_some_and(mark_ends_content)
            }
            (Some(guard), _) => {
                guard == content.end
                    && content.end - content.start == 2
                    && self
                        .rule
                        .as_ref()
                        .is_some_and(|rule| rule.source.cp == content.start)
                    && self.paragraph.as_ref().is_some_and(mark_ends_content)
            }
            _ => false,
        }
    }

    pub(super) fn into_wire(self) -> Result<ReservedNoteSeparatorStoryWire, String> {
        if !self.wire_shape() {
            return Err(unsupported("custom Word note separators are not supported"));
        }
        Ok(ReservedNoteSeparatorStoryWire {
            rule: self.rule.map(|rule| ReservedNoteSeparatorRuleWire {
                mark: rule.kind.map(RuleKind::mark),
                control: ReservedNoteSeparatorCharacterWire {
                    run: rule.character.run,
                    insertion: rule.character.insertion,
                },
            }),
            paragraph: self
                .paragraph
                .map(|paragraph| ReservedNoteSeparatorParagraphWire {
                    paragraph: paragraph.paragraph,
                    content_mark: paragraph
                        .mark
                        .map(|mark| ReservedNoteSeparatorCharacterWire {
                            run: mark.run,
                            insertion: mark.insertion,
                        }),
                    numbered: paragraph.numbered,
                    framed: paragraph.framed,
                }),
        })
    }

    /// The shared native separator consumer has metric participants, not
    /// authored text, numbering/frame cascades, or tracked-insertion payloads.
    /// sprmCSymbol (MS-DOC 2.6.1/2.9.47) can populate an otherwise empty run;
    /// check the acquired payload before admitting the new short/short class.
    /// Structure and passive-special visibility are checked by standard_rule.
    pub(super) fn metric_only(&self) -> bool {
        let character = |facts: &formatting::DirectRuleControlFacts| {
            !facts.insertion && facts.run.as_ref().is_none_or(|run| run.text.is_empty())
        };
        self.rule
            .as_ref()
            .is_none_or(|rule| character(&rule.character))
            && self.paragraph.as_ref().is_none_or(|paragraph| {
                !paragraph.numbered
                    && !paragraph.framed
                    && paragraph.mark.as_ref().is_some_and(character)
            })
    }

    /// MS-DOC 2.3.3: equal PlcfHdd CPs mean no contents and no guard.
    /// This is explicit native absence, not an omitted/default OOXML story.
    pub(super) fn empty_story(&self) -> bool {
        self.source.cp.is_empty()
            && self.source.guard_cp.is_none()
            && self.source.content.is_empty()
            && self.rule.is_none()
            && self.paragraph.is_none()
    }

    /// MS-DOC 2.3.3: a non-empty story whose only content is one paragraph
    /// mark before its guard, with no U+0003/U+0004 control. Word lays the
    /// story out as that empty paragraph: no rule ink, one line of the
    /// paragraph's own height and spacing. The shared wire represents it as
    /// a rule-free paragraph story; admission additionally requires the
    /// metric-only facts the consumer represents (no numbering, frame,
    /// tracked insertion or symbol text on the mark).
    pub(super) fn paragraph_only(&self) -> bool {
        self.source.guard_cp.is_some()
            && self.source.content == "\r"
            && self.rule.is_none()
            && self
                .paragraph
                .as_ref()
                .is_some_and(|paragraph| paragraph.source.cp + 1 == self.source.content_cp.end)
            && self.metric_only()
    }

    pub(super) fn standard_rule(&self, kind: RuleKind) -> bool {
        let expected = match kind {
            RuleKind::Short => "\u{3}\r",
            RuleKind::Full => "\u{4}\r",
        };
        self.source.cp.end - self.source.cp.start == 3
            && self.source.guard_cp.is_some()
            && self.source.content == expected
            && self.rule.as_ref().is_some_and(|rule| {
                rule.source.cp == self.source.content_cp.start
                    && rule.kind == Some(kind)
                    && rule.character.picture.passive_special()
                    && rule.character.run.is_some()
                    && !rule.character.insertion
            })
            && self.paragraph.as_ref().is_some_and(|paragraph| {
                paragraph.source.cp + 1 == self.source.content_cp.end
                    && !paragraph.numbered
                    && !paragraph.framed
                    && paragraph.mark.as_ref().is_some_and(|mark| !mark.insertion)
            })
    }
}

pub(super) fn acquire<'a>(
    headers: &'a headers::Headers<'_>,
    slot: usize,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<AcquiredSeparator<'a>, String> {
    budget.charge(std::mem::size_of::<AcquiredSeparator<'_>>())?;
    let source = headers.separator(slot);
    let marker = match source.content {
        "\u{3}" | "\u{3}\r" => Some(RuleKind::Short),
        "\u{4}" | "\u{4}\r" => Some(RuleKind::Full),
        _ => None,
    };
    let mut paragraph_style = None;
    let paragraph = if marker.is_some() || source.content == "\r" {
        // A content paragraph mark owns paragraph formatting. A separator need
        // not contain a whole paragraph; for that distinct, currently gated
        // class query its control's paragraph range, never the guard's style.
        let paragraph_cp = if source.content.ends_with('\r') {
            source.content_cp.end - 1
        } else {
            source.content_cp.start
        };
        let (_, mark_fc, mark_piece) = headers
            .story
            .position(paragraph_cp)
            .ok_or_else(|| unsupported("Word separator paragraph outside header story"))?;
        let style = formatting.paragraph_style(mark_fc)?;
        paragraph_style = Some(style);
        let paragraph = formatting.direct_paragraph(
            style,
            None,
            mark_fc,
            mark_piece.prm,
            &headers.story.prcs,
        )?;
        // Numbering marker properties are a deferred, owned cascade. A rule
        // does not activate numbering: retain its presence, not that uncharged
        // marker payload, alongside the charged visible paragraph metadata.
        let numbered = paragraph.numbering.is_some();
        let framed = paragraph.frame_gap.is_some() || paragraph.native_frame.is_some();
        let paragraph = paragraph.paragraph;
        budget.paragraph(&paragraph)?;
        // Retain the complete supported CHPX of the content mark separately
        // from paragraph font facts. The terminal guard is never its substitute.
        // A partial rule has no authored content mark and keeps that distinction.
        let mark = if source.content.ends_with('\r') {
            let mark = formatting.direct_rule_control(
                style,
                mark_fc,
                mark_piece.prm,
                &headers.story.prcs,
            )?;
            if let Some(run) = &mark.run {
                budget.charge(payload::text_run(run)?)?;
            }
            Some(mark)
        } else {
            None
        };
        Some(ParagraphFacts {
            paragraph,
            mark,
            numbered,
            framed,
            source: SourceAddress { cp: paragraph_cp },
        })
    } else {
        None
    };
    let rule = if let Some(marker) = marker {
        let style = paragraph_style
            .ok_or_else(|| unsupported("Word separator paragraph style is missing"))?;
        let (_, fc, piece) = headers
            .story
            .position(source.content_cp.start)
            .ok_or_else(|| unsupported("Word separator control outside header story"))?;
        let character =
            formatting.direct_rule_control(style, fc, piece.prm, &headers.story.prcs)?;
        if let Some(run) = &character.run {
            budget.charge(payload::text_run(run)?)?;
        }
        Some(RuleFacts {
            kind: character.picture.special.then_some(marker),
            character,
            source: SourceAddress {
                cp: source.content_cp.start,
            },
        })
    } else {
        None
    };
    Ok(AcquiredSeparator {
        source,
        rule,
        paragraph,
    })
}
