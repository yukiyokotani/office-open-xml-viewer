//! Shared direct projection for one already-tokenized Word story slice.
//!
//! Section ownership and header-field restoration remain with their respective
//! callers; PAP/CHPX resolution and hard-break normalization live here once.

use super::{ModelBudget, ParaPiece};
use crate::doc::{formatting, unsupported, Paragraph, Story, Token};
use docx_model::paragraph_breaks::visit_para_on_page_breaks;
use docx_model::{BodyElement, BreakType, DocRun};

pub(super) fn project(
    story: &Story<'_>,
    paragraphs: Vec<Paragraph>,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
    body: &mut Vec<BodyElement>,
    ending_kind: Option<&str>,
) -> Result<(), String> {
    let paragraph_count = paragraphs.len();
    for (paragraph_index, source) in paragraphs.into_iter().enumerate() {
        if source.mark == '\u{7}' {
            return Err(unsupported("direct DOC model does not yet support tables"));
        }
        let (_, mark_fc, mark_piece) = story
            .position(source.end_cp)
            .ok_or_else(|| unsupported("Word paragraph mark outside piece table"))?;
        let style = formatting.paragraph_style(mark_fc)?;
        if formatting
            .table_properties(mark_fc, mark_piece.prm, &story.prcs)?
            .depth()?
            != 0
        {
            return Err(unsupported("direct DOC model does not yet support tables"));
        }
        let direct = formatting.direct_paragraph(style, mark_fc, mark_piece.prm, &story.prcs)?;
        if direct.numbering.is_some() {
            return Err(unsupported(
                "direct DOC model does not yet support numbered paragraphs",
            ));
        }
        let mut paragraph = direct.paragraph;
        budget.paragraph(&paragraph)?;

        for (token, cp) in source.tokens {
            match token {
                Token::Text(text) => {
                    super::super::visit_text_runs(
                        &text,
                        cp,
                        story,
                        &mut Some(&mut *formatting),
                        |formatting, fc, prm| {
                            formatting.direct_text_run(style, fc, prm, &story.prcs, String::new())
                        },
                        |part, run| {
                            if let Some(mut run) = run.flatten() {
                                budget.text(&mut paragraph.runs, &mut run, part)?;
                            }
                            Ok(())
                        },
                    )?;
                }
                Token::Tab => {
                    push_control_text(&mut paragraph, story, formatting, style, cp, "\t", budget)?;
                }
                Token::LineBreak => budget.push(
                    &mut paragraph.runs,
                    DocRun::Break {
                        break_type: BreakType::Line,
                    },
                )?,
                Token::PageBreak | Token::ColumnBreak => budget.push(
                    &mut paragraph.runs,
                    DocRun::Break {
                        break_type: if matches!(token, Token::PageBreak) {
                            BreakType::Page
                        } else {
                            BreakType::Column
                        },
                    },
                )?,
                Token::Picture | Token::FloatingPicture => {
                    return Err(unsupported(
                        "direct DOC model does not yet support pictures",
                    ));
                }
                Token::NoteMarker | Token::NoteReference(_) => {
                    return Err(unsupported(
                        "direct DOC model does not yet support note content",
                    ));
                }
                Token::FieldBegin(_) | Token::FieldEnd => {
                    return Err(unsupported(
                        "direct DOC model does not retain field structures yet",
                    ));
                }
            }
        }

        match paragraph.runs.as_slice() {
            [DocRun::Break {
                break_type: BreakType::Page,
            }] => {
                let subsumed = paragraph_index + 1 == paragraph_count
                    && ending_kind.is_some_and(|kind| !matches!(kind, "continuous" | "nextColumn"));
                if !subsumed {
                    budget.push(
                        body,
                        BodyElement::PageBreak {
                            parity: None,
                            same_paragraph_as_previous: None,
                        },
                    )?;
                }
            }
            [DocRun::Break {
                break_type: BreakType::Column,
            }] => budget.push(body, BodyElement::ColumnBreak)?,
            _ => visit_para_on_page_breaks(paragraph, |piece| {
                let element = match piece {
                    ParaPiece::Para(paragraph) => {
                        budget.normalized_paragraph(&paragraph)?;
                        BodyElement::Paragraph(Box::new(paragraph))
                    }
                    ParaPiece::PageBreak {
                        same_paragraph_as_previous,
                    } => BodyElement::PageBreak {
                        parity: None,
                        same_paragraph_as_previous: same_paragraph_as_previous.then_some(true),
                    },
                    ParaPiece::ColumnBreak => BodyElement::ColumnBreak,
                };
                budget.push(body, element)
            })?,
        }
    }
    Ok(())
}

fn push_control_text(
    paragraph: &mut docx_model::DocParagraph,
    story: &Story<'_>,
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    cp: usize,
    text: &str,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let (_, fc, piece) = story
        .position(cp)
        .ok_or_else(|| unsupported("Word control outside piece table"))?;
    if let Some(mut run) =
        formatting.direct_text_run(style, fc, piece.prm, &story.prcs, String::new())?
    {
        budget.text(&mut paragraph.runs, &mut run, text)?;
    }
    Ok(())
}
