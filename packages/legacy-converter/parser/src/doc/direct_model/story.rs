//! Shared direct projection for one already-tokenized Word story slice.
//!
//! Section ownership and header-field restoration remain with their respective
//! callers; PAP/CHPX resolution and hard-break normalization live here once.

use super::{ModelBudget, ParaPiece};
use crate::doc::{formatting, pictures, unsupported, Paragraph, Story, Token};
use docx_model::paragraph_breaks::visit_para_on_page_breaks;
use docx_model::{BodyElement, BreakType, DocRun, ImageRun};

pub(super) fn project(
    story: &Story<'_>,
    paragraphs: Vec<Paragraph>,
    formatting: &mut formatting::Formatting<'_>,
    pictures: &mut pictures::Store<'_>,
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
                Token::Picture => {
                    let (_, fc, piece) = story
                        .position(cp)
                        .ok_or_else(|| unsupported("Word picture outside piece table"))?;
                    let facts = formatting.direct_inline_picture_facts(
                        style,
                        fc,
                        piece.prm,
                        &story.prcs,
                    )?;
                    if !facts.vanish {
                        let offset = facts.location.ok_or_else(|| {
                            unsupported("visible Word picture has no inline location")
                        })?;
                        let image: Option<pictures::DirectInlinePicture> =
                            pictures.direct_inline(offset, &mut budget.remaining_bytes)?;
                        if let Some(image) = image {
                            let mime_type = image.mime_type.to_string();
                            budget
                                .charge(std::mem::size_of::<ImageRun>() + mime_type.capacity())?;
                            budget.push(
                                &mut paragraph.runs,
                                DocRun::Image(Box::new(ImageRun {
                                    image_path: image.resource_key,
                                    mime_type,
                                    svg_image_path: None,
                                    src_rect: image.crop,
                                    width_pt: image.width_pt,
                                    height_pt: image.height_pt,
                                    rotation: image.rotation,
                                    flip_h: image.flip_h,
                                    flip_v: image.flip_v,
                                    anchor: false,
                                    anchor_x_pt: 0.0,
                                    anchor_y_pt: 0.0,
                                    anchor_x_from_margin: false,
                                    anchor_y_from_para: false,
                                    color_replace_from: None,
                                    duotone: None,
                                    alpha: None,
                                    wrap_mode: None,
                                    dist_top: 0.0,
                                    dist_bottom: 0.0,
                                    dist_left: 0.0,
                                    dist_right: 0.0,
                                    wrap_side: None,
                                    allow_overlap: true,
                                    anchor_x_align: None,
                                    anchor_y_align: None,
                                    anchor_x_relative_from: None,
                                    anchor_y_relative_from: None,
                                    anchor_acquisition: None,
                                })),
                            )?;
                        }
                    }
                }
                Token::FloatingPicture => {
                    return Err(unsupported(
                        "direct DOC model does not yet support floating pictures",
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
