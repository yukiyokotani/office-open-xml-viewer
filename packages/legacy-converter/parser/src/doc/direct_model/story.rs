//! Shared direct projection for one already-tokenized Word story slice.
//!
//! Section ownership and header-field restoration remain with their respective
//! callers; PAP/CHPX resolution and hard-break normalization live here once.

use super::{
    tables::{Block, Blocks, Writer},
    ModelBudget, ParaPiece,
};
use crate::doc::{
    floating, formatting, numbering, pictures, table, table_context, unsupported, Paragraph, Story,
    Token,
};
use docx_model::paragraph_breaks::visit_para_on_page_breaks;
use docx_model::{BodyElement, BreakType, DocRun, ImageRun};

pub(super) fn project(
    story: &Story<'_>,
    paragraphs: Vec<Paragraph>,
    formatting: &mut formatting::Formatting<'_>,
    numbering: &mut numbering::direct::Store,
    pictures: &mut pictures::Store<'_>,
    mut floating: Option<&mut floating::Store<'_>>,
    budget: &mut ModelBudget,
    body: &mut Vec<BodyElement>,
    ending_kind: Option<&str>,
    table_sequence: &mut usize,
) -> Result<(), String> {
    let paragraph_count = paragraphs.len();
    let mut prepared = Vec::new();
    for (paragraph_index, source) in paragraphs.into_iter().enumerate() {
        let (_, mark_fc, mark_piece) = story
            .position(source.end_cp)
            .ok_or_else(|| unsupported("Word paragraph mark outside piece table"))?;
        let table_properties = formatting.table_properties(mark_fc, mark_piece.prm, &story.prcs)?;
        let table_depth = table_properties.depth()?;
        if source.mark == '\u{7}' && table_depth == 0 {
            return Err(unsupported("Word table cell mark outside table"));
        }
        if ending_kind.is_some() && paragraph_index + 1 == paragraph_count && table_depth != 0 {
            return Err(unsupported("Word section break inside table"));
        }
        // This vector is additional retained state compared with the streaming
        // producer. ModelBudget::push charges the complete outer entry,
        // including Properties<Row>; separately charge its heap-owned cells,
        // identity operands and decoded border payload before retaining it.
        budget.charge(table_properties.retained_heap_bytes()?)?;
        budget.push(
            &mut prepared,
            PreparedParagraph {
                source,
                mark_fc,
                mark_prm: mark_piece.prm,
                table_properties,
            },
        )?;
    }

    let table_context = table_context::Index::build_with_styles(
        paragraph_count,
        prepared
            .iter()
            .map(|value| (value.table_properties.borrowed(), value.source.mark)),
        &mut |selected| Ok(formatting.resolve_table_style_id(selected)),
        &mut |bytes| budget.charge(bytes),
    )?;
    let mut tables = Writer::new(table_sequence);
    for (paragraph_index, prepared) in prepared.into_iter().enumerate() {
        let PreparedParagraph {
            source,
            mark_fc,
            mark_prm,
            table_properties,
        } = prepared;
        let style = formatting.paragraph_style(mark_fc)?;
        let table_depth = table_properties.depth()?;
        let context = table_context.paragraph(paragraph_index)?;
        if context.is_some() != (table_depth != 0) {
            return Err(unsupported(
                "Word table context disagrees with paragraph depth",
            ));
        }
        // Read the resolved row-owned selection from the production index.
        // The supported style subset is projected here, while TIstd continues
        // to trip the broader table-format admission gate.
        let table_style = context.and_then(|value| value.table_style);
        let direct =
            formatting.direct_paragraph(style, table_style, mark_fc, mark_prm, &story.prcs)?;
        let mut paragraph = direct.paragraph;
        if let Some((reference, marker)) = direct.numbering {
            paragraph.numbering = Some(Box::new(
                formatting.direct_numbering(numbering, reference, &marker, &paragraph)?,
            ));
        }
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
                            formatting.direct_text_run(
                                style,
                                table_style,
                                fc,
                                prm,
                                &story.prcs,
                                String::new(),
                            )
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
                    push_control_text(
                        &mut paragraph,
                        story,
                        formatting,
                        style,
                        table_style,
                        cp,
                        "\t",
                        budget,
                    )?;
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
                        table_style,
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
                    let store = floating.as_deref_mut().ok_or_else(|| {
                        unsupported("direct DOC model does not support header floating pictures")
                    })?;
                    let (_, fc, piece) = story
                        .position(cp)
                        .ok_or_else(|| unsupported("Word floating picture outside piece table"))?;
                    let Some(mut host) = formatting.direct_anchor_host_metrics(
                        style,
                        table_style,
                        fc,
                        piece.prm,
                        &story.prcs,
                    )?
                    else {
                        continue;
                    };
                    let image: Option<floating::DirectFloatingPicture> =
                        store.direct_picture(cp, &mut budget.remaining_bytes)?;
                    if let Some(image) = image {
                        host.anchor_occurrence_id = Some(image.occurrence_id);
                        let host_payload = std::mem::size_of::<docx_model::AnchorHostMetrics>()
                            .checked_add(host.font_family.as_ref().map_or(0, String::capacity))
                            .and_then(|bytes| {
                                bytes.checked_add(
                                    host.font_family_east_asia
                                        .as_ref()
                                        .map_or(0, String::capacity),
                                )
                            })
                            .ok_or("OUTPUT_TOO_LARGE")?;
                        budget.charge(host_payload)?;
                        budget.push(&mut paragraph.runs, DocRun::AnchorHost(host))?;
                        budget.push(&mut paragraph.runs, DocRun::Image(Box::new(image.image)))?;
                    }
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

        let mut blocks = Blocks::default();
        if table_depth != 0 {
            budget.push(&mut blocks.0, Block::Paragraph(Box::new(paragraph)))?;
        } else {
            match paragraph.runs.as_slice() {
                [DocRun::Break {
                    break_type: BreakType::Page,
                }] => {
                    let subsumed = paragraph_index + 1 == paragraph_count
                        && ending_kind
                            .is_some_and(|kind| !matches!(kind, "continuous" | "nextColumn"));
                    if !subsumed {
                        budget.push(
                            &mut blocks.0,
                            Block::PageBreak {
                                same_paragraph_as_previous: None,
                            },
                        )?;
                    }
                }
                [DocRun::Break {
                    break_type: BreakType::Column,
                }] => budget.push(&mut blocks.0, Block::ColumnBreak)?,
                _ => visit_para_on_page_breaks(paragraph, |piece| {
                    let element = match piece {
                        ParaPiece::Para(paragraph) => {
                            budget.normalized_paragraph(&paragraph)?;
                            Block::Paragraph(Box::new(paragraph))
                        }
                        ParaPiece::PageBreak {
                            same_paragraph_as_previous,
                        } => Block::PageBreak {
                            same_paragraph_as_previous: same_paragraph_as_previous.then_some(true),
                        },
                        ParaPiece::ColumnBreak => Block::ColumnBreak,
                    };
                    budget.push(&mut blocks.0, element)
                })?,
            }
        }
        tables.push(table_properties, source.mark, blocks, budget)?;
    }
    for block in tables.finish(budget)?.0 {
        let element = match block {
            Block::Paragraph(value) => BodyElement::Paragraph(value),
            Block::Table(value) => BodyElement::Table(value),
            Block::PageBreak {
                same_paragraph_as_previous,
            } => BodyElement::PageBreak {
                parity: None,
                same_paragraph_as_previous,
            },
            Block::ColumnBreak => BodyElement::ColumnBreak,
        };
        budget.push(body, element)?;
    }
    Ok(())
}

struct PreparedParagraph {
    source: Paragraph,
    mark_fc: usize,
    mark_prm: u16,
    table_properties: table::Properties,
}

fn push_control_text(
    paragraph: &mut docx_model::DocParagraph,
    story: &Story<'_>,
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    table_style: Option<usize>,
    cp: usize,
    text: &str,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let (_, fc, piece) = story
        .position(cp)
        .ok_or_else(|| unsupported("Word control outside piece table"))?;
    if let Some(mut run) = formatting.direct_text_run(
        style,
        table_style,
        fc,
        piece.prm,
        &story.prcs,
        String::new(),
    )? {
        budget.text(&mut paragraph.runs, &mut run, text)?;
    }
    Ok(())
}
