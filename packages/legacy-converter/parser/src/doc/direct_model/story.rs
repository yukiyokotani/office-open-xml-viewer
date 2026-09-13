//! Shared direct projection for one already-tokenized Word story slice.
//!
//! Section ownership and header-field restoration remain with their respective
//! callers; PAP/CHPX resolution and hard-break normalization live here once.

use super::{
    tables::{Block, Blocks, Writer},
    ModelBudget, ParaPiece,
};
use crate::doc::{
    floating, formatting, numbering, pictures, table, table_context, table_style_condition,
    unsupported, Paragraph, Story, Token,
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
        // Resolve the row-owned style and cell conditions once per original
        // paragraph. Every paragraph mark, run, control, and picture host then
        // shares the same compact formatting identity.
        let table_style = if let Some(context) = context {
            let matches = if let (Some(_), Some(options)) =
                (context.source_cell_index, context.table_style_options)
            {
                let (horizontal, vertical) = formatting.table_style_bands(context.table_style)?;
                let options = table_style_condition::Options::new(options, horizontal, vertical)?;
                let column = table_context
                    .logical_column(paragraph_index)?
                    .ok_or_else(|| unsupported("Word table cell lacks source column context"))?;
                table_style_condition::select(
                    &table_context,
                    context.table_id,
                    context.row_index,
                    column,
                    options,
                )?
            } else {
                [None; 5]
            };
            formatting.table_formatting_key(
                context.table_style,
                context.source_cell_index.and(context.table_style_options),
                matches,
            )?
        } else {
            None
        };
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
    table_style: Option<formatting::TableFormattingKey>,
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

#[cfg(test)]
mod tests {
    use super::super::table_tests::with_papx;
    use super::*;
    use crate::cfb::{test_support::build_cfb, CompoundFile};
    use crate::doc::{with_acquired_doc, Fields};
    use docx_model::CellElement;

    fn sprm(code: u16, operand: &[u8]) -> Vec<u8> {
        [code.to_le_bytes().as_slice(), operand].concat()
    }

    fn table_style_sheet() -> Vec<u8> {
        let mut header = vec![0; 18];
        header[0..2].copy_from_slice(&15u16.to_le_bytes());
        header[2..4].copy_from_slice(&10u16.to_le_bytes());
        let mut bytes = Vec::new();
        bytes.extend(18u16.to_le_bytes());
        bytes.extend(header);

        let mut normal = vec![0; 14];
        normal[2..4].copy_from_slice(&0xfff1u16.to_le_bytes());
        bytes.extend((normal.len() as u16).to_le_bytes());
        bytes.extend(normal);

        let mut style = vec![0; 14];
        style[2..4].copy_from_slice(&0xfff3u16.to_le_bytes());
        style[4..6].copy_from_slice(&3u16.to_le_bytes());
        for set in [
            &[0x88, 0x34, 1][..],
            &[1, 0][..],
            &[
                0x42, 0x2a, 1, // unconditional black
                0x85, 0xca, 5, 0x40, 0, 0x42, 0x2a, 6, // odd row red
                0x85, 0xca, 5, 0x80, 0, 0x42, 0x2a, 2, // even row blue
            ][..],
        ] {
            style.extend((set.len() as u16).to_le_bytes());
            style.extend(set);
            if set.len() % 2 != 0 {
                style.push(0);
            }
        }
        let size = style.len() as u16;
        style[6..8].copy_from_slice(&size.to_le_bytes());
        bytes.extend(size.to_le_bytes());
        bytes.extend(style);
        for _ in 2..15 {
            bytes.extend(0u16.to_le_bytes());
        }
        bytes
    }

    fn with_table_style(source: &[u8]) -> Vec<u8> {
        let cfb = CompoundFile::open(source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let stylesheet = table_style_sheet();
        word[0xa2..0xa6].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0xa6..0xaa].copy_from_slice(&(stylesheet.len() as u32).to_le_bytes());
        table.extend(stylesheet);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    fn cell() -> Vec<u8> {
        sprm(0x2416, &[1])
    }

    fn row(options: u16) -> Vec<u8> {
        [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, 1, 0xe8, 3]),
            sprm(0x563a, &1u16.to_le_bytes()),
            sprm(0x740a, &[0, 0, options as u8, (options >> 8) as u8]),
            sprm(0x2416, &[1]),
        ]
        .concat()
    }

    #[test]
    fn project_uses_ttp_options_and_source_rows_for_conditional_color_keys() {
        let text = "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}\r";
        let source = super::super::tests::source_with_typography(
            text,
            &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
            None,
            None,
            None,
            None,
        );
        let source = with_table_style(&source);
        let source = with_papx(
            &source,
            &[
                (0, 2, cell()),
                (2, 3, row(0)),
                (3, 5, cell()),
                (5, 6, row(0)),
                (6, 8, cell()),
                (8, 9, row(1 << 9)),
                (9, 10, Vec::new()),
            ],
        );
        let cfb = CompoundFile::open(&source).unwrap();
        with_acquired_doc(&cfb, |mut facts| {
            let paragraphs = crate::doc::tokenize_with_fields(
                &facts.story.text,
                &mut Fields::default(),
                0,
                true,
            );
            let mut numbering = numbering::direct::Store::default();
            numbering.begin_story()?;
            let mut budget = ModelBudget::new(1_000_000);
            let mut body = Vec::new();
            let mut table_sequence = 0;
            project(
                &facts.story,
                paragraphs,
                &mut facts.formatting,
                &mut numbering,
                &mut facts.pictures,
                Some(&mut facts.floating),
                &mut budget,
                &mut body,
                None,
                &mut table_sequence,
            )?;

            let BodyElement::Table(table) = &body[0] else {
                panic!("table")
            };
            let colors: Vec<_> = table
                .rows
                .iter()
                .map(|row| {
                    let CellElement::Paragraph(paragraph) = &row.cells[0].content[0] else {
                        panic!("paragraph")
                    };
                    let DocRun::Text(run) = &paragraph.runs[0] else {
                        panic!("text")
                    };
                    run.color.as_deref().unwrap()
                })
                .collect();
            assert_eq!(colors, ["ff0000", "0000ff", "000000"]);
            // TIstd and TTlp remain deliberately admission-gated even though
            // this internal projection verifies their acquired context.
            assert!(facts.formatting.unsupported_table_properties);
            Ok(())
        })
        .unwrap();
    }
}
