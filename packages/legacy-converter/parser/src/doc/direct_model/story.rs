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

mod borders;
mod margins;
mod preferences;

pub(super) fn project(
    story: &Story<'_>,
    paragraphs: Vec<Paragraph>,
    formatting: &mut formatting::Formatting<'_>,
    numbering: &mut numbering::direct::Store,
    pictures: &mut pictures::Store<'_>,
    mut floating: Option<(&mut floating::Store<'_>, floating::Part)>,
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
        let table_properties =
            formatting.table_properties_native(mark_fc, mark_piece.prm, &story.prcs)?;
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
    margins::resolve(&mut prepared, &table_context, formatting)?;
    borders::resolve(&mut prepared, &table_context, formatting, budget)?;
    preferences::resolve(&mut prepared, &table_context, formatting)?;
    if formatting.use_raw_table_shading() {
        resolve_table_cell_shading(&mut prepared, &table_context, formatting)?;
    }
    // Position of the depth-1 row that encloses each table paragraph (its own
    // row, or the outer row of a nested table): its TTP is the next depth-1
    // row mark ([MS-DOC] 2.4.3). Cell paragraphs may repeat that position as
    // paragraph frame properties (see table::Position::matches_cell_frame).
    let mut outer_row_positions = vec![None; prepared.len()];
    let mut next_outer: Option<&table::Position> = None;
    for (index, value) in prepared.iter().enumerate().rev() {
        let depth = value.table_properties.depth()?;
        if depth == 1 && value.table_properties.row_end {
            next_outer = Some(&value.table_properties.row.position);
        } else if depth == 0 {
            next_outer = None;
        }
        if depth != 0 {
            outer_row_positions[index] = next_outer.cloned();
        }
    }
    // Only the main story passes a floating-drawing store.
    let mut tables = Writer::with_positioned_tables(table_sequence, floating.is_some());
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
                let (horizontal, vertical, presence) =
                    formatting.table_style_selector_profile(context.table_style)?;
                let options =
                    table_style_condition::Options::new(options, horizontal, vertical, presence)?;
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
        match (context, direct.table_frame) {
            (Some(_), Some(frame)) => {
                if outer_row_positions[paragraph_index]
                    .as_ref()
                    .is_some_and(|position| position.matches_cell_frame(frame))
                {
                    // The positioned table itself carries this placement.
                    paragraph.frame_pr = None;
                } else {
                    // The DOCX renderer positions frames only in the body
                    // flow; a framed cell paragraph would silently lay out in
                    // flow.
                    formatting.unsupported_paragraph_properties = true;
                }
            }
            _ if direct.frame_gap => formatting.unsupported_paragraph_properties = true,
            _ => {}
        }
        if let Some((reference, marker)) = direct.numbering {
            paragraph.numbering = Some(Box::new(
                formatting.direct_numbering(numbering, reference, &marker, &paragraph)?,
            ));
        }
        budget.paragraph(&paragraph)?;

        for (token, cp) in source.tokens {
            let (token, link) = match token {
                Token::Linked(linked) => {
                    let linked = *linked;
                    (linked.token, Some(linked.link))
                }
                token => (token, None),
            };
            let link = link.as_ref();
            match token {
                Token::Text(text) => {
                    super::super::visit_text_runs(
                        &text,
                        cp,
                        story,
                        &mut Some(&mut *formatting),
                        |formatting, fc, prm| {
                            linked_text_run(
                                formatting,
                                style,
                                table_style,
                                fc,
                                prm,
                                &story.prcs,
                                link,
                            )
                        },
                        |part, run| {
                            if let Some(mut run) = run.flatten() {
                                let part = crate::doc::character::Properties::direct_run_text(
                                    &mut run, part,
                                )?;
                                budget.text(&mut paragraph.runs, &mut run, &part)?;
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
                        link,
                        budget,
                    )?;
                }
                Token::LineBreak => {
                    let (_, fc, piece) = story
                        .position(cp)
                        .ok_or_else(|| unsupported("Word line break outside piece table"))?;
                    if formatting.direct_line_break_clears(
                        style,
                        table_style,
                        fc,
                        piece.prm,
                        &story.prcs,
                    )? {
                        formatting.unsupported_character_properties = true;
                    }
                    budget.push(
                        &mut paragraph.runs,
                        DocRun::Break {
                            break_type: BreakType::Line,
                        },
                    )?
                }
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
                    // Textbox stories cannot host floating drawings; Word
                    // anchors them only in the main and header documents.
                    let Some((store, part)) = floating.as_mut() else {
                        return Err(unsupported(
                            "direct DOC model found a floating drawing in a textbox story",
                        ));
                    };
                    let part = *part;
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
                    let Some(drawing) =
                        store.direct_drawing(part, cp, &mut budget.remaining_bytes)?
                    else {
                        continue;
                    };
                    let mut runs = Vec::new();
                    for run in drawing.runs {
                        runs.push(match run {
                            floating::DirectRun::Image(image) => DocRun::Image(image),
                            floating::DirectRun::Shape(mut shape) => {
                                if let Some(index) = shape.text {
                                    let textboxes = store.textbox(part).ok_or_else(|| {
                                        unsupported("Word shape text lacks its textbox story")
                                    })?;
                                    shape.shape.text_box_content = textbox_content(
                                        textboxes,
                                        index,
                                        shape.spid,
                                        formatting,
                                        pictures,
                                        budget,
                                        tables.sequence(),
                                    )?;
                                }
                                DocRun::Shape(Box::new(shape.shape))
                            }
                        });
                    }
                    let occurrence_id = drawing.occurrence_id;
                    let drawing_inline = drawing.inline;
                    host.anchor_occurrence_id = Some(occurrence_id);
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
                    if !drawing_inline {
                        budget.push(&mut paragraph.runs, DocRun::AnchorHost(host))?;
                    }
                    // Group members follow one host, as DOCX wpg members do.
                    for run in runs {
                        budget.push(&mut paragraph.runs, run)?;
                    }
                }
                Token::NoteReference(reference) => {
                    let id = reference.id().to_string();
                    push_note_mark(
                        &mut paragraph,
                        story,
                        formatting,
                        style,
                        table_style,
                        cp,
                        reference.kind(),
                        &id,
                        link,
                        budget,
                    )?;
                }
                Token::NoteNumber(kind) => {
                    let (_, fc, piece) = story
                        .position(cp)
                        .ok_or_else(|| unsupported("Word note number outside piece table"))?;
                    // MS-DOC 2.3.2/2.3.5: the automatic number is a special
                    // character (sprmCFSpec), as its reference is.
                    if !formatting.passive_special_character(style, fc, piece.prm, &story.prcs)? {
                        return Err(unsupported(
                            "Word note number lacks special-character property",
                        ));
                    }
                    push_note_mark(
                        &mut paragraph,
                        story,
                        formatting,
                        style,
                        table_style,
                        cp,
                        kind,
                        "",
                        link,
                        budget,
                    )?;
                }
                Token::NoteMarker => {
                    return Err(unsupported(
                        "Word note character outside a note reference or note",
                    ));
                }
                Token::Linked(_) => {
                    return Err(unsupported("nested Word link token"));
                }
                Token::EvaluatedField(field) if field.ruby.is_some() => {
                    push_ruby(
                        &mut paragraph,
                        story,
                        formatting,
                        style,
                        table_style,
                        field.ruby.as_deref().expect("ruby form"),
                        budget,
                    )?;
                }
                Token::EvaluatedField(field) => {
                    let run = evaluated_field_run(story, formatting, style, table_style, &field)?;
                    budget.charge(
                        std::mem::size_of::<docx_model::FieldRun>()
                            .checked_add(super::payload::field_run(&run)?)
                            .ok_or("OUTPUT_TOO_LARGE")?,
                    )?;
                    budget.push(&mut paragraph.runs, DocRun::Field(Box::new(run)))?;
                }
                Token::FieldBegin(_) | Token::FieldEnd => {
                    // Only the byte converter's header restoration emits these.
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
                // A break character in the DOC text is an authored break,
                // exactly as the converted DOCX reports it.
                origin: Some(docx_model::PageBreakOrigin::Authored),
                parity: None,
                same_paragraph_as_previous,
            },
            Block::ColumnBreak => BodyElement::ColumnBreak,
        };
        budget.push(body, element)?;
    }
    Ok(())
}

fn resolve_table_cell_shading(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    formatting: &mut formatting::Formatting<'_>,
) -> Result<(), String> {
    for (table_id, table_context) in index.tables().iter().enumerate() {
        for (row_index, row_context) in table_context.rows.iter().enumerate() {
            let options = if let (Some(table_style), Some(flags)) =
                (row_context.table_style, row_context.table_style_options)
            {
                let (horizontal, vertical, presence) =
                    formatting.table_style_selector_profile(Some(table_style))?;
                Some(table_style_condition::Options::new(
                    flags, horizontal, vertical, presence,
                )?)
            } else {
                None
            };
            let row = &mut prepared
                .get_mut(row_context.ttp_id)
                .ok_or_else(|| unsupported("Word table shading TTP outside story"))?
                .table_properties
                .row;
            if row.cells.len() != row_context.source_cell_count {
                return Err(unsupported("Word table shading cell count mismatch"));
            }
            for (ordinal, cell) in row.cells.iter_mut().enumerate() {
                let matches = if let Some(options) = options {
                    table_style_condition::select(
                        index,
                        table_id,
                        row_index,
                        table_style_condition::LogicalColumn {
                            ordinal,
                            count: row_context.source_cell_count,
                        },
                        options,
                    )?
                } else {
                    [None; 5]
                };
                let key = formatting.table_formatting_key(
                    row_context.table_style,
                    row_context.table_style_options,
                    matches,
                )?;
                let style = formatting.table_cell_shading(key)?;
                cell.shading = match cell.prepared_shading.take() {
                    None | Some(table::PreparedCellShading::StyleDeferred) => style,
                    Some(table::PreparedCellShading::Explicit(value)) => Some(value),
                    Some(table::PreparedCellShading::Unsupported) => {
                        formatting.unsupported_table_properties = true;
                        None
                    }
                };
            }
        }
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
    link: Option<&super::fields::Link>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let (_, fc, piece) = story
        .position(cp)
        .ok_or_else(|| unsupported("Word control outside piece table"))?;
    if let Some(mut run) = linked_text_run(
        formatting,
        style,
        table_style,
        fc,
        piece.prm,
        &story.prcs,
        link,
    )? {
        let text = crate::doc::character::Properties::direct_run_text(&mut run, text)?;
        budget.text(&mut paragraph.runs, &mut run, &text)?;
    }
    Ok(())
}

/// A text run with the DOCX parser's link facts: `is_link`, the external
/// target and bookmark anchor, and inside a TOC result the paragraph-level
/// color and underline (see `fields::Link`).
fn linked_text_run(
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    table_style: Option<formatting::TableFormattingKey>,
    fc: usize,
    prm: u16,
    prcs: &[&[u8]],
    link: Option<&super::fields::Link>,
) -> Result<Option<docx_model::TextRun>, String> {
    let Some(link) = link else {
        return formatting.direct_text_run(style, table_style, fc, prm, prcs, String::new());
    };
    Ok(formatting
        .direct_link_text_run(style, table_style, fc, prm, prcs, link.in_toc)?
        .map(|mut run| {
            run.is_link = true;
            run.hyperlink = link.href.clone();
            run.hyperlink_anchor = link.anchor.clone();
            run
        }))
}

/// ECMA-376 17.11.6/7/16/17 as parsed by the DOCX parser: a note mark is a
/// `TextRun` with the run's properties, forced superscript, its id as
/// fallback text and `NoteRef { kind, id }`. An empty id is the in-note
/// number, which the renderer resolves to the enclosing note's number.
#[allow(clippy::too_many_arguments)]
fn push_note_mark(
    paragraph: &mut docx_model::DocParagraph,
    story: &Story<'_>,
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    table_style: Option<formatting::TableFormattingKey>,
    cp: usize,
    kind: crate::doc::notes::Kind,
    id: &str,
    link: Option<&super::fields::Link>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let (_, fc, piece) = story
        .position(cp)
        .ok_or_else(|| unsupported("Word note mark outside piece table"))?;
    let mut run = linked_text_run(
        formatting,
        style,
        table_style,
        fc,
        piece.prm,
        &story.prcs,
        link,
    )?
    .ok_or_else(|| unsupported("hidden Word note mark is not supported"))?;
    run.vert_align = Some("super".to_string());
    run.note_ref = Some(docx_model::NoteRef {
        kind: kind.tag().to_string(),
        id: id.to_string(),
    });
    budget.text(&mut paragraph.runs, &mut run, id)
}

/// Project an EQ phonetic guide exactly as the DOCX parser projects the
/// `w:ruby` Word wrote for it (ECMA-376 17.3.3.25): the base text keeps its
/// own character properties and the first base run carries the annotation
/// (guide text, `w:hps`, `w:hpsRaise`, `distributeSpace` alignment, the base
/// size and the guide runs' formatting).
fn push_ruby(
    paragraph: &mut docx_model::DocParagraph,
    story: &Story<'_>,
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    table_style: Option<formatting::TableFormattingKey>,
    form: &super::fields::RubyForm,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    use docx_model::{RubyGuideRunTypographyWire, TypographyValueStatusWire, TypographyValueWire};
    let valid = |raw: String| TypographyValueWire {
        status: TypographyValueStatusWire::Valid,
        raw: Some(raw.clone()),
        value: Some(raw),
    };
    let valid_pt = |value: f64| TypographyValueWire {
        status: TypographyValueStatusWire::Valid,
        raw: Some(value.to_string()),
        value: Some(value),
    };
    let mut guide_runs = Vec::new();
    super::super::visit_text_runs(
        &form.guide,
        form.guide_cp,
        story,
        &mut Some(&mut *formatting),
        |formatting, fc, prm| {
            formatting.direct_text_run(style, table_style, fc, prm, &story.prcs, String::new())
        },
        |part, run| {
            let run = run
                .flatten()
                .ok_or_else(|| unsupported("hidden Word phonetic guide text"))?;
            budget.charge(std::mem::size_of::<RubyGuideRunTypographyWire>() + part.len())?;
            guide_runs.push(RubyGuideRunTypographyWire {
                text: part.to_string(),
                font_family: run.font_family.or(run.font_family_east_asia),
                font_size_pt: Some(run.font_size),
                bold: run.bold,
                italic: run.italic,
                color: run.color,
                language: run.lang_default.map(|value| value.to_lowercase()),
            });
            Ok(())
        },
    )?;
    let mut base_runs = Vec::new();
    super::super::visit_text_runs(
        &form.base,
        form.base_cp,
        story,
        &mut Some(&mut *formatting),
        |formatting, fc, prm| {
            formatting.direct_text_run(style, table_style, fc, prm, &story.prcs, String::new())
        },
        |part, run| {
            if let Some(run) = run.flatten() {
                base_runs.push((part.to_string(), run));
            }
            Ok(())
        },
    )?;
    let base_size = base_runs
        .first()
        .map(|(_, run)| run.font_size)
        .ok_or_else(|| unsupported("hidden Word phonetic guide base text"))?;
    let raise_pt = f64::from(form.raise_pt);
    let annotation = docx_model::RubyAnnotation {
        text: form.guide.clone(),
        font_size_pt: f64::from(form.guide_half_points) / 2.0,
        hps_raise_pt: Some(raise_pt),
        typography: Some(docx_model::RubyTypographyWire {
            align: valid("distributeSpace".to_string()),
            base_font_size_pt: valid_pt(base_size),
            raise_pt: valid_pt(raise_pt),
            language: TypographyValueWire::default(),
            guide_runs,
        }),
    };
    for (index, (part, mut run)) in base_runs.into_iter().enumerate() {
        if index == 0 {
            if let Some(typography) = &mut run.typography_acquisition {
                typography.ruby = annotation.typography.clone();
            }
            run.ruby = Some(annotation.clone());
        }
        budget.text(&mut paragraph.runs, &mut run, &part)?;
    }
    Ok(())
}

/// Project one Word-evaluated field onto the DOCX parser's `FieldRun`
/// (`make_field_run`): the same retained character-property subset, the
/// trimmed instruction and the stored result as fallback text.
fn evaluated_field_run(
    story: &Story<'_>,
    formatting: &mut formatting::Formatting<'_>,
    style: usize,
    table_style: Option<formatting::TableFormattingKey>,
    field: &super::fields::Evaluated,
) -> Result<docx_model::FieldRun, String> {
    let checkbox = field
        .form_data_cp
        .map(|data_cp| {
            let (_, fc, piece) = story
                .position(data_cp)
                .ok_or_else(|| unsupported("Word form data outside piece table"))?;
            let data =
                formatting.direct_binary_data(style, table_style, fc, piece.prm, &story.prcs)?;
            super::fields::checkbox_state(data)
        })
        .transpose()?;
    let mut properties = |cp: usize| -> Result<docx_model::FieldRun, String> {
        let (_, fc, piece) = story
            .position(cp)
            .ok_or_else(|| unsupported("Word field outside piece table"))?;
        let run = formatting
            .direct_text_run(
                style,
                table_style,
                fc,
                piece.prm,
                &story.prcs,
                String::new(),
            )?
            .ok_or_else(|| unsupported("hidden Word evaluated field is not supported"))?;
        Ok(docx_model::FieldRun {
            field_type: field.field_type.to_string(),
            instruction: field.instruction.clone(),
            fallback_text: field.cached_result.clone(),
            bold: run.bold,
            italic: run.italic,
            underline: run.underline,
            strikethrough: run.strikethrough,
            font_size: run.font_size,
            color: run.color,
            font_family: run.font_family,
            font_family_high_ansi: run.font_family_high_ansi,
            font_slots: run.font_slots,
            font_family_east_asia: run.font_family_east_asia,
            font_hint: run.font_hint,
            rtl: run.rtl,
            cs: run.cs,
            font_family_cs: run.font_family_cs,
            font_size_cs: run.font_size_cs,
            bold_cs: run.bold_cs,
            italic_cs: run.italic_cs,
            lang_default: run.lang_default,
            lang_bidi: run.lang_bidi,
            lang_east_asia: run.lang_east_asia,
            background: run.background,
            vert_align: run.vert_align,
            all_caps: run.all_caps,
            small_caps: run.small_caps,
            double_strikethrough: run.double_strikethrough,
            highlight: run.highlight,
            emphasis_mark: run.emphasis_mark,
            typography_acquisition: run.typography_acquisition,
        })
    };
    let mut run = properties(field.format_cp)?;
    if let Some((checked, size)) = checkbox {
        // ECMA-376 17.16.17 checkBox: the DOCX parser shows a ballot box and
        // applies an explicit size to both font-size slots.
        run.fallback_text = if checked { "\u{2612}" } else { "\u{2610}" }.to_string();
        if let Some(size) = size {
            run.font_size = size;
            run.font_size_cs = Some(size);
        }
    }
    if let Some(result_cp) = field.agreeing_result_cp {
        // Without MERGEFORMAT/CHARFORMAT, ECMA-376 17.16.4.3.3 leaves the
        // formatting of a regenerated result to the application. The DOCX
        // parser uses the first instruction run; accept that only when the
        // stored result agrees, including every typography acquisition fact.
        let stored = properties(result_cp)?;
        let same = serde_json::to_value(&run).map_err(|error| error.to_string())?
            == serde_json::to_value(&stored).map_err(|error| error.to_string())?;
        if !same {
            return Err(unsupported(
                "Word evaluated field result formatting differs from its instruction",
            ));
        }
    }
    Ok(run)
}

/// Project one textbox's text (MS-DOC 2.3.6, 2.9.106) through the ordinary
/// story projection into the DOCX text box block stream (ECMA-376 17.3.4.7
/// `w:txbxContent`). The textbox range ends with its final paragraph mark.
#[allow(clippy::too_many_arguments)]
fn textbox_content(
    textboxes: &floating::textbox::Textboxes<'_>,
    index: usize,
    spid: u32,
    formatting: &mut formatting::Formatting<'_>,
    pictures: &mut pictures::Store<'_>,
    budget: &mut ModelBudget,
    table_sequence: &mut usize,
) -> Result<Vec<docx_model::TextBoxBlockWire>, String> {
    let (text, base_cp) = textboxes.text(index, spid)?;
    // Each occurrence re-tokenizes its range: charge that scratch work.
    budget.charge(text.len())?;
    let mut paragraphs = super::super::tokenize_with_fields(
        text,
        &mut super::super::Fields::default(),
        base_cp,
        true,
    );
    // Fields follow the textbox document's own Plcfld (MS-DOC 2.8.25).
    textboxes.fields.apply(base_cp, &mut paragraphs)?;
    let mut body = Vec::new();
    let mut numbering = numbering::direct::Store::default();
    numbering.begin_story()?;
    project(
        &textboxes.story,
        paragraphs,
        formatting,
        &mut numbering,
        pictures,
        None,
        budget,
        &mut body,
        None,
        table_sequence,
    )?;
    // List counters shared between textboxes and the main story, and breaks
    // inside a textbox, have no Office control yet: keep them fail-closed.
    // A bullet level has no counter, so its marker does not depend on that
    // sharing; bullets are projected (each textbox uses its own numbering
    // store, which only matters for counted levels).
    let numbered = || unsupported("direct DOC model does not yet number textbox paragraphs");
    let counted = |paragraph: &docx_model::DocParagraph| {
        paragraph
            .numbering
            .as_ref()
            .is_some_and(|numbering| numbering.format != "bullet")
    };
    let mut tables = Vec::new();
    for block in &body {
        match block {
            BodyElement::Paragraph(paragraph) if counted(paragraph) => return Err(numbered()),
            BodyElement::Paragraph(_) => {}
            BodyElement::Table(table) => tables.push(table.as_ref()),
            _ => {
                return Err(unsupported(
                    "direct DOC model does not support breaks inside textboxes",
                ))
            }
        }
    }
    while let Some(table) = tables.pop() {
        for element in table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.content)
        {
            match element {
                docx_model::CellElement::Paragraph(paragraph) if counted(paragraph) => {
                    return Err(numbered())
                }
                docx_model::CellElement::Paragraph(_) => {}
                docx_model::CellElement::Table(table) => tables.push(table.as_ref()),
            }
        }
    }
    budget.charge(
        body.len()
            .checked_mul(std::mem::size_of::<docx_model::TextBoxBlockWire>())
            .ok_or("OUTPUT_TOO_LARGE")?,
    )?;
    Ok(body
        .into_iter()
        .map(docx_model::TextBoxBlockWire::Body)
        .collect())
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

    #[derive(Clone, Copy, Default)]
    struct StyleFixture {
        first_row_color: bool,
        top_left_color: bool,
        first_column_color: bool,
        first_row_size: bool,
        first_row_alignment: bool,
        combined: bool,
        modern_raw: bool,
        inherited_conditions: bool,
        inherited_margins: bool,
        table_borders: bool,
        default_table_style: bool,
        default_table_style_indent: bool,
        conditional_borders: u16,
        conditional_borders_second: u16,
        conditional_border_nil: bool,
        conditional_border_malformed: bool,
        repeated_conditional_borders: u8,
        inherited_conditional_borders: bool,
        inherited_partial_conditional_borders: bool,
        inherited_first_column_borders: bool,
        distinct_multi_condition_order: u8,
        shading: ShadingFixture,
        margins: MarginFixture,
    }

    #[derive(Clone, Copy, Default, PartialEq, Eq)]
    enum ShadingFixture {
        #[default]
        None,
        Unconditional,
        Conditional,
    }

    #[derive(Clone, Copy, Default, PartialEq, Eq)]
    enum MarginFixture {
        #[default]
        None,
        D63e,
        ConditionalD63e,
        D634,
    }

    struct ProjectedTable {
        markers: Vec<String>,
        row_cell_counts: Vec<usize>,
        col_spans: Vec<u32>,
        colors: Vec<String>,
        sizes: Vec<f64>,
        ascii_fonts: Vec<Option<String>>,
        high_ansi_fonts: Vec<Option<String>>,
        alignments: Vec<String>,
        backgrounds: Vec<Option<String>>,
        margins: Vec<[f64; 4]>,
        margin_wires: Vec<[String; 4]>,
        borders: Vec<[Option<(String, f64)>; 4]>,
        border_styles: Vec<[Option<String>; 4]>,
        text_directions: Vec<Option<String>>,
        diagonals: Vec<[Option<(String, f64, Option<String>)>; 2]>,
        hide_marks: Vec<bool>,
        unsupported_table: bool,
        unsupported_character: bool,
        unsupported_paragraph: bool,
    }

    fn cnf(code: u16, condition: u16, properties: &[u8]) -> Vec<u8> {
        let mut value = Vec::from(code.to_le_bytes());
        value.push(u8::try_from(2 + properties.len()).unwrap());
        value.extend(condition.to_le_bytes());
        value.extend(properties);
        value
    }

    fn cell_shading(color: [u8; 3]) -> Vec<u8> {
        vec![10, 0, 0, 0, 255, color[0], color[1], color[2], 0, 0, 0]
    }

    fn cell_margin(code: u16, sides: u8, unit: u8, width: u16) -> Vec<u8> {
        cell_margin_range(code, 0, 1, sides, unit, width)
    }

    fn cell_margin_range(
        code: u16,
        first: u8,
        limit: u8,
        sides: u8,
        unit: u8,
        width: u16,
    ) -> Vec<u8> {
        let [lo, hi] = width.to_le_bytes();
        sprm(code, &[6, first, limit, sides, unit, lo, hi])
    }

    fn border_bytes(color: [u8; 3], width: u8) -> [u8; 8] {
        [color[0], color[1], color[2], 0, width, 1, 0, 0]
    }

    fn table_borders(color: [u8; 3], width: u8) -> Vec<u8> {
        let mut operand = vec![48];
        for _ in 0..6 {
            operand.extend(border_bytes(color, width));
        }
        sprm(0xd613, &operand)
    }

    fn cell_borders(first: u8, limit: u8, color: [u8; 3], width: u8) -> Vec<u8> {
        let mut operand = vec![11, first, limit, 0x0f];
        operand.extend(border_bytes(color, width));
        sprm(0xd62f, &operand)
    }

    fn conditional_borders(condition: u16) -> Vec<u8> {
        let mut properties = Vec::new();
        for (code, color) in [
            (0xd47f, [0xff, 0, 0]),
            (0xd680, [0, 0, 0xff]),
            (0xd681, [0, 0x80, 0]),
            (0xd682, [0xff, 0xff, 0]),
            (0xd683, [0xff, 0, 0xff]),
            (0xd684, [0, 0xff, 0xff]),
        ] {
            let mut operand = vec![8];
            operand.extend(border_bytes(color, 32));
            properties.extend(sprm(code, &operand));
        }
        cnf(0xd66a, condition, &properties)
    }

    fn conditional_border_side(condition: u16, code: u16, color: [u8; 3], width: u8) -> Vec<u8> {
        let mut operand = vec![8];
        operand.extend(border_bytes(color, width));
        cnf(0xd66a, condition, &sprm(code, &operand))
    }

    fn conditional_border_nil(condition: u16) -> Vec<u8> {
        let mut operand = vec![8];
        operand.extend([0xff; 8]);
        cnf(0xd66a, condition, &sprm(0xd47f, &operand))
    }

    fn malformed_conditional_border(condition: u16) -> Vec<u8> {
        cnf(0xd66a, condition, &sprm(0xd47f, &[7, 0, 0, 0, 0, 0, 0, 0]))
    }

    fn append_table_style(bytes: &mut Vec<u8>, base: u16, sets: [&[u8]; 3]) {
        let mut style = vec![0; 14];
        style[2..4].copy_from_slice(&((base << 4) | 3).to_le_bytes());
        style[4..6].copy_from_slice(&3u16.to_le_bytes());
        for set in sets {
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
    }

    fn table_style_sheet(fixture: StyleFixture) -> Vec<u8> {
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

        if fixture.default_table_style {
            for _ in 1..11 {
                bytes.extend(0u16.to_le_bytes());
            }
            let mut width_before = sprm(0xf617, &[3, 0, 0]);
            if fixture.default_table_style_indent {
                // Word's default table style also carries a zero dxa
                // sprmTWidthIndent.
                width_before.extend(sprm(0xf661, &[3, 0, 0]));
            }
            append_table_style(&mut bytes, 0xfff, [&width_before, &[], &[]]);
            for _ in 12..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        if fixture.inherited_margins {
            let base = cell_margin(0xd634, 0x02, 3, 720);
            append_table_style(&mut bytes, 0xfff, [&base, &[], &[]]);
            let child = cell_margin(0xd634, 0x02, 0, 0);
            append_table_style(&mut bytes, 1, [&child, &[], &[]]);
            append_table_style(&mut bytes, 1, [&[], &[], &[]]);
            for _ in 4..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        if fixture.inherited_conditions {
            let mut base_paragraph = vec![1, 0, 0x61, 0x24, 0];
            base_paragraph.extend(cnf(0xc666, 1, &[0x61, 0x24, 1]));
            let mut base_character = vec![0x42, 0x2a, 1];
            base_character.extend(cnf(
                0xca85,
                1,
                &[0x70, 0x68, 0xff, 0, 0, 0, 0x43, 0x4a, 28, 0],
            ));
            append_table_style(&mut bytes, 0xfff, [&[], &base_paragraph, &base_character]);

            let child_paragraph = [2, 0];
            let child_character = cnf(0xca85, 1, &[0x70, 0x68, 0, 0, 0xff, 0]);
            append_table_style(&mut bytes, 1, [&[], &child_paragraph, &child_character]);
            for _ in 3..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        if fixture.inherited_partial_conditional_borders {
            let child = conditional_border_side(1, 0xd47f, [0, 0, 0], 40)
                .into_iter()
                .chain(conditional_border_side(1, 0xd684, [0xff, 0x80, 0], 32))
                .collect::<Vec<_>>();
            append_table_style(&mut bytes, 2, [&child, &[], &[]]);
            let parent = conditional_borders(1);
            append_table_style(&mut bytes, 0xfff, [&parent, &[], &[]]);
            for _ in 3..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        if fixture.inherited_first_column_borders {
            let child = conditional_border_side(4, 0xd681, [0, 0, 0], 40);
            append_table_style(&mut bytes, 2, [&child, &[], &[]]);
            let parent = conditional_borders(4);
            append_table_style(&mut bytes, 0xfff, [&parent, &[], &[]]);
            for _ in 3..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        if fixture.inherited_conditional_borders {
            append_table_style(&mut bytes, 2, [&[], &[], &[]]);
            let parent = conditional_borders(fixture.conditional_borders);
            append_table_style(&mut bytes, 0xfff, [&parent, &[], &[]]);
            for _ in 3..15 {
                bytes.extend(0u16.to_le_bytes());
            }
            return bytes;
        }

        let mut style = vec![0; 14];
        style[2..4].copy_from_slice(&0xfff3u16.to_le_bytes());
        style[4..6].copy_from_slice(&3u16.to_le_bytes());
        let edge_character = if fixture.top_left_color {
            &[0x85, 0xca, 8, 0x00, 0x02, 0x70, 0x68, 0, 0x80, 0, 0][..]
        } else if fixture.first_row_color {
            &[0x85, 0xca, 8, 1, 0, 0x70, 0x68, 0, 0x80, 0, 0][..]
        } else if fixture.first_column_color {
            &[0x85, 0xca, 8, 4, 0, 0x70, 0x68, 0, 0x80, 0, 0][..]
        } else if fixture.first_row_size {
            &[0x85, 0xca, 6, 1, 0, 0x43, 0x4a, 28, 0][..]
        } else {
            // An empty first-row CCnf has no supported condition presence.
            &[0x85, 0xca, 2, 1, 0][..]
        };
        let mut character = vec![
            0x42, 0x2a, 1, // unconditional black
            0x85, 0xca, 5, 0x40, 0, 0x42, 0x2a, 6, // odd row red
            0x85, 0xca, 5, 0x80, 0, 0x42, 0x2a, 2, // even row blue
        ];
        character.extend(edge_character);
        let paragraph = if fixture.first_row_alignment {
            &[
                1, 0, // embedded table style istd
                0x61, 0x24, 0, // unconditional logical left
                0x66, 0xc6, 5, 1, 0, 0x61, 0x24, 1, // first-row logical center
            ][..]
        } else {
            &[1, 0][..]
        };
        let mut combined_paragraph = vec![1, 0, 0x61, 0x24, 0];
        let mut combined_character = vec![
            0x42, 0x2a, 1, // black
            0x43, 0x4a, 20, 0, // 10 pt
            0x4f, 0x4a, 0, 0, // ASCII font 0
            0x51, 0x4a, 1, 0, // high ANSI font 1
        ];
        if fixture.combined {
            for (condition, alignment) in [(1, 1), (4, 2), (0x200, 0)] {
                combined_paragraph.extend(cnf(0xc666, condition, &[0x61, 0x24, alignment]));
            }
            for (condition, properties) in [
                (1, &[0x42, 0x2a, 6, 0x43, 0x4a, 28, 0][..]),
                (4, &[0x42, 0x2a, 2, 0x43, 0x4a, 24, 0][..]),
                (0x200, &[0x70, 0x68, 0, 0x80, 0, 0, 0x43, 0x4a, 32, 0][..]),
            ] {
                combined_character.extend(cnf(0xca85, condition, properties));
            }
        }
        let (mut tapx, paragraph, character) = if fixture.combined {
            (
                Vec::new(),
                combined_paragraph.as_slice(),
                combined_character.as_slice(),
            )
        } else {
            (vec![0x88, 0x34, 1], paragraph, character.as_slice())
        };
        match fixture.margins {
            MarginFixture::None => {}
            MarginFixture::D63e => tapx.extend(cell_margin(0xd63e, 0x02, 3, 360)),
            MarginFixture::ConditionalD63e => {
                tapx.extend(cell_margin(0xd63e, 0x0f, 3, 72));
                tapx.extend(cnf(
                    0xd66a,
                    table_style_condition::FIRST_ROW,
                    &cell_margin(0xd63e, 0x0f, 3, 288),
                ));
            }
            MarginFixture::D634 => tapx.extend(cell_margin(0xd634, 0x02, 3, 360)),
        }
        if fixture.table_borders {
            tapx.extend(table_borders([0xff, 0, 0], 8));
        }
        if fixture.conditional_borders != 0 {
            tapx.extend(conditional_borders(fixture.conditional_borders));
        }
        if fixture.conditional_borders_second != 0 {
            tapx.extend(conditional_borders(fixture.conditional_borders_second));
        }
        if fixture.conditional_border_nil {
            tapx.extend(conditional_border_nil(fixture.conditional_borders));
        }
        if fixture.conditional_border_malformed {
            tapx.extend(malformed_conditional_border(fixture.conditional_borders));
        }
        match fixture.repeated_conditional_borders {
            1 => {
                tapx.extend(conditional_border_side(1, 0xd47f, [0xff, 0, 0], 24));
                tapx.extend(conditional_border_side(1, 0xd680, [0, 0, 0xff], 16));
            }
            2 => {
                tapx.extend(conditional_border_side(1, 0xd47f, [0xff, 0, 0], 24));
                tapx.extend(conditional_border_side(1, 0xd47f, [0, 0, 0xff], 16));
            }
            _ => {}
        }
        let row_left = conditional_border_side(1, 0xd681, [0xff, 0, 0xff], 32);
        let column_left = conditional_border_side(4, 0xd681, [0, 0x80, 0], 24);
        match fixture.distinct_multi_condition_order {
            1 => {
                tapx.extend(row_left);
                tapx.extend(column_left);
            }
            2 => {
                tapx.extend(column_left);
                tapx.extend(row_left);
            }
            _ => {}
        }
        match fixture.shading {
            ShadingFixture::None => {}
            ShadingFixture::Unconditional => {
                tapx.extend(sprm(0xd687, &cell_shading([0x00, 0x80, 0x00])));
            }
            ShadingFixture::Conditional => {
                tapx.extend(sprm(0xd687, &cell_shading([0x00, 0x00, 0x00])));
                for (condition, color) in [
                    (1, [0xff, 0x00, 0x00]),
                    (4, [0x00, 0x00, 0xff]),
                    (0x200, [0x00, 0x80, 0x00]),
                ] {
                    tapx.extend(cnf(0xd66a, condition, &sprm(0xd687, &cell_shading(color))));
                }
            }
        }
        for set in [tapx.as_slice(), paragraph, character] {
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

    fn with_table_style(source: &[u8], fixture: StyleFixture) -> Vec<u8> {
        let cfb = CompoundFile::open(source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let stylesheet = table_style_sheet(fixture);
        if fixture.modern_raw || fixture.shading != ShadingFixture::None {
            // Activate the post-Word-2000 Raw shading rules in the native-only
            // acquisition path without changing the fixture's stream layout.
            word[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        }
        word[0xa2..0xa6].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0xa6..0xaa].copy_from_slice(&(stylesheet.len() as u32).to_le_bytes());
        table.extend(stylesheet);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    fn with_fonts(source: &[u8], names: &[&str]) -> Vec<u8> {
        let cfb = CompoundFile::open(source).unwrap();
        let mut word = cfb.stream("WordDocument").unwrap();
        let mut table = cfb.stream("0Table").unwrap();
        let mut font_table = vec![names.len() as u8, 0, 0, 0];
        for name in names {
            let mut font = vec![0; 39];
            for unit in name.encode_utf16().chain(std::iter::once(0)) {
                font.extend(unit.to_le_bytes());
            }
            font_table.push(font.len() as u8);
            font_table.extend(font);
        }
        word[0x112..0x116].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0x116..0x11a].copy_from_slice(&(font_table.len() as u32).to_le_bytes());
        table.extend(font_table);
        build_cfb(&[("WordDocument", word), ("0Table", table)])
    }

    fn cell() -> Vec<u8> {
        sprm(0x2416, &[1])
    }

    fn row(options: u16) -> Vec<u8> {
        row_cells(options, 1)
    }

    fn row_cells(options: u16, count: u8) -> Vec<u8> {
        row_cells_with_style(options, count, 1)
    }

    fn row_cells_with_style(options: u16, count: u8, style: u16) -> Vec<u8> {
        row_cells_with_width(options, count, style, 1000)
    }

    fn row_cells_with_width(options: u16, count: u8, style: u16, width: u16) -> Vec<u8> {
        let [width_lo, width_hi] = width.to_le_bytes();
        [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, count, width_lo, width_hi]),
            sprm(0x563a, &style.to_le_bytes()),
            sprm(0x740a, &[0, 0, options as u8, (options >> 8) as u8]),
            sprm(0x2416, &[1]),
        ]
        .concat()
    }

    fn row_cells_with_horizontal_merge(options: u16, first: u8, limit: u8) -> Vec<u8> {
        let mut row = row_cells(options, 3);
        row.extend(sprm(0x5624, &[first, limit]));
        if row.len() % 2 == 0 {
            row.extend(cell());
        }
        row
    }

    fn row_cells_with_extra(options: u16, count: u8, extra: &[u8]) -> Vec<u8> {
        let mut row = row_cells(options, count);
        row.extend(extra);
        if row.len() % 2 == 0 {
            row.extend(cell());
        }
        row
    }

    fn row_cells_with_raw(options: u16, count: u8, raw: &[u8]) -> Vec<u8> {
        [
            row_cells(options, count),
            sprm(0xd670, raw),
            sprm(0x2416, &[1]),
        ]
        .concat()
    }

    fn row_cells_with_compatibility_and_raw(
        options: u16,
        count: u8,
        compatibility: &[u8],
        raw: &[u8],
    ) -> Vec<u8> {
        let mut row = row_cells(options, count);
        row.extend(sprm(0xd612, compatibility));
        row.extend(sprm(0xd670, raw));
        if row.len() % 2 == 0 {
            row.extend(cell());
        }
        row
    }

    fn offset_merged_row_with_overrides(
        options: u16,
        left: i16,
        merge: [u8; 2],
        margin: &[u8],
        raw: &[u8],
    ) -> Vec<u8> {
        let mut row = row_cells(options, 3);
        row.extend(sprm(0x9601, &left.to_le_bytes()));
        row.extend(sprm(0x5624, &merge));
        row.extend(margin);
        row.extend(sprm(0xd670, raw));
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn row_cells_with_margins(style: u16, margins: &[Vec<u8>]) -> Vec<u8> {
        let mut row = row_cells_with_style(0, 1, style);
        for margin in margins {
            row.extend(margin);
        }
        // The synthetic PAPX builder stores an odd number of bytes including
        // its two-byte style prefix. Keep the direct margin PRLs followed by a
        // harmless in-cell assertion so that framing invariant still holds.
        row.extend(sprm(0x2416, &[1]));
        row
    }

    fn row_cells_with_options_and_margins(options: u16, count: u8, margins: &[Vec<u8>]) -> Vec<u8> {
        let mut row = row_cells(options, count);
        for margin in margins {
            row.extend(margin);
        }
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn row_cells_with_border(style: u16, border: &[u8], before_tistd: bool) -> Vec<u8> {
        let mut row = [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, 1, 0xe8, 3]),
        ]
        .concat();
        if before_tistd {
            row.extend(border);
        }
        row.extend(sprm(0x563a, &style.to_le_bytes()));
        row.extend(sprm(0x740a, &[0, 0, 0, 0]));
        if !before_tistd {
            row.extend(border);
        }
        row.extend(sprm(0x2416, &[1]));
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn row_cells_with_tc80(style: u16) -> Vec<u8> {
        let mut definition = vec![26, 0, 1, 0, 0, 0xe8, 3, 0, 0, 0, 0];
        for _ in 0..4 {
            definition.extend([16, 1, 2, 0]); // blue Brc80, 2 pt
        }
        let mut row = [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0xd608, &definition),
            sprm(0x563a, &style.to_le_bytes()),
            sprm(0x740a, &[0, 0, 0, 0]),
            sprm(0x2416, &[1]),
        ]
        .concat();
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn unstyled_row_with_raw(raw: &[u8]) -> Vec<u8> {
        [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, 1, 0xe8, 3]),
            sprm(0xd670, raw),
        ]
        .concat()
    }

    fn unstyled_row_with_border(border: &[u8]) -> Vec<u8> {
        let mut row = [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, 1, 0xe8, 3]),
            border.to_vec(),
            sprm(0x2416, &[1]),
        ]
        .concat();
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn unstyled_row_with_tc80() -> Vec<u8> {
        let mut definition = vec![26, 0, 1, 0, 0, 0xe8, 3, 0, 0, 0, 0];
        for _ in 0..4 {
            definition.extend([16, 1, 2, 0]); // blue Brc80, 2 pt
        }
        let mut row = [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0xd608, &definition),
            sprm(0x2416, &[1]),
        ]
        .concat();
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn old_cell_borders(first: u8, limit: u8, color: u8, width: u8) -> Vec<u8> {
        sprm(0xd620, &[7, first, limit, 0x0f, width, 1, color, 0])
    }

    fn row_cells_with_compatibility_shading(options: u16, count: u8, shading: &[u8]) -> Vec<u8> {
        [
            row_cells(options, count),
            sprm(0xd612, shading),
            sprm(0x2416, &[1]),
        ]
        .concat()
    }

    fn try_project_table(
        text: &str,
        runs: &[(usize, usize, Vec<u8>)],
        fixture: StyleFixture,
    ) -> Result<ProjectedTable, String> {
        let source = super::super::tests::source_with_typography(
            text,
            &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
            None,
            None,
            None,
            None,
        );
        let source = if fixture.combined {
            with_fonts(&source, &["Times New Roman", "Courier New"])
        } else {
            source
        };
        let source = with_table_style(&source, fixture);
        let source = with_papx(&source, runs);
        let cfb = CompoundFile::open(&source).unwrap();
        with_acquired_doc(&cfb, true, |mut facts| {
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
                Some((&mut facts.floating, floating::Part::Main)),
                &mut budget,
                &mut body,
                None,
                &mut table_sequence,
            )?;

            let mut colors = Vec::new();
            let mut markers = Vec::new();
            let mut row_cell_counts = Vec::new();
            let mut col_spans = Vec::new();
            let mut sizes = Vec::new();
            let mut ascii_fonts = Vec::new();
            let mut high_ansi_fonts = Vec::new();
            let mut alignments = Vec::new();
            let mut backgrounds = Vec::new();
            let mut margins = Vec::new();
            let mut margin_wires = Vec::new();
            let mut borders = Vec::new();
            let mut border_styles = Vec::new();
            let mut text_directions = Vec::new();
            let mut diagonals = Vec::new();
            let mut hide_marks = Vec::new();
            for element in &body {
                let BodyElement::Table(table) = element else {
                    continue;
                };
                for row in &table.rows {
                    row_cell_counts.push(row.cells.len());
                    for cell in &row.cells {
                        col_spans.push(cell.col_span);
                        text_directions.push(cell.text_direction.clone());
                        diagonals.push([&cell.borders.tl2br, &cell.borders.tr2bl].map(|b| {
                            b.as_ref()
                                .map(|b| (b.style.clone(), b.width, b.color.clone()))
                        }));
                        hide_marks.push(cell.hide_mark);
                        backgrounds.push(cell.background.clone());
                        margins.push([
                            cell.margin_top.unwrap(),
                            cell.margin_left.unwrap(),
                            cell.margin_bottom.unwrap(),
                            cell.margin_right.unwrap(),
                        ]);
                        let wire = cell.table_cell_layout.margins.as_ref().unwrap();
                        margin_wires.push([
                            wire.top.as_ref().unwrap().value.clone().unwrap(),
                            wire.left.as_ref().unwrap().value.clone().unwrap(),
                            wire.bottom.as_ref().unwrap().value.clone().unwrap(),
                            wire.right.as_ref().unwrap().value.clone().unwrap(),
                        ]);
                        borders.push(std::array::from_fn(|side| {
                            let border = match side {
                                0 => &cell.borders.top,
                                1 => &cell.borders.left,
                                2 => &cell.borders.bottom,
                                3 => &cell.borders.right,
                                _ => unreachable!(),
                            };
                            border
                                .as_ref()
                                .map(|value| (value.color.clone().unwrap_or_default(), value.width))
                        }));
                        border_styles.push(std::array::from_fn(|side| {
                            match side {
                                0 => &cell.borders.top,
                                1 => &cell.borders.left,
                                2 => &cell.borders.bottom,
                                3 => &cell.borders.right,
                                _ => unreachable!(),
                            }
                            .as_ref()
                            .map(|value| value.style.clone())
                        }));
                        let CellElement::Paragraph(paragraph) = &cell.content[0] else {
                            panic!("paragraph")
                        };
                        let DocRun::Text(run) = &paragraph.runs[0] else {
                            panic!("text")
                        };
                        markers.push(run.text.clone());
                        colors.push(run.color.clone().unwrap_or_default());
                        sizes.push(run.font_size);
                        ascii_fonts.push(run.font_family.clone());
                        high_ansi_fonts.push(run.font_family_high_ansi.clone());
                        alignments.push(paragraph.alignment.clone());
                    }
                }
            }
            assert!(!margins.is_empty(), "table");
            Ok(ProjectedTable {
                markers,
                row_cell_counts,
                col_spans,
                colors,
                sizes,
                ascii_fonts,
                high_ansi_fonts,
                alignments,
                backgrounds,
                margins,
                margin_wires,
                borders,
                border_styles,
                text_directions,
                diagonals,
                hide_marks,
                unsupported_table: facts.formatting.unsupported_table_properties,
                unsupported_character: facts.formatting.unsupported_character_properties,
                unsupported_paragraph: facts.formatting.unsupported_paragraph_properties,
            })
        })
    }

    fn project_table(
        text: &str,
        runs: &[(usize, usize, Vec<u8>)],
        fixture: StyleFixture,
    ) -> ProjectedTable {
        try_project_table(text, runs, fixture).unwrap()
    }

    fn projected_table(fixture: StyleFixture) -> ProjectedTable {
        let text = "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}d\u{7}\u{7}\r";
        project_table(
            text,
            &[
                (0, 2, cell()),
                (2, 3, row(1 << 5)),
                (3, 5, cell()),
                (5, 6, row(1 << 5)),
                (6, 8, cell()),
                (8, 9, row(1 << 5)),
                (9, 11, cell()),
                (11, 12, row(1 << 9)),
                (12, 13, Vec::new()),
            ],
            fixture,
        )
    }

    #[test]
    fn native_story_applies_style_then_post_tistd_table_and_cell_borders() {
        let green_table = table_borders([0, 0x80, 0], 24);
        let blue_cell = cell_borders(0, 1, [0, 0, 0xff], 16);
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells_with_border(1, &green_table, true)),
                (3, 5, cell()),
                (5, 6, row_cells_with_border(1, &green_table, false)),
                (6, 8, cell()),
                (8, 9, row_cells_with_border(1, &blue_cell, false)),
                (9, 10, Vec::new()),
            ],
            StyleFixture {
                table_borders: true,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.borders,
            [
                std::array::from_fn(|_| Some(("ff0000".into(), 1.0))),
                std::array::from_fn(|_| Some(("008000".into(), 3.0))),
                std::array::from_fn(|_| Some(("0000ff".into(), 2.0))),
            ]
        );
        assert!(!projected.unsupported_table);
    }

    fn conditional_border_grid_with_rows(
        fixture: StyleFixture,
        rows: [Vec<u8>; 3],
    ) -> ProjectedTable {
        try_conditional_border_grid_with_rows(fixture, rows).unwrap()
    }

    fn try_conditional_border_grid_with_rows(
        fixture: StyleFixture,
        rows: [Vec<u8>; 3],
    ) -> Result<ProjectedTable, String> {
        try_project_table(
            "a\u{7}b\u{7}c\u{7}\u{7}d\u{7}e\u{7}f\u{7}\u{7}g\u{7}h\u{7}i\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 7, rows[0].clone()),
                (7, 9, cell()),
                (9, 11, cell()),
                (11, 13, cell()),
                (13, 14, rows[1].clone()),
                (14, 16, cell()),
                (16, 18, cell()),
                (18, 20, cell()),
                (20, 21, rows[2].clone()),
                (21, 22, Vec::new()),
            ],
            fixture,
        )
    }

    fn conditional_border_grid(condition: u16, options: u16) -> ProjectedTable {
        conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: condition,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(options, 3)),
        )
    }

    #[test]
    fn native_story_maps_first_row_conditional_borders_to_region_edges() {
        let projected = conditional_border_grid(table_style_condition::FIRST_ROW, 1 << 5);
        assert!(
            !projected.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                none.clone(),
                none.clone(),
                none.clone(),
                none,
            ]
        );
    }

    #[test]
    fn horizontal_merge_projects_conditional_borders_to_source_span_edges() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_ROW,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells_with_horizontal_merge(1 << 5, 0, 2)),
        );
        assert!(
            projected.unsupported_table,
            "horizontal merge geometry remains gated independently of border ownership"
        );
        assert_eq!(projected.row_cell_counts, [2, 2, 2]);
        assert_eq!(projected.col_spans, [2, 1, 2, 1, 2, 1]);
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                none.clone(),
                none,
            ]
        );
    }

    #[test]
    fn first_row_right_merge_transfers_the_source_end_right_border() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_ROW,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells_with_horizontal_merge(1 << 5, 1, 3)),
        );
        assert!(
            projected.unsupported_table,
            "horizontal merge geometry remains gated independently of border ownership"
        );
        assert_eq!(projected.row_cell_counts, [2, 2, 2]);
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                none,
                [None, None, None, None],
            ]
        );
    }

    #[test]
    fn merged_tables_keep_unverified_conditions_gated() {
        for (condition, option) in [
            (table_style_condition::LAST_ROW, 1 << 6),
            (table_style_condition::FIRST_COLUMN, 1 << 7),
            (table_style_condition::LAST_COLUMN, 1 << 8),
        ] {
            let projected = conditional_border_grid_with_rows(
                StyleFixture {
                    conditional_borders: condition,
                    ..StyleFixture::default()
                },
                std::array::from_fn(|_| row_cells_with_horizontal_merge(option, 0, 2)),
            );
            assert!(projected.unsupported_table);
            assert!(projected
                .borders
                .iter()
                .all(|sides| sides.iter().all(Option::is_none)));
        }
    }

    #[test]
    fn horizontal_merge_with_direct_start_or_end_border_remains_gated() {
        for (first, limit) in [(0, 1), (1, 2)] {
            let mut row = row_cells_with_horizontal_merge(1 << 5, 0, 2);
            row.extend(cell_borders(first, limit, [0, 0, 0xff], 16));
            if row.len() % 2 == 0 {
                row.extend(cell());
            }
            let projected = conditional_border_grid_with_rows(
                StyleFixture {
                    conditional_borders: table_style_condition::FIRST_ROW,
                    ..StyleFixture::default()
                },
                std::array::from_fn(|_| row.clone()),
            );
            assert!(projected.unsupported_table);
            assert_ne!(projected.borders[0][0], Some(("ff0000".into(), 4.0)));
        }
    }

    #[test]
    fn native_story_maps_first_column_conditional_borders_to_region_edges() {
        let projected = conditional_border_grid(table_style_condition::FIRST_COLUMN, 1 << 7);
        assert!(
            !projected.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none,
            ]
        );
    }

    #[test]
    fn native_story_maps_last_row_conditional_borders_to_region_edges() {
        let projected = conditional_border_grid(table_style_condition::LAST_ROW, 1 << 6);
        assert!(
            !projected.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                none.clone(),
                none.clone(),
                none.clone(),
                none.clone(),
                none.clone(),
                none,
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
            ]
        );
    }

    #[test]
    fn native_story_maps_last_column_conditional_borders_to_region_edges() {
        let projected = conditional_border_grid(table_style_condition::LAST_COLUMN, 1 << 8);
        assert!(
            !projected.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        let none = [None, None, None, None];
        assert_eq!(
            projected.borders,
            [
                none.clone(),
                none.clone(),
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none,
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
            ]
        );
    }

    #[test]
    fn disabled_conditional_borders_do_not_change_cells() {
        for condition in [
            table_style_condition::FIRST_ROW,
            table_style_condition::LAST_ROW,
            table_style_condition::FIRST_COLUMN,
            table_style_condition::LAST_COLUMN,
        ] {
            let projected = conditional_border_grid(condition, 0);
            assert!(projected
                .borders
                .iter()
                .all(|sides| sides.iter().all(Option::is_none)));
        }
    }

    #[test]
    fn singleton_conditional_border_regions_have_no_inside_edges() {
        for (condition, options) in [
            (table_style_condition::FIRST_ROW, 1 << 5),
            (table_style_condition::LAST_ROW, 1 << 6),
            (table_style_condition::FIRST_COLUMN, 1 << 7),
            (table_style_condition::LAST_COLUMN, 1 << 8),
        ] {
            let projected = project_table(
                "a\u{7}\u{7}\r",
                &[
                    (0, 2, cell()),
                    (2, 3, row_cells(options, 1)),
                    (3, 4, Vec::new()),
                ],
                StyleFixture {
                    conditional_borders: condition,
                    ..StyleFixture::default()
                },
            );
            assert_eq!(
                projected.borders,
                [[
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ]]
            );
        }
    }

    #[test]
    fn competing_cross_family_edge_conditions_gate_singleton_borders() {
        for (fixture, options) in [
            (
                StyleFixture {
                    first_row_color: true,
                    conditional_borders: table_style_condition::LAST_ROW,
                    ..StyleFixture::default()
                },
                (1 << 5) | (1 << 6),
            ),
            (
                StyleFixture {
                    first_column_color: true,
                    conditional_borders: table_style_condition::LAST_COLUMN,
                    ..StyleFixture::default()
                },
                (1 << 7) | (1 << 8),
            ),
        ] {
            let projected = project_table(
                "a\u{7}\u{7}\r",
                &[
                    (0, 2, cell()),
                    (2, 3, row_cells(options, 1)),
                    (3, 4, Vec::new()),
                ],
                fixture,
            );
            assert!(projected.unsupported_table);
            assert_eq!(projected.colors, ["008000"]);
            assert!(projected.borders[0].iter().all(Option::is_none));
        }
    }

    #[test]
    fn eligible_corner_condition_gates_conditional_border_projection() {
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells((1 << 5) | (1 << 6) | (1 << 7), 1)),
                (3, 4, Vec::new()),
            ],
            StyleFixture {
                conditional_borders: table_style_condition::LAST_ROW,
                top_left_color: true,
                ..StyleFixture::default()
            },
        );
        assert!(projected.unsupported_table);
        assert!(projected
            .borders
            .iter()
            .all(|sides| sides.iter().all(Option::is_none)));
    }

    #[test]
    fn conditional_borders_override_only_the_active_style_region() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                table_borders: true,
                conditional_borders: table_style_condition::FIRST_ROW,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(1 << 5, 3)),
        );
        let red = Some(("ff0000".into(), 1.0));
        assert_eq!(
            projected.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                std::array::from_fn(|_| red.clone()),
                std::array::from_fn(|_| red.clone()),
                std::array::from_fn(|_| red.clone()),
                std::array::from_fn(|_| red.clone()),
                std::array::from_fn(|_| red.clone()),
                std::array::from_fn(|_| red.clone()),
            ]
        );
    }

    #[test]
    fn conditional_border_projection_gates_unverified_table_shapes_and_direct_layers() {
        let condition = table_style_condition::FIRST_ROW;
        let options = 1 << 5;
        for rows in [
            [
                row_cells(options, 3),
                row_cells_with_width(options, 3, 1, 900),
                row_cells(options, 3),
            ],
            [
                row_cells(options, 3),
                row_cells_with_extra(options, 3, &sprm(0x9601, &120i16.to_le_bytes())),
                row_cells(options, 3),
            ],
            [
                row_cells(options, 3),
                row_cells_with_extra(options, 3, &sprm(0x3404, &[1])),
                row_cells(options, 3),
            ],
            [
                row_cells(options, 3),
                row_cells(0, 3),
                row_cells(options, 3),
            ],
            [
                row_cells_with_extra(options, 3, &sprm(0x560b, &1u16.to_le_bytes())),
                row_cells(options, 3),
                row_cells(options, 3),
            ],
            [
                row_cells_with_extra(options, 3, &sprm(0x5624, &[0, 2])),
                row_cells(options, 3),
                row_cells(options, 3),
            ],
            [
                row_cells_with_extra(options, 3, &cell_borders(0, 1, [0, 0, 0xff], 16)),
                row_cells(options, 3),
                row_cells(options, 3),
            ],
            [
                row_cells_with_extra(options, 3, &table_borders([0, 0x80, 0], 24)),
                row_cells(options, 3),
                row_cells(options, 3),
            ],
        ] {
            let projected = match try_conditional_border_grid_with_rows(
                StyleFixture {
                    conditional_borders: condition,
                    ..StyleFixture::default()
                },
                rows,
            ) {
                Ok(projected) => projected,
                // A styled right-to-left row is rejected outright by the
                // preferred-indent projection gate.
                Err(error) => {
                    assert!(error.contains("right-to-left"), "{error}");
                    continue;
                }
            };
            assert!(projected.unsupported_table);
            assert_ne!(
                projected.borders[0][0],
                Some(("ff0000".into(), 4.0)),
                "conditional top border must remain gated"
            );
        }
    }

    #[test]
    fn unsupported_conditional_border_nil_does_not_project_partial_patches() {
        for fixture in [StyleFixture {
            conditional_borders: table_style_condition::FIRST_ROW,
            conditional_border_nil: true,
            ..StyleFixture::default()
        }] {
            let projected = conditional_border_grid_with_rows(
                fixture,
                std::array::from_fn(|_| row_cells(1 << 5, 3)),
            );
            assert!(projected.unsupported_table);
            assert!(projected
                .borders
                .iter()
                .all(|sides| sides.iter().all(Option::is_none)));
        }
    }

    #[test]
    fn first_column_and_first_row_conditional_borders_project_in_specified_order() {
        let rows = std::array::from_fn(|_| row_cells((1 << 5) | (1 << 7), 3));
        let row_then_column = conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_ROW,
                conditional_borders_second: table_style_condition::FIRST_COLUMN,
                ..StyleFixture::default()
            },
            rows.clone(),
        );
        let column_then_row = conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_COLUMN,
                conditional_borders_second: table_style_condition::FIRST_ROW,
                ..StyleFixture::default()
            },
            rows,
        );
        assert_eq!(row_then_column.borders, column_then_row.borders);
        assert!(
            !row_then_column.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        let none = [None, None, None, None];
        assert_eq!(
            row_then_column.borders,
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                ],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("00ffff".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none.clone(),
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                none.clone(),
                none,
            ]
        );
    }

    #[test]
    fn remaining_row_column_border_pairs_project_exact_region_edges_in_both_orders() {
        for (row_condition, row_option, column_condition, column_option, corner) in [
            (
                table_style_condition::FIRST_ROW,
                1 << 5,
                table_style_condition::LAST_COLUMN,
                1 << 8,
                2,
            ),
            (
                table_style_condition::LAST_ROW,
                1 << 6,
                table_style_condition::FIRST_COLUMN,
                1 << 7,
                6,
            ),
            (
                table_style_condition::LAST_ROW,
                1 << 6,
                table_style_condition::LAST_COLUMN,
                1 << 8,
                8,
            ),
        ] {
            let rows = std::array::from_fn(|_| row_cells(row_option | column_option, 3));
            let row_then_column = conditional_border_grid_with_rows(
                StyleFixture {
                    conditional_borders: row_condition,
                    conditional_borders_second: column_condition,
                    ..StyleFixture::default()
                },
                rows.clone(),
            );
            let column_then_row = conditional_border_grid_with_rows(
                StyleFixture {
                    conditional_borders: column_condition,
                    conditional_borders_second: row_condition,
                    ..StyleFixture::default()
                },
                rows,
            );
            assert_eq!(row_then_column.borders, column_then_row.borders);
            assert_eq!(
                row_then_column.borders[corner],
                [
                    Some(("ff0000".into(), 4.0)),
                    Some((
                        if column_condition == table_style_condition::FIRST_COLUMN {
                            "008000"
                        } else {
                            "00ffff"
                        }
                        .into(),
                        4.0,
                    )),
                    Some(("0000ff".into(), 4.0)),
                    Some((
                        if column_condition == table_style_condition::LAST_COLUMN {
                            "ffff00"
                        } else {
                            "00ffff"
                        }
                        .into(),
                        4.0,
                    )),
                ]
            );
            let row_inside_v = if row_condition == table_style_condition::FIRST_ROW {
                0
            } else {
                6
            };
            assert_eq!(
                row_then_column.borders[row_inside_v][3],
                Some(("00ffff".into(), 4.0))
            );
            let column = if column_condition == table_style_condition::FIRST_COLUMN {
                0
            } else {
                2
            };
            let column_inside_h = if row_condition == table_style_condition::FIRST_ROW {
                3 + column
            } else {
                column
            };
            assert_eq!(
                row_then_column.borders[column_inside_h][2],
                Some(("ff00ff".into(), 4.0))
            );
        }
    }

    #[test]
    fn inherited_first_column_borders_replace_per_side_on_the_full_region() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                inherited_first_column_borders: true,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(1 << 7, 3)),
        );
        assert_eq!(
            [
                projected.borders[0].clone(),
                projected.borders[3].clone(),
                projected.borders[6].clone()
            ],
            [
                [
                    Some(("ff0000".into(), 4.0)),
                    Some(("000000".into(), 5.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("000000".into(), 5.0)),
                    Some(("ff00ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
                [
                    Some(("ff00ff".into(), 4.0)),
                    Some(("000000".into(), 5.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
            ]
        );
        assert!(projected
            .borders
            .iter()
            .enumerate()
            .filter(|(index, _)| index % 3 != 0)
            .all(|(_, sides)| sides.iter().all(Option::is_none)));
    }

    #[test]
    fn inherited_conditional_borders_project_without_ungating_conditional_shading() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_ROW,
                inherited_partial_conditional_borders: true,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(1 << 5, 3)),
        );
        assert!(
            !projected.unsupported_table,
            "supported TIstd/TTlp selection is admitted"
        );
        assert_eq!(
            projected.borders[..3],
            [
                [
                    Some(("000000".into(), 5.0)),
                    Some(("008000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ff8000".into(), 4.0)),
                ],
                [
                    Some(("000000".into(), 5.0)),
                    Some(("ff8000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ff8000".into(), 4.0)),
                ],
                [
                    Some(("000000".into(), 5.0)),
                    Some(("ff8000".into(), 4.0)),
                    Some(("0000ff".into(), 4.0)),
                    Some(("ffff00".into(), 4.0)),
                ],
            ]
        );
        assert!(projected.borders[3..]
            .iter()
            .all(|sides| sides.iter().all(Option::is_none)));
    }

    #[test]
    fn first_row_left_overrides_first_column_left_in_both_record_orders() {
        let rows = std::array::from_fn(|_| row_cells((1 << 5) | (1 << 7), 3));
        for order in [1, 2] {
            let projected = conditional_border_grid_with_rows(
                StyleFixture {
                    distinct_multi_condition_order: order,
                    ..StyleFixture::default()
                },
                rows.clone(),
            );
            assert_eq!(projected.borders[0][1], Some(("ff00ff".into(), 4.0)));
            assert_eq!(projected.borders[3][1], Some(("008000".into(), 3.0)));
            assert!(projected.borders[1][1].is_none());
        }
    }

    #[test]
    fn repeated_first_row_records_compose_disjoint_sides() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                repeated_conditional_borders: 1,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(1 << 5, 3)),
        );
        assert_eq!(
            projected.borders[0],
            [
                Some(("ff0000".into(), 3.0)),
                None,
                Some(("0000ff".into(), 2.0)),
                None,
            ]
        );
        assert!(projected.borders[3..]
            .iter()
            .all(|sides| sides.iter().all(Option::is_none)));
    }

    #[test]
    fn repeated_first_row_side_uses_later_serialized_value() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                repeated_conditional_borders: 2,
                ..StyleFixture::default()
            },
            std::array::from_fn(|_| row_cells(1 << 5, 3)),
        );
        for cell in &projected.borders[..3] {
            assert_eq!(cell[0], Some(("0000ff".into(), 2.0)));
            assert!(cell[1..].iter().all(Option::is_none));
        }
    }

    #[test]
    fn malformed_conditional_border_operand_is_rejected() {
        let text = "a\u{7}\u{7}\r";
        let error = try_project_table(
            text,
            &[(0, 2, cell()), (2, 3, row(1 << 5)), (3, 4, Vec::new())],
            StyleFixture {
                conditional_borders: table_style_condition::FIRST_ROW,
                conditional_border_malformed: true,
                ..StyleFixture::default()
            },
        )
        .err()
        .expect("malformed border must fail");
        assert_eq!(
            error,
            "UNSUPPORTED:invalid Word conditional table-style border"
        );
    }

    #[test]
    fn late_border_payload_obeys_model_budget_before_retention() {
        let value = table::PreparedBorder::read(&border_bytes([0xff, 0, 0], 8), false).unwrap();
        let mut insufficient = ModelBudget::new(0);
        assert!(matches!(
            borders::materialize_border(value, &mut insufficient),
            Err(message) if message == "OUTPUT_TOO_LARGE"
        ));

        let mut adequate = ModelBudget::new(1024);
        let border = borders::materialize_border(value, &mut adequate).unwrap();
        let spec = border.direct_spec();
        assert_eq!(spec.color.as_deref(), Some("ff0000"));
        assert_eq!(spec.width, 1.0);
    }

    #[test]
    fn native_story_keeps_tc80_border_ownership_explicit_under_table_style() {
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells_with_tc80(1)),
                (3, 4, Vec::new()),
            ],
            StyleFixture {
                table_borders: true,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.borders,
            [std::array::from_fn(|_| Some(("0000ff".into(), 2.0)))]
        );
        assert!(projected.unsupported_table);
    }

    #[test]
    fn unstyled_native_nil_cell_border_keeps_existing_projection() {
        let mut nil = cell_borders(0, 1, [0, 0, 0], 8);
        nil[6..].fill(0xff);
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, unstyled_row_with_border(&nil)),
                (3, 4, Vec::new()),
            ],
            StyleFixture::default(),
        );
        assert_eq!(
            projected.border_styles,
            [std::array::from_fn(|_| Some("nil".into()))]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn unused_default_table_style_does_not_gate_unstyled_native_borders() {
        let mut nil = cell_borders(0, 1, [0, 0, 0], 8);
        nil[6..].fill(0xff);
        let old = old_cell_borders(0, 1, 5, 16);
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, unstyled_row_with_tc80()),
                (3, 5, cell()),
                (5, 6, unstyled_row_with_border(&old)),
                (6, 8, cell()),
                (8, 9, unstyled_row_with_border(&nil)),
                (9, 10, Vec::new()),
            ],
            StyleFixture {
                default_table_style: true,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.borders,
            [
                std::array::from_fn(|_| Some(("0000ff".into(), 2.0))),
                std::array::from_fn(|_| Some(("ff00ff".into(), 2.0))),
                std::array::from_fn(|_| Some((String::new(), 0.5))),
            ]
        );
        assert_eq!(
            projected.border_styles,
            [
                std::array::from_fn(|_| Some("single".into())),
                std::array::from_fn(|_| Some("single".into())),
                std::array::from_fn(|_| Some("nil".into())),
            ]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_resolves_d632_above_d63e_and_d63e_above_direct_d634() {
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (
                    2,
                    3,
                    row_cells_with_margins(1, &[cell_margin(0xd634, 0x0f, 0, 0)]),
                ),
                (3, 5, cell()),
                (
                    5,
                    6,
                    row_cells_with_margins(1, &[cell_margin(0xd632, 0x02, 0, 0)]),
                ),
                (6, 8, cell()),
                (
                    8,
                    9,
                    row_cells_with_margins(1, &[cell_margin(0xd632, 0x02, 3, 720)]),
                ),
                (9, 10, Vec::new()),
            ],
            StyleFixture {
                margins: MarginFixture::D63e,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.margins,
            [
                [0.0, 18.0, 0.0, 0.0],
                [0.0, 0.0, 0.0, 5.4],
                [0.0, 36.0, 0.0, 5.4],
            ]
        );
        assert_eq!(
            projected.margin_wires,
            [
                ["0", "360", "0", "0"],
                ["0", "0", "0", "108"],
                ["0", "720", "0", "108"],
            ]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_projects_bounded_first_row_d63e_on_all_physical_sides() {
        for (options, first) in [(1 << 5, 14.4), (0, 3.6)] {
            let projected = conditional_border_grid_with_rows(
                StyleFixture {
                    margins: MarginFixture::ConditionalD63e,
                    ..StyleFixture::default()
                },
                std::array::from_fn(|_| row_cells(options, 3)),
            );
            assert_eq!(projected.margins.len(), 9);
            for margin in &projected.margins[..3] {
                assert_eq!(*margin, [first, first, first, first]);
            }
            for margin in &projected.margins[3..] {
                assert_eq!(*margin, [3.6, 3.6, 3.6, 3.6]);
            }
            assert!(
                !projected.unsupported_table,
                "supported TIstd/TTlp selection is admitted"
            );
        }
    }

    #[test]
    fn native_story_keeps_direct_d632_dxa_nil_and_omission_above_first_row_d63e() {
        let direct = [
            cell_margin_range(0xd632, 0, 1, 0x02, 3, 432),
            cell_margin_range(0xd632, 1, 2, 0x02, 0, 0),
        ];
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                margins: MarginFixture::ConditionalD63e,
                ..StyleFixture::default()
            },
            [
                row_cells_with_options_and_margins(1 << 5, 3, &direct),
                row_cells(1 << 5, 3),
                row_cells(1 << 5, 3),
            ],
        );

        assert_eq!(projected.margins[0][1], 21.6);
        assert_eq!(projected.margins[1][1], 0.0);
        assert_eq!(projected.margins[2][1], 14.4);
    }

    #[test]
    fn native_story_keeps_direct_right_d632_dxa_nil_and_omission_above_first_row_d63e() {
        let direct = [
            cell_margin_range(0xd632, 0, 1, 0x08, 3, 432),
            cell_margin_range(0xd632, 1, 2, 0x08, 0, 0),
        ];
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                margins: MarginFixture::ConditionalD63e,
                ..StyleFixture::default()
            },
            [
                row_cells_with_options_and_margins(1 << 5, 3, &direct),
                row_cells(1 << 5, 3),
                row_cells(1 << 5, 3),
            ],
        );

        assert_eq!(projected.margins[0][3], 21.6);
        assert_eq!(projected.margins[1][3], 0.0);
        assert_eq!(projected.margins[2][3], 14.4);
    }

    #[test]
    fn native_story_gates_conditional_margins_for_merged_table_shape() {
        let projected = conditional_border_grid_with_rows(
            StyleFixture {
                margins: MarginFixture::ConditionalD63e,
                ..StyleFixture::default()
            },
            [
                row_cells_with_horizontal_merge(1 << 5, 0, 2),
                row_cells(1 << 5, 3),
                row_cells(1 << 5, 3),
            ],
        );

        assert!(projected.unsupported_table);
        for margin in projected.margins {
            assert_eq!(margin, [3.6, 3.6, 3.6, 3.6]);
        }
    }

    #[test]
    fn native_story_keeps_d634_nil_distinct_from_omission_in_style_inheritance() {
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells_with_style(0, 1, 2)),
                (3, 5, cell()),
                (5, 6, row_cells_with_style(0, 1, 3)),
                (6, 7, Vec::new()),
            ],
            StyleFixture {
                inherited_margins: true,
                ..StyleFixture::default()
            },
        );
        assert_eq!(projected.margins[0][1], 0.0);
        assert_eq!(projected.margins[1][1], 36.0);
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_distinguishes_direct_d634_nil_dxa_zero_and_omission() {
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}c\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells_with_style(0, 1, 1)),
                (3, 5, cell()),
                (
                    5,
                    6,
                    row_cells_with_margins(1, &[cell_margin(0xd634, 0x02, 0, 0)]),
                ),
                (6, 8, cell()),
                (
                    8,
                    9,
                    row_cells_with_margins(1, &[cell_margin(0xd634, 0x02, 3, 0)]),
                ),
                (9, 10, Vec::new()),
            ],
            StyleFixture {
                margins: MarginFixture::D634,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.margins.iter().map(|m| m[1]).collect::<Vec<_>>(),
            [18.0, 0.0, 0.0]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_resolves_raw_nil_auto_explicit_and_omitted_cell_shading() {
        let mut raw = vec![30];
        raw.extend([255, 255, 255, 255, 255, 255, 255, 255, 0, 0]);
        raw.extend([0, 0, 0, 255, 0, 0, 0, 255, 0, 0]);
        raw.extend(&cell_shading([0xff, 0x00, 0x00])[1..]);
        let projected = project_table(
            "a\u{7}b\u{7}c\u{7}d\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 8, cell()),
                (8, 9, row_cells_with_raw(0, 4, &raw)),
                (9, 10, Vec::new()),
            ],
            StyleFixture {
                shading: ShadingFixture::Unconditional,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.backgrounds,
            [
                Some("008000".into()),
                None,
                Some("ff0000".into()),
                Some("008000".into()),
            ]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_uses_modern_compatibility_shading_only_for_explicit_raw_nil() {
        let compatibility = cell_shading([0, 0, 0xff]);
        let raw_nil = [10, 255, 255, 255, 255, 255, 255, 255, 255, 0, 0];
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (
                    2,
                    3,
                    row_cells_with_compatibility_and_raw(0, 1, &compatibility, &raw_nil),
                ),
                (3, 4, Vec::new()),
            ],
            StyleFixture {
                shading: ShadingFixture::Unconditional,
                ..StyleFixture::default()
            },
        );

        assert_eq!(projected.backgrounds, [Some("0000ff".into())]);
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_treats_raw_ipat_nil_as_no_fill_without_conflating_shd_nil() {
        let raw = [
            10, // one Shd
            0, 0, 0, 255, // automatic foreground
            0x12, 0x34, 0x56, 0, // RGB background
            0xff, 0xff, // ordinary ipatNil, not the ShdNil sentinel
        ];
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, unstyled_row_with_raw(&raw)),
                (3, 4, Vec::new()),
            ],
            StyleFixture {
                modern_raw: true,
                ..StyleFixture::default()
            },
        );

        assert_eq!(projected.backgrounds, [None]);
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn native_story_keeps_older_version_compatibility_cell_shading() {
        let projected = project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (
                    2,
                    3,
                    row_cells_with_compatibility_shading(0, 1, &cell_shading([0x12, 0x34, 0x56])),
                ),
                (3, 4, Vec::new()),
            ],
            StyleFixture::default(),
        );
        assert_eq!(projected.backgrounds, [Some("123456".into())]);
    }

    #[test]
    fn native_story_layers_conditional_cell_shading_in_doc_order() {
        let options = (1 << 5) | (1 << 7);
        let projected = project_table(
            "a\u{7}b\u{7}\u{7}c\u{7}d\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 5, row_cells(options, 2)),
                (5, 7, cell()),
                (7, 9, cell()),
                (9, 10, row_cells(options, 2)),
                (10, 11, Vec::new()),
            ],
            StyleFixture {
                shading: ShadingFixture::Conditional,
                ..StyleFixture::default()
            },
        );
        assert_eq!(
            projected.backgrounds,
            [
                Some("008000".into()),
                Some("ff0000".into()),
                Some("0000ff".into()),
                Some("000000".into()),
            ]
        );
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn project_uses_ttp_options_and_source_rows_for_conditional_color_keys() {
        let projected = projected_table(StyleFixture {
            first_row_color: true,
            ..StyleFixture::default()
        });
        // The first three entries reproduce the Office row-first color
        // control (green, red, blue); the fourth keeps the disabled-band
        // row-local TTlp regression in the same acquisition path.
        assert_eq!(projected.colors, ["008000", "ff0000", "0000ff", "000000"]);
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn project_does_not_select_or_exclude_an_empty_first_row_condition() {
        let projected = projected_table(StyleFixture::default());
        // Office keeps the ordinary row bands unshifted when the enabled
        // first-row CCnf is empty: red, blue, red.
        assert_eq!(projected.colors, ["ff0000", "0000ff", "ff0000", "000000"]);
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn project_applies_first_row_pjc_and_uses_its_cross_family_presence() {
        let projected = projected_table(StyleFixture {
            first_row_alignment: true,
            ..StyleFixture::default()
        });
        // Word 16.112.4 applies the first-row PCnf and excludes that row from
        // the CHPX horizontal bands. The final row disables bands via TTlp.
        assert_eq!(projected.colors, ["000000", "ff0000", "0000ff", "000000"]);
        assert_eq!(projected.alignments, ["center", "left", "left", "left"]);
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn project_applies_first_row_size_and_uses_its_cross_family_presence() {
        let projected = projected_table(StyleFixture {
            first_row_size: true,
            ..StyleFixture::default()
        });
        // Word 16.112.4 treats a supported size-only CCnf as first-row
        // presence, excluding that row from the horizontal color bands.
        assert_eq!(projected.colors, ["000000", "ff0000", "0000ff", "000000"]);
        assert_eq!(projected.sizes, [14.0, 10.0, 10.0, 10.0]);
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn project_inherits_conditional_fields_through_selected_child_table_style() {
        let projected = project_table(
            "a\u{7}\u{7}b\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, row_cells_with_style(1 << 5, 1, 2)),
                (3, 5, cell()),
                (5, 6, row_cells_with_style(1 << 5, 1, 2)),
                (6, 7, Vec::new()),
            ],
            StyleFixture {
                inherited_conditions: true,
                ..StyleFixture::default()
            },
        );

        assert_eq!(projected.colors, ["0000ff", "000000"]);
        assert_eq!(projected.sizes, [14.0, 10.0]);
        assert_eq!(projected.alignments, ["center", "left"]);
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn project_layers_combined_supported_table_style_properties_in_observed_order() {
        let text = "a\u{7}b\u{7}c\u{7}\u{7}d\u{7}e\u{7}f\u{7}\u{7}g\u{7}h\u{7}i\u{7}\u{7}\r";
        let options = (1 << 5) | (1 << 7);
        let projected = project_table(
            text,
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 7, row_cells(options, 3)),
                (7, 9, cell()),
                (9, 11, cell()),
                (11, 13, cell()),
                (13, 14, row_cells(options, 3)),
                (14, 16, cell()),
                (16, 18, cell()),
                (18, 20, cell()),
                (20, 21, row_cells(options, 3)),
                (21, 22, Vec::new()),
            ],
            StyleFixture {
                combined: true,
                ..StyleFixture::default()
            },
        );

        assert_eq!(
            projected.colors,
            [
                "008000", "ff0000", "ff0000", "0000ff", "000000", "000000", "0000ff", "000000",
                "000000",
            ]
        );
        assert_eq!(
            projected.sizes,
            [16.0, 14.0, 14.0, 12.0, 10.0, 10.0, 12.0, 10.0, 10.0]
        );
        assert_eq!(
            projected.alignments,
            ["left", "center", "center", "right", "left", "left", "right", "left", "left",]
        );
        assert!(projected
            .ascii_fonts
            .iter()
            .all(|font| font.as_deref() == Some("Times New Roman")));
        assert!(projected
            .high_ansi_fonts
            .iter()
            .all(|font| font.as_deref() == Some("Courier New")));
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn actual_story_keeps_source_conditions_across_ragged_horizontal_merge_and_direct_overrides() {
        let mut raw = vec![30];
        raw.extend([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0, 0]);
        raw.extend([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0, 0]);
        raw.extend(&cell_shading([0xff, 0xff, 0])[1..]);
        let projected = project_table(
            "A\u{7}B\u{7}C\u{7}\u{7}D\u{7}E\u{7}F\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 7, row_cells((1 << 5) | (1 << 7), 3)),
                (7, 9, cell()),
                (9, 11, cell()),
                (11, 13, cell()),
                (
                    13,
                    14,
                    offset_merged_row_with_overrides(
                        1 << 7,
                        1800,
                        [0, 2],
                        &cell_margin_range(0xd632, 2, 3, 0x02, 3, 720),
                        &raw,
                    ),
                ),
                (14, 15, Vec::new()),
            ],
            StyleFixture {
                combined: true,
                shading: ShadingFixture::Conditional,
                margins: MarginFixture::D63e,
                ..StyleFixture::default()
            },
        );

        assert_eq!(projected.markers, ["A", "B", "C", "D", "F"]);
        assert_eq!(projected.row_cell_counts, [3, 2]);
        assert_eq!(projected.col_spans, [1, 2, 2, 4, 1]);
        assert_eq!(
            projected.colors,
            ["008000", "ff0000", "ff0000", "0000ff", "000000"]
        );
        assert_eq!(projected.sizes, [16.0, 14.0, 14.0, 12.0, 10.0]);
        assert_eq!(
            projected.alignments,
            ["left", "center", "center", "right", "left"]
        );
        assert!(projected
            .ascii_fonts
            .iter()
            .all(|font| font.as_deref() == Some("Times New Roman")));
        assert!(projected
            .high_ansi_fonts
            .iter()
            .all(|font| font.as_deref() == Some("Courier New")));
        assert_eq!(
            projected.backgrounds,
            [
                Some("008000".into()),
                Some("ff0000".into()),
                Some("ff0000".into()),
                Some("0000ff".into()),
                Some("ffff00".into()),
            ]
        );
        // The merged output routes both paragraph conditions and retained cell
        // facts through its first source cell. This guards that routing without
        // generalizing other merged-style precedence.
        assert_eq!(
            projected.margins.iter().map(|m| m[1]).collect::<Vec<_>>(),
            [18.0, 18.0, 18.0, 18.0, 36.0]
        );
        assert_eq!(
            projected
                .margin_wires
                .iter()
                .map(|m| m[1].as_str())
                .collect::<Vec<_>>(),
            ["360", "360", "360", "360", "720"]
        );
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    /// A one-cell row selecting table style `style`, with `before` authored
    /// before sprmTIstd and `after` after sprmTTlp.
    fn styled_row(style: u16, before: &[u8], after: &[u8]) -> Vec<u8> {
        let mut row = [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, 1, 0xe8, 3]),
        ]
        .concat();
        row.extend(before);
        row.extend(sprm(0x563a, &style.to_le_bytes()));
        row.extend(sprm(0x740a, &[0xff, 0xff, 0xa0, 0x04]));
        row.extend(after);
        row.extend(sprm(0x2416, &[1]));
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1]));
        }
        row
    }

    fn try_default_styled_table(before: &[u8], after: &[u8]) -> Result<ProjectedTable, String> {
        try_project_table(
            "a\u{7}\u{7}\r",
            &[
                (0, 2, cell()),
                (2, 3, styled_row(11, before, after)),
                (3, 4, Vec::new()),
            ],
            StyleFixture {
                default_table_style: true,
                default_table_style_indent: true,
                ..StyleFixture::default()
            },
        )
    }

    #[test]
    fn native_story_admits_default_style_selection_with_rsid_and_row_preferences() {
        let after = [
            sprm(0xd635, &[5, 0, 1, 3, 0xe8, 3]),
            // [MS-DOC] 2.9.28: ignored because the cell width is ftsDxa.
            sprm(0xd639, &[3, 0, 1, 1]),
            // Preferred indent differs from the physical origin; see
            // table::PreferredIndent.
            sprm(0xf661, &[3, 0x6d, 0]),
            // ftsNil before and a zero dxa after both match the grid.
            sprm(0xf617, &[0, 0, 0]),
            sprm(0xf618, &[3, 0, 0]),
            sprm(0x7479, &[1, 2, 3, 4]),
        ]
        .concat();
        let projected = try_default_styled_table(&sprm(0x7479, &[5, 6, 7, 8]), &after).unwrap();
        assert_eq!(projected.markers, ["a"]);
        assert!(!projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }

    #[test]
    fn native_story_gates_tistd_after_an_unimplemented_replacement() {
        // TVertAlign is replaced by TIstd ([MS-DOC] 2.6.3), but that
        // replacement is not implemented. The same record after TIstd is an
        // ordinary direct override.
        let alignment = sprm(0xd62c, &[3, 0, 1, 1]);
        assert!(
            try_default_styled_table(&alignment, &[])
                .unwrap()
                .unsupported_table
        );
        assert!(
            !try_default_styled_table(&[], &alignment)
                .unwrap()
                .unsupported_table
        );
    }

    #[test]
    fn native_story_rejects_row_preferences_it_cannot_represent() {
        // fNoWrap without an ftsDxa preferred cell width changes wrapping.
        let error = try_default_styled_table(&[], &sprm(0xd639, &[3, 0, 1, 1]))
            .err()
            .unwrap();
        assert!(error.contains("no-wrap"), "{error}");
        // A leading preferred width without a matching physical grid slot.
        let error = try_default_styled_table(&[], &sprm(0xf617, &[3, 0x7c, 0]))
            .err()
            .unwrap();
        assert!(error.contains("preferred row part"), "{error}");
        let error = try_default_styled_table(&[], &sprm(0xf618, &[1, 0, 0]))
            .err()
            .unwrap();
        assert!(error.contains("preferred row part"), "{error}");
    }

    #[test]
    fn native_story_projects_cell_hide_mark() {
        // [MS-DOC] 2.9.26 bArg -> ECMA-376 Part 1 §17.4.21 hideMark.
        for (value, expected) in [(1u8, true), (0, false)] {
            let projected =
                try_default_styled_table(&[], &sprm(0xd642, &[3, 0, 1, value])).unwrap();
            assert!(!projected.unsupported_table, "{value}");
            assert_eq!(projected.hide_marks, [expected]);
        }
    }

    #[test]
    fn native_story_projects_cell_text_flow_as_ecma_text_direction() {
        // [MS-DOC] 2.9.323 TextFlow -> ECMA-376 Part 1 §17.18.93.
        for (text_flow, expected) in [
            (0u8, None),
            (1, Some("tbRl")),
            (3, Some("btLr")),
            (5, Some("tbRlV")),
        ] {
            let projected =
                try_default_styled_table(&[], &sprm(0x7629, &[0, 1, text_flow, 0])).unwrap();
            assert!(!projected.unsupported_table, "{text_flow}");
            assert_eq!(projected.text_directions, [expected.map(str::to_owned)]);
        }
        // grpfTFlrtbv is valid but the shared renderer cannot display it.
        let error = try_default_styled_table(&[], &sprm(0x7629, &[0, 1, 4, 0]))
            .err()
            .unwrap();
        assert!(error.contains("grpfTFlrtbv"), "{error}");
        // An undefined TextFlow value is invalid.
        assert!(try_default_styled_table(&[], &sprm(0x7629, &[0, 1, 2, 0])).is_err());
        // Authored before TIstd, its replacement is unobserved.
        assert!(
            try_default_styled_table(&sprm(0x7629, &[0, 1, 5, 0]), &[])
                .unwrap()
                .unsupported_table
        );
    }

    fn cell_border_sides(sides: u8, border: [u8; 8]) -> Vec<u8> {
        let mut operand = vec![11, 0, 1, sides];
        operand.extend(border);
        sprm(0xd62f, &operand)
    }

    #[test]
    fn native_story_projects_post_tistd_nil_cell_borders_as_explicit_absence() {
        // All six sides, including both diagonals, carry NilBrc.
        let nil = cell_border_sides(0x3f, [0xff; 8]);
        let projected = try_default_styled_table(&[], &nil).unwrap();
        assert!(!projected.unsupported_table);
        assert_eq!(
            projected.border_styles,
            [std::array::from_fn(|_| Some("nil".to_string()))]
        );

        // Nil diagonals are the default absence of a diagonal.
        assert_eq!(projected.diagonals, [[None, None]]);
    }

    #[test]
    fn native_story_projects_cell_diagonals_as_ecma_tl2br_and_tr2bl() {
        // [MS-DOC] 2.9.305: 0x10 top-left to bottom-right, 0x20 top-right to
        // bottom-left (ECMA-376 §17.4.73 / §17.4.79).
        let after = [
            cell_border_sides(0x10, border_bytes([0xff, 0, 0], 8)),
            cell_border_sides(0x20, border_bytes([0, 0, 0xff], 4)),
        ]
        .concat();
        let projected = try_default_styled_table(&[], &after).unwrap();
        assert!(!projected.unsupported_table);
        assert_eq!(
            projected.diagonals,
            [[
                Some(("single".to_string(), 1.0, Some("ff0000".to_string()))),
                Some(("single".to_string(), 0.5, Some("0000ff".to_string()))),
            ]]
        );
    }

    #[test]
    fn native_story_reads_word_written_tset_brc80_diagonal_bit() {
        // sample-19.doc row 35: sprmTSetBrc80 bit 0x10 (outside 2.9.304's
        // edge bits) followed by an equal sprmTSetBrc 0x10; the DOCX pair has
        // `w:tl2br w:val="single" w:sz="4" w:color="auto"`.
        let brc80 = sprm(0xd620, &[7, 0, 1, 0x10, 4, 1, 0, 0]);
        let modern = cell_border_sides(0x10, [0, 0, 0, 0xff, 4, 1, 0, 0]);
        let projected = try_default_styled_table(&[], &[brc80.clone(), modern].concat()).unwrap();
        assert!(!projected.unsupported_table);
        assert_eq!(
            projected.diagonals,
            [[Some(("single".to_string(), 0.5, None)), None]]
        );
        // Alone, the 80 diagonal is an old direct value, which stays gated
        // under a table style like any other sprmTSetBrc80 value.
        assert!(
            try_default_styled_table(&[], &brc80)
                .unwrap()
                .unsupported_table
        );
        // Bit 0x20 of the 80 operand has no evidence.
        let tr2bl80 = sprm(0xd620, &[7, 0, 1, 0x20, 4, 1, 0, 0]);
        assert!(
            try_default_styled_table(&[], &tr2bl80)
                .unwrap()
                .unsupported_table
        );
    }

    #[test]
    fn native_story_admits_rtl_direct_borders_when_the_style_has_none() {
        let rtl = [
            sprm(0x560b, &1u16.to_le_bytes()),
            // Equal to the projected origin, so both indent readings agree.
            sprm(0xf661, &[3, 0, 0]),
            cell_borders(0, 1, [0, 0, 0xff], 16),
        ]
        .concat();
        let projected = try_default_styled_table(&[], &rtl).unwrap();
        assert!(!projected.unsupported_table);
        assert_eq!(
            projected.borders,
            [std::array::from_fn(|_| Some(("0000ff".into(), 2.0)))]
        );

        // The default style's inherited zero indent equals the origin too.
        let rtl_inherited = sprm(0x560b, &1u16.to_le_bytes());
        assert!(
            !try_default_styled_table(&[], &rtl_inherited)
                .unwrap()
                .unsupported_table
        );
        let moved = sprm(0x9601, &200i16.to_le_bytes());
        let error = try_default_styled_table(&moved, &rtl_inherited)
            .err()
            .unwrap();
        assert!(error.contains("right-to-left"), "{error}");

        // A differing preferred indent is not covered by the LTR evidence.
        let rtl_indented = [
            sprm(0x560b, &1u16.to_le_bytes()),
            sprm(0xf661, &[3, 0x6d, 0]),
        ]
        .concat();
        let error = try_default_styled_table(&[], &rtl_indented).err().unwrap();
        assert!(error.contains("right-to-left"), "{error}");
    }

    #[test]
    fn native_story_checks_inherited_width_before_against_the_leading_grid() {
        let project = |second_after: &[u8]| {
            try_project_table(
                "a\u{7}\u{7}b\u{7}\u{7}\r",
                &[
                    (0, 2, cell()),
                    (2, 3, styled_row(11, &[], &[])),
                    (3, 5, cell()),
                    (
                        5,
                        6,
                        styled_row(11, &sprm(0x9601, &360i16.to_le_bytes()), second_after),
                    ),
                    (6, 7, Vec::new()),
                ],
                StyleFixture {
                    default_table_style: true,
                    default_table_style_indent: true,
                    ..StyleFixture::default()
                },
            )
        };
        // The default style's zero preferred leading width disagrees with the
        // second row's 360-twip physical leading grid slot.
        let error = project(&[]).err().unwrap();
        assert!(error.contains("preferred row part"), "{error}");
        // A direct preference equal to that slot overrides the style value.
        let projected = project(&sprm(0xf617, &[3, 0x68, 0x01])).unwrap();
        assert_eq!(projected.markers, ["a", "b"]);
        assert!(!projected.unsupported_table);
    }

    #[test]
    fn native_story_drops_cell_frames_that_repeat_the_table_position() {
        // Paragraph-relative vertical and margin-relative horizontal anchors,
        // Y offset 158 twips (YAS_plusOne 159) and 187-twip side distances.
        let position = [
            sprm(0x360d, &[0x60]),
            sprm(0x940f, &159i16.to_le_bytes()),
            sprm(0x9410, &187u16.to_le_bytes()),
            sprm(0x941e, &187u16.to_le_bytes()),
        ]
        .concat();
        let framed_cell = |pc: u8, y: i16| {
            [
                cell(),
                sprm(0x261b, &[pc]),
                sprm(0x8419, &y.to_le_bytes()),
                sprm(0x842f, &187u16.to_le_bytes()),
                sprm(0x2423, &[2]),
            ]
            .concat()
        };
        let project = |cell_papx: Vec<u8>| {
            try_project_table(
                "a\u{7}\u{7}\r",
                &[
                    (0, 2, cell_papx),
                    (2, 3, styled_row(11, &[], &position)),
                    (3, 4, Vec::new()),
                ],
                StyleFixture {
                    default_table_style: true,
                    default_table_style_indent: true,
                    ..StyleFixture::default()
                },
            )
            .unwrap()
        };
        let same = project(framed_cell(0x60, 159));
        assert!(!same.unsupported_paragraph);
        assert!(!same.unsupported_table);
        // A frame that disagrees with the table position stays gated.
        assert!(project(framed_cell(0x60, 200)).unsupported_paragraph);
        assert!(project(framed_cell(0x50, 159)).unsupported_paragraph);
    }

    #[test]
    fn native_story_drops_no_overlap_cell_frames_mirroring_asymmetric_positions() {
        // Paragraph-relative anchors at a zero offset, right/bottom-only
        // wrapping distances and no-overlap, as Word writes a floating table
        // whose OOXML form carries only tblpPr.
        let position = [
            sprm(0x360d, &[0x20]),
            sprm(0x940f, &1i16.to_le_bytes()),
            sprm(0x941e, &187u16.to_le_bytes()),
            sprm(0x941f, &72u16.to_le_bytes()),
            sprm(0x3465, &[1]),
        ]
        .concat();
        let framed_cell = |no_overlap: u8| {
            [
                cell(),
                sprm(0x261b, &[0x20]),
                sprm(0x8419, &1i16.to_le_bytes()),
                sprm(0x2423, &[2]),
                sprm(0x2462, &[no_overlap]),
                // Keeps the fixture PAPX at the odd length the FKP builder needs.
                cell(),
            ]
            .concat()
        };
        let project = |cell_papx: Vec<u8>| {
            try_project_table(
                "a\u{7}\u{7}\r",
                &[
                    (0, 2, cell_papx),
                    (2, 3, styled_row(11, &[], &position)),
                    (3, 4, Vec::new()),
                ],
                StyleFixture {
                    default_table_style: true,
                    default_table_style_indent: true,
                    ..StyleFixture::default()
                },
            )
            .unwrap()
        };
        let mirrored = project(framed_cell(1));
        assert!(!mirrored.unsupported_paragraph);
        assert!(!mirrored.unsupported_table);
        // A no-overlap flag that differs from the table's is not a mirror.
        assert!(project(framed_cell(0)).unsupported_paragraph);
    }
}
