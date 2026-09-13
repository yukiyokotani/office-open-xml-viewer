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

    #[derive(Clone, Copy, Default)]
    struct StyleFixture {
        first_row_color: bool,
        first_row_size: bool,
        first_row_alignment: bool,
        combined: bool,
    }

    struct ProjectedTable {
        colors: Vec<String>,
        sizes: Vec<f64>,
        ascii_fonts: Vec<Option<String>>,
        high_ansi_fonts: Vec<Option<String>>,
        alignments: Vec<String>,
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

        let mut style = vec![0; 14];
        style[2..4].copy_from_slice(&0xfff3u16.to_le_bytes());
        style[4..6].copy_from_slice(&3u16.to_le_bytes());
        let first_row = if fixture.first_row_color {
            &[0x85, 0xca, 8, 1, 0, 0x70, 0x68, 0, 0x80, 0, 0][..]
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
        character.extend(first_row);
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
        let (tapx, paragraph, character) = if fixture.combined {
            (
                &[][..],
                combined_paragraph.as_slice(),
                combined_character.as_slice(),
            )
        } else {
            (&[0x88, 0x34, 1][..], paragraph, character.as_slice())
        };
        for set in [tapx, paragraph, character] {
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
        [
            sprm(0x2416, &[1]),
            sprm(0x2417, &[1]),
            sprm(0x7621, &[0, count, 0xe8, 3]),
            sprm(0x563a, &1u16.to_le_bytes()),
            sprm(0x740a, &[0, 0, options as u8, (options >> 8) as u8]),
            sprm(0x2416, &[1]),
        ]
        .concat()
    }

    fn project_table(
        text: &str,
        runs: &[(usize, usize, Vec<u8>)],
        fixture: StyleFixture,
    ) -> ProjectedTable {
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
            let mut colors = Vec::new();
            let mut sizes = Vec::new();
            let mut ascii_fonts = Vec::new();
            let mut high_ansi_fonts = Vec::new();
            let mut alignments = Vec::new();
            for row in &table.rows {
                for cell in &row.cells {
                    let CellElement::Paragraph(paragraph) = &cell.content[0] else {
                        panic!("paragraph")
                    };
                    let DocRun::Text(run) = &paragraph.runs[0] else {
                        panic!("text")
                    };
                    colors.push(run.color.clone().unwrap());
                    sizes.push(run.font_size);
                    ascii_fonts.push(run.font_family.clone());
                    high_ansi_fonts.push(run.font_family_high_ansi.clone());
                    alignments.push(paragraph.alignment.clone());
                }
            }
            Ok(ProjectedTable {
                colors,
                sizes,
                ascii_fonts,
                high_ansi_fonts,
                alignments,
                unsupported_table: facts.formatting.unsupported_table_properties,
                unsupported_character: facts.formatting.unsupported_character_properties,
                unsupported_paragraph: facts.formatting.unsupported_paragraph_properties,
            })
        })
        .unwrap()
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
    fn project_uses_ttp_options_and_source_rows_for_conditional_color_keys() {
        let projected = projected_table(StyleFixture {
            first_row_color: true,
            ..StyleFixture::default()
        });
        // The first three entries reproduce the Office row-first color
        // control (green, red, blue); the fourth keeps the disabled-band
        // row-local TTlp regression in the same acquisition path.
        assert_eq!(projected.colors, ["008000", "ff0000", "0000ff", "000000"]);
        // TIstd and TTlp remain deliberately admission-gated even though
        // this internal projection verifies their acquired context.
        assert!(projected.unsupported_table);
    }

    #[test]
    fn project_does_not_select_or_exclude_an_empty_first_row_condition() {
        let projected = projected_table(StyleFixture::default());
        // Office keeps the ordinary row bands unshifted when the enabled
        // first-row CCnf is empty: red, blue, red.
        assert_eq!(projected.colors, ["ff0000", "0000ff", "ff0000", "000000"]);
        assert!(projected.unsupported_table);
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
        assert!(projected.unsupported_table);
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
        assert!(projected.unsupported_table);
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
        // TIstd and TTlp remain admission-gated independently of the verified
        // internal formatting projection.
        assert!(projected.unsupported_table);
        assert!(!projected.unsupported_character);
        assert!(!projected.unsupported_paragraph);
    }
}
