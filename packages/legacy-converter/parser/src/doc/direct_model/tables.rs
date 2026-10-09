//! Direct projection for tables planned by the shared MS-DOC table grammar.

use super::{payload, ModelBudget};
use crate::doc::{
    table::{cell_text_flow, Color, PreferredIndent, PreferredWidth, Properties},
    table_structure::{
        self, Assembler, FrameKey, LogicalTable, Payload, PlannedRow, RawEvent, RawTable,
    },
    unsupported,
};
use docx_model::{
    CellBorders, CellElement, DocParagraph, DocTable, DocTableCell, DocTableRow, TableBorders,
    TableCellLayoutAcquisitionWire, TableGridAcquisitionWire, TableGridColumnAcquisitionWire,
    TableLayoutAcquisitionWire, TableLayoutKindAcquisitionWire, TableMarginAcquisitionWire,
    TablePropertyExceptionAcquisitionWire, TableRowHeightAcquisitionWire,
    TableRowLayoutAcquisitionWire, TableWidthAcquisitionWire,
};

pub(super) enum Block {
    Paragraph(Box<DocParagraph>),
    Table(Box<DocTable>),
    PageBreak {
        same_paragraph_as_previous: Option<bool>,
    },
    ColumnBreak,
}

#[derive(Default)]
pub(super) struct Blocks(pub(super) Vec<Block>);

impl Payload for Blocks {
    fn append<A: FnMut(usize) -> Result<(), String>>(
        &mut self,
        mut other: Self,
        admit: &mut A,
    ) -> Result<(), String> {
        reserve(&mut self.0, other.0.len(), admit)?;
        self.0.append(&mut other.0);
        Ok(())
    }
    fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

pub(super) struct Writer<'a> {
    structure: Assembler<Blocks>,
    sequence: &'a mut usize,
    positioned_tables: bool,
}

impl<'a> Writer<'a> {
    /// A writer that rejects absolutely positioned tables.
    #[cfg(test)]
    pub(super) fn new(sequence: &'a mut usize) -> Self {
        Self::with_positioned_tables(sequence, false)
    }

    /// `positioned_tables` admits [MS-DOC] 2.6.3 table positioning as a
    /// floating table. Only the main document story enables it: Word ignores
    /// OOXML table positioning in notes, comments and text boxes
    /// ([MS-OI29500] 2.1.162), and DOC header/footer or note positioning has
    /// not been observed, so those stories keep positioned tables gated.
    pub(super) fn with_positioned_tables(sequence: &'a mut usize, positioned_tables: bool) -> Self {
        Self {
            structure: Assembler::new(),
            sequence,
            positioned_tables,
        }
    }

    /// The document-wide table identity counter, for a nested story (such as
    /// a textbox) projected while this story's tables are still open.
    pub(super) fn sequence(&mut self) -> &mut usize {
        self.sequence
    }

    /// `first_frame` supplies the [MS-DOC] 2.4.3 frame key when the grammar
    /// reaches a row's first-cell first paragraph. The story producer replays
    /// the key its context index acquired, so both passes segment alike.
    pub(super) fn push(
        &mut self,
        props: Properties,
        mark: char,
        paragraph: Blocks,
        first_frame: &mut impl FnMut() -> Result<FrameKey, String>,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        let remaining = std::cell::Cell::new(budget.remaining_bytes);
        let sequence = &mut *self.sequence;
        let positioned_tables = self.positioned_tables;
        self.structure.push_raw(
            props,
            mark,
            paragraph,
            None,
            first_frame,
            |raw, admit| project_tables(raw, sequence, positioned_tables, &remaining, admit),
            &mut |bytes| charge_cell(&remaining, bytes),
        )?;
        budget.remaining_bytes = remaining.get();
        Ok(())
    }

    pub(super) fn finish(self, budget: &mut ModelBudget) -> Result<Blocks, String> {
        let remaining = std::cell::Cell::new(budget.remaining_bytes);
        let positioned_tables = self.positioned_tables;
        let blocks = self.structure.finish_raw(
            |raw, admit| project_tables(raw, self.sequence, positioned_tables, &remaining, admit),
            &mut |bytes| charge_cell(&remaining, bytes),
        )?;
        budget.remaining_bytes = remaining.get();
        Ok(blocks)
    }
}

fn project_tables<A: FnMut(usize) -> Result<(), String>>(
    RawEvent(raw_tables): RawEvent<Blocks>,
    sequence: &mut usize,
    positioned_tables: bool,
    remaining: &std::cell::Cell<usize>,
    admit: &mut A,
) -> Result<Blocks, String> {
    let mut output = Blocks::default();
    for raw in raw_tables {
        let cell_frame = if positioned_tables {
            homogeneous_cell_frame(&raw)
        } else {
            None
        };
        let plan = table_structure::plan(raw, admit)?;
        let table = project_table(plan, cell_frame, *sequence, positioned_tables, remaining)?;
        *sequence = sequence.checked_add(1).ok_or("OUTPUT_TOO_LARGE")?;
        reserve(&mut output.0, 1, &mut |n| charge_cell(remaining, n))?;
        output.0.push(Block::Table(Box::new(table)));
    }
    Ok(output)
}

fn project_table(
    plan: LogicalTable<Blocks>,
    cell_frame: Option<Box<docx_model::FramePr>>,
    sequence: usize,
    positioned_tables: bool,
    remaining: &std::cell::Cell<usize>,
) -> Result<DocTable, String> {
    let first = &plan.rows[0].source;
    if first.shading.is_some() {
        return Err(unsupported(
            "direct DOC model cannot retain table-level shading",
        ));
    }
    let (tblp_pr, overlap) = first.position.direct();
    // [MS-DOC] 2.6.3/2.7.13: nondefault position or wrapping properties make
    // the table absolutely positioned; the shared model lays such a table
    // out of the ordinary flow (ECMA-376 Part 1 17.4.57).
    let tap_positioned = tblp_pr.is_some();
    if tap_positioned {
        if !positioned_tables {
            return Err(unsupported(
                "direct DOC model cannot position a table outside the main story",
            ));
        }
        first.position.check_direct_floating()?;
    }
    let cell_frame = if tap_positioned { None } else { cell_frame };
    let ordinary_flow = !tap_positioned && cell_frame.is_none();
    let (alignment, physical) = first.alignment;
    let first_bidi = first.bidi;
    let first_autofit = first.autofit;
    let margins = first.margins;
    let table_preferred = first.preferred_width;
    let alignment = if physical && first_bidi {
        2 - alignment
    } else {
        alignment
    };
    // Only the main-story writer enables positioned_tables. Nested and
    // other-story RTL P/O disagreements have no established placement owner
    // in this bounded direct projection and retain the equality fallback.
    let leading_ordinary_table =
        positioned_tables && plan.depth == 1 && ordinary_flow && alignment == 0;
    let centered_acquired_table = positioned_tables && centered_acquired_geometry(&plan);
    let row_count = plan.rows.len();
    let mut col_widths = Vec::new();
    reserve(&mut col_widths, plan.grid.len() - 1, &mut |n| {
        charge_cell(remaining, n)
    })?;
    col_widths.extend(
        plan.grid
            .windows(2)
            .map(|edge| f64::from(edge[1] - edge[0]) / 20.0),
    );
    let mut rows = Vec::new();
    reserve(&mut rows, plan.rows.len(), &mut |n| {
        charge_cell(remaining, n)
    })?;
    for (row_index, planned) in plan.rows.into_iter().enumerate() {
        if planned.source.shading.is_some() {
            return Err(unsupported(
                "direct DOC model cannot retain row table-property shading",
            ));
        }
        check_row_preferences(&planned, leading_ordinary_table || centered_acquired_table)?;
        let mut cells = Vec::new();
        reserve(&mut cells, planned.cells.len(), &mut |n| {
            charge_cell(remaining, n)
        })?;
        for cell in planned.cells {
            let source = &cell.source;
            let facts = source.shading.as_ref().map(|s| s.direct_facts());
            let background = match facts {
                None => None,
                Some(f) if f.pattern == "clear" => {
                    source.shading.as_ref().and_then(|s| s.direct_background())
                }
                Some(f) if f.pattern == "nil" => None,
                Some(f) if f.pattern == "solid" => match f.foreground {
                    Color::Rgb([r, g, b]) => Some(format!("{r:02x}{g:02x}{b:02x}")),
                    Color::Auto => {
                        return Err(unsupported(
                            "direct DOC model cannot resolve automatic solid cell shading",
                        ));
                    }
                },
                Some(_) => {
                    return Err(unsupported(
                        "direct DOC model cannot retain patterned cell shading",
                    ));
                }
            };
            if source.flags & ((1 << 12) | (1 << 14)) != 0 {
                return Err(unsupported(
                    "direct DOC model cannot retain cell fit/hide facts",
                ));
            }
            // [MS-DOC] 2.9.305 diagonal sides 0x10 (top left to bottom right)
            // and 0x20 (top right to bottom left) are ECMA-376 §17.4.73 tl2br
            // and §17.4.79 tr2bl. A cleared (none/Nil) diagonal is absence.
            let diagonal = |side: usize| {
                source.borders[side]
                    .as_ref()
                    .filter(|border| !border.is_cleared())
                    .map(|border| border.direct_spec())
            };
            let (tl2br, tr2bl) = (diagonal(4), diagonal(5));
            if (tl2br.is_some() || tr2bl.is_some()) && (cell.vertical != 0 || source.flags & 3 != 0)
            {
                // Each merged DOC cell carries its own TC; no control shows
                // which member's diagonal Word draws across a merged box.
                return Err(unsupported(
                    "direct DOC model cannot place diagonals on merged cells",
                ));
            }
            let align = (source.flags >> 7) & 3;
            if align > 2 {
                return Err(unsupported("invalid Word vertical cell alignment"));
            }
            // [MS-DOC] 2.9.317 TCGRF textFlow (from TC80 or sprmTTextFlow)
            // is a 2.9.323 TextFlow; ECMA-376 Part 1 §17.18.93 names the same
            // arrangements (observed in a Word DOC/DOCX pair: 5 = tbRlV).
            let text_direction = match cell_text_flow(source.flags) {
                0 => None,
                1 => Some("tbRl"),
                3 => Some("btLr"),
                5 => Some("tbRlV"),
                // The shared renderer lays lrTbV cells out without rotating
                // their East Asian glyphs, so projecting it would misdisplay.
                4 => {
                    return Err(unsupported(
                        "direct DOC model cannot display grpfTFlrtbv cell text flow",
                    ));
                }
                _ => return Err(unsupported("invalid Word cell text flow")),
            };
            let mut content = Vec::new();
            reserve(&mut content, cell.content.0.len().max(1), &mut |n| {
                charge_cell(remaining, n)
            })?;
            if cell.vertical == 1 {
                charge_cell(remaining, std::mem::size_of::<DocParagraph>())?;
                content.push(CellElement::Paragraph(Box::default()));
            } else {
                for (block_index, block) in cell.content.0.into_iter().enumerate() {
                    content.push(match block {
                        Block::Paragraph(value) => CellElement::Paragraph(value),
                        Block::Table(mut value) => {
                            // The observed native grid-frame owner starts the
                            // host cell. Later insertion/preceding-paragraph
                            // precedence has not been established; retain its
                            // paragraph facts and the existing limitation.
                            if block_index != 0 && value.table_layout.cell_frame.take().is_some() {
                                value.table_layout.ordinary_flow = value.tblp_pr.is_none();
                            }
                            CellElement::Table(value)
                        }
                        Block::PageBreak { .. } | Block::ColumnBreak => {
                            return Err(unsupported("table cell promoted a paragraph break"))
                        }
                    });
                }
            }
            if text_direction.is_some()
                && content
                    .iter()
                    .any(|element| matches!(element, CellElement::Table(_)))
            {
                // The shared renderer keeps a rotated cell holding a nested
                // table horizontal; projecting it would misdisplay.
                return Err(unsupported(
                    "direct DOC model cannot display rotated cells containing tables",
                ));
            }
            let fallback = |side: usize| match side {
                0 => Some(if row_index == 0 { 0 } else { 4 }),
                1 => Some(if cell.source_index == 0 { 1 } else { 5 }),
                2 => Some(if row_index + 1 == row_count { 2 } else { 4 }),
                3 => Some(if cell.source_end == planned.source_cell_count {
                    3
                } else {
                    5
                }),
                _ => None,
            };
            let border = |side: usize| {
                source.borders[side]
                    .as_ref()
                    .or_else(|| fallback(side).and_then(|s| planned.source.borders[s].as_ref()))
                    .map(|b| b.direct_spec())
            };
            let margins = std::array::from_fn::<_, 4, _>(|s| {
                source.margins[s].unwrap_or(planned.source.margins[s])
            });
            cells.push(DocTableCell {
                content,
                col_span: cell.grid_span as u32,
                v_merge: match cell.vertical {
                    1 => Some(false),
                    3 => Some(true),
                    _ => None,
                },
                borders: CellBorders {
                    top: border(0),
                    left: border(1),
                    bottom: border(2),
                    right: border(3),
                    inside_h: None,
                    inside_v: None,
                    tl2br,
                    tr2bl,
                },
                background,
                v_align: ["top", "center", "bottom"][align as usize].into(),
                width_pt: source.preferred.and_then(|w| match w {
                    PreferredWidth::Dxa(value) => Some(f64::from(value) / 20.0),
                    _ => None,
                }),
                width_pct: source.preferred.and_then(|w| match w {
                    PreferredWidth::Percent(value) => Some(f64::from(value)),
                    _ => None,
                }),
                // [MS-DOC] §2.9.28 fNoWrap maps to ECMA-376 §17.4.29.
                // The shared AutoFit model handles auto/pct preferences;
                // absolute widths keep the specified width priority.
                no_wrap: source.no_wrap.then_some(true),
                margin_top: Some(f64::from(margins[0]) / 20.0),
                margin_left: Some(f64::from(margins[1]) / 20.0),
                margin_bottom: Some(f64::from(margins[2]) / 20.0),
                margin_right: Some(f64::from(margins[3]) / 20.0),
                table_cell_layout: TableCellLayoutAcquisitionWire {
                    preferred_width: source.preferred.map(width),
                    margins: Some(margin_wire(margins)),
                },
                text_direction: text_direction.map(str::to_owned),
                hide_mark: source.hide_mark,
            });
        }
        let height = planned.source.height;
        rows.push(DocTableRow {
            cells,
            grid_before: planned.grid_before as u32,
            grid_after: planned.grid_after as u32,
            row_height: (height != 0).then(|| f64::from(height.abs()) / 20.0),
            row_height_rule: if height < 0 {
                "exact"
            } else if height > 0 {
                "atLeast"
            } else {
                "auto"
            }
            .into(),
            is_header: planned.is_header,
            cant_split: planned.source.cant_split,
            table_row_layout: TableRowLayoutAcquisitionWire {
                height: (height != 0).then(|| TableRowHeightAcquisitionWire {
                    value: Some(height.abs().to_string()),
                    rule: if height < 0 { "exact" } else { "atLeast" }.into(),
                    rule_authored: true,
                }),
                before_width: (planned.grid_before != 0)
                    .then(|| physical_width(planned.width_before)),
                after_width: (planned.grid_after != 0).then(|| physical_width(planned.width_after)),
                justification: None,
                cell_spacing: None,
                style_cell_spacing: None,
                style_cell_margins: None,
                exception: (planned.source.preferred_width != table_preferred).then(|| {
                    TablePropertyExceptionAcquisitionWire {
                        // None must actively clear an inherited table preference;
                        // absence in tblPrEx would inherit it instead.
                        preferred_width: Some(width(
                            planned
                                .source
                                .preferred_width
                                .unwrap_or(PreferredWidth::Auto),
                        )),
                        ..Default::default()
                    }
                }),
            },
        });
    }
    let mut grid_columns = Vec::new();
    reserve(&mut grid_columns, plan.grid.len() - 1, &mut |n| {
        charge_cell(remaining, n)
    })?;
    grid_columns.extend(
        plan.grid
            .windows(2)
            .map(|e| TableGridColumnAcquisitionWire {
                width: Some((e[1] - e[0]).to_string()),
            }),
    );
    let table = DocTable {
        col_widths,
        rows,
        borders: TableBorders::default(),
        cell_margin_top: f64::from(margins[0]) / 20.0,
        cell_margin_left: f64::from(margins[1]) / 20.0,
        cell_margin_bottom: f64::from(margins[2]) / 20.0,
        cell_margin_right: f64::from(margins[3]) / 20.0,
        jc: ["left", "center", "right"][alignment as usize].into(),
        tbl_ind: Some(f64::from(plan.origin) / 20.0),
        layout: Some(if first_autofit { "autofit" } else { "fixed" }.into()),
        width_pt: table_preferred.and_then(|w| match w {
            PreferredWidth::Dxa(value) if value > 0 => Some(f64::from(value) / 20.0),
            _ => None,
        }),
        width_pct: table_preferred.and_then(|w| match w {
            PreferredWidth::Percent(value) if value > 0 => Some(f64::from(value)),
            _ => None,
        }),
        bidi_visual: Some(first_bidi),
        tblp_pr,
        overlap,
        table_layout: TableLayoutAcquisitionWire {
            effective_style_id: None,
            cell_frame,
            ordinary_flow,
            logical_sequence_id: format!("legacy-doc/table/{sequence}"),
            logical_row_offset: 0,
            logical_total_rows: row_count,
            grid: TableGridAcquisitionWire {
                authored: true,
                columns: grid_columns,
                required_column_count: (plan.grid.len() - 1) as u32,
            },
            preferred_width: table_preferred.map(width),
            layout: Some(TableLayoutKindAcquisitionWire {
                kind: Some(if first_autofit { "autofit" } else { "fixed" }.into()),
            }),
            cell_spacing: None,
            cell_margins: Some(margin_wire(margins)),
        },
    };
    charge_cell(
        remaining,
        std::mem::size_of::<DocTable>() + payload::table(&table)?,
    )?;
    Ok(table)
}

/// [MS-DOC] 2.4.3 supplies row identity, not a general nested-frame election
/// rule. Native Word import controls with homogeneous margin/text, auto-size,
/// around-wrapped cell paragraphs move the complete grid, borders and fills
/// together for left/center/right/absolute X and nonzero Y. This bounded owner
/// requires every physical paragraph (including later discarded continuations)
/// to agree with the source first-cell frame. Mixed frames, recursive children,
/// TAP mirrors and later host-cell insertion do not acquire this owner. The
/// observed class is fixed-layout, horizontal and LTR. Other grid policies
/// retain their paragraph facts and residual warning. The original shared
/// column solver determines the acquired grid extent.
fn homogeneous_cell_frame(raw: &RawTable<Blocks>) -> Option<Box<docx_model::FramePr>> {
    if raw.depth != 2 {
        return None;
    }
    let key = raw.rows.first()?.first.frame;
    let frame = crate::doc::paragraph::table_row_frame(key)?;
    if frame.drop_cap != "none"
        || frame.h_anchor != "margin"
        || frame.v_anchor != "text"
        || frame.wrap != "around"
        || frame.w.is_some()
        || frame.h.is_some()
        || frame.h_rule != "auto"
    {
        return None;
    }
    let mut any = false;
    for row in &raw.rows {
        if row.source.position.specifies_nondefault_placement()
            || row.first.frame != key
            || row.source.autofit
            || row.source.bidi
            || row
                .source
                .cells
                .iter()
                .any(|cell| cell_text_flow(cell.flags) != 0)
        {
            return None;
        }
        for cell in &row.cells {
            for block in &cell.0 {
                let Block::Paragraph(paragraph) = block else {
                    return None;
                };
                if paragraph.frame_pr.as_deref() != Some(&frame) {
                    return None;
                }
                any = true;
            }
        }
    }
    any.then(|| Box::new(frame))
}

/// A bounded acquired-geometry projection for whole-frame centered RTL tables.
/// [MS-DOC] 2.6.3 sprmTDxaAbs (-4) and TPc supply physical frame placement;
/// ECMA-376 17.4.57 centers the complete laid-out table. Uniform logical O is
/// retained in tblInd/bidiVisual, then removed by the existing child-flow to
/// positioned-frame translation. It is not added to the external center.
/// This is a projection policy, not a normative P/O precedence rule. Admit
/// only homogeneous main-story, depth-one leading rows with complete positive
/// geometry and absent or agreeing absolute width preferences. AutoFit retains its
/// existing shared solver: cancellation of uniform O holds for its final width
/// as well. Mixed translations, row parts, competing preferences and discarded
/// horizontal continuations need a separate ownership proof and stay refused.
fn centered_acquired_geometry(plan: &LogicalTable<Blocks>) -> bool {
    let first = &plan.rows[0].source;
    let Some(position) = first.position.direct().0 else {
        return false;
    };
    let width = plan.grid.last().copied().unwrap_or(plan.origin) - plan.origin;
    plan.depth == 1
        && position.tblp_x_spec.as_deref() == Some("center")
        && matches!(position.horz_anchor.as_str(), "margin" | "page")
        && position.vert_anchor == "page"
        && position.tblp_y_spec.is_none()
        && position.tblp_y >= 0.0
        && matches!(first.preferred_indent, Some(PreferredIndent::Dxa(_)))
        && matches!(first.preferred_width, None | Some(PreferredWidth::Dxa(_)))
        && plan.grid.windows(2).all(|edges| edges[1] > edges[0])
        && plan.rows.iter().all(|row| {
            let source = &row.source;
            source.bidi
                && source.alignment == first.alignment
                && matches!(source.alignment, (0, false) | (2, true))
                && source.origin() == plan.origin
                && source.preferred_indent == first.preferred_indent
                && source.position == first.position
                && source.autofit == first.autofit
                && source.preferred_width == first.preferred_width
                && source.preferred_width.is_none_or(|preference| matches!(preference, PreferredWidth::Dxa(value) if i32::from(value) == width))
                && row.grid_before == 0
                && row.grid_after == 0
                && row.cells.len() == row.source_cell_count
                && row.cells.iter().all(|cell| {
                    cell.source_end == cell.source_index + 1
                        && cell.source.flags & 3 == 0
                        && cell.source.width > 0
                        && cell.source.preferred.is_none_or(|preference| matches!(preference, PreferredWidth::Dxa(value) if i32::from(value) == cell.source.width))
                })
        })
}

/// Validate the row preferences that the projection represents through the
/// acquired logical row geometry instead of a separate model field.
fn check_row_preferences(
    planned: &PlannedRow<Blocks>,
    acquired_geometry_table: bool,
) -> Result<(), String> {
    let source = &planned.source;
    let (alignment, physical) = source.alignment;
    let logical_alignment = if physical && source.bidi {
        2 - alignment
    } else {
        alignment
    };
    let acquired_leading_origin = acquired_geometry_table && logical_alignment == 0;
    if source.bidi
        && !acquired_leading_origin
        && source.preferred_indent.is_some()
        && !matches!(
            source.preferred_indent,
            Some(PreferredIndent::Dxa(value)) if i32::from(value) == source.origin()
        )
    {
        // MS-DOC 2.6.3/2.9.321 specify acquired logical row geometry;
        // 2.9.102 defines a separate preference. Retaining O rather than P is
        // the bounded direct-projection policy documented on PreferredIndent,
        // not a normative P/O precedence claim. Only ordinary leading main-
        // story, depth-one RTL tables and the bounded whole-frame centered
        // profile above use that policy through bidiVisual. Other placement
        // classes retain the exact-equality fallback.
        return Err(unsupported(
            "direct DOC model cannot place a right-to-left table with a preferred indent",
        ));
    }
    // [MS-DOC] 2.6.3 sprmTWidthBefore/sprmTWidthAfter are the preferred widths
    // of the same leading/trailing row parts whose physical widths the grid
    // projection emits as wBefore/wAfter. Admit them only where both agree, so
    // the projection is the same whichever one Word lays out from. ftsNil is
    // the documented absence of a preference.
    for (preference, grid, physical) in [
        (
            source.preferred_before,
            planned.grid_before,
            planned.width_before,
        ),
        (
            source.preferred_after,
            planned.grid_after,
            planned.width_after,
        ),
    ] {
        match preference {
            None | Some(None) => {}
            Some(Some(PreferredWidth::Dxa(value)))
                if i32::from(value) == if grid == 0 { 0 } else { physical } => {}
            Some(Some(_)) => {
                return Err(unsupported(
                    "direct DOC model cannot reconcile a preferred row part width with its grid",
                ))
            }
        }
    }
    Ok(())
}

fn physical_width(value: i32) -> TableWidthAcquisitionWire {
    TableWidthAcquisitionWire {
        kind: Some("dxa".into()),
        value: Some(value.to_string()),
    }
}
fn width(value: PreferredWidth) -> TableWidthAcquisitionWire {
    TableWidthAcquisitionWire {
        kind: Some(value.kind().into()),
        value: Some(value.value().to_string()),
    }
}
fn margin_wire(v: [u16; 4]) -> TableMarginAcquisitionWire {
    TableMarginAcquisitionWire {
        top: Some(physical_width(v[0].into())),
        left: Some(physical_width(v[1].into())),
        bottom: Some(physical_width(v[2].into())),
        right: Some(physical_width(v[3].into())),
        start: None,
        end: None,
    }
}
fn charge_cell(remaining: &std::cell::Cell<usize>, n: usize) -> Result<(), String> {
    remaining.set(remaining.get().checked_sub(n).ok_or("OUTPUT_TOO_LARGE")?);
    Ok(())
}
fn reserve<T, A: FnMut(usize) -> Result<(), String>>(
    v: &mut Vec<T>,
    add: usize,
    admit: &mut A,
) -> Result<(), String> {
    let target = v.len().checked_add(add).ok_or("OUTPUT_TOO_LARGE")?;
    if target > v.capacity() {
        let old = v.capacity();
        let growth = target - old;
        admit(
            growth
                .checked_mul(std::mem::size_of::<T>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )?;
        v.try_reserve_exact(add)
            .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
        admit(
            v.capacity()
                .saturating_sub(target)
                .checked_mul(std::mem::size_of::<T>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn unframed() -> Result<FrameKey, String> {
        Ok(FrameKey::default())
    }
    fn cell(depth: u32) -> Properties {
        let mut p = Properties::default();
        p.apply(0x6649, &depth.to_le_bytes()).unwrap();
        p
    }
    fn row(depth: u32, widths: &[u16]) -> Properties {
        let mut p = cell(depth);
        p.row_end = true;
        p.inner_row = true;
        for (i, w) in widths.iter().enumerate() {
            let [a, b] = w.to_le_bytes();
            p.row.apply(0x7621, &[i as u8, 1, a, b]).unwrap();
        }
        p
    }
    fn paragraph(text: &str) -> Blocks {
        let mut p = DocParagraph::default();
        p.runs
            .push(docx_model::DocRun::Text(Box::new(docx_model::TextRun {
                text: text.into(),
                ..Default::default()
            })));
        Blocks(vec![Block::Paragraph(Box::new(p))])
    }
    #[test]
    fn preferred_indent_checks_final_cell_sum_not_union_grid_extent() {
        let project = |preferred: i16, shifted: bool| -> Result<Blocks, String> {
            let mut sequence = 0;
            let mut writer = Writer::with_positioned_tables(&mut sequence, true);
            let mut budget = ModelBudget::new(1_000_000);
            for origin in if shifted { vec![0, 360] } else { vec![0] } {
                let mut end = row(1, &[400, 600]);
                end.row.bidi = true;
                end.row.left = origin;
                end.row.preferred_indent = Some(PreferredIndent::Dxa(if shifted && origin == 0 {
                    30_680
                } else {
                    preferred
                }));
                for text in ["a", "b"] {
                    writer.push(
                        cell(1),
                        '\u{7}',
                        paragraph(text),
                        &mut unframed,
                        &mut budget,
                    )?;
                }
                writer.push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)?;
            }
            writer.finish(&mut budget)
        };
        // MS-DOC 2.9.102: P + the final cell-width sum may equal 31680.
        // The shifted second row widens the union grid without widening its
        // own cells. It must not falsely invalidate the legal preference.
        for shifted in [false, true] {
            let blocks = project(30_680, shifted).unwrap();
            let Block::Table(table) = &blocks.0[0] else {
                panic!()
            };
            assert_eq!(table.rows.len(), if shifted { 2 } else { 1 });
            assert_eq!(
                table.col_widths.iter().sum::<f64>(),
                if shifted { 68.0 } else { 50.0 }
            );
            let error = project(30_681, shifted).err().unwrap();
            assert!(error.contains("preferred indent plus row width"), "{error}");
        }
    }

    #[test]
    fn differing_rtl_preference_requires_main_story_top_level_owner() {
        let project = |main_story: bool, nested: bool, preferred: i16| -> Result<Blocks, String> {
            let mut sequence = 0;
            let mut writer = Writer::with_positioned_tables(&mut sequence, main_story);
            let mut budget = ModelBudget::new(1_000_000);
            let depth = if nested { 2 } else { 1 };
            if nested {
                writer.push(
                    cell(1),
                    '\r',
                    paragraph("parent"),
                    &mut unframed,
                    &mut budget,
                )?;
            }
            let mut content = cell(depth);
            content.inner_cell = nested;
            let mut end = row(depth, &[1000]);
            end.row.bidi = true;
            end.row.preferred_indent = Some(PreferredIndent::Dxa(preferred));
            let mark = if nested { '\r' } else { '\u{7}' };
            writer.push(
                content,
                mark,
                paragraph("child"),
                &mut unframed,
                &mut budget,
            )?;
            writer.push(end, mark, Blocks::default(), &mut unframed, &mut budget)?;
            if nested {
                writer.push(
                    cell(1),
                    '\u{7}',
                    Blocks::default(),
                    &mut unframed,
                    &mut budget,
                )?;
                writer.push(
                    row(1, &[1000]),
                    '\u{7}',
                    Blocks::default(),
                    &mut unframed,
                    &mut budget,
                )?;
            }
            writer.finish(&mut budget)
        };
        assert!(project(true, false, 109).is_ok());
        for (main_story, nested) in [(false, false), (true, true)] {
            assert!(project(main_story, nested, 0).is_ok());
            let error = project(main_story, nested, 109).err().unwrap();
            assert!(error.contains("right-to-left"), "{error}");
        }
    }

    #[test]
    fn centered_rtl_table_retains_acquired_geometry_with_a_distinct_preference() {
        let project = |origin: i32, preferred: i16, autofit: bool, variant: u8| {
            let mut sequence = 0;
            let mut writer = Writer::with_positioned_tables(&mut sequence, variant != 1);
            let mut budget = ModelBudget::new(1_000_000);
            for index in 0..2 {
                let widths: &[u16] = if index == 1 && variant == 0 {
                    &[1200]
                } else {
                    &[400, 800]
                };
                let mut end = row(1, widths);
                end.row.bidi = true;
                end.row.left = origin;
                end.row.autofit = autofit;
                end.row.preferred_indent = Some(PreferredIndent::Dxa(preferred));
                end.row.preferred_width = Some(PreferredWidth::Dxa(1200));
                for source in &mut end.row.cells {
                    source.preferred = Some(PreferredWidth::Dxa(source.width as u16));
                }
                for (code, value) in [
                    (0x360d, vec![0x50]),
                    (0x940e, (-4i16).to_le_bytes().to_vec()),
                    (0x940f, 401i16.to_le_bytes().to_vec()),
                    (0x9410, 100u16.to_le_bytes().to_vec()),
                    (0x941e, 200u16.to_le_bytes().to_vec()),
                ] {
                    end.row.position.apply(code, &value).unwrap();
                }
                if variant == 7 {
                    end.row.preferred_width = None;
                    for source in &mut end.row.cells {
                        source.preferred = None;
                    }
                }
                if index == 1 {
                    match variant {
                        2 => end.row.left += 20,
                        3 => end.row.alignment = (1, false),
                        4 => end.row.preferred_indent = Some(PreferredIndent::Dxa(preferred + 20)),
                        5 => end.row.cells[0].preferred = Some(PreferredWidth::Percent(1000)),
                        6 => {
                            end.row.cells[0].flags = 2;
                            end.row.cells[1].flags = 1;
                        }
                        _ => {}
                    }
                }
                for text in ["A", "B"].into_iter().take(end.row.cells.len()) {
                    writer.push(
                        cell(1),
                        '\u{7}',
                        paragraph(text),
                        &mut unframed,
                        &mut budget,
                    )?;
                }
                writer.push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)?;
            }
            writer.finish(&mut budget)
        };
        // Uniform acquired O is retained. Whole-frame centering supplies the
        // external placement even when O/P differ or either one is negative.
        for (origin, preferred) in [(200, 120), (-200, 120), (200, -120)] {
            for autofit in [false, true] {
                let blocks = project(origin, preferred, autofit, 0).unwrap();
                let Block::Table(table) = &blocks.0[0] else {
                    panic!()
                };
                assert_eq!(table.tbl_ind, Some(f64::from(origin) / 20.0));
                assert_eq!(table.col_widths, vec![20.0, 40.0]);
                assert_eq!(table.rows[1].cells.len(), 1);
                assert_eq!(table.rows[1].cells[0].col_span, 2);
                assert_eq!(table.bidi_visual, Some(true));
                assert_eq!(table.jc, "left");
                assert_eq!(
                    table.layout.as_deref(),
                    Some(if autofit { "autofit" } else { "fixed" })
                );
                assert!(!table.table_layout.ordinary_flow);
                let position = table.tblp_pr.as_ref().unwrap();
                assert_eq!(position.tblp_x_spec.as_deref(), Some("center"));
                assert_eq!(
                    (position.horz_anchor.as_str(), position.vert_anchor.as_str()),
                    ("margin", "page")
                );
                assert_eq!(
                    (position.left_from_text, position.right_from_text),
                    (5.0, 10.0)
                );
            }
        }
        // Different owners, translations, alignment or competing width
        // preferences have no centered homogeneous admission proof.
        assert!(project(200, 120, true, 7).is_ok());
        for variant in 1..=6 {
            assert!(
                project(200, 120, true, variant).is_err(),
                "variant {variant}"
            );
        }
    }

    #[test]
    fn plain_table_preserves_zero_grid_slots_merges_and_cell_break_runs() {
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("b"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(
                row(1, &[0, 1000]),
                '\u{7}',
                Blocks::default(),
                &mut unframed,
                &mut budget,
            )
            .unwrap();
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        assert_eq!(table.col_widths, vec![0.0, 50.0]);
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(table.table_layout.logical_row_offset, 0);
        assert_eq!(table.table_layout.logical_total_rows, 1);
        assert!(table.table_layout.ordinary_flow);
    }
    #[test]
    fn tcgrf_text_flow_projects_and_rotated_cells_with_nested_tables_fail_closed() {
        // [MS-DOC] 2.9.317 TCGRF textFlow authored by a TC80 (grpfTFbtlr).
        let mut end = row(1, &[1000]);
        end.row.cells[0].flags |= 3 << 2;
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .unwrap();
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        assert_eq!(
            table.rows[0].cells[0].text_direction.as_deref(),
            Some("btLr")
        );

        // The same flow on a cell holding a nested table is rejected.
        let mut end = row(1, &[1000]);
        end.row.cells[0].flags |= 1 << 2;
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        // Nested cells and rows end at paragraph marks carrying
        // fInnerTableCell / fInnerTtp ([MS-DOC] 2.4.3).
        let mut inner_cell = cell(2);
        inner_cell.inner_cell = true;
        writer
            .push(
                inner_cell,
                '\r',
                paragraph("inner"),
                &mut unframed,
                &mut budget,
            )
            .unwrap();
        writer
            .push(
                row(2, &[500]),
                '\r',
                Blocks::default(),
                &mut unframed,
                &mut budget,
            )
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        let error = writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .and_then(|_| writer.finish(&mut budget).map(|_| ()))
            .err()
            .unwrap();
        assert!(error.contains("rotated cells containing tables"), "{error}");
    }
    #[test]
    fn cell_diagonals_project_and_merged_cell_diagonals_fail_closed() {
        let diagonal =
            || Some(crate::doc::border::Border::read(&[0, 0, 0, 0xff, 4, 1, 0, 0], false).unwrap());
        let mut end = row(1, &[1000]);
        end.row.cells[0].borders[5] = diagonal();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .unwrap();
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        let borders = &table.rows[0].cells[0].borders;
        assert!(borders.tl2br.is_none());
        assert_eq!(borders.tr2bl.as_ref().unwrap().style, "single");

        // The primary of a horizontal merge (TCGRF horzMerge 2) is rejected.
        let mut end = row(1, &[500, 500]);
        end.row.cells[0].flags |= 2;
        end.row.cells[1].flags |= 1;
        end.row.cells[0].borders[4] = diagonal();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("b"), &mut unframed, &mut budget)
            .unwrap();
        let error = writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .and_then(|_| writer.finish(&mut budget).map(|_| ()))
            .err()
            .unwrap();
        assert!(error.contains("diagonals on merged cells"), "{error}");
    }
    #[test]
    fn no_overlap_alone_remains_ordinary_but_positioned_table_fails_closed() {
        let mut end = row(1, &[1000]);
        end.row.apply(0x3465, &[1]).unwrap();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .unwrap();
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        assert_eq!(table.overlap.as_deref(), Some("never"));
        assert!(table.table_layout.ordinary_flow);

        let mut end = row(1, &[1000]);
        end.row.apply(0x940e, &721i16.to_le_bytes()).unwrap();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .unwrap();
        assert!(writer
            .finish(&mut budget)
            .err()
            .unwrap()
            .contains("outside the main story"));
    }

    fn positioned(sprms: &[(u16, &[u8])]) -> Result<DocTable, String> {
        let mut end = row(1, &[1000]);
        for (code, operand) in sprms {
            end.row.apply(*code, operand).unwrap();
        }
        let mut sequence = 0;
        let mut writer = Writer::with_positioned_tables(&mut sequence, true);
        let mut budget = ModelBudget::new(1_000_000);
        writer.push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)?;
        writer.push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)?;
        let mut body = writer.finish(&mut budget)?;
        let Some(Block::Table(table)) = body.0.pop() else {
            panic!("table")
        };
        Ok(*table)
    }

    #[test]
    fn main_story_positioned_table_leaves_the_ordinary_flow() {
        // Paragraph-relative vertical and margin-relative horizontal anchors,
        // a centered X and a 219-twip Y offset (YAS_plusOne 220).
        let table = positioned(&[
            (0x360d, &[0x60]),
            (0x940e, &(-4i16).to_le_bytes()),
            (0x940f, &220i16.to_le_bytes()),
            (0x9410, &180u16.to_le_bytes()),
            (0x941e, &180u16.to_le_bytes()),
            (0x3465, &[1]),
        ])
        .unwrap();
        assert!(!table.table_layout.ordinary_flow);
        let position = table.tblp_pr.unwrap();
        assert_eq!(position.vert_anchor, "text");
        assert_eq!(position.horz_anchor, "margin");
        assert_eq!(position.tblp_x_spec.as_deref(), Some("center"));
        assert_eq!(position.tblp_y, 10.95);
        assert_eq!(position.left_from_text, 9.0);
        assert_eq!(position.right_from_text, 9.0);
        assert_eq!(table.overlap.as_deref(), Some("never"));

        // Reserved anchors mean "not absolutely positioned".
        let table = positioned(&[(0x360d, &[0xf0]), (0x940f, &220i16.to_le_bytes())]).unwrap();
        assert!(table.tblp_pr.is_none());
        assert!(table.table_layout.ordinary_flow);
    }

    #[test]
    fn positioned_tables_without_established_doc_display_fail_closed() {
        // sprmTDyaAbs zero is the inline vertical alignment value.
        let error = positioned(&[(0x360d, &[0x50]), (0x940e, &721i16.to_le_bytes())])
            .err()
            .unwrap();
        assert!(error.contains("inline vertical"), "{error}");
        // The DOC counterpart of the MS-OI29500 2.1.162 ignored tblpPr.
        for x in [0i16, 1] {
            let error = positioned(&[
                (0x360d, &[0x10]),
                (0x940e, &x.to_le_bytes()),
                (0x940f, &1i16.to_le_bytes()),
                (0x9410, &180u16.to_le_bytes()),
            ])
            .err()
            .unwrap();
            assert!(error.contains("zero-offset"), "{error}");
        }
        // A paragraph-relative vertical anchor is outside that exception.
        assert!(positioned(&[(0x360d, &[0x20]), (0x940f, &1i16.to_le_bytes())]).is_ok());
    }

    #[test]
    fn preferred_widths_project_without_replacing_physical_grid_geometry() {
        let mut end = row(1, &[1000, 2000]);
        end.row.preferred_width = Some(PreferredWidth::Percent(2500));
        end.row.cells[0].preferred = Some(PreferredWidth::Dxa(720));
        end.row.cells[1].preferred = Some(PreferredWidth::Percent(1250));
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("b"), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
            .unwrap();
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        assert_eq!(table.col_widths, [50.0, 100.0]);
        assert_eq!(table.width_pt, None);
        assert_eq!(table.width_pct, Some(2500.0));
        assert_eq!(
            table
                .table_layout
                .preferred_width
                .as_ref()
                .unwrap()
                .kind
                .as_deref(),
            Some("pct")
        );
        assert_eq!(table.rows[0].cells[0].width_pt, Some(36.0));
        assert_eq!(table.rows[0].cells[0].width_pct, None);
        assert_eq!(table.rows[0].cells[1].width_pt, None);
        assert_eq!(table.rows[0].cells[1].width_pct, Some(1250.0));
    }

    #[test]
    fn no_wrap_reaches_the_shared_cell_model_for_auto_and_percent_widths() {
        for preferred in [None, Some(PreferredWidth::Percent(1250))] {
            let mut end = row(1, &[1000]);
            end.row.cells[0].preferred = preferred;
            end.row.cells[0].no_wrap = true;
            let mut sequence = 0;
            let mut writer = Writer::new(&mut sequence);
            let mut budget = ModelBudget::new(1_000_000);
            writer
                .push(
                    cell(1),
                    '\u{7}',
                    paragraph("words in a cell"),
                    &mut unframed,
                    &mut budget,
                )
                .unwrap();
            writer
                .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
                .unwrap();
            let body = writer.finish(&mut budget).unwrap();
            let Block::Table(table) = &body.0[0] else {
                panic!()
            };
            assert_eq!(table.rows[0].cells[0].no_wrap, Some(true));
        }
    }

    #[test]
    fn merged_cell_preference_projects_from_the_primary_source_only() {
        fn projected(d635_cell: u8, preferred: u16) -> Box<DocTable> {
            let mut end = cell(1);
            end.row_end = true;
            end.inner_row = true;
            let mut definition = vec![70, 0, 3];
            for boundary in [0i16, 1500, 6000, 9000] {
                definition.extend_from_slice(&boundary.to_le_bytes());
            }
            definition.extend_from_slice(&[0; 60]); // Three ftsNil TC80 records.
            end.row.apply(0xd608, &definition).unwrap();
            let [low, high] = preferred.to_le_bytes();
            end.row
                .apply(0xd635, &[5, d635_cell, d635_cell + 1, 3, low, high])
                .unwrap();
            // MS-DOC 2.6.3 sprmTMerge: cell 0 is the primary; formatting of
            // continuation cell 1 is not applied to the merged layout region.
            end.row.apply(0x5624, &[0, 2]).unwrap();

            let mut sequence = 0;
            let mut writer = Writer::new(&mut sequence);
            let mut budget = ModelBudget::new(1_000_000);
            for text in ["a", "b", "c"] {
                writer
                    .push(
                        cell(1),
                        '\u{7}',
                        paragraph(text),
                        &mut unframed,
                        &mut budget,
                    )
                    .unwrap();
            }
            writer
                .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
                .unwrap();
            let body = writer.finish(&mut budget).unwrap();
            let Block::Table(table) = body.0.into_iter().next().unwrap() else {
                panic!()
            };
            table
        }

        for preferred in [1500, 3000] {
            let table = projected(0, preferred);
            let expected = preferred.to_string();
            assert_eq!(table.col_widths, [75.0, 225.0, 150.0]);
            assert_eq!(table.rows[0].cells.len(), 2);
            assert_eq!(table.rows[0].cells[0].col_span, 2);
            assert_eq!(
                table.rows[0].cells[0].width_pt,
                Some(f64::from(preferred) / 20.0)
            );
            assert_eq!(
                table.rows[0].cells[0]
                    .table_cell_layout
                    .preferred_width
                    .as_ref()
                    .and_then(|width| width.value.as_deref()),
                Some(expected.as_str())
            );
        }
        for continuation_preferred in [1500, 6000] {
            let table = projected(1, continuation_preferred);
            assert_eq!(table.col_widths, [75.0, 225.0, 150.0]);
            assert_eq!(table.rows[0].cells.len(), 2);
            assert_eq!(table.rows[0].cells[0].col_span, 2);
            assert_eq!(table.rows[0].cells[0].width_pt, None);
            assert!(table.rows[0].cells[0]
                .table_cell_layout
                .preferred_width
                .is_none());
        }
    }

    #[test]
    fn merged_raw_tc80_preference_projects_from_the_primary_source_only() {
        fn projected(tc80_cell: Option<usize>, autofit: bool) -> Box<DocTable> {
            let mut end = cell(1);
            end.row_end = true;
            end.inner_row = true;
            let mut definition = vec![70, 0, 3];
            for boundary in [0i16, 1500, 6000, 9000] {
                definition.extend_from_slice(&boundary.to_le_bytes());
            }
            for source_cell in 0..3 {
                let (flags, preferred) = if tc80_cell == Some(source_cell) {
                    (3u16 << 9, 3000u16)
                } else {
                    (0, 0)
                };
                definition.extend_from_slice(&flags.to_le_bytes());
                definition.extend_from_slice(&preferred.to_le_bytes());
                definition.extend_from_slice(&[0; 16]);
            }
            end.row.apply(0xd608, &definition).unwrap();
            end.row.apply(0xf614, &[3, 0x70, 0x17]).unwrap();
            end.row.apply(0x3615, &[u8::from(autofit)]).unwrap();
            // Keep merge ownership independent from the TC80 formatting bits.
            end.row.apply(0x5624, &[0, 2]).unwrap();

            let mut sequence = 0;
            let mut writer = Writer::new(&mut sequence);
            let mut budget = ModelBudget::new(1_000_000);
            for text in ["a", "b", "c"] {
                writer
                    .push(
                        cell(1),
                        '\u{7}',
                        paragraph(text),
                        &mut unframed,
                        &mut budget,
                    )
                    .unwrap();
            }
            writer
                .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
                .unwrap();
            let body = writer.finish(&mut budget).unwrap();
            let Block::Table(table) = body.0.into_iter().next().unwrap() else {
                panic!()
            };
            table
        }

        for autofit in [false, true] {
            for tc80_cell in [None, Some(0), Some(1)] {
                let table = projected(tc80_cell, autofit);
                assert_eq!(table.col_widths, [75.0, 225.0, 150.0]);
                assert_eq!(table.width_pt, Some(300.0));
                assert_eq!(
                    table.layout.as_deref(),
                    Some(if autofit { "autofit" } else { "fixed" })
                );
                assert_eq!(
                    table
                        .table_layout
                        .preferred_width
                        .as_ref()
                        .and_then(|width| width.value.as_deref()),
                    Some("6000")
                );
                assert_eq!(
                    table
                        .table_layout
                        .layout
                        .as_ref()
                        .and_then(|layout| layout.kind.as_deref()),
                    Some(if autofit { "autofit" } else { "fixed" })
                );
                assert_eq!(table.rows[0].cells.len(), 2);
                assert_eq!(table.rows[0].cells[0].col_span, 2);
                let primary = &table.rows[0].cells[0];
                assert_eq!(primary.width_pt, (tc80_cell == Some(0)).then_some(150.0));
                assert_eq!(
                    primary
                        .table_cell_layout
                        .preferred_width
                        .as_ref()
                        .and_then(|width| width.value.as_deref()),
                    (tc80_cell == Some(0)).then_some("3000")
                );
            }
        }
    }

    #[test]
    fn row_nil_preference_projects_an_auto_exception_to_clear_table_width() {
        let mut first = row(1, &[1000]);
        first.row.preferred_width = Some(PreferredWidth::Dxa(2000));
        let second = row(1, &[1000]);
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        for end in [first, second] {
            writer
                .push(cell(1), '\u{7}', paragraph("x"), &mut unframed, &mut budget)
                .unwrap();
            writer
                .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
                .unwrap();
        }
        let body = writer.finish(&mut budget).unwrap();
        let Block::Table(table) = &body.0[0] else {
            panic!()
        };
        let reset = table.rows[1]
            .table_row_layout
            .exception
            .as_ref()
            .unwrap()
            .preferred_width
            .as_ref()
            .unwrap();
        assert_eq!(reset.kind.as_deref(), Some("auto"));
        assert_eq!(reset.value.as_deref(), Some("0"));
    }

    #[test]
    fn zero_table_preference_is_lexical_only_but_zero_cell_preference_is_public() {
        for preferred in [PreferredWidth::Dxa(0), PreferredWidth::Percent(0)] {
            let mut end = row(1, &[1000]);
            end.row.preferred_width = Some(preferred);
            end.row.cells[0].preferred = Some(preferred);
            let mut sequence = 0;
            let mut writer = Writer::new(&mut sequence);
            let mut budget = ModelBudget::new(1_000_000);
            writer
                .push(cell(1), '\u{7}', paragraph("x"), &mut unframed, &mut budget)
                .unwrap();
            writer
                .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
                .unwrap();
            let body = writer.finish(&mut budget).unwrap();
            let Block::Table(table) = &body.0[0] else {
                panic!()
            };
            assert_eq!((table.width_pt, table.width_pct), (None, None));
            assert_eq!(
                table
                    .table_layout
                    .preferred_width
                    .as_ref()
                    .unwrap()
                    .value
                    .as_deref(),
                Some("0")
            );
            let cell = &table.rows[0].cells[0];
            match preferred {
                PreferredWidth::Dxa(_) => {
                    assert_eq!((cell.width_pt, cell.width_pct), (Some(0.0), None))
                }
                PreferredWidth::Percent(_) => {
                    assert_eq!((cell.width_pt, cell.width_pct), (None, Some(0.0)))
                }
                PreferredWidth::Auto => unreachable!(),
            }
        }
    }

    #[test]
    fn preferred_widths_follow_the_merged_leader_cell() {
        fn configured_row() -> Properties {
            let mut end = row(1, &[1000, 2000]);
            end.row.preferred_width = Some(PreferredWidth::Percent(2500));
            end.row.cells[0].preferred = Some(PreferredWidth::Dxa(720));
            end.row.cells[0].flags = (end.row.cells[0].flags & !3) | 2;
            end.row.cells[1].preferred = Some(PreferredWidth::Percent(1250));
            end.row.cells[1].flags = (end.row.cells[1].flags & !3) | 1;
            end
        }

        let mut sequence = 0;
        let mut direct_writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        direct_writer
            .push(cell(1), '\u{7}', paragraph(""), &mut unframed, &mut budget)
            .unwrap();
        direct_writer
            .push(cell(1), '\u{7}', paragraph(""), &mut unframed, &mut budget)
            .unwrap();
        direct_writer
            .push(
                configured_row(),
                '\u{7}',
                Blocks::default(),
                &mut unframed,
                &mut budget,
            )
            .unwrap();
        let direct = direct_writer.finish(&mut budget).unwrap();
        let Block::Table(direct_table) = &direct.0[0] else {
            panic!()
        };
        let direct_table = serde_json::to_value(direct_table).unwrap();
        // MS-DOC 2.9.317: the leader's formatting extends across the merged
        // set; the continuation's conflicting preference is not projected.
        for (pointer, expected) in [
            ("/widthPt", None),
            ("/widthPct", Some(serde_json::json!(2500.0))),
            (
                "/__tableLayout/preferredWidth",
                Some(serde_json::json!({"kind": "pct", "value": "2500"})),
            ),
            ("/rows/0/cells/0/widthPt", Some(serde_json::json!(36.0))),
            ("/rows/0/cells/0/widthPct", None),
            (
                "/rows/0/cells/0/__tableCellLayout/preferredWidth",
                Some(serde_json::json!({"kind": "dxa", "value": "720"})),
            ),
            ("/rows/0/cells/0/colSpan", Some(serde_json::json!(2))),
        ] {
            assert_eq!(
                direct_table.pointer(pointer),
                expected.as_ref(),
                "{pointer}"
            );
        }
        assert_eq!(direct_table["colWidths"], serde_json::json!([50.0, 100.0]));
        assert_eq!(
            direct_table["rows"][0]["cells"].as_array().unwrap().len(),
            1
        );
    }
}
