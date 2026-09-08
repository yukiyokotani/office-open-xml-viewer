//! Direct projection for tables planned by the shared MS-DOC table grammar.

use super::{payload, ModelBudget};
use crate::doc::{
    table::{Color, Properties},
    table_structure::{Assembler, Event, LogicalTable, Payload},
    unsupported,
};
use docx_model::{
    CellBorders, CellElement, DocParagraph, DocTable, DocTableCell, DocTableRow,
    TableBorders, TableCellLayoutAcquisitionWire, TableGridAcquisitionWire,
    TableGridColumnAcquisitionWire, TableLayoutAcquisitionWire, TableLayoutKindAcquisitionWire,
    TableMarginAcquisitionWire, TablePropertyExceptionAcquisitionWire,
    TableRowHeightAcquisitionWire, TableRowLayoutAcquisitionWire, TableWidthAcquisitionWire,
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
}

impl<'a> Writer<'a> {
    pub(super) fn new(sequence: &'a mut usize) -> Self {
        Self {
            structure: Assembler::new(),
            sequence,
        }
    }

    pub(super) fn push(
        &mut self,
        props: Properties,
        mark: char,
        paragraph: Blocks,
        budget: &mut ModelBudget,
    ) -> Result<(), String> {
        let remaining = std::cell::Cell::new(budget.remaining_bytes);
        let sequence = &mut *self.sequence;
        self.structure.push(
            props,
            mark,
            paragraph,
            |Event(plans)| project_tables(plans, sequence, &remaining),
            &mut |bytes| charge_cell(&remaining, bytes),
        )?;
        budget.remaining_bytes = remaining.get();
        Ok(())
    }

    pub(super) fn finish(self, budget: &mut ModelBudget) -> Result<Blocks, String> {
        let remaining = std::cell::Cell::new(budget.remaining_bytes);
        let blocks = self.structure.finish(
            |Event(plans)| project_tables(plans, self.sequence, &remaining),
            &mut |bytes| charge_cell(&remaining, bytes),
        )?;
        budget.remaining_bytes = remaining.get();
        Ok(blocks)
    }
}

fn project_tables(
    plans: Vec<LogicalTable<Blocks>>,
    sequence: &mut usize,
    remaining: &std::cell::Cell<usize>,
) -> Result<Blocks, String> {
    let mut output = Blocks::default();
    for plan in plans {
        let table = project_table(plan, *sequence, remaining)?;
        *sequence = sequence.checked_add(1).ok_or("OUTPUT_TOO_LARGE")?;
        reserve(&mut output.0, 1, &mut |n| charge_cell(remaining, n))?;
        output.0.push(Block::Table(Box::new(table)));
    }
    Ok(output)
}

fn project_table(
    plan: LogicalTable<Blocks>,
    sequence: usize,
    remaining: &std::cell::Cell<usize>,
) -> Result<DocTable, String> {
    let first = &plan.rows[0].source;
    if first.shading.is_some() {
        return Err(unsupported(
            "direct DOC model cannot retain table-level shading",
        ));
    }
    let (tblp_pr, overlap) = first.position.direct();
    if tblp_pr.is_some() {
        return Err(unsupported(
            "direct DOC model cannot yet classify positioned table flow",
        ));
    }
    let (alignment, physical) = first.alignment;
    let first_bidi = first.bidi;
    let first_autofit = first.autofit;
    let margins = first.margins;
    let alignment = if physical && first_bidi {
        2 - alignment
    } else {
        alignment
    };
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
        let mut cells = Vec::new();
        reserve(&mut cells, planned.cells.len(), &mut |n| {
            charge_cell(remaining, n)
        })?;
        for cell in planned.cells {
            let source = &cell.source;
            let facts = source.shading.as_ref().map(|s| s.direct_facts());
            let background = match facts {
                None => None,
                Some(f) if matches!(f.pattern, "clear" | "nil") => {
                    source.shading.as_ref().and_then(|s| s.direct_background())
                }
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
            if source.flags & ((1 << 12) | (1 << 14)) != 0
                || source.borders[4].is_some()
                || source.borders[5].is_some()
            {
                return Err(unsupported(
                    "direct DOC model cannot retain cell fit/hide/diagonal facts",
                ));
            }
            let align = (source.flags >> 7) & 3;
            if align > 2 {
                return Err(unsupported("invalid Word vertical cell alignment"));
            }
            let mut content = Vec::new();
            reserve(&mut content, cell.content.0.len().max(1), &mut |n| {
                charge_cell(remaining, n)
            })?;
            if cell.vertical == 1 {
                charge_cell(remaining, std::mem::size_of::<DocParagraph>())?;
                content.push(CellElement::Paragraph(Box::new(DocParagraph::default())));
            } else {
                for block in cell.content.0 {
                    content.push(match block {
                        Block::Paragraph(value) => CellElement::Paragraph(value),
                        Block::Table(value) => CellElement::Table(value),
                        Block::PageBreak { .. } | Block::ColumnBreak => {
                            return Err(unsupported("table cell promoted a paragraph break"))
                        }
                    });
                }
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
                },
                background,
                v_align: ["top", "center", "bottom"][align as usize].into(),
                width_pt: Some(f64::from(cell.width) / 20.0),
                width_pct: None,
                margin_top: Some(f64::from(margins[0]) / 20.0),
                margin_left: Some(f64::from(margins[1]) / 20.0),
                margin_bottom: Some(f64::from(margins[2]) / 20.0),
                margin_right: Some(f64::from(margins[3]) / 20.0),
                table_cell_layout: TableCellLayoutAcquisitionWire {
                    preferred_width: Some(width(cell.width)),
                    margins: Some(margin_wire(margins)),
                },
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
                before_width: (planned.grid_before != 0).then(|| width(planned.width_before)),
                after_width: (planned.grid_after != 0).then(|| width(planned.width_after)),
                justification: None,
                cell_spacing: None,
                style_cell_spacing: None,
                style_cell_margins: None,
                exception: None::<TablePropertyExceptionAcquisitionWire>,
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
        width_pt: (plan.total > 0).then(|| f64::from(plan.total) / 20.0),
        width_pct: None,
        bidi_visual: Some(first_bidi),
        tblp_pr,
        overlap,
        table_layout: TableLayoutAcquisitionWire {
            effective_style_id: None,
            ordinary_flow: true,
            logical_sequence_id: format!("legacy-doc/table/{sequence}"),
            logical_row_offset: 0,
            logical_total_rows: row_count,
            grid: TableGridAcquisitionWire {
                authored: true,
                columns: grid_columns,
                required_column_count: (plan.grid.len() - 1) as u32,
            },
            preferred_width: Some(width(plan.total)),
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

fn width(value: i32) -> TableWidthAcquisitionWire {
    TableWidthAcquisitionWire {
        kind: Some("dxa".into()),
        value: Some(value.to_string()),
    }
}
fn margin_wire(v: [u16; 4]) -> TableMarginAcquisitionWire {
    TableMarginAcquisitionWire {
        top: Some(width(v[0].into())),
        left: Some(width(v[1].into())),
        bottom: Some(width(v[2].into())),
        right: Some(width(v[3].into())),
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
    fn plain_table_preserves_zero_grid_slots_merges_and_cell_break_runs() {
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut budget)
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("b"), &mut budget)
            .unwrap();
        writer
            .push(row(1, &[0, 1000]), '\u{7}', Blocks::default(), &mut budget)
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
    fn no_overlap_alone_remains_ordinary_but_positioned_table_fails_closed() {
        let mut end = row(1, &[1000]);
        end.row.apply(0x3465, &[1]).unwrap();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        writer
            .push(cell(1), '\u{7}', paragraph("a"), &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut budget)
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
            .push(cell(1), '\u{7}', paragraph("a"), &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut budget)
            .unwrap();
        assert!(writer
            .finish(&mut budget)
            .err().unwrap()
            .contains("classify positioned"));
    }
}
