//! Direct projection for tables planned by the shared MS-DOC table grammar.

use super::{payload, ModelBudget};
use crate::doc::{
    table::{Color, PreferredWidth, Properties},
    table_structure::{Assembler, Event, LogicalTable, Payload},
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
    let table_preferred = first.preferred_width;
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
                width_pt: source.preferred.and_then(|w| match w {
                    PreferredWidth::Dxa(value) => Some(f64::from(value) / 20.0),
                    _ => None,
                }),
                width_pct: source.preferred.and_then(|w| match w {
                    PreferredWidth::Percent(value) => Some(f64::from(value)),
                    _ => None,
                }),
                margin_top: Some(f64::from(margins[0]) / 20.0),
                margin_left: Some(f64::from(margins[1]) / 20.0),
                margin_bottom: Some(f64::from(margins[2]) / 20.0),
                margin_right: Some(f64::from(margins[3]) / 20.0),
                table_cell_layout: TableCellLayoutAcquisitionWire {
                    preferred_width: source.preferred.map(width),
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
            ordinary_flow: true,
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
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
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
            .err()
            .unwrap()
            .contains("classify positioned"));
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
            .push(cell(1), '\u{7}', paragraph("a"), &mut budget)
            .unwrap();
        writer
            .push(cell(1), '\u{7}', paragraph("b"), &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut budget)
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
    fn row_nil_preference_projects_an_auto_exception_to_clear_table_width() {
        let mut first = row(1, &[1000]);
        first.row.preferred_width = Some(PreferredWidth::Dxa(2000));
        let second = row(1, &[1000]);
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        for end in [first, second] {
            writer
                .push(cell(1), '\u{7}', paragraph("x"), &mut budget)
                .unwrap();
            writer
                .push(end, '\u{7}', Blocks::default(), &mut budget)
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
                .push(cell(1), '\u{7}', paragraph("x"), &mut budget)
                .unwrap();
            writer
                .push(end, '\u{7}', Blocks::default(), &mut budget)
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
    fn direct_preferred_width_fields_match_the_byte_route_docx_parser() {
        fn configured_row() -> Properties {
            let mut end = row(1, &[1000, 2000]);
            end.row.preferred_width = Some(PreferredWidth::Percent(2500));
            end.row.cells[0].preferred = Some(PreferredWidth::Dxa(720));
            end.row.cells[0].flags = (end.row.cells[0].flags & !3) | 2;
            end.row.cells[1].preferred = Some(PreferredWidth::Percent(1250));
            end.row.cells[1].flags = (end.row.cells[1].flags & !3) | 1;
            end
        }

        let mut xml_writer = crate::doc::table_output::Writer::new(100_000);
        xml_writer.push(cell(1), '\u{7}', "<w:p/>".into()).unwrap();
        xml_writer.push(cell(1), '\u{7}', "<w:p/>".into()).unwrap();
        xml_writer
            .push(configured_row(), '\u{7}', String::new())
            .unwrap();
        let document_xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}</w:body></w:document>"#,
            xml_writer.finish().unwrap()
        );
        // MS-DOC 2.9.317: the leader's formatting extends across the merged
        // set. The continuation's conflicting preference is not serialized.
        assert_eq!(document_xml.matches("<w:tcW ").count(), 1);
        assert!(document_xml.contains("<w:tcW w:w=\"720\" w:type=\"dxa\"/>"));
        assert!(document_xml.contains("<w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"2000\"/>"));
        let mut package = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut package));
            archive
                .start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            archive.write_all(document_xml.as_bytes()).unwrap();
            archive.finish().unwrap();
        }
        let parsed: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&package).unwrap()).unwrap();
        let byte_table = &parsed["body"][0];

        let mut sequence = 0;
        let mut direct_writer = Writer::new(&mut sequence);
        let mut budget = ModelBudget::new(1_000_000);
        direct_writer
            .push(cell(1), '\u{7}', paragraph(""), &mut budget)
            .unwrap();
        direct_writer
            .push(cell(1), '\u{7}', paragraph(""), &mut budget)
            .unwrap();
        direct_writer
            .push(configured_row(), '\u{7}', Blocks::default(), &mut budget)
            .unwrap();
        let direct = direct_writer.finish(&mut budget).unwrap();
        let Block::Table(direct_table) = &direct.0[0] else {
            panic!()
        };
        let direct_table = serde_json::to_value(direct_table).unwrap();
        for pointer in [
            "/widthPt",
            "/widthPct",
            "/__tableLayout/preferredWidth",
            "/rows/0/cells/0/widthPt",
            "/rows/0/cells/0/widthPct",
            "/rows/0/cells/0/__tableCellLayout/preferredWidth",
            "/rows/0/cells/0/colSpan",
        ] {
            assert_eq!(
                direct_table.pointer(pointer),
                byte_table.pointer(pointer),
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
