//! Assemble MS-DOC 2.4.3 table/cell/row marks into ECMA-376 17.4 tables.
//! Row grids share the union of explicit cell edges, without fitted widths.
use super::{
    table::Properties,
    table_structure::{Assembler, Event, LogicalTable, Payload},
    unsupported,
};

#[derive(Default)]
struct Xml(String);
impl Payload for Xml {
    fn append<A: FnMut(usize) -> Result<(), String>>(
        &mut self,
        other: Self,
        _admit: &mut A,
    ) -> Result<(), String> {
        self.0.push_str(&other.0);
        Ok(())
    }
    fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

pub struct Writer {
    structure: Assembler<Xml>,
    charged: usize,
    limit: usize,
}

impl Writer {
    pub fn new(limit: usize) -> Self {
        Self {
            structure: Assembler::new(),
            charged: 0,
            limit,
        }
    }
    fn charge(&mut self, bytes: usize) -> Result<(), String> {
        self.charged = self
            .charged
            .checked_add(bytes)
            .filter(|n| *n <= self.limit)
            .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
        Ok(())
    }
    pub fn push(&mut self, props: Properties, mark: char, paragraph: String) -> Result<(), String> {
        let paragraph_len = paragraph.len();
        let limit = self.limit;
        let charged = &mut self.charged;
        let retained = self.structure.push(
            props,
            mark,
            Xml(paragraph),
            |event| emit(event, limit, charged),
            &mut |_| Ok(()),
        )?;
        if retained {
            self.charge(paragraph_len)?;
        }
        Ok(())
    }
    pub fn finish(mut self) -> Result<String, String> {
        let limit = self.limit;
        let charged = &mut self.charged;
        Ok(self
            .structure
            .finish(|event| emit(event, limit, charged), &mut |_| Ok(()))?
            .0)
    }
}

fn emit(event: Event<Xml>, limit: usize, charged: &mut usize) -> Result<Xml, String> {
    match event {
        Event(tables) => {
            let mut xml = String::new();
            let content = tables
                .iter()
                .flat_map(|table| table.rows.iter())
                .flat_map(|row| row.cells.iter())
                .map(|cell| cell.content.0.len())
                .try_fold(0usize, |sum, len| {
                    sum.checked_add(len).ok_or("OUTPUT_TOO_LARGE")
                })?;
            for table in tables {
                serialize(table, &mut xml, limit)?;
            }
            *charged = charged
                .checked_add(xml.len().saturating_sub(content))
                .filter(|n| *n <= limit)
                .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
            Ok(Xml(xml))
        }
    }
}

fn margins(xml: &mut String, values: [u16; 4]) {
    for (side, width) in ["top", "left", "bottom", "right"].into_iter().zip(values) {
        xml.push_str(&format!("<w:{side} w:w=\"{width}\" w:type=\"dxa\"/>"));
    }
}

fn serialize(table: LogicalTable<Xml>, xml: &mut String, limit: usize) -> Result<(), String> {
    let first = &table.rows[0].source;
    let grid = &table.grid;
    let origin = table.origin;
    let total = table.total;
    let (jc, physical) = first.alignment;
    let jc = if physical && first.bidi { 2 - jc } else { jc };
    xml.push_str("<w:tbl><w:tblPr>");
    xml.push_str(&first.position.xml());
    xml.push_str(&format!("<w:bidiVisual w:val=\"{}\"/><w:tblW w:w=\"{total}\" w:type=\"dxa\"/><w:jc w:val=\"{}\"/><w:tblInd w:w=\"{origin}\" w:type=\"dxa\"/>",u8::from(first.bidi),["left","center","right"][jc as usize]));
    if let Some(shading) = &first.shading {
        xml.push_str(&shading.xml());
    }
    xml.push_str(&format!(
        "<w:tblLayout w:type=\"{}\"/><w:tblCellMar>",
        if first.autofit { "autofit" } else { "fixed" }
    ));
    margins(xml, first.margins);
    xml.push_str("</w:tblCellMar></w:tblPr><w:tblGrid>");
    for edges in grid.windows(2) {
        xml.push_str(&format!("<w:gridCol w:w=\"{}\"/>", edges[1] - edges[0]));
    }
    xml.push_str("</w:tblGrid>");
    for (row_index, planned) in table.rows.iter().enumerate() {
        let row = &planned.source;
        xml.push_str("<w:tr>");
        if row.shading != first.shading {
            // ECMA-376 17.4.30: table-level exceptions belong to the row;
            // changing shading must not split the table's shared grid.
            xml.push_str("<w:tblPrEx>");
            if let Some(shading) = &row.shading {
                xml.push_str(&shading.xml());
            } else {
                xml.push_str("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"auto\"/>");
            }
            xml.push_str("</w:tblPrEx>");
        }
        xml.push_str("<w:trPr>");
        let before = planned.grid_before;
        if before != 0 {
            xml.push_str(&format!(
                "<w:gridBefore w:val=\"{before}\"/><w:wBefore w:w=\"{}\" w:type=\"dxa\"/>",
                planned.width_before
            ));
        }
        let after = planned.grid_after;
        if after != 0 {
            xml.push_str(&format!(
                "<w:gridAfter w:val=\"{after}\"/><w:wAfter w:w=\"{}\" w:type=\"dxa\"/>",
                planned.width_after
            ));
        }
        if row.cant_split {
            xml.push_str("<w:cantSplit/>");
        }
        if row.height != 0 {
            xml.push_str(&format!(
                "<w:trHeight w:val=\"{}\" w:hRule=\"{}\"/>",
                row.height.abs(),
                if row.height < 0 { "exact" } else { "atLeast" }
            ));
        }
        if planned.is_header {
            xml.push_str("<w:tblHeader/>");
        }
        xml.push_str("</w:trPr>");
        for planned_cell in &planned.cells {
            let c = &planned_cell.source;
            let i = planned_cell.source_index;
            let end_cell = planned_cell.source_end;
            let width = planned_cell.width;
            let span = planned_cell.grid_span;
            xml.push_str(&format!(
                "<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/>"
            ));
            if span > 1 {
                xml.push_str(&format!("<w:gridSpan w:val=\"{span}\"/>"));
            }
            let vertical = planned_cell.vertical;
            if vertical == 1 || vertical == 3 {
                xml.push_str(&format!(
                    "<w:vMerge w:val=\"{}\"/>",
                    if vertical == 1 { "continue" } else { "restart" }
                ));
            }
            xml.push_str("<w:tcBorders>");
            for (side, name) in ["top", "left", "bottom", "right", "tl2br", "tr2bl"]
                .into_iter()
                .enumerate()
            {
                let fallback = match side {
                    0 => Some(if row_index == 0 { 0 } else { 4 }),
                    1 => Some(if i == 0 { 1 } else { 5 }),
                    2 => Some(if row_index + 1 == table.rows.len() {
                        2
                    } else {
                        4
                    }),
                    3 => Some(if end_cell == planned.source_cell_count {
                        3
                    } else {
                        5
                    }),
                    _ => None,
                };
                if let Some(b) = c.borders[side]
                    .as_ref()
                    .or_else(|| fallback.and_then(|s| row.borders[s].as_ref()))
                {
                    xml.push_str(&b.xml(name));
                }
            }
            xml.push_str("</w:tcBorders>");
            if let Some(shading) = &c.shading {
                xml.push_str(&shading.xml());
            }
            xml.push_str("<w:tcMar>");
            margins(
                xml,
                std::array::from_fn(|s| c.margins[s].unwrap_or(row.margins[s])),
            );
            xml.push_str("</w:tcMar>");
            if c.flags & (1 << 12) != 0 {
                xml.push_str("<w:tcFitText/>");
            }
            let align = (c.flags >> 7) & 3;
            if align > 2 {
                return Err(unsupported("invalid Word vertical cell alignment"));
            }
            xml.push_str(&format!(
                "<w:vAlign w:val=\"{}\"/>",
                ["top", "center", "bottom"][align as usize]
            ));
            if c.flags & (1 << 14) != 0 {
                xml.push_str("<w:hideMark/>");
            }
            xml.push_str("</w:tcPr>");
            if vertical == 1 {
                xml.push_str("<w:p/>");
            } else {
                xml.push_str(&planned_cell.content.0);
            }
            xml.push_str("</w:tc>");
            if xml.len() > limit {
                return Err("OUTPUT_TOO_LARGE".into());
            }
        }
        xml.push_str("</w:tr>");
    }
    xml.push_str("</w:tbl>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn floating_table_position_precedes_direction_without_rewriting_the_grid() {
        let mut r = row(1, &[1000]);
        r.row.apply(0x360d, &[0x20]).unwrap(); // Column/text anchors.
        r.row.apply(0x940e, &721i16.to_le_bytes()).unwrap();
        r.row.apply(0x940f, &1i16.to_le_bytes()).unwrap();
        r.row.apply(0x941e, &187u16.to_le_bytes()).unwrap();
        r.row.apply(0x3465, &[1]).unwrap();
        let mut w = Writer::new(100000);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        w.push(r, '\u{7}', String::new()).unwrap();
        let xml = w.finish().unwrap();
        assert!(xml.contains("<w:tblPr><w:tblpPr "));
        assert!(xml.contains("w:horzAnchor=\"text\" w:vertAnchor=\"text\""));
        assert!(xml.contains("w:tblpX=\"720\" w:tblpY=\"0\""));
        assert!(xml.contains("w:rightFromText=\"187\""));
        assert!(xml.contains("<w:tblOverlap w:val=\"never\"/><w:bidiVisual"));
        assert!(xml.contains("<w:gridCol w:w=\"1000\"/>"));
    }

    #[test]
    fn cell_shading_is_retained_between_borders_and_margins() {
        let mut r = row(1, &[1000]);
        assert!(r
            .row
            .apply(0xd612, &[10, 0, 0, 0, 255, 0x12, 0x34, 0x56, 0, 0, 0])
            .unwrap());
        let mut w = Writer::new(100000);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        w.push(r, '\u{7}', String::new()).unwrap();
        let xml = w.finish().unwrap();
        assert!(xml.contains(
            "</w:tcBorders><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"123456\"/><w:tcMar>"
        ));
    }

    #[test]
    fn table_shading_does_not_become_cell_override() {
        let mut r = row(1, &[1000]);
        assert!(r
            .row
            .apply(0xd660, &[10, 0, 0, 0, 255, 0x12, 0x34, 0x56, 0, 0, 0])
            .unwrap());
        let mut w = Writer::new(100000);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        w.push(r, '\u{7}', String::new()).unwrap();
        let xml = w.finish().unwrap();
        let shade = xml.find("<w:shd ").unwrap();
        assert!(shade < xml.find("<w:tblLayout ").unwrap());
        assert_eq!(xml.matches("<w:shd ").count(), 1);
    }

    #[test]
    fn varying_table_shading_uses_row_exceptions_without_splitting_the_grid() {
        for first_shaded in [false, true] {
            let mut w = Writer::new(100000);
            for shaded in [first_shaded, !first_shaded, first_shaded] {
                let mut r = row(1, &[1000]);
                if shaded {
                    r.row
                        .apply(0xd660, &[10, 0, 0, 0, 255, 0x12, 0x34, 0x56, 0, 0, 0])
                        .unwrap();
                }
                w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
                w.push(r, '\u{7}', String::new()).unwrap();
            }
            let xml = w.finish().unwrap();
            assert_eq!(xml.matches("<w:tbl>").count(), 1);
            assert_eq!(xml.matches("<w:tr>").count(), 3);
            assert_eq!(xml.matches("<w:tblPrEx>").count(), 1);
            let fill = if first_shaded { "auto" } else { "123456" };
            assert!(xml.contains(&format!("<w:tr><w:tblPrEx><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{fill}\"/></w:tblPrEx><w:trPr>")));
        }
    }

    fn cell(depth: u32, end: bool) -> Properties {
        let mut p = Properties::default();
        p.apply(0x6649, &depth.to_le_bytes()).unwrap();
        p.inner_cell = end;
        p
    }
    fn row(depth: u32, widths: &[u16]) -> Properties {
        let mut p = cell(depth, false);
        p.row_end = true;
        p.inner_row = true;
        for (i, w) in widths.iter().enumerate() {
            let [a, b] = w.to_le_bytes();
            p.row.apply(0x7621, &[i as u8, 1, a, b]).unwrap();
        }
        p
    }
    #[test]
    fn consumes_row_marks_and_builds_shared_grid_without_extra_paragraphs() {
        let mut w = Writer::new(100000);
        w.push(cell(1, false), '\u{7}', "<w:p>A</w:p>".into())
            .unwrap();
        w.push(cell(1, false), '\u{7}', "<w:p>B</w:p>".into())
            .unwrap();
        w.push(row(1, &[1000, 2000]), '\u{7}', "<w:p>MARK</w:p>".into())
            .unwrap();
        w.push(cell(1, false), '\u{7}', "<w:p>C</w:p>".into())
            .unwrap();
        w.push(row(1, &[3000]), '\u{7}', String::new()).unwrap();
        let xml = w.finish().unwrap();
        assert_eq!(xml.matches("<w:p>").count(), 3);
        assert!(!xml.contains("MARK"));
        assert!(xml.contains("<w:gridSpan w:val=\"2\"/>"));
    }
    #[test]
    fn nested_table_stays_inside_parent_cell() {
        let mut w = Writer::new(100000);
        w.push(cell(2, true), '\r', "<w:p>nested</w:p>".into())
            .unwrap();
        w.push(row(2, &[500]), '\r', String::new()).unwrap();
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        w.push(row(1, &[1000]), '\u{7}', String::new()).unwrap();
        let xml = w.finish().unwrap();
        assert_eq!(xml.matches("<w:tbl>").count(), 2);
        assert!(xml.contains("</w:tbl><w:p/></w:tc>"));
    }
    #[test]
    fn zero_width_cells_keep_their_grid_slots() {
        let mut w = Writer::new(100000);
        for _ in 0..3 {
            w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        }
        w.push(row(1, &[0, 1000, 0]), '\u{7}', String::new())
            .unwrap();
        let xml = w.finish().unwrap();
        assert_eq!(xml.matches("<w:gridCol w:w=\"0\"/>").count(), 2);
        assert_eq!(xml.matches("<w:tc>").count(), 3);
    }
    #[test]
    fn rejects_unclosed_rows_and_output_expansion_before_success() {
        let mut w = Writer::new(100000);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        assert!(w.finish().is_err());
        let mut w = Writer::new(10);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        w.push(row(1, &[1000]), '\u{7}', String::new()).unwrap();
        assert!(w.finish().unwrap_err().contains("OUTPUT_TOO_LARGE"));
        let mut w = Writer::new(100000);
        w.structure.set_row_count(100000);
        w.push(cell(1, false), '\u{7}', "<w:p/>".into()).unwrap();
        assert!(w
            .push(row(1, &[1000]), '\u{7}', String::new())
            .unwrap_err()
            .contains("row budget"));
    }
}
