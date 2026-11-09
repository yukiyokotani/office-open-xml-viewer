//! Tables (ODF Part 3 §9.1; table, column, row and cell properties
//! §17.15-§17.18), measured against LibreOffice 26.8 PDF export.
//!
//! A table maps to a fixed-layout DOCX table with the authored column grid.
//! LibreOffice collapses borders: each cell edge is drawn once, centred on
//! the column edge, a cell's text starts at its edge plus its padding, and
//! the table's top and bottom borders lie inside the table. Spanned cells
//! (`table:number-columns-spanned`, `-rows-spanned`) become DOCX grid spans
//! and vertical merges; covered cells (§9.1.5) are the merged positions.

use super::*;
use docx_model::{
    BorderSpec, CellBorders, DocTableCell, DocTableRow, TableBorders,
    TableCellLayoutAcquisitionWire, TableGridAcquisitionWire, TableGridColumnAcquisitionWire,
    TableLayoutAcquisitionWire, TableLayoutKindAcquisitionWire, TableMarginAcquisitionWire,
    TableWidthAcquisitionWire,
};

/// Implementation bounds on one table's grid.
const MAX_COLUMNS: usize = 1024;
const MAX_ROWS: usize = 100_000;

/// A property group's attributes.
type Attributes = HashMap<(String, String), String>;

/// LibreOffice's fit of absolute column widths `a` (twips) to a left-aligned
/// table `total` twips wide (`SwXMLTableContext::MakeTable_`, LibreOffice
/// 26.8): a shortfall is shared in proportion to the widths; an excess
/// gives every column MINLAY (23) and shares the rest in proportion to the
/// full widths; the last column takes what remains.
pub(super) fn redistribute(a: &[i64], total: i64) -> Vec<i64> {
    const MINLAY: i64 = 23;
    let sum: i64 = a.iter().sum::<i64>().max(1);
    let (base, spare): (Vec<i64>, i64) = if sum < total {
        (a.to_vec(), total - sum)
    } else {
        (vec![MINLAY; a.len()], total - MINLAY * a.len() as i64)
    };
    let mut out = base.clone();
    let mut last = base[a.len() - 1] + spare;
    for i in 0..a.len() - 1 {
        let extra = a[i] * spare / sum;
        out[i] += extra;
        last -= extra;
    }
    out[a.len() - 1] = last;
    out
}

fn twips(points: f64) -> String {
    ((points * 20.0).round() as i64).to_string()
}

fn dxa(points: f64) -> TableWidthAcquisitionWire {
    TableWidthAcquisitionWire {
        kind: Some("dxa".into()),
        value: Some(twips(points)),
    }
}

/// A cell border (§20.183) in DOCX terms. LibreOffice holds border widths in
/// whole twips, rounded half up (measured: 0.02 pt draws nothing).
fn cell_border(value: &str, line_width: Option<&str>) -> Option<BorderSpec> {
    let spec = parse_border(value, line_width)?;
    let twips = (spec.width * 20.0 + 0.5).floor();
    if twips <= 0.0 {
        return None;
    }
    Some(BorderSpec {
        width: twips / 20.0,
        color: spec.color,
        style: spec.style.to_string(),
    })
}

/// A spanned cell still covering later rows.
#[derive(Clone)]
struct VerticalSpan {
    remaining: u32,
    columns: u32,
    borders: CellBorders,
    background: Option<String>,
    width: f64,
}

impl<'a> Reader<'a> {
    /// The attributes of one property group of a table-family style,
    /// following its parents (§16.2).
    fn table_group(
        &mut self,
        family: Family,
        name: Option<&str>,
        group: Group,
    ) -> Result<Attributes, String> {
        let Some(style) = name.and_then(|name| self.sheet.find(self.scope, family, name)) else {
            let mut out = Attributes::new();
            if let Some(default) = self.sheet.default_style(family) {
                if let Some(properties) = default.group(group) {
                    out.extend(properties.attributes.clone());
                }
            }
            return Ok(out);
        };
        let mut chain = self.style_chain(style).to_vec();
        if let Some(default) = self.sheet.default_style(family) {
            chain.push(default);
        }
        self.charge_cascade(&chain, group)?;
        let mut out = Attributes::new();
        for style in chain.iter().rev() {
            if let Some(properties) = style.group(group) {
                for (key, value) in &properties.attributes {
                    out.insert(key.clone(), value.clone());
                }
            }
        }
        Ok(out)
    }

    /// A `table:table` (§9.1.2).
    pub(super) fn table(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        self.enter()?;
        let result = self.table_inner(node);
        self.leave();
        result
    }

    fn table_inner(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        self.charge_construction(std::mem::size_of::<DocTable>())?;
        let properties = self.table_group(
            Family::Table,
            node.attribute((NS_TABLE, "style-name")),
            "table-properties",
        )?;
        let get = |ns: &str, name: &str| {
            properties
                .get(&(ns.to_string(), name.to_string()))
                .map(String::as_str)
        };
        let length = |ns: &str, name: &str| get(ns, name).and_then(|v| length_pt(v).ok());
        let text_width = self.text_width - self.section_indent.0 - self.section_indent.1;
        let margin_left = length(NS_FO, "margin-left").unwrap_or(0.0);
        let margin_right = length(NS_FO, "margin-right").unwrap_or(0.0);
        let top = length(NS_FO, "margin-top").unwrap_or(0.0);
        let bottom = length(NS_FO, "margin-bottom").unwrap_or(0.0);
        // §20.234 table:align: "margins" fills the area between the margins.
        let align = get(NS_TABLE, "align").unwrap_or("margins");
        let authored_width = length(NS_STYLE, "width");
        let relative = get(NS_STYLE, "rel-width").and_then(|v| percent(v).ok());
        let table_width = match align {
            "margins" => text_width - margin_left - margin_right,
            _ => match (relative, authored_width) {
                (Some(r), _) => r * text_width,
                (None, Some(w)) => w,
                (None, None) => text_width - margin_left - margin_right,
            },
        }
        .max(1.0);
        let jc = match align {
            "center" => "center",
            "right" => "right",
            _ => "left",
        };
        let table_background = get(NS_FO, "background-color")
            .filter(|v| *v != "transparent")
            .and_then(|v| hex_color(v).ok());

        // Columns (§9.1.6), possibly grouped (§9.1.10-§9.1.12), and rows
        // (§9.1.3), possibly grouped or repeated as headers.
        let mut columns: Vec<(Option<f64>, Option<f64>, Option<String>)> = Vec::new();
        let mut rows: Vec<(roxmltree::Node<'a, 'a>, bool)> = Vec::new();
        self.collect_table(node, &mut columns, &mut rows, false)?;
        if rows.is_empty() {
            return Ok(());
        }
        // Rows may hold more cells than the declared columns.
        let row_width = |row: roxmltree::Node<'_, '_>| -> usize {
            row.children()
                .filter(|n| {
                    xml::is(*n, NS_TABLE, "table-cell")
                        || xml::is(*n, NS_TABLE, "covered-table-cell")
                })
                .map(|cell| {
                    cell.attribute((NS_TABLE, "number-columns-repeated"))
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1)
                        .min(MAX_COLUMNS)
                })
                .sum()
        };
        let needed = rows
            .iter()
            .map(|(row, _)| row_width(*row))
            .max()
            .unwrap_or(0)
            .min(MAX_COLUMNS);
        while columns.len() < needed {
            columns.push((None, None, None));
        }
        if columns.is_empty() {
            return Ok(());
        }
        let col_widths = column_widths(&columns, table_width, jc);
        let total: f64 = col_widths.iter().sum();
        let table_width = if total > 0.0 { total } else { table_width };

        let mut vertical: Vec<Option<VerticalSpan>> = vec![None; col_widths.len()];
        let mut doc_rows = Vec::with_capacity(rows.len());
        let row_count = rows.len();
        for (row, header) in rows {
            let row_properties = self.table_group(
                Family::TableRow,
                row.attribute((NS_TABLE, "style-name")),
                "table-row-properties",
            )?;
            let row_get = |name: (&str, &str)| {
                row_properties
                    .get(&(name.0.to_string(), name.1.to_string()))
                    .map(String::as_str)
            };
            let (row_height, rule) = match (
                row_get((NS_STYLE, "row-height")).and_then(|v| length_pt(v).ok()),
                row_get((NS_STYLE, "min-row-height")).and_then(|v| length_pt(v).ok()),
            ) {
                (Some(height), _) if height > 0.0 => (Some(height), "exact"),
                (_, Some(height)) if height > 0.0 => (Some(height), "atLeast"),
                _ => (None, "auto"),
            };
            let row_background = row_get((NS_FO, "background-color"))
                .filter(|v| *v != "transparent")
                .and_then(|v| hex_color(v).ok())
                .or_else(|| table_background.clone());
            let cant_split = row_get((NS_FO, "keep-together")) == Some("always");
            let mut cells: Vec<DocTableCell> = Vec::new();
            let mut column = 0usize;
            let mut skip = 0u32;
            let mut cell_nodes = Vec::new();
            for cell in row.children().filter(|n| n.is_element()) {
                let covered = xml::is(cell, NS_TABLE, "covered-table-cell");
                if !covered && !xml::is(cell, NS_TABLE, "table-cell") {
                    continue;
                }
                let repeat = cell
                    .attribute((NS_TABLE, "number-columns-repeated"))
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
                    .clamp(1, MAX_COLUMNS);
                for _ in 0..repeat {
                    cell_nodes.push((cell, covered));
                }
            }
            for (cell, covered) in cell_nodes {
                if column >= col_widths.len() {
                    break;
                }
                if covered {
                    if skip > 0 {
                        skip -= 1;
                        column += 1;
                        continue;
                    }
                    if let Some(span) = vertical[column].as_mut().filter(|s| s.remaining > 0) {
                        span.remaining -= 1;
                        let continuation = DocTableCell {
                            content: vec![CellElement::Paragraph(Box::default())],
                            col_span: span.columns,
                            v_merge: Some(false),
                            borders: span.borders.clone(),
                            background: span.background.clone(),
                            v_align: "top".into(),
                            width_pt: Some(span.width),
                            ..DocTableCell::default()
                        };
                        skip = span.columns - 1;
                        cells.push(continuation);
                        column += 1;
                        continue;
                    }
                    // A covered cell outside any span: an empty cell.
                    cells.push(DocTableCell {
                        content: vec![CellElement::Paragraph(Box::default())],
                        col_span: 1,
                        v_align: "top".into(),
                        width_pt: Some(col_widths[column]),
                        ..DocTableCell::default()
                    });
                    column += 1;
                    continue;
                }
                let span_columns = cell
                    .attribute((NS_TABLE, "number-columns-spanned"))
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
                    .clamp(1, col_widths.len() - column);
                let span_rows = cell
                    .attribute((NS_TABLE, "number-rows-spanned"))
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(1)
                    .max(1);
                let width: f64 = col_widths[column..column + span_columns].iter().sum();
                let mut doc_cell = self.table_cell(cell, width, row_background.clone())?;
                doc_cell.col_span = span_columns as u32;
                if span_rows > 1 {
                    doc_cell.v_merge = Some(true);
                    vertical[column] = Some(VerticalSpan {
                        remaining: span_rows - 1,
                        columns: span_columns as u32,
                        borders: doc_cell.borders.clone(),
                        background: doc_cell.background.clone(),
                        width,
                    });
                } else {
                    vertical[column] = None;
                }
                skip = span_columns as u32 - 1;
                cells.push(doc_cell);
                column += 1;
            }
            // Missing cells at the row's end.
            while column < col_widths.len() {
                if skip > 0 {
                    skip -= 1;
                    column += 1;
                    continue;
                }
                cells.push(DocTableCell {
                    content: vec![CellElement::Paragraph(Box::default())],
                    col_span: 1,
                    v_align: "top".into(),
                    width_pt: Some(col_widths[column]),
                    ..DocTableCell::default()
                });
                column += 1;
            }
            self.charge_construction(std::mem::size_of::<DocTableRow>())?;
            doc_rows.push(DocTableRow {
                cells,
                row_height,
                row_height_rule: rule.into(),
                is_header: header,
                cant_split,
                // LibreOffice cuts rows at a page its own way (see the flag).
                cuts_in_place: true,
                ..DocTableRow::default()
            });
        }
        // Where a page break falls between two rows, LibreOffice draws their
        // shared edge again above the row continuing on the next page, and
        // below the row ending the page only where that row has a bottom
        // rule of its own (measured). The lower cell carries the wider rule,
        // as the DOCX layout repeats a row's top rule there; without a break
        // the layout draws the shared edge once at the wider width either
        // way.
        for index in 1..doc_rows.len() {
            let (above, below) = doc_rows.split_at_mut(index);
            if above[index - 1].cells.len() != below[0].cells.len() {
                continue;
            }
            for (upper, lower) in above[index - 1].cells.iter().zip(&mut below[0].cells) {
                let width = |side: &Option<BorderSpec>| side.as_ref().map_or(0.0, |b| b.width);
                if width(&upper.borders.bottom) > width(&lower.borders.top) {
                    lower.borders.top = upper.borders.bottom.clone();
                }
            }
        }
        self.next_table += 1;
        let tbl_ind = match jc {
            "left" => Some(margin_left + self.section_indent.0),
            _ => None,
        };
        let table = DocTable {
            col_widths: col_widths.clone(),
            rows: doc_rows,
            borders: TableBorders::default(),
            jc: jc.into(),
            tbl_ind,
            layout: Some("fixed".into()),
            width_pt: Some(table_width),
            table_layout: TableLayoutAcquisitionWire {
                ordinary_flow: true,
                outer_edge_indent: true,
                logical_sequence_id: format!("odt/table/{}", self.next_table),
                logical_total_rows: row_count,
                grid: TableGridAcquisitionWire {
                    authored: true,
                    columns: col_widths
                        .iter()
                        .map(|w| TableGridColumnAcquisitionWire {
                            width: Some(twips(*w)),
                        })
                        .collect(),
                    required_column_count: col_widths.len() as u32,
                },
                preferred_width: Some(dxa(table_width)),
                layout: Some(TableLayoutKindAcquisitionWire {
                    kind: Some("fixed".into()),
                }),
                ..TableLayoutAcquisitionWire::default()
            },
            ..DocTable::default()
        };
        let break_before = get(NS_FO, "break-before") == Some("page") && self.story == Story::Body;
        let break_after = get(NS_FO, "break-after") == Some("page") && self.story == Story::Body;
        // A page style named on the table starts a page with it.
        if self.story == Story::Body {
            if let Some(master) = node
                .attribute((NS_TABLE, "style-name"))
                .and_then(|name| self.sheet.find(self.scope, Family::Table, name))
                .and_then(|style| {
                    style
                        .style()
                        .attributes
                        .get(&(NS_STYLE.to_string(), "master-page-name".to_string()))
                })
                .filter(|name| !name.is_empty() && self.pages.has_master(name))
            {
                if self.has_body_content() {
                    let kind = self.pages.start_kind(master);
                    self.current_master = master.to_string();
                    self.start_section(page::SectionStart {
                        master: master.to_string(),
                        kind,
                        columns: self.pages.columns(master),
                        page_start: None,
                        block: 0,
                    });
                } else {
                    // The first block names the initial page style.
                    self.current_master = master.to_string();
                    if let Some(first) = self.sections.last_mut() {
                        first.master = master.to_string();
                        first.columns = self.pages.columns(master);
                    }
                }
                self.apply_page_geometry();
            }
        }
        self.blocks.push(Block::Table(Box::new(TableBlock {
            table,
            top,
            bottom,
            page_break_before: break_before,
            page_break_after: break_after,
        })));
        Ok(())
    }

    /// Columns and rows of a table, flattening column and row groups.
    #[allow(clippy::type_complexity)]
    fn collect_table(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        columns: &mut Vec<(Option<f64>, Option<f64>, Option<String>)>,
        rows: &mut Vec<(roxmltree::Node<'a, 'a>, bool)>,
        header: bool,
    ) -> Result<(), String> {
        for child in node.children().filter(|n| n.is_element()) {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TABLE), "table-column") => {
                    let repeat = child
                        .attribute((NS_TABLE, "number-columns-repeated"))
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1)
                        .clamp(1, MAX_COLUMNS);
                    let properties = self.table_group(
                        Family::TableColumn,
                        child.attribute((NS_TABLE, "style-name")),
                        "table-column-properties",
                    )?;
                    let absolute = properties
                        .get(&(NS_STYLE.to_string(), "column-width".to_string()))
                        .and_then(|v| length_pt(v).ok());
                    let share = properties
                        .get(&(NS_STYLE.to_string(), "rel-column-width".to_string()))
                        .and_then(|v| v.trim_end_matches('*').parse::<f64>().ok())
                        .filter(|v| *v > 0.0);
                    if columns.len() + repeat > MAX_COLUMNS {
                        return Err("OpenDocument table exceeds the column limit".into());
                    }
                    columns.extend(std::iter::repeat_n((absolute, share, None), repeat));
                }
                (
                    Some(NS_TABLE),
                    "table-columns" | "table-header-columns" | "table-column-group",
                ) => self.collect_table(child, columns, rows, header)?,
                (Some(NS_TABLE), "table-header-rows") => {
                    self.collect_table(child, columns, rows, true)?
                }
                (Some(NS_TABLE), "table-rows" | "table-row-group") => {
                    self.collect_table(child, columns, rows, header)?
                }
                (Some(NS_TABLE), "table-row") => {
                    let repeat = child
                        .attribute((NS_TABLE, "number-rows-repeated"))
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1)
                        .clamp(1, MAX_ROWS);
                    if rows.len() + repeat > MAX_ROWS {
                        return Err("OpenDocument table exceeds the row limit".into());
                    }
                    rows.extend(std::iter::repeat_n((child, header), repeat));
                }
                _ => {}
            }
        }
        Ok(())
    }

    /// One `table:table-cell` (§9.1.4) and its content.
    fn table_cell(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        width: f64,
        row_background: Option<String>,
    ) -> Result<DocTableCell, String> {
        let properties = self.table_group(
            Family::TableCell,
            node.attribute((NS_TABLE, "style-name")),
            "table-cell-properties",
        )?;
        let get = |ns: &str, name: &str| {
            properties
                .get(&(ns.to_string(), name.to_string()))
                .map(String::as_str)
        };
        let padding_all = get(NS_FO, "padding").and_then(|v| length_pt(v).ok());
        let padding = |side: &str| {
            get(NS_FO, &format!("padding-{side}"))
                .and_then(|v| length_pt(v).ok())
                .or(padding_all)
                .unwrap_or(0.0)
                .max(0.0)
        };
        let margins = [
            padding("top"),
            padding("left"),
            padding("bottom"),
            padding("right"),
        ];
        let border_all = get(NS_FO, "border");
        let line_all = get(NS_STYLE, "border-line-width");
        let side = |name: &str| {
            get(NS_FO, &format!("border-{name}"))
                .or(border_all)
                .and_then(|value| {
                    cell_border(
                        value,
                        get(NS_STYLE, &format!("border-line-width-{name}")).or(line_all),
                    )
                })
        };
        let background = get(NS_FO, "background-color")
            .filter(|v| *v != "transparent")
            .and_then(|v| hex_color(v).ok())
            .or(row_background);
        let v_align = match get(NS_STYLE, "vertical-align") {
            Some("middle") => "center",
            Some("bottom") => "bottom",
            _ => "top",
        };
        let text_direction = match (
            get(NS_STYLE, "writing-mode"),
            get(NS_STYLE, "rotation-angle"),
        ) {
            (Some("tb-rl" | "tb"), _) => Some("tbRl".to_string()),
            (_, Some("90" | "90deg")) => Some("btLr".to_string()),
            (_, Some("270" | "270deg")) => Some("tbRl".to_string()),
            _ => None,
        };
        // The cell's content in its own story.
        let inner = (width - margins[1] - margins[3]).max(1.0);
        let elements = self.contained_story(node, inner)?;
        let mut content = self.finish_cell(elements);
        for element in &mut content {
            if let CellElement::Paragraph(paragraph) = element {
                // LibreOffice cuts a cell paragraph at a page without its
                // widow/orphan or keep-together controls (measured).
                paragraph.widow_control = false;
                paragraph.keep_lines = false;
            }
        }
        if content.is_empty() {
            content.push(CellElement::Paragraph(Box::default()));
        }
        let margin_wire = TableMarginAcquisitionWire {
            top: Some(dxa(margins[0])),
            left: Some(dxa(margins[1])),
            bottom: Some(dxa(margins[2])),
            right: Some(dxa(margins[3])),
            start: None,
            end: None,
        };
        Ok(DocTableCell {
            content,
            col_span: 1,
            borders: CellBorders {
                top: side("top"),
                left: side("left"),
                bottom: side("bottom"),
                right: side("right"),
                ..CellBorders::default()
            },
            background,
            v_align: v_align.into(),
            width_pt: Some(width),
            margin_top: Some(margins[0]),
            margin_left: Some(margins[1]),
            margin_bottom: Some(margins[2]),
            margin_right: Some(margins[3]),
            table_cell_layout: TableCellLayoutAcquisitionWire {
                preferred_width: Some(dxa(width)),
                margins: Some(margin_wire),
            },
            text_direction,
            ..DocTableCell::default()
        })
    }
}

/// Column widths (pt) for a table `width` wide (LibreOffice's column import,
/// sw/source/filter/xml/xmltbli.cxx `MakeTable_`, measured): relative
/// weights divide the width; absolute widths read in whole twips are kept
/// where they add up to the table, else a left-aligned table redistributes
/// the difference (`redistribute`) and other alignments scale them.
fn column_widths(
    columns: &[(Option<f64>, Option<f64>, Option<String>)],
    width: f64,
    jc: &str,
) -> Vec<f64> {
    let count = columns.len();
    if columns.iter().all(|c| c.1.is_some()) {
        let total: f64 = columns.iter().map(|c| c.1.unwrap_or(1.0)).sum();
        return columns
            .iter()
            .map(|c| c.1.unwrap_or(1.0) / total * width)
            .collect();
    }
    if columns.iter().all(|c| c.0.is_some()) {
        let widths: Vec<i64> = columns
            .iter()
            .map(|c| ((c.0.unwrap_or(0.0) * 20.0 + 0.5).floor() as i64).max(1))
            .collect();
        let total = (width * 20.0 + 0.5).floor() as i64;
        let sum: i64 = widths.iter().sum();
        let widths = if sum == total {
            widths
        } else if jc == "left" && total >= 23 * count as i64 {
            redistribute(&widths, total)
        } else {
            let scale = total as f64 / sum.max(1) as f64;
            widths
                .iter()
                .map(|w| (*w as f64 * scale).round() as i64)
                .collect()
        };
        return widths.iter().map(|w| *w as f64 / 20.0).collect();
    }
    // Mixed or missing widths: known absolute widths kept, the rest share
    // what remains equally.
    let known: f64 = columns.iter().filter_map(|c| c.0).sum();
    let unknown = columns.iter().filter(|c| c.0.is_none()).count();
    let share = if unknown > 0 {
        ((width - known) / unknown as f64).max(width / count as f64 / 4.0)
    } else {
        0.0
    };
    columns.iter().map(|c| c.0.unwrap_or(share)).collect()
}
