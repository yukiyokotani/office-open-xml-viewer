//! Tables (ODF Part 3 §9.1; table, column, row and cell properties §17.15-
//! §17.18), measured against LibreOffice 26.8 PDF export (tables addendum).
//!
//! A table maps to a fixed-layout DOCX table with the authored column grid.
//! LibreOffice collapses borders: each cell edge is drawn once, centred on
//! the column edge, a cell's text starts at its edge plus its padding, and
//! the table's top and bottom borders lie inside the table. Spacing around a
//! table adds the neighbouring paragraphs' margins and the table's own; a
//! cell keeps its first paragraph's space before and its last paragraph's
//! space after. Only the corpus's table shape is admitted: no merged or
//! repeated cells, nested tables, lists or headings in cells, double borders
//! or vertical text.

use super::*;
use docx_model::{
    BorderSpec, CellBorders, CellElement, DocTable, DocTableCell, DocTableRow, TableBorders,
    TableCellLayoutAcquisitionWire, TableGridAcquisitionWire, TableGridColumnAcquisitionWire,
    TableLayoutAcquisitionWire, TableLayoutKindAcquisitionWire, TableMarginAcquisitionWire,
    TableRowHeightAcquisitionWire, TableRowLayoutAcquisitionWire, TableWidthAcquisitionWire,
};

/// A property group's attributes.
type Attributes = HashMap<(String, String), String>;

/// Implementation bounds on one table's grid.
const MAX_COLUMNS: usize = 64;
const MAX_ROWS: usize = 10_000;

/// A table collected among the body's paragraphs, with the vertical margins
/// that join its neighbours' spacing.
pub(super) struct TableBlock {
    pub table: DocTable,
    pub top: f64,
    pub bottom: f64,
}

/// The properties of a table-family automatic style (§16.2: automatic
/// styles carry every property they set; table styles have no parents in
/// the admitted subset).
fn group<'s>(
    sheet: &'s StyleSheet,
    family: Family,
    name: Option<&str>,
    group: &str,
) -> Result<Option<&'s Attributes>, String> {
    let Some(name) = name else {
        return Ok(None);
    };
    let style = sheet
        .find(Scope::Content, family, name)
        .ok_or_else(|| invalid("a table names a missing style"))?;
    let style = style.style();
    if style.parent.is_some() {
        return Err(unsupported("table-style-parent"));
    }
    // Other style-element attributes (a page style to start, a class) are
    // not modelled; a display name only names the style in the UI.
    if style.attributes.keys().any(|key| {
        key.0 != NS_STYLE || !matches!(key.1.as_str(), "name" | "family" | "display-name")
    })
    {
        return Err(unsupported("table-style-attribute"));
    }
    for (key, properties) in &style.groups {
        // LibreOffice writes an empty `style:background-image` beside a
        // background colour; it draws nothing.
        let inert = properties.elements.iter().all(|(name, elements)| {
            name == "background-image"
                && elements
                    .iter()
                    .all(|e| e.attributes.is_empty() && e.children.is_empty())
        });
        if key != group || !inert {
            return Err(unsupported("table-style-group"));
        }
    }
    Ok(style.groups.get(group).map(|p| &p.attributes))
}

impl Reader<'_> {
    /// `group` of a table-family style, charged by the bytes each use scans
    /// (a style shared by many cells is scanned once per cell).
    fn table_group(
        &self,
        family: Family,
        name: Option<&str>,
        group_name: &str,
    ) -> Result<Option<&Attributes>, String> {
        let properties = group(self.sheet, family, name, group_name)?;
        let bytes: usize = properties.map_or(0, |p| {
            p.iter()
                .map(|(key, value)| key.0.len() + key.1.len() + value.len())
                .sum()
        });
        self.charge_work(1 + bytes)?;
        Ok(properties)
    }
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

/// `fo:border` (§20.183): "none", or a width, "solid" and a colour.
fn border(value: &str) -> Result<Option<BorderSpec>, String> {
    if value == "none" {
        return Ok(None);
    }
    let mut parts = value.split_whitespace();
    let (Some(width), Some(style), Some(color), None) =
        (parts.next(), parts.next(), parts.next(), parts.next())
    else {
        return Err(unsupported("table-border"));
    };
    let width = length_pt(width)?;
    if style != "solid" || !(width > 0.0 && width <= 12.0) {
        return Err(unsupported("table-border"));
    }
    Ok(Some(BorderSpec {
        width,
        color: Some(hex_color(color)?),
        style: "single".into(),
    }))
}

impl<'a> Reader<'a> {
    /// A `table:table` in the body (§9.1.2).
    pub(super) fn table(&mut self, node: roxmltree::Node<'_, '_>) -> Result<(), String> {
        if !self.layout_flags.table_settings {
            return Err(unsupported("table-setting"));
        }
        for attribute in node.attributes() {
            match (attribute.namespace(), attribute.name()) {
                (Some(NS_TABLE), "name" | "style-name") | (Some(NS_XML), "id") => {}
                _ => return Err(unsupported("table-attribute")),
            }
        }
        self.charge_construction(std::mem::size_of::<DocTable>())?;
        // The measured model: collapsing borders (§20.302), which
        // LibreOffice also reads where no style states a model (it writes
        // "collapsing" back; the controls state none), and rows that do not
        // split unless their default says so.
        let mut row_keep_default = false;
        if let Some(default) = self.sheet.default_style(Family::Table) {
            for (key, properties) in &default.style().groups {
                self.charge_work(
                    1 + properties
                        .attributes
                        .iter()
                        .map(|(key, value)| key.0.len() + key.1.len() + value.len())
                        .sum::<usize>(),
                )?;
                for (attribute, value) in &properties.attributes {
                    match (key.as_str(), attribute.0.as_str(), attribute.1.as_str()) {
                        ("table-properties", NS_TABLE, "border-model") if value == "collapsing" => {
                        }
                        _ => return Err(unsupported("table-default-style")),
                    }
                }
            }
        }
        if let Some(default) = self.sheet.default_style(Family::TableRow) {
            for (key, properties) in &default.style().groups {
                self.charge_work(
                    1 + properties
                        .attributes
                        .iter()
                        .map(|(key, value)| key.0.len() + key.1.len() + value.len())
                        .sum::<usize>(),
                )?;
                for (attribute, value) in &properties.attributes {
                    match (key.as_str(), attribute.0.as_str(), attribute.1.as_str()) {
                        ("table-row-properties", NS_FO, "keep-together") => {
                            row_keep_default = match value.as_str() {
                                "always" => true,
                                "auto" => false,
                                _ => return Err(unsupported("table-default-style")),
                            }
                        }
                        _ => return Err(unsupported("table-default-style")),
                    }
                }
            }
        }
        for family in [Family::TableColumn, Family::TableCell] {
            if self.sheet.default_style(family).is_some() {
                return Err(unsupported("table-default-style"));
            }
        }

        // Table properties (§17.15).
        let properties = self.table_group(
            Family::Table,
            node.attribute((NS_TABLE, "style-name")),
            "table-properties",
        )?;
        let mut width = None;
        let mut relative = None;
        let (mut margin_left, mut top, mut bottom) = (0.0, 0.0, 0.0);
        let mut jc = "left";
        let mut break_before = false;
        for (key, value) in properties.into_iter().flatten() {
            match (key.0.as_str(), key.1.as_str()) {
                (NS_STYLE, "width") => width = Some(length_pt(value)?),
                (NS_STYLE, "rel-width") => relative = Some(percent(value)?),
                (NS_FO, "margin-left") => margin_left = length_pt(value)?,
                (NS_FO, "margin-right") if length_pt(value)? == 0.0 => {}
                (NS_FO, "margin-top") => top = length_pt(value)?,
                (NS_FO, "margin-bottom") => bottom = length_pt(value)?,
                (NS_TABLE, "align") => {
                    jc = match value.as_str() {
                        "left" => "left",
                        "center" => "center",
                        "right" => "right",
                        _ => return Err(unsupported("table-align")),
                    }
                }
                (NS_STYLE, "writing-mode") if matches!(value.as_str(), "page" | "lr-tb") => {}
                (NS_FO, "break-before") => {
                    break_before = match value.as_str() {
                        "page" => true,
                        "auto" => false,
                        _ => return Err(unsupported("table-break")),
                    }
                }
                (NS_FO, "break-after") if value == "auto" => {}
                (NS_TABLE, "border-model") if value == "collapsing" => {}
                (NS_FO, "background-color") if value == "transparent" => {}
                _ => return Err(unsupported("table-property")),
            }
        }
        // The table's x and width (measured): a 100% relative width fills
        // the text area, else the authored width; aligned left it starts at
        // the margin-left, centred or right-aligned it is placed by its width.
        let text_width = self.text_width;
        let table_width = match relative {
            Some(r) if (r - 1.0).abs() < 1e-9 => text_width,
            Some(_) => return Err(unsupported("table-rel-width")),
            None => width.ok_or_else(|| unsupported("table-width"))?,
        };
        if jc != "left" && margin_left != 0.0 {
            return Err(unsupported("table-align"));
        }
        // LibreOffice keeps the authored width where it overflows the text
        // area into the page margins (measured); beyond the page edge it is
        // not measured.
        let (start, end) = match jc {
            "left" => (margin_left, margin_left + table_width),
            "center" => (
                (text_width - table_width) / 2.0,
                (text_width + table_width) / 2.0,
            ),
            _ => (text_width - table_width, text_width),
        };
        if table_width.is_nan()
            || table_width <= 0.0
            || start < -self.page_left_margin
            || end > text_width + self.page_right_margin
        {
            return Err(unsupported("table-geometry"));
        }

        // Columns (§9.1.6) and rows (§9.1.3), header rows first (§9.1.10).
        let mut columns: Vec<(Option<f64>, Option<f64>)> = Vec::new();
        let mut rows: Vec<(roxmltree::Node<'_, '_>, bool)> = Vec::new();
        for child in node.children() {
            if child.is_text() {
                if child.text().is_some_and(|t| !t.trim().is_empty()) {
                    return Err(invalid("text in a table"));
                }
                continue;
            }
            if !child.is_element() {
                continue;
            }
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TABLE), "table-column") if rows.is_empty() => {
                    let mut repeat = 1usize;
                    for attribute in child.attributes() {
                        match (attribute.namespace(), attribute.name()) {
                            (Some(NS_TABLE), "style-name") => {}
                            (Some(NS_TABLE), "number-columns-repeated") => {
                                repeat = attribute
                                    .value()
                                    .parse()
                                    .ok()
                                    .filter(|n| (1..=MAX_COLUMNS).contains(n))
                                    .ok_or_else(|| unsupported("table-columns"))?;
                            }
                            _ => return Err(unsupported("table-column")),
                        }
                    }
                    let mut absolute = None;
                    let mut share = None;
                    for (key, value) in self.table_group(
                        Family::TableColumn,
                        child.attribute((NS_TABLE, "style-name")),
                        "table-column-properties",
                    )?
                    .into_iter()
                    .flatten()
                    {
                        match (key.0.as_str(), key.1.as_str()) {
                            (NS_STYLE, "column-width") => absolute = Some(length_pt(value)?),
                            (NS_STYLE, "rel-column-width") => {
                                share = Some(
                                    value
                                        .strip_suffix('*')
                                        .and_then(|v| v.parse::<f64>().ok())
                                        .filter(|v| *v > 0.0 && v.is_finite())
                                        .ok_or_else(|| invalid("bad relative column width"))?,
                                )
                            }
                            _ => return Err(unsupported("table-column")),
                        }
                    }
                    if columns.len() + repeat > MAX_COLUMNS {
                        return Err(unsupported("table-columns"));
                    }
                    columns.extend(std::iter::repeat_n((absolute, share), repeat));
                }
                (Some(NS_TABLE), "table-header-rows") if rows.is_empty() => {
                    for row in child.children().filter(|n| n.is_element()) {
                        if !xml::is(row, NS_TABLE, "table-row") {
                            return Err(unsupported("table-child"));
                        }
                        rows.push((row, true));
                    }
                }
                (Some(NS_TABLE), "table-row") => rows.push((child, false)),
                (Some(NS_TEXT), "soft-page-break") => {}
                _ => return Err(unsupported("table-child")),
            }
        }
        if columns.is_empty() || rows.is_empty() || rows.len() > MAX_ROWS {
            return Err(unsupported("table-shape"));
        }
        // LibreOffice sizes columns by their relative weights wherever every
        // column has one, also beside absolute widths (measured: 4 cm/6 cm
        // columns weighted 1000*/3000* in a 10 cm table draw 2.5/7.5 cm).
        let weighted = columns.iter().all(|c| c.1.is_some());
        if !weighted && columns.iter().any(|c| c.1.is_some()) {
            return Err(unsupported("table-columns"));
        }
        let col_widths: Vec<f64> = if relative.is_some() || weighted {
            let total: f64 = columns
                .iter()
                .map(|c| c.1.ok_or_else(|| unsupported("table-columns")))
                .sum::<Result<f64, String>>()?;
            columns
                .iter()
                .map(|c| c.1.expect("checked") / total * table_width)
                .collect()
        } else {
            let widths: Vec<f64> = columns
                .iter()
                .map(|c| c.0.ok_or_else(|| unsupported("table-columns")))
                .collect::<Result<_, _>>()?;
            // The grid must add up to the table's width (as LibreOffice
            // writes it); a mismatch is not measured.
            if (widths.iter().sum::<f64>() - table_width).abs() > 0.5 {
                return Err(unsupported("table-columns"));
            }
            widths
        };

        let enclosing_area = self.area;
        let mut doc_rows = Vec::with_capacity(rows.len());
        let row_count = rows.len();
        for (row, header) in rows {
            if header && doc_rows.iter().any(|r: &DocTableRow| !r.is_header) {
                return Err(unsupported("table-header-rows"));
            }
            for attribute in row.attributes() {
                match (attribute.namespace(), attribute.name()) {
                    (Some(NS_TABLE), "style-name") => {}
                    _ => return Err(unsupported("table-row")),
                }
            }
            let mut height = None;
            let mut cant_split = row_keep_default;
            for (key, value) in self.table_group(
                Family::TableRow,
                row.attribute((NS_TABLE, "style-name")),
                "table-row-properties",
            )?
            .into_iter()
            .flatten()
            {
                match (key.0.as_str(), key.1.as_str()) {
                    // Both together resolve in an order that was not
                    // measured.
                    (NS_STYLE, "row-height" | "min-row-height") => {
                        let rule = if key.1 == "row-height" { "exact" } else { "atLeast" };
                        if height.replace((length_pt(value)?, rule)).is_some() {
                            return Err(unsupported("table-row-height"));
                        }
                    }
                    (NS_FO, "keep-together") => {
                        cant_split = match value.as_str() {
                            "always" => true,
                            "auto" => false,
                            _ => return Err(unsupported("table-row")),
                        }
                    }
                    (NS_FO, "background-color") if value == "transparent" => {}
                    _ => return Err(unsupported("table-row")),
                }
            }
            let mut cells = Vec::new();
            for cell in row.children() {
                if cell.is_text() {
                    if cell.text().is_some_and(|t| !t.trim().is_empty()) {
                        return Err(invalid("text in a table row"));
                    }
                    continue;
                }
                if !cell.is_element() {
                    continue;
                }
                if !xml::is(cell, NS_TABLE, "table-cell") {
                    return Err(unsupported("table-cell"));
                }
                let column = cells.len();
                if column >= col_widths.len() {
                    return Err(unsupported("table-shape"));
                }
                cells.push(self.table_cell(cell, col_widths[column])?);
            }
            if cells.len() != col_widths.len() {
                return Err(unsupported("table-shape"));
            }
            doc_rows.push(DocTableRow {
                cells,
                row_height: height.map(|h| h.0),
                row_height_rule: height.map_or("auto", |h| h.1).into(),
                is_header: header,
                cant_split,
                // LibreOffice splits a kept-together row taller than a page.
                over_page_continues: cant_split,
                table_row_layout: TableRowLayoutAcquisitionWire {
                    height: height.map(|(value, rule)| TableRowHeightAcquisitionWire {
                        value: Some(twips(value)),
                        rule: rule.into(),
                        rule_authored: true,
                    }),
                    ..TableRowLayoutAcquisitionWire::default()
                },
                ..DocTableRow::default()
            });
        }
        self.area = enclosing_area;
        // Collapsed borders on a shared edge: the wider one is drawn in
        // LibreOffice and in the DOCX layout alike; two of equal width that
        // differ in colour resolve by rules that were not compared.
        let conflict = |a: &Option<BorderSpec>, b: &Option<BorderSpec>| match (a, b) {
            (Some(a), Some(b)) => a.width == b.width && a.color != b.color,
            _ => false,
        };
        for (index, row) in doc_rows.iter().enumerate() {
            self.charge_work(row.cells.len())?;
            let across = row
                .cells
                .windows(2)
                .any(|pair| conflict(&pair[0].borders.right, &pair[1].borders.left));
            let down = doc_rows.get(index + 1).is_some_and(|next| {
                row.cells
                    .iter()
                    .zip(&next.cells)
                    .any(|(cell, below)| conflict(&cell.borders.bottom, &below.borders.top))
            });
            if across || down {
                return Err(unsupported("table-border-conflict"));
            }
        }
        // LibreOffice counts a row's explicit height (fixed or minimum) from
        // the top of the border above it (measured). The DOCX layout does so
        // for an exact row (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT) but adds
        // that border to an atLeast row, so a minimum height is given net of
        // it; the border is drawn once at the thicker width. Word also adds
        // the cells' bottom margins to an exact row, which LibreOffice was not
        // measured to do, so exact rows admit no bottom padding.
        for index in 0..doc_rows.len() {
            let width = |cell: &DocTableCell, top: bool| {
                let side = if top {
                    &cell.borders.top
                } else {
                    &cell.borders.bottom
                };
                side.as_ref().map_or(0.0, |b| b.width)
            };
            match doc_rows[index].row_height_rule.as_str() {
                "exact" => {
                    if doc_rows[index]
                        .cells
                        .iter()
                        .any(|c| c.margin_bottom != Some(0.0))
                    {
                        return Err(unsupported("table-row-padding"));
                    }
                }
                "atLeast" => {
                    let above = index
                        .checked_sub(1)
                        .map(|i| {
                            doc_rows[i]
                                .cells
                                .iter()
                                .map(|c| width(c, false))
                                .fold(0.0, f64::max)
                        })
                        .unwrap_or(0.0);
                    let shared = doc_rows[index]
                        .cells
                        .iter()
                        .map(|c| width(c, true))
                        .fold(above, f64::max);
                    let row = &mut doc_rows[index];
                    let height = row.row_height.expect("an atLeast row has a height") - shared;
                    if height <= 0.0 {
                        return Err(unsupported("table-row"));
                    }
                    row.row_height = Some(height);
                    if let Some(wire) = row.table_row_layout.height.as_mut() {
                        wire.value = Some(twips(height));
                    }
                }
                _ => {}
            }
        }
        self.next_table += 1;
        let table = DocTable {
            col_widths: col_widths.clone(),
            rows: doc_rows,
            borders: TableBorders::default(),
            jc: jc.into(),
            tbl_ind: (jc == "left").then_some(margin_left),
            layout: Some("fixed".into()),
            width_pt: Some(table_width),
            // A table that already starts a page breaks no further, as a
            // paragraph's break does not (measured).
            page_break_before: break_before,
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
        self.charge_construction(std::mem::size_of::<DocTableRow>() * row_count)?;
        self.blocks
            .push(Block::Table(Box::new(TableBlock { table, top, bottom })));
        Ok(())
    }

    /// One `table:table-cell` (§9.1.4) and its content, collected apart from
    /// the body's blocks.
    fn table_cell(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        width: f64,
    ) -> Result<DocTableCell, String> {
        for attribute in node.attributes() {
            match (attribute.namespace(), attribute.name()) {
                (Some(NS_TABLE), "style-name") => {}
                // The value type only matters in spreadsheets; a text cell
                // shows its paragraphs.
                (Some(NS_OFFICE), "value-type") if attribute.value() == "string" => {}
                _ => return Err(unsupported("table-cell")),
            }
        }
        let mut margins = [0.0f64; 4];
        let mut sides: [Option<BorderSpec>; 4] = [None, None, None, None];
        let mut background = None;
        let mut v_align = "top";
        // Shorthands first: a side set beside them wins (§20.183, §20.217,
        // after XSL's precedence of specific over shorthand properties).
        let mut entries: Vec<_> = self.table_group(
            Family::TableCell,
            node.attribute((NS_TABLE, "style-name")),
            "table-cell-properties",
        )?
        .into_iter()
        .flatten()
        .collect();
        entries.sort_by(|(a, _), (b, _)| {
            let side = |key: &(String, String)| !matches!(key.1.as_str(), "padding" | "border");
            (side(a), a).cmp(&(side(b), b))
        });
        for (key, value) in entries {
            match (key.0.as_str(), key.1.as_str()) {
                (NS_FO, "padding") => margins = [length_pt(value)?; 4],
                (NS_FO, "padding-top") => margins[0] = length_pt(value)?,
                (NS_FO, "padding-left") => margins[1] = length_pt(value)?,
                (NS_FO, "padding-bottom") => margins[2] = length_pt(value)?,
                (NS_FO, "padding-right") => margins[3] = length_pt(value)?,
                (NS_FO, "border") => {
                    let spec = border(value)?;
                    sides = [spec.clone(), spec.clone(), spec.clone(), spec];
                }
                (NS_FO, "border-top") => sides[0] = border(value)?,
                (NS_FO, "border-left") => sides[1] = border(value)?,
                (NS_FO, "border-bottom") => sides[2] = border(value)?,
                (NS_FO, "border-right") => sides[3] = border(value)?,
                (NS_FO, "background-color") => {
                    background = (value != "transparent")
                        .then(|| hex_color(value))
                        .transpose()?
                }
                (NS_STYLE, "vertical-align") => {
                    v_align = match value.as_str() {
                        "top" | "automatic" => "top",
                        "middle" => "center",
                        "bottom" => "bottom",
                        _ => return Err(unsupported("table-cell-property")),
                    }
                }
                (NS_STYLE, "writing-mode") if matches!(value.as_str(), "page" | "lr-tb") => {}
                _ => return Err(unsupported("table-cell-property")),
            }
        }
        if margins.iter().any(|m| *m < 0.0) || margins[1] + margins[3] >= width {
            return Err(unsupported("table-cell-geometry"));
        }

        // The cell's paragraphs, in their own content area.
        let saved = std::mem::take(&mut self.blocks);
        self.next_area += 1;
        self.area = self.next_area;
        for child in node.children() {
            if child.is_text() {
                if child.text().is_some_and(|t| !t.trim().is_empty()) {
                    return Err(invalid("text in a table cell"));
                }
                continue;
            }
            if !child.is_element() {
                continue;
            }
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "p") => self.paragraph(child, None)?,
                (Some(NS_TEXT), "soft-page-break") => {}
                _ => return Err(unsupported("table-cell-content")),
            }
        }
        let blocks = std::mem::replace(&mut self.blocks, saved);
        let mut content = finish_cell(blocks, self.layout_flags)?;
        if content.is_empty() {
            return Err(unsupported("table-cell-empty"));
        }
        // LibreOffice centres or bottom-aligns a cell's content with the
        // first paragraph's top and the last one's bottom margin; the DOCX
        // layout aligns without them (measured: 12 pt above moves centred
        // text 6 pt, 10 pt below moves bottom-aligned text 10 pt).
        let outer = |element: Option<&CellElement>, top: bool| match element {
            Some(CellElement::Paragraph(p)) => {
                if top {
                    p.space_before
                } else {
                    p.space_after
                }
            }
            _ => 0.0,
        };
        if v_align != "top" && (outer(content.first(), true) != 0.0 || outer(content.last(), false) != 0.0)
        {
            return Err(unsupported("table-cell-align-spacing"));
        }
        if let Some(fill) = &background {
            // §20.395 automatic colour over the cell's background, as over a
            // paragraph's (cell paragraphs have none of their own).
            let resolved = libreoffice_automatic_color(fill)?;
            for element in &mut content {
                let CellElement::Paragraph(paragraph) = element else {
                    continue;
                };
                for run in &mut paragraph.runs {
                    if let DocRun::Text(text) = run {
                        if text.color.is_none() && text.background.is_none() {
                            text.color = Some(resolved.to_string());
                            text.color_auto = false;
                        }
                    }
                }
            }
        }
        let margin_wire = TableMarginAcquisitionWire {
            top: Some(dxa(margins[0])),
            left: Some(dxa(margins[1])),
            bottom: Some(dxa(margins[2])),
            right: Some(dxa(margins[3])),
            start: None,
            end: None,
        };
        let [top, left, bottom, right] = sides;
        Ok(DocTableCell {
            content: std::mem::take(&mut content),
            col_span: 1,
            borders: CellBorders {
                top,
                left,
                bottom,
                right,
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
            ..DocTableCell::default()
        })
    }
}

/// A cell's paragraphs with LibreOffice's spacing inside a cell (measured):
/// the first keeps its space before, consecutive ones are separated as body
/// paragraphs are (`LayoutFlags::gap`), and the last keeps its space after.
fn finish_cell(blocks: Vec<Block>, flags: LayoutFlags) -> Result<Vec<CellElement>, String> {
    let mut out = Vec::with_capacity(blocks.len());
    let mut previous: Option<Collected> = None;
    let count = blocks.len();
    for (index, block) in blocks.into_iter().enumerate() {
        let Block::Paragraph(collected) = block else {
            return Err(unsupported("table-cell-content"));
        };
        let mut collected = *collected;
        if collected.page_break_after || collected.paragraph.page_break_before {
            return Err(unsupported("table-cell-break"));
        }
        // With TableRowKeep a cell paragraph kept with the next keeps its
        // row with the next one; unmeasured.
        if collected.paragraph.keep_next {
            return Err(unsupported("table-cell-keep"));
        }
        let mut paragraph = std::mem::take(&mut collected.paragraph);
        paragraph.space_before = match &previous {
            None => collected.top,
            Some(prev) => {
                let same = prev.named_style.is_some() && prev.named_style == collected.named_style;
                if same && prev.contextual != collected.contextual {
                    return Err(unsupported("contextual-spacing-mixed"));
                }
                if same && prev.contextual {
                    0.0
                } else {
                    flags.gap(prev.bottom, collected.top)
                }
            }
        };
        paragraph.space_after = if index + 1 == count {
            collected.bottom
        } else {
            0.0
        };
        if paragraph.shading.is_some() {
            return Err(unsupported("table-cell-paragraph-fill"));
        }
        paragraph.contextual_spacing = false;
        out.push(CellElement::Paragraph(Box::new(paragraph)));
        previous = Some(collected);
    }
    Ok(out)
}
