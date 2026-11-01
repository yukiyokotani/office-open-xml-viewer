//! Tables (ODF Part 3 §9.1; table, column, row and cell properties §17.15-
//! §17.18), measured against LibreOffice 26.8 PDF export (tables addendum).
//!
//! A table maps to a fixed-layout DOCX table with the authored column grid.
//! LibreOffice collapses borders: each cell edge is drawn once, centred on
//! the column edge, a cell's text starts at its edge plus its padding, and
//! the table's top and bottom borders lie inside the table. The paragraphs
//! around a table keep their own margins; cell paragraphs have none.
//!
//! Only what the corpus's tables use is modelled. Every other table feature
//! rejects with a broad rule rather than a precise model: table margins,
//! breaks and kept paragraphs at tables, repeated header rows, authored
//! row heights and rows kept together, cell backgrounds and centred or
//! bottom-aligned cells, merged or repeated cells, nested tables, lists,
//! headings and pictures in cells, cell text outside ASCII and CJK or above
//! 24 pt, double or dashed borders and vertical text.

use super::*;
use docx_model::{
    BorderSpec, CellBorders, CellElement, DocTable, DocTableCell, DocTableRow, TableBorders,
    TableCellLayoutAcquisitionWire, TableGridAcquisitionWire, TableGridColumnAcquisitionWire,
    TableLayoutAcquisitionWire, TableLayoutKindAcquisitionWire, TableMarginAcquisitionWire,
    TableWidthAcquisitionWire,
};

/// A property group's attributes.
type Attributes = HashMap<(String, String), String>;

/// Implementation bounds on one table's grid.
const MAX_COLUMNS: usize = 64;
const MAX_ROWS: usize = 10_000;

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
    }) {
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

/// LibreOffice's fit of absolute column widths `a` (twips) to a left-aligned
/// table `total` twips wide (`SwXMLTableContext::MakeTable_`, LibreOffice
/// 26.8): a shortfall is shared in proportion to the widths; an excess
/// gives every column MINLAY (23) and shares the rest in proportion to the
/// full widths; the last column takes what remains.
pub(super) fn redistribute(a: &[i64], total: i64) -> Vec<i64> {
    const MINLAY: i64 = 23;
    let sum: i64 = a.iter().sum();
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
    // LibreOffice holds border widths in whole twips, rounded half up
    // (measured: 0.02 pt draws nothing and lays out as no border, 0.05 pt
    // as one twip); a width at a half twip is left unresolved.
    let scaled = width * 20.0;
    if (scaled.fract() - 0.5).abs() < 1e-6 {
        return Err(unsupported("table-border"));
    }
    let twips = (scaled + 0.5).floor();
    if twips == 0.0 {
        return Ok(None);
    }
    Ok(Some(BorderSpec {
        width: twips / 20.0,
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
        // The measured model: collapsing borders (§20.302), stated by the
        // default table style or the table's own (without either LibreOffice
        // separates them, insetting text by the border; measured), and rows
        // that split at a page.
        let mut collapsing = false;
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
                        ("table-properties", NS_TABLE, "border-model") => {
                            collapsing = match value.as_str() {
                                "collapsing" => true,
                                "separating" => false,
                                _ => return Err(unsupported("table-default-style")),
                            }
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
                        ("table-row-properties", NS_FO, "keep-together") if value == "auto" => {}
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
        // Without an alignment LibreOffice fills the text area whatever the
        // width (measured); that is not modelled.
        let mut jc = "";
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
                // Page breaks at tables, like the table features below that
                // no corpus table uses, reject rather than being modelled.
                (NS_FO, "break-before" | "break-after") if value == "auto" => {}
                (NS_TABLE, "border-model") if value == "collapsing" => collapsing = true,
                (NS_FO, "background-color") if value == "transparent" => {}
                _ => return Err(unsupported("table-property")),
            }
        }
        if !collapsing {
            return Err(unsupported("table-border-model"));
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
        if jc.is_empty() {
            return Err(unsupported("table-align"));
        }
        // A table's top and bottom margins join the neighbouring paragraphs'
        // spacing (measured), but at page breaks and with keep-with-next
        // LibreOffice places them where the DOCX layout, carrying them in
        // those paragraphs, cannot (measured: a 710 pt top margin moves a
        // kept paragraph a page later). No corpus table has either.
        if top != 0.0 || bottom != 0.0 {
            return Err(unsupported("table-margin"));
        }
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

        // Columns (§9.1.6) and rows (§9.1.3).
        let mut columns: Vec<(Option<f64>, Option<u32>)> = Vec::new();
        let mut rows: Vec<roxmltree::Node<'_, '_>> = Vec::new();
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
                    for (key, value) in self
                        .table_group(
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
                                // LibreOffice reads integral weights.
                                share = Some(
                                    value
                                        .strip_suffix('*')
                                        .and_then(|v| v.parse::<u32>().ok())
                                        .filter(|v| *v > 0)
                                        .ok_or_else(|| unsupported("table-columns"))?,
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
                // Repeated header rows (§9.1.10) are not in the corpus.
                (Some(NS_TABLE), "table-header-rows") => {
                    return Err(unsupported("table-header-rows"));
                }
                (Some(NS_TABLE), "table-row") => rows.push(child),
                (Some(NS_TEXT), "soft-page-break") => {}
                _ => return Err(unsupported("table-child")),
            }
        }
        if columns.is_empty() || rows.is_empty() || rows.len() > MAX_ROWS {
            return Err(unsupported("table-shape"));
        }
        // LibreOffice's column import (sw/source/filter/xml/xmltbli.cxx,
        // MakeTable_; measured): a relative-width table divides its width by
        // the columns' relative weights, each raised to at least 23 (MINLAY)
        // first, so a smaller weight was not measured. An absolute table
        // reads widths in whole twips (rounded half up) and keeps them where
        // they add up to the table's twips, as every grid LibreOffice writes
        // does; otherwise a left-aligned table redistributes the difference
        // (`redistribute`; a 0.8 cm column next to 9.2 cm in a 10 cm table
        // draws 0.834 cm). Other alignments were not traced, and weights
        // beside absolute widths follow other rules: both reject.
        let col_widths: Vec<f64> = if relative.is_some() {
            let shares = columns
                .iter()
                .map(|c| {
                    c.1.filter(|share| (23..=65535).contains(share))
                        .map(f64::from)
                })
                .collect::<Option<Vec<f64>>>()
                .ok_or_else(|| unsupported("table-columns"))?;
            // Weights adding up past 65535 are normalised against that cap
            // (not measured).
            let total: f64 = shares.iter().sum();
            if total > 65535.0 {
                return Err(unsupported("table-columns"));
            }
            shares.iter().map(|s| s / total * table_width).collect()
        } else {
            let twips = |points: f64| {
                let scaled = points * 20.0;
                // A half twip rounds by floating-point error.
                if (scaled.fract() - 0.5).abs() < 1e-6 {
                    return None;
                }
                let value = (scaled + 0.5).floor();
                (23.0..=65535.0).contains(&value).then_some(value)
            };
            let widths = columns
                .iter()
                .map(|c| match c {
                    (Some(width), None) => twips(*width).map(|w| w as i64),
                    _ => None,
                })
                .collect::<Option<Vec<i64>>>()
                .ok_or_else(|| unsupported("table-columns"))?;
            let total = twips(table_width)
                .map(|w| w as i64)
                .filter(|w| *w >= 23 * widths.len() as i64)
                .ok_or_else(|| unsupported("table-columns"))?;
            let widths = if widths.iter().sum::<i64>() == total {
                widths
            } else if jc == "left" {
                redistribute(&widths, total)
            } else {
                return Err(unsupported("table-columns"));
            };
            widths.iter().map(|w| *w as f64 / 20.0).collect()
        };

        let enclosing_area = self.area;
        let mut doc_rows = Vec::with_capacity(rows.len());
        let row_count = rows.len();
        for row in rows {
            for attribute in row.attributes() {
                match (attribute.namespace(), attribute.name()) {
                    (Some(NS_TABLE), "style-name") => {}
                    _ => return Err(unsupported("table-row")),
                }
            }
            for (key, value) in self
                .table_group(
                    Family::TableRow,
                    row.attribute((NS_TABLE, "style-name")),
                    "table-row-properties",
                )?
                .into_iter()
                .flatten()
            {
                match (key.0.as_str(), key.1.as_str()) {
                    // Authored row heights and rows kept together are not
                    // modelled (a minimum height cut at a page differs).
                    (NS_FO, "keep-together") if value == "auto" => {}
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
                row_height_rule: "auto".into(),
                // LibreOffice cuts rows at a page its own way (see the flag).
                cuts_in_place: true,
                ..DocTableRow::default()
            });
        }
        self.area = enclosing_area;
        // LibreOffice continues rows whose vertical padding or borders run
        // past a page; the DOCX layout cannot split that fixed space and does
        // not converge. Each row's fixed space (a cell's vertical padding and
        // borders) is held to a quarter of the page body, leaving room for a
        // row's first line (bounded in `table_cell`).
        let bound = self.text_height / 4.0;
        let fixed = |row: &DocTableRow| {
            let rule = |side: &Option<BorderSpec>| side.as_ref().map_or(0.0, |b| b.width);
            row.cells
                .iter()
                .map(|cell| {
                    cell.margin_top.unwrap_or(0.0)
                        + cell.margin_bottom.unwrap_or(0.0)
                        + rule(&cell.borders.top)
                        + rule(&cell.borders.bottom)
                })
                .fold(0.0, f64::max)
        };
        // LibreOffice starts the content of every cell in a row below the
        // row's largest top padding (measured: 0 and 20 pt both start 20 pt
        // down); the DOCX layout pads each cell apart.
        if doc_rows.iter().any(|row| {
            row.cells.windows(2).any(|pair| {
                pair[0].margin_top != pair[1].margin_top
                    || pair[0].margin_bottom != pair[1].margin_bottom
            })
        }) {
            return Err(unsupported("table-cell-padding"));
        }
        for row in &doc_rows {
            let space = fixed(row);
            if space.partial_cmp(&bound) == Some(std::cmp::Ordering::Greater) || space.is_nan() {
                return Err(unsupported("table-row-height"));
            }
        }
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
        // Where a page break falls between two rows, LibreOffice draws their
        // shared edge again above the row continuing on the next page, and
        // below the row ending the page only where that row has a bottom
        // rule of its own (measured: a 12 pt bottom rule alone moves the
        // continued rows 12 pt down; a top rule alone adds nothing at the
        // page's foot). The lower cell carries the wider rule, as the DOCX
        // layout repeats a row's top rule there; without a break the layout
        // draws the shared edge once at the wider width either way.
        for index in 1..doc_rows.len() {
            let (above, below) = doc_rows.split_at_mut(index);
            for (upper, lower) in above[index - 1].cells.iter().zip(&mut below[0].cells) {
                let width = |side: &Option<BorderSpec>| side.as_ref().map_or(0.0, |b| b.width);
                if width(&upper.borders.bottom) > width(&lower.borders.top) {
                    lower.borders.top = upper.borders.bottom.clone();
                }
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
        self.blocks.push(Block::Table(Box::new(table)));
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
        // Shorthands first: a side set beside them wins (§20.183, §20.217,
        // after XSL's precedence of specific over shorthand properties).
        let mut entries: Vec<_> = self
            .table_group(
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
                // Cell backgrounds and centred or bottom-aligned content are
                // not in the corpus.
                (NS_FO, "background-color") if value == "transparent" => {}
                (NS_STYLE, "vertical-align") if matches!(value.as_str(), "top" | "automatic") => {}
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
        let mut content = finish_cell(blocks)?;
        if content.is_empty() {
            return Err(unsupported("table-cell-empty"));
        }
        // A row's first line must fit a page beside its fixed space (held to
        // a quarter page): LibreOffice
        // continues a row whose first line does not, the DOCX layout cannot
        // advance. Each line's contribution is held to a quarter page too,
        // counting a font at four times its size (room for tall font
        // metrics) and its line height.
        let mark = |p: &DocParagraph| {
            p.paragraph_mark_font_facts
                .as_ref()
                .and_then(|facts| facts.font_size)
                .into_iter()
                .chain(p.default_font_size)
                .fold(0.0, f64::max)
                * 4.0
        };
        // Fonts above 24 pt and pictures in cells are not in the corpus.
        for element in &content {
            let CellElement::Paragraph(p) = element else {
                continue;
            };
            for run in &p.runs {
                match run {
                    DocRun::Text(text) if text.font_size > 24.0 => {
                        return Err(unsupported("table-cell-line"));
                    }
                    DocRun::Image(_) => return Err(unsupported("table-cell-content")),
                    _ => {}
                }
            }
        }
        let mut tallest: f64 = 0.0;
        // The cell's text width after its padding.
        let available = width - margins[1] - margins[3];
        for element in &content {
            let CellElement::Paragraph(p) = element else {
                continue;
            };
            // The paragraph mark sizes an empty paragraph's line; a
            // proportional line height scales the line, a fixed or minimum
            // one sets it at least.
            let (factor, floor) = match &p.line_spacing {
                Some(spacing) if spacing.rule == "auto" => (spacing.value.max(1.0), 0.0),
                Some(spacing) => (1.0, spacing.value),
                None => (1.0, 0.0),
            };
            let line = p
                .runs
                .iter()
                .map(|run| match run {
                    DocRun::Text(text) => text.font_size * 4.0,
                    _ => 0.0,
                })
                .fold(mark(p), f64::max)
                * factor;
            tallest = tallest.max(line).max(floor);
            // Indents leaving no width in the cell (LibreOffice clips the
            // text) or reaching outside it were not measured.
            let first = p.indent_left + p.indent_first;
            if p.indent_left < 0.0
                || first < 0.0
                || p.indent_right < 0.0
                || p.indent_left.max(first) + p.indent_right >= available
            {
                return Err(unsupported("table-cell-indent"));
            }
        }
        if tallest.partial_cmp(&(self.text_height / 4.0)) != Some(std::cmp::Ordering::Less) {
            return Err(unsupported("table-cell-line"));
        }
        // LibreOffice clips a glyph wider than the cell's line; the DOCX
        // layout paints it whole. Cell text is held to ASCII and CJK, whose
        // glyphs keep within an em (other characters are not in the corpus:
        // U+2E3B advances three), and a line must be at least 1.25 em wide.
        let narrow = |c: char| {
            c.is_ascii()
                || matches!(c, '\u{3000}'..='\u{9fff}' | '\u{f900}'..='\u{faff}' | '\u{ff00}'..='\u{ffef}')
        };
        for element in &content {
            let CellElement::Paragraph(p) = element else {
                continue;
            };
            let line =
                available - p.indent_left.max(p.indent_left + p.indent_first) - p.indent_right;
            for run in &p.runs {
                if let DocRun::Text(text) = run {
                    if !text.text.chars().all(narrow) {
                        return Err(unsupported("table-cell-text"));
                    }
                    if line < text.font_size * 1.25 {
                        return Err(unsupported("table-cell-indent"));
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
            v_align: "top".into(),
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

/// A cell's paragraphs. LibreOffice repeats a cell paragraph's top margin
/// where the cell continues on the next page, which the DOCX layout does
/// not, and aligns centred or bottom-aligned content with the outer margins
/// (measured), so cell paragraphs admit no vertical margins (none in the
/// corpus has one).
fn finish_cell(blocks: Vec<Block>) -> Result<Vec<CellElement>, String> {
    let mut out = Vec::with_capacity(blocks.len());
    for block in blocks {
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
        if collected.top != 0.0 || collected.bottom != 0.0 {
            return Err(unsupported("table-cell-spacing"));
        }
        let mut paragraph = std::mem::take(&mut collected.paragraph);
        paragraph.space_before = 0.0;
        paragraph.space_after = 0.0;
        if paragraph.shading.is_some() {
            return Err(unsupported("table-cell-paragraph-fill"));
        }
        paragraph.contextual_spacing = false;
        // LibreOffice cuts a cell paragraph at a page without its
        // widow/orphan or keep-together controls (measured: four-line
        // paragraphs split three/one with widows and orphans 2 and with
        // keep-together always).
        paragraph.widow_control = false;
        paragraph.keep_lines = false;
        out.push(CellElement::Paragraph(Box::new(paragraph)));
    }
    Ok(out)
}
