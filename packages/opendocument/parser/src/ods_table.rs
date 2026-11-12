//! One `table:table` (ODF Part 3 §9.1) as a worksheet: columns and rows with
//! their repeat counts (§19.675 / §19.680), cells and their values (§19.385
//! `office:value-type`), spans (§19.676), styles, drawings and annotations.
//!
//! Repeats are walked as runs, never expanded: a repeated row or cell is
//! materialized only when it carries content or a format that differs from
//! its column's default, so LibreOffice's trailing million-row filler costs
//! one run.

use super::{
    cells::{CellStyle, PartialFont, XfBuilder},
    column_letters, date_serial, draw, duration_days, pt_to_px, Context, OdsSheet, MAX_COLS,
    MAX_ROWS, NS_TABLE,
};
use crate::budget;
use crate::odf_props::{get, hex_color, length_pt};
use crate::ods_number::DataKind;
use crate::package::OdfPackage;
use crate::xml::{self, NS_OFFICE, NS_TEXT, NS_XLINK};
use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;
use xlsx_model::{Cell, CellValue, Hyperlink, MergeCell, Row, Run, Worksheet};

const NS_CALCEXT: &str = crate::odf_props::NS_CALCEXT;
const NS_CONFIG: &str = crate::xml::NS_CONFIG;
const NS_DC: &str = "http://purl.org/dc/elements/1.1/";

/// LibreOffice's default row height (0.452 cm) and column width (2.258 cm),
/// used when a table declares no row or column style.
const DEFAULT_ROW_PT: f64 = 12.8;
const DEFAULT_COL_PT: f64 = 64.0;

#[derive(Clone, Debug, Default)]
pub struct ViewSettings {
    pub freeze_rows: u32,
    pub freeze_cols: u32,
    pub show_grid: Option<bool>,
    pub show_zeros: Option<bool>,
    pub tab_color: Option<String>,
}

/// `settings.xml` view settings (LibreOffice's `ooo:view-settings`): frozen
/// panes (split mode 2), grid and zero display, per table.
pub fn read_view_settings(
    root: roxmltree::Node<'_, '_>,
) -> Result<HashMap<String, ViewSettings>, String> {
    let mut tables = HashMap::new();
    let item = |node: roxmltree::Node<'_, '_>, name: &str| -> Option<String> {
        node.children()
            .find(|n| {
                xml::is(*n, NS_CONFIG, "config-item")
                    && n.attribute((NS_CONFIG, "name")) == Some(name)
            })
            .and_then(|n| n.text())
            .map(str::to_string)
    };
    let Some(view) = root
        .descendants()
        .find(|n| {
            xml::is(*n, NS_CONFIG, "config-item-map-indexed")
                && n.attribute((NS_CONFIG, "name")) == Some("Views")
        })
        .and_then(|views| {
            views
                .children()
                .find(|n| xml::is(*n, NS_CONFIG, "config-item-map-entry"))
        })
    else {
        return Ok(tables);
    };
    // View settings first, then the document configuration
    // (`ooo:configuration-settings`), where LibreOffice keeps them too.
    let configuration = root.descendants().find(|n| {
        xml::is(*n, NS_CONFIG, "config-item-set")
            && n.attribute((NS_CONFIG, "name")) == Some("ooo:configuration-settings")
    });
    let setting = |name: &str| {
        item(view, name)
            .or_else(|| configuration.and_then(|c| item(c, name)))
            .map(|v| v == "true")
    };
    let global_grid = setting("ShowGrid");
    let global_zeros = setting("ShowZeroValues");
    if let Some(map) = view.children().find(|n| {
        xml::is(*n, NS_CONFIG, "config-item-map-named")
            && n.attribute((NS_CONFIG, "name")) == Some("Tables")
    }) {
        for entry in map
            .children()
            .filter(|n| xml::is(*n, NS_CONFIG, "config-item-map-entry"))
        {
            budget::tick(1)?;
            let Some(name) = entry.attribute((NS_CONFIG, "name")) else {
                continue;
            };
            let number = |key: &str| item(entry, key).and_then(|v| v.parse::<u32>().ok());
            let mut settings = ViewSettings {
                show_grid: item(entry, "ShowGrid").map(|v| v == "true").or(global_grid),
                show_zeros: global_zeros,
                ..ViewSettings::default()
            };
            if number("HorizontalSplitMode") == Some(2) {
                settings.freeze_cols = number("HorizontalSplitPosition").unwrap_or(0);
            }
            if number("VerticalSplitMode") == Some(2) {
                settings.freeze_rows = number("VerticalSplitPosition").unwrap_or(0);
            }
            if let Some(color) = item(entry, "TabColor").and_then(|v| v.parse::<i64>().ok()) {
                // -1 (COL_AUTO) is no tab color.
                if color >= 0 {
                    settings.tab_color = Some(format!("{:06X}", color & 0xFF_FFFF));
                }
            }
            tables.insert(name.to_string(), settings);
        }
    }
    Ok(tables)
}

/// The column run holding 1-based column `col` (runs are contiguous and in
/// column order).
fn column_at(columns: &[Column], col: u32) -> Option<&Column> {
    let at = columns.partition_point(|c| c.last < col);
    columns.get(at).filter(|c| c.first <= col)
}

/// A copy of `cell` at (`row`, `col`).
fn copy_cell(cell: &Cell, row: u32, col: u32) -> Cell {
    Cell {
        col,
        row,
        value: copy_value(&cell.value),
        style_index: cell.style_index,
        formula: cell.formula.clone(),
        show_phonetic: cell.show_phonetic,
    }
}

fn copy_value(value: &CellValue) -> CellValue {
    match value {
        CellValue::Empty => CellValue::Empty,
        CellValue::Number { number } => CellValue::Number { number: *number },
        CellValue::Bool { bool } => CellValue::Bool { bool: *bool },
        CellValue::Error { error } => CellValue::Error {
            error: error.clone(),
        },
        CellValue::Shared { si } => CellValue::Shared { si: *si },
        CellValue::Text {
            text,
            runs,
            phonetic_runs,
            phonetic_pr,
        } => CellValue::Text {
            text: text.clone(),
            runs: runs.clone(),
            phonetic_runs: phonetic_runs
                .iter()
                .map(|r| xlsx_model::PhoneticRun {
                    sb: r.sb,
                    eb: r.eb,
                    text: r.text.clone(),
                })
                .collect(),
            phonetic_pr: phonetic_pr.clone(),
        },
    }
}

/// A run of rows: (first, last, height px, custom height, (hidden, outline level)).
type RowBlock = (u32, u32, u32, bool, (bool, u8));
struct Column {
    first: u32,
    last: u32,
    px: u32,
    hidden: bool,
    level: u8,
    default_xf: u32,
    /// The column's `table:default-cell-style-name`, resolved once.
    default_style: Option<CellStyle>,
}

impl Column {
    /// The style of a cell of this column that names none, nor its row.
    fn cell_style(&self) -> CellStyle {
        self.default_style.unwrap_or(CellStyle::Default)
    }
}

/// Drawings found in cells or `table:shapes`, resolved after the grid.
pub(crate) struct Anchored<'a, 'input> {
    pub node: roxmltree::Node<'a, 'input>,
    /// 0-based anchor cell, `None` for page-anchored shapes.
    pub cell: Option<(u32, u32)>,
}

pub(crate) fn read_table<'a, 'input>(
    ctx: &mut Context,
    package: &OdfPackage,
    table: roxmltree::Node<'a, 'input>,
    index: usize,
    view: Option<&ViewSettings>,
) -> Result<(OdsSheet, Vec<draw::PendingChart>), String> {
    let name = table
        .attribute((NS_TABLE, "name"))
        .unwrap_or("")
        .to_string();
    let table_chain = ctx
        .store
        .chain("table", table.attribute((NS_TABLE, "style-name")));
    let hidden = get(&table_chain, "table:table:display") == Some("false");
    let tab_color = get(&table_chain, "table:table:tab-color")
        .or_else(|| get(&table_chain, "table:other:tab-color"))
        .and_then(hex_color)
        .or_else(|| view.and_then(|v| v.tab_color.clone()));
    let right_to_left = get(&table_chain, "table:style:writing-mode") == Some("rl-tb");

    // ── Columns ────────────────────────────────────────────────────────
    let mut columns: Vec<Column> = Vec::new();
    let mut next_col = 1u32;
    collect_columns(ctx, table, 0, false, &mut columns, &mut next_col)?;
    if columns.is_empty() {
        columns.push(Column {
            first: 1,
            last: MAX_COLS,
            px: pt_to_px(DEFAULT_COL_PT),
            hidden: false,
            level: 0,
            default_xf: 0,
            default_style: None,
        });
    }
    promote_column_formats(ctx, table, &mut columns)?;
    // Column runs whose default cell style carries §16.3 conditions: the
    // only runs an unstyled blank cell must visit.
    let mapped_runs: Vec<usize> = columns
        .iter()
        .enumerate()
        .filter(|(_, c)| c.cell_style().has_maps(&ctx.store))
        .map(|(index, _)| index)
        .collect();

    // ── Rows ───────────────────────────────────────────────────────────
    let mut walker = RowWalker {
        rows: Vec::new(),
        heights: Vec::new(),
        merges: Vec::new(),
        hyperlinks: Vec::new(),
        comments: Vec::new(),
        anchored: Vec::new(),
        validations: Vec::new(),
        style_maps: Vec::new(),
        style_map_last: HashMap::new(),
        mapped_runs,
        next_row: 1,
    };
    walk_rows(ctx, table, 0, false, &columns, &mut walker)?;
    for shapes in table.children().filter(|n| xml::is(*n, NS_TABLE, "shapes")) {
        for node in shapes.children().filter(|n| n.is_element()) {
            budget::tick(1)?;
            walker.anchored.push(Anchored { node, cell: None });
        }
    }

    // ── Geometry ───────────────────────────────────────────────────────
    // Default row height: the height covering the most rows.
    // A hidden row is a zero-height row (ECMA-376 §18.3.1.73 `hidden`); when
    // most rows are hidden (LibreOffice's "hide the rest of the sheet") the
    // default height is zero, as SpreadsheetML's `zeroHeight`.
    for entry in &mut walker.heights {
        if entry.4 .0 {
            entry.2 = 0;
        }
    }
    let mut weight: HashMap<u32, u64> = HashMap::new();
    for (first, last, px, _, _) in &walker.heights {
        *weight.entry(*px).or_default() += u64::from(last - first + 1);
    }
    let default_row_px = weight
        .iter()
        .max_by_key(|(px, count)| (**count, std::cmp::Reverse(**px)))
        .map(|(px, _)| *px)
        .unwrap_or_else(|| pt_to_px(DEFAULT_ROW_PT));
    let default_row_pt = px_to_pt(default_row_px);
    let mut row_heights = BTreeMap::new();
    for (first, last, px, _, _) in &walker.heights {
        if *px != default_row_px {
            // The model keeps one height per row: one step each.
            for row in budget::each(*first, *last) {
                row_heights.insert(row?, px_to_pt(*px));
            }
        }
    }
    // Row records: content rows, plus outlined rows (the outline gutter
    // needs their level). Hidden rows without content are zero-height
    // rows of `row_heights`.
    let mut rows = std::mem::take(&mut walker.rows);
    {
        let mut existing: std::collections::BTreeSet<u32> = rows.iter().map(|r| r.index).collect();
        for (first, last, _, custom, (hidden, level)) in &walker.heights {
            if *level == 0 {
                continue;
            }
            for row in budget::each(*first, *last) {
                let row = row?;
                if existing.insert(row) {
                    budget::cell()?;
                    rows.push(Row {
                        index: row,
                        height: None,
                        custom_height: *custom,
                        cells: Vec::new(),
                        outline_level: *level,
                        collapsed: false,
                        hidden: *hidden,
                    });
                }
            }
        }
        rows.sort_by_key(|r| r.index);
        let mut blocks = walker.heights.iter().peekable();
        for row in &mut rows {
            while blocks.peek().is_some_and(|b| b.1 < row.index) {
                blocks.next();
            }
            if let Some((first, _, _, custom, (hidden, level))) = blocks.peek() {
                if *first <= row.index {
                    row.hidden = *hidden;
                    row.outline_level = *level;
                    row.custom_height = *custom;
                }
            }
            row.height = row_heights.get(&row.index).copied().filter(|h| *h > 0.0);
        }
    }
    let default_col_px = {
        let mut weight: HashMap<u32, u64> = HashMap::new();
        for c in &columns {
            *weight.entry(c.px).or_default() += u64::from(c.last - c.first + 1);
        }
        weight
            .into_iter()
            .max_by_key(|(px, count)| (*count, std::cmp::Reverse(*px)))
            .map(|(px, _)| px)
            .unwrap_or_else(|| pt_to_px(DEFAULT_COL_PT))
    };
    let mut worksheet = Worksheet::placeholder(&name, String::new());
    worksheet.parse_error = None;
    // Column metadata: the per-column maps the model keeps (one entry per
    // hidden or outlined column), and the style runs.
    for c in &columns {
        budget::tick(1)?;
        if c.hidden || c.level > 0 {
            for col in budget::each(c.first, c.last) {
                let col = col?;
                if c.hidden {
                    worksheet.col_hidden.insert(col, true);
                }
                if c.level > 0 {
                    worksheet.col_outline_levels.insert(col, c.level);
                }
            }
        }
        if c.default_xf != 0 {
            worksheet
                .col_style_ranges
                .push(xlsx_model::ColumnStyleRange {
                    min: c.first,
                    max: c.last,
                    style_index: c.default_xf,
                });
        }
    }
    // Hidden columns have no width in SpreadsheetML terms.
    let column_px: Vec<(u32, u32, u32)> = columns
        .iter()
        .map(|c| (c.first, c.last, if c.hidden { 0 } else { c.px }))
        .collect();
    // The sheet's copy of the workbook default font family.
    let font = ctx.xfs.default_font_model();
    worksheet.default_font_family = font.name.clone();
    worksheet.default_font_size = Some(if font.size > 0.0 { font.size } else { 10.0 });
    worksheet.default_font_bold = font.bold.then_some(true);
    worksheet.default_font_italic = font.italic.then_some(true);
    worksheet.normal_font_size = worksheet.default_font_size;
    worksheet.default_row_height = default_row_pt;
    worksheet.default_row_height_custom = true;
    worksheet.row_heights = row_heights;
    worksheet.merge_cells = walker.merges;
    worksheet.hyperlinks = walker.hyperlinks;
    worksheet.comment_refs = walker.comments.iter().map(|c| c.cell_ref.clone()).collect();
    worksheet.comments = walker.comments;
    worksheet.date1904 = ctx.date1904;
    worksheet.right_to_left = right_to_left;
    worksheet.tab_color = tab_color.clone();
    if let Some(view) = view {
        worksheet.freeze_rows = view.freeze_rows;
        worksheet.freeze_cols = view.freeze_cols;
        if let Some(grid) = view.show_grid {
            worksheet.show_gridlines = grid;
        }
        if let Some(zeros) = view.show_zeros {
            worksheet.show_zeros = zeros;
        }
    }
    if worksheet.col_outline_levels.values().any(|l| *l > 0)
        || rows.iter().any(|r| r.outline_level > 0)
    {
        worksheet.outline_pr = Some(xlsx_model::OutlinePr::default());
    }
    let geometry = draw::Grid::new(&column_px, default_col_px, &worksheet, default_row_px);
    let charts = draw::read_drawings(
        ctx,
        package,
        &walker.anchored,
        &geometry,
        &mut worksheet,
        index,
    )?;
    Ok((
        OdsSheet {
            name,
            visibility: if hidden {
                xlsx_model::SheetVisibility::Hidden
            } else {
                xlsx_model::SheetVisibility::Visible
            },
            tab_color,
            worksheet,
            rows,
            column_px,
            default_col_px,
            validation_refs: walker.validations,
            style_map_refs: walker.style_maps,
        },
        charts,
    ))
}

/// A column whose every cell, down to the last row, paints the same blank
/// format is a formatted column: LibreOffice writes a SpreadsheetML column
/// style as that format on each row's cells (its columns carry no default
/// cell style past the used columns). The column then takes that format as
/// its default, so the cells need not be materialized row by row.
fn promote_column_formats(
    ctx: &mut Context,
    table: roxmltree::Node<'_, '_>,
    columns: &mut Vec<Column>,
) -> Result<(), String> {
    let width = columns.last().map_or(0, |c| c.last) as usize;
    if width == 0 {
        return Ok(());
    }
    let mut rows: Vec<roxmltree::Node<'_, '_>> = Vec::new();
    collect_row_nodes(table, &mut rows)?;
    // Only a table that reaches the last row can format whole columns, and
    // only columns whose last row paints differently from the column
    // default are candidates; the scan below visits those columns only.
    let mut total = 0u64;
    for row in &rows {
        budget::tick(1)?;
        total = total.saturating_add(u64::from(repeat(*row, "number-rows-repeated")?));
    }
    if total < u64::from(MAX_ROWS) {
        return Ok(());
    }
    let mut candidates: Vec<(usize, usize)> = Vec::new();
    if let Some(last_row) = rows.last() {
        let row_default = last_row
            .attribute((NS_TABLE, "default-cell-style-name"))
            .map(|name| CellStyle::named(&ctx.store, name));
        let mut col = 1usize;
        for cell in last_row.children().filter(|n| {
            xml::is(*n, NS_TABLE, "table-cell") || xml::is(*n, NS_TABLE, "covered-table-cell")
        }) {
            if col > width {
                break;
            }
            let last = col
                .saturating_add(repeat(cell, "number-columns-repeated")? as usize - 1)
                .min(width);
            let own = cell
                .attribute((NS_TABLE, "style-name"))
                .map(|name| CellStyle::named(&ctx.store, name));
            let mut c = col;
            while c <= last {
                budget::tick(1)?;
                let column = column_at(columns, c as u32);
                let segment_last = column.map_or(last, |k| (k.last as usize).min(last));
                let style = own
                    .or(row_default)
                    .unwrap_or_else(|| column.map_or(CellStyle::Default, Column::cell_style));
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style)?;
                let default = column.map_or(0, |k| k.default_xf);
                if ctx.xfs.blank_signature(xf) != ctx.xfs.blank_signature(default) {
                    candidates.push((c, segment_last));
                }
                c = segment_last + 1;
            }
            col = last + 1;
        }
    }
    if candidates.is_empty() {
        return Ok(());
    }
    // Per column (at most 16,384): the blank xf seen so far (None:
    // conflicting), and rows.
    let mut seen: Vec<(Option<u32>, u32, bool)> = vec![(None, 0, true); width + 1];
    let mut next_row = 1u32;
    for row in rows {
        if next_row > MAX_ROWS {
            break;
        }
        let count = repeat(row, "number-rows-repeated")?.min(MAX_ROWS - next_row + 1);
        next_row += count;
        budget::tick(1)?;
        let row_default = row
            .attribute((NS_TABLE, "default-cell-style-name"))
            .map(|name| CellStyle::named(&ctx.store, name));
        let mut col = 1usize;
        for cell in row.children().filter(|n| {
            xml::is(*n, NS_TABLE, "table-cell") || xml::is(*n, NS_TABLE, "covered-table-cell")
        }) {
            if col > width {
                break;
            }
            let span = repeat(cell, "number-columns-repeated")? as usize;
            let last = col.saturating_add(span - 1).min(width);
            let has_content = cell.attribute((NS_OFFICE, "value-type")).is_some()
                || cell.children().any(|n| n.is_element());
            let own = cell
                .attribute((NS_TABLE, "style-name"))
                .map(|name| CellStyle::named(&ctx.store, name));
            let mut c = col;
            while c <= last {
                budget::tick(1)?;
                let column = column_at(columns, c as u32);
                let segment_last = column.map_or(last, |k| (k.last as usize).min(last));
                let style = own
                    .or(row_default)
                    .unwrap_or_else(|| column.map_or(CellStyle::Default, Column::cell_style));
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style)?;
                // Candidates are disjoint and in column order: only those
                // overlapping the segment are visited, each column once.
                let from = candidates.partition_point(|(_, b)| *b < c);
                let overlapping: Vec<(usize, usize)> = candidates[from..]
                    .iter()
                    .take_while(|(a, _)| *a <= segment_last)
                    .map(|(a, b)| ((*a).max(c), (*b).min(segment_last)))
                    .collect();
                for (a, b) in &overlapping {
                    budget::tick(b - a + 1)?;
                }
                let slots = overlapping.into_iter().flat_map(|(a, b)| a..=b);
                for index in slots {
                    let slot = &mut seen[index];
                    if has_content {
                        // Content cells keep their own format; they do not
                        // decide the column's.
                        slot.1 += count;
                        continue;
                    }
                    match slot.0 {
                        None if slot.2 && slot.1 == 0 => slot.0 = Some(xf),
                        Some(prev)
                            if ctx.xfs.blank_signature(prev) == ctx.xfs.blank_signature(xf) => {}
                        _ => slot.2 = false,
                    }
                    slot.1 += count;
                }
                c = segment_last + 1;
            }
            col = last + 1;
        }
    }
    if next_row <= MAX_ROWS {
        // The table ends before the last row: rows past it are unformatted.
        return Ok(());
    }
    // Split column runs where the promoted format changes.
    let mut promoted: Vec<Column> = Vec::new();
    for column in columns.drain(..) {
        let mut start = column.first;
        let format_of = |c: u32| -> Option<u32> {
            let slot = seen.get(c as usize)?;
            (slot.2 && slot.1 >= MAX_ROWS).then_some(slot.0).flatten()
        };
        let mut current = format_of(start);
        for c in budget::each(column.first, column.last + 1) {
            let c = c?;
            let next = if c <= column.last { format_of(c) } else { None };
            if c > column.last || next != current {
                let xf = current.filter(|xf| {
                    ctx.xfs.blank_signature(*xf) != ctx.xfs.blank_signature(column.default_xf)
                });
                promoted.push(Column {
                    first: start,
                    last: c - 1,
                    px: column.px,
                    hidden: column.hidden,
                    level: column.level,
                    default_xf: xf.unwrap_or(column.default_xf),
                    // Unstyled cells still resolve through the authored
                    // default; every blank one paints like `xf`.
                    default_style: column.default_style,
                });
                start = c;
                current = next;
            }
        }
    }
    *columns = promoted;
    Ok(())
}

fn collect_row_nodes<'a, 'input>(
    parent: roxmltree::Node<'a, 'input>,
    out: &mut Vec<roxmltree::Node<'a, 'input>>,
) -> Result<(), String> {
    for node in parent.children().filter(|n| n.is_element()) {
        budget::tick(1)?;
        if node.tag_name().namespace() != Some(NS_TABLE) {
            continue;
        }
        match node.tag_name().name() {
            "table-row-group" | "table-header-rows" | "table-rows" => collect_row_nodes(node, out)?,
            "table-row" => out.push(node),
            _ => {}
        }
    }
    Ok(())
}

fn px_to_pt(px: u32) -> f64 {
    f64::from(px) * 72.0 / 96.0
}

/// A repeat or span count (§19.675 / §19.676 / §19.680 / §19.681): one
/// beyond the implementation bound rejects instead of becoming 1.
fn repeat(node: roxmltree::Node<'_, '_>, name: &str) -> Result<u32, String> {
    xml::count_attribute(node, NS_TABLE, name)
}

fn collect_columns(
    ctx: &mut Context,
    parent: roxmltree::Node<'_, '_>,
    level: u8,
    group_hidden: bool,
    columns: &mut Vec<Column>,
    next: &mut u32,
) -> Result<(), String> {
    for node in parent.children().filter(|n| n.is_element()) {
        if node.tag_name().namespace() != Some(NS_TABLE) {
            continue;
        }
        match node.tag_name().name() {
            "table-column-group" => {
                let collapsed = node.attribute((NS_TABLE, "display")) == Some("false");
                collect_columns(
                    ctx,
                    node,
                    level.saturating_add(1).min(7),
                    group_hidden || collapsed,
                    columns,
                    next,
                )?;
            }
            "table-header-columns" | "table-columns" => {
                collect_columns(ctx, node, level, group_hidden, columns, next)?
            }
            "table-column" => {
                if *next > MAX_COLS {
                    continue;
                }
                let count = repeat(node, "number-columns-repeated")?;
                let last = next.saturating_add(count - 1).min(MAX_COLS);
                let chain = ctx
                    .store
                    .chain("table-column", node.attribute((NS_TABLE, "style-name")));
                let width = get(&chain, "col:style:column-width")
                    .and_then(length_pt)
                    .unwrap_or(DEFAULT_COL_PT);
                let visibility = node.attribute((NS_TABLE, "visibility"));
                budget::tick(1)?;
                let default_style = node
                    .attribute((NS_TABLE, "default-cell-style-name"))
                    .map(|name| CellStyle::named(&ctx.store, name));
                let default_xf = ctx.xfs.xf_for(
                    &ctx.store,
                    &mut ctx.data,
                    &ctx.faces,
                    default_style.unwrap_or(CellStyle::Default),
                )?;
                columns.push(Column {
                    first: *next,
                    last,
                    px: pt_to_px(width),
                    hidden: matches!(visibility, Some("collapse") | Some("filter"))
                        || group_hidden && visibility != Some("visible"),
                    level,
                    default_xf,
                    default_style,
                });
                *next = last + 1;
            }
            _ => {}
        }
    }
    Ok(())
}

struct RowWalker<'a, 'input> {
    rows: Vec<Row>,
    /// (first row, last row, px, custom height, (hidden, outline level)).
    heights: Vec<RowBlock>,
    merges: Vec<MergeCell>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<xlsx_model::XlsxComment>,
    anchored: Vec<Anchored<'a, 'input>>,
    validations: Vec<(String, xlsx_model::CellRange)>,
    /// Ranges of cells whose style has `style:map` conditions.
    style_maps: Vec<(CellStyle, xlsx_model::CellRange)>,
    /// (style, first column, last column) → its last range, to extend it
    /// down over the next rows.
    style_map_last: HashMap<(CellStyle, u32, u32), usize>,
    /// Indices of the column runs whose default style has conditions.
    mapped_runs: Vec<usize>,
    next_row: u32,
}

fn walk_rows<'a, 'input>(
    ctx: &mut Context,
    parent: roxmltree::Node<'a, 'input>,
    level: u8,
    group_hidden: bool,
    columns: &[Column],
    walker: &mut RowWalker<'a, 'input>,
) -> Result<(), String> {
    for node in parent.children().filter(|n| n.is_element()) {
        budget::tick(1)?;
        if node.tag_name().namespace() != Some(NS_TABLE) {
            continue;
        }
        match node.tag_name().name() {
            "table-row-group" => {
                let collapsed = node.attribute((NS_TABLE, "display")) == Some("false");
                walk_rows(
                    ctx,
                    node,
                    level.saturating_add(1).min(7),
                    group_hidden || collapsed,
                    columns,
                    walker,
                )?;
            }
            "table-header-rows" | "table-rows" => {
                walk_rows(ctx, node, level, group_hidden, columns, walker)?
            }
            "table-row" => read_row(ctx, node, level, group_hidden, columns, walker)?,
            _ => {}
        }
    }
    Ok(())
}

fn read_row<'a, 'input>(
    ctx: &mut Context,
    node: roxmltree::Node<'a, 'input>,
    level: u8,
    group_hidden: bool,
    columns: &[Column],
    walker: &mut RowWalker<'a, 'input>,
) -> Result<(), String> {
    let first = walker.next_row;
    if first > MAX_ROWS {
        return Ok(());
    }
    let count = repeat(node, "number-rows-repeated")?;
    let last = first.saturating_add(count - 1).min(MAX_ROWS);
    walker.next_row = last.saturating_add(1);
    let chain = ctx
        .store
        .chain("table-row", node.attribute((NS_TABLE, "style-name")));
    let height = get(&chain, "row:style:row-height")
        .or_else(|| get(&chain, "row:style:min-row-height"))
        .and_then(length_pt)
        .unwrap_or(DEFAULT_ROW_PT);
    let optimal = get(&chain, "row:style:use-optimal-row-height") == Some("true");
    let visibility = node.attribute((NS_TABLE, "visibility"));
    let hidden = matches!(visibility, Some("collapse") | Some("filter"))
        || group_hidden && visibility != Some("visible");
    walker
        .heights
        .push((first, last, pt_to_px(height), !optimal, (hidden, level)));
    let row_default = node
        .attribute((NS_TABLE, "default-cell-style-name"))
        .map(|name| CellStyle::named(&ctx.store, name));

    // The cells of one occurrence of this row: (column, the cell to copy).
    let mut templates: Vec<(u32, Rc<Cell>)> = Vec::new();
    let mut col = 1u32;
    for cell in node.children().filter(|n| n.is_element()) {
        budget::tick(1)?;
        let covered = xml::is(cell, NS_TABLE, "covered-table-cell");
        if !covered && !xml::is(cell, NS_TABLE, "table-cell") {
            continue;
        }
        if col > MAX_COLS {
            break;
        }
        let repeat_count = repeat(cell, "number-columns-repeated")?;
        let cell_last = col.saturating_add(repeat_count - 1).min(MAX_COLS);
        // The style the cell or its row names; a cell that names none takes
        // each column's default.
        let explicit = cell
            .attribute((NS_TABLE, "style-name"))
            .map(|name| CellStyle::named(&ctx.store, name))
            .or(row_default);
        let CellRead {
            value,
            formula,
            shown_empty_zero,
        } = cell_value(ctx, cell, explicit, col, columns)?;
        let has_value = !matches!(value, CellValue::Empty);
        // A blank cell (no value type, no content, no own or row style)
        // takes its column's default everywhere: it paints nothing of its
        // own, so only runs whose default style has conditions are visited.
        let unstyled_blank = !has_value
            && formula.is_none()
            && explicit.is_none()
            && cell.attribute((NS_OFFICE, "value-type")).is_none()
            && !cell.children().any(|n| n.is_element());
        if unstyled_blank {
            let mut at = walker
                .mapped_runs
                .partition_point(|run| columns[*run].last < col);
            while let Some(run) = walker
                .mapped_runs
                .get(at)
                .copied()
                .filter(|run| columns[*run].first <= cell_last)
            {
                budget::tick(1)?;
                let column = &columns[run];
                record_style_map(
                    walker,
                    column.cell_style(),
                    (first, column.first.max(col)),
                    (last, column.last.min(cell_last)),
                )?;
                at += 1;
            }
            // Columns past the declared runs take the `Default` style.
            let declared = columns.last().map_or(0, |c| c.last);
            if cell_last > declared && CellStyle::Default.has_maps(&ctx.store) {
                record_style_map(
                    walker,
                    CellStyle::Default,
                    (first, col.max(declared + 1)),
                    (last, cell_last),
                )?;
            }
        }
        // The cell of the last format seen in this authored cell; the
        // columns it repeats over share it.
        let mut prototype: Option<Rc<Cell>> = None;
        // The column style the prototype's rich-text runs were read over.
        let mut prototype_column: Option<CellStyle> = None;
        let rich = matches!(&value, CellValue::Text { runs: Some(_), .. }) && explicit.is_none();
        // Walk the repeated cells as segments of constant column default.
        let mut c = if unstyled_blank { cell_last + 1 } else { col };
        while c <= cell_last {
            budget::tick(1)?;
            let column = column_at(columns, c);
            let segment_last = column.map_or(cell_last, |k| k.last.min(cell_last));
            let column_xf = column.map_or(0, |k| k.default_xf);
            let column_style = column.map_or(CellStyle::Default, Column::cell_style);
            let style = explicit.unwrap_or(column_style);
            if style.has_maps(&ctx.store) {
                record_style_map(walker, style, (first, c), (last, segment_last))?;
            }
            let xf = ctx
                .xfs
                .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style)?;
            let xf = adjust_xf_for_value(&mut ctx.xfs, xf, cell, shown_empty_zero)?;
            // An empty cell is materialized only when it paints differently
            // from its column's default (fill or borders); its font, number
            // format and alignment have nothing to show.
            let paints = ctx.xfs.blank_signature(xf) != ctx.xfs.blank_signature(column_xf);
            if has_value || paints || formula.is_some() {
                let template = match prototype.take() {
                    // Runs read over a column style serve only that style.
                    Some(t)
                        if t.style_index == Some(xf)
                            && (!rich || prototype_column == Some(column_style)) =>
                    {
                        t
                    }
                    _ => {
                        // Rich text of a cell without its own or a row
                        // style resolves its runs over each column's
                        // default style: a later column with another
                        // style reads them again.
                        let value = if rich && c != col {
                            cell_value(ctx, cell, explicit, c, columns)?.value
                        } else {
                            copy_value(&value)
                        };
                        Rc::new(Cell {
                            col,
                            row: first,
                            value,
                            style_index: Some(xf),
                            formula: formula.clone(),
                            show_phonetic: false,
                        })
                    }
                };
                for c in budget::each(c, segment_last) {
                    budget::cell()?;
                    templates.push((c?, template.clone()));
                }
                prototype = Some(template);
                prototype_column = Some(column_style);
            }
            c = segment_last + 1;
        }
        // §19.676 / §19.678 spans: a merge from each cell of the block. The
        // span is read once; a cell without one costs nothing here.
        if !covered {
            let span_cols = repeat(cell, "number-columns-spanned")?;
            let span_rows = repeat(cell, "number-rows-spanned")?;
            if span_cols > 1 || span_rows > 1 {
                for row in budget::each(first, last) {
                    let row = row?;
                    for c in budget::each(col, cell_last) {
                        let c = c?;
                        walker.merges.push(MergeCell {
                            top: row,
                            left: c,
                            bottom: row.saturating_add(span_rows - 1).min(MAX_ROWS),
                            right: c.saturating_add(span_cols - 1).min(MAX_COLS),
                        });
                    }
                }
            }
        }
        if let Some(name) = cell.attribute((NS_TABLE, "content-validation-name")) {
            walker.validations.push((
                name.to_string(),
                xlsx_model::CellRange {
                    top: first,
                    left: col,
                    bottom: last,
                    right: cell_last,
                },
            ));
        }
        // Links and annotations are cell content: every repeated cell of
        // the block carries its own copy (§19.675 / §19.680), one step
        // each. Drawings are anchored once, at the block's first cell, as
        // LibreOffice anchors them.
        if !covered || cell.children().any(|n| n.is_element()) {
            let link = cell
                .descendants()
                .find(|n| xml::is(*n, NS_TEXT, "a"))
                .and_then(|a| a.attribute((NS_XLINK, "href")))
                .map(|href| hyperlink(col, first, href));
            let note = match cell
                .children()
                .find(|n| xml::is(*n, NS_OFFICE, "annotation"))
            {
                Some(a) => Some(comment(a, super::paragraph_text(a)?, first, col)),
                None => None,
            };
            if link.is_some() || note.is_some() {
                for row in budget::each(first, last) {
                    let row = row?;
                    for c in budget::each(col, cell_last) {
                        let c = c?;
                        if let Some(link) = &link {
                            walker.hyperlinks.push(Hyperlink {
                                col: c,
                                row,
                                ..link.clone()
                            });
                        }
                        if let Some(note) = &note {
                            walker.comments.push(xlsx_model::XlsxComment {
                                cell_ref: format!("{}{}", column_letters(c), row),
                                ..note.clone()
                            });
                        }
                    }
                }
            }
            for child in cell.children().filter(|n| is_drawing(*n)) {
                budget::tick(1)?;
                walker.anchored.push(Anchored {
                    node: child,
                    cell: Some((first - 1, col - 1)),
                });
            }
        }
        col = cell_last + 1;
    }
    if !templates.is_empty() {
        for row in budget::each(first, last) {
            let row = row?;
            let mut cells: Vec<Cell> = Vec::with_capacity(templates.len());
            for (c, template) in &templates {
                // The first occurrence was counted with its templates.
                if row != first {
                    budget::cell()?;
                }
                cells.push(copy_cell(template, row, *c));
            }
            walker.rows.push(Row {
                index: row,
                height: None,
                custom_height: false,
                cells,
                outline_level: 0,
                collapsed: false,
                hidden: false,
            });
        }
    }
    Ok(())
}

/// Note the cells (rows × columns) whose style carries `style:map`
/// conditions, extending the style's range from the rows above.
fn record_style_map(
    walker: &mut RowWalker<'_, '_>,
    style: CellStyle,
    (top, left): (u32, u32),
    (bottom, right): (u32, u32),
) -> Result<(), String> {
    let key = (style, left, right);
    if let Some(index) = walker.style_map_last.get(&key) {
        let range = &mut walker.style_maps[*index].1;
        if range.bottom.saturating_add(1) == top {
            range.bottom = bottom;
            return Ok(());
        }
    }
    budget::tick(1)?;
    walker.style_map_last.insert(key, walker.style_maps.len());
    walker.style_maps.push((
        style,
        xlsx_model::CellRange {
            top,
            left,
            bottom,
            right,
        },
    ));
    Ok(())
}

/// The format of a formula result LibreOffice displays as nothing: the
/// numeric zero stays a number and its zero section is empty (§18.8.30).
const HIDDEN_ZERO_CODE: &str = "0;-0;;@";

/// A cell whose value type has no matching data style is shown as
/// LibreOffice shows that type (§19.385): dates and times with the default
/// ISO-like date / time formats, percentages with `0%` decimals of the value.
///
/// A formula whose cached result is the number 0 with an empty displayed
/// paragraph is a result LibreOffice shows as nothing: an empty-cell or
/// empty-string result (a lookup that lands on an empty cell,
/// `IFERROR(…;"")` kept from a SpreadsheetML import), or a zero its format
/// hides. Measured with LibreOffice 26.8: all three are written as
/// `office:value="0"` with `<text:p/>`, while a displayed zero is written
/// with `<text:p>0</text:p>`. The value stays the number LibreOffice
/// calculates with; only its zero section is emptied.
fn adjust_xf_for_value(
    xfs: &mut XfBuilder,
    xf: u32,
    cell: roxmltree::Node<'_, '_>,
    shown_empty_zero: bool,
) -> Result<u32, String> {
    if shown_empty_zero {
        return xfs.with_num_fmt(xf, HIDDEN_ZERO_CODE, DataKind::Number);
    }
    let kind = xfs.kind(xf);
    match cell.attribute((NS_OFFICE, "value-type")) {
        Some("date") if !matches!(kind, Some(DataKind::Date) | Some(DataKind::Time)) => {
            let has_time = cell
                .attribute((NS_OFFICE, "date-value"))
                .is_some_and(|v| v.contains('T'));
            let code = if has_time {
                "yyyy-mm-dd hh:mm:ss"
            } else {
                "yyyy-mm-dd"
            };
            xfs.with_num_fmt(xf, code, DataKind::Date)
        }
        Some("time") if !matches!(kind, Some(DataKind::Date) | Some(DataKind::Time)) => {
            xfs.with_num_fmt(xf, "[hh]:mm:ss", DataKind::Time)
        }
        Some("percentage") if kind.is_none() => xfs.with_num_fmt(xf, "0%", DataKind::Percentage),
        _ => Ok(xf),
    }
}

/// Whether the cell has displayed paragraphs and they hold nothing.
fn paragraphs_show_nothing(cell: roxmltree::Node<'_, '_>) -> bool {
    let mut paragraphs = cell
        .children()
        .filter(|n| xml::is(*n, NS_TEXT, "p") || xml::is(*n, NS_TEXT, "h"))
        .peekable();
    paragraphs.peek().is_some()
        && paragraphs.all(|p| {
            !p.descendants().any(|n| {
                n.is_text() && n.text().is_some_and(|t| !t.is_empty())
                    || n.tag_name().namespace() == Some(NS_TEXT)
                        && matches!(n.tag_name().name(), "s" | "tab" | "line-break")
            })
        })
}

fn is_drawing(node: roxmltree::Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().namespace() == Some(crate::xml::NS_DRAW)
}

fn hyperlink(col: u32, row: u32, href: &str) -> Hyperlink {
    if let Some(location) = href.strip_prefix('#') {
        // `#Sheet2.A1` → `Sheet2!A1`.
        let location = match super::split_sheet(location) {
            (Some(sheet), cell) => format!("{}{}", super::sheet_prefix(&sheet), cell),
            (None, cell) => cell.to_string(),
        };
        Hyperlink {
            col,
            row,
            url: None,
            location: Some(location),
            display: None,
        }
    } else {
        Hyperlink {
            col,
            row,
            url: Some(href.to_string()),
            location: None,
            display: None,
        }
    }
}

fn comment(
    annotation: roxmltree::Node<'_, '_>,
    text: String,
    row: u32,
    col: u32,
) -> xlsx_model::XlsxComment {
    let author = annotation
        .children()
        .find(|n| xml::is(*n, NS_DC, "creator"))
        .and_then(|n| n.text())
        .map(str::to_string);
    let date = annotation
        .children()
        .find(|n| xml::is(*n, NS_DC, "date"))
        .and_then(|n| n.text())
        .map(str::to_string);
    xlsx_model::XlsxComment {
        kind: xlsx_model::XlsxCommentKind::Note,
        cell_ref: format!("{}{}", column_letters(col), row),
        id: None,
        person_id: None,
        author,
        date,
        root_text: None,
        text,
        resolved: None,
        replies: Vec::new(),
    }
}

/// What one authored cell holds.
struct CellRead {
    value: CellValue,
    formula: Option<String>,
    /// A formula's numeric zero that LibreOffice displays as nothing (see
    /// [`adjust_xf_for_value`]).
    shown_empty_zero: bool,
}

/// The value of a cell (§19.385 value types) and its formula. `explicit`
/// is the style the cell or its row names; without one the cell at `col`
/// takes its column's default.
fn cell_value(
    ctx: &mut Context,
    cell: roxmltree::Node<'_, '_>,
    explicit: Option<CellStyle>,
    col: u32,
    columns: &[Column],
) -> Result<CellRead, String> {
    let formula = cell
        .attribute((NS_TABLE, "formula"))
        .map(super::extras::convert_formula);
    let value_type = cell.attribute((NS_OFFICE, "value-type"));
    let calc_type = cell.attribute((NS_CALCEXT, "value-type"));
    let number = |name: &str| {
        cell.attribute((NS_OFFICE, name))
            .and_then(|v| v.trim().parse::<f64>().ok())
    };
    let read = |value: CellValue, formula: Option<String>| CellRead {
        value,
        formula,
        shown_empty_zero: false,
    };
    if calc_type == Some("error") {
        let error = super::paragraph_text(cell)?;
        return Ok(read(CellValue::Error { error }, formula));
    }
    let value = match value_type {
        Some("float") | Some("percentage") | Some("currency") => match number("value") {
            Some(n) => {
                let shown_empty_zero =
                    n == 0.0 && formula.is_some() && paragraphs_show_nothing(cell);
                return Ok(CellRead {
                    value: CellValue::Number { number: n },
                    formula,
                    shown_empty_zero,
                });
            }
            None => CellValue::Empty,
        },
        Some("date") => match cell
            .attribute((NS_OFFICE, "date-value"))
            .and_then(|v| date_serial(v, ctx.date1904))
        {
            Some(n) => CellValue::Number { number: n },
            None => CellValue::Empty,
        },
        Some("time") => match cell
            .attribute((NS_OFFICE, "time-value"))
            .and_then(duration_days)
        {
            Some(n) => CellValue::Number { number: n },
            None => CellValue::Empty,
        },
        Some("boolean") => match cell.attribute((NS_OFFICE, "boolean-value")) {
            Some("true") | Some("1") => CellValue::Bool { bool: true },
            Some(_) => CellValue::Bool { bool: false },
            None => CellValue::Empty,
        },
        Some("string") | Some("void") | None => {
            let has_paragraphs = cell
                .children()
                .any(|n| xml::is(n, NS_TEXT, "p") || xml::is(n, NS_TEXT, "h"));
            if let Some(text) = cell.attribute((NS_OFFICE, "string-value")) {
                text_value(text.to_string(), None)
            } else if has_paragraphs {
                let style = explicit.unwrap_or_else(|| {
                    column_at(columns, col).map_or(CellStyle::Default, Column::cell_style)
                });
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style)?;
                let (text, runs) = rich_text(ctx, cell, xf, style)?;
                if value_type.is_none() && text.is_empty() {
                    CellValue::Empty
                } else {
                    text_value(text, runs)
                }
            } else if value_type == Some("string") {
                text_value(String::new(), None)
            } else {
                CellValue::Empty
            }
        }
        Some(_) => CellValue::Empty,
    };
    Ok(read(value, formula))
}

fn text_value(text: String, runs: Option<Vec<Run>>) -> CellValue {
    CellValue::Text {
        text,
        runs,
        phonetic_runs: Vec::new(),
        phonetic_pr: None,
    }
}

/// Cell text with runs (§6.1.7) when a paragraph or `text:span` style
/// formats text. Each formatting context resolves what its own style chain
/// sets once, over what the context around it sets (the cell's style, its
/// paragraph, the enclosing spans: §16.2), so nesting costs one style chain
/// per span. A run's font is that, then the default style of the
/// innermost context's family (the initial family), then the table-cell
/// default where LibreOffice keeps the document font. Each run's font is a
/// copy (its family name included).
fn rich_text(
    ctx: &mut Context,
    cell: roxmltree::Node<'_, '_>,
    xf: u32,
    style: CellStyle,
) -> Result<(String, Option<Vec<Run>>), String> {
    let cell_font = ctx.xfs.font_of_xf(xf);
    let cell_default: Vec<&crate::odf_props::Style> =
        ctx.store.default_style("table-cell").into_iter().collect();
    let document_font = super::cells::font_of(&cell_default, &ctx.faces, None);
    // A percentage of the cell's style scales the table-cell default, as
    // the cell's own font does, whatever context its text sits in.
    let cell_partial = Rc::new(TextContext {
        set: PartialFont::of(
            &ctx.store
                .chain_own_of("table-cell", style.style(&ctx.store)),
            &ctx.faces,
        )
        .with_size_base(document_font.size),
        family: "table-cell",
    });
    // A paragraph without a style of its own keeps the cell's formatting:
    // Calc formats cell text from the cell style, and LibreOffice's
    // paragraph-family default (where it keeps its Writer fonts) never
    // reaches unstyled cell text.
    let mut runs: Vec<(String, Option<Rc<TextContext>>)> = Vec::new();
    let mut first = true;
    for p in cell
        .children()
        .filter(|n| xml::is(*n, NS_TEXT, "p") || xml::is(*n, NS_TEXT, "h"))
    {
        budget::tick(1)?;
        if !first {
            push_run(&mut runs, "\n", &None);
        }
        first = false;
        let paragraph = p
            .attribute((NS_TEXT, "style-name"))
            .and_then(|name| text_context(ctx, "paragraph", name, &cell_partial));
        collect_runs(ctx, p, &paragraph, &cell_partial, &mut runs)?;
    }
    let text: String = runs.iter().map(|(t, _)| t.as_str()).collect();
    if runs.iter().all(|(_, context)| context.is_none()) {
        return Ok((text, None));
    }
    let family_default = |family: &str| {
        let default: Vec<&crate::odf_props::Style> =
            ctx.store.default_style(family).into_iter().collect();
        PartialFont::of(&default, &ctx.faces)
    };
    let defaults = [
        ("paragraph", family_default("paragraph")),
        ("text", family_default("text")),
    ];
    let mut out = Vec::with_capacity(runs.len());
    for (text, context) in runs {
        budget::tick(1)?;
        let font = match &context {
            Some(context) => {
                let default = defaults
                    .iter()
                    .find(|(family, _)| *family == context.family)
                    .map(|(_, d)| d.clone())
                    .unwrap_or_default();
                super::cells::run_font_of(&context.set.over(&default).realize(&document_font))
            }
            None => super::cells::run_font_of(&cell_font),
        };
        out.push(Run {
            text,
            font: Some(font),
        });
    }
    Ok((text, Some(out)))
}

/// A formatting context of cell text: what it and the contexts around it
/// set, and its style family.
struct TextContext {
    set: PartialFont,
    family: &'static str,
}

/// The context of style `name` of `family` inside `outer`. `None` when the
/// document does not define the style (the context changes nothing).
fn text_context(
    ctx: &Context,
    family: &'static str,
    name: &str,
    outer: &TextContext,
) -> Option<Rc<TextContext>> {
    ctx.store.find(family, name)?;
    let chain = ctx.store.chain_own(family, name);
    Some(Rc::new(TextContext {
        set: PartialFont::of(&chain, &ctx.faces).over(&outer.set),
        family,
    }))
}

fn push_run(
    runs: &mut Vec<(String, Option<Rc<TextContext>>)>,
    text: &str,
    context: &Option<Rc<TextContext>>,
) {
    if let Some(last) = runs.last_mut() {
        let same = match (&last.1, context) {
            (None, None) => true,
            (Some(a), Some(b)) => Rc::ptr_eq(a, b),
            _ => false,
        };
        if same {
            last.0.push_str(text);
            return;
        }
    }
    runs.push((text.to_string(), context.clone()));
}

fn collect_runs(
    ctx: &Context,
    node: roxmltree::Node<'_, '_>,
    context: &Option<Rc<TextContext>>,
    cell: &Rc<TextContext>,
    runs: &mut Vec<(String, Option<Rc<TextContext>>)>,
) -> Result<(), String> {
    for child in node.children() {
        // One step per node: each may add a run or a formatting context.
        budget::tick(1)?;
        if child.is_text() {
            push_run(runs, child.text().unwrap_or(""), context);
            continue;
        }
        if child.tag_name().namespace() != Some(NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            "s" => {
                let mut spaces = String::new();
                budget::push_spaces(&mut spaces, crate::odf_text::space_count(child))?;
                push_run(runs, &spaces, context);
            }
            "tab" => push_run(runs, "\t", context),
            "line-break" => push_run(runs, "\n", context),
            "span" => {
                let outer = context.as_ref().unwrap_or(cell);
                let inner = child
                    .attribute((NS_TEXT, "style-name"))
                    .and_then(|name| text_context(ctx, "text", name, outer));
                match inner {
                    Some(inner) => collect_runs(ctx, child, &Some(inner), cell, runs)?,
                    None => collect_runs(ctx, child, context, cell, runs)?,
                }
            }
            "note" | "ruby-text" => {}
            _ => collect_runs(ctx, child, context, cell, runs)?,
        }
    }
    Ok(())
}
